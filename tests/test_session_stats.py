from __future__ import annotations

import os
import json
import sys
import tempfile
import unittest
import urllib.error
from io import BytesIO
from pathlib import Path

BACKEND_ROOT = Path(__file__).resolve().parents[1] / "web" / "backend"
sys.path.insert(0, str(BACKEND_ROOT))

from app.session_stats import OpenAIExtractor, SessionStatsService, SessionTextFile


class FakeExtractor:
    def __init__(self, payloads):
        self.payloads = list(payloads)

    def extract(self, *, filename: str, content: str, month: str):
        if not self.payloads:
            raise AssertionError("fake extractor payload exhausted")
        return self.payloads.pop(0)


class SessionStatsServiceTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.db_path = Path(self.tmp.name) / "stats.sqlite"
        self.service = SessionStatsService(self.db_path)

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def test_import_counts_kp_as_player_with_extra_host_dimension(self):
        extractor = FakeExtractor([
            {
                "title": "三黄门风云",
                "duration_hours": 7,
                "kp": {"name": "汐蒂仙", "qq": "1950566630"},
                "players": [{"name": "白穆"}, {"name": "黑猫约娜"}],
                "confidence": 0.93,
                "warnings": [],
            }
        ])

        result = self.service.import_text_files(
            month="2026-05",
            files=[SessionTextFile(filename="三黄门风云.txt", content="结团原文")],
            extractor=extractor,
            model_name="fake-model",
        )

        self.assertEqual(result["success_count"], 1)
        self.assertEqual(result["failure_count"], 0)

        overview = self.service.get_overview("2026-05")
        self.assertEqual(overview["session_count"], 1)
        self.assertEqual(overview["participant_count"], 3)
        self.assertAlmostEqual(overview["total_game_hours"], 21.0)
        self.assertAlmostEqual(overview["total_host_hours"], 7.0)

        players = {row["name"]: row for row in self.service.list_players("2026-05")}
        self.assertEqual(players["汐蒂仙"]["qq"], "1950566630")
        self.assertEqual(players["汐蒂仙"]["game_count"], 1)
        self.assertAlmostEqual(players["汐蒂仙"]["game_hours"], 7.0)
        self.assertEqual(players["汐蒂仙"]["reincarnation_count"], 1)
        self.assertEqual(players["汐蒂仙"]["host_count"], 1)
        self.assertAlmostEqual(players["汐蒂仙"]["host_hours"], 7.0)

        self.assertEqual(players["白穆"]["game_count"], 1)
        self.assertEqual(players["白穆"]["reincarnation_count"], 1)
        self.assertEqual(players["白穆"]["host_count"], 0)

        sessions = self.service.list_sessions("2026-05")
        self.assertEqual(len(sessions), 1)
        self.assertEqual(sessions[0]["title"], "三黄门风云")
        self.assertEqual(sessions[0]["kp_name"], "汐蒂仙")
        self.assertEqual(sessions[0]["pl_count"], 2)

    def test_duplicate_file_content_is_not_counted_twice(self):
        payload = {
            "title": "重复团",
            "duration_hours": 5,
            "kp": {"name": "主持", "qq": None},
            "players": [{"name": "玩家"}],
            "confidence": 0.9,
            "warnings": [],
        }

        first = self.service.import_text_files(
            month="2026-05",
            files=[SessionTextFile(filename="重复.txt", content="同一份原文")],
            extractor=FakeExtractor([payload]),
            model_name="fake-model",
        )
        second = self.service.import_text_files(
            month="2026-05",
            files=[SessionTextFile(filename="重复.txt", content="同一份原文")],
            extractor=FakeExtractor([]),
            model_name="fake-model",
        )

        self.assertEqual(first["success_count"], 1)
        self.assertEqual(second["success_count"], 0)
        self.assertEqual(second["failure_count"], 0)
        self.assertEqual(second["skip_count"], 1)
        self.assertTrue(second["items"][0]["skipped"])
        self.assertIn("重复导入", second["items"][0]["reason"])
        self.assertEqual(self.service.get_overview("2026-05")["session_count"], 1)

    def test_missing_required_fields_goes_to_import_errors(self):
        result = self.service.import_text_files(
            month="2026-05",
            files=[SessionTextFile(filename="坏文件.txt", content="缺字段")],
            extractor=FakeExtractor([
                {
                    "title": "坏文件",
                    "duration_hours": None,
                    "kp": {"name": "主持", "qq": None},
                    "players": [{"name": "玩家"}],
                    "confidence": 0.4,
                    "warnings": ["缺少时长"],
                }
            ]),
            model_name="fake-model",
        )

        self.assertEqual(result["success_count"], 0)
        self.assertEqual(result["failure_count"], 1)
        self.assertEqual(self.service.get_overview("2026-05")["session_count"], 0)
        errors = self.service.list_errors("2026-05")
        self.assertEqual(len(errors), 1)
        self.assertEqual(errors[0]["filename"], "坏文件.txt")
        self.assertIn("时长", errors[0]["reason"])
        self.assertIn("坏文件", errors[0]["raw_payload"])

    def test_invalid_player_item_does_not_partially_insert_session(self):
        result = self.service.import_text_files(
            month="2026-05",
            files=[SessionTextFile(filename="坏玩家.txt", content="坏玩家原文")],
            extractor=FakeExtractor([
                {
                    "title": "坏玩家",
                    "duration_hours": 2,
                    "kp": {"name": "主持", "qq": None},
                    "players": [{"name": "玩家"}, {}],
                    "confidence": 0.4,
                    "warnings": [],
                }
            ]),
            model_name="fake-model",
        )

        self.assertEqual(result["success_count"], 0)
        self.assertEqual(result["failure_count"], 1)
        self.assertIn("玩家", result["items"][0]["reason"])
        self.assertEqual(self.service.get_overview("2026-05")["session_count"], 0)
        self.assertEqual(self.service.list_players("2026-05"), [])

    def test_kp_with_same_qq_is_merged_even_when_name_changes(self):
        self.service.import_text_files(
            month="2026-05",
            files=[SessionTextFile(filename="一.txt", content="原文1")],
            extractor=FakeExtractor([
                {
                    "title": "第一团",
                    "duration_hours": 4,
                    "kp": {"name": "汐蒂仙", "qq": "1950566630"},
                    "players": [{"name": "白穆"}],
                    "confidence": 0.9,
                    "warnings": [],
                }
            ]),
            model_name="fake-model",
        )
        self.service.import_text_files(
            month="2026-05",
            files=[SessionTextFile(filename="二.txt", content="原文2")],
            extractor=FakeExtractor([
                {
                    "title": "第二团",
                    "duration_hours": 6,
                    "kp": {"name": "汐蒂仙改名", "qq": "1950566630"},
                    "players": [{"name": "黑猫约娜"}],
                    "confidence": 0.9,
                    "warnings": [],
                }
            ]),
            model_name="fake-model",
        )

        hosts = [row for row in self.service.list_players("2026-05") if row["qq"] == "1950566630"]
        self.assertEqual(len(hosts), 1)
        self.assertEqual(hosts[0]["game_count"], 2)
        self.assertEqual(hosts[0]["host_count"], 2)
        self.assertAlmostEqual(hosts[0]["game_hours"], 10.0)

    def test_kp_with_same_name_but_different_qq_is_not_merged(self):
        self.service.import_text_files(
            month="2026-05",
            files=[SessionTextFile(filename="甲.txt", content="甲")],
            extractor=FakeExtractor([
                {
                    "title": "甲",
                    "duration_hours": 1,
                    "kp": {"name": "同名主持", "qq": "111111"},
                    "players": [{"name": "玩家甲"}],
                    "confidence": 0.9,
                    "warnings": [],
                }
            ]),
            model_name="fake-model",
        )
        self.service.import_text_files(
            month="2026-05",
            files=[SessionTextFile(filename="乙.txt", content="乙")],
            extractor=FakeExtractor([
                {
                    "title": "乙",
                    "duration_hours": 1,
                    "kp": {"name": "同名主持", "qq": "222222"},
                    "players": [{"name": "玩家乙"}],
                    "confidence": 0.9,
                    "warnings": [],
                }
            ]),
            model_name="fake-model",
        )

        hosts = [row for row in self.service.list_players("2026-05") if row["name"] == "同名主持"]
        self.assertEqual({row["qq"] for row in hosts}, {"111111", "222222"})
        self.assertEqual(sum(row["host_count"] for row in hosts), 2)

    def test_player_with_same_qq_is_merged_even_when_name_changes(self):
        self.service.import_text_files(
            month="2026-05",
            files=[SessionTextFile(filename="一.txt", content="一")],
            extractor=FakeExtractor([
                {
                    "title": "第一团",
                    "duration_hours": 2,
                    "kp": {"name": "主持甲", "qq": "999001"},
                    "players": [{"name": "玩家旧名", "qq": "888001"}],
                    "confidence": 0.9,
                    "warnings": [],
                }
            ]),
            model_name="fake-model",
        )
        self.service.import_text_files(
            month="2026-05",
            files=[SessionTextFile(filename="二.txt", content="二")],
            extractor=FakeExtractor([
                {
                    "title": "第二团",
                    "duration_hours": 3,
                    "kp": {"name": "主持乙", "qq": "999002"},
                    "players": [{"name": "玩家新名", "qq": "888001"}],
                    "confidence": 0.9,
                    "warnings": [],
                }
            ]),
            model_name="fake-model",
        )

        players = [row for row in self.service.list_players("2026-05") if row["qq"] == "888001"]
        self.assertEqual(len(players), 1)
        self.assertEqual(players[0]["game_count"], 2)
        self.assertAlmostEqual(players[0]["game_hours"], 5.0)

    def test_import_batch_records_file_success_and_failure_counts(self):
        result = self.service.import_text_files(
            month="2026-05",
            files=[
                SessionTextFile(filename="成功.txt", content="成功"),
                SessionTextFile(filename="失败.txt", content="失败"),
            ],
            extractor=FakeExtractor([
                {
                    "title": "成功",
                    "duration_hours": 1,
                    "kp": {"name": "主持", "qq": None},
                    "players": [{"name": "玩家"}],
                    "confidence": 0.9,
                    "warnings": [],
                },
                {
                    "title": "失败",
                    "duration_hours": None,
                    "kp": {"name": "主持", "qq": None},
                    "players": [{"name": "玩家"}],
                    "confidence": 0.1,
                    "warnings": [],
                },
            ]),
            model_name="fake-model",
        )

        self.assertEqual(result["success_count"], 1)
        self.assertEqual(result["failure_count"], 1)
        batch = self.service.list_import_batches("2026-05")[0]
        self.assertEqual(batch["file_count"], 2)
        self.assertEqual(batch["success_count"], 1)
        self.assertEqual(batch["failure_count"], 1)

    def test_missing_openai_key_imports_as_failure_without_stats(self):
        old_key = os.environ.pop("OPENAI_API_KEY", None)
        old_third_party_key = os.environ.pop("THIRD_PARTY_API_KEY", None)
        try:
            result = self.service.import_text_files(
                month="2026-05",
                files=[SessionTextFile(filename="未配置.txt", content="原文")],
                extractor=None,
                model_name="gpt-test",
            )
        finally:
            if old_key is not None:
                os.environ["OPENAI_API_KEY"] = old_key
            if old_third_party_key is not None:
                os.environ["THIRD_PARTY_API_KEY"] = old_third_party_key

        self.assertEqual(result["success_count"], 0)
        self.assertEqual(result["failure_count"], 1)
        self.assertIn("OPENAI_API_KEY", result["items"][0]["reason"])
        self.assertEqual(self.service.get_overview("2026-05")["session_count"], 0)

    def test_delete_session_removes_participation_from_stats(self):
        result = self.service.import_text_files(
            month="2026-05",
            files=[SessionTextFile(filename="删除.txt", content="原文")],
            extractor=FakeExtractor([
                {
                    "title": "待删除",
                    "duration_hours": 3,
                    "kp": {"name": "主持", "qq": None},
                    "players": [{"name": "玩家"}],
                    "confidence": 0.9,
                    "warnings": [],
                }
            ]),
            model_name="fake-model",
        )
        session_id = result["items"][0]["session_id"]

        self.assertTrue(self.service.delete_session(session_id))
        self.assertEqual(self.service.get_overview("2026-05")["session_count"], 0)
        self.assertEqual(self.service.list_players("2026-05"), [])

    def test_cleanup_exact_duplicate_sessions_for_month_keeps_oldest_session(self):
        duplicate_payloads = [
            {
                "title": "duplicate round one",
                "duration_hours": 2,
                "kp": {"name": "host", "qq": None},
                "players": [{"name": "player a"}],
                "confidence": 0.9,
                "warnings": [],
            },
            {
                "title": "duplicate round two",
                "duration_hours": 2,
                "kp": {"name": "host", "qq": None},
                "players": [{"name": "player b"}],
                "confidence": 0.9,
                "warnings": [],
            },
            {
                "title": "unique round",
                "duration_hours": 4,
                "kp": {"name": "host", "qq": None},
                "players": [{"name": "player c"}],
                "confidence": 0.9,
                "warnings": [],
            },
        ]
        first = self.service.import_text_files(
            month="2026-05",
            files=[SessionTextFile(filename="A.txt", content="same source text")],
            extractor=FakeExtractor([duplicate_payloads[0]]),
            model_name="fake-model",
        )
        with self.service._connect() as conn:
            batch_id = self.service._create_batch(conn, "2026-05", "fake-model", 1)
            second_id = self.service._insert_session(
                conn,
                batch_id=batch_id,
                month="2026-05",
                file=SessionTextFile(filename="B.txt", content="same source text"),
                content_hash=self.service._content_hash("same source text"),
                payload=duplicate_payloads[1],
                model_name="fake-model",
            )
        second = {"items": [{"session_id": second_id}]}
        unique = self.service.import_text_files(
            month="2026-05",
            files=[SessionTextFile(filename="C.txt", content="different source text")],
            extractor=FakeExtractor([duplicate_payloads[2]]),
            model_name="fake-model",
        )
        other_month = self.service.import_text_files(
            month="2026-06",
            files=[SessionTextFile(filename="D.txt", content="same source text")],
            extractor=FakeExtractor([duplicate_payloads[0]]),
            model_name="fake-model",
        )

        result = self.service.cleanup_exact_duplicate_sessions_for_month("2026-05")

        duplicate_hash = self.service._content_hash("same source text")
        self.assertEqual(result["deleted_count"], 1)
        self.assertEqual(result["kept"], [
            {
                "content_hash": duplicate_hash,
                "session_id": first["items"][0]["session_id"],
                "source_filename": "A.txt",
                "duplicate_count": 1,
            }
        ])
        self.assertEqual(result["deleted"], [
            {
                "content_hash": duplicate_hash,
                "session_id": second["items"][0]["session_id"],
                "source_filename": "B.txt",
            }
        ])
        remaining_ids = {row["id"] for row in self.service.list_sessions("2026-05")}
        self.assertEqual(remaining_ids, {first["items"][0]["session_id"], unique["items"][0]["session_id"]})
        self.assertEqual(self.service.get_overview("2026-05")["session_count"], 2)
        self.assertEqual(self.service.get_overview("2026-06")["session_count"], 1)
        self.assertIn(other_month["items"][0]["session_id"], {row["id"] for row in self.service.list_sessions("2026-06")})

    def test_openai_extractor_uses_structured_schema_and_base_url(self):
        calls = []

        class FakeResponse:
            def __enter__(self):
                return self

            def __exit__(self, exc_type, exc, tb):
                return False

            def read(self):
                return json.dumps({
                    "output_text": json.dumps({
                        "title": "测试团",
                        "duration_hours": 1,
                        "kp": {"name": "主持", "qq": None},
                        "players": [{"name": "玩家", "qq": None}],
                        "confidence": 0.9,
                        "warnings": [],
                    }, ensure_ascii=False)
                }).encode("utf-8")

        def fake_urlopen(request, timeout):
            calls.append((request, timeout))
            return FakeResponse()

        extractor = OpenAIExtractor(
            "gpt-test",
            api_key="test-key",
            base_url="https://example.test/v1",
            timeout=7,
            urlopen=fake_urlopen,
        )

        result = extractor.extract(filename="团.txt", content="原文", month="2026-05")

        self.assertEqual(result["title"], "测试团")
        request, timeout = calls[0]
        self.assertEqual(request.full_url, "https://example.test/v1/responses")
        self.assertEqual(timeout, 7)
        payload = json.loads(request.data.decode("utf-8"))
        self.assertEqual(payload["text"]["format"]["type"], "json_schema")
        self.assertEqual(payload["text"]["format"]["schema"]["required"], ["title", "duration_hours", "kp", "players", "confidence", "warnings"])

    def test_openai_extractor_uses_third_party_key_fallback(self):
        calls = []
        old_openai_key = os.environ.pop("OPENAI_API_KEY", None)
        old_third_party_key = os.environ.get("THIRD_PARTY_API_KEY")
        os.environ["THIRD_PARTY_API_KEY"] = "third-party-test-key"

        class FakeResponse:
            def __enter__(self):
                return self

            def __exit__(self, exc_type, exc, tb):
                return False

            def read(self):
                return json.dumps({
                    "choices": [
                        {
                            "message": {
                                "content": json.dumps({
                                    "title": "fallback key round",
                                    "duration_hours": 2,
                                    "kp": {"name": "host", "qq": None},
                                    "players": [{"name": "player", "qq": None}],
                                    "confidence": 0.8,
                                    "warnings": [],
                                })
                            }
                        }
                    ]
                }).encode("utf-8")

        def fake_urlopen(request, timeout):
            calls.append(request)
            return FakeResponse()

        try:
            extractor = OpenAIExtractor(
                "deepseek-v4-pro",
                base_url="https://example.test/v1",
                urlopen=fake_urlopen,
            )
            result = extractor.extract(filename="round.txt", content="raw", month="2026-06")
        finally:
            if old_openai_key is not None:
                os.environ["OPENAI_API_KEY"] = old_openai_key
            else:
                os.environ.pop("OPENAI_API_KEY", None)
            if old_third_party_key is not None:
                os.environ["THIRD_PARTY_API_KEY"] = old_third_party_key
            else:
                os.environ.pop("THIRD_PARTY_API_KEY", None)

        self.assertEqual(result["title"], "fallback key round")
        self.assertEqual(calls[0].headers["Authorization"], "Bearer third-party-test-key")

    def test_deepseek_model_uses_chat_completions_json_schema(self):
        calls = []

        class FakeResponse:
            def __enter__(self):
                return self

            def __exit__(self, exc_type, exc, tb):
                return False

            def read(self):
                return json.dumps({
                    "choices": [
                        {
                            "message": {
                                "content": json.dumps({
                                    "title": "chat round",
                                    "duration_hours": 2,
                                    "kp": {"name": "host", "qq": None},
                                    "players": [{"name": "player", "qq": None}],
                                    "confidence": 0.8,
                                    "warnings": [],
                                })
                            }
                        }
                    ]
                }).encode("utf-8")

        def fake_urlopen(request, timeout):
            calls.append((request, timeout))
            return FakeResponse()

        extractor = OpenAIExtractor(
            "deepseek-v4-pro",
            api_key="test-key",
            base_url="https://example.test/v1",
            timeout=9,
            urlopen=fake_urlopen,
        )

        result = extractor.extract(filename="round.txt", content="raw", month="2026-06")

        self.assertEqual(result["title"], "chat round")
        request, timeout = calls[0]
        self.assertEqual(request.full_url, "https://example.test/v1/chat/completions")
        self.assertEqual(timeout, 9)
        payload = json.loads(request.data.decode("utf-8"))
        self.assertEqual(payload["messages"][0]["role"], "user")
        self.assertEqual(payload["response_format"]["type"], "json_schema")
        self.assertEqual(payload["response_format"]["json_schema"]["schema"]["required"], ["title", "duration_hours", "kp", "players", "confidence", "warnings"])

    def test_chat_completions_falls_back_to_json_object_when_schema_is_rejected(self):
        calls = []

        class FakeResponse:
            def __enter__(self):
                return self

            def __exit__(self, exc_type, exc, tb):
                return False

            def read(self):
                return json.dumps({
                    "choices": [
                        {
                            "message": {
                                "content": json.dumps({
                                    "title": "fallback round",
                                    "duration_hours": 3,
                                    "kp": {"name": "host", "qq": None},
                                    "players": [{"name": "player", "qq": None}],
                                    "confidence": 0.7,
                                    "warnings": [],
                                })
                            }
                        }
                    ]
                }).encode("utf-8")

        def fake_urlopen(request, timeout):
            calls.append(json.loads(request.data.decode("utf-8")))
            if len(calls) == 1:
                raise urllib.error.HTTPError(
                    request.full_url,
                    400,
                    "bad response_format",
                    hdrs=None,
                    fp=BytesIO(b'{"error":"unsupported json_schema"}'),
                )
            return FakeResponse()

        extractor = OpenAIExtractor(
            "deepseek-v4-pro",
            api_key="test-key",
            base_url="https://example.test/v1",
            urlopen=fake_urlopen,
        )

        result = extractor.extract(filename="round.txt", content="raw", month="2026-06")

        self.assertEqual(result["title"], "fallback round")
        self.assertEqual(calls[0]["response_format"]["type"], "json_schema")
        self.assertEqual(calls[1]["response_format"]["type"], "json_object")

    def test_openai_extractor_retries_urlerror_wrapped_read_timeout(self):
        calls = []
        old_retries = os.environ.get("OPENAI_REQUEST_RETRIES")
        os.environ["OPENAI_REQUEST_RETRIES"] = "1"

        def fake_urlopen(request, timeout):
            calls.append((request, timeout))
            raise urllib.error.URLError(TimeoutError("The read operation timed out"))

        try:
            extractor = OpenAIExtractor(
                "deepseek-v4-pro",
                api_key="test-key",
                base_url="https://example.test/v1",
                timeout=5,
                urlopen=fake_urlopen,
            )

            with self.assertRaises(RuntimeError) as raised:
                extractor.extract(filename="round.txt", content="raw", month="2026-06")
        finally:
            if old_retries is None:
                os.environ.pop("OPENAI_REQUEST_RETRIES", None)
            else:
                os.environ["OPENAI_REQUEST_RETRIES"] = old_retries

        self.assertEqual(len(calls), 2)
        self.assertIn("模型请求超时", str(raised.exception))
        self.assertNotIn("The read operation timed out", str(raised.exception))


if __name__ == "__main__":
    unittest.main()
