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

from app.session_stats import (
    OpenAIExtractor,
    SessionStatsService,
    SessionTextFile,
    clean_person_identity,
    infer_duration_hours_fallback,
)


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

    def test_flaw_repairs_duration_fallback_kp_card_split_and_null_players(self):
        # 模拟生产库真实瑕疵：LLM 未抽出时长、KP 名字混带『+卡名（卡名）』、玩家含全 null 乱码项且混带括号 QQ
        raw_content = "十镇冬日11月第一车·21H结团\n时长：11.10~11.15 总计21H\n【KP】林雨霁（3909862519）+林若真（卡名）\n"
        result = self.service.import_text_files(
            month="2026-11",
            files=[SessionTextFile(filename="十镇冬日11月第一车·21H结团.txt", content=raw_content, source_file_id="/file_nov_1")],
            extractor=FakeExtractor([
                {
                    "title": "十镇冬日",
                    "duration_hours": None,
                    "kp": {"name": "林雨霁（3909862519）+林若真（卡名）", "qq": None},
                    "players": [
                        {"name": "梅普露（2577535590）", "qq": None},
                        {"name": "卡名：花塚硝子", "qq": None},
                        {"name": None, "qq": None},
                        {"name": "林雨霁", "qq": "3909862519"},  # 误混入 PL 的 KP 本人
                    ],
                    "confidence": 0.85,
                    "warnings": [],
                }
            ]),
            model_name="fake-model",
        )
        self.assertEqual(result["success_count"], 1, result)
        sid = result["items"][0]["session_id"]
        detail = self.service.get_session_detail(sid)
        self.assertAlmostEqual(detail["duration_hours"], 21.0)
        self.assertEqual(detail["kp"]["name"], "林雨霁")
        self.assertEqual(detail["kp"]["qq"], "3909862519")
        pls = [p for p in detail["participants"] if not p["is_host"]]
        self.assertEqual(len(pls), 2)
        self.assertEqual({(p["name"], p["qq"]) for p in pls}, {("梅普露", "2577535590"), ("花塚硝子", None)})

        # 同一群文件 source_file_id 在群内修订了正文（例如从 21h 改成 22h），重新同步时应原地覆盖而非记成两场团
        updated = self.service.import_text_files(
            month="2026-11",
            files=[SessionTextFile(filename="十镇冬日修订版.txt", content=raw_content + "\n修订", source_file_id="/file_nov_1")],
            extractor=FakeExtractor([
                {
                    "title": "十镇冬日（修订）",
                    "duration_hours": 22,
                    "kp": {"name": "林雨霁", "qq": "3909862519"},
                    "players": [{"name": "梅普露", "qq": "2577535590"}, {"name": "花塚硝子", "qq": None}],
                    "confidence": 0.95,
                    "warnings": [],
                }
            ]),
            model_name="fake-model",
        )
        self.assertEqual(updated["success_count"], 1)
        self.assertEqual(self.service.get_overview("2026-11")["session_count"], 1)
        self.assertAlmostEqual(self.service.get_overview("2026-11")["total_host_hours"], 22.0)

        # 支持手动编辑补录 KP 与 PL 名单并实时更新统计
        patched = self.service.update_session(
            updated["items"][0]["session_id"],
            kp={"name": "林雨霁", "qq": "3909862519"},
            players=[
                {"name": "梅普露", "qq": "2577535590"},
                {"name": "花塚硝子", "qq": "123450000"},
                {"name": "新补录调查员", "qq": None},
            ],
        )
        self.assertEqual(len([p for p in patched["participants"] if not p["is_host"]]), 3)
        self.assertEqual(self.service.get_overview("2026-11")["participant_count"], 4)

    def test_identity_and_duration_helpers(self):
        self.assertEqual(clean_person_identity("艾丝特1546003281", None), ("艾丝特", "1546003281"))
        self.assertEqual(clean_person_identity("Jinbert（1462158343）+方汪汪（卡名）", None), ("Jinbert", "1462158343"))
        # 纯数字卡名与 +1 增益后缀不应误伤
        self.assertEqual(clean_person_identity("10086", None), ("10086", None))
        self.assertEqual(clean_person_identity("苏早琴+1", None), ("苏早琴+1", None))
        # .ra3.5h 检定不应误抽成 5h
        self.assertIsNone(infer_duration_hours_fallback("普通结团报告.txt", "检定记录 .ra3.5h 成功"))
        self.assertEqual(infer_duration_hours_fallback("从地狱归来的刮刮乐8h结团.txt", "无时长正文"), 8.0)
        self.assertEqual(infer_duration_hours_fallback("隔离区h(31小时)第一车升b.txt", "无"), 31.0)

        # 验证原本无 QQ 的玩家在后续补填 QQ 时原地升级归并，不分裂成两条同名玩家
        self.service.import_text_files(
            month="2026-11",
            files=[SessionTextFile(filename="t1.txt", content="c1")],
            extractor=FakeExtractor([
                {"title": "T1", "duration_hours": 5, "kp": {"name": "KP甲", "qq": "100001"}, "players": [{"name": "无号旅人", "qq": None}], "confidence": 0.9, "warnings": []}
            ]),
            model_name="fake-model",
        )
        self.service.import_text_files(
            month="2026-11",
            files=[SessionTextFile(filename="t2.txt", content="c2")],
            extractor=FakeExtractor([
                {"title": "T2", "duration_hours": 7, "kp": {"name": "KP甲", "qq": "100001"}, "players": [{"name": "无号旅人", "qq": "77788899"}], "confidence": 0.9, "warnings": []}
            ]),
            model_name="fake-model",
        )
        travellers = [p for p in self.service.list_players("2026-11") if p["name"] == "无号旅人"]
        self.assertEqual(len(travellers), 1)
        self.assertEqual(travellers[0]["qq"], "77788899")
        self.assertEqual(travellers[0]["game_count"], 2)
        self.assertAlmostEqual(travellers[0]["game_hours"], 12.0)


if __name__ == "__main__":
    unittest.main()
