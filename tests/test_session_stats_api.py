from __future__ import annotations

import os
import sys
import tempfile
import time
import unittest
from pathlib import Path

from fastapi.testclient import TestClient

BACKEND_ROOT = Path(__file__).resolve().parents[1] / "web" / "backend"
sys.path.insert(0, str(BACKEND_ROOT))

from app.session_stats import SessionStatsService, SessionTextFile
import app.main as main_app


class FakeExtractor:
    def __init__(self, payloads):
        self.payloads = list(payloads)

    def extract(self, *, filename: str, content: str, month: str):
        return self.payloads.pop(0)


class SessionStatsApiTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.db_path = Path(self.tmp.name) / "stats.sqlite"
        self.old_service = getattr(main_app, "SESSION_STATS_SERVICE", None)
        self.old_extractor = getattr(main_app, "SESSION_STATS_EXTRACTOR", None)
        main_app.SESSION_STATS_SERVICE = SessionStatsService(self.db_path)
        main_app.SESSION_STATS_EXTRACTOR = None
        self.client = TestClient(main_app.app)

    def tearDown(self) -> None:
        if self.old_service is None:
            if hasattr(main_app, "SESSION_STATS_SERVICE"):
                delattr(main_app, "SESSION_STATS_SERVICE")
        else:
            main_app.SESSION_STATS_SERVICE = self.old_service
        main_app.SESSION_STATS_EXTRACTOR = self.old_extractor
        self.tmp.cleanup()

    def test_players_endpoint_is_public_and_returns_empty_stats(self):
        response = self.client.get("/api/session-stats/players?month=2026-05")

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json()["items"], [])

    def test_public_stats_endpoints_return_imported_session_data(self):
        main_app.SESSION_STATS_SERVICE.import_text_files(
            month="2026-05",
            files=[SessionTextFile(filename="团.txt", content="原文")],
            extractor=FakeExtractor([
                {
                    "title": "测试团",
                    "duration_hours": 2,
                    "kp": {"name": "主持", "qq": "123456"},
                    "players": [{"name": "玩家", "qq": None}],
                    "confidence": 0.9,
                    "warnings": [],
                }
            ]),
            model_name="fake-model",
        )

        overview = self.client.get("/api/session-stats/overview?month=2026-05")
        players = self.client.get("/api/session-stats/players?month=2026-05&sort=hours")
        sessions = self.client.get("/api/session-stats/sessions?month=2026-05")

        self.assertEqual(overview.status_code, 200)
        self.assertEqual(overview.json()["session_count"], 1)
        self.assertEqual(overview.json()["participant_count"], 2)
        self.assertEqual(players.status_code, 200)
        self.assertEqual(players.json()["count"], 2)
        self.assertEqual(sessions.status_code, 200)
        self.assertEqual(sessions.json()["items"][0]["title"], "测试团")

    def test_import_requires_admin_configuration(self):
        old_password = os.environ.pop("ADMIN_PASSWORD", None)
        try:
            response = self.client.post(
                "/api/session-stats/import",
                data={"month": "2026-05"},
                files=[("files", ("结团.txt", "原文", "text/plain"))],
            )
        finally:
            if old_password is not None:
                os.environ["ADMIN_PASSWORD"] = old_password

        self.assertEqual(response.status_code, 503)
        self.assertIn("ADMIN_PASSWORD", response.json()["detail"])

    def test_errors_endpoint_requires_admin_configuration(self):
        old_password = os.environ.pop("ADMIN_PASSWORD", None)
        try:
            response = self.client.get("/api/session-stats/errors?month=2026-05")
        finally:
            if old_password is not None:
                os.environ["ADMIN_PASSWORD"] = old_password

        self.assertEqual(response.status_code, 503)

    def test_import_job_reports_progress_and_final_counts(self):
        old_password = os.environ.get("ADMIN_PASSWORD")
        os.environ["ADMIN_PASSWORD"] = "test-pass"
        main_app.SESSION_STATS_EXTRACTOR = FakeExtractor([
            {
                "title": "第一团",
                "duration_hours": 2,
                "kp": {"name": "主持一", "qq": "111"},
                "players": [{"name": "玩家一", "qq": None}],
                "confidence": 0.9,
                "warnings": [],
            },
            {
                "title": "第二团",
                "duration_hours": 3,
                "kp": {"name": "主持二", "qq": "222"},
                "players": [{"name": "玩家二", "qq": None}],
                "confidence": 0.9,
                "warnings": [],
            },
        ])
        try:
            login = self.client.post("/api/admin/login", json={"password": "test-pass"})
            self.assertEqual(login.status_code, 200)

            started = self.client.post(
                "/api/session-stats/import-jobs",
                data={"month": "2026-05"},
                files=[
                    ("files", ("一.txt", "原文一", "text/plain")),
                    ("files", ("二.txt", "原文二", "text/plain")),
                ],
            )
            self.assertEqual(started.status_code, 200)
            payload = started.json()
            self.assertEqual(payload["status"], "running")
            self.assertEqual(payload["total_count"], 2)
            job_id = payload["job_id"]

            final = None
            for _ in range(30):
                polled = self.client.get(f"/api/session-stats/import-jobs/{job_id}")
                self.assertEqual(polled.status_code, 200)
                final = polled.json()
                self.assertLessEqual(final["processed_count"], 2)
                if final["status"] == "completed":
                    break
                time.sleep(0.05)

            self.assertIsNotNone(final)
            self.assertEqual(final["status"], "completed")
            self.assertEqual(final["processed_count"], 2)
            self.assertEqual(final["success_count"], 2)
            self.assertEqual(final["failure_count"], 0)
            self.assertEqual(len(final["items"]), 2)
            self.assertEqual(main_app.SESSION_STATS_SERVICE.get_overview("2026-05")["session_count"], 2)
        finally:
            if old_password is None:
                os.environ.pop("ADMIN_PASSWORD", None)
            else:
                os.environ["ADMIN_PASSWORD"] = old_password


if __name__ == "__main__":
    unittest.main()
