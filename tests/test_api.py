import pytest
from fastapi.testclient import TestClient
from api.main import app
from database import Database

client = TestClient(app)


class TestRosterAPI:
    @pytest.fixture(autouse=True)
    def setup_test_db(self, tmp_path):
        # Override test DB
        test_db_path = str(tmp_path / "test_api.db")
        from api import roster
        roster.db = Database(test_db_path)
        yield
        if hasattr(roster.db.conn, "close"):
            roster.db.conn.close()

    def test_health_check(self):
        res = client.get("/api/health")
        assert res.status_code == 200
        assert res.json()["status"] == "ok"

    def test_commit_and_get_daily_report(self):
        staged_cadets = [
            {
                "temp_id": "c1",
                "rank": "курсант",
                "fio": "Иванов Иван Иванович",
                "short_fio": "Иванов И.И.",
                "group_num": "101",
                "course": 1,
                "gender": "male",
                "status": "active",
                "needs_manual_review": False,
                "warnings": [],
                "sort_order": 0
            },
            {
                "temp_id": "c2",
                "rank": "курсант",
                "fio": "Смирнова Анна Сергеевна",
                "short_fio": "Смирнова А.С.",
                "group_num": "101",
                "course": 1,
                "gender": "female",
                "status": "active",
                "needs_manual_review": False,
                "warnings": [],
                "sort_order": 1
            }
        ]

        # 1. Commit
        res_commit = client.post("/api/roster/commit", json={"cadets": staged_cadets})
        assert res_commit.status_code == 200
        assert res_commit.json()["imported_count"] == 2

        # 2. Get daily report
        res_daily = client.get("/api/roster/daily?report_date=2026-11-04")
        assert res_daily.status_code == 200
        data = res_daily.json()
        assert len(data["cadets"]) == 2
        assert data["balance"]["is_sum_balanced"] is True
        assert data["balance"]["total"]["listed_count"] == 2
        assert data["balance"]["male"]["listed_count"] == 1
        assert data["balance"]["female"]["listed_count"] == 1

    def test_transform_staging(self):
        cadets = [
            {
                "temp_id": "1",
                "rank": "ряд.",
                "fio": "Иванов И.И.",
                "short_fio": "Иванов И.И.",
                "group_num": "101",
                "course": 1,
                "gender": "male",
                "status": "active",
                "needs_manual_review": False,
                "warnings": [],
                "sort_order": 0
            }
        ]

        # Increment all
        res = client.post("/api/roster/transform", json={
            "cadets": cadets,
            "action": "increment_all"
        })
        assert res.status_code == 200
        updated = res.json()["cadets"]
        assert updated[0]["course"] == 2

        # Bulk set course
        res_bulk = client.post("/api/roster/transform", json={
            "cadets": updated,
            "action": "bulk_set_course",
            "target_course": 4
        })
        assert res_bulk.status_code == 200
        assert res_bulk.json()["cadets"][0]["course"] == 4
