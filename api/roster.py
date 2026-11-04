from fastapi import APIRouter, UploadFile, File, HTTPException, Query
from fastapi.responses import Response
from typing import List, Optional, Dict, Any
from pydantic import BaseModel
import tempfile
import os

from core.models import StagingCadetDTO, TriSliceBalance
from core.parsers import (
    parse_roster_excel,
    increment_course_all,
    bulk_set_course,
    remove_graduates,
    toggle_gender
)
from core.balance import compute_tri_slice_balance
from core.exporter import export_report_to_excel
from database import Database

router = APIRouter(prefix="/api/roster", tags=["Roster"])
db = Database()


class ParsePreviewResponse(BaseModel):
    cadets: List[StagingCadetDTO]
    total_count: int
    graduated_count: int
    needs_review_count: int
    males_count: int
    females_count: int


class CommitRosterRequest(BaseModel):
    cadets: List[StagingCadetDTO]


class CommitRosterResponse(BaseModel):
    status: str = "success"
    imported_count: int


class TransformRequest(BaseModel):
    cadets: List[StagingCadetDTO]
    action: str  # "increment_all", "bulk_set_course", "remove_graduates", "toggle_gender"
    target_course: Optional[int] = 1
    selected_temp_ids: Optional[List[str]] = None
    target_temp_id: Optional[str] = None


class AddAbsenceRequest(BaseModel):
    report_date: str
    cadet_id: int
    type_id: int
    note: str = ""


class DailyReportResponse(BaseModel):
    report_date: str
    cadets: List[Dict[str, Any]]
    absences: List[Dict[str, Any]]
    categories: List[str]
    balance: TriSliceBalance


@router.post("/parse-preview", response_model=ParsePreviewResponse)
async def parse_preview(file: UploadFile = File(...)):
    """Receives .xlsx file, parses cadet rows into staging DTOs, detects gender & course,
    and returns intermediate staging state without committing to DB.
    """
    if not file.filename.endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="Поддерживаются только файлы формата .xlsx или .xlsm")

    contents = await file.read()
    try:
        staging_cadets = parse_roster_excel(contents)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Ошибка парсинга Excel: {str(e)}")

    graduated = sum(1 for c in staging_cadets if c.status == "graduated" or c.course > 5)
    needs_rev = sum(1 for c in staging_cadets if c.needs_manual_review)
    males = sum(1 for c in staging_cadets if c.gender == "male")
    females = sum(1 for c in staging_cadets if c.gender == "female")

    return ParsePreviewResponse(
        cadets=staging_cadets,
        total_count=len(staging_cadets),
        graduated_count=graduated,
        needs_review_count=needs_rev,
        males_count=males,
        females_count=females
    )


@router.post("/commit", response_model=CommitRosterResponse)
def commit_roster(payload: CommitRosterRequest):
    """Commits corrected staging cadets into active database."""
    if not payload.cadets:
        raise HTTPException(status_code=400, detail="Список курсантов пуст")

    db.clear_and_import_cadets(payload.cadets)
    return CommitRosterResponse(imported_count=len(payload.cadets))


@router.post("/transform", response_model=ParsePreviewResponse)
def transform_staging(payload: TransformRequest):
    """Applies bulk transformation (e.g. +1 course, bulk course override, remove graduates)
    to staging items in memory and returns updated preview.
    """
    cadets = payload.cadets
    act = payload.action

    if act == "increment_all":
        cadets = increment_course_all(cadets)
    elif act == "bulk_set_course":
        course = payload.target_course or 1
        cadets = bulk_set_course(cadets, course, payload.selected_temp_ids)
    elif act == "remove_graduates":
        cadets = remove_graduates(cadets)
    elif act == "toggle_gender":
        if payload.target_temp_id:
            cadets = [
                toggle_gender(c) if c.temp_id == payload.target_temp_id else c
                for c in cadets
            ]
        elif payload.selected_temp_ids:
            sel_set = set(payload.selected_temp_ids)
            cadets = [
                toggle_gender(c) if c.temp_id in sel_set else c
                for c in cadets
            ]
    else:
        raise HTTPException(status_code=400, detail=f"Неизвестное действие: {act}")

    graduated = sum(1 for c in cadets if c.status == "graduated" or c.course > 5)
    needs_rev = sum(1 for c in cadets if c.needs_manual_review)
    males = sum(1 for c in cadets if c.gender == "male")
    females = sum(1 for c in cadets if c.gender == "female")

    return ParsePreviewResponse(
        cadets=cadets,
        total_count=len(cadets),
        graduated_count=graduated,
        needs_review_count=needs_rev,
        males_count=males,
        females_count=females
    )


@router.get("/daily", response_model=DailyReportResponse)
def get_daily_report(report_date: Optional[str] = None):
    """Returns daily attendance, cadets list, absence categories, and 3-slice balance."""
    from datetime import date
    target_date = report_date or date.today().strftime("%Y-%m-%d")

    cadets = db.get_all_cadets()
    absences = db.get_absences_for_date(target_date)
    categories = db.get_distinct_categories()

    balance = compute_tri_slice_balance(cadets, absences, categories)

    return DailyReportResponse(
        report_date=target_date,
        cadets=cadets,
        absences=absences,
        categories=categories,
        balance=balance
    )


@router.post("/absence")
def add_absence(req: AddAbsenceRequest):
    db.set_absence(req.report_date, req.cadet_id, req.type_id, req.note)
    return {"status": "success"}


@router.delete("/absence/{absence_id}")
def delete_absence(absence_id: int):
    db.remove_absence(absence_id)
    return {"status": "success"}


@router.get("/export-excel")
def download_excel_report(report_date: Optional[str] = None):
    """Generates the multi-sheet .xlsx report with active formulas."""
    from datetime import date
    target_date = report_date or date.today().strftime("%Y-%m-%d")

    cadets = db.get_all_cadets()
    absences = db.get_absences_for_date(target_date)
    categories = db.get_distinct_categories()

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = tmp.name

    try:
        export_report_to_excel(
            output_path=tmp_path,
            report_date=target_date,
            cadets=cadets,
            absences=absences,
            unit_title=db.get_setting("unit_title", "3 курс 1 факультета"),
            chief_title=db.get_setting("chief_title", "Начальник 3 курса 1 факультета"),
            chief_name=db.get_setting("chief_name", "майор                          В.Медков"),
            active_categories=categories
        )
        with open(tmp_path, "rb") as f:
            data = f.read()
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

    filename = f"Строевая_записка_{target_date}.xlsx"
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )
