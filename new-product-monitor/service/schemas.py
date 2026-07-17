from __future__ import annotations

from pydantic import BaseModel, Field


class GenerateWeeklyReportRequest(BaseModel):
    recordId: str = Field(..., min_length=1)
    secret: str | None = None


class GenerateWeeklyReportResponse(BaseModel):
    ok: bool
    status: str
    recordId: str
    message: str = ""
