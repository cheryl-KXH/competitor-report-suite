from __future__ import annotations

from pydantic import BaseModel


class AcceptedResponse(BaseModel):
    ok: bool
    status: str
    recordId: str
    message: str = ""


class FinalizeResponse(BaseModel):
    ok: bool
    status: str
    recordId: str
    message: str
    reportUrl: str
    folderUrl: str
    archivedFileCount: int
    archivedFileNames: str


class HealthResponse(BaseModel):
    ok: bool
