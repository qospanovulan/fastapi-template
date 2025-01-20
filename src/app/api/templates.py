import logging
import os.path
from dataclasses import dataclass
from pathlib import Path
from typing import Annotated

from fastapi import HTTPException, status, APIRouter, Depends, UploadFile, File
from fastapi.responses import FileResponse

from app.application.protocols.database import DatabaseGateway, UoW
from app.application.templates import create_or_get_template, get_all_templates, get_template
from app.config.base import Settings

templates_router = APIRouter()


@dataclass
class TemplateInfo:
    id: int
    name: str
    file_path: str


@dataclass
class SomeResult:
    template_id: int

@dataclass
class ErrorResponse:
    detail: str
    status_code: int


@dataclass
class OkResponse:
    detail: str
    status_code: int


@templates_router.get('/', responses={
    404: {"model": ErrorResponse, "description": "Templates not found"}
})
def get_templates_list(
        database: Annotated[DatabaseGateway, Depends()],
        uow: Annotated[UoW, Depends()]
) -> list[TemplateInfo]:

    templates = get_all_templates(database, uow)

    if not templates:
        raise HTTPException(status_code=404, detail="Templates not found")

    return [TemplateInfo(id=t.id, name=t.name, file_path=t.file_path) for t in templates]


@templates_router.get('/{template_id}/', responses={
    404: {"model": ErrorResponse, "description": "Template not found"}

})
def get_template_by_id(
        database: Annotated[DatabaseGateway, Depends()],
        uow: Annotated[UoW, Depends()],
        template_id: int = None,
) -> FileResponse:

    template = get_template(database, uow, template_id)

    if not template:
        raise HTTPException(status_code=404, detail="Template not found")

    path = Path(template.file_path)
    if not path.is_file():
        raise HTTPException(status_code=404, detail="File not found")

    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=template.name,
    )



@templates_router.post("/create/", response_model=SomeResult, responses={
    409: {"model": ErrorResponse, "description": "Template already exists."}
})
def add_template(
        database: Annotated[DatabaseGateway, Depends()],
        uow: Annotated[UoW, Depends()],
        settings: Annotated[Settings, Depends()],
        name: str = None,
        file: UploadFile = File(...),
) -> SomeResult:

    if not name:
        name = file.filename
    else:
        file_extension = Path(file.filename).suffix

        if not name.endswith(('.xlsx', '.pdf')):
            name = f"{name}{file_extension}"


    file_path = os.path.join(settings.template_path, name)

    template_id, created = create_or_get_template(database, uow, name, file_path)

    if created:
        with Path(file_path).open("wb") as f:
            f.write(file.file.read())
    else:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=f'Template with name "{name}" already exists'
        )

    return SomeResult(
        template_id=template_id,
    )


@templates_router.post("/update/")
def update_template(
        database: Annotated[DatabaseGateway, Depends()],
        uow: Annotated[UoW, Depends()],
        settings: Annotated[Settings, Depends()],
        name: str = None,
        file: UploadFile = File(...),
) -> SomeResult:
    if not name:
        name = file.filename
    else:
        file_extension = Path(file.filename).suffix

        if not name.endswith(('.xlsx', '.pdf')):
            name = f"{name}{file_extension}"

    file_path = os.path.join(settings.template_path, name)

    template_id, created = create_or_get_template(database, uow, name, file_path)

    with Path(file_path).open("wb") as f:
        f.write(file.file.read())

    return SomeResult(
        template_id=template_id,
    )

@templates_router.post("/delete/", responses={
    404: {"model": ErrorResponse, "description": "Template not found."}
})
def delete_template(
        database: Annotated[DatabaseGateway, Depends()],
        uow: Annotated[UoW, Depends()],
        settings: Annotated[Settings, Depends()],
        name: str
) -> OkResponse:


    file_path = os.path.join(settings.template_path, name)

    deleted = delete_template(database, uow, name)

    if not deleted:
        raise HTTPException(
            detail="Template not found",
            status_code=status.HTTP_404_NOT_FOUND
        )
    try:
        os.remove(file_path)
    except Exception as e:
        logging.error(f"Error while deleting file {file_path}: {e}")

    return OkResponse(
        detail="Template deleted!",
        status_code=status.HTTP_200_OK
    )



