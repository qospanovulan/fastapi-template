from sqlalchemy import Integer, String, Column, MetaData, Table
from sqlalchemy.orm import registry

from app.application.models import Template

metadata_obj = MetaData()
mapper_registry = registry()

template = Table(
    "template",
    metadata_obj,
    Column("id", Integer, primary_key=True),
    Column("name", String()),
    Column("file_path", String()),
)


mapper_registry.map_imperatively(Template, template)
