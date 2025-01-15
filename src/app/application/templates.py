from .models import Template
from .protocols.database import DatabaseGateway, UoW


def create_or_get_template(
        database: DatabaseGateway,
        uow: UoW,
        name: str,
        file_path: str,
) -> (int, bool):

    template = database.get_template_by_name(name)

    if template:

        return template.id, False

    template = Template(name=name, file_path=file_path)
    database.add_template(template)
    uow.commit()
    return template.id, True


def get_all_templates(database: DatabaseGateway, uow: UoW):
    templates = database.get_templates()

    return templates

def get_template(database: DatabaseGateway, uow: UoW, template_id: int):
    template = database.get_template_by_id(template_id=template_id)

    return template
