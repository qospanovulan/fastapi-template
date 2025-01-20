from sqlalchemy.orm import Session

from app.application.models import Template
from app.application.protocols.database import DatabaseGateway


class SqlaGateway(DatabaseGateway):
    def __init__(self, session: Session):
        self.session = session

    def get_template_by_name(self, name: str):
        return self.session.query(Template).filter_by(name=name).first()

    def get_template_by_id(self, template_id: int):
        return self.session.query(Template).filter_by(id=template_id).first()

    def get_templates(self):
        return self.session.query(Template).all()

    def add_template(self, template: Template) -> None:
        self.session.add(template)
        return

    def delete_template_by_id(self, template_id: int):
        template = self.session.query(Template).filter_by(id=template_id).first()
        self.session.delete(template)
        self.session.commit()
        return
