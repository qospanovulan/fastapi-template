from abc import ABC, abstractmethod

from app.application.models import Template


class UoW(ABC):
    @abstractmethod
    def commit(self):
        raise NotImplementedError

    @abstractmethod
    def flush(self):
        raise NotImplementedError


class DatabaseGateway(ABC):

    @abstractmethod
    def get_template_by_name(self, name: str):
        raise NotImplementedError

    @abstractmethod
    def get_template_by_id(self, template_id: int):
        raise NotImplementedError

    @abstractmethod
    def get_templates(self):
        raise NotImplementedError

    @abstractmethod
    def add_template(self, template: Template) -> None:
        raise NotImplementedError

    @abstractmethod
    def delete_template_by_id(self, template_id: int):
        raise NotImplementedError


