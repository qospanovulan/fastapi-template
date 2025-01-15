from dataclasses import dataclass, field


@dataclass
class Template:
    id: int = field(init=False)
    name: str
    file_path: str
