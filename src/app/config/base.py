import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent.parent.parent.parent

class Settings:
    TEMPLATE_FOLDER: str = os.environ.get('TEMPLATE_FOLDER')

    @property
    def template_path(self):
        return os.path.join(BASE_DIR, self.TEMPLATE_FOLDER)
