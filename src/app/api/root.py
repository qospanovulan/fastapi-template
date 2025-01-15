from fastapi import APIRouter

from .index import index_router
from .templates import templates_router

root_router = APIRouter()

root_router.include_router(
    templates_router,
    prefix="/templates",
)
root_router.include_router(
    index_router,
)
