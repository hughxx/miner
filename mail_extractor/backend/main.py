from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from router import extract

app = FastAPI(title="邮件提取服务")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(extract.router, prefix="/api", tags=["提取"])

if __name__ == "__main__":
    import uvicorn
    from config import settings
    uvicorn.run(app, host=settings.host, port=settings.port)
