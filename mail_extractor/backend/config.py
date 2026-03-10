from pydantic_settings import BaseSettings
from typing import Optional
import os

class Settings(BaseSettings):
    # 服务配置
    host: str = "0.0.0.0"
    port: int = 8000

    # 外部服务URL（需根据实际配置）
    img_to_url_api: str = "http://localhost:8001/Img_to_url"
    ocr_api: str = "http://localhost:8002/ocr"
    llm_api: str = "http://localhost:8003/llm"
    upload_to_db_api: str = "http://localhost:8004/upload_to_db"

    # 敏感词文件路径
    sensitive_words_file: str = "sensitive_words.txt"

    # 线程池配置
    max_workers: int = 4

    class Config:
        env_file = ".env"

settings = Settings()
