from pydantic import BaseModel
from typing import Optional, List
from enum import Enum
from datetime import datetime

class TaskStatus(str, Enum):
    PENDING = "pending"
    PROCESSING = "processing"
    COMPLETED = "completed"
    FAILED = "failed"

class EmailImage(BaseModel):
    position: int
    base64: str

class EmailData(BaseModel):
    subject: str
    sender: str
    received_time: str
    conversation_topic: str
    html_content: str
    images: List[EmailImage] = []

class ExtractRequest(BaseModel):
    emails: List[EmailData]
    options: dict = {"desensitize": True}

class Task(BaseModel):
    task_id: str
    email_index: int
    status: TaskStatus = TaskStatus.PENDING
    progress: int = 0
    message: str = "等待处理"
    result: Optional[dict] = None
    error: Optional[str] = None
    created_at: datetime = None
    updated_at: datetime = None

    class Config:
        arbitrary_types_allowed = True
