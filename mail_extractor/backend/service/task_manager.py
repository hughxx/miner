import uuid
from datetime import datetime
from typing import Dict, Optional
from models import Task, TaskStatus, EmailData
from concurrent.futures import ThreadPoolExecutor

class TaskManager:
    def __init__(self, max_workers: int = 4):
        self.tasks: Dict[str, Task] = {}
        self.executor = ThreadPoolExecutor(max_workers=max_workers)

    def create_task(self, email_index: int) -> str:
        task_id = f"task_{uuid.uuid4().hex[:8]}"
        task = Task(
            task_id=task_id,
            email_index=email_index,
            created_at=datetime.now(),
            updated_at=datetime.now()
        )
        self.tasks[task_id] = task
        return task_id

    def get_task(self, task_id: str) -> Optional[Task]:
        return self.tasks.get(task_id)

    def update_task(self, task_id: str, **kwargs):
        if task_id in self.tasks:
            task = self.tasks[task_id]
            for key, value in kwargs.items():
                setattr(task, key, value)
            task.updated_at = datetime.now()

    def submit_task(self, task_id: str, email_data: EmailData, process_email):
        """提交任务到线程池"""
        future = self.executor.submit(process_email, task_id, email_data)
        return future

task_manager = TaskManager()
