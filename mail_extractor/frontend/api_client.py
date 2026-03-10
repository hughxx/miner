import httpx
from typing import Dict
import json

class APIClient:
    def __init__(self, base_url: str = "http://localhost:8000"):
        self.base_url = base_url

    async def extract_emails(self, json_data: dict) -> Dict:
        """提交邮件数据"""
        files = {
            'file': (
                'emails.json',
                json.dumps(json_data, ensure_ascii=False).encode('utf-8'),
                'application/json'
            )
        }

        async with httpx.AsyncClient(timeout=120.0) as client:
            resp = await client.post(f"{self.base_url}/api/extract", files=files)
            resp.raise_for_status()
            return resp.json()

    async def get_task_status(self, task_id: str) -> Dict:
        """查询任务状态"""
        async with httpx.AsyncClient(timeout=10.0) as client:
            resp = await client.get(f"{self.base_url}/api/task/{task_id}")
            resp.raise_for_status()
            return resp.json()
