# 邮件提取工具 - 实现规格说明书

## 1. 项目初始化

### 1.1 环境依赖

**后端（Python 3.10+）**：
```txt
fastapi>=0.104.0
uvicorn>=0.24.0
python-multipart>=0.0.6
lxml>=4.9.0
beautifulsoup4>=4.12.0
httpx>=0.25.0
pydantic>=2.0.0
```

**前端（Python 3.10+）**：
```txt
pywin32>=306
```

---

## 2. 后端实现

### 2.1 入口文件

**backend/main.py**
```python
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
    uvicorn.run(app, host="0.0.0.0", port=8000)
```

---

### 2.2 配置管理

**backend/config.py**
```python
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
```

---

### 2.3 数据模型

**backend/models.py**
```python
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
```

---

### 2.4 任务管理

**backend/service/task_manager.py**
```python
import uuid
from datetime import datetime
from typing import Dict, Optional
from models import Task, TaskStatus, EmailData
from concurrent.futures import ThreadPoolExecutor
import asyncio

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
```

---

### 2.5 脱敏处理

**backend/service/desensitize.py**
```python
import re
from typing import List
import os

class Desensitizer:
    def __init__(self, words_file: str):
        self.sensitive_words: List[str] = []
        self._load_words(words_file)

    def _load_words(self, words_file: str):
        if os.path.exists(words_file):
            with open(words_file, 'r', encoding='utf-8') as f:
                self.sensitive_words = [line.strip() for line in f if line.strip()]

    def desensitize(self, text: str) -> str:
        """替换敏感词为 ***"""
        result = text
        for word in self.sensitive_words:
            if word:  # 避免空字符串
                result = result.replace(word, "***")
        return result

    def desensitize_html(self, html: str) -> str:
        """HTML内容脱敏"""
        # 简单处理：直接对整个HTML字符串替换
        return self.desensitize(html)

# 全局实例
desensitizer = Desensitizer("sensitive_words.txt")
```

---

### 2.6 HTML解析与图片处理

**backend/service/html_parser.py**
```python
from bs4 import BeautifulSoup
import base64
import httpx
from typing import List, Dict, Tuple
from models import EmailImage

class HTMLParser:
    def __init__(self, img_to_url_api: str, ocr_api: str):
        self.img_to_url_api = img_to_url_api
        self.ocr_api = ocr_api

    async def process_html(self, html: str, images: List[EmailImage]) -> Tuple[str, str]:
        """
        处理HTML和图片
        返回: (处理后的HTML, 所有OCR结果拼接)
        """
        soup = BeautifulSoup(html, 'lxml')

        # 建立图片位置映射
        img_map = {img.position: img.base64 for img in images}

        # 处理所有img标签
        ocr_results = []
        for idx, img_tag in enumerate(soup.find_all('img')):
            position = idx
            if position in img_map:
                base64_data = img_map[position]

                # 1. 图片转超链接
                url = await self._upload_image(base64_data)
                if url:
                    # 替换图片为超链接文本
                    link_text = f"[图片{idx+1}: {url}]"
                    img_tag.replace_with(link_text)

                # 2. OCR识别
                ocr_text = await self._ocr_image(base64_data)
                if ocr_text:
                    ocr_results.append(f"[图片{idx+1} OCR]: {ocr_text}")

        # 将OCR结果追加到HTML末尾
        if ocr_results:
            ocr_div = soup.new_tag('div')
            ocr_div['style'] = 'margin-top: 20px; padding: 10px; background: #f5f5f5;'
            ocr_div.string = "\n".join(ocr_results)
            soup.append(ocr_div)

        return str(soup), "\n".join(ocr_results)

    async def _upload_image(self, base64_data: str) -> str:
        """调用Img_to_url接口"""
        try:
            # 解析base64获取文件名和内容
            header, data = base64_data.split(',', 1) if ',' in base64_data else ('', base64_data)
            # 提取mime类型
            mime = 'image/png'
            if 'jpeg' in header or 'jpg' in header:
                mime = 'image/jpeg'

            # 这里需要根据实际API格式调整
            files = {'file': ('image.png', data, mime)}
            async with httpx.AsyncClient() as client:
                resp = await client.post(self.img_to_url_api, files=files)
                if resp.status_code == 200:
                    return resp.json().get('url', '')
        except Exception as e:
            print(f"图片上传失败: {e}")
        return ''

    async def _ocr_image(self, base64_data: str) -> str:
        """调用OCR接口"""
        try:
            # 提取纯base64数据
            if ',' in base64_data:
                base64_str = base64_data.split(',')[1]
            else:
                base64_str = base64_data

            async with httpx.AsyncClient() as client:
                resp = await client.post(self.ocr_api, json={'base64_str': base64_str})
                if resp.status_code == 200:
                    return resp.json().get('ocr_result', '')
        except Exception as e:
            print(f"OCR识别失败: {e}")
        return ''
```

---

### 2.7 LLM提取

**backend/service/llm_extractor.py**
```python
import httpx
from typing import Optional

class LLMExtractor:
    def __init__(self, llm_api: str):
        self.llm_api = llm_api

    async def extract(self, content: str, html_content: str) -> dict:
        """
        调用LLM提取字段
        返回: {name, description, reason, solution, process}
        """
        prompt = self._build_prompt(content, html_content)

        try:
            async with httpx.AsyncClient() as client:
                resp = await client.post(self.llm_api, json={'prompt': prompt})
                if resp.status_code == 200:
                    result = resp.json().get('resp', '')
                    return self._parse_llm_result(result)
        except Exception as e:
            print(f"LLM调用失败: {e}")

        return self._empty_result()

    def _build_prompt(self, content: str, html_content: str) -> str:
        return f"""请从以下邮件内容中提取关键信息，JSON格式输出：
{{
    "name": "邮件主题",
    "description": "描述",
    "reason": "原因",
    "solution": "解决方案",
    "process": "处理过程"
}}

邮件内容：
{content}
{html_content}"""

    def _parse_llm_result(self, response: str) -> dict:
        """解析LLM返回的JSON"""
        import json
        import re

        # 尝试提取JSON
        try:
            # 去掉可能的markdown代码块
            json_str = re.sub(r'^```json', '', response)
            json_str = re.sub(r'^```', '', json_str)
            json_str = json_str.strip()
            return json.loads(json_str)
        except:
            pass

        return self._empty_result()

    def _empty_result(self) -> dict:
        return {
            "name": "",
            "description": "",
            "reason": "",
            "solution": "",
            "process": ""
        }
```

---

### 2.8 提取路由

**backend/router/extract.py**
```python
from fastapi import APIRouter, UploadFile, File, HTTPException
from typing import List
import json
from datetime import datetime

from models import ExtractRequest, TaskStatus
from service.task_manager import task_manager
from service.html_parser import HTMLParser
from service.desensitizer import desensitizer
from service.llm_extractor import LLMExtractor
from config import settings

router = APIRouter()

# 初始化服务实例
html_parser = HTMLParser(settings.img_to_url_api, settings.ocr_api)
llm_extractor = LLMExtractor(settings.llm_api)

@router.post("/extract")
async def extract_emails(file: UploadFile = File(...)):
    """提交邮件提取任务"""
    # 1. 解析JSON文件
    try:
        content = await file.read()
        data = json.loads(content.decode('utf-8'))
        request = ExtractRequest(**data)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"JSON解析失败: {str(e)}")

    # 2. 为每封邮件创建任务
    task_ids = []
    for idx, email in enumerate(request.emails):
        task_id = task_manager.create_task(idx)
        task_ids.append(task_id)

        # 3. 异步处理
        task_manager.submit_task(task_id, email, process_email_task)

    return {"task_ids": task_ids, "total": len(task_ids)}


@router.get("/task/{task_id}")
async def get_task_status(task_id: str):
    """查询任务状态"""
    task = task_manager.get_task(task_id)
    if not task:
        raise HTTPException(status_code=404, detail="任务不存在")

    return {
        "task_id": task.task_id,
        "status": task.status.value,
        "progress": task.progress,
        "message": task.message,
        "result": task.result,
        "error": task.error
    }


def process_email_task(task_id: str, email_data):
    """处理单封邮件（在线程池中执行）"""
    import asyncio

    # 更新状态为处理中
    task_manager.update_task(task_id, status=TaskStatus.PROCESSING, progress=0)

    try:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        loop.run_until_complete(_process_email_async(task_id, email_data))
    except Exception as e:
        task_manager.update_task(
            task_id,
            status=TaskStatus.FAILED,
            error=str(e)
        )


async def _process_email_async(task_id: str, email_data):
    """异步处理邮件"""

    # 1. 脱敏
    task_manager.update_task(task_id, message="正在进行脱敏处理...", progress=10)
    html_content = email_data.html_content
    if email_data.options.get('desensitize', True):
        html_content = desensitizer.desensitize_html(html_content)
    task_manager.update_task(task_id, progress=20)

    # 2. HTML解析 + 图片处理 + OCR
    task_manager.update_task(task_id, message="正在处理图片...", progress=30)
    processed_html, ocr_text = await html_parser.process_html(
        html_content,
        email_data.images
    )
    task_manager.update_task(task_id, progress=60)

    # 3. 构建完整内容
    full_content = f"""
主题: {email_data.subject}
发件人: {email_data.sender}
时间: {email_data.received_time}

正文:
{processed_html}

OCR识别结果:
{ocr_text}
"""

    # 4. LLM提取
    task_manager.update_task(task_id, message="正在进行LLM提取...", progress=70)
    extracted = await llm_extractor.extract(full_content, processed_html)
    task_manager.update_task(task_id, progress=80)

    # 5. 补充字段
    result = {
        "name": extracted.get('name', email_data.subject),
        "description": extracted.get('description', ''),
        "reason": extracted.get('reason', ''),
        "solution": extracted.get('solution', ''),
        "process": extracted.get('process', ''),
        "status": "open",
        "creator": "system",
        "create_time": datetime.now().isoformat()
    }

    # 6. 上传到数据库
    task_manager.update_task(task_id, message="正在入库...", progress=90)
    try:
        async with httpx.AsyncClient() as client:
            resp = await client.post(
                settings.upload_to_db_api,
                json=[result]
            )
            batch_id = resp.json().get('batch_id', '')
    except Exception as e:
        print(f"数据库上传失败: {e}")
        batch_id = ''

    # 7. 完成
    task_manager.update_task(
        task_id,
        status=TaskStatus.COMPLETED,
        progress=100,
        message="处理完成",
        result=result
    )
```

---

## 3. 前端实现

### 3.1 主界面

**frontend/main.py**
```python
import tkinter as tk
from tkinter import ttk
from email_window import EmailWindow

class MainWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("邮件提取工具")
        self.root.geometry("400x300")

        self._setup_ui()

    def _setup_ui(self):
        # 居中框架
        frame = tk.Frame(self.root)
        frame.place(relx=0.5, rely=0.5, anchor="center")

        # 邮件提取按钮
        btn_extract = tk.Button(
            frame,
            text="邮件提取",
            width=15,
            height=2,
            command=self._open_email_window
        )
        btn_extract.pack(pady=10)

        # 待规划按钮（禁用）
        btn_plan = tk.Button(
            frame,
            text="待规划",
            width=15,
            height=2,
            state="disabled"  # 禁用
        )
        btn_plan.pack(pady=10)

    def _open_email_window(self):
        EmailWindow(self.root)

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = MainWindow()
    app.run()
```

---

### 3.2 Outlook客户端

**frontend/outlook_client.py**
```python
import win32com.client
from datetime import datetime
from typing import List, Dict, Tuple

class OutlookClient:
    def __init__(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")

    def get_folders(self) -> List[Dict]:
        """获取所有文件夹"""
        folders = []

        # 根文件夹
        root = self.namespace.Folders
        for folder in root:
            folders.append({
                "name": folder.Name,
                "entry_id": folder.EntryID
            })
            # 子文件夹
            try:
                for subfolder in folder.Folders:
                    folders.append({
                        "name": f"{folder.Name}/{subfolder.Name}",
                        "entry_id": subfolder.EntryID
                    })
            except:
                pass

        return folders

    def get_default_folder(self, folder_type: int = 6) -> win32com.client.CDispatch:
        """获取默认文件夹
        6 = 收件箱, 4 = 已发送, 3 = 发件箱, etc.
        """
        return self.namespace.GetDefaultFolder(folder_type)

    def get_folder_by_entry_id(self, entry_id: str):
        """通过EntryID获取文件夹"""
        return self.namespace.GetFolderFromID(entry_id)

    def get_emails(
        self,
        folder,
        start_date: datetime = None,
        end_date: datetime = None,
        subject_filter: str = None
    ) -> List[Dict]:
        """获取邮件列表"""
        emails = []

        try:
            items = folder.Items
            # 排序
            items.Sort("[ReceivedTime]", True)

            for item in items:
                if item.Class != 43:  # 43 = IPM.Note (邮件)
                    continue

                email_data = {
                    "subject": item.Subject or "",
                    "sender": item.SenderEmailAddress or "",
                    "received_time": item.ReceivedTime,
                    "conversation_topic": item.ConversationTopic or "",
                    "entry_id": item.EntryID,
                    "html_content": item.HTMLBody or ""
                }

                # 解析图片（简化版）
                email_data["images"] = self._extract_images(item)

                # 日期筛选
                if start_date and email_data["received_time"] < start_date:
                    continue
                if end_date and email_data["received_time"] > end_date:
                    continue

                # 主题筛选
                if subject_filter and subject_filter.lower() not in email_data["subject"].lower():
                    continue

                emails.append(email_data)

        except Exception as e:
            print(f"获取邮件失败: {e}")

        return emails

    def _extract_images(self, mail_item) -> List[Dict]:
        """提取邮件中的图片（简化版）"""
        images = []

        try:
            # 尝试获取附件中的图片
            attachments = mail_item.Attachments
            for idx, att in enumerate(attachments):
                if att.Type == 1:  # olByValue
                    try:
                        # 获取图片内容
                        import base64
                        import tempfile

                        temp_path = tempfile.NamedTemporaryFile(delete=False, suffix='.tmp')
                        temp_path.close()
                        att.SaveAsFile(temp_path.name)

                        with open(temp_path.name, 'rb') as f:
                            data = base64.b64encode(f.read()).decode()
                            images.append({
                                "position": idx,
                                "base64": f"data:image/png;base64,{data}"
                            })

                        import os
                        os.unlink(temp_path.name)
                    except:
                        pass
        except:
            pass

        return images

    def deduplicate_by_conversation(self, emails: List[Dict]) -> List[Dict]:
        """按ConversationTopic去重，保留最新"""
        latest = {}

        for email in emails:
            topic = email["conversation_topic"]
            if not topic:
                topic = email["entry_id"]  # 无主题则用ID

            if topic not in latest:
                latest[topic] = email
            else:
                # 比较时间，保留最新的
                if email["received_time"] > latest[topic]["received_time"]:
                    latest[topic] = email

        return list(latest.values())
```

---

### 3.3 API客户端

**frontend/api_client.py**
```python
import httpx
from typing import List, Dict, Optional

class APIClient:
    def __init__(self, base_url: str = "http://localhost:8000"):
        self.base_url = base_url

    async def extract_emails(self, json_data: dict) -> Dict:
        """提交邮件数据"""
        files = {'file': ('emails.json', json.dumps(json_data).encode('utf-8'), 'application/json')}

        async with httpx.AsyncClient() as client:
            resp = await client.post(f"{self.base_url}/api/extract", files=files)
            resp.raise_for_status()
            return resp.json()

    async def get_task_status(self, task_id: str) -> Dict:
        """查询任务状态"""
        async with httpx.AsyncClient() as client:
            resp = await client.get(f"{self.base_url}/api/task/{task_id}")
            resp.raise_for_status()
            return resp.json()
```

---

### 3.4 邮件提取窗口

**frontend/email_window.py**
```python
import tkinter as tk
from tkinter import ttk, messagebox
import asyncio
from datetime import datetime
from outlook_client import OutlookClient
from api_client import APIClient
import json

class EmailWindow:
    def __init__(self, parent):
        self.parent = parent
        self.window = tk.Toplevel(parent)
        self.window.title("邮件提取")
        self.window.geometry("1000x700")

        self.outlook_client = OutlookClient()
        self.api_client = APIClient()

        self.folders = []
        self.emails = []
        self.selected_emails = []
        self.task_ids = []

        self._setup_ui()
        self._load_folders()

    def _setup_ui(self):
        # 筛选条件区域
        filter_frame = tk.LabelFrame(self.window, text="筛选条件", padx=10, pady=10)
        filter_frame.pack(fill="x", padx=10, pady=5)

        # 日期范围
        tk.Label(filter_frame, text="起始日期:").grid(row=0, column=0, sticky="w")
        self.start_date_var = tk.StringVar(value="2021-01-01")
        tk.Entry(filter_frame, textvariable=self.start_date_var, width=15).grid(row=0, column=1, padx=5)

        tk.Label(filter_frame, text="截止日期:").grid(row=0, column=2, sticky="w")
        self.end_date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        tk.Entry(filter_frame, textvariable=self.end_date_var, width=15).grid(row=0, column=3, padx=5)

        # 文件夹选择
        tk.Label(filter_frame, text="文件夹:").grid(row=1, column=0, sticky="w", pady=5)
        self.folder_var = tk.StringVar()
        self.folder_combo = ttk.Combobox(filter_frame, textvariable=self.folder_var, width=30)
        self.folder_combo.grid(row=1, column=1, columnspan=2, padx=5, pady=5)

        # 主题搜索
        tk.Label(filter_frame, text="主题搜索:").grid(row=2, column=0, sticky="w")
        self.subject_var = tk.StringVar()
        tk.Entry(filter_frame, textvariable=self.subject_var, width=30).grid(row=2, column=1, columnspan=2, padx=5)

        # 按钮
        tk.Button(filter_frame, text="刷新列表", command=self._load_emails).grid(row=3, column=0, pady=10)
        tk.Button(filter_frame, text="提取选中", command=self._extract_selected).grid(row=3, column=1, pady=10)

        # 邮件列表
        list_frame = tk.Frame(self.window)
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)

        columns = ("select", "subject", "sender", "received_time", "conversation_topic")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=15)

        self.tree.heading("select", text="选择")
        self.tree.heading("subject", text="主题")
        self.tree.heading("sender", text="发件人")
        self.tree.heading("received_time", text="收件时间")
        self.tree.heading("conversation_topic", text="会话主题")

        self.tree.column("select", width=50)
        self.tree.column("subject", width=300)
        self.tree.column("sender", width=150)
        self.tree.column("received_time", width=150)
        self.tree.column("conversation_topic", width=200)

        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 任务状态区域
        status_frame = tk.LabelFrame(self.window, text="任务状态", padx=10, pady=10)
        status_frame.pack(fill="x", padx=10, pady=5)

        self.task_label = tk.Label(status_frame, text="无任务")
        self.task_label.pack(side="left")

        tk.Button(status_frame, text="刷新状态", command=self._refresh_status).pack(side="right")

    def _load_folders(self):
        """加载文件夹列表"""
        try:
            self.folders = self.outlook_client.get_folders()
            self.folder_combo['values'] = [f["name"] for f in self.folders]
            if self.folders:
                self.folder_var.set(self.folders[0]["name"])  # 默认收件箱
        except Exception as e:
            messagebox.showerror("错误", f"无法连接Outlook: {e}")

    def _load_emails(self):
        """加载邮件列表"""
        try:
            # 解析日期
            start_date = datetime.strptime(self.start_date_var.get(), "%Y-%m-%d")
            end_date = datetime.strptime(self.end_date_var.get(), "%Y-%m-%d")
            end_date = end_date.replace(hour=23, minute=59, second=59)

            # 获取文件夹
            folder_name = self.folder_var.get()
            folder = self.outlook_client.get_default_folder(6)  # 默认收件箱

            # 获取邮件
            self.emails = self.outlook_client.get_emails(
                folder,
                start_date=start_date,
                end_date=end_date,
                subject_filter=self.subject_var.get()
            )

            # 去重
            self.emails = self.outlook_client.deduplicate_by_conversation(self.emails)

            # 显示
            self._display_emails()

        except Exception as e:
            messagebox.showerror("错误", f"加载邮件失败: {e}")

    def _display_emails(self):
        """显示邮件列表"""
        # 清空
        for item in self.tree.get_children():
            self.tree.delete(item)

        # 插入数据
        for email in self.emails:
            received = email["received_time"].strftime("%Y-%m-%d %H:%M") if email["received_time"] else ""
            self.tree.insert("", "end", values=(
                "",
                email["subject"][:50] + "..." if len(email["subject"]) > 50 else email["subject"],
                email["sender"],
                received,
                email["conversation_topic"][:30] + "..." if len(email["conversation_topic"]) > 30 else email["conversation_topic"]
            ), tags=(email["entry_id"],))

    def _extract_selected(self):
        """提取选中的邮件"""
        # 获取选中项
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请先选择邮件")
            return

        # 获取选中的邮件数据
        selected_emails = []
        for item in selected_items:
            entry_id = self.tree.item(item)["tags"][0]
            for email in self.emails:
                if email["entry_id"] == entry_id:
                    selected_emails.append({
                        "subject": email["subject"],
                        "sender": email["sender"],
                        "received_time": email["received_time"].strftime("%Y-%m-%d %H:%M:%S") if email["received_time"] else "",
                        "conversation_topic": email["conversation_topic"],
                        "html_content": email["html_content"],
                        "images": email["images"]
                    })
                    break

        # 构建请求数据
        json_data = {
            "emails": selected_emails,
            "options": {"desensitize": True}
        }

        try:
            # 调用API
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            result = loop.run_until_complete(self.api_client.extract_emails(json_data))

            self.task_ids = result.get("task_ids", [])
            self.task_label.config(text=f"任务已提交: {', '.join(self.task_ids)}")
            messagebox.showinfo("成功", f"已提交 {len(self.task_ids)} 个任务")

        except Exception as e:
            messagebox.showerror("错误", f"提交失败: {e}")

    def _refresh_status(self):
        """刷新任务状态"""
        if not self.task_ids:
            messagebox.showwarning("警告", "没有任务")
            return

        try:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)

            status_text = []
            for task_id in self.task_ids:
                result = loop.run_until_complete(self.api_client.get_task_status(task_id))
                status = result.get("status", "unknown")
                progress = result.get("progress", 0)
                status_text.append(f"{task_id}: {status} ({progress}%)")

            self.task_label.config(text="\n".join(status_text))

        except Exception as e:
            messagebox.showerror("错误", f"查询失败: {e}")
```

---

## 4. 敏感词配置

**backend/sensitive_words.txt**
```
张三
李四
13800138000
13900139000
zhangsan@company.com
lisi@company.com
身份证
银行卡
```

---

## 5. 启动方式

### 后端
```bash
cd backend
pip install -r requirements.txt
python main.py
```

### 前端
```bash
pip install pywin32
python main.py
```

---

## 6. 待优化项

- [ ] 前端窗口样式美化
- [ ] 错误重试机制
- [ ] 日志记录
- [ ] 单元测试
- [ ] 打包为exe
