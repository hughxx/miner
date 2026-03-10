from fastapi import APIRouter, UploadFile, File, HTTPException
from typing import List
import json
from datetime import datetime
import asyncio
import httpx

from models import ExtractRequest, TaskStatus
from service.task_manager import task_manager
from service.html_parser import HTMLParser
from service.desensitize import desensitizer
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
