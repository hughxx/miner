# 邮件提取工具 - 接口规格说明书

## 1. 技术栈

| 层级 | 技术 | 说明 |
|------|------|------|
| 前端 | Python 3.x + tkinter | 桌面客户端界面 |
| 前端 | pywin32 | Outlook COM接口 |
| 后端 | FastAPI | Web服务框架 |
| 后端 | uvicorn | ASGI服务器 |
| 后端 | httpx | 异步HTTP客户端 |
| 后端 | python-multipart | 文件上传解析 |
| 后端 | lxml + BeautifulSoup | HTML解析 |

---

## 2. 目录结构

```
mail_extractor/
├── frontend/
│   ├── main.py              # 主界面入口
│   ├── email_window.py      # 邮件提取窗口
│   ├── outlook_client.py    # Outlook操作封装
│   └── api_client.py        # 后端API调用
├── backend/
│   ├── main.py              # FastAPI主程序
│   ├── config.py            # 配置文件
│   ├── models.py            # 数据模型
│   ├── router/
│   │   └── extract.py       # 提取相关路由
│   ├── service/
│   │   ├── task_manager.py  # 任务管理
│   │   ├── desensitize.py   # 脱敏处理
│   │   ├── html_parser.py   # HTML解析+图片处理
│   │   └── llm_extractor.py # LLM提取
│   └── sensitive_words.txt  # 敏感词列表
└── docs/
    ├── SPEC.md              # 需求规格
    └── INTERFACE.md         # 本文档
```

---

## 3. API接口设计

### 3.1 前端 → 后端

#### POST /api/extract

**功能**：提交邮件数据，返回task_id

**Content-Type**：`multipart/form-data`

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| file | Binary | 是 | JSON文件（UTF-8编码） |

**请求体JSON结构**（在file参数中）：

```json
{
  "emails": [
    {
      "subject": "邮件主题",
      "sender": "发件人邮箱",
      "received_time": "2024-01-01 10:00:00",
      "conversation_topic": "会话主题",
      "html_content": "<html>...</html>",
      "images": [
        {
          "position": 0,
          "base64": "data:image/png;base64,..."
        }
      ]
    }
  ],
  "options": {
    "desensitize": true
  }
}
```

**响应**（200 OK）：

```json
{
  "task_ids": ["task_001", "task_002"],
  "total": 2
}
```

---

#### GET /api/task/{task_id}

**功能**：查询单个任务状态

**路径参数**：

| 参数 | 类型 | 说明 |
|------|------|------|
| task_id | string | 任务ID |

**响应**（200 OK）：

```json
{
  "task_id": "task_001",
  "status": "processing",
  "progress": 60,
  "message": "正在进行OCR识别...",
  "result": null
}
```

**status枚举**：
| 值 | 说明 |
|-----|------|
| pending | 等待处理 |
| processing | 处理中 |
| completed | 处理完成 |
| failed | 处理失败 |

**progress说明**：
| 阶段 | progress值 |
|------|-------------|
| 待处理 | 0 |
| 脱敏完成 | 20 |
| 图片处理完成 | 40 |
| OCR完成 | 60 |
| LLM提取完成 | 80 |
| 入库完成 | 100 |

---

### 3.2 后端 → 外部服务

#### Img_to_url

| 项目 | 内容 |
|------|------|
| 调用方式 | HTTP POST |
| URL | 外部提供 |
| 输入 | file: 二进制文件流 + filename: 文件名 |
| 输出 | url: string (http链接) |

---

#### ocr

| 项目 | 内容 |
|------|------|
| 调用方式 | HTTP POST |
| URL | 外部提供 |
| 输入 | base64_str: string |
| 输出 | ocr_result: string |

---

#### llm

| 项目 | 内容 |
|------|------|
| 调用方式 | HTTP POST |
| URL | 外部提供 |
| 输入 | prompt: string |
| 输出 | resp: string |

---

#### upload_to_db

| 项目 | 内容 |
|------|------|
| 调用方式 | HTTP POST |
| URL | 外部提供 |
| 输入 | json_list: list[dict] |
| 输出 | batch_id: string |

**json_list元素结构**：

```json
{
  "name": "邮件主题",
  "description": "描述内容",
  "reason": "原因",
  "solution": "解决方案",
  "process": "处理过程",
  "status": "open",
  "creator": "system",
  "create_time": "2024-01-01 10:00:00"
}
```

---

## 4. 数据模型

### 4.1 Task模型

```python
class TaskStatus(str, Enum):
    PENDING = "pending"
    PROCESSING = "processing"
    COMPLETED = "completed"
    FAILED = "failed"

class Task:
    task_id: str
    email_index: int          # 对应emails数组的下标
    status: TaskStatus
    progress: int            # 0-100
    message: str
    result: Optional[dict]    # 提取结果
    error: Optional[str]     # 错误信息
    created_at: datetime
    updated_at: datetime
```

### 4.2 处理阶段定义

| 阶段 | progress | 说明 |
|------|----------|------|
| INIT | 0 | 任务创建 |
| DESENSITIZE | 20 | 脱敏完成 |
| IMAGE_PROCESS | 40 | 图片转超链接完成 |
| OCR_DONE | 60 | OCR识别完成 |
| LLM_DONE | 80 | LLM提取完成 |
| COMPLETE | 100 | 入库完成 |

---

## 5. 敏感词配置

### 5.1 文件位置

`backend/sensitive_words.txt`

### 5.2 文件格式

纯文本，每行一个敏感词：

```
张三
13800138000
zhangsan@company.com
身份证
银行卡
```

---

## 6. 前端界面交互

### 6.1 主界面 → 邮件提取窗口

```python
# 点击"邮件提取"按钮
def on_extract_click():
    EmailWindow()  # 新窗口
```

### 6.2 邮件提取窗口流程

```
1. 初始化
   ↓
2. 连接Outlook，加载文件夹列表
   ↓
3. 用户选择筛选条件（日期、文件夹、关键词）
   ↓
4. 点击"刷新列表"
   ↓
5. 获取邮件 → 去重 → 显示列表
   ↓
6. 用户多选邮件 → 点击"提取选中"
   ↓
7. 构建JSON → POST /api/extract
   ↓
8. 拿到task_ids，显示在界面
   ↓
9. 用户点击"刷新状态"
   ↓
10. 轮询GET /api/task/{task_id}
    ↓
11. 状态变为completed，显示result
```

---

## 7. 异常处理

| 场景 | 处理方式 |
|------|----------|
| Outlook连接失败 | 弹窗提示"无法连接Outlook，请确保Outlook已启动" |
| 后端服务不可达 | 弹窗提示"无法连接服务器" |
| 任务失败 | 显示error信息，前端可重试 |
| 网络超时 | 弹窗提示"请求超时，请重试" |

---

## 8. 配置项

### 8.1 后端配置（config.py）

```python
# 服务配置
HOST = "0.0.0.0"
PORT = 8000

# 外部服务URL
IMG_TO_URL_API = "http://localhost:8001/Img_to_url"
OCR_API = "http://localhost:8002/ocr"
LLM_API = "http://localhost:8003/llm"
UPLOAD_TO_DB_API = "http://localhost:8004/upload_to_db"

# 敏感词文件路径
SENSITIVE_WORDS_FILE = "sensitive_words.txt"
```

---

## 9. 待定项

- [ ] 外部服务URL需根据实际部署环境配置
- [ ] LLM的prompt模板需根据实际需求设计
- [ ] 日志配置
- [ ] 部署文档
