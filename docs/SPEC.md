# 邮件提取工具 - 需求规格说明书

## 1. 项目概述

| 项目 | 内容 |
|------|------|
| 项目名称 | 邮件提取工具 (MailExtractor) |
| 项目类型 | Python桌面客户端 + Flask后端服务 |
| 核心功能 | 从Outlook提取邮件内容，转化为结构化数据存入数据库 |
| 目标用户 | 需要批量处理邮件信息的办公人员 |

---

## 2. 功能需求

### 2.1 前端 - Python客户端

#### 2.1.1 主界面
- 窗口标题：`邮件提取工具`
- 窗口尺寸：800x600（可调整）
- 布局：垂直居中，两个按钮水平排列

| 按钮 | 功能 | 状态 |
|------|------|------|
| 邮件提取 | 打开邮件提取窗口 | 可点击 |
| 待规划 | 无功能 | 禁用（灰色） |

#### 2.1.2 邮件提取窗口

**筛选条件区域**：

| 条件 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| 起始日期 | 日期选择器 | 2021-01-01 | 必选 |
| 截止日期 | 日期选择器 | 今日 | 必选 |
| 文件夹 | 下拉选择 | 收件箱 | 可选，默认读取系统文件夹列表 |
| 主题搜索 | 文本输入框 | 空 | 可选，模糊匹配 |

**邮件列表区域**：
- 表格形式展示，每列：复选框 | 主题 | 发件人 | 收件时间 | ConversationTopic
- 支持多选（复选框）
- 双击行可预览邮件内容

**操作按钮**：
| 按钮 | 功能 |
|------|------|
| 刷新列表 | 根据筛选条件重新加载邮件 |
| 提取选中 | 将选中的邮件转为Word并上传到后端 |

**任务状态区域**：
- 显示最近一次操作的task_id
- "刷新状态"按钮：查询task处理进度
- 状态显示：pending / processing / completed / failed

#### 2.1.3 数据处理逻辑

**邮件去重规则**：
- 同一ConversationTopic的邮件，只保留最新（ReceivedTime最大）的一封

**Word文件生成**：
- 邮件HTML → Word文档
- 保留原始格式（图片、样式）

---

### 2.2 后端 - Flask服务

#### 2.2.1 接口设计

| 接口 | 方法 | 参数 | 返回 |
|------|------|------|------|
| /api/extract | POST | file (Word文件) | task_id |
| /api/task/{task_id} | GET | task_id | {status, progress, result} |

#### 2.2.2 处理流程（异步）

```
1. 接收Word文件
   ↓
2. 脱敏处理
   - 读取配置文件中的敏感词列表
   - 替换邮件内容中的敏感词为 ***
   ↓
3. 图片处理
   - 提取Word中的内嵌图片
   - 调用 Img_to_url(file) 转换为超链接
   - 替换Word中的图片为超链接文本
   ↓
4. OCR识别
   - 对Word中的图片调用 ocr(base64_str)
   - 识别结果附加到邮件内容后面
   ↓
5. LLM提取
   - 调用 llm(prompt) 提取字段
   - prompt需包含：name, description, reason, solution, process
   ↓
6. 数据入库
   - 补充字段：
     * status = "open"
     * creator = "当前用户"（可固定为系统用户）
     * create_time = 当前时间戳
   - 调用 upload_to_db([json_dict])
   ↓
7. 更新task状态为 completed
```

#### 2.2.3 数据库字段

| 字段 | 类型 | 说明 |
|------|------|------|
| name | string | 邮件名称/标题 |
| description | string | 描述 |
| reason | string | 原因 |
| solution | string | 解决方案 |
| process | string | 处理过程 |
| status | string | open / closed |
| creator | string | 创建人 |
| create_time | datetime | 创建时间 |

---

## 3. 技术架构

### 3.1 技术栈

| 层级 | 技术 | 说明 |
|------|------|------|
| 前端 | Python + tkinter | 桌面客户端 |
| 前端 | pywin32 | Outlook COM接口 |
| 前端 | python-docx | Word文件生成 |
| 后端 | Flask | Web服务 |
| 后端 | Flask-CORS | 跨域支持 |
| 后端 | celery | 异步任务队列 |

### 3.2 目录结构

```
mail_extractor/
├── frontend/
│   ├── main.py              # 主界面入口
│   ├── email_window.py      # 邮件提取窗口
│   ├── outlook_client.py    # Outlook操作封装
│   ├── word_converter.py   # HTML转Word
│   └── api_client.py        # 后端API调用
├── backend/
│   ├── app.py               # Flask主程序
│   ├── tasks.py             # Celery异步任务
│   ├── config.py            # 配置文件
│   ├── processor/
│   │   ├── desensitize.py  # 脱敏处理
│   │   ├── image_handler.py # 图片处理
│   │   └── llm_extractor.py # LLM提取
│   └── sensitive_words.txt  # 敏感词列表
├── docs/
│   └── SPEC.md              # 本规格文档
└── README.md
```

---

## 4. 已知约束

1. **一次处理一封邮件**：虽然upload_to_db支持批量，但为了task状态追踪，每次提取单独处理
2. **不处理附件**：仅处理邮件正文中的图文内容
3. **敏感词配置**：从后端项目配置文件读取
4. **日期范围**：固定起始2021-01-01，截止为当天

---

## 5. 待确认项（开发时可调整）

- [ ] 后端API端口号（建议5000）
- [ ] celery broker选择（redis/rabbitmq）
- [ ] 前端窗口样式细节
- [ ] 任务状态保存方式（内存/数据库）
