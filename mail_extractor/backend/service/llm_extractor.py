import httpx
import json
import re
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
