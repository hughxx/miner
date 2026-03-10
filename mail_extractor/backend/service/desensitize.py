import os
from typing import List

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
