from bs4 import BeautifulSoup
import httpx
from typing import List, Tuple

class HTMLParser:
    def __init__(self, img_to_url_api: str, ocr_api: str):
        self.img_to_url_api = img_to_url_api
        self.ocr_api = ocr_api

    async def process_html(self, html: str, images: List[dict]) -> Tuple[str, str]:
        """
        处理HTML和图片
        返回: (处理后的HTML, 所有OCR结果拼接)
        """
        soup = BeautifulSoup(html, 'lxml')

        # 建立图片位置映射
        img_map = {img['position']: img['base64'] for img in images}

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

            # 根据实际API格式调整
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
