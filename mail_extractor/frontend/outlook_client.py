import win32com.client
from datetime import datetime
from typing import List, Dict
import base64
import tempfile
import os

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
        6 = 收件箱, 4 = 已发送, 3 = 发件箱
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
            print(f"[DEBUG] 邮件总数: {len(items)}")

            # 先不过滤，看看有多少邮件
            count = 0
            for item in items:
                count += 1
                if count > 50:  # 只打印前50封的调试信息
                    break
                print(f"[DEBUG] Item Class: {item.Class}, Subject: {item.Subject[:30] if item.Subject else 'None'}")

            # 按时间倒序
            items.Sort("[ReceivedTime]", True)

            for item in items:
                if item.Class != 43:  # 43 = IPM.Note (邮件)
                    continue

                received_time = item.ReceivedTime
                email_data = {
                    "subject": item.Subject or "",
                    "sender": item.SenderEmailAddress or "",
                    "received_time": received_time,
                    "conversation_topic": item.ConversationTopic or "",
                    "entry_id": item.EntryID,
                    "html_content": item.HTMLBody or ""
                }

                # 解析图片
                email_data["images"] = self._extract_images(item)

                # 日期筛选
                if start_date and received_time and received_time < start_date:
                    continue
                if end_date and received_time and received_time > end_date:
                    continue

                # 主题筛选
                if subject_filter and subject_filter.lower() not in email_data["subject"].lower():
                    continue

                emails.append(email_data)

        except Exception as e:
            print(f"获取邮件失败: {e}")

        return emails

    def _extract_images(self, mail_item) -> List[Dict]:
        """提取邮件中的图片"""
        images = []

        try:
            attachments = mail_item.Attachments
            for idx, att in enumerate(attachments):
                if att.Type == 1:  # olByValue
                    try:
                        # 保存到临时文件
                        temp_path = tempfile.NamedTemporaryFile(delete=False, suffix='.tmp')
                        temp_path.close()
                        att.SaveAsFile(temp_path.name)

                        # 读取并转为base64
                        with open(temp_path.name, 'rb') as f:
                            data = base64.b64encode(f.read()).decode()

                        # 根据附件名判断mime类型
                        filename = att.FileName
                        mime = 'image/png'
                        if filename.lower().endswith(('.jpg', '.jpeg')):
                            mime = 'image/jpeg'
                        elif filename.lower().endswith('.gif'):
                            mime = 'image/gif'

                        images.append({
                            "position": idx,
                            "base64": f"data:{mime};base64,{data}"
                        })

                        os.unlink(temp_path.name)
                    except Exception as e:
                        print(f"提取图片失败: {e}")
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
                current_time = email["received_time"]
                existing_time = latest[topic]["received_time"]
                if current_time and existing_time and current_time > existing_time:
                    latest[topic] = email

        return list(latest.values())
