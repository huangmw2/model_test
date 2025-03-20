import os
import json
from typing import Any, Dict
from package.logger import log_message

class JsonDataHandler:
    def __init__(self,file_dir):
        self.user_data_path = os.path.join(file_dir,"user_data.json")
        self.user_data = self.read_json(self.user_data_path)

    def read_json(self,file_path: str) -> Dict[str, Any]:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            log = f"文件 {file_path} 不存在"
            log_message(log,"ERROR")
            return {}
        except json.JSONDecodeError:
            log = f"文件 {file_path} 格式错误"
            log_message(log,"ERROR")
            return {}
        
    def write_json(self) -> None:
        self.clear_json()
        with open(self.user_data_path, 'w', encoding='utf-8') as f:
            json.dump(self.user_data, f, ensure_ascii=False, indent=2)
        log = f"数据已保存至 {self.user_data_path}"
        log_message(log,"DEBUG")

    def clear_json(self):
        open(self.user_data_path, 'w').close()

    def update_data(self, key: str, value: Any) -> None:
        """更新 self.user_data 中的内容，不立即写入文件"""
        self.user_data[key] = value  # 先存入内存
        log = f"数据已更新到内存: {key} = {value}"
        log_message(log,"DEBUG")

    def find_data(self, key: str) -> Any:
        """查找 self.user_data 中的某个键值"""
        return self.user_data.get(key, None)

    def delete_data(self, key: str) -> None:
        """从 self.user_data 中删除某个键值"""
        if key in self.user_data:
            del self.user_data[key]
            log = f"数据 {key} 已从内存中删除"
            log_message(log,"DEBUG")
        else:
            log = f"键 {key} 不存在"
            log_message(log,"DEBUG")