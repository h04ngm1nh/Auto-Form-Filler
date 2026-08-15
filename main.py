#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Google Forms Automated Form Filler (CLI Script)
Author: Senior Backend & Automation Engineer
Description: 
    Kịch bản tự động hóa gửi dữ liệu hàng loạt từ file Excel/CSV vào Google Forms 
    thông qua giao thức HTTP POST Request trực tiếp (Bypass UI/Headless).
    Hỗ trợ xử lý Checkbox (multi-select), Date, Dropdown, Radio và tự động báo cáo kết quả.
"""

import os
import sys
import json
import time
import random
import argparse
import logging
from typing import List, Dict, Any, Tuple, Optional
import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util import Retry
from tqdm import tqdm

# Cấu hình lại mã hóa UTF-8 cho console Windows để tránh lỗi UnicodeEncodeError khi log tiếng Việt
if sys.platform.startswith("win"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
        sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    except AttributeError:
        pass

# Thiết lập hệ thống Log hiển thị trên Console
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("GoogleFormFiller")

# Danh sách các User-Agent phổ biến để mô phỏng ngẫu nhiên các trình duyệt thực tế
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/121.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.2 Safari/605.1.15"
]

class ConfigLoader:
    """Class đảm nhận việc đọc, phân tích và xác thực cấu hình config.json"""
    
    @staticmethod
    def load(config_path: str) -> Dict[str, Any]:
        """Đọc file config.json và kiểm tra tính hợp lệ của các trường bắt buộc"""
        if not os.path.exists(config_path):
            raise FileNotFoundError(f"Không tìm thấy file cấu hình: {config_path}")
            
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
            
        # Xác thực cấu hình
        if "form_id" not in config or not config["form_id"]:
            raise ValueError("Thiếu hoặc trống trường 'form_id' trong file cấu hình.")
        if "mappings" not in config or not config["mappings"]:
            raise ValueError("Thiếu hoặc trống trường 'mappings' trong file cấu hình.")
            
        # Bổ sung cấu hình settings mặc định nếu chưa khai báo
        if "settings" not in config:
            config["settings"] = {}
        defaults = {
            "min_delay": 1.0,
            "max_delay": 3.0,
            "max_retries": 3,
            "retry_backoff": 2.0,
            "timeout": 10
        }
        for key, val in defaults.items():
            config["settings"].setdefault(key, val)
            
        return config


class DataNormalizer:
    """Class xử lý làm sạch và chuẩn hóa dữ liệu từ Excel/CSV phù hợp với giao thức của Google Forms"""

    @staticmethod
    def clean_cell_value(val: Any) -> str:
        """Chuẩn hóa dữ liệu ô đơn lẻ, chuyển float dạng .0 về int và loại bỏ khoảng trắng"""
        if pd.isna(val):
            return ""
        # Trường hợp giá trị float nhưng biểu diễn số nguyên (ví dụ: số điện thoại hoặc ID bị đọc nhầm dạng số)
        if isinstance(val, float):
            if val.is_integer():
                return str(int(val))
        return str(val).strip()

    @classmethod
    def normalize_row(cls, row: pd.Series, mappings: Dict[str, Any]) -> List[Tuple[str, str]]:
        """
        Chuyển đổi một dòng dữ liệu trong file Excel thành danh sách payload HTTP POST
        Sử dụng cấu trúc danh sách tuple (List[Tuple[str, str]]) để hỗ trợ gửi lặp key (như Checkbox)
        """
        payload: List[Tuple[str, str]] = []
        
        for col_name, field_cfg in mappings.items():
            if col_name not in row:
                continue
                
            raw_val = row[col_name]
            entry_id = field_cfg["entry_id"]
            field_type = field_cfg.get("type", "text").lower()
            
            # Bỏ qua nếu ô trống
            if pd.isna(raw_val) or str(raw_val).strip() == "":
                continue

            if field_type == "checkbox":
                # Đối với câu hỏi Checkbox, cho phép nhiều giá trị ngăn cách bởi ký tự định nghĩa (mặc định dấu phẩy)
                separator = field_cfg.get("separator", ",")
                values = [v.strip() for v in str(raw_val).split(separator) if v.strip()]
                for val in values:
                    payload.append((entry_id, val))
                    
            elif field_type == "date":
                # Đối với câu hỏi Date, Google Forms yêu cầu tách thành _year, _month, _day
                try:
                    dt = pd.to_datetime(raw_val)
                    payload.append((f"{entry_id}_year", str(dt.year)))
                    payload.append((f"{entry_id}_month", f"{dt.month:02d}"))
                    payload.append((f"{entry_id}_day", f"{dt.day:02d}"))
                except Exception as e:
                    logger.warning(f"Lỗi định dạng ngày tháng tại cột '{col_name}', giá trị: {raw_val}. Chi tiết: {e}")
                    # Gửi fallback dạng text thông thường nếu parse lỗi
                    payload.append((entry_id, cls.clean_cell_value(raw_val)))
            else:
                # Text, Radio, Dropdown, Email, Phone...
                payload.append((entry_id, cls.clean_cell_value(raw_val)))
                
        return payload


class FormSubmitter:
    """Class đảm nhận kết nối HTTP, quản lý phiên và thực hiện gửi request POST"""

    def __init__(self, form_id: str, settings: Dict[str, Any]):
        self.form_id = form_id
        self.settings = settings
        self.endpoint = f"https://docs.google.com/forms/d/e/{form_id}/formResponse"
        
        # Khởi tạo session để tái sử dụng kết nối HTTP (Keep-Alive), tối ưu hóa tốc độ
        self.session = requests.Session()
        
        # Thiết lập cơ chế tự động thử lại ở cấp độ Network/HTTP (dành cho lỗi kết nối hoặc 5xx)
        retries = Retry(
            total=self.settings["max_retries"],
            backoff_factor=self.settings["retry_backoff"],
            status_forcelist=[429, 500, 502, 503, 504],
            raise_on_status=False
        )
        self.session.mount("https://", HTTPAdapter(max_retries=retries))

    def _get_headers(self) -> Dict[str, str]:
        """Tạo bộ headers giả lập trình duyệt thật để bypass các hệ thống kiểm duyệt cơ bản"""
        return {
            "User-Agent": random.choice(USER_AGENTS),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7",
            "Content-Type": "application/x-www-form-urlencoded",
            "Origin": "https://docs.google.com",
            "Referer": f"https://docs.google.com/forms/d/e/{self.form_id}/viewform",
            "Connection": "keep-alive"
        }

    def submit(self, payload: List[Tuple[str, str]]) -> Tuple[bool, str]:
        """
        Thực hiện gửi POST Request lên Google Forms
        Trả về: Tuple (trạng thái gửi thành công, thông báo chi tiết)
        """
        headers = self._get_headers()
        
        try:
            # Gửi request với dữ liệu dạng URL-encoded
            response = self.session.post(
                url=self.endpoint,
                data=payload,
                headers=headers,
                timeout=self.settings["timeout"]
            )
            
            # Phân tích phản hồi
            if response.status_code == 200:
                # Kiểm tra nội dung trang HTML để khẳng định việc gửi thành công
                # Google Forms sẽ phản hồi mã 200 kèm form HTML trống/báo lỗi nếu dữ liệu không hợp lệ.
                # Khi gửi thành công, thường sẽ chứa các từ khóa xác nhận.
                success_keywords = ["Your response has been recorded", "Câu trả lời của bạn đã được ghi lại", "formResponse"]
                response_text = response.text
                
                is_recorded = any(keyword in response_text for keyword in success_keywords)
                if is_recorded:
                    return True, "Gửi thành công (Phản hồi được ghi nhận)"
                else:
                    # Nếu không tìm thấy từ khóa xác nhận, có thể form bị dính Validation Error (Ví dụ: Thiếu trường bắt buộc)
                    return False, "Google Forms trả về trang biểu mẫu lỗi (có thể do sai định dạng trường dữ liệu bắt buộc)"
            else:
                return False, f"HTTP Error Status Code: {response.status_code}"
                
        except requests.exceptions.RequestException as e:
            return False, f"Lỗi kết nối mạng: {str(e)}"


class ExecutionController:
    """Class điều phối toàn bộ quy trình đọc file, phân phối request và ghi xuất báo cáo"""

    def __init__(self, data_path: str, config_path: str, output_path: str):
        self.data_path = data_path
        self.config_path = config_path
        self.output_path = output_path
        
    def run(self):
        logger.info("=== BẮT ĐẦU QUY TRÌNH GỬI DỮ LIỆU TỰ ĐỘNG ===")
        
        # 1. Đọc và kiểm tra file cấu hình
        try:
            config = ConfigLoader.load(self.config_path)
            logger.info("Nạp file cấu hình thành công.")
        except Exception as e:
            logger.error(f"Lỗi nạp cấu hình: {e}")
            sys.exit(1)
            
        # 2. Đọc file dữ liệu Excel/CSV bằng Pandas
        try:
            if self.data_path.endswith(".csv"):
                df = pd.read_csv(self.data_path)
            else:
                df = pd.read_excel(self.data_path)
            logger.info(f"Đọc thành công file dữ liệu: {self.data_path}. Tổng số dòng: {len(df)}")
        except Exception as e:
            logger.error(f"Lỗi đọc file dữ liệu: {e}")
            sys.exit(1)
            
        # 3. Khởi tạo đối tượng submitter
        submitter = FormSubmitter(config["form_id"], config["settings"])
        
        # Khai báo cột ghi nhận trạng thái
        status_list: List[str] = []
        error_details: List[str] = []
        
        # 4. Vòng lặp gửi dữ liệu tích hợp thanh tiến trình trực quan
        min_delay = config["settings"]["min_delay"]
        max_delay = config["settings"]["max_delay"]
        
        logger.info("Đang thực hiện gửi dữ liệu...")
        for index, row in tqdm(df.iterrows(), total=len(df), desc="Tiến trình gửi", unit="dòng"):
            row_num = index + 1
            
            # Chuẩn hóa dữ liệu sang định dạng payload
            payload = DataNormalizer.normalize_row(row, config["mappings"])
            
            # Thực thi gửi
            success, msg = submitter.submit(payload)
            
            if success:
                status_list.append("Success")
                error_details.append(msg)
            else:
                status_list.append("Failed")
                error_details.append(msg)
                logger.warning(f"Dòng #{row_num} gửi thất bại. Chi tiết: {msg}")
                
            # Tạo độ trễ ngẫu nhiên (Random Jitter) để tránh bị chặn IP / phát hiện bot
            if index < len(df) - 1:
                delay = random.uniform(min_delay, max_delay)
                time.sleep(delay)
                
        # 5. Xuất báo cáo Excel kết quả
        df["Submission_Status"] = status_list
        df["Submission_Error_Detail"] = error_details
        
        try:
            # Ghi đè hoặc tạo mới file Excel kết quả
            df.to_excel(self.output_path, index=False, engine="openpyxl")
            logger.info(f"Đã xuất báo cáo kết quả gửi dữ liệu thành công: {self.output_path}")
        except Exception as e:
            logger.error(f"Không thể ghi file báo cáo: {e}")
            
        # Thống kê nhanh
        success_count = status_list.count("Success")
        failed_count = status_list.count("Failed")
        logger.info("=== QUY TRÌNH HOÀN THÀNH ===")
        logger.info(f"Thành công: {success_count}/{len(df)} dòng.")
        logger.info(f"Thất bại: {failed_count}/{len(df)} dòng.")


def main():
    parser = argparse.ArgumentParser(
        description="Google Forms Direct HTTP Auto Filler Tool.",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument(
        "-i", "--input", 
        required=False,
        default="sample_data.xlsx",
        help="Đường dẫn tới file Excel/CSV chứa dữ liệu cần nạp."
    )
    parser.add_argument(
        "-c", "--config", 
        required=False,
        default="config.json",
        help="Đường dẫn tới file cấu hình config.json (mặc định: config.json)."
    )
    parser.add_argument(
        "-o", "--output", 
        required=False,
        default="result_log.xlsx",
        help="Đường dẫn lưu kết quả báo cáo Excel (mặc định: result_log.xlsx)."
    )
    
    args = parser.parse_args()
    
    # Kiểm tra sự tồn tại của file dữ liệu
    if not os.path.exists(args.input):
        logger.error(f"Không tìm thấy file dữ liệu đầu vào: {args.input}")
        logger.info("Vui lòng cung cấp đường dẫn chính xác bằng tham số: python main.py -i <path_to_excel_or_csv>")
        sys.exit(1)
        
    controller = ExecutionController(
        data_path=args.input, 
        config_path=args.config, 
        output_path=args.output
    )
    controller.run()


if __name__ == "__main__":
    main()
