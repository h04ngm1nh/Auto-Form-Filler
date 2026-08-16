#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Google Forms Automated Form Filler (CLI Script) - Refactored Version
Author: Senior Software Architect & Automation Specialist
Description:
    Kịch bản tự động hóa gửi dữ liệu hàng loạt từ file Excel/CSV vào Google Forms
    thông qua giao thức HTTP POST Request trực tiếp (Bypass UI/Headless).
    Hệ thống được thiết kế với khả năng phục hồi lỗi cao (Fault-tolerant), chống chặn bot,
    quản lý phiên nhất quán, phân tích phản hồi thông minh và cơ chế lưu vết checkpoint.
"""

import os
import re
import sys
import json
import time
import random
import datetime
import argparse
import logging
from typing import List, Dict, Any, Tuple, Optional, Set
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

# Danh sách các User-Agent phổ biến của các trình duyệt máy tính hiện đại
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:123.0) Gecko/20100101 Firefox/123.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.3 Safari/605.1.15",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36 Edg/122.0.0.0"
]


# ==============================================================================
# CUSTOM EXCEPTIONS FOR DATA VALIDATION
# ==============================================================================

class ValidationError(Exception):
    """Lớp ngoại lệ cơ sở cho tất cả lỗi xác thực dữ liệu tại local."""
    pass


class MissingRequiredFieldValidationError(ValidationError):
    """Ngoại lệ ném ra khi một trường bắt buộc bị trống trong dữ liệu nguồn."""
    pass


class DateValidationError(ValidationError):
    """Ngoại lệ ném ra khi định dạng ngày tháng trong file dữ liệu không thể parse."""
    pass


# ==============================================================================
# CONFIGURATION LOADER
# ==============================================================================

class ConfigLoader:
    """Class đảm nhận việc đọc, phân tích và xác thực cấu hình config.json."""

    @staticmethod
    def load(config_path: str) -> Dict[str, Any]:
        """
        Đọc file config.json và kiểm tra tính hợp lệ của các cấu hình bắt buộc.
        
        Args:
            config_path (str): Đường dẫn tới file config.json.
            
        Returns:
            Dict[str, Any]: Cấu hình đã được xác thực và gán giá trị mặc định.
        """
        if not os.path.exists(config_path):
            raise FileNotFoundError(f"Không tìm thấy file cấu hình: {config_path}")

        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)

        # Xác thực các trường cấu hình cốt lõi
        if "form_id" not in config or not config["form_id"]:
            raise ValueError("Thiếu hoặc trống trường 'form_id' trong file cấu hình.")
        if "mappings" not in config or not config["mappings"]:
            raise ValueError("Thiếu hoặc trống trường 'mappings' trong file cấu hình.")

        # Thiết lập các giá trị cấu hình mặc định cho settings nếu chưa khai báo
        if "settings" not in config:
            config["settings"] = {}
            
        defaults = {
            "min_delay": 2.0,
            "max_delay": 5.0,
            "max_retries": 3,
            "retry_backoff": 2.0,
            "timeout": 10,
            "session_rotation_limit": 10,
            "success_keywords": [
                "Your response has been recorded",
                "Câu trả lời của bạn đã được ghi lại"
            ]
        }
        for key, val in defaults.items():
            config["settings"].setdefault(key, val)

        # Chuẩn hóa cấu hình mappings để đảm bảo có đủ các flag mặc định
        for col_name, field_cfg in config["mappings"].items():
            field_cfg.setdefault("required", False)
            field_cfg.setdefault("type", "text")
            if field_cfg["type"].lower() == "checkbox":
                field_cfg.setdefault("separator", ";")

        return config


# ==============================================================================
# DATA NORMALIZER & LOCAL VALIDATOR
# ==============================================================================

class DataNormalizer:
    """Class xử lý làm sạch, validate và chuẩn hóa dữ liệu từ nguồn trước khi gửi."""

    @staticmethod
    def clean_cell_value(val: Any) -> str:
        """
        Chuẩn hóa dữ liệu ô đơn lẻ, chuyển float dạng .0 về int và loại bỏ khoảng trắng.
        
        Args:
            val (Any): Giá trị thô từ Pandas DataFrame.
            
        Returns:
            str: Chuỗi văn bản đã được làm sạch hoặc chuỗi rỗng.
        """
        if pd.isna(val):
            return ""
        
        # Trường hợp Pandas đọc số điện thoại hoặc mã ID thành float (ví dụ 123.0)
        if isinstance(val, float):
            if val.is_integer():
                return str(int(val))
            return str(val).strip()
            
        return str(val).strip()

    @classmethod
    def normalize_row(cls, row: pd.Series, mappings: Dict[str, Any]) -> List[Tuple[str, str]]:
        """
        Chuyển đổi và validate một dòng dữ liệu từ Excel thành cấu trúc payload HTTP POST.
        Kiểm tra nghiêm ngặt tính hợp lệ của trường Ngày tháng và các trường Bắt buộc.
        
        Args:
            row (pd.Series): Dòng dữ liệu hiện tại trong file Excel.
            mappings (Dict[str, Any]): Ánh xạ cột từ cấu hình config.json.
            
        Returns:
            List[Tuple[str, str]]: Danh sách các tuple (key, value) biểu diễn payload URL-encoded.
            
        Raises:
            MissingRequiredFieldValidationError: Nếu trường bắt buộc bị bỏ trống.
            DateValidationError: Nếu trường ngày tháng sai định dạng không thể parse.
        """
        payload: List[Tuple[str, str]] = []

        for col_name, field_cfg in mappings.items():
            entry_id = field_cfg["entry_id"]
            field_type = field_cfg.get("type", "text").lower()
            is_required = field_cfg.get("required", False)

            # Trường hợp cột cấu hình không tồn tại trong file Excel
            if col_name not in row:
                if is_required:
                    raise MissingRequiredFieldValidationError(
                        f"Cột bắt buộc '{col_name}' không tồn tại trong dữ liệu Excel."
                    )
                continue

            raw_val = row[col_name]
            cleaned_val = cls.clean_cell_value(raw_val)

            # Kiểm tra trường bắt buộc (Required Fields) ngay tại local
            if cleaned_val == "":
                if is_required:
                    raise MissingRequiredFieldValidationError(
                        f"Trường bắt buộc '{col_name}' bị trống ở dòng này."
                    )
                continue

            # Xử lý kiểu checkbox (chọn nhiều đáp án)
            if field_type == "checkbox":
                # Hỗ trợ phân tích đáp án được lưu dạng JSON list hoặc split theo dấu phân tách an toàn (mặc định ';')
                if cleaned_val.startswith("[") and cleaned_val.endswith("]"):
                    try:
                        values = json.loads(cleaned_val)
                        if not isinstance(values, list):
                            values = [values]
                    except json.JSONDecodeError:
                        # Fallback về split theo separator nếu parse JSON lỗi
                        separator = field_cfg.get("separator", ";")
                        values = [v.strip() for v in cleaned_val.split(separator) if v.strip()]
                else:
                    separator = field_cfg.get("separator", ";")
                    values = [v.strip() for v in cleaned_val.split(separator) if v.strip()]
                
                for val in values:
                    payload.append((entry_id, val))

            # Xử lý kiểu date (ngày tháng)
            elif field_type == "date":
                try:
                    # Parse nghiêm ngặt bằng pandas, ném exception nếu không nhận dạng được định dạng ngày
                    dt = pd.to_datetime(raw_val)
                    payload.append((f"{entry_id}_year", str(dt.year)))
                    payload.append((f"{entry_id}_month", f"{dt.month:02d}"))
                    payload.append((f"{entry_id}_day", f"{dt.day:02d}"))
                except Exception as e:
                    raise DateValidationError(
                        f"Lỗi định dạng ngày tháng tại cột '{col_name}' với giá trị '{raw_val}'. Chi tiết: {e}"
                    )

            # Xử lý các kiểu dữ liệu văn bản/lựa chọn đơn khác (text, radio, dropdown, email, phone)
            else:
                payload.append((entry_id, cleaned_val))

        return payload


# ==============================================================================
# FORM SUBMITTER WITH SESSION ROTATION & ROBUST PARSER
# ==============================================================================

class FormSubmitter:
    """Class quản lý kết nối HTTP, xoay vòng session ổn định và xác thực phản hồi."""

    def __init__(self, form_id: str, settings: Dict[str, Any]):
        self.form_id = form_id
        self.settings = settings
        self.endpoint = f"https://docs.google.com/forms/d/e/{form_id}/formResponse"
        
        # Thư mục lưu vết HTML phản hồi khi gặp lỗi để phục vụ debug
        self.error_logs_dir = "./logs/error_responses"
        
        # Trạng thái quản lý phiên (Session Lifecycle)
        self.session: Optional[requests.Session] = None
        self.current_ua: Optional[str] = None
        self.request_counter = 0
        
        # Khởi tạo Session đầu tiên
        self._rotate_session()

    def _rotate_session(self) -> None:
        """Đóng session cũ và thiết lập một session mới với User-Agent và bộ vân tay cố định."""
        if self.session:
            logger.info("Đang đóng phiên kết nối cũ để thực hiện xoay vòng Session...")
            try:
                self.session.close()
            except Exception:
                pass

        self.session = requests.Session()
        self.current_ua = random.choice(USER_AGENTS)
        self.request_counter = 0
        
        # Thiết lập cơ chế tự động thử lại ở tầng mạng
        retries = Retry(
            total=self.settings["max_retries"],
            backoff_factor=self.settings["retry_backoff"],
            status_forcelist=[429, 500, 502, 503, 504],
            raise_on_status=False
        )
        self.session.mount("https://", HTTPAdapter(max_retries=retries))
        logger.info(f"Đã thiết lập Session mới. Cố định User-Agent: {self.current_ua}")

    def _get_headers(self) -> Dict[str, str]:
        """
        Tạo bộ headers giả lập trình duyệt Chromium đồng bộ với User-Agent hiện tại.
        
        Returns:
            Dict[str, str]: Bộ HTTP Headers chuẩn hóa chống chặn.
        """
        # Xác định nền tảng (Platform) tương ứng với User-Agent để đồng bộ Client Hint
        ua_platform = '"Windows"'
        if "Macintosh" in self.current_ua:
            ua_platform = '"macOS"'
            
        return {
            "User-Agent": self.current_ua,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
            "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7",
            "Content-Type": "application/x-www-form-urlencoded",
            "Origin": "https://docs.google.com",
            "Referer": f"https://docs.google.com/forms/d/e/{self.form_id}/viewform",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "same-origin",
            "Sec-Fetch-User": "?1",
            "sec-ch-ua": f'"Chromium";v="122", "Not(A:Brand";v="24", "Google Chrome";v="122"',
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": ua_platform
        }

    def _save_error_html(self, html_content: str, line_num: int) -> str:
        """
        Lưu mã nguồn HTML phản hồi khi lỗi vào thư mục cục bộ để phục vụ debug.
        
        Args:
            html_content (str): Nội dung HTML phản hồi từ Google.
            line_num (int): Vị trí số dòng dữ liệu trong Excel.
            
        Returns:
            str: Đường dẫn file HTML đã được lưu.
        """
        if not os.path.exists(self.error_logs_dir):
            os.makedirs(self.error_logs_dir, exist_ok=True)
            
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"form_error_{timestamp}_line_{line_num}.html"
        filepath = os.path.join(self.error_logs_dir, filename)
        
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(html_content)
            
        return filepath

    def submit(self, payload: List[Tuple[str, str]], line_num: int) -> Tuple[bool, str]:
        """
        Thực hiện gửi POST Request lên Google Forms và xác thực phản hồi thành công kép.
        
        Args:
            payload (List[Tuple[str, str]]): Payload dữ liệu đã được định dạng.
            line_num (int): Số thứ tự dòng Excel phục vụ debug.
            
        Returns:
            Tuple[bool, str]: Trạng thái (thành công/thất bại) và thông báo chi tiết.
        """
        # Kiểm tra giới hạn vòng xoay Session (Session Rotation Limit)
        rotation_limit = self.settings.get("session_rotation_limit", 10)
        if self.request_counter >= rotation_limit:
            logger.info(f"Đạt giới hạn gửi của Session ({rotation_limit} requests). Tiến hành xoay vòng.")
            self._rotate_session()

        headers = self._get_headers()
        self.request_counter += 1

        try:
            response = self.session.post(
                url=self.endpoint,
                data=payload,
                headers=headers,
                timeout=self.settings["timeout"]
            )
            
            # Phân tích phản hồi thông minh (Robust Response Parser)
            # Tránh lỗi False Positive khi Google Forms trả về trang biểu mẫu lỗi nhưng status_code vẫn là 200
            html = response.text
            
            # Bộ từ khóa lỗi nghiệp vụ và cấu trúc lỗi DOM của Google Forms
            error_keywords = [
                "Mục này là bắt buộc", "This is a required question",
                "Trường này không hợp lệ", "Invalid entry",
                "Giá trị phải là", "Must be",
                "Có lỗi xảy ra", "An error occurred",
                "freebirdFormviewerViewResponseError", "hasError", "error-message"
            ]
            
            # Lấy danh sách từ khóa thành công từ cấu hình hoặc fallback mặc định
            success_keywords = self.settings.get("success_keywords", [
                "Your response has been recorded",
                "Câu trả lời của bạn đã được ghi lại"
            ])
            if not isinstance(success_keywords, list):
                success_keywords = [str(success_keywords)]

            # So khớp không phân biệt hoa thường để tăng độ chính xác
            html_lower = html.lower()
            is_recorded = any(str(k).lower() in html_lower for k in success_keywords)
            has_error_flag = any(k in html for k in error_keywords)

            # Điều kiện thành công kép: HTTP 200 OK + Có từ khóa thành công + Không có từ khóa lỗi
            if response.status_code == 200:
                if is_recorded and not has_error_flag:
                    return True, "Gửi thành công (Phản hồi được ghi nhận)"
                elif has_error_flag:
                    err_file = self._save_error_html(html, line_num)
                    return False, f"Google Forms trả về lỗi Validation (HTML lỗi lưu tại: {err_file})"
                else:
                    # Fallback kiểm tra từ khóa formResponse chung nếu Google Forms đổi mẫu cấu trúc tiếng khác
                    if "formResponse" in html and not has_error_flag:
                        return True, "Gửi thành công (Phân tích fallback)"
                    
                    err_file = self._save_error_html(html, line_num)
                    return False, f"Không phát hiện dấu hiệu thành công (HTML lỗi lưu tại: {err_file})"
            else:
                # Xoay session ngay lập tức nếu dính mã lỗi từ chối của Google (ví dụ 429) để reset IP/Session
                if response.status_code == 429:
                    logger.warning("Dính lỗi 429 Too Many Requests. Tiến hành đổi Session lập tức.")
                    self._rotate_session()
                err_file = self._save_error_html(html, line_num)
                return False, f"HTTP Error Status Code: {response.status_code} (HTML lỗi lưu tại: {err_file})"

        except requests.exceptions.RequestException as e:
            # Lỗi mạng vật lý, tiến hành xoay session để chuẩn bị cho request tiếp theo
            logger.error(f"Lỗi kết nối mạng tại dòng #{line_num}: {e}. Xoay vòng Session.")
            self._rotate_session()
            return False, f"Lỗi kết nối mạng: {str(e)}"


# ==============================================================================
# EXECUTION CONTROLLER WITH CHECKPOINTING & RESUME STATE
# ==============================================================================

class ExecutionController:
    """Class điều phối toàn bộ luồng công việc: Checkpoint, Resume, Run, Export."""

    def __init__(self, data_path: str, config_path: str, output_path: str):
        self.data_path = data_path
        self.config_path = config_path
        self.output_path = output_path
        self.checkpoint_path = "checkpoint_log.csv"
        
        # Lưu các chỉ số dòng đã gửi thành công để bỏ qua trong chế độ Resume
        self.success_row_indices: Set[int] = set()
        self.is_resume_mode = False

    def _load_checkpoint(self) -> None:
        """Đọc file checkpoint và hỏi người dùng có muốn Resume từ vị trí gián đoạn không."""
        if not os.path.exists(self.checkpoint_path):
            return

        try:
            checkpoint_df = pd.read_csv(self.checkpoint_path)
            if checkpoint_df.empty:
                return

            # Xác định các dòng đã gửi thành công trước đó
            success_rows = checkpoint_df[checkpoint_df["Status"] == "Success"]
            self.success_row_indices = set(success_rows["Row_Index"].tolist())

            if not self.success_row_indices:
                return

            print(f"\n[!] CẢNH BÁO: Phát hiện file checkpoint `{self.checkpoint_path}`.")
            print(f"    Tìm thấy {len(self.success_row_indices)} dòng dữ liệu đã được gửi thành công trước đó.")
            choice = input("    Bạn có muốn tiếp tục (Resume) chạy tiếp từ vị trí bị gián đoạn không? [Y/n]: ").strip().lower()

            if choice in ["y", "yes", ""]:
                self.is_resume_mode = True
                logger.info("Chế độ phục hồi lỗi (Resume) được kích hoạt.")
            else:
                # Xóa file checkpoint cũ nếu chọn chạy lại từ đầu
                os.remove(self.checkpoint_path)
                logger.info("Đã xóa file checkpoint cũ. Bắt đầu quy trình chạy mới hoàn toàn.")
        except Exception as e:
            logger.error(f"Không thể đọc file checkpoint: {e}. Tiến hành chạy mới hoàn toàn.")

    def _append_to_checkpoint(self, row_idx: int, status: str, message: str, identifier: str) -> None:
        """
        Ghi append kết quả dòng hiện tại lập tức vào file checkpoint_log.csv.
        
        Args:
            row_idx (int): Chỉ số dòng Excel (0-indexed).
            status (str): Trạng thái gửi (Success/Failed).
            message (str): Thông điệp phản hồi hoặc mô tả lỗi.
            identifier (str): Giá trị định danh của dòng (ví dụ cột Họ tên hoặc cột đầu tiên).
        """
        file_exists = os.path.exists(self.checkpoint_path)
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Tạo dataframe cho bản ghi hiện tại
        record = pd.DataFrame([{
            "Row_Index": row_idx,
            "Identifier": identifier,
            "Status": status,
            "Message": message,
            "Timestamp": timestamp
        }])

        try:
            # Ghi append trực tiếp
            record.to_csv(
                self.checkpoint_path, 
                mode="a", 
                header=not file_exists, 
                index=False, 
                encoding="utf-8"
            )
        except Exception as e:
            logger.error(f"Lỗi ghi nhận checkpoint tại dòng index {row_idx}: {e}")

    def _calculate_jitter_delay(self, min_delay: float, max_delay: float) -> float:
        """
        Tính toán khoảng trễ ngẫu nhiên phi tuyến tính bằng phân phối log-normal kết hợp.
        Mô phỏng hành vi ngẫu nhiên có độ dài nghỉ bất thường của con người.
        
        Args:
            min_delay (float): Khoảng trễ tối thiểu cấu hình.
            max_delay (float): Khoảng trễ tối đa cấu hình.
            
        Returns:
            float: Khoảng thời gian trễ tính bằng giây.
        """
        base_delay = random.uniform(min_delay, max_delay)
        
        # Thêm phân phối Log-Normal để thỉnh thoảng tạo ra đợt nghỉ dài ngẫu nhiên
        # mu=0, sigma=0.5 tạo ra giá trị nhỏ trung bình nhưng thỉnh thoảng có giá trị lớn đột biến
        extra_jitter = random.lognormvariate(mu=0, sigma=0.5) - 1.0
        extra_jitter = max(0.0, extra_jitter)
        
        return base_delay + extra_jitter

    def _export_final_report(self, df: pd.DataFrame) -> None:
        """
        Đối chiếu file checkpoint và xuất báo cáo Excel kết quả cuối cùng.
        
        Args:
            df (pd.DataFrame): DataFrame dữ liệu gốc.
        """
        if not os.path.exists(self.checkpoint_path):
            logger.error("Không tìm thấy file checkpoint để xuất báo cáo cuối cùng.")
            return

        try:
            checkpoint_df = pd.read_csv(self.checkpoint_path)
            
            # Lọc bản ghi checkpoint cuối cùng của mỗi Row_Index (nếu chạy đè nhiều lần)
            checkpoint_df = checkpoint_df.drop_duplicates(subset=["Row_Index"], keep="last")
            
            # Tạo map từ Row_Index sang kết quả
            status_map = dict(zip(checkpoint_df["Row_Index"], checkpoint_df["Status"]))
            msg_map = dict(zip(checkpoint_df["Row_Index"], checkpoint_df["Message"]))

            # Gán kết quả vào DataFrame gốc
            df["Submission_Status"] = [status_map.get(i, "Not Processed") for i in range(len(df))]
            df["Submission_Error_Detail"] = [msg_map.get(i, "") for i in range(len(df))]

            # Ghi file Excel báo cáo
            df.to_excel(self.output_path, index=False, engine="openpyxl")
            logger.info(f"Đã xuất báo cáo kết quả gửi dữ liệu thành công: {self.output_path}")
            
            # Thống kê kết quả
            success_count = list(status_map.values()).count("Success")
            failed_count = list(status_map.values()).count("Failed")
            logger.info("=== BÁO CÁO THỐNG KÊ ===")
            logger.info(f"Tổng số bản ghi xử lý: {len(status_map)}")
            logger.info(f"Gửi thành công: {success_count}")
            logger.info(f"Thất bại (Gồm lỗi logic dữ liệu & lỗi mạng): {failed_count}")
            
        except Exception as e:
            logger.error(f"Lỗi trong quá trình tổng hợp xuất báo cáo Excel: {e}")

    def run(self) -> None:
        """Khởi chạy quy trình điều phối và gửi biểu mẫu hàng loạt."""
        logger.info("=== BẮT ĐẦU QUY TRÌNH GỬI DỮ LIỆU TỰ ĐỘNG KHÁNG LỖI ===")
        
        # 1. Đọc và kiểm tra file cấu hình
        try:
            config = ConfigLoader.load(self.config_path)
            logger.info("Nạp cấu hình config.json thành công.")
        except Exception as e:
            logger.error(f"Lỗi nạp cấu hình: {e}")
            sys.exit(1)

        # 2. Đọc file dữ liệu Excel/CSV
        try:
            if self.data_path.endswith(".csv"):
                df = pd.read_csv(self.data_path)
            else:
                df = pd.read_excel(self.data_path)
            logger.info(f"Đọc dữ liệu thành công: {self.data_path}. Tổng số dòng: {len(df)}")
        except Exception as e:
            logger.error(f"Lỗi đọc file dữ liệu đầu vào: {e}")
            sys.exit(1)

        # Kiểm tra file checkpoint và cấu hình trạng thái Resume
        self._load_checkpoint()

        # 3. Khởi tạo Submitter
        submitter = FormSubmitter(config["form_id"], config["settings"])
        
        min_delay = config["settings"]["min_delay"]
        max_delay = config["settings"]["max_delay"]

        logger.info("Đang thực hiện gửi dữ liệu...")
        
        # Duyệt qua các dòng bằng tqdm hiển thị progress bar
        for index, row in tqdm(df.iterrows(), total=len(df), desc="Tiến trình", unit="dòng"):
            row_num = index + 1
            
            # Lấy định danh dòng (mặc định lấy cột đầu tiên để ghi log checkpoint cho dễ đọc)
            row_identifier = str(row.iloc[0]) if len(row) > 0 else f"Row_{row_num}"

            # Nếu ở chế độ Resume và dòng này đã gửi thành công trước đó thì bỏ qua
            if self.is_resume_mode and index in self.success_row_indices:
                continue

            # Bước A: Chuẩn hóa và Validate dữ liệu tại local
            try:
                payload = DataNormalizer.normalize_row(row, config["mappings"])
            except ValidationError as e:
                # Đánh dấu lỗi logic dữ liệu và ghi nhận checkpoint ngay lập tức không gửi request rác
                error_msg = f"Lỗi Logic Dữ Liệu Local: {str(e)}"
                logger.warning(f"Dòng #{row_num} [{row_identifier}] bị skip: {error_msg}")
                self._append_to_checkpoint(index, "Failed", error_msg, row_identifier)
                continue
            except Exception as e:
                error_msg = f"Lỗi xử lý không xác định tại local: {str(e)}"
                logger.error(f"Dòng #{row_num} [{row_identifier}] bị skip: {error_msg}")
                self._append_to_checkpoint(index, "Failed", error_msg, row_identifier)
                continue

            # Bước B: Gửi dữ liệu qua HTTP POST
            success, msg = submitter.submit(payload, row_num)
            
            # Ghi nhận trạng thái lập tức vào checkpoint_log.csv
            status_str = "Success" if success else "Failed"
            self._append_to_checkpoint(index, status_str, msg, row_identifier)
            
            if not success:
                logger.warning(f"Dòng #{row_num} [{row_identifier}] gửi thất bại. Chi tiết: {msg}")

            # Tạo độ trễ ngẫu nhiên Jitter phi tuyến tính tránh bị phát hiện bot
            if index < len(df) - 1:
                delay = self._calculate_jitter_delay(min_delay, max_delay)
                time.sleep(delay)

        # 4. Quy trình tổng hợp xuất báo cáo Excel cuối cùng từ checkpoint
        logger.info("Đang tiến hành tổng hợp báo cáo Excel cuối cùng...")
        self._export_final_report(df)
        logger.info("=== QUY TRÌNH HOÀN THÀNH ===")


# ==============================================================================
# MAIN ENTRYPOINT
# ==============================================================================

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Google Forms Direct HTTP Auto Filler Tool - Resilient CLI Version.",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument(
        "-i", "--input", 
        required=False,
        default="sample_data.xlsx",
        help="Đường dẫn tới file Excel/CSV chứa dữ liệu."
    )
    parser.add_argument(
        "-c", "--config", 
        required=False,
        default="config.json",
        help="Đường dẫn tới file cấu hình config.json."
    )
    parser.add_argument(
        "-o", "--output", 
        required=False,
        default="result_log.xlsx",
        help="Đường dẫn lưu kết quả báo cáo Excel."
    )
    
    args = parser.parse_args()
    
    # Kiểm tra sự tồn tại của file dữ liệu nguồn
    if not os.path.exists(args.input):
        logger.error(f"Không tìm thấy file dữ liệu đầu vào: {args.input}")
        sys.exit(1)
        
    controller = ExecutionController(
        data_path=args.input, 
        config_path=args.config, 
        output_path=args.output
    )
    controller.run()


if __name__ == "__main__":
    main()
