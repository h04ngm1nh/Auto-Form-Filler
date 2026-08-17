#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Google Forms Automated Form Filler (CLI Script) - Ponytail Refactored Version
Author: Antigravity (under Ponytail Senior Developer guidelines)
Description:
    Kịch bản tự động hóa gửi dữ liệu hàng loạt từ file Excel/CSV vào Google Forms
    thông qua HTTP POST Request trực tiếp. Codebase tối giản, không boilerplate dư thừa.
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
import pandas as pd
import requests

# Cấu hình lại mã hóa UTF-8 cho console Windows tránh UnicodeEncodeError
if sys.platform.startswith("win"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
        sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    except AttributeError:
        pass

# Thiết lập log console
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger("GoogleFormFiller")

# Modern Desktop User-Agents
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:123.0) Gecko/20100101 Firefox/123.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.3 Safari/605.1.15",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36 Edg/122.0.0.0"
]

def load_config(config_path):
    """Đọc và chuẩn hóa cấu hình config.json"""
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Không tìm thấy file cấu hình: {config_path}")

    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)

    # Chuẩn hóa cấu trúc pre-filled URL sang cấu trúc chuẩn của ứng dụng
    if "form_url" in config and "field_mappings" in config:
        url = config["form_url"]
        match = re.search(r'/d/e/([^/]+)/', url)
        config["form_id"] = match.group(1) if match else (url.split('/')[-2] if '/' in url else url)
        config["mappings"] = config["field_mappings"]
        if "success_keywords" in config:
            config.setdefault("settings", {})["success_keywords"] = config["success_keywords"]

    if "form_id" not in config or not config["form_id"]:
        raise ValueError("Thiếu 'form_id' trong file cấu hình.")
    if "mappings" not in config or not config["mappings"]:
        raise ValueError("Thiếu 'mappings' trong file cấu hình.")

    settings = config.setdefault("settings", {})
    defaults = {
        "min_delay": 2.0, "max_delay": 5.0, "max_retries": 3,
        "retry_backoff": 2.0, "timeout": 10, "session_rotation_limit": 10,
        "success_keywords": ["Your response has been recorded", "Câu trả lời của bạn đã được ghi lại"]
    }
    for k, v in defaults.items():
        settings.setdefault(k, v)

    for key, field_cfg in config["mappings"].items():
        if isinstance(field_cfg, dict):
            field_cfg.setdefault("required", False)
            field_cfg.setdefault("type", "text")
            if field_cfg["type"].lower() == "checkbox":
                field_cfg.setdefault("separator", ";")
    return config

def clean_cell_value(val):
    """Chuẩn hóa giá trị từ ô dữ liệu Excel/CSV"""
    if pd.isna(val):
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).strip()

def normalize_row(row, mappings):
    """Validate dữ liệu nguồn và chuyển đổi thành payload gửi HTTP POST"""
    payload = []
    for key, cfg in mappings.items():
        if key.startswith("entry."):
            entry_id = key
            col_name = cfg.get("column_name") if isinstance(cfg, dict) else cfg
        else:
            col_name = key
            entry_id = cfg.get("entry_id") if isinstance(cfg, dict) else ""

        is_required = cfg.get("required", False) if isinstance(cfg, dict) else False
        field_type = cfg.get("type", "text").lower() if isinstance(cfg, dict) else "text"

        if col_name not in row:
            if is_required:
                raise ValueError(f"Cột bắt buộc '{col_name}' không tồn tại trong dữ liệu.")
            continue

        val = clean_cell_value(row[col_name])
        if not val:
            if is_required:
                raise ValueError(f"Trường bắt buộc '{col_name}' bị trống.")
            continue

        if field_type == "checkbox":
            separator = cfg.get("separator", ";") if isinstance(cfg, dict) else ";"
            if val.startswith("[") and val.endswith("]"):
                try:
                    vals = json.loads(val)
                    if not isinstance(vals, list):
                        vals = [vals]
                except json.JSONDecodeError:
                    vals = [v.strip() for v in val.split(separator) if v.strip()]
            else:
                vals = [v.strip() for v in val.split(separator) if v.strip()]
            for v in vals:
                payload.append((entry_id, v))
        elif field_type == "date":
            try:
                dt = pd.to_datetime(row[col_name])
                payload.append((f"{entry_id}_year", str(dt.year)))
                payload.append((f"{entry_id}_month", f"{dt.month:02d}"))
                payload.append((f"{entry_id}_day", f"{dt.day:02d}"))
            except Exception as e:
                raise ValueError(f"Ngày tháng '{row[col_name]}' tại cột '{col_name}' không thể parse: {e}")
        else:
            payload.append((entry_id, val))
    return payload

class FormSubmitter:
    """Quản lý HTTP session, xoay vòng User-Agent và gửi dữ liệu biểu mẫu"""
    def __init__(self, form_id, settings):
        self.form_id = form_id
        self.settings = settings
        self.endpoint = f"https://docs.google.com/forms/d/e/{form_id}/formResponse"
        self.error_logs_dir = "./logs/error_responses"
        self.session = None
        self.current_ua = None
        self.request_counter = 0
        self._rotate_session()

    def _rotate_session(self):
        if self.session:
            try:
                self.session.close()
            except Exception:
                pass
        self.session = requests.Session()
        self.current_ua = random.choice(USER_AGENTS)
        self.request_counter = 0
        logger.info(f"Đã xoay vòng Session mới. User-Agent: {self.current_ua}")

    def _get_headers(self):
        return {
            "User-Agent": self.current_ua,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7",
            "Content-Type": "application/x-www-form-urlencoded",
            "Referer": f"https://docs.google.com/forms/d/e/{self.form_id}/viewform",
        }

    def _save_error_html(self, html, line_num):
        os.makedirs(self.error_logs_dir, exist_ok=True)
        filepath = os.path.join(self.error_logs_dir, f"form_error_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}_line_{line_num}.html")
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(html)
        return filepath

    def submit(self, payload, line_num):
        if self.request_counter >= self.settings.get("session_rotation_limit", 10):
            logger.info("Xoay vòng Session do vượt giới hạn requests.")
            self._rotate_session()

        headers = self._get_headers()
        self.request_counter += 1
        logger.info(f"Dòng #{line_num} - Gửi payload: {dict(payload)}")

        max_retries = self.settings.get("max_retries", 3)
        retry_backoff = self.settings.get("retry_backoff", 2.0)
        timeout = self.settings.get("timeout", 10)

        for attempt in range(max_retries):
            try:
                response = self.session.post(self.endpoint, data=payload, headers=headers, timeout=timeout)
                html = response.text
                html_lower = html.lower()

                error_keywords = ["Mục này là bắt buộc", "This is a required question", "Trường này không hợp lệ", "Invalid entry", "freebirdFormviewerViewResponseError", "hasError", "error-message"]
                success_keywords = self.settings.get("success_keywords", ["Your response has been recorded", "Câu trả lời của bạn đã được ghi lại"])

                is_success = any(str(k).lower() in html_lower for k in success_keywords) or ("formResponse" in html and not any(k in html for k in error_keywords))
                has_error = any(k in html for k in error_keywords)

                if response.status_code == 200:
                    if is_success and not has_error:
                        return True, "Gửi thành công"
                    err_file = self._save_error_html(html, line_num)
                    return False, f"Google Forms trả về lỗi Validation (HTML lỗi tại: {err_file})"
                else:
                    if response.status_code == 429:
                        self._rotate_session()
                    err_file = self._save_error_html(html, line_num)
                    return False, f"Lỗi HTTP {response.status_code} (HTML lỗi tại: {err_file})"
            except requests.RequestException as e:
                if attempt == max_retries - 1:
                    self._rotate_session()
                    return False, f"Lỗi kết nối mạng: {e}"
                time.sleep(retry_backoff * (attempt + 1))

def record_checkpoint(filepath, row_idx, identifier, status, message):
    """Ghi nhận tức thời kết quả của dòng vào file checkpoint CSV"""
    file_exists = os.path.exists(filepath)
    record = pd.DataFrame([{
        "Row_Index": row_idx,
        "Identifier": identifier,
        "Status": status,
        "Message": message,
        "Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }])
    try:
        record.to_csv(filepath, mode="a", header=not file_exists, index=False, encoding="utf-8")
    except Exception as e:
        logger.error(f"Lỗi ghi nhận checkpoint: {e}")

def bootstrap_from_prefilled_url(url, output_excel, output_config):
    """Khởi tạo file cấu hình config.json và Excel mẫu từ link Pre-filled"""
    import urllib.parse
    logger.info("Đang bắt đầu phân tích Link Pre-filled...")
    try:
        parsed = urllib.parse.urlparse(url)
        path_parts = parsed.path.split('/')
        form_id = path_parts[path_parts.index('e') + 1] if 'e' in path_parts else re.search(r'/d/e/([^/]+)/', parsed.path).group(1)

        query_params = urllib.parse.parse_qs(parsed.query)
        mappings = {}
        headers = []
        dummy1, dummy2 = {}, {}

        for k, v in query_params.items():
            if k.startswith("entry."):
                raw_val = v[0] if v else ""
                col_name = raw_val.strip() if raw_val.strip() else f"Cột_{k}"

                original_col = col_name
                counter = 1
                while col_name in headers:
                    col_name = f"{original_col}_{counter}"
                    counter += 1

                mappings[k] = col_name
                headers.append(col_name)
                dummy1[col_name] = raw_val

                col_lower = col_name.lower()
                if "email" in col_lower:
                    dummy2[col_name] = "nguyenvana@example.com"
                elif any(x in col_lower for x in ["thoại", "sđt", "phone"]):
                    dummy2[col_name] = "0912345678"
                elif any(x in col_lower for x in ["ngày", "date", "sinh"]):
                    dummy2[col_name] = "1998-05-20"
                else:
                    dummy2[col_name] = f"Mẫu_{col_name}"

        if not mappings:
            raise ValueError("Không tìm thấy tham số pre-filled 'entry.xxxx'.")

        pd.DataFrame([dummy1, dummy2]).to_excel(output_excel, index=False, engine="openpyxl")
        with open(output_config, "w", encoding="utf-8") as f:
            json.dump({"form_url": f"https://docs.google.com/forms/d/e/{form_id}/formResponse", "field_mappings": mappings}, f, ensure_ascii=False, indent=2)

        logger.info(f"Khởi tạo thành công:\n1. Excel mẫu: {output_excel}\n2. Config mẫu: {output_config}")
    except Exception as e:
        logger.error(f"Bootstrap thất bại: {e}")
        sys.exit(1)

def main():
    parser = argparse.ArgumentParser(description="Google Forms Direct HTTP Auto Filler Tool - Ponytail Edition.")
    parser.add_argument("-i", "--input", default="data/input/Data_to_input.xlsx", help="Đường dẫn file dữ liệu Excel/CSV.")
    parser.add_argument("-c", "--config", default="config.json", help="Đường dẫn file cấu hình config.json.")
    parser.add_argument("-o", "--output", default="data/output/result_log.xlsx", help="Đường dẫn file báo cáo Excel đầu ra.")
    parser.add_argument("--init-from-url", default=None, help="Khởi tạo từ URL pre-filled.")
    parser.add_argument("--output-excel", default="data_template.xlsx", help="Đường dẫn file Excel mẫu đầu ra.")
    parser.add_argument("--output-config", default="config.json", help="Đường dẫn file config mẫu đầu ra.")

    args = parser.parse_args()

    if args.init_from_url:
        bootstrap_from_prefilled_url(args.init_from_url, args.output_excel, args.output_config)
        sys.exit(0)

    if not os.path.exists(args.input):
        logger.error(f"Không tìm thấy file dữ liệu đầu vào: {args.input}")
        sys.exit(1)

    # 1. Nạp cấu hình
    try:
        config = load_config(args.config)
        logger.info("Nạp file cấu hình config.json thành công.")
    except Exception as e:
        logger.error(f"Lỗi nạp cấu hình: {e}")
        sys.exit(1)

    # 2. Đọc file dữ liệu đầu vào
    df = None
    while df is None:
        try:
            df = pd.read_csv(args.input) if args.input.endswith(".csv") else pd.read_excel(args.input)
            logger.info(f"Đọc dữ liệu thành công: {args.input}. Tổng số dòng: {len(df)}")
        except PermissionError:
            print(f"\n[!] CẢNH BÁO: File {args.input} đang bị khóa hoặc được mở bởi một chương trình khác.")
            input("    Vui lòng đóng file lại và nhấn Enter để thử lại...")
        except Exception as e:
            logger.error(f"Lỗi đọc file dữ liệu: {e}")
            sys.exit(1)

    # 3. Phục hồi checkpoint
    checkpoint_path = "logs/checkpoint_log.csv"
    success_row_indices = set()
    is_resume_mode = False

    if os.path.exists(checkpoint_path):
        try:
            checkpoint_df = pd.read_csv(checkpoint_path)
            if not checkpoint_df.empty:
                success_rows = checkpoint_df[checkpoint_df["Status"] == "Success"]
                success_row_indices = set(success_rows["Row_Index"].tolist())

                if success_row_indices:
                    print(f"\n[!] CẢNH BÁO: Phát hiện checkpoint `{checkpoint_path}` với {len(success_row_indices)} dòng đã hoàn thành.")
                    choice = input("    Bạn có muốn Resume chạy tiếp từ vị trí gián đoạn không? [Y/n]: ").strip().lower()
                    if choice in ["y", "yes", ""]:
                        is_resume_mode = True
                        logger.info("Kích hoạt chế độ Resume.")
                    else:
                        os.remove(checkpoint_path)
                        logger.info("Bắt đầu chạy mới hoàn toàn.")
        except Exception as e:
            logger.error(f"Không thể đọc checkpoint: {e}. Tiến hành chạy mới hoàn toàn.")

    # 4. Gửi dữ liệu biểu mẫu
    submitter = FormSubmitter(config["form_id"], config["settings"])
    min_delay = config["settings"]["min_delay"]
    max_delay = config["settings"]["max_delay"]

    from tqdm import tqdm
    for index, row in tqdm(df.iterrows(), total=len(df), desc="Tiến trình", unit="dòng"):
        row_num = index + 1
        row_identifier = str(row.iloc[0]) if len(row) > 0 else f"Row_{row_num}"

        if is_resume_mode and index in success_row_indices:
            continue

        # A. Validate offline tại local
        try:
            payload = normalize_row(row, config["mappings"])
        except Exception as e:
            error_msg = f"Lỗi Validate Local: {e}"
            logger.warning(f"Dòng #{row_num} [{row_identifier}] bị skip: {error_msg}")
            record_checkpoint(checkpoint_path, index, row_identifier, "Failed", error_msg)
            continue

        # B. Gửi HTTP POST request
        success, msg = submitter.submit(payload, row_num)
        status_str = "Success" if success else "Failed"
        record_checkpoint(checkpoint_path, index, row_identifier, status_str, msg)

        if not success:
            logger.warning(f"Dòng #{row_num} [{row_identifier}] gửi thất bại: {msg}")

        # C. Delay Jitter phi tuyến tính
        if index < len(df) - 1:
            delay = random.uniform(min_delay, max_delay) + max(0.0, random.lognormvariate(0, 0.5) - 1.0)
            time.sleep(delay)

    # 5. Xuất báo cáo Excel kết quả cuối cùng từ checkpoint
    logger.info("Đang tiến hành tổng hợp báo cáo Excel...")
    if os.path.exists(checkpoint_path):
        try:
            checkpoint_df = pd.read_csv(checkpoint_path).drop_duplicates(subset=["Row_Index"], keep="last")
            status_map = dict(zip(checkpoint_df["Row_Index"], checkpoint_df["Status"]))
            msg_map = dict(zip(checkpoint_df["Row_Index"], checkpoint_df["Message"]))

            df["Submission_Status"] = [status_map.get(i, "Not Processed") for i in range(len(df))]
            df["Submission_Error_Detail"] = [msg_map.get(i, "") for i in range(len(df))]

            write_success = False
            current_output = args.output
            while not write_success:
                try:
                    df.to_excel(current_output, index=False, engine="openpyxl")
                    write_success = True
                    logger.info(f"Báo cáo được xuất thành công tại: {current_output}")
                except PermissionError:
                    print(f"\n[!] CẢNH BÁO: Không thể ghi file báo cáo vào: {current_output} (file đang mở).")
                    choice = input("    Vui lòng đóng file Excel lại và nhấn Enter để thử lại, hoặc nhập 'save' để lưu thành file phụ: ").strip().lower()
                    if choice == "save":
                        base, ext = os.path.splitext(args.output)
                        current_output = f"{base}_copy_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"
                except Exception as e:
                    logger.error(f"Lỗi khi ghi báo cáo Excel: {e}")
                    break

            success_count = list(status_map.values()).count("Success")
            failed_count = list(status_map.values()).count("Failed")
            logger.info("=== BÁO CÁO THỐNG KÊ ===")
            logger.info(f"Tổng số bản ghi: {len(status_map)}")
            logger.info(f"Thành công: {success_count}")
            logger.info(f"Thất bại: {failed_count}")
        except Exception as e:
            logger.error(f"Lỗi trong quá trình tổng hợp báo cáo: {e}")
    logger.info("=== QUY TRÌNH HOÀN THÀNH ===")

if __name__ == "__main__":
    main()
