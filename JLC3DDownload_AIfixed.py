import configparser
import json
import os
import re
import subprocess
import sys
import threading
import tkinter as tk
import webbrowser
from datetime import datetime
from tkinter import filedialog, messagebox, scrolledtext, ttk
from typing import Any, Iterable, Optional

import requests


# =====================================================
# 配置
# =====================================================
APP_DIR = os.path.join(os.path.expanduser("~"), ".jlc3d")
CONFIG_FILE = os.path.join(APP_DIR, "config.ini")


def ensure_app_dir() -> None:
    os.makedirs(APP_DIR, exist_ok=True)


def save_download_path(path: str) -> None:
    ensure_app_dir()
    cfg = configparser.ConfigParser()
    cfg["PATH"] = {"DownloadPath": path}
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        cfg.write(f)


def load_download_path() -> Optional[str]:
    if not os.path.exists(CONFIG_FILE):
        return None
    cfg = configparser.ConfigParser()
    cfg.read(CONFIG_FILE, encoding="utf-8")
    return cfg.get("PATH", "DownloadPath", fallback=None)


def default_desktop() -> str:
    return os.path.join(os.path.expanduser("~"), "Desktop")


# =====================================================
# API
# =====================================================
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/javascript, */*; q=0.01",
}

EXACT_COMPONENT_API = "https://easyeda.com/api/products/{lcsc_id}/components?version=6.4.19.5"
BASE_API = "https://pro.lceda.cn/api"
STEP_URLS = [
    "https://modules.easyeda.com/qAxj6KHrDKw4blvCG8QJPs7Y/{uuid}",
    "https://modules.lceda.cn/qAxj6KHrDKw4blvCG8QJPs7Y/{uuid}",
]

SESSION = requests.Session()
SESSION.headers.update(HEADERS)


class JLC3DError(Exception):
    pass


def normalize_input(text: str) -> str:
    return text.strip()


def normalize_code(code: str) -> str:
    return normalize_input(code).upper()


def normalize_match_key(text: str) -> str:
    text = normalize_input(text).upper()
    return re.sub(r"[^A-Z0-9]+", "", text)


def looks_like_lcsc_code(text: str) -> bool:
    text = normalize_input(text)
    if not text:
        return False
    return text.upper().startswith("C") and text[1:].isdigit()


def _clean_display_name(value: Any) -> Optional[str]:
    if not isinstance(value, str):
        return None
    value = value.strip()
    return value or None


def _iter_candidate_names(data: dict[str, Any]) -> Iterable[str]:
    direct_keys = (
        "productModel",
        "product_model",
        "componentModel",
        "component_model",
        "model",
        "title",
        "displayTitle",
        "display_title",
        "productTitle",
        "product_title",
        "productName",
        "product_name",
        "componentName",
        "component_name",
        "name",
        "productDesc",
        "product_desc",
    )
    nested_dict_keys = (
        "productInfo",
        "product_info",
        "basicInfo",
        "basic_info",
        "componentInfo",
        "component_info",
    )

    for key in direct_keys:
        value = _clean_display_name(data.get(key))
        if value:
            yield value

    for key in nested_dict_keys:
        nested = data.get(key)
        if not isinstance(nested, dict):
            continue
        for nested_key in direct_keys:
            value = _clean_display_name(nested.get(nested_key))
            if value:
                yield value


def extract_component_display_name(component_data: dict[str, Any], fallback_code: str) -> str:
    candidates = list(dict.fromkeys(_iter_candidate_names(component_data)))
    code_upper = normalize_code(fallback_code)

    for name in candidates:
        if normalize_code(name) != code_upper:
            return name

    if candidates:
        return candidates[0]
    return fallback_code


def extract_component_code(data: dict[str, Any]) -> Optional[str]:
    keys = (
        "componentCode",
        "productCode",
        "number",
        "code",
        "lcsc",
        "lcscNumber",
        "component_id",
    )
    for key in keys:
        value = data.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()

    nested_dict_keys = (
        "productInfo",
        "product_info",
        "basicInfo",
        "basic_info",
        "componentInfo",
        "component_info",
    )
    for dict_key in nested_dict_keys:
        nested = data.get(dict_key)
        if not isinstance(nested, dict):
            continue
        for key in keys:
            value = nested.get(key)
            if isinstance(value, str) and value.strip():
                return value.strip()

    return None


def _request_json(method: str, url: str, **kwargs: Any) -> dict[str, Any]:
    r = SESSION.request(method, url, timeout=kwargs.pop("timeout", 15), **kwargs)
    r.raise_for_status()
    try:
        return r.json()
    except ValueError as e:
        raise JLC3DError(f"接口返回不是合法 JSON: {url}") from e


def get_component_data(code: str) -> dict[str, Any]:
    code = normalize_code(code)
    url = EXACT_COMPONENT_API.format(lcsc_id=code)
    obj = _request_json("GET", url)

    if not obj or obj.get("success") is False:
        msg = obj.get("message") or obj.get("msg") or "未找到该器件"
        raise JLC3DError(str(msg))

    result = obj.get("result")
    if not isinstance(result, dict) or not result:
        raise JLC3DError("未找到该器件")

    return result


def search_products(keyword: str, page_size: int = 30) -> list[dict[str, Any]]:
    keyword = normalize_input(keyword)
    payload = {
        "keyword": keyword,
        "needAggs": "true",
        "currPage": "1",
        "pageSize": str(page_size),
    }
    obj = _request_json("POST", f"{BASE_API}/eda/product/search", data=payload)
    products = ((obj.get("result") or {}).get("productList") or [])
    return [item for item in products if isinstance(item, dict)]


def rank_search_products(products: list[dict[str, Any]], keyword: str) -> list[dict[str, Any]]:
    keyword_raw = normalize_input(keyword)
    keyword_upper = normalize_code(keyword_raw)
    keyword_key = normalize_match_key(keyword_raw)

    def item_score(item: dict[str, Any]) -> tuple[int, int, int, int, int]:
        names = list(_iter_candidate_names(item))
        code = extract_component_code(item) or ""
        values = names + ([code] if code else [])

        exact_code = 0
        exact_name = 0
        prefix_name = 0
        contains_name = 0
        key_contains = 0

        for value in values:
            value_upper = normalize_code(value)
            value_key = normalize_match_key(value)
            if value_upper == keyword_upper:
                if value == code:
                    exact_code = 1
                else:
                    exact_name = 1
            if value_key and keyword_key:
                if value_key == keyword_key:
                    exact_name = max(exact_name, 1)
                if value_key.startswith(keyword_key) or keyword_key.startswith(value_key):
                    prefix_name = 1
                if keyword_key in value_key:
                    key_contains = 1
            if keyword_upper and keyword_upper in value_upper:
                contains_name = 1

        has_device = 1 if item.get("hasDevice") else 0
        return (exact_code, exact_name, prefix_name, key_contains, contains_name + has_device)

    return sorted(products, key=item_score, reverse=True)


def get_model_uuid_by_device(device_uuid: str) -> str:
    obj = _request_json(
        "POST",
        f"{BASE_API}/devices/searchByIds",
        data={"uuids[]": device_uuid},
    )
    result = obj.get("result") or []
    if not result:
        raise JLC3DError("未找到器件详情")

    attrs = result[0].get("attributes") or {}
    model_uuid = attrs.get("3D Model")
    if not model_uuid:
        raise JLC3DError("该器件没有 3D 模型")
    return str(model_uuid)


def get_model_uuid_from_component_data(component_data: dict[str, Any]) -> str:
    package_detail = component_data.get("packageDetail") or {}
    data_str = package_detail.get("dataStr")
    if isinstance(data_str, str):
        try:
            data_str = json.loads(data_str)
        except json.JSONDecodeError:
            data_str = None

    if isinstance(data_str, dict):
        shape_lines = data_str.get("shape") or []
        for line in shape_lines:
            if not isinstance(line, str) or not line.startswith("SVGNODE~"):
                continue
            raw_json = line.split("~", 1)[1]
            try:
                node = json.loads(raw_json)
            except json.JSONDecodeError:
                continue
            attrs = node.get("attrs") or {}
            model_uuid = attrs.get("uuid")
            if isinstance(model_uuid, str) and model_uuid.strip():
                return model_uuid.strip()

    attrs = component_data.get("attributes") or {}
    model_uuid = attrs.get("3D Model")
    if isinstance(model_uuid, str) and model_uuid.strip():
        return model_uuid.strip()

    raise JLC3DError("该器件没有 3D 模型")


def resolve_model_info(query: str) -> tuple[str, str]:
    query = normalize_input(query)
    if not query:
        raise JLC3DError("请输入元器件编号或名称")

    errors: list[str] = []

    if looks_like_lcsc_code(query):
        try:
            component_data = get_component_data(query)
            model_uuid = get_model_uuid_from_component_data(component_data)
            display_name = extract_component_display_name(component_data, fallback_code=query)
            return model_uuid, display_name
        except Exception as e:
            errors.append(f"精确编号查询失败: {e}")

    try:
        products = search_products(query)
        if not products:
            raise JLC3DError("未找到匹配的器件")

        ranked_products = rank_search_products(products, query)
        candidate_errors: list[str] = []

        for item in ranked_products[:10]:
            display_name = extract_component_display_name(item, fallback_code=query)
            code_candidate = extract_component_code(item)
            device_uuid = item.get("hasDevice")

            if code_candidate:
                try:
                    component_data = get_component_data(code_candidate)
                    model_uuid = get_model_uuid_from_component_data(component_data)
                    display_name = extract_component_display_name(component_data, fallback_code=display_name)
                    return model_uuid, display_name
                except Exception as e:
                    candidate_errors.append(f"{display_name}/{code_candidate}: {e}")

            if device_uuid:
                try:
                    model_uuid = get_model_uuid_by_device(str(device_uuid))
                    return model_uuid, display_name
                except Exception as e:
                    candidate_errors.append(f"{display_name}/device={device_uuid}: {e}")

        if candidate_errors:
            errors.append("名称搜索候选均失败: " + " | ".join(candidate_errors[:5]))
        else:
            errors.append("名称搜索结果中没有可用 3D 模型")
    except Exception as e:
        errors.append(f"名称搜索失败: {e}")

    raise JLC3DError("；".join(errors))


def get_model_uuid(code: str) -> str:
    model_uuid, _display_name = resolve_model_info(code)
    return model_uuid


def download_step_file(model_uuid: str) -> bytes:
    last_error: Optional[Exception] = None

    for url_tpl in STEP_URLS:
        url = url_tpl.format(uuid=model_uuid)
        try:
            r = SESSION.get(url, timeout=(8, 20))
            r.raise_for_status()
            if not r.content:
                raise JLC3DError("返回内容为空")
            return r.content
        except Exception as e:
            last_error = e

    raise JLC3DError(f"STEP 下载失败: {last_error}")


def sanitize_filename(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*]', "_", name)
    name = name.rstrip(" .")
    name = re.sub(r"\s+", " ", name).strip()
    return name or "model"


# =====================================================
# UI
# =====================================================
class JLC3DApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("嘉立创 3D 模型下载器")
        self.app_width = 600
        self.app_height = 420
        self.center_window(self, self.app_width, self.app_height)
        self.configure(bg="#f8f9fa")

        self.download_path = load_download_path() or default_desktop()
        self.last_download_file: Optional[str] = None

        self._build_menu()
        self._build_ui()

    def center_window(self, target: tk.Tk | tk.Toplevel, width: int, height: int) -> None:
        screen_width = target.winfo_screenwidth()
        screen_height = target.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        target.geometry(f"{width}x{height}+{x}+{y}")

    def _build_menu(self) -> None:
        menu = tk.Menu(self)
        self.config(menu=menu)

        file_menu = tk.Menu(menu, tearoff=0)
        menu.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="修改下载路径", command=self.choose_path)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.quit)

        help_menu = tk.Menu(menu, tearoff=0)
        menu.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="关于", command=self.show_about)

    def _build_ui(self) -> None:
        main_container = tk.Frame(self, bg="#f8f9fa", padx=20, pady=10)
        main_container.pack(fill="both", expand=True)

        center_row = tk.Frame(main_container, bg="#f8f9fa")
        center_row.pack(pady=(40, 40))

        tk.Label(
            center_row,
            text="编号/名称:",
            bg="#f8f9fa",
            font=("Microsoft YaHei", 12, "bold"),
        ).pack(side="left", padx=(0, 10))

        self.entry = ttk.Entry(center_row, font=("Arial", 14), width=20)
        self.entry.insert(0, "C8734")
        self.entry.pack(side="left", padx=5)
        self.entry.bind("<Return>", lambda _e: self.start_download())

        self.btn_download = tk.Button(
            center_row,
            text="立即下载",
            bg="#28a745",
            fg="white",
            font=("Microsoft YaHei", 11, "bold"),
            relief="flat",
            width=10,
            height=1,
            cursor="hand2",
            command=self.start_download,
        )
        self.btn_download.pack(side="left", padx=15)

        self.log = scrolledtext.ScrolledText(
            main_container,
            height=8,
            font=("Consolas", 10),
            bg="white",
            relief="solid",
            borderwidth=1,
        )
        self.log.pack(fill="both", expand=True)

        bottom_frame = tk.Frame(main_container, bg="#f8f9fa", pady=15)
        bottom_frame.pack(fill="x")

        self.path_label = tk.Label(
            bottom_frame,
            text=f"保存至: {self.download_path}",
            bg="#f8f9fa",
            fg="#495057",
            font=("Microsoft YaHei", 10),
            anchor="w",
        )
        self.path_label.pack(side="left", fill="x", expand=True)

        btn_style = ttk.Style()
        btn_style.configure("Small.TButton", font=("Microsoft YaHei", 10))

        ttk.Button(
            bottom_frame,
            text="定位文件",
            style="Small.TButton",
            width=10,
            command=self.locate_file,
        ).pack(side="right")

    def log_msg(self, msg: str) -> None:
        now = datetime.now().strftime("%H:%M:%S")
        self.log.insert(tk.END, f"[{now}] {msg}\n")
        self.log.see(tk.END)

    def choose_path(self) -> None:
        p = filedialog.askdirectory()
        if p:
            self.download_path = p
            save_download_path(p)
            self.path_label.config(text=f"保存至: {p}")

    def locate_file(self) -> None:
        if not self.last_download_file or not os.path.exists(self.last_download_file):
            messagebox.showinfo("提示", "还没有可定位的下载文件")
            return

        path = self.last_download_file
        if sys.platform.startswith("win"):
            subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
        elif sys.platform.startswith("darwin"):
            subprocess.call(["open", "-R", path])
        else:
            subprocess.call(["xdg-open", os.path.dirname(path)])

    def show_about(self) -> None:
        about = tk.Toplevel(self)
        about.title("关于")
        self.center_window(about, 420, 260)
        about.resizable(False, False)

        text = tk.Text(about, wrap="word", padx=15, pady=15, font=("Microsoft YaHei", 11))
        text.pack(fill="both", expand=True)

        content = (
            "嘉立创 3D 模型下载器\n"
            "修复版：1.4\n"
            "作者：ChatGPT 修复\n\n"
            "项目地址：\n"
            "https://github.com/zhanghankui/JLC3DDownload\n\n"
            "基于zhutongxueya版本由ChatGPT修复：\n"
            "https://github.com/zhutongxueya/JLC3DDownload\n\n"
            "修复点：\n"
            "1. 支持编号和名称搜索\n"
            "2. 名称搜索会尝试多个候选器件\n"
            "3. 直接使用 3D UUID 下载 STEP\n"
            "4. 保存名优先使用器件型号\n"
        )
        text.insert("1.0", content)
        text.config(state="disabled")

        start = "5.0"
        end = "5.end"
        text.config(state="normal")
        text.tag_add("link", start, end)
        text.tag_config("link", foreground="blue", underline=True)
        text.tag_bind(
            "link",
            "<Button-1>",
            lambda _e: webbrowser.open("https://github.com/zhutongxueya/JLC3DDownload"),
        )
        text.config(state="disabled")

    def set_download_button_state(self, enabled: bool) -> None:
        if enabled:
            self.btn_download.config(state="normal", bg="#28a745", text="立即下载")
        else:
            self.btn_download.config(state="disabled", bg="#6c757d", text="下载中...")

    def start_download(self) -> None:
        self.set_download_button_state(False)
        threading.Thread(target=self.download_task, daemon=True).start()

    def download_task(self) -> None:
        query = normalize_input(self.entry.get())
        if not query:
            self.after(0, lambda: messagebox.showwarning("提示", "请输入元器件编号或名称"))
            self.after(0, lambda: self.set_download_button_state(True))
            return

        try:
            self.after(0, lambda: self.log_msg(f"查询器件〖{query}〗…"))
            model_uuid, display_name = resolve_model_info(query)
            self.after(0, lambda: self.log_msg(f"器件名称：{display_name}"))
            self.after(0, lambda: self.log_msg(f"解析 3D UUID 成功：{model_uuid}"))

            self.after(0, lambda: self.log_msg("下载 STEP 文件…"))
            data = download_step_file(model_uuid)

            os.makedirs(self.download_path, exist_ok=True)
            safe_name = sanitize_filename(display_name)
            filepath = os.path.join(self.download_path, f"{safe_name}.step")
            with open(filepath, "wb") as f:
                f.write(data)

            self.last_download_file = filepath
            self.after(0, lambda: self.log_msg(f"下载完成 ✔\n保存至：{filepath}"))
        except Exception as e:
            self.after(0, lambda: self.log_msg(f"错误：{e}"))
        finally:
            self.after(0, lambda: self.set_download_button_state(True))


if __name__ == "__main__":
    JLC3DApp().mainloop()
