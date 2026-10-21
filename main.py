#!/usr/bin/env python3

# -*- coding: utf-8 -*-

"""

网页自动化工具 v1.0

"""



import sys

import json

import time
import re
import os
import shutil
import subprocess

import uuid

from pathlib import Path

from datetime import datetime


def _configure_console_output():
    """避免 Windows 打包 exe 在 GBK 控制台输出特殊符号时崩溃。"""
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(errors="replace")
            except Exception:
                pass


_configure_console_output()



# 打包为 exe 时用 exe 所在目录，源码运行时用脚本所在目录。
# Linux DEB 启动器通过 CHROMEAUTORUN_HOME 将可写数据放到用户目录，避免
# 配置、浏览器配置文件和输出写入 /opt 下的只读程序目录。

if getattr(sys, 'frozen', False):

    APP_DIR = Path(sys.executable).resolve().parent

else:

    APP_DIR = Path(__file__).resolve().parent

_runtime_data_dir = os.environ.get('CHROMEAUTORUN_HOME', '').strip()
if os.name != 'nt' and _runtime_data_dir:
    BASE_DIR = Path(_runtime_data_dir).expanduser().resolve()
    BASE_DIR.mkdir(parents=True, exist_ok=True)
else:
    BASE_DIR = APP_DIR

CONFIG_FILE = BASE_DIR / 'config.json'

EXCEL_CONFIG_FILE = BASE_DIR / 'config.xlsx'

RECORDINGS  = BASE_DIR / 'recordings'

OUTPUT      = BASE_DIR / 'output'

CHROME_PROFILE_DIR = BASE_DIR / 'chrome_profile'



OUTPUT.mkdir(exist_ok=True)   # 确保 output/ 存在



# ─── 全局浏览器实例（跨菜单保持） ───────────────────────

_driver = None

_saved_page_url = None     # 登录后保存的页面 URL，用于崩溃后重新导航

_saved_iframe_src = None   # 首次切入报表 iframe 时保存的 src，用于刷新后重定位


def _resolve_download_dir():
    """解析浏览器下载目录。

    优先级:
    1) config.json 的 download_dir
    2) 系统默认下载目录 USERPROFILE/Downloads
    """
    try:
        if CONFIG_FILE.exists():
            cfg = json.loads(CONFIG_FILE.read_text(encoding='utf-8'))
            d = str(cfg.get('download_dir', '') or '').strip()
            if d:
                p = Path(d).expanduser()
                p.mkdir(parents=True, exist_ok=True)
                return p.resolve()
    except Exception:
        pass

    p = Path.home() / 'Downloads'
    p.mkdir(parents=True, exist_ok=True)
    return p.resolve()



DEFAULT_LOGIN_SELECTORS = {

    "credit_code_selector": "input[placeholder='请输入社会信用代码']",

    "username_selector": "input[placeholder='请选择用户名']",

    "password_selector": "input[placeholder='请输入密码'], input[type='password']",

    "submit_selector": "button.login-btn, button.el-button--primary",

    "success_indicator": ".el-menu, .avue-layout, .main-content",

}



SCENE_MATRIX_ORDER = {

    "county": [

        "data_date",

        "bill_financing",

        "country",

        "province",

        "city",

        "county",

        "institution_class",

    ],

    "org": [

        "data_date",

        "bill_financing",

        "country",

        "institution_name",

        "province",

        "city",

        "county",

    ],

    "county2": [

        "data_date",

        "bill_financing",

        "country",

        "province",

        "city",

        "county",

        "institution_class",

    ],

}

SCENE_FILL_ORDER = {
    "county": [
        "data_date",
        "bill_financing",
        "country",
        "province",
        "city",
        "county",
        "institution_class",
    ],
    "org": [
        "data_date",
        "bill_financing",
        "country",
        "institution_name",
        "province",
        "city",
        "county",
    ],
    "county2": [
        "data_date",
        "bill_financing",
        "country",
        "province",
        "city",
        "county",
        "institution_class",
    ],
}

GEO_FIELDS = {"province", "city", "county"}
FIELD_METHODS = {"sendkey", "checkbox_dropdown"}


def _default_field_method(scene, field):
    """字段默认填充方式；配置 field.xxx.method 可覆盖。"""
    if field == "data_date":
        return "sendkey"
    if scene == "county" and field in GEO_FIELDS:
        return "checkbox_dropdown"
    return "sendkey"

REPORT_LABEL_MAP = {

    "data_date": "数据日期:",

    "bill_financing": "是否含票据融资:",

    "country": "全国:",

    "province": "省份:",

    "city": "地市:",

    "county": "县区:",

    "institution_class": "机构类型:",

    "institution_name": "机构名称:",

}



SCENE_SHEET_MAP = {

    "county": "county",

    "org": "org",

    "county2": "county2",

}

SCENE_REPORT_DEFAULTS = {
    "county": {
        "report_title_keyword": "地区汇总-金融机构在当地的经办机构",
    },
    "org": {
        "report_title_keyword": "单家机构",
    },
    "county2": {
        "report_title_keyword": "地区汇总-注册在当地的地方性法人金融机构",
    },
}







# ═══════════════════════════════════════════════════

#  工具函数

# ═══════════════════════════════════════════════════



def load_config():

    """从 config.json 读取登录配置。



    调用方: auto_login（自动登录时读取 url/credit_code/username/password）。

    返回值: dict | None。

    """

    if not CONFIG_FILE.exists():

        print(f'❌ 配置文件不存在: {CONFIG_FILE}')

        return None

    try:

        return json.loads(CONFIG_FILE.read_text(encoding='utf-8'))

    except json.JSONDecodeError as e:

        print(f'❌ config.json 格式错误: {e}')

        return None







def _init_template_sheets(wb):
    """按当前约定重建模板 sheet（login/county/org/county2/standalone_fill/fr_widget_map）。"""
    for name in ("login", "county", "org", "county2", "standalone_fill", "fr_widget_map"):
        if name in wb.sheetnames:
            del wb[name]

    ws = wb.create_sheet("login")
    ws.append(["key", "value", "note"])

    def add_kv(ws, key, value, note):
        ws.append([key, value, note])

    def add_field(ws, field, mode, values, method, value_note='填"不选"表示清空该字段', validate_options=None):
        add_kv(ws, f"field.{field}.mode", mode, "fixed=单值；list=多值用|分隔")
        add_kv(ws, f"field.{field}.values", values, value_note)
        add_kv(ws, f"field.{field}.method", method, "sendkey=普通输入；checkbox_dropdown=多选checkbox/tag")
        if validate_options is not None:
            add_kv(ws, f"field.{field}.validate_options", validate_options, "yes=校验机构候选池；no=不校验")

    def add_matrix_sheet(name, scene_note, methods):
        wsx = wb.create_sheet(name)
        wsx.append(["key", "value", "note"])
        defaults = SCENE_REPORT_DEFAULTS.get(name, {})
        add_kv(wsx, "task.mode", "matrix", "固定为matrix")
        add_kv(wsx, "task.report_title_keyword", defaults.get("report_title_keyword", ""), "按外层报表标签页标题匹配iframe；建议填写能唯一定位报表的关键词")
        add_kv(wsx, "task.report_title_index", "", "同名/包含匹配多个时可填1/2/3指定第几个；留空则报歧义不执行")
        add_kv(wsx, "task.tree_path", "", "刷新后导航树的路径，用 > 分隔层级，如：科技贷款统计 > 单家机构 > ...(报表)")
        add_kv(wsx, "task.refresh_interval", "0", "定时刷新间隔（条数），0=不自动刷新；建议仅在出错时触发恢复")
        add_kv(wsx, "task.query", "yes", "yes=执行查询；no=仅填充不查询")
        add_kv(wsx, "task.export", "yes", "yes=查询后导出；no=不导出")
        add_kv(wsx, "task.query_wait_seconds", "8", "查询后等待秒数")
        add_kv(wsx, "task.field_settle_seconds", "0", "每次 sendkey/下拉成功设置后的等待秒数；0=沿用默认短等待")
        add_kv(wsx, "task.max_items", "0", f"0=不限制；{scene_note}")
        add_field(wsx, "data_date", "list", "2026-03-31", methods.get("data_date", "sendkey"), "多值用|分隔，格式YYYY-MM-DD")
        add_field(wsx, "bill_financing", "fixed", "是", methods.get("bill_financing", "sendkey"))
        add_field(wsx, "country", "fixed", "全国", methods.get("country", "sendkey"))
        add_field(wsx, "province", "fixed", "广东省", methods.get("province", "sendkey"))
        add_field(wsx, "city", "fixed", "惠州市", methods.get("city", "sendkey"))
        add_field(wsx, "county", "list", "不选|惠城区", methods.get("county", "sendkey"))
        if name == "org":
            add_field(wsx, "institution_name", "list", "中国银行股份有限公司", methods.get("institution_name", "sendkey"), validate_options="yes")
        else:
            add_field(wsx, "institution_class", "list", "不选", methods.get("institution_class", "sendkey"), validate_options="yes")
        return wsx

    county_methods = {
        "province": "checkbox_dropdown",
        "city": "checkbox_dropdown",
        "county": "checkbox_dropdown",
    }
    add_matrix_sheet("county", "地区汇总报表查询导出（分县区）", county_methods)
    add_matrix_sheet("org", "单家机构报表查询导出（分机构）", {})
    add_matrix_sheet("county2", "注册当地法人机构报表查询导出", {})

    ws4 = wb.create_sheet("standalone_fill")
    ws4.append(["enabled", "report_name", "condition_index", "field_label", "field_type", "field_value", "match_mode", "placeholder", "note"])

    ws5 = wb.create_sheet("fr_widget_map")
    ws5.append(["field_key", "label", "label_widget", "input_widget", "note"])


def _ensure_query_config_schema(wb):
    """补齐现有 config.xlsx 的查询配置结构，尽量不改用户已有配置值。"""
    from openpyxl import Workbook

    changed = False

    def append_row_if_missing(ws, key, value, note):
        nonlocal changed
        existing = {
            str(ws.cell(r, 1).value or '').strip()
            for r in range(1, ws.max_row + 1)
        }
        if key not in existing:
            ws.append([key, value, note])
            changed = True

    def ensure_headers(ws):
        nonlocal changed
        headers = ["key", "value", "note"]
        for cidx, val in enumerate(headers, start=1):
            if ws.cell(1, cidx).value != val:
                ws.cell(1, cidx).value = val
                changed = True

    def copy_template_sheet(sheet_name):
        nonlocal changed
        tmp = Workbook()
        if tmp.sheetnames:
            del tmp[tmp.sheetnames[0]]
        _init_template_sheets(tmp)
        src = tmp[sheet_name]
        dst = wb.create_sheet(sheet_name)
        for row in src.iter_rows(values_only=True):
            dst.append(row)
        changed = True
        return dst

    def ensure_matrix_sheet(sheet_name, fields, method_defaults, validate_defaults):
        if sheet_name not in wb.sheetnames:
            copy_template_sheet(sheet_name)
            return
        ws = wb[sheet_name]
        ensure_headers(ws)
        defaults = SCENE_REPORT_DEFAULTS.get(sheet_name, {})
        append_row_if_missing(
            ws,
            "task.report_title_keyword",
            defaults.get("report_title_keyword", ""),
            "按外层报表标签页标题匹配iframe；建议填写能唯一定位报表的关键词",
        )
        append_row_if_missing(
            ws,
            "task.report_title_index",
            "",
            "同名/包含匹配多个时可填1/2/3指定第几个；留空则报歧义不执行",
        )
        append_row_if_missing(
            ws,
            "task.field_settle_seconds",
            "0",
            "每次 sendkey/下拉成功设置后的等待秒数；0=沿用默认短等待",
        )
        for field in fields:
            append_row_if_missing(
                ws,
                f"field.{field}.method",
                method_defaults.get(field, _default_field_method(sheet_name, field)),
                "sendkey=普通输入；checkbox_dropdown=多选checkbox/tag",
            )
            if field in validate_defaults:
                append_row_if_missing(
                    ws,
                    f"field.{field}.validate_options",
                    validate_defaults[field],
                    "yes=校验机构候选池；no=不校验",
                )

    county_fields = ["data_date", "bill_financing", "country", "province", "city", "county", "institution_class"]
    org_fields = ["data_date", "bill_financing", "country", "institution_name", "province", "city", "county"]
    county_methods = {"province": "checkbox_dropdown", "city": "checkbox_dropdown", "county": "checkbox_dropdown"}
    ensure_matrix_sheet("county", county_fields, county_methods, {"institution_class": "yes"})
    ensure_matrix_sheet("org", org_fields, {}, {"institution_name": "yes"})
    ensure_matrix_sheet("county2", county_fields, {}, {"institution_class": "yes"})
    return changed


def reset_config_template():
    """补齐 config.xlsx 缺失配置，绝不覆盖用户已有值。"""
    try:
        from openpyxl import Workbook, load_workbook
        if EXCEL_CONFIG_FILE.exists():
            wb = load_workbook(EXCEL_CONFIG_FILE)
        else:
            wb = Workbook()
            if wb.sheetnames:
                del wb[wb.sheetnames[0]]

        tmp = Workbook()
        if tmp.sheetnames:
            del tmp[tmp.sheetnames[0]]
        _init_template_sheets(tmp)

        for name in tmp.sheetnames:
            src = tmp[name]
            if name not in wb.sheetnames:
                dst = wb.create_sheet(name)
                for row in src.iter_rows(values_only=True):
                    dst.append(list(row))
                continue

            # 现有工作表只追加缺失的 key，不修改表头、已有行或用户填写的值。
            dst = wb[name]
            existing_keys = {
                str(dst.cell(r, 1).value or '').strip()
                for r in range(1, dst.max_row + 1)
            }
            for row in src.iter_rows(min_row=2, values_only=True):
                key = str(row[0] or '').strip() if row else ''
                if key and key not in existing_keys:
                    dst.append(list(row))
                    existing_keys.add(key)

        wb.save(EXCEL_CONFIG_FILE)
        print(f'✅ 已补齐缺失配置项（未覆盖已有值）: {EXCEL_CONFIG_FILE.name}')
    except Exception as e:
        print(f'❌ 重置配置模板失败: {e}')


def open_config_excel():
    """用系统默认程序打开 config.xlsx，便于直接维护配置。"""
    if not EXCEL_CONFIG_FILE.exists():
        build_runtime_config_from_excel()
        if not EXCEL_CONFIG_FILE.exists():
            print(f'❌ 配置文件不存在且自动创建失败: {EXCEL_CONFIG_FILE}')
            return
    try:
        if os.name == 'nt':
            os.startfile(str(EXCEL_CONFIG_FILE))
        else:
            xdg_open = shutil.which('xdg-open')
            if not xdg_open:
                raise FileNotFoundError('未找到 xdg-open，请安装 xdg-utils')
            subprocess.Popen(
                [xdg_open, str(EXCEL_CONFIG_FILE)],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        print(f'✅ 已打开配置文件: {EXCEL_CONFIG_FILE}')
    except Exception as e:
        print(f'❌ 打开配置文件失败: {e}')

def build_runtime_config_from_excel():

    """从 config.xlsx 的 login sheet（A列key / B列value）读取登录配置，生成运行时 dict。



    若 config.xlsx 不存在则自动创建模板文件（含 login/county/org/county2/standalone_fill/fr_widget_map 等 sheet）。

    同时填充 DEFAULT_LOGIN_SELECTORS 中的选择器默认值。



    调用方: sync_config_json_from_excel。

    返回值: dict | None。

    """

    if not EXCEL_CONFIG_FILE.exists():

        try:

            from openpyxl import Workbook

            wb = Workbook()

            if wb.sheetnames:
                del wb[wb.sheetnames[0]]
            _init_template_sheets(wb)


            wb.save(EXCEL_CONFIG_FILE)

            print(f'⚠️  已自动创建模板: {EXCEL_CONFIG_FILE.name}')

            print('   请填写 login/county/org/county2/standalone_fill 相关 sheet 后重试')

        except Exception as e:

            print(f'❌ 配置文件不存在且创建模板失败: {e}')

        return None



    try:

        from openpyxl import load_workbook

    except ImportError:

        print('❌ 未安装 openpyxl，请执行: pip install openpyxl')

        return None



    try:

        wb = load_workbook(EXCEL_CONFIG_FILE, data_only=False)

        try:
            if _ensure_query_config_schema(wb):
                wb.save(EXCEL_CONFIG_FILE)
                print('✅ 已补齐 config.xlsx 查询配置结构（报表title/method/机构候选校验行 / county2 sheet）')
        except Exception as e:
            print(f'⚠️  config.xlsx 结构补齐失败（不影响旧配置默认运行）: {e}')

        ws = wb["login"] if "login" in wb.sheetnames else wb.active

    except Exception as e:

        print(f'❌ 读取 config.xlsx 失败: {e}')

        return None



    row_map = {}

    for row in ws.iter_rows(min_row=2, values_only=True):

        key = str(row[0]).strip() if row and row[0] is not None else ''

        val = str(row[1]).strip() if row and len(row) > 1 and row[1] is not None else ''

        if key:

            row_map[key] = val



    cfg = {

        "url": row_map.get("url", ""),

        "credit_code": row_map.get("credit_code", ""),

        "username": row_map.get("username", ""),

        "password": row_map.get("password", ""),

        "login": dict(DEFAULT_LOGIN_SELECTORS),

    }



    return cfg





def sync_config_json_from_excel():

    """将 config.xlsx 的 login 配置同步写入 config.json。



    这是启动时的关键步骤——后续所有登录操作读取的是 config.json 而非 xlsx，

    因此每次登录前都需要同步以确保 json 文件与 xlsx 一致。



    调用方: auto_login（登录前同步）、main（启动时同步）。

    返回值: dict | None。

    """

    cfg = build_runtime_config_from_excel()

    if not cfg:

        return None

    try:

        CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding='utf-8')

    except Exception as e:

        print(f'❌ 写入 config.json 失败: {e}')

        return None

    return cfg





def _split_values(text):

    """将竖线分隔的字符串拆成列表（如 "值1|值2|值3" → ["值1", "值2", "值3"]）。



    用于解析 config.xlsx 中 field.X.values 等多值配置项。



    调用方: parse_matrix_config。

    """

    if not text:

        return []

    return [x.strip() for x in str(text).split('|') if x and x.strip()]


def _is_yes(value):
    """解析配置表中的 yes/no 开关。"""
    return str(value or '').strip().lower() in ('yes', 'y', '1', 'true')


def _normalize_date_text(value):
    """将日期值规范为 YYYY-MM-DD，无法识别时返回原值字符串。"""
    if value is None:
        return ''
    text = str(value).strip()
    if not text:
        return ''
    # 常见格式: 2026-03-31 / 2026-03-31 00:00:00 / 2026/03/31 00:00:00
    m = re.match(r'^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:\s+.*)?$', text)
    if m:
        y, mo, d = m.groups()
        return f'{int(y):04d}-{int(mo):02d}-{int(d):02d}'
    return text





def _read_kv_sheet(ws):

    """从 xlsx sheet 读取两列键值对（A列key / B列value，首行为表头自动跳过）。



    通用工具函数——所有只需要两列配置的 sheet 都通过它读取。



    调用方: build_runtime_config_from_excel、load_task_config、environment_check。

    返回值: dict[str, str]。

    """

    row_map = {}

    for row in ws.iter_rows(min_row=2, values_only=True):

        key = str(row[0]).strip() if row and row[0] is not None else ''

        val = str(row[1]).strip() if row and len(row) > 1 and row[1] is not None else ''

        if key:

            row_map[key] = val

    return row_map





def _read_table_sheet(ws):

    """从 xlsx sheet 读取多列表格数据（首行为表头），每行转为 {列名: 值} 字典。



    自动跳过全空行。



    调用方: load_fr_widget_map。

    返回值: list[dict]。

    """

    headers = []

    for cell in next(ws.iter_rows(min_row=1, max_row=1, values_only=True), []):

        headers.append(str(cell).strip() if cell is not None else '')



    rows = []

    for raw in ws.iter_rows(min_row=2, values_only=True):

        item = {}

        has_value = False

        for idx, key in enumerate(headers):

            if not key:

                continue

            val = raw[idx] if idx < len(raw) else ''

            val = str(val).strip() if val is not None else ''

            if val:

                has_value = True

            item[key] = val

        if has_value:

            rows.append(item)

    return rows





def _truthy(value, default=False):

    """将 yes/no/1/0/true/false/是 等配置字符串转为 Python bool。



    用于解析 config.xlsx 中的开关类字段（enabled、multi_condition 等）。



    调用方: environment_check。

    """

    if value is None or value == '':

        return default

    return str(value).strip().lower() in ('yes', 'y', '1', 'true', '是')





def _int_setting(value, default=0, minimum=0):

    """将配置字符串转为整数，解析失败时返回默认值，并确保不低于最小值。



    用于解析 config.xlsx 中的数值类字段（max_items、sort_order 等）。



    调用方: environment_check。

    """

    try:

        parsed = int(str(value).strip())

    except Exception:

        parsed = default

    return max(minimum, parsed)





def load_task_config(scene):

    """从 config.xlsx 读取指定场景（county/org/county2）的键值配置。



    调用方: run_matrix_export（按场景读取查询导出配置）。

    参数 scene: 对应 config.xlsx 中的 sheet 名。

    返回值: dict[str, str] | None。

    """

    if not EXCEL_CONFIG_FILE.exists():

        print(f'❌ 配置文件不存在: {EXCEL_CONFIG_FILE}')

        return None

    try:

        from openpyxl import load_workbook

        wb = load_workbook(EXCEL_CONFIG_FILE, data_only=True)

        sheet_name = SCENE_SHEET_MAP.get(scene, scene)

        if sheet_name not in wb.sheetnames:

            print(f'❌ 配置sheet不存在: {sheet_name}')

            return None

        ws = wb[sheet_name]

    except Exception as e:

        print(f'❌ 读取配置失败: {e}')

        return None

    return _read_kv_sheet(ws)





def parse_matrix_config(row_map, scene=None):

    """解析 county/org/county2 sheet 的矩阵查询配置，转换为结构化执行计划。



    从 load_task_config 返回的扁平键值对中提取各字段的 mode

    （fixed/list）和 values，组装成可执行的查询遍历计划。



    调用方: run_matrix_export。

    参数 row_map: load_task_config 返回的 {key: value} 扁平字典。

    返回值: dict {query, export, query_wait_seconds, max_items, fields}。

    异常: ValueError（mode 不合法、fixed 多值等）。

    """

    mode = (row_map.get('task.mode', 'matrix') or '').strip().lower()

    if mode and mode != 'matrix':

        raise ValueError('task.mode 当前仅支持 matrix')



    query_yes = (row_map.get('task.query', 'yes') or '').strip().lower() in ('yes', 'y', '1', 'true')

    export_yes = (row_map.get('task.export', 'yes') or '').strip().lower() in ('yes', 'y', '1', 'true')

    scene_defaults = SCENE_REPORT_DEFAULTS.get(scene, {})
    report_title_keyword = (row_map.get('task.report_title_keyword', scene_defaults.get('report_title_keyword', '')) or '').strip()
    report_title_index_raw = (row_map.get('task.report_title_index', '') or '').strip()
    report_title_index = 0
    if report_title_index_raw:
        try:
            report_title_index = int(report_title_index_raw)
        except Exception:
            raise ValueError(f'task.report_title_index 不合法: {report_title_index_raw}，应填写正整数或留空')
        if report_title_index < 1:
            raise ValueError('task.report_title_index 必须为正整数或留空')

    tree_path = (row_map.get('task.tree_path', '') or '').strip()

    try:

        query_wait_seconds = float((row_map.get('task.query_wait_seconds', '6') or '6').strip())

    except Exception:

        query_wait_seconds = 6.0

    query_wait_seconds = max(1.0, query_wait_seconds)

    try:
        field_settle_seconds = float((row_map.get('task.field_settle_seconds', '0') or '0').strip())
    except Exception:
        field_settle_seconds = 0.0
    field_settle_seconds = max(0.0, min(30.0, field_settle_seconds))

    try:

        max_items = int((row_map.get('task.max_items', '0') or '0').strip())

    except Exception:

        max_items = 0

    max_items = max(0, max_items)

    try:
        refresh_interval = int((row_map.get('task.refresh_interval', '0') or '0').strip())
    except Exception:
        refresh_interval = 0
    refresh_interval = max(0, refresh_interval)  # 0 表示不自动刷新，仅出错时触发恢复



    fields = {}

    if scene in SCENE_FILL_ORDER:
        all_fields = list(SCENE_FILL_ORDER[scene])
    else:
        all_fields = list(dict.fromkeys(SCENE_FILL_ORDER["county"] + SCENE_FILL_ORDER["org"]))

    for f in all_fields:

        mode_key = f'field.{f}.mode'

        val_key = f'field.{f}.values'

        method_key = f'field.{f}.method'

        validate_key = f'field.{f}.validate_options'

        # 兼容旧key: org_type / institution

        if f == 'institution_class':

            mode_key = mode_key if mode_key in row_map else 'field.org_type.mode'

            val_key = val_key if val_key in row_map else 'field.org_type.values'

            method_key = method_key if method_key in row_map else 'field.org_type.method'

            validate_key = validate_key if validate_key in row_map else 'field.org_type.validate_options'

        if f == 'institution_name':

            mode_key = mode_key if mode_key in row_map else 'field.institution.mode'

            val_key = val_key if val_key in row_map else 'field.institution.values'

            method_key = method_key if method_key in row_map else 'field.institution.method'

            validate_key = validate_key if validate_key in row_map else 'field.institution.validate_options'



        m = (row_map.get(mode_key, 'fixed') or 'fixed').strip().lower()

        vals = _split_values(row_map.get(val_key, ''))

        method = (row_map.get(method_key, '') or '').strip().lower()

        validate_options = _is_yes(row_map.get(validate_key, 'no')) and f in ('institution_class', 'institution_name')

        if not method:

            method = _default_field_method(scene, f)

        if m not in ('fixed', 'list'):

            raise ValueError(f'field.{f}.mode 不合法: {m}，仅支持 fixed/list；如需遍历多个值请使用 list')

        if m == 'fixed' and len(vals) > 1:

            raise ValueError(f'field.{f}.mode=fixed 只能有一个值')

        if m == 'fixed' and len(vals) == 0:

            vals = ['']

        if method not in FIELD_METHODS:

            raise ValueError(f'field.{f}.method 不合法: {method}，仅支持 sendkey/checkbox_dropdown')

        fields[f] = {'mode': m, 'values': vals, 'method': method, 'validate_options': validate_options}



    return {

        'query': query_yes,

        'export': export_yes,

        'report_title_keyword': report_title_keyword,

        'report_title_index': report_title_index,

        'query_wait_seconds': query_wait_seconds,

        'field_settle_seconds': field_settle_seconds,

        'refresh_interval': refresh_interval,

        'tree_path': tree_path,

        'max_items': max_items,

        'fields': fields,

    }







def load_standalone_config():
    """从 config.xlsx 的 standalone_fill sheet 加载独立填充配置。

    格式:
    | enabled | report_name | condition_index | field_label | field_type | field_value | match_mode | placeholder |
    enabled=yes/no  report_name=表格名称(留空通用)  condition_index=条件序号(1,2,3...)
    field_type: text/select/date  match_mode: exact/contains

    返回值: dict {rules: [{enabled, report_name, condition_index, field_label, ...}]} | None。
    """
    if not EXCEL_CONFIG_FILE.exists():
        print(f'\u2753 未找到 {EXCEL_CONFIG_FILE.name}')
        return None
    try:
        from openpyxl import load_workbook
        wb = load_workbook(EXCEL_CONFIG_FILE, data_only=True)
    except Exception as e:
        print(f'\u2753 读取 config.xlsx 失败: {e}')
        return None
    if 'standalone_fill' not in wb.sheetnames:
        print('\u2753 config.xlsx 缺少 standalone_fill sheet')
        return None
    ws = wb['standalone_fill']
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    rules = []
    for row in rows:
        if not row or len(row) < 4:
            continue
        enabled = str(row[0] or '').strip().lower() if row[0] is not None else ''
        if enabled and enabled not in ('yes', 'y', '1', 'true'):
            continue
        report_name = str(row[1] or '').strip() if len(row) > 1 else ''
        cond_idx = str(row[2] or '1').strip() if len(row) > 2 else '1'
        label = str(row[3] or '').strip() if len(row) > 3 else ''
        if not label:
            continue
        ftype = str(row[4] or 'text').strip().lower() if len(row) > 4 else 'text'
        value = str(row[5] or '').strip() if len(row) > 5 else ''
        if ftype == 'date' or label == '数据日期':
            value = _normalize_date_text(value)
        match_mode = str(row[6] or 'exact').strip().lower() if len(row) > 6 else 'exact'
        placeholder = str(row[7] or '').strip() if len(row) > 7 else ''
        try:
            cond_idx_int = int(cond_idx)
        except ValueError:
            cond_idx_int = 1
        rules.append({
            'enabled': True,
            'report_name': report_name,
            'condition_index': cond_idx_int,
            'field_label': label,
            'field_type': ftype,
            'field_value': value,
            'match_mode': match_mode,
            'placeholder': placeholder,
        })
    return {'rules': rules}


def _standalone_fill_js():
    """返回 JS 填充脚本（在指定 reportBox 内，按条件序号+字段标签定位填充）。"""
    return r"""
const reportName = arguments[0];
const conditionIndex = arguments[1];  // 1-based
const labelText = arguments[2];
const fieldType = (arguments[3] || 'text').toLowerCase();
const value = arguments[4] || '';
const matchMode = (arguments[5] || 'exact').toLowerCase();
const placeholder = (arguments[6] || '').trim();

function normText(s) {
    var x = (s === null || s === undefined) ? '' : String(s);
    return x.replace(/[:：\s]/g, '').trim().toLowerCase();
}

function getFieldText(e) {
    var t = (e.getAttribute('title') || '').trim();
    if (t) return t;
    var lbl = e.querySelector('label');
    return lbl ? (lbl.textContent || '').trim() : '';
}

function matchFieldText(target, actual) {
    var t0 = (target === null || target === undefined) ? '' : String(target).trim();
    var a0 = (actual === null || actual === undefined) ? '' : String(actual).trim();
    if (!t0 || !a0) return false;
    if (a0 === t0) return true;
    var t = normText(t0);
    var a = normText(a0);
    if (!t || !a) return false;
    return a === t || a.includes(t) || t.includes(a);
}

function findReportBoxByName(name) {
    var target = (name === null || name === undefined) ? '' : String(name).trim();
    var strongs = Array.from(document.querySelectorAll('.reportBox .reportTop strong'));
    var strong = strongs.find(function(s) {
        var t = (s.textContent || '').trim();
        return t === target || t.includes(target) || target.includes(t);
    });
    if (!strong) return null;
    return strong.closest('.reportBox');
}

function findConditionBlock(box, conditionIndex) {
    if (!box) return null;
    var condMarkers = Array.from(box.querySelectorAll('p.condition strong'));
    var marker = condMarkers.find(function(s) {
        var txt = (s.textContent || '').trim();
        return txt.indexOf('条件' + conditionIndex) >= 0;
    });
    if (!marker) return null;
    var p = marker.closest('p.condition');
    if (!p) return null;
    return p;
}

function collectConditionSiblings(p) {
    var nodes = [];
    if (!p) return nodes;
    var cur = p.nextElementSibling;
    while (cur) {
        if (cur.matches && cur.matches('p.condition')) break;
        nodes.push(cur);
        cur = cur.nextElementSibling;
    }
    return nodes;
}

function findInputInCondition(p, box, labelText, placeholder) {
    var sibs = collectConditionSiblings(p);
    // 1) placeholder 精确定位（优先）
    if (placeholder) {
        for (var i = 0; i < sibs.length; i++) {
            var inp0 = sibs[i].querySelector('input[placeholder="' + placeholder + '"]');
            if (inp0) return inp0;
        }
    }
    // 2) title/label 定位到 el-form-item，再取 input
    for (var j = 0; j < sibs.length; j++) {
        var items = Array.from(sibs[j].querySelectorAll('.el-form-item'));
        var item = items.find(function(e) { return matchFieldText(labelText, getFieldText(e)); });
        if (item) {
            var inp1 = placeholder ? item.querySelector('input[placeholder="' + placeholder + '"]') : null;
            if (!inp1) inp1 = item.querySelector('input.el-input__inner, input');
            if (inp1) return inp1;
        }
    }
    // 3) 兜底：仅在没有条件锚点时，退到 reportBox 范围
    if (!p) {
        if (placeholder) {
            var inp2 = box.querySelector('input[placeholder="' + placeholder + '"]');
            if (inp2) return inp2;
        }
        var allItemsBox = Array.from(box.querySelectorAll('.el-form-item'));
        var item2 = allItemsBox.find(function(e) { return matchFieldText(labelText, getFieldText(e)); });
        if (item2) return item2.querySelector('input.el-input__inner, input');
    }
    return null;
}

// 定位 reportBox
const box = findReportBoxByName(reportName);
if (!box) return {ok:false, msg:'reportBox_not_found'};

// 定位条件（主路径: p.condition strong=条件N）
var condP = findConditionBlock(box, conditionIndex);
var input = findInputInCondition(condP, box, labelText, placeholder);
if (!input) return {ok:false, msg:'input_not_found'};

// 清空/跳过
if (value === '' || value === '\u4e0d\u9009') {
    input.focus();
    input.value = '';
    input.dispatchEvent(new Event('input', {bubbles:true}));
    input.dispatchEvent(new Event('change', {bubbles:true}));
    return {ok:true, actual:''};
}

// select/readonly
if (fieldType === 'select' || input.readOnly) {
    function getVisibleDropdowns() {
        return Array.from(document.querySelectorAll('.el-select-dropdown, .el-select-dropdown.el-popper'))
            .filter(isVisible);
    }
    function setFrom(arr) {
        return new Set(arr);
    }
    function diffDropdowns(after, beforeSet) {
        return after.filter(function(d) { return !beforeSet.has(d); });
    }

    var beforeOpen = getVisibleDropdowns();
    var beforeSet = setFrom(beforeOpen);

    // 先展开下拉: 优先点击当前字段所属 el-select 的触发区域
    var selectRoot = input.closest('.el-select');
    if (selectRoot) {
        var trigger = selectRoot.querySelector('.el-input__suffix, .el-input__inner');
        if (trigger) {
            trigger.dispatchEvent(new MouseEvent('mousedown', {bubbles:true}));
            trigger.click();
        } else {
            selectRoot.click();
        }
    } else {
        input.click();
    }

    function pickOptionFromList(opts) {
        var opt;
        if (matchMode === 'contains') {
            opt = opts.find(function(e) { return (e.innerText || '').includes(value); });
        } else {
            opt = opts.find(function(e) { return (e.innerText || '').trim() === value; }) ||
                  opts.find(function(e) { return (e.innerText || '').includes(value); });
        }
        if (opt) {
            var span = opt.querySelector('span');
            if (span) span.click();
            opt.click();
            return true;
        }
        return false;
    }

    function isVisible(el) {
        if (!el) return false;
        var st = window.getComputedStyle(el);
        if (st.display === 'none' || st.visibility === 'hidden') return false;
        var r = el.getBoundingClientRect();
        return r.width > 0 && r.height > 0;
    }

    function getNearestDropdownOptions() {
        var inputRect = input.getBoundingClientRect();
        var dropdowns = Array.from(document.querySelectorAll('.el-select-dropdown, .el-select-dropdown.el-popper'))
            .filter(isVisible);
        if (!dropdowns.length) return [];
        var best = null;
        var bestScore = Number.POSITIVE_INFINITY;
        for (var i = 0; i < dropdowns.length; i++) {
            var r = dropdowns[i].getBoundingClientRect();
            // 优先同列/近列，再比较纵向距离
            var cx = (r.left + r.right) / 2;
            var ix = (inputRect.left + inputRect.right) / 2;
            var dx = Math.abs(cx - ix);
            var dy = Math.abs(r.top - inputRect.bottom);
            var score = dx * 10 + dy;
            if (score < bestScore) {
                bestScore = score;
                best = dropdowns[i];
            }
        }
        if (!best) return [];
        return Array.from(best.querySelectorAll('.el-select-dropdown__item, li[role="option"]'))
            .filter(function(e) { return isVisible(e) && (e.innerText || '').trim(); });
    }

    function getNewlyOpenedOptions() {
        var afterOpen = getVisibleDropdowns();
        var created = diffDropdowns(afterOpen, beforeSet);
        var pool = created.length ? created : afterOpen;
        var opts = [];
        for (var i = 0; i < pool.length; i++) {
            opts = opts.concat(Array.from(pool[i].querySelectorAll('.el-select-dropdown__item, li[role="option"]'))
                .filter(function(e) { return isVisible(e) && (e.innerText || '').trim(); }));
        }
        return opts;
    }

    function getVisibleGlobalOptions() {
        return Array.from(document.querySelectorAll('.el-select-dropdown__item, li[role="option"]'))
            .filter(function(e) { return isVisible(e) && (e.innerText || '').trim(); });
    }

    // 特殊值：全部/不选/空 -> 保持当前值（不强制改动）
    if (value === '' || value === '\u4e0d\u9009' || value === '\u5168\u90e8') {
        return {ok:true, actual:input.value || ''};
    }

    var newOpts = getNewlyOpenedOptions();
    if (newOpts.length && pickOptionFromList(newOpts)) return {ok:true, actual:input.value || ''};

    var nearOpts = getNearestDropdownOptions();
    if (nearOpts.length && pickOptionFromList(nearOpts)) return {ok:true, actual:input.value || ''};

    var globalOpts = getVisibleGlobalOptions();
    if (globalOpts.length && pickOptionFromList(globalOpts)) return {ok:true, actual:input.value || ''};

    // 重试一次展开与选择
    if (selectRoot) {
        var trigger2 = selectRoot.querySelector('.el-input__suffix, .el-input__inner');
        if (trigger2) {
            trigger2.dispatchEvent(new MouseEvent('mousedown', {bubbles:true}));
            trigger2.click();
        } else {
            selectRoot.click();
        }
    } else {
        input.click();
    }
    beforeOpen = getVisibleDropdowns();
    beforeSet = setFrom(beforeOpen);
    // 再触发一次，确保有"新打开"可识别
    if (selectRoot) {
        var trigger3 = selectRoot.querySelector('.el-input__suffix, .el-input__inner');
        if (trigger3) {
            trigger3.dispatchEvent(new MouseEvent('mousedown', {bubbles:true}));
            trigger3.click();
        }
    } else {
        input.click();
    }
    newOpts = getNewlyOpenedOptions();
    if (newOpts.length && pickOptionFromList(newOpts)) return {ok:true, actual:input.value || ''};

    nearOpts = getNearestDropdownOptions();
    if (nearOpts.length && pickOptionFromList(nearOpts)) return {ok:true, actual:input.value || ''};
    globalOpts = getVisibleGlobalOptions();
    if (globalOpts.length && pickOptionFromList(globalOpts)) return {ok:true, actual:input.value || ''};

    return {ok:false, msg:'dropdown_not_match'};
}

// text/date
input.focus();
input.value = value;
input.dispatchEvent(new Event('input', {bubbles:true}));
input.dispatchEvent(new Event('change', {bubbles:true}));
input.dispatchEvent(new KeyboardEvent('keydown', {key:'Enter', bubbles:true}));
input.dispatchEvent(new KeyboardEvent('keyup', {key:'Enter', bubbles:true}));
return {ok:true, actual:input.value || ''};
"""


def run_standalone_fill():
    """菜单8: 独立填充 — 按 standalone_fill 配置填充所有匹配的报表。"""
    if not driver_alive():
        print('\u2753 浏览器未启动，请先执行「1. 启动浏览器」')
        return

    config = load_standalone_config()
    if not config or not config['rules']:
        print('\u2753 standalone_fill 配置为空')
        return

    rules = config['rules']

    # ===== Phase 1: DOM 预检 =====
    print('\n--- 页面 DOM 预检 ---')
    diag_js = r"""
var boxes = document.querySelectorAll('.reportBox');
return Array.from(boxes).map(function(box) {
    var nameEl = box.querySelector('.reportTop strong');
    var name = nameEl ? (nameEl.textContent || '').trim() : '';
    var conds = box.querySelectorAll('p.condition');
    var condInfo = [];
    Array.from(conds).forEach(function(p) {
        var block = p.nextElementSibling;
        var flds = [];
        if (block) {
            var items = block.querySelectorAll('.el-form-item');
            Array.from(items).forEach(function(f) {
                var inp = f.querySelector('input.el-input__inner');
                flds.push({
                    title: f.getAttribute('title') || '',
                    has_input: !!inp,
                    readonly: inp ? inp.readOnly : null,
                    placeholder: inp ? (inp.placeholder || '') : ''
                });
            });
        }
        condInfo.push({
            text: (p.textContent || '').trim().slice(0, 50),
            fields: flds
        });
    });
    return {name: name, condition_count: conds.length, conditions: condInfo};
});
"""
    try:
        boxes_info = _driver.execute_script(diag_js) or []
        print(f'   页面共 {len(boxes_info)} 个报表块 (reportBox):')
        for bi, box in enumerate(boxes_info):
            print(f'    [{bi+1}] "{box.get("name", "")}"  ({box.get("condition_count", 0)} 个条件)')
            for ci, cond in enumerate(box.get('conditions', [])):
                print(f'        条件[{ci+1}]: {cond.get("text", "")}  ({len(cond.get("fields",[]))} 个字段)')
    except Exception as e:
        print(f'   \u26a0\ufe0f DOM预检异常: {e}')
        boxes_info = []

    if not boxes_info:
        print('\u2753 页面上未找到 reportBox')
        return

    # ===== Phase 2: Config 分析 =====
    print('\n--- 配置分析 ---')
    cfg_by_report = {}
    for r in rules:
        rn = r['report_name'] or '\u901a\u7528'
        cfg_by_report.setdefault(rn, []).append(r)

    print(f'   配置中已启用规则共 {len(rules)} 条，涉及 {len(cfg_by_report)} 个报表:')
    for rn, rlist in sorted(cfg_by_report.items()):
        conds_set = sorted(set(r['condition_index'] for r in rlist))
        print(f'    "{rn}": {len(rlist)} 条规则, {len(conds_set)} 个条件 (序号 {conds_set})')

    # ===== Phase 3: Config vs DOM 对比 =====
    print('\n--- 配置与页面匹配检查 ---')
    box_map = {b.get('name', ''): b for b in boxes_info}
    unmatched = []
    for rn in cfg_by_report:
        if rn == '\u901a\u7528':
            continue
        if rn not in box_map:
            print(f'  \u26a0\ufe0f 配置中的 "{rn}" 未在页面上找到对应的 reportBox，将跳过')
            unmatched.append(rn)
    if unmatched:
        print(f'  共 {len(unmatched)} 个报表匹配失败，但程序将继续执行')

    # ===== Phase 4: 按 reportBox 分组填充 =====
    print('\n--- 开始填充 ---')

    run_id = f"standalone_fill_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:8]}"
    run_dir = OUTPUT / run_id
    log_file = run_dir / "run_log.jsonl"
    run_dir.mkdir(parents=True, exist_ok=True)

    fill_js = _standalone_fill_js()
    read_js = r"""
function normText(s) {
    var x = (s === null || s === undefined) ? '' : String(s);
    return x.replace(/[:：\s]/g, '').trim().toLowerCase();
}
function getFieldText(e) {
    var t = (e.getAttribute('title') || '').trim();
    if (t) return t;
    var lbl = e.querySelector('label');
    return lbl ? (lbl.textContent || '').trim() : '';
}
function matchFieldText(target, actual) {
    var t0 = (target === null || target === undefined) ? '' : String(target).trim();
    var a0 = (actual === null || actual === undefined) ? '' : String(actual).trim();
    if (!t0 || !a0) return false;
    if (a0 === t0) return true;
    var t = normText(t0);
    var a = normText(a0);
    if (!t || !a) return false;
    return a === t || a.includes(t) || t.includes(a);
}
function findReportBoxByName(name) {
    var target = (name === null || name === undefined) ? '' : String(name).trim();
    var strongs = Array.from(document.querySelectorAll('.reportBox .reportTop strong'));
    var strong = strongs.find(function(s) {
        var t = (s.textContent || '').trim();
        return t === target || t.includes(target) || target.includes(t);
    });
    if (!strong) return null;
    return strong.closest('.reportBox');
}
function findConditionBlock(box, conditionIndex) {
    if (!box) return null;
    var condMarkers = Array.from(box.querySelectorAll('p.condition strong'));
    var marker = condMarkers.find(function(s) {
        var txt = (s.textContent || '').trim();
        return txt.indexOf('条件' + conditionIndex) >= 0;
    });
    if (!marker) return null;
    var p = marker.closest('p.condition');
    if (!p) return null;
    return p;
}
function collectConditionSiblings(p) {
    var nodes = [];
    if (!p) return nodes;
    var cur = p.nextElementSibling;
    while (cur) {
        if (cur.matches && cur.matches('p.condition')) break;
        nodes.push(cur);
        cur = cur.nextElementSibling;
    }
    return nodes;
}
function findInputInCondition(p, box, labelText, placeholder) {
    var sibs = collectConditionSiblings(p);
    if (placeholder) {
        for (var i = 0; i < sibs.length; i++) {
            var inp0 = sibs[i].querySelector('input[placeholder="' + placeholder + '"]');
            if (inp0) return inp0;
        }
    }
    for (var j = 0; j < sibs.length; j++) {
        var items = Array.from(sibs[j].querySelectorAll('.el-form-item'));
        var item = items.find(function(e) { return matchFieldText(labelText, getFieldText(e)); });
        if (item) {
            var inp1 = placeholder ? item.querySelector('input[placeholder="' + placeholder + '"]') : null;
            if (!inp1) inp1 = item.querySelector('input.el-input__inner, input');
            if (inp1) return inp1;
        }
    }
    if (!p) {
        if (placeholder) {
            var inp2 = box.querySelector('input[placeholder="' + placeholder + '"]');
            if (inp2) return inp2;
        }
        var allItemsBox = Array.from(box.querySelectorAll('.el-form-item'));
        var item2 = allItemsBox.find(function(e) { return matchFieldText(labelText, getFieldText(e)); });
        if (item2) return item2.querySelector('input.el-input__inner, input');
    }
    return null;
}
var box = findReportBoxByName(arguments[0]);
if (!box) return null;
var ps = Array.from(box.querySelectorAll('p.condition'));
var p = ps[arguments[1] - 1];
var ph = (arguments[3] || '').trim();
var condP = findConditionBlock(box, arguments[1]) || p || null;
var inp = findInputInCondition(condP, box, arguments[2], ph);
return inp ? (inp.value || '') : null;
"""

    total_ok = 0
    total_fail = 0
    total_skip = 0

    # 拆分通用规则和指定报表规则
    specific_rules = {}
    general_rules = []
    for rn, rlist in cfg_by_report.items():
        if rn == '\u901a\u7528':
            general_rules = rlist
        else:
            specific_rules[rn] = rlist

    # 构建填充列表: (report_name, box_info, rules_list)
    fill_queue = []

    # 指定报表
    for rn, rlist in sorted(specific_rules.items()):
        if rn in unmatched:
            continue
        box = box_map.get(rn)
        if not box:
            continue
        combined = list(rlist)
        if general_rules:
            combined.extend(general_rules)
        fill_queue.append((rn, box, combined))

    # 如果只有通用规则，应用到所有报表
    if not fill_queue and general_rules:
        for rn, box in sorted(box_map.items(), key=lambda x: x[0]):
            if rn:
                fill_queue.append((rn, box, list(general_rules)))

    if not fill_queue:
        print('\u2753 没有可填充的报表（配置中的报表名与页面不匹配）')
        return

    interrupted = False
    interrupt_ts = None

    try:
        for rn, box, rlist in fill_queue:
            cond_count = box.get('condition_count', 0)
            print(f'\n=== 报表: {rn} ({cond_count} 个条件) ===')
            try:
                for cidx, c in enumerate(box.get('conditions', []), start=1):
                    titles = [str(f.get('title', '')).strip() for f in c.get('fields', []) if str(f.get('title', '')).strip()]
                    if titles:
                        print(f'    条件{cidx}字段: {titles}')
            except Exception:
                pass

            # 按条件序号分组
            by_cond = {}
            for r in rlist:
                by_cond.setdefault(r['condition_index'], []).append(r)

            max_cond = max(by_cond.keys())
            if max_cond > cond_count:
                print(f'  \u26a0\ufe0f 配置中的最大条件序号({max_cond})超过页面实际条件数({cond_count})，超出部分将跳过')

            for cond_num in sorted(by_cond.keys()):
                if cond_num > cond_count:
                    print(f'  [条件{cond_num}] 跳过（页面只有 {cond_count} 个条件）')
                    continue

                fields = by_cond[cond_num]
                print(f'  [条件{cond_num}] ({len(fields)} 个字段)')

                for field in fields:
                    label = field['field_label']
                    ftype = field['field_type']
                    value = field['field_value']
                    if ftype == 'date' or label == '数据日期':
                        value = _normalize_date_text(value)
                    match_mode = field['match_mode']
                    placeholder = field['placeholder']
                    value_str = str(value or '').strip()

                    # 语义约定:
                    # - 空值: 不改该配置（跳过）
                    # - 不选: 清空该单元格
                    if value_str == '':
                        print(f'    ↷ {label} 空值，按配置不改动')
                        _append_jsonl(log_file, {
                            "ts": datetime.now().isoformat(timespec='seconds'),
                            "type": "field_skip",
                            "report_name": rn,
                            "condition": cond_num,
                            "field": label,
                            "target": "",
                            "actual": str(_driver.execute_script(read_js, rn, cond_num, label, placeholder) or ''),
                            "reason": "empty_value_noop"
                        })
                        total_skip += 1
                        continue

                    # 预读当前值，已匹配则跳过
                    current_val = _driver.execute_script(read_js, rn, cond_num, label, placeholder)
                    current_str = str(current_val or '').strip()
                    target_str = value_str if value_str != '\u4e0d\u9009' else ''
                    already = (current_str == target_str) or (value_str == '\u4e0d\u9009' and not current_val)

                    if already:
                        print(f'    \u2713 {label} \u2192 {value_str if value_str else "清空"} (已匹配，跳过)')
                        _append_jsonl(log_file, {
                            "ts": datetime.now().isoformat(timespec='seconds'),
                            "type": "field_skip",
                            "report_name": rn,
                            "condition": cond_num,
                            "field": label,
                            "target": target_str,
                            "actual": current_str,
                            "reason": "already_matched"
                        })
                        total_skip += 1
                        continue

                    # 填充
                    result = _driver.execute_script(fill_js, rn, cond_num, label, ftype, value_str, match_mode, placeholder)
                    time.sleep(0.2)

                    if result and result.get('ok'):
                        actual = result.get('actual', '')
                        print(f'    \u2713 {label} \u2192 {value_str if value_str != "不选" else "清空"} (实际: {actual})')
                        total_ok += 1
                    else:
                        msg = result.get('msg', 'js_error') if result else 'exec_error'
                        print(f'    \u2717 {label} \u2192 {value_str if value_str != "不选" else "清空"} | {msg}')
                        total_fail += 1

                    _append_jsonl(log_file, {
                        "ts": datetime.now().isoformat(timespec='seconds'),
                        "type": "standalone_field",
                        "report_name": rn,
                        "condition": cond_num,
                        "field": label,
                        "target": target_str,
                        "actual": result.get('actual', '') if result else '',
                        "ok": bool(result and result.get('ok')),
                        "msg": result.get('msg', '') if result else 'exec_error',
                    })

    except KeyboardInterrupt:
        interrupted = True
        interrupt_ts = datetime.now().isoformat(timespec='seconds')
        _append_jsonl(log_file, {
            "ts": interrupt_ts,
            "type": "run_interrupted",
            "ok": total_ok,
            "fail": total_fail,
            "skip": total_skip,
        })
        print('\n⚠️ 已收到停止指令，当前报表组批量修改已中断，返回主菜单')

    _append_jsonl(log_file, {
        "ts": datetime.now().isoformat(timespec='seconds'),
        "type": "run_summary",
        "ok": total_ok,
        "fail": total_fail,
        "skip": total_skip,
        "interrupted": interrupted,
        "interrupt_ts": interrupt_ts,
    })
    status_text = "已中断" if interrupted else "完成"
    print(f'\n\u2705 报表组批量修改{status_text}: 成功 {total_ok}, 失败 {total_fail}, 跳过 {total_skip}')
    print(f'\U0001f4dd \u65e5\u5fd7: {run_dir}')


def _append_jsonl(path, data):
    """追加一行 JSON 到日志文件。"""
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, 'a', encoding='utf-8') as f:
            f.write(json.dumps(data, ensure_ascii=False) + '\n')
    except Exception:
        pass


def _download_snapshot():
    """记录当前下载目录文件状态，用于导出后判断新增文件。"""
    download_dir = _resolve_download_dir()
    snapshot = {}
    try:
        for p in download_dir.iterdir():
            if p.is_file():
                st = p.stat()
                snapshot[str(p)] = {
                    "name": p.name,
                    "path": str(p),
                    "size": st.st_size,
                    "mtime": st.st_mtime,
                    "is_partial": p.suffix.lower() == ".crdownload",
                }
    except Exception:
        pass
    return download_dir, snapshot


def _inspect_download_file(path):
    """读取文件头，判断下载文件是否像有效 Excel 或 HTML 响应。"""
    p = Path(path)
    info = {
        "path": str(p),
        "extension": p.suffix.lower(),
        "file_kind": "unknown",
        "valid_excel": None,
    }
    try:
        head = p.read_bytes()[:512]
    except Exception as e:
        info["error"] = str(e)
        return info

    head_lstrip = head.lstrip()
    head_hex = head[:16].hex()
    info["head_hex"] = head_hex

    if head.startswith(b"PK"):
        info["file_kind"] = "zip_xlsx_like"
        info["valid_excel"] = p.suffix.lower() in (".xlsx", ".xlsm", ".xltx", ".xltm")
    elif head.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        info["file_kind"] = "ole_xls_like"
        info["valid_excel"] = p.suffix.lower() == ".xls"
    elif head_lstrip[:5].lower() == b"<html" or head_lstrip[:9].lower().startswith(b"<!doctype"):
        info["file_kind"] = "html_response"
        info["valid_excel"] = False
    elif head_lstrip[:1] == b"<":
        info["file_kind"] = "xml_or_html_like"
        info["valid_excel"] = False
    else:
        info["file_kind"] = "unknown_binary"
        info["valid_excel"] = False if p.suffix.lower() in (".xlsx", ".xls") else None

    if p.suffix.lower() == ".xlsx" and not head.startswith(b"PK"):
        info["validation_error"] = "invalid_xlsx_header"
    if p.suffix.lower() == ".xls" and not (
        head.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1") or head_lstrip[:5].lower() == b"<html"
    ):
        info["validation_error"] = "unexpected_xls_header"
    return info


def _wait_download_delta(before, timeout=20.0):
    """等待下载目录出现新增或更新文件，并返回可审计信息。"""
    download_dir = _resolve_download_dir()
    started = time.time()
    last_files = []
    while time.time() - started < timeout:
        current = {}
        try:
            for p in download_dir.iterdir():
                if not p.is_file():
                    continue
                st = p.stat()
                key = str(p)
                prev = before.get(key)
                changed = prev is None or prev.get("size") != st.st_size or prev.get("mtime") != st.st_mtime
                if changed:
                    inspect = _inspect_download_file(p) if p.suffix.lower() != ".crdownload" else {"file_kind": "partial_download", "valid_excel": False}
                    current[key] = {
                        "name": p.name,
                        "path": key,
                        "size": st.st_size,
                        "mtime": datetime.fromtimestamp(st.st_mtime).isoformat(timespec='seconds'),
                        "is_partial": p.suffix.lower() == ".crdownload",
                        "extension": p.suffix.lower(),
                        "inspection": inspect,
                    }
        except Exception as e:
            return {"ok": False, "download_dir": str(download_dir), "error": str(e), "files": last_files}

        last_files = list(current.values())
        complete = [f for f in last_files if not f.get("is_partial") and f.get("size", 0) > 0]
        partial = [f for f in last_files if f.get("is_partial")]
        if complete and not partial:
            return {"ok": True, "download_dir": str(download_dir), "files": complete, "partial_files": partial}
        time.sleep(0.5)

    return {"ok": False, "download_dir": str(download_dir), "files": last_files, "timeout": True}


def _rename_download_files(download_res, seq, scene, strategy):
    """给本轮导出的文件加序号，避免 FineReport 同名文件覆盖。"""
    renamed = []
    files = download_res.get("files", []) if isinstance(download_res, dict) else []
    for idx, info in enumerate(files, 1):
        if info.get("is_partial"):
            renamed.append(info)
            continue
        src = Path(info.get("path", ""))
        if not src.exists() or not src.is_file():
            info["rename_error"] = "source_missing"
            renamed.append(info)
            continue

        safe_original = _safe_filename(src.stem, maxlen=80) or "download"
        suffix = 1
        target = src.with_name(f"{safe_original}_{suffix}{src.suffix}")
        while target.exists() and target.resolve() != src.resolve():
            suffix += 1
            target = src.with_name(f"{safe_original}_{suffix}{src.suffix}")

        try:
            if target.resolve() != src.resolve():
                src.replace(target)
            st = target.stat()
            updated = dict(info)
            updated.update({
                "original_name": info.get("name"),
                "original_path": info.get("path"),
                "name": target.name,
                "path": str(target),
                "size": st.st_size,
                "mtime": datetime.fromtimestamp(st.st_mtime).isoformat(timespec='seconds'),
                "inspection": _inspect_download_file(target),
                "renamed": True,
            })
            renamed.append(updated)
        except Exception as e:
            info["rename_error"] = str(e)
            renamed.append(info)

    if isinstance(download_res, dict):
        download_res["files"] = renamed
        download_res["renamed"] = True
    return download_res


def driver_alive():
    """检查浏览器是否存活。"""
    global _driver
    if _driver is None:
        return False
    try:
        _driver.current_url
        return True
    except Exception:
        try:
            _driver.quit()
        except Exception:
            pass
        _driver = None
        return False


def _check_page_alive():
    """快速检查浏览器 tab 是否存活（非崩溃状态）。

    通过执行简单 JS 检测页面是否可交互，同时排除 Chrome 崩溃页面
    和登录状态过期弹窗。
    调用方: run_matrix_export 主循环（每轮开始前），_refresh_and_reestablish（刷新后验证）。
    返回值: bool。
    """
    global _driver
    if _driver is None:
        return False
    try:
        # 先检查 WebDriver 连接是否正常
        _driver.current_url
        # 再检查页面是否可执行 JS
        ready = _driver.execute_script("return document.readyState || ''")
        if not ready:
            return False
        # 排除 Chrome 崩溃页面关键字
        body_text = _driver.execute_script(
            "return (document.body && document.body.innerText || '').substring(0, 200)"
        ) or ""
        crash_keywords = ['Aw, snap!', '无法打开此网页', 'ERR_', '网页崩溃', '页面崩溃']
        for kw in crash_keywords:
            if kw in body_text:
                return False
        # 排除登录状态过期弹窗（页面存活但无法操作，需要重新登录）
        expired = _driver.execute_script("""
var boxes = document.querySelectorAll('.el-message-box__message p');
for (var i = 0; i < boxes.length; i++) {
    var t = (boxes[i].textContent || '').trim();
    if (t.indexOf('登录状态已过期') >= 0 || t.indexOf('登录状态已过期') >= 0) return true;
}
var wrappers = document.querySelectorAll('.el-message-box__wrapper');
for (var j = 0; j < wrappers.length; j++) {
    var s = window.getComputedStyle(wrappers[j]);
    if (s.display === 'none' || s.visibility === 'hidden') continue;
    var msg = wrappers[j].querySelector('.el-message-box__message');
    if (msg && (msg.textContent || '').indexOf('登录状态已过期') >= 0) return true;
}
return false;
""")
        if expired:
            return False
        return True
    except Exception:
        return False


def _save_iframe_src():
    """尝试保存当前 iframe 的 src 地址，供刷新后重定位使用。"""
    global _saved_iframe_src
    try:
        _saved_iframe_src = _driver.execute_script(
            "try { return window.frameElement ? window.frameElement.src : window.location.href; } catch(e) { return window.location.href; }"
        ) or ""
    except Exception:
        _saved_iframe_src = ""


def _navigate_saved_url():
    """使用保存的页面 URL 重新导航，返回是否成功。"""
    global _saved_page_url, _driver
    if not _saved_page_url:
        return False
    if _driver is None:
        return False
    try:
        _driver.get(_saved_page_url)
        # 等待页面基本加载完成
        for _ in range(30):  # 最多等 15 秒
            try:
                ready = _driver.execute_script("return document.readyState || ''")
                if ready == 'complete':
                    time.sleep(1.0)  # 额外等 1 秒确保 SPA 渲染
                    return True
            except Exception:
                pass
            time.sleep(0.5)
        return False
    except Exception as e:
        print(f'  ⚠️ 导航至保存的 URL 失败: {e}')
        return False


def _refresh_and_reestablish(scene, plan=None):
    """刷新页面并重新切入报表 iframe。

    策略: 优先使用 _navigate_saved_url() 重新导航（比 driver.refresh() 更可靠），
    失败时回退到 driver.refresh()。重新导航成功后调用 _switch_to_report_frame()
    重建 iframe 上下文。

    调用方: run_matrix_export 主循环（定时刷新、错误恢复）。
    返回值: bool — 是否成功恢复页面和 iframe 上下文。
    """
    global _driver
    if _driver is None:
        return False

    # 先切回默认上下文，避免在已崩溃的 iframe 中操作
    try:
        _driver.switch_to.default_content()
    except Exception:
        pass

    # 方案 A: 优先 F5 刷新（和手动刷新一致）
    navigated = False
    try:
        _driver.refresh()
        navigated = True
        print('  🔄 F5 刷新...')
    except Exception:
        pass

    # 方案 B: 回退到导航保存的 URL
    if not navigated and _saved_page_url:
        try:
            _driver.get(_saved_page_url)
            navigated = True
            print('  🔄 导航到保存的 URL...')
        except Exception:
            pass

    if not navigated:
        print('  ✗ 页面刷新失败（navigate 和 refresh 均不可用）')
        return False

    # 等待页面加载
    time.sleep(2.0)
    for _ in range(20):
        try:
            ready = _driver.execute_script("return document.readyState || ''")
            if ready == 'complete':
                break
        except Exception:
            pass
        time.sleep(0.5)

    # 等待额外稳定时间
    time.sleep(1.5)

    tree_path = plan.get('tree_path', '') if plan else ''

    # ── 恢复策略：先树导航（登录通常还在），失败才处理登录 ──
    if tree_path:
        # 快速检测是否掉登录
        need_login = _driver.execute_script("""
// 检查登录过期弹窗
var boxes = document.querySelectorAll('.el-message-box__message p');
for (var i = 0; i < boxes.length; i++) {
    var t = (boxes[i].textContent || '').trim();
    if (t.indexOf('登录状态已过期') >= 0) return true;
}
// 检查是否直接跳到了登录页
var form = document.querySelector('form.login-form');
if (form && window.getComputedStyle(form).display !== 'none') return true;
return false;
""")
        if need_login:
            print('  ⚠️ 检测到登录异常，先处理登录...')
            if not _ensure_logged_in():
                print('  ✗ 登录恢复失败')
                return False

        # 先尝试树导航（大多数情况登录还在，直接走这步）
        if not _navigate_tree_to_report(tree_path):
            # 树导航失败可能是登录过期导致页面跳转
            print('  ⚠️ 树导航失败，尝试处理后重试...')
            if _ensure_logged_in():
                _navigate_tree_to_report(tree_path)
    else:
        # 无 tree_path，兜底检查登录状态
        if not _ensure_logged_in():
            print('  ✗ 登录状态恢复失败')
            return False

    # 重新切入报表 iframe
    if scene and plan is not None:
        ok = _switch_to_report_frame(scene, plan)
        if ok:
            _save_iframe_src()
        return ok

    # 无 scene 信息时，只确认页面存活
    return _check_page_alive()


def _find_linux_chrome_binary():
    """返回 Linux 上可用的 Chrome/Chromium 路径。"""
    configured = os.environ.get('CHROMEAUTORUN_CHROME_BINARY', '').strip()
    if configured:
        path = Path(configured).expanduser()
        return str(path) if path.is_file() else None

    for name in ('google-chrome', 'google-chrome-stable', 'chromium', 'chromium-browser', 'chrome'):
        path = shutil.which(name)
        if path:
            return path
    return None


def _find_linux_chromedriver():
    """依次查找用户指定、随 DEB 打包、系统 PATH 中的 ChromeDriver。"""
    configured = os.environ.get('SE_CHROMEDRIVER', '').strip()
    if configured:
        return str(Path(configured).expanduser())

    bundled = APP_DIR / 'drivers' / 'chromedriver'
    if bundled.is_file():
        return str(bundled)

    return shutil.which('chromedriver')


def launch_browser():
    """菜单1: 启动 Chrome 浏览器。Windows/Linux 均优先使用随包驱动。"""
    global _driver
    if driver_alive():
        print('⚠️  浏览器已启动')
        return

    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service

    if os.name == 'nt':
        chrome_driver = BASE_DIR / 'chromedriver.exe'
        if not chrome_driver.exists():
            print(f'❌ 未找到 chromedriver.exe: {chrome_driver}')
            return
        service = Service(str(chrome_driver))
    else:
        chrome_driver = _find_linux_chromedriver()
        if chrome_driver:
            service = Service(chrome_driver)
            print(f'ℹ️ 使用 ChromeDriver: {chrome_driver}')
        else:
            # 源码运行时若没有本地驱动，才交给 Selenium Manager 自动下载。
            service = Service()
            print('ℹ️ 未找到本地 chromedriver，将由 Selenium Manager 自动匹配驱动（需要联网）')

    options = Options()
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_experimental_option('excludeSwitches', ['enable-automation', 'enable-logging'])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument('--no-first-run')
    options.add_argument('--disable-default-apps')
    options.add_argument('--start-maximized')
    options.add_argument('--log-level=3')
    options.add_argument(f'--user-data-dir={CHROME_PROFILE_DIR}')
    if os.name != 'nt':
        chrome_binary = _find_linux_chrome_binary()
        if chrome_binary:
            options.binary_location = chrome_binary
            print(f'ℹ️ 使用浏览器: {chrome_binary}')
        else:
            print('⚠️ 未在 PATH 中找到 Chrome/Chromium；请先安装浏览器，或设置 CHROMEAUTORUN_CHROME_BINARY')
    download_dir = str(_resolve_download_dir().resolve())
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "download.extensions_to_open": "",
        "download.open_pdf_in_system_reader": False,
        "profile.default_content_settings.popups": 0,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "profile.default_content_setting_values.notifications": 2,
        "profile.content_settings.exceptions.automatic_downloads": {
            "http://*/*,*": {"setting": 1},
            "https://*/*,*": {"setting": 1},
        },
        "safebrowsing.enabled": True,
    }
    options.add_experimental_option('prefs', prefs)

    try:
        _driver = webdriver.Chrome(service=service, options=options)
        try:
            _driver.execute_cdp_cmd('Browser.setDownloadBehavior', {
                'behavior': 'allow',
                'downloadPath': download_dir,
                'eventsEnabled': True,
            })
        except Exception:
            try:
                _driver.execute_cdp_cmd('Page.setDownloadBehavior', {
                    'behavior': 'allow',
                    'downloadPath': download_dir,
                })
            except Exception:
                pass
        _driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
            'source': 'Object.defineProperty(navigator, "webdriver", {get: () => undefined})'
        })
        print(f'✅ 浏览器已启动（下载目录: {download_dir}，配置目录: {CHROME_PROFILE_DIR}）')
    except Exception as e:
        print(f'❌ 启动浏览器失败: {e}')
        if os.name != 'nt':
            print('   请确认 Chrome/Chromium 已安装且可在图形桌面运行，并检查 ChromeDriver 与浏览器版本是否匹配。')
        _driver = None


def _restart_browser_session():
    """关闭当前 WebDriver 浏览器并重新启动。"""
    global _driver
    print('  🔁 重启浏览器会话...')
    if _driver is not None:
        try:
            _driver.quit()
        except Exception as e:
            print(f'  ⚠️ 关闭旧浏览器失败，继续重启: {e}')
        finally:
            _driver = None
    time.sleep(1.0)
    launch_browser()
    return driver_alive()


def _login_wait_css(selector, timeout=15):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    return WebDriverWait(_driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
    )


def _login_read_value(element):
    return _driver.execute_script("return arguments[0] ? (arguments[0].value || '') : '';", element) or ''


def _login_set_value(element, value):
    _driver.execute_script("""
const el = arguments[0];
const nextValue = arguments[1] || '';
if (!el) return false;
el.focus();
const proto = el.tagName === 'TEXTAREA'
    ? window.HTMLTextAreaElement.prototype
    : window.HTMLInputElement.prototype;
const descriptor = Object.getOwnPropertyDescriptor(proto, 'value');
if (descriptor && descriptor.set) {
    descriptor.set.call(el, nextValue);
} else {
    el.value = nextValue;
}
el.dispatchEvent(new Event('input', { bubbles: true }));
el.dispatchEvent(new Event('change', { bubbles: true }));
el.dispatchEvent(new Event('blur', { bubbles: true }));
return true;
""", element, value)


def _login_fill_text(selector, value, field_name, timeout=15):
    element = _login_wait_css(selector, timeout=timeout)
    _driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
    disabled = bool(_driver.execute_script("return !!arguments[0].disabled;", element))
    readonly = bool(_driver.execute_script("return !!arguments[0].readOnly;", element))

    if disabled:
        raise RuntimeError(f'{field_name} 输入框处于 disabled 状态')

    if not readonly:
        try:
            element.click()
            element.clear()
            element.send_keys(value)
        except Exception:
            _login_set_value(element, value)
    else:
        _login_set_value(element, value)

    actual = _login_read_value(element).strip()
    if actual != value:
        _login_set_value(element, value)
        actual = _login_read_value(element).strip()
    if actual != value:
        raise RuntimeError(f'{field_name} 回读不一致: target={value}, actual={actual}')
    return element


def _login_select_username(selector, username, timeout=20):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait

    wait = WebDriverWait(_driver, timeout)
    input_el = _login_wait_css(selector, timeout=timeout)
    _driver.execute_script("arguments[0].scrollIntoView({block:'center'});", input_el)

    if _login_read_value(input_el).strip() == username:
        return input_el

    readonly = bool(_driver.execute_script("return !!arguments[0].readOnly;", input_el))
    if not readonly:
        return _login_fill_text(selector, username, '用户名', timeout=timeout)

    trigger = _driver.execute_script("return arguments[0].closest('.el-select') || arguments[0];", input_el)

    for _ in range(3):
        try:
            trigger.click()
        except Exception:
            _driver.execute_script("arguments[0].click();", trigger)

        def _visible_options(driver):
            options = driver.find_elements(By.CSS_SELECTOR, '.el-select-dropdown__item, li[role=\"option\"]')
            return [opt for opt in options if opt.is_displayed() and opt.text.strip()]

        options = wait.until(_visible_options)
        for option in options:
            text = option.text.strip()
            if text == username or username in text:
                try:
                    option.click()
                except Exception:
                    _driver.execute_script("arguments[0].click();", option)
                time.sleep(0.5)
                actual = _login_read_value(input_el).strip()
                if actual == username or username in actual:
                    return input_el
        time.sleep(1)

    actual = _login_read_value(input_el).strip()
    raise RuntimeError(f'用户名下拉未匹配到目标项: target={username}, actual={actual}')


def auto_login():
    """菜单4: 导航至金数系统并自动登录。"""
    if not driver_alive():
        print('❌ 浏览器未启动，请先执行「1. 启动浏览器」')
        return

    cfg = sync_config_json_from_excel()
    if not cfg:
        print('❌ 无法加载登录配置')
        return

    url = cfg.get('url', '').strip()
    if not url:
        print('❌ config.xlsx login sheet 中 url 为空')
        return

    credit_code = cfg.get('credit_code', '').strip()
    username = cfg.get('username', '').strip()
    password = cfg.get('password', '').strip()

    if not all([credit_code, username, password]):
        print('❌ config.xlsx login sheet 缺少 credit_code/username/password')
        return

    try:
        _driver.get(url)
        time.sleep(2)
    except Exception as e:
        print(f'❌ 导航失败: {e}')
        return

    try:
        sel = DEFAULT_LOGIN_SELECTORS
        login_url = _driver.current_url

        # 社会信用代码
        _login_fill_text(sel['credit_code_selector'], credit_code, '社会信用代码')
        time.sleep(1)

        # 用户名（该页通常是 readonly 下拉）
        _login_select_username(sel['username_selector'], username)

        # 密码
        _login_fill_text(sel['password_selector'], password, '密码')

        # 登录按钮
        btn = _login_wait_css(sel['submit_selector'], timeout=15)
        try:
            btn.click()
        except Exception:
            _driver.execute_script("arguments[0].click();", btn)

        # 等待成功标志或 URL 变化
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.common.by import By
        try:
            WebDriverWait(_driver, 30).until(
                lambda d: (
                    d.current_url != login_url
                    or len(d.find_elements(By.CSS_SELECTOR, sel['success_indicator'])) > 0
                    or len(d.find_elements(By.CSS_SELECTOR, 'form.login-form')) == 0
                )
            )
            print('✅ 登录成功')
            global _saved_page_url
            _saved_page_url = _driver.current_url
        except Exception:
            print('⚠️  登录后未检测到成功标志，可能页面跳转较慢或选择器不匹配')
    except Exception as e:
        print(f'❌ 登录过程异常: {e}')


def _detect_login_state():
    """统一检测当前页面登录状态。门户首页有报表入口也视为已登录。"""
    global _driver
    default = {
        "expired": False,
        "on_login_page": False,
        "has_logged_in_ui": False,
        "has_portal_entry": False,
        "has_report_tree": False,
        "logged_in": False,
    }
    if _driver is None:
        return default
    try:
        _driver.switch_to.default_content()
    except Exception:
        pass
    try:
        state = _driver.execute_script("""
var state = {
    expired: false,
    on_login_page: false,
    has_logged_in_ui: false,
    has_portal_entry: false,
    has_report_tree: false
};

var boxes = document.querySelectorAll('.el-message-box__message p');
for (var i = 0; i < boxes.length; i++) {
    var t = (boxes[i].textContent || '').trim();
    if (t.indexOf('登录状态已过期') >= 0) state.expired = true;
}
var wrappers = document.querySelectorAll('.el-message-box__wrapper');
for (var j = 0; j < wrappers.length; j++) {
    var style = window.getComputedStyle(wrappers[j]);
    if (style.display === 'none' || style.visibility === 'hidden') continue;
    var msg = wrappers[j].querySelector('.el-message-box__message');
    if (msg && (msg.textContent || '').indexOf('登录状态已过期') >= 0) state.expired = true;
}

var form = document.querySelector('form.login-form');
if (form) {
    var formStyle = window.getComputedStyle(form);
    state.on_login_page = formStyle.display !== 'none' && formStyle.visibility !== 'hidden';
}

state.has_report_tree = !!document.querySelector('.el-tree, .org_tree');
state.has_logged_in_ui = !!document.querySelector('#app .avue-layout, #app .el-menu, .main-content, .avue-left, .avue-tags, .el-tabs__header');

var links = document.querySelectorAll('a, span, div');
for (var k = 0; k < links.length; k++) {
    if (links[k].offsetParent === null) continue;
    var txt = (links[k].textContent || '').trim();
    if (txt === '报表服务中心' || txt.indexOf('报表服务中心') >= 0) {
        state.has_portal_entry = true;
        break;
    }
}
return state;
""")
        if not isinstance(state, dict):
            state = {}
    except Exception:
        state = {}

    merged = dict(default)
    merged.update({k: bool(v) for k, v in state.items() if k in merged})
    merged["logged_in"] = (
        not merged["expired"]
        and not merged["on_login_page"]
        and (
            merged["has_logged_in_ui"]
            or merged["has_portal_entry"]
            or merged["has_report_tree"]
        )
    )
    return merged


def _navigate_to_report_page():
    """从门户首页导航到报表树页面。

    步骤: 门户首页 → 点击"报表服务中心" → 展开"数据共享" → 点击"共享数据查看"
    登录后不会直接到报表页，需要走这几步。
    如果已在报表页则跳过。

    调用方: _ensure_logged_in（重新登录后），test_portal_navigate（测试菜单）。
    返回值: bool。
    """
    global _driver, _saved_page_url
    if _driver is None:
        return False

    # 检查是否已经在报表页面（有 el-tree 或 org_tree）
    already_there = _driver.execute_script("""
var tree = document.querySelector('.el-tree, .org_tree');
if (tree) return true;
// Also check if we have the app layout (sidebar + content)
var sidebar = document.querySelector('.avue-left, .el-menu');
var tabs = document.querySelector('.avue-tags, .el-tabs__header');
if (sidebar && tabs) return true;
return false;
""")
    if already_there:
        return True  # 已经在报表页面

    # ── Step 1: 门户首页 → 点击 "报表服务中心" ──
    on_portal = _driver.execute_script("""
var links = document.querySelectorAll('a');
for (var i = 0; i < links.length; i++) {
    var span = links[i].querySelector('span');
    if (span && (span.textContent || '').trim() === '报表服务中心') {
        return true;
    }
}
return false;
""")
    if on_portal:
        print('  📍 在门户首页，点击「报表服务中心」...')
        clicked = _driver.execute_script("""
var links = document.querySelectorAll('a');
for (var i = 0; i < links.length; i++) {
    var span = links[i].querySelector('span');
    if (span && (span.textContent || '').trim() === '报表服务中心') {
        links[i].click();
        return true;
    }
}
return false;
""")
        if not clicked:
            print('  ⚠️ 未找到「报表服务中心」链接')
            return False

        # 等待页面跳转
        time.sleep(2)
        for _ in range(20):
            ready = _driver.execute_script("return document.readyState || ''")
            if ready == 'complete':
                break
            time.sleep(0.5)
        time.sleep(1)

    # ── Step 2: 展开 "数据共享" 侧边栏菜单 ──
    has_sidebar = _driver.execute_script("""
var submenus = document.querySelectorAll('.el-submenu');
for (var i = 0; i < submenus.length; i++) {
    var title = submenus[i].querySelector('.el-submenu__title span');
    if (title && (title.textContent || '').trim() === '数据共享') {
        return true;
    }
}
return false;
""")
    if has_sidebar:
        print('  📍 在报表管理页面，展开「数据共享」...')
        _driver.execute_script("""
var submenus = document.querySelectorAll('.el-submenu');
for (var i = 0; i < submenus.length; i++) {
    var title = submenus[i].querySelector('.el-submenu__title span');
    if (title && (title.textContent || '').trim() === '数据共享') {
        // Expand if not already open
        if (!submenus[i].classList.contains('is-active') && !submenus[i].classList.contains('is-opened')) {
            var titleDiv = submenus[i].querySelector('.el-submenu__title');
            if (titleDiv) titleDiv.click();
        }
        break;
    }
}
""")
        time.sleep(0.8)

    # ── Step 3: 点击 "共享数据查看" ──
        print('  📍 点击「共享数据查看」...')
        clicked2 = _driver.execute_script("""
var items = document.querySelectorAll('.el-menu-item');
for (var i = 0; i < items.length; i++) {
    var span = items[i].querySelector('span');
    var txt = span ? (span.textContent || '').trim() : (items[i].textContent || '').trim();
    if (txt === '共享数据查看') {
        items[i].click();
        return true;
    }
}
// Fallback: look for any element with this text
var all = document.querySelectorAll('li, span, div');
for (var j = 0; j < all.length; j++) {
    if ((all[j].textContent || '').trim() === '共享数据查看' && all[j].offsetParent !== null) {
        all[j].click();
        return true;
    }
}
return false;
""")
        if not clicked2:
            print('  ⚠️ 未找到「共享数据查看」菜单项')

        # 等待报表树加载
        time.sleep(1)
        for _ in range(15):
            has_tree = _driver.execute_script("""
var tree = document.querySelector('.el-tree, .org_tree');
return !!tree;
""")
            if has_tree:
                print('  ✅ 报表树已加载')
                _saved_page_url = _driver.current_url
                return True
            time.sleep(0.5)

        print('  ⚠️ 报表树未在预期时间内加载')
        return False

    # 不在已知页面，可能已经在目标页
    return True


def _ensure_logged_in():
    """检测并处理登录状态过期，自动重新登录。

    检测场景：
    1. Element UI 弹窗 "登录状态已过期" → 点击"重新登录" → 填写表单 → 登录
    2. 页面直接显示登录表单 → 直接填写表单 → 登录

    调用方: _refresh_and_reestablish（刷新后确保已登录再切 iframe）。
    返回值: bool — 是否已登录（或重新登录成功）。
    """
    global _driver, _saved_page_url
    if _driver is None:
        return False

    try:
        _driver.switch_to.default_content()
    except Exception:
        pass

    # 读取登录凭据（从 config.json 缓存，避免频繁读 xlsx）
    cfg = {}
    try:
        if CONFIG_FILE.exists():
            cfg = json.loads(CONFIG_FILE.read_text(encoding='utf-8'))
    except Exception:
        pass
    credit_code = str(cfg.get('credit_code', '') or '').strip()
    username = str(cfg.get('username', '') or '').strip()
    password = str(cfg.get('password', '') or '').strip()

    if not all([credit_code, username, password]):
        return True  # 无凭据可填，假设当前已登录

    sel = DEFAULT_LOGIN_SELECTORS

    state = _detect_login_state()
    if state.get("logged_in"):
        _saved_page_url = _driver.current_url
        return True

    # ── 检测登录过期弹窗 ──
    try:
        expired = state.get("expired") or _driver.execute_script("""
var boxes = document.querySelectorAll('.el-message-box__message p');
for (var i = 0; i < boxes.length; i++) {
    var t = (boxes[i].textContent || '').trim();
    if (t.indexOf('登录状态已过期') >= 0 || t.indexOf('登录状态已过期') >= 0) {
        return true;
    }
}
// Also check for the whole message box
var wrappers = document.querySelectorAll('.el-message-box__wrapper');
for (var j = 0; j < wrappers.length; j++) {
    var style = window.getComputedStyle(wrappers[j]);
    if (style.display === 'none' || style.visibility === 'hidden') continue;
    var msg = wrappers[j].querySelector('.el-message-box__message');
    if (msg && (msg.textContent || '').indexOf('登录状态已过期') >= 0) {
        return true;
    }
}
return false;
""")
        if expired:
            print('  ⚠️ 检测到登录状态已过期，自动重新登录...')

            # 点击"重新登录"按钮（el-message-box__btns 中 primary 按钮）
            clicked = _driver.execute_script("""
var btns = document.querySelectorAll('.el-message-box__btns button');
for (var i = 0; i < btns.length; i++) {
    var t = (btns[i].textContent || '').trim();
    if (t === '重新登录' || t === '重新登录') {
        btns[i].click();
        return true;
    }
}
// Fallback: click the primary button
var primary = document.querySelector('.el-message-box__btns .el-button--primary');
if (primary) { primary.click(); return true; }
return false;
""")
            if clicked:
                time.sleep(1.5)
            else:
                print('  ⚠️ 未找到"重新登录"按钮，尝试直接定位登录表单')
    except Exception as e:
        print(f'  ⚠️ 检测登录过期弹窗异常: {e}')

    # ── 检测是否在登录页面（无 session） ──
    state = _detect_login_state()
    on_login_page = state.get("on_login_page") or _driver.execute_script("""
// Check for login form
var form = document.querySelector('form.login-form');
if (!form) return false;
// Check that it's visible
var style = window.getComputedStyle(form);
if (style.display === 'none' || style.visibility === 'hidden') return false;
return true;
""")

    if not on_login_page:
        # Check if we're neither on login page nor have logged-in UI
        state = _detect_login_state()
        if state.get("logged_in"):
            _saved_page_url = _driver.current_url
            return True  # Already logged in, no issue
        # No app UI and no login form - tricky state, try navigating
        print('  ⚠️ 未检测到已登录状态，尝试重新导航...')
        try:
            _driver.get(_saved_page_url or _driver.current_url)
            time.sleep(3)
            # Re-check after navigation
            on_login_page = _driver.execute_script("""
var form = document.querySelector('form.login-form');
return !!(form && window.getComputedStyle(form).display !== 'none');
""")
        except Exception:
            pass

    if on_login_page:
        print('  🔑 在登录页面，执行自动登录...')
        try:
            # 社会信用代码
            _login_fill_text(sel['credit_code_selector'], credit_code, '社会信用代码')
            time.sleep(0.8)

            # 用户名
            _login_select_username(sel['username_selector'], username)

            # 密码
            _login_fill_text(sel['password_selector'], password, '密码')

            # 登录按钮
            from selenium.webdriver.common.by import By
            btn = _login_wait_css(sel['submit_selector'], timeout=15)
            try:
                btn.click()
            except Exception:
                _driver.execute_script("arguments[0].click();", btn)

            # 等待登录成功
            from selenium.webdriver.support.ui import WebDriverWait
            login_url = _driver.current_url
            try:
                WebDriverWait(_driver, 30).until(
                    lambda d: (
                        d.current_url != login_url
                        or len(d.find_elements(By.CSS_SELECTOR, sel['success_indicator'])) > 0
                        or len(d.find_elements(By.CSS_SELECTOR, 'form.login-form')) == 0
                    )
                )
                print('  ✅ 自动重新登录成功')
                _saved_page_url = _driver.current_url
                # 登录后导航到报表页面（门户→报表服务中心→数据共享→共享数据查看）
                _navigate_to_report_page()
                return True
            except Exception:
                print('  ⚠️ 重新登录后未检测到成功标志，继续尝试...')
                # Check if login form disappeared (login likely succeeded)
                still_on_login = _driver.execute_script("""
var form = document.querySelector('form.login-form');
return !!(form && window.getComputedStyle(form).display !== 'none');
""")
                if not still_on_login:
                    print('  ✅ 登录表单已消失，视为登录成功')
                    _saved_page_url = _driver.current_url
                    _navigate_to_report_page()
                    return True
                return False

        except Exception as e:
            print(f'  ✗ 自动重新登录失败: {e}')
            return False

    return True


def _navigate_tree_to_report(path):
    """按路径导航左侧 el-tree 并点击叶子节点打开报表 Tab。

    路径格式: "关键词A > 关键词B > 关键词C"
    每一级用关键词模糊匹配（包含即可），匹配到多个时自动选第一个。
    逐级展开中间节点（等待异步加载），最后点击叶子节点打开 Tab。

    调用方: _refresh_and_reestablish（刷新后恢复 Tab）。
    返回值: bool — 是否成功打开报表 Tab。
    """
    global _driver
    if not path or not path.strip():
        return True

    keywords = [n.strip() for n in path.split('>') if n.strip()]
    if not keywords:
        return True

    print(f'  🌲 导航树: {" > ".join(keywords)}')

    # 先展开根节点"共享报表查看"（刷新后可能缩着，只点展开箭头不点内容）
    _driver.execute_script("""
var allNodes = document.querySelectorAll('.el-tree-node');
for (var i = 0; i < allNodes.length; i++) {
    var node = allNodes[i];
    if (node.offsetParent === null) continue;
    var spans = node.querySelectorAll('.el-tree-node__content span');
    for (var s = 0; s < spans.length; s++) {
        var txt = (spans[s].textContent || '').trim();
        if (txt === '共享报表查看') {
            if (!node.classList.contains('is-expanded')) {
                var icon = node.querySelector(':scope > .el-tree-node__content .el-tree-node__expand-icon');
                if (icon) { icon.click(); }
            }
            return node.classList.contains('is-expanded') ? 'ok' : 'expanding';
        }
    }
}
return 'not_found';
""")
    time.sleep(0.5)

    # 清空之前可能残留的访问标记
    _driver.execute_script("""
var nodes = document.querySelectorAll('.el-tree-node[data-nav-visited]');
for (var i = 0; i < nodes.length; i++) {
    nodes[i].removeAttribute('data-nav-visited');
}
""")

    for i, keyword in enumerate(keywords):
        is_last = (i == len(keywords) - 1)

        # 轮询等待节点出现（异步加载可能延迟）
        found = False
        prev_visible_count = _driver.execute_script(
            "return document.querySelectorAll('.el-tree-node').length"
        )
        for attempt in range(20):  # 最多 20 次 (~10s)
            result = _driver.execute_script("""
var keyword = arguments[0];
var clickLeaf = arguments[1];

// 搜索可见树节点，跳过已访问的（避免重复匹配上级节点）
var allNodes = document.querySelectorAll('.el-tree-node');
var matches = [];
for (var i = 0; i < allNodes.length; i++) {
    var node = allNodes[i];
    if (node.offsetParent === null) continue;
    // 跳过之前已展开过的节点（避免每级都匹配到同一个父节点）
    if (node.hasAttribute('data-nav-visited')) continue;

    var titles = [];
    var titleEl = node.querySelector('.org_tree_node [title]');
    if (titleEl) titles.push(titleEl.getAttribute('title') || '');
    var spanEls = node.querySelectorAll('.el-tree-node__content span');
    for (var s = 0; s < spanEls.length; s++) {
        var txt = (spanEls[s].textContent || '').trim();
        if (txt) titles.push(txt);
    }

    for (var t = 0; t < titles.length; t++) {
        if (titles[t].indexOf(keyword) >= 0 || keyword.indexOf(titles[t]) >= 0) {
            matches.push({node: node, title: titles[t]});
            break;
        }
    }
}

if (matches.length === 0) return {ok: false, error: 'no_match', keyword: keyword, totalNodes: allNodes.length};

// 多个匹配时选最合适的
var best = matches[0];
var found = best.node;
if (matches.length > 1) {
    if (clickLeaf) {
        // 最后一级：优先选叶子节点（.cpt 文件），其次选已展开的
        for (var m = 0; m < matches.length; m++) {
            var icon = matches[m].node.querySelector(':scope > .el-tree-node__content .el-tree-node__expand-icon');
            if (icon && icon.classList.contains('is-leaf')) {
                best = matches[m]; found = best.node; break;
            }
        }
    }
    if (found === matches[0].node) {
        // 中间级：优先选已展开的子节点
        for (var m = 0; m < matches.length; m++) {
            if (matches[m].node.classList.contains('is-expanded')) {
                best = matches[m]; found = best.node; break;
            }
        }
    }
}

var expandIcon = found.querySelector(':scope > .el-tree-node__content .el-tree-node__expand-icon');
var isLeaf = expandIcon && expandIcon.classList.contains('is-leaf');
var isExpanded = found.classList.contains('is-expanded');

// 标记已访问，后续级别不再匹配此节点
found.setAttribute('data-nav-visited', '1');

if (clickLeaf || isLeaf) {
    var content = found.querySelector(':scope > .el-tree-node__content');
    if (content) {
        try { content.click(); } catch(e) {}
        return {ok: true, clicked: true, keyword: keyword, matched: best.title, total: matches.length};
    }
    return {ok: false, error: 'no_content', keyword: keyword};
} else {
    // 非叶子节点：只展开，不点内容（点内容可能缩回去）
    if (!isExpanded && expandIcon) {
        try { expandIcon.click(); } catch(e) {}
        return {ok: true, expanded: true, keyword: keyword, matched: best.title, total: matches.length};
    }
    if (isExpanded) {
        return {ok: true, already_expanded: true, keyword: keyword, matched: best.title, total: matches.length};
    }
    return {ok: false, error: 'cannot_expand', keyword: keyword};
}
""", keyword, is_last)

            if result and result.get('ok'):
                matched = result.get('matched', keyword)
                total_matches = result.get('total', 1)
                hint = f' ({total_matches}个匹配)' if total_matches > 1 else ''
                if result.get('clicked'):
                    print(f'  ✓ 点击: {matched}{hint}')
                    found = True
                    break
                elif result.get('expanded') or result.get('already_expanded'):
                    if attempt == 0:
                        print(f'  ✓ 展开: {matched}{hint}')
                    found = True
                    # 展开后等待子节点出现（监控节点数增长）
                    for _ in range(15):
                        time.sleep(0.4)
                        cur_count = _driver.execute_script(
                            "return document.querySelectorAll('.el-tree-node').length"
                        )
                        if cur_count > prev_visible_count:
                            prev_visible_count = cur_count
                            break
                    break
            else:
                err = result.get('error', 'unknown') if result else 'exec_error'
                err_extra = result.get('totalNodes', 0) if result else 0
                if attempt < 19:
                    # 检查是否有新节点出现（说明上一级正在加载）
                    cur_count = _driver.execute_script(
                        "return document.querySelectorAll('.el-tree-node').length"
                    )
                    if cur_count > prev_visible_count:
                        prev_visible_count = cur_count
                        # 有新节点了，快速重试
                        time.sleep(0.2)
                    else:
                        time.sleep(0.5)
                else:
                    print(f'  ⚠️ 未找到 [{keyword}]: {err} (可见节点: {err_extra})')
                    return False

        if not found:
            return False

    # 等待 Tab 出现
    print('  等待报表 Tab 打开...')
    for _ in range(15):
        tabs = _driver.execute_script("""
var items = document.querySelectorAll('.el-tabs__item');
var visible = [];
for (var i = 0; i < items.length; i++) {
    if (items[i].offsetParent !== null) {
        visible.push((items[i].textContent || '').trim());
    }
}
return visible;
""")
        if tabs and len(tabs) > 0:
            print(f'  ✓ Tab 已打开')
            return True
        time.sleep(0.5)

    print('  ⚠️ Tab 未在预期时间内打开')
    return True  # 即使 Tab 未检测到也不阻塞，后续 _switch_to_report_frame 会处理


def _expand_all_report_tree(max_rounds=8):
    """尽量展开左侧报表树，给 tree_path 二次导航做兜底。"""
    global _driver
    if _driver is None:
        return False
    print('  🌲 全量展开报表树作为兜底...')
    expanded_any = False
    try:
        for _ in range(max(1, int(max_rounds or 1))):
            clicked = _driver.execute_script("""
var icons = Array.from(document.querySelectorAll('.el-tree-node__expand-icon'))
    .filter(function(icon) {
        if (!icon || icon.offsetParent === null) return false;
        if (icon.classList.contains('is-leaf')) return false;
        var node = icon.closest('.el-tree-node');
        return node && !node.classList.contains('is-expanded');
    });
icons.slice(0, 80).forEach(function(icon) {
    try { icon.click(); } catch(e) {}
});
return icons.length;
""")
            if not clicked:
                break
            expanded_any = True
            time.sleep(0.8)
        return True
    except Exception as e:
        print(f'  ⚠️ 全量展开报表树失败: {e}')
        return expanded_any


def _switch_to_report_frame(scene, plan=None):
    """切换到报表 iframe，优先用外层报表 tab 标题区分场景。"""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait

    _driver.switch_to.default_content()
    scene_defaults = SCENE_REPORT_DEFAULTS.get(scene, {})
    report_title_keyword = ""
    report_title_index = 0
    if plan:
        report_title_keyword = str(plan.get("report_title_keyword", "") or "").strip()
        report_title_index = int(plan.get("report_title_index", 0) or 0)
    if not report_title_keyword:
        report_title_keyword = scene_defaults.get("report_title_keyword", "")

    def _clean_title(text):
        text = re.sub(r"\s+", "", str(text or ""))
        for suffix in ("(报表)", "（报表）"):
            if text.endswith(suffix):
                text = text[:-len(suffix)]
        return text

    def _tab_info(tab, idx):
        tab_text = (
            tab.text
            or tab.get_attribute("innerText")
            or tab.get_attribute("textContent")
            or ""
        ).strip()
        pane_id = tab.get_attribute("aria-controls") or ""
        iframe_src = ""
        if pane_id:
            try:
                pane = _driver.find_element(By.ID, pane_id)
                found = pane.find_elements(By.TAG_NAME, "iframe")
                if found:
                    iframe_src = found[0].get_attribute("src") or ""
            except Exception:
                iframe_src = ""
        return {"idx": idx, "tab": tab, "text": tab_text, "pane_id": pane_id, "iframe_src": iframe_src}

    def _print_ambiguous(candidates, match_type):
        print(f'  ⚠️ 报表标签匹配到多个候选（{match_type}），为避免导错表已停止。')
        print(f'     task.report_title_keyword={report_title_keyword!r}')
        print('     可把 keyword 写得更完整，或设置 task.report_title_index=1/2/3 指定第几个。')
        for c in candidates:
            print(f'     [{c["idx"]}] tab={c["text"]!r} pane={c["pane_id"]!r} src={c["iframe_src"]!r}')

    def _choose_tab(candidates, match_type):
        if not candidates:
            return None
        if report_title_index:
            pos = report_title_index - 1
            if pos < len(candidates):
                chosen = candidates[pos]
                print(f'  报表标签{match_type}匹配多个，按 task.report_title_index={report_title_index} 选择 [{chosen["idx"]}]')
                return chosen
            print(f'  ⚠️ task.report_title_index={report_title_index} 超出候选数量 {len(candidates)}')
            _print_ambiguous(candidates, match_type)
            return False
        if len(candidates) == 1:
            return candidates[0]
        _print_ambiguous(candidates, match_type)
        return False

    # 优先根据外层 Element UI tab 标题定位 pane，再切入该 pane 下的 iframe。
    # county 和 county2 都有"机构类型"，只靠 iframe 内容会误判。
    if report_title_keyword:
        tabs = _driver.find_elements(By.CSS_SELECTOR, "[role='tab'], .el-tabs__item")
        tab_infos = []
        for idx, tab in enumerate(tabs, start=1):
            try:
                info = _tab_info(tab, idx)
                if info["text"]:
                    tab_infos.append(info)
            except Exception:
                continue

        keyword_clean = _clean_title(report_title_keyword)
        exact_matches = [
            info for info in tab_infos
            if _clean_title(info["text"]) == keyword_clean
        ]
        contains_matches = [
            info for info in tab_infos
            if keyword_clean and keyword_clean in _clean_title(info["text"])
        ]
        chosen = _choose_tab(exact_matches, "完全") if exact_matches else _choose_tab(contains_matches, "包含")
        if chosen is False:
            return False
        if chosen:
            try:
                tab = chosen["tab"]
                pane_id = chosen["pane_id"]
                try:
                    _driver.execute_script(
                        "arguments[0].scrollIntoView({block:'center', inline:'center'});",
                        tab,
                    )
                except Exception:
                    pass
                try:
                    tab.click()
                except Exception:
                    _driver.execute_script("arguments[0].click();", tab)

                if not pane_id:
                    raise RuntimeError("匹配到tab但缺少aria-controls，无法确定对应iframe")

                def _find_pane_iframe(driver):
                    pane = driver.find_element(By.ID, pane_id)
                    found = pane.find_elements(By.TAG_NAME, "iframe")
                    return found[0] if found else False

                iframe = WebDriverWait(_driver, 10).until(_find_pane_iframe)
                _driver.switch_to.default_content()
                _driver.switch_to.frame(iframe)
                title = _driver.execute_script("return document.title || ''")
                print(f'  切换到报表 iframe（场景: {scene}, tab: {chosen["text"]}, title: {title}）')
                return True
            except Exception:
                _driver.switch_to.default_content()
                print(f'  ⚠️ 按报表标题找到tab但切换iframe失败，转入iframe标题扫描: {chosen.get("text", "")}')

        # 如果外层 tab 没有可用 aria-controls，则退一步直接扫描 iframe 的 document.title。
        for iframe in _driver.find_elements(By.TAG_NAME, 'iframe'):
            try:
                _driver.switch_to.default_content()
                _driver.switch_to.frame(iframe)
                title = _driver.execute_script("return document.title || ''")
                body_text = _driver.find_element(By.TAG_NAME, 'body').text
                merged = _clean_title(f"{title}\n{body_text}")
                if keyword_clean and keyword_clean in merged:
                    print(f'  切换到报表 iframe（场景: {scene}, title: {title}）')
                    return True
            except Exception:
                _driver.switch_to.default_content()
                continue
        _driver.switch_to.default_content()

    _driver.switch_to.default_content()
    expected_title = report_title_keyword or "(未配置)"
    print(f'  ⚠️ 未找到标题包含"{expected_title}"的报表 iframe')
    return False


def _fr_click_export():
    """三步导出: 点击导出 → 鼠标悬停Excel → 点击原样导出。"""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.action_chains import ActionChains
    from selenium.common.exceptions import StaleElementReferenceException
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    last_err = None
    for _ in range(3):
        try:
            wait = WebDriverWait(_driver, 8)

            # 1. 点击导出（每次重定位，避免 stale）
            export_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space(.)='导出'] | //*[normalize-space(text())='导出']")))
            try:
                export_btn.click()
            except Exception:
                _driver.execute_script("arguments[0].click();", export_btn)
            time.sleep(0.25)

            # 2. 悬停 Excel
            excel_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//*[normalize-space(text())='Excel']")))
            ActionChains(_driver).move_to_element(excel_btn).perform()
            time.sleep(0.25)

            # 3. 点击原样导出
            original_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[normalize-space(text())='原样导出']")))
            try:
                original_btn.click()
            except Exception:
                _driver.execute_script("arguments[0].click();", original_btn)
            time.sleep(0.25)

            # 关闭可能残留的菜单
            try:
                ActionChains(_driver).send_keys('\ue00c').perform()
            except Exception:
                pass
            print('  已点击 导出 → Excel → 原样导出')
            return
        except StaleElementReferenceException as e:
            last_err = e
            time.sleep(0.35)
            continue
        except Exception as e:
            last_err = e
            time.sleep(0.35)
            continue

    if last_err:
        raise last_err
    raise RuntimeError('export_click_failed')


def _fr_set_field(label_text, value, field_type='text', placeholder='', select_strategy='sendkey'):
    """通过 JS 设置字段值。"""
    js = r"""
var labelText = arguments[0];
var value = arguments[1] || '';
var fieldType = (arguments[2] || 'text').toLowerCase();
var placeholder = (arguments[3] || '').trim();
var selectStrategy = (arguments[4] || 'sendkey').toLowerCase();

function normText(s) {
    var x = (s === null || s === undefined) ? '' : String(s);
    return x.replace(/[:：\s]/g, '').trim().toLowerCase();
}
function textMatch(target, actual) {
    var t0 = (target === null || target === undefined) ? '' : String(target).trim();
    var a0 = (actual === null || actual === undefined) ? '' : String(actual).trim();
    if (!t0 || !a0) return false;
    if (t0 === a0) return true;
    var t = normText(t0);
    var a = normText(a0);
    if (!t || !a) return false;
    return a === t || a.indexOf(t) >= 0 || t.indexOf(a) >= 0;
}
function getTagTexts(root) {
    if (!root) return [];
    return Array.from(root.querySelectorAll('.fr-tageditor-container .fr-tageditor-tag'))
        .map(function(el) { return (el.textContent || '').trim(); })
        .filter(Boolean);
}
function getVisibleFRDropdowns() {
    return Array.from(document.querySelectorAll('.fr-checkbox-list.fr-combo-list, .fr-combo-list'))
        .filter(function(el) {
            if (!el || el.offsetParent === null) return false;
            var style = window.getComputedStyle(el);
            return style.display !== 'none' && style.visibility !== 'hidden';
        });
}
function pickBestFRDropdown(before, input, wantCheckbox) {
    var after = getVisibleFRDropdowns().filter(function(el) {
        var isCheckbox = el.classList.contains('fr-checkbox-list');
        return wantCheckbox ? isCheckbox : !isCheckbox;
    });
    var created = after.filter(function(el) { return before.indexOf(el) < 0; });
    var candidates = created.length ? created : after;
    if (!candidates.length) return null;
    var root = input.closest('.fr-trigger-editor') || input;
    var rect = root.getBoundingClientRect();
    var best = null;
    var bestScore = Number.POSITIVE_INFINITY;
    candidates.forEach(function(el) {
        var r = el.getBoundingClientRect();
        var dy = Math.abs(r.top - rect.bottom);
        var dx = Math.abs(r.left - rect.left);
        var widthPenalty = Math.abs((r.width || 0) - (rect.width || 0));
        var score = dy * 1000 + dx * 10 + widthPenalty;
        if (score < bestScore) {
            bestScore = score;
            best = el;
        }
    });
    return best;
}
function dropdownAligned(el, input) {
    if (!el || !input) return false;
    var root = input.closest('.fr-trigger-editor') || input;
    var rect = root.getBoundingClientRect();
    var r = el.getBoundingClientRect();
    var dx = Math.abs(r.left - rect.left);
    var dy = Math.abs(r.top - rect.bottom);
    return dx <= Math.max(12, rect.width * 0.25) && dy <= 80;
}
function closeStaleFRDropdowns(input, wantCheckbox) {
    var existing = pickBestFRDropdown([], input, wantCheckbox);
    if (!existing || dropdownAligned(existing, input)) return existing;
    try {
        document.dispatchEvent(new KeyboardEvent('keydown', {key:'Escape', code:'Escape', bubbles:true}));
        document.dispatchEvent(new KeyboardEvent('keyup', {key:'Escape', code:'Escape', bubbles:true}));
    } catch (e) {}
    try {
        if (document.activeElement && document.activeElement.blur) document.activeElement.blur();
    } catch (e) {}
    return null;
}
function openFRDropdown(root, input, wantCheckbox) {
    var before = getVisibleFRDropdowns();
    var existing = closeStaleFRDropdowns(input, wantCheckbox);
    if (existing) return existing;
    function fireMouse(el) {
        if (!el) return false;
        try { el.scrollIntoView({block:'center', inline:'center'}); } catch (e) {}
        var r = el.getBoundingClientRect();
        var x = r.left + Math.max(1, r.width / 2);
        var y = r.top + Math.max(1, r.height / 2);
        ['mouseover', 'mousemove', 'mousedown', 'mouseup', 'click'].forEach(function(type) {
            try {
                el.dispatchEvent(new MouseEvent(type, {
                    bubbles: true,
                    cancelable: true,
                    view: window,
                    clientX: x,
                    clientY: y
                }));
            } catch (e) {}
        });
        try { el.click(); } catch (e) {}
        return true;
    }
    var center = root ? root.querySelector('.fr-trigger-center, .fr-date-trigger-center') : null;
    var btn = root ? root.querySelector(
        '.fr-trigger-btn-up, .fr-trigger-btn-over, .fr-trigger-btn-down, .frtriggerbtn-up, .frtriggerbtn-over, .frtriggerbtn-down'
    ) : null;
    if (!fireMouse(center)) {
        if (!fireMouse(btn)) {
            fireMouse(input);
        }
    }
    var first = pickBestFRDropdown(before, input, wantCheckbox);
    if (first) return first;
    // Hover can switch the outer button class; try the current outer node after center.
    if (btn && btn !== center) {
        fireMouse(btn);
    }
    return pickBestFRDropdown(before, input, wantCheckbox);
}
function clearFRTagEditor(root, input) {
    var attempts = 0;
    while (attempts < 3) {
        var tags = getTagTexts(root);
        if (!tags.length) return [];
        var dropdown = openFRDropdown(root, input, true);
        if (!dropdown) break;
        var ctrl = dropdown.querySelector('.fr-checkbox-control .fr-widget-click');
        if (ctrl) {
            try { ctrl.click(); } catch (e) {}
        } else {
            var selected = Array.from(dropdown.querySelectorAll('.fr-combo-list-item.fr-combo-selected .fr-widget-click, .fr-combo-list-item.fr-combo-selected'));
            selected.forEach(function(el) { try { el.click(); } catch (e) {} });
        }
        attempts += 1;
    }
    return getTagTexts(root);
}
function selectFRTagValue(root, input, targetValue) {
    var dropdown = openFRDropdown(root, input, true);
    if (!dropdown) return {ok:false, msg:'fr_checkbox_dropdown_not_found'};
    var items = Array.from(dropdown.querySelectorAll('.fr-combo-list-item'))
        .filter(function(el) {
            var txt = (el.getAttribute('title') || el.innerText || '').trim();
            return txt && txt.indexOf('全选/不选') < 0;
        });
    var opt = items.find(function(el) {
        var txt = (el.getAttribute('title') || el.innerText || '').trim();
        return txt === targetValue;
    }) || items.find(function(el) {
        var txt = (el.getAttribute('title') || el.innerText || '').trim();
        return txt.indexOf(targetValue) >= 0 || targetValue.indexOf(txt) >= 0;
    });
    if (!opt) {
        return {
            ok:false,
            msg:'checkbox_dropdown_not_match',
            options: items.map(function(el) { return (el.getAttribute('title') || el.innerText || '').trim(); })
        };
    }
    var clickTarget = opt.querySelector('.fr-widget-click') || opt;
    try { clickTarget.click(); } catch (e) { opt.click(); }
    return {ok:true, actual:getTagTexts(root).join('|'), set_mode:'checkbox_dropdown'};
}
function selectFRComboValue(root, input, targetValue) {
    setInputValue(input, targetValue);
    var dropdown = openFRDropdown(root, input, false);
    if (dropdown) {
        var items = Array.from(dropdown.querySelectorAll('.fr-combo-list-item'))
            .filter(function(el) {
                var txt = (el.getAttribute('title') || el.innerText || '').trim();
                return !!txt;
            });
        var opt = items.find(function(el) {
            var txt = (el.getAttribute('title') || el.innerText || '').trim();
            return txt === targetValue;
        }) || items.find(function(el) {
            var txt = (el.getAttribute('title') || el.innerText || '').trim();
            return textMatch(targetValue, txt);
        });
        if (opt) {
            try { opt.dispatchEvent(new MouseEvent('mousedown', {bubbles:true, cancelable:true, view:window})); } catch (e) {}
            try { opt.click(); } catch (e) {}
            try { opt.dispatchEvent(new MouseEvent('mouseup', {bubbles:true, cancelable:true, view:window})); } catch (e) {}
            commitInputValue(input);
            clickBlankToCommit(input);
            return {
                ok: textMatch(targetValue, input.value || ''),
                actual: input.value || '',
                set_mode: 'sendkey_combo_click'
            };
        }
    }
    try {
        input.dispatchEvent(new KeyboardEvent('keydown', {key:'Enter', code:'Enter', bubbles:true, cancelable:true}));
        input.dispatchEvent(new KeyboardEvent('keyup', {key:'Enter', code:'Enter', bubbles:true, cancelable:true}));
    } catch (e) {}
    commitInputValue(input);
    clickBlankToCommit(input);
    return {
        ok: textMatch(targetValue, input.value || ''),
        actual: input.value || '',
        msg: textMatch(targetValue, input.value || '') ? '' : 'sendkey_select_not_committed',
        set_mode: 'sendkey_enter_fallback'
    };
}

function findElementInput(labelText, placeholder) {
    // 1) Element UI: .el-form-item
    var items = document.querySelectorAll('.el-form-item');
    var item = null;
    for (var i = 0; i < items.length; i++) {
        var t = (items[i].getAttribute('title') || '').trim();
        if (textMatch(labelText, t)) { item = items[i]; break; }
        var labels = items[i].querySelectorAll('label');
        for (var j = 0; j < labels.length; j++) {
            if (textMatch(labelText, (labels[j].textContent || '').trim())) {
                item = items[i];
                break;
            }
        }
        if (item) break;
    }
    if (item) {
        var inp = placeholder ? item.querySelector('input[placeholder="' + placeholder + '"]') : null;
        if (!inp) inp = item.querySelector('input.el-input__inner, input');
        if (inp) return inp;
    }

    // 2) FineReport: pre.fr-label + fr-trigger-editor
    var labels = Array.from(document.querySelectorAll('pre.fr-label'));
    var labelPre = labels.find(function(pre) {
        return textMatch(labelText, (pre.textContent || '').trim());
    });
    if (!labelPre) return null;

    var labelBox = labelPre.closest('[widgetname]') || labelPre.parentElement;
    if (!labelBox || !labelBox.parentElement) return null;
    var parent = labelBox.parentElement;
    var labelRect = labelBox.getBoundingClientRect();
    var editors = Array.from(parent.querySelectorAll('.fr-trigger-editor'));
    if (!editors.length) return null;
    var best = null;
    var bestScore = Number.POSITIVE_INFINITY;
    for (var k = 0; k < editors.length; k++) {
        var r = editors[k].getBoundingClientRect();
        var dy = Math.abs(r.top - labelRect.top);
        var dx = r.left - labelRect.right;
        if (dx < -5) continue;
        var score = dy * 1000 + Math.max(0, dx);
        if (score < bestScore) {
            bestScore = score;
            best = editors[k];
        }
    }
    if (!best) return null;
    var finp = placeholder ? best.querySelector('input[placeholder="' + placeholder + '"]') : null;
    if (!finp) finp = best.querySelector('input.fr-trigger-texteditor, input');
    return finp || null;
}

var input = findElementInput(labelText, placeholder);
if (!input) return {ok:false, msg:'input_not_found'};
var frRoot = input.closest('.fr-trigger-editor');
var isFRTagEditor = !!(frRoot && frRoot.querySelector('.fr-tageditor-container'));

function clearInputValue(inp) {
    inp.focus();
    try { inp.click(); } catch (e) {}
    try { inp.select(); } catch (e) {}
    ['keydown', 'keyup'].forEach(function(type) {
        try {
            inp.dispatchEvent(new KeyboardEvent(type, {
                key:'a',
                code:'KeyA',
                ctrlKey:true,
                bubbles:true,
                cancelable:true
            }));
        } catch (e) {}
    });
    inp.value = '';
    inp.dispatchEvent(new Event('input', {bubbles:true}));
    inp.dispatchEvent(new Event('change', {bubbles:true}));
    ['keydown', 'keyup'].forEach(function(type) {
        try {
            inp.dispatchEvent(new KeyboardEvent(type, {
                key:'Backspace',
                code:'Backspace',
                bubbles:true,
                cancelable:true
            }));
        } catch (e) {}
    });
    try { inp.blur(); } catch (e) {}
}
function clickBlankToCommit(inp) {
    try {
        if (inp) {
            inp.dispatchEvent(new Event('change', {bubbles:true}));
            inp.blur();
        }
    } catch (e) {}
    var blank = document.querySelector('.reportPane, .parameter-container, #content-container, body') || document.body;
    if (!blank) return;
    try {
        var r = blank.getBoundingClientRect();
        var x = r.left + Math.min(Math.max(30, r.width * 0.15), Math.max(30, r.width - 30));
        var y = r.top + Math.min(Math.max(30, r.height * 0.15), Math.max(30, r.height - 30));
        ['mousedown', 'mouseup', 'click'].forEach(function(type) {
            blank.dispatchEvent(new MouseEvent(type, {
                bubbles:true,
                cancelable:true,
                view:window,
                clientX:x,
                clientY:y
            }));
        });
    } catch (e) {}
}
function closeOpenDropdowns() {
    try {
        document.dispatchEvent(new KeyboardEvent('keydown', {key:'Escape', code:'Escape', bubbles:true}));
        document.dispatchEvent(new KeyboardEvent('keyup', {key:'Escape', code:'Escape', bubbles:true}));
    } catch (e) {}
}
function clearSelectValue(input) {
    var root = input.closest('.el-select, .el-input, .fr-trigger-editor') || input.parentElement;
    try {
        if (root) {
            var clearBtn = root.querySelector('.el-icon-circle-close, .el-input__clear, .fr-trigger-clear, .fr-trigger-close');
            if (clearBtn) {
                try { clearBtn.dispatchEvent(new MouseEvent('mousedown', {bubbles:true, cancelable:true, view:window})); } catch (e) {}
                try { clearBtn.click(); } catch (e) {}
            }
            var tagCloses = root.querySelectorAll('.el-select__tags .el-tag .el-tag__close');
            tagCloses.forEach(function(el) { try { el.click(); } catch (e) {} });
        }
    } catch (e) {}
    clearInputValue(input);
    clickBlankToCommit(input);
    closeOpenDropdowns();
    clearInputValue(input);
    clickBlankToCommit(input);
}
function commitInputValue(inp) {
    inp.dispatchEvent(new Event('change', {bubbles:true}));
    try { inp.blur(); } catch (e) {}
}
function setInputValue(inp, val) {
    inp.focus();
    try { inp.click(); } catch (e) {}
    inp.value = val;
    inp.dispatchEvent(new Event('input', {bubbles:true}));
    inp.dispatchEvent(new Event('change', {bubbles:true}));
    commitInputValue(inp);
}

// 清空/跳过
if (value === '' || value === '\\u4e0d\\u9009') {
    if (isFRTagEditor) {
        var remain = clearFRTagEditor(frRoot, input);
        clearInputValue(input);
        return {ok: remain.length === 0, actual: remain.join('|'), msg: remain.length ? 'tag_clear_incomplete' : ''};
    }
    clearInputValue(input);
    return {ok:true, actual:''};
}

// select/readonly
if (fieldType === 'select' || input.readOnly) {
    // 先清空输入框，避免多选场景被叠加
    if (isFRTagEditor) {
        clearFRTagEditor(frRoot, input);
    } else {
        clearSelectValue(input);
    }
    if (selectStrategy === 'sendkey') {
        if (frRoot && !isFRTagEditor) {
            return selectFRComboValue(frRoot, input, value);
        }
        setInputValue(input, value);
        return {ok:true, actual:(input.value || ''), set_mode:'sendkey'};
    }
    if (selectStrategy === 'checkbox_dropdown') {
        if (frRoot) {
            return selectFRTagValue(frRoot, input, value);
        }
        return {ok:false, msg:'checkbox_dropdown_not_supported', set_mode:'checkbox_dropdown'};
    }
    return {ok:false, msg:'unsupported_method', set_mode:selectStrategy};
}

// text/date
clearInputValue(input);
setInputValue(input, value);
return {ok:true, actual:input.value || ''};
"""
    return _driver.execute_script(js, label_text, value, field_type, placeholder, select_strategy)


def _fr_get_field_value(label_text, placeholder=''):
    """获取字段当前值。"""
    js = r"""
var labelText = arguments[0];
var placeholder = (arguments[1] || '').trim();

function normText(s) {
    var x = (s === null || s === undefined) ? '' : String(s);
    return x.replace(/[:：\s]/g, '').trim().toLowerCase();
}
function textMatch(target, actual) {
    var t0 = (target === null || target === undefined) ? '' : String(target).trim();
    var a0 = (actual === null || actual === undefined) ? '' : String(actual).trim();
    if (!t0 || !a0) return false;
    if (t0 === a0) return true;
    var t = normText(t0);
    var a = normText(a0);
    if (!t || !a) return false;
    return a === t || a.indexOf(t) >= 0 || t.indexOf(a) >= 0;
}
function getTagTexts(root) {
    if (!root) return [];
    return Array.from(root.querySelectorAll('.fr-tageditor-container .fr-tageditor-tag'))
        .map(function(el) { return (el.textContent || '').trim(); })
        .filter(Boolean);
}
function findInput(labelText, placeholder) {
    var items = document.querySelectorAll('.el-form-item');
    var item = null;
    for (var i = 0; i < items.length; i++) {
        var t = (items[i].getAttribute('title') || '').trim();
        if (textMatch(labelText, t)) { item = items[i]; break; }
        var labels = items[i].querySelectorAll('label');
        for (var j = 0; j < labels.length; j++) {
            if (textMatch(labelText, (labels[j].textContent || '').trim())) { item = items[i]; break; }
        }
        if (item) break;
    }
    if (item) {
        var inp = placeholder ? item.querySelector('input[placeholder="' + placeholder + '"]') : null;
        if (!inp) inp = item.querySelector('input.el-input__inner, input');
        if (inp) return inp;
    }
    var labels2 = Array.from(document.querySelectorAll('pre.fr-label'));
    var labelPre = labels2.find(function(pre) {
        return textMatch(labelText, (pre.textContent || '').trim());
    });
    if (!labelPre) return null;
    var labelBox = labelPre.closest('[widgetname]') || labelPre.parentElement;
    if (!labelBox || !labelBox.parentElement) return null;
    var parent = labelBox.parentElement;
    var labelRect = labelBox.getBoundingClientRect();
    var editors = Array.from(parent.querySelectorAll('.fr-trigger-editor'));
    var best = null, bestScore = Number.POSITIVE_INFINITY;
    for (var k = 0; k < editors.length; k++) {
        var r = editors[k].getBoundingClientRect();
        var dy = Math.abs(r.top - labelRect.top);
        var dx = r.left - labelRect.right;
        if (dx < -5) continue;
        var score = dy * 1000 + Math.max(0, dx);
        if (score < bestScore) { bestScore = score; best = editors[k]; }
    }
    if (!best) return null;
    var finp = placeholder ? best.querySelector('input[placeholder="' + placeholder + '"]') : null;
    if (!finp) finp = best.querySelector('input.fr-trigger-texteditor, input');
    return finp || null;
}
var inp = findInput(labelText, placeholder);
if (!inp) return null;
var frRoot = inp.closest('.fr-trigger-editor');
if (frRoot && frRoot.querySelector('.fr-tageditor-container')) {
    return getTagTexts(frRoot).join('|');
}
return inp.value || '';
"""
    return _driver.execute_script(js, label_text, placeholder)


def _fr_get_all_options(label_text, placeholder='', want_checkbox=None, filter_text=''):
    """获取下拉框选项文本；filter_text 非空时先输入文本触发筛选。"""
    js = r"""
var labelText = arguments[0];
var placeholder = (arguments[1] || '').trim();
var wantCheckbox = arguments[2];
var filterText = (arguments[3] || '').trim();

function normText(s) {
    var x = (s === null || s === undefined) ? '' : String(s);
    return x.replace(/[:：\s]/g, '').trim().toLowerCase();
}
function textMatch(target, actual) {
    var t0 = (target === null || target === undefined) ? '' : String(target).trim();
    var a0 = (actual === null || actual === undefined) ? '' : String(actual).trim();
    if (!t0 || !a0) return false;
    if (t0 === a0) return true;
    var t = normText(t0);
    var a = normText(a0);
    if (!t || !a) return false;
    return a === t || a.indexOf(t) >= 0 || t.indexOf(a) >= 0;
}
function getVisibleFRDropdowns() {
    return Array.from(document.querySelectorAll('.fr-checkbox-list.fr-combo-list, .fr-combo-list'))
        .filter(function(el) {
            if (!el || el.offsetParent === null) return false;
            var style = window.getComputedStyle(el);
            return style.display !== 'none' && style.visibility !== 'hidden';
        });
}
function pickBestFRDropdown(input, before) {
    var all = getVisibleFRDropdowns().filter(function(el) {
        if (wantCheckbox === null || wantCheckbox === undefined) return true;
        var isCheckbox = el.classList.contains('fr-checkbox-list');
        return wantCheckbox ? isCheckbox : !isCheckbox;
    });
    var created = before ? all.filter(function(el) { return before.indexOf(el) < 0; }) : [];
    var candidates = created.length ? created : all;
    if (!candidates.length) return null;
    var root = input.closest('.fr-trigger-editor') || input;
    var rect = root.getBoundingClientRect();
    var best = null, bestScore = Number.POSITIVE_INFINITY;
    candidates.forEach(function(el) {
        var r = el.getBoundingClientRect();
        var score = Math.abs(r.top - rect.bottom) * 1000 + Math.abs(r.left - rect.left) * 10 + Math.abs((r.width || 0) - (rect.width || 0));
        if (score < bestScore) { bestScore = score; best = el; }
    });
    return best;
}
function dropdownAligned(el, input) {
    if (!el || !input) return false;
    var root = input.closest('.fr-trigger-editor') || input;
    var rect = root.getBoundingClientRect();
    var r = el.getBoundingClientRect();
    var dx = Math.abs(r.left - rect.left);
    var dy = Math.abs(r.top - rect.bottom);
    return dx <= Math.max(12, rect.width * 0.25) && dy <= 80;
}
function closeStaleFRDropdowns(input) {
    var existing = pickBestFRDropdown(input, null);
    if (!existing || dropdownAligned(existing, input)) return existing;
    try {
        document.dispatchEvent(new KeyboardEvent('keydown', {key:'Escape', code:'Escape', bubbles:true}));
        document.dispatchEvent(new KeyboardEvent('keyup', {key:'Escape', code:'Escape', bubbles:true}));
    } catch (e) {}
    try {
        if (document.activeElement && document.activeElement.blur) document.activeElement.blur();
    } catch (e) {}
    return null;
}
function fireMouse(el) {
    if (!el) return false;
    try { el.scrollIntoView({block:'center', inline:'center'}); } catch (e) {}
    var r = el.getBoundingClientRect();
    var x = r.left + Math.max(1, r.width / 2);
    var y = r.top + Math.max(1, r.height / 2);
    ['mouseover', 'mousemove', 'mousedown', 'mouseup', 'click'].forEach(function(type) {
        try {
            el.dispatchEvent(new MouseEvent(type, {bubbles:true, cancelable:true, view:window, clientX:x, clientY:y}));
        } catch (e) {}
    });
    try { el.click(); } catch (e) {}
    return true;
}
function applyFilter(input) {
    if (!filterText || !input) return;
    try { input.focus(); } catch (e) {}
    try { input.click(); } catch (e) {}
    input.value = filterText;
    try { input.dispatchEvent(new Event('input', {bubbles:true})); } catch (e) {}
    try { input.dispatchEvent(new Event('change', {bubbles:true})); } catch (e) {}
    ['keydown', 'keyup'].forEach(function(type) {
        try {
            input.dispatchEvent(new KeyboardEvent(type, {
                key: filterText.slice(-1) || '',
                bubbles: true,
                cancelable: true
            }));
        } catch (e) {}
    });
}
function findInput(labelText, placeholder) {
    var items = document.querySelectorAll('.el-form-item');
    var item = null;
    for (var i = 0; i < items.length; i++) {
        var t = (items[i].getAttribute('title') || '').trim();
        if (textMatch(labelText, t)) { item = items[i]; break; }
        var labels = items[i].querySelectorAll('label');
        for (var j = 0; j < labels.length; j++) {
            if (textMatch(labelText, (labels[j].textContent || '').trim())) { item = items[i]; break; }
        }
        if (item) break;
    }
    if (item) {
        var inp = placeholder ? item.querySelector('input[placeholder="' + placeholder + '"]') : null;
        if (!inp) inp = item.querySelector('input.el-input__inner, input');
        if (inp) return inp;
    }
    var labels2 = Array.from(document.querySelectorAll('pre.fr-label'));
    var labelPre = labels2.find(function(pre) {
        return textMatch(labelText, (pre.textContent || '').trim());
    });
    if (!labelPre) return null;
    var labelBox = labelPre.closest('[widgetname]') || labelPre.parentElement;
    if (!labelBox || !labelBox.parentElement) return null;
    var parent = labelBox.parentElement;
    var labelRect = labelBox.getBoundingClientRect();
    var editors = Array.from(parent.querySelectorAll('.fr-trigger-editor'));
    var best = null, bestScore = Number.POSITIVE_INFINITY;
    for (var k = 0; k < editors.length; k++) {
        var r = editors[k].getBoundingClientRect();
        var dy = Math.abs(r.top - labelRect.top);
        var dx = r.left - labelRect.right;
        if (dx < -5) continue;
        var score = dy * 1000 + Math.max(0, dx);
        if (score < bestScore) { bestScore = score; best = editors[k]; }
    }
    if (!best) return null;
    var finp = placeholder ? best.querySelector('input[placeholder="' + placeholder + '"]') : null;
    if (!finp) finp = best.querySelector('input.fr-trigger-texteditor, input');
    return finp || null;
}
var input = findInput(labelText, placeholder);
if (!input) return [];
var existing = closeStaleFRDropdowns(input);
if (existing) {
    applyFilter(input);
    var existingOpts = Array.from(existing.querySelectorAll('.el-select-dropdown__item, li[role="option"], .fr-combo-list-item'))
        .filter(function(e) { return e.offsetParent !== null; })
        .map(function(e) { return (e.getAttribute('title') || e.innerText || '').trim(); })
        .filter(Boolean);
    return existingOpts;
}
var before = getVisibleFRDropdowns();
var root = input.closest('.fr-trigger-editor');
if (root) {
    var center = root.querySelector('.fr-trigger-center, .fr-date-trigger-center');
    var btn = root.querySelector('.fr-trigger-btn-up, .fr-trigger-btn-over, .fr-trigger-btn-down, .frtriggerbtn-up, .frtriggerbtn-over, .frtriggerbtn-down');
    if (!fireMouse(center)) {
        if (!fireMouse(btn)) fireMouse(input);
    }
} else {
    fireMouse(input);
}
var dropdown = pickBestFRDropdown(input, before);
if (!dropdown || !dropdownAligned(dropdown, input)) return [];
applyFilter(input);
var scope = dropdown;
var opts = Array.from(scope.querySelectorAll('.el-select-dropdown__item, li[role="option"], .fr-combo-list-item'))
    .filter(function(e) { return e.offsetParent !== null; });
return opts.map(function(e) { return (e.getAttribute('title') || e.innerText || '').trim(); });
"""
    return _driver.execute_script(js, label_text, placeholder, want_checkbox, filter_text)


def _option_text_match(target, actual):
    """候选文本匹配：先精确，再做包含兜底。"""
    t = str(target or '').strip()
    a = str(actual or '').strip()
    if not t or not a:
        return False
    return a == t or t in a or a in t


def _fr_wait_options_ready(label_text, min_count=1, timeout=6.0, want_checkbox=True, close_on_ready=True, target_value='', filter_text=''):
    """等待级联下拉候选出现，主要用于省份->地市->县区。"""
    started = time.time()
    last_options = []
    while time.time() - started < timeout:
        try:
            opts = _fr_get_all_options(label_text, want_checkbox=want_checkbox, filter_text=filter_text)
            last_options = [str(x).strip() for x in (opts or []) if str(x).strip()]
            filtered = [x for x in last_options if x != '全选/不选']
            target_hit = True
            if str(target_value or '').strip():
                target_hit = any(_option_text_match(target_value, x) for x in filtered)
            if len(filtered) >= min_count and target_hit:
                if close_on_ready:
                    try:
                        _fr_commit_active_editor()
                    except Exception:
                        pass
                return {"ok": True, "options": filtered, "target_in_options": target_hit}
        except Exception:
            pass
        time.sleep(0.4)
    filtered = [x for x in last_options if x != '全选/不选']
    target_text = str(target_value or '').strip()
    target_hit = any(_option_text_match(target_text, x) for x in filtered) if target_text else False
    return {"ok": False, "options": last_options, "target_in_options": target_hit}


def _load_option_cache(label_text, log_file=None, scene='', field='', cache_key='', timeout=8.0):
    """打开普通下拉并缓存候选池，用于机构名称/机构类型的 sendkey 校验。"""
    started = time.time()
    last_options = []
    stable_count = 0
    last_signature = None
    while time.time() - started < timeout:
        try:
            opts = _fr_get_all_options(label_text, want_checkbox=False)
            options = [str(x).strip() for x in (opts or []) if str(x).strip() and str(x).strip() != '全选/不选']
            signature = '|'.join(options)
            if options and signature == last_signature:
                stable_count += 1
            else:
                stable_count = 0
                last_signature = signature
            if options:
                last_options = options
            if options and stable_count >= 2:
                break
        except Exception:
            pass
        time.sleep(0.4)
    try:
        _fr_commit_active_editor()
    except Exception:
        pass
    if log_file:
        _append_jsonl(log_file, {
            "ts": datetime.now().isoformat(timespec='seconds'),
            "type": "option_cache_load",
            "scene": scene,
            "field": field,
            "label": label_text,
            "cache_key": cache_key,
            "option_count": len(last_options),
            "options": last_options[:50],
        })
    return last_options


def _fr_commit_active_editor():
    """让当前 FineReport 编辑器失焦并触发 change，不使用 Enter，避免触发查询快捷键。"""
    try:
        _driver.execute_script("""
var el = document.activeElement;
if (el) {
  try {
    el.dispatchEvent(new Event('change', {bubbles:true}));
    el.blur();
  } catch (e) {}
}
var pane = document.querySelector('.reportPane, .parameter-container, body');
if (pane) {
  try { pane.dispatchEvent(new MouseEvent('mousedown', {bubbles:true, cancelable:true, view:window})); } catch (e) {}
  try { pane.dispatchEvent(new MouseEvent('mouseup', {bubbles:true, cancelable:true, view:window})); } catch (e) {}
  try { pane.dispatchEvent(new MouseEvent('click', {bubbles:true, cancelable:true, view:window})); } catch (e) {}
}
""")
        time.sleep(0.25)
    except Exception:
        pass
    try:
        from selenium.webdriver.common.action_chains import ActionChains
        blank = _driver.execute_script("""
var btn = document.querySelector("[widgetname='SEARCH'] button, [widgetname='SEARCH'], button.fr-btn-text");
if (btn) return btn;
return document.querySelector('.reportPane, #content-container, body') || document.body;
""")
        if blank:
            if _driver.execute_script("return !!(arguments[0] && arguments[0].closest && arguments[0].closest('[widgetname=\"SEARCH\"]'))", blank):
                rect = _driver.execute_script("""
var r = arguments[0].getBoundingClientRect();
return {w:r.width, h:r.height};
""", blank)
                offset_x = max(1, min(10, int((rect.get('w') or 20) / 2)))
                offset_y = -max(24, int((rect.get('h') or 21) * 1.2))
                ActionChains(_driver).move_to_element_with_offset(blank, offset_x, offset_y).click().perform()
            else:
                ActionChains(_driver).move_to_element_with_offset(blank, 20, 20).click().perform()
            time.sleep(0.25)
    except Exception:
        pass


def _fr_accept_possible_dialog():
    """处理查询后可能出现的浏览器 alert 或页面确认按钮。"""
    try:
        alert = _driver.switch_to.alert
        text = alert.text
        alert.accept()
        time.sleep(0.3)
        return {"accepted": True, "type": "alert", "text": text}
    except Exception:
        pass

    try:
        js = """
var texts = ['确定', '确认', 'OK', 'ok'];
var nodes = Array.from(document.querySelectorAll('button, .fr-btn-text, .el-button, span, div'));
function visible(el) {
  if (!el || el.offsetParent === null) return false;
  var s = window.getComputedStyle(el);
  if (s.display === 'none' || s.visibility === 'hidden' || Number(s.opacity || '1') === 0) return false;
  var r = el.getBoundingClientRect();
  return r.width > 0 && r.height > 0;
}
for (var i = 0; i < nodes.length; i++) {
  var t = (nodes[i].innerText || nodes[i].textContent || '').trim();
  if (visible(nodes[i]) && texts.indexOf(t) >= 0) {
    try { nodes[i].click(); } catch (e) {}
    return {accepted:true, type:'page_button', text:t};
  }
}
return {accepted:false};
"""
        res = _driver.execute_script(js)
        if res and res.get("accepted"):
            time.sleep(0.3)
            return res
    except Exception:
        pass
    return {"accepted": False}


def _fr_click_query():
    """点击查询按钮。"""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    # 优先命中 FineReport 参数区 SEARCH 控件
    fr_btns = _driver.find_elements(
        By.XPATH,
        "//*[@widgetname='SEARCH']//button[contains(normalize-space(.), '查询')]"
    )
    for btn in fr_btns:
        try:
            if btn.is_displayed() and btn.is_enabled():
                btn.click()
                return True
        except Exception:
            continue

    btns = _driver.find_elements(By.XPATH, "//*[normalize-space(text())='查询']")
    for btn in btns:
        try:
            if btn.is_displayed() and btn.is_enabled():
                btn.click()
                return True
        except Exception:
            continue
    # fallback: 找所有可见的 button/span 包含"查询"文本
    try:
        fallback_btn = WebDriverWait(_driver, 0.5).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(normalize-space(.), '查询')]"))
        )
        fallback_btn.click()
        return True
    except Exception:
        pass
    try:
        fallback_btn = WebDriverWait(_driver, 0.5).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '查询')]/parent::button | //span[contains(text(), '查询')]"))
        )
        fallback_btn.click()
        return True
    except Exception:
        pass
    return False


def _fr_is_loading():
    """检测页面是否处于查询加载中状态。"""
    js = r"""
const selectors = [
  '.loading-whole-indicator',
  '.load-indicator',
  '.el-loading-mask',
  '.el-loading-spinner',
  '.el-icon-loading',
  '.is-loading',
  '[aria-busy="true"]'
];
function visible(el) {
  if (!el) return false;
  const style = window.getComputedStyle(el);
  if (style.display === 'none' || style.visibility === 'hidden' || Number(style.opacity || '1') === 0) return false;
  const rect = el.getBoundingClientRect();
  return rect.width > 0 && rect.height > 0;
}
for (const sel of selectors) {
  const nodes = document.querySelectorAll(sel);
  for (const n of nodes) {
    if (visible(n)) return true;
  }
}
return false;
"""
    try:
        return bool(_driver.execute_script(js))
    except Exception:
        return False


def _fr_is_search_button_ready():
    """检测参数区 SEARCH 查询按钮是否恢复可用。"""
    js = r"""
const root = document.querySelector('[widgetname="SEARCH"]');
if (!root) return null;
const cls = root.className || '';
if (cls.includes('ui-state-disabled')) return false;
if (cls.includes('ui-state-enabled')) return true;
const btn = root.querySelector('button');
if (!btn) return true;
return !btn.disabled;
"""
    try:
        return _driver.execute_script(js)
    except Exception:
        return None


def _fr_wait_query_done(max_wait_seconds, settle_seconds=0.8):
    """等待查询结束（尽量基于加载状态动态判断，而不是固定长等待）。"""
    timeout = max(2.0, float(max_wait_seconds))
    settle = max(0.3, float(settle_seconds))
    started = time.time()
    stable_since = None
    saw_loading = False

    while True:
        now = time.time()
        elapsed = now - started
        loading = _fr_is_loading()
        search_ready = _fr_is_search_button_ready()
        if loading:
            saw_loading = True
            stable_since = None
        else:
            if stable_since is None:
                stable_since = now
            stable_elapsed = now - stable_since
            # 若能读取 SEARCH 状态，则要求按钮恢复可用；否则仅依赖 loading 消失
            search_ok = (search_ready is None) or bool(search_ready)
            # 若见过 loading，则要求连续稳定 settle 秒；若没见过，至少给 1 秒缓冲
            if stable_elapsed >= settle and (saw_loading or elapsed >= 1.0) and search_ok:
                return {
                    "ok": True,
                    "elapsed_seconds": round(elapsed, 3),
                    "saw_loading": saw_loading,
                    "loading": False,
                    "search_ready": search_ready,
                    "timeout": False,
                }

        if elapsed >= timeout:
            return {
                "ok": (not loading) and ((search_ready is None) or bool(search_ready)),
                "elapsed_seconds": round(elapsed, 3),
                "saw_loading": saw_loading,
                "loading": bool(loading),
                "search_ready": search_ready,
                "timeout": True,
            }
        time.sleep(0.15)


def _query_wait_indicates_stall(wait_result):
    """仅在等待耗尽且页面仍显示查询加载时判定为卡住。"""
    if not isinstance(wait_result, dict) or not wait_result.get("timeout"):
        return False
    return bool(wait_result.get("loading", wait_result.get("saw_loading", False)))


def _should_skip_existing_matrix_field(current_value, target_value, want_clear=False):
    """按当前页面实际字段值判断是否可跳过设置。"""
    current = str(current_value or '').strip()
    target = str(target_value or '').strip()
    return not current if want_clear else current == target


def _field_settle_delay(field_method, set_mode, configured_seconds):
    """返回字段成功设置后等待前端提交/联动的秒数。"""
    method = str(field_method or '').strip().lower()
    configured = max(0.0, float(configured_seconds or 0))
    if configured > 0 and method in {'sendkey', 'checkbox_dropdown'}:
        return configured
    if method == 'sendkey':
        return 0.35
    if str(set_mode or '').strip().lower() == 'checkbox_dropdown':
        return 0.6
    return 0.2


def walk_scene(fields, idx=0, current=None, results=None, max_items=0):
    """递归遍历所有字段组合，生成查询坐标列表。"""
    if results is None:
        results = []
    if current is None:
        current = {}

    field_names = list(fields.keys())
    if idx >= len(field_names):
        results.append(dict(current))
        return results

    fname = field_names[idx]
    finfo = fields[fname]
    mode = finfo['mode']
    vals = finfo['values']

    if mode == 'fixed':
        current[fname] = vals[0] if vals else ''
        walk_scene(fields, idx + 1, current, results, max_items)
    elif mode == 'list':
        for v in vals:
            if max_items > 0 and len(results) >= max_items:
                break
            current[fname] = v
            walk_scene(fields, idx + 1, current, results, max_items)
    return results


def _recover_report(scene, plan, validate_label=''):
    """恢复报表页面: 直接重走登录/门户/tree_path/iframe。"""
    return _full_recover_report(
        scene,
        plan,
        validate_label=validate_label,
        expand_tree_on_retry=True,
        restart_browser=True,
    )


def _handle_matrix_item_exception(exc, scene, plan, seq, log_file, record_item_summary, item_started):
    """Handle an unexpected per-item crash and try to restore the report page."""
    msg = str(exc)
    print(f'  ✗ 当前条目异常: {msg}')
    _append_jsonl(log_file, {
        "ts": datetime.now().isoformat(timespec='seconds'),
        "type": "step_error",
        "seq": seq,
        "scene": scene,
        "error": msg,
        "stage": "item_exception",
    })
    print('  🔄 尝试恢复页面后继续下一条...')
    if _recover_report(scene, plan):
        record_item_summary(seq, item_started, False, "item_exception_recovered")
        print('  ✓ 页面已恢复，继续下一条')
        return True

    record_item_summary(seq, item_started, False, "item_exception_recovery_failed")
    print('  ✗ 页面恢复失败，继续记录失败')
    return False


def _is_recoverable_field_lookup_error(msg):
    """字段输入区整体找不到时，恢复入口可尝试重建报表页面。"""
    text = str(msg or '').strip().lower()
    return text in {'input_not_found', 'reportbox_not_found'}


def _download_result_ok(download_res):
    """下载必须明确成功且至少确认到一个新增文件。"""
    if not isinstance(download_res, dict):
        return False
    return bool(download_res.get("ok")) and bool(download_res.get("files") or [])


def _should_abort_after_field_lookup_retries(msg, use_recovery):
    """增强恢复入口连续找不到报表输入区时，应停止批量而不是继续空转。"""
    return bool(use_recovery) and _is_recoverable_field_lookup_error(msg)


def _recovery_step_passed(recovered, field_available=True):
    """恢复步骤必须完成，且可选字段验证通过才算成功。"""
    return bool(recovered) and bool(field_available)


def _report_field_available(label_text):
    """验证当前报表 iframe 中能找到指定字段输入区。"""
    if not str(label_text or '').strip():
        return True
    try:
        probe = _fr_set_field(label_text, '__probe__', field_type='text', select_strategy='probe')
        if isinstance(probe, dict):
            return probe.get('msg') != 'input_not_found'
    except Exception:
        pass
    try:
        return _fr_get_field_value(label_text) is not None
    except Exception:
        return False


def _click_tree_path_leaf(tree_path):
    """全量展开后显式点击 tree_path 最末级叶子，触发报表 iframe 加载。"""
    keywords = [part.strip() for part in str(tree_path or '').split('>') if part.strip()]
    if not keywords or _driver is None:
        return False
    leaf_keyword = keywords[-1]
    try:
        result = _driver.execute_script("""
var keyword = arguments[0];
var nodes = Array.from(document.querySelectorAll('.el-tree-node')).filter(function(node) {
    if (!node || node.offsetParent === null) return false;
    var icon = node.querySelector(':scope > .el-tree-node__content .el-tree-node__expand-icon');
    return !!(icon && icon.classList.contains('is-leaf'));
});
var matches = nodes.filter(function(node) {
    var texts = [];
    var titleEl = node.querySelector('.org_tree_node [title]');
    if (titleEl) texts.push(titleEl.getAttribute('title') || '');
    node.querySelectorAll(':scope > .el-tree-node__content span').forEach(function(span) {
        texts.push((span.textContent || '').trim());
    });
    return texts.some(function(text) { return text === keyword; }) ||
        texts.some(function(text) { return text && (text.indexOf(keyword) >= 0 || keyword.indexOf(text) >= 0); });
});
if (!matches.length) return {ok:false, error:'leaf_not_found'};
var node = matches[0];
var content = node.querySelector(':scope > .el-tree-node__content');
if (!content) return {ok:false, error:'leaf_content_not_found'};
try { content.click(); } catch (e) { return {ok:false, error:'leaf_click_failed'}; }
return {ok:true, matched:(content.textContent || '').trim(), count:matches.length};
""", leaf_keyword)
        if result and result.get('ok'):
            hint = f' ({result.get("count", 1)}个匹配)' if result.get('count', 1) > 1 else ''
            print(f'  ✓ 全量展开后点击末级报表: {result.get("matched", leaf_keyword)}{hint}')
            return True
        print(f'  ⚠️ 全量展开后未能点击末级报表: {leaf_keyword}')
    except Exception as e:
        print(f'  ⚠️ 全量展开后点击末级报表异常: {e}')
    return False


def _wait_for_recovery_step(step_name, seconds=1.2):
    """给门户/报表 SPA 在恢复阶段间留出稳定渲染时间。"""
    print(f'  ⏳ {step_name}完成，等待 {seconds:.1f} 秒稳定...')
    time.sleep(seconds)


def _full_recover_report(scene, plan, validate_label='', expand_tree_on_retry=True, restart_browser=False):
    """强恢复: 登录/门户导航/树路径/iframe/字段验证。"""
    tree_path = plan.get('tree_path', '') if plan else ''
    print('  🔄 强恢复：重新确认登录并回到报表页...')

    if restart_browser:
        if not _restart_browser_session():
            print('  ✗ 强恢复失败：浏览器重启失败')
            return False
        _wait_for_recovery_step('浏览器重启')
        auto_login()
        _wait_for_recovery_step('自动登录')

    try:
        _driver.switch_to.default_content()
    except Exception:
        pass

    if not _ensure_logged_in():
        print('  ✗ 强恢复失败：登录状态无法恢复')
        return False
    _wait_for_recovery_step('登录状态确认')
    if not _navigate_to_report_page():
        print('  ✗ 强恢复失败：无法进入报表服务中心/共享数据查看')
        return False
    _wait_for_recovery_step('报表服务中心导航')

    if tree_path:
        navigated = _navigate_tree_to_report(tree_path)
        expanded_tree = False
        if not navigated and expand_tree_on_retry:
            print('  ⚠️ tree_path 首次导航失败，尝试全量展开树后重试...')
            _expand_all_report_tree()
            _wait_for_recovery_step('报表树全量展开')
            navigated = _navigate_tree_to_report(tree_path)
            expanded_tree = True
        if not navigated:
            print('  ✗ 强恢复失败：tree_path 无法打开目标报表')
            return False
        if expanded_tree and _click_tree_path_leaf(tree_path):
            _wait_for_recovery_step('末级报表节点点击')
        _wait_for_recovery_step('tree_path 导航')

    if not _switch_to_report_frame(scene, plan):
        print('  ✗ 强恢复失败：无法切入目标报表 iframe')
        return False
    _wait_for_recovery_step('报表 iframe 切换')
    _save_iframe_src()

    if validate_label and not _report_field_available(validate_label):
        print(f'  ✗ 强恢复失败：目标字段仍不可用: {validate_label}')
        return False

    print('  ✅ 强恢复成功')
    return True


def _matrix_recovery_enabled(recovery_mode):
    mode = str(recovery_mode or 'plain').strip().lower()
    if mode in ('', 'plain'):
        return False
    if mode == 'recover':
        return True
    raise ValueError(f'未知恢复模式: {recovery_mode}')


def run_matrix_export(scene, recovery_mode='plain'):
    """按场景配置执行矩阵查询导出。

    字段填充方式由 config.xlsx 对应 sheet 的 field.<字段>.method 控制。
    """
    try:
        use_recovery = _matrix_recovery_enabled(recovery_mode)
    except ValueError as e:
        print(f'❌ {e}')
        return

    if not driver_alive():
        print('❌ 浏览器未启动，请先执行「1. 启动浏览器」')
        return

    field_methods_mode = 'config'

    # 加载配置
    config = load_task_config(scene)
    if not config:
        return
    try:
            plan = parse_matrix_config(config, scene=scene)
    except ValueError as e:
        print(f'❌ 配置解析失败: {e}')
        return

    fields = plan['fields']
    if not fields:
        print('❌ 配置中无有效字段')
        return

    # 组合遍历顺序也按场景优先级重排
    preferred = [f for f in SCENE_MATRIX_ORDER.get(scene, []) if f in fields]
    tail = [f for f in fields.keys() if f not in preferred]
    ordered_keys = preferred + tail
    fields = {k: fields[k] for k in ordered_keys}

    query_enabled = plan['query']
    export_enabled = plan['export']
    query_wait = plan['query_wait_seconds']
    field_settle_seconds = plan.get('field_settle_seconds', 0.0)
    max_items = plan['max_items']
    refresh_interval = plan.get('refresh_interval', 30) if use_recovery else 0

    print(f'\\n场景: {scene}')
    print(f'  恢复兜底: {"启用" if use_recovery else "关闭"}')
    print('  字段填充方式: config.xlsx 的 field.xxx.method')
    print(f'  字段: {", ".join(fields.keys())}')
    print(f'  查询: {"是" if query_enabled else "否"}  |  导出: {"是" if export_enabled else "否"}')
    if field_settle_seconds > 0:
        print(f'  字段设置后等待: {field_settle_seconds:.1f} 秒')
    if max_items > 0:
        print(f'  最大执行次数: {max_items}')

    # 生成查询坐标
    coords = walk_scene(fields, max_items=max_items)
    total = len(coords)
    print(f'\\n共 {total} 个查询组合')

    if total == 0:
        print('❌ 无查询组合')
        return

    # 重置页面状态：关闭任何残留菜单/弹窗
    try:
        from selenium.webdriver.common.action_chains import ActionChains
        ActionChains(_driver).send_keys('\ue00c').perform()
        time.sleep(0.2)
    except Exception:
        pass

    # 切换到报表 iframe
    if not _switch_to_report_frame(scene, plan):
        return
    _save_iframe_src()

    # 准备输出目录
    run_id = f'{scene}_{field_methods_mode}_{datetime.now().strftime("%Y%m%d_%H%M%S")}_{uuid.uuid4().hex[:8]}'
    run_dir = OUTPUT / run_id
    log_file = run_dir / 'run_log.jsonl'
    run_dir.mkdir(parents=True, exist_ok=True)
    run_started = time.time()
    run_started_iso = datetime.now().isoformat(timespec='seconds')

    ok_count = 0
    fail_count = 0
    option_cache = {}
    abort_reason = ""

    def validate_from_option_cache(fname, label, target_value, seq=None):
        cache_key = f'{scene}|{fname}'
        cached_options = option_cache.get(cache_key, [])
        matched_option = next((opt for opt in cached_options if _option_text_match(target_value, opt)), '')
        target_in_options = bool(matched_option)
        warning = "" if target_in_options else "sendkey_target_not_in_options"
        _append_jsonl(log_file, {
            "ts": datetime.now().isoformat(timespec='seconds'),
            "type": "sendkey_option_validate",
            "seq": seq,
            "scene": scene,
            "field": fname,
            "label": label,
            "target_value": target_value,
            "target_in_options": target_in_options,
            "matched_option": matched_option,
            "option_count": len(cached_options),
            "options": cached_options[:50],
            "cache_key": cache_key,
            "warning": warning,
        })
        if not target_in_options:
            print(f'  ⚠️ {fname} sendkey候选未匹配: {target_value}，可能查询无数据')
        return target_in_options

    def record_item_summary(seq, item_started, ok, status):
        item_elapsed = round(time.time() - item_started, 3)
        _append_jsonl(log_file, {
            "ts": datetime.now().isoformat(timespec='seconds'),
            "type": "item_summary",
            "seq": seq,
            "scene": scene,
            "field_methods_mode": field_methods_mode,
            "ok": bool(ok),
            "status": status,
            "elapsed_seconds": item_elapsed,
        })

    for fname, finfo in fields.items():
        if fname not in ('institution_class', 'institution_name'):
            continue
        if str(finfo.get('method', _default_field_method(scene, fname)) or '').strip().lower() != 'sendkey':
            continue
        if not bool(finfo.get('validate_options', False)):
            continue
        label = REPORT_LABEL_MAP.get(fname, fname)
        cache_key = f'{scene}|{fname}'
        option_cache[cache_key] = _load_option_cache(
            label,
            log_file=log_file,
            scene=scene,
            field=fname,
            cache_key=cache_key,
        )
        print(f'  ✓ 已预加载 {fname} 候选池: {len(option_cache[cache_key])} 项')

    interrupted = False
    interrupt_ts = None

    consecutive_errors = 0
    items_since_refresh = 0
    MAX_CONSECUTIVE_ERRORS = 3
    query_stall_retry_counts = {}

    # 遍历所有查询组合
    try:
        work_items = list(enumerate(coords, 1))
        for run_index, (seq, coord) in enumerate(work_items):
            try:
                item_started = time.time()
                print(f'\\n[ {seq}/{total} ] {coord}')
    
                # ── 页面存活检查 ──
                if use_recovery and not _check_page_alive():
                    print('  ⚠️ 页面无响应，尝试恢复...')
                    _append_jsonl(log_file, {
                        "ts": datetime.now().isoformat(timespec='seconds'),
                        "type": "page_recovery", "seq": seq, "scene": scene,
                        "reason": "page_not_alive",
                    })
                    if _recover_report(scene, plan):
                        items_since_refresh = 0
                        consecutive_errors = 0
                    else:
                        print('  ✗ 页面恢复失败，跳过当前条目')
                        fail_count += 1
                        record_item_summary(seq, item_started, False, "page_dead")
                        continue
    
                # ── 定时刷新 ──
                if refresh_interval > 0 and items_since_refresh >= refresh_interval:
                    print(f'  🔄 已处理 {items_since_refresh} 条，执行定时强恢复...')
                    _append_jsonl(log_file, {
                        "ts": datetime.now().isoformat(timespec='seconds'),
                        "type": "page_refresh", "seq": seq, "scene": scene,
                        "reason": "periodic",
                        "items_since_refresh": items_since_refresh,
                    })
                    if _recover_report(scene, plan):
                        print('  ✓ 定时强恢复完成')
                        items_since_refresh = 0
                    else:
                        print('  ⚠️ 定时强恢复失败，继续尝试...')
    
                # 防止上一轮查询尚未结束，先做一次短等待
                if _fr_is_loading():
                    print('  检测到上次查询仍在加载，先等待结束...')
                    prev_wait = _fr_wait_query_done(max(3.0, min(query_wait, 10.0)))
                    prev_elapsed = prev_wait.get("elapsed_seconds", 0)
                    if prev_wait.get("timeout"):
                        print(f'  ⚠️ 上次查询加载状态未完全消失（{prev_elapsed:.2f}s），继续尝试当前轮')
                    else:
                        print(f'  ✓ 上次查询已结束（{prev_elapsed:.2f}s）')
    
                # 字段执行顺序: 场景优先级在前，其余字段兜底在后
                ordered_fields = [f for f in SCENE_FILL_ORDER.get(scene, []) if f in coord]
                ordered_fields.extend([f for f in coord.keys() if f not in ordered_fields])
                first_row_cascade_init_done = set()
    
                # 每条任务都重新设置所有字段，确保下拉选择已真正提交到报表内部状态。
                all_ok = True
                fatal_field_lookup_failure = False
                last_field_lookup_error = ""
                max_verify_rounds = 4 if use_recovery else 3
                for vround in range(1, max_verify_rounds + 1):
                    round_all_matched = True
                    recovered_field_lookup = False
                    if vround > 1:
                        print(f'  进行第{vround}轮查询前校验补填...')
                    for fname in ordered_fields:
                        fval = coord.get(fname, '')
                        label = REPORT_LABEL_MAP.get(fname, fname)
                        if fname == 'data_date':
                            fval = _normalize_date_text(fval)
                        fval_str = str(fval or '').strip()
                        want_clear = (fval_str == '不选')
                        exec_value = '' if want_clear else fval_str
                        finfo = fields.get(fname, {})
                        raw_mode = str(finfo.get('mode', '') or '').strip().lower()
                        field_method = str(finfo.get('method', _default_field_method(scene, fname)) or '').strip().lower()
                        validate_options = bool(finfo.get('validate_options', False)) and fname in ('institution_class', 'institution_name')
                        if fname == 'data_date':
                            ftype = 'date'
                            field_method = 'sendkey'
                        elif raw_mode in ('fixed', 'list'):
                            ftype = 'select'
                        else:
                            ftype = 'select'

                        current_value = _fr_get_field_value(label)
                        if _should_skip_existing_matrix_field(current_value, exec_value, want_clear=want_clear):
                            print(f'  ✓ {fname}={exec_value if exec_value else "清空"} (当前页面已匹配，跳过)')
                            _append_jsonl(log_file, {
                                "ts": datetime.now().isoformat(timespec='seconds'),
                                "type": "field_skip", "seq": seq, "scene": scene,
                                "field": fname, "label": label,
                                "target_value": exec_value, "actual_value": str(current_value or '').strip(), "matched": True,
                                "field_method": field_method,
                                "reason": "page_value_matched",
                            })
                            continue
    
                        if seq == 1 and vround == 1:
                            init_map = {
                                'province': 'province',
                                'city': 'city',
                                'county': 'county',
                            }
                            if fname in init_map and fname not in first_row_cascade_init_done:
                                current_before_init = _fr_get_field_value(label)
                                if not str(current_before_init or '').strip() and field_method == 'checkbox_dropdown':
                                    wait_ready = _fr_wait_options_ready(label, min_count=1, timeout=8.0, want_checkbox=(field_method == 'checkbox_dropdown'))
                                    _append_jsonl(log_file, {
                                        "ts": datetime.now().isoformat(timespec='seconds'),
                                        "type": "cascade_init",
                                        "scene": scene,
                                        "field": fname,
                                        "ok": bool(wait_ready.get("ok")),
                                        "method": field_method,
                                        "want_checkbox": field_method == 'checkbox_dropdown',
                                        "option_count": len(wait_ready.get("options", []) or []),
                                        "options": wait_ready.get("options", [])[:50],
                                    })
                                    if wait_ready.get("ok"):
                                        print(f'  ✓ {fname} 第一条候选初始化完成')
                                    else:
                                        print(f'  ⚠️ {fname} 第一条候选初始化未确认，继续原有重试')
                                first_row_cascade_init_done.add(fname)
    
                        # FineReport 级联下拉有异步加载：先点开当前字段并等待候选出现，
                        # 再进入真正设置流程，避免点三角后立刻读取到空候选或上一个字段的残留候选。
                        pre_wait_dropdown = field_method == 'checkbox_dropdown'
                        if pre_wait_dropdown and exec_value:
                            want_checkbox_options = field_method == 'checkbox_dropdown'
                            strict_option_match = True
                            wait_ready = _fr_wait_options_ready(
                                label,
                                min_count=1,
                                timeout=8.0,
                                want_checkbox=want_checkbox_options,
                                close_on_ready=False,
                                target_value=exec_value if strict_option_match else '',
                            )
                            _append_jsonl(log_file, {
                                "ts": datetime.now().isoformat(timespec='seconds'),
                                "type": "pre_field_options_ready",
                                "scene": scene,
                                "field": fname,
                                "label": label,
                                "ok": bool(wait_ready.get("ok")),
                                "target_value": exec_value if strict_option_match else '',
                                "target_in_options": wait_ready.get("target_in_options"),
                                "option_count": len(wait_ready.get("options", []) or []),
                                "options": wait_ready.get("options", [])[:50],
                                "want_checkbox": want_checkbox_options,
                                "field_method": field_method,
                                "set_mode": field_method,
                                "error": "" if wait_ready.get("ok") else ("target_not_in_options" if strict_option_match else "options_not_ready"),
                            })
                            if wait_ready.get("ok"):
                                print(f'  ✓ {fname} 候选已打开并就绪')
                            else:
                                if strict_option_match:
                                    print(f'  ✗ {fname} 候选暂未包含配置值: {exec_value}，等待后重试本条')
                                    round_all_matched = False
                                    break
                                print(f'  ⚠️ {fname} 候选等待超时，继续进入原有重试流程')
    
                        result = None
                        retry_errors = {
                            'fr_checkbox_dropdown_not_found',
                            'checkbox_dropdown_not_match',
                            'tag_clear_incomplete',
                            'sendkey_select_not_committed',
                        }
                        for attempt in range(1, 6):
                            result = _fr_set_field(label, exec_value, field_type=ftype, select_strategy=field_method)
                            msg = result.get('msg', '') if result else 'exec_error'
                            if result and result.get('ok'):
                                break
                            if msg not in retry_errors or attempt == 5:
                                break
                            opts = result.get('options', []) if isinstance(result, dict) else []
                            opt_hint = f' 候选={opts[:6]}' if opts else ''
                            print(f'  ↻ {fname} 第{attempt}次设置未命中({msg})，等待联动后重试...{opt_hint}')
                            _fr_commit_active_editor()
                            time.sleep(1.2 if fname in {'province', 'city', 'county'} else 0.8)
    
                        set_mode_for_delay = result.get('set_mode', '') if result else ''
                        time.sleep(_field_settle_delay(field_method, set_mode_for_delay, field_settle_seconds))
    
                        if result and result.get('ok'):
                            actual = result.get('actual', '')
                            set_mode = result.get('set_mode', field_method if ftype == 'select' else ftype)
                            print(f'  ✓ {fname}={exec_value if exec_value else "清空"} (实际: {actual})')
                            _append_jsonl(log_file, {
                                "ts": datetime.now().isoformat(timespec='seconds'),
                                "type": "field_set", "seq": seq, "scene": scene,
                                "field": fname, "label": label,
                                "target_value": exec_value, "actual_value": actual, "matched": True,
                                "verify_round": vround,
                                "field_method": field_method,
                                "set_mode": set_mode,
                            })
                            if field_method == 'sendkey' and validate_options and exec_value:
                                validate_from_option_cache(fname, label, exec_value, seq=seq)
                            if fname in ('province', 'city') and field_method == 'checkbox_dropdown':
                                next_fname = 'city' if fname == 'province' else 'county'
                                next_label = REPORT_LABEL_MAP.get(next_fname, next_fname)
                                next_method = str(fields.get(next_fname, {}).get('method', _default_field_method(scene, next_fname)) or '').strip().lower()
                                wait_ready = _fr_wait_options_ready(next_label, min_count=1, timeout=8.0, want_checkbox=(next_method == 'checkbox_dropdown'))
                                _append_jsonl(log_file, {
                                    "ts": datetime.now().isoformat(timespec='seconds'),
                                    "type": "cascade_ready",
                                    "scene": scene,
                                    "from_field": fname,
                                    "next_field": next_fname,
                                    "ok": bool(wait_ready.get("ok")),
                                    "method": next_method,
                                    "want_checkbox": next_method == 'checkbox_dropdown',
                                    "option_count": len(wait_ready.get("options", []) or []),
                                    "options": wait_ready.get("options", [])[:50],
                                })
                                if wait_ready.get("ok"):
                                    print(f'  ✓ {next_fname} 候选已就绪')
                                else:
                                    print(f'  ⚠️ {next_fname} 候选未确认就绪，继续重试流程')
                        else:
                            round_all_matched = False
                            msg = result.get('msg', 'js_error') if result else 'exec_error'
                            set_mode = result.get('set_mode', field_method if ftype == 'select' else ftype) if result else (field_method if ftype == 'select' else ftype)
                            print(f'  ✗ {fname}={exec_value if exec_value else "清空"} | {msg}')
                            _append_jsonl(log_file, {
                                "ts": datetime.now().isoformat(timespec='seconds'),
                                "type": "field_set", "seq": seq, "scene": scene,
                                "field": fname, "label": label,
                            "target_value": exec_value, "actual_value": '', "matched": False,
                            "error": msg, "verify_round": vround,
                            "field_method": field_method,
                            "set_mode": set_mode,
                                "options": result.get('options', []) if isinstance(result, dict) else [],
                            })
                            if use_recovery and _is_recoverable_field_lookup_error(msg):
                                last_field_lookup_error = msg
                                _append_jsonl(log_file, {
                                    "ts": datetime.now().isoformat(timespec='seconds'),
                                    "type": "page_recovery", "seq": seq, "scene": scene,
                                    "reason": msg,
                                    "stage": "field_lookup",
                                    "field": fname,
                                    "label": label,
                                })
                                print(f'  🔄 {fname} 输入框不可用({msg})，尝试恢复报表页面...')
                                if _recover_report(scene, plan, validate_label=label):
                                    print('  ✓ 页面已恢复，重新开始本条查询前校验')
                                    consecutive_errors = 0
                                    items_since_refresh = 0
                                    recovered_field_lookup = True
                                    round_all_matched = False
                                    all_ok = True
                                    break
                                print('  ✗ 页面恢复失败，当前条目按字段设置失败处理')
                                fatal_field_lookup_failure = _should_abort_after_field_lookup_retries(msg, use_recovery)
    
                    if recovered_field_lookup:
                        if vround == max_verify_rounds:
                            print('  ⚠️ 页面已恢复，但校验轮次已用完，停止本轮批量避免继续空转')
                            all_ok = False
                            fatal_field_lookup_failure = _should_abort_after_field_lookup_retries(last_field_lookup_error, use_recovery)
                        continue

                    if round_all_matched:
                        all_ok = True
                        if vround > 1:
                            print('  ✓ 查询前校验完成：全部字段已匹配')
                        break
    
                    if vround == max_verify_rounds:
                        print('  ⚠️ 查询前校验结束：仍有字段未匹配')
                        all_ok = False
    
                if not all_ok:
                    fail_count += 1
                    if query_enabled:
                        print('  字段设置失败，跳过查询/导出')
                        record_item_summary(seq, item_started, False, "field_set_failed")
                        if fatal_field_lookup_failure:
                            abort_reason = f'field_lookup_unrecoverable:{last_field_lookup_error or "unknown"}'
                            _append_jsonl(log_file, {
                                "ts": datetime.now().isoformat(timespec='seconds'),
                                "type": "run_abort",
                                "seq": seq,
                                "scene": scene,
                                "reason": abort_reason,
                            })
                            print(f'  ❌ 报表输入区恢复失败，停止后续批量: {abort_reason}')
                            break
                        continue
    
                if not query_enabled:
                    print(f'  跳过查询（task.query=no）')
                    ok_count += 1
                    record_item_summary(seq, item_started, True, "query_disabled")
                    consecutive_errors = 0
                    items_since_refresh += 1
                    continue
    
                # 点击查询
                if not _fr_click_query():
                    print('  ✗ 未找到可点击的"查询"按钮')
                    _append_jsonl(log_file, {
                        "ts": datetime.now().isoformat(timespec='seconds'),
                        "type": "step_error", "seq": seq, "scene": scene,
                        "error": "query_button_not_found",
                    })
                    consecutive_errors += 1
                    if use_recovery and _recover_report(scene, plan) and _fr_click_query():
                        print('  ✓ 恢复后查询按钮已可用')
                        consecutive_errors = 0
                        items_since_refresh = 0
                    elif not use_recovery:
                        print('  ✗ 正式入口不执行恢复兜底，跳过当前条目')
                        fail_count += 1
                        record_item_summary(seq, item_started, False, "query_button_not_found")
                        continue
                    else:
                        print('  ✗ 恢复失败，跳过当前条目')
                        fail_count += 1
                        record_item_summary(seq, item_started, False, "query_button_not_found")
                        continue
    
                print(f'  等待查询完成（最大 {query_wait:.1f} 秒）...')
                wait_res = _fr_wait_query_done(query_wait)
                elapsed = wait_res.get("elapsed_seconds", query_wait)
                query_stalled = _query_wait_indicates_stall(wait_res)
                dialog_res = _fr_accept_possible_dialog() if not query_stalled else {"accepted": False}
                if wait_res.get("timeout"):
                    if query_stalled:
                        print(f'  ⚠️ 查询加载超过 {elapsed:.2f}s，判定页面卡住')
                    else:
                        print(f'  ⚠️ 查询等待超时（{elapsed:.2f}s），页面已不在加载，继续后续流程')
                else:
                    print(f'  ✓ 查询完成（{elapsed:.2f}s）')
                if dialog_res.get("accepted"):
                    print(f'  ✓ 已处理查询后确认: {dialog_res.get("text", "")}')
    
                _append_jsonl(log_file, {
                    "ts": datetime.now().isoformat(timespec='seconds'),
                    "type": "query_done", "seq": seq, "scene": scene,
                    "field_methods_mode": field_methods_mode,
                    "query_wait_seconds": query_wait,
                    "query_elapsed_seconds": elapsed,
                    "query_timeout": bool(wait_res.get("timeout")),
                    "query_saw_loading": bool(wait_res.get("saw_loading")),
                    "query_loading_at_timeout": bool(wait_res.get("loading")),
                    "query_stalled": query_stalled,
                    "query_search_ready": wait_res.get("search_ready"),
                    "dialog": dialog_res,
                })

                if query_stalled and use_recovery:
                    retry_count = query_stall_retry_counts.get(seq, 0) + 1
                    query_stall_retry_counts[seq] = retry_count
                    _append_jsonl(log_file, {
                        "ts": datetime.now().isoformat(timespec='seconds'),
                        "type": "page_recovery", "seq": seq, "scene": scene,
                        "reason": "query_stalled",
                        "query_wait_seconds": query_wait,
                        "query_elapsed_seconds": elapsed,
                        "retry_count": retry_count,
                    })
                    if retry_count <= MAX_CONSECUTIVE_ERRORS:
                        print(f'  🔄 查询仍在加载，执行强恢复后重试当前条（第 {retry_count}/{MAX_CONSECUTIVE_ERRORS} 次）...')
                        validate_label = REPORT_LABEL_MAP.get(ordered_fields[0], '') if ordered_fields else ''
                        if _recover_report(scene, plan, validate_label=validate_label):
                            print(f'  ✓ 页面已恢复，将重新执行 [{seq}/{total}]')
                            work_items.insert(run_index + 1, (seq, coord))
                            consecutive_errors = 0
                            items_since_refresh = 0
                            continue

                    print('  ✗ 查询卡住后的页面恢复失败，跳过当前条目')
                    fail_count += 1
                    record_item_summary(seq, item_started, False, "query_stalled_recovery_failed")
                    continue
    
                # 保存快照
                snap_dir = run_dir / 'snapshots'
                snap_dir.mkdir(exist_ok=True)
                snap_file = snap_dir / f'query_snapshot_{seq:04d}.json'
                try:
                    import json as _json
                    snap_data = {
                        "seq": seq, "scene": scene, "coord": coord,
                        "ts": datetime.now().isoformat(timespec='seconds'),
                        "log_file": str(log_file),
                    }
                    snap_file.write_text(_json.dumps(snap_data, ensure_ascii=False, indent=2), encoding='utf-8')
                except Exception as e:
                    print(f'  ⚠️ 快照写入失败: {e}')
    
                _append_jsonl(log_file, {
                    "ts": datetime.now().isoformat(timespec='seconds'),
                    "type": "snapshot", "path": str(snap_file)
                })
    
                # 导出
                if export_enabled:
                    try:
                        download_dir, before_download = _download_snapshot()
                        if not _switch_to_report_frame(scene, plan):
                            raise RuntimeError("导出前切换报表iframe失败")
                        _fr_click_export()
                        download_res = _wait_download_delta(before_download, timeout=25.0)
                        if _download_result_ok(download_res):
                            download_res = _rename_download_files(download_res, seq, scene, field_methods_mode)
                        new_files = download_res.get("files", [])
                        if _download_result_ok(download_res):
                            for f in new_files:
                                print(f'  ✓ 下载文件路径: {f.get("path")} ({f.get("size")} bytes)')
                        else:
                            print(f'  ⚠️ 未确认新增下载文件，目录: {download_res.get("download_dir", str(download_dir))}')
                        _append_jsonl(log_file, {
                            "ts": datetime.now().isoformat(timespec='seconds'),
                            "type": "export_done", "seq": seq, "scene": scene,
                            "field_methods_mode": field_methods_mode,
                            "download": download_res,
                        })
                        if not _download_result_ok(download_res):
                            raise RuntimeError(f'download_not_confirmed: timeout={download_res.get("timeout")}, files={len(new_files)}')
                        # 关闭残留菜单后重新切回 iframe（导出菜单可能改变了焦点）
                        try:
                            from selenium.webdriver.common.action_chains import ActionChains
                            ActionChains(_driver).send_keys('\ue00c').perform()
                        except Exception:
                            pass
                        time.sleep(0.3)
                        _switch_to_report_frame(scene, plan)
                    except Exception as e:
                        print(f'  ✗ 导出失败: {e}')
                        _append_jsonl(log_file, {
                            "ts": datetime.now().isoformat(timespec='seconds'),
                            "type": "step_error", "seq": seq, "error": str(e),
                        })
                        consecutive_errors += 1
                        if not use_recovery:
                            fail_count += 1
                            record_item_summary(seq, item_started, False, "export_failed")
                            continue
                        # 三级恢复后重试导出
                        print('  🔄 尝试恢复后重试导出...')
                        if _recover_report(scene, plan):
                            consecutive_errors = 0
                            items_since_refresh = 0
                            try:
                                download_dir2, before_download2 = _download_snapshot()
                                if not _switch_to_report_frame(scene, plan):
                                    raise RuntimeError("导出前切换报表iframe失败")
                                _fr_click_export()
                                download_res2 = _wait_download_delta(before_download2, timeout=25.0)
                                if _download_result_ok(download_res2):
                                    download_res2 = _rename_download_files(download_res2, seq, scene, field_methods_mode)
                                for f2 in (download_res2.get("files") or []):
                                    print(f'  ✓ 下载文件路径: {f2.get("path")} ({f2.get("size")} bytes)')
                                _append_jsonl(log_file, {
                                    "ts": datetime.now().isoformat(timespec='seconds'),
                                    "type": "export_done", "seq": seq, "scene": scene,
                                    "field_methods_mode": field_methods_mode,
                                    "download": download_res2,
                                })
                                if not _download_result_ok(download_res2):
                                    print('  ✗ 重试导出仍未确认新增下载文件')
                                    fail_count += 1
                                    record_item_summary(seq, item_started, False, "export_failed")
                                    continue
                                try:
                                    from selenium.webdriver.common.action_chains import ActionChains
                                    ActionChains(_driver).send_keys('').perform()
                                except Exception:
                                    pass
                                time.sleep(0.3)
                                _switch_to_report_frame(scene, plan)
                            except Exception as e2:
                                print(f'  ✗ 重试导出仍失败: {e2}')
                                fail_count += 1
                                record_item_summary(seq, item_started, False, "export_failed")
                                continue
                        else:
                            print('  ✗ 恢复失败，跳过当前条目')
                            fail_count += 1
                            record_item_summary(seq, item_started, False, "export_failed")
                            continue
    
                ok_count += 1
                consecutive_errors = 0
                items_since_refresh += 1
                record_item_summary(seq, item_started, True, "ok")
    
            except Exception as e:
                fail_count += 1
                if use_recovery:
                    if _handle_matrix_item_exception(e, scene, plan, seq, log_file, record_item_summary, item_started):
                        items_since_refresh = 0
                        consecutive_errors = 0
                else:
                    msg = str(e)
                    print(f'  ✗ 当前条目异常: {msg}')
                    _append_jsonl(log_file, {
                        "ts": datetime.now().isoformat(timespec='seconds'),
                        "type": "step_error",
                        "seq": seq,
                        "scene": scene,
                        "error": msg,
                        "stage": "item_exception_plain",
                    })
                    record_item_summary(seq, item_started, False, "item_exception")
                continue
    except KeyboardInterrupt:
        interrupted = True
        interrupt_ts = datetime.now().isoformat(timespec='seconds')
        _append_jsonl(log_file, {
            "ts": interrupt_ts,
            "type": "run_interrupted",
            "scene": scene,
            "field_methods_mode": field_methods_mode,
            "ok": ok_count,
            "fail": fail_count,
            "total": total,
        })
        print('\n⚠️ 已收到停止指令，当前任务已中断，返回主菜单')

    # 总结
    run_elapsed = round(time.time() - run_started, 3)
    avg_elapsed = round(run_elapsed / total, 3) if total else 0
    avg_success_elapsed = round(run_elapsed / ok_count, 3) if ok_count else 0
    summary = {
        "ts": datetime.now().isoformat(timespec='seconds'),
        "run_started": run_started_iso,
        "run_finished": datetime.now().isoformat(timespec='seconds'),
        "type": "run_summary",
        "scene": scene,
        "field_methods_mode": field_methods_mode,
        "total": total,
        "ok": ok_count,
        "fail": fail_count,
        "elapsed_seconds": run_elapsed,
        "elapsed_minutes": round(run_elapsed / 60, 3),
        "avg_seconds_per_item": avg_elapsed,
        "avg_seconds_per_success": avg_success_elapsed,
        "query_enabled": query_enabled,
        "export_enabled": export_enabled,
        "query_wait_seconds": query_wait,
        "abort_reason": abort_reason,
        "interrupted": interrupted,
        "interrupt_ts": interrupt_ts,
    }
    _append_jsonl(log_file, {
        **summary,
    })
    try:
        (run_dir / 'run_summary.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
    except Exception as e:
        print(f'⚠️ 汇总耗时文件写入失败: {e}')
    status_text = "已中断" if interrupted else "执行完成"
    print(f'\\n✅ {scene} [{field_methods_mode}] {status_text}: 成功 {ok_count}, 失败 {fail_count} / 共 {total}')
    print(f'⏱️ 总耗时: {run_elapsed:.2f}s ({run_elapsed / 60:.2f}分钟) | 平均每条: {avg_elapsed:.2f}s | 平均成功: {avg_success_elapsed:.2f}s')
    print(f'\\U0001f4dd 日志: {run_dir}')


def _safe_filename(text, maxlen=40):

    """将任意文本转为安全的文件名（只保留字母数字和 -_.，其余替换为 _，限制长度）。



    调用方: _iframe_label、extract_current_html、extract_all_tabs_html。

    """

    return ''.join(c if c.isalnum() or c in '-_.' else '_' for c in text)[:maxlen]





def _iframe_label(iframe, fallback):

    """从 iframe 元素提取名称标签（优先级: name > id > src文件名 > fallback）。



    调用方: _extract_frames_recursive。

    返回值: str（经过 _safe_filename 处理的安全文件名片段）。

    """

    try:

        raw = (

            iframe.get_attribute('name') or

            iframe.get_attribute('id') or

            (iframe.get_attribute('src') or '').rstrip('/').split('/')[-1] or

            fallback

        )

        return _safe_filename(raw) or fallback

    except Exception:

        return fallback





def _extract_frames_recursive(parent_label, ts, url_part, saved, depth, max_depth=5):

    """递归抓取当前框架内所有 iframe 的 HTML 源码，最深 max_depth=5 层。



    每抓完一个 iframe 后安全退回到父框架（parent_frame），而非 default_content，

    确保嵌套层级正确。每个 iframe 输出为独立的 .html 文件。



    调用方: extract_current_html、extract_all_tabs_html。

    """

    from selenium.webdriver.common.by import By



    if depth >= max_depth:

        return



    try:

        iframes = _driver.find_elements(By.TAG_NAME, 'iframe')

    except Exception:

        return



    for i, iframe in enumerate(iframes, 1):

        label      = _iframe_label(iframe, f'f{i}')

        full_label = f'{parent_label}__{label}' if parent_label else label

        indent     = '  ' * depth



        try:

            _driver.switch_to.frame(iframe)



            frame_out = OUTPUT / f'{url_part}_{ts}_{full_label}.html'

            frame_out.write_text(_driver.page_source, encoding='utf-8')

            saved.append((f'{indent}iframe [{full_label}]', frame_out))



            # 递归进入子 iframe

            _extract_frames_recursive(full_label, ts, url_part, saved, depth + 1, max_depth)



        except Exception as e:

            print(f'   ⚠️  {indent}iframe [{full_label}] 跳过: {e}')

        finally:

            # 返回上一层（而非 default_content，保持递归层级正确）

            try:

                _driver.switch_to.parent_frame()

            except Exception:

                _driver.switch_to.default_content()





def extract_all_tabs_html():

    """菜单3: 列出所有浏览器标签页，用户选择后提取每个标签页的 HTML（含 iframe 递归）。



    调用方: main 菜单分发。

    """

    if not driver_alive():

        print('❌ 浏览器未启动')

        return



    handles  = _driver.window_handles

    original = _driver.current_window_handle



    # 列出所有标签页

    print(f'\n共 {len(handles)} 个标签页:')

    tab_urls = []

    for i, h in enumerate(handles, 1):

        try:

            _driver.switch_to.window(h)

            url = _driver.current_url

            active = ' ← 当前' if h == original else ''

            print(f'  {i}. {url or "空白标签页"}{active}')

            tab_urls.append((h, url))

        except Exception:

            tab_urls.append((h, ''))



    _driver.switch_to.window(original)



    choice = input('\n提取哪些标签页? (直接 Enter = 全部, 输入序号如 1,3): ').strip()

    if choice:

        try:

            idxs = [int(x) - 1 for x in choice.split(',')]

            selected = [tab_urls[i] for i in idxs if 0 <= i < len(tab_urls)]

        except Exception:

            print('❌ 序号格式错误')

            return

    else:

        selected = tab_urls



    ts = datetime.now().strftime('%Y%m%d_%H%M%S')

    all_saved = []



    for tab_i, (handle, url) in enumerate(selected, 1):

        print(f'\n── 标签页 {tab_i}/{len(selected)}: {url or "空白"} ──')

        try:

            _driver.switch_to.window(handle)

            _driver.switch_to.default_content()



            url_part = _safe_filename(url.rstrip('/').split('/')[-1] or f'tab{tab_i}')

            prefix   = f'{url_part}_{ts}' if len(selected) == 1 else f'tab{tab_i}_{url_part}_{ts}'

            saved    = []



            main_out = OUTPUT / f'{prefix}.html'

            main_out.write_text(_driver.page_source, encoding='utf-8')

            saved.append(('主框架', main_out))



            _extract_frames_recursive('', ts, prefix, saved, depth=0)

            _driver.switch_to.default_content()



            for label, path in saved:

                kb = path.stat().st_size // 1024

                print(f'   ✅ {label:<25} {path.resolve()}  ({kb} KB)')

            all_saved.extend(saved)

        except Exception as e:

            print(f'   ❌ 提取失败: {e}')



    _driver.switch_to.window(original)

    print(f'\n共保存 {len(all_saved)} 个文件，保存目录: {OUTPUT.resolve()}')


def quick_start():
    """一键启动: 启动浏览器 → 登录 → 跳转报表树页面。

    合并 1 + 2 + t2 的流程，自动处理登录过期。
    卡住时报错暂停，方便人工介入。
    """
    # Step 1: 启动浏览器
    if not driver_alive():
        print('📌 启动浏览器...')
        launch_browser()
        if not driver_alive():
            print('❌ 浏览器启动失败，请检查 chromedriver.exe')
            return
    else:
        print('✅ 浏览器已运行')

    # Step 2: 检查是否已登录
    try:
        _driver.switch_to.default_content()
    except Exception:
        pass

    state = _detect_login_state()
    already_logged_in = state.get("logged_in")

    if already_logged_in:
        print('✅ 检测到已登录状态')
        # 检查是否有登录过期弹窗
        if state.get("expired"):
            print('⚠️ 检测到登录过期弹窗，重新登录...')
            if not _ensure_logged_in():
                print('❌ 重新登录失败，请人工介入')
                return
    else:
        # 未登录，执行完整登录流程
        print('📌 未登录，执行自动登录...')
        auto_login()

        # 登录后再次检查
        time.sleep(2)
        state_after_login = _detect_login_state()
        logged_in_now = state_after_login.get("logged_in")
        if not logged_in_now:
            print('⚠️ 登录后未检测到成功标志')
            # 检查是否在登录页（登录失败）
            on_login_page = state_after_login.get("on_login_page")
            if on_login_page:
                print('❌ 仍在登录页，自动登录可能失败，请人工登录后重试')
                return
            # 检查过期弹窗
            if state_after_login.get("expired"):
                print('⚠️ 登录过期，尝试重新登录...')
                if not _ensure_logged_in():
                    print('❌ 重新登录失败，请人工介入')
                    return

    # Step 3: 导航到报表树页面
    print('📌 导航到报表树页面...')
    if not _navigate_to_report_page():
        print('❌ 导航到报表页失败，请人工检查当前页面状态')
        print('   可能需要手动点击: 报表服务中心 → 数据共享 → 共享数据查看')
        return

    print('\n🎉 一键启动完成！已到达报表树页面，可以执行查询导出。')


def test_soft_refresh():
    """测试: 软刷新 — 重置树 → 重新导航 → 切 iframe → 填日期。

    不刷新整个页面，只点击「共享报表查看」重置树状态，
    然后重新按 tree_path 导航到目标报表。
    """
    if not driver_alive():
        print('❌ 浏览器未启动，请先执行「1. 启动浏览器」')
        return

    config = load_task_config('org')
    if not config:
        print('❌ 无法加载 org 配置')
        return
    try:
        plan = parse_matrix_config(config, scene='org')
    except ValueError as e:
        print(f'❌ 配置解析失败: {e}')
        return

    tree_path = plan.get('tree_path', '')
    if not tree_path:
        print('❌ org 配置中 task.tree_path 为空')
        return

    date_val = plan.get('fields', {}).get('data_date', {}).get('values', ['2026-03-31'])
    date_str = date_val[0] if isinstance(date_val, list) and date_val else str(date_val)

    print(f'\\n🧪 软刷新测试（不重载页面）')
    print(f'  tree_path: {tree_path}')
    print(f'  data_date: {date_str}')

    # Step 1: 切回主页面，直接树导航（保留已有展开状态）
    print('\\n📌 Step 1: 树导航 + 切 iframe...')
    try:
        _driver.switch_to.default_content()
    except Exception:
        pass

    if not _navigate_tree_to_report(tree_path):
        print('❌ 树导航失败')
        return

    if not _switch_to_report_frame('org', plan):
        print('❌ 切入 iframe 失败')
        return
    _save_iframe_src()
    print('✅ 已切入报表 iframe')

    # Step 2: 填日期验证
    print('\\n📌 Step 2: 填充 data_date 验证...')
    label = REPORT_LABEL_MAP.get('data_date', 'data_date')
    date_norm = _normalize_date_text(date_str)
    current = _fr_get_field_value(label)
    print(f'  当前 data_date: {current}')
    result = _fr_set_field(label, date_norm, field_type='date', select_strategy='sendkey')
    if result and result.get('ok'):
        print(f'  ✅ 日期已填充: {result.get("actual", "")}')
    else:
        print(f'  ❌ 日期填充失败')
        return
    time.sleep(0.5)
    after = _fr_get_field_value(label)
    print(f'  回读 data_date: {after}')
    if after and date_norm in str(after or ''):
        print(f'\\n🎉 软刷新测试通过！')
    else:
        print(f'\\n⚠️ 日期回读不匹配')


def test_hard_refresh():
    """测试: 硬刷新 — F5 整页重载 → 完整恢复 → 填日期。

    模拟页面崩溃后的完整恢复流程:
    F5刷新 → 登录检测 → 门户导航 → 树导航 → 切 iframe → 填日期
    """
    if not driver_alive():
        print('❌ 浏览器未启动，请先执行「1. 启动浏览器」')
        return

    config = load_task_config('org')
    if not config:
        print('❌ 无法加载 org 配置')
        return
    try:
        plan = parse_matrix_config(config, scene='org')
    except ValueError as e:
        print(f'❌ 配置解析失败: {e}')
        return

    tree_path = plan.get('tree_path', '')
    if not tree_path:
        print('❌ org 配置中 task.tree_path 为空')
        return

    date_val = plan.get('fields', {}).get('data_date', {}).get('values', ['2026-03-31'])
    date_str = date_val[0] if isinstance(date_val, list) and date_val else str(date_val)

    print(f'\\n🧪 硬刷新测试（F5 整页重载）')
    print(f'  tree_path: {tree_path}')
    print(f'  data_date: {date_str}')

    # Step 1: F5 整页刷新 + 完整恢复
    print('\\n📌 Step 1: F5 整页重载 + 完整恢复流程...')
    if not _refresh_and_reestablish('org', plan):
        print('❌ 硬刷新恢复失败')
        return
    print('✅ 硬刷新恢复成功，已切入报表 iframe')

    # Step 2: 填日期验证
    print('\\n📌 Step 2: 填充 data_date 验证...')
    label = REPORT_LABEL_MAP.get('data_date', 'data_date')
    date_norm = _normalize_date_text(date_str)
    current = _fr_get_field_value(label)
    print(f'  当前 data_date: {current}')
    result = _fr_set_field(label, date_norm, field_type='date', select_strategy='sendkey')
    if result and result.get('ok'):
        print(f'  ✅ 日期已填充: {result.get("actual", "")}')
    else:
        print(f'  ❌ 日期填充失败')
        return
    time.sleep(0.5)
    after = _fr_get_field_value(label)
    print(f'  回读 data_date: {after}')
    if after and date_norm in str(after or ''):
        print(f'\\n🎉 硬刷新测试通过！')
    else:
        print(f'\\n⚠️ 日期回读不匹配')


def test_portal_navigate():
    """测试: 从门户首页导航到报表树页面。

    模拟登录后的页面跳转流程:
    门户首页 → 报表服务中心 → 数据共享 → 共享数据查看
    """
    if not driver_alive():
        print('❌ 浏览器未启动，请先执行「1. 启动浏览器」')
        return

    print('\\n🧪 测试: 门户首页 → 报表服务中心 → 数据共享 → 共享数据查看')
    print('  注意: 请确保当前已在门户首页（登录后的 landing page）')

    if _navigate_to_report_page():
        print('\\n🎉 测试通过！已成功导航到报表树页面。')
    else:
        print('\\n❌ 测试失败，导航过程中出现问题。')


def extract_current_html():

    """菜单2: 提取当前浏览器标签页的完整 HTML（含所有嵌套 iframe 递归抓取）。



    输出: output/<url片段>_<时间戳>.html 及各级 iframe 的 HTML 文件。



    调用方: main 菜单分发。

    """

    if not driver_alive():

        print('❌ 浏览器未启动，请先执行「1. 启动浏览器」')

        return



    try:

        url = _driver.current_url

    except Exception as e:

        print(f'❌ 获取页面失败: {e}')

        return



    ts       = datetime.now().strftime('%Y%m%d_%H%M%S')

    url_part = _safe_filename(url.rstrip('/').split('/')[-1] or 'page')

    saved    = []



    _driver.switch_to.default_content()



    # 主框架

    try:

        main_out = OUTPUT / f'{url_part}_{ts}.html'

        main_out.write_text(_driver.page_source, encoding='utf-8')

        saved.append(('主框架', main_out))

    except Exception as e:

        print(f'❌ 主框架提取失败: {e}')

        return



    # 递归抓所有 iframe

    _extract_frames_recursive('', ts, url_part, saved, depth=0)



    # 确保最终回到主框架

    _driver.switch_to.default_content()



    print(f'\n✅ 共保存 {len(saved)} 个文件  |  保存目录: {OUTPUT.resolve()}  |  来源: {url}')

    for label, path in saved:

        kb = path.stat().st_size // 1024

        print(f'   {label:<30} {path.resolve()}  ({kb} KB)')





# ═══════════════════════════════════════════════════

#  7. 环境自检

# ═══════════════════════════════════════════════════



def environment_check():

    """菜单9: 环境自检——检查工作目录、output可写性、chromedriver、依赖、config.xlsx 各 sheet。



    生成 output/env_check_<时间戳>.json 诊断报告。



    调用方: main 菜单分发。

    """

    checks = []



    def add(name, ok, detail=''):

        checks.append({"name": name, "ok": bool(ok), "detail": detail})

        status = '✓' if ok else '✗'

        print(f'   {status} {name}: {detail}')



    print('\n开始环境自检...')

    add('工作目录', BASE_DIR.exists(), str(BASE_DIR))



    try:

        OUTPUT.mkdir(exist_ok=True)

        probe = OUTPUT / f'.write_test_{datetime.now().strftime("%H%M%S")}.tmp'

        probe.write_text('ok', encoding='utf-8')

        probe.unlink(missing_ok=True)

        add('output目录可写', True, str(OUTPUT))

    except Exception as e:

        add('output目录可写', False, str(e))



    if os.name == 'nt':
        driver_exe = BASE_DIR / 'chromedriver.exe'
        add('chromedriver.exe', driver_exe.exists(), str(driver_exe) if driver_exe.exists() else '未找到')
    else:
        driver_path = _find_linux_chromedriver()
        if driver_path:
            add('ChromeDriver', Path(driver_path).is_file(), driver_path)
        else:
            add('ChromeDriver', True, '未找到本地驱动；启动时由 Selenium Manager 自动匹配（需要网络）')
        browser_path = _find_linux_chrome_binary()
        add('Chrome/Chromium', bool(browser_path), browser_path or '未找到，请安装浏览器或设置 CHROMEAUTORUN_CHROME_BINARY')



    try:

        import selenium  # noqa: F401

        add('selenium依赖', True, '已安装')

    except Exception as e:

        add('selenium依赖', False, str(e))



    try:

        from openpyxl import load_workbook

        add('openpyxl依赖', True, '已安装')

    except Exception as e:

        add('openpyxl依赖', False, str(e))

        load_workbook = None



    if not EXCEL_CONFIG_FILE.exists():

        add('config.xlsx', False, '未找到')

    elif load_workbook is None:

        add('config.xlsx', False, 'openpyxl不可用，无法读取')

    else:

        try:

            wb = load_workbook(EXCEL_CONFIG_FILE, data_only=True)

            required_sheets = ['login', 'county', 'org', 'county2', 'standalone_fill']

            missing_sheets = [s for s in required_sheets if s not in wb.sheetnames]

            add('config.xlsx sheet', not missing_sheets, '缺少: ' + ', '.join(missing_sheets) if missing_sheets else ', '.join(required_sheets))



            if 'fr_widget_map' in wb.sheetnames:

                wm = _read_table_sheet(wb['fr_widget_map'])

                ok_wm = all((r.get('label') or '').strip() and (r.get('input_widget') or '').strip() for r in wm)

                add('fr_widget_map', ok_wm, f'{len(wm)} 条映射' + ('' if ok_wm else ' (存在空 label 或空 input_widget)'))

            else:

                add('fr_widget_map', True, '未配置（将使用 DOM 兜底定位）')



            if 'login' in wb.sheetnames:

                login_cfg = _read_kv_sheet(wb['login'])

                missing_login = [k for k in ('url', 'credit_code', 'username', 'password') if not login_cfg.get(k)]

                add('login配置', not missing_login, '缺少值: ' + ', '.join(missing_login) if missing_login else '完整')



            for scene in ('county', 'org', 'county2'):

                if scene in wb.sheetnames:

                    row_map = _read_kv_sheet(wb[scene])

                    try:

                        parse_matrix_config(row_map, scene=scene)

                        add(f'{scene}配置解析', True, '通过')

                    except Exception as e:

                        add(f'{scene}配置解析', False, str(e))

            if 'standalone_fill' in wb.sheetnames:

                try:

                    cfg = load_standalone_config()

                    add('standalone_fill配置', cfg is not None, f'{len(cfg.get("rules",[]))} 个字段' if cfg else '解析失败')

                except Exception as e:

                    add('standalone_fill配置', False, str(e))

        except Exception as e:

            add('config.xlsx读取', False, str(e))



    if driver_alive():

        try:

            caps = _driver.capabilities or {}

            browser = f"{caps.get('browserName', 'chrome')} {caps.get('browserVersion', '')}".strip()

            add('当前浏览器连接', True, f'{browser} | {_driver.current_url}')

        except Exception as e:

            add('当前浏览器连接', False, str(e))

    else:

        add('当前浏览器连接', True, '未启动，跳过连接检查；如需验证Chrome/driver匹配，请先执行菜单1')



    report = {

        "ts": datetime.now().isoformat(timespec='seconds'),

        "base_dir": str(BASE_DIR),

        "checks": checks,

        "ok": all(c["ok"] for c in checks),

    }

    path = OUTPUT / f'env_check_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json'

    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')

    print(f'\n自检结果: {"通过" if report["ok"] else "存在问题"}')

    print(f'报告文件: {path}')





# ═══════════════════════════════════════════════════

def browser_smoke_test():
    """菜单12: 验证 Chrome/ChromeDriver 启动、本地按钮交互或指定网页读取。"""
    print('\n浏览器冒烟测试：')
    print('  1. 本地按钮交互（不联网）')
    print('  2. 网页只读检查（默认百度）')
    print('  0. 返回')
    try:
        mode = input('请选择: ').strip()
    except (KeyboardInterrupt, EOFError):
        print('\n已取消浏览器冒烟测试')
        return

    if mode == '0':
        return
    if mode not in ('1', '2'):
        print('❌ 无效输入')
        return

    test_url = None
    host_label = None
    if mode == '2':
        from urllib.parse import urlsplit

        try:
            entered_url = input('请输入网址（回车使用 https://www.baidu.com）: ').strip()
        except (KeyboardInterrupt, EOFError):
            print('\n已取消网页检查')
            return
        test_url = entered_url or 'https://www.baidu.com'
        if '://' not in test_url:
            test_url = 'https://' + test_url
        try:
            parsed_url = urlsplit(test_url)
            host = parsed_url.hostname
            if parsed_url.scheme.lower() not in ('http', 'https') or not host:
                raise ValueError('仅支持有效的 http:// 或 https:// 网址')
            host_label = host
            if parsed_url.port:
                host_label = f'{host}:{parsed_url.port}'
        except ValueError as e:
            print(f'❌ 网址无效: {e}')
            return

    browser_was_running = driver_alive()
    if not browser_was_running:
        print('ℹ️ 浏览器尚未启动，正在启动 Chrome...')
        launch_browser()
    if not driver_alive():
        print('❌ Chrome/ChromeDriver 未能启动，冒烟测试无法继续')
        return

    try:
        original_handle = _driver.current_window_handle
    except Exception as e:
        print(f'❌ 无法读取当前浏览器标签页: {type(e).__name__}')
        return

    created_test_tab = False
    test_handle = original_handle
    if browser_was_running:
        try:
            _driver.switch_to.new_window('tab')
            created_test_tab = True
            test_handle = _driver.current_window_handle
        except Exception as e:
            print(f'❌ 无法创建独立测试标签页: {type(e).__name__}')
            return

    try:
        if mode == '1':
            from urllib.parse import quote
            from selenium.webdriver.common.by import By
            from selenium.webdriver.support import expected_conditions as EC
            from selenium.webdriver.support.ui import WebDriverWait

            test_page = '''<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><title>ChromeAutoRun 本地冒烟测试</title></head>
<body><button id="smoke-test-button" type="button"
onclick="document.getElementById('smoke-test-result').textContent='BUTTON_OK'">测试按钮</button>
<span id="smoke-test-result">NOT_CLICKED</span></body></html>'''
            data_url = 'data:text/html;charset=utf-8,' + quote(test_page, safe='')
            _driver.get(data_url)
            button = WebDriverWait(_driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'smoke-test-button'))
            )
            button.click()
            WebDriverWait(_driver, 5).until(
                lambda driver: driver.find_element(By.ID, 'smoke-test-result').text == 'BUTTON_OK'
            )
            print('✅ 本地按钮点击成功：Chrome、ChromeDriver、Selenium 点击和页面脚本均可用')
        else:
            from selenium.common.exceptions import TimeoutException
            from selenium.webdriver.support.ui import WebDriverWait

            previous_page_load_timeout = None
            try:
                previous_page_load_timeout = _driver.timeouts.page_load
            except Exception:
                pass
            try:
                _driver.set_page_load_timeout(30)
                try:
                    _driver.get(test_url)
                except TimeoutException:
                    print('⚠️ 页面加载超过 30 秒，继续检查已加载内容')

                try:
                    WebDriverWait(_driver, 10).until(
                        lambda driver: driver.execute_script('return document.readyState')
                        in ('interactive', 'complete')
                    )
                except TimeoutException:
                    pass

                page_info = _driver.execute_script('''
return {
    title: document.title || '',
    readyState: document.readyState || '',
    chromeError: (document.querySelector('#main-frame-error .error-code, .error-code') || {}).textContent || '',
    interactiveElements: document.querySelectorAll(
        'button,input,select,textarea,a[href],[role="button"]'
    ).length
};''')
                title = (page_info.get('title') or '').replace('\n', ' ').strip()
                if len(title) > 100:
                    title = title[:97] + '...'
                ready_state = page_info.get('readyState', '')
                chrome_error = (page_info.get('chromeError') or '').strip()
                element_count = page_info.get('interactiveElements', 0)
                if chrome_error:
                    print(f'❌ 网页加载失败: {host_label}（{chrome_error}）')
                elif ready_state in ('interactive', 'complete'):
                    print(f'✅ 网页已打开: {host_label}')
                else:
                    print(f'⚠️ 已连接到 {host_label}，但页面尚未就绪（{ready_state or "未知状态"}）')
                print(f'   页面标题: {title or "（空）"}')
                print(f'   可交互元素数: {element_count}')
                print('   本项只读取页面状态和元素数量，不会点击网页控件或提交表单。')
            finally:
                if previous_page_load_timeout is not None:
                    try:
                        _driver.set_page_load_timeout(previous_page_load_timeout)
                    except Exception:
                        pass
    except Exception as e:
        print(f'❌ 冒烟测试失败: {type(e).__name__}')
        if mode == '2':
            print(f'   检查地址、网络连通性、证书，以及 Chrome 与 ChromeDriver 版本是否匹配（目标: {host_label}）。')
        else:
            print('   请检查 Chrome 与 ChromeDriver 版本是否匹配。')
    finally:
        if created_test_tab:
            try:
                if _driver is not None and test_handle in _driver.window_handles:
                    _driver.switch_to.window(test_handle)
                    _driver.close()
            except Exception:
                print('⚠️ 测试标签页未能自动关闭')
            try:
                if _driver is not None and original_handle in _driver.window_handles:
                    _driver.switch_to.window(original_handle)
            except Exception:
                print('⚠️ 测试结束后未能恢复原标签页，请检查浏览器窗口')


#  主菜单

# ═══════════════════════════════════════════════════



def browser_status_line():

    """生成菜单顶部显示的浏览器状态行字符串。



    调用方: print_menu。

    返回值: str（如 "✅ [3个标签] https://example.com/..." | "⭕ 未启动" | "⭕ 连接丢失"）。

    """

    if not driver_alive():

        return '⭕ 未启动'

    try:

        handles = _driver.window_handles

        url     = _driver.current_url

        tab_info = f'[{len(handles)}个标签]' if len(handles) > 1 else ''

        display  = '空白标签页' if url in ('data:,', 'about:blank', '') else url

        short    = display if len(display) <= 38 else display[:35] + '...'

        return f'✅ {tab_info} {short}'.strip()

    except Exception:

        return '⭕ 连接丢失'





def print_menu():

    """打印主菜单界面（含浏览器状态行和实验菜单项）。



    调用方: main 主循环。

    """

    bstat = browser_status_line()

    print()

    print('╔══════════════════════════════════════════════╗')

    print('║          网页自动化工具  v1.0                ║')

    print(f'║  浏览器: {bstat:<36}║')

    print('╠══════════════════════════════════════════════╣')

    print('║  准备阶段                                    ║')
    print('║  1.  启动浏览器                              ║')
    print('║  2.  导航至金数系统并登陆                    ║')
    print('║  2.5 一键启动+登陆+跳转报表页                ║')
    print('║                                              ║')
    print('║  执行阶段                                    ║')
    print('║  3.  地区汇总报表查询导出                    ║')
    print('║  4.  单家机构报表查询导出                    ║')
    print('║  5.  注册当地法人机构报表查询导出            ║')
    print('║  6.  报表组批量修改                          ║')
    print('║                                              ║')
    print('║  诊断阶段                                    ║')
    print('║  7.  提取当前页面 HTML                       ║')
    print('║  8.  提取所有标签页 HTML                     ║')
    print('║  9.  环境自检                                ║')
    print('║  12. 浏览器冒烟测试                          ║')
    print('║                                              ║')
    print('║  维护阶段                                    ║')
    print('║  10. 重置配置模板                            ║')
    print('║  11. 打开 config.xlsx                        ║')

    print('║  r.  刷新状态                                ║')

    print('║  0.  退出                                    ║')

    print('╚══════════════════════════════════════════════╝')





def main():

    """程序入口: 同步配置 → 进入主菜单循环 → 分发用户选择到各功能函数。



    菜单项对应关系:

    1→launch_browser, 2→auto_login,
    3→run_matrix_export('county', recovery_mode='recover'),
    4→run_matrix_export('org', recovery_mode='recover'),
    5→run_matrix_export('county2', recovery_mode='recover'),
    6→run_standalone_fill。
    7→extract_current_html, 8→extract_all_tabs_html, 9→environment_check,
    12→browser_smoke_test,
    10→reset_config_template, 11→open_config_excel, r→刷新, 0→退出。

    """

    cfg = sync_config_json_from_excel()

    if cfg:

        missing = [k for k in ('url', 'credit_code', 'username', 'password') if not cfg.get(k)]

        if missing:

            print(f'⚠️  config.xlsx 缺少字段值: {", ".join(missing)}')



    while True:

        print_menu()

        try:

            choice = input('请选择: ').strip()

        except (KeyboardInterrupt, EOFError):

            choice = '0'



        print()

        if choice == '1':

            launch_browser()

        elif choice == '2':

            auto_login()

        elif choice == '2.5':

            quick_start()

        elif choice == '3':

            run_matrix_export('county', recovery_mode='recover')

        elif choice == '4':

            run_matrix_export('org', recovery_mode='recover')

        elif choice == '5':

            run_matrix_export('county2', recovery_mode='recover')

        elif choice == '6':

            run_standalone_fill()

        elif choice == '7':

            extract_current_html()

        elif choice == '8':

            extract_all_tabs_html()

        elif choice == '9':

            environment_check()

        elif choice == '10':

            reset_config_template()

        elif choice == '11':

            open_config_excel()

        elif choice == '12':

            browser_smoke_test()

        elif choice in ('3p', '3P'):

            run_matrix_export('county', recovery_mode='plain')

        elif choice in ('4p', '4P'):

            run_matrix_export('org', recovery_mode='plain')

        elif choice in ('5p', '5P'):

            run_matrix_export('county2', recovery_mode='plain')

        elif choice in ('r', 'R'):

            continue

        elif choice == '0':

            if driver_alive():

                try:

                    _driver.quit()

                except Exception:

                    pass

            print('👋 已退出')

            break

        else:

            print('❌ 无效输入')

            continue





if __name__ == '__main__':

    main()

