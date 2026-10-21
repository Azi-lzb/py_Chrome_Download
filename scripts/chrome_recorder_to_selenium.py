#!/usr/bin/env python3
"""
Chrome DevTools Recorder JSON  →  Python Selenium 脚本

用法:
    python chrome_recorder_to_selenium.py 录制文件.json
    python chrome_recorder_to_selenium.py 录制文件.json -o 输出.py

输出的脚本依赖:
    pip install selenium
    还需要 ChromeDriver: https://chromedriver.chromium.org/
"""

import json
import sys
import argparse
from pathlib import Path


# ─── 选择器转换 ─────────────────────────────────────────────────────────────

def parse_selectors(raw):
    """
    Chrome Recorder selectors 格式:
        [ ["css..."], ["xpath//..."], ["pierce/..."], ["text/..."] ]
    每个内层 list 是一组等价选择器（通常只有一条），转成 (By, value) 元组列表
    """
    result = []
    seen = set()
    for group in raw:
        for sel in group:
            if not sel:
                continue
            if sel.startswith('xpath/'):
                # xpath//... → 去掉 'xpath/' 前缀，保留 //...
                item = ('XPATH', sel[6:])
            elif sel.startswith('text/'):
                text = sel[5:].replace('"', '\\"')
                item = ('XPATH', f'//*[normalize-space(.)="{text}"]')
            elif sel.startswith('pierce/'):
                # Shadow DOM 穿透，Selenium 不支持，降级为普通 CSS
                item = ('CSS_SELECTOR', sel[7:])
            elif sel.startswith('aria/'):
                label = sel[5:].replace('"', '\\"')
                item = ('XPATH', f'//*[@aria-label="{label}"]')
            else:
                item = ('CSS_SELECTOR', sel)
            if item not in seen:
                seen.add(item)
                result.append(item)
    return result


# ─── 步骤转换 ────────────────────────────────────────────────────────────────

def step_comment(step, idx):
    t = step.get('type', '')
    extra = ''
    if t == 'navigate':
        extra = f' → {step.get("url", "")}'
    elif t in ('click', 'doubleClick', 'hover', 'change'):
        sels = step.get('selectors', [])
        if sels and sels[0]:
            extra = f'  [{sels[0][0]}]'
    return f'# 步骤 {idx}: {t}{extra}'


def convert_step(step, pad='    '):
    t = step.get('type', '')
    raw = step.get('selectors', [])
    sels = parse_selectors(raw)
    sels_repr = repr(sels)
    lines = []

    if t == 'setViewport':
        w, h = step.get('width', 1280), step.get('height', 720)
        lines.append(f'{pad}driver.set_window_size({w}, {h})')

    elif t == 'navigate':
        url = step.get('url', '')
        lines.append(f'{pad}driver.get({repr(url)})')
        lines.append(f'{pad}wait_load(driver)')

    elif t == 'click':
        lines.append(f'{pad}find_el(driver, {sels_repr}).click()')
        lines.append(f'{pad}time.sleep(0.5)')

    elif t == 'doubleClick':
        lines.append(f'{pad}ActionChains(driver).double_click(find_el(driver, {sels_repr})).perform()')
        lines.append(f'{pad}time.sleep(0.5)')

    elif t == 'hover':
        lines.append(f'{pad}ActionChains(driver).move_to_element(find_el(driver, {sels_repr})).perform()')
        lines.append(f'{pad}time.sleep(0.3)')

    elif t == 'change':
        value = step.get('value', '')
        lines.append(f'{pad}el = find_el(driver, {sels_repr})')
        lines.append(f'{pad}if el.tag_name.lower() == "select":')
        lines.append(f'{pad}    Select(el).select_by_visible_text({repr(value)})')
        lines.append(f'{pad}else:')
        lines.append(f'{pad}    el.clear()')
        lines.append(f'{pad}    el.send_keys({repr(value)})')
        lines.append(f'{pad}time.sleep(0.3)')

    elif t == 'keyDown':
        key = step.get('key', '')
        lines.append(f'{pad}ActionChains(driver).key_down(Keys.{key.upper()}).perform()')

    elif t == 'keyUp':
        key = step.get('key', '')
        lines.append(f'{pad}ActionChains(driver).key_up(Keys.{key.upper()}).perform()')

    elif t == 'scroll':
        dx = step.get('distanceX', 0)
        dy = step.get('distanceY', 0)
        if sels:
            lines.append(f'{pad}driver.execute_script("arguments[0].scrollBy({dx}, {dy})", find_el(driver, {sels_repr}))')
        else:
            lines.append(f'{pad}driver.execute_script("window.scrollBy({dx}, {dy})")')

    elif t == 'waitForElement':
        ms = step.get('timeout', 5000)
        sec = max(1, ms // 1000)
        if sels:
            by, val = sels[0]
            lines.append(f'{pad}WebDriverWait(driver, {sec}).until(EC.presence_of_element_located((By.{by}, {repr(val)})))')
        else:
            lines.append(f'{pad}time.sleep({sec})')

    elif t == 'waitForExpression':
        expr = step.get('expression', '')
        ms = step.get('timeout', 5000)
        sec = max(1, ms // 1000)
        lines.append(f'{pad}WebDriverWait(driver, {sec}).until(lambda d: d.execute_script("return Boolean({expr})"))')

    elif t in ('emulateNetworkConditions', 'close', 'setCookies'):
        lines.append(f'{pad}# [{t} 已跳过]')

    else:
        lines.append(f'{pad}# [未处理的步骤类型: {t}]  原始: {json.dumps(step, ensure_ascii=False)}')

    return lines


# ─── 辅助函数模板 ───────────────────────────────────────────────────────────

HELPER = '''\
def find_el(driver, selectors, timeout=10):
    """依次尝试多个 (By, value) 选择器，返回第一个找到的元素"""
    errors = []
    for by, value in selectors:
        try:
            by_obj = getattr(By, by)
            el = WebDriverWait(driver, min(timeout, 3)).until(
                EC.presence_of_element_located((by_obj, value)))
            return el
        except Exception as e:
            errors.append(f"  {by}: {value}  →  {e}")
            continue
    raise Exception("找不到元素，已尝试:\\n" + "\\n".join(errors))


def wait_load(driver, timeout=20):
    """等待页面加载完成"""
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete")
    time.sleep(0.5)
'''


# ─── 主转换逻辑 ──────────────────────────────────────────────────────────────

def convert(data: dict) -> str:
    title = data.get('title', '录制的流程')
    steps = data.get('steps', [])

    lines = []
    lines.append(f'# 流程: {title}')
    lines.append(f'# 由 Chrome DevTools Recorder JSON 转换')
    lines.append(f'# 依赖: pip install selenium  +  ChromeDriver')
    lines.append('')
    lines.append('from selenium import webdriver')
    lines.append('from selenium.webdriver.common.by import By')
    lines.append('from selenium.webdriver.support.ui import WebDriverWait, Select')
    lines.append('from selenium.webdriver.support import expected_conditions as EC')
    lines.append('from selenium.webdriver.common.action_chains import ActionChains')
    lines.append('from selenium.webdriver.common.keys import Keys')
    lines.append('import time')
    lines.append('')
    lines.append('')
    lines.append(HELPER)
    lines.append('driver = webdriver.Chrome()')
    lines.append('')
    lines.append('try:')

    for i, step in enumerate(steps, 1):
        lines.append(f'    {step_comment(step, i)}')
        step_lines = convert_step(step)
        lines.extend(step_lines)
        lines.append('')

    lines.append('    print("✅ 执行完成")')
    lines.append('')
    lines.append('except Exception as e:')
    lines.append('    print(f"❌ 出错: {e}")')
    lines.append('    import traceback; traceback.print_exc()')
    lines.append('')
    lines.append('finally:')
    lines.append('    input("按 Enter 键关闭浏览器...")')
    lines.append('    driver.quit()')

    return '\n'.join(lines)


# ─── 入口 ───────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Chrome Recorder JSON → Python Selenium')
    parser.add_argument('input', help='Chrome Recorder 导出的 JSON 文件路径')
    parser.add_argument('-o', '--output', help='输出 Python 文件路径（默认同名 .py）')
    args = parser.parse_args()

    src = Path(args.input)
    if not src.exists():
        print(f'❌ 文件不存在: {src}')
        sys.exit(1)

    data = json.loads(src.read_text(encoding='utf-8'))
    code = convert(data)

    out = Path(args.output) if args.output else src.with_suffix('.py')
    out.write_text(code, encoding='utf-8')
    print(f'✅ 已生成: {out}')
    print(f'   步骤数: {len(data.get("steps", []))}')


if __name__ == '__main__':
    main()
