import sys
import os
import winreg  # [改动] 引入 winreg 模块，用于读取 Windows 注册表


# 辅助函数：弹出黄色 CMD 错误窗口（仅 Windows）

def _popup_error(lines: list[str]):
    body = " & ".join(
        f'echo {ln}' if ln.strip() else "echo."
        for ln in lines
    )
    os.system(f'start cmd /k "color 4E & {body} & echo. & pause"')



# 1. 检测 playwright / openpyxl 是否可导入
#    （打包后这两个库已嵌入 exe，正常不会触发）

try:
    from playwright.sync_api import sync_playwright
    import openpyxl
except ImportError as _err:
    if sys.platform == "win32":
        _popup_error([
            "==============================================",
            " [致命错误] 程序核心组件加载失败！",
            "----------------------------------------------",
            f" 缺失模块: {_err}",
            "",
            " 可能原因：您运行的不是完整的 exe 文件，",
            " 或 exe 文件已损坏。",
            "",
            " 解决方法：",
            "   请重新下载完整的 CNKIBug.exe 文件，",
            "   不要解压、不要移动内部文件，直接双击运行。",
            "==============================================",
        ])
    else:
        print(f"[FATAL] 缺少依赖: {_err}")
        print("请运行: pip install playwright openpyxl && playwright install chromium")
    sys.exit(1)

import time
import random

_EDGE_PATHS = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe"),
]

def _edge_installed() -> bool:
    return any(os.path.isfile(p) for p in _EDGE_PATHS)



def get_real_desktop_path():
    if sys.platform != "win32":
        return os.path.join(os.path.expanduser("~"), "Desktop")
        
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
        )
        val, _ = winreg.QueryValueEx(key, "Desktop")
        winreg.CloseKey(key)

        return os.path.expandvars(val)
    except Exception:
        return os.path.join(os.path.expanduser("~"), "Desktop")


def check_env():
    if sys.platform != "win32":
       
        playwright_path = os.path.join(
            os.path.expanduser("~"), "AppData", "Local", "ms-playwright"
        )
        if not os.path.exists(playwright_path):
            print("\n[环境缺失] 请先在终端运行: playwright install chromium\n")
            sys.exit(0)
        return

    if not _edge_installed():
        _popup_error([
            "==============================================",
            " [环境缺失] 未检测到 Microsoft Edge 浏览器！",
            "----------------------------------------------",
            " 本程序需要使用 Microsoft Edge 来抓取网页数据。",
            " Windows 10/11 通常已预装，若您已卸载请重新安装。",
            "",
            " 请用浏览器打开以下地址，下载并安装 Edge：",
            "",
            "   https://www.microsoft.com/zh-cn/edge/download",
            "",
            " 安装完成后，关闭此窗口，重新双击程序即可！",
            "==============================================",
        ])
        sys.exit(0)


def scrape_cnki(keyword: str, max_pages: int = 3):
    results = []

    with sync_playwright() as p:
        browser = None

        # 优先尝试系统 Edge，再回退普通 Chromium+
        try:
            browser = p.chromium.launch(channel="msedge", headless=False)
            print("[*] 已启动 Microsoft Edge 浏览器")
        except Exception as _e1:
            print(f"[!] Edge 启动失败 ({_e1})，尝试备用 Chromium...")
            try:
                browser = p.chromium.launch(headless=False)
                print("[*] 已启动备用 Chromium 浏览器")
            except Exception as _e2:
                # 两种方式均失败 → 弹窗提示
                if sys.platform == "win32":
                    _popup_error([
                        "==============================================",
                        " [错误] 浏览器启动失败！",
                        "----------------------------------------------",
                        " 程序找到了 Edge，但无法正常启动它。",
                        "",
                        " 可能原因：",
                        "   · Edge 浏览器文件损坏",
                        "   · 系统权限不足",
                        "   · 安全软件阻止了浏览器启动",
                        "",
                        " 建议：",
                        "   1. 重新安装 Microsoft Edge",
                        "      https://www.microsoft.com/zh-cn/edge/download",
                        "   2. 以管理员身份运行本程序",
                        "   3. 暂时关闭杀毒软件后重试",
                        "==============================================",
                    ])
                else:
                    print(f"[FATAL] 浏览器启动失败: {_e2}")
                #用 raise抛出异常代替sys.exit(1)，防止主窗口直接闪退
                raise RuntimeError(f"浏览器启动彻底失败: {_e2}")

        try:
            context = browser.new_context(
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/120.0.0.0 Safari/537.36"
                )
            )
            page = context.new_page()

            print(f"[*] 目标关键词：{keyword}")
            page.goto("https://kns.cnki.net/kns8s/")
            time.sleep(random.uniform(2, 4))

            page.fill("input.search-input", keyword)
            time.sleep(random.uniform(0.5, 1.5))
            page.click("input.search-btn")

            for current_page in range(1, max_pages + 1):
                print(f"[*] 读取第 {current_page} 页...")
                page.wait_for_selector(
                    "table.result-table-list tbody tr", timeout=15000
                )
                time.sleep(random.uniform(2, 5))

                rows = page.query_selector_all("table.result-table-list tbody tr")
                for row in rows:
                    title_el = row.query_selector("td.name a")
                    if title_el:
                        title = title_el.inner_text().strip()
                        results.append([title])
                        print(f"  -> 抓取到: {title}")

                if current_page < max_pages:
                    next_btn = page.query_selector("a#PageNext")
                    if next_btn:
                        next_btn.click()
                        time.sleep(random.uniform(4, 8))
                    else:
                        print("[!] 没找到下一页按钮，可能已到最后一页。")
                        break

        except Exception as e:
            print(f"[x] 抓取过程出错（可能遇到验证码或页面结构变化）: {e}")
        finally:
            #增加防空指针判断，防止启动失败时引发 AttributeError
            if browser:
                browser.close()

    # 调用注册表读取函数，直达真实桌面，无视 OneDrive 劫持
    try:
        real_desktop = get_real_desktop_path()
        os.makedirs(real_desktop, exist_ok=True)
        filename = os.path.join(real_desktop, f"cnki_titles_{keyword}.xlsx")
    except Exception:
        #终极兜底：保存在软件当前的同级目录
        filename = os.path.join(os.getcwd(), f"cnki_titles_{keyword}.xlsx")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "论文标题"
    ws.append(["论文标题"])
    for row in results:
        ws.append(row)
    wb.save(filename)

    full_path = os.path.abspath(filename)
    print("\n" + "=" * 50)
    print(f"[*] 共抓取 {len(results)} 条数据。")
    print(f"[*] 文件已保存至：")
    print(f"    >>> {full_path} <<<")
    print("=" * 50 + "\n")

if __name__ == "__main__":
    try:
        if sys.platform == "win32":
            os.system("cls")

        print("=" * 50)
        print("  CNKI_Bug_dev  |  copyright by Kaffu_Alcaid")
        print("  Version 0.0.5")
        print("=" * 50)
        print("  本软件用于抓取中国知网的论文标题\n")
        print("按Ctrl+C 退出程序")

        check_env()

        search_word = input("请输入你要搜索的关键词: ").strip()
        if not search_word:
            print("[!] 关键词不能为空，程序退出。")
            sys.exit(0)

        user_input_pages = input("请输入想抓取的总页数（纯数字，值不要太大）: ").strip()
        target_pages = int(user_input_pages)
        if target_pages <= 0:
            raise ValueError("页数必须大于 0")

        scrape_cnki(search_word, max_pages=target_pages)

    except ValueError:
        print("\n" + "!" * 40)
        print("  错误：页数请输入【纯数字】，例如 3 或 10")
        print("!" * 40)
    except KeyboardInterrupt:
        print("\n[*] 用户中断，程序退出。")
    except Exception as e:
        print("\n" + "!" * 40)
        print(f"  程序遇到未知错误: {e}")
        print("!" * 40)
    finally:
        input("\n按 [回车键 Enter] 退出程序...")
