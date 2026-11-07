import DrissionPage
from DrissionPage import ChromiumOptions
import ddddocr
import re
import sys
import json
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from datetime import datetime
import os


class SecondClassroomAutomator:
    def __init__(self, web_url, user_id, user_pwd, incognito=False, save_path=""):
        """初始化配置和基础对象，新增 incognito 参数"""
        self.web_url = web_url
        self.user_id = user_id
        self.user_pwd = user_pwd
        # 【保存位置】留空 = 当前文件夹（默认文件名不变）；填文件夹 = 存到该文件夹；填 .xlsx 完整路径 = 存成该文件名
        self.file_name = self._resolve_save_path(save_path)

        self.requirements = {
            "思想成长": 32, "创新创业": 48, "社会实践": 16,
            "志愿公益": 16, "公益志愿": 16, "文体活动": 8,
            "财经素养": 0, "社会参与": 0
        }

        self.ocr = ddddocr.DdddOcr(show_ad=False)

        options = ChromiumOptions()

        # 优先 Chrome，没有 Chrome 就换 Edge
        browser_path = None

        chrome_candidates = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        ]

        edge_candidates = [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            os.path.join(
                os.environ.get("LOCALAPPDATA", ""),
                "Microsoft", "Edge", "Application", "msedge.exe",
            ),
        ]

        for path in chrome_candidates:
            if os.path.exists(path):
                browser_path = path
                print(f"[INFO] 检测到 Chrome，使用：{path}")
                break

        if not browser_path:
            for path in edge_candidates:
                if path and os.path.exists(path):
                    browser_path = path
                    print(f"[INFO] 未检测到 Chrome，已切换到 Edge：{path}")
                    break

        if browser_path:
            options.set_browser_path(browser_path)
        else:
            print("[WARN] 未找到 Chrome 或 Edge，将使用 DrissionPage 默认浏览器路径。")

        # 无痕模式配置
        if incognito:
            print("[INFO] 已开启浏览器无痕模式（不读取本地 Cookie 缓存）。")
            options.incognito(True)
        else:
            print("[INFO] 使用正常模式（可能会读取本地 Cookie 自动登录）。")

        # 初始化浏览器，传入配置
        self.browser = DrissionPage.Chromium(options)
        self.tab = self.browser.latest_tab
        # ================================================

    @staticmethod
    def _resolve_save_path(save_path):
        """把用户填写的保存位置解析成完整的文件名"""
        target = (save_path or "").strip().strip('"').strip("'")
        target = os.path.expanduser(target)
        if not target:
            return "第二课堂学时统计.xlsx"
        if target.lower().endswith(".xlsx"):
            return target
        return os.path.join(target, "第二课堂学时统计.xlsx")

    def _clean_old_file(self):
        if os.path.exists(self.file_name):
            try:
                os.remove(self.file_name)
                print(f"[INFO] 已清理旧的 Excel 文件: {self.file_name}")
            except PermissionError:
                print(f"[ERROR] 旧的 Excel 文件 '{self.file_name}' 正在被打开，无法删除。请先关闭它。")
                sys.exit(1)

    def _check_login_status(self):
        is_logged_in = self.tab.ele('.uni-tabbar__bd', timeout=3)
        return bool(is_logged_in)

    def _do_login(self):
        print("[INFO] 启动自动登录模块（账密模式）...")
        max_retries = 5
        for attempt in range(max_retries):
            print(f"\n--- 第 {attempt + 1} 次尝试登录 ---")
            try:
                login_way = self.tab.ele('xpath://div[contains(text(), "账号密码")]')
                if login_way:
                    login_way.wait.clickable(timeout=3)
                    login_way.click()

                self.tab.ele('#userid').clear().input(self.user_id)
                self.tab.ele('#password').clear().input(self.user_pwd)

                captcha_img = self.tab.ele('@title=点我刷新')
                img_bytes = captcha_img.get_screenshot(as_bytes='png')
                raw_text = self.ocr.classification(img_bytes)

                clean_text = re.sub(r'[xX×]', '*', raw_text)
                clean_text = re.sub(r'÷', '/', clean_text).replace('=', '').replace(' ', '')

                try:
                    answer = eval(clean_text)
                    print(f"[INFO] 识别算式: {clean_text} = {answer}")
                except Exception:
                    print(f"[WARNING] 算式解析失败 '{raw_text}'，刷新重试...")
                    captcha_img.click()
                    self.tab.wait(1)
                    continue

                self.tab.ele('#captchaInput').clear().input(str(answer))
                login_btn = self.tab.ele('xpath://button[normalize-space(.)="登 录" or normalize-space(.)="登录"]')
                if not login_btn:
                    login_btn = self.tab.ele('css:button.ant-btn-primary')

                login_btn.wait.clickable(timeout=5)
                login_btn.click()
                print("[INFO] 已点击登录，正在等待验证结果...")

                self.tab.wait(1.5)
                if not self.tab.ele('#userid', timeout=1):
                    print("[SUCCESS] 账号密码验证通过！")
                    return True
                else:
                    print("[ERROR] 登录失败，刷新验证码重试...")
                    captcha_img.click()
                    self.tab.wait(1)

            except Exception as e:
                print(f"[WARNING] 登录发生意外错误: {e}")
                self.tab.wait(2)
        return False

    def _wait_for_manual_login(self):
        print("\n" + "=" * 50)
        print("[NOTICE] 未检测到账号或密码！")
        print("[NOTICE] 程序已暂停，请在弹出的浏览器中使用微信扫码登录。")
        print("=" * 50)

        while True:
            user_input = input("\n[ACTION] 扫码完成后，请在控制台按【回车键】继续（输入 q 退出程序）：")
            if user_input.strip().lower() == 'q':
                print("[INFO] 用户选择退出程序。")
                sys.exit(0)

            print("[INFO] 正在检测登录状态，请稍候...")
            self.tab.wait(1)

            if self._check_login_status():
                print("[SUCCESS] 检测到登录成功，继续执行后续操作！")
                break
            else:
                print("[WARNING] 似乎还没有登录成功（可能是误触回车，或者二维码未扫描）。")
                print("[WARNING] 请重新扫码，并在完成后再次按回车。")


    def _fetch_data(self):
        print("\n[INFO] 准备进入'我的'页面...")
        my_tab = self.tab.ele('text:我的') or self.tab.eles('.uni-tabbar__item')[-1]
        self.tab.run_js('arguments[0].click();', my_tab)

        print("[INFO] 正在等待'我的'页面加载...")
        self.tab.wait.ele_displayed('text=成绩单', timeout=5)

        score_btn = self.tab.ele('text=成绩单') or self.tab.ele('text:成绩单')
        if not score_btn:
            print("[ERROR] 找不到成绩单按钮，页面可能未完全加载，请重试！")
            sys.exit(1)

        print("[INFO] 开启网络监听，准备获取接口数据...")
        self.tab.listen.start('by-classify-list')

        print("[INFO] 点击'成绩单'按钮...")
        score_btn.click()

        print("[INFO] 等待数据接口响应...")
        res = self.tab.listen.wait(timeout=10)

        if not res:
            print("[ERROR] 未能捕获到成绩单数据接口。")
            sys.exit(1)

        print("[SUCCESS] 成功捕获数据包！开始解析 JSON...")
        json_data = res.response.body
        if isinstance(json_data, str):
            json_data = json.loads(json_data)

        data_list = json_data.get("data", [])
        print(f"[INFO] 成功获取到 {len(data_list)} 条学时数据！")
        return data_list

    def _export_to_excel(self, data_list):
        print("\n[INFO] 开始生成 Excel 表格...")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "第二课堂学时统计"

        color_red_bg = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
        color_red_font = Font(color="C00000", size=14)
        color_green_bg = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
        color_green_font = Font(color="274E13", size=14)
        header_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ws.append([f"数据记录截止时间：{current_time}"])
        ws.merge_cells('A1:E1')
        ws['A1'].font = Font(bold=True, size=12, color="808080")
        ws['A1'].alignment = Alignment(horizontal="center")
        ws.row_dimensions[1].height = 20

        headers = ["学时类型", "毕业需求", "目前学时", "差值", "完成度"]
        ws.append(headers)

        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=2, column=col_num)
            cell.font = Font(bold=True, size=16)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = header_fill
            ws.column_dimensions[get_column_letter(col_num)].width = 18
        ws.row_dimensions[2].height = 30

        for idx, item in enumerate(data_list):
            row_idx = idx + 3
            h_type = item.get("classifyName", "未知类型").strip()
            required_hours = self.requirements.get(h_type, 0)
            current_hours = item.get("classifyHours", 0)
            if current_hours is None:
                current_hours = 0

            ws.cell(row=row_idx, column=1).value = h_type
            ws.cell(row=row_idx, column=2).value = required_hours
            ws.cell(row=row_idx, column=3).value = current_hours

            diff_formula = f'=IF(B{row_idx}-C{row_idx}<=0,"已超额",B{row_idx}-C{row_idx})'
            ws.cell(row=row_idx, column=4).value = diff_formula

            status_formula = f'=IF(B{row_idx}<=C{row_idx},"已完成","未完成")'
            ws.cell(row=row_idx, column=5).value = status_formula

            for col in range(1, 6):
                cell = ws.cell(row=row_idx, column=col)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(size=14)
            ws.row_dimensions[row_idx].height = 25

        max_row = len(data_list) + 2
        ws.conditional_formatting.add(f'B3:C{max_row}', FormulaRule(formula=['$E3="未完成"'], font=color_red_font))
        ws.conditional_formatting.add(f'B3:C{max_row}', FormulaRule(formula=['$E3="已完成"'], font=color_green_font))
        ws.conditional_formatting.add(f'D3:E{max_row}',
                                      FormulaRule(formula=['$E3="未完成"'], fill=color_red_bg, font=color_red_font))
        ws.conditional_formatting.add(f'D3:E{max_row}',
                                      FormulaRule(formula=['$E3="已完成"'], fill=color_green_bg, font=color_green_font))

        ws.freeze_panes = "A3"

        save_dir = os.path.dirname(os.path.abspath(self.file_name))
        if save_dir and not os.path.exists(save_dir):
            os.makedirs(save_dir)
            print(f"[INFO] 保存文件夹不存在，已自动创建: {save_dir}")

        try:
            wb.save(self.file_name)
            print(f"[SUCCESS] 完美！彩色 Excel 表格生成成功！文件名: {os.path.abspath(self.file_name)}")
        except Exception as e:
            print(f"[ERROR] 保存发生未知错误: {e}")

    def run(self):
        try:
            self._clean_old_file()

            print("[INFO] 正在打开网页...")
            self.tab.get(self.web_url)
            self.tab.wait.load_start()

            if not self._check_login_status():
                if self.user_id.strip() and self.user_pwd.strip():
                    if not self._do_login():
                        print("[FATAL] 达到最大重试次数，登录失败，程序退出。")
                        return
                else:
                    self._wait_for_manual_login()
            else:
                print("[INFO] 检测到已登录状态（Cookie 有效），直接跳转至主页。")

            data_list = self._fetch_data()
            self._export_to_excel(data_list)

        finally:
            print("[INFO] 任务结束，关闭浏览器。")
            self.tab.close()


# ================= 启动入口 =================
if __name__ == "__main__":
    # 你的配置信息
    WEB_URL = "https://ekapp.gxufe.edu.cn"  # 你的网址

    # 【保存位置】生成的表格放在哪，留空就保存在当前文件夹（默认，文件名不变）
    # 想换个文件夹：SAVE_PATH = r"D:\我的文件\第二课堂"
    # 想连文件名一起改：SAVE_PATH = r"D:\我的文件\我的学时统计.xlsx"
    SAVE_PATH = ""

    with open(file="account.json", mode='r', encoding='utf-8') as f:
        account = json.load(f)
    YOUR_ID = account["YOUR_ID"]
    YOUR_PWD = account["YOUR_PWD"]

    # 【关键开关】设置为 True 开启无痕模式（每次都是新环境），False 则使用正常模式（享受记住密码的便利）
    USE_INCOGNITO = True
    automator = SecondClassroomAutomator(WEB_URL, YOUR_ID, YOUR_PWD, incognito=USE_INCOGNITO, save_path=SAVE_PATH)
    automator.run()