import os
import sys
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
import time
import base64
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options


class PptCrawlerApp:
    def __init__(self, root):
        self.firstppt=True
        self.root = root
        self.root.title("课程爬取爬取")
        self.root.geometry("700x600")

        self.DRIVER_PATH = r"D:/edgedriver_win64/msedgedriver.exe"

        # Cookie 输入区
        frm_cookie = tk.Frame(self.root)
        frm_cookie.pack(pady=10, fill='x', padx=10)
        tk.Label(frm_cookie, text="请输入 Cookie：").pack(side='left')
        self.cookie_entry = tk.Entry(frm_cookie)
        self.cookie_entry.pack(side='left', fill='x', expand=True, padx=5)
        self.start_button = tk.Button(frm_cookie, text="开始获取课程列表", command=self.on_start_clicked)
        self.start_button.pack(side='left', padx=5)

        # 课程列表区
        lbl_courses = tk.Label(self.root, text="课程列表：")
        lbl_courses.pack(anchor='w', padx=10)
        self.course_listbox = tk.Listbox(self.root, height=8)
        self.course_listbox.pack(fill='x', padx=10)
        self.select_course_button = tk.Button(self.root, text="选择所选课程", command=self.on_select_course, state='disabled')
        self.select_course_button.pack(pady=5)

        # 回放列表区
        lbl_replays = tk.Label(self.root, text="回放列表（仅“回放”状态）：")
        lbl_replays.pack(anchor='w', padx=10)
        self.replay_listbox = tk.Listbox(self.root, height=8)
        self.replay_listbox.pack(fill='x', padx=10)
        self.select_replay_button = tk.Button(self.root, text="选择所选回放并开始爬取", command=self.on_select_replay, state='disabled')
        self.select_replay_button.pack(pady=5)

        # 进度区
        frm_progress = tk.Frame(self.root)
        frm_progress.pack(fill='x', pady=10, padx=10)
        tk.Label(frm_progress, text="爬取进度：").pack(side='left')
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_bar = ttk.Progressbar(frm_progress, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(side='left', fill='x', expand=True, padx=5)
        self.progress_label = tk.Label(frm_progress, text="0%")
        self.progress_label.pack(side='left')

        # 状态日志区
        self.status_text = tk.Text(self.root, height=8, state='disabled')
        self.status_text.pack(fill='both', padx=10, pady=5, expand=True)

        # Selenium driver placeholder
        self.driver = None
        # 存储临时用户数据目录，以便结束时清理
        self.temp_user_data_dir = None
        # 存储课程及回放信息
        self.courses = []
        self.replays = []

        # 目标 URL
        self.target_url = "https://classroom.guet.edu.cn/education/?tenant_code=21#/home?menu_code=C-wdkc,-1"
        self.log("""警告：
此应用仅可用于个人的课程学习，请尊重版权，勿将爬取内容在网上传播!""")

    def log(self, msg: str):
        """在状态框中写日志，并刷新UI"""
        self.status_text.configure(state='normal')
        self.status_text.insert('end', msg + '\n')
        self.status_text.see('end')
        self.status_text.configure(state='disabled')
        self.root.update()

    def on_start_clicked(self):
        cookie = self.cookie_entry.get().strip()
        if not cookie:
            messagebox.showerror("错误", "Cookie 不能为空！")
            return
        # 禁用按钮，防止重复点击
        self.start_button.config(state='disabled')
        self.log("开始初始化浏览器并登录...")
        self.start_crawling(cookie)

    def start_crawling(self, cookie_str: str):
        try:
            if self.firstppt is True:
                # 创建临时用户数据目录，避免与已有浏览器进程冲突
                #self.temp_user_data_dir = tempfile.mkdtemp(prefix="selenium_edge_")
                #print('开始初始化')
                edge_options = Options()
                edge_options.add_argument("--disable-blink-features=AutomationControlled")  # 避免被检测为自动化脚本
                edge_options.add_argument(
                    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36 Edg/132.0.0.0")  # 模拟正常浏览器的User-Agent
                edge_options.add_argument("--headless=new")  # 启用无头模式
                edge_options.add_argument("--disable-gpu")  # 禁用 GPU 加速
                edge_options.add_argument("--remote-debugging-port=9222")  # 启用远程调试端口
                edge_options.add_argument("--window-size=1920,1080")
                service = Service(self.DRIVER_PATH)
                self.driver = webdriver.Edge(service=service, options=edge_options)

                # 先打开域名主页，方便设置 cookie
                self.driver.get(self.target_url)
                time.sleep(3)
                # 解析并添加 cookie
                # 构建 cookie 列表
                #print("开始添加cookie")
                cookies = []
                for item in cookie_str.split(";"):
                    item = item.strip()  # 去掉多余的空格
                    item = item.strip('\n')
                    if "=" in item:  # 确保 item 中有 "="
                        key, value = item.split("=", 1)  # 使用 1 来限制分割成两个部分
                        cookies.append({"name": key, "value": value, "domain": "guet.edu.cn", "path": "/"})

                # 输出 cookies 列表
                #print(cookies)

                for cookie in cookies:
                    self.driver.add_cookie(cookie)


            self.driver.refresh()
            time.sleep(3)
            self.driver.get(self.target_url)
            time.sleep(3)
            #print("加载完毕")
            time.sleep(2)
            self.log("已打开目标页面，切换 iframe ...")
            # 切换到 iframe
            try:
                iframe = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.ID, "inlineFrameExample"))
                )
                self.driver.switch_to.frame(iframe)
            except Exception as e:
                self.log(f"未能切换到 iframe: {e}")
                self.cleanup_driver()
                self.enable_start_button()
                return

            # 等待并获取课程元素
            elems = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.courseInfo"))
            )
            self.courses.clear()
            for idx, ce in enumerate(elems):
                try:
                    title = ce.find_element(By.CSS_SELECTOR, ".course-label").text.strip()
                    teacher = ce.find_element(By.CSS_SELECTOR, ".course-desc-p").text.strip()
                    self.courses.append({"title": title, "teacher": teacher, "element": ce})
                    self.log(f"课程 {idx+1}: {title} - {teacher}")
                except Exception:
                    continue

            # 更新 GUI 课程列表
            self.course_listbox.delete(0, 'end')
            for c in self.courses:
                self.course_listbox.insert('end', f"{c['title']} - {c['teacher']}")
            self.select_course_button.config(state='normal')
            self.log("请在列表中选择课程，然后点击“选择所选课程”")
        except Exception as ex:
            self.log(f"初始化或获取课程列表时出错: {ex}")
            self.cleanup_driver()
            self.enable_start_button()

    def enable_start_button(self):
        self.start_button.config(state='normal')
        self.root.update()

    def on_select_course(self):
        sel = self.course_listbox.curselection()
        if not sel:
            messagebox.showerror("错误", "请先选择课程！")
            return
        index = sel[0]
        selected = self.courses[index]
        self.log(f"选择了课程：{selected['title']} - {selected['teacher']}")
        self.select_course_button.config(state='disabled')
        self.fetch_replays(selected)

    def fetch_replays(self, course_item: dict):
        time.sleep(3)
        try:
            # 点击进入课程
            try:
                btn = course_item["element"].find_element(By.CSS_SELECTOR, ".el-button--primary")
                btn.click()
            except Exception as e:
                self.log(f"点击课程进入失败: {e}")
                self.select_course_button.config(state='normal')
                return
            time.sleep(2)
            # 切换到新窗口
            handles = self.driver.window_handles
            if len(handles) > 1:
                self.driver.switch_to.window(handles[-1])
            time.sleep(1)

            # 获取所有 <p> 标签，即回放记录
            p_elements = self.driver.find_elements(By.CSS_SELECTOR, ".content-inner-one > p")
            self.replays.clear()
            for idx, p in enumerate(p_elements):
                try:
                    spans = p.find_elements(By.TAG_NAME, "span")
                    time_info = spans[1].text.strip()
                    teacher = spans[2].get_attribute("title") or spans[2].text.strip()
                    status = spans[4].text.strip()
                    self.replays.append({"time": time_info, "teacher": teacher, "status": status, "element": p})
                    self.log(f"回放 {idx+1}: {time_info} - {teacher} ({status})")
                except Exception:
                    continue

            # 更新 GUI 回放列表，仅显示 status == "回放"
            self.replay_listbox.delete(0, 'end')
            for r in self.replays:
                if r["status"] == "回放":
                    self.replay_listbox.insert('end', f"{r['time']} - {r['teacher']}")
            self.select_replay_button.config(state='normal')
            self.log("请在列表中选择回放，然后点击“选择所选回放并开始爬取”")
        except Exception as ex:
            self.log(f"获取回放列表出错: {ex}")
            self.select_course_button.config(state='normal')

    def on_select_replay(self):
        sel = self.replay_listbox.curselection()
        if not sel:
            messagebox.showerror("错误", "请先选择回放！")
            return
        index = sel[0]
        # 过滤后的列表
        filtered = [r for r in self.replays if r["status"] == "回放"]
        selected = filtered[index]
        self.log(f"选择了回放：{selected['time']} - {selected['teacher']}")
        self.select_replay_button.config(state='disabled')
        self.crawl_ppt(selected)

    def crawl_ppt(self, replay_item: dict):
        try:
            # 点击回放按钮打开回放窗口
            try:
                parent = self.driver.current_window_handle
                #print('开始回放按钮')
                spans = replay_item["element"].find_elements(By.TAG_NAME, "span")
                btn = spans[4]
                #print('回放按钮查找成功，进行点击')
                btn.click()
                time.sleep(3)
            except Exception as e:
                self.log(f"点击回放进入失败: {e}")
                self.select_replay_button.config(state='normal')
                return
            # 3. 切到新窗口（总是最后一个）
            # WebDriverWait(self.driver, 10).until(lambda d: len(d.window_handles) > 1)
            # new_win = [h for h in self.driver.window_handles if h != parent][0]
            # self.driver.switch_to.window(new_win)

            self.driver.switch_to.window(self.driver.window_handles[-1])

            time.sleep(3)

            # 4. 在新窗口做完事以后，先关闭它
            self.driver.close()

            # 5. 切回父窗口
            #self.driver.switch_to.window(parent)
            self.driver.switch_to.window(self.driver.window_handles[-1])

            # 6. 重新定位按钮并再次点击
            status_button = replay_item["element"].find_elements(By.TAG_NAME, "span")[4]
            #print("准备再次点击")
            status_button.click()
            #print("准备再次点击成功")
            time.sleep(3)
            # 切换到新打开的窗口/标签
            all_windows = self.driver.window_handles  # 获取所有窗口句柄
            self.driver.switch_to.window(all_windows[-1])  # 切换到最新打开的窗口

            # 准备 PPT
            prs = Presentation()

            #print('开始进行爬取')
            # change_item = driver.find_element(By.CLASS_NAME, "change-item__img")
            change_item = WebDriverWait(self.driver, 20).until(
                EC.visibility_of_element_located((By.CLASS_NAME, "change-item__img"))
            )
            #print('找到标签')
            change_item.click()

            # 获取总页数信息
            try:
                page_info_elem = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, 'ppt_page_con'))
                )
                page_info = page_info_elem.text.strip()
                current_page, total_pages = map(int, page_info.split('/'))
            except Exception:
                self.log("无法获取页数信息，停止爬取")
                self.cleanup_driver()
                self.enable_start_button()
                return

            # 从第一页开始循环抓取
            for page_no in range(1, total_pages + 1):
                # 等待 canvas 出现并抓图
                try:
                    WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.ID, 'ppt_canvas'))
                    )
                    img_data = self.driver.execute_script(
                        "return document.getElementById('ppt_canvas').toDataURL('image/png').substring(22);"
                    )
                    img_bytes = base64.b64decode(img_data)
                    slide_layout = prs.slide_layouts[5]
                    slide = prs.slides.add_slide(slide_layout)
                    image_stream = BytesIO(img_bytes)
                    left = Inches(1); top = Inches(1); height = Inches(5)
                    slide.shapes.add_picture(image_stream, left, top, height=height)
                    self.log(f"已抓取第 {page_no} 页 / 共 {total_pages} 页")
                except Exception as e:
                    self.log(f"第 {page_no} 页抓取失败: {e}")
                # 更新进度条
                pct = page_no / total_pages * 100
                self.progress_var.set(pct)
                self.progress_label.config(text=f"{int(pct)}%")
                self.root.update()
                # 如果不是最后一页，点击“下一页”
                if page_no < total_pages:
                    try:
                        next_btn = self.driver.find_element(By.CLASS_NAME, 'ppt_btn_next')
                        next_btn.click()
                        time.sleep(0.1)
                    except Exception as e:
                        self.log(f"翻页到第 {page_no+1} 页失败: {e}")
                        break

            # 提示用户输入文件名
            filename = simpledialog.askstring("保存文件", "请输入 PPT 文件名（不带后缀）：", parent=self.root)
            if not filename:
                filename = f"output_{int(time.time())}"
            save_path = f"PPT\\{filename}.pptx"
            prs.save(save_path)
            self.log(f"PPT 已保存为: {save_path}")
            messagebox.showinfo("完成", f"PPT 已保存为当前目录下的：{save_path}")

            # 结束时关闭浏览器并清理
            # self.cleanup_driver()

            # 重置 UI 状态
            self.progress_var.set(0)
            self.progress_label.config(text="0%")
            self.start_button.config(state='active')
            self.select_course_button.config(state='disabled')
            self.select_replay_button.config(state='disabled')
            self.course_listbox.delete(0, 'end')
            self.replay_listbox.delete(0, 'end')
            self.firstppt=False
            self.log("流程结束，可再次重试")

        except Exception as ex:
            self.log(f"爬取过程中出错: {ex}")
            self.cleanup_driver()
            self.start_button.config(state='normal')
            self.select_course_button.config(state='disabled')
            self.select_replay_button.config(state='disabled')
        # self.driver.quit()
        # os.system("taskkill /F /IM msedge.exe")
        # # 然后结束主循环并退出程序
        # self.root.destroy()
        # # 如果希望彻底退出，避免残留线程，可再调用 sys.exit()
        # sys.exit(0)

        self.driver.close()
        self.driver.switch_to.window(self.driver.window_handles[-1])
        self.driver.close()
        self.driver.switch_to.window(self.driver.window_handles[-1])



    def cleanup_driver(self):
        # 关闭浏览器并删除临时用户数据目录
        try:
            if self.driver:
                self.driver.quit()
                os.system("taskkill /F /IM msedge.exe")
                # 然后结束主循环并退出程序
                self.root.destroy()
                # 如果希望彻底退出，避免残留线程，可再调用 sys.exit()
                sys.exit(0)
        except:
            pass

def main():
    root = tk.Tk()
    PptCrawlerApp(root)  # 不必保存到变量也能正常工作
    root.mainloop()

if __name__ == "__main__":
    main()
