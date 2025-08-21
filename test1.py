import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
# from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import threading
import time
import winsound
import random
import re
from urllib.parse import urlparse, parse_qs
import os
import json

# 如使用非默认安装位置的 Chrome 测试版，可在此指定二进制路径；留空则使用系统默认
CHROME_BINARY_PATH = r"D:\迅雷下载\chrome-win64\chrome-win64\chrome.exe"  # 例如 r"C:\\Program Files\\Google\\Chrome Beta\\Application\\chrome.exe"
# 可选：指定用户数据目录以保持登录态（填写你自己的 Chrome 配置目录）
CHROME_USER_DATA_DIR = ""  # 例如 r"C:\\Users\\你的用户名\\AppData\\Local\\Google\\Chrome Beta\\User Data"

CONFIG_FILE = "config.json"

class TaobaoMonitorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("淘宝/天猫 商品监控工具")
        self.root.geometry("540x520")
        self.root.wm_attributes("-topmost", False)

        tk.Label(root, text="商品链接：").pack(pady=3)
        self.url_entry = tk.Entry(root, width=62)
        self.url_entry.pack()

        top_row = tk.Frame(root)
        top_row.pack(pady=3)
        tk.Label(top_row, text="刷新间隔（秒）：").pack(side=tk.LEFT)
        self.interval_entry = tk.Entry(top_row, width=8)
        self.interval_entry.insert(0, "3")
        self.interval_entry.pack(side=tk.LEFT, padx=5)

        tk.Label(top_row, text="SKU关键词：").pack(side=tk.LEFT)
        self.sku_keywords_entry = tk.Entry(top_row, width=28)
        self.sku_keywords_entry.pack(side=tk.LEFT)

        opts_row = tk.Frame(root)
        opts_row.pack(pady=3)
        self.sound_var = tk.BooleanVar(value=True)
        tk.Checkbutton(opts_row, text="声音提示", variable=self.sound_var).pack(side=tk.LEFT)
        tk.Label(opts_row, text="次数").pack(side=tk.LEFT)
        self.sound_times_entry = tk.Entry(opts_row, width=5)
        self.sound_times_entry.insert(0, "3")
        self.sound_times_entry.pack(side=tk.LEFT, padx=(0,10))
        self.topmost_var = tk.BooleanVar(value=False)
        tk.Checkbutton(opts_row, text="窗口置顶", variable=self.topmost_var, command=self.apply_topmost).pack(side=tk.LEFT)
        # 新增选项
        self.cart_as_instock_var = tk.BooleanVar(value=True)
        tk.Checkbutton(opts_row, text="加入购物车视为有货", variable=self.cart_as_instock_var).pack(side=tk.LEFT, padx=(10,0))

        debug_row = tk.Frame(root)
        debug_row.pack(pady=3)
        self.debug_log_var = tk.BooleanVar(value=False)
        tk.Checkbutton(debug_row, text="调试日志", variable=self.debug_log_var).pack(side=tk.LEFT)

        path_row = tk.Frame(root)
        path_row.pack(pady=3)
        tk.Label(path_row, text="Chrome二进制：").pack(side=tk.LEFT)
        self.chrome_path_entry = tk.Entry(path_row, width=42)
        self.chrome_path_entry.insert(0, CHROME_BINARY_PATH)
        self.chrome_path_entry.pack(side=tk.LEFT)
        tk.Button(path_row, text="浏览...", command=self.pick_chrome_path).pack(side=tk.LEFT, padx=5)

        prof_row = tk.Frame(root)
        prof_row.pack(pady=3)
        tk.Label(prof_row, text="用户数据目录：").pack(side=tk.LEFT)
        self.profile_entry = tk.Entry(prof_row, width=42)
        self.profile_entry.insert(0, CHROME_USER_DATA_DIR)
        self.profile_entry.pack(side=tk.LEFT)
        tk.Button(prof_row, text="选择...", command=self.pick_profile_dir).pack(side=tk.LEFT, padx=5)

        btn_row = tk.Frame(root)
        btn_row.pack(pady=6)
        self.start_btn = tk.Button(btn_row, text="开始监控", command=self.start_monitoring)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        self.stop_btn = tk.Button(btn_row, text="停止监控", state=tk.DISABLED, command=self.stop_monitoring)
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        self.login_btn = tk.Button(btn_row, text="打开登录页", command=self.open_login_page)
        self.login_btn.pack(side=tk.LEFT, padx=5)
        self.open_btn = tk.Button(btn_row, text="打开商品页", command=self.open_product_page)
        self.open_btn.pack(side=tk.LEFT, padx=5)

        self.status_label = tk.Label(root, text="状态：未运行", fg="blue")
        self.status_label.pack(pady=6)

        # 日志框
        self.log_text = tk.Text(root, height=10, width=70, state=tk.DISABLED)
        self.log_text.pack(pady=4)

        self.driver = None
        self.running = False
        self.error_backoff_s = 0

        # 加载配置
        self.load_config()
        self.apply_topmost()

    def log(self, text):
        ts = time.strftime("%H:%M:%S")
        self.root.after(0, lambda: self._append_log(f"[{ts}] {text}\n"))

    def _append_log(self, text):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, text)
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def apply_topmost(self):
        try:
            self.root.wm_attributes("-topmost", bool(self.topmost_var.get()))
        except Exception:
            pass

    def pick_chrome_path(self):
        path = filedialog.askopenfilename(title="选择 Chrome 可执行文件", filetypes=[["Chrome", "chrome.exe"], ["所有文件", "*"]])
        if path:
            self.chrome_path_entry.delete(0, tk.END)
            self.chrome_path_entry.insert(0, path)

    def pick_profile_dir(self):
        path = filedialog.askdirectory(title="选择用户数据目录")
        if path:
            self.profile_entry.delete(0, tk.END)
            self.profile_entry.insert(0, path)

    def save_config(self):
        cfg = {
            "chrome_binary": self.chrome_path_entry.get().strip(),
            "profile_dir": self.profile_entry.get().strip(),
            "interval": self.interval_entry.get().strip(),
            "sku_keywords": self.sku_keywords_entry.get().strip(),
            "sound": bool(self.sound_var.get()),
            "sound_times": self.sound_times_entry.get().strip(),
            "topmost": bool(self.topmost_var.get()),
            "cart_as_instock": bool(self.cart_as_instock_var.get()),
            "debug_log": bool(self.debug_log_var.get()),
            "last_url": self.url_entry.get().strip(),
        }
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def load_config(self):
        try:
            if not os.path.exists(CONFIG_FILE):
                return
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            self.chrome_path_entry.delete(0, tk.END)
            self.chrome_path_entry.insert(0, cfg.get("chrome_binary", ""))
            self.profile_entry.delete(0, tk.END)
            self.profile_entry.insert(0, cfg.get("profile_dir", ""))
            self.interval_entry.delete(0, tk.END)
            self.interval_entry.insert(0, cfg.get("interval", "3"))
            self.sku_keywords_entry.delete(0, tk.END)
            self.sku_keywords_entry.insert(0, cfg.get("sku_keywords", ""))
            self.sound_var.set(bool(cfg.get("sound", True)))
            self.sound_times_entry.delete(0, tk.END)
            self.sound_times_entry.insert(0, str(cfg.get("sound_times", "3")))
            self.topmost_var.set(bool(cfg.get("topmost", False)))
            self.cart_as_instock_var.set(bool(cfg.get("cart_as_instock", True)))
            self.debug_log_var.set(bool(cfg.get("debug_log", False)))
            if cfg.get("last_url"):
                self.url_entry.delete(0, tk.END)
                self.url_entry.insert(0, cfg.get("last_url"))
        except Exception:
            pass

    def beep(self):
        try:
            if not bool(self.sound_var.get()):
                return
            times = int(self.sound_times_entry.get().strip() or "1")
        except Exception:
            times = 3
        for _ in range(max(1, times)):
            winsound.Beep(1200, 500)
            time.sleep(0.05)

    def check_buy_button(self):
        """检查页面是否有立即购买按钮，返回元素或 None"""
        selectors = [
            ("css", "#J_LinkBuy"),
            ("css", ".tb-btn-buy"),
            ("css", "#J_BuyButton"),
            ("css", ".tm-btn-buy"),
            ("css", "button.buy"),
            ("xpath", "//*[contains(@class,'btn') and contains(., '立即购买')]"),
            ("xpath", "//a[contains(., '立即购买')]|//button[contains(., '立即购买')]"),
        ]
        for how, expr in selectors:
            try:
                el = self.driver.find_element(By.CSS_SELECTOR, expr) if how == "css" else self.driver.find_element(By.XPATH, expr)
                if el:
                    if self.debug_log_var.get():
                        self.log(f"找到立即购买候选: {how}:{expr}, enabled={el.is_enabled()} visible={el.is_displayed()}")
                    return el
            except Exception:
                continue
        if self.debug_log_var.get():
            self.log("未找到立即购买按钮")
        return None

    def find_cart_button(self):
        """尝试找到加入购物车按钮"""
        selectors = [
            ("css", "#J_LinkBasket"),
            ("css", "#J_AddToCart"),
            ("css", ".tb-btn-add a, .tb-btn-add button, .tm-btn-add"),
            ("xpath", "//a[contains(., '加入购物车')]|//button[contains(., '加入购物车')]"),
        ]
        for how, expr in selectors:
            try:
                el = self.driver.find_element(By.CSS_SELECTOR, expr) if how == "css" else self.driver.find_element(By.XPATH, expr)
                if el:
                    if self.debug_log_var.get():
                        self.log(f"找到加入购物车候选: {how}:{expr}, enabled={el.is_enabled()} visible={el.is_displayed()}")
                    return el
            except Exception:
                continue
        if self.debug_log_var.get():
            self.log("未找到加入购物车按钮")
        return None

    def update_status(self, text, color):
        self.root.after(0, lambda: self.status_label.config(text=text, fg=color))

    def is_in_stock(self):
        """更稳健的库存判断：按钮可点击且未禁用，且页面不存在显式售罄提示；可选：加入购物车可用视为有货"""
        try:
            buy_btn = self.check_buy_button()
            if buy_btn and buy_btn.is_displayed():
                classes = (buy_btn.get_attribute("class") or "").lower()
                disabled_attr = buy_btn.get_attribute("disabled")
                enabled = ("disabled" not in classes) and ("tb-disabled" not in classes) and (disabled_attr is None) and buy_btn.is_enabled()
                if self.debug_log_var.get():
                    self.log(f"立即购买状态: enabled={enabled}, classes={classes}, disabled_attr={disabled_attr}")
                if enabled:
                    return True
        except Exception as e:
            if self.debug_log_var.get():
                self.log(f"检测立即购买异常: {e}")
        # 若允许将加入购物车视为有货
        if bool(self.cart_as_instock_var.get()):
            try:
                cart_btn = self.find_cart_button()
                if cart_btn and cart_btn.is_displayed() and cart_btn.is_enabled():
                    if self.debug_log_var.get():
                        self.log("加入购物车可用，视为有货")
                    return True
            except Exception as e:
                if self.debug_log_var.get():
                    self.log(f"检测加入购物车异常: {e}")
        # 若出现明显售罄/无货字样，则认为无货
        try:
            page = self.driver.page_source
            for bad in ["已售罄", "无货", "下架", "到货通知", "补货中", "暂不支持购买", "暂无库存"]:
                if bad in page:
                    if self.debug_log_var.get():
                        self.log(f"检测到无货提示：{bad}")
                    return False
        except Exception:
            pass
        return False

    def select_sku_by_keywords(self, keywords):
        """尽力根据输入的关键词选择 SKU 选项。keywords 为小写关键词列表"""
        if not keywords:
            return
        try:
            prop_selectors = [".J_Prop", ".tb-prop", ".tm-sale-prop", ".sku"]
            for selector in prop_selectors:
                sections = self.driver.find_elements(By.CSS_SELECTOR, selector)
                for sec in sections:
                    try:
                        options = sec.find_elements(By.CSS_SELECTOR, "li, li a")
                        for opt in options:
                            try:
                                text = opt.text.strip().lower()
                                if not text:
                                    continue
                                if any(k in text for k in keywords):
                                    # 点击 li 或内部 a
                                    target = opt
                                    if opt.tag_name.lower() != "li":
                                        try:
                                            target = opt.find_element(By.XPATH, "ancestor::li[1]")
                                        except Exception:
                                            pass
                                    if target:
                                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target)
                                        target.click()
                                        time.sleep(0.1)
                            except Exception:
                                continue
                    except Exception:
                        continue
        except Exception:
            pass

    def build_chrome_options(self):
        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)
        if self.chrome_path_entry.get().strip():
            options.binary_location = self.chrome_path_entry.get().strip()
        if self.profile_entry.get().strip():
            options.add_argument(f"--user-data-dir={self.profile_entry.get().strip()}")
        # 尝试降低被自动化识别的概率（可选）
        options.add_experimental_option("excludeSwitches", ["enable-automation"])  # type: ignore
        options.add_experimental_option("useAutomationExtension", False)  # type: ignore
        options.add_argument("--disable-blink-features=AutomationControlled")
        return options

    def ensure_driver(self):
        if self.driver is None:
            self.driver = webdriver.Chrome(options=self.build_chrome_options())

    def normalize_product_url(self, url: str) -> str:
        """将移动端链接尽量转换为 PC 端商品链接"""
        try:
            parsed = urlparse(url)
            host = (parsed.netloc or "").lower()
            qs = parse_qs(parsed.query)
            item_id = None
            # 常见参数 id / itemId
            for key in ["id", "itemId", "item_id"]:
                if key in qs and qs[key]:
                    item_id = qs[key][0]
                    break
            if not item_id:
                m = re.search(r"[?&](?:id|itemId|item_id)=(\d+)", url)
                if m:
                    item_id = m.group(1)
            if not item_id:
                return url
            if "tmall" in host:
                return f"https://detail.tmall.com/item.htm?id={item_id}"
            if "taobao" in host:
                return f"https://item.taobao.com/item.htm?id={item_id}"
            return url
        except Exception:
            return url

    def open_login_page(self):
        try:
            self.ensure_driver()
            self.driver.get("https://login.taobao.com/member/login.jhtml")
            self.update_status("状态：请在新页面完成登录", "orange")
            self.log("已打开登录页")
        except Exception:
            self.update_status("状态：打开登录页失败", "red")
            self.log("打开登录页失败")

    def open_product_page(self):
        url = self.url_entry.get().strip()
        if not url:
            messagebox.showwarning("提示", "请输入商品链接！")
            return
        try:
            self.ensure_driver()
            self.driver.get(self.normalize_product_url(url))
            self.update_status("状态：已打开商品页", "green")
            self.log("已打开商品页")
        except Exception:
            self.update_status("状态：打开商品页失败", "red")
            self.log("打开商品页失败")

    def monitor(self, url, interval):
        url = self.normalize_product_url(url)
        self.ensure_driver()
        self.driver.get(url)

        self.update_status("状态：监控中...", "green")
        self.log("开始监控")
        found_stock = False

        # 预处理：如填写了 SKU 关键词，先尝试选择
        try:
            keywords_str = self.sku_keywords_entry.get().strip()
            keywords = [k.lower() for k in keywords_str.split() if k]
        except Exception:
            keywords = []

        while self.running:
            try:
                self.driver.refresh()
                # 等待页面主结构加载完成
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

                if keywords:
                    self.select_sku_by_keywords(keywords)

                if self.is_in_stock():
                    self.update_status("状态：发现有货！", "red")
                    self.log("检测到可下单/可加入购物车")
                    self.beep()
                    self.root.after(0, lambda: messagebox.showinfo("提醒", "发现有货！请尽快下单！"))
                    found_stock = True
                    break
                else:
                    self.update_status("状态：暂无库存，继续监控...", "blue")
                    if self.debug_log_var.get():
                        self.log("本轮未发现库存")

            except Exception as e:
                self.update_status("状态：可能未登录/页面异常，已退避重试", "orange")
                self.log(f"异常：{e}")
                # 指数退避：最多退避到 60 秒
                self.error_backoff_s = min(60, max(1, (self.error_backoff_s or 1) * 2))

            # 将退避时间叠加到下一次刷新间隔，并在成功一次轮询后清零
            wait_s = max(0.5, interval + random.uniform(-0.5, 0.5)) + (self.error_backoff_s or 0)
            time.sleep(wait_s)
            self.error_backoff_s = 0

        # 结束后：如发现有货则保留页面，反之才完全停止与关闭
        if found_stock:
            # 恢复按钮状态，但不关闭浏览器
            self.running = False
            self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.stop_btn.config(state=tk.DISABLED))
        else:
            self.root.after(0, self.stop_monitoring)

    def start_monitoring(self):
        url = self.url_entry.get().strip()
        if not url:
            messagebox.showwarning("提示", "请输入商品链接！")
            return
        try:
            interval = int(self.interval_entry.get())
        except ValueError:
            messagebox.showwarning("提示", "刷新间隔必须是数字！")
            return

        self.running = True
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)

        threading.Thread(target=self.monitor, args=(url, interval), daemon=True).start()
        self.save_config()
        self.log("已开始后台监控线程")

    def stop_monitoring(self):
        self.running = False
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
        self.driver = None
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.status_label.config(text="状态：已停止", fg="black")
        self.log("已停止监控并关闭浏览器")


if __name__ == "__main__":
    root = tk.Tk()
    app = TaobaoMonitorApp(root)
    root.mainloop()
