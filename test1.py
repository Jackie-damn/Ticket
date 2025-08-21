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
        self.interval_entry.insert(0, "0.1")
        self.interval_entry.pack(side=tk.LEFT, padx=5)

        tk.Label(top_row, text="SKU关键词：").pack(side=tk.LEFT)
        self.sku_keywords_entry = tk.Entry(top_row, width=28)
        self.sku_keywords_entry.pack(side=tk.LEFT)

        # 新增：最短整页刷新间隔
        refresh_row = tk.Frame(root)
        refresh_row.pack(pady=3)
        tk.Label(refresh_row, text="最短整页刷新（秒）：").pack(side=tk.LEFT)
        self.refresh_min_entry = tk.Entry(refresh_row, width=8)
        self.refresh_min_entry.insert(0, "2")
        self.refresh_min_entry.pack(side=tk.LEFT, padx=5)

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
        self.auto_buy_var = tk.BooleanVar(value=True)
        tk.Checkbutton(debug_row, text="检测到可下单自动下单", variable=self.auto_buy_var).pack(side=tk.LEFT, padx=(10,0))

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
            "refresh_min": self.refresh_min_entry.get().strip(),
            "sku_keywords": self.sku_keywords_entry.get().strip(),
            "sound": bool(self.sound_var.get()),
            "sound_times": self.sound_times_entry.get().strip(),
            "topmost": bool(self.topmost_var.get()),
            "cart_as_instock": bool(self.cart_as_instock_var.get()),
            "debug_log": bool(self.debug_log_var.get()),
            "auto_buy": bool(self.auto_buy_var.get()),
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
            self.interval_entry.insert(0, cfg.get("interval", "0.1"))
            self.refresh_min_entry.delete(0, tk.END)
            self.refresh_min_entry.insert(0, cfg.get("refresh_min", "2"))
            self.sku_keywords_entry.delete(0, tk.END)
            self.sku_keywords_entry.insert(0, cfg.get("sku_keywords", ""))
            self.sound_var.set(bool(cfg.get("sound", True)))
            self.sound_times_entry.delete(0, tk.END)
            self.sound_times_entry.insert(0, str(cfg.get("sound_times", "3")))
            self.topmost_var.set(bool(cfg.get("topmost", False)))
            self.cart_as_instock_var.set(bool(cfg.get("cart_as_instock", True)))
            self.debug_log_var.set(bool(cfg.get("debug_log", False)))
            self.auto_buy_var.set(bool(cfg.get("auto_buy", True)))
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

    def normalize_text(self, s: str) -> str:
        try:
            t = (s or "").strip().lower()
            # 去除常见空白与标点
            for ch in [" ", "\n", "\t", "\r", "-", "_", ",", ".", "。", "，", "、", "（", "）", "(", ")", "[", "]", "{", "}", "|", ":", ";", "'", '"']:
                t = t.replace(ch, "")
            return t
        except Exception:
            return s or ""

    def extract_option_text(self, el) -> str:
        """尽可能多地提取一个选项的可读文本，用于匹配关键词"""
        texts = []
        try:
            txt = (el.text or "").strip()
            if txt:
                texts.append(txt)
        except Exception:
            pass
        for attr in ["title", "aria-label", "data-value", "data-title", "data-text", "data-tip"]:
            try:
                v = el.get_attribute(attr)
                if v:
                    texts.append(v)
            except Exception:
                pass
        try:
            img = el.find_element(By.CSS_SELECTOR, "img")
            alt = img.get_attribute("alt")
            if alt:
                texts.append(alt)
        except Exception:
            pass
        merged = " ".join(texts)
        return merged

    def list_sku_sections(self):
        """返回页面上可能的 SKU 区块元素列表"""
        sections = []
        selectors = [
            ".J_Prop", ".tm-sale-prop", ".tb-prop", ".J_TSaleProp",
            "dl[ts-prop]", "dl[data-property]", ".sku-prop", ".skuWrap", ".sku-prop-item",
            "[data-property]", ".sku", "div[role='radiogroup']",
        ]
        for sel in selectors:
            try:
                sections.extend(self.driver.find_elements(By.CSS_SELECTOR, sel))
            except Exception:
                continue
        # 去重
        uniq = []
        seen = set()
        for s in sections:
            try:
                key = s.id
                if key not in seen:
                    seen.add(key)
                    uniq.append(s)
            except Exception:
                continue
        return uniq

    def is_option_selected(self, li_el):
        try:
            cls = (li_el.get_attribute("class") or "").lower()
            if any(flag in cls for flag in [
                "selected", "tb-selected", "tm-selected", "is-selected", "isactive", "active", "checked", "cur", "on"
            ]):
                return True
            try:
                a = li_el.find_element(By.CSS_SELECTOR, "a, button, label, input")
                acls = (a.get_attribute("class") or "").lower()
                if any(flag in acls for flag in ["selected", "active", "checked", "cur", "on"]):
                    return True
            except Exception:
                pass
        except Exception:
            pass
        return False

    def is_option_available(self, li_el):
        try:
            cls = (li_el.get_attribute("class") or "").lower()
            if any(x in cls for x in [
                "disabled", "sold-out", "tm-disabled", "tb-out-of-stock", "tb-skuprop-disable", "nostock", "forbid", "ban", "unavailable"
            ]):
                return False
            if li_el.get_attribute("aria-disabled") in ("true", "1"):
                return False
            if li_el.get_attribute("disabled") is not None:
                return False
        except Exception:
            pass
        return True

    def find_clickable_ancestor(self, el):
        try:
            for xp in [
                "ancestor::li[1]", "ancestor::label[1]", "ancestor::button[1]", "ancestor::a[1]", "ancestor::div[1]"
            ]:
                try:
                    anc = el.find_element(By.XPATH, xp)
                    if anc:
                        return anc
                except Exception:
                    continue
        except Exception:
            pass
        return el

    def click_by_title_tokens(self, keywords):
        """优先通过属性包含关系匹配并点击：title/aria-label/data-value/data-title 等。
        返回点击成功次数。
        """
        if not keywords:
            return 0
        clicked = 0
        for k in keywords:
            token = (k or "").strip()
            if not token:
                continue
            css_list = [
                f"[title*='{token}']",
                f"[aria-label*='{token}']",
                f"[data-value*='{token}']",
                f"[data-title*='{token}']",
                f"[data-text*='{token}']",
                f"span[title*='{token}']",
            ]
            for css in css_list:
                try:
                    elems = self.driver.find_elements(By.CSS_SELECTOR, css)
                    for e in elems[:5]:
                        anc = self.find_clickable_ancestor(e)
                        if not anc.is_displayed():
                            continue
                        if not self.is_option_available(anc):
                            continue
                        if self.click_option_element(anc):
                            clicked += 1
                except Exception:
                    continue
        return clicked

    def click_global_by_keywords(self, keywords):
        """兜底：全局搜索包含关键词的可点击元素并尝试点击。
        返回点击成功次数。
        """
        if not keywords:
            return 0
        clicked = 0
        # 先尝试属性选择器直接匹配原始关键词（更适合中文与括号等符号）
        clicked += self.click_by_title_tokens(keywords)
        ABC = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        abc = "abcdefghijklmnopqrstuvwxyz"
        for k in keywords:
            token_raw = (k or "").strip()
            if not token_raw:
                continue
            token_norm = self.normalize_text(token_raw)
            # XPath：原始关键词匹配
            xpaths_raw = [
                f"//*[contains(normalize-space(.), '{token_raw}')]",
                f"//*[@title and contains(@title, '{token_raw}')]",
                f"//*[@aria-label and contains(@aria-label, '{token_raw}')]",
                f"//*[@data-value and contains(@data-value, '{token_raw}')]",
            ]
            for xp in xpaths_raw:
                try:
                    elems = self.driver.find_elements(By.XPATH, xp)
                    for e in elems[:5]:
                        anc = self.find_clickable_ancestor(e)
                        if not anc.is_displayed():
                            continue
                        if not self.is_option_available(anc):
                            continue
                        if self.click_option_element(anc):
                            clicked += 1
                except Exception:
                    continue
            # XPath：规范化后关键词匹配（英文环境更友好）
            if token_norm and token_norm != token_raw:
                xp = f"//*[contains(translate(normalize-space(.), '{ABC}', '{abc}'), '{token_norm}')]"
                try:
                    elems = self.driver.find_elements(By.XPATH, xp)
                    for e in elems[:5]:
                        anc = self.find_clickable_ancestor(e)
                        if not anc.is_displayed():
                            continue
                        if not self.is_option_available(anc):
                            continue
                        if self.click_option_element(anc):
                            clicked += 1
                except Exception:
                    pass
        return clicked

    def click_option_element(self, el):
        try:
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
            # 尝试点击内部可点击元素
            for css in ["a", "button", "input", "span", "div"]:
                try:
                    child = el.find_element(By.CSS_SELECTOR, css)
                    child.click()
                    time.sleep(0.05)
                    return True
                except Exception:
                    continue
            # 直接点击自身
            el.click()
            time.sleep(0.05)
            return True
        except Exception:
            return False

    def ensure_all_sku_selected(self, keywords):
        """确保每个 SKU 区块至少选择一个可用选项。优先匹配关键词，否则选择第一个可用。
        返回 (all_selected: bool, selected_count: int)
        """
        kw_tokens = [self.normalize_text(k) for k in (keywords or []) if k]
        total_selected = 0
        sections = self.list_sku_sections()
        for idx, sec in enumerate(sections):
            try:
                options = sec.find_elements(By.CSS_SELECTOR, "li, [role='radio'], .sku-item, .item, span[title]")
                options = [o for o in options if o.is_displayed()]
                if not options:
                    # 该区块下无常见选项节点，使用属性兜底
                    extra = self.click_by_title_tokens(keywords or [])
                    if extra > 0:
                        total_selected += 1
                        if self.debug_log_var.get():
                            self.log(f"规格区块{idx+1}: 通过属性匹配兜底点击 {extra} 次")
                    continue
                # 已选？
                already = False
                for opt in options:
                    if self.is_option_selected(opt):
                        already = True
                        break
                if already:
                    total_selected += 1
                    if self.debug_log_var.get():
                        self.log(f"规格区块{idx+1}: 已有选项，无需更改")
                    continue

                # 为每个选项计算匹配分数
                best = None
                best_score = -1
                best_text = ""
                debug_opts = []
                for opt in options:
                    if not self.is_option_available(opt):
                        continue
                    raw = self.extract_option_text(opt)
                    norm = self.normalize_text(raw)
                    score = 0
                    for k in kw_tokens:
                        if k and (k in norm):
                            score += 1
                    debug_opts.append((raw, score))
                    if score > best_score:
                        best_score = score
                        best = opt
                        best_text = raw

                if self.debug_log_var.get():
                    self.log(f"规格区块{idx+1} 候选: " + " | ".join([f"{t}(score={s})" for t, s in debug_opts[:6]]))

                if best is None or best_score <= 0:
                    # 无匹配关键词，尝试属性兜底
                    extra = self.click_by_title_tokens(keywords or [])
                    if extra > 0:
                        total_selected += 1
                        if self.debug_log_var.get():
                            self.log(f"规格区块{idx+1}: 属性兜底点击 {extra} 次")
                        continue
                    # 仍未命中则退化为第一个可用可见选项
                    for opt in options:
                        if self.is_option_available(opt):
                            best = opt
                            best_text = self.extract_option_text(opt)
                            break

                if best is not None:
                    ok = self.click_option_element(best)
                    if ok:
                        total_selected += 1
                        if self.debug_log_var.get():
                            self.log(f"已选择规格：{(best_text or '').strip()}")
                else:
                    if self.debug_log_var.get():
                        self.log("该规格无可选项或均售罄")
            except Exception as e:
                if self.debug_log_var.get():
                    self.log(f"选择规格异常：{e}")
        need = len(sections)
        return (total_selected >= need and need > 0, total_selected)

    def auto_submit_order(self, keywords=None):
        """自动点击下单：补全规格 -> 点击立即购买 -> 订单确认页 -> 提交订单"""
        try:
            # 0) 尝试补全规格
            try:
                all_ok, cnt = self.ensure_all_sku_selected(keywords)
                if self.debug_log_var.get():
                    self.log(f"规格补全结果 all_ok={all_ok}, 已选={cnt}")
                if cnt == 0:
                    # 兜底：全局关键词点击
                    extra = self.click_global_by_keywords(keywords or [])
                    if self.debug_log_var.get():
                        self.log(f"全局关键词兜底点击次数：{extra}")
            except Exception:
                pass

            # 1) 点击立即购买
            buy_btn = self.check_buy_button()
            if buy_btn and buy_btn.is_enabled():
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", buy_btn)
                buy_btn.click()
                if self.debug_log_var.get():
                    self.log("已点击立即购买，等待订单确认页或校验提示")
            else:
                try:
                    buy_btn = WebDriverWait(self.driver, 5).until(lambda d: self.check_buy_button())
                    if buy_btn and buy_btn.is_enabled():
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", buy_btn)
                        buy_btn.click()
                except Exception:
                    pass

            # 1.1) 若页面出现“请选择”提示，说明规格未全选，补选后再次点击
            try:
                tip = self.driver.find_elements(By.XPATH, "//*[contains(text(),'请选择')]")
                if tip:
                    if self.debug_log_var.get():
                        self.log("检测到‘请选择’提示，尝试自动补选并重试")
                    try:
                        self.ensure_all_sku_selected(keywords)
                        extra = self.click_global_by_keywords(keywords or [])
                        if self.debug_log_var.get():
                            self.log(f"补选后全局兜底点击次数：{extra}")
                    except Exception:
                        pass
                    buy_btn = self.check_buy_button()
                    if buy_btn:
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", buy_btn)
                        self.driver.execute_script("arguments[0].click();", buy_btn)
            except Exception:
                pass

            # 2) 等待跳转到订单确认页（常见为 buy.taobao.com 或 buy.tmall.com）
            WebDriverWait(self.driver, 25).until(
                lambda d: ("buy.tmall.com" in d.current_url) or ("buy.taobao.com" in d.current_url) or ("trade.taobao.com" in d.current_url)
            )
            if self.debug_log_var.get():
                self.log(f"已进入订单确认页: {self.driver.current_url}")

            # 2.1) 可选：若存在“同意/阅读协议”勾选框，则勾选
            try:
                agree = self.driver.find_elements(By.XPATH, "//label[contains(.,'同意') or contains(.,'已阅读')]/preceding-sibling::input[@type='checkbox']")
                for cb in agree:
                    if cb.is_displayed() and not cb.is_selected():
                        self.driver.execute_script("arguments[0].click();", cb)
            except Exception:
                pass

            # 2.2) 可选：若地址未选中，点击第一个地址卡片
            try:
                addr = self.driver.find_elements(By.CSS_SELECTOR, ".address, .addr-item, .J_Addr, .J_AddrItem")
                if addr:
                    first = addr[0]
                    if first and first.is_displayed():
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", first)
                        self.driver.execute_script("arguments[0].click();", first)
            except Exception:
                pass

            # 3) 提交订单
            submit_selectors = [
                ("css", "#submitOrderPC_1, #submitOrder_1, .go-btn, .btn-order, .place-order .btn, button[type='submit']"),
                ("xpath", "//a[contains(., '提交订单')]|//button[contains(., '提交订单')]"),
            ]
            submit_clicked = False
            for how, expr in submit_selectors:
                try:
                    if how == "css":
                        btn = WebDriverWait(self.driver, 8).until(lambda d: d.find_element(By.CSS_SELECTOR, expr))
                    else:
                        btn = WebDriverWait(self.driver, 8).until(lambda d: d.find_element(By.XPATH, expr))
                    if btn and btn.is_displayed() and btn.is_enabled():
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                        btn.click()
                        submit_clicked = True
                        break
                except Exception:
                    continue
            if not submit_clicked:
                # 兜底：JS 点击包含“提交订单”的任意元素
                try:
                    btn = self.driver.find_element(By.XPATH, "//*[contains(., '提交订单')]")
                    self.driver.execute_script("arguments[0].click();", btn)
                    submit_clicked = True
                except Exception:
                    pass

            if submit_clicked:
                self.log("已尝试提交订单。如有支付弹窗/需确认短信等，请手动完成。")
                return True
            else:
                self.log("未找到提交订单按钮，可能需完善选择器或未完成必填信息。")
                return False
        except Exception as e:
            self.log(f"自动下单异常: {e}")
            return False

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

        last_full_refresh = time.time() # 记录上一次整页刷新时间
        while self.running:
            try:
                now = time.time()
                # 尝试在不整页刷新的情况下检测一次
                WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

                if keywords:
                    self.select_sku_by_keywords(keywords)

                if self.is_in_stock():
                    self.update_status("状态：发现有货！", "red")
                    self.log("检测到可下单/可加入购物车")
                    if bool(self.auto_buy_var.get()):
                        ok = self.auto_submit_order(keywords)
                        if ok:
                            self.beep()
                            self.root.after(0, lambda: messagebox.showinfo("提醒", "已自动尝试下单，若有支付验证请手动完成！"))
                        else:
                            self.beep()
                            self.root.after(0, lambda: messagebox.showinfo("提醒", "发现库存但自动下单失败，请手动下单！"))
                    else:
                        self.beep()
                        self.root.after(0, lambda: messagebox.showinfo("提醒", "发现有货！请尽快下单！"))
                    found_stock = True
                    break
                else:
                    self.update_status("状态：暂无库存，继续监控...", "blue")
                    if self.debug_log_var.get():
                        self.log("本轮未发现库存")

                # 到达最短整页刷新时间，则刷新一次
                refresh_min = float(self.refresh_min_entry.get().strip() or 2)
                if (now - last_full_refresh) + 1e-6 >= (refresh_min + (self.error_backoff_s or 0)):
                    if self.debug_log_var.get():
                        self.log("达到最短整页刷新间隔，执行刷新")
                    self.driver.refresh()
                    last_full_refresh = time.time()
                    # 刷新后等待页面结构
                    WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

            except Exception as e:
                self.update_status("状态：可能未登录/页面异常，已退避重试", "orange")
                self.log(f"异常：{e}")
                self.error_backoff_s = min(60, max(1, (self.error_backoff_s or 1) * 2))

            # 0.1 秒级轮询节奏
            base_interval = float(self.interval_entry.get().strip() or 0.1)
            wait_s = max(0.05, base_interval)
            time.sleep(wait_s)

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
            interval = float(self.interval_entry.get())
            if interval <= 0:
                raise ValueError
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

    def build_chrome_options(self):
        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)
        if self.chrome_path_entry.get().strip():
            options.binary_location = self.chrome_path_entry.get().strip()
        if self.profile_entry.get().strip():
            options.add_argument(f"--user-data-dir={self.profile_entry.get().strip()}")
        # 可选：降低被自动化识别的概率
        options.add_experimental_option("excludeSwitches", ["enable-automation"])  # type: ignore
        options.add_experimental_option("useAutomationExtension", False)  # type: ignore
        options.add_argument("--disable-blink-features=AutomationControlled")
        return options

    def ensure_driver(self):
        if self.driver is None:
            self.driver = webdriver.Chrome(options=self.build_chrome_options())

    def normalize_product_url(self, url: str) -> str:
        """将移动端/带多余参数的链接转换为标准 PC 商品链接"""
        try:
            parsed = urlparse(url)
            host = (parsed.netloc or "").lower()
            qs = parse_qs(parsed.query)
            item_id = None
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


if __name__ == "__main__":
    root = tk.Tk()
    app = TaobaoMonitorApp(root)
    root.mainloop()
