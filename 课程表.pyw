# -*- coding: utf-8 -*-
"""我的课程表 —— 桌面课程/考试/日程管理,带 Windows 系统通知提醒。

功能:
  · 课程表:周视图时间网格,支持周次范围(如 1-16 周),自动高亮今天、显示当前时间线
  · 考试安排:列表管理,显示倒计时
  · 我的日程:自定义事项,每条可单独设置提前提醒时间
  · 提醒:到期前通过 Windows 系统通知(右下角 Toast)提醒,每条只提醒一次
  · 设置:默认提前量、学期第一周、开机自启、测试通知

数据保存在 %APPDATA%\\课程表\\data.json,可直接编辑或备份。
"""

import base64
import datetime as dt
import json
import logging
import math
import os
import queue
import re
import socket
import subprocess
import sys
import threading
import time
import traceback
import uuid
import winreg
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# ---------------------------------------------------------------- 基础配置

APP_NAME = "课程表"
APP_VERSION = "1.0.0"
APP_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), APP_NAME)
DATA_FILE = os.path.join(APP_DIR, "data.json")
LOG_FILE = os.path.join(APP_DIR, "app.log")

SINGLE_INSTANCE_PORT = 59231
CHECK_INTERVAL = 15          # 提醒扫描间隔(秒)

DAY_NAMES = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]

# 学校作息:节次(数字, 起止时间);午休/晚休作为独立时间带插入
PERIODS = [
    (1, "08:00", "08:45"), (2, "08:55", "09:40"),
    (3, "10:00", "10:45"), (4, "10:55", "11:40"),
    (5, "14:00", "14:45"), (6, "14:55", "15:40"),
    (7, "16:00", "16:45"), (8, "16:55", "17:40"),
    (9, "19:00", "19:45"), (10, "19:55", "20:40"),
    (11, "21:00", "21:45"), (12, "21:55", "22:40"),
]
BREAK_AFTER = {4: "午休", 8: "晚休"}
PERIOD_ROW_H = 50          # 每节课行高(逻辑像素)
BREAK_H = 24               # 休息带高度(逻辑像素)

COURSE_COLORS = [
    "#4F86F7", "#00A98F", "#E9615E", "#8E6FC1", "#E8912D",
    "#16A085", "#3B7DD8", "#D35FB7", "#5D8A00", "#C2543B",
    "#2C7A7B", "#B8437F",
]

# ---- Apple 风格配色 ----
BG = "#F5F5F7"       # 页面底色
CARD = "#FFFFFF"     # 卡片
ACCENT = "#0071E3"   # 主题蓝
TXT = "#1D1D1F"      # 主文字
TXT2 = "#86868B"     # 次要文字
LINE = "#F0F0F2"     # 网格细线
LINE2 = "#E5E5EA"    # 网格粗线


def blend_hex(hex1, hex2, t):
    """把 hex1 向 hex2 混合 t(0~1),用于生成浅色底/深色字。"""
    a = [int(hex1[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(hex2[i:i + 2], 16) for i in (1, 3, 5)]
    return "#%02X%02X%02X" % tuple(round(a[i] * (1 - t) + b[i] * t) for i in range(3))


def card_colors(base):
    """课程块的苹果风配色:(浅底, 主文字, 次文字)。"""
    return (blend_hex(base, "#FFFFFF", 0.84),
            blend_hex(base, "#000000", 0.45),
            blend_hex(base, "#000000", 0.22))

os.makedirs(APP_DIR, exist_ok=True)
logging.basicConfig(
    filename=LOG_FILE, level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s", encoding="utf-8",
)


def log_exc(tag):
    logging.error("%s\n%s", tag, traceback.format_exc())


# ---------------------------------------------------------------- 数据层

DEFAULT_DATA = {
    "courses": [
        {"id": "c1", "name": "高等数学", "weekday": 0, "start": "08:00", "end": "09:40",
         "location": "教三-301", "teacher": "", "weeks": "1-16", "color": 0, "lead": None},
        {"id": "c2", "name": "大学英语", "weekday": 2, "start": "10:00", "end": "11:40",
         "location": "文B-202", "teacher": "", "weeks": "1-16", "color": 1, "lead": None},
    ],
    "exams": [
        {"id": "e1", "title": "高等数学期中考试", "date": "2026-11-12", "time": "14:00",
         "location": "教一-101", "lead": 60, "note": "带学生证、计算器"},
    ],
    "events": [
        {"id": "v1", "title": "班会", "date": "2026-10-26", "time": "19:00",
         "location": "", "lead": 30, "note": "示例日程,可删除"},
    ],
    "settings": {"default_lead": 15, "reminder_on": True,
                 "semester_start": "2026-09-07", "autostart": False},
    "fired": [],
}

_lock = threading.Lock()
_data = None


def load_data():
    global _data
    try:
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            _data = json.load(f)
    except Exception:
        logging.warning("读取数据失败,使用默认数据")
        _data = None
    if not isinstance(_data, dict):
        _data = {}
    for k, v in DEFAULT_DATA.items():
        _data.setdefault(k, json.loads(json.dumps(v)) if k != "fired" else [])
    _data["settings"] = {**DEFAULT_DATA["settings"], **_data.get("settings", {})}
    return _data


def save_data():
    with _lock:
        try:
            tmp = DATA_FILE + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(_data, f, ensure_ascii=False, indent=2)
            os.replace(tmp, DATA_FILE)
        except Exception:
            log_exc("保存数据失败")


def new_id(prefix):
    return f"{prefix}{uuid.uuid4().hex[:8]}"


def next_color_index():
    used = {c.get("color", 0) for c in _data["courses"]}
    for i in range(len(COURSE_COLORS)):
        if i not in used:
            return i
    return len(_data["courses"]) % len(COURSE_COLORS)


# ---------------------------------------------------------------- 工具函数

def parse_hhmm(s, default=None):
    """'HH:MM' -> 分钟数;非法返回 default。"""
    m = re.fullmatch(r"\s*(\d{1,2})[:：点h](\d{1,2})?\s*", str(s or ""))
    if not m:
        return default
    h, mi = int(m.group(1)), int(m.group(2) or 0)
    if h > 23 or mi > 59:
        return default
    return h * 60 + mi


def fmt_minutes(m):
    return f"{m // 60:02d}:{m % 60:02d}"


def parse_date(s):
    try:
        return dt.datetime.strptime(str(s or "").strip(), "%Y-%m-%d").date()
    except ValueError:
        return None


def semester_week(day, semester_start_str):
    """返回 day 所在的教学周(1 起);未设置或早于第一周返回 None。"""
    start = parse_date(semester_start_str)
    if not start:
        return None
    start -= dt.timedelta(days=start.weekday())   # 对齐到第一周的周一
    delta = (day - start).days
    if delta < 0:
        return None
    return delta // 7 + 1


def parse_weeks(s):
    """'1-16' / '3' / '' -> (a, b) 或 None(每周)。"""
    s = str(s or "").strip()
    if not s:
        return None
    m = re.fullmatch(r"(\d+)\s*[-~]\s*(\d+)", s)
    if m:
        return int(m.group(1)), int(m.group(2))
    if re.fullmatch(r"\d+", s):
        return int(s), int(s)
    return None


def week_in_range(week, weeks_str):
    if week is None:
        return True
    rng = parse_weeks(weeks_str)
    return True if rng is None else rng[0] <= week <= rng[1]


def parse_int_in(s, lo, hi, err):
    if not s:
        return None
    try:
        v = int(s)
    except ValueError:
        raise ValueError(err)
    if not (lo <= v <= hi):
        raise ValueError(err)
    return v


# ---------------------------------------------------------------- Windows 通知

TOAST_AUMID = r"{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe"

TOAST_PS_TEMPLATE = r'''
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
$xmlText = @"
<toast scenario="reminder">
  <visual><binding template="ToastGeneric">
    <text>__TITLE__</text>
    <text>__BODY__</text>
  </binding></visual>
  <audio src="ms-winsoundevent:Notification.Default"/>
</toast>
"@
$doc = New-Object Windows.Data.Xml.Dom.XmlDocument
$doc.LoadXml($xmlText)
$toast = New-Object Windows.UI.Notifications.ToastNotification($doc)
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("__AUMID__").Show($toast)
'''.replace("__AUMID__", TOAST_AUMID)


def xml_escape(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace("'", "&apos;").replace('"', "&quot;"))


def show_toast(title, body):
    """通过 PowerShell 发送 Windows 系统通知(不阻塞、不闪黑框)。"""
    try:
        ps = TOAST_PS_TEMPLATE.replace("__TITLE__", xml_escape(title)).replace("__BODY__", xml_escape(body))
        enc = base64.b64encode(ps.encode("utf-16-le")).decode("ascii")
        subprocess.Popen(
            ["powershell", "-NoProfile", "-EncodedCommand", enc],
            creationflags=0x08000000,  # CREATE_NO_WINDOW
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        logging.info("通知已发送: %s | %s", title, body)
    except Exception:
        log_exc("发送通知失败")


# ---------------------------------------------------------------- 到期计算(提醒引擎与状态栏共用)

def compute_due_items(now):
    """返回 [(key, occ_datetime, lead, title, body)],key 用于去重。"""
    out = []
    today = now.date()
    default_lead = int(_data["settings"].get("default_lead", 15) or 15)
    start_str = _data["settings"].get("semester_start")

    for c in _data["courses"]:
        st = parse_hhmm(c.get("start"))
        if st is None or not (0 <= c.get("weekday", 0) <= 6):
            continue
        days_ahead = (c.get("weekday", 0) - today.weekday()) % 7
        occ_date = today + dt.timedelta(days=days_ahead)
        if not week_in_range(semester_week(occ_date, start_str), c.get("weeks")):
            continue
        occ = dt.datetime.combine(occ_date, dt.time(st // 60, st % 60))
        if now >= occ:
            continue
        lead = int(c.get("lead") or default_lead)
        key = f"course:{c['id']}:{occ_date.isoformat()}"
        body = f"{fmt_minutes(st)} {c.get('location', '')} {c.get('teacher', '')}".strip()
        out.append((key, occ, lead, c["name"], body))

    for kind, items in (("exam", _data["exams"]), ("event", _data["events"])):
        for it in items:
            d = parse_date(it.get("date"))
            t = parse_hhmm(it.get("time"))
            if not d or t is None:
                continue
            occ = dt.datetime.combine(d, dt.time(t // 60, t % 60))
            if now >= occ:
                continue
            lead = int(it.get("lead") or default_lead)
            key = f"{kind}:{it['id']}:{d.isoformat()}"
            loc = f" {it['location']}" if it.get("location") else ""
            out.append((key, occ, lead, it["title"], f"{it['time']}{loc}".strip()))
    return out


class ReminderEngine(threading.Thread):
    """定时扫描:课程(每周循环)、考试、日程,到点发 Windows 系统通知。"""

    def __init__(self, ui_queue):
        super().__init__(daemon=True)
        self.ui_queue = ui_queue

    def run(self):
        time.sleep(3)   # 等 UI 初始化
        while True:
            try:
                self.scan()
            except Exception:
                log_exc("提醒扫描出错")
            time.sleep(CHECK_INTERVAL)

    def scan(self):
        now = dt.datetime.now()
        fired = set(_data.get("fired", []))
        today_iso = now.date().isoformat()
        fired = {k for k in fired if k.rsplit(":", 1)[-1] >= today_iso}   # 清理过期记录
        newly = []
        remind_on = _data["settings"].get("reminder_on", True)

        for key, occ, lead, title, body in compute_due_items(now):
            if key in fired:
                continue
            remind_at = occ - dt.timedelta(minutes=lead)
            if remind_at <= now < occ:
                if remind_on:
                    show_toast(f"⏰ {lead}分钟后:{title}", body)
                fired.add(key)
                newly.append(key)
            elif now >= occ:
                fired.add(key)   # 已错过,静默标记,不再提醒
                newly.append(key)
        if newly:
            with _lock:
                _data["fired"] = sorted(fired)
            save_data()
        self.ui_queue.put(("tick", None))


# ---------------------------------------------------------------- 对话框

class FormDialog(tk.Toplevel):
    """简单表单对话框:子类实现 build_form(),返回 {字段名: 控件}。"""

    def __init__(self, parent, title, on_ok):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.on_ok = on_ok
        frm = ttk.Frame(self, padding=14)
        frm.pack(fill="both", expand=True)
        self.w = self.build_form(frm)
        btns = ttk.Frame(frm)
        btns.grid(row=100, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="取消", command=self.destroy).pack(side="right", padx=4)
        ttk.Button(btns, text="保存", command=self.do_ok).pack(side="right")
        self.bind("<Return>", lambda e: self.do_ok())
        self.bind("<Escape>", lambda e: self.destroy())
        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 3
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        self.wait_window()

    def build_form(self, frm):
        raise NotImplementedError

    def add_row(self, frm, row, label, widget):
        ttk.Label(frm, text=label).grid(row=row, column=0, sticky="ne", pady=4, padx=(0, 8))
        widget.grid(row=row, column=1, sticky="we", pady=4)
        return widget

    def do_ok(self):
        try:
            values = self.collect()
        except ValueError as e:
            messagebox.showwarning("填写有误", str(e), parent=self)
            return
        self.on_ok(values)
        self.destroy()


class LeadField:
    @staticmethod
    def make(parent, lead):
        row = ttk.Frame(parent)
        sp = ttk.Spinbox(row, from_=0, to=10080, width=6, increment=5)
        sp.set("" if lead is None else lead)
        ttk.Label(row, text="分钟前提醒(留空=用默认)").pack(side="left")
        sp.pack(side="left", padx=(0, 6))
        return row, sp


class CourseDialog(FormDialog):
    def __init__(self, parent, on_ok, course=None):
        self.course = course or {}
        super().__init__(parent, "编辑课程" if course else "添加课程", on_ok)

    def build_form(self, frm):
        c = self.course
        frm.columnconfigure(1, weight=1)
        name = self.add_row(frm, 0, "课程名称", ttk.Entry(frm, width=28))
        name.insert(0, c.get("name", ""))
        wd = self.add_row(frm, 1, "星期", ttk.Combobox(frm, state="readonly", width=25, values=DAY_NAMES))
        wd.current(c.get("weekday", 0))
        st = self.add_row(frm, 2, "开始时间", ttk.Entry(frm, width=28))
        st.insert(0, c.get("start", "08:00"))
        et = self.add_row(frm, 3, "结束时间", ttk.Entry(frm, width=28))
        et.insert(0, c.get("end", "09:40"))
        loc = self.add_row(frm, 4, "上课地点", ttk.Entry(frm, width=28))
        loc.insert(0, c.get("location", ""))
        tea = self.add_row(frm, 5, "教师(可选)", ttk.Entry(frm, width=28))
        tea.insert(0, c.get("teacher", ""))
        wk = self.add_row(frm, 6, "周次(可选)", ttk.Entry(frm, width=28))
        wk.insert(0, c.get("weeks", ""))
        ttk.Label(frm, text="留空=每周都有;填 1-16 表示只在第 1~16 周上课", foreground="#888")\
            .grid(row=7, column=1, sticky="w")
        lead_row, lead = LeadField.make(frm, c.get("lead"))
        self.add_row(frm, 8, "提前提醒", lead_row)
        self.w = dict(name=name, weekday=wd, start=st, end=et,
                      location=loc, teacher=tea, weeks=wk, lead=lead)
        return self.w

    def collect(self):
        w = self.w
        name = w["name"].get().strip()
        st = parse_hhmm(w["start"].get())
        et = parse_hhmm(w["end"].get())
        if not name:
            raise ValueError("请填写课程名称")
        if st is None:
            raise ValueError("开始时间格式应为 HH:MM,例如 08:00")
        if et is None:
            raise ValueError("结束时间格式应为 HH:MM,例如 09:40")
        if et <= st:
            raise ValueError("结束时间应晚于开始时间")
        lead = parse_int_in(w["lead"].get().strip(), 0, 720, "提前提醒应为 0~720 的整数分钟")
        return {
            "id": self.course.get("id") or new_id("c"),
            "name": name, "weekday": w["weekday"].current(),
            "start": fmt_minutes(st), "end": fmt_minutes(et),
            "location": w["location"].get().strip(), "teacher": w["teacher"].get().strip(),
            "weeks": w["weeks"].get().strip(), "color": self.course.get("color", next_color_index()),
            "lead": lead,
        }


class DatedDialog(FormDialog):
    """考试 / 日程共用的表单。"""

    LABELS = {"exam": ("考试名称", "考试日期"), "event": ("事项名称", "日期")}

    def __init__(self, parent, kind, on_ok, item=None):
        self.kind = kind
        self.item = item or {}
        noun = "考试" if kind == "exam" else "日程"
        super().__init__(parent, ("编辑" if item else "添加") + noun, on_ok)

    def build_form(self, frm):
        it, k = self.item, self.kind
        frm.columnconfigure(1, weight=1)
        name_l, date_l = self.LABELS[k]
        name = self.add_row(frm, 0, name_l, ttk.Entry(frm, width=28))
        name.insert(0, it.get("title", ""))
        date = self.add_row(frm, 1, date_l + "(YYYY-MM-DD)", ttk.Entry(frm, width=28))
        date.insert(0, it.get("date", dt.date.today().isoformat()))
        tm = self.add_row(frm, 2, "开始时间", ttk.Entry(frm, width=28))
        tm.insert(0, it.get("time", "09:00"))
        loc = self.add_row(frm, 3, "地点(可选)", ttk.Entry(frm, width=28))
        loc.insert(0, it.get("location", ""))
        lead_row, lead = LeadField.make(frm, it.get("lead"))
        self.add_row(frm, 4, "提前提醒", lead_row)
        note = self.add_row(frm, 5, "备注(可选)", tk.Text(frm, width=28, height=3))
        note.insert("1.0", it.get("note", ""))
        self.w = dict(name=name, date=date, time=tm, location=loc, lead=lead, note=note)
        return self.w

    def collect(self):
        w = self.w
        name = w["name"].get().strip()
        d = parse_date(w["date"].get())
        t = parse_hhmm(w["time"].get())
        if not name:
            raise ValueError("请填写名称")
        if not d:
            raise ValueError("日期格式应为 YYYY-MM-DD,例如 2026-11-12")
        if t is None:
            raise ValueError("时间格式应为 HH:MM,例如 14:00")
        lead = parse_int_in(w["lead"].get().strip(), 0, 10080, "提前提醒应为 0~10080 的整数分钟(最多 7 天)")
        return {
            "id": self.item.get("id") or new_id("e" if self.kind == "exam" else "v"),
            "title": name, "date": d.isoformat(), "time": fmt_minutes(t),
            "location": w["location"].get().strip(), "lead": lead,
            "note": w["note"].get("1.0", "end").strip(),
        }


# ---------------------------------------------------------------- Apple 风控件

class RoundedButton(tk.Canvas):
    """胶囊按钮:primary 蓝底白字,secondary 浅灰底深字,悬停微变色。"""

    PALETTE = {"primary": (ACCENT, "#FFFFFF", blend_hex(ACCENT, "#000000", 0.12)),
               "secondary": ("#E8E8ED", TXT, "#DEDEE3")}

    def __init__(self, parent, text, command=None, kind="secondary", padh=20):
        self.text, self.command = text, command
        self.normal, self.fg, self.hover = self.PALETTE[kind]
        fpx = parent.winfo_fpixels("1i") / 72 * 10
        tw = sum(fpx * (1.0 if ord(ch) > 0x2E80 else 0.56) for ch in text)
        h = int(fpx * 1.9)
        w = int(tw + padh * 2)
        try:
            bgc = parent["bg"]
        except tk.TclError:
            bgc = BG
        super().__init__(parent, width=w, height=h, bg=bgc, highlightthickness=0,
                         cursor="hand2")
        self.bind("<Button-1>", self._click)
        self.bind("<Enter>", lambda e: self._paint(self.hover))
        self.bind("<Leave>", lambda e: self._paint(self.normal))
        self._paint(self.normal)

    def _paint(self, bgc):
        self.delete("all")
        w, h = self.winfo_reqwidth(), self.winfo_reqheight()
        self.create_polygon(*_rrect_pts(0, 0, w, h, h / 2), smooth=True, fill=bgc, outline="")
        self.create_text(w / 2, h / 2, text=self.text, fill=self.fg,
                         font=("Microsoft YaHei UI", 10))

    def _click(self, e):
        if self.command:
            self.command()


def _rrect_pts(x0, y0, x1, y1, r):
    return [x0 + r, y0, x1 - r, y0, x1, y0, x1, y0 + r, x1, y1 - r, x1, y1,
            x1 - r, y1, x0 + r, y1, x0, y1, x0, y1 - r, x0, y0 + r, x0, y0]


class SegmentedBar(tk.Canvas):
    """Apple 分段控件:灰色圆角轨道,选中段为白色浮起圆角块。"""

    def __init__(self, parent, tabs, on_change):
        self.tabs, self.on_change = tabs, on_change
        self.selected = tabs[0][0]
        try:
            bgc = parent["bg"]
        except tk.TclError:
            bgc = BG
        fpx = parent.winfo_fpixels("1i") / 72 * 10
        widest = max(sum(fpx * (1.0 if ord(ch) > 0x2E80 else 0.56) for ch in label)
                     for _, label in tabs)
        self.seg_w = max(int(widest + 48), 108)
        self.h = int(fpx * 2.1)
        total = self.seg_w * len(tabs) + 8
        super().__init__(parent, width=total, height=self.h, bg=bgc,
                         highlightthickness=0, cursor="hand2")
        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Button-1>", self._click)
        self._draw()

    def select(self, key, notify=False):
        if key != self.selected:
            self.selected = key
            self._draw()
            if notify:
                self.on_change(key)

    def _draw(self):
        self.delete("all")
        w, h = self.winfo_reqwidth(), self.winfo_reqheight()
        self.create_polygon(*_rrect_pts(0, 0, w, h, h / 2), smooth=True,
                            fill="#E9E9EB", outline="")
        for i, (key, label) in enumerate(self.tabs):
            x0, x1 = 4 + i * self.seg_w, 4 + (i + 1) * self.seg_w
            if key == self.selected:
                self.create_polygon(*_rrect_pts(x0 + 2, 3, x1 - 2, h - 3, (h - 6) / 2),
                                    smooth=True, fill="#FFFFFF", outline="")
            self.create_text((x0 + x1) / 2, h / 2, text=label,
                             fill=TXT if key == self.selected else "#6E6E73",
                             font=("Microsoft YaHei UI", 10,
                                   "bold" if key == self.selected else "normal"))

    def _click(self, e):
        i = int((e.x - 4) // self.seg_w)
        if 0 <= i < len(self.tabs):
            self.select(self.tabs[i][0], notify=True)


# ---------------------------------------------------------------- 主界面

class App(tk.Tk):
    AUTOSTART_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
    AUTOSTART_NAME = "我的课程表"

    def __init__(self, ui_queue):
        super().__init__()
        self.ui_queue = ui_queue
        self.week_offset = 0
        self.S = system_scale_factor()          # 屏幕缩放系数(高 DPI 屏上界面等比放大)
        self.row_h = int(PERIOD_ROW_H * self.S)  # 每节课行高(随缩放)
        self.title("我的课程表")
        icon = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.ico")
        if os.path.exists(icon):
            try:
                self.iconbitmap(icon)
            except Exception:
                pass
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w = min(int(1180 * self.S), int(sw * 0.95))
        h = min(int(760 * self.S), int(sh * 0.92))
        self.geometry(f"{w}x{h}+80+50")
        self.minsize(int(900 * self.S), int(560 * self.S))
        self._build_style()
        self._build_ui()
        self.after(500, self._poll_queue)
        self.after(60_000, self._refresh_every_min)
        self.refresh_all()

    # ---- 样式 ----
    def _build_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except Exception:
            pass
        self.option_add("*Font", ("Microsoft YaHei UI", 10))
        style.configure("Treeview", background=CARD, fieldbackground=CARD,
                        borderwidth=0, rowheight=int(34 * self.S),
                        font=("Microsoft YaHei UI", 10))
        style.configure("Treeview.Heading", background=BG, foreground="#6E6E73",
                        font=("Microsoft YaHei UI", 10, "bold"), relief="flat",
                        padding=(8, 8))
        style.map("Treeview.Heading", background=[("active", BG)])
        style.map("Treeview", background=[("selected", "#E4EEFB")],
                  foreground=[("selected", TXT)])
        style.configure("TSpinbox", arrowcolor=TXT)
        style.configure("Vertical.TScrollbar", background="#D9D9DE", troughcolor=BG,
                        bordercolor=BG, arrowcolor="#86868B")

    # ---- 界面骨架:顶部分段控件 + 页面栈 ----
    def _build_ui(self):
        container = tk.Frame(self, bg=BG)
        container.pack(fill="both", expand=True)

        bar_holder = tk.Frame(container, bg=BG)
        bar_holder.pack(fill="x", pady=(12, 4))
        self.segbar = SegmentedBar(bar_holder,
                                   [("schedule", "课程表"), ("exams", "考试安排"),
                                    ("events", "我的日程"), ("settings", "设置")],
                                   self._switch_page)
        self.segbar.pack(anchor="center")

        page_holder = tk.Frame(container, bg=BG)
        page_holder.pack(fill="both", expand=True)
        self.pages = {}
        for key, builder in (("schedule", self._build_schedule_page),
                             ("exams", self._build_exams_page),
                             ("events", self._build_events_page),
                             ("settings", self._build_settings_page)):
            page = tk.Frame(page_holder, bg=BG)
            page.place(relwidth=1, relheight=1)
            builder(page)
            self.pages[key] = page
        self.pages["schedule"].tkraise()

        self.status = tk.Label(self, anchor="w", padx=14, pady=5, bg=BG, fg=TXT2,
                               font=("Microsoft YaHei UI", 9))
        self.status.pack(fill="x", side="bottom")

    def _switch_page(self, key):
        self.pages[key].tkraise()
        if key == "schedule":
            self.after(30, self.draw_schedule)

    # ---- 页1 课程表 ----
    def _build_schedule_page(self, parent):
        bar = tk.Frame(parent, bg=BG)
        bar.pack(fill="x", padx=16, pady=(8, 8))
        RoundedButton(bar, "＋ 添加课程", kind="primary", command=self.add_course).pack(side="left")
        RoundedButton(bar, "上一周", command=lambda: self.shift_week(-1)).pack(side="left", padx=(16, 6))
        RoundedButton(bar, "本周", command=lambda: self.shift_week(0)).pack(side="left", padx=(0, 6))
        RoundedButton(bar, "下一周", command=lambda: self.shift_week(1)).pack(side="left")
        self.week_label = tk.Label(bar, bg=BG, fg=TXT2, font=("Microsoft YaHei UI", 10))
        self.week_label.pack(side="right")

        holder = tk.Frame(parent, bg=BG)
        holder.pack(fill="both", expand=True, padx=16, pady=(0, 10))
        self.canvas = tk.Canvas(holder, bg=BG, highlightthickness=0)
        vsb = ttk.Scrollbar(holder, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<Configure>", lambda e: self.draw_schedule())
        self.canvas.bind("<MouseWheel>", self._on_wheel)
        self.canvas.bind("<Button-1>", self._on_canvas_click)

    def _on_wheel(self, e):
        self.canvas.yview_scroll(-1 if e.delta > 0 else 1, "units")

    def shift_week(self, d):
        self.week_offset = 0 if d == 0 else self.week_offset + d
        self.draw_schedule()

    def monday_of_view(self):
        today = dt.date.today()
        monday = today - dt.timedelta(days=today.weekday())
        return monday + dt.timedelta(weeks=self.week_offset)

    def draw_schedule(self):
        cv = self.canvas
        cv.delete("all")
        w = cv.winfo_width()
        if w < 50:
            return
        S, period_h = self.S, self.row_h
        break_h = int(BREAK_H * S)
        left_w, top_h = int(64 * S), int(52 * S)
        card_r, margin = int(16 * S), int(6 * S)
        card_x1 = w - margin
        col_w = (card_x1 - left_w) / 7
        monday = self.monday_of_view()
        today = dt.date.today()
        week = semester_week(monday, _data["settings"].get("semester_start"))

        self.week_label.config(text=(
            f"第 {week} 教学周 · {monday.isoformat()} ~ {(monday + dt.timedelta(days=6)).isoformat()}"
            + ("(本周)" if self.week_offset == 0 else "")))

        def rrect(x0, y0, x1, y1, r, **kw):
            return cv.create_polygon(*_rrect_pts(x0, y0, x1, y1, r), smooth=True, **kw)

        # 先做布局:各节次行的 y 与当前时间线定位段
        ys = {}
        seg = []          # (开始分钟, 结束分钟, y0, y1)
        y = top_h
        for idx, (num, ps, pe) in enumerate(PERIODS):
            ps_m, pe_m = parse_hhmm(ps), parse_hhmm(pe)
            ys[num] = y
            seg.append((ps_m, pe_m, y, y + period_h))
            y += period_h
            if num in BREAK_AFTER and idx + 1 < len(PERIODS):
                seg.append((pe_m, parse_hhmm(PERIODS[idx + 1][1]), y, y + break_h))
                y += break_h
        grid_h = y - top_h

        # 白色大圆角卡片承载整个网格(从日历列左缘开始,不遮挡左侧时间轴)
        rrect(left_w, int(2 * S), card_x1, top_h + grid_h + int(10 * S),
              r=card_r, fill=CARD, outline="")

        # 日期表头:今天为主题蓝药丸,其余为纯文字
        for i, dn in enumerate(DAY_NAMES):
            d = monday + dt.timedelta(days=i)
            x0 = left_w + i * col_w
            is_today = (d == today and self.week_offset == 0)
            pad = int(5 * S)
            if is_today:
                rrect(x0 + pad, int(7 * S), x0 + col_w - pad, top_h - int(7 * S),
                      r=int(12 * S), fill=ACCENT, outline="")
            cv.create_text(x0 + col_w / 2, top_h / 2 + int(1 * S), text=f"{dn} {d.month}/{d.day}",
                           fill="#FFFFFF" if is_today else TXT,
                           font=("Microsoft YaHei UI", 10, "bold"))

        # 左轴:节号 + 起止时间(与卡片保持间隙)
        for num, ps, pe in PERIODS:
            y0 = ys[num]
            ax = left_w - int(12 * S)
            cv.create_text(ax, y0 + int(5 * S), text=str(num), anchor="ne",
                           fill=TXT, font=("Microsoft YaHei UI", 10, "bold"))
            cv.create_text(ax, y0 + int(25 * S), text=ps, anchor="ne",
                           fill=TXT2, font=("Microsoft YaHei UI", 7))
            cv.create_text(ax, y0 + int(37 * S), text=pe, anchor="ne",
                           fill=TXT2, font=("Microsoft YaHei UI", 7))

        # 节次分隔线与午休/晚休带
        for idx, (num, ps, pe) in enumerate(PERIODS):
            y0 = ys[num]
            if num in BREAK_AFTER:
                yb = y0 + period_h
                rrect(left_w + int(8 * S), yb + int(3 * S), card_x1 - int(8 * S),
                      yb + break_h - int(3 * S), r=int(9 * S), fill="#F2F2F4", outline="")
                cv.create_text((left_w + card_x1) / 2, yb + break_h / 2, text=BREAK_AFTER[num],
                               fill=TXT2, font=("Microsoft YaHei UI", 8))
            elif idx + 1 < len(PERIODS):
                cv.create_line(left_w, y0 + period_h, card_x1, y0 + period_h, fill=LINE)

        for i in range(8):
            x = left_w + i * col_w
            cv.create_line(x, top_h, x, top_h + grid_h, fill=LINE)

        # 课程块:马卡龙浅底 + 同色系深字,按节次跨度放置
        self._block_map = []
        avail = col_w - 18
        for c in _data["courses"]:
            st = parse_hhmm(c.get("start"))
            et = parse_hhmm(c.get("end"))
            if st is None or et is None or not (0 <= c.get("weekday", 0) <= 6):
                continue
            if not week_in_range(week, c.get("weeks")):
                continue
            covered = [num for num, ps, pe in PERIODS
                       if st < parse_hhmm(pe) and et > parse_hhmm(ps)]
            if not covered:
                continue
            x0 = left_w + c["weekday"] * col_w + int(5 * S)
            x1 = left_w + (c["weekday"] + 1) * col_w - int(5 * S)
            y0 = ys[covered[0]] + int(3 * S)
            y1 = ys[covered[-1]] + period_h - int(3 * S)
            fill, t_main, t_sub = card_colors(COURSE_COLORS[c.get("color", 0) % len(COURSE_COLORS)])
            rrect(x0, y0, x1, y1, r=int(12 * S), fill=fill, outline="")

            # 内容条目:名称加粗,其余小字;放不下时依次省略教师、地点
            entries = [(c["name"], 9, True), (f"{c['start']}–{c['end']}", 8, False)]
            if c.get("location"):
                entries.append((c["location"], 8, False))
            if c.get("teacher"):
                entries.append((c["teacher"], 8, False))

            def nlines(text, pt, bold):
                fpx = pt * S * 4 / 3 * (1.06 if bold else 1.0)
                tw = sum(fpx * (1.0 if ord(ch) > 0x2E80 else 0.56) for ch in text)
                return max(1, math.ceil(tw / avail))

            heights = [nlines(t, p, b) * p * S * 4 / 3 * 1.32 for t, p, b in entries]
            while len(entries) > 2 and sum(heights) > (y1 - y0) - int(10 * S):
                entries.pop()
                heights.pop()
            total = sum(heights)
            ty = y0 + max(int(4 * S), ((y1 - y0) - total) / 2)
            for (text, pt, bold), h in zip(entries, heights):
                cv.create_text((x0 + x1) / 2, ty, text=text, anchor="n", width=avail,
                               fill=t_main if bold else t_sub,
                               font=("Microsoft YaHei UI", pt,
                                     "bold" if bold else "normal"))
                ty += h
            self._block_map.append((x0, y0, x1, y1, c))

        # 当前时间线:按节次/休息带精确插值(红线 + 时间药丸)
        if self.week_offset == 0:
            now = dt.datetime.now()
            mins = now.hour * 60 + now.minute
            for t0, t1, ya, yb in seg:
                if t0 is not None and t1 is not None and t0 <= mins <= t1:
                    yy = ya + (mins - t0) / max(1, t1 - t0) * (yb - ya)
                    x = left_w + today.weekday() * col_w
                    cv.create_line(x + int(2 * S), yy, x + col_w - int(2 * S), yy,
                                   fill="#FF3B30", width=int(2 * S))
                    pw, ph = int(42 * S), int(18 * S)
                    rrect(x + col_w / 2 - pw / 2, yy - ph / 2, x + col_w / 2 + pw / 2,
                          yy + ph / 2, r=ph / 2, fill="#FF3B30", outline="")
                    cv.create_text(x + col_w / 2, yy, text=now.strftime("%H:%M"),
                                   fill="#FFFFFF", font=("Microsoft YaHei UI", 7, "bold"))
                    break

        cv.configure(scrollregion=(0, 0, w, top_h + grid_h + int(16 * S)))

    def _on_canvas_click(self, e):
        # 后绘制的块在上层,倒序匹配优先命中可见块
        for x0, y0, x1, y1, c in reversed(getattr(self, "_block_map", [])):
            if x0 <= e.x <= x1 and y0 <= e.y <= y1:
                CourseDialog(self, self._save_course, c)
                return

    # ---- 页2 考试 ----
    EXAM_COLS = ("title", "date", "time", "location", "lead", "countdown", "note")
    EXAM_HEADS = ("考试名称", "日期", "时间", "地点", "提前提醒", "倒计时", "备注")
    EXAM_W = (220, 110, 70, 130, 100, 90, 260)

    def _build_exams_page(self, parent):
        bar = tk.Frame(parent, bg=BG)
        bar.pack(fill="x", padx=16, pady=(8, 8))
        RoundedButton(bar, "＋ 添加考试", kind="primary", command=self.add_exam).pack(side="left")
        RoundedButton(bar, "编辑", command=lambda: self.edit_sel("exam")).pack(side="left", padx=(10, 6))
        RoundedButton(bar, "删除", command=lambda: self.del_sel("exam")).pack(side="left")
        self.exam_tree = self._make_tree(parent, self.EXAM_COLS, self.EXAM_HEADS, self.EXAM_W)
        self.exam_tree.bind("<Double-1>", lambda e: self.edit_sel("exam"))

    # ---- 页3 日程 ----
    EVENT_COLS = ("title", "date", "time", "lead", "note")
    EVENT_HEADS = ("事项", "日期", "时间", "提前提醒", "备注")
    EVENT_W = (260, 110, 70, 100, 420)

    def _build_events_page(self, parent):
        bar = tk.Frame(parent, bg=BG)
        bar.pack(fill="x", padx=16, pady=(8, 8))
        RoundedButton(bar, "＋ 添加日程", kind="primary", command=self.add_event).pack(side="left")
        RoundedButton(bar, "编辑", command=lambda: self.edit_sel("event")).pack(side="left", padx=(10, 6))
        RoundedButton(bar, "删除", command=lambda: self.del_sel("event")).pack(side="left")
        self.event_tree = self._make_tree(parent, self.EVENT_COLS, self.EVENT_HEADS, self.EVENT_W)
        self.event_tree.bind("<Double-1>", lambda e: self.edit_sel("event"))

    def _make_tree(self, parent, cols, heads, widths):
        holder = tk.Frame(parent, bg=BG)
        holder.pack(fill="both", expand=True, padx=16, pady=(0, 10))
        card = tk.Frame(holder, bg=CARD)
        card.pack(fill="both", expand=True)
        tree = ttk.Treeview(card, columns=cols, show="headings", selectmode="browse",
                            displaycolumns=cols[1:])
        vsb = ttk.Scrollbar(card, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        for c, h, wd in zip(cols, heads, widths):
            tree.heading(c, text=h)
            tree.column(c, width=wd, anchor="w")
        tree.tag_configure("past", foreground="#A9A9AE")
        tree.tag_configure("soon", foreground="#D70015")
        vsb.pack(side="right", fill="y")
        tree.pack(fill="both", expand=True, padx=2, pady=2)
        return tree

    # ---- 页4 设置 ----
    def _build_settings_page(self, parent):
        f = tk.Frame(parent, bg=BG)
        f.pack(fill="both", expand=True, padx=24, pady=(10, 12))

        card = tk.Frame(f, bg=CARD)
        card.pack(fill="x")
        g = tk.Frame(card, bg=CARD, padx=20, pady=16)
        g.pack(fill="both", expand=True)

        tk.Label(g, text="默认提前提醒(分钟)", bg=CARD, fg=TXT).grid(row=0, column=0, sticky="w", pady=7)
        self.lead_var = tk.StringVar()
        ttk.Spinbox(g, from_=0, to=720, increment=5, width=8, textvariable=self.lead_var)\
            .grid(row=0, column=1, sticky="w", pady=7, padx=(12, 0))

        tk.Label(g, text="学期第一周的周一(YYYY-MM-DD,用于计算教学周和周次过滤)",
                 bg=CARD, fg=TXT).grid(row=1, column=0, sticky="w", pady=7)
        self.start_var = tk.StringVar()
        ttk.Entry(g, width=14, textvariable=self.start_var)\
            .grid(row=1, column=1, sticky="w", pady=7, padx=(12, 0))

        self.remind_var = tk.BooleanVar()
        tk.Checkbutton(g, text="开启提醒(关闭后不再弹系统通知)", variable=self.remind_var,
                       bg=CARD, fg=TXT, activebackground=CARD, activeforeground=TXT,
                       highlightthickness=0, anchor="w")\
            .grid(row=2, column=0, columnspan=2, sticky="w", pady=7)

        self.auto_var = tk.BooleanVar()
        tk.Checkbutton(g, text="开机自动启动(登录 Windows 后自动运行,保证提醒不漏)",
                       variable=self.auto_var, command=self.toggle_autostart,
                       bg=CARD, fg=TXT, activebackground=CARD, activeforeground=TXT,
                       highlightthickness=0, anchor="w")\
            .grid(row=3, column=0, columnspan=2, sticky="w", pady=7)

        btns = tk.Frame(g, bg=CARD)
        btns.grid(row=4, column=0, columnspan=2, sticky="w", pady=(14, 4))
        RoundedButton(btns, "保存设置", kind="primary", command=self.save_settings).pack(side="left")
        RoundedButton(btns, "发送测试通知", command=lambda: show_toast(
            "通知测试成功 ✅", "到课前会在右下角这样提醒你")).pack(side="left", padx=(10, 0))
        RoundedButton(btns, "导出备份", command=self.export_data).pack(side="left", padx=(10, 0))
        RoundedButton(btns, "导入备份", command=self.import_data).pack(side="left", padx=(10, 0))
        RoundedButton(btns, "打开数据文件夹", command=self.open_data_dir).pack(side="left", padx=(10, 0))

        tip = ("使用说明\n"
               "  · 课程表:单击课程块即可编辑;『＋ 添加课程』新增;周次栏填如 1-16 表示只在那些周显示,非本周课程自动隐藏。\n"
               "  · 考试/日程:开始时间前的『提前提醒』分钟数到点后,右下角弹 Windows 系统通知,每条只提醒一次。\n"
               "  · 提醒只在程序运行时有效,建议勾选『开机自动启动』并保持程序开启(可最小化到任务栏)。\n"
               "  · 导出/导入备份:把全部课程、考试、日程、设置打包为一个 JSON 文件,可在另一台电脑恢复。\n"
               f"  · 数据文件:{DATA_FILE}")
        tk.Label(f, text=tip, bg=BG, fg=TXT2, justify="left", font=("Microsoft YaHei UI", 9))\
            .pack(anchor="w", pady=(16, 0))

    def export_data(self):
        path = filedialog.asksaveasfilename(
            parent=self, title="导出备份", defaultextension=".json",
            initialfile=f"课程表备份-{dt.date.today().isoformat()}.json",
            filetypes=[("JSON 文件", "*.json")])
        if not path:
            return
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as src, open(path, "w", encoding="utf-8") as dst:
                dst.write(src.read())
            messagebox.showinfo("已导出", f"数据已导出到:\n{path}", parent=self)
        except Exception:
            log_exc("导出数据失败")
            messagebox.showerror("导出失败", "导出失败,详见日志。", parent=self)

    def import_data(self):
        path = filedialog.askopenfilename(parent=self, title="导入备份",
                                          filetypes=[("JSON 文件", "*.json")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as fp:
                loaded = json.load(fp)
            if not isinstance(loaded, dict):
                raise ValueError
        except Exception:
            messagebox.showwarning("导入失败", "所选文件不是有效的课程表备份。", parent=self)
            return
        if not messagebox.askyesno("导入数据", "导入将覆盖当前全部数据,确定继续吗?", parent=self):
            return
        global _data
        with _lock:
            _data = loaded
        for k, v in DEFAULT_DATA.items():
            _data.setdefault(k, json.loads(json.dumps(v)) if k != "fired" else [])
        _data["settings"] = {**DEFAULT_DATA["settings"], **_data.get("settings", {})}
        save_data()
        self.refresh_all()
        messagebox.showinfo("已导入", "数据导入完成。", parent=self)

    def open_data_dir(self):
        subprocess.Popen(["explorer", APP_DIR])

    # ---- 增删改 ----
    def add_course(self):
        CourseDialog(self, self._save_course)

    def _save_course(self, values):
        for i, c in enumerate(_data["courses"]):
            if c["id"] == values["id"]:
                _data["courses"][i] = values
                break
        else:
            _data["courses"].append(values)
        save_data()
        self.refresh_all()

    def add_exam(self):
        DatedDialog(self, "exam", lambda v: self._save_dated("exams", v))

    def add_event(self):
        DatedDialog(self, "event", lambda v: self._save_dated("events", v))

    def _save_dated(self, key, values):
        items = _data[key]
        for i, it in enumerate(items):
            if it["id"] == values["id"]:
                items[i] = values
                break
        else:
            items.append(values)
        save_data()
        self.refresh_all()

    def _sel_id(self, kind):
        tree = self.exam_tree if kind == "exam" else self.event_tree
        sel = tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先在列表中选中一条记录", parent=self)
            return None
        return str(tree.item(sel[0])["values"][0])

    def edit_sel(self, kind):
        key = "exams" if kind == "exam" else "events"
        iid = self._sel_id(kind)
        if not iid:
            return
        item = next((it for it in _data[key] if it["id"] == iid), None)
        if item:
            DatedDialog(self, kind, lambda v: self._save_dated(key, v), item)

    def del_sel(self, kind):
        key = "exams" if kind == "exam" else "events"
        iid = self._sel_id(kind)
        if not iid:
            return
        item = next((it for it in _data[key] if it["id"] == iid), None)
        if not item:
            return
        if messagebox.askyesno("删除", f"确定删除「{item['title']}」吗?", parent=self):
            _data[key] = [it for it in _data[key] if it["id"] != iid]
            save_data()
            self.refresh_all()

    # ---- 设置 ----
    def load_settings_ui(self):
        s = _data["settings"]
        self.lead_var.set(str(s.get("default_lead", 15)))
        self.start_var.set(s.get("semester_start", ""))
        self.remind_var.set(bool(s.get("reminder_on", True)))
        self.auto_var.set(self.get_autostart())

    def save_settings(self):
        s = _data["settings"]
        try:
            s["default_lead"] = max(0, min(720, int(self.lead_var.get())))
        except ValueError:
            messagebox.showwarning("填写有误", "默认提前提醒应为整数分钟", parent=self)
            return
        d = parse_date(self.start_var.get())
        s["semester_start"] = d.isoformat() if d else ""
        s["reminder_on"] = bool(self.remind_var.get())
        self.set_autostart(bool(self.auto_var.get()))
        save_data()
        self.refresh_all()
        messagebox.showinfo("已保存", "设置已保存。", parent=self)

    def toggle_autostart(self):
        self.set_autostart(bool(self.auto_var.get()))

    @classmethod
    def autostart_command(cls):
        pythonw = sys.executable.replace("python.exe", "pythonw.exe")
        return f'"{pythonw}" "{os.path.abspath(__file__)}"'

    def get_autostart(self):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.AUTOSTART_KEY) as k:
                winreg.QueryValueEx(k, self.AUTOSTART_NAME)
                return True
        except OSError:
            return False

    def set_autostart(self, enable):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.AUTOSTART_KEY, 0, winreg.KEY_SET_VALUE) as k:
                if enable:
                    winreg.SetValueEx(k, self.AUTOSTART_NAME, 0, winreg.REG_SZ, self.autostart_command())
                else:
                    try:
                        winreg.DeleteValue(k, self.AUTOSTART_NAME)
                    except FileNotFoundError:
                        pass
            logging.info("开机自启 = %s", enable)
        except OSError:
            log_exc("设置开机自启失败")

    # ---- 刷新 / 状态栏 ----
    def refresh_all(self):
        self.draw_schedule()
        self.refresh_exams()
        self.refresh_events()
        self.load_settings_ui()
        self.update_status()

    def refresh_exams(self):
        tree = self.exam_tree
        tree.delete(*tree.get_children())
        now = dt.datetime.now()
        items = sorted(_data["exams"], key=lambda it: (it.get("date", ""), it.get("time", "")))
        for it in items:
            d = parse_date(it.get("date"))
            m = parse_hhmm(it.get("time"), 0)
            occ = dt.datetime.combine(d, dt.time(m // 60, m % 60)) if d else None
            if occ and occ < now:
                cd, tag = "已结束", ("past",)
            elif occ:
                days = (occ.date() - now.date()).days
                cd = "今天" if days == 0 else f"{days} 天"
                tag = ("soon",) if days <= 3 else ()
            else:
                cd, tag = "", ()
            lead = f"{it['lead']} 分钟" if it.get("lead") else f"默认 {_data['settings'].get('default_lead', 15)} 分"
            tree.insert("", "end", tags=tag,
                        values=(it["id"], it["title"], it.get("date", ""), it.get("time", ""),
                                it.get("location", ""), lead, cd, it.get("note", "")))

    def refresh_events(self):
        tree = self.event_tree
        tree.delete(*tree.get_children())
        now = dt.datetime.now()
        items = sorted(_data["events"], key=lambda it: (it.get("date", ""), it.get("time", "")))
        for it in items:
            d = parse_date(it.get("date"))
            m = parse_hhmm(it.get("time"), 0)
            occ = dt.datetime.combine(d, dt.time(m // 60, m % 60)) if d else None
            if occ and occ < now:
                tag = ("past",)
            elif occ and (occ - now).total_seconds() <= 86400:
                tag = ("soon",)
            else:
                tag = ()
            lead = f"{it['lead']} 分钟" if it.get("lead") else f"默认 {_data['settings'].get('default_lead', 15)} 分"
            tree.insert("", "end", tags=tag,
                        values=(it["id"], it["title"], it.get("date", ""), it.get("time", ""),
                                lead, it.get("note", "")))

    def update_status(self):
        today = dt.date.today()
        week = semester_week(today, _data["settings"].get("semester_start"))
        today_courses = [c for c in _data["courses"]
                         if c.get("weekday") == today.weekday() and week_in_range(week, c.get("weeks"))]
        now = dt.datetime.now()
        nxt = None
        for key, occ, lead, title, body in compute_due_items(now):
            if nxt is None or occ < nxt[0]:
                nxt = (occ, title)
        wk_txt = f"第 {week} 教学周" if week else "未设置学期周"
        nxt_txt = f"下一项:{nxt[1]} {nxt[0].strftime('%m-%d %H:%M')}" if nxt else "暂无后续安排"
        remind = "提醒开启" if _data["settings"].get("reminder_on", True) else "提醒已关闭"
        self.status.config(text=f"  今天 {today.month}月{today.day}日 {DAY_NAMES[today.weekday()]} · {wk_txt} · "
                                f"今日 {len(today_courses)} 节课 · {nxt_txt} · {remind}")

    # ---- 后台消息 ----
    def _poll_queue(self):
        try:
            while True:
                kind, _ = self.ui_queue.get_nowait()
                if kind == "tick":
                    self.update_status()
                    self.refresh_exams()
                    self.refresh_events()
        except queue.Empty:
            pass
        except Exception:
            log_exc("UI 轮询异常")
        self.after(1000, self._poll_queue)

    def _refresh_every_min(self):
        self.draw_schedule()
        self.update_status()
        self.after(60_000, self._refresh_every_min)


# ---------------------------------------------------------------- 入口

def enable_dpi_awareness():
    """启用 Per-Monitor v2 DPI 感知,高缩放屏上窗口清晰、不发虚。"""
    import ctypes
    try:
        # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
        if ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):
            return
    except Exception:
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def system_scale_factor():
    """返回当前屏幕缩放系数(96 DPI = 1.0),失败返回 1.0。"""
    try:
        import ctypes
        dpi = ctypes.windll.user32.GetDpiForSystem()
        if dpi >= 96:
            return dpi / 96.0
    except Exception:
        pass
    return 1.0


def main():
    # 单实例:第二个实例直接退出
    try:
        lock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        lock.bind(("127.0.0.1", SINGLE_INSTANCE_PORT))
    except OSError:
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("我的课程表", "课程表程序已经在运行了(请看任务栏)。")
        return

    enable_dpi_awareness()

    global _data
    _data = load_data()
    if not os.path.exists(DATA_FILE):
        save_data()          # 首次运行先把数据文件落盘,方便用户找到
    q = queue.Queue()
    ReminderEngine(q).start()
    try:
        app = App(q)
        app.mainloop()
    except Exception:
        log_exc("主界面异常")
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("我的课程表", f"程序出现异常,详见日志:\n{LOG_FILE}")


if __name__ == "__main__":
    main()
