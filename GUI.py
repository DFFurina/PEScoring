# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import re
import os
import json
import warnings
import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

warnings.filterwarnings("ignore")

CONFIG_FILE = "体育评分配置.json"

# ====================== 默认配置（满分改为15分） ======================
DEFAULT_CONFIG = {
    '姓名列': 'C', '性别列': 'D',
    '第一类项目列': 'E', '第一类成绩列': 'F', '第一类得分列': 'G',
    '第二类项目列': 'I', '第二类成绩列': 'J', '第二类得分列': 'K',
    '第三类项目列': 'M', '第三类成绩列': 'N', '第三类得分列': 'O',
    '第四类项目列': 'Q', '第四类成绩列': 'R', '第四类得分列': 'S',
    '总分列': 'T',
    '满分标准': 15.0,               # 修改为15分满分
    '优秀标准': 12.0,
    '是否加排名列': False,
    '缺考背景色': 'FFFF00',
    '立定跳远单位': '厘米',
    '实心球单位': '米',
    '自定义项目': [],
    '总分小数位数': 2,               # 新增：总分保留小数位
    '缺考字体颜色': 'FF0000',        # 新增：缺考标红字体颜色（16进制）
    '排名并列方式': 'min',           # 新增：min（并列）或dense（连续）
    '是否自动排序班级': True,        # 新增：输出时按总分降序排序
    '统计包含优秀率': True,          # 新增：统计表是否加优秀率列
    '统计文件名后缀': '_统计汇总',      # 新增：统计文件后缀
    '当前版本': '1.0.0',              # 你当前的程序版本，发布新版时改这里
    '忽略更新版本': ''                # 用户选择“不再提醒”时记录被忽略的版本号

}

RUN_BONUS = {
    '女生800米_1分秒数': 205,
    '女生800米_0.5分秒数': 219,
    '男生1000米_1分秒数': 220,
    '男生1000米_0.5分秒数': 234,
}

STANDARD_FULL = {
    '引体向上满分次数': 11,
    '实心球男生满分米数': 9.7,
    '实心球女生满分米数': 6.8,
    '立定跳远男生满分米数': 2.49,
    '立定跳远女生满分米数': 1.99,
    '50米跑男生满分秒数': 7.1,
    '50米跑女生满分秒数': 8.1,
    '25米游泳男生满分秒数': 22.0,
    '25米游泳女生满分秒数': 25.0,
    '足球运球男生满分秒数': 7.6,
    '足球运球女生满分秒数': 8.5,
    '仰卧起坐满分次数': 50,
    '4分钟跳绳男生满分次数': 400,
    '4分钟跳绳女生满分次数': 405,
    '200米游泳满分秒数': 276
}

PINGPONG_SCORES = {25:3.00,24:2.88,23:2.76,22:2.64,21:2.52,20:2.40,19:2.28,18:2.16,17:2.04,16:1.92,
                   15:1.80,14:1.68,13:1.56,12:1.44,11:1.32,10:1.20,9:1.08,8:0.96,7:0.84,6:0.72,
                   5:0.60,4:0.48,3:0.36,2:0.24,1:0.12}

VOLLEYBALL_LIMITS = [45,43,40,37,34,31,29,26,23,20,18,16,14,12,10,8,6,5,4,3]
VOLLEYBALL_SCORES = [3.00,2.85,2.70,2.55,2.40,2.25,2.10,1.95,1.80,1.65,1.50,1.35,1.20,1.05,0.90,0.75,0.60,0.45,0.30,0.15]

HEADERS = ['序号', '班级', '姓名', '性别',
           '第一类项目', '第一类成绩', '第一类得分', '',
           '第二类项目', '第二类成绩', '第二类得分', '',
           '第三类项目', '第三类成绩', '第三类得分', '',
           '第四类项目', '第四类成绩', '第四类得分', '总分']

# ====================== 工具函数 ======================
def col_to_num(col):
    num = 0
    for c in col.upper():
        num = num * 26 + (ord(c) - ord('A') + 1)
    return num - 1

def time_to_seconds(s):
    if pd.isna(s) or str(s).strip() == '': return np.nan
    s = str(s).strip().replace('′', "'").replace('″', '"').replace(':', "'").replace('’', "'")
    s = re.sub(r'[^\d\'"]', '', s)
    try:
        if "'" in s:
            m, rest = s.split("'")
            sec = float(rest.replace('"', ''))
            return int(m) * 60 + sec
        return float(s)
    except:
        return np.nan

def to_num(x):
    if pd.isna(x): return np.nan
    return float(re.sub(r'[^\d.]', '', str(x)))

# ====================== 评分函数（已补全仰卧起坐、引体向上） ======================
def get_score(gender, project, value):
    if pd.isna(value) or str(value).strip() in ['', '-']:
        return "缺考"

    p = str(project).strip()
    p_map = {
        '800米':'800/1000米','1000米':'800/1000米','800/1000米':'800/1000米',
        '实心球':'实心球','双手头上前掷实心球':'实心球',
        '仰卧起坐':'仰卧起坐','一分钟仰卧起坐':'仰卧起坐',
        '4分钟跳绳':'4min跳绳','4min跳绳':'4min跳绳',
        '50米跑':'50m跑','排球垫球':'排球','排球':'排球',
        '乒乓':'乒乓球','乒乓球':'乒乓球',
        '武术操':'武术','武术':'武术','体操':'体操','羽毛球':'羽毛球','篮球':'篮球','足球':'足球运球'
    }
    p = p_map.get(p, p)

    val = to_num(value)
    if np.isnan(val):
        return "缺考"

    # 单位转换
    if p == '立定跳远' and DEFAULT_CONFIG['立定跳远单位'] == '厘米':
        val /= 100
    if p == '实心球' and DEFAULT_CONFIG['实心球单位'] == '厘米':
        val /= 100

    # 自定义项目（用户添加的）
    for custom in DEFAULT_CONFIG['自定义项目']:
        if custom['name'] == p:
            std = custom['full_value']
            if std == 0: return 0.0
            if custom['bigger_better']:
                ratio = val / std
            else:
                ratio = std / val if val > 0 else 0
            score = ratio * 3
            return round(max(0, min(3, score)), 2)

    # 写死的特殊项目：武术、体操（满分3分线性制）
    if p == '武术':
        std = 10.0
        ratio = val / std
        score = ratio * 3
        return round(max(0, min(3, score)), 2)

    if p == '体操':
        std = 20.0
        ratio = val / std
        score = ratio * 3
        return round(max(0, min(3, score)), 2)

    # 800/1000米特殊
    if p == '800/1000米':
        sec = time_to_seconds(value)
        if np.isnan(sec): return "缺考"
        if gender == '女':
            bonus = 1.0 if sec <= RUN_BONUS['女生800米_1分秒数'] else (0.5 if sec <= RUN_BONUS['女生800米_0.5分秒数'] else 0.0)
            table = [(220,6.0),(228,5.7),(236,5.4),(244,5.1),(252,4.8),(257,4.5),(262,4.2),
                     (267,3.9),(272,3.6),(278,3.3),(284,3.0),(290,2.7),(296,2.4),(302,2.1),
                     (308,1.8),(314,1.5),(320,1.2),(326,0.9),(332,0.6)]
        else:
            bonus = 1.0 if sec <= RUN_BONUS['男生1000米_1分秒数'] else (0.5 if sec <= RUN_BONUS['男生1000米_0.5分秒数'] else 0.0)
            table = [(235,6.0),(243,5.7),(251,5.4),(259,5.1),(267,4.8),(272,4.5),(277,4.2),
                     (282,3.9),(287,3.6),(293,3.3),(299,3.0),(305,2.7),(311,2.4),(317,2.1),
                     (323,1.8),(329,1.5),(335,1.2),(341,0.9),(347,0.6)]
        base = 0.0
        for lim, sc in table:
            if sec <= lim:
                base = sc
                break
        return round(base + bonus, 2)

    if p == '乒乓球':
        n = int(val or 0)
        if n >= 25: return 3.00
        return PINGPONG_SCORES.get(n, 0.0)

    if p == '排球':
        for l, s in zip(VOLLEYBALL_LIMITS, VOLLEYBALL_SCORES):
            if val >= l: return s
        return 0.0

    # 内置线性项目（严格比例，已补全仰卧起坐、引体向上）
    key = p
    if p == '实心球':
        key = '实心球男生满分米数' if gender == '男' else '实心球女生满分米数'
    elif p == '立定跳远':
        key = '立定跳远男生满分米数' if gender == '男' else '立定跳远女生满分米数'
    elif p == '50m跑':
        key = '50米跑男生满分秒数' if gender == '男' else '50米跑女生满分秒数'
    elif p == '25m游泳':
        key = '25米游泳男生满分秒数' if gender == '男' else '25米游泳女生满分秒数'
    elif p == '足球运球':
        key = '足球运球男生满分秒数' if gender == '男' else '足球运球女生满分秒数'
    elif p == '4min跳绳':
        key = '4分钟跳绳男生满分次数' if gender == '男' else '4分钟跳绳女生满分次数'
    elif p == '仰卧起坐':
        key = '仰卧起坐满分次数'
    elif p == '引体向上':
        key = '引体向上满分次数'

    if key in STANDARD_FULL:
        std = STANDARD_FULL[key]
        if std == 0: return 0.0
        ratio = val / std
        score = ratio * 3
        return round(max(0, min(3, score)), 2)

    return 3.00

# ====================== GUI 类（新增更多配置项） ======================
class SportsScoreGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("上海中考体育评分")
        self.root.geometry("1000x800")

        self.input_file = ""
        self.output_file = ""
        self.stats_file = ""
        self.sheets = []
        self.sheet_vars = []
        self.config_entries = {}
        self.standard_entries = {}
        self.custom_projects = []

        self.var_statistics = tk.IntVar()

        self.create_widgets()
        self.load_config()
        self.update_custom_list()
        self.check_for_update()


    def create_widgets(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)

        # 选项卡1: 主要操作
        tab_main = ttk.Frame(notebook)
        notebook.add(tab_main, text="主要操作")

        ttk.Label(tab_main, text="输入文件:", font=("Arial", 12)).grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.input_entry = ttk.Entry(tab_main, width=70)
        self.input_entry.grid(row=0, column=1, padx=10, pady=10)
        ttk.Button(tab_main, text="浏览...", command=self.select_input).grid(row=0, column=2, padx=10, pady=10)

        ttk.Label(tab_main, text="选择班级:", font=("Arial", 12)).grid(row=1, column=0, padx=10, pady=10, sticky="ne")
        self.sheet_frame = ttk.Frame(tab_main)
        self.sheet_frame.grid(row=1, column=1, columnspan=2, padx=10, pady=10, sticky="w")

        ttk.Label(tab_main, text="输出文件:", font=("Arial", 12)).grid(row=2, column=0, padx=10, pady=10, sticky="e")
        self.output_entry = ttk.Entry(tab_main, width=70)
        self.output_entry.grid(row=2, column=1, padx=10, pady=10)
        ttk.Button(tab_main, text="浏览...", command=self.select_output).grid(row=2, column=2, padx=10, pady=10)

        stats_frame = ttk.LabelFrame(tab_main, text="统计汇总")
        stats_frame.grid(row=3, column=0, columnspan=3, padx=10, pady=10, sticky="ew")
        ttk.Checkbutton(stats_frame, text="生成班级统计汇总（单独文件）", variable=self.var_statistics, command=self.toggle_stats_path).pack(anchor="w", padx=10, pady=5)
        self.stats_entry = ttk.Entry(stats_frame, width=70)
        self.stats_entry.pack(padx=10, pady=5, fill="x")
        self.stats_entry.config(state="disabled")
        ttk.Button(stats_frame, text="浏览...", command=self.select_stats_output).pack(pady=5)

        ttk.Button(tab_main, text="开始评分", command=self.process, style="Accent.TButton").grid(row=4, column=1, pady=30)
        ttk.Button(tab_main, text="保存当前配置", command=self.save_config).grid(row=5, column=1, pady=10)

        # 选项卡2: 列配置
        tab_config = ttk.Frame(notebook)
        notebook.add(tab_config, text="列配置")

        canvas_config = tk.Canvas(tab_config)
        scrollbar_config = ttk.Scrollbar(tab_config, orient="vertical", command=canvas_config.yview)
        scrollable_frame_config = ttk.Frame(canvas_config)

        scrollable_frame_config.bind("<Configure>", lambda e: canvas_config.configure(scrollregion=canvas_config.bbox("all")))

        canvas_config.create_window((0, 0), window=scrollable_frame_config, anchor="nw")
        canvas_config.configure(yscrollcommand=scrollbar_config.set)

        config_frame = ttk.LabelFrame(scrollable_frame_config, text="自定义列字母（大写字母，如 C、T）")
        config_frame.pack(fill="both", expand=True, padx=20, pady=20)

        row = 0
        for label, default in DEFAULT_CONFIG.items():
            if label in ['满分标准', '优秀标准', '是否加排名列', '缺考背景色', '立定跳远单位', '实心球单位', '自定义项目']:
                continue
            ttk.Label(config_frame, text=label + ":").grid(row=row, column=0, padx=10, pady=5, sticky="e")
            entry = ttk.Entry(config_frame, width=10)
            entry.insert(0, str(default))
            entry.grid(row=row, column=1, padx=10, pady=5)
            self.config_entries[label] = entry
            row += 1

        canvas_config.pack(side="left", fill="both", expand=True)
        scrollbar_config.pack(side="right", fill="y")

        # 评分标准配置选项卡
        tab_standard = ttk.Frame(notebook)
        notebook.add(tab_standard, text="评分标准配置")

        canvas_standard = tk.Canvas(tab_standard)
        scrollbar_standard = ttk.Scrollbar(tab_standard, orient="vertical", command=canvas_standard.yview)
        scrollable_frame_standard = ttk.Frame(canvas_standard)

        scrollable_frame_standard.bind("<Configure>", lambda e: canvas_standard.configure(scrollregion=canvas_standard.bbox("all")))

        canvas_standard.create_window((0, 0), window=scrollable_frame_standard, anchor="nw")
        canvas_standard.configure(yscrollcommand=scrollbar_standard.set)

        run_frame = ttk.LabelFrame(scrollable_frame_standard, text="800/1000米附加分条件（秒数）")
        run_frame.pack(fill="x", padx=20, pady=10)

        row = 0
        for label, default in RUN_BONUS.items():
            ttk.Label(run_frame, text=label + ":").grid(row=row, column=0, padx=10, pady=5, sticky="e")
            entry = ttk.Entry(run_frame, width=15)
            entry.insert(0, str(default))
            entry.grid(row=row, column=1, padx=10, pady=5)
            self.standard_entries[label] = entry
            row += 1

        other_frame = ttk.LabelFrame(scrollable_frame_standard, text="其他项目满分标准")
        other_frame.pack(fill="both", expand=True, padx=20, pady=10)

        row = 0
        for label, default in STANDARD_FULL.items():
            ttk.Label(other_frame, text=label + ":").grid(row=row, column=0, padx=10, pady=5, sticky="e")
            entry = ttk.Entry(other_frame, width=15)
            entry.insert(0, str(default))
            entry.grid(row=row, column=1, padx=10, pady=5)
            self.standard_entries[label] = entry
            row += 1

        # 单位配置
        unit_frame = ttk.LabelFrame(scrollable_frame_standard, text="单位配置")
        unit_frame.pack(fill="x", padx=20, pady=10)

        ttk.Label(unit_frame, text="立定跳远单位:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
        var_jump = tk.StringVar(value=DEFAULT_CONFIG['立定跳远单位'])
        ttk.Combobox(unit_frame, textvariable=var_jump, values=['厘米', '米'], state="readonly", width=10).grid(row=0, column=1, padx=10, pady=5)
        self.standard_entries['立定跳远单位'] = var_jump

        ttk.Label(unit_frame, text="实心球单位:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        var_ball = tk.StringVar(value=DEFAULT_CONFIG['实心球单位'])
        ttk.Combobox(unit_frame, textvariable=var_ball, values=['米', '厘米'], state="readonly", width=10).grid(row=1, column=1, padx=10, pady=5)
        self.standard_entries['实心球单位'] = var_ball

       # 其他配置（新增更多项）
        extra_frame = ttk.LabelFrame(scrollable_frame_standard, text="其他配置")
        extra_frame.pack(fill="x", padx=20, pady=10)

        row = 0
        ttk.Label(extra_frame, text="满分标准:").grid(row=row, column=0, padx=10, pady=5, sticky="e")
        entry_full = ttk.Entry(extra_frame, width=15)
        entry_full.insert(0, str(DEFAULT_CONFIG['满分标准']))
        entry_full.grid(row=row, column=1, padx=10, pady=5)
        self.standard_entries['满分标准'] = entry_full
        row += 1

        ttk.Label(extra_frame, text="优秀标准:").grid(row=row, column=0, padx=10, pady=5, sticky="e")
        entry_excellent = ttk.Entry(extra_frame, width=15)
        entry_excellent.insert(0, str(DEFAULT_CONFIG['优秀标准']))
        entry_excellent.grid(row=row, column=1, padx=10, pady=5)
        self.standard_entries['优秀标准'] = entry_excellent
        row += 1

        ttk.Label(extra_frame, text="是否加排名列:").grid(row=row, column=0, padx=10, pady=5, sticky="e")
        var_rank = tk.BooleanVar(value=DEFAULT_CONFIG['是否加排名列'])
        ttk.Checkbutton(extra_frame, variable=var_rank).grid(row=row, column=1, padx=10, pady=5)
        self.standard_entries['是否加排名列'] = var_rank
        row += 1

        ttk.Label(extra_frame, text="缺考背景色（16进制）:").grid(row=row, column=0, padx=10, pady=5, sticky="e")
        entry_bg = ttk.Entry(extra_frame, width=15)
        entry_bg.insert(0, DEFAULT_CONFIG['缺考背景色'])
        entry_bg.grid(row=row, column=1, padx=10, pady=5)
        self.standard_entries['缺考背景色'] = entry_bg
        row += 1

        ttk.Label(extra_frame, text="缺考字体颜色（16进制）:").grid(row=row, column=0, padx=10, pady=5, sticky="e")
        entry_font = ttk.Entry(extra_frame, width=15)
        entry_font.insert(0, DEFAULT_CONFIG['缺考字体颜色'])
        entry_font.grid(row=row, column=1, padx=10, pady=5)
        self.standard_entries['缺考字体颜色'] = entry_font
        row += 1

        ttk.Label(extra_frame, text="总分小数位数:").grid(row=row, column=0, padx=10, pady=5, sticky="e")
        entry_decimal = ttk.Entry(extra_frame, width=15)
        entry_decimal.insert(0, str(DEFAULT_CONFIG['总分小数位数']))
        entry_decimal.grid(row=row, column=1, padx=10, pady=5)
        self.standard_entries['总分小数位数'] = entry_decimal
        row += 1

        ttk.Label(extra_frame, text="排名并列方式:").grid(row=row, column=0, padx=10, pady=5, sticky="e")
        var_rank_method = tk.StringVar(value=DEFAULT_CONFIG['排名并列方式'])
        ttk.Combobox(extra_frame, textvariable=var_rank_method, values=['min', 'dense'], state="readonly", width=10).grid(row=row, column=1, padx=10, pady=5)
        self.standard_entries['排名并列方式'] = var_rank_method
        row += 1

        ttk.Label(extra_frame, text="是否自动排序班级:").grid(row=row, column=0, padx=10, pady=5, sticky="e")
        var_sort = tk.BooleanVar(value=DEFAULT_CONFIG['是否自动排序班级'])
        ttk.Checkbutton(extra_frame, variable=var_sort).grid(row=row, column=1, padx=10, pady=5)
        self.standard_entries['是否自动排序班级'] = var_sort
        row += 1

        ttk.Label(extra_frame, text="统计包含优秀率:").grid(row=row, column=0, padx=10, pady=5, sticky="e")
        var_excellent_rate = tk.BooleanVar(value=DEFAULT_CONFIG['统计包含优秀率'])
        ttk.Checkbutton(extra_frame, variable=var_excellent_rate).grid(row=row, column=1, padx=10, pady=5)
        self.standard_entries['统计包含优秀率'] = var_excellent_rate
        row += 1

        ttk.Label(extra_frame, text="统计文件名后缀:").grid(row=row, column=0, padx=10, pady=5, sticky="e")
        entry_suffix = ttk.Entry(extra_frame, width=20)
        entry_suffix.insert(0, DEFAULT_CONFIG['统计文件名后缀'])
        entry_suffix.grid(row=row, column=1, padx=10, pady=5)
        self.standard_entries['统计文件名后缀'] = entry_suffix

        # 自定义项目
        custom_frame = ttk.LabelFrame(scrollable_frame_standard, text="自定义项目（满分3分制）")
        custom_frame.pack(fill="both", expand=True, padx=20, pady=10)

        add_frame = ttk.Frame(custom_frame)
        add_frame.pack(fill="x", pady=5)

        ttk.Label(add_frame, text="项目名:").grid(row=0, column=0, padx=5)
        self.custom_name = ttk.Entry(add_frame, width=15)
        self.custom_name.grid(row=0, column=1, padx=5)

        ttk.Label(add_frame, text="满分值:").grid(row=0, column=2, padx=5)
        self.custom_full = ttk.Entry(add_frame, width=10)
        self.custom_full.grid(row=0, column=3, padx=5)

        ttk.Label(add_frame, text="单位:").grid(row=0, column=4, padx=5)
        self.custom_unit = ttk.Entry(add_frame, width=10)
        self.custom_unit.grid(row=0, column=5, padx=5)

        ttk.Label(add_frame, text="越大越好:").grid(row=0, column=6, padx=5)
        self.custom_bigger = tk.BooleanVar(value=True)
        ttk.Checkbutton(add_frame, variable=self.custom_bigger).grid(row=0, column=7, padx=5)

        ttk.Button(add_frame, text="添加项目", command=self.add_custom_project).grid(row=0, column=8, padx=10)

        self.custom_listbox = tk.Listbox(custom_frame, height=10)
        self.custom_listbox.pack(fill="both", expand=True, padx=10, pady=5)

        ttk.Button(custom_frame, text="删除选中项目", command=self.delete_custom_project).pack(pady=5)

        canvas_standard.pack(side="left", fill="both", expand=True)
        scrollbar_standard.pack(side="right", fill="y")

    def add_custom_project(self):
        name = self.custom_name.get().strip()
        if not name:
            messagebox.showwarning("警告", "请输入项目名")
            return
        try:
            full = float(self.custom_full.get().strip())
        except:
            messagebox.showwarning("警告", "满分值必须是数字")
            return
        unit = self.custom_unit.get().strip() or "个"
        bigger = self.custom_bigger.get()

        project = {"name": name, "full_value": full, "unit": unit, "bigger_better": bigger}
        self.custom_projects.append(project)
        self.update_custom_list()
        self.custom_name.delete(0, tk.END)
        self.custom_full.delete(0, tk.END)
        self.custom_unit.delete(0, tk.END)

    def delete_custom_project(self):
        selected = self.custom_listbox.curselection()
        if selected:
            del self.custom_projects[selected[0]]
            self.update_custom_list()

    def update_custom_list(self):
        self.custom_listbox.delete(0, tk.END)
        for proj in self.custom_projects:
            direction = "越大越好" if proj['bigger_better'] else "越小越好"
            text = f"{proj['name']} | 满分 {proj['full_value']} {proj['unit']} | {direction}"
            self.custom_listbox.insert(tk.END, text)

    def toggle_stats_path(self):
        if self.var_statistics.get():
            self.stats_entry.config(state="normal")
            if self.input_file:
                dir_name = os.path.dirname(self.input_file)
                base_name = os.path.basename(self.input_file).rsplit('.', 1)[0]
                suggested = os.path.join(dir_name, f"{base_name}_统计汇总.xlsx")
                self.stats_file = suggested
                self.stats_entry.delete(0, tk.END)
                self.stats_entry.insert(0, suggested)
        else:
            self.stats_entry.config(state="disabled")
            self.stats_file = ""

    def select_stats_output(self):
        if self.var_statistics.get():
            file = filedialog.asksaveasfilename(title="保存统计汇总", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if file:
                self.stats_file = file
                self.stats_entry.delete(0, tk.END)
                self.stats_entry.insert(0, file)

    def save_config(self):
        config = {}
        config['当前版本'] = DEFAULT_CONFIG['当前版本']
        config['忽略更新版本'] = DEFAULT_CONFIG.get('忽略更新版本', '')

        for label, entry in self.config_entries.items():
            value = entry.get().strip().upper()
            if value:
                config[label] = value

        for label, entry in self.standard_entries.items():
            if isinstance(entry, (tk.StringVar, tk.BooleanVar)):
                config[label] = entry.get()
            else:
                value = entry.get().strip()
                if value:
                    try:
                        config[label] = float(value)
                    except:
                        config[label] = value

        config['自定义项目'] = self.custom_projects

        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
        messagebox.showinfo("成功", "配置已保存！")

    def load_config(self):
        if not os.path.exists(CONFIG_FILE):
            return
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                saved = json.load(f)

            for label, entry in self.config_entries.items():
                if label in saved:
                    entry.delete(0, tk.END)
                    entry.insert(0, saved[label])

            for label, entry in self.standard_entries.items():
                if label in saved:
                    if isinstance(entry, (tk.StringVar, tk.BooleanVar)):
                        entry.set(saved[label])
                    else:
                        entry.delete(0, tk.END)
                        entry.insert(0, str(saved[label]))

            if '自定义项目' in saved:
                self.custom_projects = saved['自定义项目']
                self.update_custom_list()
                DEFAULT_CONFIG['自定义项目'] = self.custom_projects[:]
        except Exception as e:
            messagebox.showerror("错误", f"加载配置失败：{e}")

    def select_input(self):
        file = filedialog.askopenfilename(title="选择原始Excel文件", filetypes=[("Excel files", "*.xlsx *.xls")])
        if file:
            self.input_file = file
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, file)
            self.load_sheets()

            dir_name = os.path.dirname(file)
            base_name = os.path.basename(file).rsplit('.', 1)[0]
            suggested = os.path.join(dir_name, f"{base_name}_已评分.xlsx")
            self.output_file = suggested
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, suggested)

            if self.var_statistics.get():
                stats_suggested = os.path.join(dir_name, f"{base_name}_统计汇总.xlsx")
                self.stats_file = stats_suggested
                self.stats_entry.delete(0, tk.END)
                self.stats_entry.insert(0, stats_suggested)

    def load_sheets(self):
        try:
            xl = pd.ExcelFile(self.input_file)
            self.sheets = xl.sheet_names
            for widget in self.sheet_frame.winfo_children():
                widget.destroy()
            self.sheet_vars = []
            for sheet in self.sheets:
                var = tk.IntVar(value=1)
                cb = ttk.Checkbutton(self.sheet_frame, text=sheet, variable=var)
                cb.pack(anchor="w")
                self.sheet_vars.append(var)
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败：{e}")

    def select_output(self):
        file = filedialog.asksaveasfilename(title="保存评分结果", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file:
            self.output_file = file
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, file)

    def get_current_config(self):
        config = {}
        for label, entry in self.config_entries.items():
            value = entry.get().strip().upper()
            if value:
                config[label] = value

        name_col = config.get('姓名列', 'C')
        gender_col = config.get('性别列', 'D')
        proj_cols = [config.get('第一类项目列', 'E'), config.get('第二类项目列', 'I'),
                     config.get('第三类项目列', 'M'), config.get('第四类项目列', 'Q')]
        res_cols = [config.get('第一类成绩列', 'F'), config.get('第二类成绩列', 'J'),
                    config.get('第三类成绩列', 'N'), config.get('第四类成绩列', 'R')]
        score_cols = [config.get('第一类得分列', 'G'), config.get('第二类得分列', 'K'),
                      config.get('第三类得分列', 'O'), config.get('第四类得分列', 'S')]
        total_col = config.get('总分列', 'T')

        # 更新全局配置
        DEFAULT_CONFIG['立定跳远单位'] = self.standard_entries['立定跳远单位'].get()
        DEFAULT_CONFIG['实心球单位'] = self.standard_entries['实心球单位'].get()
        DEFAULT_CONFIG['自定义项目'] = self.custom_projects

        try:
            DEFAULT_CONFIG['满分标准'] = float(self.standard_entries['满分标准'].get())
            DEFAULT_CONFIG['优秀标准'] = float(self.standard_entries['优秀标准'].get())
            DEFAULT_CONFIG['是否加排名列'] = self.standard_entries['是否加排名列'].get()
            DEFAULT_CONFIG['缺考背景色'] = self.standard_entries['缺考背景色'].get().upper()
        except:
            pass

        # 更新附加分标准
        global RUN_BONUS, STANDARD_FULL
        for label, entry in self.standard_entries.items():
            if label in RUN_BONUS or label in STANDARD_FULL:
                try:
                    RUN_BONUS[label] = float(entry.get())
                except:
                    pass

        return (name_col, gender_col, proj_cols, res_cols, score_cols, total_col)

    def process(self):
        if not self.input_file:
            messagebox.showwarning("警告", "请先选择输入文件！")
            return
        if not self.output_file:
            messagebox.showwarning("警告", "请先选择输出文件！")
            return

        selected_indices = [i for i, var in enumerate(self.sheet_vars) if var.get() == 1]
        if not selected_indices:
            messagebox.showwarning("警告", "请至少选择一个班级！")
            return

        selected_sheets = [self.sheets[i] for i in selected_indices]

        NAME_COL, GENDER_COL, PROJ_COLS, RES_COLS, SCORE_COLS, TOTAL_COL = self.get_current_config()

        NAME_N = col_to_num(NAME_COL)
        GENDER_N = col_to_num(GENDER_COL)
        PROJ_N = [col_to_num(c) for c in PROJ_COLS]
        RES_N = [col_to_num(c) for c in RES_COLS]
        SCORE_N = [col_to_num(c) for c in SCORE_COLS]
        TOTAL_N = col_to_num(TOTAL_COL)

        # 从配置读取更多自定义项
        decimal_places = int(DEFAULT_CONFIG.get('总分小数位数', 2))
        rank_method = DEFAULT_CONFIG.get('排名并列方式', 'min')
        auto_sort = DEFAULT_CONFIG.get('是否自动排序班级', True)
        excellent_rate_in_stats = DEFAULT_CONFIG.get('统计包含优秀率', True)
        stats_suffix = DEFAULT_CONFIG.get('统计文件名后缀', '_统计汇总')

        messagebox.showinfo("开始", f"正在处理 {len(selected_sheets)} 个班级，请稍等...")

        with pd.ExcelWriter(self.output_file, engine='openpyxl') as writer:
            for sheet_name in selected_sheets:
                df = pd.read_excel(self.input_file, sheet_name=sheet_name, header=None, skiprows=1, dtype=str)

                max_col = max(TOTAL_N, max(SCORE_N + RES_N + PROJ_N, default=0))
                while len(df.columns) <= max_col:
                    df[len(df.columns)] = np.nan

                missing_rows = []

                for i in range(len(df)):
                    if pd.isna(df.iat[i, NAME_N]): 
                        continue
                    gender = str(df.iat[i, GENDER_N]).strip()
                    if gender not in ['男', '女']: 
                        continue

                    scores_num = []
                    has_missing = False

                    for k in range(4):
                        proj = df.iat[i, PROJ_N[k]]
                        res = df.iat[i, RES_N[k]]
                        sc = get_score(gender, proj, res)
                        df.iat[i, SCORE_N[k]] = sc
                        if sc == "缺考":
                            scores_num.append(0.0)
                            has_missing = True
                        else:
                            scores_num.append(sc if isinstance(sc, (int, float)) else 0.0)

                    total = round(sum(scores_num), decimal_places)
                    df.iat[i, TOTAL_N] = total
                    if has_missing:
                        missing_rows.append(i + 2)

                # 先设置列名（关键！避免 KeyError: '总分'）
                df.columns = HEADERS[:len(df.columns)]

                # 再加排名列（可选 + 自定义方式 + 自动排序）
                if DEFAULT_CONFIG['是否加排名列']:
                    df['总分'] = pd.to_numeric(df['总分'], errors='coerce')
                    df['排名'] = df['总分'].rank(method=rank_method, ascending=False)
                    if auto_sort:
                        df = df.sort_values('总分', ascending=False).reset_index(drop=True)
                    df['排名'] = df['排名'].astype(int)

                df.to_excel(writer, sheet_name=sheet_name, index=False)

        # 标红（使用自定义颜色）
        wb = load_workbook(self.output_file)
        red_font = Font(color=DEFAULT_CONFIG['缺考字体颜色'])
        fill = PatternFill("solid", fgColor=DEFAULT_CONFIG['缺考背景色'])
        for sheet_name in selected_sheets:
            ws = wb[sheet_name]
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    col_idx = cell.column - 1
                    if col_idx in SCORE_N and cell.value == "缺考":
                        cell.font = red_font
                        cell.fill = fill
                    if col_idx == TOTAL_N and cell.row in missing_rows:
                        cell.font = red_font
                        cell.fill = fill

            # 自动列宽
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                ws.column_dimensions[column].width = min(max_length + 2, 40)

        wb.save(self.output_file)

        # 班级统计汇总（单独文件，可选优秀率）
        if self.var_statistics.get() and self.stats_file:
            stats = []
            full_threshold = DEFAULT_CONFIG['满分标准']
            excellent_threshold = DEFAULT_CONFIG['优秀标准']
            for sheet_name in selected_sheets:
                df = pd.read_excel(self.output_file, sheet_name=sheet_name)
                total_students = len(df)
                valid_total = pd.to_numeric(df['总分'], errors='coerce')
                avg = valid_total.mean()
                full = (valid_total >= full_threshold).sum()
                excellent = (valid_total >= excellent_threshold).sum()
                missing = (df.iloc[:, [6,10,14,18]] == "缺考").sum().sum()
                max_score = valid_total.max()
                min_score = valid_total.min()

                row_data = [sheet_name, total_students, round(avg, 2), full, excellent, missing, max_score, min_score]
                if excellent_rate_in_stats and total_students > 0:
                    excellent_rate_pct = round(excellent / total_students * 100, 2)
                    row_data.append(excellent_rate_pct)

                stats.append(row_data)

            columns = ['班级', '人数', '平均分', '满分人数', '优秀人数', '缺考项目数', '最高分', '最低分']
            if excellent_rate_in_stats:
                columns.append('优秀率(%)')

            stats_df = pd.DataFrame(stats, columns=columns)
            stats_df.to_excel(self.stats_file, index=False)
            os.startfile(self.stats_file)

        messagebox.showinfo("成功", "所有任务完成！")
        os.startfile(self.output_file)

    def check_for_update(self):
        current_version = DEFAULT_CONFIG['当前版本']
        ignore_version = DEFAULT_CONFIG.get('忽略更新版本', '')

        try:
            response = requests.get("https://1427.tech/projects/PEScoring/latest", timeout=5)
            if response.status_code == 200:
                latest_version = response.text.strip()
                if latest_version != current_version and latest_version != ignore_version:
                    result = messagebox.askyesnocancel(
                        "发现新版本",
                        f"当前版本：{current_version}\n最新版本：{latest_version}\n\n是否打开下载页面？"
                    )
                    if result is True:  # 是 → 打开浏览器
                        import webbrowser
                        webbrowser.open("https://1427.tech/projects/PEScoring")
                    elif result is False:  # 否 → 不再提醒此版本
                        DEFAULT_CONFIG['忽略更新版本'] = latest_version
                        self.save_config()  # 保存到配置文件
        except:
            pass  # 网络错误或超时，静默忽略

if __name__ == "__main__":
    root = tk.Tk()
    app = SportsScoreGUI(root)
    root.mainloop()


