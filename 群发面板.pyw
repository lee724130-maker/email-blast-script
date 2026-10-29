# -*- coding: utf-8 -*-
"""
邮件群发面板 —— tkinter 桌面版（MVP）
双击本文件即可打开面板（本机已装 Python；若闪退见「使用说明.md」用命令行启动看报错）。
操作流：选发信账号 → 导入名单 → 核对列识别 → 勾选收件人 → 预览 → 群发。
真正发送复用同目录 mailer.py + panel_core.py。
"""
import sys
import os
import io
import re
import queue
import threading
from datetime import datetime, timedelta

# .pyw（pythonw）下没有控制台，mailer.connect_smtp 里的 print 会打到 None，先把 stdout 兜住
class _Noop:
    def write(self, *a): return 0
    def flush(self): pass
if sys.stdout is None:
    sys.stdout = _Noop()
if sys.stderr is None:
    sys.stderr = _Noop()

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import panel_core as pc
import mailer

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _HAVE_DND = True
except Exception:
    TkinterDnD = None
    _HAVE_DND = False


def _split_tcl_list(s):
    """把 tkdnd 传入的路径列表字符串拆成普通列表，正确处理 {带空格 的路径}。"""
    items, buf, i, n = [], "", 0, len(s)
    while i < n:
        c = s[i]
        if c == "{":
            i += 1
            depth, tok = 1, ""
            while i < n and depth > 0:
                if s[i] == "{":
                    depth += 1
                elif s[i] == "}":
                    depth -= 1
                    if depth == 0:
                        i += 1
                        break
                tok += s[i] if s[i] not in "{}" or depth > 1 else ""
                i += 1 if depth > 0 else 0
            if buf:
                buf += tok
                items.append(buf)
                buf = ""
            else:
                items.append(tok)
            continue
        if c in " \t\r\n":
            if buf:
                items.append(buf)
                buf = ""
            i += 1
            continue
        buf += c
        i += 1
    if buf:
        items.append(buf)
    return items


FIXED_OPT = "（固定值，右侧填写）"
LOCAL_OPT = "（用邮箱前缀）"


class AccountForm(tk.Toplevel):
    """新增/编辑单个发信账号。返回 self.result（dict）或 None。"""
    def __init__(self, master, profile=None, ses=False):
        super().__init__(master)
        self.title("编辑发信账号")
        self.geometry("460x520")
        self.resizable(False, False)
        self.result = None
        p = dict(profile or (pc.default_ses_profile() if ses else pc.default_profile()))
        self.p = p

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)
        r = 0
        self.e = {}
        for key, label in [("名称", "账号备注名"), ("发信人地址", "发信地址(From)"),
                           ("发信人名称", "发件人显示名"), ("回复地址", "回复到(Reply-To)"),
                           ("邮件主题", "邮件主题(可用{{变量}})")]:
            ttk.Label(frm, text=label).grid(row=r, column=0, sticky="w", pady=2)
            var = tk.StringVar(value=str(p.get(key, "")))
            ttk.Entry(frm, textvariable=var, width=42).grid(row=r, column=1, sticky="we", pady=2)
            self.e[key] = var
            r += 1

        ttk.Label(frm, text="发送通道").grid(row=r, column=0, sticky="w", pady=6)
        self.chan = ttk.Combobox(frm, values=["smtp", "api"], width=10, state="readonly")
        self.chan.set(p.get("通道", "smtp"))
        self.chan.grid(row=r, column=1, sticky="w", pady=6)
        self.chan.bind("<<ComboboxSelected>>", lambda e: self._switch())
        r += 1

        # SMTP 区
        self.smtp_frm = ttk.LabelFrame(frm, text="SMTP 设置", padding=8)
        for key, label, typ in [("SMTP服务器", "服务器", str), ("SMTP端口", "端口", int),
                                ("SMTP用户名", "用户名(=发信地址)", str), ("SMTP密码", "密码/授权码", str)]:
            var = tk.StringVar(value=str(p.get(key, "")))
            self.e[key] = var
        self.ssl = tk.BooleanVar(value=bool(p.get("使用SSL", True)))
        r += 1

        # SES 区
        self.ses_frm = ttk.LabelFrame(frm, text="腾讯云 SES 设置", padding=8)
        ses = p.get("腾讯云SES", {})
        for key in ["SecretId", "SecretKey", "Region", "模板ID"]:
            self.e["SES_" + key] = tk.StringVar(value=str(ses.get(key, "")))

        ttk.Label(frm, text="间隔/封(秒)").grid(row=r + 6, column=0, sticky="w")
        self.e["每封间隔秒"] = tk.StringVar(value=str(p.get("每封间隔秒", 2.0)))
        ttk.Entry(frm, textvariable=self.e["每封间隔秒"], width=10).grid(row=r + 6, column=1, sticky="w")

        btns = ttk.Frame(self)
        btns.pack(pady=10)
        ttk.Button(btns, text="保存", command=self._save).pack(side="left", padx=6)
        ttk.Button(btns, text="取消", command=self.destroy).pack(side="left", padx=6)
        frm.columnconfigure(1, weight=1)
        self._switch()

    def _build_subframes(self):
        pass

    def _switch(self):
        api = self.chan.get() == "api"
        self.smtp_frm.destroy(); self.ses_frm.destroy()
        parent = self.winfo_children()[0]
        if api:
            f = ttk.LabelFrame(parent, text="腾讯云 SES（需已在控制台建好并审核通过的模板ID）", padding=8)
            for i, (key, label) in enumerate([("SecretId", "SecretId"), ("SecretKey", "SecretKey"),
                                              ("Region", "Region"), ("模板ID", "模板ID")]):
                ttk.Label(f, text=label).grid(row=i, column=0, sticky="w", pady=1)
                ttk.Entry(f, textvariable=self.e["SES_" + key], width=40).grid(row=i, column=1, sticky="we")
            f.grid(row=8, column=0, columnspan=2, sticky="we", pady=6)
            f.columnconfigure(1, weight=1)
        else:
            f = ttk.LabelFrame(parent, text="SMTP（阿里云/企业微信Exmail/任意SMTP 都用这里）", padding=8)
            for i, (key, label) in enumerate([("SMTP服务器", "服务器"), ("SMTP端口", "端口"),
                                              ("SMTP用户名", "用户名=发信地址"), ("SMTP密码", "密码/授权码")]):
                ttk.Label(f, text=label).grid(row=i, column=0, sticky="w", pady=1)
                ttk.Entry(f, textvariable=self.e[key], width=40).grid(row=i, column=1, sticky="we")
            ttk.Checkbutton(f, text="使用SSL(端口465勾选；25/587不勾)", variable=self.ssl).grid(
                row=4, column=1, sticky="w")
            f.grid(row=8, column=0, columnspan=2, sticky="we", pady=6)
            f.columnconfigure(1, weight=1)

    def _save(self):
        chan = self.chan.get()
        prof = {
            "名称": self.e["名称"].get().strip() or "未命名",
            "通道": chan,
            "发信人地址": self.e["发信人地址"].get().strip(),
            "发信人名称": self.e["发信人名称"].get().strip(),
            "回复地址": self.e["回复地址"].get().strip(),
            "邮件主题": self.e["邮件主题"].get().strip(),
            "每封间隔秒": _tofloat(self.e["每封间隔秒"].get(), 2.0),
            "每100封额外暂停秒": 60,
        }
        if chan == "api":
            prof["腾讯云SES"] = {
                "SecretId": self.e["SES_SecretId"].get().strip(),
                "SecretKey": self.e["SES_SecretKey"].get().strip(),
                "Region": self.e["SES_Region"].get().strip() or "ap-guangzhou",
                "模板ID": _toint(self.e["SES_模板ID"].get(), 0),
            }
        else:
            prof.update({
                "SMTP服务器": self.e["SMTP服务器"].get().strip(),
                "SMTP端口": _toint(self.e["SMTP端口"].get(), 465),
                "使用SSL": bool(self.ssl.get()),
                "SMTP用户名": self.e["SMTP用户名"].get().strip(),
                "SMTP密码": self.e["SMTP密码"].get(),
            })
        if not prof["发信人地址"]:
            messagebox.showwarning("提示", "发信地址不能为空", parent=self)
            return
        self.result = prof
        self.destroy()


def _tofloat(s, d):
    try: return float(s)
    except Exception: return d


def _toint(s, d):
    try: return int(float(s))
    except Exception: return d


class App(TkinterDnD.Tk if _HAVE_DND else tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("邮件群发面板  ·  多通道 · 智能识别名单")
        self.geometry("1000x720")
        self.profiles = []
        self.headers = []
        self.records = []
        self.email_col = None
        self.var_map = {}
        self.items = []
        self.selected = set()
        self._send_order = []
        self.sender = None
        self.q = queue.Queue()
        self.dedup_mode = tk.StringVar(value="email")  # 去重依据：email=按邮箱 / row=按整行
        self.group_of = {}          # 收件人下标 -> 分组名
        self.scheduled_at = None    # 定时发送目标时间
        self._sched_job = None      # after() 句柄

        self._build()
        self._load_accounts()
        self._enable_dnd()
        self.after(120, self._poll)

    # ---------------- UI ----------------
    def _build(self):
        pad = dict(padx=8, pady=4)

        top = ttk.LabelFrame(self, text="① 选择发信账号", padding=8)
        top.pack(fill="x", **pad)
        self.acc_combo = ttk.Combobox(top, state="readonly", width=42)
        self.acc_combo.pack(side="left")
        self.acc_combo.bind("<<ComboboxSelected>>", lambda e: self._on_account())
        ttk.Button(top, text="管理账号", command=self._manage_accounts).pack(side="left", padx=6)
        ttk.Button(top, text="刷新", command=self._load_accounts).pack(side="left", padx=2)
        self.acc_info = tk.StringVar(value="")
        ttk.Label(top, textvariable=self.acc_info, foreground="#555").pack(side="left", padx=12)

        imp = ttk.LabelFrame(self, text="② 导入客户名单", padding=8)
        imp.pack(fill="x", **pad)
        self.imp_frame = imp
        ttk.Button(imp, text="导入 CSV", command=self._imp_csv).pack(side="left")
        ttk.Button(imp, text="导入 Excel(.xlsx)", command=self._imp_xlsx).pack(side="left", padx=6)
        ttk.Button(imp, text="粘贴/表格编辑…", command=self._imp_paste).pack(side="left")
        self.header_is_title = tk.BooleanVar(value=True)
        ttk.Checkbutton(imp, text="Excel 首行为表头", variable=self.header_is_title).pack(side="left", padx=10)
        ttk.Button(imp, text="清空", command=self._clear_list).pack(side="left", padx=6)
        ttk.Label(imp, text="   去重依据:").pack(side="left")
        self._dedup_disp = {"按邮箱": "email", "按整行": "row"}
        self._dedup_rev = {v: k for k, v in self._dedup_disp.items()}
        self.dedup_combo = ttk.Combobox(imp, values=list(self._dedup_disp.keys()), state="readonly", width=7)
        self.dedup_combo.set(self._dedup_rev[self.dedup_mode.get()])
        self.dedup_combo.pack(side="left")
        self.dedup_combo.bind("<<ComboboxSelected>>", self._on_dedup_mode)
        self.imp_info = tk.StringVar(value="尚未导入名单")
        ttk.Label(imp, textvariable=self.imp_info, foreground="#555").pack(side="right")

        mapf = ttk.LabelFrame(self, text="③ 字段识别（自动猜测，可手动纠正；选“固定值”则在右侧填）", padding=8)
        mapf.pack(fill="x", **pad)
        self.map_inner = ttk.Frame(mapf)
        self.map_inner.pack(fill="x")
        self.map_widgets = {}

        sel = ttk.LabelFrame(self, text="④ 选择要群发的对象（单击行勾选/取消；只发勾上且可发送的）", padding=8)
        sel.pack(fill="both", expand=True, **pad)
        bar = ttk.Frame(sel)
        bar.pack(fill="x")
        ttk.Button(bar, text="全选", command=lambda: self._select_all(True)).pack(side="left")
        ttk.Button(bar, text="全不选", command=lambda: self._select_all(False)).pack(side="left", padx=4)
        ttk.Button(bar, text="只选可发送", command=self._select_ok).pack(side="left")
        ttk.Label(bar, text="   过滤:").pack(side="left")
        self.filter_var = tk.StringVar()
        self.filter_var.trace_add("write", lambda *a: self._render_tree())
        ttk.Entry(bar, textvariable=self.filter_var, width=14).pack(side="left")
        ttk.Label(bar, text="  分组:").pack(side="left")
        self.group_filter = tk.StringVar(value="全部")
        self.group_combo = ttk.Combobox(bar, textvariable=self.group_filter, values=["全部"],
                                        state="readonly", width=10)
        self.group_combo.pack(side="left")
        self.group_combo.bind("<<ComboboxSelected>>", lambda e: self._render_tree())
        ttk.Button(bar, text="给勾选设分组", command=self._set_group_for_selected).pack(side="left", padx=4)
        ttk.Button(bar, text="选中该分组", command=self._select_group).pack(side="left")
        ttk.Button(bar, text="清除分组", command=self._clear_group).pack(side="left", padx=2)
        ttk.Button(bar, text="删除勾选", command=self._delete_selected).pack(side="left", padx=2)
        self.sel_info = tk.StringVar(value="")
        ttk.Label(bar, textvariable=self.sel_info).pack(side="right")

        cols = ("sel", "email", "name", "group", "status")
        self.tree = ttk.Treeview(sel, columns=cols, show="headings", height=9)
        for c, t, w, anchor in [("sel", "选", 40, "center"), ("email", "邮箱", 260, "w"),
                                ("name", "姓名", 100, "w"), ("group", "分组", 90, "w"),
                                ("status", "状态", 180, "w")]:
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor=anchor)
        self.tree.pack(fill="both", expand=True, pady=4)
        self.tree.bind("<Button-1>", self._on_tree_click)
        for tag, bg in [("bad", "#ffe1e1"), ("off", "#eeeeee")]:
            self.tree.tag_configure(tag, background=bg)
        self._row_menu = tk.Menu(self, tearoff=0)
        self._row_menu.add_command(label="从名单删除此行", command=self._menu_delete_row)
        self._menu_idx = None
        self.tree.bind("<Button-3>", self._on_rightclick)

        snd = ttk.LabelFrame(self, text="⑤ 预览与发送", padding=8)
        snd.pack(fill="x", **pad)
        row0 = ttk.Frame(snd); row0.pack(fill="x")
        ttk.Label(row0, text="主题:").pack(side="left")
        self.subject_var = tk.StringVar()
        ttk.Entry(row0, textvariable=self.subject_var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row0, text="预览前3封", command=self._preview).pack(side="left")
        ttk.Button(row0, text="编辑模板…", command=self.edit_template).pack(side="left", padx=6)
        row1 = ttk.Frame(snd); row1.pack(fill="x", pady=6)
        self.dry_run = tk.BooleanVar(value=True)
        self.skip_dup = tk.BooleanVar(value=True)
        ttk.Checkbutton(row1, text="试运行（不实际发信）", variable=self.dry_run).pack(side="left")
        ttk.Checkbutton(row1, text="跳过已发过(去重)", variable=self.skip_dup).pack(side="left", padx=10)
        self.send_btn = ttk.Button(row1, text="开始群发", command=self._start_send)
        self.send_btn.pack(side="left", padx=6)
        self.stop_btn = ttk.Button(row1, text="停止", command=self._stop_send, state="disabled")
        self.stop_btn.pack(side="left")
        self.pb = ttk.Progressbar(row1, length=220, mode="determinate")
        self.pb.pack(side="left", padx=10)
        self.status = tk.StringVar(value="就绪")
        ttk.Label(row1, textvariable=self.status).pack(side="right")
        row2 = ttk.Frame(snd); row2.pack(fill="x", pady=(2, 0))
        ttk.Button(row2, text="重发失败项", command=self.resend_failed).pack(side="left")
        ttk.Button(row2, text="发送记录…", command=self.show_log).pack(side="left", padx=6)
        ttk.Label(row2, text="（重发失败项：从发送记录里挑出失败的邮箱，用当前名单/模板重发；跳过“已发过”限制）",
                  foreground="#777").pack(side="left", padx=6)
        row3 = ttk.Frame(snd); row3.pack(fill="x", pady=(2, 0))
        ttk.Button(row3, text="导出给命令行 (recipients.csv + config.json)", command=self._export_cli).pack(side="left")
        ttk.Label(row3, text="（面板调好名单/账号/字段识别后点这里，会在文件夹根目录生成命令行可用的两个文件，之后用 py mailer.py --send 跑，内容与面板一致）",
                  foreground="#777").pack(side="left", padx=6)

        sch = ttk.LabelFrame(self, text="⑥ 定时发送", padding=8)
        sch.pack(fill="x", **pad)
        rows = ttk.Frame(sch); rows.pack(fill="x")
        ttk.Label(rows, text="发送时间:").pack(side="left")
        self.sched_date = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.sched_time = tk.StringVar(value=(datetime.now() + timedelta(minutes=5)).strftime("%H:%M"))
        ttk.Entry(rows, textvariable=self.sched_date, width=11).pack(side="left", padx=(4, 2))
        ttk.Entry(rows, textvariable=self.sched_time, width=6).pack(side="left")
        ttk.Label(rows, text="(YYYY-MM-DD  HH:MM)").pack(side="left", padx=4)
        ttk.Button(rows, text="启用定时", command=self._enable_sched).pack(side="left", padx=6)
        ttk.Button(rows, text="取消定时", command=self._cancel_sched).pack(side="left")
        self.sched_status = tk.StringVar(value="未设定定时发送")
        ttk.Label(rows, textvariable=self.sched_status, foreground="#06c").pack(side="left", padx=10)
        ttk.Label(sch, text="说明：定时发送需保持面板开启、电脑不休眠；到点会按④当前勾选自动真实发送（跳过确认、忽略“试运行”）。",
                  foreground="#777").pack(anchor="w", pady=(4, 0))

        self.log = tk.Text(self, height=7, wrap="word", state="disabled")
        self.log.pack(fill="x", **pad)

    # ---------------- 账号 ----------------
    def _load_accounts(self):
        try:
            self.profiles = pc.load_accounts()
        except Exception as e:
            self.profiles = [pc.default_profile()]
            self._log(f"读取 accounts.json 失败：{e}")
        names = [p.get("名称", f"账号{i+1}") for i, p in enumerate(self.profiles)]
        self.acc_combo["values"] = names
        if names:
            self.acc_combo.current(0)
            self._on_account()

    def _cur(self):
        i = self.acc_combo.current()
        if 0 <= i < len(self.profiles):
            return self.profiles[i]
        return None

    def _on_account(self):
        p = self._cur()
        if not p:
            return
        chan = p.get("通道", "smtp")
        self.acc_info.set(f"通道={chan}  发信={p.get('发信人地址','')}  "
                          f"{'服务器='+p.get('SMTP服务器','') if chan=='smtp' else 'SES模板ID='+str(p.get('腾讯云SES',{}).get('模板ID',''))}")
        self.subject_var.set(p.get("邮件主题", ""))
        self._rebuild_map()

    def _on_dedup_mode(self, *_):
        self.dedup_mode.set(self._dedup_disp.get(self.dedup_combo.get(), "email"))

    def _manage_accounts(self):
        AccountManager(self, on_change=self._load_accounts)

    # ---------------- 导入 ----------------
    def _after_import(self, headers, records):
        if not records:
            messagebox.showinfo("提示", "没有解析到任何记录")
            return
        html = pc.load_template()
        subject = self.subject_var.get() or (self._cur() or {}).get("邮件主题", "")
        email_col, var_map = pc.detect_mapping(html, subject, headers, records)
        dups = pc.find_dup_groups(records, email_col, self.dedup_mode.get())
        removed = 0
        if dups:
            keep = self._resolve_dups(records, email_col, dups)
            if keep is None:
                self._log("已取消：重复名单处理时取消，本次导入未生效。")
                return
            before = len(records)
            records = [rec for i, rec in enumerate(records) if i in keep]
            removed = before - len(records)
        self.headers, self.records = headers, records
        self.email_col, self.var_map = email_col, var_map
        self._auto_group_from_columns()
        self._rebuild_map()
        self.imp_info.set(f"已导入 {len(records)} 行 · 列: {', '.join(headers[:8])}")
        self._recompute()
        self._refresh_group_combo()
        ok = sum(1 for it in self.items if it["ok"])
        self._log(f"导入完成：{len(self.items)} 条记录，其中 {ok} 条可直接发送。")
        if dups:
            self._log(f"检测到 {len(dups)} 个重复邮箱，已按你的选择保留 {len(dups)} 条、去掉 {removed} 条重复。")
        auto = [v for v, src in self.var_map.items() if src[0] != "col"]
        if auto:
            self._log("⚠ 名单里缺少这些列，已用占位值自动填充："
                      + "、".join("{{%s}}" % v for v in auto)
                      + "。要个性化请到③改成“固定值/用邮箱前缀”，或在②的编辑器里加列。")

    def _resolve_dups(self, records, email_col, dups):
        dlg = DupResolveDialog(self, records, email_col, dups)
        self.wait_window(dlg)
        return dlg.keep

    def _imp_csv(self):
        p = filedialog.askopenfilename(filetypes=[("CSV", "*.csv"), ("所有文件", "*.*")])
        if p:
            try:
                self._after_import(*pc.parse_csv(p))
            except Exception as e:
                messagebox.showerror("导入失败", str(e))

    def _imp_xlsx(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx"), ("所有文件", "*.*")])
        if p:
            try:
                self._after_import(*pc.parse_xlsx(p, self.header_is_title.get()))
            except Exception as e:
                messagebox.showerror("导入失败", f"{e}\n（Excel 需要 openpyxl，若报缺库：py -m pip install openpyxl）")

    def _imp_paste(self):
        initial = (self.headers, self.records) if self.records else None
        dlg = TableEditorDialog(self, initial=initial)
        self.wait_window(dlg)
        if dlg.result is not None:
            headers, records = dlg.result
            if records:
                self._after_import(headers, records)
            else:
                messagebox.showinfo("提示", "表格里没有数据行")

    # ---------------- 拖放导入 ----------------
    def _enable_dnd(self):
        if not _HAVE_DND:
            self._log("提示：未安装 tkinterdnd2，文件拖拽导入不可用（用“导入”按钮即可）。启用：py -m pip install tkinterdnd2")
            return
        targets = [self, getattr(self, "imp_frame", None), getattr(self, "tree", None)]
        try:
            for w in targets:
                if w is None:
                    continue
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self._on_drop_files)
            self._log("文件拖拽已启用：把 .csv / .xlsx / .txt 名单拖到窗口任意位置即可导入（支持一次拖多个）。")
        except Exception as e:
            self._log("拖拽注册失败（仍可用按钮导入）：%s" % e)

    def _on_drop_files(self, event):
        paths = [p for p in _split_tcl_list(event.data) if p]
        files = [p for p in paths if os.path.isfile(p)]
        if not files:
            self._log("拖入的不是可识别的名单文件，已忽略。")
            return
        self._import_files(files)

    def _import_files(self, files):
        parsed = []
        for f in files:
            ext = os.path.splitext(f)[1].lower()
            try:
                if ext == ".csv":
                    h, r = pc.parse_csv(f)
                elif ext in (".xlsx", ".xls"):
                    h, r = pc.parse_xlsx(f, self.header_is_title.get())
                elif ext == ".txt":
                    with open(f, encoding="utf-8", errors="ignore") as fh:
                        h, r = pc.parse_pasted(fh.read())
                else:
                    self._log("忽略不支持的文件类型：%s" % os.path.basename(f))
                    continue
            except Exception as e:
                self._log("读取失败 %s：%s" % (os.path.basename(f), e))
                continue
            if r:
                parsed.append((h, r, os.path.basename(f)))
        if not parsed:
            messagebox.showwarning("拖放", "拖入的文件里没有可导入的名单")
            return
        if len(parsed) == 1:
            h, r, name = parsed[0]
            self._after_import(h, r)
            self._log("已拖入并导入：%s" % name)
            return
        base_h = parsed[0][0]
        merged = list(parsed[0][1])
        names = [parsed[0][2]]
        for h, r, n in parsed[1:]:
            for rec in r:
                merged.append({k: rec.get(k, "") for k in base_h})
            names.append(n)
        self._after_import(base_h, merged)
        self._log("已拖入并合并 %d 个文件：%s" % (len(parsed), "、".join(names)))

    def _clear_list(self):
        self.headers, self.records, self.items, self.selected = [], [], [], set()
        self._recompute()
        self.imp_info.set("尚未导入名单")

    # ---------------- 列映射面板 ----------------
    def _rebuild_map(self):
        for w in self.map_inner.winfo_children():
            w.destroy()
        self.map_widgets = {}
        if not self.headers:
            ttk.Label(self.map_inner, text="（先导入名单，这里会自动出现需要映射的字段）").grid(row=0, column=0)
            return
        html = pc.load_template()
        subject = self.subject_var.get()
        vars_ = pc.needed_vars(html, subject)

        # 收件人邮箱列
        ttk.Label(self.map_inner, text="收件人邮箱列 →").grid(row=0, column=0, sticky="w", padx=4, pady=2)
        opts = ["列:" + h for h in self.headers]
        ec = ttk.Combobox(self.map_inner, values=opts, state="readonly", width=24)
        if self.email_col and ("列:" + self.email_col) in opts:
            ec.set("列:" + self.email_col)
        elif opts:
            ec.set(opts[0])
        ec.grid(row=0, column=1, sticky="w", padx=4, pady=2)
        ec.bind("<<ComboboxSelected>>", lambda e: self._recompute())
        self.map_widgets["__email__"] = (ec, None)

        for i, var in enumerate(vars_, start=1):
            ttk.Label(self.map_inner, text="{{%s}} 取自 →" % var).grid(row=i, column=0, sticky="w", padx=4, pady=2)
            opts2 = ["列:" + h for h in self.headers] + [LOCAL_OPT, FIXED_OPT]
            cb = ttk.Combobox(self.map_inner, values=opts2, state="readonly", width=24)
            ent = ttk.Entry(self.map_inner, width=22)
            cur = self.var_map.get(var)
            if cur and cur[0] == "col" and ("列:" + cur[1]) in opts2:
                cb.set("列:" + cur[1])
            elif cur and cur[0] == "fixed":
                cb.set(FIXED_OPT)
                ent.delete(0, "end"); ent.insert(0, cur[1])
            elif cur and cur[0] == "emaillocal":
                cb.set(LOCAL_OPT)
            else:
                cb.set(opts2[0])
            cb.grid(row=i, column=1, sticky="w", padx=4, pady=2)
            ent.grid(row=i, column=2, sticky="w", padx=4, pady=2)
            cb.bind("<<ComboboxSelected>>", lambda e, v=var: self._recompute())
            ent.bind("<KeyRelease>", lambda e, v=var: self._recompute())
            self.map_widgets[var] = (cb, ent)

    def _read_varmap(self):
        email_col = None
        ec, _ = self.map_widgets.get("__email__", (None, None))
        if ec:
            v = ec.get()
            email_col = v[2:] if v.startswith("列:") else None
        var_map = {}
        for var, (cb, ent) in self.map_widgets.items():
            if var == "__email__":
                continue
            v = cb.get()
            if v == FIXED_OPT:
                var_map[var] = ("fixed", ent.get())
            elif v == LOCAL_OPT:
                var_map[var] = ("emaillocal",)
            elif v.startswith("列:"):
                var_map[var] = ("col", v[2:])
            else:
                var_map[var] = ("fixed", "")
        return email_col, var_map

    def _recompute(self):
        if not self.records:
            self.items = []
            self._render_tree()
            return
        self.email_col, self.var_map = self._read_varmap()
        html = pc.load_template()
        subject = self.subject_var.get()
        self.items = pc.build_sendable(self.records, self.email_col, self.var_map, html, subject)
        for i, it in enumerate(self.items):
            if it["ok"]:
                self.selected.add(i)
            else:
                self.selected.discard(i)
        self._render_tree()

    # ---------------- 收件人表格 ----------------
    def _render_tree(self):
        self.tree.delete(*self.tree.get_children())
        filt = self.filter_var.get().strip().lower()
        gsel = getattr(self, "group_filter", None)
        gfilter = gsel.get() if gsel else "全部"
        for i, it in enumerate(self.items):
            if filt and filt not in it["email"].lower() and filt not in it["row"].get("姓名", "").lower():
                continue
            grp = self.group_of.get(i, "")
            if gfilter != "全部" and grp != gfilter:
                continue
            mark = "☑" if i in self.selected else "☐"
            name = it["row"].get("姓名", "")
            status = "" if it["ok"] else it["err"]
            tags = () if it["ok"] else ("bad",)
            self.tree.insert("", "end", iid=str(i),
                             values=(mark, it["email"], name, grp, status), tags=tags)
        ok = sum(1 for x in self.items if x["ok"])
        self.sel_info.set("共 %d · 可发送 %d · 已选 %d%s" % (
            len(self.items), ok, len(self.selected),
            (" · 分组 %d 个" % len(self._group_names())) if self.group_of else ""))

    def _on_tree_click(self, ev):
        iid = self.tree.identify_row(ev.y)
        if not iid:
            return
        idx = int(iid)
        if idx in self.selected:
            self.selected.discard(idx)
        else:
            if self.items[idx]["ok"]:
                self.selected.add(idx)
            else:
                messagebox.showinfo("无法选择", self.items[idx]["err"])
                return
        self._render_tree()

    def _select_all(self, val):
        self.selected = {i for i, it in enumerate(self.items) if val and it["ok"]}
        self._render_tree()

    def _select_ok(self):
        self._select_all(True)

    # ---------------- 收件人分组 / 标签 ----------------
    def _group_names(self):
        return sorted({g for g in self.group_of.values() if g})

    def _refresh_group_combo(self):
        vals = ["全部"] + self._group_names()
        self.group_combo["values"] = vals
        if self.group_filter.get() not in vals:
            self.group_filter.set("全部")

    def _auto_group_from_columns(self):
        """导入后：若名单里有“分组/标签/组/类别/group/tag”这类列，用它自动填充每个收件人的分组。"""
        self.group_of = {}
        gcol = None
        for h in self.headers:
            hl = str(h).lower()
            if h in ("分组", "标签", "组", "类别", "分类") or "group" in hl or "tag" in hl or h == "组别":
                gcol = h
                break
        if not gcol:
            return
        for i, rec in enumerate(self.records):
            g = str(rec.get(gcol, "")).strip()
            if g:
                self.group_of[i] = g
        if self.group_of:
            self._log("已按“%s”列自动分组，共 %d 个分组。" % (gcol, len(self._group_names())))

    def _set_group_for_selected(self):
        if not self.selected:
            messagebox.showinfo("分组", "请先在④勾选要设分组的收件人")
            return
        name = simpledialog.askstring("设分组", "给选中的 %d 个收件人设置分组名（留空=清除其分组）："
                                      % len(self.selected), parent=self)
        if name is None:
            return
        name = name.strip()
        for i in list(self.selected):
            if name:
                self.group_of[i] = name
            else:
                self.group_of.pop(i, None)
        self._refresh_group_combo()
        self._render_tree()
        self._log("已为 %d 个收件人%s分组%s。"
                  % (len(self.selected), "设置“%s”" % name if name else "清除", ""))

    def _select_group(self):
        g = self.group_filter.get()
        if g == "全部":
            messagebox.showinfo("分组", "请先在“分组”下拉里选一个具体分组")
            return
        self.selected = {i for i, it in enumerate(self.items)
                         if it["ok"] and self.group_of.get(i) == g}
        self._render_tree()

    def _clear_group(self):
        g = self.group_filter.get()
        if g == "全部":
            self.group_of = {}
            self._log("已清空所有分组")
        else:
            for i in [i for i, v in self.group_of.items() if v == g]:
                del self.group_of[i]
            self._log("已清除分组“%s”" % g)
        self._refresh_group_combo()
        self._render_tree()

    # ---------------- 删除收件人 ----------------
    def _remove_indices(self, idxset):
        idxset = {i for i in idxset if 0 <= i < len(self.records)}
        if not idxset:
            return
        keep_emails = {self.items[i]["email"] for i in self.selected
                       if i not in idxset and i < len(self.items)}
        keep = [i for i in range(len(self.records)) if i not in idxset]
        self.records = [self.records[i] for i in keep]
        ng = {}
        for newpos, old in enumerate(keep):
            if old in self.group_of:
                ng[newpos] = self.group_of[old]
        self.group_of = ng
        self.selected = set()
        self._recompute()
        self.selected = {i for i, it in enumerate(self.items)
                         if it["ok"] and it["email"] in keep_emails}
        self._refresh_group_combo()
        self._render_tree()

    def _delete_selected(self):
        if not self.selected:
            messagebox.showinfo("删除", "请先在④勾选(☑)要删除的收件人")
            return
        n = len(self.selected)
        if not messagebox.askyesno(
                "删除", "将从名单中移除当前勾选的 %d 个收件人（这是“删除”，不是取消勾选）。确定？" % n):
            return
        self._remove_indices(set(self.selected))
        self._log("已从名单删除 %d 个收件人。" % n)

    def _on_rightclick(self, ev):
        iid = self.tree.identify_row(ev.y)
        if not iid:
            return
        self._menu_idx = int(iid)
        try:
            self._row_menu.tk_popup(ev.x_root, ev.y_root)
        finally:
            self._row_menu.grab_release()

    def _menu_delete_row(self):
        idx = self._menu_idx
        if idx is None or idx >= len(self.records):
            return
        self._remove_indices({idx})
        self._menu_idx = None
        self._log("已删除 1 个收件人。")

    # ---------------- 预览 ----------------
    def _preview(self):
        ready = [it for i, it in enumerate(self.items) if it["ok"] and i in self.selected] or \
                [it for it in self.items if it["ok"]]
        if not ready:
            messagebox.showinfo("提示", "没有可预览的记录（先导入并勾选）")
            return
        html = pc.load_template()
        subj_tpl = self.subject_var.get()
        txt = []
        for it in ready[:3]:
            subj = mailer.render(subj_tpl, it["row"]) or subj_tpl
            body = mailer.render(html, it["row"]) or html
            txt.append(f"收件人：{it['email']}\n主题：{subj}\n正文：\n{mailer.html_to_text(body)}\n" + "-" * 50)
        _TextPopup(self, "预览（前 3 封）", "\n\n".join(txt))

    # ---------------- 模板编辑 ----------------
    def edit_template(self):
        dlg = TemplateEditorDialog(self)
        self.wait_window(dlg)

    def reload_template(self):
        if self.records:
            html = pc.load_template()
            subject = self.subject_var.get()
            self.email_col, self.var_map = pc.detect_mapping(html, subject, self.headers, self.records)
            self._rebuild_map()
            self._recompute()
        else:
            self._rebuild_map()
        self._log("模板已更新，字段识别与预览已按新模板刷新。")

    # ---------------- 发送 ----------------
    def _start_send(self):
        p = self._cur()
        if not p:
            messagebox.showwarning("提示", "请先选择发信账号")
            return
        send_pairs = [(i, self.items[i]) for i in sorted(self.selected) if self.items[i]["ok"]]
        if not send_pairs:
            messagebox.showwarning("提示", "没有勾选任何可发送的收件人")
            return
        self._launch_send(send_pairs, skip_dup=self.skip_dup.get())

    def _launch_send(self, send_pairs, skip_dup, confirm=True, force_real=False):
        p = self._cur()
        if not p:
            messagebox.showwarning("提示", "请先选择发信账号")
            return
        dry = self.dry_run.get() and not force_real
        if confirm and not dry and not messagebox.askyesno(
                "确认群发", f"将通过【{p.get('名称')}】向 {len(send_pairs)} 位收件人真实发送，确认？"):
            return
        self._send_order = [i for i, _ in send_pairs]
        items = [it for _, it in send_pairs]
        prof = dict(p)
        prof["邮件主题"] = self.subject_var.get() or prof.get("邮件主题", "")
        self.sender = pc.Sender(
            prof, items, skip_sent=skip_dup, dry_run=dry,
            on_progress=lambda i, t: self.q.put(("prog", i, t)),
            on_status=lambda pos, st: self.q.put(("status", pos, st)),
            on_log=lambda m: self.q.put(("log", m)),
            on_done=lambda ok, fail, skip: self.q.put(("done", ok, fail, skip)),
        )
        self.send_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.status.set("发送中…")
        threading.Thread(target=self.sender.run, daemon=True).start()

    def resend_failed(self):
        rows = pc.load_log_rows()
        fails = {(r.get("收件人") or "").strip().lower()
                 for r in rows if (r.get("状态") or "").startswith("失败")}
        if not fails:
            messagebox.showinfo("重发失败项", "发送记录里没有失败项。")
            return
        if not self.items:
            messagebox.showwarning(
                "重发失败项",
                "当前没有已导入并映射的名单，无法按当前模板重发。\n请先导入原来那份名单（失败邮箱在其中的行会被重新发送）。")
            return
        pairs = [(i, it) for i, it in enumerate(self.items)
                 if it["ok"] and it["email"].strip().lower() in fails]
        if not pairs:
            messagebox.showinfo(
                "重发失败项",
                "这些失败邮箱不在当前名单里（名单可能已更换），无法用当前映射重发。\n"
                "失败邮箱示例：" + "、".join(list(fails)[:5]))
            return
        self._log("开始重发 %d 个失败收件人…" % len(pairs))
        self._launch_send(pairs, skip_dup=False)

    def show_log(self):
        dlg = SendLogViewer(self)
        self.wait_window(dlg)

    # ---------------- 导出给命令行 ----------------
    def _export_cli(self):
        sendable = [it for it in self.items if it["ok"]]
        if not sendable:
            messagebox.showwarning("导出", "当前没有可发送的行（先导入名单并在③把变量补齐）。")
            return
        p = self._cur()
        if not p:
            messagebox.showwarning("导出", "请先选择发信账号。")
            return
        if not messagebox.askyesno(
                "导出给命令行",
                "将在面板文件夹根目录生成/覆盖两个文件：\n"
                "  · recipients.csv —— 已解析的可发送名单（%d 行，含邮箱与各模板变量列的实际值）\n"
                "  · config.json    —— 当前账号【%s】的命令行配置\n"
                "之后可在命令行运行：py mailer.py --send\n"
                "（会覆盖同名旧文件）继续？" % (len(sendable), p.get("名称", ""))):
            return
        try:
            rp, n = pc.export_recipients_csv(
                self.items, pc.load_template(), self.subject_var.get(),
                os.path.join(pc.BASE_DIR, "recipients.csv"))
            cp = pc.export_config_json(p, os.path.join(pc.BASE_DIR, "config.json"))
            self._log("已导出 recipients.csv（%d 行）+ config.json 到 %s。命令行：py mailer.py --send" % (n, pc.BASE_DIR))
            messagebox.showinfo(
                "导出完成",
                "已生成两个文件（同文件夹）：\n  recipients.csv（%d 行）\n  config.json（账号：%s）\n\n"
                "命令行用法（在本文件夹开 cmd）：\n"
                "  py mailer.py --preview 3\n"
                "  py mailer.py --test 你的邮箱\n"
                "  py mailer.py --send" % (n, p.get("名称", "")))
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    # ---------------- 定时发送 ----------------
    def _enable_sched(self):
        self._cancel_sched(silent=True)
        s = (self.sched_date.get().strip() + " " + self.sched_time.get().strip())
        try:
            when = datetime.strptime(s, "%Y-%m-%d %H:%M")
        except ValueError:
            messagebox.showwarning("定时", "时间格式应为 YYYY-MM-DD  HH:MM")
            return
        if when <= datetime.now():
            messagebox.showwarning("定时", "设定的时间已过，请填一个将来的时间")
            return
        self.scheduled_at = when
        self.sched_status.set("已设定：%s 到点自动发送" % when.strftime("%Y-%m-%d %H:%M"))
        self._log("定时发送已设定：%s。请保持面板开启、电脑不休眠。" % s)
        self._schedule_tick()

    def _cancel_sched(self, silent=False):
        if self._sched_job:
            try:
                self.after_cancel(self._sched_job)
            except Exception:
                pass
            self._sched_job = None
        self.scheduled_at = None
        if not silent:
            self.sched_status.set("未设定定时发送")
            self._log("定时发送已取消。")

    def _schedule_tick(self):
        self._sched_job = self.after(15000, self._sched_check)

    def _sched_check(self):
        self._sched_job = None
        if not self.scheduled_at:
            return
        now = datetime.now()
        if now >= self.scheduled_at:
            self._fire_scheduled()
        else:
            remain = int((self.scheduled_at - now).total_seconds() // 60)
            self.sched_status.set("定时：%s（还需约 %d 分钟）" % (self.scheduled_at.strftime("%H:%M"), remain))
            self._schedule_tick()

    def _fire_scheduled(self):
        self.scheduled_at = None
        pairs = [(i, self.items[i]) for i in sorted(self.selected)
                 if i < len(self.items) and self.items[i]["ok"]]
        if not pairs:
            self.sched_status.set("定时到点，但无勾选可发送项，已取消")
            self._log("定时发送触发，但没有勾选可发送的收件人，已跳过。")
            return
        self._log("⏰ 定时发送触发：将向 %d 位收件人真实发送…" % len(pairs))
        self.sched_status.set("定时已触发，发送中…")
        self._launch_send(pairs, skip_dup=self.skip_dup.get(), confirm=False, force_real=True)

    def _stop_send(self):
        if self.sender:
            self.sender.cancel = True
            self._log("正在停止…（等当前这封发完）")

    def _poll(self):
        try:
            while True:
                msg = self.q.get_nowait()
                kind = msg[0]
                if kind == "prog":
                    _, i, t = msg
                    self.pb["maximum"] = t
                    self.pb["value"] = i
                    self.status.set(f"进度 {i}/{t}")
                elif kind == "status":
                    _, pos, st = msg
                    orig = self._send_order[pos] if pos < len(self._send_order) else None
                    if orig is not None and self.tree.exists(str(orig)):
                        vals = list(self.tree.item(str(orig), "values"))
                        vals[4] = st
                        self.tree.item(str(orig), values=vals,
                                       tags=() if st == "成功" else ("bad" if st.startswith("失败") else ("off",)))
                elif kind == "log":
                    self._log(msg[1])
                elif kind == "done":
                    _, ok, fail, skip = msg
                    self.status.set(f"完成：成功 {ok}，失败 {fail}，跳过 {skip}")
                    self.send_btn.config(state="normal")
                    self.stop_btn.config(state="disabled")
                    self._log(f"===== 发送结束：成功 {ok} / 失败 {fail} / 跳过 {skip} =====")
                    if fail:
                        self._log("提示：有 %d 封失败，可用「重发失败项」按当前名单重发，或点「发送记录…」查看明细。" % fail)
        except queue.Empty:
            pass
        self.after(120, self._poll)

    def _log(self, s):
        self.log.config(state="normal")
        self.log.insert("end", s + "\n")
        self.log.see("end")
        self.log.config(state="disabled")


class AccountManager(tk.Toplevel):
    def __init__(self, master, on_change=None):
        super().__init__(master)
        self.master_ = master
        self.on_change = on_change
        self.title("发信账号管理")
        self.geometry("520x420")
        top = ttk.Frame(self, padding=10); top.pack(fill="both", expand=True)
        self.lb = tk.Listbox(top, width=52)
        self.lb.pack(side="left", fill="both", expand=True)
        side = ttk.Frame(top); side.pack(side="left", fill="y", padx=8)
        for text, cmd in [("新增 SMTP", lambda: self._edit(-1, ses=False)),
                          ("新增 腾讯SES", lambda: self._edit(-1, ses=True)),
                          ("编辑选中", lambda: self._edit(self._idx())),
                          ("复制选中", self._dup),
                          ("删除选中", self._del),
                          ("关闭", self.destroy)]:
            ttk.Button(side, text=text, command=cmd).pack(fill="x", pady=3)
        self._refresh()

    def _profiles(self):
        return self.master_.profiles

    def _refresh(self):
        self.lb.delete(0, "end")
        for p in self._profiles():
            self.lb.insert("end", f"[{p.get('通道','?')}] {p.get('名称','未命名')}")
        if self.on_change:
            self.on_change()

    def _idx(self):
        s = self.lb.curselection()
        return s[0] if s else -1

    def _edit(self, i, ses=False):
        prof = None if i < 0 else dict(self._profiles()[i])
        f = AccountForm(self, prof, ses=ses)
        self.wait_window(f)
        if f.result:
            ps = self._profiles()
            if i < 0:
                ps.append(f.result)
            else:
                ps[i] = f.result
            pc.save_accounts(ps)
            self._refresh()

    def _dup(self):
        i = self._idx()
        if i < 0:
            return
        ps = self._profiles()
        c = dict(ps[i]); c["名称"] = c.get("名称", "") + "（副本）"
        ps.insert(i + 1, c)
        pc.save_accounts(ps); self._refresh()

    def _del(self):
        i = self._idx()
        if i < 0:
            return
        if not messagebox.askyesno("确认", "删除该账号？", parent=self):
            return
        ps = self._profiles(); del ps[i]
        pc.save_accounts(ps); self._refresh()


class TableEditorDialog(tk.Toplevel):
    """小型名单编辑器：粘贴→自动分列成可编辑表格，可增删行/列、双击改单元格、
    把首行设为表头。确定后 result=(headers, records)，交给主流程做列识别。"""

    def __init__(self, master, initial=None):
        super().__init__(master)
        self.result = None
        self.headers = []          # list[str]
        self.rows = []             # list[list[str]]
        if initial:
            ih, irec = initial
            self.headers = list(ih)
            self.rows = [[str(r.get(h, "")) for h in ih] for r in irec]
        self.active_col = None     # 形如 'c2'
        self._editor = None        # (Entry, row, coli)
        self.mode_var = getattr(master, "dedup_mode", None)  # 与主窗口共享去重依据
        if self.mode_var is None:
            self.mode_var = tk.StringVar(value="email")
        self.title("名单编辑器 · 当前名单（可继续补充/编辑，确定后更新）")
        self.geometry("900x560")
        self.transient(master)

        top = ttk.Frame(self, padding=(10, 8, 10, 0))
        top.pack(fill="x")
        ttk.Label(top, text="下面就是当前群发名单，可加行/删行/双击改；把新邮箱粘到下框再选“追加”或“覆盖”：").pack(side="left")
        self.header_is_title = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="首行作为表头(列名)", variable=self.header_is_title).pack(side="left", padx=10)

        self.txt = tk.Text(self, height=4, wrap="none")
        self.txt.pack(fill="x", padx=10, pady=(4, 6))

        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=10)
        for text, cmd in [("解析并覆盖", self._parse), ("解析并追加", self._append),
                          ("加行", self._add_row),
                          ("删选中行", self._del_row), ("加列", self._add_col),
                          ("删选中列", self._del_col), ("首行设为表头", self._promote_header),
                          ("一键去重", self._dedup)]:
            ttk.Button(bar, text=text, command=cmd).pack(side="left", padx=2)
        ttk.Label(bar, text="  去重依据:").pack(side="left")
        self._m_disp = {"按邮箱": "email", "按整行": "row"}
        self._m_rev = {v: k for k, v in self._m_disp.items()}
        self.mode_combo = ttk.Combobox(bar, values=list(self._m_disp.keys()), state="readonly", width=6)
        self.mode_combo.set(self._m_rev.get(self.mode_var.get(), "按邮箱"))
        self.mode_combo.pack(side="left")
        self.mode_combo.bind("<<ComboboxSelected>>", self._on_mode)
        self.dup_info = tk.StringVar(value="")
        ttk.Label(bar, textvariable=self.dup_info, foreground="#b00020").pack(side="left", padx=8)
        ttk.Label(bar, text="   双击单元格改内容；点列标题选中该列").pack(side="left")

        quick = ttk.Frame(self)
        quick.pack(fill="x", padx=10, pady=(2, 0))
        ttk.Label(quick, text="逐条添加:").pack(side="left")
        self.qfields = {}
        for label in ("邮箱", "姓名", "称呼", "订单号", "专属优惠码"):
            ttk.Label(quick, text=label).pack(side="left", padx=(6, 1))
            ent = ttk.Entry(quick, width=11)
            ent.pack(side="left")
            self.qfields[label] = ent
        ttk.Button(quick, text="添加→", command=self._quick_add).pack(side="left", padx=6)
        ttk.Label(quick, text="（邮箱必填，填完点“添加”会按列自动归位，不怕顺序）",
                  foreground="#777").pack(side="left")

        cf = ttk.Frame(self)
        cf.pack(fill="both", expand=True, padx=10, pady=6)
        self.tree = ttk.Treeview(cf, show="headings", selectmode="extended")
        vs = ttk.Scrollbar(cf, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(cf, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.tag_configure("dup", background="#ffe0e0")
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        cf.rowconfigure(0, weight=1)
        cf.columnconfigure(0, weight=1)
        self.tree.bind("<Double-1>", self._on_double)
        self.tree.bind("<Button-1>", self._on_click)

        btns = ttk.Frame(self)
        btns.pack(pady=8)
        ttk.Button(btns, text="确定", command=self._ok).pack(side="left", padx=8)
        ttk.Button(btns, text="取消", command=self.destroy).pack(side="left")

        self.update_idletasks()
        try:
            self.wait_visibility()
            self.grab_set()
        except tk.TclError:
            pass
        self.txt.focus_force()
        self._canonicalize()
        self._refresh()

    # ---- 渲染 ----
    def _col_ids(self):
        return ["c%d" % i for i in range(len(self.headers))]

    def _refresh(self):
        ids = self._col_ids()
        self.tree["columns"] = ids
        for i, cid in enumerate(ids):
            self.tree.heading(cid, text=self.headers[i])
            self.tree.column(cid, width=90 if len(ids) > 6 else 150, stretch=False, anchor="w")
        self.tree.delete(*self.tree.get_children())
        seen = set()
        dupn = 0
        for ri, row in enumerate(self.rows):
            tags = ()
            k = self._dup_key(row)
            if k:
                if k in seen:
                    tags = ("dup",)
                    dupn += 1
                else:
                    seen.add(k)
            self.tree.insert("", "end", iid=str(ri), values=row, tags=tags)
        if hasattr(self, "dup_info"):
            self.dup_info.set(("重复 %d 行(灰底)，可点“一键去重”" % dupn) if dupn else "")

    # ---- 解析粘贴文本 ----
    def _parse_text(self):
        raw = self.txt.get("1.0", "end").rstrip("\n")
        lines = [ln for ln in raw.splitlines() if ln.strip()]
        if not lines:
            messagebox.showinfo("提示", "先粘贴内容再点解析", parent=self)
            return None
        blk = pc.try_parse_block(raw)
        if blk:
            bh, brecs = blk
            bbody = [[str(r.get(h, "")) for h in bh] for r in brecs]
            return bh, bbody
        if any("\t" in ln for ln in lines):
            delim = "\t"
        elif sum(";" in ln for ln in lines) >= len(lines) * 0.5:
            delim = ";"
        elif sum("," in ln for ln in lines) >= len(lines) * 0.5:
            delim = ","
        else:
            delim = None  # 按空白切分
        grid = []
        for ln in lines:
            parts = ln.split() if delim is None else [p.strip() for p in ln.split(delim)]
            grid.append([p for p in parts])
        width = max(len(r) for r in grid)
        for r in grid:
            r += [""] * (width - len(r))
        if self.header_is_title.get():
            headers = [h if str(h).strip() else "第%d列" % (i + 1) for i, h in enumerate(grid[0])]
            body = grid[1:]
        else:
            headers = ["第%d列" % (i + 1) for i in range(width)]
            body = grid
            if width == 1 and all("@" in (r[0] or "") for r in body):
                headers = ["邮箱"]
        return headers, body

    def _parse(self):
        got = self._parse_text()
        if not got:
            return
        self.headers, self.rows = got
        self.active_col = None
        self._canonicalize()
        self._refresh()

    def _append(self):
        got = self._parse_text()
        if not got:
            return
        h2, body = got
        if not self.headers:
            self.headers, self.rows = h2, [list(r) for r in body]
            self._canonicalize()
            self._refresh()
            messagebox.showinfo("已加载", "已载入 %d 行。" % len(body), parent=self)
            return
        # 确定/建立“邮箱”列，之后每行按内容归位，绝不新增重复列
        em = self._email_col_index()
        if em is None:
            em = self._ensure_col("邮箱")
        elif self.headers[em] != "邮箱":
            self.headers[em] = "邮箱"
        width = len(self.headers)
        added = 0
        for cells in body:
            cells = [str(c).strip() for c in cells]
            ei = None
            for j, c in enumerate(cells):
                if "@" in c and "." in c.split("@")[-1]:
                    ei = j
                    break
            mapped = [""] * width
            if ei is not None:
                mapped[em] = cells[ei]
            others = [cells[j] for j in range(len(cells)) if j != ei]
            fill = [k for k in range(width) if k != em]
            for v, k in zip(others, fill):
                mapped[k] = v
            self.rows.append(mapped)
            added += 1
        self._canonicalize()
        self._refresh()
        messagebox.showinfo("追加完成",
                            "已追加 %d 行（邮箱自动进“邮箱”列，其余按顺序对齐，不怕输入顺序）。" % added,
                            parent=self)

    # ---- 按内容识别列并规范命名 ----
    def _name_col_index(self, exclude=None):
        for i, h in enumerate(self.headers):
            if i != exclude and h == "姓名":
                return i
        best, bs = None, 0.5
        for i in range(len(self.headers)):
            if i == exclude:
                continue
            vals = [str(r[i]).strip() for r in self.rows if i < len(r) and str(r[i]).strip()]
            if not vals:
                continue
            good = 0
            for v in vals:
                if "@" in v or v.isdigit() or v in ("先生", "女士", "小姐", "老师"):
                    continue
                if len(v) <= 12 and re.search(r"[A-Za-z\u4e00-\u9fff]", v) and not re.search(r"\d", v):
                    good += 1
            s = good / len(vals)
            if s > bs:
                best, bs = i, s
        return best

    def _canonicalize(self):
        if not self.headers:
            return
        em = self._email_col_index()
        if em is not None:
            self.headers[em] = "邮箱"
        nm = self._name_col_index(exclude=em)
        if nm is not None and nm != em:
            self.headers[nm] = "姓名"

    _KW = {"邮箱": ["mail", "email", "邮箱", "邮件"], "姓名": ["姓名", "名字", "name"],
           "称呼": ["称呼", "称谓", "性别"], "订单号": ["订单", "单号", "order"],
           "专属优惠码": ["优惠", "券", "码", "coupon", "promo"]}

    def _ensure_col(self, name):
        for i, h in enumerate(self.headers):
            if h == name:
                return i
        if name == "邮箱":
            ei = self._email_col_index()
            if ei is not None:
                if self.headers[ei] != "邮箱":
                    self.headers[ei] = "邮箱"
                return ei
        hl = {h: h.lower() for h in self.headers}
        for kw in self._KW.get(name, []):
            for i, h in enumerate(self.headers):
                if kw.lower() in hl[h]:
                    return i
        self.headers.append(name)
        for r in self.rows:
            r.append("")
        return len(self.headers) - 1

    def _quick_add(self):
        vals = {lab: self.qfields[lab].get().strip()
                for lab in ("邮箱", "姓名", "称呼", "订单号", "专属优惠码")}
        if not vals["邮箱"]:
            messagebox.showinfo("提示", "请先填写邮箱", parent=self)
            return
        if not self.headers:
            self.headers = ["邮箱", "姓名", "称呼", "订单号", "专属优惠码"]
        row = [""] * len(self.headers)
        for lab, v in vals.items():
            if not v:
                continue
            idx = self._ensure_col(lab)
            while len(row) <= idx:
                row.append("")
            row[idx] = v
        # _ensure_col 可能新增了列，补齐既有行长度
        for r in self.rows:
            while len(r) < len(self.headers):
                r.append("")
        self.rows.append(row)
        self._refresh()
        for e in self.qfields.values():
            e.delete(0, "end")

    # ---- 行 / 列 增删 ----
    def _add_row(self):
        self.rows.append([""] * len(self.headers))
        self._refresh()

    def _del_row(self):
        sel = {int(i) for i in self.tree.selection() if i.isdigit()}
        if not sel:
            messagebox.showinfo("提示", "先选中要删除的行", parent=self)
            return
        self.rows = [r for i, r in enumerate(self.rows) if i not in sel]
        self._refresh()

    def _add_col(self):
        name = simpledialog.askstring(
            "加列", "列名（用模板变量名最省事：姓名 / 称呼 / 订单号 / 专属优惠码）：", parent=self)
        if not name:
            return
        self.headers.append(name.strip())
        for r in self.rows:
            r.append("")
        self._refresh()

    def _del_col(self):
        if self.active_col is None:
            messagebox.showinfo("提示", "先点一个列标题，选中要删除的列", parent=self)
            return
        ci = int(self.active_col[1:])
        if len(self.headers) <= 1:
            messagebox.showinfo("提示", "至少保留一列", parent=self)
            return
        self.headers.pop(ci)
        for r in self.rows:
            if len(r) > ci:
                r.pop(ci)
        self.active_col = None
        self._refresh()

    def _promote_header(self):
        if not self.rows:
            return
        if not messagebox.askyesno("确认", "把当前第一行数据设为列名(表头)？", parent=self):
            return
        self.headers = [h if str(h).strip() else "第%d列" % (i + 1) for i, h in enumerate(self.rows[0])]
        self.rows = self.rows[1:]
        self._refresh()

    # ---- 去重依据 / 邮箱列定位 / 一键去重 ----
    def _mode(self):
        return self.mode_var.get() if self.mode_var else "email"

    def _on_mode(self, *_):
        self.mode_var.set(self._m_disp.get(self.mode_combo.get(), "email"))
        self._refresh()

    def _email_col_index(self):
        if not self.rows or not self.headers:
            return None
        for ci, h in enumerate(self.headers):
            hl = str(h).lower()
            if "邮" in h or "mail" in hl or "email" in hl:
                return ci
        best, bs = None, 0.3
        for ci in range(len(self.headers)):
            vals = [str(r[ci]).strip() for r in self.rows if ci < len(r) and str(r[ci]).strip()]
            if not vals:
                continue
            ratio = sum(1 for v in vals if "@" in v and "." in v.split("@")[-1]) / len(vals)
            if ratio > bs:
                best, bs = ci, ratio
        return best

    def _dup_key(self, row):
        if self._mode() == "row":
            k = "||".join(str(v).strip().lower() for v in row)
            return k if k.replace("|", "").strip() else None
        ci = self._email_col_index()
        e = str(row[ci]).strip().lower() if (ci is not None and ci < len(row)) else ""
        return e if "@" in e else None

    def _dedup(self):
        seen, drop = set(), []
        for ri, r in enumerate(self.rows):
            k = self._dup_key(r)
            if k:
                if k in seen:
                    drop.append(ri)
                else:
                    seen.add(k)
        label = "完全相同的整行" if self._mode() == "row" else "重复邮箱"
        if not drop:
            messagebox.showinfo("提示", "没有%s" % label.replace("完全相同的整行", "完全相同的行"), parent=self)
            return
        if not messagebox.askyesno("确认", "发现 %d 行%s，保留每组第一条并删除其余？" % (len(drop), label), parent=self):
            return
        dset = set(drop)
        self.rows = [r for i, r in enumerate(self.rows) if i not in dset]
        self._refresh()

    # ---- 单元格就地编辑 ----
    def _on_click(self, ev):
        self._close_editor()
        if self.tree.identify_region(ev.x, ev.y) == "heading":
            cid = self.tree.identify_column(ev.x)
            if cid in self._col_ids():
                self.active_col = cid

    def _on_double(self, ev):
        if self.tree.identify_region(ev.x, ev.y) != "cell":
            return
        iid = self.tree.identify_row(ev.y)
        col = self.tree.identify_column(ev.x)
        if iid and col in self._col_ids():
            self._edit_cell(iid, col)

    def _edit_cell(self, iid, col):
        self._close_editor()
        r = int(iid)
        ci = int(col[1:])
        bb = self.tree.bbox(iid, col)
        if not bb:
            return
        x, y, w, h = bb
        e = tk.Entry(self.tree, highlightthickness=1, relief="solid")
        e.insert(0, self.rows[r][ci] if ci < len(self.rows[r]) else "")
        e.place(x=x, y=y, width=w, height=h)
        e.focus_set()
        e.select_range(0, "end")
        self._editor = (e, r, ci)
        e.bind("<Return>", lambda ev: self._commit())
        e.bind("<Escape>", lambda ev: self._close_editor())
        e.bind("<FocusOut>", lambda ev: self._commit())

    def _commit(self):
        if not self._editor:
            return
        e, r, ci = self._editor
        val = e.get()
        self._editor = None
        e.destroy()
        if 0 <= r < len(self.rows):
            while len(self.rows[r]) <= ci:
                self.rows[r].append("")
            self.rows[r][ci] = val
            self.tree.item(str(r), values=self.rows[r])

    def _close_editor(self):
        if self._editor:
            try:
                self._editor[0].destroy()
            except Exception:
                pass
            self._editor = None

    # ---- 输出 ----
    def _ok(self):
        self._close_editor()
        headers = [h if str(h).strip() else "第%d列" % (i + 1) for i, h in enumerate(self.headers)]
        self.headers = headers
        recs = []
        for row in self.rows:
            if not any(str(v).strip() for v in row):
                continue
            recs.append({headers[i]: (str(row[i]).strip() if i < len(row) else "")
                         for i in range(len(headers))})
        self.result = (headers, recs)
        self.destroy()


class DupResolveDialog(tk.Toplevel):
    """导入名单里出现重复邮箱时，逐组让用户挑选保留哪一条。确定后 self.keep=保留下标集合。"""

    def __init__(self, master, records, email_col, dups):
        super().__init__(master)
        self.records, self.email_col, self.dups = records, email_col, dups
        self.keep = None
        self.title("发现重复邮箱 —— 每组选一个保留")
        self.geometry("740x540")
        self.transient(master)

        head = ttk.Frame(self, padding=(10, 8, 10, 0))
        head.pack(fill="x")
        ttk.Label(head, text="共 %d 个邮箱出现多次，为每组选择保留哪一条（其余会被去掉）：" % len(dups)).pack(side="left")
        ttk.Button(head, text="全部保留第一条", command=lambda: self._quick(True)).pack(side="right")
        ttk.Button(head, text="全部保留最后一条", command=lambda: self._quick(False)).pack(side="right", padx=4)

        canvas = tk.Canvas(self, highlightthickness=0)
        vs = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        iw = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=vs.set)
        canvas.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=6)
        vs.pack(side="right", fill="y")
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(iw, width=e.width))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))

        self.vars = {}
        for email, idxs in dups:
            lf = ttk.LabelFrame(inner, text="%s  （出现 %d 次）" % (email, len(idxs)), padding=6)
            lf.pack(fill="x", pady=4, padx=4)
            var = tk.IntVar(value=idxs[0])
            self.vars[email] = var
            for j in idxs:
                ttk.Radiobutton(lf, text="第%d行   %s" % (j + 1, pc.summarize_record(records[j], email_col)),
                                variable=var, value=j).pack(anchor="w")

        btm = ttk.Frame(self)
        btm.pack(pady=8)
        ttk.Button(btm, text="确定", command=self._ok).pack(side="left", padx=8)
        ttk.Button(btm, text="取消", command=self.destroy).pack(side="left")

        self.update_idletasks()
        try:
            self.wait_visibility()
            self.grab_set()
        except tk.TclError:
            pass

    def _quick(self, first):
        for email, idxs in self.dups:
            self.vars[email].set(idxs[0] if first else idxs[-1])

    def _ok(self):
        drop = set()
        for email, idxs in self.dups:
            chosen = self.vars[email].get()
            for j in idxs:
                if j != chosen:
                    drop.add(j)
        self.keep = {i for i in range(len(self.records)) if i not in drop}
        self.destroy()

    def destroy(self):
        try:
            self.unbind_all("<MouseWheel>")
        except Exception:
            pass
        tk.Toplevel.destroy(self)


class SendLogViewer(tk.Toplevel):
    """发送记录查看器：读取发送记录.csv，可按状态/关键词过滤，并一键重发失败项。"""

    def __init__(self, master):
        super().__init__(master)
        self.app = master
        self.title("发送记录")
        self.geometry("900x560")
        self.transient(master)
        self.rows = []

        bar = ttk.Frame(self, padding=8)
        bar.pack(fill="x")
        ttk.Label(bar, text="状态:").pack(side="left")
        self.filt = tk.StringVar(value="全部")
        ttk.Combobox(bar, textvariable=self.filt, values=["全部", "成功", "失败"],
                     state="readonly", width=6).pack(side="left", padx=(2, 8))
        self.filt.trace_add("write", lambda *a: self._refresh())
        ttk.Label(bar, text="搜索:").pack(side="left")
        self.kw = tk.StringVar()
        ttk.Entry(bar, textvariable=self.kw, width=24).pack(side="left", padx=(2, 8))
        self.kw.trace_add("write", lambda *a: self._refresh())
        ttk.Button(bar, text="刷新", command=self._reload).pack(side="left", padx=2)
        ttk.Button(bar, text="重发失败项", command=self._resend).pack(side="left", padx=2)
        self.count = tk.StringVar(value="")
        ttk.Label(bar, textvariable=self.count, foreground="#555").pack(side="right")

        cols = ("time", "email", "name", "status", "note")
        self.tree = ttk.Treeview(self, columns=cols, show="headings")
        vs = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        for c, t, w in [("time", "时间", 150), ("email", "收件人", 240), ("name", "姓名", 100),
                        ("status", "状态", 90), ("note", "说明", 260)]:
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="w")
        self.tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=(0, 8))
        vs.pack(side="left", fill="y", pady=(0, 8))
        self.tree.tag_configure("bad", background="#ffe1e1")

        self._reload()
        try:
            self.wait_visibility(); self.grab_set()
        except tk.TclError:
            pass

    def _reload(self):
        self.rows = pc.load_log_rows()
        self._refresh()

    def _resend(self):
        self.app.resend_failed()
        self._reload()

    def _refresh(self):
        f = self.filt.get()
        kw = self.kw.get().strip().lower()
        self.tree.delete(*self.tree.get_children())
        tot = ok = bad = 0
        for r in self.rows:
            status = (r.get("状态") or "").strip()
            email = (r.get("收件人") or "").strip()
            tot += 1
            if status.startswith("成功"):
                ok += 1
            elif status.startswith("失败"):
                bad += 1
            if f == "成功" and not status.startswith("成功"):
                continue
            if f == "失败" and not status.startswith("失败"):
                continue
            if kw and kw not in email.lower() and kw not in (r.get("姓名") or "").lower() \
                    and kw not in (r.get("说明") or "").lower():
                continue
            tag = ("bad",) if status.startswith("失败") else ()
            self.tree.insert("", "end", values=(r.get("时间", ""), email, r.get("姓名", ""),
                                                status, r.get("说明", "")), tags=tag)
        self.count.set("共 %d 条 | 成功 %d | 失败 %d" % (tot, ok, bad))


class TemplateEditorDialog(tk.Toplevel):
    """邮件模板编辑器：改主题/正文(HTML 源码 或 简易正文)，浏览器所见即所得预览，保存回写 template.html。"""

    def __init__(self, master):
        super().__init__(master)
        self.app = master
        self.saved = False
        self.changed_subject = None
        self.title("编辑邮件模板")
        self.geometry("980x660")
        self.transient(master)

        html0 = pc.load_template() or pc.DEFAULT_TEMPLATE_HTML
        self.subject0 = master.subject_var.get()

        # 主题
        top = ttk.Frame(self, padding=(10, 8, 10, 0)); top.pack(fill="x")
        ttk.Label(top, text="邮件主题:").pack(side="left")
        self.subject_var = tk.StringVar(value=self.subject0)
        ttk.Entry(top, textvariable=self.subject_var).pack(side="left", fill="x", expand=True, padx=6)

        # 可用变量（基础 + 模板/主题已有 + 名单列）
        base = {"姓名", "称呼", "订单号", "专属优惠码", "邮箱"}
        avail = sorted(base | set(pc.needed_vars(html0, self.subject0)) | set(getattr(master, "headers", [])))[:12]
        vb = ttk.Frame(self); vb.pack(fill="x", padx=10, pady=(6, 0))
        ttk.Label(vb, text="插入变量 → 目标:").pack(side="left")
        self.target = tk.StringVar(value="正文HTML")
        for t in ("主题", "正文HTML", "简易正文"):
            ttk.Radiobutton(vb, text=t, variable=self.target, value=t).pack(side="left")
        for var in avail:
            ttk.Button(vb, text="{{%s}}" % var, width=len(var) + 3,
                       command=lambda v=var: self._insert(v)).pack(side="left", padx=2)

        # Notebook
        self.nb = ttk.Notebook(self); self.nb.pack(fill="both", expand=True, padx=10, pady=6)
        src_frame = ttk.Frame(self.nb); self.tab_src = str(src_frame)
        self.nb.add(src_frame, text="HTML 源码")
        self.src = tk.Text(src_frame, wrap="word", font=("Consolas", 10))
        self.src.pack(fill="both", expand=True)
        self.src.insert("1.0", html0)

        sim = ttk.Frame(self.nb, padding=8); self.tab_simple = str(sim)
        self.nb.add(sim, text="简易正文（不写代码）")
        ttk.Label(sim, text="称呼行:").pack(anchor="w")
        self.greet = ttk.Entry(sim); self.greet.pack(fill="x")
        self.greet.insert(0, "亲爱的 {{姓名}}{{称呼}}：")
        ttk.Label(sim, text="正文（空行分段，段内换行会转成换行；可直接打 {{姓名}} 等变量）:").pack(anchor="w", pady=(6, 0))
        self.body = tk.Text(sim, height=8, wrap="word"); self.body.pack(fill="both", expand=True)
        self.body.insert("1.0", "您好！这里是正文，可用 {{姓名}} {{订单号}} {{专属优惠码}} 等变量。")
        ttk.Label(sim, text="落款:").pack(anchor="w", pady=(6, 0))
        self.signoff = ttk.Entry(sim); self.signoff.pack(fill="x")
        self.signoff.insert(0, "某某公司 敬上")
        ttk.Button(sim, text="用简易正文生成 HTML → 写入源码", command=self._gen_from_simple).pack(anchor="w", pady=8)

        # 预览
        pv = ttk.LabelFrame(self, text="预览（所见即所得）", padding=8); pv.pack(fill="x", padx=10, pady=4)
        ttk.Label(pv, text="代入收件人:").pack(side="left")
        self.samples = self._build_samples()
        self.sample_var = tk.StringVar()
        self.sample_combo = ttk.Combobox(pv, textvariable=self.sample_var, state="readonly",
                                         values=[s[0] for s in self.samples], width=40)
        if self.samples:
            self.sample_combo.current(0)
        self.sample_combo.pack(side="left", padx=6)
        ttk.Button(pv, text="浏览器预览", command=self._preview_browser).pack(side="left", padx=4)
        ttk.Button(pv, text="文本预览", command=self._preview_text).pack(side="left")

        btns = ttk.Frame(self); btns.pack(pady=8)
        ttk.Button(btns, text="保存模板", command=self._save).pack(side="left", padx=8)
        ttk.Button(btns, text="关闭", command=self.destroy).pack(side="left")

        try:
            self.wait_visibility(); self.grab_set()
        except tk.TclError:
            pass

    def _src(self):
        return self.src.get("1.0", "end-1c")

    def _insert(self, var):
        token = "{{%s}}" % var
        t = self.target.get()
        if t == "主题":
            self.subject_var.set(self.subject_var.get() + token)
        elif t == "简易正文":
            self.body.insert("insert", token); self.body.focus_set()
        else:
            self.src.insert("insert", token); self.src.focus_set()

    def _gen_from_simple(self):
        h = pc.build_basic_html(self.greet.get(), self.body.get("1.0", "end-1c"), self.signoff.get())
        self.src.delete("1.0", "end"); self.src.insert("1.0", h)
        self.target.set("正文HTML")

    def _build_samples(self):
        out = []
        for it in getattr(self.app, "items", []):
            if it.get("ok"):
                out.append(("%s (%s)" % (it["email"], it["row"].get("姓名", "")), it["row"]))
            if len(out) >= 50:
                break
        return out

    def _current_row(self):
        sel = self.sample_var.get()
        for label, row in self.samples:
            if label == sel:
                return dict(row)
        return {"邮箱": "someone@example.com", "姓名": "张三", "称呼": "先生",
                "订单号": "A2026001", "专属优惠码": "VIP888"}

    def _preview_browser(self):
        row = self._current_row()
        r_html = mailer.render(self._src(), row) or self._src()
        r_subj = mailer.render(self.subject_var.get(), row) or self.subject_var.get()
        full = "<title>%s</title>\n%s" % (r_subj, r_html)
        try:
            p = os.path.join(pc.BASE_DIR, "_预览.html")
            with open(p, "w", encoding="utf-8") as f:
                f.write(full)
            import webbrowser
            webbrowser.open("file:///" + p.replace("\\", "/"))
        except Exception as e:
            messagebox.showerror("预览失败", str(e))

    def _preview_text(self):
        row = self._current_row()
        r_html = mailer.render(self._src(), row) or self._src()
        _TextPopup(self, "文本预览", mailer.html_to_text(r_html))

    def _save(self):
        if self.nb.select() == self.tab_simple:
            self._gen_from_simple()
        html = self._src()
        if not html.strip():
            messagebox.showwarning("提示", "正文不能为空")
            return
        subj = self.subject_var.get()
        cols = set(getattr(self.app, "headers", [])) | {"邮箱"}
        unknown = [v for v in pc.needed_vars(html, subj) if v not in cols]
        if unknown and getattr(self.app, "records", None):
            if not messagebox.askyesno("确认", "这些模板变量在名单里没有对应列：%s。\n"
                                       "（发送时该变量会用占位值或标为缺变量）仍要保存？"
                                       % "、".join("{{%s}" % v for v in unknown)):
                return
        try:
            pc.save_template(html)
        except Exception as e:
            messagebox.showerror("保存失败", str(e)); return
        # 主题一并存进当前账号，避免切账号/重启后丢失
        idx = self.app.acc_combo.current()
        if 0 <= idx < len(self.app.profiles):
            self.app.profiles[idx]["邮件主题"] = subj
            try:
                pc.save_accounts(self.app.profiles)
            except Exception:
                pass
        self.changed_subject = subj
        self.saved = True
        messagebox.showinfo("已保存", "模板已保存（旧版已备份为 template.html.bak）。")

    def destroy(self):
        if self.saved:
            self.app.subject_var.set(self.changed_subject)
            self.app.reload_template()
        tk.Toplevel.destroy(self)


class _TextPopup(tk.Toplevel):
    def __init__(self, master, title, content):
        super().__init__(master)
        self.title(title)
        self.geometry("720x520")
        t = tk.Text(self, wrap="word")
        t.pack(fill="both", expand=True)
        t.insert("1.0", content)
        t.config(state="disabled")
        ttk.Button(self, text="关闭", command=self.destroy).pack(pady=6)


def _selftest():
    import traceback
    msgs = []
    res_path = os.path.join(pc.BASE_DIR, "selftest_result.txt")

    def write(status, detail):
        with open(res_path, "w", encoding="utf-8") as f:
            f.write("%s\n%s\n" % (status, detail))

    try:
        app = App()
        app.update()
        msgs.append("GUI 构造+渲染 OK")
        msgs.append("tkinterdnd2 加载 _HAVE_DND=%s" % _HAVE_DND)
        app.destroy()

        import openpyxl, tempfile
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["邮箱", "姓名"])
        ws.append(["t@example.com", "测试"])
        tf = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tf.close()
        wb.save(tf.name)
        h, r = pc.parse_xlsx(tf.name)
        os.unlink(tf.name)
        assert r and r[0].get("邮箱") == "t@example.com", "xlsx 读回失败"
        msgs.append("openpyxl Excel 读写 OK")

        html = pc.load_template()
        assert html.strip(), "template.html 为空/缺失"
        item = {"email": "t@example.com",
                "row": {"姓名": "测试", "称呼": "先生", "订单号": "A1",
                        "专属优惠码": "X1", "邮箱": "t@example.com"},
                "ok": True, "err": ""}
        prof = dict(pc.load_accounts()[0])
        prof["邮件主题"] = "Hi {{姓名}}"
        done = {}
        s = pc.Sender(prof, [item], skip_sent=False, dry_run=True,
                      on_status=lambda i, st: done.setdefault("st", st),
                      on_done=lambda o, f, k: done.update({"ok": o, "fail": f, "skip": k}))
        s.run()
        assert done.get("ok") == 1, "dry 发送未成功: %r" % done
        msgs.append("发送管线(dry, 不联网) OK %r" % done)

        write("PASS", "\n".join(msgs))
        sys.exit(0)
    except SystemExit:
        raise
    except Exception:
        msgs.append("FAIL:\n" + traceback.format_exc())
        write("FAIL", "\n".join(msgs))
        sys.exit(1)


def _testsend():
    import traceback
    args = [a for a in sys.argv[1:] if a != "--testsend"]
    email = args[0] if args else ""
    acc_key = args[1] if len(args) > 1 else None
    res = os.path.join(pc.BASE_DIR, "testsend_result.txt")
    if not email:
        with open(res, "w", encoding="utf-8") as fh:
            fh.write("用法：邮件群发面板.exe --testsend <收件邮箱> [账号名关键字]\n（未提供收件邮箱，已取消）\n")
        sys.exit(2)
    msgs = []
    try:
        profs = pc.load_accounts()
        if acc_key:
            prof = next((p for p in profs if acc_key in p.get("名称", "")), profs[0])
        else:
            prof = next((p for p in profs if p.get("通道") == "api"), profs[0])
        html = pc.load_template()
        subj = prof.get("邮件主题", "测试邮件")
        row = {"邮箱": email, "姓名": "测试", "称呼": "先生",
               "订单号": "TEST-" + datetime.now().strftime("%H%M%S"), "专属优惠码": "TEST"}
        for v in pc.needed_vars(html, subj):
            row.setdefault(v, "测试")
        item = {"email": email, "row": row, "ok": True, "err": ""}
        done = {}
        s = pc.Sender(prof, [item], skip_sent=False, dry_run=False,
                      on_status=lambda i, st: done.setdefault("st", st),
                      on_log=lambda m: msgs.append(m),
                      on_done=lambda o, f, k: done.update({"ok": o, "fail": f, "skip": k}))
        s.run()
        ok = done.get("ok") == 1
        with open(res, "w", encoding="utf-8") as fh:
            fh.write(("PASS" if ok else "FAIL") + "\n")
            fh.write("账号=%s 通道=%s 收件=%s\n结果=%r\n%s\n"
                     % (prof.get("名称"), prof.get("通道"), email, done, "\n".join(msgs)))
        sys.exit(0 if ok else 1)
    except SystemExit:
        raise
    except Exception:
        with open(res, "w", encoding="utf-8") as fh:
            fh.write("ERROR\n" + traceback.format_exc())
        sys.exit(2)


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        _selftest()
    elif "--testsend" in sys.argv:
        _testsend()
    else:
        App().mainloop()
