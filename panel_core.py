# -*- coding: utf-8 -*-
"""
邮件群发面板 —— 核心逻辑层（不依赖 tkinter，方便单独测试与复用）。
职责：
  - 多套发信账号(accounts.json)读写
  - 名单导入：CSV / Excel(.xlsx) / 粘贴文本，统一成 (表头, 记录列表)
  - 智能列识别：猜出哪列是邮箱、哪列对应模板变量({{姓名}}{{称呼}}...)
  - 组装可发送行 + 变量完整性校验
  - 发送执行器 Sender（api / smtp 双通道，支持试运行、去重、逐封回调）
所有真正的“发信”动作复用同目录 mailer.py 的成熟实现。
"""
import os
import re
import csv
import io
import json
import sys

import mailer  # 复用：render / ses_send_one / build_message / connect_smtp / classify_error / EMAIL_RE / VAR_RE / read_text / sanitize / SendError

# 打包成 exe（PyInstaller frozen）时，数据文件放在 exe 同目录，而非临时解包目录
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ACCOUNTS_PATH = os.path.join(BASE_DIR, "accounts.json")
TEMPLATE_HTML_PATH = os.path.join(BASE_DIR, "template.html")
LOG_PATH = os.path.join(BASE_DIR, "发送记录.csv")


# ---------------------------------------------------------------- 账号配置

def load_accounts():
    with open(ACCOUNTS_PATH, encoding="utf-8") as f:
        data = json.load(f)
    return data.get("profiles", [])


def save_accounts(profiles):
    with open(ACCOUNTS_PATH, "w", encoding="utf-8") as f:
        json.dump({"profiles": profiles}, f, ensure_ascii=False, indent=2)


def default_profile():
    return {
        "名称": "新账号",
        "通道": "smtp",
        "发信人地址": "",
        "SMTP服务器": "smtp.exmail.qq.com",
        "SMTP端口": 465,
        "使用SSL": True,
        "SMTP用户名": "",
        "SMTP密码": "",
        "发信人名称": "某某公司",
        "回复地址": "",
        "邮件主题": "亲爱的{{姓名}}{{称呼}}，这是专程寄给您的一封信",
        "每封间隔秒": 2.0,
        "每100封额外暂停秒": 60,
    }


def default_ses_profile():
    p = default_profile()
    p["通道"] = "api"
    p["SMTP服务器"] = ""
    p["腾讯云SES"] = {"SecretId": "", "SecretKey": "", "Region": "ap-guangzhou", "模板ID": 0}
    return p


# ---------------------------------------------------------------- 模板与变量

def load_template():
    if not os.path.exists(TEMPLATE_HTML_PATH):
        return ""
    return mailer.read_text(TEMPLATE_HTML_PATH)


def needed_vars(html, subject):
    """模板正文 + 主题里出现的所有 {{变量}}，排除收件人自身用的‘邮箱’由单独列处理。"""
    found = set(mailer.VAR_RE.findall(html or "")) | set(mailer.VAR_RE.findall(subject or ""))
    return sorted(found)


# ---------------------------------------------------------------- 模板编辑

_TEMPLATE_WRAPPER = """<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="font-family:'Microsoft YaHei',Arial,sans-serif;line-height:1.8;color:#333;max-width:640px;margin:0 auto;">

  <p>__GREETING__</p>

__BODY__

  <p style="margin-top:32px;">
    此致<br>
    __SIGNOFF__
  </p>

  <hr style="border:none;border-top:1px solid #ddd;margin:24px 0;">
  <p style="font-size:12px;color:#999;">__UNSUB__</p>

</body>
</html>
"""


def _esc_keep_vars(text):
    """纯文本转义 & < >，但保留 {{变量}} 字面量（转义只影响尖括号，不含大括号）。"""
    s = str(text or "")
    s = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return s


def build_basic_html(greeting, body_text, signoff,
                     unsubscribe="如不希望继续收到此类邮件，请回复“退订”，我们会立即将您从列表中移除。"):
    """由简易字段生成同款排版的 HTML。正文按空行分段，段内换行转 <br>。"""
    paras = []
    for block in str(body_text or "").split("\n\n"):
        block = block.strip("\n")
        if not block.strip():
            continue
        inner = _esc_keep_vars(block).replace("\n", "<br>\n    ")
        paras.append("  <p>" + inner + "</p>")
    body_html = "\n\n".join(paras) if paras else "  <p>（正文为空）</p>"
    out = (_TEMPLATE_WRAPPER
           .replace("__GREETING__", _esc_keep_vars(greeting))
           .replace("__BODY__", body_html)
           .replace("__SIGNOFF__", _esc_keep_vars(signoff))
           .replace("__UNSUB__", _esc_keep_vars(unsubscribe)))
    return out


DEFAULT_TEMPLATE_HTML = build_basic_html(
    "亲爱的 {{姓名}}{{称呼}}：",
    "您好！感谢您一直以来的支持。您的订单 {{订单号}} 已有最新进展，"
    "我们为您准备了专属优惠码 {{专属优惠码}}，下单时输入即可享受折扣。\n\n"
    "如有任何问题，欢迎随时回复本邮件联系我们。",
    "某某公司 敬上",
)


def save_template(html):
    """保存 template.html；先备份旧文件为 template.html.bak。返回写入路径。"""
    import shutil
    if os.path.exists(TEMPLATE_HTML_PATH):
        try:
            shutil.copy2(TEMPLATE_HTML_PATH, TEMPLATE_HTML_PATH + ".bak")
        except Exception:
            pass
    with open(TEMPLATE_HTML_PATH, "w", encoding="utf-8") as f:
        f.write(html)
    return TEMPLATE_HTML_PATH


# ---------------------------------------------------------------- 名单导入


def _clean(v):
    if v is None:
        return ""
    return str(v).strip()


_KV_LINE = re.compile(r"^\s*([^\s：:]{1,15})\s*[：:]\s*(.+?)\s*$")
_KNOWN_KEYS = {
    "邮箱", "邮件", "email", "mail", "e-mail", "姓名", "名字", "name",
    "称呼", "称谓", "性别", "订单", "订单号", "单号", "优惠", "优惠码",
    "兑换码", "券", "专属优惠码", "手机", "电话", "地址", "公司", "备注",
}
_KEY_ALIAS = {
    "邮件": "邮箱", "mail": "邮箱", "email": "邮箱", "e-mail": "邮箱",
    "名字": "姓名", "name": "姓名",
    "性别": "称呼", "称谓": "称呼",
    "订单": "订单号", "单号": "订单号",
    "优惠": "专属优惠码", "优惠码": "专属优惠码", "兑换码": "专属优惠码", "券": "专属优惠码",
}


def try_parse_block(text):
    """把「标签：值 / 空行或重复键分隔多条」的块状文本转成 (headers, records)。
    识别不到就返回 None。"""
    if not text:
        return None
    lines_all = [ln for ln in text.splitlines() if ln.strip()]
    if len(lines_all) < 2:
        return None
    hits = [_KV_LINE.match(ln) for ln in lines_all]
    kv_lines = [(m.group(1).strip(), m.group(2).strip()) for m in hits if m]
    if len(kv_lines) < 2 or len(kv_lines) / len(lines_all) < 0.6:
        return None
    seen_keys = {k.lower() for k, _ in kv_lines}
    if not (seen_keys & _KNOWN_KEYS or {k for k, _ in kv_lines} & _KNOWN_KEYS):
        return None
    records, cur, keys_in_cur = [], {}, set()
    for k, v in kv_lines:
        if k in keys_in_cur:                    # 同键再出现 = 下一条
            records.append(cur); cur = {}; keys_in_cur = set()
        cur[k] = v; keys_in_cur.add(k)
    if cur:
        records.append(cur)
    if len(records) < 2:
        blocks = re.split(r"\n[ \t]*\n", text.strip())
        if len(blocks) > 1:
            records = []
            for blk in blocks:
                d = {}
                for ln in blk.splitlines():
                    m = _KV_LINE.match(ln.strip())
                    if m:
                        d[m.group(1).strip()] = m.group(2).strip()
                if d:
                    records.append(d)
    if not records:
        return None
    norm = []
    headers = []
    for r in records:
        m = {}
        for k, v in r.items():
            kk = _KEY_ALIAS.get(k.lower()) or _KEY_ALIAS.get(k) or k
            m[kk] = v
            if kk not in headers:
                headers.append(kk)
        norm.append(m)
    return headers, norm


def parse_csv(path):
    text = mailer.read_text(path)
    block = try_parse_block(text)
    if block:
        return block
    reader = csv.reader(io.StringIO(text))
    rows = [r for r in reader if any(_clean(c) for c in r)]
    if not rows:
        return [], []
    headers = [_clean(h) or f"第{i+1}列" for i, h in enumerate(rows[0])]
    records = []
    for r in rows[1:]:
        rec = {headers[i]: _clean(r[i]) if i < len(r) else "" for i in range(len(headers))}
        records.append(rec)
    return headers, records


def parse_xlsx(path, header_row=True):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    raw = []
    for row in ws.iter_rows(values_only=True):
        cells = [_clean(c) for c in row]
        if any(cells):
            raw.append(cells)
    wb.close()
    if not raw:
        return [], []
    width = max(len(r) for r in raw)
    if header_row:
        headers = [raw[0][i] if i < len(raw[0]) and raw[0][i] else f"第{i+1}列" for i in range(width)]
        body = raw[1:]
    else:
        headers = [f"第{i+1}列" for i in range(width)]
        body = raw
    records = []
    for r in body:
        rec = {headers[i]: (r[i] if i < len(r) else "") for i in range(width)}
        records.append(rec)
    return headers, records


def parse_pasted(text):
    """粘贴内容：块状「标签：值」优先识别；否则像表格(制表符)按 TSV；否则当作邮箱清单抓邮箱。"""
    block = try_parse_block(text)
    if block:
        return block
    lines = [ln for ln in text.splitlines() if ln.strip()]
    tabbed = sum(1 for ln in lines if "\t" in ln)
    if lines and tabbed >= max(1, len(lines) // 2):  # 多数行带制表符 → TSV 表格
        reader = csv.reader(io.StringIO("\n".join(lines)), delimiter="\t")
        rows = [[_clean(c) for c in r] for r in reader if any(_clean(c) for c in r)]
        headers = [h or f"第{i+1}列" for i, h in enumerate(rows[0])]
        records = [{headers[i]: (r[i] if i < len(r) else "") for i in range(len(headers))} for r in rows[1:]]
        return headers, records
    emails = mailer.EMAIL_RE and re.findall(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", text)
    seen, records = set(), []
    for e in emails:
        if e.lower() not in seen:
            seen.add(e.lower())
            records.append({"邮箱": e})
    return ["邮箱"], records


# ---------------------------------------------------------------- 智能列识别

_KEYWORDS = {
    "邮箱": ["邮箱", "邮件", "电子邮箱", "email", "e-mail", "mail", "地址", "address"],
    "姓名": ["姓名", "名字", "名称", "name", "联系人", "收件人", "客户姓名", "全名"],
    "称呼": ["称呼", "称谓", "title", "salutation", "性别", "先生女士", "问候"],
    "订单号": ["订单", "order", "单号", "编号", "订单编号"],
    "专属优惠码": ["优惠", "coupon", "promo", "折扣", "兑换码", "券", "码"],
}
_HONORIFICS = {"先生", "女士", "小姐", "老师", "同学", "同志", "博士", "医生",
               "经理", "总", "男", "女", "m", "f", "male", "female", "先生/女士"}


def _nonempty(vals):
    return [str(v).strip() for v in vals if str(v).strip() != ""]


def _is_email(s):
    return bool(s) and "@" in s and bool(mailer.EMAIL_RE.match(s))


def _email_score(recs, col):
    vs = _nonempty([r.get(col, "") for r in recs])
    return sum(1 for v in vs if _is_email(v)) / len(vs) if vs else 0.0


def _honorific_score(recs, col):
    vs = _nonempty([r.get(col, "") for r in recs])
    return sum(1 for v in vs if v in _HONORIFICS) / len(vs) if vs else 0.0


def _name_score(recs, col):
    vs = _nonempty([r.get(col, "") for r in recs])
    if not vs:
        return 0.0
    good = 0
    for v in vs:
        if "@" in v or v in _HONORIFICS or v.isdigit():
            continue
        if len(v) <= 12 and re.search(r"[A-Za-z\u4e00-\u9fff]", v) and not re.search(r"\d", v):
            good += 1
    return good / len(vs)


def _code_score(recs, col):
    vs = _nonempty([r.get(col, "") for r in recs])
    if not vs:
        return 0.0
    good = sum(1 for v in vs if re.fullmatch(r"[A-Za-z0-9_\-]{2,}", v) and re.search(r"\d", v))
    return good / len(vs)


_SCORERS = {"称呼": _honorific_score, "姓名": _name_score,
            "订单号": _code_score, "专属优惠码": _code_score}


def _find_col_keyword(var, headers, records, exclude):
    hlow = {h: h.lower() for h in headers}
    kws = _KEYWORDS.get(var, [])
    for h in headers:
        if h not in exclude and h == var:
            return h
    for h in headers:
        if h in exclude:
            continue
        if any(kw in hlow[h] for kw in kws):
            return h
    for h in headers:
        if h not in exclude and var in h:
            return h
    return None


def _find_col_byvalue(var, headers, records, exclude):
    scorer = _SCORERS.get(var)
    if not scorer or not records:
        return None
    best, bestscore = None, 0.6  # 需要足够置信度才自动认
    for h in headers:
        if h in exclude:
            continue
        s = scorer(records, h)
        if s > bestscore:
            best, bestscore = h, s
    return best


def detect_mapping(html, subject, headers, records):
    """两遍识别：先列名关键词，再列内容特征。已认领的列不会被后面的变量抢走。
    返回 (email_col, var_map)，var_map: 变量 -> ("col",名) / ("fixed",值)。"""
    email_col = _find_col_keyword("邮箱", headers, records, set())
    if not email_col and records:
        best, bs = None, 0.5
        for h in headers:
            s = _email_score(records, h)
            if s > bs:
                best, bs = h, s
        email_col = best
    if not email_col:
        email_col = headers[0] if headers else None
    claimed = {email_col} if email_col else set()
    var_map = {}
    for var in needed_vars(html, subject):
        col = _find_col_keyword(var, headers, records, claimed) or _find_col_byvalue(var, headers, records, claimed)
        if col:
            var_map[var] = ("col", col)
            claimed.add(col)
        elif var == "邮箱" and email_col:
            var_map[var] = ("col", email_col)
        elif var == "姓名":
            var_map[var] = ("fixed", "客户")
        elif var == "称呼":
            var_map[var] = ("fixed", "你好")
        else:
            var_map[var] = ("fixed", "-")
    return email_col, var_map


# ---------------------------------------------------------------- 重复检测

def _row_key(rec):
    return "||".join(str(v).strip().lower() for v in rec.values())


def find_dup_groups(records, email_col, mode="email"):
    """重复分组。mode="email" 按规范化邮箱；mode="row" 按整行内容完全相同才算重复
    （同邮箱但姓名/订单等不同则不合并）。返回 [(展示用邮箱, [原始下标...])]，仅含出现>1次的组。"""
    buckets = {}
    for i, rec in enumerate(records):
        if mode == "row":
            key = _row_key(rec)
            if key.replace("|", "").strip() == "":
                continue
        else:
            e = (rec.get(email_col, "") if email_col else "").strip()
            if not _is_email(e):
                continue
            key = e.lower()
        buckets.setdefault(key, []).append(i)
    out = []
    for key, idxs in buckets.items():
        if len(idxs) > 1:
            show = (records[idxs[0]].get(email_col, "") if email_col else "").strip() or key
            out.append((show, idxs))
    return out


def summarize_record(rec, email_col):
    parts = [f"{k}:{str(v).strip()}" for k, v in rec.items() if k != email_col and str(v).strip()]
    return "   ".join(parts) if parts else "(仅邮箱，无其它信息)"


# ---------------------------------------------------------------- 组装可发送行

def resolve(source, record, email_val):
    kind = source[0]
    if kind == "col":
        return record.get(source[1], "").strip()
    if kind == "fixed":
        return (source[1] or "").strip()
    if kind == "emaillocal":
        return email_val.split("@")[0] if email_val else ""
    return ""


def build_sendable(records, email_col, var_map, html, subject_tpl):
    """把记录转成 [{email,row,ok,err}]；ok=False 表示邮箱无效或缺变量。"""
    vars_ = needed_vars(html, subject_tpl)
    out = []
    for rec in records:
        email = (rec.get(email_col, "") if email_col else "").strip()
        row = {}
        missing = []
        for var in vars_:
            val = resolve(var_map[var], rec, email)
            if not val:
                missing.append(var)
            row[var] = val
        row["邮箱"] = email
        valid_email = bool(mailer.EMAIL_RE.match(email))
        if not valid_email:
            out.append({"email": email, "row": row, "ok": False, "err": "邮箱无效"})
        elif missing:
            out.append({"email": email, "row": row, "ok": False, "err": "缺变量:" + "/".join(missing)})
        else:
            out.append({"email": email, "row": row, "ok": True, "err": ""})
    return out


# ---------------------------------------------------------------- 发送记录（去重）

def load_log_rows():
    """读取发送记录.csv 全部行，返回 list[dict]（键：时间/收件人/姓名/状态/说明）。无文件返回 []。"""
    if not os.path.exists(LOG_PATH):
        return []
    text = mailer.read_text(LOG_PATH)
    return list(csv.DictReader(io.StringIO(text)))


# ---------------------------------------------------------------- 导出给命令行

def export_recipients_csv(items, html, subject, path):
    """把已解析的可发送行导出成命令行可直接用的 recipients.csv。
    列 = 邮箱 + 模板所需变量；每格写实际解析值（③的固定值/兜底也落成实值），
    这样命令行 mailer.py 的“列名精确匹配”规则一定能对上。"""
    vars_ = needed_vars(html, subject)
    cols = ["邮箱"] + [v for v in vars_ if v != "邮箱"]
    rows = []
    for it in items:
        if not it.get("ok"):
            continue
        r = it.get("row", {})
        rows.append({c: str(r.get(c, "")).strip() for c in cols})
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)
    return path, len(rows)


def export_config_json(profile, path):
    """把面板的一套账号转成命令行 mailer.py 需要的 config.json。"""
    api = profile.get("通道") == "api"
    cfg = {
        "说明": "由邮件群发面板导出的命令行配置（py mailer.py --send 用这个）。",
        "发送方式": "api" if api else "smtp",
        "发信人地址": profile.get("发信人地址", ""),
        "发信人名称": profile.get("发信人名称", ""),
        "回复地址": profile.get("回复地址", ""),
        "邮件主题": profile.get("邮件主题", ""),
        "每封间隔秒": profile.get("每封间隔秒", 2.0),
        "每100封额外暂停秒": profile.get("每100封额外暂停秒", 60),
    }
    if api:
        ses = profile.get("腾讯云SES", {}) or {}
        cfg["腾讯云SES"] = {
            "SecretId": ses.get("SecretId", ""),
            "SecretKey": ses.get("SecretKey", ""),
            "Region": ses.get("Region", "ap-guangzhou"),
            "模板ID": ses.get("模板ID", 0),
        }
    else:
        cfg["SMTP服务器"] = profile.get("SMTP服务器", "")
        cfg["SMTP端口"] = profile.get("SMTP端口", 465)
        cfg["使用SSL"] = profile.get("使用SSL", True)
        cfg["SMTP用户名"] = profile.get("SMTP用户名", "")
        cfg["SMTP密码"] = profile.get("SMTP密码", "")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    return path


def load_sent():
    done = set()
    if not os.path.exists(LOG_PATH):
        return done
    text = mailer.read_text(LOG_PATH)
    for r in csv.DictReader(io.StringIO(text)):
        if (r.get("状态") or "").startswith("成功"):
            done.add((r.get("收件人") or "").strip().lower())
    return done


def _append_log(email, name, status, note):
    new = not os.path.exists(LOG_PATH)
    with open(LOG_PATH, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["时间", "收件人", "姓名", "状态", "说明"])
        import datetime
        w.writerow([datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), email, name, status, note])


# ---------------------------------------------------------------- 发送执行器

class Sender:
    """在后台线程里调用 run()；通过回调把进度回传给界面（回调里不要碰 tkinter）。"""

    def __init__(self, profile, items, interval=2.0, batch_every=100, batch_pause=60,
                 skip_sent=True, dry_run=False,
                 on_progress=None, on_status=None, on_log=None, on_done=None):
        self.cfg = dict(profile)
        self.cfg.setdefault("发送方式", "api" if profile.get("通道") == "api" else "smtp")
        self.items = items            # build_sendable 的输出（仅传 ok=True 的进来）
        self.interval = float(profile.get("每封间隔秒", interval))
        self.batch_every = int(batch_every)
        self.batch_pause = int(profile.get("每100封额外暂停秒", batch_pause))
        self.skip_sent = skip_sent
        self.dry_run = dry_run
        self.cancel = False
        self.on_progress = on_progress or (lambda *a: None)
        self.on_status = on_status or (lambda *a: None)
        self.on_log = on_log or (lambda *a: None)
        self.on_done = on_done or (lambda *a: None)

    def run(self):
        import time
        cfg = self.cfg
        html = load_template()
        subject_tpl = cfg.get("邮件主题", "")
        use_api = cfg.get("发送方式") == "api"
        sent = load_sent() if self.skip_sent else set()
        server = None
        ok = fail = skip = 0
        todo = [it for it in self.items]
        try:
            if not self.dry_run and not use_api:
                server = mailer.connect_smtp(cfg)
            for idx, it in enumerate(todo, start=1):
                if self.cancel:
                    self.on_log("已取消，停止发送。")
                    break
                email = it["email"]
                row = it["row"]
                if email.lower() in sent:
                    skip += 1
                    self.on_status(idx - 1, "跳过(已发过)")
                    self.on_progress(idx, len(todo))
                    continue
                try:
                    subject = mailer.render(subject_tpl, row)
                    body_html = mailer.render(html, row)
                    if subject is None or body_html is None:
                        raise ValueError("变量渲染失败")
                    note = ""
                    if self.dry_run:
                        note = "试运行(未实际发送)"
                    elif use_api:
                        note = "MessageId=" + mailer.ses_send_one(cfg, email, subject, row)
                    else:
                        msg = mailer.build_message(cfg, row, subject, body_html)
                        server.sendmail(cfg["发信人地址"], [email], msg.as_string())
                        note = "SMTP已投递"
                    ok += 1
                    self.on_status(idx - 1, "成功")
                    self.on_log(f"[{idx}/{len(todo)}] 成功 -> {email} {note}")
                    if not self.dry_run:
                        _append_log(email, row.get("姓名", ""), "成功", note)
                except Exception as e:
                    status = mailer.classify_error(e) if isinstance(e, mailer.SendError) else "失败(临时)"
                    fail += 1
                    self.on_status(idx - 1, status)
                    self.on_log(f"[{idx}/{len(todo)}] {status} -> {email}: {e}")
                    if not self.dry_run:
                        _append_log(email, row.get("姓名", ""), status, str(e)[:200])
                    if (not use_api) and not self.dry_run and status == "失败(临时)":
                        try:
                            server.quit()
                        except Exception:
                            pass
                        time.sleep(5)
                        try:
                            server = mailer.connect_smtp(cfg)
                        except Exception as e2:
                            self.on_log(f"[严重] 无法重连 SMTP，中止。剩余未发：{e2}")
                            break
                self.on_progress(idx, len(todo))
                if idx % self.batch_every == 0 and idx < len(todo) and self.batch_pause > 0:
                    self.on_log(f"[限速] 暂停 {self.batch_pause} 秒 ...")
                    time.sleep(self.batch_pause)
                else:
                    time.sleep(self.interval if not self.dry_run else 0)
        finally:
            try:
                if server:
                    server.quit()
            except Exception:
                pass
        self.on_done(ok, fail, skip)
