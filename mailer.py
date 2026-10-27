#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
邮件群发脚本（通用版）
- 纯 Python 标准库，无需安装任何第三方包（Python 3.6+）
- 服务商无关：腾讯云 SES / 阿里云 DirectMail / 任意 SMTP 都可用，只改 config.json
- 安全机制：名单校验、发前预览、防重复发送、逐封记录结果

用法：
  python mailer.py --preview 3        预览前 3 封渲染结果（不发送，推荐先跑这个）
  python mailer.py --test 你的邮箱    把第一封渲染后的邮件发给你自己检查效果
  python mailer.py --send             正式群发（自动跳过上次已成功的收件人）
"""

import argparse
import csv
import hashlib
import hmac
import io
import json
import os
import re
import smtplib
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
RECIPIENTS_PATH = os.path.join(BASE_DIR, "recipients.csv")
TEMPLATE_HTML_PATH = os.path.join(BASE_DIR, "template.html")
TEMPLATE_TEXT_PATH = os.path.join(BASE_DIR, "template.txt")  # 可选
SEND_LOG_PATH = os.path.join(BASE_DIR, "发送记录.csv")

VAR_RE = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


# ---------- 基础工具 ----------

def read_text(path):
    """读取文本文件，自动识别 utf-8-sig / gbk / utf-8 编码（兼容 Excel 另存的 CSV）"""
    with open(path, "rb") as f:
        raw = f.read()
    for enc in ("utf-8-sig", "gbk", "utf-8"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    raise SystemExit(f"[错误] 无法识别文件编码，请用 Excel 另存为 UTF-8 格式的 CSV: {path}")


def load_config():
    if not os.path.exists(CONFIG_PATH):
        raise SystemExit("[错误] 未找到 config.json。"
                         "请先把 config.example.json 改名为 config.json，"
                         "再用记事本填入你自己的配置（见使用说明第三步）")
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


def render(text, row, context=""):
    """把 {{列名}} 替换成该行数据；变量缺失时返回 None 并报告"""
    missing = []

    def sub(m):
        key = m.group(1).strip()
        if key in row and str(row[key]).strip() != "":
            return str(row[key]).strip()
        missing.append(key)
        return m.group(0)

    out = VAR_RE.sub(sub, text)
    if missing:
        print(f"[警告] {context}缺少变量: {', '.join(sorted(set(missing)))}")
        return None
    return out


def sanitize(value):
    """防止换行符注入邮件头（发件人名称/主题里的换行会被攻击者利用）"""
    return re.sub(r"[\r\n]+", " ", str(value)).strip()


def html_to_text(html):
    """粗略把 HTML 转成纯文本，作为不支持 HTML 的邮箱的兜底"""
    text = re.sub(r"(?is)<(script|style).*?</\1>", "", html)
    text = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</tr>", "\n", text)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


# ---------- 名单读取与校验 ----------

def load_recipients():
    text = read_text(RECIPIENTS_PATH)
    reader = csv.DictReader(io.StringIO(text))
    if not reader.fieldnames or "邮箱" not in reader.fieldnames:
        raise SystemExit("[错误] recipients.csv 必须包含列名「邮箱」，现有列: %s" % reader.fieldnames)
    rows = [dict(r) for r in reader]
    return rows


def validate(rows):
    """返回 (合格名单, 问题清单)。合格名单已按邮箱去重。"""
    problems = []
    seen = {}
    valid = []
    for i, row in enumerate(rows, start=2):  # 第 1 行是表头
        email = (row.get("邮箱") or "").strip()
        name = (row.get("姓名") or "").strip()
        if not email:
            problems.append(f"第 {i} 行: 邮箱为空")
            continue
        if not EMAIL_RE.match(email):
            problems.append(f"第 {i} 行: 邮箱格式不正确 -> {email}")
            continue
        if not name:
            problems.append(f"第 {i} 行 ({email}): 姓名为空，会导致称呼出错")
            continue
        if email.lower() in seen:
            problems.append(f"第 {i} 行: 邮箱重复（首次出现于第 {seen[email.lower()]} 行）-> {email}")
            continue
        seen[email.lower()] = i
        valid.append(row)
    return valid, problems


def load_already_sent():
    """读取发送记录，返回已成功的邮箱集合（用于断点续发，防止重复轰炸客户）"""
    done = set()
    if not os.path.exists(SEND_LOG_PATH):
        return done
    text = read_text(SEND_LOG_PATH)
    for r in csv.DictReader(io.StringIO(text)):
        if (r.get("状态") or "").startswith("成功"):
            done.add((r.get("收件人") or "").strip().lower())
    return done


def log_result(email, name, status, note):
    new = not os.path.exists(SEND_LOG_PATH)
    with open(SEND_LOG_PATH, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["时间", "收件人", "姓名", "状态", "说明"])
        w.writerow([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), email, name, status, note])


# ---------- 邮件构建与发送 ----------

def build_message(cfg, row, subject, html_body):
    from_addr = cfg["发信人地址"]
    from_name = sanitize(cfg.get("发信人名称", ""))
    to_addr = row["邮箱"].strip()
    to_name = sanitize(row.get("姓名", ""))

    msg = MIMEMultipart("alternative")
    msg["Subject"] = Header(subject, "utf-8")
    msg["From"] = formataddr((from_name, from_addr))
    msg["To"] = f"{to_name} <{to_addr}>" if to_name else to_addr

    text_body = html_to_text(html_body)
    msg.attach(MIMEText(text_body, "plain", "utf-8"))
    msg.attach(MIMEText(html_body, "html", "utf-8"))
    return msg, subject


def connect_smtp(cfg):
    host, port = cfg["SMTP服务器"], int(cfg["SMTP端口"])
    print(f"[连接] {host}:{port} ...")
    if cfg.get("使用SSL", True):
        server = smtplib.SMTP_SSL(host, port, timeout=30)
    else:
        server = smtplib.SMTP(host, port, timeout=30)
        server.ehlo()
        server.starttls()
        server.ehlo()
    server.login(cfg["SMTP用户名"], cfg["SMTP_PASSWORD"] if "SMTP_PASSWORD" in cfg else cfg["SMTP密码"])
    print("[连接] 登录成功")
    return server


class SendError(Exception):
    """API 发送错误。permanent=True 表示这个收件人不用再试了"""

    def __init__(self, message, permanent=False):
        super().__init__(message)
        self.permanent = permanent


def classify_error(e):
    """把发送异常分成「永久失败（别再发了）」和「临时失败（可以稍后重试）」"""
    if isinstance(e, SendError):
        return "失败(永久)" if e.permanent else "失败(临时)"
    code = getattr(e, "smtp_code", None)
    if isinstance(code, int) and 500 <= code < 600:
        return "失败(永久)"
    return "失败(临时)"


# ---------- 腾讯云 SES API 通道（个人实名账号不支持 SMTP，走这里） ----------

# 出现这些字样 = 临时性错误（限流/内部抖动），等一下还能再试
_TEMP_CODES = ("frequency", "throttl", "limitexceeded", "highrejection",
               "internalerror", "temporary", "timeout", "serverbusy")


def ses_api_call(cfg, action, params):
    """调用腾讯云 API（TC3-HMAC-SHA256 签名），返回 Response 字典"""
    ses = cfg.get("腾讯云SES") or {}
    secret_id = (ses.get("SecretId") or "").strip()
    secret_key = (ses.get("SecretKey") or "").strip()
    if (not secret_id or not secret_key
            or "在这里填" in secret_id or "在这里填" in secret_key):
        raise SendError("config.json 未配置腾讯云SES的 SecretId / SecretKey"
                        "（或仍是模板里的占位文字，请先填写）", permanent=True)
    region = ses.get("Region", "ap-guangzhou")
    host, service = "ses.tencentcloudapi.com", "ses"
    algorithm, version = "TC3-HMAC-SHA256", "2020-10-02"

    ts = int(time.time())
    date = time.strftime("%Y-%m-%d", time.gmtime(ts))
    payload = json.dumps(params, ensure_ascii=False).encode("utf-8")

    content_type = "application/json; charset=utf-8"
    canonical_headers = (f"content-type:{content_type}\n"
                         f"host:{host}\n"
                         f"x-tc-action:{action.lower()}\n")
    signed_headers = "content-type;host;x-tc-action"
    canonical_request = ("POST\n/\n\n" + canonical_headers + "\n" + signed_headers + "\n"
                         + hashlib.sha256(payload).hexdigest())
    string_to_sign = (f"{algorithm}\n{ts}\n{date}/{service}/tc3_request\n"
                      + hashlib.sha256(canonical_request.encode("utf-8")).hexdigest())

    def _hmac(key, msg):
        return hmac.new(key, msg.encode("utf-8"), hashlib.sha256).digest()

    k_date = _hmac(("TC3" + secret_key).encode("utf-8"), date)
    k_service = _hmac(k_date, service)
    k_signing = _hmac(k_service, "tc3_request")
    signature = hmac.new(k_signing, string_to_sign.encode("utf-8"),
                         hashlib.sha256).hexdigest()
    authorization = (f"{algorithm} Credential={secret_id}/{date}/{service}/tc3_request, "
                     f"SignedHeaders={signed_headers}, Signature={signature}")

    headers = {
        "Authorization": authorization,
        "Content-Type": content_type,
        "Host": host,
        "X-TC-Action": action,
        "X-TC-Region": region,
        "X-TC-Timestamp": str(ts),
        "X-TC-Version": version,
    }
    req = urllib.request.Request("https://" + host, data=payload,
                                 headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            body = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", "replace")
        try:
            body = json.loads(raw)
        except Exception:
            raise SendError(f"HTTP {e.code}: {raw[:200]}")
    except urllib.error.URLError as e:
        raise SendError(f"网络错误: {e.reason}")

    obj = body.get("Response", {})
    err = obj.get("Error")
    if err:
        code, msg = err.get("Code", ""), err.get("Message", "")
        permanent = not any(t in code.lower() for t in _TEMP_CODES)
        raise SendError(f"SES {code}: {msg}", permanent=permanent)
    return obj


def ses_send_one(cfg, to_addr, subject, row):
    """通过 SES 发送一封个性化邮件（模板在 SES 服务端渲染），返回 MessageId"""
    ses = cfg.get("腾讯云SES") or {}
    try:
        template_id = int(ses.get("模板ID", 0))
    except (TypeError, ValueError):
        template_id = 0
    if not template_id:
        raise SendError("config.json 未配置腾讯云SES 模板ID（需先在控制台创建模板并审核通过）",
                        permanent=True)
    template_data = {str(k): (v or "").strip() for k, v in row.items() if k}
    params = {
        "FromEmailAddress": f'{sanitize(cfg.get("发信人名称", ""))} <{cfg["发信人地址"]}>',
        "Destination": [to_addr],
        "Subject": subject,
        "Template": {
            "TemplateID": template_id,
            "TemplateData": json.dumps(template_data, ensure_ascii=False),
        },
        "TriggerType": 0,  # 0=非触发类（营销/批量类邮件）
    }
    reply = (cfg.get("回复地址") or "").strip()
    if reply:
        params["ReplyToAddresses"] = reply
    obj = ses_api_call(cfg, "SendEmail", params)
    return obj.get("MessageId", "")


# ---------- 主流程 ----------

def main():
    ap = argparse.ArgumentParser(description="邮件群发脚本")
    mode = ap.add_mutually_exclusive_group()
    mode.add_argument("--preview", type=int, metavar="N", help="预览前 N 封渲染结果，不发送")
    mode.add_argument("--test", metavar="邮箱", help="把渲染后的邮件发给你自己检查")
    ap.add_argument("--row", type=int, default=1, help="配合 --test：用名单第 N 行的数据渲染（默认第 1 行）")
    mode.add_argument("--send", action="store_true", help="正式群发")
    ap.add_argument("--limit", type=int, default=0, help="本次最多发送 N 封（0 = 不限制），用于小批量试发")
    ap.add_argument("--skip-bad", action="store_true", help="名单有问题时跳过问题行继续，而不是中止")
    args = ap.parse_args()
    if not (args.preview or args.test or args.send):
        args.preview = 3  # 默认行为：安全的预览模式

    cfg = load_config()
    use_api = str(cfg.get("发送方式", "smtp")).lower() == "api"

    # 1. 读名单 + 校验
    rows = load_recipients()
    valid, problems = validate(rows)
    print(f"[名单] 共 {len(rows)} 行，合格 {len(valid)} 人，问题 {len(problems)} 处")
    if problems:
        for p in problems[:20]:
            print("   -", p)
        if len(problems) > 20:
            print(f"   ... 其余 {len(problems) - 20} 处略")
        if not args.skip_bad:
            print("[中止] 名单存在上述问题。修好名单再跑，或加 --skip-bad 跳过问题行。")
            sys.exit(1)
        print("[继续] 已按 --skip-bad 忽略问题行")

    if not valid:
        raise SystemExit("[错误] 没有任何合格的收件人")

    # 2. 读模板
    if not os.path.exists(TEMPLATE_HTML_PATH):
        raise SystemExit(f"[错误] 缺少模板文件: {TEMPLATE_HTML_PATH}")
    html_tpl = read_text(TEMPLATE_HTML_PATH)

    # 模板变量与名单列的一致性检查（只做提示，不阻断）
    tpl_keys = set(VAR_RE.findall(html_tpl)) | set(VAR_RE.findall(cfg["邮件主题"]))
    csv_cols = set(rows[0].keys())
    unknown = tpl_keys - csv_cols
    if unknown:
        print(f"[警告] 模板里用到了名单中不存在的列: {', '.join(sorted(unknown))}")

    # 3-a. 预览模式
    if args.preview:
        print(f"\n========== 预览前 {min(args.preview, len(valid))} 封（不会发送） ==========")
        for row in valid[: args.preview]:
            html = render(html_tpl, row, context=f"[正文] (发给 {row['邮箱']}) ")
            subject = render(cfg["邮件主题"], row, context="[主题] ")
            print(f"\n收件人: {row['邮箱']}  ({row.get('姓名', '')}{row.get('称呼', '')})")
            print(f"主  题: {subject}")
            print(f"正  文:\n{html_to_text(html or html_tpl)}")
            print("-" * 60)
        print("\n确认无误后，用 --test 自己收一封实测，再用 --send 正式群发。")
        return

    # 3-b. 测试模式：发给自己
    if args.test:
        if not (1 <= args.row <= len(valid)):
            raise SystemExit(f"[错误] --row {args.row} 超出名单范围（共 {len(valid)} 行）")
        sample = valid[args.row - 1]
        html = render(html_tpl, sample, context="[正文] ")
        subject = render(cfg["邮件主题"], sample, context=f"[主题] (发给 {sample['邮箱']}) ")
        if html is None or subject is None:
            raise SystemExit(f"[错误] 名单第 {args.row} 行变量缺失，中止测试")
        if use_api:
            try:
                message_id = ses_send_one(cfg, args.test, subject, sample)
            except SendError as e:
                raise SystemExit(f"[错误] SES API 发送失败: {e}")
            print(f"\n[完成] 测试邮件已通过 SES API 发往 {args.test}（MessageId: {message_id}）")
        else:
            msg = build_message(cfg, sample, subject, html)
            server = connect_smtp(cfg)
            try:
                server.sendmail(cfg["发信人地址"], [args.test], msg.as_string())
            finally:
                server.quit()
            print(f"\n[完成] 测试邮件已发往 {args.test}")
        print(f"       渲染样例（基于名单第 {args.row} 行 {sample['邮箱']}）: 主题「{subject}」")
        print("       请检查: 称呼是否正确、格式是否错乱、是否进了垃圾箱。")
        return

    # 3-c. 正式群发
    already = load_already_sent()
    todo = [r for r in valid if r["邮箱"].strip().lower() not in already]
    if already:
        print(f"[续发] 上次已成功 {len(already)} 封，本次跳过，待发 {len(todo)} 封")
    if args.limit > 0:
        todo = todo[: args.limit]
        print(f"[限量] 本次最多发送 {args.limit} 封，实际待发 {len(todo)} 封")
    if not todo:
        print("[完成] 没有需要发送的邮件。")
        return

    confirm = input(f"即将向 {len(todo)} 位收件人发送，输入 yes 回车确认: ")
    if confirm.strip().lower() != "yes":
        print("已取消。")
        return

    server = connect_smtp(cfg) if not use_api else None
    ok = fail = 0
    interval = float(cfg.get("每封间隔秒", 1.0))
    batch_every, batch_pause = 100, int(cfg.get("每100封额外暂停秒", 60))
    try:
        for idx, row in enumerate(todo, start=1):
            email = row["邮箱"].strip()
            try:
                html = render(html_tpl, row, context=f"[正文] (发给 {email}) ")
                if html is None:
                    raise ValueError("正文变量缺失")
                subject = render(cfg["邮件主题"], row, context=f"[主题] (发给 {email}) ")
                if subject is None:
                    raise SystemExit(f"[错误] 主题变量缺失，中止，避免发出残缺邮件: {email}")
                note = ""
                if use_api:
                    note = "MessageId=" + ses_send_one(cfg, email, subject, row)
                else:
                    msg = build_message(cfg, row, subject, html)
                    server.sendmail(cfg["发信人地址"], [email], msg.as_string())
                log_result(email, row.get("姓名", ""), "成功", note)
                ok += 1
                print(f"[{idx}/{len(todo)}] 成功 -> {email} ({row.get('姓名', '')})")
            except SystemExit:
                raise
            except Exception as e:
                status = classify_error(e)
                log_result(email, row.get("姓名", ""), status, str(e)[:200])
                fail += 1
                print(f"[{idx}/{len(todo)}] {status} -> {email}: {e}")
                if status == "失败(永久)":
                    # 收件方明确拒绝（如邮箱不存在），继续发下一封
                    pass
                elif use_api:
                    # API 通道无连接可重建，等 5 秒后继续
                    time.sleep(5)
                else:
                    # SMTP 临时故障（限流/网络），重建连接再继续
                    try:
                        server.quit()
                    except Exception:
                        pass
                    time.sleep(5)
                    try:
                        server = connect_smtp(cfg)
                    except Exception as e2:
                        print(f"[严重] 无法重建 SMTP 连接，中止发送。已完成 {ok} 封，下次运行会自动续发。错误: {e2}")
                        break
            if idx % batch_every == 0 and idx < len(todo) and batch_pause > 0:
                print(f"[限速] 已发 {batch_every} 封，额外暂停 {batch_pause} 秒 ...")
                time.sleep(batch_pause)
            elif idx < len(todo):
                time.sleep(interval)
    finally:
        try:
            server.quit()
        except Exception:
            pass

    print(f"\n========== 发送结束 ==========")
    print(f"成功 {ok} 封，失败 {fail} 封")
    print(f"逐封明细已记录到: {SEND_LOG_PATH}")
    if fail:
        print("失败的邮件可在修复后重新运行 --send，已成功的会自动跳过。")


if __name__ == "__main__":
    main()
