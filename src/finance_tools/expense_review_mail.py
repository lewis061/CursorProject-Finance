"""Monthly expense-review mail: read Excel, attach fixed files, send via SMTP."""

from __future__ import annotations

import argparse
import logging
import mimetypes
import os
import re
import smtplib
import ssl
from dataclasses import dataclass, field
from datetime import date, datetime
from email.message import EmailMessage
from pathlib import Path

import pandas as pd

from finance_tools.paths import (
    ATTACHMENT_DIR,
    CONTENT_FILE,
    LOGS,
    RECIPIENT_DIR,
    ROOT,
    ensure_dirs,
)

VENUES = ("台北", "南投", "高雄")
SENDER_ACCOUNT = "hilai.aiagent@grand-hilai.com.tw"
MIS_ADDRESS = "mis@grand-hilai.com.tw"

WEEKDAY_ZH = ("一", "二", "三", "四", "五", "六", "日")

# Example: 115.9客務部 費用檢討
SUBJECT_TEMPLATE = "{roc_year}.{review_month}{department} 費用檢討"
BODY_TEMPLATE = """TO 各單位主管

{review_month}月份費用檢討為{meeting_with_weekday}進行
因轉換新系統，本期若分類帳有顯示” 耗用結轉”， 請參考工作頁”耗用結轉”中耗用科目
煩請於{deadline}前完成回傳，謝謝配合。
"""


@dataclass
class ScheduleInfo:
    review_month: int
    western_year: int
    roc_year: int
    meeting_raw: str
    meeting_with_weekday: str
    deadline: str


@dataclass
class DepartmentMail:
    venue: str
    department: str
    to_addrs: list[str] = field(default_factory=list)
    cc_addrs: list[str] = field(default_factory=list)
    attachments: list[Path] = field(default_factory=list)


@dataclass
class PreparedMail:
    venue: str
    department: str
    to_addrs: list[str]
    cc_addrs: list[str]
    subject: str
    body: str
    attachments: list[Path]
    skipped_reason: str | None = None


@dataclass
class SmtpConfig:
    host: str
    port: int
    username: str
    password: str
    sender: str
    use_starttls: bool = True


def _setup_logging(verbose: bool = False) -> Path:
    ensure_dirs()
    log_path = LOGS / f"expense_review_mail_{datetime.now():%Y%m%d_%H%M%S}.log"
    level = logging.DEBUG if verbose else logging.INFO
    stream = logging.StreamHandler()
    try:
        stream.stream.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass
    logging.basicConfig(
        level=level,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            stream,
        ],
        force=True,
    )
    return log_path


def _clean_email(value: object) -> str | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return None
    text = text.rstrip(">")
    if "@" not in text:
        return None
    return text


def _dedupe_preserve(addrs: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for a in addrs:
        key = a.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(a)
    return out


def to_roc_year(western_year: int) -> int:
    return western_year - 1911


def resolve_review_period(
    send_on: date | None = None,
    review_month: int | None = None,
) -> tuple[int, int]:
    """Return (western_year, review_month). Sending month reviews previous month."""
    d = send_on or date.today()
    if review_month is None:
        if d.month == 1:
            return d.year - 1, 12
        return d.year, d.month - 1

    year = d.year
    # e.g. send in Jan for December review
    if review_month > d.month:
        year -= 1
    return year, review_month


def _as_md_text(value: object) -> str:
    """Normalize Excel cell to 'M月D日' text (handles datetime cells)."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, datetime):
        return f"{value.month}月{value.day}日"
    if isinstance(value, date):
        return f"{value.month}月{value.day}日"
    if hasattr(value, "to_pydatetime"):
        try:
            dt = value.to_pydatetime()
            return f"{dt.month}月{dt.day}日"
        except Exception:
            pass
    text = str(value).strip()
    if re.search(r"\d{1,2}\s*月\s*\d{1,2}\s*日", text):
        return re.sub(r"\s+", "", text)
    m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", text)
    if m:
        return f"{int(m.group(2))}月{int(m.group(3))}日"
    return text


def load_schedule(
    review_month: int,
    western_year: int,
    content_file: Path = CONTENT_FILE,
) -> ScheduleInfo:
    df = pd.read_excel(content_file)
    cols = {str(c).strip(): c for c in df.columns}
    month_col = cols.get("費用檢討月份")
    meeting_col = cols.get("開會日期")
    deadline_col = cols.get("資料繳交日期")
    if not month_col or not meeting_col or not deadline_col:
        raise ValueError(
            f"MAIL內文.xlsx 缺少必要欄位，實際欄位: {list(df.columns)}"
        )

    match = df[df[month_col].astype(int) == int(review_month)]
    if match.empty:
        raise ValueError(
            f"MAIL內文.xlsx 找不到費用檢討月份={review_month} 的資料"
        )
    row = match.iloc[0]
    meeting_raw = _as_md_text(row[meeting_col])
    deadline = _as_md_text(row[deadline_col])
    meeting_with_weekday = format_meeting_with_weekday(
        meeting_raw, reference_year=western_year, raw_value=row[meeting_col]
    )
    return ScheduleInfo(
        review_month=int(review_month),
        western_year=western_year,
        roc_year=to_roc_year(western_year),
        meeting_raw=meeting_raw,
        meeting_with_weekday=meeting_with_weekday,
        deadline=deadline,
    )


def format_meeting_with_weekday(
    meeting_raw: str,
    reference_year: int,
    raw_value: object | None = None,
) -> str:
    """Turn '8月26日' into '8月26日(二)'."""
    if isinstance(raw_value, datetime):
        d = raw_value.date()
        return f"{d.month}月{d.day}日({WEEKDAY_ZH[d.weekday()]})"
    if isinstance(raw_value, date):
        return f"{raw_value.month}月{raw_value.day}日({WEEKDAY_ZH[raw_value.weekday()]})"
    if raw_value is not None and hasattr(raw_value, "to_pydatetime"):
        try:
            d = raw_value.to_pydatetime().date()
            return f"{d.month}月{d.day}日({WEEKDAY_ZH[d.weekday()]})"
        except Exception:
            pass

    m = re.fullmatch(r"\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日\s*", meeting_raw)
    if not m:
        return meeting_raw
    month, day = int(m.group(1)), int(m.group(2))
    try:
        d = date(reference_year, month, day)
    except ValueError:
        return meeting_raw
    today = date.today()
    if month - today.month >= 6:
        try:
            d = date(reference_year - 1, month, day)
        except ValueError:
            pass
    return f"{month}月{day}日({WEEKDAY_ZH[d.weekday()]})"


def load_department_mails(
    venue: str,
    recipient_dir: Path = RECIPIENT_DIR,
    attachment_dir: Path = ATTACHMENT_DIR,
) -> list[DepartmentMail]:
    path = recipient_dir / f"MAIL_{venue}.xlsx"
    if not path.exists():
        raise FileNotFoundError(path)

    df = pd.read_excel(path, header=None)
    if df.shape[0] < 2 or df.shape[1] < 2:
        raise ValueError(f"{path.name} 格式異常: {df.shape}")

    depts: list[DepartmentMail] = []
    for col in range(1, df.shape[1]):
        dept_name = df.iloc[1, col]
        if dept_name is None or (isinstance(dept_name, float) and pd.isna(dept_name)):
            continue
        department = str(dept_name).strip()
        if not department:
            continue

        to_addrs: list[str] = []
        cc_addrs: list[str] = []
        for row_idx in range(2, df.shape[0]):
            label = df.iloc[row_idx, 0]
            label_text = (
                ""
                if label is None or (isinstance(label, float) and pd.isna(label))
                else str(label).strip()
            )
            addr = _clean_email(df.iloc[row_idx, col])
            if not addr:
                continue
            if label_text.startswith("收件"):
                to_addrs.append(addr)
            elif label_text.startswith("副"):
                cc_addrs.append(addr)

        attachments = find_attachments(venue, department, attachment_dir)
        depts.append(
            DepartmentMail(
                venue=venue,
                department=department,
                to_addrs=_dedupe_preserve(to_addrs),
                cc_addrs=_dedupe_preserve(cc_addrs),
                attachments=attachments,
            )
        )
    return depts


def find_attachments(
    venue: str, department: str, attachment_dir: Path = ATTACHMENT_DIR
) -> list[Path]:
    """Resolve attachment files for a department.

    If the department name contains「合併」(e.g. 中餐廳合併), attach every
    workbook whose filename contains the base name (中餐廳).
    Otherwise attach the single best-matching file.
    """
    folder = attachment_dir / venue
    if not folder.exists():
        return []
    candidates = sorted(folder.glob("*.xlsx"))

    if "合併" in department:
        base = department.replace("合併", "").strip()
        if not base:
            return []
        merged = [p for p in candidates if base in p.stem]
        return merged

    exact = [p for p in candidates if p.stem.endswith(department) or p.stem == department]
    if exact:
        return [exact[0]]
    loose = [p for p in candidates if department in p.stem]
    if len(loose) == 1:
        return loose
    return []


def render_body(schedule: ScheduleInfo) -> str:
    return BODY_TEMPLATE.format(
        review_month=schedule.review_month,
        meeting_with_weekday=schedule.meeting_with_weekday,
        deadline=schedule.deadline,
    )


def render_subject(schedule: ScheduleInfo, department: str) -> str:
    return SUBJECT_TEMPLATE.format(
        roc_year=schedule.roc_year,
        review_month=schedule.review_month,
        department=department,
    )


def prepare_mails(
    schedule: ScheduleInfo,
    venues: tuple[str, ...] = VENUES,
    test_mode: bool = True,
) -> list[PreparedMail]:
    body = render_body(schedule)
    prepared: list[PreparedMail] = []

    for venue in venues:
        for dept in load_department_mails(venue):
            subject = render_subject(schedule, dept.department)
            if not dept.attachments:
                prepared.append(
                    PreparedMail(
                        venue=dept.venue,
                        department=dept.department,
                        to_addrs=[],
                        cc_addrs=[],
                        subject=subject,
                        body=body,
                        attachments=[],
                        skipped_reason="找不到對應夾檔",
                    )
                )
                continue

            if test_mode:
                to_addrs = [MIS_ADDRESS]
                cc_addrs: list[str] = []
            else:
                to_addrs = list(dept.to_addrs)
                cc_addrs = list(dept.cc_addrs)
                if MIS_ADDRESS.lower() not in {a.lower() for a in to_addrs + cc_addrs}:
                    cc_addrs.append(MIS_ADDRESS)

            if not to_addrs:
                prepared.append(
                    PreparedMail(
                        venue=dept.venue,
                        department=dept.department,
                        to_addrs=[],
                        cc_addrs=cc_addrs,
                        subject=subject,
                        body=body,
                        attachments=dept.attachments,
                        skipped_reason="沒有收件者",
                    )
                )
                continue

            prepared.append(
                PreparedMail(
                    venue=dept.venue,
                    department=dept.department,
                    to_addrs=to_addrs,
                    cc_addrs=cc_addrs,
                    subject=subject,
                    body=body,
                    attachments=dept.attachments,
                )
            )
    return prepared


def _load_dotenv_file(path: Path) -> None:
    if not path.exists():
        return
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        os.environ.setdefault(key, value)


def load_smtp_config() -> SmtpConfig:
    """Load SMTP settings from env / secrets/mail_smtp.env."""
    _load_dotenv_file(ROOT / "secrets" / "mail_smtp.env")

    host = os.environ.get("SMTP_HOST", "").strip()
    port = int(os.environ.get("SMTP_PORT", "587"))
    username = os.environ.get("SMTP_USERNAME", SENDER_ACCOUNT).strip()
    password = os.environ.get("SMTP_PASSWORD", "").strip()
    sender = os.environ.get("SMTP_SENDER", SENDER_ACCOUNT).strip()
    use_starttls = os.environ.get("SMTP_STARTTLS", "1").strip() not in {"0", "false", "False"}

    missing = [name for name, val in (("SMTP_HOST", host), ("SMTP_PASSWORD", password)) if not val]
    if missing:
        raise RuntimeError(
            "缺少 SMTP 設定: "
            + ", ".join(missing)
            + "。請在 secrets/mail_smtp.env 填入伺服器寄信資訊。"
        )
    return SmtpConfig(
        host=host,
        port=port,
        username=username,
        password=password,
        sender=sender,
        use_starttls=use_starttls,
    )


def build_email_message(mail: PreparedMail, sender: str) -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = ", ".join(mail.to_addrs)
    if mail.cc_addrs:
        msg["Cc"] = ", ".join(mail.cc_addrs)
    msg["Subject"] = mail.subject
    msg.set_content(mail.body)

    for path in mail.attachments:
        if not path.exists():
            continue
        ctype, _ = mimetypes.guess_type(str(path))
        if ctype is None:
            maintype, subtype = "application", "octet-stream"
        else:
            maintype, subtype = ctype.split("/", 1)
        msg.add_attachment(
            path.read_bytes(),
            maintype=maintype,
            subtype=subtype,
            filename=path.name,
        )
    return msg


def send_via_smtp(
    mail: PreparedMail,
    smtp: SmtpConfig | None = None,
    dry_run: bool = False,
) -> None:
    if mail.skipped_reason:
        logging.warning(
            "略過 %s/%s: %s", mail.venue, mail.department, mail.skipped_reason
        )
        return

    logging.info(
        "%s %s/%s | Subject=%s | To=%s | Cc=%s | 附件(%s)=%s",
        "DRY-RUN" if dry_run else "SEND",
        mail.venue,
        mail.department,
        mail.subject,
        ";".join(mail.to_addrs),
        ";".join(mail.cc_addrs),
        len(mail.attachments),
        ";".join(p.name for p in mail.attachments),
    )
    if dry_run:
        return

    cfg = smtp or load_smtp_config()
    msg = build_email_message(mail, cfg.sender)
    recipients = _dedupe_preserve(mail.to_addrs + mail.cc_addrs)

    context = ssl.create_default_context()
    with smtplib.SMTP(cfg.host, cfg.port, timeout=60) as server:
        server.ehlo()
        if cfg.use_starttls:
            server.starttls(context=context)
            server.ehlo()
        if cfg.username:
            server.login(cfg.username, cfg.password)
        server.send_message(msg, from_addr=cfg.sender, to_addrs=recipients)


def run(
    *,
    test_mode: bool = True,
    dry_run: bool = False,
    venues: tuple[str, ...] = VENUES,
    review_month: int | None = None,
    limit: int | None = None,
    verbose: bool = False,
) -> int:
    log_path = _setup_logging(verbose=verbose)
    western_year, month = resolve_review_period(review_month=review_month)
    logging.info(
        "費用檢討=%s.%s (民國%s) | test_mode=%s | dry_run=%s",
        western_year,
        month,
        to_roc_year(western_year),
        test_mode,
        dry_run,
    )
    logging.info("寄件帳號=%s（SMTP 伺服器寄出）", SENDER_ACCOUNT)
    logging.info("log=%s", log_path)

    schedule = load_schedule(month, western_year=western_year)
    logging.info(
        "開會=%s | 繳交=%s",
        schedule.meeting_with_weekday,
        schedule.deadline,
    )
    logging.info(
        "主旨範例: %s",
        render_subject(schedule, "客務部"),
    )
    logging.info("內文範例:\n%s", render_body(schedule))

    mails = prepare_mails(schedule, venues=venues, test_mode=test_mode)
    if limit is not None:
        mails = mails[:limit]

    smtp: SmtpConfig | None = None
    if not dry_run:
        smtp = load_smtp_config()

    sent = 0
    skipped = 0
    failed = 0
    for mail in mails:
        try:
            if mail.skipped_reason:
                skipped += 1
            send_via_smtp(mail, smtp=smtp, dry_run=dry_run)
            if not mail.skipped_reason and not dry_run:
                sent += 1
        except Exception:
            failed += 1
            logging.exception("寄送失敗: %s/%s", mail.venue, mail.department)

    logging.info(
        "完成: prepared=%s sent=%s skipped=%s failed=%s",
        len(mails),
        sent,
        skipped,
        failed,
    )
    return 1 if failed else 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="費用檢討每月自動寄信")
    p.add_argument(
        "--test",
        action="store_true",
        default=True,
        help="測試模式：只寄給 mis@（預設開啟）",
    )
    p.add_argument(
        "--prod",
        action="store_true",
        help="正式模式：依 MAIL_台北/南投/高雄 寄給各部門，並同步 CC mis@",
    )
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="只組信與記錄，不實際寄出",
    )
    p.add_argument(
        "--venues",
        nargs="+",
        choices=list(VENUES),
        default=list(VENUES),
        help="要處理的館別",
    )
    p.add_argument(
        "--review-month",
        type=int,
        help="覆寫費用檢討月份（預設=寄信月的上個月）",
    )
    p.add_argument(
        "--limit",
        type=int,
        help="最多處理幾封（測試用）",
    )
    p.add_argument("-v", "--verbose", action="store_true")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    test_mode = not args.prod
    return run(
        test_mode=test_mode,
        dry_run=args.dry_run,
        venues=tuple(args.venues),
        review_month=args.review_month,
        limit=args.limit,
        verbose=args.verbose,
    )


if __name__ == "__main__":
    raise SystemExit(main())
