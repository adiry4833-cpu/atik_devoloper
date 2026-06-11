import telebot
from telebot import types
import json
import os
import re
import time
import threading
import datetime
import requests
import phonenumbers
import random
import csv
import io
import tempfile
import openpyxl
import xlrd
from bs4 import BeautifulSoup
from phonenumbers import region_code_for_number, geocoder

_PID_FILE = "/tmp/ar_otp_bot.pid"
_my_pid = os.getpid()
if os.path.exists(_PID_FILE):
    try:
        _old_pid = int(open(_PID_FILE).read().strip())
        if _old_pid != _my_pid:
            try:
                os.kill(_old_pid, 9)
                time.sleep(5)
                print(f"[START] Killed old instance PID {_old_pid}")
            except ProcessLookupError:
                pass
    except Exception:
        pass
open(_PID_FILE, "w").write(str(_my_pid))

API_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
# <<SYNC:SUPER_ADMIN_IDS:START>>
SUPER_ADMIN_IDS = [
    6664150885,
    8523774444,
]
# <<SYNC:SUPER_ADMIN_IDS:END>>
ADMIN_IDS = list(SUPER_ADMIN_IDS)
CHANNEL_2 = ""

# ── Panel 1 (Mahofuza) ───────────────────────────────────────────────────────
P1_BASE_URL = "http://91.232.105.47/ints"
P1_LOGIN_PAGE = P1_BASE_URL + "/login"
P1_SIGNIN_URL = P1_BASE_URL + "/signin"
P1_CDR_PAGE = P1_BASE_URL + "/agent/SMSCDRStats"
P1_CDR_DATA_URL = P1_BASE_URL + "/agent/res/data_smscdr.php"
P1_USER_NAME = "Mahofuza"
P1_PASSWORD = "Mahofuza"

# ── Panel 2 (Sagardas50 / XISORA) ────────────────────────────────────────────
P2_BASE_URL = "http://94.23.31.29/sms"
P2_SIGNIN_URL = P2_BASE_URL + "/signmein"
P2_REPORTS_PAGE = P2_BASE_URL + "/client/Reports"
P2_DATA_URL = P2_BASE_URL + "/client/ajax/dt_reports.php"
P2_USER_NAME = "Sagardas50"
P2_PASSWORD = "Sagardas50"

# ── Panel 3 (Rabbi1_FD) ───────────────────────────────────────────────────────
P3_BASE_URL = "http://168.119.13.175/ints"
P3_LOGIN_PAGE = P3_BASE_URL + "/login"
P3_SIGNIN_URL = P3_BASE_URL + "/signin"
P3_CDR_PAGE = P3_BASE_URL + "/agent/SMSCDRStats"
P3_CDR_DATA_URL = P3_BASE_URL + "/agent/res/data_smscdr.php"
P3_USER_NAME = "Rabbi1_FD"
P3_PASSWORD = "Rabbi1_FD"

# ── Panel 4 (Rabbi12) ─────────────────────────────────────────────────────────
P4_BASE_URL = "http://144.217.71.192/ints"
P4_LOGIN_PAGE = P4_BASE_URL + "/login"
P4_SIGNIN_URL = P4_BASE_URL + "/signin"
P4_CDR_PAGE = P4_BASE_URL + "/agent/SMSCDRStats"
P4_CDR_DATA_URL = P4_BASE_URL + "/agent/res/data_smscdr.php"
P4_USER_NAME = "Rabbi12"
P4_PASSWORD = "Rabbi12"

# ── Panel 5 (Rabbi12_v2 / 51.75.144.178) ─────────────────────────────────────
P5_BASE_URL = "http://51.75.144.178/ints"
P5_LOGIN_PAGE = P5_BASE_URL + "/login"
P5_SIGNIN_URL = P5_BASE_URL + "/signin"
P5_CDR_PAGE = P5_BASE_URL + "/agent/SMSCDRStats"
P5_CDR_DATA_URL = P5_BASE_URL + "/agent/res/data_smscdr.php"
P5_USER_NAME = "Rabbi12"
P5_PASSWORD = "Rabbi12@"

# ── Panel 6 (Sagardas50 / TrueSMS.net — SMSRanges) ───────────────────────────
P6_BASE_URL = "https://truesms.net"
P6_LOGIN_PAGE = P6_BASE_URL + "/login"
P6_SIGNIN_URL = P6_BASE_URL + "/signin"
P6_CDR_PAGE = P6_BASE_URL + "/agent/SMSRanges"
P6_CDR_DATA_URL = P6_BASE_URL + "/agent/res/data_smsranges.php"
P6_USER_NAME = "Sagardas50"
P6_PASSWORD = "Sagardas50"

# ── Builtin extra panels (hardcoded, always started automatically) ─────────────
# <<SYNC:_BUILTIN_PANELS:START>>
_BUILTIN_PANELS = [
    {'id': 'bp1', 'host': '139.99.69.196', 'base_url': 'http://139.99.69.196/ints', 'url_hint': 'http://139.99.69.196/ints/agent/SMSCDRStats', 'username': 'Mahofuza12', 'password': 'Mahofuza12', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': None},
    {'id': 'bp2', 'host': '139.99.9.4', 'base_url': 'http://139.99.9.4/ints', 'url_hint': 'http://139.99.9.4/ints/agent/SMSCDRStats', 'username': 'Rabbi12', 'password': 'Rabbi12', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': None},
    {'id': 'bp3', 'host': '54.36.173.235', 'base_url': 'http://54.36.173.235/ints', 'url_hint': 'http://54.36.173.235/ints/agent/SMSCDRStats', 'username': 'Rabbi12', 'password': 'Rabbi@', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': None},
    {'id': 'bp4', 'host': '54.39.104.241', 'base_url': 'http://54.39.104.241/ints', 'url_hint': 'http://54.39.104.241/ints/agent/SMSCDRStats', 'username': 'Rabbi5', 'password': 'Rabbi5', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': None},
    {'id': 'bp5', 'host': '213.32.24.208', 'base_url': 'http://213.32.24.208/ints', 'url_hint': 'http://213.32.24.208/ints/agent/SMSCDRStats', 'username': 'mahofuza', 'password': 'mahofuza@', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': None},
    {'id': 'bp6', 'host': '15.235.182.3', 'base_url': 'http://15.235.182.3/konekta', 'url_hint': 'http://15.235.182.3/konekta/agent/SMSCDRReports', 'username': 'Rabbi200', 'password': 'Rabbi200', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': None},
    {'id': 'bp7', 'host': 'nexor-iprn.com', 'base_url': 'https://nexor-iprn.com', 'url_hint': 'https://nexor-iprn.com/agent/SMSCDRStats', 'username': 'Rabbi12', 'password': 'Rabbi12@', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': None},
    {'id': 'bp8', 'host': '51.77.52.79', 'base_url': 'http://51.77.52.79/ints', 'url_hint': 'http://51.77.52.79/ints/agent/SMSCDRStats', 'username': 'Rabbi12', 'password': 'Rabbi12', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': None},
    {'id': 'bp9', 'host': '51.210.208.26', 'base_url': 'http://51.210.208.26/ints', 'url_hint': 'http://51.210.208.26/ints/agent/SMSCDRStats', 'username': 'Dasbabu50_FD', 'password': 'Dasbabu50_FD', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': None},
    {'id': 'bp10', 'host': 'ivasms.com', 'base_url': 'https://ivasms.com', 'url_hint': 'https://ivasms.com/portal/sms/received', 'username': 'mdrashub2@gmail.com', 'password': 'Rabbi+nnn', 'engine': 'iva_sms', 'data_path': '/portal/sms/received', 'admin_id': None},
    {'id': 'bp11', 'host': '139.99.68.231', 'base_url': 'http://139.99.68.231/ints', 'url_hint': 'http://139.99.68.231/ints/agent/SMSCDRStats', 'username': 'Rabbi12', 'password': 'Rabbi12', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': None},
    {'id': 'bp12', 'host': '51.75.144.178', 'base_url': 'http://51.75.144.178/ints', 'url_hint': 'http://51.75.144.178/ints/agent/SMSCDRStats', 'username': 'Rabbi12', 'password': 'Rabbi12@', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': None},
    {'id': 'd20591', 'host': '54.39.104.241', 'base_url': 'http://54.39.104.241/ints', 'url_hint': 'http://54.39.104.241/ints/client/SMSCDRStats', 'username': 'Atik9898', 'password': 'Atik9898', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': 6664150885},
    {'id': 'd20984', 'host': '54.39.104.241', 'base_url': 'http://54.39.104.241/ints', 'url_hint': 'http://54.39.104.241/ints/agent/SMSCDRStats', 'username': 'Rabbi5', 'password': 'Rabbi5', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': 6664150885},
    {'id': 'pscall1', 'host': 'pscall.net', 'base_url': 'http://pscall.net/restapi', 'url_hint': 'http://pscall.net/restapi/smsreport', 'username': 'api:pscall.net', 'password': '', 'api_key': 'SFNSQz1SS2NygIF6QlBR', 'api_key_param': 'key', 'engine': 'api_key', 'data_path': '/smsreport', 'admin_id': None},
    {'id': 'd34527', 'host': '151.80.19.204', 'base_url': 'http://151.80.19.204/ints', 'url_hint': 'http://151.80.19.204/ints/agent/SMSCDRStats', 'username': 'Atik9898', 'password': 'Atik9898', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': 6664150885},
    {'id': 'd39020', 'host': '139.99.69.196', 'base_url': 'http://139.99.69.196/ints', 'url_hint': 'http://139.99.69.196/ints/agent/SMSCDRStats', 'username': 'Mahofuza12', 'password': 'Mahofuza12', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': 6664150885},
    {'id': 'd42653', 'host': '54.39.104.241', 'base_url': 'http://54.39.104.241/ints', 'url_hint': 'http://54.39.104.241/ints/agent/SMSCDRStats', 'username': 'Rabbi5', 'password': 'Rabbi5', 'engine': 'ints_smsranges', 'data_path': '/agent/res/data_smsranges.php', 'admin_id': 6664150885},
    {'id': 'd26084', 'host': '91.232.105.47', 'base_url': 'http://91.232.105.47/ints', 'url_hint': 'http://91.232.105.47/ints/agent/SMSCDRStats', 'username': 'Mahofuza', 'password': 'Mahofuza12@', 'engine': 'ints_smscdr', 'data_path': '/agent/res/data_smscdr.php', 'admin_id': 6664150885},
]
# <<SYNC:_BUILTIN_PANELS:END>>


POLL_INTERVAL = 8
DATA_FILE = "stock_data.json"
USERS_FILE = "users.json"
SEEN_FILE = "seen_otps.json"

bot = telebot.TeleBot(API_TOKEN, threaded=True, num_threads=40)

# ── Persistent helpers ────────────────────────────────────────────────────────


def load_json(path, default):
    if os.path.exists(path):
        try:
            with open(path) as f:
                return json.load(f)
        except Exception:
            pass
    return default


def save_json(path, data):
    with open(path, "w") as f:
        json.dump(data, f, indent=2)


stock = load_json(
    DATA_FILE,
    {
        "whatsapp": {},
        "facebook": {},
        "telegram": {},
        "instagram": {},
        "pc clone": {},
        "binance": {},
    },
)
users = load_json(USERS_FILE, [])
seen_otps = load_json(SEEN_FILE, {})

USER_NAMES_FILE = "user_names.json"
user_names = load_json(USER_NAMES_FILE, {})

ADMINS_FILE = "admins.json"
ADMIN_EXPIRY_FILE = "admin_expiry.json"
ADMIN_SETTINGS_FILE = "admin_settings.json"

_extra_admins = load_json(ADMINS_FILE, [])
for _aid in _extra_admins:
    if _aid not in ADMIN_IDS:
        ADMIN_IDS.append(_aid)

# {str(uid): expiry_unix_timestamp}  — None means permanent
_admin_expiry = load_json(ADMIN_EXPIRY_FILE, {})

# per-admin settings: {str(uid): {otp_group_link, otp_group_id, channel2, bot_link}}
_admin_settings = load_json(ADMIN_SETTINGS_FILE, {})


def is_super_admin(uid):
    return uid in SUPER_ADMIN_IDS


def save_admins():
    save_json(ADMINS_FILE, [a for a in ADMIN_IDS if a not in SUPER_ADMIN_IDS])
    _sync_settings_to_botpy()


def save_admin_expiry():
    save_json(ADMIN_EXPIRY_FILE, _admin_expiry)


def save_admin_settings():
    save_json(ADMIN_SETTINGS_FILE, _admin_settings)


def get_admin_setting(uid, key, default=""):
    """Return per-admin setting, fall back to global group_settings."""
    return _admin_settings.get(str(uid), {}).get(key, _group_settings.get(key, default))


def add_admin(uid, months=None):
    """Add uid as admin. months=None means permanent (super admin only can set)."""
    if uid in SUPER_ADMIN_IDS:
        return False  # already super admin
    if uid not in ADMIN_IDS:
        ADMIN_IDS.append(uid)
    if months:
        expiry = time.time() + months * 30 * 24 * 3600
        _admin_expiry[str(uid)] = expiry
    else:
        _admin_expiry.pop(str(uid), None)
    save_admins()
    save_admin_expiry()
    return True


def remove_admin(uid):
    if uid in SUPER_ADMIN_IDS:
        return False
    if uid in ADMIN_IDS:
        ADMIN_IDS.remove(uid)
        _admin_expiry.pop(str(uid), None)
        _admin_settings.pop(str(uid), None)
        save_admins()
        save_admin_expiry()
        save_admin_settings()
        return True
    return False


def _admin_expiry_checker():
    """Background thread: remove expired admins every 10 minutes."""
    while True:
        time.sleep(600)
        now = time.time()
        to_remove = [
            int(uid) for uid, exp in list(_admin_expiry.items())
            if exp and now >= exp
        ]
        for uid in to_remove:
            remove_admin(uid)
            try:
                bot.send_message(
                    uid,
                    "⚠️ <b>Admin Access Expired!</b>\n\n"
                    "Tomar admin access er meiad shesh hoy geche.\n"
                    "Notun access er jonno admin-er sathe jogajog koro.",
                    parse_mode="HTML",
                )
            except Exception:
                pass


threading.Thread(target=_admin_expiry_checker, daemon=True).start()

GROUP_SETTINGS_FILE = "group_settings.json"
# <<SYNC:_group_settings_defaults:START>>
_group_settings = load_json(GROUP_SETTINGS_FILE, {
    'otp_group_id': -1003850531522,
    'otp_group_link': 'https://t.me/+BC0-N3KJkiYyOTE1',
    'auto_delete': True,
    'auto_delete_seconds': 3600,
    'channel2': 'https://t.me/easy_earn_with_Atik',
    'bot_link': 'https://t.me/hot_otp_bot',
    'support_id': '',
    'group_otp_send': True,
})
# <<SYNC:_group_settings_defaults:END>>

CHANNEL_1 = _group_settings["otp_group_link"]
OTP_GROUP_ID = _group_settings["otp_group_id"]


def save_group_settings():
    save_json(GROUP_SETTINGS_FILE, _group_settings)
    _sync_settings_to_botpy()


def get_otp_group_id():
    return _group_settings.get("otp_group_id")


def get_otp_group_link():
    return _group_settings.get("otp_group_link", "")


def _extract_username(link):
    """Extract @username from a t.me link for use with get_chat_member."""
    if not link:
        return None
    link = link.strip().rstrip("/")
    if "joinchat" in link or "/+" in link:
        return None
    if "t.me/" in link:
        uname = link.split("t.me/")[-1].split("/")[0]
        if uname:
            return "@" + uname
    return None


def _check_member(chat_ref, user_id):
    """Returns True if member, False if not, None if cannot check."""
    if not chat_ref:
        return None
    try:
        m = bot.get_chat_member(chat_ref, user_id)
        return m.status not in ("left", "kicked")
    except Exception:
        return None


def get_channel2():
    return _group_settings.get("channel2", "")


def get_bot_link():
    return _group_settings.get("bot_link", "")


def is_auto_delete():
    return _group_settings.get("auto_delete", True)


def _schedule_delete(chat_id, msg_id):
    delay = _group_settings.get("auto_delete_seconds", 3600)
    def _do_delete():
        try:
            bot.delete_message(chat_id, msg_id)
        except Exception:
            pass
    threading.Timer(delay, _do_delete).start()

# ── Message templates ──────────────────────────────────────────────────────────

TEMPLATES_FILE = "message_templates.json"
# <<SYNC:_DEFAULT_TEMPLATES:START>>
_DEFAULT_TEMPLATES = {
    'start': '🔥 <b>𝗔𝗥 𝗢𝗧𝗣 𝗕𝗢𝗧-𝗲 𝗦𝗔𝗚𝗢𝗧𝗢𝗠!</b> 🔥\n\n╔═════════════════════════════╗\n   🧾 <b>USER DASHBOARD</b>\n╠═════════════════════════════╣\n  👤 <b>User:</b> {uname}\n  🆔 <b>ID:</b> <code>{uid}</code>\n  📊 <b>Status:</b> 💎 Premium\n  🚀 <b>Workers:</b> 0\n╚═════════════════════════════╝\n\n╔══════════════════╗\n 𝗡𝗶𝗰𝗵𝗲𝗿 𝗰𝗵𝗮𝗻𝗻𝗲𝗹𝗲 <b>𝗝𝗢𝗜𝗡</b> 𝗵𝗼𝘆𝗲\n <b>𝗩𝗘𝗥𝗜𝗙𝗬</b> 𝗯𝗮𝘁𝗮𝗻𝗲 𝗰𝗹𝗶𝗰𝗸 𝗸𝗼𝗿𝗼!\n╚══════════════════╝\n\n🤖 <i>𝙋𝙤𝙬𝙚𝙧𝙚𝙙 𝙗𝙮</i>  <b>𝗔𝗥 𝗢𝗧𝗣 𝗕𝗢𝗧</b>',
    'otp_group': '📱 <b><i>{svc}</i></b>{flag} | {number} | {flag}\n\n🔑 <b>KEY</b> : <b><i>{otp}</i></b>\n\n╭─────────────────╮\n📩 <b>MESSAGE</b>\n\n<i>{sms}</i>\n╰─────────────────╯\n💬 <i>Thanks for using</i> <b>@hot_otp_bot</b>',
    'otp_dm': '📱 <b><i>{svc}</i></b>{flag} | {number} | {flag}\n\n🔑 <b>KEY</b> : <b><i>{otp}</i></b>\n\n╭─────────────────╮\n📩 <b>MESSAGE</b>\n\n<i>{sms}</i>\n╰─────────────────╯\n💬 <i>Thanks for using</i> <b>@hot_otp_bot</b>',
    'verify_success': '🔥 <b>VERIFICATION COMPLETE!</b> 🔥\n\n╔═════════════════════════════╗\n   ✅ <b>ACCESS GRANTED</b>\n╠═════════════════════════════╣\n  👋 <b>Welcome, {vname}!</b>\n  🆔 <b>ID:</b> <code>{uid}</code>\n  📊 <b>Status:</b> 💎 Premium\n╚═════════════════════════════╝\n\n⚡ <b>𝗘𝗸𝗸𝗵𝗼𝗻 𝗻𝘂𝗺𝗯𝗮𝗿 𝗻𝗶𝘁𝗲 𝗽𝗮𝗿𝗯𝗲!</b> ⚡',
    'number_assigned': '✅ <b>Number Assigned Successfully !</b>\n\n🔧 <b>Platform :</b> {svc}\n🌍 <b>Country :</b> {flag} {country}\n\n📞 <b>Number :</b> <code>{number}</code>\n\n⏱ <b>Auto code fetch :</b> 10:00s',
    'broadcast': '🔥 <b>𝗔𝗥 𝗢𝗧𝗣 𝗕𝗢𝗧 — 𝗕𝗥𝗢𝗔𝗗𝗖𝗔𝗦𝗧!</b> 🔥\n⚡━━━━━━━━━━━━━━━━⚡\n\n📢 {text} 📢\n\n⚡━━━━━━━━━━━━━━━━⚡\n🤖🔥 <i>𝙋𝙤𝙬𝙚𝙧𝙚𝙙 𝙗𝙮</i>  <b>𝗔𝗥 𝗢𝗧𝗣 𝗕𝗢𝗧</b>  🔥🤖',
}
# <<SYNC:_DEFAULT_TEMPLATES:END>>
_templates = load_json(TEMPLATES_FILE, dict(_DEFAULT_TEMPLATES))
for _k, _v in _DEFAULT_TEMPLATES.items():
    if _k not in _templates:
        _templates[_k] = _v
_edit_template_state = {}


def _fmt_pyval(val, indent=0):
    """Format a Python value as readable source code."""
    pad = "    " * indent
    inner = "    " * (indent + 1)
    if isinstance(val, dict):
        if not val:
            return "{}"
        lines = ["{"]
        for k, v in val.items():
            lines.append(f"{inner}{repr(k)}: {repr(v)},")
        lines.append(f"{pad}}}")
        return "\n".join(lines)
    elif isinstance(val, list):
        if not val:
            return "[]"
        lines = ["["]
        for item in val:
            lines.append(f"{inner}{repr(item)},")
        lines.append(f"{pad}]")
        return "\n".join(lines)
    return repr(val)


def _sync_block(source, marker_name, new_content):
    """Replace content between <<SYNC:X:START>> and <<SYNC:X:END>> markers."""
    start_marker = f"# <<SYNC:{marker_name}:START>>"
    end_marker   = f"# <<SYNC:{marker_name}:END>>"
    s = source.find(start_marker)
    e = source.find(end_marker)
    if s == -1 or e == -1:
        return source
    return (
        source[:s + len(start_marker)] + "\n" +
        new_content + "\n" +
        source[e:]
    )


def _sync_settings_to_botpy():
    """Auto-patch bot.py so its hardcoded defaults always match live settings.
    This ensures that when bot.py is pushed to Railway/any server, all panels,
    admin IDs, and other settings are already baked in — no data loss on redeploy.
    """
    try:
        bot_file = os.path.abspath(__file__)
        with open(bot_file, "r", encoding="utf-8") as f:
            source = f.read()

        # Sync message templates
        source = _sync_block(
            source, "_DEFAULT_TEMPLATES",
            f"_DEFAULT_TEMPLATES = {_fmt_pyval(_templates)}"
        )
        # Sync services list
        source = _sync_block(
            source, "_DEFAULT_SERVICES",
            f"_DEFAULT_SERVICES = {_fmt_pyval(_services)}"
        )
        # Sync group settings defaults
        source = _sync_block(
            source, "_group_settings_defaults",
            f"_group_settings = load_json(GROUP_SETTINGS_FILE, {_fmt_pyval(_group_settings)})"
        )

        # ── Sync SUPER_ADMIN_IDS ──────────────────────────────────────────────
        source = _sync_block(
            source, "SUPER_ADMIN_IDS",
            f"SUPER_ADMIN_IDS = {_fmt_pyval(SUPER_ADMIN_IDS)}"
        )

        # ── Sync _BUILTIN_PANELS (merge with dynamic panels) ─────────────────
        # Build merged panel list: start with all dynamic panels, then add any
        # BUILTIN panels not already present (by id). Dynamic panels added via
        # /addpanel are promoted to hardcoded — so Railway always has them.
        _seen_ids = set()
        _merged_panels = []
        for _p in (_dynamic_panels + _BUILTIN_PANELS):
            if _p.get("id") and _p["id"] not in _seen_ids:
                _seen_ids.add(_p["id"])
                # Strip runtime-only keys that shouldn't be hardcoded
                _clean = {k: v for k, v in _p.items() if k not in ("cookie_str",)}
                _merged_panels.append(_clean)
        source = _sync_block(
            source, "_BUILTIN_PANELS",
            f"_BUILTIN_PANELS = {_fmt_pyval(_merged_panels)}"
        )

        with open(bot_file, "w", encoding="utf-8") as f:
            f.write(source)
        print("[SYNC] ✅ bot.py auto-patched with latest settings (panels, admins, templates)")
    except Exception as e:
        print(f"[SYNC] ❌ Failed to patch bot.py: {e}")


def save_templates():
    save_json(TEMPLATES_FILE, _templates)
    _sync_settings_to_botpy()


def get_template(key):
    return _templates.get(key, _DEFAULT_TEMPLATES.get(key, ""))


_TEMPLATE_LABELS = {
    "start": "🚀 Start / Welcome মেসেজ",
    "otp_group": "📲 OTP মেসেজ (Group)",
    "otp_dm": "📲 OTP মেসেজ (DM/User)",
    "verify_success": "✅ Verify Success মেসেজ",
    "number_assigned": "☎️ Number Assigned মেসেজ",
    "broadcast": "📢 Broadcast মেসেজ",
}

_TEMPLATE_VARS = {
    "start": "{uname} = ইউজার নাম, {uid} = ইউজার আইডি",
    "otp_group": "{svc} = সার্ভিস, {number} = নম্বর, {country} = দেশ, {flag} = ফ্ল্যাগ, {otp} = OTP কোড, {sms} = পূর্ণ SMS টেক্সট",
    "otp_dm": "{svc} = সার্ভিস, {number} = নম্বর, {country} = দেশ, {flag} = ফ্ল্যাগ, {otp} = OTP কোড, {sms} = পূর্ণ SMS টেক্সট",
    "verify_success": "{vname} = ইউজার নাম, {uid} = ইউজার আইডি",
    "number_assigned": "{svc} = সার্ভিস, {country} = দেশ, {flag} = ফ্ল্যাগ, {number} = নম্বর",
    "broadcast": "{text} = broadcast content",
}

# ── End Message templates ──────────────────────────────────────────────────────

SERVICES_FILE = "services.json"
# <<SYNC:_DEFAULT_SERVICES:START>>
_DEFAULT_SERVICES = [
    {'label': 'Instagram →', 'key': 'instagram'},
    {'label': 'Facebook 💎', 'key': 'facebook'},
    {'label': 'WhatsApp', 'key': 'whatsapp'},
    {'label': 'PC Clone 💎', 'key': 'pc clone'},
]
# <<SYNC:_DEFAULT_SERVICES:END>>
_services = load_json(SERVICES_FILE, list(_DEFAULT_SERVICES))
_addservice_state = {}
_countdowns = {}

USER_MAP_FILE = "user_map.json"
_raw_user_map = load_json(USER_MAP_FILE, {})
user_map: dict[str, int] = {k: int(v) for k, v in _raw_user_map.items()}
user_map_lock = threading.Lock()
assigned_time: dict[str, float] = {}

OTP_STATS_FILE = "otp_stats.json"
otp_stats: dict[str, int] = load_json(OTP_STATS_FILE, {})
otp_stats_lock = threading.Lock()


def _save_otp_stats():
    with otp_stats_lock:
        save_json(OTP_STATS_FILE, otp_stats)

# Tracks last service+country per user so OTP message buttons know what to request
_user_last_svc: dict[int, tuple] = {}   # uid -> (svc, scnt)
# Tracks last "active number/OTP" message_id per user so buttons can be stripped
_user_last_num_msg: dict[int, int] = {} # uid -> message_id


def _save_user_map():
    with user_map_lock:
        save_json(USER_MAP_FILE, user_map)


def register_number(user_id, number):
    clean = re.sub(r"\D", "", str(number))
    with user_map_lock:
        user_map[clean] = user_id
        assigned_time[clean] = time.time()
    _save_user_map()


def mask_number(number):
    s = re.sub(r"\D", "", str(number))
    _bold = lambda t: "".join(
        chr(0x1D5D4 + ord(c) - 65) if 'A' <= c <= 'Z' else
        chr(0x1D5EE + ord(c) - 97) if 'a' <= c <= 'z' else c
        for c in str(t)
    )
    if len(s) >= 7:
        return f"{s[:3]}⚜{_bold('ATIK')}⚜{s[-4:]}"
    return s


# ── OTP Messages ──────────────────────────────────────────────────────────────


def _ensure_code_tag(text, value):
    """Wrap `value` in <code> if not already wrapped."""
    v = str(value)
    if f"<code>{v}</code>" in text:
        return text
    return text.replace(v, f"<code>{v}</code>", 1)


def _send_with_retry(fn, max_retries=3, **kwargs):
    """Call fn(**kwargs) with up to max_retries on 429 rate-limit errors.
    Returns (result, rate_limit_seconds) tuple:
      - result: the API result or None on failure
      - rate_limit_seconds: >0 if all retries exhausted due to rate limit, else 0
    """
    last_wait = 0
    for attempt in range(max_retries):
        try:
            return fn(**kwargs), 0
        except Exception as e:
            err = str(e)
            if "429" in err or "Too Many Requests" in err:
                try:
                    last_wait = int(re.search(r"retry after (\d+)", err).group(1))
                except Exception:
                    last_wait = 30
                capped = min(last_wait, 90)
                print(f"[RETRY] 429 for chat={kwargs.get('chat_id','?')} retry_after={last_wait}s — waiting {capped}s (attempt {attempt+1}/{max_retries})")
                time.sleep(capped)
            else:
                raise
    print(f"[RETRY] All {max_retries} attempts failed for chat={kwargs.get('chat_id','?')} — rate limited {last_wait}s")
    return None, last_wait


def send_otp_message(chat_id, otp, number, seconds, service="", sms=""):
    svc = service.upper() if service else "—"
    c_name, flag = get_country_details(number)
    otp_str = str(otp)
    sms_str = str(sms) if sms else "—"
    _grp_vars = dict(svc=svc, number=mask_number(number), country=c_name, flag=flag, otp=otp_str, sms=sms_str)
    _dm_vars  = dict(svc=svc, number=(number if str(number).startswith("+") else "+" + str(number)),
                     country=c_name, flag=flag, otp=otp_str, sms=sms_str)

    def _build_message(key, vars_dict):
        """Return (text, used_default). Falls back to default on any error."""
        try:
            txt = get_template(key).format(**vars_dict)
            return _ensure_code_tag(txt, otp_str), False
        except Exception as e:
            print(f"[TEMPLATE] ⚠️ Custom template '{key}' format error, using default: {e}")
        txt = _DEFAULT_TEMPLATES[key].format(**vars_dict)
        return _ensure_code_tag(txt, otp_str), True

    def _try_send(label, chat_id, text, markup):
        """Send message; if Telegram rejects it return (None, err_str)."""
        try:
            result, rl = _send_with_retry(bot.send_message,
                                          chat_id=chat_id, text=text,
                                          parse_mode="HTML", reply_markup=markup)
            return result, rl, None
        except Exception as e:
            return None, 0, str(e)

    if chat_id == get_otp_group_id():
        markup = types.InlineKeyboardMarkup()
        markup.row(
            types.InlineKeyboardButton(
                f"📋🔒 {otp_str}",
                copy_text=types.CopyTextButton(text=otp_str)
            )
        )
        _btns = []
        if get_bot_link():
            _btns.append(types.InlineKeyboardButton("📞 𝗡𝗨𝗠𝗕𝗘𝗥 ↗", url=get_bot_link()))
        if get_channel2():
            _btns.append(types.InlineKeyboardButton("📣 𝗖𝗛𝗔𝗡𝗡𝗘𝗟 ↗", url=get_channel2()))
        if _btns:
            markup.row(*_btns)

        message, used_default = _build_message("otp_group", _grp_vars)
        sent, rl, err = _try_send("GROUP", chat_id, message, markup)

        # If custom template caused a send error, retry with default
        if err and not used_default:
            print(f"[OTP-GROUP] ⚠️ Send failed (custom template HTML error?): {err} — retrying with default")
            message = _ensure_code_tag(_DEFAULT_TEMPLATES["otp_group"].format(**_grp_vars), otp_str)
            sent, rl, err = _try_send("GROUP-DEFAULT", chat_id, message, markup)

        if err:
            print(f"[OTP-GROUP] ❌ Exception sending to group {chat_id}: {err}")
        elif sent:
            print(f"[OTP-GROUP] ✅ Sent OTP={otp_str} num={mask_number(number)} svc={svc} to group {chat_id}")
            if is_auto_delete():
                _schedule_delete(chat_id, sent.message_id)
        else:
            print(f"[OTP-GROUP] ❌ FAILED to send OTP={otp_str} num={mask_number(number)} — rate limited {rl}s")
    else:
        uid = chat_id  # DM: chat_id == user_id
        last_svc_info = _user_last_svc.get(uid)
        dm_markup = types.InlineKeyboardMarkup(row_width=2)
        if last_svc_info:
            _svc, _scnt = last_svc_info
            dm_markup.add(
                types.InlineKeyboardButton("🔄 𝗚𝗲𝘁 𝗡𝗲𝘄 𝗡𝘂𝗺𝗯𝗲𝗿", callback_data=f"n:{_svc}:{_scnt}"),
                types.InlineKeyboardButton("🌍 𝗖𝗵𝗮𝗻𝗴𝗲 𝗖𝗼𝘂𝗻𝘁𝗿𝘆", callback_data=f"s:{_svc}"),
            )
        if get_otp_group_link():
            dm_markup.add(
                types.InlineKeyboardButton("📢 𝗢𝗧𝗣 𝗚𝗿𝗼𝘂𝗽", url=get_otp_group_link()),
            )

        # Delete the previous "Number Assigned" message when OTP arrives
        prev_msg_id = _user_last_num_msg.get(uid)
        if prev_msg_id:
            try:
                bot.delete_message(chat_id=uid, message_id=prev_msg_id)
            except Exception:
                pass
            _user_last_num_msg.pop(uid, None)

        message, used_default = _build_message("otp_dm", _dm_vars)
        result, rl, err = _try_send("DM", chat_id, message, dm_markup)

        # If custom template caused a send error, retry with default
        if err and not used_default:
            print(f"[OTP-DM] ⚠️ Send failed (custom template HTML error?): {err} — retrying with default")
            message = _ensure_code_tag(_DEFAULT_TEMPLATES["otp_dm"].format(**_dm_vars), otp_str)
            result, rl, err = _try_send("DM-DEFAULT", chat_id, message, dm_markup)

        if err:
            print(f"[OTP-DM] ❌ Exception sending to user {chat_id}: {err}")
        elif result:
            print(f"[OTP-DM] ✅ Sent OTP={otp_str} to user {chat_id}")
            # Do NOT store OTP message in _user_last_num_msg —
            # that tracker is only for "Number Assigned" messages
        else:
            print(f"[OTP-DM] ❌ FAILED to send OTP={otp_str} to user {chat_id} — rate limited {rl}s")


def is_group_otp_send_enabled():
    return _group_settings.get("group_otp_send", True)


def _dispatch_otp(otp, number, seconds, service="", sms=""):
    grp = get_otp_group_id()
    clean = re.sub(r"\D", "", str(number))
    with user_map_lock:
        uid = user_map.get(clean)
    print(f"[DISPATCH] OTP={otp} num={number} svc={service} group={grp} user_dm={uid} grp_send={is_group_otp_send_enabled()}")
    if grp and is_group_otp_send_enabled():
        send_otp_message(grp, otp, number, seconds, service, sms=sms)
    elif grp and not is_group_otp_send_enabled():
        print(f"[DISPATCH] ℹ️ Group OTP send is DISABLED — skipping group send (DM only mode)")
    else:
        print(f"[DISPATCH] ⚠️ No OTP group configured — skipping group send!")
    if uid:
        send_otp_message(uid, otp, number, seconds, service, sms=sms)
        # Track OTP receive count per user
        with otp_stats_lock:
            otp_stats[str(uid)] = otp_stats.get(str(uid), 0) + 1
        _save_otp_stats()
        # Auto-release: number এ OTP আসলে সাথে সাথে delete হয়ে যাবে
        with user_map_lock:
            user_map.pop(clean, None)
            assigned_time.pop(clean, None)
        _save_user_map()
        print(f"[DISPATCH] 🗑️ Auto-released number {number} after OTP delivery to user {uid}")
    else:
        print(f"[DISPATCH] ℹ️ No user DM mapping for {number} — DM skipped")


def send_status_message(chat_id, status_text):
    message = (
        "⚙️ <b>𝗦𝗧𝗔𝗧𝗨𝗦 𝗔𝗟𝗘𝗥𝗧</b> ⚙️\n"
        "🔥━━━━━━━━━━━━━━🔥\n\n"
        f"📛 {status_text} 📛\n\n"
        "🔥━━━━━━━━━━━━━━🔥\n"
        "🤖⚡ <b>𝗔𝗥 𝗢𝗧𝗣 𝗕𝗢𝗧 — 𝗔𝗖𝗧𝗜𝗩𝗘</b> ⚡🤖"
    )
    try:
        bot.send_message(chat_id=chat_id, text=message, parse_mode="HTML")
    except Exception as e:
        print(f"[MONITOR] Status send error: {e}")


# ── Country helpers ───────────────────────────────────────────────────────────


def get_country_details(num_str):
    try:
        num_str = str(num_str).strip()
        if not num_str.startswith("+"):
            num_str = "+" + num_str
        parsed = phonenumbers.parse(num_str)
        country_code = region_code_for_number(parsed)
        country_name = geocoder.description_for_number(parsed, "en")
        flag = "".join(chr(ord(c.upper()) + 127397) for c in country_code)
        return country_name, flag
    except Exception:
        return "Unknown", "🌐"


# ── Stock helpers ─────────────────────────────────────────────────────────────


def save_stock():
    save_json(DATA_FILE, stock)


def register_user(chat_id, first_name="", last_name="", username=""):
    if chat_id not in users:
        users.append(chat_id)
        save_json(USERS_FILE, users)
    full = f"{first_name} {last_name}".strip()
    if full and username:
        display = f"{full} (@{username})"
    elif full:
        display = full
    elif username:
        display = f"@{username}"
    else:
        display = None
    if display:
        user_names[str(chat_id)] = display
        save_json(USER_NAMES_FILE, user_names)


# ── Panel sessions ────────────────────────────────────────────────────────────

_p1_session = None
_p1_sesskey = None
_p1_lock = threading.Lock()

_p2_session = None
_p2_lock = threading.Lock()

_p3_session = None
_p3_csstr = None
_p3_lock = threading.Lock()

_p4_session = None
_p4_sesskey = None
_p4_lock = threading.Lock()

_p5_session = None
_p5_sesskey = None
_p5_lock = threading.Lock()

_p6_session = None
_p6_sesskey = None
_p6_lock = threading.Lock()


# ── Panel stats (for /panels command) ─────────────────────────────────────────
_panel_stats = {
    "p1": {
        "name": "Mahofuza",
        "host": "91.232.105.47",
        "status": "⏳",
        "count": 0,
        "last": None,
        "errors": 0,
    },
    "p2": {
        "name": "Sagardas50",
        "host": "94.23.31.29",
        "status": "⏳",
        "count": 0,
        "last": None,
        "errors": 0,
    },
    "p3": {
        "name": "Rabbi1_FD",
        "host": "168.119.13.175",
        "status": "⏳",
        "count": 0,
        "last": None,
        "errors": 0,
    },
    "p4": {
        "name": "Rabbi12",
        "host": "144.217.71.192",
        "status": "⏳",
        "count": 0,
        "last": None,
        "errors": 0,
    },
    "p5": {
        "name": "Rabbi12_v2",
        "host": "51.75.144.178",
        "status": "⏳",
        "count": 0,
        "last": None,
        "errors": 0,
    },
    "p6": {
        "name": "TrueSMS/Ranges",
        "host": "truesms.net",
        "status": "⏳",
        "count": 0,
        "last": None,
        "errors": 0,
    },
}
_stats_lock = threading.Lock()


def _record_fetch(pid, count):
    with _stats_lock:
        _panel_stats[pid]["status"] = "🟢"
        _panel_stats[pid]["count"] = count
        _panel_stats[pid]["last"] = time.time()
        _panel_stats[pid]["errors"] = 0


def _record_error(pid):
    with _stats_lock:
        _panel_stats[pid]["status"] = "🔴"
        _panel_stats[pid]["errors"] += 1


# ── Demo OTP state ─────────────────────────────────────────────────────────────
_demo_active = False
_demo_lock = threading.Lock()
_demo_cfg_id_counter = 1
_demo_configs: list = [
    {"id": 1, "name": "Config 1", "active": False, "numbers": ["8801700000000"], "digits": 6, "services": ["Facebook"], "interval": 30}
]
_demo_next_fire: dict = {}
_demo_svc_state: dict = {}
_demo_cfg_temp: dict = {}

seen_lock = threading.Lock()

# ── Dynamic panel system ───────────────────────────────────────────────────────
DYNAMIC_PANELS_FILE = "dynamic_panels.json"
_dynamic_panels = load_json(DYNAMIC_PANELS_FILE, [])
_dynamic_sessions = {}
_dynamic_locks = {}
_addpanel_state = {}
_testpanel_state = {}
_pending_force_add = {}   # panel_id → panel dict (login failed, user wants force-add)
_pending_excel = {}  # uid → {'numbers': [...], 'filename': str}

# Migrate old panels that use panel_type → new engine/data_path format
def _migrate_dynamic_panels():
    changed = False
    for p in _dynamic_panels:
        if "panel_type" in p and "engine" not in p:
            pt = p.pop("panel_type", "smscdr")
            p["engine"] = "ints_smsranges" if pt == "smsranges" else "ints_smscdr"
            p["data_path"] = (
                "/agent/res/data_smsranges.php" if pt == "smsranges"
                else "/agent/res/data_smscdr.php"
            )
            changed = True
        if "engine" not in p:
            p["engine"] = "ints_smscdr"
            p["data_path"] = "/agent/res/data_smscdr.php"
            changed = True
    if changed:
        save_json(DYNAMIC_PANELS_FILE, _dynamic_panels)
        print(f"[MIGRATE] Migrated {len(_dynamic_panels)} dynamic panels to universal format")

_migrate_dynamic_panels()


def save_dynamic_panels():
    save_json(DYNAMIC_PANELS_FILE, _dynamic_panels)
    _sync_settings_to_botpy()


def _get_dp_lock(pid):
    if pid not in _dynamic_locks:
        _dynamic_locks[pid] = threading.Lock()
    return _dynamic_locks[pid]


# ── Universal Panel Engine ─────────────────────────────────────────────────────
# Supports any SMS panel: INTS (math captcha), Xisora, or custom panels.
# Auto-detects login page, captcha, signin path, token, and data endpoint.

_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

# All known login page paths to try (in order)
_LOGIN_PAGE_PATHS = [
    "/login", "/signin", "/signmein",
    "/ints/login", "/sms/login", "/konekta/login",
    "/admin/login", "/user/login", "/agent/login", "/panel/login",
    "/index.php", "/",
]

# All known signin (POST) paths to try
_SIGNIN_PATHS = [
    "/signin", "/signmein", "/login",
    "/ints/signin", "/sms/signin", "/konekta/signin",
    "/admin/signin", "/user/signin", "/panel/signin",
    "/ints/login", "/sms/login", "/konekta/login",
    "/index.php", "/signIn", "/auth/login", "/auth/signin",
]

# All known data endpoints: (path, param_style, engine_name)
# param_style: "ints" or "xisora"
_DATA_ENDPOINTS = [
    ("/agent/res/data_smscdr.php",              "ints",   "ints_smscdr"),
    ("/agent/res/data_smsranges.php",            "ints",   "ints_smsranges"),
    ("/agent/res/data_smscdrreports.php",        "ints",   "ints_smscdr"),
    ("/ints/agent/res/data_smscdr.php",          "ints",   "ints_smscdr"),
    ("/ints/agent/res/data_smsranges.php",       "ints",   "ints_smsranges"),
    ("/sms/agent/res/data_smscdr.php",           "ints",   "ints_smscdr"),
    ("/sms/agent/res/data_smsranges.php",        "ints",   "ints_smsranges"),
    ("/konekta/agent/res/data_smscdr.php",       "ints",   "ints_smscdr"),
    ("/konekta/agent/res/data_smsranges.php",    "ints",   "ints_smsranges"),
    ("/client/ajax/dt_reports.php",              "xisora", "xisora"),
    ("/client/ajax/dt_smscdr.php",               "xisora", "xisora"),
    ("/api/sms/cdr",                             "ints",   "ints_smscdr"),
]

# Dashboard pages to probe for sesskey/csstr token
_DASHBOARD_PATHS = [
    "/agent/SMSCDRStats", "/agent/SMSRanges", "/agent/SMSCDRReports",
    "/ints/agent/SMSCDRStats", "/ints/agent/SMSRanges", "/ints/agent/SMSCDRReports",
    "/sms/agent/SMSCDRStats", "/sms/agent/SMSRanges",
    "/konekta/agent/SMSCDRStats", "/konekta/agent/SMSRanges", "/konekta/agent/SMSCDRReports",
    "/agent/", "/dashboard", "/admin/", "/home",
]


def _extract_panel_base_url(raw_url: str) -> str | None:
    """
    Extract the panel base URL (scheme+host+any-path-prefix) from ANY panel URL.
    Handles paths like /konekta/agent/SMSCDRReports, /ints/login, etc.
    """
    url = raw_url.strip().split("?")[0].split("#")[0].rstrip("/")
    if not re.match(r"https?://", url, re.IGNORECASE):
        return None
    # Strip from first occurrence of known endpoint/action segment onwards
    cleaned = re.sub(
        r"/(?:agent|login|signin|signmein|client|api|dashboard|auth)(?:/.*)?$",
        "", url, flags=re.IGNORECASE,
    )
    return cleaned.rstrip("/") or url


def _univ_build_url(base_endpoint: str, token: str, date_str: str, style: str) -> str:
    if style == "xisora":
        return (
            f"{base_endpoint}"
            f"?fdate1={date_str}%2000:00:00&fdate2={date_str}%2023:59:59"
            f"&ftermination=&fclient=&fnum=&fcli="
            f"&fgdate=0&fgtermination=0&fgclient=0&fgnumber=0&fgcli=0&fg=0"
        )
    base_q = (
        f"{base_endpoint}"
        f"?fdate1={date_str}%2000:00:00&fdate2={date_str}%2023:59:59"
        f"&frange=&fclient=&fnum=&fcli=&fgdate=&fgmonth="
        f"&fgrange=&fgclient=&fgnumber=&fgcli=&fg=0"
    )
    # Only append sesskey when token is actually present (cookie-based panels don't need it)
    if token:
        base_q += f"&sesskey={token}"
    return base_q


def _univ_extract_token(html: str) -> str:
    sk = re.search(r"sesskey=([A-Za-z0-9+/=]+)", html)
    if sk:
        return sk.group(1)
    cs = re.search(r"csstr=([a-f0-9]{16,})", html)
    if cs:
        return cs.group(1)
    return ""


def _univ_is_login_page(url: str, text: str) -> bool:
    """Return True if response looks like still-on-login-page."""
    u = (url or "").lower()
    t = (text or "").lower()[:800]
    if any(w in t for w in ("invalid password", "incorrect password", "wrong password",
                             "login failed", "invalid username", "invalid credentials",
                             "authentication failed", "wrong credentials",
                             "username or password")):
        return True
    # Login form still visible = still on login page (only if very short response)
    if "type=\"password\"" in (text or "").lower() and len(text) < 300:
        return True
    # URL still looks like a login/sign-in page (catches /sign-in with hyphen too)
    if any(w in u for w in ("/login", "/signin", "/signmein", "/sign-in", "/sign_in")):
        if len(text) < 400:
            return True
    return False


def _univ_detect_form_fields(html: str):
    """Auto-detect login form field names from HTML. Returns (user_field, pass_field).
    Only matches visible text/email inputs — skips hidden, radio, checkbox, submit."""
    _SKIP_NAMES = {"password", "_token", "csrf_token", "token", "capt", "captcha",
                   "theme-style", "theme_style", "remember", "remember_me", "submit"}

    # Detect password field name
    pf_m = re.search(
        r'<input[^>]+type=["\']password["\'][^>]*name=["\']([^"\']+)["\']'
        r'|<input[^>]+name=["\']([^"\']+)["\'][^>]*type=["\']password["\']',
        html, re.IGNORECASE,
    )
    pass_field = (pf_m.group(1) or pf_m.group(2)).strip() if pf_m else "password"
    _SKIP_NAMES.add(pass_field)

    # 1st priority: exact well-known names
    for name in ("username", "user", "login", "email", "uname", "usr", "user_name"):
        if re.search(rf'name=["\']({re.escape(name)})["\']', html, re.IGNORECASE):
            return name, pass_field

    # 2nd priority: any text/email input whose name doesn't look like a non-user field
    for m in re.finditer(
        r'<input[^>]+type=["\'](?:text|email)["\'][^>]*name=["\']([^"\']+)["\']'
        r'|<input[^>]+name=["\']([^"\']+)["\'][^>]*type=["\'](?:text|email)["\']',
        html, re.IGNORECASE,
    ):
        candidate = (m.group(1) or m.group(2) or "").strip()
        if candidate.lower() not in _SKIP_NAMES and candidate:
            return candidate, pass_field

    # fallback
    return "username", pass_field


def _universal_login(panel):
    """Login to any SMS panel. Returns (session, token, engine, data_path) or (None,)*4."""
    pid = panel["id"]
    base = panel["base_url"].rstrip("/")
    username = panel["username"]
    password = panel["password"]
    # url_hint: original full URL the user provided (may contain path like /konekta/agent/...)
    url_hint = panel.get("url_hint", "")

    sess = requests.Session()
    sess.headers.update({"User-Agent": _UA})
    sess.verify = False

    # ── Step 1: Find login page ──────────────────────────────────────────────
    # Build a prioritized list: try the hint URL first, then known paths
    login_page_candidates = []
    if url_hint:
        # Try sibling paths of the hint URL (same prefix, different suffix)
        hint_base = _extract_panel_base_url(url_hint) or base
        for lp in ["/login", "/signin", "/signmein", "/"]:
            login_page_candidates.append(hint_base + lp)
    for lp in _LOGIN_PAGE_PATHS:
        login_page_candidates.append(base + lp)
    # Deduplicate while preserving order
    seen_lp = set()
    login_page_candidates = [x for x in login_page_candidates if not (x in seen_lp or seen_lp.add(x))]

    login_html = ""
    login_url_used = ""
    for candidate in login_page_candidates:
        try:
            r = sess.get(candidate, timeout=12, verify=False)
            txt_lo = r.text.lower()
            if r.status_code == 200 and (
                "password" in txt_lo or "username" in txt_lo or "login" in txt_lo
            ):
                login_html = r.text
                login_url_used = candidate
                print(f"[{pid}] Login page found: {candidate}")
                break
        except Exception:
            continue

    if not login_html:
        print(f"[{pid}] ❌ Login page not found at {base}")
        return None, None, None, None

    # ── Step 2: Build post data ──────────────────────────────────────────────
    user_field, pass_field = _univ_detect_form_fields(login_html)
    post_data: dict = {user_field: username, pass_field: password}
    print(f"[{pid}] Form fields: {user_field}={username}, {pass_field}=***")

    # Math captcha
    m_cap = re.search(r"[Ww]hat is (\d+) \+ (\d+)", login_html)
    if m_cap:
        ans = int(m_cap.group(1)) + int(m_cap.group(2))
        post_data["capt"] = ans
        print(f"[{pid}] Math captcha: {m_cap.group(1)}+{m_cap.group(2)}={ans}")

    # Collect ALL hidden fields from the login form (CSRF tokens, session seeds, etc.)
    for hf in re.finditer(
        r'<input[^>]+type=["\']hidden["\'][^>]*name=["\']([^"\']+)["\'][^>]*value=["\']([^"\']*)["\']'
        r'|<input[^>]+name=["\']([^"\']+)["\'][^>]*type=["\']hidden["\'][^>]*value=["\']([^"\']*)["\']',
        login_html, re.IGNORECASE,
    ):
        n = (hf.group(1) or hf.group(3) or "").strip()
        v = (hf.group(2) or hf.group(4) or "").strip()
        if n and n.lower() not in (user_field.lower(), pass_field.lower()):
            post_data[n] = v

    # ── Step 3: Try signin paths ─────────────────────────────────────────────
    # Build candidate signin URLs: derive from login page URL first, then fallbacks
    login_dir = re.sub(r"/[^/]+$", "", login_url_used)  # directory of login page
    signin_candidates = []
    for sp in ["/signin", "/signmein", "/login"]:
        signin_candidates.append(login_dir + sp)   # same directory
    for sp in _SIGNIN_PATHS:
        signin_candidates.append(base + sp)
    # Deduplicate
    seen_sp = set()
    signin_candidates = [x for x in signin_candidates if not (x in seen_sp or seen_sp.add(x))]

    logged_sess = None
    login_resp_text = ""
    for sp_url in signin_candidates:
        try:
            r2 = sess.post(
                sp_url, data=post_data, timeout=12, allow_redirects=True, verify=False,
                headers={"Referer": login_url_used},
            )
            if r2.status_code in (200, 201, 302) and not _univ_is_login_page(r2.url, r2.text):
                logged_sess = sess
                login_resp_text = r2.text
                print(f"[{pid}] ✅ Signed in via {sp_url} → {r2.url}")
                break
        except Exception:
            continue

    if not logged_sess:
        print(f"[{pid}] ❌ Login failed — all signin paths exhausted")
        return None, None, None, None

    # ── Step 3b: Validate session by probing the original URL (soft check) ──────
    # Only used as a hint — never hard-fails a successful login POST
    if url_hint:
        try:
            chk = logged_sess.get(url_hint, timeout=10, verify=False, allow_redirects=True)
            if _univ_is_login_page(chk.url, chk.text):
                print(f"[{pid}] ⚠️ Hint URL looks like login page after signin — proceeding anyway")
            elif len(chk.text) < 50:
                print(f"[{pid}] ⚠️ Hint page very short ({len(chk.text)}b) — proceeding anyway")
            else:
                print(f"[{pid}] ✅ Session validated via {url_hint} ({len(chk.text)}b)")
                login_resp_text = login_resp_text or chk.text
        except Exception as e:
            print(f"[{pid}] ⚠️ Session validation skipped: {e}")

    # ── Step 4: Extract session token ────────────────────────────────────────
    token = _univ_extract_token(login_resp_text)
    if not token:
        # Probe dashboard pages — include hint URL itself
        dash_candidates = []
        if url_hint:
            dash_candidates.append(url_hint)
        hint_base2 = _extract_panel_base_url(url_hint) if url_hint else base
        for dp in _DASHBOARD_PATHS:
            if hint_base2 and hint_base2 != base:
                dash_candidates.append(hint_base2 + dp)
            dash_candidates.append(base + dp)
        for dash_url in dash_candidates:
            try:
                rd = logged_sess.get(dash_url, timeout=10, verify=False)
                token = _univ_extract_token(rd.text)
                if token:
                    print(f"[{pid}] Token found via {dash_url}")
                    break
            except Exception:
                continue
    print(f"[{pid}] Token: {'found (' + token[:8] + '...)' if token else 'empty (cookie-based)'}")

    # ── Step 5: Probe data endpoints ─────────────────────────────────────────
    today = time.strftime("%Y-%m-%d")
    hint_base3 = _extract_panel_base_url(url_hint) if url_hint else None

    # Step 5a: Scrape the dashboard/hint page HTML to extract AJAX data URLs
    # Many panels embed the data URL directly in their JS (ajax: "/path/to/data.php")
    scraped_ep_candidates = []
    pages_to_scrape = []
    if url_hint:
        pages_to_scrape.append(url_hint)
    for dp in _DASHBOARD_PATHS:
        if hint_base3 and hint_base3 != base:
            pages_to_scrape.append(hint_base3 + dp)
        pages_to_scrape.append(base + dp)
    for scrape_url in pages_to_scrape[:8]:  # limit to avoid slow startup
        try:
            rp = logged_sess.get(scrape_url, timeout=10, verify=False)
            if rp.status_code != 200:
                continue
            pg = rp.text
            # Look for ajax/url patterns pointing to data PHP files
            for m in re.finditer(
                r'''["']([^"']*(?:data_sms|dt_reports|dt_sms|cdr|reports)[^"']*\.php)['"''',
                pg, re.IGNORECASE
            ):
                raw = m.group(1)
                # Convert to absolute URL
                if raw.startswith("http"):
                    abs_ep = raw
                elif raw.startswith("/"):
                    parsed_host = re.match(r"(https?://[^/]+)", scrape_url)
                    abs_ep = (parsed_host.group(1) if parsed_host else base) + raw
                else:
                    abs_ep = base + "/" + raw.lstrip("/")
                style = "xisora" if "dt_reports" in raw or "dt_sms" in raw else "ints"
                eng = "xisora" if style == "xisora" else "ints_smscdr"
                scraped_ep_candidates.append((abs_ep, raw, style, eng))
                print(f"[{pid}] 🔎 Scraped data URL from {scrape_url}: {raw}")
        except Exception:
            continue

    # Step 5b: Build known-path candidates (hint_base first, then base)
    known_ep_candidates = []
    for ep_path, style, eng_name in _DATA_ENDPOINTS:
        if hint_base3 and hint_base3 != base:
            known_ep_candidates.append((hint_base3 + ep_path, ep_path, style, eng_name))
        known_ep_candidates.append((base + ep_path, ep_path, style, eng_name))

    # Combine: scraped first (highest confidence), then known paths
    all_ep_candidates = scraped_ep_candidates + known_ep_candidates

    for full_ep, ep_path, style, eng_name in all_ep_candidates:
        try:
            test_url = _univ_build_url(full_ep, token, today, style)
            rr = logged_sess.get(
                test_url, timeout=10, verify=False,
                headers={"X-Requested-With": "XMLHttpRequest"},
            )
            body = rr.text.strip()
            print(f"[{pid}] Probe {full_ep} → HTTP {rr.status_code}, body={len(body)}b, starts={body[:30]!r}")
            if rr.status_code == 200 and body and not body.startswith("<"):
                try:
                    data = json.loads(body)
                    if "aaData" in data:
                        resolved_path = ep_path if full_ep.startswith(base) else ("/" + ep_path.lstrip("/"))
                        print(f"[{pid}] ✅ Data endpoint: {full_ep} (engine={eng_name})")
                        if hint_base3 and hint_base3 != base and full_ep.startswith(hint_base3):
                            panel["base_url"] = hint_base3
                        return logged_sess, token, eng_name, resolved_path
                except Exception:
                    pass
        except Exception as probe_err:
            print(f"[{pid}] Probe error {full_ep}: {probe_err}")
            continue

    # ── Step 5c: HTML table scraping fallback ────────────────────────────────
    # For panels that render data directly in HTML (no AJAX JSON endpoint)
    html_scrape_url = None
    html_pages_to_try = []
    if url_hint:
        html_pages_to_try.append(url_hint)
    for dp in _DASHBOARD_PATHS:
        if hint_base3 and hint_base3 != base:
            html_pages_to_try.append(hint_base3 + dp)
        html_pages_to_try.append(base + dp)
    for pg_url in html_pages_to_try[:6]:
        try:
            rp = logged_sess.get(pg_url, timeout=10, verify=False)
            if rp.status_code != 200 or _univ_is_login_page(rp.url, rp.text):
                continue
            from bs4 import BeautifulSoup
            soup = BeautifulSoup(rp.text, "lxml")
            tables = soup.find_all("table")
            for tbl in tables:
                rows = tbl.find_all("tr")
                if len(rows) >= 3:  # at least header + 2 data rows
                    print(f"[{pid}] 🔎 HTML scraping fallback: found table with {len(rows)} rows at {pg_url}")
                    html_scrape_url = pg_url
                    break
            if html_scrape_url:
                break
        except Exception:
            continue
    if html_scrape_url:
        return logged_sess, token, "html_scrape", html_scrape_url

    # Login succeeded but no endpoint matched → return with INTS default
    print(f"[{pid}] ⚠️ Login OK but no data endpoint found — using default")
    return logged_sess, token, "ints_smscdr", "/agent/res/data_smscdr.php"


def _universal_fetch(panel):
    """Fetch OTPs from any panel using the universal engine."""
    pid = panel["id"]
    base = panel["base_url"].rstrip("/")
    engine = panel.get("engine", "ints_smscdr")
    # IVA SMS has its own fetcher
    if engine == "iva_sms":
        return _iva_fetch(panel)
    # API Key panels have their own fetcher
    if engine == "api_key":
        return _api_key_fetch(panel)
    data_path = panel.get("data_path", "/agent/res/data_smscdr.php")
    style = "xisora" if engine == "xisora" else "ints"
    found = {}

    with _get_dp_lock(pid):
        sd = _dynamic_sessions.get(pid, {})
        if not sd.get("session"):
            sess, tok, det_eng, det_path = _universal_login(panel)
            if not sess:
                _record_error(pid)
                return found
            if det_eng and (det_eng != engine or det_path != data_path):
                panel["engine"] = det_eng
                panel["data_path"] = det_path
                engine = det_eng
                data_path = det_path
                style = "xisora" if engine == "xisora" else "ints"
                save_dynamic_panels()
            _dynamic_sessions[pid] = {"session": sess, "token": tok}
            sd = _dynamic_sessions[pid]

        sess = sd["session"]
        token = sd.get("token", "")
        today = time.strftime("%Y-%m-%d")
        full_ep = base + data_path
        hdrs = {"X-Requested-With": "XMLHttpRequest"}

        def _do_get():
            return sess.get(
                _univ_build_url(full_ep, token, today, style),
                headers=hdrs, timeout=15, verify=False,
            )

        # ── HTML scraping engine ──────────────────────────────────────────────
        if engine == "html_scrape":
            page_url = data_path  # data_path is the full page URL for this engine
            try:
                rp = sess.get(page_url, timeout=15, verify=False)
                if rp.status_code != 200 or _univ_is_login_page(rp.url, rp.text):
                    print(f"[{pid}] Session expired (html_scrape) — re-login")
                    _dynamic_sessions[pid] = {}
                    sess2, tok2, det_eng, det_path = _universal_login(panel)
                    if not sess2:
                        _record_error(pid)
                        return found
                    panel["engine"] = det_eng
                    panel["data_path"] = det_path
                    engine = det_eng
                    data_path = det_path
                    save_dynamic_panels()
                    _dynamic_sessions[pid] = {"session": sess2, "token": tok2}
                    rp = sess2.get(det_path if det_eng == "html_scrape" else page_url,
                                   timeout=15, verify=False)
                from bs4 import BeautifulSoup
                soup = BeautifulSoup(rp.text, "lxml")
                # Find header row to map column positions
                tbl = None
                for t in soup.find_all("table"):
                    if len(t.find_all("tr")) >= 3:
                        tbl = t
                        break
                if not tbl:
                    _record_fetch(pid, 0)
                    return found
                headers = [th.get_text(strip=True).lower()
                           for th in (tbl.find("tr").find_all(["th", "td"]))]
                # Heuristic column detection
                num_col = next((i for i, h in enumerate(headers)
                                if any(w in h for w in ("number", "msisdn", "phone", "mobile", "num"))), None)
                txt_col = next((i for i, h in enumerate(headers)
                                if any(w in h for w in ("message", "sms", "text", "body", "msg"))), None)
                svc_col = next((i for i, h in enumerate(headers)
                                if any(w in h for w in ("service", "route", "dest", "sender"))), None)
                row_count = 0
                for tr in tbl.find_all("tr")[1:]:  # skip header
                    cells = tr.find_all(["td", "th"])
                    if not cells:
                        continue
                    row_count += 1
                    number = cells[num_col].get_text(strip=True) if num_col is not None and num_col < len(cells) else ""
                    sms_txt = cells[txt_col].get_text(strip=True) if txt_col is not None and txt_col < len(cells) else ""
                    service = cells[svc_col].get_text(strip=True) if svc_col is not None and svc_col < len(cells) else ""
                    # Fallback: scan all cells for phone-like pattern and OTP
                    if not number:
                        for c in cells:
                            ct = c.get_text(strip=True)
                            if re.match(r"^\+?\d{7,15}$", ct):
                                number = ct
                                break
                    if not sms_txt:
                        for c in cells:
                            ct = c.get_text(strip=True)
                            if len(ct) > 10 and re.search(r"\d{4,8}", ct):
                                sms_txt = ct
                                break
                    otp = extract_otp_from_sms(sms_txt)
                    if number and otp:
                        key = f"{number}:{sms_txt}"
                        found[key] = (number, otp, sms_txt, service)
                _record_fetch(pid, row_count)
                if found:
                    print(f"[{pid}] ✅ HTML-scraped {row_count} rows, {len(found)} OTPs")
            except Exception as e:
                print(f"[{pid}] HTML scrape error: {e}")
                _record_error(pid)
                _dynamic_sessions[pid] = {}
            return found

        # ── JSON / AJAX engine (default) ──────────────────────────────────────
        try:
            r = _do_get()
            body = r.text.strip()
            if r.status_code != 200 or not body or body.startswith("<") or "Direct Script" in body:
                print(f"[{pid}] Session expired — re-login")
                _dynamic_sessions[pid] = {}
                sess2, tok2, det_eng, det_path = _universal_login(panel)
                if not sess2:
                    _record_error(pid)
                    return found
                if det_eng:
                    panel["engine"] = det_eng
                    panel["data_path"] = det_path
                    engine = det_eng
                    data_path = det_path
                    style = "xisora" if engine == "xisora" else "ints"
                    full_ep = base + data_path
                    save_dynamic_panels()
                _dynamic_sessions[pid] = {"session": sess2, "token": tok2}
                sd = _dynamic_sessions[pid]
                sess = sess2
                token = tok2
                r = _do_get()
                body = r.text.strip()

            rows = json.loads(body).get("aaData", [])
            for row in rows:
                parsed = _univ_parse_row(row, engine)
                if not parsed:
                    continue
                number, service, sms_txt = parsed
                otp = extract_otp_from_sms(sms_txt)
                if otp:
                    key = f"{number}:{sms_txt}"
                    found[key] = (number, otp, sms_txt, service)
            _record_fetch(pid, len(rows))
            if found:
                print(f"[{pid}] ✅ Fetched {len(rows)} rows, {len(found)} OTPs")
        except Exception as e:
            print(f"[{pid}] Fetch error: {e}")
            _record_error(pid)
            _dynamic_sessions[pid] = {}
    return found


def _univ_parse_row(row, engine):
    """Parse one aaData row. Returns (number, service, sms_text) or None."""
    try:
        if not row:
            return None
        cells = [str(c).strip() for c in row]
        # Standard INTS layout: [date, range, number, cli/service, client, sms, ...]
        # extract_otp_from_sms already enforces real-SMS validation (≥4 alpha chars)
        if len(cells) > 5:
            number  = cells[2]
            service = cells[3]
            sms_txt = cells[5]
            if number and extract_otp_from_sms(sms_txt):
                return number, service, sms_txt
        # Shorter rows (some panels only have 5 cols)
        if len(cells) > 4:
            number  = cells[2]
            service = cells[3]
            sms_txt = cells[4]
            if number and extract_otp_from_sms(sms_txt):
                return number, service, sms_txt
    except Exception:
        pass
    return None


# ── API Key Panel Engine ───────────────────────────────────────────────────────

_API_KEY_ENDPOINTS = [
    "/api/sms",
    "/api/messages",
    "/api/received",
    "/api/sms/received",
    "/api/v1/sms",
    "/api/v1/messages",
    "/api/inbox",
    "/api/sms/list",
    "/api/data",
    "/sms",
    "/api/otp",
]

_API_KEY_PARAMS = ["api_key", "token", "key", "apikey", "access_token"]


def _api_key_test(base_url, api_key):
    """Try common API endpoints with key. Returns (path, param_style) or (None, None).
    param_style: 'bearer', 'api_key=xxx', 'token=xxx', etc.
    """
    base = base_url.rstrip("/")
    hdrs = {"Accept": "application/json", "User-Agent": _UA}
    for path in _API_KEY_ENDPOINTS:
        # Try query param variants
        for param in _API_KEY_PARAMS:
            try:
                r = requests.get(
                    f"{base}{path}?{param}={api_key}",
                    timeout=8, verify=False, headers=hdrs,
                )
                if r.status_code == 200:
                    try:
                        data = r.json()
                        if isinstance(data, list):
                            print(f"[APIKEY-TEST] ✅ {path}?{param}= → array")
                            return path, param
                        if isinstance(data, dict):
                            for dk in ("data", "sms", "messages", "records", "result", "items", "list"):
                                if dk in data and isinstance(data[dk], list):
                                    print(f"[APIKEY-TEST] ✅ {path}?{param}= → dict.{dk}")
                                    return path, f"{param}|resp={dk}"
                            if any(k in data for k in ("status", "success", "ok", "error")):
                                print(f"[APIKEY-TEST] ✅ {path}?{param}= → status obj")
                                return path, param
                    except Exception:
                        pass
            except Exception:
                continue
        # Try Authorization: Bearer header
        try:
            r = requests.get(
                f"{base}{path}", timeout=8, verify=False,
                headers={**hdrs, "Authorization": f"Bearer {api_key}"},
            )
            if r.status_code == 200:
                try:
                    data = r.json()
                    if isinstance(data, (list, dict)):
                        print(f"[APIKEY-TEST] ✅ {path} → Bearer header")
                        return path, "bearer"
                except Exception:
                    pass
        except Exception:
            continue
    return None, None


def _api_key_fetch(panel):
    """Fetch OTPs from an API-key authenticated panel."""
    pid      = panel["id"]
    base     = panel["base_url"].rstrip("/")
    api_key  = panel.get("api_key", "")
    data_path = panel.get("data_path", "/api/sms")
    param_style = panel.get("api_key_param", "api_key")
    found = {}
    try:
        hdrs = {"Accept": "application/json", "User-Agent": _UA}
        if param_style == "bearer":
            url = f"{base}{data_path}"
            hdrs["Authorization"] = f"Bearer {api_key}"
        elif "|resp=" in param_style:
            pname = param_style.split("|resp=")[0]
            url = f"{base}{data_path}?{pname}={api_key}"
        else:
            url = f"{base}{data_path}?{param_style}={api_key}"

        r = requests.get(url, timeout=15, verify=False, headers=hdrs)
        if r.status_code != 200:
            print(f"[{pid}] API key fetch HTTP {r.status_code}")
            _record_error(pid)
            return found

        raw = r.json()
        rows = []
        if isinstance(raw, list):
            rows = raw
        elif isinstance(raw, dict):
            resp_key = param_style.split("|resp=")[1] if "|resp=" in param_style else None
            if resp_key and resp_key in raw:
                rows = raw[resp_key]
            else:
                for dk in ("data", "sms", "messages", "records", "result", "items", "list"):
                    if dk in raw and isinstance(raw[dk], list):
                        rows = raw[dk]
                        break

        for row in rows:
            number = sms_txt = service = ""
            if isinstance(row, dict):
                number  = str(row.get("number") or row.get("phone") or row.get("msisdn") or
                              row.get("from") or row.get("sender") or row.get("to") or "").strip()
                sms_txt = str(row.get("message") or row.get("sms") or row.get("text") or
                              row.get("body") or row.get("content") or "").strip()
                service = str(row.get("service") or row.get("sender") or row.get("source") or "").strip()
            elif isinstance(row, list):
                if len(row) > 2: number  = str(row[2]).strip()
                if len(row) > 5: sms_txt = str(row[5]).strip()
                elif len(row) > 3: sms_txt = str(row[3]).strip()
                if len(row) > 3: service = str(row[3]).strip()
            otp = extract_otp_from_sms(sms_txt)
            if number and otp:
                key = f"{number}:{sms_txt}"
                found[key] = (number, otp, sms_txt, service)

        _record_fetch(pid, len(rows))
        if found:
            print(f"[{pid}] ✅ API key: {len(rows)} rows, {len(found)} OTPs")
    except Exception as e:
        print(f"[{pid}] API key fetch error: {e}")
        _record_error(pid)
    return found


# ─────────────────────────────────────────────────────────────────────────────

def _start_dynamic_panel(panel):
    pid = panel["id"]
    with _stats_lock:
        _panel_stats[pid] = {
            "name": panel.get("username", pid),
            "host": panel.get("host", ""),
            "status": "⏳",
            "count": 0,
            "last": None,
            "errors": 0,
        }

    def monitor():
        global seen_otps
        print(f"[{pid}-MONITOR] Started. Pre-loading existing records...")
        existing = _universal_fetch(panel)
        with seen_lock:
            for key in existing:
                seen_otps[key] = True
            save_json(SEEN_FILE, seen_otps)
        print(f"[{pid}-MONITOR] Pre-loaded {len(existing)} records. Watching for new ones...")
        while True:
            try:
                process_new_otps(_universal_fetch(panel))
            except Exception as e:
                print(f"[{pid}-MONITOR] Loop error: {e}")
            time.sleep(POLL_INTERVAL)

    threading.Thread(target=monitor, daemon=True).start()


# ─── IVA SMS (ivasms.com) engine ─────────────────────────────────────────────

_iva_scrapers: dict = {}
_iva_lock_map: dict = {}


def _iva_get_lock(pid):
    if pid not in _iva_lock_map:
        _iva_lock_map[pid] = threading.Lock()
    return _iva_lock_map[pid]


_IVA_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
}


def _iva_make_session(cookie_str=""):
    """Create a plain requests.Session with browser headers + optional cookies."""
    sess = requests.Session()
    sess.headers.update(_IVA_HEADERS)
    if cookie_str:
        for part in cookie_str.split(";"):
            part = part.strip()
            if "=" in part:
                k, v = part.split("=", 1)
                sess.cookies.set(k.strip(), v.strip(), domain="ivasms.com")
    return sess


# Keep old name for compatibility
def _iva_make_scraper():
    return _iva_make_session()


def _iva_set_cookies(sess, cookie_str):
    for part in cookie_str.split(";"):
        part = part.strip()
        if "=" in part:
            k, v = part.split("=", 1)
            sess.cookies.set(k.strip(), v.strip(), domain="ivasms.com")


def _iva_login(panel):
    """Login to ivasms.com using browser cookies (plain requests — no cloudscraper).
    Railway/cloud server IPs are blocked by Cloudflare for email/password login,
    so we ONLY support cookie-based login here.
    """
    pid     = panel["id"]
    cookies = panel.get("cookie_str", "")
    base    = "https://ivasms.com"

    if not cookies:
        print(f"[{pid}] ❌ IVA SMS: No cookie provided — cannot login from server IP")
        return False

    sess = _iva_make_session(cookies)
    try:
        r = sess.get(
            f"{base}/portal/sms/received",
            timeout=20,
            allow_redirects=True,
        )
        if r.status_code == 200 and "login" not in r.url.lower():
            _iva_scrapers[pid] = sess
            print(f"[{pid}] ✅ IVA SMS: Cookie login OK ({r.url})")
            return True
        print(f"[{pid}] ❌ IVA SMS: Cookie login failed — status={r.status_code} url={r.url}")
        return False
    except Exception as e:
        print(f"[{pid}] ❌ IVA SMS cookie login error: {e}")
        return False


def _iva_parse_page(html):
    """Parse ivasms.com SMS received page → {key: (number, otp, sms_txt, service)}."""
    found = {}

    # 1) Embedded JS array (e.g. var data = [...])
    for pat in [
        r'(?:data|rows|messages|smsList|records|smsData)\s*[:=]\s*(\[.*?\])\s*[,;]',
        r'\.DataTable\([^)]*data\s*:\s*(\[.*?\])',
    ]:
        m = re.search(pat, html, re.DOTALL | re.I)
        if m:
            try:
                records = json.loads(m.group(1))
                for rec in records:
                    if not isinstance(rec, dict):
                        continue
                    num = str(rec.get("number", rec.get("phone", rec.get("msisdn", ""))))
                    txt = str(rec.get("message", rec.get("sms", rec.get("text", rec.get("body", "")))))
                    svc = str(rec.get("service", rec.get("cli", rec.get("sender", "IVA"))))
                    otp = extract_otp_from_sms(txt)
                    if num and otp:
                        found[f"{num}:{txt}"] = (num, otp, txt, svc)
                if found:
                    return found
            except Exception:
                pass

    # 2) HTML table: look for rows with phone number + SMS content
    from html.parser import HTMLParser
    class _TblParser(HTMLParser):
        def __init__(self):
            super().__init__()
            self.rows, self._cur_row, self._cur_cell, self._in_td = [], [], [], False
        def handle_starttag(self, tag, attrs):
            if tag == "tr":   self._cur_row = []
            elif tag == "td": self._in_td = True; self._cur_cell = []
        def handle_endtag(self, tag):
            if tag == "td":
                self._cur_row.append("".join(self._cur_cell).strip())
                self._in_td = False
            elif tag == "tr" and self._cur_row:
                self.rows.append(self._cur_row)
        def handle_data(self, data):
            if self._in_td: self._cur_cell.append(data)

    parser = _TblParser()
    try:
        parser.feed(html)
    except Exception:
        pass

    phone_re = re.compile(r"^\+?\d{7,15}$")
    for row in parser.rows:
        clean = [re.sub(r"\s+", " ", c).strip() for c in row]
        nums = [c for c in clean if phone_re.match(c)]
        smses = [c for c in clean if extract_otp_from_sms(c)]
        if nums and smses:
            n, t = nums[0], smses[0]
            found[f"{n}:{t}"] = (n, extract_otp_from_sms(t), t, "IVA")

    return found


def _iva_parse_rows(rows, default_svc="IVA"):
    """Convert a list of dicts (DataTables / JSON) to found-dict entries."""
    found = {}
    for rec in rows:
        if not isinstance(rec, dict):
            # DataTables may return list-of-lists
            continue
        num = str(rec.get("number", rec.get("phone", rec.get("msisdn",
              rec.get("sender", rec.get("from", ""))))))
        txt = str(rec.get("message", rec.get("sms", rec.get("text",
              rec.get("body", rec.get("content", ""))))))
        svc = str(rec.get("service", rec.get("cli", rec.get("application",
              rec.get("app", rec.get("shortcode", default_svc))))))
        otp = extract_otp_from_sms(txt)
        if num and otp:
            found[f"{num}:{txt}"] = (num, otp, txt, svc)
    return found


def _iva_dt_post(scraper, url, csrf_token, page_html="", start=0, length=100):
    """Send a DataTables server-side POST and return parsed rows."""
    today = time.strftime("%Y-%m-%d")
    payload = {
        "draw": "1",
        "start": str(start),
        "length": str(length),
        "search[value]": "",
        "search[regex]": "false",
        "_token": csrf_token,
        "start_date": today,
        "end_date": today,
        # Common DataTables column ordering params (harmless if unused)
        "order[0][column]": "0",
        "order[0][dir]": "desc",
    }
    hdrs = {
        "X-Requested-With": "XMLHttpRequest",
        "Accept": "application/json, text/javascript, */*",
        "Referer": "https://ivasms.com/portal/sms/received",
        "Content-Type": "application/x-www-form-urlencoded",
    }
    r = scraper.post(url, data=payload, headers=hdrs, timeout=25)
    if r.status_code != 200:
        return None, r.status_code
    ct = r.headers.get("Content-Type", "")
    if "json" not in ct:
        return None, 0
    try:
        return r.json(), 200
    except Exception:
        return None, 0


def _iva_fetch(panel):
    """Fetch latest OTPs from ivasms.com.

    Strategy (in order):
      1. Load /portal/sms/received, grab CSRF token + discover AJAX URL
      2. Try DataTables POST to the discovered/candidate AJAX endpoints
      3. Fall back to plain HTML table parse of the page
    """
    pid  = panel["id"]
    base = "https://ivasms.com"
    found = {}

    with _iva_get_lock(pid):
        scraper = _iva_scrapers.get(pid)
        if not scraper:
            if not _iva_login(panel):
                _record_error(pid)
                return found
            scraper = _iva_scrapers[pid]

        today = time.strftime("%Y-%m-%d")

        # ── Step 1: load the page ─────────────────────────────────────────────
        try:
            r = scraper.get(f"{base}/portal/sms/received",
                            params={"start_date": today, "end_date": today},
                            timeout=30)
        except Exception as e:
            print(f"[{pid}] IVA page load error: {e}")
            _iva_scrapers.pop(pid, None)
            _record_error(pid)
            return found

        # Session expired?
        if r.status_code != 200 or "login" in r.url.lower():
            print(f"[{pid}] IVA session expired → re-login")
            _iva_scrapers.pop(pid, None)
            if not _iva_login(panel):
                _record_error(pid)
                return found
            scraper = _iva_scrapers[pid]
            try:
                r = scraper.get(f"{base}/portal/sms/received",
                                params={"start_date": today, "end_date": today},
                                timeout=30)
            except Exception as e:
                print(f"[{pid}] IVA page reload error: {e}")
                _record_error(pid)
                return found

        html = r.text

        # ── Step 2: extract CSRF token from page ──────────────────────────────
        csrf = ""
        for pat in [
            r'<meta[^>]+name=["\']csrf-token["\'][^>]*content=["\']([^"\']+)["\']',
            r'"_token"\s*:\s*"([^"]+)"',
            r"_token[\"']\s*:\s*[\"']([^\"']+)[\"']",
            r'<input[^>]+name=["\']_token["\'][^>]*value=["\']([^"\']+)["\']',
        ]:
            m = re.search(pat, html, re.I)
            if m:
                csrf = m.group(1)
                break

        # ── Step 3: discover DataTables AJAX URL from page JS ─────────────────
        # Looks for patterns like: ajax: '/portal/sms/received/data'  or  url: '...'
        ajax_url_candidates = []
        for pat in [
            r"""ajax\s*:\s*['"](\/portal\/sms\/[^'"]+)['"]""",
            r"""url\s*:\s*['"](\/portal\/sms\/[^'"]+)['"]""",
            r"""action\s*=\s*['"](\/portal\/sms\/[^'"]+)['"]""",
            r"""fetch\(['"](\/portal\/sms\/[^'"?]+)""",
        ]:
            for m in re.finditer(pat, html, re.I):
                ep = m.group(1)
                if ep not in ajax_url_candidates:
                    ajax_url_candidates.append(ep)

        # Add hardcoded fallback candidates (most common Laravel SMS panel patterns)
        for fallback in [
            "/portal/sms/received/data",
            "/portal/sms/received-data",
            "/portal/received/sms/data",
            "/portal/sms/datatable",
            "/portal/sms/ajax",
            "/portal/sms/list",
            "/portal/api/sms/received",
            "/portal/sms/received",          # POST to same URL (common)
        ]:
            if fallback not in ajax_url_candidates:
                ajax_url_candidates.append(fallback)

        # ── Step 4: try DataTables POST on each candidate ─────────────────────
        for ep in ajax_url_candidates:
            ep_url = base + ep if ep.startswith("/") else ep
            try:
                js, code = _iva_dt_post(scraper, ep_url, csrf, html)
                if js is None:
                    continue
                # DataTables standard: {"draw":N, "data":[...], "recordsTotal":N}
                rows = (js if isinstance(js, list)
                        else js.get("data", js.get("records",
                             js.get("rows", js.get("sms", [])))))
                if isinstance(rows, list) and rows:
                    # Rows may be list-of-dicts or list-of-lists
                    if isinstance(rows[0], dict):
                        found = _iva_parse_rows(rows)
                    else:
                        # list-of-lists: try to figure out columns from header
                        # Best-effort: assume [id, number, message, service, ...]
                        for row in rows:
                            if len(row) >= 3:
                                num = str(row[1]) if len(row) > 1 else ""
                                txt = str(row[2]) if len(row) > 2 else ""
                                svc = str(row[3]) if len(row) > 3 else "IVA"
                                otp = extract_otp_from_sms(txt)
                                if num and otp:
                                    found[f"{num}:{txt}"] = (num, otp, txt, svc)
                    if found:
                        print(f"[{pid}] ✅ IVA DataTables hit: {ep_url} → {len(found)} OTPs")
                        # Remember this working endpoint for next time
                        if panel.get("iva_ajax_url") != ep_url:
                            panel["iva_ajax_url"] = ep_url
                            save_dynamic_panels()
                        break
            except Exception as ex:
                print(f"[{pid}] IVA AJAX probe {ep}: {ex}")
                continue

        # ── Step 5: fall back to static HTML table parse ─────────────────────
        if not found:
            found = _iva_parse_page(html)
            if found:
                print(f"[{pid}] ✅ IVA HTML-table parse → {len(found)} OTPs")

        _record_fetch(pid, len(found))
        if found:
            print(f"[{pid}] ✅ IVA SMS total: {len(found)} OTPs found")
        else:
            # Debug: log first 300 chars of page so we know what we're getting
            preview = html[:300].replace("\n", " ").strip()
            print(f"[{pid}] IVA SMS: 0 OTPs. Page preview: {preview}")

    return found


# ─────────────────────────────────────────────────────────────────────────────


def extract_otp_from_sms(sms_text):
    if not sms_text:
        return None
    sms_text = str(sms_text).strip()
    if len(sms_text) < 4:
        return None
    # Pure short numeric codes (4-8 digits only) — treat directly as OTP
    # WhatsApp/Telegram panels sometimes store just the raw code
    if re.match(r"^\d{4,8}$", sms_text):
        return sms_text
    # WhatsApp format: "123-456" or "123 456" (digits separated by dash/space)
    wa_m = re.match(r"^(\d{3})[- ](\d{3})$", sms_text)
    if wa_m:
        return wa_m.group(1) + wa_m.group(2)
    # Must have at least 1 letter for longer strings — pure long digit strings
    # (e.g. phone numbers "88017XXXXXXX") are not OTPs
    if not re.search(r"[a-zA-Z]", sms_text):
        return None
    cleaned = re.sub(r"(?<=\d) (?=\d)", "", sms_text)
    cleaned = re.sub(r"(\d)-(\d)", r"\1\2", cleaned)
    cleaned = re.sub(r"(\d)\.(\d)", r"\1\2", cleaned)
    m = re.search(r"\b(\d{4,8})\b", cleaned)
    return m.group(1) if m else None


# ── Panel 1 login & fetch (Mahofuza) ─────────────────────────────────────────


def p1_login():
    global _p1_session, _p1_sesskey
    sess = requests.Session()
    sess.headers.update({"User-Agent": "Mozilla/5.0"})
    try:
        r = sess.get(P1_LOGIN_PAGE, timeout=15)
        m = re.search(r"What is (\d+) \+ (\d+)", r.text)
        if not m:
            print("[P1] Could not find captcha")
            return False
        answer = int(m.group(1)) + int(m.group(2))
        r2 = sess.post(
            P1_SIGNIN_URL,
            data={"username": P1_USER_NAME, "password": P1_PASSWORD, "capt": answer},
            timeout=15,
            allow_redirects=True,
        )
        if "login" in r2.url.lower() or "login" in r2.text.lower()[:500]:
            print("[P1] Login failed — still on login page")
            return False
        r3 = sess.get(
            P1_CDR_PAGE, timeout=15, headers={"Referer": P1_BASE_URL + "/agent/"}
        )
        sk = re.search(r"sesskey=([A-Za-z0-9+/=]+)", r3.text)
        _p1_sesskey = sk.group(1) if sk else ""
        _p1_session = sess
        print(f"[P1] Logged in. sesskey={_p1_sesskey}")
        return True
    except Exception as e:
        print(f"[P1] Login error: {e}")
        return False


def fetch_panel1():
    global _p1_session, _p1_sesskey
    found = {}
    with _p1_lock:
        try:
            today = time.strftime("%Y-%m-%d")

            def build_url():
                return (
                    f"{P1_CDR_DATA_URL}"
                    f"?fdate1={today}%2000:00:00"
                    f"&fdate2={today}%2023:59:59"
                    f"&frange=&fclient=&fnum=&fcli=&fgdate=&fgmonth="
                    f"&fgrange=&fgclient=&fgnumber=&fgcli=&fg=0"
                    f"&sesskey={_p1_sesskey or ''}"
                )

            headers = {"Referer": P1_CDR_PAGE, "X-Requested-With": "XMLHttpRequest"}
            if _p1_session is None:
                if not p1_login():
                    return found
            r = _p1_session.get(build_url(), headers=headers, timeout=15)
            body = r.text.strip()
            if (
                r.status_code != 200
                or not body
                or body.startswith("<")
                or "Direct Script" in body
            ):
                print(f"[P1] Bad response ({r.status_code}), re-logging in.")
                _p1_session = None
                if not p1_login():
                    return found
                r = _p1_session.get(build_url(), headers=headers, timeout=15)
                body = r.text.strip()
            rows = json.loads(body).get("aaData", [])
            for row in rows:
                number = str(row[2]).strip()
                service = str(row[3]).strip()
                sms_txt = str(row[5]).strip()
                otp = extract_otp_from_sms(sms_txt)
                if otp:
                    key = f"{number}:{sms_txt}"
                    found[key] = (number, otp, sms_txt, service)
            _record_fetch("p1", len(rows))
            if found:
                print(f"[P1] ✅ Fetched {len(found)} records.")
        except Exception as e:
            print(f"[P1] Fetch error: {e}")
            _record_error("p1")
            _p1_session = None
    return found


# ── Panel 2 login & fetch (Sagardas50 / XISORA) ──────────────────────────────


def p2_login():
    global _p2_session
    sess = requests.Session()
    sess.headers.update({"User-Agent": "Mozilla/5.0"})
    try:
        r = sess.post(
            P2_SIGNIN_URL,
            data={"username": P2_USER_NAME, "password": P2_PASSWORD},
            timeout=15,
            allow_redirects=True,
        )
        if "signin" in r.url.lower() or "login" in r.url.lower():
            print("[P2] Login failed — still on login page")
            return False
        _p2_session = sess
        print(f"[P2] Logged in. URL={r.url}")
        return True
    except Exception as e:
        print(f"[P2] Login error: {e}")
        return False


def fetch_panel2():
    global _p2_session
    found = {}
    with _p2_lock:
        try:
            today = time.strftime("%Y-%m-%d")
            url = (
                f"{P2_DATA_URL}"
                f"?fdate1={today}%2000:00:00"
                f"&fdate2={today}%2023:59:59"
                f"&ftermination=&fclient=&fnum=&fcli="
                f"&fgdate=0&fgtermination=0&fgclient=0&fgnumber=0&fgcli=0&fg=0"
            )
            headers = {"Referer": P2_REPORTS_PAGE, "X-Requested-With": "XMLHttpRequest"}
            if _p2_session is None:
                if not p2_login():
                    return found
            r = _p2_session.get(url, headers=headers, timeout=15)
            body = r.text.strip()
            if r.status_code != 200 or not body or body.startswith("<"):
                print(f"[P2] Bad response ({r.status_code}), re-logging in.")
                _p2_session = None
                if not p2_login():
                    return found
                r = _p2_session.get(url, headers=headers, timeout=15)
                body = r.text.strip()
            rows = json.loads(body).get("aaData", [])
            for row in rows:
                if not isinstance(row[0], str):
                    continue
                number = str(row[2]).strip()
                service = str(row[3]).strip()
                sms_txt = str(row[10]).strip()
                otp = extract_otp_from_sms(sms_txt)
                if otp:
                    key = f"{number}:{sms_txt}"
                    found[key] = (number, otp, sms_txt, service)
            _record_fetch("p2", len(rows))
            if found:
                print(f"[P2] ✅ Fetched {len(found)} records.")
        except Exception as e:
            print(f"[P2] Fetch error: {e}")
            _record_error("p2")
            _p2_session = None
    return found


# ── Shared OTP processor ──────────────────────────────────────────────────────


def process_new_otps(current):
    global seen_otps
    for key, (number, otp, sms_txt, service) in current.items():
        with seen_lock:
            if key in seen_otps:
                continue
            seen_otps[key] = True
            save_json(SEEN_FILE, seen_otps)
        clean = re.sub(r"\D", "", str(number))
        with user_map_lock:
            t_start = assigned_time.get(clean)
        seconds = int(time.time() - t_start) if t_start else 0
        _dispatch_otp(otp, number, seconds, service, sms=sms_txt)
        print(
            f"[MONITOR] ✅ Forwarded OTP={otp} for {number} ({service}) in {seconds}s"
        )


# ── Global OTP monitors ───────────────────────────────────────────────────────


def panel1_monitor():
    global seen_otps
    print("[P1-MONITOR] Started. Pre-loading existing records...")
    existing = fetch_panel1()
    with seen_lock:
        for key in existing:
            seen_otps[key] = True
        save_json(SEEN_FILE, seen_otps)
    print(f"[P1-MONITOR] Pre-loaded {len(existing)} records. Watching for new ones...")
    while True:
        try:
            process_new_otps(fetch_panel1())
        except Exception as e:
            print(f"[P1-MONITOR] Loop error: {e}")
        time.sleep(POLL_INTERVAL)


def panel2_monitor():
    global seen_otps
    print("[P2-MONITOR] Started. Pre-loading existing records...")
    existing = fetch_panel2()
    with seen_lock:
        for key in existing:
            seen_otps[key] = True
        save_json(SEEN_FILE, seen_otps)
    print(f"[P2-MONITOR] Pre-loaded {len(existing)} records. Watching for new ones...")
    while True:
        try:
            process_new_otps(fetch_panel2())
        except Exception as e:
            print(f"[P2-MONITOR] Loop error: {e}")
        time.sleep(POLL_INTERVAL)


# ── Panel 3 login & fetch (Rabbi1_FD) ────────────────────────────────────────


def p3_login():
    global _p3_session, _p3_csstr
    sess = requests.Session()
    sess.headers.update({"User-Agent": "Mozilla/5.0"})
    try:
        r = sess.get(P3_LOGIN_PAGE, timeout=15)
        m = re.search(r"What is (\d+) \+ (\d+)", r.text)
        if not m:
            print("[P3] Could not find captcha")
            return False
        answer = int(m.group(1)) + int(m.group(2))
        r2 = sess.post(
            P3_SIGNIN_URL,
            data={"username": P3_USER_NAME, "password": P3_PASSWORD, "capt": answer},
            timeout=15,
            allow_redirects=True,
        )
        if "login" in r2.url.lower() or "signin" in r2.url.lower():
            print("[P3] Login failed — still on login page")
            return False
        r3 = sess.get(
            P3_CDR_PAGE, timeout=15, headers={"Referer": P3_BASE_URL + "/agent/"}
        )
        cs = re.search(r"csstr=([a-f0-9]+)", r3.text)
        _p3_csstr = cs.group(1) if cs else ""
        _p3_session = sess
        print(f"[P3] Logged in. csstr={_p3_csstr}")
        return True
    except Exception as e:
        print(f"[P3] Login error: {e}")
        return False


def fetch_panel3():
    global _p3_session, _p3_csstr
    found = {}
    with _p3_lock:
        try:
            today = time.strftime("%Y-%m-%d")

            def build_url():
                return (
                    f"{P3_CDR_DATA_URL}"
                    f"?fdate1={today}%2000:00:00"
                    f"&fdate2={today}%2023:59:59"
                    f"&frange=&fclient=&fnum=&fcli=&fgdate=&fgmonth="
                    f"&fgrange=&fgclient=&fgnumber=&fgcli=&fg=0"
                    f"&csstr={_p3_csstr or ''}"
                )

            headers = {"Referer": P3_CDR_PAGE, "X-Requested-With": "XMLHttpRequest"}
            if _p3_session is None:
                if not p3_login():
                    return found
            r = _p3_session.get(build_url(), headers=headers, timeout=15)
            body = r.text.strip()
            if (
                r.status_code != 200
                or not body
                or body.startswith("<")
                or "Direct Script" in body
            ):
                print(f"[P3] Bad response ({r.status_code}), re-logging in.")
                _p3_session = None
                if not p3_login():
                    return found
                r = _p3_session.get(build_url(), headers=headers, timeout=15)
                body = r.text.strip()
            rows = json.loads(body).get("aaData", [])
            for row in rows:
                if not isinstance(row[0], str):
                    continue
                number = str(row[2]).strip()
                service = str(row[3]).strip()
                sms_txt = str(row[5]).strip()
                otp = extract_otp_from_sms(sms_txt)
                if otp:
                    key = f"{number}:{sms_txt}"
                    found[key] = (number, otp, sms_txt, service)
            _record_fetch("p3", len(rows))
            if found:
                print(f"[P3] ✅ Fetched {len(found)} records.")
        except Exception as e:
            print(f"[P3] Fetch error: {e}")
            _record_error("p3")
            _p3_session = None
    return found


def panel3_monitor():
    global seen_otps
    print("[P3-MONITOR] Started. Pre-loading existing records...")
    existing = fetch_panel3()
    with seen_lock:
        for key in existing:
            seen_otps[key] = True
        save_json(SEEN_FILE, seen_otps)
    print(f"[P3-MONITOR] Pre-loaded {len(existing)} records. Watching for new ones...")
    while True:
        try:
            process_new_otps(fetch_panel3())
        except Exception as e:
            print(f"[P3-MONITOR] Loop error: {e}")
        time.sleep(POLL_INTERVAL)


# ── Panel 4 login & fetch (Rabbi12 / 144.217.71.192) ─────────────────────────


def p4_login():
    global _p4_session, _p4_sesskey
    sess = requests.Session()
    sess.headers.update({"User-Agent": "Mozilla/5.0"})
    try:
        r = sess.get(P4_LOGIN_PAGE, timeout=15)
        m = re.search(r"What is (\d+) \+ (\d+)", r.text)
        if not m:
            print("[P4] Could not find captcha")
            return False
        answer = int(m.group(1)) + int(m.group(2))
        r2 = sess.post(
            P4_SIGNIN_URL,
            data={"username": P4_USER_NAME, "password": P4_PASSWORD, "capt": answer},
            timeout=15,
            allow_redirects=True,
        )
        if "SMSDashboard" not in r2.url and "agent" not in r2.url:
            print(f"[P4] Login failed: {r2.url}")
            return False
        r3 = sess.get(
            P4_CDR_PAGE, timeout=15, headers={"Referer": P4_BASE_URL + "/agent/"}
        )
        sk = re.search(r"sesskey=([A-Za-z0-9+/=]+)", r3.text)
        _p4_sesskey = sk.group(1) if sk else ""
        _p4_session = sess
        print(f"[P4] Logged in. sesskey={_p4_sesskey}")
        return True
    except Exception as e:
        print(f"[P4] Login error: {e}")
        return False


def fetch_panel4():
    global _p4_session, _p4_sesskey
    found = {}
    with _p4_lock:
        if not _p4_session and not p4_login():
            return found
        today = time.strftime("%Y-%m-%d")

        def build_url():
            return (
                f"{P4_CDR_DATA_URL}"
                f"?fdate1={today}%2000:00:00&fdate2={today}%2023:59:59"
                f"&frange=&fclient=&fnum=&fcli=&fgdate=&fgmonth="
                f"&fgrange=&fgclient=&fgnumber=&fgcli=&fg=0"
                f"&sesskey={_p4_sesskey}"
            )

        headers = {"Referer": P4_CDR_PAGE, "X-Requested-With": "XMLHttpRequest"}
        try:
            r = _p4_session.get(build_url(), headers=headers, timeout=15)
            body = r.text.strip()
            if (
                r.status_code != 200
                or not body
                or body.startswith("<")
                or "Direct Script" in body
            ):
                print(f"[P4] Bad response ({r.status_code}), re-logging in.")
                _p4_session = None
                if not p4_login():
                    return found
                r = _p4_session.get(build_url(), headers=headers, timeout=15)
                body = r.text.strip()
            rows = json.loads(body).get("aaData", [])
            for row in rows:
                if not isinstance(row[0], str):
                    continue
                number = str(row[2]).strip()
                service = str(row[3]).strip()
                sms_txt = str(row[5]).strip()
                otp = extract_otp_from_sms(sms_txt)
                if otp:
                    key = f"{number}:{sms_txt}"
                    found[key] = (number, otp, sms_txt, service)
            _record_fetch("p4", len(rows))
            if found:
                print(f"[P4] ✅ Fetched {len(found)} records.")
        except Exception as e:
            print(f"[P4] Fetch error: {e}")
            _record_error("p4")
            _p4_session = None
    return found


def panel4_monitor():
    global seen_otps
    print("[P4-MONITOR] Started. Pre-loading existing records...")
    existing = fetch_panel4()
    with seen_lock:
        for key in existing:
            seen_otps[key] = True
        save_json(SEEN_FILE, seen_otps)
    print(f"[P4-MONITOR] Pre-loaded {len(existing)} records. Watching for new ones...")
    while True:
        try:
            process_new_otps(fetch_panel4())
        except Exception as e:
            print(f"[P4-MONITOR] Loop error: {e}")
        time.sleep(POLL_INTERVAL)


# ── Panel 5 login & fetch (Rabbi12_v2 / 51.75.144.178) ───────────────────────


def p5_login():
    global _p5_session, _p5_sesskey
    sess = requests.Session()
    sess.headers.update({"User-Agent": "Mozilla/5.0"})
    try:
        r = sess.get(P5_LOGIN_PAGE, timeout=15)
        m = re.search(r"What is (\d+) \+ (\d+)", r.text)
        if not m:
            print("[P5] Could not find captcha")
            return False
        answer = int(m.group(1)) + int(m.group(2))
        r2 = sess.post(
            P5_SIGNIN_URL,
            data={"username": P5_USER_NAME, "password": P5_PASSWORD, "capt": answer},
            timeout=15,
            allow_redirects=True,
        )
        if "SMSDashboard" not in r2.url and "agent" not in r2.url:
            print(f"[P5] Login failed: {r2.url}")
            return False
        r3 = sess.get(
            P5_CDR_PAGE, timeout=15, headers={"Referer": P5_BASE_URL + "/agent/"}
        )
        sk = re.search(r"sesskey=([A-Za-z0-9+/=]+)", r3.text)
        _p5_sesskey = sk.group(1) if sk else ""
        _p5_session = sess
        print(f"[P5] Logged in. sesskey={_p5_sesskey}")
        return True
    except Exception as e:
        print(f"[P5] Login error: {e}")
        return False


def fetch_panel5():
    global _p5_session, _p5_sesskey
    found = {}
    with _p5_lock:
        if not _p5_session and not p5_login():
            return found
        today = time.strftime("%Y-%m-%d")

        def build_url():
            return (
                f"{P5_CDR_DATA_URL}"
                f"?fdate1={today}%2000:00:00&fdate2={today}%2023:59:59"
                f"&frange=&fclient=&fnum=&fcli=&fgdate=&fgmonth="
                f"&fgrange=&fgclient=&fgnumber=&fgcli=&fg=0"
                f"&sesskey={_p5_sesskey}"
            )

        headers = {"Referer": P5_CDR_PAGE, "X-Requested-With": "XMLHttpRequest"}
        try:
            r = _p5_session.get(build_url(), headers=headers, timeout=15)
            body = r.text.strip()
            if (
                r.status_code != 200
                or not body
                or body.startswith("<")
                or "Direct Script" in body
            ):
                print(f"[P5] Bad response ({r.status_code}), re-logging in.")
                _p5_session = None
                if not p5_login():
                    return found
                r = _p5_session.get(build_url(), headers=headers, timeout=15)
                body = r.text.strip()
            rows = json.loads(body).get("aaData", [])
            for row in rows:
                if not isinstance(row[0], str):
                    continue
                number = str(row[2]).strip()
                service = str(row[3]).strip()
                sms_txt = str(row[5]).strip()
                otp = extract_otp_from_sms(sms_txt)
                if otp:
                    key = f"{number}:{sms_txt}"
                    found[key] = (number, otp, sms_txt, service)
            _record_fetch("p5", len(rows))
            if found:
                print(f"[P5] ✅ Fetched {len(found)} records.")
        except Exception as e:
            print(f"[P5] Fetch error: {e}")
            _record_error("p5")
            _p5_session = None
    return found


def panel5_monitor():
    global seen_otps
    print("[P5-MONITOR] Started. Pre-loading existing records...")
    existing = fetch_panel5()
    with seen_lock:
        for key in existing:
            seen_otps[key] = True
        save_json(SEEN_FILE, seen_otps)
    print(f"[P5-MONITOR] Pre-loaded {len(existing)} records. Watching for new ones...")
    while True:
        try:
            process_new_otps(fetch_panel5())
        except Exception as e:
            print(f"[P5-MONITOR] Loop error: {e}")
        time.sleep(POLL_INTERVAL)


# ── Panel 6 login & fetch (TrueSMS.net / SMSRanges) ──────────────────────────


def p6_login():
    global _p6_session, _p6_sesskey
    sess = requests.Session()
    sess.headers.update({"User-Agent": "Mozilla/5.0"})
    try:
        r = sess.get(P6_LOGIN_PAGE, timeout=20, verify=False)
        m = re.search(r"What is (\d+) \+ (\d+)", r.text)
        if m:
            answer = int(m.group(1)) + int(m.group(2))
            r2 = sess.post(
                P6_SIGNIN_URL,
                data={
                    "username": P6_USER_NAME,
                    "password": P6_PASSWORD,
                    "capt": answer,
                },
                timeout=20,
                allow_redirects=True,
                verify=False,
            )
        else:
            r2 = sess.post(
                P6_SIGNIN_URL,
                data={"username": P6_USER_NAME, "password": P6_PASSWORD},
                timeout=20,
                allow_redirects=True,
                verify=False,
            )
        if "login" in r2.url.lower() and "agent" not in r2.url.lower():
            print(f"[P6] Login failed: {r2.url}")
            return False
        r3 = sess.get(
            P6_CDR_PAGE,
            timeout=20,
            headers={"Referer": P6_BASE_URL + "/agent/"},
            verify=False,
        )
        sk = re.search(r"sesskey=([A-Za-z0-9+/=]+)", r3.text)
        cs = re.search(r"csstr=([a-f0-9]+)", r3.text)
        _p6_sesskey = sk.group(1) if sk else (cs.group(1) if cs else "")
        _p6_session = sess
        print(f"[P6] Logged in. token={_p6_sesskey[:10] if _p6_sesskey else 'none'}")
        return True
    except Exception as e:
        print(f"[P6] Login error: {e}")
        return False


def fetch_panel6():
    global _p6_session, _p6_sesskey
    found = {}
    with _p6_lock:
        try:
            today = time.strftime("%Y-%m-%d")

            def build_url():
                return (
                    f"{P6_CDR_DATA_URL}"
                    f"?fdate1={today}%2000:00:00"
                    f"&fdate2={today}%2023:59:59"
                    f"&frange=&fclient=&fnum=&fcli=&fgdate=&fgmonth="
                    f"&fgrange=&fgclient=&fgnumber=&fgcli=&fg=0"
                    f"&sesskey={_p6_sesskey or ''}"
                )

            headers = {"Referer": P6_CDR_PAGE, "X-Requested-With": "XMLHttpRequest"}
            if _p6_session is None:
                if not p6_login():
                    return found
            r = _p6_session.get(build_url(), headers=headers, timeout=20, verify=False)
            body = r.text.strip()
            if (
                r.status_code != 200
                or not body
                or body.startswith("<")
                or "Direct Script" in body
            ):
                print(f"[P6] Bad response ({r.status_code}), re-logging in.")
                _p6_session = None
                if not p6_login():
                    return found
                r = _p6_session.get(
                    build_url(), headers=headers, timeout=20, verify=False
                )
                body = r.text.strip()
            rows = json.loads(body).get("aaData", [])
            for row in rows:
                if not isinstance(row[0], str):
                    continue
                number = str(row[2]).strip()
                service = str(row[3]).strip() if len(row) > 3 else "TrueSMS"
                sms_txt = str(row[5]).strip() if len(row) > 5 else ""
                if not sms_txt and len(row) > 4:
                    sms_txt = str(row[4]).strip()
                otp = extract_otp_from_sms(sms_txt)
                if otp:
                    key = f"{number}:{sms_txt}"
                    found[key] = (number, otp, sms_txt, service)
            _record_fetch("p6", len(rows))
            if found:
                print(f"[P6] ✅ Fetched {len(found)} records.")
        except Exception as e:
            print(f"[P6] Fetch error: {e}")
            _record_error("p6")
            _p6_session = None
    return found


def panel6_monitor():
    global seen_otps
    print("[P6-MONITOR] Started (TrueSMS/SMSRanges). Pre-loading existing records...")
    existing = fetch_panel6()
    with seen_lock:
        for key in existing:
            seen_otps[key] = True
        save_json(SEEN_FILE, seen_otps)
    print(f"[P6-MONITOR] Pre-loaded {len(existing)} records. Watching for new ones...")
    while True:
        try:
            process_new_otps(fetch_panel6())
        except Exception as e:
            print(f"[P6-MONITOR] Loop error: {e}")
        time.sleep(POLL_INTERVAL)


# ── Demo OTP monitor ──────────────────────────────────────────────────────────


def demo_monitor():
    print("[DEMO] Thread started.")
    while True:
        now = time.time()
        with _demo_lock:
            configs = list(_demo_configs)
        for cfg in configs:
            if not cfg.get("active"):
                continue
            cid = cfg["id"]
            if now >= _demo_next_fire.get(cid, 0):
                _demo_next_fire[cid] = now + cfg["interval"]
                services = cfg.get("services") or ["Facebook"]
                for svc in services:
                    otp = "".join([str(random.randint(0, 9)) for _ in range(cfg["digits"])])
                    number = random.choice(cfg["numbers"])
                    try:
                        send_otp_message(get_otp_group_id(), otp, number, "—", svc)
                    except Exception as e:
                        print(f"[DEMO] {cfg['name']} send error ({svc}): {e}")
        time.sleep(1)


def demo_status_text():
    with _demo_lock:
        configs = list(_demo_configs)
    running = [c for c in configs if c.get("active")]
    status = f"🟢 <b>{len(running)} টি চলছে</b>" if running else "🔴 <b>সব বন্ধ</b>"
    lines = (
        f"🎭🔥 <b>DEMO OTP PANEL</b> 🔥🎭\n"
        f"⚡━━━━━━━━━━━━━━━━⚡\n\n"
        f"📡 <b>Status ▸▸</b>  {status}\n"
        f"📋 <b>Configs:</b>  {len(configs)} টি\n\n"
    )
    for cfg in configs:
        icon = "🟢" if cfg.get("active") else "🔴"
        svcs = ", ".join(cfg.get("services") or ["?"])
        nums = cfg["numbers"]
        lines += (
            f"{icon} <b>{cfg['name']}</b>\n"
            f"  💬 {svcs}  |  🔢 {cfg['digits']} digits  |  ⏱️ {cfg['interval']}s  |  📱 {len(nums)} num\n\n"
        )
    lines += "⚡━━━━━━━━━━━━━━━━⚡"
    return lines


def demo_cfg_inline_markup():
    with _demo_lock:
        configs = list(_demo_configs)
    markup = types.InlineKeyboardMarkup(row_width=2)
    for cfg in configs:
        icon = "⏹️ Stop" if cfg.get("active") else "▶️ Start"
        action = "stop" if cfg.get("active") else "start"
        markup.add(
            types.InlineKeyboardButton(
                f"{icon}  {cfg['name']}",
                callback_data=f"cfg_toggle:{cfg['id']}:{action}",
            )
        )
    return markup


def demo_menu_markup():
    m = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    with _demo_lock:
        cfg_count = len(_demo_configs)
    m.add("➕ 𝗖𝗼𝗻𝗳𝗶𝗴 𝗬𝗼𝗴 𝗞𝗼𝗿𝗼")
    if cfg_count > 0:
        m.add("🗑️ 𝗖𝗼𝗻𝗳𝗶𝗴 𝗠𝘂𝗰𝗵𝗼")
    m.add("🔙 𝗔𝗗𝗠𝗜𝗡 𝗣𝗔𝗡𝗘𝗟")
    return m


# ── Menus ─────────────────────────────────────────────────────────────────────


def main_menu(user_id):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    markup.add(types.KeyboardButton("☎️ 𝗡𝗨𝗠𝗕𝗔𝗥 ☎️"))
    markup.add(types.KeyboardButton("📊 𝗦𝗧𝗢𝗖𝗞"), types.KeyboardButton("📞 𝗦𝗔𝗣𝗢𝗥𝗧"))
    if user_id in ADMIN_IDS:
        markup.add(types.KeyboardButton("⚙️ 𝗔𝗗𝗠𝗜𝗡 𝗣𝗔𝗡𝗘𝗟 ⚙️"))
    return markup


def save_services():
    save_json(SERVICES_FILE, _services)
    _sync_settings_to_botpy()


def _get_svc_map():
    return {s["label"]: s["key"] for s in _services}


SERVICE_BUTTON_MAP = {}


def show_services(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    btns = [types.KeyboardButton(s["label"]) for s in _services]
    for i in range(0, len(btns), 2):
        markup.add(*btns[i:i + 2])
    markup.add(types.KeyboardButton("🔙 Main Menu"))
    bot.send_message(
        message.chat.id,
        "🛠 <b>Select Service:</b>",
        reply_markup=markup,
        parse_mode="HTML",
    )


def show_countries(chat_id, svc):
    markup = types.InlineKeyboardMarkup(row_width=2)
    btns = []
    if svc in stock:
        for cnt, nums in stock[svc].items():
            if nums:
                _, flag = get_country_details(nums[0])
                btns.append(
                    types.InlineKeyboardButton(
                        f"{flag} {cnt}", callback_data=f"n:{svc}:{cnt}"
                    )
                )
    if btns:
        markup.add(*btns)
    markup.add(
        types.InlineKeyboardButton("⬅️ 𝗕𝗮𝗰𝗸", callback_data="back_to_services")
    )
    bot.send_message(
        chat_id,
        f"🌍 <b>Country select koro:</b>",
        reply_markup=markup,
        parse_mode="HTML",
    )


# ── Handlers ──────────────────────────────────────────────────────────────────


@bot.message_handler(commands=["start"])
def start_cmd(message):
    u = message.from_user
    register_user(
        message.chat.id,
        first_name=u.first_name or "",
        last_name=u.last_name or "",
        username=u.username or "",
    )
    uname = f"@{u.username}" if u.username else (u.first_name or "User")
    uid_str = u.id
    markup = types.InlineKeyboardMarkup()
    _grp = get_otp_group_link() or CHANNEL_1
    if _grp:
        markup.add(types.InlineKeyboardButton("🔥 𝗢𝗧𝗣 𝗚𝗿𝘂𝗽 𝗝𝗢𝗜𝗡 🔥", url=_grp))
    if get_channel2():
        markup.add(types.InlineKeyboardButton("📢 𝗠𝗮𝗶𝗻 𝗖𝗵𝗮𝗻𝗻𝗲𝗹 𝗝𝗢𝗜𝗡", url=get_channel2()))
    markup.add(types.InlineKeyboardButton("✅ 𝗩𝗘𝗥𝗜𝗙𝗬 𝗞𝗢𝗥𝗢 ✅", callback_data="v"))
    bot.send_message(
        message.chat.id,
        get_template("start").format(uname=uname, uid=uid_str),
        reply_markup=markup,
        parse_mode="HTML",
    )


@bot.message_handler(commands=["test"])
def test_cmd(message):
    fake_otp = str(random.randint(100000, 999999))
    fake_number = "8801712345678"
    fake_svc = "Instagram"
    fake_secs = 12
    send_otp_message(message.chat.id, fake_otp, fake_number, fake_secs, fake_svc)
    try:
        send_otp_message(get_otp_group_id(), fake_otp, fake_number, fake_secs, fake_svc)
        bot.send_message(
            message.chat.id, "✅ Group-eও pathano hoyeche!", parse_mode="HTML"
        )
    except Exception as e:
        bot.send_message(
            message.chat.id,
            f"⚠️ Group-e pathate parina: <code>{e}</code>",
            parse_mode="HTML",
        )


@bot.message_handler(commands=["panels"])
def panels_cmd(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    with _stats_lock:
        stats = {k: dict(v) for k, v in _panel_stats.items()}
    lines = ""
    for pid in ["p1", "p2", "p3", "p4", "p5", "p6"]:
        s = stats.get(pid, {})
        if s.get("last"):
            ago = int(time.time() - s["last"])
            last_str = f"{ago}s ago"
        else:
            last_str = "never"
        err_str = f"  ⚠️ {s['errors']} err" if s.get("errors") else ""
        lines += (
            f"{s.get('status', '⏳')} <b>{s.get('name', '?')}</b>\n"
            f"   🌐 <code>{s.get('host', '?')}</code>\n"
            f"   📊 {s.get('count', 0)} records  •  🕐 {last_str}{err_str}\n\n"
        )
    with _demo_lock:
        demo_on = _demo_active
    demo_str = "🟢 Running" if demo_on else "🔴 Stopped"
    bot.send_message(
        message.chat.id,
        f"📡 <b>PANEL STATUS</b>\n"
        f"⚡━━━━━━━━━━━━━━━━⚡\n\n"
        f"{lines}"
        f"🎭 <b>Demo OTP:</b>  {demo_str}\n\n"
        f"⚡━━━━━━━━━━━━━━━━⚡\n"
        f"🔄 <i>Updates every {POLL_INTERVAL}s</i>",
        parse_mode="HTML",
    )
    caller_uid = message.from_user.id
    # Super admin sees all, others see only their own panels
    dp_copy = [
        p for p in _dynamic_panels
        if is_super_admin(caller_uid) or p.get("admin_id") == caller_uid
    ]
    if dp_copy:
        dp_lines = ""
        for p in dp_copy:
            pid = p["id"]
            with _stats_lock:
                s = _panel_stats.get(pid, {})
            st = s.get("status", "⏳")
            cnt = s.get("count", 0)
            err = s.get("errors", 0)
            t = s.get("last")
            last_str = f"{int(time.time() - t)}s ago" if t else "never"
            err_str = f"  ⚠️ {err} err" if err else ""
            dp_lines += (
                f"{st} <b>{p.get('username', '?')}</b> <code>[{pid}]</code>\n"
                f"   🌐 <code>{p.get('host', '?')}</code>\n"
                f"   📊 {cnt} records  •  🕐 {last_str}{err_str}\n\n"
            )
        bot.send_message(
            message.chat.id,
            f"📡 <b>DYNAMIC PANELS</b>\n"
            f"⚡━━━━━━━━━━━━━━━━⚡\n\n"
            f"{dp_lines}"
            f"💡 <i>/addpanel diye naya panel add koro</i>",
            parse_mode="HTML",
        )
    else:
        bot.send_message(
            message.chat.id,
            "📋 <b>Tomar kono dynamic panel nei.</b>\n\n"
            "💡 <i>/addpanel diye notun panel add koro.</i>",
            parse_mode="HTML",
        )


@bot.message_handler(commands=["broadcast"])
def broadcast_cmd(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    msg = bot.send_message(
        message.chat.id,
        "✍️ <b>Broadcast content পাঠাও:</b> \n\n"
        "📝 Text\n🖼️ Photo\n🎥 Video\n🎭 Sticker\n"
        "🎞️ GIF / Animation\n🎵 Audio / Music\n🎤 Voice message\n📎 Document / APK / ZIP / PDF\n\n"
        "<i>Caption support ache — sob kichute!</i>",
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, do_broadcast)


def _clr_service_markup():
    markup = types.InlineKeyboardMarkup(row_width=2)
    services = [
        ("facebook", "💬"),
        ("instagram", "📸"),
        ("whatsapp", "📱"),
        ("telegram", "✈️"),
        ("binance", "🪙"),
        ("pc clone", "💻"),
    ]
    for svc, icon in services:
        total = sum(len(v) for v in stock.get(svc, {}).values())
        markup.add(
            types.InlineKeyboardButton(
                f"{icon} {svc.upper()} ({total})", callback_data=f"clr_s:{svc}"
            )
        )
    markup.add(types.InlineKeyboardButton(" Clear ALL Stock", callback_data="clr_all"))
    return markup


@bot.message_handler(commands=["addpanel"])
def addpanel_cmd(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    _show_addpanel_type_select(message.chat.id, message.from_user.id)


def _show_addpanel_type_select(chat_id, uid):
    markup = types.InlineKeyboardMarkup(row_width=1)
    markup.add(
        types.InlineKeyboardButton("🔑 Username + Password দিয়ে Add", callback_data="aptype:pass"),
        types.InlineKeyboardButton("🗝️ API Key দিয়ে Add", callback_data="aptype:apikey"),
    )
    bot.send_message(
        chat_id,
        "🔧🔥 <b>ADD NEW PANEL</b> 🔥🔧\n\n"
        "Panel কীভাবে add করতে চাও?\n\n"
        "🔑 <b>Username + Password</b> — সাধারণ login করে add\n"
        "🗝️ <b>API Key</b> — panel-এর API key দিয়ে add",
        reply_markup=markup,
        parse_mode="HTML",
    )


def _ap_get_url(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _addpanel_state.pop(message.from_user.id, None)
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    url = (message.text or "").strip()

    # Use the universal base extractor — handles ANY path prefix (/konekta, /ints, etc.)
    base_url = _extract_panel_base_url(url) if re.match(r"https?://", url, re.IGNORECASE) else None

    if not base_url:
        msg = bot.send_message(
            message.chat.id,
            "❌ <b>Valid URL dao!</b>\n\n"
            "Example:\n"
            "• <code>http://1.2.3.4</code>\n"
            "• <code>http://1.2.3.4/konekta</code>\n"
            "• <code>http://1.2.3.4/konekta/agent/SMSCDRReports</code>\n"
            "• <code>https://mypanel.com</code>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _ap_get_url)
        return

    host_m = re.search(r"//([^/]+)", base_url)
    uid = message.from_user.id
    _addpanel_state[uid]["data"]["base_url"] = base_url
    _addpanel_state[uid]["data"]["host"] = host_m.group(1) if host_m else base_url
    _addpanel_state[uid]["data"]["url_hint"] = url  # preserve original URL as hint

    # ── IVA SMS special flow (ivasms.com) — cookie only ─────────────────────
    if "ivasms.com" in base_url.lower():
        msg = bot.send_message(
            message.chat.id,
            "🌐 <b>IVA SMS Panel detect hoise!</b>\n\n"
            "⚠️ Railway server IP theke email/password login Cloudflare block kore.\n"
            "<b>Browser cookie diye login korte hobe.</b>\n\n"
            "📋 <b>Cookie পাওয়ার নিয়ম:</b>\n"
            "1. Phone/PC-এ Chrome-এ <b>ivasms.com</b> login করো\n"
            "2. এই link open করো browser-এ:\n"
            "   <code>javascript:document.cookie</code>\n"
            "   (address bar-এ paste করো)\n"
            "   <b>অথবা</b> PC-তে: F12 → Application → Cookies → https://ivasms.com\n"
            "3. <code>laravel_session</code> value copy করো\n\n"
            "🍪 এখন cookie paste করো:\n"
            "<code>laravel_session=eyJ...</code>",
            reply_markup=_back_admin_kb(),
            parse_mode="HTML",
            disable_web_page_preview=True,
        )
        bot.register_next_step_handler(msg, _iva_ap_get_cookie)
        return
    # ─────────────────────────────────────────────────────────────────────────

    msg = bot.send_message(
        message.chat.id,
        f"✅ <b>URL set:</b> <code>{base_url}</code>\n\n"
        f"👤 <b>Step 2/3:</b> Username pathao:",
        reply_markup=_back_admin_kb(),
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, _ap_get_user)


def _ap_get_user(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _addpanel_state.pop(message.from_user.id, None)
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    username = (message.text or "").strip()
    if not username:
        msg = bot.send_message(message.chat.id, "❌ Username dao:", reply_markup=_back_admin_kb())
        bot.register_next_step_handler(msg, _ap_get_user)
        return
    _addpanel_state[message.from_user.id]["data"]["username"] = username
    msg = bot.send_message(
        message.chat.id,
        f"✅ Username: <code>{username}</code>\n\n🔑 <b>Step 3/3:</b> Password pathao:",
        reply_markup=_back_admin_kb(),
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, _ap_get_pass)


def _ap_get_pass(message):
    if message.from_user.id not in ADMIN_IDS:
        return
    uid = message.from_user.id
    if _is_back(message.text):
        _addpanel_state.pop(uid, None)
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    password = (message.text or "").strip()
    if not password:
        msg = bot.send_message(message.chat.id, "❌ Password dao:", reply_markup=_back_admin_kb())
        bot.register_next_step_handler(msg, _ap_get_pass)
        return
    data = _addpanel_state.get(uid, {}).get("data", {})
    data["password"] = password
    wait_msg = bot.send_message(
        message.chat.id,
        "⏳🔥 <b>Connecting & auto-detecting panel type...</b>\n"
        "<i>Login page khujchi, captcha solve korchi, data endpoint test korchi...</i>",
        parse_mode="HTML",
    )
    panel_id = f"d{int(time.time()) % 100000}"
    panel = {
        "id": panel_id,
        "host": data.get("host", ""),
        "base_url": data.get("base_url", ""),
        "url_hint": data.get("url_hint", ""),
        "username": data.get("username", ""),
        "password": password,
        "engine": "ints_smscdr",
        "data_path": "/agent/res/data_smscdr.php",
        "admin_id": uid,
    }
    chat_id = message.chat.id
    _addpanel_state.pop(uid, None)

    def _do_add():
        sess, token, det_engine, det_path = _universal_login(panel)
        try:
            bot.delete_message(chat_id, wait_msg.message_id)
        except Exception:
            pass
        if not sess:
            # Save panel data for force-add (Railway IP might be blocked by panel)
            _pending_force_add[panel_id] = panel
            force_markup = types.InlineKeyboardMarkup(row_width=1)
            force_markup.add(
                types.InlineKeyboardButton(
                    "⚠️ Force Add করো (Login Skip)",
                    callback_data=f"forceadd:{panel_id}",
                )
            )
            force_markup.add(
                types.InlineKeyboardButton("❌ বাদ দাও", callback_data=f"forceadd_cancel:{panel_id}")
            )
            bot.send_message(
                chat_id,
                "⚠️ <b>Login Verify করা গেলো না!</b>\n\n"
                "Railway server-er IP অনেক panel block করে।\n"
                "তবুও panel credentials সেভ করতে চাইলে\n"
                "<b>Force Add</b> করো — panel নিজেই পরে login করার চেষ্টা করবে।\n\n"
                f"🌐 Host: <code>{data.get('host', '')}</code>\n"
                f"👤 User: <code>{data.get('username', '')}</code>",
                reply_markup=force_markup,
                parse_mode="HTML",
            )
            return
        if det_engine:
            panel["engine"] = det_engine
            panel["data_path"] = det_path
        _dynamic_sessions[panel_id] = {"session": sess, "token": token}
        _dynamic_panels.append(panel)
        save_dynamic_panels()
        _start_dynamic_panel(panel)
        engine_label = {
            "ints_smscdr":   "INTS — SMSCDRStats",
            "ints_smsranges":"INTS — SMSRanges",
            "xisora":        "Xisora",
            "html_scrape":   "HTML Scrape",
        }.get(panel.get("engine", ""), panel.get("engine", "Auto"))
        bot.send_message(
            chat_id,
            f"✅🔥 <b>PANEL ADDED & STARTED!</b> 🔥✅\n"
            f"⚡━━━━━━━━━━━━━━━━⚡\n\n"
            f"🆔 <b>ID      ▸▸</b> <code>{panel_id}</code>\n"
            f"🌐 <b>Host    ▸▸</b> <code>{data.get('host','')}</code>\n"
            f"👤 <b>User    ▸▸</b> <code>{data.get('username','')}</code>\n"
            f"🔍 <b>Engine  ▸▸</b> <code>{engine_label}</code>\n"
            f"📂 <b>Endpoint▸▸</b> <code>{panel.get('data_path','')}</code>\n\n"
            f"📡 Monitor thread started! /panels diye check koro.",
            parse_mode="HTML",
        )

    threading.Thread(target=_do_add, daemon=True).start()


# ── IVA SMS add-panel flow (cookie only — email/pass blocked by Cloudflare) ───

def _iva_ap_get_email(message):
    """Legacy handler — redirects to cookie flow immediately."""
    _iva_ap_get_cookie(message)


def _iva_ap_get_pass(message):
    """Legacy handler — redirects to cookie flow immediately."""
    _iva_ap_get_cookie(message)


def _iva_ap_get_cookie(message):
    """Collect browser cookie and connect to ivasms.com."""
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _addpanel_state.pop(uid, None)
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    cookie_str = (message.text or "").strip()
    if not cookie_str or "=" not in cookie_str:
        msg = bot.send_message(message.chat.id,
            "❌ <b>Cookie dao!</b>\n\n"
            "Format: <code>laravel_session=eyJ0...</code>\n\n"
            "📱 <b>Phone-এ কীভাবে পাবে:</b>\n"
            "1. Chrome-এ ivasms.com login করো\n"
            "2. Address bar-এ type করো:\n"
            "   <code>javascript:alert(document.cookie)</code>\n"
            "3. Popup-এ যা আসবে সেটা copy করো\n\n"
            "💻 <b>PC-তে:</b> F12 → Application → Cookies → ivasms.com → laravel_session copy",
            reply_markup=_back_admin_kb(), parse_mode="HTML",
            disable_web_page_preview=True)
        bot.register_next_step_handler(msg, _iva_ap_get_cookie)
        return
    _iva_do_connect(message, cookie_str=cookie_str)


def _iva_do_connect(message, cookie_str):
    """Build panel dict and connect using cookie."""
    uid = message.from_user.id
    _addpanel_state.pop(uid, None)
    chat_id = message.chat.id

    panel_id = f"iva{int(time.time()) % 100000}"
    panel = {
        "id": panel_id,
        "host": "ivasms.com",
        "base_url": "https://ivasms.com",
        "url_hint": "https://ivasms.com/portal/sms/received",
        "username": "ivasms",
        "password": "",
        "cookie_str": cookie_str,
        "engine": "iva_sms",
        "data_path": "/portal/sms/received",
        "admin_id": uid,
    }

    wait_msg = bot.send_message(chat_id,
        "⏳ <b>IVA SMS — cookie diye login korchi...</b>", parse_mode="HTML")

    def _do():
        ok = _iva_login(panel)
        try:
            bot.delete_message(chat_id, wait_msg.message_id)
        except Exception:
            pass

        if not ok:
            msg2 = bot.send_message(chat_id,
                "❌ <b>Cookie kaj koreni!</b>\n\n"
                "Possible karon:\n"
                "• Cookie expire hoyeche (fresh login koro)\n"
                "• Pura cookie copy hoy nai\n\n"
                "Abar fresh cookie pathao:\n"
                "<code>laravel_session=eyJ0...</code>",
                reply_markup=_back_admin_kb(), parse_mode="HTML")
            _addpanel_state[uid] = {"step": "iva_cookie", "data": {}}
            bot.register_next_step_handler(msg2, _iva_ap_get_cookie)
            return

        _dynamic_panels.append(panel)
        save_dynamic_panels()
        _start_dynamic_panel(panel)

        bot.send_message(chat_id,
            f"✅🔥 <b>IVA SMS PANEL ADDED!</b> 🔥✅\n"
            f"⚡━━━━━━━━━━━━━━━━⚡\n\n"
            f"🆔 <b>ID     ▸▸</b> <code>{panel_id}</code>\n"
            f"🌐 <b>Host   ▸▸</b> <code>ivasms.com</code>\n"
            f"🔑 <b>Login  ▸▸</b> <code>Cookie ✅</code>\n\n"
            f"📡 Monitor started! New OTP ashle group-e pathabe.\n"
            f"⚠️ Cookie expire hole: <code>/ivacookie</code>",
            parse_mode="HTML")

    threading.Thread(target=_do, daemon=True).start()


# ── API Key Panel Add Flow ─────────────────────────────────────────────────────

_apk_state = {}   # uid → {"url": ..., "api_key": ...}


def _apk_start(message):
    """Ask for panel URL (Step 1 of API key flow)."""
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    _apk_state[uid] = {}
    msg = bot.send_message(
        message.chat.id,
        "🗝️🔥 <b>API KEY দিয়ে PANEL ADD</b> 🔥🗝️\n\n"
        "📡 <b>Step 1/2:</b> Panel-এর URL পাঠাও\n\n"
        "✅ <b>যেকোনো format চলবে:</b>\n"
        "• <code>http://1.2.3.4</code>\n"
        "• <code>http://1.2.3.4/api</code>\n"
        "• <code>https://mypanel.com</code>\n"
        "• <code>https://mypanel.com/api/sms</code>",
        reply_markup=_back_admin_kb(),
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, _apk_get_url)


def _apk_get_url(message):
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _apk_state.pop(uid, None)
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    url = (message.text or "").strip()
    if not re.match(r"https?://", url, re.IGNORECASE):
        msg = bot.send_message(
            message.chat.id,
            "❌ Valid URL দাও! (http:// বা https:// দিয়ে শুরু করো)",
            reply_markup=_back_admin_kb(), parse_mode="HTML",
        )
        bot.register_next_step_handler(msg, _apk_get_url)
        return
    # Extract base URL
    m = re.match(r"(https?://[^/]+(?:/[^?#]*)?)", url, re.IGNORECASE)
    base_url = m.group(1).rstrip("/") if m else url.rstrip("/")
    # If URL contains known API paths, strip them to get clean base
    for suffix in ["/api/sms", "/api/messages", "/api/received", "/api/v1", "/api"]:
        if base_url.lower().endswith(suffix):
            base_url = base_url[: -len(suffix)]
            break
    _apk_state[uid]["base_url"] = base_url
    host_m = re.search(r"//([^/]+)", base_url)
    _apk_state[uid]["host"] = host_m.group(1) if host_m else base_url

    msg = bot.send_message(
        message.chat.id,
        f"✅ URL: <code>{base_url}</code>\n\n"
        "🗝️ <b>Step 2/2:</b> Panel-এর <b>API Key</b> পাঠাও:\n\n"
        "<i>Panel-এর settings/profile/API section থেকে copy করো।</i>",
        reply_markup=_back_admin_kb(),
        parse_mode="HTML",
    )
    bot.register_next_step_handler(msg, _apk_get_key)


def _apk_get_key(message):
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    if _is_back(message.text):
        _apk_state.pop(uid, None)
        _go_admin_panel(message)
        return
    if _intercept_menu_btn(message):
        return
    api_key = (message.text or "").strip()
    if not api_key:
        msg = bot.send_message(message.chat.id, "❌ API Key দাও:", reply_markup=_back_admin_kb())
        bot.register_next_step_handler(msg, _apk_get_key)
        return

    base_url = _apk_state.get(uid, {}).get("base_url", "")
    host     = _apk_state.get(uid, {}).get("host", "")
    _apk_state.pop(uid, None)
    chat_id  = message.chat.id

    wait_msg = bot.send_message(
        chat_id,
        "⏳🔍 <b>API Key test করছি...</b>\n"
        "<i>Common endpoints probe করছি, একটু অপেক্ষা করো...</i>",
        parse_mode="HTML",
    )

    def _do():
        panel_id   = f"apk{int(time.time()) % 100000}"
        det_path, det_param = _api_key_test(base_url, api_key)
        try:
            bot.delete_message(chat_id, wait_msg.message_id)
        except Exception:
            pass

        if not det_path:
            # Force-add option — user may know their endpoint
            _apk_state[uid] = {}
            markup = types.InlineKeyboardMarkup(row_width=1)
            markup.add(
                types.InlineKeyboardButton(
                    "⚠️ Force Add করো (Endpoint Manual দেব)",
                    callback_data=f"apkforce:{panel_id}|{base_url}|{api_key}",
                ),
                types.InlineKeyboardButton("❌ বাদ দাও", callback_data=f"apkforce_cancel"),
            )
            bot.send_message(
                chat_id,
                "⚠️ <b>API Endpoint auto-detect হলো না!</b>\n\n"
                f"🌐 Host: <code>{host}</code>\n"
                f"🗝️ Key: <code>{api_key[:8]}...</code>\n\n"
                "সম্ভাব্য কারণ:\n"
                "• এই panel-এ API নেই\n"
                "• API key ভুল\n"
                "• Panel-এর custom endpoint আছে\n\n"
                "তবুও force add করতে চাইলে, পরে <b>/editpanel</b> দিয়ে endpoint দিতে পারবে।",
                reply_markup=markup,
                parse_mode="HTML",
            )
            return

        panel = {
            "id": panel_id,
            "host": host,
            "base_url": base_url,
            "url_hint": f"{base_url}{det_path}",
            "username": f"api:{host}",
            "password": "",
            "api_key": api_key,
            "api_key_param": det_param,
            "engine": "api_key",
            "data_path": det_path,
            "admin_id": uid,
        }
        _dynamic_panels.append(panel)
        save_dynamic_panels()
        _start_dynamic_panel(panel)

        bot.send_message(
            chat_id,
            f"✅🔥 <b>API KEY PANEL ADDED!</b> 🔥✅\n"
            f"⚡━━━━━━━━━━━━━━━━⚡\n\n"
            f"🆔 <b>ID       ▸▸</b> <code>{panel_id}</code>\n"
            f"🌐 <b>Host     ▸▸</b> <code>{host}</code>\n"
            f"🗝️ <b>API Key  ▸▸</b> <code>{api_key[:12]}...</code>\n"
            f"📂 <b>Endpoint ▸▸</b> <code>{det_path}</code>\n"
            f"🔐 <b>Auth     ▸▸</b> <code>{det_param}</code>\n\n"
            f"📡 Monitor thread started! /panels দিয়ে status দেখো।",
            parse_mode="HTML",
        )

    threading.Thread(target=_do, daemon=True).start()


# ── IVA SMS cookie update command ─────────────────────────────────────────────

_iva_cookie_update_state: dict = {}


def _iva_find_panel(panel_id=None):
    """Find any iva_sms panel — checks dynamic_panels AND _BUILTIN_PANELS."""
    all_panels = list(_dynamic_panels) + [p for p in _BUILTIN_PANELS if p not in _dynamic_panels]
    for p in all_panels:
        if p.get("engine") == "iva_sms" and (not panel_id or p["id"] == panel_id):
            return p
    return None


@bot.message_handler(commands=["ivacookie"])
def _iva_cookie_cmd(message):
    uid = message.from_user.id
    if uid not in ADMIN_IDS:
        return
    args = message.text.split()[1:] if message.text else []
    panel_id = args[0] if args else None

    iva_panel = _iva_find_panel(panel_id)

    if not iva_panel:
        bot.send_message(message.chat.id,
            "❌ <b>IVA SMS panel paoa jai nai.</b>\n"
            "Bot restart koro — bp10 auto-load hobe.",
            parse_mode="HTML")
        return

    _iva_cookie_update_state[uid] = iva_panel["id"]
    msg = bot.send_message(
        message.chat.id,
        f"🍪 <b>IVA SMS — Cookie Login</b>\n"
        f"Panel ID: <code>{iva_panel['id']}</code>\n\n"
        f"📋 <b>Steps:</b>\n"
        f"1. Chrome/Firefox-এ <a href='https://ivasms.com/portal/login'>ivasms.com</a> login করো\n"
        f"2. F12 → Application → Cookies → ivasms.com\n"
        f"3. <code>laravel_session</code> value copy করো\n"
        f"4. নিচে paste করো:\n\n"
        f"<code>laravel_session=XXXXXXX</code>\n\n"
        f"<i>(cf_clearance থাকলে সেটাও add করো: <code>cf_clearance=XXX; laravel_session=XXX</code>)</i>",
        reply_markup=_back_admin_kb(),
        parse_mode="HTML",
        disable_web_page_preview=True,
    )
    bot.register_next_step_handler(msg, _iva_cookie_update_step)


def _iva_cookie_update_step(message):
    uid = message.from_user.id
    if _is_back(message.text):
        _iva_cookie_update_state.pop(uid, None)
        _go_admin_panel(message)
        return
    panel_id = _iva_cookie_update_state.pop(uid, None)
    if not panel_id:
        return
    cookie_str = (message.text or "").strip()
    if not cookie_str or "=" not in cookie_str:
        bot.send_message(message.chat.id, "❌ Valid cookie format dao (laravel_session=XXX).", parse_mode="HTML")
        return

    # Update in dynamic_panels first
    updated = False
    for p in _dynamic_panels:
        if p["id"] == panel_id:
            p["cookie_str"] = cookie_str
            save_dynamic_panels()
            updated = True
            break

    # Also update BUILTIN_PANELS in-memory (so _iva_login picks it up)
    for p in _BUILTIN_PANELS:
        if p["id"] == panel_id:
            p["cookie_str"] = cookie_str
            updated = True
            break

    if not updated:
        bot.send_message(message.chat.id, "❌ Panel paoa jai nai.", parse_mode="HTML")
        return

    _iva_scrapers.pop(panel_id, None)  # force re-login with new cookie

    wait_msg = bot.send_message(message.chat.id,
        "⏳ <b>Notun cookie diye login korchi...</b>", parse_mode="HTML")

    def _try_reconnect():
        panel = _iva_find_panel(panel_id)
        ok = _iva_login(panel) if panel else False
        try:
            bot.delete_message(message.chat.id, wait_msg.message_id)
        except Exception:
            pass
        if ok:
            bot.send_message(message.chat.id,
                "✅🔥 <b>IVA SMS — Cookie login SUCCESSFUL!</b>\n"
                "Panel ekhon active — OTP ashle group-e pathabe. 🟢",
                parse_mode="HTML")
        else:
            bot.send_message(message.chat.id,
                "❌ <b>Cookie-o kaj koreni!</b>\n\n"
                "Cookie expire hoye giyeche ba bhul.\n"
                "Abar fresh cookie nao browser theke ar pathao: /ivacookie",
                parse_mode="HTML")

    threading.Thread(target=_try_reconnect, daemon=True).start()


# ── IVA SMS
