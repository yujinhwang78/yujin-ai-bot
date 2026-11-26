import os
import io
import re
import json
import copy
import uuid
import unicodedata
import base64
import zipfile
import asyncio
import sys
import time
import datetime as dt
from zoneinfo import ZoneInfo
import logging
import imaplib
import smtplib
import email
import email.utils
from email.header import decode_header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

import httpx
import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import PatternFill
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.utils import ImageReader
from pypdf import PdfReader, PdfWriter

try:
    # 구글 드라이브 자동 업로드용(설정 안 해두면 그냥 이 기능만 꺼짐).
    # 서비스 계정은 개인 "내 드라이브"에는 저장 용량이 아예 없어서(구글 정책상 제약,
    # storageQuotaExceeded 오류) 유진님 개인 계정 권한으로 직접 올리는 OAuth 방식을 씀.
    from google.oauth2.credentials import Credentials as _gdrive_user_credentials
    from google.auth.transport.requests import Request as _gdrive_auth_request
    from googleapiclient.discovery import build as _gdrive_build
    from googleapiclient.http import MediaInMemoryUpload as _gdrive_media
    from googleapiclient.http import MediaIoBaseDownload as _gdrive_media_download
except ImportError:
    _gdrive_user_credentials = None
    _gdrive_auth_request = None
    _gdrive_build = None
    _gdrive_media = None
    _gdrive_media_download = None
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters,
)
from anthropic import Anthropic

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)
# httpx가 요청 주소를 INFO로 찍으면 텔레그램 봇 토큰이 Render 로그에 그대로 남음(2026-11-18 발견).
# 요청 주소 로그는 끄고 경고/오류만 남김.
for _noisy in ("httpx", "httpx2", "httpcore"):
    logging.getLogger(_noisy).setLevel(logging.WARNING)

TELEGRAM_TOKEN = os.environ["TELEGRAM_TOKEN"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
ALLOWED_USER_ID = int(os.environ["ALLOWED_USER_ID"])

GMAIL_ADDRESS = os.environ.get("GMAIL_ADDRESS")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD")
MAIL_CHECK_INTERVAL = int(os.environ.get("MAIL_CHECK_INTERVAL", "120"))

# 메일 보낼 때 "보낸사람"으로 표시할 주소. 지정 안 하면 GMAIL_ADDRESS 그대로 사용.
# (구글 계정에 "다른 이메일 주소로 보내기" 별칭으로 인증된 주소여야 정상 발송돼요)
MAIL_FROM_ADDRESS = os.environ.get("MAIL_FROM_ADDRESS", GMAIL_ADDRESS)

# 보낸사람은 그대로 GMAIL_ADDRESS로 두되(별도 별칭 인증 없이 바로 되는 방식), 담당자가
# "답장"을 누르면 유진님의 실제 업무 메일(네이버, 구글로 자동전달 중)로 가도록 지정하는 주소
MAIL_REPLY_TO = os.environ.get("MAIL_REPLY_TO", MAIL_FROM_ADDRESS)

MODEL_NAME = os.environ.get("MODEL_NAME", "claude-haiku-4-5-20251001")

# 가입증명서를 구글 드라이브(유진님 PC에 동기화된 폴더)에도 자동으로 올려주는 기능용 설정.
# 유진님의 개인 구글 계정(bubal5503@gmail.com) 권한으로 직접 올리는 OAuth 방식. 처음 한 번만
# 유진님이 로그인/동의해서 refresh token을 받아두면, 이후로는 그 토큰으로 계속 자동 갱신하며
# 접속해서 유진님 계정의 "내 드라이브" 용량으로 파일을 저장함(서비스 계정은 이게 안 됨).
GDRIVE_OAUTH_CLIENT_ID = os.environ.get("GDRIVE_OAUTH_CLIENT_ID")
GDRIVE_OAUTH_CLIENT_SECRET = os.environ.get("GDRIVE_OAUTH_CLIENT_SECRET")
GDRIVE_OAUTH_REFRESH_TOKEN = os.environ.get("GDRIVE_OAUTH_REFRESH_TOKEN")
GDRIVE_FOLDER_ID = os.environ.get("GDRIVE_FOLDER_ID")

# 카카오 REST API 키 (카카오맵 대중교통/도보/자전거 길찾기 + 카카오모빌리티 자동차 길찾기 공용)
KAKAO_REST_API_KEY = os.environ.get("KAKAO_REST_API_KEY")

client = Anthropic(api_key=ANTHROPIC_API_KEY)

# 가입증명서 자동 생성용 자산(폰트/템플릿) 경로
# assets 폴더가 있으면 그 안에서, 없으면(루트에 바로 올린 경우) bot.py와 같은 위치에서 찾음
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 통합파일/증권번호/브랜드 설정처럼 "재배포해도 사라지면 안 되는" 데이터를 저장할 위치.
# Render에서 영구 디스크(Persistent Disk)를 만들어 PERSIST_DIR 환경변수로 그 마운트 경로를
# 지정해두면 그쪽에 저장하고, 설정 안 해두면(기본값) 지금처럼 코드 폴더에 저장함(재배포시 초기화).
PERSIST_DIR = os.environ.get("PERSIST_DIR", _BASE_DIR)
os.makedirs(PERSIST_DIR, exist_ok=True)

# 집/회사 등 저장된 위치를 담아두는 파일 (재시작해도 남아있도록 디스크에 저장)
PLACES_FILE = os.path.join(PERSIST_DIR, "places.json")


def _find_file_ci(directory: str, filename: str) -> str | None:
    """Render 서버(Linux)는 대소문자를 구분해서, GitHub 웹에서 파일 이름을 바꿀 때
    확장자가 의도치 않게 .PDF처럼 대문자로 바뀌어 있으면 .pdf로 찾다가 못 찾는 경우가 생김.
    그래서 폴더 안 파일 목록을 대소문자 구분 없이 한 번 더 확인함."""
    try:
        entries = os.listdir(directory)
    except OSError:
        return None
    target = filename.lower()
    for entry in entries:
        if entry.lower() == target:
            return os.path.join(directory, entry)
    return None


def _find_asset(filename: str) -> str:
    # GitHub 웹으로 파일을 올릴 때 이름 첫 글자가 의도치 않게 대문자로 바뀌는 경우가 있어서
    # (예: batang.ttc -> Batang.ttc), 정확한 이름으로 못 찾으면 대소문자 구분 없이 한 번 더 찾아봄
    for base in (os.path.join(_BASE_DIR, "assets"), _BASE_DIR):
        candidate = os.path.join(base, filename)
        if os.path.exists(candidate):
            return candidate
        ci_match = _find_file_ci(base, filename)
        if ci_match:
            return ci_match
    return os.path.join(_BASE_DIR, filename)


CERT_FONT_PATH = _find_asset("batang.ttc")
CERT_FONT_INDEX = 1  # batang.ttc 안에는 바탕/바탕체/궁서/궁서체 4종류가 들어있는데, 서식 PDF들이 실제로
# 쓰는 폰트가 '바탕체(BatangChe)'라서(pdfplumber로 서식 글자를 직접 확인함) 그 인덱스를 지정해서 씀
CERT_TEMPLATE_PATH = _find_asset("cert_template.pdf")  # 브랜드별 서식이 없을 때 쓰는 기본값(트레몰로 서식)

# 담당자에게 보내는 이메일 맨 아래에 붙는 명함/서명. 바뀌면 여기만 수정하면 됨.
_CERT_EMAIL_SIGNATURE = (
    "황유진 대표\n"
    "다온인슈 기업보험 전문컨설팅 / 대표이사\n\n"
    "우)06620 서울특별시 서초구 강남대로 375 서초현대타워 10층\n"
    "Tel 02-6953-8129 Mobile 010-9810-5503\n"
    "Fax 0303-0950-5503\n"
    "Email yujin.hwang@daonins.co.kr"
)

# 담당자에게 가입증명서를 이메일로 보낼 때 본문 문구. {store_name} 자리에 매장명이 자동으로 들어감.
# 문구를 바꾸고 싶으면 여기만 수정하면 됨.
_CERT_EMAIL_BODY_TEMPLATE = (
    "안녕하십니까.\n\n"
    "요청해주신 '{store_name}' 신규 매장 가입증명서를 첨부와 같이 보내드립니다.\n\n"
    "확인 부탁드리며, 문의사항 있으시면 언제든 연락 주시기 바랍니다.\n\n"
    "감사합니다.\n\n"
    "--\n" + _CERT_EMAIL_SIGNATURE
)

# 같은 담당자 앞으로 신규매장이 여러 곳일 때, 이메일 한 통에 가입증명서를 모두 첨부해서 보낼 때 쓰는 본문.
# {store_list} 자리에 "- 매장명1\n- 매장명2" 형태로 매장 목록이 자동으로 들어감.
_CERT_EMAIL_BODY_TEMPLATE_MULTI = (
    "안녕하십니까.\n\n"
    "요청해주신 아래 신규 매장 가입증명서를 첨부와 같이 보내드립니다.\n\n"
    "{store_list}\n\n"
    "확인 부탁드리며, 문의사항 있으시면 언제든 연락 주시기 바랍니다.\n\n"
    "감사합니다.\n\n"
    "--\n" + _CERT_EMAIL_SIGNATURE
)
CERT_PAGE_W, CERT_PAGE_H = 595.2, 841.92
CERT_FONT_SCALE = 4  # 텍스트를 이미지로 그릴 때 선명하게 보이도록 확대 비율

# 브랜드별 가입증명서 서식에서 5개 항목(보험기간/점포명/주소/보상한도액/보험료) 줄의 세로 위치(top, pt)와
# '보험료' 값 칸의 오른쪽 한계(premium_x1 - 바로 뒤 '원' 글자를 안 지우면서 값을 덮을 수 있는 최대 x좌표).
# 같은 디자인이라도 브랜드마다 '계약자'/'주소' 줄 길이가 달라 아래 항목들이 조금씩 밀려 있고, '원' 글자
# 시작 위치도 살짝 달라서, 실제 발급된 예시 증명서에서 좌표를 직접 재서 브랜드별로 따로 저장해둠
# (pdfplumber로 측정). premium_x1을 너무 좁게 잡으면 새로 쓰는 금액이 기존 서식 금액보다 짧을 때
# 예전 숫자 일부가 안 가려지고 남아 보이는 문제가 생김.
DEFAULT_CERT_ROWS = {"period": 201.4, "store_name": 295.3, "address": 313.3, "limit": 367.6, "premium": 439.8, "premium_x1": 160}
CERT_ROWS_BY_BRAND = {
    "트레몰로": {"period": 201.4, "store_name": 295.3, "address": 313.3, "limit": 367.6, "premium": 439.8, "premium_x1": 162.7},
    "월메이드": {"period": 183.4, "store_name": 277.1, "address": 300.6, "limit": 360.1, "premium": 432.4, "premium_x1": 168.7},
    "올리비아로렌&오뷔엘알&주얼리": {"period": 183.4, "store_name": 277.1, "address": 295.3, "limit": 349.6, "premium": 421.8, "premium_x1": 168.7},
}

# 브랜드별 정산 통합파일을 서버에 계속 보관/갱신하는 폴더
MASTERS_DIR = os.path.join(PERSIST_DIR, "masters")
# 브랜드별 증권번호를 저장해두는 파일 (파일명에서 못 찾을 때 사용)
POLICY_NUMBERS_FILE = os.path.join(PERSIST_DIR, "policy_numbers.json")
# 정산양식을 보내는 담당자 이름 -> 이메일 주소. 신규매장 가입증명서를 그 담당자에게
# 자동으로 이메일로 보내줄 때 씀(엑셀 '접수자' 칸의 이름으로 찾음)
CONTACTS_FILE = os.path.join(PERSIST_DIR, "contacts.json")
# 메일 알림에서 무시할 보낸사람 키워드(발신 이메일 주소/이름에 포함되면 매칭) 목록.
# 듀오링고 등 광고성 메일이 여기 걸리면 '새 메일 도착' 알림 자체를 안 보냄
# (2026-09-29 유진님 요청). /mutesender, /unmutesender, /mutedsenders 로 관리함.
MUTED_SENDERS_FILE = os.path.join(PERSIST_DIR, "muted_senders.json")
_DEFAULT_MUTED_SENDERS = ["duolingo"]
# 담당자에게 이메일로 보내기 전에 유진님 확인을 거치는 가입증명서들을 보관하는 곳.
# "보내기"를 누르기 전까지는 실제로 발송되지 않고 여기에 대기함(재배포/재시작해도
# 안 없어지도록 디스크에 저장).
PENDING_CERTS_DIR = os.path.join(PERSIST_DIR, "pending_certs")
PENDING_CERTS_INDEX_FILE = os.path.join(PERSIST_DIR, "pending_certs.json")
# 브랜드별로 파일을 보내드릴 때 쓸 표시용 파일명(엑셀 A1 셀 이름과 다르게 쓰고 싶을 때)
BRAND_NAMES_FILE = os.path.join(PERSIST_DIR, "brand_names.json")
# 담당자마다 A1 셀에 브랜드명을 다르게 적어 보내는 경우, 같은 브랜드로 취급하도록 이름을 통일시켜주는 매핑
BRAND_ALIASES_FILE = os.path.join(PERSIST_DIR, "brand_aliases.json")
# 브랜드별 가입증명서 '원본 서식' PDF 폴더. 실제로 발급된 예시 증명서를 그대로 서식으로 써서,
# 증권번호/계약자/보험종목 등은 그 서식에 이미 올바르게 적혀 있는 값을 그대로 두고,
# 매장마다 달라지는 5개 항목(보험기간/점포명/주소/보상한도액/보험료)만 덮어써서 새로 만듦.
# assets/cert_templates 에 미리 등록해둔 브랜드(트레몰로/월메이드/올리비아로렌&오뷔엘알&주얼리)는
# 코드와 함께 배포되어 재배포해도 사라지지 않음. 새 브랜드는 /setcerttemplate 로 직접 추가 가능(이땐
# 서버 디스크에 저장되어 재배포 시 다시 등록해야 함). /setcerttemplate로 등록한 서식은
# PERSIST_DIR(영구 디스크)에 저장해서 재배포해도 남아있게 함.
CERT_TEMPLATES_DIR = os.path.join(PERSIST_DIR, "cert_templates")
CERT_BUNDLED_TEMPLATES_DIR = os.path.join(_BASE_DIR, "assets", "cert_templates")

SYSTEM_PROMPT = (
    "당신은 황유진 대표님의 개인 AI 비서 '김대리'입니다. 대표님은 다온인슈 기업보험 전문컨설팅의 "
    "대표이사이고 기업보험 전문 컨설턴트입니다(세정 등 법인 고객의 신규/폐점 매장 정산, 가입증명서 "
    "발급 업무를 하십니다). 업무(보험, 정산, 문서, 메일)와 일상 이야기(고민, 건강, 취미, 잡담) 모두 "
    "편하고 자연스러운 한국어 말투로 대화하세요. 딱딱한 보고서체보다는 믿을 수 있는 동료처럼 따뜻하게 "
    "말하되, 존중하는 말투(~요체)를 유지하세요. 질문이 단순하면 짧게, 설명이 필요하거나 고민 상담처럼 "
    "깊이 있는 대화가 필요하면 충분히 자세하게 답하세요. 모르는 것은 아는 척하지 말고 솔직히 말하세요. "
    "날씨, 최신 뉴스, 맛집, 가격 등 최신 정보가 필요한 질문은 웹 검색 도구를 적극 활용하세요. "
    "법률/세무/투자 판단은 확정적으로 단정하지 말고 판단에 필요한 사실과 전문가 확인이 필요함을 함께 알려주세요."
)

CHAT_TOOLS = [
    {"type": "web_search_20250305", "name": "web_search", "max_uses": 3}
]

# 일반 대화 기억: 예전엔 20개까지만 기억하고 재배포하면 사라졌음(2026-11-06, 일상+업무 대화를
# 하고 싶다는 유진님 요청으로 늘림). 파일로 저장해서 재시작해도 이어지도록 함. 단, Render에서
# 영구 디스크(Persistent Disk)를 PERSIST_DIR로 연결해두지 않았다면 재배포 시 파일이 같이 사라짐.
MAX_HISTORY = 80
CHAT_HISTORY_FILE = os.path.join(PERSIST_DIR, "chat_history.json")


def _load_conversations() -> dict:
    try:
        with open(CHAT_HISTORY_FILE, "r", encoding="utf-8") as f:
            raw = json.load(f)
        return {int(k): v for k, v in raw.items() if isinstance(v, list)}
    except (FileNotFoundError, json.JSONDecodeError, ValueError, OSError):
        return {}


def _save_conversations() -> None:
    try:
        with open(CHAT_HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump({str(k): v[-MAX_HISTORY:] for k, v in conversations.items()}, f, ensure_ascii=False)
    except OSError:
        logger.exception("대화 기록 저장 중 오류")


conversations: dict[int, list[dict]] = _load_conversations()

# 메일 감시용: 마지막으로 확인한 IMAP UID (재시작하면 초기화됨)
last_uid_seen: int | None = None


def is_allowed(update: Update) -> bool:
    user = update.effective_user
    return user is not None and user.id == ALLOWED_USER_ID


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        await update.message.reply_text("죄송해요, 이 봇은 개인용이라 사용할 수 없어요.")
        return
    await update.message.reply_text(
        "안녕하세요! AI 비서예요. 👋\n\n"
        "그냥 편하게 메시지를 보내면 대화할 수 있어요.\n\n"
        "사용 가능한 명령어:\n"
        "/remind <분> <내용> - 알림 예약 (예: /remind 30 회의 참석)\n"
        "/mail <받는사람이메일> - 메일 작성해서 보내기 (파일 첨부 가능)\n"
        "/route - 두 지점 사이 길찾기 (자동차/대중교통/도보/자전거 비교)\n"
        "/sethome - 집 위치 등록\n"
        "/setwork - 회사 위치 등록\n"
        "/towork - 집→회사 길찾기 (등록 필요)\n"
        "/tohome - 회사→집 길찾기 (등록 필요)\n"
        "/reset - 지금까지의 대화 기억 지우기\n"
        "/setcontact <담당자 이름> <이메일> - ㈜세정 담당자 이메일 등록\n"
        "/contacts - 등록된 담당자 연락처 목록 보기\n"
        "/pending - 대기 중인(아직 안 보낸) 메일 목록 보기\n"
        "/waiting - 담당자 답을 기다리는 세정 건 보기, /cancelwait 매장명 - 그 대기 취소\n\n"
        "길찾기는 명령어 없이 그냥 '강남역까지 얼마나 걸려?'처럼 물어보셔도 알아들어요.\n\n"
        "새 이메일이 오면 자동으로 요약해서 알려드려요. 📬 대화 중에 '내 메일에서 ~ 찾아줘'처럼 물어보셔도 돼요.\n"
        "㈜세정 변경신청서(엑셀)가 오면 처브 담당자들에게 보낼지 물어보고, 처브 보험가입증명서/배서증명서(PDF)가 "
        "오면 통합관리리스트에 적고 드라이브 정산관리 폴더에 저장한 뒤 세정 담당자에게 보낼지 물어봐요.\n"
        "날씨, 최신 뉴스, 맛집 등도 그냥 물어보시면 웹 검색해서 답해드려요.\n"
        "📎(첨부) 버튼으로 '위치'를 공유해주시면, 그 위치 기준으로 근처 맛집도 찾아드려요.\n\n"
        "변경신청서 엑셀이나 처브 증명서 PDF는 텔레그램으로 직접 올려주셔도 똑같이 처리해요. 📄"
    )


async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    conversations.pop(update.effective_chat.id, None)
    _save_conversations()
    await update.message.reply_text("대화 기록을 초기화했어요.")


async def remind(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return

    try:
        minutes = float(context.args[0])
        text = " ".join(context.args[1:]).strip()
        if not text:
            raise ValueError
    except (IndexError, ValueError):
        await update.message.reply_text(
            "사용법: /remind <분> <내용>\n예: /remind 30 회의 참석"
        )
        return

    chat_id = update.effective_chat.id
    context.job_queue.run_once(
        send_reminder, when=minutes * 60, chat_id=chat_id, data=text
    )
    await update.message.reply_text(f"⏰ {minutes}분 후에 알려드릴게요: {text}")


async def send_reminder(context: ContextTypes.DEFAULT_TYPE) -> None:
    job = context.job
    await context.bot.send_message(chat_id=job.chat_id, text=f"⏰ 알림: {job.data}")


async def _reverse_geocode(lat: float, lon: float) -> str:
    url = "https://nominatim.openstreetmap.org/reverse"
    params = {"lat": lat, "lon": lon, "format": "json", "accept-language": "ko"}
    headers = {"User-Agent": "yujin-ai-bot/1.0"}
    async with httpx.AsyncClient(timeout=10) as http_client:
        resp = await http_client.get(url, params=params, headers=headers)
        resp.raise_for_status()
        data = resp.json()
    return data.get("display_name") or f"위도 {lat}, 경도 {lon}"


def _load_places() -> dict:
    try:
        with open(PLACES_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _save_places(places: dict) -> None:
    with open(PLACES_FILE, "w", encoding="utf-8") as f:
        json.dump(places, f, ensure_ascii=False, indent=2)


async def _kakao_address_search(query: str) -> dict | None:
    """정확한 주소 문자열을 좌표로 변환 (지번/도로명 주소 전용, 상세 동/호수는 제외 권장)"""
    if not KAKAO_REST_API_KEY:
        return None
    url = "https://dapi.kakao.com/v2/local/search/address.json"
    headers = {"Authorization": f"KakaoAK {KAKAO_REST_API_KEY}"}
    params = {"query": query}
    async with httpx.AsyncClient(timeout=10) as http_client:
        resp = await http_client.get(url, headers=headers, params=params)
        resp.raise_for_status()
        data = resp.json()
    docs = data.get("documents") or []
    if not docs:
        return None
    d = docs[0]
    return {
        "name": d.get("address_name"),
        "address": d.get("address_name"),
        "x": d["x"],
        "y": d["y"],
    }


async def _kakao_geocode_keyword(query: str) -> dict | None:
    """주소나 장소명을 좌표로 변환 (카카오맵 키워드 검색)"""
    if not KAKAO_REST_API_KEY:
        return None
    url = "https://dapi.kakao.com/v2/local/search/keyword.json"
    headers = {"Authorization": f"KakaoAK {KAKAO_REST_API_KEY}"}
    params = {"query": query, "size": 1}
    async with httpx.AsyncClient(timeout=10) as http_client:
        resp = await http_client.get(url, headers=headers, params=params)
        resp.raise_for_status()
        data = resp.json()
    docs = data.get("documents") or []
    if not docs:
        return None
    d = docs[0]
    return {
        "name": d.get("place_name") or d.get("road_address_name") or d.get("address_name"),
        "address": d.get("road_address_name") or d.get("address_name"),
        "x": d["x"],
        "y": d["y"],
    }


async def _extract_place_name_llm(text: str) -> str | None:
    """문장에서 장소명/주소만 뽑아냄 ('현재 위치는 강남역이야' -> '강남역')"""
    system = (
        "사용자 문장에서 언급된 장소명이나 주소만 정확히 추출해서 그 텍스트만 답하세요. "
        "다른 설명, 문장부호, 따옴표 없이 장소명만 출력하세요. "
        "장소를 찾을 수 없으면 정확히 NONE 이라고만 답하세요."
    )
    try:
        response = client.messages.create(
            model=MODEL_NAME,
            max_tokens=30,
            system=system,
            messages=[{"role": "user", "content": text}],
        )
        raw = "".join(b.text for b in response.content if b.type == "text").strip()
        if not raw or raw.upper() == "NONE":
            return None
        return raw
    except Exception:
        logger.exception("장소명 추출 중 오류")
        return None


async def _resolve_place(text: str, places: dict) -> dict | None:
    """사용자가 입력한 텍스트를 위치 정보로 변환. '집'/'회사'는 저장된 위치 사용"""
    t = text.strip()
    if t in ("집", "우리집", "집으로", "자택"):
        return places.get("home")
    if t in ("회사", "직장", "사무실", "회사로"):
        return places.get("work")
    try:
        # 정확한 주소(도로명/지번)는 주소 검색 API가 훨씬 정확함. 안 되면 장소명 검색으로 재시도
        place = await _kakao_address_search(t)
        if place:
            return place
        place = await _kakao_geocode_keyword(t)
        if place:
            return place
        # 문장 형태로 입력한 경우("현재 위치는 강남역이야") 장소명만 뽑아서 재시도
        extracted = await _extract_place_name_llm(t)
        if extracted and extracted != t:
            place = await _kakao_address_search(extracted)
            if place:
                return place
            return await _kakao_geocode_keyword(extracted)
        return None
    except Exception:
        logger.exception("장소 검색 중 오류")
        return None


async def _kakao_driving_route(sx: str, sy: str, ex: str, ey: str) -> dict | None:
    if not KAKAO_REST_API_KEY:
        return None
    url = "https://apis-navi.kakaomobility.com/v1/directions"
    headers = {"Authorization": f"KakaoAK {KAKAO_REST_API_KEY}"}
    params = {
        "origin": f"{sx},{sy}",
        "destination": f"{ex},{ey}",
        "priority": "RECOMMEND",
    }
    try:
        async with httpx.AsyncClient(timeout=10) as http_client:
            resp = await http_client.get(url, headers=headers, params=params)
            resp.raise_for_status()
            data = resp.json()
        routes = data.get("routes") or []
        if not routes or routes[0].get("result_code") != 0:
            return None
        return routes[0].get("summary")
    except Exception:
        logger.exception("자동차 길찾기 조회 중 오류")
        return None


async def _kakao_transit_route(sx: str, sy: str, ex: str, ey: str) -> dict | None:
    if not KAKAO_REST_API_KEY:
        return None
    url = "https://dapi.kakao.com/v2/routing/publictraffic"
    headers = {"Authorization": f"KakaoAK {KAKAO_REST_API_KEY}"}
    params = {"start_x": sx, "start_y": sy, "end_x": ex, "end_y": ey}
    try:
        async with httpx.AsyncClient(timeout=10) as http_client:
            resp = await http_client.get(url, headers=headers, params=params)
            resp.raise_for_status()
            return resp.json()
    except Exception:
        logger.exception("대중교통 길찾기 조회 중 오류")
        return None


async def _kakao_walk_or_bicycle_route(sx: str, sy: str, ex: str, ey: str, mode: str) -> dict | None:
    if not KAKAO_REST_API_KEY:
        return None
    path = "walk" if mode == "walk" else "bicycle"
    url = f"https://dapi.kakao.com/v2/routing/{path}"
    headers = {"Authorization": f"KakaoAK {KAKAO_REST_API_KEY}"}
    params = {"start_x": sx, "start_y": sy, "end_x": ex, "end_y": ey}
    try:
        async with httpx.AsyncClient(timeout=10) as http_client:
            resp = await http_client.get(url, headers=headers, params=params)
            resp.raise_for_status()
            return resp.json()
    except Exception:
        logger.exception("도보/자전거 길찾기 조회 중 오류")
        return None


def _format_driving(summary: dict | None) -> str:
    if not summary:
        return "🚗 자동차: 정보 없음 (승인 대기 중이거나 경로를 찾지 못했어요)"
    minutes = round(summary.get("duration", 0) / 60)
    km = round(summary.get("distance", 0) / 1000, 1)
    text = f"🚗 자동차: 약 {minutes}분 ({km}km)"
    toll = (summary.get("fare") or {}).get("toll")
    if toll:
        text += f", 통행료 {toll}원"
    return text


def _format_transit(data: dict | None) -> str:
    if not data or data.get("status") != "OK" or not data.get("routes"):
        return "🚌 대중교통: 정보 없음"
    route = data["routes"][0]
    props = route["properties"]
    minutes = round(props.get("totalTime", 0) / 60)
    line = f"🚌 대중교통: 약 {minutes}분"
    if props.get("transfers") is not None:
        line += f", 환승 {props['transfers']}회"
    fare = (props.get("fare") or {}).get("value")
    if fare:
        line += f", 요금 {fare}원"

    vehicle_names = []
    for step in route.get("steps", []):
        sp = step.get("properties", {})
        if sp.get("type") in ("BUS", "SUBWAY"):
            kind = "버스" if sp["type"] == "BUS" else "지하철"
            for v in sp.get("vehicles", []):
                label = f"{kind} {v.get('name', '')}".strip()
                if label not in vehicle_names:
                    vehicle_names.append(label)
    if vehicle_names:
        line += "\n   이용 노선: " + ", ".join(vehicle_names)
    return line


def _format_walk_or_bicycle(data: dict | None, label: str, emoji: str) -> str:
    if not data or data.get("status") != "OK":
        return f"{emoji} {label}: 정보 없음"
    props = data["route"]["properties"]
    minutes = round(props.get("totalTime", 0) / 60)
    km = round(props.get("totalDistance", 0) / 1000, 1)
    return f"{emoji} {label}: 약 {minutes}분 ({km}km)"


async def _reply_route_comparison(update: Update, start: dict, end: dict) -> None:
    start_label = start.get("name") or start.get("address") or "출발지"
    end_label = end.get("name") or end.get("address") or "도착지"

    if not KAKAO_REST_API_KEY:
        await update.message.reply_text("길찾기 기능이 아직 설정되지 않았어요.")
        return

    await update.message.reply_text(f"🔎 '{start_label}' → '{end_label}' 경로를 찾고 있어요...")

    sx, sy, ex, ey = start["x"], start["y"], end["x"], end["y"]

    driving, transit, walk, bicycle = await asyncio.gather(
        _kakao_driving_route(sx, sy, ex, ey),
        _kakao_transit_route(sx, sy, ex, ey),
        _kakao_walk_or_bicycle_route(sx, sy, ex, ey, "walk"),
        _kakao_walk_or_bicycle_route(sx, sy, ex, ey, "bicycle"),
    )

    lines = [
        f"📍 {start_label} → {end_label}",
        "",
        _format_driving(driving),
        _format_transit(transit),
        _format_walk_or_bicycle(walk, "도보", "🚶"),
        _format_walk_or_bicycle(bicycle, "자전거", "🚴"),
    ]
    await update.message.reply_text("\n".join(lines))


async def _advance_route_query(update: Update, context: ContextTypes.DEFAULT_TYPE, place: dict) -> None:
    """진행 중인 route_query 상태에 place를 반영하고 다음 단계로 진행 (텍스트 입력/위치 공유 공통 처리)"""
    route_query = context.user_data["route_query"]
    state = route_query["state"]
    label = place.get("name") or place.get("address") or "위치"

    if state == "await_current_location":
        dest_text = route_query.pop("pending_destination_text", None)
        if dest_text:
            places = _load_places()
            dest_place = await _resolve_place(dest_text, places)
            context.user_data.pop("route_query", None)
            if not dest_place:
                await update.message.reply_text(
                    f"현재 위치는 확인했는데, '{dest_text}'는 찾지 못했어요. 목적지를 다시 알려주세요."
                )
                context.user_data["route_query"] = {"state": "to", "from": place}
                return
            await _reply_route_comparison(update, place, dest_place)
        else:
            route_query["from"] = place
            route_query["state"] = "to"
            await update.message.reply_text(f"현재 위치: {label}\n\n도착지는 어디인가요?")
        return

    if state == "from":
        route_query["from"] = place
        route_query["state"] = "to"
        await update.message.reply_text(f"출발지: {label}\n\n도착지는 어디인가요?")
        return

    # state == "to"
    context.user_data.pop("route_query", None)
    await _reply_route_comparison(update, route_query["from"], place)


ROUTE_KEYWORDS = (
    "길찾", "가는길", "가는법", "가는방법", "어떻게가", "어떻게가나",
    "얼마나걸려", "얼마나걸리", "출근길", "퇴근길", "이동시간", "차로가", "대중교통으로",
    "버스로가", "도보로가", "걸어서가", "경로",
)


def _has_route_keyword(text: str) -> bool:
    normalized = text.replace(" ", "")
    return any(kw in normalized for kw in ROUTE_KEYWORDS)


async def _extract_route_intent(text: str) -> dict | None:
    """자유롭게 쓴 문장에서 길찾기 요청인지, 출발지/목적지가 무엇인지 뽑아냄"""
    system = (
        "사용자 문장이 길찾기(이동 경로) 요청인지 판단하고, 아래 JSON 형식으로만 답하세요. "
        "설명이나 다른 텍스트는 절대 포함하지 마세요.\n"
        '{"is_route": true 또는 false, '
        '"origin": "current"(현재 위치를 말하거나 출발지가 명시되지 않은 경우) 또는 "home" 또는 "work" 또는 구체적 장소명 문자열 또는 null, '
        '"destination": 구체적 장소명 문자열 또는 null(불명확한 경우)}\n'
        "길찾기 요청이 아니면 is_route를 false로 하세요."
    )
    try:
        response = client.messages.create(
            model=MODEL_NAME,
            max_tokens=200,
            system=system,
            messages=[{"role": "user", "content": text}],
        )
        raw = "".join(b.text for b in response.content if b.type == "text").strip()
        raw = raw.strip("`")
        if raw.lower().startswith("json"):
            raw = raw[4:].strip()
        data = json.loads(raw)
        if not isinstance(data, dict):
            return None
        return data
    except Exception:
        logger.exception("길찾기 의도 분석 중 오류")
        return None


async def _handle_natural_route_request(update: Update, context: ContextTypes.DEFAULT_TYPE, user_text: str) -> bool:
    """자연어 길찾기 요청을 감지해서 처리. 처리했으면 True 반환"""
    if not KAKAO_REST_API_KEY:
        return False
    if not _has_route_keyword(user_text):
        return False

    intent = await _extract_route_intent(user_text)
    if not intent or not intent.get("is_route"):
        return False

    places = _load_places()
    origin = intent.get("origin")
    destination_text = intent.get("destination")

    # 출발지가 "현재 위치"를 가리키는 경우: 최근 공유된 위치가 있으면 재사용, 없으면 위치 공유 요청
    if origin in (None, "current"):
        recent = context.user_data.get("location")
        if isinstance(recent, dict) and recent.get("x") and recent.get("y"):
            from_place = {
                "name": recent.get("address"),
                "address": recent.get("address"),
                "x": recent["x"],
                "y": recent["y"],
            }
        else:
            context.user_data["route_query"] = {
                "state": "await_current_location",
                "pending_destination_text": destination_text,
            }
            extra = f" (목적지: {destination_text})" if destination_text else ""
            await update.message.reply_text(
                f"현재 위치를 확인할게요.{extra}\n📎(첨부) 버튼으로 위치를 공유해주세요."
            )
            return True
    else:
        from_place = await _resolve_place(origin, places)
        if not from_place:
            context.user_data["route_query"] = {"state": "from"}
            await update.message.reply_text(
                f"'{origin}' 위치를 찾지 못했어요. 출발지를 다시 알려주세요."
            )
            return True

    if not destination_text:
        context.user_data["route_query"] = {"state": "to", "from": from_place}
        await update.message.reply_text("도착지는 어디인가요?")
        return True

    to_place = await _resolve_place(destination_text, places)
    if not to_place:
        context.user_data["route_query"] = {"state": "to", "from": from_place}
        await update.message.reply_text(
            f"'{destination_text}'를 찾지 못했어요. 도착지를 다시 알려주세요."
        )
        return True

    await _reply_route_comparison(update, from_place, to_place)
    return True


async def route_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    if not KAKAO_REST_API_KEY:
        await update.message.reply_text("길찾기 기능이 아직 설정되지 않았어요.")
        return
    context.user_data["route_query"] = {"state": "from"}
    await update.message.reply_text(
        "출발지를 알려주세요.\n"
        "주소나 장소명을 입력하거나, 📎(첨부)로 위치를 공유하거나, '집' 또는 '회사'라고 입력해주세요."
    )


async def set_home(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    context.user_data["saving_place"] = "home"
    await update.message.reply_text("🏠 집 위치를 설정할게요. 위치를 공유하거나 주소/건물명을 입력해주세요.")


async def set_work(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    context.user_data["saving_place"] = "work"
    await update.message.reply_text("🏢 회사 위치를 설정할게요. 위치를 공유하거나 주소/건물명을 입력해주세요.")


async def commute_to_work(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    places = _load_places()
    if not places.get("home") or not places.get("work"):
        await update.message.reply_text(
            "먼저 /sethome 과 /setwork 으로 집과 회사 위치를 등록해주세요."
        )
        return
    await _reply_route_comparison(update, places["home"], places["work"])


async def commute_to_home(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    places = _load_places()
    if not places.get("home") or not places.get("work"):
        await update.message.reply_text(
            "먼저 /sethome 과 /setwork 으로 집과 회사 위치를 등록해주세요."
        )
        return
    await _reply_route_comparison(update, places["work"], places["home"])


async def handle_location(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return

    loc = update.message.location
    try:
        address = await _reverse_geocode(loc.latitude, loc.longitude)
    except Exception:
        logger.exception("위치 확인 중 오류")
        address = f"위도 {loc.latitude}, 경도 {loc.longitude}"

    place = {"name": address, "address": address, "x": str(loc.longitude), "y": str(loc.latitude)}

    # 집/회사 위치 저장 중이면 그쪽으로 처리
    saving = context.user_data.get("saving_place")
    if saving:
        places = _load_places()
        places[saving] = place
        _save_places(places)
        context.user_data.pop("saving_place", None)
        label = "집" if saving == "home" else "회사"
        await update.message.reply_text(f"✅ {label} 위치를 저장했어요: {address}")
        return

    # 길찾기 진행 중이면 그쪽으로 처리
    if context.user_data.get("route_query"):
        await _advance_route_query(update, context, place)
        return

    context.user_data["location"] = {"address": address, "x": str(loc.longitude), "y": str(loc.latitude)}
    await update.message.reply_text(
        f"📍 위치를 받았어요: {address}\n"
        "이제 '근처 맛집 추천해줘'처럼 물어보시면 이 위치를 기준으로 찾아드릴게요."
    )


# ===================== 가입증명서 자동 생성 =====================
# 정산양식 엑셀(신규매장/폐점매장 시트)을 브랜드별 통합파일과 비교해 새로 추가된 매장을
# 찾아, 예시 PDF와 같은 양식의 가입증명서를 자동으로 만들어주는 기능. 원본 서식(배경/직인
# 이미지)은 그대로 두고 매장별로 달라지는 값(점포명/주소/보험기간/보험료 등)만 흰색으로
# 덮은 뒤 새로 그려넣는 방식.

# 브랜드에 등록된 증권번호가 없을 때 예시 양식과 동일하게 쓸 기본 증권번호
DEFAULT_POLICY_NO = "82509565736000"


def _cert_bundled_template_candidates(brand: str) -> list:
    """코드와 함께 배포되어 재배포해도 안 사라지는 브랜드별 서식을 찾을 수 있는 후보 (디렉터리, 파일명) 목록.
    GitHub에 assets 폴더를 따로 안 만들고 bot.py 등과 같은 위치(루트)에 그냥 올리는 경우가 많아서,
    assets/cert_templates 안과 루트(cert_template_브랜드명.pdf) 둘 다 확인함."""
    return [
        (CERT_BUNDLED_TEMPLATES_DIR, f"{brand}.pdf"),
        (_BASE_DIR, f"cert_template_{brand}.pdf"),
    ]


def _resolve_cert_bundled_template(brand: str) -> str | None:
    for directory, filename in _cert_bundled_template_candidates(brand):
        exact = os.path.join(directory, filename)
        if os.path.exists(exact):
            return exact
        # GitHub 웹에서 파일명을 바꿀 때 확장자가 .PDF로 바뀌어 있는 경우가 있어 대소문자 구분 없이도 확인
        found = _find_file_ci(directory, filename)
        if found:
            return found
    return None


def _cert_template_path_for_brand(brand: str) -> str:
    """브랜드별 가입증명서 원본 서식 경로. 순서: 사용자가 /setcerttemplate로 직접 등록한 것(서버 디스크,
    재배포시 사라짐) -> 코드와 함께 배포되는 것(재배포에도 안전) -> 아무것도 없으면 기본(트레몰로) 서식."""
    custom_path = os.path.join(CERT_TEMPLATES_DIR, f"{brand}.pdf")
    if os.path.exists(custom_path):
        return custom_path
    bundled = _resolve_cert_bundled_template(brand)
    if bundled:
        return bundled
    return CERT_TEMPLATE_PATH


def _cert_rows_for_brand(brand: str) -> dict:
    return CERT_ROWS_BY_BRAND.get(brand, DEFAULT_CERT_ROWS)


def _has_cert_template(brand: str) -> bool:
    """이 브랜드 전용 가입증명서 서식이 등록되어 있는지. 없으면 트레몰로 서식을 임시로 써서
    계약자/증권번호 등이 실제와 다르게 나올 수 있어 사용자에게 알려줘야 함."""
    if os.path.exists(os.path.join(CERT_TEMPLATES_DIR, f"{brand}.pdf")):
        return True
    return _resolve_cert_bundled_template(brand) is not None


# 가입증명서용 폰트(NotoSerifKR)에 글자가 비어있는(빈 도형) 특수문자들을 폰트가
# 지원하는 형태로 미리 바꿔줌. 안 하면 그 글자만 화면/PDF에서 통째로 빈 칸으로 사라짐.
# (예: '계약자' 줄에 '㈜세정'처럼 회사명 앞에 동그라미 표시가 들어가는 경우)
FONT_CHAR_FALLBACKS = {
    "㈜": "(주)",
}


def _sanitize_for_font(text: str) -> str:
    for bad, good in FONT_CHAR_FALLBACKS.items():
        text = text.replace(bad, good)
    return text


def _render_text_image(text: str, font_size_pt: float):
    text = _sanitize_for_font(text)
    px_size = int(font_size_pt * CERT_FONT_SCALE)
    font = ImageFont.truetype(CERT_FONT_PATH, px_size, index=CERT_FONT_INDEX)
    ascent, descent = font.getmetrics()
    bbox = font.getbbox(text)
    text_w = bbox[2] - bbox[0]
    img_h = ascent + descent
    img = Image.new("RGBA", (max(text_w, 1) + 4, img_h), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)
    draw.text((-bbox[0] + 2, 0), text, font=font, fill=(0, 0, 0, 255))
    return img, img.width / CERT_FONT_SCALE, img.height / CERT_FONT_SCALE, descent / CERT_FONT_SCALE


POLICY_END_MMDD = (11, 12)  # ㈜세정 재산종합보험 증권 만기(갱신)일. 증권 보험기간은 11.12 00:01 ~ 11.12 00:01


def _cert_end_time(data: dict) -> str:
    """가입증명서 종기 시각. data에 end_time이 있으면 그대로, 없으면 종기일로 판단."""
    if data.get("end_time"):
        return data["end_time"]
    end_d = _parse_date_val(data.get("end_date"))
    if end_d and (end_d.month, end_d.day) == POLICY_END_MMDD:
        return "00:01"
    return "24:00"


def _build_certificate_pdf(data: dict, brand: str) -> bytes:
    """브랜드별 실제 발급 예시 증명서를 원본 서식으로 그대로 쓰고, 매장마다 달라지는 5개 항목만
    흰색으로 덮은 뒤 새로 그려넣음. 증권번호/계약자/보험종목 등 나머지는 서식에 이미 올바르게
    적혀 있으므로 건드리지 않음(브랜드마다 회사명 문구가 달라도 신경 쓸 필요 없음).
    data keys: start_date, end_date, store_name, address, stock_amt(int), facility_amt(int), premium(int)"""
    rows = _cert_rows_for_brand(brand)
    buf = io.BytesIO()
    c = pdf_canvas.Canvas(buf, pagesize=(CERT_PAGE_W, CERT_PAGE_H))

    def y_of(bottom):
        return CERT_PAGE_H - bottom

    def cover(x0, top, bottom, x1=545):
        c.setFillColorRGB(1, 1, 1)
        c.rect(x0 - 2, CERT_PAGE_H - bottom - 2, (x1 - x0) + 4, (bottom - top) + 4, fill=1, stroke=0)

    def draw_text(x0, bottom, s, size=12):
        # 이미지 자체가 이미 '글꼴 상 맨 아래(descent)'까지 포함해서 그려지므로, 이미지의 맨 아래를
        # 그 줄의 바텀 위치(baseline_y)에 바로 맞추면 됨. 예전에는 여기서 descent_pt를 한 번 더 빼서
        # 삽입한 글자들이 서식 원문 글자보다 항상 살짝(약 2pt) 아래로 처지는 문제가 있었음(바탕체로
        # 바꾸면서 서식과 나란히 놓고 보니 확실히 티가 남 — 사용자 확인 사례).
        img, w_pt, h_pt, descent_pt = _render_text_image(s, size)
        baseline_y = y_of(bottom)
        c.drawImage(ImageReader(img), x0, baseline_y, width=w_pt, height=h_pt, mask="auto")

    def row_bottom(top):
        return top + 12.0

    period_top = rows["period"]
    period_bottom = row_bottom(period_top)
    cover(122.7, period_top, period_bottom, x1=183)
    draw_text(122.7, period_bottom - 0.05, data["start_date"])
    cover(254.7, period_top, period_bottom, x1=315)
    draw_text(254.7, period_bottom - 0.05, data["end_date"])
    # 종기 시각: 브랜드 서식마다 원문 시각이 달라서(월메이드 서식만 '24:00'으로 고정돼 있었음 →
    # 11/12 만기 건이 24:00으로 발급되는 오류, 2026-10-13 충무상설직영특판점 사례) 서식 값을 쓰지 않고
    # 항상 덮어씀. 증권 만기일(11/12)에 끝나면 증권과 동일하게 00:01, 그 전에 끝나는 단기 매장은 24:00.
    cover(320.76, period_top, period_bottom, x1=351)
    draw_text(320.76, period_bottom - 0.05, _cert_end_time(data))

    store_top = rows["store_name"]
    store_bottom = row_bottom(store_top)
    cover(122.7, store_top, store_bottom)
    draw_text(122.7, store_bottom - 0.05, data["store_name"])

    addr_top = rows["address"]
    addr_bottom = row_bottom(addr_top)
    cover(122.7, addr_top, addr_bottom)
    draw_text(122.7, addr_bottom - 0.05, data["address"])

    # 재물부문(재고/시설비품/건물) 항목별로, 정산양식에 금액이 기재된 항목만 골라서 나열함
    # (예: 건물 값이 없으면 '건물' 자체를 언급하지 않고 재고/시설비품만 표시). 셋 다 없으면
    # '보상한도액' 줄 자체를(라벨 포함) 지움.
    limit_top = rows["limit"]
    limit_bottom = row_bottom(limit_top)
    property_items = []
    if data.get("stock_amt", 0):
        property_items.append(f'재고 {data["stock_amt"]:,}원')
    if data.get("facility_amt", 0):
        property_items.append(f'시설/비품 {data["facility_amt"]:,}원')
    if data.get("building_amt", 0):
        property_items.append(f'건물 {data["building_amt"]:,}원')
    if property_items:
        cover(188.7, limit_top, limit_bottom)  # '재고'/'시설/비품' 등 기존 라벨+금액을 통째로 지우고
        draw_text(188.7, limit_bottom - 0.05, ", ".join(property_items))  # 있는 항목만 새로 나열
    else:
        cover(56.7, limit_top, limit_bottom)  # '보상한도액 :' 라벨까지 통째로 지움

    # 영업배상 부문(평수 기준 산정)이 정산양식에 기재 안 돼있으면, 서식에 고정으로 박혀 있는
    # '배상부문 ...' / '구내치료비 ...' 두 줄을 흰색으로 지움(이 두 줄은 모든 브랜드가 항상
    # 재물부문 줄에서 정확히 +18pt, +36pt 아래에 있어서 좌표를 새로 재지 않아도 계산 가능함).
    if not data.get("has_liability", True):
        liability_row1 = limit_top + 18.0
        liability_row2 = limit_top + 36.3
        cover(122.7, liability_row1, liability_row1 + 18.3)
        cover(122.7, liability_row2, liability_row2 + 18.3)

    premium_top = rows["premium"]
    premium_bottom = row_bottom(premium_top)
    cover(122.7, premium_top, premium_bottom, x1=rows.get("premium_x1", 160))
    draw_text(122.7, premium_bottom - 0.05, f'{data["premium"]:,}')

    c.save()
    buf.seek(0)

    overlay_reader = PdfReader(buf)
    template_reader = PdfReader(_cert_template_path_for_brand(brand))
    writer = PdfWriter()
    page = template_reader.pages[0]
    page.merge_page(overlay_reader.pages[0])
    writer.add_page(page)

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


# 브랜드/담당자마다 같은 항목을 다른 이름으로 적는 경우가 있어서, 정규화 후 이 표로 한 번 더 통일함
# (예: 올리비아로렌 취합본은 '재고' 대신 '재고자산'이라고 씀 — 이걸 못 맞추면 재고 금액이 통째로 빠짐)
HEADER_SYNONYMS = {
    "재고자산": "재고",
    "시설": "시설/비품",
}


def _norm_header(s) -> str:
    if s is None:
        return ""
    s = str(s).replace("\n", "").strip()
    s = re.sub(r"\([^)]*\)", "", s)  # 괄호 안 요율/단가 등은 파일마다 달라서 제거
    s = s.replace(" ", "")
    return HEADER_SYNONYMS.get(s, s)


def _find_header_row(ws, search_upto: int = 6) -> int | None:
    """'매장명' 라벨이 있는 행(하위 헤더 행)을 찾음. 파일마다 헤더가 2~3행 또는 3~4행일 수 있어 위치를 직접 스캔."""
    for r in range(1, min(ws.max_row, search_upto) + 1):
        row_vals = [_norm_header(c.value) for c in ws[r]]
        if "매장명" in row_vals:
            return r
    return None


_STORE_CODE_RE = re.compile(r"^[A-Za-z]{2,4}\d{2,5}[A-Za-z]?$")
_KOREAN_RE = re.compile(r"[가-힣]")


def _looks_like_store_code(v) -> bool:
    return bool(_STORE_CODE_RE.match(str(v).strip()))


def _looks_like_korean_text(v) -> bool:
    return bool(_KOREAN_RE.search(str(v)))


def _is_placeholder_code(code) -> bool:
    """'미개설(추후전달)', '코드미개설상태'처럼 매장코드가 아직 배정되지 않아서 담당자가
    임시로 적어넣은 안내 문구인 경우를 가려냄. 이런 값을 진짜 매장코드처럼 dedup 키에
    그대로 쓰면, 나중에 실제 코드가 배정된 뒤 같은 매장이 다시 파일에 실려 오거나(담당자가
    예전 이력을 안 지우고 계속 붙여보내는 경우가 있음) 문구가 미세하게 달라지면(예: '추후
    전달' vs '추후전달') 매번 새 매장으로 잘못 인식해서 중복 등록/가입증명서 재발급으로
    이어짐. 매장코드 칸의 각 줄이 전부 실제 코드 패턴(_STORE_CODE_RE)이어야 '진짜 코드'로
    인정하고, 하나라도 아니면 placeholder로 봄."""
    if not code:
        return True
    segments = [s.strip() for s in str(code).split("\n") if s.strip()]
    if not segments:
        return True
    return not all(_looks_like_store_code(seg) for seg in segments)


def _is_company_refund(vals: dict) -> bool:
    """폐점매장 행의 비고가 '㈜세정으로 환급'(본사 환급)인지 판단. 본사 환급 건은 매장주 서류
    (사업자등록증/신분증/통장사본/개인정보동의서)가 필요 없으므로 폐점서류 저장·알림을 하지 않음
    (유진님 지시, 2026-10-14). 비고가 비어 있거나 '매장주에게 환급'이면 기존대로 서류를 처리함."""
    for v in vals.values():
        if isinstance(v, str):
            t = re.sub(r"\s+", "", unicodedata.normalize("NFC", v))
            if "환급" in t and "세정" in t and "매장주" not in t:
                return True
    return False


def _store_identity_key(vals: dict) -> tuple:
    """매장코드가 없어도 같은 매장인지 알아볼 수 있도록, 매장명+사업자번호(없으면 주소)로
    만드는 보조 식별 키. 공백/줄바꿈 차이는 무시함.
    유니코드 정규화(NFC)도 꼭 해줘야 함 - 같은 한글이라도 담당자 PC/입력기에 따라 완성형(NFC)과
    조합형(NFD)으로 다르게 저장될 수 있어서, 눈으로는 완전히 똑같아 보여도 문자열 비교로는
    다른 값으로 취급돼 중복을 못 잡는 경우가 실제로 있었음(2026-10-02, 월메이드 조치원점 -
    지난달엔 매장코드 WMD0032로 정상 등록됐는데, 이번 달 파일에 매장명/주소가 똑같아 보이는
    채 매장코드만 '코드미개설상태'로 다시 들어와서 신규매장으로 중복 등록되고 가입증명서까지
    다시 만들어짐. 드라이브 폴더명 매칭 때 겪었던 것과 같은 종류의 문제라 같은 방식으로 고침)."""
    def _norm(s) -> str:
        s = unicodedata.normalize("NFC", str(s or ""))
        return re.sub(r"\s+", "", s)
    name = _norm(vals.get("매장명"))
    biz_no = _norm(vals.get("사업자번호"))
    addr = _norm(vals.get("매장주소"))
    return (name, biz_no or addr)


def _build_header_map(ws) -> tuple[dict, int]:
    """헤더 행을 자동으로 찾아 '정규화된 헤더명 -> 0-based 컬럼 인덱스' 매핑과 데이터 시작 행을 반환.
    '매장명' 등 주요 라벨이 있는 행(main_row) 바로 아래 행에 병합된 그룹(재물부문 등)의
    하위 항목명(재고/시설/비품/건물 등)이 있으므로 두 행을 합쳐서 매핑을 만들고,
    실제 데이터는 그 두 행 다음부터 시작함."""
    main_row = _find_header_row(ws)
    if main_row is None:
        return {}, 4

    sub_row = main_row + 1
    row_main = [c.value for c in ws[main_row]]
    row_sub = [c.value for c in ws[sub_row]] if sub_row <= ws.max_row else []
    max_col = max(len(row_main), len(row_sub))

    header_map = {}
    for idx in range(max_col):
        vmain = row_main[idx] if idx < len(row_main) else None
        vsub = row_sub[idx] if idx < len(row_sub) else None
        # 병합된 그룹 칸(예: '재물부문(요율 0.0665%)')은 main_row에, 그 아래 세부 항목명
        # (재고/시설/비품/건물 등)은 sub_row에 있으므로 더 구체적인 sub_row를 우선함
        label = vsub if vsub not in (None, "") else vmain
        norm = _norm_header(label)
        if norm and norm not in header_map:
            header_map[norm] = idx

    data_start = sub_row + 1

    # 일부 브랜드의 '상설' 시트처럼, 헤더 글자는 '매장명'/'매장코드' 순서대로 적혀 있는데
    # 실제 데이터는 그 반대로 입력돼 있는 경우가 실제로 있었음(담당자 쪽 서식 실수). 이걸 그대로
    # 믿으면 가입증명서에 매장코드와 매장명이 뒤바뀌어 나가는 심각한 오류로 이어지므로, 실제
    # 데이터 몇 줄의 값 형태(코드형 vs 한글명)를 보고 뒤바뀐 게 확실하면 자동으로 바로잡음.
    name_idx = header_map.get("매장명")
    code_idx = header_map.get("매장코드")
    if name_idx is not None and code_idx is not None and name_idx != code_idx:
        checked = 0
        swapped_looking = 0
        for r in range(data_start, min(ws.max_row, data_start + 30) + 1):
            name_val = ws.cell(row=r, column=name_idx + 1).value
            code_val = ws.cell(row=r, column=code_idx + 1).value
            if not name_val or not code_val:
                continue
            checked += 1
            if _looks_like_store_code(name_val) and _looks_like_korean_text(code_val):
                swapped_looking += 1
            if checked >= 5:
                break
        if checked and swapped_looking == checked:
            header_map["매장명"], header_map["매장코드"] = code_idx, name_idx

    return header_map, data_start


_SIMPLE_REF_RE = re.compile(r"^=([^!]+)!(\$?[A-Z]+\$?\d+)$")


def _resolve_simple_formula(formula, wb_values):
    """'=최초가입전체리스트!E108'처럼 다른 시트의 셀 하나를 그대로 참조하는 단순 수식이면,
    그 시트의 실제 값을 대신 찾아서 반환함. 매장명 칸이 이런 수식으로 돼 있는 폐점매장 행이
    있는데, 파일이 한 번도 진짜 엑셀에서 재계산/저장된 적이 없으면 캐시된 값이 없어서
    data_only=True로 읽어도 매장명이 빈 값(None)으로 나옴 -> 중복 확인(dedup)이
    그 행을 못 찾아서 같은 매장이 두 번 추가되는 버그로 이어짐. 이걸 막기 위한 보강."""
    if not isinstance(formula, str) or not formula.startswith("="):
        return None
    m = _SIMPLE_REF_RE.match(formula.strip())
    if not m:
        return None
    sheet_name = m.group(1).strip().strip("'")
    cell_ref = m.group(2).replace("$", "")
    if sheet_name not in wb_values.sheetnames:
        return None
    try:
        return wb_values[sheet_name][cell_ref].value
    except Exception:
        return None


def _row_values(ws, header_map: dict, r: int, formula_ws=None, wb_values=None) -> dict:
    row_cells = ws[r]
    formula_cells = formula_ws[r] if formula_ws is not None else None
    vals = {}
    for key, idx in header_map.items():
        v = row_cells[idx].value if idx < len(row_cells) else None
        if v is None and formula_cells is not None and wb_values is not None and idx < len(formula_cells):
            resolved = _resolve_simple_formula(formula_cells[idx].value, wb_values)
            if resolved is not None:
                v = resolved
        vals[key] = v
    return vals


def _extract_data_rows(ws, header_map: dict, min_row: int, formula_ws=None, wb_values=None) -> list:
    out = []
    for r in range(min_row, ws.max_row + 1):
        vals = _row_values(ws, header_map, r, formula_ws, wb_values)
        if not vals.get("매장명"):
            continue
        vals["_excel_row"] = r
        out.append(vals)
    return out


def _val_key(v) -> str:
    """셀 값을 dedup 비교용 문자열로 정규화 (datetime이든 문자열이든 동일하게 비교 가능하게)"""
    if v is None:
        return ""
    if hasattr(v, "isoformat"):
        return v.isoformat()
    return str(v).strip()


def _dedup_key(vals: dict) -> tuple:
    """매장코드가 같아도 보험시작일이 다르면(단기 특판 연장 등) 별개 건으로 취급.
    보험시작일이 없으면 접수일자로 대체."""
    code = vals.get("매장코드")
    period_key = _val_key(vals.get("보험시작일")) or _val_key(vals.get("접수일자"))
    if code:
        return ("code", str(code).strip(), period_key)
    name = str(vals.get("매장명") or "").strip()
    return ("namedate", name, period_key)


def _get_brand_name(wb) -> str | None:
    """첫 시트에서 '*. 브랜드명' 형식의 셀을 찾아 브랜드명을 반환.
    보통 A1 셀에 있지만, 담당자마다 앞에 빈 열을 하나 더 두는 등 서식이 조금씩 달라서
    B1처럼 다른 칸에 적혀 있는 경우도 실제로 있었음. 그래서 A1만 보지 않고 첫 시트의
    맨 위 몇 행/열을 넓게 훑어서 '*.'로 시작하는 칸을 찾음."""
    ws = wb.worksheets[0]
    for row in ws.iter_rows(min_row=1, max_row=3, max_col=6):
        for cell in row:
            v = cell.value
            if not v:
                continue
            m = re.match(r"\*\.\s*(.+)", str(v).strip())
            if m:
                return m.group(1).strip()
    return None


def _find_type_sheets(wb, keyword: str) -> dict:
    """시트 이름에 keyword(예: '신규매장')가 포함된 시트를 '정상'/'상설' 등으로 분류.
    담당자가 개별로 보내는 신청서는 '신규매장(정상)'이 아니라 그냥 '신규매장'처럼
    괄호 구분 없이 오는 경우가 많은데, 이때는 일반(정상) 매장 신청으로 간주함."""
    result = {}
    for name in wb.sheetnames:
        if keyword not in name:
            continue
        if "상설" in name:
            result["상설"] = wb[name]
        else:
            result.setdefault("정상", wb[name])
    return result


def _strip_external_links(wb) -> None:
    """일부 담당자 취합본에는 예전에 다른 파일(예: 카카오톡으로 받은 엑셀)을 참조하던 '외부 워크북
    링크'가 남아있는 경우가 있음(수식이 [1]시트명! 형태). openpyxl로 다시 저장하면 이 외부 링크의
    내부 참조가 깨져서 엑셀에서 '내용에 문제가 있습니다' 복구 경고가 뜨는 원인이 됨. 다행히 같은
    이름의 시트가 통합문서 안에 이미 있는 경우가 대부분이라, [숫자] 표시만 지워서 로컬 시트를
    가리키도록 바꾸고, 더 이상 쓰이지 않는 외부 링크 정의 자체도 제거함."""
    if not getattr(wb, "_external_links", None):
        return
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("=") and re.search(r"\[\d+\]", cell.value):
                    cell.value = re.sub(r"\[\d+\]", "", cell.value)
    wb._external_links = []


SKIP_WRITE_FIELDS = {"월별", "순번"}


def _find_target_row(ws, header_map: dict, min_row: int) -> int:
    """실제 데이터(매장명)가 채워진 가장 마지막 행 바로 다음 행을 찾음. 새 매장은 항상 접수
    순서대로 맨 아래에 쌓여야 하므로, 중간에 (예전에 지워졌거나 서식만 남은) 빈 행이 있어도
    거기 끼워넣지 않고 항상 전체에서 가장 마지막에 채워진 행 다음에 추가함."""
    name_idx = header_map.get("매장명")
    if name_idx is None:
        return ws.max_row + 1
    last_filled = min_row - 1
    for r in range(min_row, ws.max_row + 1):
        row_cells = ws[r]
        if name_idx < len(row_cells) and row_cells[name_idx].value:
            last_filled = r
    return last_filled + 1


def _find_template_row(ws, header_map: dict, min_row: int, target_row: int) -> int:
    """target_row 바로 위에서 실제 데이터(매장명)가 채워진 가장 가까운 행을 찾아 수식 기준 행으로 씀.
    (일부 파일은 빈 줄마다 수식이 끝까지 미리 채워져 있지 않고 중간에 끊겨 있어서,
    바로 위 빈 줄이 아니라 '가장 최근 실제로 채워진 행'을 기준으로 삼아야 수식이 안전하게 이어짐)"""
    name_idx = header_map.get("매장명")
    r = target_row - 1
    while r >= min_row:
        if name_idx is None or ws.cell(row=r, column=name_idx + 1).value:
            return r
        r -= 1
    return target_row


YELLOW_FILL = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
NO_FILL = PatternFill(fill_type=None)

# 통합파일에 새로 써넣는 날짜 칸들. 담당자가 접수일자를 '26.07.16'처럼 문자로 직접 타이핑해서
# 보내는 경우가 있어서, 그대로 쓰면 셀마다 표기 형식이 제각각이 됨. 항상 0000-00-00 형식으로
# 통일해서 보이도록, 실제 날짜값으로 바꿔서 쓰고 셀 서식도 맞춰줌.
DATE_FIELDS = {"접수일자", "보험시작일", "보험종기일", "폐점일"}


def _write_row(ws, header_map: dict, target_row: int, values: dict, template_row: int) -> None:
    """각 컬럼마다 '기준 행(template_row)'의 셀이 수식이면 그 수식을 target_row로 복사(번역)하고,
    수식이 아니면(원본 입력값 칸이면) values의 값을 그대로 채워 넣음. 필드명을 하드코딩해서
    구분하지 않고, 실제로 그 칸이 수식인지 아닌지를 셀 단위로 직접 확인하기 때문에 파일마다
    수식이 있는 칸이 달라도(예: 폐점매장의 매장명이 수식인 경우 등) 안전하게 동작함.
    (음영 처리는 이 함수가 아니라 _recolor_by_period가 접수일자 기준으로 별도로 담당함)"""
    for key, idx in header_map.items():
        if key in SKIP_WRITE_FIELDS:
            continue
        col = idx + 1
        target_cell = ws.cell(row=target_row, column=col)
        if isinstance(target_cell, MergedCell):
            continue

        template_cell = ws.cell(row=template_row, column=col)
        template_val = template_cell.value
        is_formula = isinstance(template_val, str) and template_val.startswith("=")

        if is_formula:
            if target_row != template_row:
                target_cell.value = Translator(
                    template_val, origin=template_cell.coordinate
                ).translate_formula(target_cell.coordinate)
            continue

        if key in values and values[key] not in (None, ""):
            v = values[key]
            if key in DATE_FIELDS:
                parsed = _parse_date_val(v)
                if parsed is not None:
                    v = parsed
                    target_cell.number_format = "yyyy-mm-dd"
            target_cell.value = v

    # 새로 만든 행(기존 서식 범위를 벗어난 경우)엔 순번을 이어서 채워줌
    if target_row != template_row:
        seq_idx = header_map.get("순번")
        if seq_idx is not None:
            prev = ws.cell(row=template_row, column=seq_idx + 1).value
            if isinstance(prev, (int, float)):
                ws.cell(row=target_row, column=seq_idx + 1).value = int(prev) + 1


def _current_period(today: dt.date | None = None) -> tuple[dt.date, dt.date]:
    """'해당월 16일 ~ 다음달 15일'을 하나의 정산 기간으로 봄(예: 7/27이면 7/16~8/15).
    오늘이 1~15일 사이면 아직 저번 기간(전달 16일~이번달 15일)이 진행 중인 걸로 봄."""
    today = today or dt.date.today()
    if today.day >= 16:
        start = dt.date(today.year, today.month, 16)
        if today.month == 12:
            end = dt.date(today.year + 1, 1, 15)
        else:
            end = dt.date(today.year, today.month + 1, 15)
    else:
        end = dt.date(today.year, today.month, 15)
        if today.month == 1:
            start = dt.date(today.year - 1, 12, 16)
        else:
            start = dt.date(today.year, today.month - 1, 16)
    return start, end


def _format_period_label(period_start: dt.date, period_end: dt.date) -> str:
    """정산기간을 통합파일에 실제로 쓰여있는 '26년\\n7월16일~\\n8월15일' 형식의 라벨 문자열로
    만듦. 연도가 바뀌는 기간(예: 12월16일~1월15일)은 두 연도를 다 표시함."""
    start_yy = str(period_start.year)[2:]
    if period_start.year == period_end.year:
        return f"{start_yy}년\n{period_start.month}월{period_start.day}일~\n{period_end.month}월{period_end.day}일"
    end_yy = str(period_end.year)[2:]
    return (
        f"{start_yy}년\n{period_start.month}월{period_start.day}일~\n"
        f"{end_yy}년\n{period_end.month}월{period_end.day}일"
    )


def _col_a_label_blocks(ws) -> list[tuple[str, int, int]]:
    """컬럼 A(월별)에 있는 정산기간 라벨 칸들을 전부 찾아 (라벨텍스트, 시작행, 끝행) 목록으로
    반환함. 여러 행이 병합되어 하나의 라벨을 이루는 경우와, 병합 없이 단독으로 있는 라벨(예:
    그 기간에 항목이 하나뿐인 경우) 둘 다 포함함."""
    merge_end = {mc.min_row: mc.max_row for mc in ws.merged_cells.ranges if mc.min_col == 1 and mc.max_col == 1}
    blocks = []
    for r in range(1, ws.max_row + 1):
        cell = ws.cell(row=r, column=1)
        if isinstance(cell, MergedCell):
            continue
        v = cell.value
        if v in (None, ""):
            continue
        blocks.append((str(v), r, merge_end.get(r, r)))
    return blocks


ROLLOVER_BUFFER_ROWS = 25


def _rollover_period_block(ws, header_map: dict, target_row: int, period_label: str) -> int:
    """새 항목을 target_row에 넣기 전에 호출함. 이 항목이 속한 정산기간(period_label)에 맞는
    라벨 블록으로 target_row를 맞춰줌.
    - 통합파일은 보통 몇 달치 정산기간 블록(라벨 + 여분 행)을 미리 만들어두는 서식이라, 새
      기간으로 넘어갈 때 그 미리 만들어둔 블록까지 건너뛰어 들어가야 하는데, 지금까지는 항상
      '마지막으로 채워진 행 바로 다음'에만 넣어서 새 기간 블록을 건너뛰고 예전 기간 라벨
      아래에 잘못 들어가는 문제가 있었음.
    - 방금 끝난 기간 블록에 남아있던 미사용 여분 행은 삭제하고, 그만큼 아래 블록들을 앞으로
      당겨서 서식이 계속 깔끔하게 이어지도록 함.
    - 컬럼 A에 정산기간 라벨이 아예 없는(이 서식을 안 쓰는) 시트는 손대지 않고 그대로 둠."""
    name_idx = header_map.get("매장명")
    if name_idx is None:
        return target_row

    blocks = _col_a_label_blocks(ws)
    if not blocks:
        return target_row  # 이 서식(기간 블록 미리 만들어두기)을 안 쓰는 시트는 그대로 둠

    # 라벨 글자가 완전히 같아야만 같은 기간으로 보면, 통합파일 쪽 라벨의 종료일이 실제 정산기간과
    # 조금 다르게 적혀 있을 때(예: '10월16일~11월12일' vs 봇이 계산한 '10월16일~11월15일') 이미
    # 있는 블록을 못 찾고 새 라벨을 병합된 칸에 쓰려다 오류가 났음(2026-10-21, 세정 추가가입
    # 건). 그래서 '시작월/시작일'만 뽑아서 비교함.
    def _start_md(label: str):
        m = re.search(r"(\d{1,2})월\s*(\d{1,2})일", str(label))
        return (int(m.group(1)), int(m.group(2))) if m else None

    want_md = _start_md(period_label)

    def _same_period(label: str) -> bool:
        return label == period_label or (want_md is not None and _start_md(label) == want_md)

    for label, start, end in blocks:
        if start <= target_row <= end:
            if _same_period(label):
                return target_row  # 이미 맞는 기간 블록 안에 있음
            break  # target_row가 속한 블록이 있는데 기간이 다르면 아래에서 처리

    match = next(((start, end) for label, start, end in blocks if _same_period(label)), None)

    if match is None:
        # 미리 만들어둔 블록이 없음(향후 기간이 다 소진됨) -> 새로 라벨 + 여분 행을 만듦
        if isinstance(ws.cell(row=target_row, column=1), MergedCell):
            return target_row  # 이미 다른 기간 블록 안쪽이라 새 라벨을 못 씀 - 오류 대신 그대로 둠
        ws.cell(row=target_row, column=1).value = period_label
        end_row = target_row + ROLLOVER_BUFFER_ROWS - 1
        if end_row > target_row:
            ws.merge_cells(f"A{target_row}:A{end_row}")
        return target_row

    block_start, _block_end = match
    if block_start == target_row:
        return target_row
    if block_start < target_row:
        return target_row  # 이상 상황(라벨 블록이 이미 지나간 위치) - 안전하게 그대로 둠

    gap = block_start - target_row
    _delete_blank_rows_and_shift(ws, target_row, gap)
    return target_row


def _delete_blank_rows_and_shift(ws, start_row: int, count: int) -> None:
    """start_row부터 count개의 빈 여분 행을 지우고 그 아래 모든 행을 count칸 위로 당김.
    수식은 Translator로 새 위치에 맞게 다시 옮기고, 컬럼 A(월별)의 정산기간 라벨 병합도
    같이 당김. start_row~start_row+count-1 구간은 실제 데이터가 없는 빈 행이어야 안전함
    (호출하는 쪽인 _rollover_period_block에서 이미 그렇게 보장함)."""
    max_row = ws.max_row
    max_col = ws.max_column

    col_a_merges = [mc for mc in list(ws.merged_cells.ranges) if mc.min_col == 1 and mc.max_col == 1]

    to_shift = []
    for mc in col_a_merges:
        if mc.min_row >= start_row:
            to_shift.append(mc)
        elif mc.max_row >= start_row:
            # 방금 끝난 기간의 라벨 블록이 삭제 구간까지 넘어와 있던 경우 -> 실제 데이터가
            # 있는 부분(삭제 구간 시작 직전)까지로 줄임(당기지는 않음)
            ws.unmerge_cells(str(mc))
            new_end = start_row - 1
            if new_end > mc.min_row:
                ws.merge_cells(f"A{mc.min_row}:A{new_end}")

    for mc in to_shift:
        ws.unmerge_cells(str(mc))

    for r in range(start_row + count, max_row + 1):
        dst_row = r - count
        for c in range(1, max_col + 1):
            src = ws.cell(row=r, column=c)
            dst = ws.cell(row=dst_row, column=c)
            if isinstance(dst, MergedCell):
                continue
            val = None if isinstance(src, MergedCell) else src.value
            if isinstance(val, str) and val.startswith("="):
                try:
                    val = Translator(val, origin=src.coordinate).translate_formula(dst.coordinate)
                except Exception:
                    pass
            dst.value = val
            if not isinstance(src, MergedCell):
                dst.number_format = src.number_format
                dst.fill = copy.copy(src.fill)
                dst.font = copy.copy(src.font)
                dst.alignment = copy.copy(src.alignment)
                dst.border = copy.copy(src.border)

    for r in range(max(max_row - count + 1, start_row), max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            if isinstance(cell, MergedCell):
                continue
            cell.value = None

    for mc in to_shift:
        new_min = mc.min_row - count
        new_max = mc.max_row - count
        if new_max >= new_min and new_min >= 1:
            ws.merge_cells(f"A{new_min}:A{new_max}")

    if max_row - count >= start_row:
        try:
            ws.delete_rows(max_row - count + 1, count)
        except Exception:
            pass


def _parse_date_val(date_val) -> dt.date | None:
    """접수일자 셀 값을 날짜로 변환. datetime/date 객체는 그대로, 문자열('26.07.16',
    '2026-07-16', '2026.07.16' 등)은 여러 형식을 시도해서 파싱. 담당자가 개별로 셀에
    직접 타이핑한 파일은 날짜가 진짜 날짜형이 아니라 문자열로 들어오는 경우가 있음."""
    if date_val is None:
        return None
    if isinstance(date_val, dt.datetime):
        return date_val.date()
    if isinstance(date_val, dt.date):
        return date_val
    s = str(date_val).strip()
    if not s:
        return None
    s = s.split(" ")[0]  # '2026-11-12 00:00:00' 처럼 시각이 붙은 문자열
    for fmt in ("%y.%m.%d", "%Y.%m.%d", "%y-%m-%d", "%Y-%m-%d", "%y/%m/%d", "%Y/%m/%d", "%m/%d/%Y"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _is_stale_recv_date(vals: dict, period_start: dt.date) -> str | None:
    """담당자가 보내는 정산엑셀에는 이번에 새로 접수된 건뿐 아니라 그동안의 전체 이력이 통째로
    들어있는 경우가 많음. 그중 매장코드가 통합파일에 없는 행은 규칙상 '신규'로 처리하지만,
    접수일자가 지난 정산기간보다도 훨씬 이전(한 기간 이상 전)이면 사실은 예전에 이미 다른 방식으로
    처리된 건이 파일에만 남아있는 것일 수 있음. 이런 경우를 조용히 신규 처리하지 말고 사용자에게
    표시해서 실수로 잘못 등록되는 걸 막기 위한 확인용 함수. 오래된 날짜면 'YYYY-MM-DD' 문자열을,
    아니면 None을 반환함."""
    recv = vals.get("접수일자") or vals.get("보험시작일")
    d = _parse_date_val(recv)
    if d is None:
        return None
    threshold = period_start - dt.timedelta(days=31)
    if d < threshold:
        return d.isoformat()
    return None


# 음영(fill) 처리에서만 예외로 둘 칸. 값은 절대 건드리면 안 되는 '월별'(정산기간 라벨, 여러
# 행에 병합돼 있어 그 블록 전체를 대표하는 텍스트라 행 단위로 색칠하면 오히려 어색함)만 제외하고,
# '순번'은 _write_row에서는 값을 안 건드릴 뿐 음영까지 뺄 이유는 없음 - 예전엔 SKIP_WRITE_FIELDS를
# 그대로 재사용해서 순번 칸 음영이 기간이 지나도 노란색으로 남아있는 버그가 있었음.
SKIP_FILL_FIELDS = {"월별"}


def _fill_row_by_period(ws, header_map: dict, row: int, row_date, period_start: dt.date, period_end: dt.date) -> None:
    """한 행에 대해 접수일자(row_date)가 이번 정산 기간 안이면 노란색, 아니면 흰색(음영 없음)으로
    칠함. _recolor_by_period와 완전히 같은 규칙을, 새로 쓴 행 하나에 즉시 적용할 때 씀
    (전체 재계산 pass를 기다리지 않고 그 자리에서 바로 색이 반영되도록 하기 위한 안전장치)."""
    in_period = row_date is not None and period_start <= row_date <= period_end
    fill = YELLOW_FILL if in_period else NO_FILL
    for key, idx in header_map.items():
        if key in SKIP_FILL_FIELDS:
            continue
        cell = ws.cell(row=row, column=idx + 1)
        if isinstance(cell, MergedCell):
            continue
        cell.fill = fill


def _recolor_by_period(ws, header_map: dict, min_row: int, period_start: dt.date, period_end: dt.date) -> None:
    """접수일자가 이번 정산 기간(16일~다음달 15일) 안이면 노란색, 기간이 지난 항목이면 다시
    흰색(음영 없음)으로 되돌림. 월별 칸(정산기간 라벨)만 그대로 두고, 순번을 포함한 나머지
    칸은 전부 같이 칠함. 신규매장/폐점매장 시트 양쪽 다 이 규칙을 적용해서, 이번 달에 새로
    들어온 신규/폐점 매장을 한눈에 볼 수 있게 함."""
    date_idx = header_map.get("접수일자")
    name_idx = header_map.get("매장명")
    if date_idx is None or name_idx is None:
        return
    for r in range(min_row, ws.max_row + 1):
        name_val = ws.cell(row=r, column=name_idx + 1).value
        if not name_val:
            continue
        date_val = ws.cell(row=r, column=date_idx + 1).value
        row_date = _parse_date_val(date_val)
        _fill_row_by_period(ws, header_map, r, row_date, period_start, period_end)


def _extract_rate(ws, header_map: dict, field: str, pattern: str, default: float, min_row: int) -> float:
    idx = header_map.get(field)
    if idx is None:
        return default
    for r in range(min_row, min(ws.max_row, min_row + 50) + 1):
        cell = ws.cell(row=r, column=idx + 1)
        if isinstance(cell.value, str) and cell.value.startswith("="):
            m = re.search(pattern, cell.value)
            if m:
                return float(m.group(1))
    return default


def _compute_new_store_cert_values(vals: dict, rate1_pct: float, rate2: float) -> dict:
    stock = float(vals.get("재고") or 0)
    facility = float(vals.get("시설/비품") or 0)
    building = float(vals.get("건물") or 0)
    total = stock + facility + building
    property_premium = total * rate1_pct / 100
    pyeong = float(vals.get("평수") or 0)
    liability_premium = pyeong * rate2
    start = vals.get("보험시작일")
    end = vals.get("보험종기일")
    # 날짜형 셀뿐 아니라 담당자가 '2026-11-12'처럼 문자로 타이핑한 날짜도 인식(예전엔 문자면
    # 날짜를 못 읽어 보험료 0원, 날짜가 입력 형식 그대로 찍히는 문제가 생길 수 있었음)
    start_d = _parse_date_val(start)
    end_d = _parse_date_val(end)
    if start_d and end_d:
        days = (end_d - start_d).days
        premium = (property_premium + liability_premium) * days / 365
    else:
        premium = 0
    return {
        "stock_amt": int(stock),
        "facility_amt": int(facility),
        "building_amt": int(building),
        "premium": round(premium),
        "start_date": start_d.strftime("%Y.%m.%d") if start_d else str(start or ""),
        "start_date_yymmdd": start_d.strftime("%y%m%d") if start_d else "",
        "end_date": end_d.strftime("%Y.%m.%d") if end_d else str(end or ""),
        # 재고/시설비품/건물 중 하나라도 기재돼 있으면 재물부문 있음, 평수가 기재돼 있으면 영업배상 부문 있음.
        # 둘 다 없으면 가입증명서에서 해당 줄 자체를 지워야 함(사용자 요청).
        "has_property": bool(stock or facility or building),
        "has_liability": bool(pyeong),
        # 보험종기일이 이미 지난(오늘 이전) 건은 통합파일엔 기록해두되, 가입증명서는 만들 필요
        # 없음(사용자 요청) - 예: 뒤늦게 딸려온, 이미 끝난 단기 행사장 계약 등
        "period_ended": bool(end_d and end_d < dt.date.today()),
    }


_PREMIUM_MISMATCH_TOLERANCE = 1  # 반올림 오차 정도는 무시(원 단위)


def _to_number(v) -> float | None:
    """엑셀 칸 값을 숫자로 바꿈(문자/공백/None이면 None). 수식 칸을 data_only로 읽었을 때
    이미 계산된 숫자가 들어있는 걸 전제로 함."""
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _verify_new_store_premium(vals: dict, cert_vals: dict) -> dict | None:
    """신규매장 정산양식의 '잔여 합계보험료' 칸(담당자가 보낸 파일에 이미 계산되어 있는 값)이
    봇이 직접 계산한 값(=(연간재물보험료+연간영업배상보험료)*보험기간일수/365, 가입증명서에도
    쓰는 그 값)과 같은지 확인함. 그 칸이 비어있으면(서식에 없거나 계산 전) None을 반환하고
    조용히 넘어감 - 있을 때만 비교함."""
    stated = _to_number(vals.get("잔여합계보험료"))
    if stated is None:
        return None
    expected = cert_vals.get("premium")
    if expected is None:
        return None
    if abs(stated - expected) <= _PREMIUM_MISMATCH_TOLERANCE:
        return None
    return {"store_name": str(vals.get("매장명") or "").strip(), "expected": expected, "stated": round(stated)}


def _verify_closed_store_refund(vals: dict) -> dict | None:
    """폐점매장 정산양식의 '환급보험료' 칸이, 같은 행의 '연간총보험료'/'폐점일'/'보험시작일'
    값으로 직접 계산한 값(=연간총보험료-연간총보험료*DATEDIF(보험시작일,폐점일,"d")/365 -
    유진님 정산양식 서식에 실제로 들어있는 수식 그대로)과 같은지 확인함. 필요한 값 중 하나라도
    없으면(서식이 다르거나 계산 전) None을 반환하고 조용히 넘어감."""
    annual_premium = _to_number(vals.get("연간총보험료"))
    refund_stated = _to_number(vals.get("환급보험료"))
    start = vals.get("보험시작일")
    closed = vals.get("폐점일")
    if annual_premium is None or refund_stated is None:
        return None
    if not (hasattr(start, "date") and hasattr(closed, "date")):
        return None
    used_days = (closed - start).days
    expected = annual_premium - annual_premium * used_days / 365
    if abs(refund_stated - expected) <= _PREMIUM_MISMATCH_TOLERANCE:
        return None
    return {
        "store_name": str(vals.get("매장명") or "").strip(),
        "expected": round(expected),
        "stated": round(refund_stated),
    }


def _load_policy_numbers() -> dict:
    try:
        with open(POLICY_NUMBERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _save_policy_numbers(policy_numbers: dict) -> None:
    os.makedirs(os.path.dirname(POLICY_NUMBERS_FILE), exist_ok=True)
    with open(POLICY_NUMBERS_FILE, "w", encoding="utf-8") as f:
        json.dump(policy_numbers, f, ensure_ascii=False, indent=2)


async def set_policy_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    if len(context.args) < 2:
        await update.message.reply_text(
            "사용법: /setpolicy <브랜드명> <증권번호>\n"
            "예: /setpolicy 트레몰로 82509565736000\n\n"
            "브랜드명은 엑셀 시트 A1 셀에 적힌 이름(예: '트레몰로', '월메이드')과 정확히 같아야 해요."
        )
        return
    brand = context.args[0]
    policy_no = context.args[1]
    policy_numbers = _load_policy_numbers()
    policy_numbers[brand] = policy_no
    _save_policy_numbers(policy_numbers)
    await update.message.reply_text(f"✅ '{brand}' 브랜드의 증권번호를 등록했어요: {policy_no}")


# ㈜세정·OVLR 담당자 기본 연락처(2026-11-16 유진님이 알려주심). 서버에 저장된 연락처(/setcontact)가
# 있으면 그쪽이 우선이고, 없는 사람만 여기서 채움 - 재배포/저장공간 초기화와 상관없이 항상 쓸 수 있게.
_DEFAULT_CONTACTS = {
    "김가진": "kgj15@sejung.co.kr",
    "김연희": "yhkim@sejung.co.kr",
    "안해인": "hian01@sejung.co.kr",
    "서석란": "srseo@sejung.co.kr",
    "정은주": "ejjung01@ovlr.ai",
    "김영란": "yrkim01@ovlr.ai",
    "김훈영": "hykim@ovlr.ai",
    "배순민": "smbae@sejung.co.kr",
    "박재용": "jypark@sejung.co.kr",
    "장혜진": "hjhang@sejung.co.kr",
}


def _load_contacts() -> dict:
    try:
        with open(CONTACTS_FILE, "r", encoding="utf-8") as f:
            saved = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        saved = {}
    merged = {_norm_contact_name(k): v for k, v in _DEFAULT_CONTACTS.items()}
    merged.update(saved)
    return merged


def _save_contacts(contacts: dict) -> None:
    os.makedirs(os.path.dirname(CONTACTS_FILE), exist_ok=True)
    with open(CONTACTS_FILE, "w", encoding="utf-8") as f:
        json.dump(contacts, f, ensure_ascii=False, indent=2)


def _load_muted_senders() -> list:
    try:
        with open(MUTED_SENDERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return list(_DEFAULT_MUTED_SENDERS)


def _save_muted_senders(keywords: list) -> None:
    os.makedirs(os.path.dirname(MUTED_SENDERS_FILE), exist_ok=True)
    with open(MUTED_SENDERS_FILE, "w", encoding="utf-8") as f:
        json.dump(keywords, f, ensure_ascii=False, indent=2)


def _is_muted_sender(sender: str, subject: str = "", list_unsubscribe: str = "") -> bool:
    """보낸사람 표시(이름+이메일 통째로)나 제목에 무시 키워드가 포함되면 True.
    'List-Unsubscribe' 헤더(광고/구독성 메일에 거의 항상 붙는 헤더)가 있는 것도 광고성으로
    간주함 - 단, 이미 등록된 정산 담당자 연락처(contacts.json)에서 보낸 메일은 절대 무시하지
    않음(실수로 중요한 정산 메일을 놓치는 사고를 막기 위함)."""
    sender_email = (email.utils.parseaddr(sender)[1] or "").lower()
    contacts = _load_contacts()
    if sender_email and sender_email in set(contacts.values()):
        return False
    haystack = f"{sender} {subject}".lower()
    keywords = _load_muted_senders()
    if any(kw.lower() in haystack for kw in keywords if kw):
        return True
    if list_unsubscribe:
        return True
    return False


async def mute_sender_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    if not context.args:
        await update.message.reply_text(
            "사용법: /mutesender <키워드>\n예: /mutesender duolingo\n\n"
            "보낸사람 이름/이메일이나 제목에 이 키워드가 들어있으면 '새 메일 도착' 알림을 안 보내요."
        )
        return
    keyword = " ".join(context.args).strip()
    keywords = _load_muted_senders()
    if keyword.lower() not in [k.lower() for k in keywords]:
        keywords.append(keyword)
        _save_muted_senders(keywords)
    await update.message.reply_text(f"🔕 앞으로 '{keyword}'가 포함된 메일은 알림을 안 보낼게요.")


async def unmute_sender_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    if not context.args:
        await update.message.reply_text("사용법: /unmutesender <키워드>\n예: /unmutesender duolingo")
        return
    keyword = " ".join(context.args).strip()
    keywords = _load_muted_senders()
    new_keywords = [k for k in keywords if k.lower() != keyword.lower()]
    if len(new_keywords) == len(keywords):
        await update.message.reply_text(f"'{keyword}'는 무시 목록에 없어요. /mutedsenders 로 확인해주세요.")
        return
    _save_muted_senders(new_keywords)
    await update.message.reply_text(f"🔔 '{keyword}' 무시를 해제했어요.")


async def list_muted_senders_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    keywords = _load_muted_senders()
    if not keywords:
        await update.message.reply_text("무시 중인 보낸사람 키워드가 없어요.")
        return
    lines = "\n".join(f"- {k}" for k in keywords)
    await update.message.reply_text(
        f"무시 중인 보낸사람/제목 키워드:\n{lines}\n\n"
        "(이 키워드가 포함된 메일과, 'List-Unsubscribe' 헤더가 붙은 광고성 메일은 알림을 안 보내요. "
        "단, 등록된 담당자 연락처로 온 메일은 절대 무시 안 해요.)"
    )


def _norm_contact_name(name: str) -> str:
    return re.sub(r"\s+", "", str(name or "")).strip()


async def set_contact_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    if len(context.args) < 2:
        await update.message.reply_text(
            "사용법: /setcontact <담당자 이름> <이메일>\n"
            "예: /setcontact 안해인 hian01@sejung.co.kr\n\n"
            "이름은 정산양식 엑셀의 '접수자' 칸에 적힌 이름과 정확히 같아야 해요. "
            "등록해두면, 그 담당자가 텔레그램으로 직접 올린 정산양식에서도 신규매장 "
            "가입증명서를 자동으로 그 담당자 이메일로 보내드려요."
        )
        return
    name = context.args[0]
    email_addr = context.args[1]
    contacts = _load_contacts()
    contacts[_norm_contact_name(name)] = email_addr
    _save_contacts(contacts)
    await update.message.reply_text(f"✅ '{name}' 담당자의 이메일을 등록했어요: {email_addr}")


async def list_contacts_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    contacts = _load_contacts()
    if not contacts:
        await update.message.reply_text("등록된 담당자 연락처가 없어요. /setcontact로 등록해주세요.")
        return
    lines = "\n".join(f"- {name}: {addr}" for name, addr in contacts.items())
    await update.message.reply_text(f"등록된 담당자 연락처:\n{lines}")


def _load_pending_certs() -> dict:
    try:
        with open(PENDING_CERTS_INDEX_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _save_pending_certs(pending: dict) -> None:
    os.makedirs(os.path.dirname(PENDING_CERTS_INDEX_FILE), exist_ok=True)
    with open(PENDING_CERTS_INDEX_FILE, "w", encoding="utf-8") as f:
        json.dump(pending, f, ensure_ascii=False, indent=2)


def _queue_pending_certs(items: list, target_email: str, subject: str, body: str, label: str = "가입증명서",
                         sent_keys: list | None = None) -> str:
    """가입증명서(한 담당자 앞으로 1건 이상)를 바로 보내지 않고 대기시켜둠.
    items: [{"store_name":..., "out_name":..., "pdf_bytes":...}, ...]
    유진님이 '보내기'를 누르면 그때 첨부를 모두 묶어서 실제 발송함.
    sent_keys: 발송이력 키(_sj_sent_key). 이미 보낸 매장이면 보내기 전에 한 번 더 물어봄(2026-11-23)."""
    os.makedirs(PENDING_CERTS_DIR, exist_ok=True)
    cert_id = uuid.uuid4().hex[:10]
    attachments = []
    for i, item in enumerate(items):
        pdf_path = os.path.join(PENDING_CERTS_DIR, f"{cert_id}_{i}.pdf")
        with open(pdf_path, "wb") as f:
            f.write(item["pdf_bytes"])
        attachments.append({"out_name": item["out_name"], "pdf_path": pdf_path})
    pending = _load_pending_certs()
    pending[cert_id] = {
        "store_names": [item["store_name"] for item in items],
        "target_email": target_email,
        "subject": subject,
        "body": body,
        "attachments": attachments,
        "label": label,
        "sent_keys": list(sent_keys or []),
    }
    _save_pending_certs(pending)
    return cert_id


def _remove_pending_cert(cert_id: str) -> dict | None:
    pending = _load_pending_certs()
    entry = pending.pop(cert_id, None)
    if entry is None:
        return None
    _save_pending_certs(pending)
    for a in entry.get("attachments", []):
        try:
            os.remove(a["pdf_path"])
        except OSError:
            pass
    return entry


async def handle_cert_confirmation(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    if query is None:
        return
    if query.from_user is None or query.from_user.id != ALLOWED_USER_ID:
        await query.answer()
        return

    data = query.data or ""
    if ":" not in data:
        await query.answer()
        return
    if data.startswith("sjdup:"):
        # ㈜세정 중복 확인 버튼 [덮어쓰기][새로 추가][무시] (2026-11-23)
        await _sj_handle_dup_callback(query)
        return
    action, cert_id = data.split(":", 1)

    pending = _load_pending_certs()
    entry = pending.get(cert_id)
    if entry is None:
        await query.answer()
        await query.edit_message_text("이미 처리됐거나 없는 요청이에요.")
        return

    if action in ("sendcert", "sendcertf"):
        await query.answer()
        names_text = ", ".join(entry["store_names"])
        sent_keys = entry.get("sent_keys") or []
        if action == "sendcert" and sent_keys:
            # 이미 보낸 매장이면 다시 보낼지 한 번 더 확인(2026-11-23). 원래 메시지·버튼은 그대로 둠.
            already = await asyncio.to_thread(_sj_sent_lookup, sent_keys)
            if already:
                lines = []
                for k, rec in already.items():
                    parts = k.split("|")
                    store = " ".join(parts[2:4]) if len(parts) >= 4 else k
                    try:
                        when = dt.datetime.strptime(rec.get("at", "")[:10], "%Y-%m-%d")
                        when_txt = f"{when.month}월{when.day}일"
                    except ValueError:
                        when_txt = rec.get("at") or "?"
                    lines.append(f"- {store}: 이미 {when_txt} 발송됨({rec.get('to') or ''})")
                await _send_message_retrying(
                    context.bot, ALLOWED_USER_ID,
                    f"⚠️ '{names_text}' {entry.get('label', '가입증명서')}\n" + "\n".join(lines) + "\n다시 보낼까요?",
                    reply_markup=InlineKeyboardMarkup([[
                        InlineKeyboardButton("📨 다시 보내기", callback_data=f"sendcertf:{cert_id}"),
                        InlineKeyboardButton("🗑 보내지 않기", callback_data=f"cancelcert:{cert_id}"),
                    ]]),
                )
                return
        try:
            mail_attachments = []
            for a in entry["attachments"]:
                with open(a["pdf_path"], "rb") as f:
                    mail_attachments.append((a["out_name"], f.read()))
            _send_email(
                entry["target_email"],
                subject=entry["subject"],
                body=entry["body"],
                attachments=mail_attachments,
            )
            _remove_pending_cert(cert_id)
            if sent_keys:
                await asyncio.to_thread(_sj_sent_record, sent_keys, entry["target_email"])
            await query.edit_message_text(f"✅ '{names_text}' {entry.get('label', '가입증명서')}를 {entry['target_email']}로 보냈어요.")
        except Exception:
            logger.exception("대기 중이던 가입증명서 발송 중 오류")
            await query.edit_message_text(f"⚠️ '{names_text}' {entry.get('label', '가입증명서')} 발송에 실패했어요. 이 메시지에서 다시 눌러 시도해주세요.")
    elif action == "holdcert":
        # 메시지와 버튼을 그대로 남겨둠 -> 명령어 없이도, 나중에 이 메시지를 찾아 스크롤해서
        # 그냥 다시 '보내기'를 누르기만 하면 됨. 살짝 알려주는 팝업만 띄우고 메시지는 안 건드림.
        await query.answer("⏸ 대기 상태예요. 나중에 이 메시지에서 다시 눌러 보내시면 돼요.", show_alert=True)
    elif action == "cancelcert":
        await query.answer()
        names_text = ", ".join(entry["store_names"])
        _remove_pending_cert(cert_id)
        await query.edit_message_text(f"🗑 '{names_text}' {entry.get('label', '가입증명서')} 발송을 취소했어요.")


async def list_pending_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    pending = _load_pending_certs()
    if not pending:
        await update.message.reply_text("대기 중인 가입증명서가 없어요.")
        return
    for cert_id, entry in pending.items():
        keyboard = InlineKeyboardMarkup([[
            InlineKeyboardButton("✅ 보내기", callback_data=f"sendcert:{cert_id}"),
            InlineKeyboardButton("🗑 취소", callback_data=f"cancelcert:{cert_id}"),
        ]])
        names_text = ", ".join(entry["store_names"])
        await update.message.reply_text(
            f"'{names_text}' → {entry['target_email']}",
            reply_markup=keyboard,
        )


def _load_brand_names() -> dict:
    try:
        with open(BRAND_NAMES_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _save_brand_names(brand_names: dict) -> None:
    os.makedirs(os.path.dirname(BRAND_NAMES_FILE), exist_ok=True)
    with open(BRAND_NAMES_FILE, "w", encoding="utf-8") as f:
        json.dump(brand_names, f, ensure_ascii=False, indent=2)


def _display_name(brand: str) -> str:
    """봇이 파일을 보내드릴 때 쓸 이름. 따로 지정해둔 게 있으면 그 이름, 없으면 브랜드명(A1 셀 이름) 그대로."""
    return _load_brand_names().get(brand, brand)


async def set_brand_name_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    if len(context.args) < 2:
        await update.message.reply_text(
            "사용법: /setbrandname <브랜드명> <파일에 표시할 이름>\n"
            "예: /setbrandname 올리비아로렌&주얼리 올리비아로렌&오뷔엘알&주얼리\n\n"
            "브랜드명(첫 번째)은 /brands 로 확인되는 이름과 정확히 같아야 해요. "
            "엑셀 A1 셀이나 신규/폐점 매장 비교 기준은 그대로 두고, 봇이 보내드리는 통합파일의 이름만 바뀌어요."
        )
        return
    brand = context.args[0]
    new_name = " ".join(context.args[1:])
    brand_names = _load_brand_names()
    brand_names[brand] = new_name
    _save_brand_names(brand_names)
    await update.message.reply_text(f"✅ 이제 '{brand}' 통합파일을 보내드릴 때 '{new_name}.xlsx'로 보내드릴게요.")


def _load_brand_aliases() -> dict:
    try:
        with open(BRAND_ALIASES_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _save_brand_aliases(aliases: dict) -> None:
    os.makedirs(os.path.dirname(BRAND_ALIASES_FILE), exist_ok=True)
    with open(BRAND_ALIASES_FILE, "w", encoding="utf-8") as f:
        json.dump(aliases, f, ensure_ascii=False, indent=2)


def _resolve_brand_alias(brand: str) -> str:
    """담당자마다 A1 셀에 브랜드명을 조금씩 다르게 적어 보내는 경우가 있어서(예: '올리비아로렌&주얼리'
    vs '올리비아로렌&오뷔엘알&주얼리'), 등록해둔 별칭이 있으면 진짜(정식) 브랜드명으로 바꿔줌.
    이걸 안 하면 같은 브랜드인데 이름이 살짝 달라서 별도의 새 통합파일로 쪼개져 버림."""
    return _load_brand_aliases().get(brand, brand)


async def set_brand_alias_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    if len(context.args) < 2:
        await update.message.reply_text(
            "사용법: /setbrandalias <다르게 인식된 이름> <진짜 브랜드명>\n"
            "예: /setbrandalias 올리비아로렌&주얼리 올리비아로렌&오뷔엘알&주얼리\n\n"
            "어떤 담당자가 보낸 엑셀의 브랜드명이 <다르게 인식된 이름>으로 나오면, 이제부터 <진짜 브랜드명>의 "
            "통합파일로 합쳐서 처리해요. 진짜 브랜드명은 /brands 로 확인할 수 있어요."
        )
        return
    alias = context.args[0]
    canonical = " ".join(context.args[1:])
    aliases = _load_brand_aliases()
    aliases[alias] = canonical
    _save_brand_aliases(aliases)
    await update.message.reply_text(f"✅ 앞으로 '{alias}'로 인식되는 엑셀은 '{canonical}' 통합파일로 합쳐서 처리할게요.")


async def set_cert_template_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """새 브랜드의 실제 발급된 예시 가입증명서 PDF를 그 브랜드의 서식으로 등록.
    사용법: 예시 PDF를 먼저 보낸 뒤, 그 PDF 메시지에 '답장(reply)'하면서 '/setcerttemplate 브랜드명'을 입력."""
    if not is_allowed(update):
        return
    if not context.args:
        await update.message.reply_text(
            "사용법: 먼저 그 브랜드로 실제 발급된 예시 가입증명서 PDF를 보내주세요. 그 다음 그 PDF 메시지에 "
            "답장(reply)하면서 아래처럼 입력해주세요.\n\n/setcerttemplate <브랜드명>\n예: /setcerttemplate 트레몰로\n\n"
            "브랜드명은 /brands 로 확인되는 이름과 정확히 같아야 해요."
        )
        return
    reply = update.message.reply_to_message
    if not reply or not reply.document:
        await update.message.reply_text(
            "가입증명서 예시 PDF 파일에 답장(reply)하는 방식으로 이 명령어를 보내주세요.\n"
            "(PDF를 먼저 보내고, 그 메시지를 길게 눌러 '답장'을 선택한 뒤 명령어를 입력)"
        )
        return
    doc = reply.document
    if not (doc.file_name or "").lower().endswith(".pdf"):
        await update.message.reply_text("PDF 파일에 답장해주세요.")
        return
    brand = " ".join(context.args)
    try:
        tg_file = await doc.get_file()
        file_bytes = bytes(await tg_file.download_as_bytearray())
    except Exception:
        logger.exception("가입증명서 서식 다운로드 중 오류")
        await update.message.reply_text("파일을 받아오는 중 오류가 발생했어요. 다시 시도해주세요.")
        return
    os.makedirs(CERT_TEMPLATES_DIR, exist_ok=True)
    with open(os.path.join(CERT_TEMPLATES_DIR, f"{brand}.pdf"), "wb") as f:
        f.write(file_bytes)
    await update.message.reply_text(
        f"✅ '{brand}' 가입증명서 서식으로 등록했어요. 이제부터 이 브랜드는 이 서식을 기준으로, "
        "보험기간/점포명/주소/보상한도액/보험료만 바뀌어서 인증서가 나가요.\n\n"
        "⚠️ 다만 이 서식은 트레몰로 서식과 항목 위치가 조금 다를 수 있어서 값이 살짝 어긋나 보일 수 있어요. "
        "그런 경우 알려주시면 위치를 정확히 맞춰서 코드에 반영해드릴게요.\n"
        "또한 이 서식은 재배포하면 사라지니, 재배포 후에는 다시 등록해주셔야 해요."
    )


async def list_brands_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return
    if not os.path.isdir(MASTERS_DIR):
        await update.message.reply_text("등록된 브랜드가 아직 없어요.")
        return
    brands = sorted(f[:-5] for f in os.listdir(MASTERS_DIR) if f.endswith(".xlsx"))
    if not brands:
        await update.message.reply_text("등록된 브랜드가 아직 없어요.")
        return
    policy_numbers = _load_policy_numbers()
    brand_names = _load_brand_names()
    aliases = _load_brand_aliases()
    lines = []
    for b in brands:
        status = "증권번호 등록됨" if b in policy_numbers else "⚠️ 증권번호 미등록"
        if b in brand_names:
            status += f", 표시명: {brand_names[b]}"
        alias_list = [a for a, c in aliases.items() if c == b]
        if alias_list:
            status += f", 별칭: {', '.join(alias_list)}"
        lines.append(f"- {b} ({status})")
    await update.message.reply_text("📋 등록된 브랜드 목록:\n" + "\n".join(lines))


async def send_master_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """지금 저장되어 있는 통합파일을 새로 올릴 필요 없이 그대로(최신 파일명 등 반영해서) 다시 받아봄"""
    if not is_allowed(update):
        return
    if not context.args:
        await update.message.reply_text(
            "사용법: /sendmaster <브랜드명>\n예: /sendmaster 트레몰로\n\n"
            "지금 저장되어 있는 통합파일을 다시 보내드려요. 정확한 브랜드명은 /brands 로 확인할 수 있어요."
        )
        return
    brand = " ".join(context.args)
    # 통합파일은 이제 서버가 아니라 구글 드라이브('가입증명서및정산/정산_{브랜드}' 폴더)에
    # 있으므로, 거기서 먼저 찾아보고 없을 때만(연결 문제 등) 예전 서버 백업을 씀.
    master_bytes, _drive_id, _drive_name = _load_brand_master_from_drive(brand)
    if master_bytes is None:
        master_path = os.path.join(MASTERS_DIR, f"{brand}.xlsx")
        if not os.path.exists(master_path):
            await update.message.reply_text(f"'{brand}' 통합파일을 찾지 못했어요. /brands 로 정확한 브랜드명을 확인해주세요.")
            return
        with open(master_path, "rb") as f:
            master_bytes = f.read()
    display_name = _display_name(brand)
    await update.message.reply_document(
        document=io.BytesIO(master_bytes),
        filename=f"{display_name}.xlsx",
        caption=f"📎 '{display_name}' 통합파일이에요.",
    )


def _sync_brand_excel(file_bytes: bytes) -> dict | None:
    """엑셀을 읽어 브랜드를 판별하고, 저장된 통합파일과 비교해 신규/폐점 매장을 반영.
    반환: {"brand", "cold_start", "new_stores": [...], "closed_count": int, "master_bytes": bytes}
    브랜드를 못 찾으면 None."""
    try:
        input_wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception:
        logger.exception("엑셀 파일을 여는 중 오류")
        return None

    brand = _get_brand_name(input_wb)
    if not brand:
        return None
    brand = _resolve_brand_alias(brand)

    # 통합파일은 이제 서버에 따로 보관하지 않고, 유진님이 실제로 쓰시는 구글 드라이브의
    # '가입증명서및정산/정산_{브랜드}' 폴더 안 취합본 엑셀을 직접 찾아서 그 자리에서 갱신함
    # (2026-09-29, 유진님 요청 - 더 이상 텔레그램으로 통합파일을 따로 안 보내고, PC에 자동
    # 동기화되는 그 파일에 바로 반영해달라고 하심).
    drive_bytes, drive_file_id, drive_file_name = _load_brand_master_from_drive(brand)
    drive_master_missing = drive_bytes is None

    if drive_master_missing:
        # 드라이브에서 '정산_{brand}' 폴더나 그 안의 취합본 엑셀을 못 찾음(폴더가 아직
        # 없거나, 드라이브 연결 자체에 문제가 있을 수 있음) - 데이터 유실을 막기 위해
        # 예전처럼 서버에 임시로 보관해두고, 안내 메시지로 확인을 요청함.
        os.makedirs(MASTERS_DIR, exist_ok=True)
        master_path = os.path.join(MASTERS_DIR, f"{brand}.xlsx")
        if not os.path.exists(master_path):
            with open(master_path, "wb") as f:
                f.write(file_bytes)
            return {
                "brand": brand, "cold_start": True, "new_stores": [], "closed_count": 0,
                "master_bytes": file_bytes, "drive_master_missing": True,
            }
        # 쓰기용(수식 보존)과 읽기용(중복 확인은 계산된 값 기준) 두 벌로 마스터 파일을 엶.
        # 폐점매장 시트처럼 '매장명' 칸 자체가 VLOOKUP 수식인 행이 실제로 존재하기 때문에,
        # 중복 판별은 반드시 계산된 값(data_only=True) 기준으로 해야 정확함.
        master_wb = openpyxl.load_workbook(master_path, data_only=False)
        master_wb_values = openpyxl.load_workbook(master_path, data_only=True)
    else:
        master_wb = openpyxl.load_workbook(io.BytesIO(drive_bytes), data_only=False)
        master_wb_values = openpyxl.load_workbook(io.BytesIO(drive_bytes), data_only=True)

    policy_numbers = _load_policy_numbers()
    policy_no = policy_numbers.get(brand, "")

    # 오래된 접수일자인 '신규' 건을 나중에 알림으로 표시하기 위해, 그리고 정산기간 라벨
    # 블록을 맞춰넣기 위해 미리 계산해둠
    period_start, period_end = _current_period()
    current_period_label = _format_period_label(period_start, period_end)

    new_stores = []
    closed_count = 0
    closed_stores = []
    premium_mismatches = []

    input_new_sheets = _find_type_sheets(input_wb, "신규매장")
    master_new_sheets = _find_type_sheets(master_wb, "신규매장")
    master_new_sheets_values = _find_type_sheets(master_wb_values, "신규매장")

    # 담당자가 보내는 개별 신청서는 '신규매장(정상)/신규매장(상설)' 구분 없이 그냥 '신규매장'
    # 하나로만 오는 경우가 많아서(_find_type_sheets 설명 참고), 이때는 일단 '정상'으로 간주해서
    # 비교함. 그런데 그 안에 사실은 이미 '상설' 쪽에 등록된 매장이 섞여 들어오면, '정상' 쪽
    # 기존 목록에는 없으니 신규로 잘못 잡혀서 똑같은 매장이 중복 등록되는 사고가 남. 이를 막기
    # 위해 어떤 서브타입을 처리하든 정상/상설 등 모든 서브타입에 이미 등록된 매장 전체를 합쳐서
    # 중복 여부를 확인함(실제로 새 행을 쓰는 시트는 여전히 해당 서브타입 시트 하나뿐).
    all_new_existing_keys = set()
    all_new_existing_identities = set()
    for _st, _mws in master_new_sheets.items():
        _mws_values = master_new_sheets_values.get(_st)
        if _mws_values is None:
            continue
        _h, _mr = _build_header_map(_mws)
        if not _h:
            continue
        _existing_rows = _extract_data_rows(_mws_values, _h, _mr, formula_ws=_mws, wb_values=master_wb_values)
        all_new_existing_keys |= {_dedup_key(v) for v in _existing_rows}
        all_new_existing_identities |= {_store_identity_key(v) for v in _existing_rows}

    skipped_placeholder_stores = []

    for sub_type, input_ws in input_new_sheets.items():
        master_ws = master_new_sheets.get(sub_type)
        master_ws_values = master_new_sheets_values.get(sub_type)
        if master_ws is None or master_ws_values is None:
            continue
        input_header, input_min_row = _build_header_map(input_ws)
        master_header, master_min_row = _build_header_map(master_ws)
        if not input_header or not master_header:
            continue
        existing_keys = set(all_new_existing_keys)
        existing_identities = set(all_new_existing_identities)
        rate1 = _extract_rate(master_ws, master_header, "연간재물보험료", r"\*([\d.]+)%", 0.0665, master_min_row)
        rate2 = _extract_rate(master_ws, master_header, "연간영업배상보험료", r"\*([\d.]+)", 1793, master_min_row)

        for vals in _extract_data_rows(input_ws, input_header, input_min_row):
            # 매장코드가 아직 '미개설(추후전달)' 같은 임시 문구인 채로 왔는데, 매장명+사업자번호
            # (또는 주소)가 이미 등록된 매장과 같으면 -- 실제 코드가 나중에 배정되어 등록된 뒤
            # 담당자가 예전 임시 이력을 다시 보낸 경우일 수 있음. 이 경우 매장코드 문구 자체가
            # 매번 미세하게 달라질 수 있어(dedup 키로는 안 걸러짐) 그대로 두면 같은 매장이 계속
            # 중복 등록되므로, 여기서 별도로 걸러서 건너뜀(통합파일에 추가/증명서 발급 안 함)
            if _is_placeholder_code(vals.get("매장코드")) and _store_identity_key(vals) in existing_identities:
                skipped_placeholder_stores.append({
                    "store_name": str(vals.get("매장명") or "").strip(),
                })
                continue

            key = _dedup_key(vals)
            if key in existing_keys:
                continue
            existing_keys.add(key)
            existing_identities.add(_store_identity_key(vals))

            # 오래된 접수일자 경고는 정산파일에 실제로 적혀 있던 원래 접수일자 기준으로
            # 판단해야 하므로, 접수일자를 오늘 날짜로 덮어쓰기 전에 미리 계산해둠
            stale_recv_date = _is_stale_recv_date(vals, period_start)

            # 접수일자는 정산파일/메일에 뭐라고 적혀 있든, 실제로 봇이 이 건을 처리한 날짜
            # (메일이 온 날 또는 정산파일을 텔레그램에 업로드한 날)로 통일해서 통합파일에 기록함
            vals = dict(vals)
            vals["접수일자"] = dt.date.today()

            target_row = _find_target_row(master_ws, master_header, master_min_row)
            target_row = _rollover_period_block(master_ws, master_header, target_row, current_period_label)
            template_row = _find_template_row(master_ws, master_header, master_min_row, target_row)
            _write_row(master_ws, master_header, target_row, vals, template_row)
            # 접수일자(=오늘, 처리일) 기준으로 즉시 색칠. 보험시작일 등 다른 날짜 칸과
            # 헷갈리지 않도록, 전체 재계산(_recolor_by_period)을 기다리지 않고 이 자리에서 바로 반영.
            _fill_row_by_period(master_ws, master_header, target_row, vals.get("접수일자"), period_start, period_end)

            cert_vals = _compute_new_store_cert_values(vals, rate1, rate2)
            mismatch = _verify_new_store_premium(vals, cert_vals)
            if mismatch:
                premium_mismatches.append({**mismatch, "kind": "신규(잔여 합계보험료)"})
            address = str(vals.get("매장주소") or "").strip()
            # 엑셀 매장주소가 "...52\n(백화점내)"처럼 줄바꿈으로 나뉘어 있는 경우가 있는데,
            # 그대로 두면 가입증명서에 한 줄만 그려서 괄호 부분이 아예 안 보이게 됨(PDF 렌더링이
            # 한 줄짜리로만 처리됨). 줄바꿈을 공백으로 바꿔서 괄호 내용까지 한 줄로 전부 표시함.
            address = re.sub(r"\s*\n\s*", " ", address)
            address = re.sub(r"(?<=\S)\(", " (", address)
            new_stores.append({
                "policy_no": policy_no or DEFAULT_POLICY_NO,
                "store_code": str(vals.get("매장코드") or "").strip(),
                "store_name": str(vals.get("매장명") or "").strip(),
                "address": address,
                "received_by": str(vals.get("접수자") or "").strip(),
                "stale_recv_date": stale_recv_date,
                **cert_vals,
            })

    input_closed_sheets = _find_type_sheets(input_wb, "폐점매장")
    master_closed_sheets = _find_type_sheets(master_wb, "폐점매장")
    master_closed_sheets_values = _find_type_sheets(master_wb_values, "폐점매장")

    # 신규매장과 마찬가지로, 정상/상설 등 모든 서브타입에 이미 등록된 폐점매장을 합쳐서 중복 검사함.
    # (폐점매장 시트는 애초에 '매장코드' 칸이 없는 서식이라 미배정 코드 문제가 생기지 않으므로,
    # 신규매장과 달리 여기는 원래 이름+날짜 기준 dedup만으로 충분함)
    all_closed_existing_keys = set()
    for _st, _mws in master_closed_sheets.items():
        _mws_values = master_closed_sheets_values.get(_st)
        if _mws_values is None:
            continue
        _h, _mr = _build_header_map(_mws)
        if not _h:
            continue
        all_closed_existing_keys |= {
            _dedup_key(v)
            for v in _extract_data_rows(_mws_values, _h, _mr, formula_ws=_mws, wb_values=master_wb_values)
        }

    for sub_type, input_ws in input_closed_sheets.items():
        master_ws = master_closed_sheets.get(sub_type)
        master_ws_values = master_closed_sheets_values.get(sub_type)
        if master_ws is None or master_ws_values is None:
            continue
        input_header, input_min_row = _build_header_map(input_ws)
        master_header, master_min_row = _build_header_map(master_ws)
        if not input_header or not master_header:
            continue
        existing_keys = set(all_closed_existing_keys)

        for vals in _extract_data_rows(input_ws, input_header, input_min_row):
            key = _dedup_key(vals)
            if key in existing_keys:
                continue
            existing_keys.add(key)

            # 폐점매장도 마찬가지로 접수일자를 실제 처리일(오늘)로 통일해서 기록함
            vals = dict(vals)
            vals["접수일자"] = dt.date.today()

            target_row = _find_target_row(master_ws, master_header, master_min_row)
            target_row = _rollover_period_block(master_ws, master_header, target_row, current_period_label)
            template_row = _find_template_row(master_ws, master_header, master_min_row, target_row)
            _write_row(master_ws, master_header, target_row, vals, template_row)
            _fill_row_by_period(master_ws, master_header, target_row, vals.get("접수일자"), period_start, period_end)
            closed_count += 1
            refund_mismatch = _verify_closed_store_refund(vals)
            if refund_mismatch:
                premium_mismatches.append({**refund_mismatch, "kind": "폐점(환급보험료)"})
            # 이 행이 통합파일의 '폐점매장' 시트에 그대로 저장되므로(_write_row), 이 매장이
            # '신규매장'으로 기록된 적이 없어도 나중에 _all_closed_store_rows()로 브랜드/
            # 대표자명을 다시 찾을 수 있음(별도 저장소 없이 통합파일 자체가 정보 출처).
            closed_stores.append({
                "store_name": str(vals.get("매장명") or "").strip(),
                "owner_name": str(vals.get("대표자명") or "").strip(),
                "company_refund": _is_company_refund(vals),
            })

    if not new_stores and not closed_count:
        return {
            "brand": brand,
            "cold_start": False,
            "new_stores": [],
            "closed_count": 0,
            "closed_stores": [],
            "master_bytes": None,
            "skipped_placeholder_stores": skipped_placeholder_stores,
        }

    # 이번 정산 기간(16일~다음달 15일) 기준으로 신규/폐점매장 시트 전체를 다시 칠함.
    # 새로 추가된 항목뿐 아니라, 기간이 지나 더 이상 '이번 달' 항목이 아닌 예전 노란색도
    # 흰색으로 되돌려서 항상 현재 기간만 노랗게 보이도록 함.
    period_start, period_end = _current_period()
    for sheets_dict in (master_new_sheets, master_closed_sheets):
        for master_ws in sheets_dict.values():
            header, min_row = _build_header_map(master_ws)
            if header:
                _recolor_by_period(master_ws, header, min_row, period_start, period_end)

    _strip_external_links(master_wb)

    out = io.BytesIO()
    master_wb.save(out)
    master_bytes = out.getvalue()

    if drive_master_missing:
        with open(master_path, "wb") as f:
            f.write(master_bytes)
        master_saved_to_drive = False
    else:
        master_saved_to_drive = _save_brand_master_to_drive(drive_file_id, master_bytes)
        if not master_saved_to_drive:
            # 드라이브 저장이 실패하면(연결 끊김 등) 유실 방지용으로 서버에도 백업해둠
            os.makedirs(MASTERS_DIR, exist_ok=True)
            with open(os.path.join(MASTERS_DIR, f"{brand}.xlsx"), "wb") as f:
                f.write(master_bytes)

    return {
        "brand": brand,
        "cold_start": False,
        "new_stores": new_stores,
        "closed_count": closed_count,
        "closed_stores": closed_stores,
        # 드라이브 저장에 성공했으면 텔레그램으로 따로 파일을 보내지 않으므로 None으로 비워둠
        # (실패했을 때만 유실 방지용으로 파일 자체를 담아 보냄 - _sync_and_notify에서 처리)
        "master_bytes": None if master_saved_to_drive else master_bytes,
        "master_saved_to_drive": master_saved_to_drive,
        "master_drive_file_name": drive_file_name if not drive_master_missing else None,
        "drive_master_missing": drive_master_missing,
        "has_policy_no": bool(policy_no),
        "has_cert_template": _has_cert_template(brand),
        "skipped_placeholder_stores": skipped_placeholder_stores,
        "premium_mismatches": premium_mismatches,
    }


def _is_own_address(addr) -> bool:
    """유진님 본인 메일 주소인지(대소문자 무시). 포워딩/테스트로 보낸 메일의 '보낸 사람'이 본인일 때
    담당자로 잘못 인식하지 않기 위해 씀."""
    if not addr:
        return False
    own = {a.strip().lower() for a in (GMAIL_ADDRESS, MAIL_FROM_ADDRESS, MAIL_REPLY_TO,
                                       "yujin.hwang@daonins.co.kr") if a}
    return str(addr).strip().lower() in own


async def _sync_and_notify(
    bot, chat_id: int, file_bytes: bytes, requester_email: str | None = None
) -> dict | None:
    """정산양식 엑셀을 브랜드 통합파일과 동기화하고, 결과(가입증명서/갱신된 통합파일/요약)를
    텔레그램으로 보냄. 텔레그램에 직접 파일을 올린 경우와, 메일에 첨부된 파일을 자동으로
    감지한 경우 양쪽에서 공용으로 사용함.
    브랜드를 인식하지 못하면(정산양식 형식이 아니면) 아무것도 보내지 않고 None을 반환하며,
    이 경우 어떻게 안내할지는 호출한 쪽에서 정함(메일 첨부는 정산양식이 아닐 수도 있으니
    조용히 넘어가고, 텔레그램 직접 업로드는 사용자에게 안내함).
    requester_email: 메일에 첨부되어 들어온 경우, 그 메일을 보낸 사람 주소. 신규매장
    가입증명서를 만들면 이 주소로도 자동 발송함(텔레그램으로 직접 올린 경우엔 None이라
    대신 담당자 연락처 등록부(contacts.json)에서 '접수자' 이름으로 찾음)."""
    try:
        result = _sync_brand_excel(file_bytes)
    except Exception:
        logger.exception("엑셀 동기화 중 오류")
        # 원인을 로그까지 안 가고도 바로 알 수 있도록, 오류 종류/위치를 같이 알려드림
        # (2026-10-21, 세정 추가가입 정산파일 오류 때 원인을 바로 못 찾아서 추가함)
        import traceback as _tb
        _last = _tb.extract_tb(sys.exc_info()[2])[-1] if sys.exc_info()[2] else None
        _where = f"{_last.name}() {_last.lineno}번째 줄" if _last else "위치 불명"
        await _send_message_retrying(
            bot, chat_id,
            f"⚠️ 엑셀을 처리하는 중 오류가 발생했어요.\n원인: {type(sys.exc_info()[1]).__name__}: "
            f"{str(sys.exc_info()[1])[:200]}\n위치: {_where}",
        )
        return None

    if result is None:
        return None

    brand = result["brand"]

    # 아래부터는 텔레그램 전송에 전부 재시도 래퍼(_send_message_retrying/_send_document_retrying)를
    # 씀. 예전엔 raw bot.send_message/send_document를 직접 호출해서, 텔레그램 쪽 일시적
    # 네트워크 지연(telegram.error.TimedOut/NetworkError) 한 번이면 그 예외가 이 함수를 호출한
    # check_new_mail의 try/except까지 새어나가 그 뒤 단계(통합파일 저장 결과, 담당자 발송 확인
    # 버튼)가 통째로 날아가고 유진님껜 아무 알림도 안 가는 사고가 있었음(2026-10-02, 세정
    # 웰메이드 대구방촌점 건 - Render 로그에서 telegram.error.TimedOut 확인함).
    if result["cold_start"]:
        if result.get("drive_master_missing"):
            await _send_message_retrying(
                bot, chat_id,
                f"📁 구글 드라이브에서 '{brand}' 통합파일을 찾지 못해서, 이번 파일을 서버에 임시 기준으로 "
                f"등록했어요.\n'가입증명서및정산' 폴더 밑에 '정산_{brand}' 폴더가 있고 그 안에 취합본 "
                "엑셀이 있는지 확인해주세요 — 확인되면 다음부터는 그 파일에 바로 반영해드릴게요.",
            )
        else:
            await _send_message_retrying(
                bot, chat_id,
                f"📁 '{brand}' 통합파일을 처음 등록했어요. 앞으로 이 파일을 기준으로 신규/폐점 매장을 비교할게요.",
            )
        return result

    skipped_placeholder_stores = result.get("skipped_placeholder_stores") or []
    if skipped_placeholder_stores:
        names = ", ".join(s["store_name"] for s in skipped_placeholder_stores)
        await _send_message_retrying(
            bot, chat_id,
            f"ℹ️ {names}: 매장코드가 아직 '미개설(추후전달)' 같은 임시 문구인 채로 다시 들어왔는데, "
            "매장명/사업자번호가 이미 등록된 매장과 같아서 중복 등록하지 않고 건너뛰었어요.",
        )

    if not result["new_stores"] and not result["closed_count"]:
        await _send_message_retrying(bot, chat_id, f"'{brand}' 기준으로 새로운 신규/폐점 매장이 없어요.")
        return result

    if not result.get("has_cert_template"):
        await _send_message_retrying(
            bot, chat_id,
            f"⚠️ '{brand}'의 가입증명서 서식이 아직 등록되어 있지 않아, 다른 브랜드 서식으로 임시로 만들어요. "
            "증권번호/계약자 등이 실제와 다를 수 있어요. 이 브랜드로 실제 발급된 예시 인증서를 보내주시면 "
            "전용 서식으로 등록해드릴게요.",
        )

    ended_stores = []
    no_contact_stores = []
    queued_groups = []
    drive_saved = 0
    drive_failed_stores = []
    contacts = _load_contacts()
    # target_email -> {"received_by": str, "items": [{"store_name","out_name","pdf_bytes"}, ...]}
    cert_groups: dict[str, dict] = {}
    for store in result["new_stores"]:
        # 보험종기일이 이미 지난 건은 통합파일엔 기록해두되, 가입증명서는 만들지 않음(사용자 요청)
        if store.get("period_ended"):
            ended_stores.append(store)
            continue
        try:
            pdf_bytes = _build_certificate_pdf(store, brand)
        except Exception:
            logger.exception("가입증명서 생성 중 오류")
            await _send_message_retrying(bot, chat_id, f"⚠️ '{store['store_name']}' 가입증명서 생성에 실패했어요.")
            continue

        out_name = f"{store['store_code']}_{store['store_name']}_{store['start_date_yymmdd']}.pdf"
        # 예전엔 재시도 없는 bot.send_document를 직접 썼는데, 텔레그램 쪽 일시적
        # timeout(telegram.error.TimedOut)이 나면 그 예외가 이 함수를 호출한 check_new_mail의
        # try/except까지 새어나가서, 이미 보낸 증명서 PDF 이후의 모든 단계(통합파일 저장 결과
        # 요약, 담당자 발송 확인 버튼)가 통째로 날아가고 유진님껜 아무 알림도 안 갔음
        # (2026-10-02, 세정 웰메이드 대구방촌점 건 - Render 로그에서 telegram.error.TimedOut
        # 확인함). _send_document_retrying은 최대 3번 재시도하고, 그래도 실패하면 False만
        # 반환해서 예외가 새어나가지 않고 뒷단 로직이 계속 진행되도록 함.
        doc_ok = await _send_document_retrying(
            bot, chat_id, pdf_bytes, out_name,
            caption=f"📄 {store['store_name']} 가입증명서",
        )
        if not doc_ok:
            await _send_message_retrying(
                bot, chat_id,
                f"⚠️ '{store['store_name']}' 가입증명서 PDF를 텔레그램으로 보내는 데 실패했어요 "
                "(통합파일/드라이브 반영과 담당자 발송은 계속 진행할게요).",
            )

        # 담당자에게 이메일로 보낼지: 메일로 들어온 건이면 그 발신 주소로, 텔레그램으로
        # 직접 올린 건이면 엑셀 '접수자' 이름으로 등록된 연락처를 찾음. 같은 담당자 앞으로
        # 신규매장이 여러 곳이면 메일을 여러 통 나눠 보내지 않고, 한 번에 묶어서 보냄.
        # 바로 보내지 않고 유진님이 버튼으로 확인한 뒤에 실제 발송함(원치 않으면 '대기' 선택 가능)
        # 보낸 사람이 유진님 본인 주소(포워딩/테스트 등)면 담당자로 취급하지 않음 - 안 그러면 담당자
        # 연락처가 유진님 주소로 덮어써져서 가입증명서가 유진님께 가버림(2026-10-21 세정 김연희 건)
        sender_for_contact = None if _is_own_address(requester_email) else requester_email
        target_email = sender_for_contact
        received_by = store.get("received_by") or ""
        if not target_email and received_by:
            target_email = contacts.get(_norm_contact_name(received_by))
            if _is_own_address(target_email):
                target_email = None  # 이미 유진님 주소로 잘못 저장된 연락처는 무시
        if target_email and GMAIL_ADDRESS and GMAIL_APP_PASSWORD:
            # 메일로 들어와서 처음 알게 된 담당자면, 다음부터 텔레그램으로 직접 올려도
            # 이메일을 찾을 수 있도록 연락처에 미리 저장해둠(발송 여부와 무관하게 저장)
            if sender_for_contact and received_by:
                key = _norm_contact_name(received_by)
                if contacts.get(key) != sender_for_contact:
                    contacts[key] = sender_for_contact
                    _save_contacts(contacts)
            group = cert_groups.setdefault(target_email, {"received_by": "", "items": []})
            if received_by and not group["received_by"]:
                group["received_by"] = received_by
            group["items"].append({
                "store_name": store["store_name"],
                "out_name": out_name,
                "pdf_bytes": pdf_bytes,
            })
        elif received_by:
            no_contact_stores.append(received_by)

        # 구글 드라이브(설정해뒀으면 유진님 PC에 자동 동기화)에도 저장
        if _upload_to_drive(out_name, pdf_bytes, brand=brand):
            drive_saved += 1
        else:
            drive_failed_stores.append(store["store_name"])

    for target_email, group in cert_groups.items():
        items = group["items"]
        store_names = [item["store_name"] for item in items]
        received_by = group["received_by"]
        if len(items) == 1:
            subject = f"[{brand}] {store_names[0]} 가입증명서"
            body = _CERT_EMAIL_BODY_TEMPLATE.format(store_name=store_names[0])
            confirm_text = f"담당자 {received_by or target_email}님에게 '{store_names[0]}' 가입증명서를 보낼까요?"
        else:
            subject = f"[{brand}] 신규매장 가입증명서 ({len(items)}건)"
            store_list = "\n".join(f"- {name}" for name in store_names)
            body = _CERT_EMAIL_BODY_TEMPLATE_MULTI.format(store_list=store_list)
            names_text = ", ".join(store_names)
            confirm_text = f"담당자 {received_by or target_email}님에게 신규매장 {len(items)}건({names_text}) 가입증명서를 한번에 보낼까요?"
        cert_id = _queue_pending_certs(items, target_email=target_email, subject=subject, body=body)
        queued_groups.append((store_names, target_email))
        keyboard = InlineKeyboardMarkup([[
            InlineKeyboardButton("✅ 보내기", callback_data=f"sendcert:{cert_id}"),
            InlineKeyboardButton("⏸ 대기", callback_data=f"holdcert:{cert_id}"),
            InlineKeyboardButton("🗑 취소", callback_data=f"cancelcert:{cert_id}"),
        ]])
        await _send_message_retrying(
            bot, chat_id,
            confirm_text + "\n"
            "(대기를 누르면 이 메시지는 그대로 남아있으니, 나중에 다시 여기서 보내기를 누르시면 돼요)",
            reply_markup=keyboard,
        )

    if ended_stores:
        lines = "\n".join(f"- {s['store_name']}({s['store_code']}) 종기일 {s['end_date']}" for s in ended_stores)
        await _send_message_retrying(
            bot, chat_id,
            "ℹ️ 아래 매장은 보험종기일이 이미 지나서 통합파일에는 등록했지만 가입증명서는 만들지 않았어요.\n"
            + lines,
        )

    # 아래 "통합파일" 관련 줄은 전부 같은 동작(드라이브 파일에 직접 반영) 하나를 설명하는 거라,
    # "갱신했어요"처럼 모호한 표현과 "드라이브에 반영했어요"를 따로따로 두 번 말해서 마치 별개
    # 단계가 두 번 있는 것처럼 보이지 않도록, 결과를 한 줄로 명확하게 합침(2026-10-02, 유진님
    # 지적 반영). 가입증명서 PDF가 드라이브에 저장됐다는 줄도 통합파일과 헷갈리지 않게 문구를
    # '가입증명서'라고 명시함.
    summary = f"신규 매장 {len(result['new_stores'])}곳"
    if result["closed_count"]:
        summary += f", 폐점 매장 {result['closed_count']}곳"
    summary += " 처리했어요."
    if queued_groups:
        queued_lines = "\n".join(f"  - {', '.join(names)} → {addr}" for names, addr in queued_groups)
        summary += f"\n📧 담당자 발송 확인 대기 중(위 버튼 눌러주세요):\n{queued_lines}"
    if drive_saved:
        summary += f"\n💾 가입증명서 PDF도 구글 드라이브에 {drive_saved}건 저장했어요."
    if drive_failed_stores:
        summary += (
            f"\n❌ 가입증명서 PDF 드라이브 저장 실패: {', '.join(drive_failed_stores)} "
            "(드라이브 연결 설정을 확인해주세요. 텔레그램으로는 정상적으로 받으셨을 거예요)"
        )
    if result.get("master_saved_to_drive"):
        master_file_name = result.get("master_drive_file_name")
        if master_file_name:
            summary += f"\n✅ 통합파일: 드라이브의 '{master_file_name}'에 정산 내용을 바로 반영했어요(PC에도 자동 동기화될 거예요)."
        else:
            summary += "\n✅ 통합파일: 드라이브의 기존 파일에 바로 반영했어요(PC에도 자동 동기화될 거예요)."
    elif result.get("drive_master_missing"):
        summary += (
            f"\n⚠️ 통합파일: 드라이브에서 '{brand}' 통합파일을 못 찾아서 서버에만 임시로 저장했어요. "
            f"'가입증명서및정산' 폴더 밑에 '정산_{brand}' 폴더가 있는지 확인해주세요."
        )
    elif result.get("master_bytes"):
        summary += "\n❌ 통합파일: 드라이브에 저장하는 데 실패해서, 대신 파일로 보내드려요(수동으로 반영해주세요)."
    # 이 요약 메시지(신규/폐점 매장 건수 포함) 전송이 타임아웃 등으로 실패하면, 그 예외가
    # 이 함수 밖(check_new_mail의 try/except)까지 새어나가 return result에 도달하지 못하고
    # 결과가 통째로 유실됨 -> 호출한 쪽의 폐점서류 저장 로직까지 같이 건너뛰게 되는 사고로
    # 이어졌음(2026-09-18, 세정 폐점 건). 이후로는 이 함수의 나머지 전송은 전부 재시도 래퍼를
    # 쓰고, 실패해도 예외를 던지지 않아 반드시 아래 return result까지 도달하도록 함.
    await _send_message_retrying(bot, chat_id, summary)

    # 산출보험료 검증(2026-09-29 유진님 요청): 신규매장 '잔여 합계보험료', 폐점매장
    # '환급보험료'가 봇이 같은 행 값으로 직접 계산한 결과와 다르면 알려드림. 맞을 때는
    # 조용히 그대로 기재만 하고 따로 알리지 않음(정확하면 알림 없음, 틀릴 때만 알림).
    premium_mismatches = result.get("premium_mismatches") or []
    if premium_mismatches:
        lines = "\n".join(
            f"- [{m['kind']}] {m['store_name']}: 파일엔 {m['stated']:,}원, 봇 계산으론 {m['expected']:,}원"
            for m in premium_mismatches
        )
        await _send_message_retrying(
            bot, chat_id,
            "⚠️ 아래 매장은 정산양식에 적힌 보험료가 봇이 같은 행 값(재물/영업배상 항목, "
            "보험기간, 연간총보험료 등)으로 직접 계산한 값과 달라요. 통합파일 쪽 칸은 수식이라 "
            "자동으로 봇 계산값과 같게 나올 거예요 - 정산양식 원본 쪽 숫자나 입력값이 잘못됐을 "
            "수 있으니 확인해주세요.\n" + lines,
        )

    if no_contact_stores:
        names = ", ".join(sorted(set(no_contact_stores)))
        await _send_message_retrying(
            bot, chat_id,
            f"ℹ️ {names} 담당자 이메일이 등록되어 있지 않아 자동 발송을 못 했어요.\n"
            "/setcontact <담당자 이름> <이메일> 로 등록해주시면 다음부터 자동으로 보내드려요.",
        )

    # 담당자가 보낸 파일에 전체 이력이 섞여 있어서, 접수일자가 아주 오래된(지난 정산기간보다도
    # 이전) 매장이 '신규'로 잡힌 경우 - 이미 예전에 다른 방식으로 처리된 건일 수 있으니 등록/증명서
    # 발급은 그대로 진행하되, 놓치지 않도록 별도로 알려드림 (조용히 신규 처리만 하면 나중에 실수로
    # 발견하게 됨)
    stale_stores = [s for s in result["new_stores"] if s.get("stale_recv_date")]
    if stale_stores:
        lines = "\n".join(
            f"- {s['store_name']}({s['store_code']}) 접수일자 {s['stale_recv_date']}"
            for s in stale_stores
        )
        await _send_message_retrying(
            bot, chat_id,
            "⚠️ 아래 매장은 접수일자가 지난 정산기간보다도 오래돼서, 이미 예전에 처리됐는데 "
            "파일에만 이력으로 남아있던 건일 수 있어요. 확인해주세요(맞으면 그냥 두시면 되고, "
            "이미 처리된 거면 알려주시면 통합파일에서 빼드릴게요).\n" + lines,
        )

    # master_bytes는 드라이브 저장에 성공했으면 None(더 이상 텔레그램으로 통합파일을 따로
    # 안 보냄, 2026-09-29 유진님 요청) - 드라이브에서 못 찾았거나 저장이 실패했을 때만
    # 데이터 유실 방지용으로 파일 자체를 보내둠.
    if result.get("master_bytes"):
        display_name = _display_name(brand)
        doc_ok = await _send_document_retrying(
            bot, chat_id, result["master_bytes"], f"{display_name}.xlsx",
            caption=f"📎 '{display_name}' 통합파일이에요(드라이브에 반영을 못 해서 대신 보내드려요 - 확인 후 수동으로 반영해주세요).",
        )
        if not doc_ok:
            await _send_message_retrying(
                bot, chat_id,
                f"⚠️ '{display_name}' 통합파일을 텔레그램으로 전송하는 데도 실패했어요 "
                "(파일 자체는 서버에 임시로 저장돼 있어요). 다시 필요하시면 말씀해주세요.",
            )

    return result


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return

    draft = context.user_data.get("mail_draft")
    if not (draft and draft.get("state") == "attach"):
        if context.user_data.get("closure_buffers") or context.user_data.get("awaiting_closure_store_for"):
            await update.message.reply_text(
                "폐점서류는 사진 말고 '파일(문서)'로 보내주세요 — 사진으로 보내면 파일명이 없어져서 "
                "어떤 서류인지 구분을 못 해요. (텔레그램에서 첨부 아이콘 → 파일 선택으로 보내시면 돼요)"
            )
            return
        await update.message.reply_text(
            "사진은 메일 작성 중 '첨부할 파일' 단계에서만 첨부할 수 있어요. /mail로 메일 작성을 시작해보세요."
        )
        return

    photo = update.message.photo[-1]  # 가장 고화질
    try:
        tg_file = await photo.get_file()
        file_bytes = bytes(await tg_file.download_as_bytearray())
    except Exception:
        logger.exception("사진 다운로드 중 오류")
        await update.message.reply_text("사진을 받아오는 중 오류가 발생했어요. 다시 보내주세요.")
        return

    n = len(draft.get("attachments", [])) + 1
    filename = f"사진_{n}.jpg"
    draft.setdefault("attachments", []).append((filename, file_bytes))
    await update.message.reply_text(
        f"📎 첨부됨: {filename} (총 {n}개)\n계속 보내시거나, 다 되셨으면 '없음'/'완료'라고 입력해주세요."
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return

    doc = update.message.document
    filename = doc.file_name or "attachment"

    # 메일 작성 중 첨부파일 단계면, 파일 종류 상관없이 메일 첨부로 처리
    draft = context.user_data.get("mail_draft")
    if draft and draft.get("state") == "attach":
        try:
            tg_file = await doc.get_file()
            file_bytes = bytes(await tg_file.download_as_bytearray())
        except Exception:
            logger.exception("첨부파일 다운로드 중 오류")
            await update.message.reply_text("파일을 받아오는 중 오류가 발생했어요. 다시 보내주세요.")
            return
        draft.setdefault("attachments", []).append((filename, file_bytes))
        n = len(draft["attachments"])
        await update.message.reply_text(
            f"📎 첨부됨: {filename} (총 {n}개)\n계속 보내시거나, 다 되셨으면 '없음'/'완료'라고 입력해주세요."
        )
        return

    # ㈜세정 x 처브 업무(2026-11-12~): 변경신청서 엑셀 또는 처브 증명서 PDF를 직접 올린 경우.
    # (예전 정산양식 동기화/가입증명서 생성/폐점서류 수집 로직은 껐음)
    lower = filename.lower()
    if not lower.endswith((".xlsx",) + _SJ_DOC_EXTS):
        await update.message.reply_text(
            "㈜세정 변경신청서(엑셀/PDF), 처브 증명서(PDF), 해지 서류(사업자등록증·통장사본, PDF/사진)만 처리할 수 있어요."
        )
        return
    try:
        tg_file = await doc.get_file()
        file_bytes = bytes(await tg_file.download_as_bytearray())
    except Exception:
        logger.exception("파일 다운로드 중 오류")
        await update.message.reply_text("파일을 받아오는 중 오류가 발생했어요. 다시 보내주세요.")
        return

    # 여러 파일을 한 번에 올렸거나(앨범) 설명에 '보내줘' 같은 지시가 있으면, 묶음을 다 받은 뒤
    # 한꺼번에 처리함(2026-11-24: '9개 증명서를 서석란 과장한테 한번에 보내줘' 지시를 못 알아듣던 문제)
    mg = update.message.media_group_id
    caption = (update.message.caption or "").strip()
    if mg or (caption and _SJ_SEND_RX.search(caption)):
        key = str(mg or f"single-{update.message.message_id}")
        groups = context.application.bot_data.setdefault("sj_upload_groups", {})
        new_group = key not in groups
        grp = groups.setdefault(key, {"files": [], "captions": []})
        grp["files"].append((filename, file_bytes))
        grp["last"] = time.monotonic()
        if caption:
            grp["captions"].append(caption)
        if new_group:
            context.job_queue.run_once(_sj_flush_upload_group, 4, data=key, name=f"sjgrp-{key}")
        return
    await _sj_process_uploaded_file(context.bot, filename, file_bytes)


async def _sj_process_uploaded_file(bot, filename: str, file_bytes: bytes) -> None:
    """텔레그램에 올린 파일 1개 처리(신청서 / 처브 증명서 / 해지 서류). handle_document에서 떼어냄(2026-11-24)."""
    lower = filename.lower()

    async def reply(text: str) -> None:
        await _send_message_retrying(bot, ALLOWED_USER_ID, text)

    try:
        if lower.endswith(".xlsx") or _sj_pdf_is_request(_extract_pdf_text(file_bytes)):
            # 변경신청서(엑셀/PDF)를 직접 올린 경우도 메일과 똑같이 처리(리스트 기재+드라이브 저장+처브 발송 확인)
            handled = await _sj_process_request(
                bot, [(filename, file_bytes)], filename, "", None, None, source="텔레그램 업로드"
            )
            if not handled:
                await reply("처브 가입·해지 신청서 양식이 아니라서 처리하지 않았어요.")
        else:
            handled = lower.endswith(".pdf") and await _sj_handle_cert_file(bot, filename, file_bytes)
            if not handled:
                # 해지 관련 서류(사업자등록증/통장사본 등)면 가장 최근 해지신청 매장 기준으로 저장
                category = await _classify_closure_doc_deep(filename, file_bytes)
                if category not in _SJ_CLOSURE_DOC_TYPES:
                    category = None
                last = _sj_load(SEJUNG_LAST_CLOSE_FILE, {})
                recent = False
                try:
                    recent = (dt.datetime.now(KST) - dt.datetime.fromisoformat(last.get("at"))).days <= 14
                except Exception:
                    pass
                if category and last.get("stores") and recent:
                    msg = await _sj_save_closure_docs(last["brand"], last["stores"], [(filename, file_bytes)])
                    await reply(
                        (msg or "저장할 서류가 없어요.") +
                        f"\n(가장 최근 해지신청 매장 기준: {last['brand']} {', '.join(last['stores'])})"
                    )
                elif category:
                    await reply(
                        f"'{filename}'은 {category}로 보이는데, 어느 매장 건인지 모르겠어요. "
                        "해지신청서를 먼저 올려주시면 그 매장 이름으로 저장해드릴게요."
                    )
                else:
                    await reply(
                        f"'{filename}': 변경신청서, 처브 증명서, 해지 서류(사업자등록증·통장사본) 중 "
                        "어느 것으로도 보이지 않아서 처리하지 않았어요."
                    )
    except Exception as e:
        logger.exception("세정 파일 처리 중 오류")
        await reply(f"⚠️ '{filename}' 처리 중 오류가 발생했어요: {type(e).__name__} ({_sj_err_where()}): {str(e)[:150]}")


def _send_email(to_addr: str, subject: str, body: str, attachments: list | None = None,
                extra_headers: dict | None = None) -> None:
    """attachments: [(파일명, 파일바이트), ...]
    extra_headers: 답장으로 보낼 때 In-Reply-To/References 등(받는 쪽 메일함에서 같은 대화로 묶임)."""
    if attachments:
        msg = MIMEMultipart()
        msg.attach(MIMEText(body))
        for filename, file_bytes in attachments:
            part = MIMEApplication(file_bytes, Name=filename)
            part["Content-Disposition"] = f'attachment; filename="{filename}"'
            msg.attach(part)
    else:
        msg = MIMEText(body)

    msg["Subject"] = subject
    msg["From"] = MAIL_FROM_ADDRESS
    msg["To"] = to_addr
    if MAIL_REPLY_TO:
        # 보낸사람 계정은 그대로 두고, 상대방이 "답장"을 누르면 유진님 실제 업무 메일로
        # 가도록 함(구글 별칭 인증 없이도 바로 적용됨)
        msg["Reply-To"] = MAIL_REPLY_TO
    for hk, hv in (extra_headers or {}).items():
        if hv:
            msg[hk] = hv

    # 로그인은 항상 GMAIL_ADDRESS 계정으로 하고, 보낸사람 표시만 MAIL_FROM_ADDRESS로 바뀜
    # (구글의 "다른 이메일 주소로 보내기" 별칭 인증이 되어 있어야 함)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        recipients = [a.strip() for a in re.split(r"[,;]", to_addr) if a.strip()]
        server.sendmail(GMAIL_ADDRESS, recipients, msg.as_string())


_gdrive_service = None  # 매번 새로 로그인하지 않도록 한 번 만든 클라이언트를 재사용


def _get_gdrive_service():
    """구글 드라이브 클라이언트를 만들어서 반환함(설정 안 돼있으면 None).
    서비스 계정은 개인 '내 드라이브'에 저장 용량이 없어서 못 쓰고(storageQuotaExceeded),
    유진님 개인 계정(bubal5503@gmail.com)의 OAuth refresh token으로 인증함 - 그 계정의
    드라이브 용량을 그대로 씀. refresh token은 처음 한 번만 발급받아 Render 환경변수에
    저장해두면, access token은 만료될 때마다 이 코드가 자동으로 갱신함(재로그인 불필요)."""
    global _gdrive_service
    if _gdrive_service is not None:
        return _gdrive_service
    if not (
        GDRIVE_OAUTH_CLIENT_ID and GDRIVE_OAUTH_CLIENT_SECRET and GDRIVE_OAUTH_REFRESH_TOKEN
        and GDRIVE_FOLDER_ID and _gdrive_build
    ):
        return None
    try:
        creds = _gdrive_user_credentials(
            token=None,
            refresh_token=GDRIVE_OAUTH_REFRESH_TOKEN,
            client_id=GDRIVE_OAUTH_CLIENT_ID,
            client_secret=GDRIVE_OAUTH_CLIENT_SECRET,
            token_uri="https://oauth2.googleapis.com/token",
            # 'drive.file'(이 앱이 직접 만든 파일만 접근)로 잠깐 바꿔봤으나, 유진님이 이미
            # 여러 단계 깊이로 만들어두신 기존 폴더 구조(가입증명서/세정(1112)/25년/...)에는
            # 이 권한으로 접근이 안 돼서(이 앱이 만들지 않은 폴더라 안 보임) 엉뚱한 곳에 새
            # 폴더가 생기는 문제가 있었음. 기존 폴더 구조를 그대로 쓰는 게 더 중요하므로
            # 'drive'(전체 접근)로 되돌림(2026-09-22).
            # scopes는 일부러 지정하지 않음: refresh token 자체에 이미 발급 당시의 권한
            # 범위가 저장돼 있어서, refresh 요청에 scopes를 같이 넘기면 구글 서버가 이를
            # "다시 검증해야 할 요청"으로 보고 invalid_scope로 거부하는 경우가 있었음
            # (실제로 새로 발급받은 refresh token으로도 계속 invalid_scope가 재현됨,
            # 2026-09-22). scopes를 안 넘기면 refresh token에 저장된 범위를 그대로 씀.
        )
        creds.refresh(_gdrive_auth_request())
        _gdrive_service = _gdrive_build("drive", "v3", credentials=creds)
        return _gdrive_service
    except Exception:
        logger.exception("구글 드라이브 연결 중 오류")
        return None


def _gdrive_find_child_folder(service, parent_id: str, predicate) -> str | None:
    """parent_id 밑에 있는 폴더들 중 predicate(name)이 True인 첫 폴더의 id를 찾아 반환.
    없으면 None."""
    query = (
        f"'{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' "
        "and trashed=false"
    )
    resp = service.files().list(q=query, fields="files(id,name)", pageSize=200).execute()
    for f in resp.get("files", []):
        if predicate(f["name"]):
            return f["id"]
    return None


def _gdrive_find_file(service, parent_id: str, predicate) -> dict | None:
    """parent_id 밑에 있는 파일들(폴더 제외) 중 predicate(name)이 True인 첫 파일의
    {"id","name"}을 반환. 없으면 None. 폐점서류 zip처럼 '이미 올려둔 파일이 있으면 그걸
    이어서 채워야 하는' 경우에 씀."""
    query = (
        f"'{parent_id}' in parents and mimeType!='application/vnd.google-apps.folder' "
        "and trashed=false"
    )
    resp = service.files().list(q=query, fields="files(id,name)", pageSize=200).execute()
    for f in resp.get("files", []):
        if predicate(f["name"]):
            return f
    return None


def _gdrive_download_file(service, file_id: str) -> bytes | None:
    if _gdrive_media_download is None:
        return None
    try:
        request = service.files().get_media(fileId=file_id)
        buf = io.BytesIO()
        downloader = _gdrive_media_download(buf, request)
        done = False
        while not done:
            _status, done = downloader.next_chunk()
        return buf.getvalue()
    except Exception:
        logger.exception("구글 드라이브 파일 다운로드 중 오류")
        return None


def _gdrive_create_folder(service, parent_id: str, name: str) -> str:
    meta = {
        "name": name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_id],
    }
    folder = service.files().create(body=meta, fields="id").execute()
    return folder["id"]


def _gdrive_period_prefix(period_start: dt.date, period_end: dt.date) -> str:
    """유진님이 실제로 쓰시는 정산기간 폴더 이름 규칙: 'MM15_MM15' (예: 8/16~9/15 기간은
    '0815_0915'). 시작월의 15일 + 종료월의 15일을 이어붙인 형태(시작일 16일 자체가 아님)."""
    return f"{period_start.month:02d}15_{period_end.month:02d}15"


def _resolve_drive_target_folder(service, brand: str, extra_subfolder: str | None = None) -> str | None:
    """GDRIVE_FOLDER_ID(예: '가입증명서및정산' 폴더) 밑에서 이번 정산기간 폴더를 찾고(없으면
    '유진님 최근 사용 방식'대로 새로 만듦), 그 안에서 브랜드 폴더를 찾아(없으면 새로 만들어서)
    최종적으로 파일을 올릴 폴더의 id를 반환함. extra_subfolder가 주어지면 그 브랜드 폴더
    밑에 한 단계 더 들어가서(예: '폐점서류') 그 폴더를 찾거나 만듦."""
    period_start, period_end = _current_period()
    prefix = _gdrive_period_prefix(period_start, period_end)
    period_folder_id = _gdrive_find_child_folder(
        service, GDRIVE_FOLDER_ID, lambda name: name.startswith(prefix)
    )
    if period_folder_id is None:
        period_folder_id = _gdrive_create_folder(
            service, GDRIVE_FOLDER_ID, f"{prefix}가입증명서및변경신청서"
        )
    brand_folder_id = _gdrive_find_child_folder(
        service, period_folder_id, lambda name: name == brand or name in brand or brand in name
    )
    if brand_folder_id is None:
        brand_folder_id = _gdrive_create_folder(service, period_folder_id, brand)
    if not extra_subfolder:
        return brand_folder_id
    sub_folder_id = _gdrive_find_child_folder(
        service, brand_folder_id, lambda name: name == extra_subfolder
    )
    if sub_folder_id is None:
        sub_folder_id = _gdrive_create_folder(service, brand_folder_id, extra_subfolder)
    return sub_folder_id


_XLSX_MIMETYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _gdrive_find_brand_master_folder(service, brand: str) -> str | None:
    """GDRIVE_FOLDER_ID(가입증명서및정산) 밑에서 유진님이 이미 만들어두신 '정산_{brand}'
    폴더(예: '정산_올리비아로렌')를 찾아 id를 반환함. 없으면 None(새로 만들지 않음 - 유진님이
    실제 쓰시는 폴더 구조가 아닌 곳에 엉뚱하게 새로 만드는 사고를 막기 위함).
    다른 브랜드 폴더 매칭(_resolve_drive_target_folder)과 마찬가지로 완전히 똑같은 이름이
    아니어도(앞뒤 공백, 유니코드 조합 차이 등) 찾을 수 있도록 느슨하게 비교함 - 완전 일치만
    보면 눈에는 똑같아 보이는데 못 찾는 경우가 있었음(2026-09-29, '트레몰로' 폴더에서 확인됨)."""
    folder_name = f"정산_{brand}"
    target = unicodedata.normalize("NFC", folder_name).strip()

    def _matches(name: str) -> bool:
        norm = unicodedata.normalize("NFC", name).strip()
        return norm == target or target in norm or norm in target

    found = _gdrive_find_child_folder(service, GDRIVE_FOLDER_ID, _matches)
    if found is None:
        logger.warning("드라이브에서 '%s' 폴더를 못 찾음(GDRIVE_FOLDER_ID 밑 하위 폴더 이름과 비교함)", folder_name)
    return found


def _gdrive_find_latest_master_file(service, folder_id: str) -> dict | None:
    """그 폴더 안의 엑셀 파일들 중, 열려있을 때 생기는 임시 잠금 파일('~$'로 시작)은 제외하고
    가장 최근에 수정된 파일 하나를 {"id","name"}로 반환. 유진님이 파일명을 그때그때 바꿔가며
    (예: '..._260918.xlsx') 쓰시므로 이름이 아니라 수정 시각 기준으로 '현재 쓰는 파일'을 고름."""
    query = f"'{folder_id}' in parents and trashed=false and mimeType='{_XLSX_MIMETYPE}'"
    resp = service.files().list(
        q=query, fields="files(id,name,modifiedTime)", orderBy="modifiedTime desc", pageSize=50
    ).execute()
    for f in resp.get("files", []):
        if not f["name"].startswith("~$"):
            return f
    return None


def _load_brand_master_from_drive(brand: str) -> tuple[bytes | None, str | None, str | None]:
    """브랜드의 '정산_{brand}' 드라이브 폴더에서 현재 쓰는 통합파일을 찾아 내용을 내려받음.
    반환: (파일 내용 bytes, 드라이브 file id, 파일명). 못 찾으면 (None, None, None)."""
    service = _get_gdrive_service()
    if service is None:
        return None, None, None
    try:
        folder_id = _gdrive_find_brand_master_folder(service, brand)
        if folder_id is None:
            return None, None, None
        f = _gdrive_find_latest_master_file(service, folder_id)
        if f is None:
            return None, None, None
        content = _gdrive_download_file(service, f["id"])
        if content is None:
            return None, None, None
        return content, f["id"], f["name"]
    except Exception:
        logger.exception("드라이브에서 '%s' 통합파일을 찾는 중 오류", brand)
        return None, None, None


def _list_drive_brand_folders(service) -> dict:
    """GDRIVE_FOLDER_ID('가입증명서및정산') 밑의 '정산_브랜드명' 폴더를 전부 찾아
    {브랜드명: 폴더ID} 매핑으로 반환함. 가입증명서 재검색/담당자 재발송처럼 '등록된 모든
    브랜드의 통합파일'을 뒤져야 하는 기능들이 어떤 브랜드가 있는지 미리 알지 못해도 되게 함."""
    query = (
        f"'{GDRIVE_FOLDER_ID}' in parents and trashed=false "
        "and mimeType='application/vnd.google-apps.folder'"
    )
    resp = service.files().list(q=query, fields="files(id,name)", pageSize=200).execute()
    result = {}
    for f in resp.get("files", []):
        name = unicodedata.normalize("NFC", f["name"]).strip()
        if name.startswith("정산_"):
            result[name[len("정산_"):]] = f["id"]
    return result


def _save_brand_master_to_drive(file_id: str, file_bytes: bytes) -> bool:
    """찾아둔 그 파일(file_id)을 그대로 덮어써서 저장함(새 파일을 만들지 않음 - 유진님 PC의
    구글 드라이브 동기화 폴더에 있는 바로 그 파일이 갱신되게 하기 위함)."""
    service = _get_gdrive_service()
    if service is None:
        return False
    try:
        media = _gdrive_media(file_bytes, mimetype=_XLSX_MIMETYPE)
        service.files().update(fileId=file_id, media_body=media).execute()
        return True
    except Exception:
        logger.exception("통합파일을 드라이브에 저장하는 중 오류")
        return False


def _upload_to_drive(
    filename: str,
    file_bytes: bytes,
    brand: str | None = None,
    extra_subfolder: str | None = None,
    mimetype: str = "application/pdf",
) -> bool:
    """가입증명서(또는 폐점서류 압축파일 등)를 유진님이 지정해둔 구글 드라이브 폴더에 올림.
    brand가 주어지면 '이번 정산기간 폴더 / 브랜드 폴더' 안에(extra_subfolder까지 주어지면
    그 밑의 하위 폴더 안에), 아니면 지정해둔 폴더 바로 밑에 올림. 그 폴더가 유진님 PC의
    구글 드라이브 동기화 폴더 안에 있으면, PC에도 자동으로 똑같이 저장됨. 설정이 안 돼
    있거나 실패하면 조용히 False만 반환(가입증명서 발급 자체는 그대로 진행돼야 하므로)."""
    service = _get_gdrive_service()
    if service is None:
        return False
    try:
        parent_id = GDRIVE_FOLDER_ID
        if brand:
            resolved = _resolve_drive_target_folder(service, brand, extra_subfolder=extra_subfolder)
            if resolved:
                parent_id = resolved
        media = _gdrive_media(file_bytes, mimetype=mimetype)
        service.files().create(
            body={"name": filename, "parents": [parent_id]},
            media_body=media,
            fields="id",
        ).execute()
        return True
    except Exception:
        logger.exception("구글 드라이브 업로드 중 오류")
        return False


async def test_mail_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """실제 매장 데이터 없이, 가짜 샘플 매장으로 가입증명서 이메일 발송을 미리 테스트해봄
    (문구/서명/첨부가 실제로 어떻게 보이는지 확인용). 사용법: /testmail <받는이메일>
    (안 적으면 유진님 실제 메일(MAIL_REPLY_TO)로 감)"""
    if not is_allowed(update):
        return
    if not GMAIL_ADDRESS or not GMAIL_APP_PASSWORD:
        await update.message.reply_text("메일 발송 기능이 설정되어 있지 않아요.")
        return
    to_addr = context.args[0] if context.args else MAIL_REPLY_TO
    if not to_addr:
        await update.message.reply_text("사용법: /testmail <받는이메일>\n예: /testmail yujin.hwang@daonins.co.kr")
        return

    sample = {
        "policy_no": DEFAULT_POLICY_NO,
        "store_code": "TEST0000",
        "store_name": "테스트매장",
        "address": "서울특별시 서초구 강남대로 375 서초현대타워 (테스트용 샘플 주소)",
        "stock_amt": 100000000,
        "facility_amt": 10000000,
        "building_amt": 0,
        "premium": 12345,
        "start_date": "2026.01.01",
        "end_date": "2026.12.31",
        "start_date_yymmdd": "260101",
        "has_property": True,
        "has_liability": True,
    }
    try:
        pdf_bytes = _build_certificate_pdf(sample, "트레몰로")
    except Exception:
        logger.exception("테스트 가입증명서 생성 중 오류")
        await update.message.reply_text("⚠️ 테스트 가입증명서 생성에 실패했어요.")
        return

    out_name = "TEST0000_테스트매장_260101.pdf"
    body = _CERT_EMAIL_BODY_TEMPLATE.format(store_name=sample["store_name"])
    try:
        _send_email(
            to_addr,
            subject="[테스트] 테스트매장 가입증명서",
            body=body,
            attachments=[(out_name, pdf_bytes)],
        )
        await update.message.reply_text(f"✅ 테스트 이메일을 {to_addr}로 보냈어요. 메일함에서 확인해보세요.")
    except Exception:
        logger.exception("테스트 이메일 발송 중 오류")
        await update.message.reply_text("⚠️ 테스트 이메일 발송에 실패했어요.")


async def mail_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        return

    if not GMAIL_ADDRESS or not GMAIL_APP_PASSWORD:
        await update.message.reply_text("메일 발송 기능이 설정되어 있지 않아요.")
        return

    if not context.args:
        await update.message.reply_text(
            "사용법: /mail <받는사람이메일>\n예: /mail friend@example.com"
        )
        return

    to_addr = context.args[0]
    context.user_data["mail_draft"] = {"to": to_addr, "state": "subject"}
    await update.message.reply_text(
        f"보낸사람: {MAIL_FROM_ADDRESS}\n받는사람: {to_addr}\n메일 제목을 입력해주세요."
    )


def _collect_store_rows_from_workbook(result: dict, brand: str, wb, wb_values, sheet_type: str) -> None:
    """워크북 하나에서 지정한 종류(예: '신규매장'/'폐점매장') 시트를 뒤져서 result(매장명 ->
    행 정보 목록)에 누적함. 드라이브에서 읽은 워크북/로컬 백업 워크북 양쪽에서 공용으로 씀."""
    sheets = _find_type_sheets(wb, sheet_type)
    sheets_values = _find_type_sheets(wb_values, sheet_type)
    for sub_type, ws in sheets.items():
        ws_values = sheets_values.get(sub_type)
        if ws_values is None:
            continue
        header, min_row = _build_header_map(ws)
        if not header:
            continue
        rate1 = _extract_rate(ws, header, "연간재물보험료", r"\*([\d.]+)%", 0.0665, min_row)
        rate2 = _extract_rate(ws, header, "연간영업배상보험료", r"\*([\d.]+)", 1793, min_row)
        for vals in _extract_data_rows(ws_values, header, min_row, formula_ws=ws, wb_values=wb_values):
            name_raw = str(vals.get("매장명") or "")
            for seg in re.split(r"\n+", name_raw):
                seg = seg.strip()
                if not seg:
                    continue
                result.setdefault(seg, []).append({
                    "brand": brand,
                    "sub_type": sub_type,
                    "vals": vals,
                    "rate1": rate1,
                    "rate2": rate2,
                })


def _all_store_rows_by_sheet_type(sheet_type: str) -> dict:
    """등록된 모든 브랜드의 통합파일에서 지정한 종류(예: '신규매장'/'폐점매장') 시트를 다 뒤져서,
    매장명 -> 그 매장이 있던 행(들) 정보를 모아둠. 통합파일이 유일한 정보 출처가 되도록,
    신규매장 조회/폐점서류 매칭 둘 다 이 함수 하나로 통일함.
    예전엔 서버 로컬(MASTERS_DIR)에 있는 파일만 뒤졌는데, 정산 로직이 드라이브를 기준으로
    바뀐 뒤(2026-10-02)로는 통합파일이 대부분 드라이브에만 있고 서버엔 없어서, '가입증명서
    찾아줘' 같은 기능이 최신 매장을 못 찾는 문제가 있었음. 이제 드라이브를 먼저 뒤지고,
    드라이브에서 못 찾은 브랜드(연결 끊김 등)만 서버 로컬 백업으로 보충함."""
    result: dict[str, list[dict]] = {}
    seen_brands = set()

    service = _get_gdrive_service()
    if service is not None:
        try:
            brand_folders = _list_drive_brand_folders(service)
        except Exception:
            logger.exception("드라이브 브랜드 폴더 목록 조회 중 오류")
            brand_folders = {}
        for brand, folder_id in brand_folders.items():
            try:
                f = _gdrive_find_latest_master_file(service, folder_id)
                if f is None:
                    continue
                content = _gdrive_download_file(service, f["id"])
                if content is None:
                    continue
                wb = openpyxl.load_workbook(io.BytesIO(content), data_only=False)
                wb_values = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
            except Exception:
                logger.exception("드라이브에서 '%s' 통합파일을 읽는 중 오류", brand)
                continue
            seen_brands.add(brand)
            _collect_store_rows_from_workbook(result, brand, wb, wb_values, sheet_type)

    if os.path.isdir(MASTERS_DIR):
        for fname in os.listdir(MASTERS_DIR):
            if not fname.lower().endswith(".xlsx"):
                continue
            brand = fname[:-5]
            if brand in seen_brands:
                continue  # 드라이브에서 이미 찾은 브랜드는 서버에 남아있는 예전 사본을 쓰지 않음
            path = os.path.join(MASTERS_DIR, fname)
            try:
                wb = openpyxl.load_workbook(path, data_only=False)
                wb_values = openpyxl.load_workbook(path, data_only=True)
            except Exception:
                continue
            _collect_store_rows_from_workbook(result, brand, wb, wb_values, sheet_type)

    return result


def _all_registered_store_rows() -> dict:
    """등록된 모든 브랜드의 '신규매장' 시트를 다 뒤져서, 매장명 -> 그 매장이 있던 행(들) 정보를
    모아둠. '가입증명서를 다시 만들어줘' 같은 요청이 오면 여기서 검색해서 다시 만듦."""
    return _all_store_rows_by_sheet_type("신규매장")


def _all_closed_store_rows() -> dict:
    """등록된 모든 브랜드의 '폐점매장' 시트를 다 뒤져서, 매장명 -> 그 매장이 있던 행(들) 정보를
    모아둠. 그 매장이 '신규매장'으로 기록된 적이 없어도(봇 도입 전부터 있던 매장 등), 폐점
    처리되는 순간 통합파일의 '폐점매장' 시트에 이미 기록되므로, 폐점서류를 매장에 매칭시킬
    때(브랜드/대표자명 확인용) 여기서 찾음. 통합파일 자체가 유일한 정보 출처임(별도 저장 안 함)."""
    return _all_store_rows_by_sheet_type("폐점매장")


def _all_initial_list_store_rows() -> dict:
    """등록된 모든 브랜드의 '최초가입전체리스트' 시트를 다 뒤져서, 매장명 -> 그 매장이 있던
    행(들) 정보를 모아둠. 봇이 신규/폐점매장으로 한 번도 기록한 적 없는(정말 오래 전부터
    있던) 매장도 이 시트에는 원래 다 있는 경우가 많아서(유진님 요청), 폐점서류 매칭 시
    마지막 대체 조회 소스로 씀."""
    return _all_store_rows_by_sheet_type("최초가입전체리스트")


_CERT_LOOKUP_RE = re.compile(r"가입\s*증명서")
_CERT_LOOKUP_VERBS = ("찾아", "다시", "재발급", "재전송", "보내줘", "보내주세요", "올려줘", "올려주세요", "보여줘")
_CERT_SEND_TO_CONTACT_KEYWORDS = ("담당자", "메일로", "이메일로")


def _match_registered_store_in_text(user_text: str, store_index: dict):
    """문장 안에 등록된 매장명이 포함돼 있는지 확인. 여러 매장명이 겹쳐 걸리면(예: '점'으로
    끝나는 짧은 이름이 긴 이름의 일부인 경우) 가장 긴 이름을 우선함. 같은 이름으로 여러 건
    등록돼 있으면(갱신/재등록 등) 접수일자가 가장 최근인 걸 고름.
    (매장명, 선택된 entry) 튜플을 반환하거나, 못 찾으면 (None, None)."""
    matches = [name for name in store_index if name and name in user_text]
    if not matches:
        return None, None
    matches.sort(key=len, reverse=True)
    store_name = matches[0]
    entries = store_index[store_name]

    def _recv_key(entry):
        return _parse_date_val(entry["vals"].get("접수일자")) or dt.date.min

    entries.sort(key=_recv_key, reverse=True)
    return store_name, entries[0]


def _build_cert_from_registered_entry(chosen: dict) -> tuple[dict, bytes]:
    """이미 통합파일에 등록된 행(entry) 정보로 가입증명서 PDF를 다시 만듦.
    (store dict, pdf bytes) 튜플을 반환함."""
    vals = chosen["vals"]
    cert_vals = _compute_new_store_cert_values(vals, chosen["rate1"], chosen["rate2"])
    address = str(vals.get("매장주소") or "").strip()
    address = re.sub(r"\s*\n\s*", " ", address)
    address = re.sub(r"(?<=\S)\(", " (", address)
    policy_numbers = _load_policy_numbers()
    store = {
        "policy_no": policy_numbers.get(chosen["brand"]) or DEFAULT_POLICY_NO,
        "store_code": str(vals.get("매장코드") or "").strip(),
        "store_name": str(vals.get("매장명") or "").strip(),
        "address": address,
        **cert_vals,
    }
    pdf_bytes = _build_certificate_pdf(store, chosen["brand"])
    return store, pdf_bytes


async def _handle_cert_send_to_contact_request(update: Update, context: ContextTypes.DEFAULT_TYPE, user_text: str) -> bool:
    """'대구방촌점 담당자한테 보내줘'처럼, 이미 통합파일에 등록된 매장의 가입증명서를
    담당자에게 이메일로 보내달라는 요청을 감지해서 처리함. 정산파일 처리 흐름(_sync_and_notify)
    때와 똑같이 템플릿 메일을 만들고 '보내기/대기/취소' 버튼으로 확인을 받음 - 원래는 그
    흐름 중에 자동으로 뜨는 버튼인데, 중간에 텔레그램 타임아웃 등으로 끊겨서 못 받았을 때
    이 매장만 다시 트리거하는 용도(2026-10-02, 세정 웰메이드 대구방촌점 건).
    _handle_cert_lookup_request보다 먼저 확인해야 함 - '보내줘'가 양쪽 다 걸리는 verb라서,
    '담당자'/'메일로' 같은 키워드가 있으면 이쪽을 우선함."""
    if not _CERT_LOOKUP_RE.search(user_text):
        return False
    if not any(k in user_text for k in _CERT_SEND_TO_CONTACT_KEYWORDS):
        return False
    if not (GMAIL_ADDRESS and GMAIL_APP_PASSWORD):
        return False  # 메일 발송 자체가 설정 안 돼 있으면, 평소처럼 가입증명서 재전송 처리로 넘어가게 둠

    store_index = _all_registered_store_rows()
    if not store_index:
        return False
    store_name, chosen = _match_registered_store_in_text(user_text, store_index)
    if chosen is None:
        return False

    vals = chosen["vals"]
    brand = chosen["brand"]
    received_by = str(vals.get("접수자") or "").strip()
    contacts = _load_contacts()
    target_email = contacts.get(_norm_contact_name(received_by)) if received_by else None
    if not target_email:
        await update.message.reply_text(
            f"'{store_name}' 담당자({received_by or '이름 미확인'}) 이메일이 등록되어 있지 않아요. "
            "/setcontact <담당자 이름> <이메일> 로 등록해주시면 보내드릴게요."
        )
        return True

    try:
        store, pdf_bytes = _build_cert_from_registered_entry(chosen)
    except Exception:
        logger.exception("담당자 발송용 가입증명서 재생성 중 오류")
        await update.message.reply_text(f"⚠️ '{store_name}' 가입증명서를 다시 만드는 중 오류가 발생했어요.")
        return True

    out_name = f"{store['store_code']}_{store['store_name']}_{store['start_date_yymmdd']}.pdf"
    subject = f"[{brand}] {store['store_name']} 가입증명서"
    body = _CERT_EMAIL_BODY_TEMPLATE.format(store_name=store["store_name"])
    cert_id = _queue_pending_certs(
        [{"store_name": store["store_name"], "out_name": out_name, "pdf_bytes": pdf_bytes}],
        target_email=target_email, subject=subject, body=body,
    )
    keyboard = InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ 보내기", callback_data=f"sendcert:{cert_id}"),
        InlineKeyboardButton("⏸ 대기", callback_data=f"holdcert:{cert_id}"),
        InlineKeyboardButton("🗑 취소", callback_data=f"cancelcert:{cert_id}"),
    ]])
    await update.message.reply_text(
        f"담당자 {received_by or target_email}님에게 '{store['store_name']}' 가입증명서를 보낼까요?\n"
        "(대기를 누르면 이 메시지는 그대로 남아있으니, 나중에 다시 여기서 보내기를 누르시면 돼요)",
        reply_markup=keyboard,
    )
    return True


async def _handle_cert_lookup_request(update: Update, context: ContextTypes.DEFAULT_TYPE, user_text: str) -> bool:
    """'계산점 가입증명서 찾아줘'처럼, 이미 등록된 매장의 가입증명서를 다시 보내달라는
    자연스러운 문장을 감지해서 처리함. 새 명령어를 안 외우고 그냥 말로 요청할 수 있게 함.
    True를 반환하면 이미 답장까지 다 처리한 것이므로 일반 대화(Claude)로 안 넘어감."""
    if not _CERT_LOOKUP_RE.search(user_text):
        return False
    if not any(v in user_text for v in _CERT_LOOKUP_VERBS):
        return False

    chat_id = update.effective_chat.id
    store_index = _all_registered_store_rows()
    if not store_index:
        return False

    store_name, chosen = _match_registered_store_in_text(user_text, store_index)
    if chosen is None:
        return False
    entries = store_index[store_name]

    try:
        store, pdf_bytes = _build_cert_from_registered_entry(chosen)
    except Exception:
        logger.exception("가입증명서 재검색/재생성 중 오류")
        await update.message.reply_text(f"⚠️ '{store_name}' 가입증명서를 다시 만드는 중 오류가 발생했어요.")
        return True

    out_name = f"{store['store_code']}_{store['store_name']}_{store['start_date_yymmdd']}.pdf"
    extra = f" (같은 이름으로 {len(entries)}건 등록되어 있어서 가장 최근 걸로 보내드려요)" if len(entries) > 1 else ""
    await context.bot.send_document(
        chat_id=chat_id,
        document=io.BytesIO(pdf_bytes),
        filename=out_name,
        caption=f"📄 {store['store_name']} 가입증명서{extra}",
    )
    return True


def _find_registered_store(store_name: str, store_index: dict):
    """정확히 같은 이름이 없으면(끝 '점' 유무 등) 비슷한 이름도 한번 더 찾아봄.
    (매장명, entries) 튜플을 반환하거나, 못 찾으면 (None, None)."""
    if store_name in store_index:
        return store_name, store_index[store_name]
    candidates = _closure_store_name_candidates(store_name)
    for name, entries in store_index.items():
        norm_name = re.sub(r"\s+", "", name)
        if any(c and (c in norm_name or norm_name in c) for c in candidates):
            return name, entries
    return None, None


def _resolve_closure_store(store_name: str, store_index: dict):
    """폐점서류를 매장에 매칭시킬 때 필요한 브랜드/대표자명을 찾음. 순서대로 찾아봄:
    1) '신규매장' 시트로 등록된 기록(_find_registered_store)
    2) 통합파일의 '폐점매장' 시트(폐점 처리되는 순간 이미 거기 기록되므로)
    3) 통합파일의 '최초가입전체리스트' 시트(봇이 신규/폐점매장으로 한 번도 기록한 적 없는
       아주 오래된 매장도 원래 여기엔 있는 경우가 많음 - 유진님 요청)
    셋 다 통합파일 자체가 정보 출처이고 별도 저장소는 두지 않음.
    (매장명, 브랜드, 대표자명) 튜플을 반환하거나, 끝까지 못 찾으면 (None, None, None)."""
    name, entries = _find_registered_store(store_name, store_index)
    if not entries:
        name, entries = _find_registered_store(store_name, _all_closed_store_rows())
    if not entries:
        name, entries = _find_registered_store(store_name, _all_initial_list_store_rows())
    if not entries:
        return None, None, None
    entries = sorted(
        entries, key=lambda e: _parse_date_val(e["vals"].get("접수일자")) or dt.date.min, reverse=True
    )
    chosen = entries[0]
    return name, chosen["brand"], str(chosen["vals"].get("대표자명") or "").strip()


def _classify_closure_doc(filename: str) -> str | None:
    """파일명만 보고 폐점필수서류 4종류 중 뭔지 알아냄(모르면 None). handle_document에서
    '이 파일이 폐점서류로 보이는지' 먼저 가볍게 판단하는 데 씀."""
    for category, keywords in _CLOSURE_DOC_KEYWORDS.items():
        if any(kw in filename for kw in keywords):
            return category
    return None


def _start_closure_buffer(buffers: dict, store_name: str, brand: str, owner_name: str) -> dict:
    buf = {"brand": brand, "owner_name": owner_name or "", "matched": {}}
    buffers[store_name] = buf
    return buf


async def _save_and_report_closure_progress(update: Update, store_name: str, buf: dict, buffers: dict) -> None:
    """폐점서류가 한 개 추가될 때마다 호출함. 4종류가 다 모일 때까지 기다리지 않고, 지금까지
    모인 것을 그 자리에서 바로 드라이브에 저장함(드라이브에 이미 있던 이전 서류와도 자동으로
    합쳐짐 - _save_closure_zip_merged 참고). 예전엔 4종류가 다 모일 때까지 봇 메모리에만
    들고 있어서, 중간에 봇이 재배포되면 그동안 받은 서류가 통째로 날아가는 문제가 있었음
    (2026-09-18, 유진님 피드백으로 변경)."""
    zip_name = _closure_zip_name(store_name, buf["owner_name"])
    try:
        saved, merged = _save_closure_zip_merged(store_name, buf["owner_name"], buf["brand"], buf["matched"])
    except Exception:
        logger.exception("폐점서류 압축/저장 중 오류")
        saved, merged = False, dict(buf["matched"])
    missing = [c for c in _CLOSURE_DOC_KEYWORDS if c not in merged]

    if not missing:
        if saved:
            await update.message.reply_text(f"✅ '{store_name}' 폐점필수서류 4종을 다 모아서 '{zip_name}'로 구글 드라이브에 저장했어요.")
        else:
            await update.message.reply_text(f"⚠️ '{store_name}' 서류는 다 모았는데 구글 드라이브 저장에 실패했어요. 드라이브 설정을 확인해주세요.")
        buffers.pop(store_name, None)
        return

    n = 4 - len(missing)
    if saved:
        await update.message.reply_text(
            f"💾 '{store_name}' 서류를 구글 드라이브에 저장했어요 ({n}/4). 남은 서류: {', '.join(missing)}"
        )
    else:
        await update.message.reply_text(
            f"⚠️ '{store_name}' 서류를 확인했지만({n}/4) 구글 드라이브 저장에 실패했어요. "
            f"드라이브 설정을 확인해주세요. 남은 서류: {', '.join(missing)}"
        )


async def _handle_closure_doc_upload(
    update: Update, context: ContextTypes.DEFAULT_TYPE, filename: str, file_bytes: bytes,
    category: str | None = None,
) -> None:
    """폐점필수서류로 보이는 파일이 명령어 없이 그냥 올라왔을 때 처리함. 파일명에서 등록된
    매장을 바로 찾아서 자동으로 모으고(여러 매장 동시 진행 가능), 매장당 4종류가 다 모이면
    바로 압축해서 드라이브에 저장함. 파일명에 매장명이 없으면(예: 개인정보동의서), 지금 막
    모으고 있는 매장이 하나뿐일 때만 그 매장 걸로 간주하고, 여러 개/없음이면 매장명을 물어봄.
    category를 호출부에서 미리 판단해서 넘겨줄 수도 있음(파일명으로 못 알아봐서 내용을
    직접 열어서 확인한 경우)."""
    if category is None:
        category = _classify_closure_doc(filename)
    buffers = context.user_data.setdefault("closure_buffers", {})
    store_index = _all_registered_store_rows()
    # 신규매장으로 기록된 적 없는(봇 도입 전부터 있던) 매장도 파일명 매칭 대상에 포함하기 위해,
    # 통합파일의 '폐점매장'/'최초가입전체리스트' 시트에 있는 이름도 후보에 같이 넣음
    known_names = (
        set(store_index.keys())
        | set(_all_closed_store_rows().keys())
        | set(_all_initial_list_store_rows().keys())
    )

    norm_fname = re.sub(r"\s+", "", filename)
    matched_name = None
    for name in known_names:
        if any(cand and cand in norm_fname for cand in _closure_store_name_candidates(name)):
            matched_name = name
            break

    if matched_name is None:
        if len(buffers) == 1:
            matched_name = next(iter(buffers))
        else:
            context.user_data["awaiting_closure_store_for"] = (filename, file_bytes, category)
            if buffers:
                names = ", ".join(buffers.keys())
                await update.message.reply_text(
                    f"'{filename}' 파일이 어느 매장 폐점서류인지 파일명만으로는 못 찾았어요. "
                    f"지금 모으고 있는 매장이 여러 곳이에요({names}). 매장명만 답장으로 보내주세요."
                )
            else:
                await update.message.reply_text(
                    f"'{filename}' 파일이 어느 매장 폐점서류인지 파일명만으로는 못 찾았어요. 매장명만 답장으로 보내주세요."
                )
            return

    if matched_name not in buffers:
        resolved_name, brand, owner_name = _resolve_closure_store(matched_name, store_index)
        if not resolved_name:
            await update.message.reply_text(f"⚠️ '{matched_name}' 매장을 등록된 매장 목록에서 찾지 못했어요.")
            return
        _start_closure_buffer(buffers, resolved_name, brand, owner_name)
        matched_name = resolved_name

    buf = buffers[matched_name]
    buf["matched"][category] = (filename, file_bytes)
    await _save_and_report_closure_progress(update, matched_name, buf, buffers)


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed(update):
        await update.message.reply_text("죄송해요, 이 봇은 개인용이라 사용할 수 없어요.")
        return

    chat_id = update.effective_chat.id
    user_text = update.message.text

    # 집/회사 위치를 텍스트(주소/장소명)로 저장하려는 중이면 그쪽으로 처리
    saving = context.user_data.get("saving_place")
    if saving:
        place = await _resolve_place(user_text, {})  # '집'/'회사' 문구는 여기선 의미 없으니 빈 places
        if not place:
            await update.message.reply_text("위치를 찾지 못했어요. 다른 이름으로 다시 입력하거나 위치를 공유해주세요.")
            return
        places = _load_places()
        places[saving] = place
        _save_places(places)
        context.user_data.pop("saving_place", None)
        label = "집" if saving == "home" else "회사"
        await update.message.reply_text(f"✅ {label} 위치를 저장했어요: {place.get('name')}")
        return

    # 파일명만으로 매장을 못 찾은 폐점서류(예: 개인정보동의서)가 있어서 매장명을 물어본
    # 상태면, 이 답장을 매장명으로 보고 처리함
    awaiting = context.user_data.get("awaiting_closure_store_for")
    if awaiting:
        filename, file_bytes, category = awaiting
        store_index = _all_registered_store_rows()
        store_name, brand, owner_name = _resolve_closure_store(user_text.strip(), store_index)
        if not store_name:
            await update.message.reply_text(
                f"'{user_text.strip()}' 매장을 찾지 못했어요. 등록된 매장명을 정확히 다시 알려주시겠어요?"
            )
            return
        context.user_data.pop("awaiting_closure_store_for", None)
        buffers = context.user_data.setdefault("closure_buffers", {})
        if store_name not in buffers:
            _start_closure_buffer(buffers, store_name, brand, owner_name)
        buf = buffers[store_name]
        buf["matched"][category] = (filename, file_bytes)
        await _save_and_report_closure_progress(update, store_name, buf, buffers)
        return

    # 길찾기(출발지/도착지) 입력 중이면 그쪽으로 처리
    if context.user_data.get("route_query"):
        state = context.user_data["route_query"]["state"]
        place = await _resolve_place(user_text, _load_places())
        if not place:
            if state == "await_current_location":
                await update.message.reply_text(
                    "위치를 찾지 못했어요. 📎(첨부) 버튼으로 정확한 현재 위치를 공유해주시거나, "
                    "계신 곳 이름을 다시 알려주세요."
                )
            else:
                await update.message.reply_text("위치를 찾지 못했어요. 다른 이름으로 다시 입력하거나 위치를 공유해주세요.")
            return
        await _advance_route_query(update, context, place)
        return

    # 메일 작성 중이면(제목/내용/확인 단계) 일반 대화 대신 메일 작성 흐름으로 처리
    draft = context.user_data.get("mail_draft")
    if draft:
        state = draft["state"]

        if state == "subject":
            draft["subject"] = user_text
            draft["state"] = "body"
            await update.message.reply_text("메일 내용을 입력해주세요.")
            return

        if state == "body":
            draft["body"] = user_text
            draft["state"] = "attach"
            draft["attachments"] = []
            await update.message.reply_text(
                "첨부할 파일이 있으면 지금 보내주세요 (여러 개 가능해요, 엑셀/PDF/워드/한글 파일, 사진 등).\n"
                "다 보내셨으면 '없음' 또는 '완료'라고 입력해주세요."
            )
            return

        if state == "attach":
            if user_text.strip() in ("없음", "완료", "no", "done", "skip"):
                draft["state"] = "confirm"
                n = len(draft.get("attachments", []))
                attach_line = f"첨부파일: {n}개\n" if n else ""
                preview = (
                    f"보낸사람: {MAIL_FROM_ADDRESS}\n"
                    f"받는사람: {draft['to']}\n"
                    f"제목: {draft['subject']}\n"
                    f"{attach_line}\n"
                    f"{draft['body']}\n\n"
                    "이대로 보낼까요? '네' 또는 '아니오'로 답해주세요."
                )
                await update.message.reply_text(preview)
            else:
                await update.message.reply_text(
                    "파일을 보내시거나, 다 되셨으면 '없음'/'완료'라고 입력해주세요."
                )
            return

        if state == "confirm":
            context.user_data.pop("mail_draft", None)
            if user_text.strip() in ("네", "예", "ㅇㅇ", "y", "yes", "Y"):
                try:
                    _send_email(
                        draft["to"], draft["subject"], draft["body"],
                        attachments=draft.get("attachments") or None,
                    )
                    await update.message.reply_text("✅ 메일을 보냈어요!")
                except Exception:
                    logger.exception("메일 발송 중 오류")
                    await update.message.reply_text("❌ 메일 발송에 실패했어요. 잠시 후 다시 시도해주세요.")
            else:
                await update.message.reply_text("메일 발송을 취소했어요.")
            return

    # '대구방촌점 담당자한테 보내줘'처럼 담당자 발송을 요청한 경우를 먼저 확인하고(둘 다
    # '보내줘'에 걸리므로 순서가 중요함), 아니면 '계산점 가입증명서 찾아줘'처럼 자연스러운
    # 문장으로 재발급을 요청한 경우를 감지해서 처리
    # 처브 증명서를 보낼 세정 담당자를 물어본 상태면, 이 답장(이름/이메일)으로 발송 준비
    # (예전 '가입증명서 찾아줘'/'담당자한테 보내줘' 기능은 2026-11-12 처브 전환으로 껐음)
    if await _sj_handle_receiver_reply(update, user_text):
        return
    if await _sj_handle_await_contact_reply(update, user_text):
        return

    # 자연스러운 문장으로 길찾기를 요청한 경우 감지해서 처리
    if await _handle_natural_route_request(update, context, user_text):
        return

    history = conversations.setdefault(chat_id, [])
    history.append({"role": "user", "content": user_text})
    history = history[-MAX_HISTORY:]
    # 잘라낸 뒤 맨 앞이 assistant 답변이면 API가 거부하므로, 항상 user 메시지로 시작하게 맞춤
    while history and history[0].get("role") != "user":
        history.pop(0)

    await context.bot.send_chat_action(chat_id=chat_id, action="typing")

    system_prompt = SYSTEM_PROMPT
    loc = context.user_data.get("location")
    if loc:
        loc_address = loc.get("address") if isinstance(loc, dict) else loc
        system_prompt += (
            f"\n\n참고: 사용자의 최근 공유 위치는 '{loc_address}'입니다. "
            "근처 맛집/장소 등 위치 기반 질문에는 이 정보를 활용해 웹 검색하세요."
        )

    system_prompt += (
        "\n\n메일, 접수 내역, 받은 서류, 개체번호처럼 메일로 받은 정보를 물으면 search_mail 도구로 실제 "
        "메일함을 검색해서 답하세요. 검색 결과에 없는 내용은 지어내지 말고 못 찾았다고 말하세요. "
        "찾지 못하면 검색어를 바꿔 한두 번 더 시도해도 됩니다."
    )

    try:
        # 메일 검색 도구는 봇 서버에서 직접 실행해야 해서, Claude가 도구를 쓰겠다고 하면 실행
        # 결과를 돌려주고 다시 묻는 과정을 몇 번 반복함(웹 검색은 Anthropic 쪽에서 알아서 처리됨)
        tools = CHAT_TOOLS + ([MAIL_SEARCH_TOOL] if (GMAIL_ADDRESS and GMAIL_APP_PASSWORD) else [])
        working = list(history)
        response = None
        for _ in range(6):
            response = await asyncio.to_thread(
                client.messages.create,
                model=MODEL_NAME,
                max_tokens=2048,
                system=system_prompt,
                messages=working,
                tools=tools,
            )
            if response.stop_reason not in ("tool_use", "pause_turn"):
                break
            assistant_blocks = [b.model_dump(exclude_none=True) for b in response.content]
            working.append({"role": "assistant", "content": assistant_blocks})
            if response.stop_reason == "pause_turn":
                continue  # 웹 검색이 길어져 잠깐 멈춘 경우 - 그대로 이어서 요청
            tool_results = []
            for b in response.content:
                if b.type == "tool_use":
                    await context.bot.send_chat_action(chat_id=chat_id, action="typing")
                    result = await _run_chat_tool(b.name, b.input or {})
                    tool_results.append({"type": "tool_result", "tool_use_id": b.id, "content": result})
            if not tool_results:
                break
            working.append({"role": "user", "content": tool_results})
        # 도구 사용(검색) 결과가 섞여 있을 수 있어 텍스트 블록만 모아서 답변으로 사용
        reply_text = "".join(
            block.text for block in response.content if block.type == "text"
        ).strip()
        if not reply_text:
            reply_text = "죄송해요, 답변을 생성하지 못했어요. 다시 한번 물어봐주세요."
    except Exception:
        logger.exception("Claude API 호출 중 오류")
        reply_text = "죄송해요, 답변을 만드는 중에 오류가 발생했어요. 잠시 후 다시 시도해주세요."

    history.append({"role": "assistant", "content": reply_text})
    conversations[chat_id] = history
    _save_conversations()

    # 텔레그램은 메시지 하나에 4096자까지만 보낼 수 있어서, 긴 답변은 나눠서 보냄
    for i in range(0, len(reply_text), 4000):
        await update.message.reply_text(reply_text[i:i + 4000])


def _decode_mime_words(s: str) -> str:
    if not s:
        return ""
    decoded = decode_header(s)
    return "".join(
        (t.decode(enc or "utf-8", errors="ignore") if isinstance(t, bytes) else t)
        for t, enc in decoded
    )


def _get_email_body(msg: email.message.Message) -> str:
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = str(part.get("Content-Disposition"))
            if ctype == "text/plain" and "attachment" not in disp:
                charset = part.get_content_charset() or "utf-8"
                try:
                    return part.get_payload(decode=True).decode(charset, errors="ignore")
                except Exception:
                    return ""
        return ""
    else:
        charset = msg.get_content_charset() or "utf-8"
        try:
            return msg.get_payload(decode=True).decode(charset, errors="ignore")
        except Exception:
            return ""


def _extract_xlsx_attachments(msg: email.message.Message) -> list[tuple[str, bytes]]:
    """메일에서 .xlsx 첨부파일만 골라 (파일명, 파일바이트) 목록으로 반환"""
    attachments = []
    if not msg.is_multipart():
        return attachments
    for part in msg.walk():
        filename = part.get_filename()
        if not filename:
            continue
        filename = _decode_mime_words(filename)
        if not filename.lower().endswith(".xlsx"):
            continue
        try:
            payload = part.get_payload(decode=True)
        except Exception:
            continue
        if payload:
            attachments.append((filename, payload))
    return attachments


def _extract_non_xlsx_attachments(msg: email.message.Message) -> list[tuple[str, bytes]]:
    """메일에서 .xlsx가 아닌 실제 첨부파일(이미지/PDF 등)을 전부 (파일명, 파일바이트) 목록으로
    반환함. 폐점매장 처리 시 폐점필수서류(사업자등록증/주민등록증/통장사본/개인정보처리동의서)를
    찾아 압축해서 드라이브에 저장하는 데 씀."""
    attachments = []
    if not msg.is_multipart():
        return attachments
    for part in msg.walk():
        filename = part.get_filename()
        if not filename:
            continue
        filename = _decode_mime_words(filename)
        if filename.lower().endswith(".xlsx"):
            continue
        try:
            payload = part.get_payload(decode=True)
        except Exception:
            continue
        if payload:
            attachments.append((filename, payload))
    return attachments


# ===================== 대화 중 메일함 검색 도구 =====================
# 텔레그램 대화에서 "내 메일에서 소 사고 접수 온 거 최근 개체번호 알려줘"처럼 물으면, Claude가
# 이 도구로 실제 Gmail 메일함을 검색해서 본문/첨부 내용을 보고 답함(2026-11-06 유진님 요청).
# 검색만 하고, 메일을 보내거나 지우거나 읽음 표시를 바꾸지 않음(BODY.PEEK로 읽기 전용 조회).

MAIL_SEARCH_TOOL = {
    "name": "search_mail",
    "description": (
        "유진님의 Gmail 메일함 전체(받은편지함, 보관된 메일 포함)를 검색해서 최근 메일부터 제목, 보낸사람, "
        "날짜, 본문, 첨부파일 내용(엑셀/PDF 텍스트 일부)을 돌려줍니다. 사용자가 메일, 접수 내역, 받은 서류, "
        "개체번호, 정산파일 등 메일로 받은 정보를 물으면 반드시 이 도구로 먼저 찾아보세요. "
        "query에는 Gmail 검색 문법을 쓸 수 있습니다(예: '사고접수', 'from:sejung.co.kr', "
        "'subject:가입 newer_than:30d', 'has:attachment 소 사고')."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "query": {"type": "string", "description": "Gmail 검색어(Gmail 검색창에 넣는 문법 그대로)"},
            "max_results": {"type": "integer", "description": "가져올 최대 메일 수(기본 5, 최대 10)"},
        },
        "required": ["query"],
    },
}


def _attachment_text_preview(filename: str, data: bytes, limit: int = 1500) -> str:
    """첨부파일 안의 글자를 조금 뽑아냄(엑셀은 셀 값, PDF는 텍스트). 개체번호처럼 첨부에만
    있는 정보도 찾을 수 있게 하려는 용도. 실패하면 빈 문자열."""
    lower = filename.lower()
    try:
        if lower.endswith((".xlsx", ".xlsm")):
            wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
            lines = []
            for ws in wb.worksheets[:3]:
                lines.append(f"[시트 {ws.title}]")
                for row in ws.iter_rows(max_row=40, values_only=True):
                    vals = [str(v) for v in row if v not in (None, "")]
                    if vals:
                        lines.append(" | ".join(vals))
            return "\n".join(lines)[:limit]
        if lower.endswith(".pdf"):
            reader = PdfReader(io.BytesIO(data))
            text = "\n".join((p.extract_text() or "") for p in reader.pages[:3])
            return text[:limit]
    except Exception:
        return ""
    return ""


def _gmail_all_mail_box(imap) -> str:
    """Gmail의 '전체보관함' 메일함 이름을 찾음(언어 설정에 따라 이름이 달라서 \\All 표시로 찾음)."""
    try:
        status, boxes = imap.list()
        for line in boxes or []:
            s = line.decode("utf-8", errors="ignore") if isinstance(line, bytes) else str(line)
            if "\\All" in s:
                m = re.search(r'"([^"]+)"\s*$', s)
                if m:
                    return f'"{m.group(1)}"'
    except Exception:
        pass
    return "INBOX"


def _search_mail_sync(query: str, max_results: int = 5) -> str:
    if not GMAIL_ADDRESS or not GMAIL_APP_PASSWORD:
        return "메일 계정이 설정되어 있지 않아 검색할 수 없어요."
    max_results = max(1, min(int(max_results or 5), 10))
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    try:
        imap.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        imap.select(_gmail_all_mail_box(imap), readonly=True)
        # 한글 검색어는 literal로 넘겨야 Gmail이 제대로 인식함
        imap.literal = query.encode("utf-8")
        status, data = imap.uid("SEARCH", "CHARSET", "UTF-8", "X-GM-RAW")
        uids = (data[0] or b"").split() if data else []
        if not uids:
            return f"'{query}'로 검색된 메일이 없어요."
        results = []
        for uid in reversed(uids[-max_results:]):  # 최신 메일부터
            status, msg_data = imap.uid("FETCH", uid, "(BODY.PEEK[])")
            if not msg_data or not msg_data[0] or not isinstance(msg_data[0], tuple):
                continue
            msg = email.message_from_bytes(msg_data[0][1])
            subject = _decode_mime_words(msg.get("Subject", ""))
            sender = _decode_mime_words(msg.get("From", ""))
            date = msg.get("Date", "")
            body = (_get_email_body(msg) or "").strip()[:2000]
            att_parts = []
            for fname, fbytes in _extract_xlsx_attachments(msg) + _extract_non_xlsx_attachments(msg):
                preview = _attachment_text_preview(fname, fbytes)
                att_parts.append(f"- {fname}" + (f"\n{preview}" if preview else ""))
            att_text = "\n".join(att_parts) if att_parts else "(첨부 없음)"
            results.append(
                f"### 메일\n날짜: {date}\n보낸사람: {sender}\n제목: {subject}\n본문:\n{body}\n첨부:\n{att_text}"
            )
        return "\n\n".join(results)[:12000] or "검색은 됐지만 내용을 읽지 못했어요."
    finally:
        try:
            imap.logout()
        except Exception:
            pass


async def _run_chat_tool(name: str, tool_input: dict) -> str:
    if name == "search_mail":
        try:
            return await asyncio.to_thread(
                _search_mail_sync, str(tool_input.get("query") or ""), tool_input.get("max_results") or 5
            )
        except Exception as e:
            logger.exception("메일 검색 도구 실행 중 오류")
            return f"메일 검색 중 오류가 났어요: {type(e).__name__}"
    return f"알 수 없는 도구: {name}"


# 폐점필수서류로 인정하는 4종류와, 각각을 파일명에서 알아보는 키워드. 4종류가 전부 매칭돼야
# '서류가 갖춰졌다'고 보고 압축해서 저장함(유진님이 정해주신 기준).
# 해지 서류는 사업자등록증·통장사본 2가지만 받음(2026-11-13 유진님: 신분증·개인정보동의서·제3자동의서는 아예 배제)
_CLOSURE_DOC_KEYWORDS = {
    "사업자등록증": ["사업자등록증"],
    "통장사본": ["통장사본", "통장"],
}

# 파일명만으로 못 알아본 서류(예: 'OL인천구월점_251112.pdf'처럼 키워드가 전혀 없는 이름)를,
# 문서 내용을 직접 열어서 확인할 때 쓰는 텍스트 단서. PDF는 먼저 이 키워드로 확인하고(텍스트가
# 있으면 API 호출 없이 바로 판단 가능), 안 되면(스캔본 PDF나 사진) Claude에게 실제 내용을
# 보여주고 판단시킴(유진님 요청: "파일을 다 열어서 파악해보고 없는 서류를 알려줘").
_CLOSURE_DOC_CONTENT_HINTS = {
    "사업자등록증": ["사업자등록증", "사업자등록번호", "개업연월일", "법인등록번호"],
    "통장사본": ["예금주", "계좌번호", "보통예금", "예금거래통장"],
}


def _extract_pdf_text(file_bytes: bytes) -> str:
    try:
        reader = PdfReader(io.BytesIO(file_bytes))
        return "\n".join((page.extract_text() or "") for page in reader.pages)
    except Exception:
        return ""


def _classify_by_content_text(text: str) -> str | None:
    if not text:
        return None
    for category, hints in _CLOSURE_DOC_CONTENT_HINTS.items():
        if any(h in text for h in hints):
            return category
    return None


async def _classify_doc_with_claude(filename: str, file_bytes: bytes, mimetype: str) -> str | None:
    """파일명/텍스트로 못 알아본 서류를 실제 내용(이미지 또는 PDF)을 Claude에게 직접 보여주고
    폐점필수서류 4종류 중 무엇인지 판단시킴. 어디에도 해당 안 되면 None."""
    b64 = base64.b64encode(file_bytes).decode()
    if mimetype == "application/pdf":
        content_block = {
            "type": "document",
            "source": {"type": "base64", "media_type": "application/pdf", "data": b64},
        }
    else:
        content_block = {
            "type": "image",
            "source": {"type": "base64", "media_type": mimetype, "data": b64},
        }
    try:
        response = client.messages.create(
            model=MODEL_NAME,
            max_tokens=20,
            system=(
                "첨부된 문서 이미지/파일이 '사업자등록증', '통장사본'(예금계좌 확인서 포함) 중 "
                "어떤 것인지 정확히 그 이름 하나만 답하세요. 둘 다 아니면 정확히 NONE이라고만 답하세요. "
                "다른 설명은 절대 하지 마세요."
            ),
            messages=[{
                "role": "user",
                "content": [content_block, {"type": "text", "text": f"파일명: {filename}"}],
            }],
        )
        raw = "".join(b.text for b in response.content if b.type == "text").strip()
        return raw if raw in _CLOSURE_DOC_KEYWORDS else None
    except Exception:
        logger.exception("문서 내용 분류 중 오류")
        return None


async def _classify_closure_doc_deep(filename: str, file_bytes: bytes) -> str | None:
    """파일명으로 먼저 판단하고(_classify_closure_doc), 안 되면 문서 내용을 직접 열어서
    (PDF면 텍스트 추출 -> 그래도 안 되면 Claude 비전) 폐점필수서류 4종류 중 어느 것인지 알아냄.
    확장자가 pdf/jpg/jpeg/png가 아니면 애초에 폐점서류일 가능성이 낮아 바로 None."""
    by_name = _classify_closure_doc(filename)
    if by_name:
        return by_name
    lower = filename.lower()
    if lower.endswith(".pdf"):
        text = _extract_pdf_text(file_bytes)
        by_text = _classify_by_content_text(text)
        if by_text:
            return by_text
        mimetype = "application/pdf"
    elif lower.endswith((".jpg", ".jpeg")):
        mimetype = "image/jpeg"
    elif lower.endswith(".png"):
        mimetype = "image/png"
    else:
        return None
    return await _classify_doc_with_claude(filename, file_bytes, mimetype)


async def _match_closure_docs_with_content(
    store_name: str, attachments: list, allow_no_name_fallback: bool = False
) -> dict:
    """_match_closure_docs로 먼저 파일명 기준으로 매칭하고, 그래도 4종류가 안 채워지면 아직
    아무 카테고리에도 안 쓰인 첨부파일들을 직접 열어서(내용 기반) 다시 확인함. 파일명이
    'OL인천구월점_251112.pdf'처럼 아무 단서가 없는 경우를 위한 보강."""
    matched = _match_closure_docs(store_name, attachments, allow_no_name_fallback=allow_no_name_fallback)
    if len(matched) >= 4:
        return matched
    used_filenames = {fn for fn, _fb in matched.values()}
    for filename, file_bytes in attachments:
        if len(matched) >= 4:
            break
        if filename in used_filenames or not filename.lower().endswith((".pdf", ".jpg", ".jpeg", ".png")):
            continue
        category = await _classify_closure_doc_deep(filename, file_bytes)
        if category and category not in matched:
            matched[category] = (filename, file_bytes)
            used_filenames.add(filename)
    return matched


def _closure_store_name_candidates(store_name: str) -> set:
    norm_store = re.sub(r"\s+", "", store_name)
    candidates = {norm_store}
    if norm_store.endswith("점"):
        candidates.add(norm_store[:-1])
    candidates.discard("")
    return candidates


def _match_closure_docs(store_name: str, attachments: list, allow_no_name_fallback: bool = False) -> dict:
    """attachments 중 파일명에 store_name(끝의 '점'은 있든 없든)이 들어있는 것만 골라,
    폐점필수서류 4종류별로 하나씩 분류해서 {카테고리: (파일명, 파일바이트)}로 반환함.
    한 메일에 여러 매장 폐점이 나열되어 오면 보통 개인정보동의서 등도 매장명이 파일명에
    들어있어서 이름 매칭만으로 충분함. 다만 한 메일에 폐점매장이 '하나뿐'일 때는 개인정보
    동의서처럼 매장명이 아예 안 들어간 파일도 있을 수 있어서(예: "개인정보 처리에 관한
    동의서_20260826_0001.pdf"), allow_no_name_fallback=True로 호출하면 이름이 매칭 안 된
    파일도 키워드만으로 보충해서 채워줌(모호할 수 있으니 매장이 여러 개일 땐 절대 켜면 안 됨).
    4종류가 다 안 갖춰지면 호출부에서 len(결과) < 4 로 걸러서 저장을 건너뛰면 됨."""
    candidate_names = _closure_store_name_candidates(store_name)

    matched: dict = {}
    fallback_candidates: dict = {}
    for filename, file_bytes in attachments:
        norm_fname = re.sub(r"\s+", "", filename)
        name_matches = any(name in norm_fname for name in candidate_names)
        for category, keywords in _CLOSURE_DOC_KEYWORDS.items():
            if not any(kw in filename for kw in keywords):
                continue
            if name_matches:
                if category not in matched:
                    matched[category] = (filename, file_bytes)
            elif category not in fallback_candidates:
                fallback_candidates[category] = (filename, file_bytes)
            break

    if allow_no_name_fallback:
        for category, item in fallback_candidates.items():
            matched.setdefault(category, item)
    return matched


def _find_store_zip_attachment(store_name: str, attachments: list):
    """attachments 중에 매장명이 파일명에 들어간 .zip 파일이 있으면 그대로 반환함(유진님이
    이미 매장별로 직접 압축해서 보내는 경우 대응). 없으면 None."""
    candidate_names = _closure_store_name_candidates(store_name)
    for filename, file_bytes in attachments:
        if not filename.lower().endswith(".zip"):
            continue
        norm_fname = re.sub(r"\s+", "", filename)
        if any(name in norm_fname for name in candidate_names):
            return filename, file_bytes
    return None


def _zip_has_all_closure_docs(zip_bytes: bytes) -> bool:
    """이미 압축되어 온 zip 안에 폐점필수서류 4종류가 다 들어있는지(파일명 키워드 기준으로만)
    확인함."""
    try:
        zf = zipfile.ZipFile(io.BytesIO(zip_bytes))
        names = zf.namelist()
    except Exception:
        return False
    found = set()
    for name in names:
        base = os.path.basename(name)
        for category, keywords in _CLOSURE_DOC_KEYWORDS.items():
            if any(kw in base for kw in keywords):
                found.add(category)
    return len(found) >= 4


def _zip_closure_docs(matched: dict) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for _category, (filename, file_bytes) in matched.items():
            zf.writestr(filename, file_bytes)
    return buf.getvalue()


def _unzip_closure_docs(zip_bytes: bytes) -> dict:
    """드라이브에 이미 저장돼 있던 폐점서류 zip을 다시 열어서, 파일명 키워드로 재분류한
    {category: (filename, file_bytes)} 딕셔너리로 돌려줌(저장할 때도 파일명 기준으로 분류해서
    넣었으므로, 다시 열 때도 파일명만으로 충분함)."""
    out = {}
    try:
        zf = zipfile.ZipFile(io.BytesIO(zip_bytes))
        for name in zf.namelist():
            base = os.path.basename(name)
            category = _classify_closure_doc(base)
            if category and category not in out:
                out[category] = (base, zf.read(name))
    except Exception:
        logger.exception("기존 폐점서류 zip 압축 해제 중 오류")
    return out


def _save_closure_zip_merged(
    store_name: str, owner_name: str, brand: str | None, new_matched: dict
) -> tuple[bool, dict]:
    """폐점서류를 저장할 때, 드라이브에 이미 같은 매장 zip이 있으면 그 안의 기존 서류와 이번에
    새로 모인 서류를 합쳐서(겹치는 종류는 새 걸로 덮어씀) 다시 올림. 4종류가 다 안 모였어도
    매번 그 자리에서 바로 저장하는 이유: 예전엔 봇 메모리(또는 이번 메일 첨부파일)에만 다
    모일 때까지 들고 있다가 한 번에 저장해서, 중간에 봇이 재배포되거나 나머지 서류가 다른
    메일/시점에 따로 오면 이전에 받은 서류가 통째로 유실됐음(2026-09-18, 유진님 피드백으로
    변경). 반환값: (이번 저장 성공 여부, 드라이브 기존 것까지 합친 전체 matched 딕셔너리)."""
    zip_name = _closure_zip_name(store_name, owner_name)
    merged = dict(new_matched)
    service = _get_gdrive_service()
    if service is None:
        return False, merged
    try:
        parent_id = GDRIVE_FOLDER_ID
        if brand:
            resolved = _resolve_drive_target_folder(service, brand, extra_subfolder="폐점서류")
            if resolved:
                parent_id = resolved
        existing = _gdrive_find_file(service, parent_id, lambda name: name == zip_name)
        if existing is not None:
            existing_bytes = _gdrive_download_file(service, existing["id"])
            if existing_bytes is not None:
                existing_matched = _unzip_closure_docs(existing_bytes)
                merged = {**existing_matched, **new_matched}
        zip_bytes = _zip_closure_docs(merged)
        media = _gdrive_media(zip_bytes, mimetype="application/zip")
        if existing is not None:
            service.files().update(fileId=existing["id"], media_body=media).execute()
        else:
            service.files().create(
                body={"name": zip_name, "parents": [parent_id]},
                media_body=media,
                fields="id",
            ).execute()
        return True, merged
    except Exception:
        logger.exception("폐점서류 병합/저장 중 오류")
        return False, merged


def _closure_zip_name(store_name: str, owner_name: str) -> str:
    if owner_name:
        return f"{store_name}_{owner_name}.zip"
    return f"{store_name}_폐점서류.zip"


async def _send_message_retrying(bot, chat_id, text: str, retries: int = 3, **kwargs) -> None:
    """텔레그램 서버 쪽 일시적 타임아웃(telegram.error.TimedOut) 때문에 중요한 결과 알림
    (저장 성공/실패 등)이 통째로 유실되는 걸 막기 위한 재시도 래퍼. 실제로 폐점서류 저장
    자체는 문제없이 끝났는데 그 결과를 알려주는 메시지 전송만 타임아웃 나서, 예외가 바깥
    try/except까지 새어나가 알림이 아예 안 간 사고가 있었음(2026-09-18) - 이후로 알림
    전송에는 항상 이 함수를 씀."""
    delay = 1.0
    last_err = None
    for attempt in range(retries):
        try:
            await bot.send_message(chat_id=chat_id, text=text, **kwargs)
            return
        except Exception as e:
            last_err = e
            logger.warning("메시지 전송 재시도(%d/%d): %s", attempt + 1, retries, e)
            if attempt < retries - 1:
                await asyncio.sleep(delay)
                delay *= 2
    logger.error("메시지 전송 최종 실패: %s", last_err)


async def _send_document_retrying(bot, chat_id, file_bytes: bytes, filename: str, retries: int = 3, **kwargs) -> bool:
    """_send_message_retrying과 같은 이유로 만든 문서 전송용 재시도 래퍼. 통합파일(xlsx)처럼
    용량이 있는 파일은 텍스트 메시지보다 타임아웃이 더 잘 나서 특히 중요함. 성공하면 True,
    재시도까지 다 실패하면 False를 반환함(호출한 쪽에서 이 실패 때문에 뒷단 로직 전체가
    끊기지 않도록 값을 보고 판단할 수 있게)."""
    delay = 1.0
    last_err = None
    for attempt in range(retries):
        try:
            await bot.send_document(
                chat_id=chat_id, document=io.BytesIO(file_bytes), filename=filename, **kwargs
            )
            return True
        except Exception as e:
            last_err = e
            logger.warning("문서 전송 재시도(%d/%d): %s", attempt + 1, retries, e)
            if attempt < retries - 1:
                await asyncio.sleep(delay)
                delay *= 2
    logger.error("문서 전송 최종 실패: %s", last_err)
    return False


KST = ZoneInfo("Asia/Seoul")

# 구글 드라이브 refresh token이 구글 앱 '테스트' 상태라 발급 후 약 7일이면 만료되므로, 6일
# 주기로 미리 갱신 안내를 드림(2026-09-22 결정). 그 6일째가 주말/공휴일이면 다음 영업일에
# 알려달라고 하셔서(2026-09-29), 아래 공휴일 목록 기준으로 평일까지 미뤄서 계산함.
# 주의: 이 목록은 2026년 기준. 해가 바뀌면 그 해의 대한민국 법정공휴일로 갱신해야 함
# (갱신 안 해도 주말 제외 로직은 계속 동작하므로 최악의 경우 공휴일에 알림이 가는 정도의
# 문제만 생김 - 완전히 멈추진 않음).
_KR_HOLIDAYS_2026 = {
    dt.date(2026, 1, 1),   # 신정
    dt.date(2026, 2, 16),  # 설날 연휴
    dt.date(2026, 2, 17),  # 설날
    dt.date(2026, 2, 18),  # 설날 연휴
    dt.date(2026, 3, 2),   # 삼일절 대체공휴일
    dt.date(2026, 5, 5),   # 어린이날
    dt.date(2026, 5, 25),  # 부처님오신날 대체공휴일
    dt.date(2026, 6, 6),   # 현충일
    dt.date(2026, 8, 17),  # 광복절 대체공휴일
    dt.date(2026, 9, 24),  # 추석 연휴
    dt.date(2026, 9, 25),  # 추석
    dt.date(2026, 9, 26),  # 추석 연휴
    dt.date(2026, 10, 5),  # 개천절 대체공휴일
    dt.date(2026, 10, 9),  # 한글날
    dt.date(2026, 12, 25),  # 크리스마스
}


def _is_kr_business_day(d: dt.date) -> bool:
    if d.weekday() >= 5:  # 토(5)/일(6)
        return False
    if d.year == 2026 and d in _KR_HOLIDAYS_2026:
        return False
    return True


def _should_send_token_reminder(today: dt.date) -> bool:
    """오늘이 '6일 주기 기준일(1/6/11/16/21/26일, 주말·공휴일이면 다음 영업일로 미룸)'과
    같은 날인지 판단함."""
    candidates = [d for d in (1, 6, 11, 16, 21, 26) if d <= today.day]
    anchor_day = max(candidates) if candidates else 26
    try:
        target = today.replace(day=anchor_day)
    except ValueError:
        return False
    while not _is_kr_business_day(target):
        target += dt.timedelta(days=1)
    return target == today


async def check_token_reminder(context: ContextTypes.DEFAULT_TYPE) -> None:
    """구글 드라이브 refresh token을 슬슬 갱신할 때가 되면(6일 주기, 주말/공휴일이면 다음
    영업일) 유진님께 텔레그램으로 직접 안내함(2026-09-29, 유진님 요청 - 이 세션이 아니라
    봇 자체가 보내도록)."""
    if not ALLOWED_USER_ID:
        return
    today = dt.datetime.now(KST).date()
    if not _should_send_token_reminder(today):
        return
    text = (
        "🔑 슬슬 구글 드라이브 토큰을 갱신할 때예요.\n\n"
        "1. myaccount.google.com/permissions 접속(bubal5503@gmail.com으로 로그인) → "
        "검색창에 '김대리' 검색 → 나오면 눌러서 '액세스 권한 삭제' (검색해도 안 나오면 "
        "이미 연결이 끊긴 상태이니 이 단계는 건너뛰어도 돼요)\n"
        "2. developers.google.com/oauthplayground 접속\n"
        "3. 오른쪽 위 톱니바퀴 → 'OAuth 2.0 구성'에서 '내 자체 OAuth 자격증명 사용' 체크, "
        "Client ID/Secret 입력돼 있는지, '강제 프롬프트'가 '동의 화면'인지 확인\n"
        "4. 왼쪽 '1단계' scope 입력창에 정확히 https://www.googleapis.com/auth/drive 만 "
        "입력 → 'API 권한 부여' 클릭 → bubal5503@gmail.com으로 로그인 및 동의\n"
        "5. '2단계'에서 '인증 코드를 토큰으로 교환하기' 클릭\n"
        "6. 오른쪽 응답 패널에서 \"scope\": \"https://www.googleapis.com/auth/drive\"(.file "
        "아님)와 refresh_token 값이 있는지 확인\n"
        "7. Render 대시보드 → yujin-ai-bot 서비스 → Environment 탭 → "
        "GDRIVE_OAUTH_REFRESH_TOKEN 값 입력칸 클릭 → Ctrl+A → Delete → Playground의 "
        "'리프레시 토큰' 칸에서 직접 복사(옮겨 적지 말고 꼭 복사/붙여넣기로)한 새 값을 붙여넣기\n"
        "8. Save Changes → 자동 재배포 대기(2~3분) → 완료되면 정산파일이나 가입증명서 "
        "하나로 실제 테스트"
    )
    await _send_message_retrying(context.bot, ALLOWED_USER_ID, text)


# ---------------------------------------------------------------------------
# 팜스코 소 가축재해보험 사고접수 자동 기재 (2026-11-12 추가)
#  - farmscohw@gmail.com 이 보내는 "[보험 접수]" 메일을 주기적으로 읽어서, 구글 드라이브의
#    「26년 계약리스트.xlsx」(소(240430)/26년) 사고내역 시트 다음 빈 행에 기재함.
#  - PC가 꺼져 있어도 드라이브의 '같은 파일'을 덮어써서 저장 -> PC가 켜지면 드라이브 동기화로
#    바탕화면 파일에도 그대로 반영됨. 백업 파일은 만들지 않음(유진님 요청).
#  - 기재 칸: 날짜(발생일자) / 이력번호 / 농장명 / 증권번호 / 가입금액 / 사고내용(폐사·긴급출하).
#    사고접수번호·처리경과는 유진님이 직접 입력(기존 방식 유지).
#  - 중복 방지: 사고내역에 이미 있는 이력번호는 다시 쓰지 않음. 최근 CLAIMS_LOOKBACK_DAYS일
#    메일을 매번 대조하므로, PC 엑셀에서 저장하다 봇이 쓴 행이 덮여 사라져도 다음 실행 때 다시 채움.
# ---------------------------------------------------------------------------
CLAIMS_SENDER = os.environ.get("CLAIMS_SENDER", "farmscohw@gmail.com")
CLAIMS_XLSX_FILE_ID = os.environ.get("CLAIMS_XLSX_FILE_ID", "1hysBIWklqGKxW5_Un0blQUTpMMct3cXs")
CLAIMS_CHECK_INTERVAL = int(os.environ.get("CLAIMS_CHECK_INTERVAL", "600"))
CLAIMS_LOOKBACK_DAYS = int(os.environ.get("CLAIMS_LOOKBACK_DAYS", "14"))
CLAIMS_SHEET = "사고내역"
_claims_state: dict = {"last_md5": None, "seen_ids": set()}


def _claims_norm_id(v) -> str:
    """이력번호 비교용 정규화: 숫자만 남기고, 12자리 '002...' 형태면 앞의 '00'을 떼서
    시트의 10자리 숫자(예: 2214485575)와 맞춤."""
    d = re.sub(r"\D", "", str(v or ""))
    if len(d) == 12 and d.startswith("00"):
        d = d[2:]
    return d


def _claims_fmt_id(v) -> str:
    """'002 2144 8557 5' 형태로 표기(사고내역 기존 표기와 동일)."""
    d = re.sub(r"\D", "", str(v or ""))
    if len(d) == 10:
        d = "00" + d
    if len(d) == 12:
        return f"{d[:3]} {d[3:7]} {d[7:11]} {d[11]}"
    return str(v or "").strip()


def _claims_parse_mail(body: str) -> dict | None:
    def grab(label):
        m = re.search(label + r"\s*[:：]\s*(.+?)(?:\s*🐄|\n|$)", body)
        return m.group(1).strip() if m else ""
    kind = grab("발생구분")
    farm = grab("발생농장")
    animal = grab("개체번호")
    date_s = grab("발생일자")
    reason = grab("발생사유")
    if not animal or not _claims_norm_id(animal):
        return None
    m = re.search(r"(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})", date_s)
    occur = dt.datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else None
    m2 = re.search(r"\(([^)]+)\)", kind)
    accident = (m2.group(1) if m2 else kind).strip() or kind
    return {"kind": kind, "accident": accident, "farm": farm, "animal": animal,
            "occur": occur, "reason": reason}


def _claims_fetch_mails() -> list[dict]:
    """최근 CLAIMS_LOOKBACK_DAYS일 동안 온 사고접수 메일을 읽어서 파싱해 반환(오래된 순)."""
    since = (dt.datetime.now(KST) - dt.timedelta(days=CLAIMS_LOOKBACK_DAYS)).strftime("%d-%b-%Y")
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    out = []
    try:
        imap.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        imap.select("INBOX", readonly=True)  # 읽음 표시를 바꾸지 않음
        _st, data = imap.uid("search", None, f'(FROM "{CLAIMS_SENDER}" SINCE {since})')
        for uid in (data[0].split() if data and data[0] else []):
            _st, msg_data = imap.uid("fetch", uid, "(BODY.PEEK[])")
            if not msg_data or msg_data[0] is None:
                continue
            msg = email.message_from_bytes(msg_data[0][1])
            subject = _decode_mime_words(msg.get("Subject", ""))
            if "보험 접수" not in subject:
                continue
            parsed = _claims_parse_mail(_get_email_body(msg))
            if parsed:
                parsed["msg_id"] = msg.get("Message-ID", uid.decode())
                out.append(parsed)
    finally:
        try:
            imap.logout()
        except Exception:
            pass
    return out


def _claims_lookup(wb, animal_norm: str) -> dict:
    """개체별 시트들에서 개체를 찾아 농장명·가입금액·농장관리번호를 돌려주고, 농장리스트에서
    증권번호를 찾음. 사고 표시(폐사/긴급출하)를 남길 셀 위치도 같이 돌려줌."""
    info = {"farm": None, "amount": None, "mgmt": None, "policy": None, "mark": None}
    # 1) 4월 갱신 개체별: D=이력번호, K/L=농장관리번호, M=농장명(가입 기준), O=가입금액, U=사고표시
    if "4월갱신시개체별" in wb.sheetnames:
        ws = wb["4월갱신시개체별"]
        for row in ws.iter_rows(min_row=2):
            if (_claims_norm_id(row[3].value) == animal_norm
                    and "재가입" in str(row[13].value or "")
                    and isinstance(row[14].value, (int, float))):
                # N열이 '재가입(..)'인 개체만 O열이 가입금액임('계약기간지속'은 O열이 다른 값이라
                # 이 경우엔 월별 추징 시트의 보험가입금액을 씀)
                info.update(farm=row[12].value or row[4].value, amount=row[14].value,
                            mgmt=row[11].value, mark=(ws.title, row[20].coordinate))
                break
    # 2) 없거나 가입금액이 비어 있으면 월별 추징(신규 가입) 시트: C=이력번호, B=농장, K=농장관리번호, P=가입금액, T=사고표시
    if info["amount"] in (None, ""):
        for name in reversed(wb.sheetnames):
            if "추징" not in name:
                continue
            ws = wb[name]
            hit = None
            for row in ws.iter_rows(min_row=2):
                if len(row) > 19 and _claims_norm_id(row[2].value) == animal_norm and isinstance(row[15].value, (int, float)):
                    hit = row
                    break
            if hit is not None:
                info.update(farm=info["farm"] or hit[1].value, amount=hit[15].value,
                            mgmt=info["mgmt"] or hit[10].value, mark=(ws.title, hit[19].coordinate))
                break
    # 3) 증권번호/농장명: 농장리스트에서 농장관리번호(O열)로 찾음. 같은 관리번호가 여러 개면 농장명까지 맞는 행.
    if "농장리스트" in wb.sheetnames and info["mgmt"] is not None:
        cands = []
        for row in wb["농장리스트"].iter_rows(min_row=4, values_only=True):
            if len(row) > 14 and row[14] is not None and str(row[14]).strip() == str(info["mgmt"]).strip() and row[8]:
                cands.append(row)
        pick = None
        if len(cands) == 1:
            pick = cands[0]
        elif cands:
            pick = next((c for c in cands if str(c[2]).strip() == str(info["farm"]).strip()), None)
        if pick:
            info["policy"] = pick[8]
            info["farm"] = pick[2] if len(cands) == 1 else info["farm"]
    return info


def _claims_paint_red(ws, row: int, last_col: int) -> None:
    """한 행의 A열~last_col 글씨를 빨간색으로(글꼴·크기·굵기 등은 그대로 두고 색만 바꿈)."""
    from openpyxl.styles import Font
    for c in range(1, last_col + 1):
        cell = ws.cell(row, c)
        f = cell.font
        cell.font = Font(name=f.name, size=f.size, bold=f.bold, italic=f.italic,
                         underline=f.underline, strike=f.strike, vertAlign=f.vertAlign,
                         charset=f.charset, family=f.family, scheme=f.scheme, color="FFFF0000")


def _claims_apply(wb, mails: list[dict]) -> list[dict]:
    """사고내역 시트에 아직 없는 메일 건을 다음 빈 행에 기재. 기재한 건 목록을 반환."""
    ws = wb[CLAIMS_SHEET]
    existing = set()
    total_row = None
    for r in range(4, ws.max_row + 1):
        if str(ws.cell(r, 2).value or "").strip() == "지급 보험금":
            total_row = r
            break
        n = _claims_norm_id(ws.cell(r, 3).value)
        if n:
            existing.add(n)
    added = []
    for m in mails:
        key = _claims_norm_id(m["animal"])
        if key in existing:
            continue
        target = None
        last = total_row or ws.max_row + 1
        for r in range(4, last):
            if all(ws.cell(r, c).value in (None, "") for c in range(2, 11)):  # A열(처리월 병합 라벨)은 제외
                target = r
                break
        if target is None:
            m["error"] = "사고내역 시트에 빈 행이 없음(합계행 위에 행을 추가해 주세요)"
            added.append(m)
            continue
        info = _claims_lookup(wb, key)
        ws.cell(target, 2).value = m["occur"]
        if target > 4 and ws.cell(target - 1, 2).number_format:
            ws.cell(target, 2).number_format = ws.cell(target - 1, 2).number_format
        ws.cell(target, 3).value = _claims_fmt_id(m["animal"])
        ws.cell(target, 4).value = info["farm"] or m["farm"]
        ws.cell(target, 5).value = info["policy"]
        ws.cell(target, 7).value = info["amount"]
        ws.cell(target, 9).value = m["accident"]
        if not ws.cell(target, 12).value and ws.cell(target - 1, 12).value:
            ws.cell(target, 12).value = '="26-"&TEXT(ROW()-3,"000")'
        if info["mark"]:
            # 가입금액을 가져온 시트(4월갱신시개체별 U열 / 월별추징 T열)의 해당 개체 행:
            # 맨 끝 칸에 폐사·긴급출하를 적고, 그 행 전체(A열~끝 칸)를 빨간 글씨로 바꿈(기존 표시 방식과 동일)
            sh, coord = info["mark"]
            mark_cell = wb[sh][coord]
            if mark_cell.value in (None, ""):
                mark_cell.value = m["accident"]
            _claims_paint_red(wb[sh], mark_cell.row, mark_cell.column)
        existing.add(key)
        m.update(row=target, info=info)
        added.append(m)
    return added


def _claims_run_once() -> list[dict] | None:
    """한 번 실행. 기재한 건 목록(없으면 []), 드라이브 연결이 안 되면 None."""
    mails = _claims_fetch_mails()
    if not mails:
        return []
    service = _get_gdrive_service()
    if service is None:
        return None
    meta = service.files().get(fileId=CLAIMS_XLSX_FILE_ID, fields="md5Checksum").execute()
    new_ids = {m["msg_id"] for m in mails} - _claims_state["seen_ids"]
    if not new_ids and meta.get("md5Checksum") == _claims_state["last_md5"]:
        return []  # 새 메일도 없고 파일도 그대로면 다운로드 생략
    content = _gdrive_download_file(service, CLAIMS_XLSX_FILE_ID)
    if not content:
        return None
    wb = openpyxl.load_workbook(io.BytesIO(content))
    added = _claims_apply(wb, mails)
    if any("row" in a for a in added):
        # 다운로드한 뒤에 다른 곳(PC 엑셀 등)에서 파일이 바뀌었으면 덮어쓰지 않고 다음 차례에 다시 시도
        now_md5 = service.files().get(fileId=CLAIMS_XLSX_FILE_ID, fields="md5Checksum").execute().get("md5Checksum")
        if now_md5 != meta.get("md5Checksum"):
            logger.info("사고접수: 저장 직전 파일이 바뀌어 이번 차례는 건너뜀")
            return []
        buf = io.BytesIO()
        wb.save(buf)
        res = service.files().update(
            fileId=CLAIMS_XLSX_FILE_ID,
            media_body=_gdrive_media(buf.getvalue(), mimetype=_XLSX_MIMETYPE),
            fields="md5Checksum",
        ).execute()
        _claims_state["last_md5"] = res.get("md5Checksum")
    else:
        _claims_state["last_md5"] = meta.get("md5Checksum")
    _claims_state["seen_ids"] |= {m["msg_id"] for m in mails}
    return added


async def check_livestock_claims(context: ContextTypes.DEFAULT_TYPE) -> None:
    if not GMAIL_ADDRESS or not GMAIL_APP_PASSWORD:
        return
    try:
        added = await asyncio.to_thread(_claims_run_once)
    except Exception as e:
        logger.exception("사고접수 자동 기재 중 오류")
        await _send_message_retrying(
            context.bot, ALLOWED_USER_ID,
            f"⚠️ 소 사고접수 자동 기재 중 오류: {type(e).__name__}: {str(e)[:150]}",
        )
        return
    if not added:
        return
    lines = ["🐄 소 사고접수 → 「26년 계약리스트」 사고내역에 기재했어요(드라이브 원본, PC 켜지면 자동 반영)"]
    for a in added:
        if "error" in a:
            lines.append(f"⚠️ {a['farm']} / {a['animal']}: {a['error']}")
            continue
        info = a["info"]
        occ = a["occur"].strftime("%m.%d") if a["occur"] else "?"
        amt = f"{info['amount']:,}" if isinstance(info["amount"], (int, float)) else "확인필요"
        lines.append(
            f"• {a['row']}행: {info['farm'] or a['farm']} / {_claims_fmt_id(a['animal'])} / {occ} {a['accident']}"
            f" / 가입금액 {amt} / 증권 {info['policy'] or '확인필요'}"
            + (f" / 사유 {a['reason']}" if a.get("reason") else "")
        )
    lines.append("사고접수번호·처리경과는 직접 입력해 주세요.")
    await _send_message_retrying(context.bot, ALLOWED_USER_ID, "\n".join(lines))


async def check_gdrive_health(context: ContextTypes.DEFAULT_TYPE) -> None:
    """매일 아침 구글 드라이브 refresh token이 아직 유효한지 점검하고, 끊겼을 때에만 유진님께
    텔레그램으로 알림(2026-10-21, 6일 주기 갱신 알림 대체). 일시적인 네트워크 오류로 잘못
    알리지 않도록 한 번 더 시도한 뒤에도 실패할 때만 알림."""
    global _gdrive_service
    if not ALLOWED_USER_ID:
        return
    if not (
        GDRIVE_OAUTH_CLIENT_ID and GDRIVE_OAUTH_CLIENT_SECRET and GDRIVE_OAUTH_REFRESH_TOKEN
        and _gdrive_build
    ):
        return
    last_err = None
    for attempt in range(2):
        try:
            creds = _gdrive_user_credentials(
                token=None,
                refresh_token=GDRIVE_OAUTH_REFRESH_TOKEN,
                client_id=GDRIVE_OAUTH_CLIENT_ID,
                client_secret=GDRIVE_OAUTH_CLIENT_SECRET,
                token_uri="https://oauth2.googleapis.com/token",
            )
            await asyncio.to_thread(creds.refresh, _gdrive_auth_request())
            return  # 정상
        except Exception as e:  # noqa: BLE001
            last_err = e
            if attempt == 0:
                await asyncio.sleep(60)
    logger.error("구글 드라이브 연결 점검 실패: %s", last_err)
    _gdrive_service = None  # 다음 사용 때 새 토큰으로 다시 연결하도록 캐시 비움
    text = (
        "⚠️ 구글 드라이브 연결이 끊겼어요. 파일이 드라이브에 저장되지 않을 수 있어요.\n\n"
        "토큰을 새로 받아야 해요:\n"
        "1. developers.google.com/oauthplayground 접속 → '1단계'에 "
        "https://www.googleapis.com/auth/drive 입력 → 'API 권한 부여' → bubal5503@gmail.com 로그인 "
        "('확인되지 않은 앱' 경고가 나오면 고급 → 이동)\n"
        "2. '2단계'에서 '인증 코드를 토큰으로 교환하기' → refresh_token 값 복사\n"
        "3. Render → yujin-ai-bot → Environment → GDRIVE_OAUTH_REFRESH_TOKEN 값 교체 → Save Changes\n\n"
        "참고: 구글 클라우드 콘솔에서 앱 게시 상태가 '프로덕션'인지도 확인해주세요(테스트면 7일마다 끊겨요)."
    )
    await _send_message_retrying(context.bot, ALLOWED_USER_ID, text)


# ===================== ㈜세정 x 처브 업무 (2026-11-12 보험사 변경 이후) =====================
# 2026-11-12부터 ㈜세정 계약이 삼성화재 -> 처브(에이스아메리칸화재해상보험)로 바뀌면서, 봇이 직접
# 가입증명서를 만들던 예전 로직은 전부 껐음(유진님 지시). 지금 흐름:
#  1) 세정 담당자가 '변경신청서'(가입/해지 신청서 엑셀)를 보내면 -> 처브 담당자 4명에게 메일 발송
#     (확인 버튼 후). 신청서 내용과 보낸 사람은 기억해둠.
#  2) 처브에서 '보험가입증명서'(추가매장) 또는 '배서증명서'(환급/해지) PDF가 오면 ->
#     통합관리리스트의 해당 브랜드 시트(추가매장/해지매장)에 기재, 드라이브 정산관리/브랜드/월 폴더에
#     저장, 그 매장의 세정 담당자(신청서를 보낸 사람)에게 메일 발송(확인 버튼 후).
#     담당자를 모르면 텔레그램으로 유진님께 이름(또는 이메일)을 물어봄.

CHUBB_RECIPIENTS = [
    "Seo-Hyun.Cho@Chubb.com",
    "Ha-Na.Kim@Chubb.com",
    "Chang-Ha.Lee@Chubb.com",
    "Won.jang@Chubb.com",
    "Ji-Hyun.Lee@chubb.com",  # 이지현 (2026-11-16 추가)
]
SEJUNG_BRANDS = ("웰메이드", "올리비아로렌", "트레몰로")
SEJUNG_POLICY_BRAND = {"310D260840": "웰메이드", "310D260841": "올리비아로렌", "310D260842": "트레몰로"}
# 한 메일에 처브 증명서 PDF가 이보다 많으면 원계약 전체 발송으로 보고 자동 기재하지 않음(2026-11-18)
# 한 메일의 PDF가 이 개수보다 많으면 원계약 전체 증명서(대량)로 보고 자동 기재하지 않음.
# 2026-11-24: 5개 → 30개로 늘림(특판 9개 매장 추가 증명서가 막혔음). 원계약 전체는 수백 장이라 계속 걸러짐.
SEJUNG_BULK_CERT_LIMIT = int(os.getenv("SEJUNG_BULK_CERT_LIMIT", "30"))
SEJUNG_BRAND_SYNONYMS = {"월메이드": "웰메이드", "오뷔엘알": "올리비아로렌", "OVLR": "올리비아로렌", "올리비아": "올리비아로렌"}
SEJUNG_REQUESTS_FILE = os.path.join(PERSIST_DIR, "sejung_requests.json")
SEJUNG_AWAIT_FILE = os.path.join(PERSIST_DIR, "sejung_await_contact.json")
SEJUNG_PROCESSED_FILE = os.path.join(PERSIST_DIR, "sejung_processed_certs.json")
SEJUNG_AWAIT_DIR = os.path.join(PERSIST_DIR, "sejung_await_files")
# 드라이브 '정산관리' 폴더 id. 비워두면 이름('정산관리', 상위 폴더 '26년')으로 찾음.
SEJUNG_FOLDER_ID = os.environ.get("SEJUNG_FOLDER_ID")

def _chubb_request_body(items: list, brand: str | None = None) -> str:
    """처브에 보내는 신청 메일 본문. 추가/해지/둘 다에 맞게 문구와 매장 목록을 바꿈(2026-11-16 요청)."""
    adds = [i["store_name"] for i in items if i.get("kind") == "추가"]
    dels = [i["store_name"] for i in items if i.get("kind") == "해지"]
    if adds and dels:
        what = "추가가입 및 해지"
    elif adds:
        what = "추가가입"
    elif dels:
        what = "해지"
    else:
        what = "변경"
    lines = []
    if adds:
        lines.append(f"- 추가가입: {', '.join(adds)}")
    if dels:
        lines.append(f"- 해지: {', '.join(dels)}")
    store_block = ("\n".join(lines) + "\n\n") if lines else ""
    brand_text = f"({brand}) " if brand else ""
    return (
        "안녕하세요\n"
        f"㈜세정{brand_text and ' ' + brand_text}{what} 정산 양식 보내드립니다.\n\n"
        + store_block +
        "진행 부탁드리고, 서류 보내주시면 감사하겠습니다.\n"
        "수고하세요\n\n"
        "--\n" + _CERT_EMAIL_SIGNATURE
    )


def _chubb_request_subject(items: list, brand: str | None, desc: str) -> str:
    kinds = {i.get("kind") for i in items}
    what = "추가가입·해지" if {"추가", "해지"} <= kinds else "추가가입" if "추가" in kinds else "해지" if "해지" in kinds else "변경"
    return f"[㈜세정 {brand or ''}] {what} 정산 양식 송부 ({desc})"


def _sejung_cert_body(kind: str, refund_date: str | None = None, store: str | None = None) -> str:
    what = "추가매장의 보험가입증명서를" if kind == "추가" else "해지매장의 변경증서(환급)를"
    # 어떤 매장의 추가/해지 건인지 본문에 명시(2026-11-16 요청)
    store_line = ""
    if store:
        store_line = (f"- 추가(보험가입증명서): {store}\n" if kind == "추가"
                      else f"- 해지(변경증명서): {store}\n")
    # 해지 건은 처브 메일에 적힌 환급 지급 예정일을 같이 안내(2026-11-16 유진님 요청)
    refund_line = ""
    if kind != "추가" and refund_date:
        mo, d = refund_date.split("/")
        refund_line = f"환급보험료는 {int(mo)}월 {int(d)}일에 지급될 예정입니다.\n"
    return (
        "안녕하십니까~\n"
        f"보내주신 {what} 보내드립니다.\n\n"
        + store_line + refund_line + ("\n" if (store_line or refund_line) else "") +
        "업무에 참고하여 주시기 바랍니다.\n\n"
        "즐거운 하루 보내십시요\n"
        "감사합니다.\n"
        "수고하세요\n\n"
        "--\n" + _CERT_EMAIL_SIGNATURE
    )


def _sj_norm(s) -> str:
    s = unicodedata.normalize("NFC", str(s or ""))
    return re.sub(r"\s+", "", s).lower()


# 처브가 증명서에 매장명을 줄여 쓰는 경우가 있음(예: 신청서 '롯데김해아울렛직영특판점' ->
# 증명서 '롯데김해아울렛특판점', 2026-11-18). 매장 '종류'를 나타내는 말을 빼고 비교함.
_SJ_STORE_NOISE_RX = re.compile(r"직영|상설|특판|매장|점$|[()\[\]（）·\-_]")


def _sj_store_core(s) -> str:
    return _SJ_STORE_NOISE_RX.sub("", _sj_norm(s))


def _sj_store_similar(a, b) -> bool:
    """매장명이 사실상 같은지(종류 표기만 다른지). 빈 값이면 False."""
    ca, cb = _sj_store_core(a), _sj_store_core(b)
    if not ca or not cb:
        return False
    if ca == cb:
        return True
    import difflib
    return difflib.SequenceMatcher(None, ca, cb).ratio() >= 0.9


def _sj_brand_from_text(*texts) -> str | None:
    joined = " ".join(str(t or "") for t in texts)
    for pno, brand in SEJUNG_POLICY_BRAND.items():
        if pno in joined:
            return brand
    norm = unicodedata.normalize("NFC", joined)
    for brand in SEJUNG_BRANDS:
        if brand in norm:
            return brand
    for syn, brand in SEJUNG_BRAND_SYNONYMS.items():
        if syn in norm:
            return brand
    return None


def _sj_load(path: str, default):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return default


def _sj_save(path: str, data) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)


def _sj_json_value(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    return v


# ---------- 변경신청서(엑셀) 읽기 ----------

def _sj_row_by_headers(ws, header_rows: tuple, data_row: int, first_col: int, last_col: int) -> dict:
    """신청서 오른쪽의 '리스트 복사用' 줄을 헤더 이름 -> 값으로 읽음."""
    out = {}
    for col in range(first_col, last_col + 1):
        label = None
        for hr in header_rows:
            v = ws.cell(row=hr, column=col).value
            if v not in (None, ""):
                label = v
                break
        if label is None:
            continue
        key = _norm_header(label)
        val = ws.cell(row=data_row, column=col).value
        if val in (None, "") or (isinstance(val, str) and val.startswith("#")):
            continue
        out[key] = val
    return out


def _sj_parse_request(file_bytes: bytes, filename: str = "") -> list | None:
    """처브 '가입_해지신청서' 엑셀이면 [ {kind, brand, store_name, fields}, ... ] 반환.
    신청서 형식이 아니면 None."""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception:
        return None
    names = wb.sheetnames
    if not any(n.startswith(("추가가입_배서", "해지_배서", "비금액_배서")) for n in names):
        return None

    brand_hint_texts = [filename]
    for n in names:
        ws = wb[n]
        for addr in ("F9", "O8", "D37", "D49"):
            try:
                brand_hint_texts.append(ws[addr].value)
            except Exception:
                pass
    brand = _sj_brand_from_text(*brand_hint_texts)

    items = []
    for n in names:
        ws = wb[n]
        if n.startswith("추가가입_배서"):
            store = ws["F19"].value
            if not store:
                continue
            fields = _sj_row_by_headers(ws, (5, 4), 6, 25, 43)  # Y~AQ
            # 보험기간·주소는 신청서 본문 칸을 우선함. 오른쪽 '리스트 복사用' 줄(AN6/AO6)은 종기가
            # 실제 변경기간과 상관없이 원보험 종기(2027-11-12)로 고정돼 있고 주소(AF6)도 비어 있는 경우가 많음
            # (2026-11-18 WMC 특판 9건: 실제 종기 2027-02-28인데 1년으로 기재되고 주소 누락).
            def _cell(addr):
                v = ws[addr].value
                return v.strip() if isinstance(v, str) else v
            period_start, period_end = _cell("F17"), _cell("L17")
            fields.update({
                "매장명": str(store).strip(),
                "평수": ws["O19"].value or fields.get("평수"),
                "건물": ws["F23"].value,
                "시설/비품": ws["F24"].value,
                "재고": ws["F25"].value,
                "보험시작일": (period_start if _sj_to_date(period_start) else None)
                              or ws["O16"].value or fields.get("보험시작일"),
                "보험종기일": (period_end if _sj_to_date(period_end) else None) or fields.get("보험종기일"),
                "매장주소": _cell("F11") or fields.get("매장주소"),
                "창고주소": _cell("F12") or fields.get("창고주소"),
            })
            items.append({"kind": "추가", "brand": brand, "store_name": str(store).strip(), "fields": fields})
        elif n.startswith("해지_배서"):
            store = ws["D19"].value
            if not store:
                continue
            fields = _sj_row_by_headers(ws, (11, 10), 12, 22, 35)  # V~AI
            fields.update({
                "매장명": str(store).strip(),
                "폐점일": ws["O16"].value or fields.get("폐점일"),
                "평수": ws["D20"].value,
            })
            items.append({"kind": "해지", "brand": brand, "store_name": str(store).strip(), "fields": fields})
        elif n.startswith("비금액_배서"):
            store = ws["F19"].value
            if store:
                items.append({"kind": "비금액", "brand": brand, "store_name": str(store).strip(), "fields": {}})
    return items


def _sj_remember_request(item: dict, sender_email: str | None, sender_name: str | None, filename: str) -> None:
    requests = _sj_load(SEJUNG_REQUESTS_FILE, {})
    key = f"{item.get('brand') or ''}|{_sj_norm(item['store_name'])}"
    requests[key] = {
        "kind": item["kind"],
        "brand": item.get("brand"),
        "store_name": item["store_name"],
        "fields": {k: _sj_json_value(v) for k, v in (item.get("fields") or {}).items()},
        "sender_email": None if _is_own_address(sender_email) else sender_email,
        "sender_name": sender_name,
        "filename": filename,
        "received": dt.date.today().isoformat(),
    }
    _sj_save(SEJUNG_REQUESTS_FILE, requests)


def _sj_find_request(brand: str | None, store_name: str) -> dict | None:
    requests = _sj_load(SEJUNG_REQUESTS_FILE, {})
    target = _sj_norm(store_name)
    best = None
    for key, rec in requests.items():
        if brand and rec.get("brand") and rec.get("brand") != brand:
            continue
        name = _sj_norm(rec.get("store_name"))
        if not name:
            continue
        if name == target:
            return rec
        if name in target or target in name or _sj_store_similar(name, target):
            best = best or rec
    return best


def _queue_confirm_keyboard(cert_id: str) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ 보내기", callback_data=f"sendcert:{cert_id}"),
        InlineKeyboardButton("⏸ 대기", callback_data=f"holdcert:{cert_id}"),
        InlineKeyboardButton("🗑 취소", callback_data=f"cancelcert:{cert_id}"),
    ]])


async def _sj_handle_request_file(bot, filename: str, file_bytes: bytes,
                                  sender_email: str | None, sender_name: str | None) -> bool:
    """변경신청서 엑셀이면 기억해두고 처브 4명에게 보낼지 확인 버튼을 띄움. 처리했으면 True."""
    items = _sj_parse_request(file_bytes, filename)
    if items is None:
        return False
    if not items:
        await _send_message_retrying(
            bot, ALLOWED_USER_ID,
            f"⚠️ '{filename}'은 변경신청서 양식인데 매장명이 비어 있어서 어떤 매장 건인지 못 찾았어요. "
            "그래도 처브로 보낼지 아래 버튼으로 정해주세요.",
        )
    for item in items:
        _sj_remember_request(item, sender_email, sender_name, filename)

    brand = next((i.get("brand") for i in items if i.get("brand")), None) or "브랜드 미확인"
    desc = ", ".join(f"{i['kind']} {i['store_name']}" for i in items) or "매장 미확인"
    subject = _chubb_request_subject(items, brand, desc)
    cert_id = _queue_pending_certs(
        [{"store_name": desc, "out_name": filename, "pdf_bytes": file_bytes}],
        target_email=", ".join(CHUBB_RECIPIENTS), subject=subject, body=_chubb_request_body(items, brand),
        label="변경신청서",
    )
    who = sender_email or "텔레그램 업로드"
    await bot.send_message(
        chat_id=ALLOWED_USER_ID,
        text=(
            f"📨 ㈜세정 변경신청서가 왔어요 ({brand}, {desc}, 보낸 사람: {who}).\n"
            f"처브 담당자 {len(CHUBB_RECIPIENTS)}명({', '.join(CHUBB_RECIPIENTS)})에게 보낼까요?"
        ),
        reply_markup=_queue_confirm_keyboard(cert_id),
    )
    return True


def _sj_pdf_is_request(text: str) -> bool:
    t = (text or "").replace(" ", "").replace("\n", "")
    return "변경승인신청서" in t or "RequestforPolicyAmend" in t


_SJ_REQUEST_WORDS = ("신청", "요청")
_SJ_REQUEST_KINDS = ("가입", "해지", "폐점", "배서", "변경", "추가")


def _sj_extract_request_from_text(subject: str, body: str, pdf_texts: list) -> list:
    """엑셀 신청서가 없을 때(PDF 신청서/본문만 있는 경우) 메일 제목·본문·PDF 글자에서
    매장별 신청 내용을 Claude로 뽑음. [{store_name, kind, fields}] 반환(못 찾으면 빈 목록).
    fields 키는 통합관리리스트 칸 이름(매장주소, 재고, 보험시작일 등)에 맞춤."""
    system = (
        "아래는 ㈜세정 담당자가 보낸 보험 변경신청 메일과 첨부된 '보험계약 변경 승인신청서' PDF 글자예요"
        "(PDF 글자 순서는 뒤섞여 있을 수 있어요). 신청 대상 매장별로 아래 JSON으로만 답하세요.\n"
        '{"stores": [{"store_name": 매장명, "kind": "추가" 또는 "해지" 또는 "비금액", '
        '"address": 소재지 주소, "warehouse": 창고주소, "stock": 재고자산 금액 숫자, "facility": 시설/비품 금액 숫자, '
        '"building": 건물 금액 숫자, "area_py": 면적(평) 숫자, "start_date": 변경기간 시작일 "YYYY-MM-DD", '
        '"end_date": 변경기간 종료일 "YYYY-MM-DD", "effective_date": 유효일 "YYYY-MM-DD", '
        '"property_refund": 재물 환급보험료 숫자, "liability_refund": 배상 환급보험료 숫자, '
        '"refund_account": 환급 계좌/위임받는 분 문구}]}\n'
        "가입/추가가입(소재지 추가)은 추가, 해지/폐점/환급(소재지 제외)은 해지예요. 같은 매장에 해지와 추가가 "
        "같이 있으면 두 개로 나눠 적으세요. 금액이 '-'면 0, 모르면 null로 두세요.\n"
        "start_date/end_date는 반드시 '변경기간(Period)' 칸의 날짜예요. '원보험기간(Policy Period, 예: "
        "2026.11.12 ~ 2027.11.12)'은 계약 전체 기간이라 절대 쓰지 마세요. 변경기간을 못 찾으면 null로 두세요."
    )
    content = f"제목: {subject}\n본문:\n{(body or '')[:3000]}\n\nPDF 내용:\n" + "\n---\n".join(t[:4000] for t in pdf_texts)
    try:
        resp = client.messages.create(
            # 매장이 여러 개(9개 등)면 답이 길어져 1500토큰에서 잘려 JSON이 깨졌음 -> 늘림(2026-11-24)
            model=MODEL_NAME, max_tokens=6000, system=system,
            messages=[{"role": "user", "content": content[:40000]}],
        )
        raw = "".join(b.text for b in resp.content if b.type == "text")
        m = re.search(r"\{.*\}", raw, re.S)
        data = json.loads(m.group(0)) if m else {}
    except Exception:
        logger.exception("변경신청 메일 내용 읽기 중 오류")
        return []
    out = []
    for st in data.get("stores") or []:
        name = str(st.get("store_name") or "").strip()
        if not _sj_valid_store_name(name):
            continue
        kind = st.get("kind") if st.get("kind") in ("추가", "해지", "비금액") else "추가"
        f = {"매장명": name, "매장주소": st.get("address"), "창고주소": st.get("warehouse")}
        if kind == "추가":
            f.update({"재고": st.get("stock"), "시설/비품": st.get("facility"), "건물": st.get("building"),
                      "평수": st.get("area_py"), "보험시작일": st.get("start_date"), "보험종기일": st.get("end_date")})
        else:
            f.update({"폐점일": st.get("effective_date") or st.get("start_date"),
                      "재물부문환급": st.get("property_refund"), "영업부문환급": st.get("liability_refund")})
            if st.get("refund_account"):
                f["비고"] = f"환급: {st.get('refund_account')}"
        out.append({"store_name": name, "kind": kind, "fields": {k: v for k, v in f.items() if v not in (None, "")}})
    return out


# ---------- PDF 신청서를 규칙대로 직접 읽기 (2026-11-18 유진님 요청) ----------
# 신청서가 PDF로 와도 엑셀과 같은 기준으로 기재: 보험기간은 '변경기간(Period)' 칸, 폐점일은 '유효일',
# 주소는 '소재지'·'창고주소' 칸. Claude가 뽑은 값보다 이 규칙값을 우선함(Claude가 원보험기간
# 2027-11-12를 종기로 잘못 쓰는 경우를 막기 위함).
_SJ_PDF_D = r"(\d{4})\s*[년.\-/]\s*(\d{1,2})\s*[월.\-/]\s*(\d{1,2})\s*일?"


def _sj_pdf_date(m, i: int):
    try:
        return dt.datetime(int(m.group(i)), int(m.group(i + 1)), int(m.group(i + 2))).strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        return None


def _sj_pdf_clean_addr(s):
    # 신청서 칸이 좁아 주소가 글자 단위로 줄바꿈된 것을 붙임(2026-11-26).
    # 예) '…장유로 469, 롯\n데아울렛김해점' -> '…장유로 469, 롯데아울렛김해점' (예전엔 '롯 데'로 띄어져 저장됨)
    s = re.sub(r"[ \t]*\n[ \t]*", "", s or "")
    s = re.sub(r"\s+", " ", s).strip()
    # 칸이 비어서 다른 글자까지 잡힌 경우는 버림
    if not s or len(s) > 120 or re.search(r"Policy|KRW|Insured|\d{3}-\d{2}-\d{5}|일간|매장명", s):
        return None
    return s


def _sj_valid_store_name(name) -> bool:
    """매장명으로 쓸 수 있는 글자인지. '~ 일간', '면적', '(P)'처럼 양식 글자를 매장명으로 잘못 읽은 경우를 거름
    (2026-11-23 웰메이드 특판 9개 신청서가 모두 '~ 일간'으로 읽혀 처브 메일 제목까지 잘못 나간 건)."""
    s = re.sub(r"\s+", "", str(name or ""))
    if len(s) < 2 or not re.search(r"[가-힣A-Za-z]", s):
        return False
    bad = ("일간", "면적", "(P)", "Section", "Sum", "Insured", "Premium", "목적물", "변경기간", "~")
    return not any(b in s for b in bad)


_SJ_CODE_RX = r"[A-Z]{2,4}\d[A-Z0-9]{2,5}"
_SJ_FILENAME_SKIP = ("오뷔엘알", "올리비아로렌", "웰메이드", "트레몰로", "세정", "신청서", "증명서", "가입", "해지")


def _sj_store_from_filename(fn: str | None) -> str | None:
    """파일명에서 매장명. 'WMO0171_WMCNC고잔점(특판)_26.11.12.pdf' -> 'WMCNC고잔점(특판)',
    '◆ 오뷔엘알_OLM7201_OL롯데김해아울렛상설특판점_20261126.pdf' -> 'OL롯데김해아울렛상설특판점'
    (2026-11-26: 앞에 '◆ 브랜드'가 붙은 파일명에서 브랜드를 매장명으로 읽던 문제 수정 - 매장코드 바로 뒤 칸을 우선)."""
    base = re.sub(r"\.[A-Za-z0-9]+$", "", unicodedata.normalize("NFC", fn or ""))
    base = re.sub(r"^\[(?:가입|해지)\]", "", base)
    parts = [p.strip(" ◆◇■□●○") for p in base.split("_")]
    parts = [p for p in parts if p]
    ok = lambda p: (_sj_valid_store_name(p) and re.search(r"[가-힣]", p)
                    and not re.fullmatch(_SJ_CODE_RX, p) and not re.fullmatch(r"[\d.\-]+", p)
                    and p not in _SJ_FILENAME_SKIP)
    for i, p in enumerate(parts):
        if re.fullmatch(_SJ_CODE_RX, p) and i + 1 < len(parts) and ok(parts[i + 1]):
            return parts[i + 1]
    for p in parts:
        if ok(p):
            return p
    return None


def _sj_code_from_filename(fn: str | None) -> str | None:
    """파일명의 매장코드. 'WMO01A1_WMC천안모다…pdf' -> 'WMO01A1', '◆ 오뷔엘알_OLM7201_…pdf' -> 'OLM7201'
    (2026-11-26: 맨 앞 칸만 보던 것을 모든 칸에서 찾도록 수정)."""
    base = re.sub(r"\.[A-Za-z0-9]+$", "", unicodedata.normalize("NFC", fn or ""))
    base = re.sub(r"^\[(?:가입|해지)\]", "", base)
    for p in base.split("_"):
        p = p.strip(" ◆◇■□●○")
        if re.fullmatch(_SJ_CODE_RX, p):
            return p
    return None


def _sj_layout_request_fields(page) -> dict:
    """처브 '보험계약 변경 승인신청서' 한 장을 표의 칸 위치 그대로(layout) 읽어 칸별 값을 꺼냄(2026-11-26).
    '매장명' 칸 값(옆의 '면적' 칸 전까지), 면적(평수), 재물 가입금액(건물 / 시설(인테리어)/비품 / 재고자산).
    빈 칸이면 그 값은 넣지 않음('-'는 0)."""
    try:
        t = unicodedata.normalize("NFC", page.extract_text(extraction_mode="layout") or "")
    except Exception:
        logger.exception("신청서 표(layout) 읽기 실패")
        return {}
    out = {}

    def krw(label):
        m = re.search(r"(?m)^[ \t]*" + label + r"[ \t]{2,}(KRW[ \t]*[\d,]+|-)(?=[ \t]|$)", t)
        if not m:
            return None
        v = m.group(1)
        return 0 if v == "-" else int(re.sub(r"\D", "", v))

    m = re.search(r"(?m)^[ \t]*매장명[ \t]{2,}(\S(?:.*?\S)?)[ \t]{2,}(?:면적|가입[ \t]*시[ \t]*납입보험료)", t)
    if m and _sj_valid_store_name(m.group(1)):
        out["매장명"] = m.group(1).strip()
    m = re.search(r"면적(?:[ \t]*\(P\))?[ \t]+([\d,.]+)(?=[ \t]|$)", t, re.M)
    if m:
        v = m.group(1).replace(",", "")
        try:
            out["평수"] = float(v) if "." in v else int(v)
        except ValueError:
            pass
    for key, label in (("건물", r"건물"), ("시설/비품", r"시설\(인테리어\)[ \t]*/[ \t]*비품"), ("재고", r"재고자산")):
        v = krw(label)
        if v is not None:
            out[key] = v
    return out


def _sj_parse_request_pdf(file_bytes: bytes, filename: str | None = None) -> list:
    """PDF '보험계약 변경 승인신청서'를 장(page)별로 읽어 [{kind, store_name, fields}] 반환.
    매장명을 못 읽은 장도 store_name=None으로 넣음(Claude 결과와 맞춰 쓰기 위해).
    2026-11-24: 매장명은 '매장명' 칸 글자를 먼저 쓰고, 못 읽으면 파일명(매장코드_매장명_날짜)에서 가져옴.
    양식 글자('~ 일간' 등)는 매장명으로 쓰지 않음."""
    try:
        reader = PdfReader(io.BytesIO(file_bytes))
        pages = [unicodedata.normalize("NFC", p.extract_text() or "") for p in reader.pages]
    except Exception:
        return []
    file_store = _sj_store_from_filename(filename)
    file_code = _sj_code_from_filename(filename)
    out = []
    for page_idx, t in enumerate(pages):
        lay = _sj_layout_request_fields(reader.pages[page_idx])
        flat = t.replace(" ", "").replace("\n", "")
        if "변경승인신청서" not in flat and "RequestforPolicyAmend" not in flat:
            continue
        if "소재지추가" in flat or "IncreasedTSI" in flat:
            kind = "추가"
        elif "소재지제외" in flat or "DecreasedTSI" in flat:
            kind = "해지"
        else:
            continue
        f = {}
        m = re.search(r"변경기간\s*Period\s*" + _SJ_PDF_D + r"\s*~?\s*" + _SJ_PDF_D, t)
        start, end = (_sj_pdf_date(m, 1), _sj_pdf_date(m, 4)) if m else (None, None)
        m = re.search(r"유효일\s*Ef+ective\s*Date\s*" + _SJ_PDF_D, t, re.I)
        effective = _sj_pdf_date(m, 1) if m else None
        m = re.search(r"소재지\s*Location\s*of\s*Insured\s*(.*?)\s*원보험기간", t, re.S)
        addr = _sj_pdf_clean_addr(m.group(1)) if m else None
        # PDF 만드는 방식에 따라 '창고주소 건물 내 621-81-19659'처럼 사업자번호와 한 줄로 붙기도 함
        m = re.search(r"창고주소\s*(.*?)\s*(?:법인등록번호|\d{3}-\d{2}-\d{5})", t, re.S)
        wh = _sj_pdf_clean_addr(m.group(1)) if m else None
        # 1순위(2026-11-26): 신청서 표를 칸 위치 그대로 읽은 '매장명' 칸 값.
        # 그냥 글자만 뽑으면 '매장명 OL롯데김해아울렛상설특판점 24'처럼 옆 칸(면적 24)이 붙어 매장이 2개로 읽혔음
        store = lay.get("매장명")
        if kind == "추가":
            # 2순위: 한 장짜리 신청서면 파일명(매장코드_매장명_날짜)의 매장명
            if not store and file_store and len(pages) == 1:
                store = file_store
            # 3순위: 글자만 뽑은 '매장명 …' 줄(면적 칸 앞까지만)
            if not store:
                m = re.search(r"매장명[ \t]*([^\n]+)", t)
                if m:
                    cand = re.split(r"[ \t]{2,}|면적|\(P\)", m.group(1).strip())[0].strip()
                    if _sj_valid_store_name(cand):
                        store = cand
            # 4순위: 예전 방식('면적 (P)' 앞 줄) - 양식 글자면 버림
            if not store:
                m = re.search(r"(?:^|\n)([^\n]*?)\s*면적\s*\n?\s*\(P\)", t)
                if m and _sj_valid_store_name(m.group(1)):
                    store = m.group(1).strip()
            # 표에서 읽은 평수·재물 가입금액(재고/시설·비품/건물)은 AI가 읽은 값보다 우선
            for k in ("평수", "재고", "시설/비품", "건물"):
                if lay.get(k) is not None:
                    f[k] = lay[k]
            if start or effective:
                f["보험시작일"] = start or effective
            if end:
                f["보험종기일"] = end
            if wh:
                f["창고주소"] = wh
        else:
            if not store:
                m = re.search(r"매장명[ \t]+([^\n]+?)[ \t]+가입\s*시\s*납입보험료", t)
                if m and _sj_valid_store_name(m.group(1)):
                    store = m.group(1).strip()
            if effective or start:
                f["폐점일"] = effective or start
        if not store and file_store and len(pages) == 1:
            store = file_store  # 한 장짜리 신청서는 파일명의 매장명을 씀
        if file_code and len(pages) == 1:
            f["매장코드"] = file_code  # 신청서 파일명 'WMO0171_매장명_날짜.pdf'의 매장코드(2026-11-24)
        if not store and not addr:
            continue  # 매장명·주소가 다 빈 장은 양식만 붙어 온 빈 장으로 봄
        if addr:
            f["매장주소"] = addr
        if store:
            f["매장명"] = store
        out.append({"kind": kind, "store_name": store, "fields": f})
    return out


def _sj_apply_pdf_rules(items: list, rule_items: list) -> list:
    """Claude가 뽑은 매장 목록(items)에 PDF 규칙값(rule_items)을 덮어씀.
    매장명이 비슷하면 그 장 값으로, 매장명을 못 읽었으면 같은 종류가 한 장씩뿐일 때만 맞춤.
    Claude가 놓친 매장(규칙으로 매장명을 읽은 장)은 새로 추가함."""
    used = set()
    for it in items:
        cands = [i for i, r in enumerate(rule_items) if i not in used and r["kind"] == it["kind"]]
        hit = next((i for i in cands if rule_items[i]["store_name"] and (
            _sj_norm(rule_items[i]["store_name"]) == _sj_norm(it["store_name"])
            or _sj_store_similar(rule_items[i]["store_name"], it["store_name"]))), None)
        if hit is None:
            same_kind_items = [x for x in items if x["kind"] == it["kind"]]
            if len(cands) == 1 and len(same_kind_items) == 1 and not rule_items[cands[0]]["store_name"]:
                hit = cands[0]
        if hit is None:
            continue
        used.add(hit)
        fields = it.setdefault("fields", {})
        for k, v in rule_items[hit]["fields"].items():
            if k == "매장명":
                continue
            # 주소는 칸 줄바꿈 때문에 띄어쓰기를 정확히 알 수 없어서('…매장\n앞)'이 '매장 앞)'인지 '매장앞)'인지),
            # AI가 읽은 주소와 띄어쓰기만 다르면 AI 쪽(문맥으로 띄어쓴 주소)을 씀(2026-11-26)
            if (k in ("매장주소", "창고주소") and fields.get(k)
                    and re.sub(r"\s", "", str(fields[k])) == re.sub(r"\s", "", str(v))):
                continue
            fields[k] = v
    for i, r in enumerate(rule_items):
        if i in used or not r["store_name"]:
            continue
        if any(x["kind"] == r["kind"] and _sj_store_similar(x["store_name"], r["store_name"]) for x in items):
            continue
        items.append({"kind": r["kind"], "store_name": r["store_name"], "fields": dict(r["fields"])})
    return items


def _sj_list_info(wb, store_name: str) -> dict:
    """최초가입전체리스트에서 매장코드/대표자명/사업자번호/전화번호/주소를 찾아줌(없으면 빈 dict)."""
    target = _sj_norm(store_name)
    for n in wb.sheetnames:
        if "최초가입" not in unicodedata.normalize("NFC", n):
            continue
        ws = wb[n]
        for r in range(4, ws.max_row + 1):
            if _sj_norm(ws.cell(row=r, column=5).value) == target:
                g = lambda c: ws.cell(row=r, column=c).value
                return {k: v for k, v in {"매장코드": g(3), "대표자명": g(6), "사업자번호": g(7),
                                           "전화번호": g(13), "매장주소": g(14)}.items() if v not in (None, "")}
    return {}


# ---------- 중복 방지 (2026-11-23 유진님 요청) ----------
# 세정 담당자가 같은 신청 메일을 두 번 보내도 통합관리리스트에 같은 매장이 두 번 적히지 않게 함.
#  1단계: 같은 메일(Message-ID)·같은 첨부(SHA-256)는 아예 처리하지 않음 -> 처리이력은 드라이브 JSON
#        (Render는 재배포하면 로컬 파일이 지워져서 드라이브 '가입증명서및정산' 폴더에 둠)
#  2단계: 행을 적기 직전에 브랜드+매장명+구분+개시일/폐점일로 같은 건을 찾아
#        ① 금액까지 같으면 건너뜀 ② 일자·금액이 다르면 [덮어쓰기][새로 추가][무시] 버튼 ③ 없으면 기재
#  3단계: 건너뛴 매장은 드라이브 재저장·처브 발송을 하지 않고, 이미 보낸 매장은 보내기 전에 한 번 더 확인
SEJUNG_HISTORY_NAME = "세정_처리이력.json"
SEJUNG_DUP_CONFIRM_FILE = os.path.join(PERSIST_DIR, "sejung_dup_confirm.json")
# 같은 건인지 볼 때 '보험료'를 정하는 칸들(이 값이 다르면 정정일 수 있어 확인받음)
_SJ_DUP_AMOUNT_KEYS = {
    "추가": ("재고", "시설/비품", "건물", "평수", "보험종기일"),
    "해지": ("재물부문환급", "영업부문환급"),
}
_SJ_KEY_PAREN_RX = re.compile(r"[(\[（【][^)\]）】]*[)\]）】]")


def _sj_store_key(name) -> str:
    """중복 비교용 매장명: 괄호(안 글자 포함)·공백을 빼고, 끝의 '점'을 뗌. 예) '롯데 김해점' = '롯데김해'."""
    s = unicodedata.normalize("NFC", str(name or ""))
    s = _SJ_KEY_PAREN_RX.sub("", s)
    s = re.sub(r"[\s()\[\]（）【】]", "", s).lower()
    if len(s) > 1 and s.endswith("점"):
        s = s[:-1]
    return s


def _sj_dup_num(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        t = re.sub(r"[,\s원]", "", v)
        if t in ("", "-"):
            return 0.0 if t == "-" else None
        try:
            return float(t)
        except ValueError:
            return None
    return None


def _sj_dup_same(a, b) -> bool:
    """두 칸 값이 같은지(날짜는 날짜끼리, 금액은 쉼표 빼고 1원 차이까지 같다고 봄)."""
    if a in (None, "") and b in (None, ""):
        return True
    da, db = _sj_to_date(a), _sj_to_date(b)
    if da and db:
        return da.date() == db.date()
    na, nb = _sj_dup_num(a), _sj_dup_num(b)
    if na is not None and nb is not None:
        return abs(na - nb) <= 1
    if a in (None, "") or b in (None, ""):
        # 한쪽만 비어 있으면 다름(단, 금액 0과 빈칸은 같게 봄)
        other = b if a in (None, "") else a
        return _sj_dup_num(other) == 0.0
    return _sj_norm(a) == _sj_norm(b)


def _sj_master_sheet(wb, brand: str, kind: str):
    """통합관리리스트에서 '{브랜드}_추가매장'/'{브랜드}_해지매장' 시트와 헤더(칸 이름->열 번호)."""
    want = "추가" if kind == "추가" else "해지"
    for n in wb.sheetnames:
        nn = unicodedata.normalize("NFC", n)
        if brand and brand in nn and want in nn:
            ws = wb[n]
            header = {}
            for col in range(1, ws.max_column + 1):
                label = ws.cell(row=3, column=col).value or ws.cell(row=2, column=col).value
                if label not in (None, ""):
                    header.setdefault(_norm_header(label), col)
            return ws, header
    return None, {}


def _sj_dup_key(item_or_values: dict, kind: str | None = None, brand: str | None = None) -> str:
    """브랜드 + 매장명 + 구분 + 개시일(추가)/폐점일(해지)."""
    v = item_or_values
    kind = kind or v.get("kind")
    brand = brand or v.get("brand") or ""
    fields = v.get("fields") if isinstance(v.get("fields"), dict) else v
    name = v.get("store_name") or fields.get("매장명")
    d = _sj_to_date(fields.get("보험시작일" if kind == "추가" else "폐점일"))
    return f"{brand}|{_sj_store_key(name)}|{kind}|{d.date().isoformat() if d else ''}"


def _sj_find_dup(wb, brand: str, kind: str, values: dict) -> dict:
    """통합관리리스트에 같은 건이 있는지 봄.
    반환 {"status": "new"} / {"status": "same", "sheet", "row"} / {"status": "differ", "sheet", "row", "diffs"}"""
    ws, header = _sj_master_sheet(wb, brand, kind)
    name_col = header.get("매장명")
    key = _sj_store_key(values.get("매장명"))
    if ws is None or not name_col or not key:
        return {"status": "new"}
    date_key = "보험시작일" if kind == "추가" else "폐점일"
    dcol = header.get(date_key)
    want_d = _sj_to_date(values.get(date_key))
    best = None
    for r in range(4, ws.max_row + 1):
        if _sj_store_key(ws.cell(row=r, column=name_col).value) != key:
            continue
        diffs = []
        if want_d is not None and dcol:
            cd = _sj_to_date(ws.cell(row=r, column=dcol).value)
            if cd is None or cd.date() != want_d.date():
                diffs.append(f"{date_key}: {_sj_fmt(cd)} → {_sj_fmt(want_d)}")
        for k in _SJ_DUP_AMOUNT_KEYS.get(kind, ()):
            nv = values.get(k)
            col = header.get(k)
            if nv in (None, "") or not col:
                continue
            ov = ws.cell(row=r, column=col).value
            if isinstance(ov, str) and ov.startswith("="):
                continue
            if not _sj_dup_same(ov, nv):
                diffs.append(f"{k}: {_sj_fmt(ov)} → {_sj_fmt(nv)}")
        if not diffs:
            return {"status": "same", "sheet": ws.title, "row": r}
        if best is None or len(diffs) < len(best[1]):
            best = (r, diffs)
    if best is None:
        return {"status": "new"}
    return {"status": "differ", "sheet": ws.title, "row": best[0], "diffs": best[1]}


def _sj_dedupe_items(items: list) -> tuple[list, list]:
    """한 번에 들어온 신청 안에서 같은 매장이 두 번 적혀 온 경우 첫 번째만 남김. (남긴 목록, 뺀 목록)"""
    seen, keep, dropped = set(), [], []
    for it in items:
        if it.get("kind") not in ("추가", "해지"):
            keep.append(it)
            continue
        k = _sj_dup_key(it)
        if k in seen:
            dropped.append(it)
            continue
        seen.add(k)
        keep.append(it)
    return keep, dropped


def _sj_err_where() -> str:
    """오류가 난 함수와 줄 번호(텔레그램 오류 알림용)."""
    import traceback as _tb
    tb = sys.exc_info()[2]
    last = _tb.extract_tb(tb)[-1] if tb else None
    return f"{last.name}() {last.lineno}번째 줄" if last else "위치 불명"


# ----- 처리이력(드라이브 JSON) -----

def _sj_history_folder(service) -> str | None:
    return GDRIVE_FOLDER_ID or _sj_find_root(service)


def _sj_history_load(service) -> tuple[dict, str | None]:
    """(이력 dict, 드라이브 파일 id). 파일이 없으면 빈 이력과 None."""
    data = {"mails": {}, "files": {}, "sent": {}}
    folder = _sj_history_folder(service)
    if not folder:
        return data, None
    resp = service.files().list(
        q=f"'{folder}' in parents and name='{SEJUNG_HISTORY_NAME}' and trashed=false",
        fields="files(id,name,modifiedTime)", orderBy="modifiedTime desc", pageSize=5,
    ).execute()
    files = resp.get("files", [])
    if not files:
        return data, None
    raw = _gdrive_download_file(service, files[0]["id"])
    if raw:
        try:
            loaded = json.loads(raw.decode("utf-8"))
            for k in data:
                if isinstance(loaded.get(k), dict):
                    data[k] = loaded[k]
        except (ValueError, UnicodeDecodeError):
            logger.warning("처리이력 JSON을 읽지 못해 빈 이력으로 시작함")
    return data, files[0]["id"]


def _sj_history_save(service, data: dict, file_id: str | None) -> str | None:
    body = json.dumps(data, ensure_ascii=False, indent=1).encode("utf-8")
    media = _gdrive_media(body, mimetype="application/json")
    if file_id:
        service.files().update(fileId=file_id, media_body=media).execute()
        return file_id
    folder = _sj_history_folder(service)
    if not folder:
        return None
    f = service.files().create(
        body={"name": SEJUNG_HISTORY_NAME, "parents": [folder]}, media_body=media, fields="id",
    ).execute()
    return f.get("id")


def _sj_history_check(message_id: str | None, hashes: list) -> dict | None:
    """이미 처리한 메일/첨부면 최초 처리 기록을, 아니면 None. 드라이브 연결이 안 되면 None(그냥 처리)."""
    try:
        service = _get_gdrive_service()
        if service is None:
            return None
        hist, _ = _sj_history_load(service)
    except Exception:
        logger.exception("처리이력 확인 중 오류(이력 없이 계속 처리)")
        return None
    if message_id and message_id in hist["mails"]:
        return hist["mails"][message_id]
    if hashes and all(h in hist["files"] for h in hashes):
        return hist["files"][hashes[0]]
    return None


def _sj_history_record(message_id: str | None, file_hashes: list, sender: str, stores: list) -> None:
    """처리한 메일 Message-ID와 첨부 해시를 드라이브 처리이력에 남김. file_hashes=[(해시, 파일명)]."""
    try:
        service = _get_gdrive_service()
        if service is None:
            return
        hist, fid = _sj_history_load(service)
        now = dt.datetime.now(KST).strftime("%Y-%m-%d %H:%M")
        rec = {"at": now, "sender": sender, "stores": stores}
        if message_id:
            hist["mails"][message_id] = rec
        for h, fn in file_hashes:
            hist["files"].setdefault(h, dict(rec, name=fn))
        _sj_history_save(service, hist, fid)
    except Exception:
        logger.exception("처리이력 저장 중 오류")


def _sj_sent_lookup(keys: list) -> dict:
    """이미 발송한 매장 {키: {"at", "to"}} (드라이브 처리이력의 sent)."""
    if not keys:
        return {}
    try:
        service = _get_gdrive_service()
        if service is None:
            return {}
        hist, _ = _sj_history_load(service)
    except Exception:
        logger.exception("발송이력 확인 중 오류")
        return {}
    return {k: hist["sent"][k] for k in keys if k in hist["sent"]}


def _sj_sent_record(keys: list, target: str) -> None:
    if not keys:
        return
    try:
        service = _get_gdrive_service()
        if service is None:
            return
        hist, fid = _sj_history_load(service)
        now = dt.datetime.now(KST).strftime("%Y-%m-%d %H:%M")
        for k in keys:
            hist["sent"][k] = {"at": now, "to": target}
        _sj_history_save(service, hist, fid)
    except Exception:
        logger.exception("발송이력 저장 중 오류")


def _sj_sent_key(who: str, brand: str | None, kind: str, store: str) -> str:
    """who: '처브'(신청서 발송) 또는 '세정'(증명서 발송)."""
    return f"{who}|{brand or ''}|{kind}|{_sj_store_key(store)}"


def _sj_sha256(data: bytes) -> str:
    import hashlib
    return hashlib.sha256(data).hexdigest()


# ----- ② 확인 버튼 -----

def _sj_dup_keyboard(cid: str) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("✏️ 덮어쓰기", callback_data=f"sjdup:over:{cid}"),
        InlineKeyboardButton("➕ 새로 추가", callback_data=f"sjdup:new:{cid}"),
        InlineKeyboardButton("🚫 무시", callback_data=f"sjdup:skip:{cid}"),
    ]])


def _sj_dup_apply(entry: dict, action: str) -> str:
    """확인 버튼 결과를 통합관리리스트에 반영하고 결과 문구를 돌려줌."""
    service = _get_gdrive_service()
    root = _sj_find_root(service) if service else None
    master = _sj_find_master(service, root) if root else None
    if master is None:
        return "❌ 드라이브 통합관리리스트를 못 찾아 반영하지 못했어요."
    wb = openpyxl.load_workbook(io.BytesIO(_gdrive_download_file(service, master["id"])), data_only=False)
    brand, kind, values = entry["brand"], entry["kind"], entry["values"]
    if action == "over":
        # 버튼을 누르기 전에 줄이 옮겨졌을 수 있어 같은 매장 줄인지 다시 확인
        ws, header = _sj_master_sheet(wb, brand, kind)
        row = entry.get("row")
        ncol = header.get("매장명")
        if not (ws and ncol and row and _sj_store_key(ws.cell(row=row, column=ncol).value)
                == _sj_store_key(values.get("매장명"))):
            found = _sj_find_dup(wb, brand, kind, values)
            row = found.get("row")
        if not row:
            return "⚠️ 덮어쓸 기존 줄을 찾지 못했어요. [새로 추가]로 다시 처리해 주세요."
        # 접수일자는 처음 적은 날 그대로 둠(월별 묶음이 바뀌지 않게)
        vals = {k: v for k, v in values.items() if k != "접수일자"}
        where = _sj_write_master(wb, brand, kind, vals, force_row=row)
        result = f"✏️ {entry['store_name']}: {where}을(를) 새 내용으로 덮어썼어요."
    else:
        where = _sj_write_master(wb, brand, kind, values, update_match=False)
        result = f"➕ {entry['store_name']}: {where}에 새로 적었어요."
    out = io.BytesIO()
    wb.save(out)
    if not _save_brand_master_to_drive(master["id"], out.getvalue()):
        return "❌ 통합관리리스트 저장에 실패했어요."
    return result


async def _sj_handle_dup_callback(query) -> None:
    """[덮어쓰기][새로 추가][무시] 버튼 처리. callback_data = 'sjdup:<over|new|skip>:<id>'"""
    await query.answer()
    try:
        _, action, cid = (query.data or "").split(":", 2)
    except ValueError:
        return
    waiting = _sj_load(SEJUNG_DUP_CONFIRM_FILE, {})
    entry = waiting.pop(cid, None)
    if entry is None:
        await query.edit_message_text("이미 처리됐거나 없는 확인 요청이에요.")
        return
    _sj_save(SEJUNG_DUP_CONFIRM_FILE, waiting)
    head = f"{entry['brand']} {entry['kind']} {entry['store_name']}"
    if action == "skip":
        await query.edit_message_text(f"🚫 {head}: 무시했어요(통합관리리스트는 그대로).")
        return
    try:
        msg = await asyncio.to_thread(_sj_dup_apply, entry, action)
    except Exception as e:
        logger.exception("중복 확인 반영 중 오류")
        where = _sj_err_where()
        waiting = _sj_load(SEJUNG_DUP_CONFIRM_FILE, {})
        waiting[cid] = entry  # 실패하면 다시 누를 수 있게 되돌려 둠
        _sj_save(SEJUNG_DUP_CONFIRM_FILE, waiting)
        await query.edit_message_text(
            f"⚠️ {head} 반영 중 오류: {type(e).__name__} ({where}): {str(e)[:120]}\n다시 눌러 주세요.",
            reply_markup=_sj_dup_keyboard(cid),
        )
        return
    await query.edit_message_text(f"{head}\n{msg}")


def _sj_write_requests_to_master(items: list, sender_name: str | None) -> list:
    """신청서가 들어오면 바로 통합관리리스트(드라이브 정산관리 폴더)에 기재(2026-11-12 유진님 요청).
    나중에 처브 증명서가 오면 같은 줄을 증명서 값으로 갱신함. 결과 문구 목록 반환.
    2026-11-23 중복 방지: 적기 직전에 같은 건을 찾아서 각 item['_dup']에 결과를 남김
      'same'(이미 기재됨, 건너뜀) / 'confirm'(일자·금액이 달라 버튼으로 확인) / 'new'(새로 기재)
    확인이 필요한 건은 item['_confirm']에 버튼 메시지 정보를 담아 둠(보내는 건 호출한 쪽)."""
    msgs = []
    items = [i for i in items if i.get("kind") in ("추가", "해지") and i.get("brand")]
    if not items:
        return ["ℹ️ 기재할 추가/해지 매장이나 브랜드를 확인하지 못해 통합관리리스트에는 적지 않았어요."]
    service = _get_gdrive_service()
    if service is None:
        return ["❌ 구글 드라이브에 연결되지 않아 통합관리리스트에 적지 못했어요."]
    root = _sj_find_root(service)
    master = _sj_find_master(service, root) if root else None
    if master is None:
        return ["❌ 드라이브 정산관리 폴더에서 통합관리리스트를 못 찾아 적지 못했어요."]
    try:
        wb = openpyxl.load_workbook(io.BytesIO(_gdrive_download_file(service, master["id"])), data_only=False)
        wrote = False
        confirm_waiting = _sj_load(SEJUNG_DUP_CONFIRM_FILE, {})
        for item in items:
            info = _sj_list_info(wb, item["store_name"])
            values = {"접수자": sender_name, "접수일자": dt.datetime.now(KST).date().isoformat()}
            for k, v in info.items():
                if item["kind"] == "추가" and k == "매장주소":
                    continue  # 추가 건 주소는 신청서의 새 소재지를 씀
                values[k] = v
            values.update({k: _sj_json_value(v) for k, v in (item.get("fields") or {}).items()})
            values["매장명"] = item["store_name"]
            dup = _sj_find_dup(wb, item["brand"], item["kind"], values)
            if dup["status"] == "same":
                item["_dup"] = "same"
                msgs.append(f"⏭ {item['kind']} {item['store_name']}: 이미 기재됨({dup['sheet']} {dup['row']}행)")
                continue
            if dup["status"] == "differ":
                cid = uuid.uuid4().hex[:10]
                confirm_waiting[cid] = {
                    "brand": item["brand"], "kind": item["kind"], "store_name": item["store_name"],
                    "values": {k: _sj_json_value(v) for k, v in values.items()},
                    "row": dup["row"], "sheet": dup["sheet"], "diffs": dup["diffs"],
                    "at": dt.datetime.now(KST).isoformat(),
                }
                item["_dup"] = "confirm"
                item["_confirm"] = {"id": cid, "sheet": dup["sheet"], "row": dup["row"], "diffs": dup["diffs"]}
                msgs.append(f"❓ {item['kind']} {item['store_name']}: 이미 적힌 줄({dup['sheet']} {dup['row']}행)과 "
                            "일자·금액이 달라서 확인이 필요해요(아래 버튼).")
                continue
            where = _sj_write_master(wb, item["brand"], item["kind"], values, update_match=False)
            item["_dup"] = "new"
            wrote = True
            msgs.append(f"✅ {item['kind']} {item['store_name']}: 통합관리리스트 {where}에 적었어요.")
        _sj_save(SEJUNG_DUP_CONFIRM_FILE, confirm_waiting)
        if wrote:
            out = io.BytesIO()
            wb.save(out)
            if not _save_brand_master_to_drive(master["id"], out.getvalue()):
                for it in items:
                    if it.get("_dup") == "new":
                        it["_dup"] = "error"
                return ["❌ 통합관리리스트 저장에 실패했어요."]
    except Exception as e:
        logger.exception("신청서 통합관리리스트 기재 중 오류")
        return [f"❌ 통합관리리스트 기재 실패: {type(e).__name__} ({_sj_err_where()}): {str(e)[:150]}"]
    return msgs


async def _sj_handle_request_mail(bot, msg, subject: str, body: str,
                                  sender_email: str | None, sender_name: str | None) -> bool:
    """세정 담당자의 변경신청 메일이면(엑셀 신청서, PDF 신청서, 또는 '가입/해지 신청' 메일+첨부)
    첨부파일 전부(신청서+사업자등록증/통장사본/동의서 등)를 처브 4명에게 보낼지 확인 버튼을 띄우고,
    매장별로 보낸 사람을 기억해둠. 처리했으면 True.
    (예전엔 엑셀 신청서만 알아봐서, PDF 신청서로 온 메일을 그냥 넘겨버렸음 - 2026-11-12 김가진 건)"""
    if sender_email and "chubb.com" in sender_email.lower():
        return False
    forwarded = False
    if _is_own_address(sender_email) or (sender_email or "").lower() == (GMAIL_ADDRESS or "").lower():
        # 유진님이 직접 전달(포워딩)한 메일이면, 본문의 원래 보낸 사람(From/보낸 사람/Sender 줄)을
        # 세정 담당자로 봄. 원래 보낸 사람을 못 찾으면 담당자 미확인으로 두고 나중에 물어봄.
        forwarded = True
        m = re.search(r"(?:From|보낸\s*사람|Sender)\s*:\s*([^\n<]*)<?([\w.+-]+@[\w-]+\.[\w.]+)", body or "")
        if m and "chubb.com" not in m.group(2).lower() and not _is_own_address(m.group(2)):
            sender_name = m.group(1).strip().strip('"') or sender_name
            sender_email = m.group(2)
        else:
            sender_email = None
    attachments = _extract_xlsx_attachments(msg) + _extract_non_xlsx_attachments(msg)
    message_id = (msg.get("Message-ID") or msg.get("Message-Id") or "").strip() or None
    # 세정 담당자가 직접 보낸 메일이면 접수 확인 답장을 보냄(유진님이 전달한 메일은 제외, 2026-11-23)
    ack_to = sender_email if (not forwarded and _sj_is_sejung_address(sender_email)) else None
    return await _sj_process_request(bot, attachments, subject, body, sender_email, sender_name,
                                     source="메일", message_id=message_id, ack_to=ack_to)


# ---------- 접수 확인 답장 (2026-11-23 유진님 요청) ----------
# 세정 담당자가 추가가입/해지 신청 메일을 보내면 "잘 받았고, 처리 후 증명서 보내드리겠다"는
# 답장을 같은 메일 대화로 바로 보내고, 보낸 뒤 텔레그램으로 알림. 끄려면 Render 환경변수 SEJUNG_AUTO_ACK=0
SEJUNG_AUTO_ACK = os.getenv("SEJUNG_AUTO_ACK", "1").strip() != "0"


def _sj_is_sejung_address(addr: str | None) -> bool:
    a = (addr or "").lower()
    return bool(a) and ("@sejung" in a or "@ovlr" in a) and not _is_own_address(a)


def _sj_ack_body(items: list, brand: str | None) -> str:
    adds = [i["store_name"] for i in items if i.get("kind") == "추가"]
    dels = [i["store_name"] for i in items if i.get("kind") == "해지"]
    what = "추가가입 및 해지" if adds and dels else "추가가입" if adds else "해지" if dels else "변경"
    cert = ("보험가입증명서와 변경증명서를" if adds and dels else "보험가입증명서를" if adds
            else "변경증명서를" if dels else "증명서를")
    lines = []
    if adds:
        lines.append(f"- 추가가입: {', '.join(adds)}")
    if dels:
        lines.append(f"- 해지: {', '.join(dels)}")
    store_block = ("\n".join(lines) + "\n\n") if lines else ""
    brand_text = f" {brand}" if brand else ""
    return (
        "안녕하십니까~\n"
        f"보내주신 ㈜세정{brand_text} {what} 신청 메일 잘 받았습니다.\n\n"
        + store_block +
        f"보험사에 접수하여 처리되는 대로 {cert} 보내드리겠습니다.\n\n"
        "감사합니다.\n"
        "수고하세요\n\n"
        "--\n" + _CERT_EMAIL_SIGNATURE
    )


async def _sj_send_ack(bot, to_addr: str, to_name: str | None, orig_subject: str,
                       message_id: str | None, items: list, brand: str | None) -> None:
    """접수 확인 답장을 보내고 텔레그램으로 결과를 알림. 실패해도 신청서 처리는 계속됨."""
    subj = (orig_subject or "").strip()
    subject = subj if re.match(r"(?i)^\s*re\s*:", subj) else f"RE: {subj or '㈜세정 변경신청'}"
    headers = {"In-Reply-To": message_id, "References": message_id} if message_id else None
    who = f"{to_name}님({to_addr})" if to_name else to_addr
    try:
        await asyncio.to_thread(_send_email, to_addr, subject, _sj_ack_body(items, brand), None, headers)
    except Exception as e:
        logger.exception("세정 접수 확인 메일 발송 중 오류")
        await _send_message_retrying(
            bot, ALLOWED_USER_ID,
            f"⚠️ {who}께 접수 확인 메일을 보내지 못했어요: {type(e).__name__} ({_sj_err_where()}): {str(e)[:120]}",
        )
        return
    stores = ", ".join(f"{i['kind']} {i['store_name']}" for i in items if i.get("kind") in ("추가", "해지"))
    await _send_message_retrying(
        bot, ALLOWED_USER_ID,
        f"📧 {who}께 접수 확인 메일을 보냈어요.\n- 제목: {subject}" + (f"\n- 매장: {stores}" if stores else ""),
    )


SEJUNG_LAST_CLOSE_FILE = os.path.join(PERSIST_DIR, "sejung_last_close.json")
SEJUNG_AWAIT_RECEIVER_FILE = os.path.join(PERSIST_DIR, "sejung_await_receiver.json")


def _sj_set_receiver_in_master(brand: str, items: list, name: str) -> str:
    """통합관리리스트에서 해당 매장 줄(가장 아래 줄)의 접수자 칸에 이름을 적음."""
    service = _get_gdrive_service()
    root = _sj_find_root(service) if service else None
    master = _sj_find_master(service, root) if root else None
    if master is None:
        return "❌ 통합관리리스트를 못 찾아 접수자를 적지 못했어요."
    try:
        wb = openpyxl.load_workbook(io.BytesIO(_gdrive_download_file(service, master["id"])), data_only=False)
        done = []
        for it in items:
            want = "추가" if it["kind"] == "추가" else "해지"
            ws = next((wb[n] for n in wb.sheetnames
                       if brand and brand in unicodedata.normalize("NFC", n) and want in n), None)
            if ws is None:
                continue
            hdr = {}
            for c in range(1, ws.max_column + 1):
                lab = ws.cell(row=3, column=c).value or ws.cell(row=2, column=c).value
                if lab not in (None, ""):
                    hdr.setdefault(_norm_header(lab), c)
            ncol, rcol = hdr.get("매장명"), hdr.get("접수자")
            if not (ncol and rcol):
                continue
            rows = [r for r in range(4, ws.max_row + 1) if _sj_norm(ws.cell(row=r, column=ncol).value) == _sj_norm(it["store_name"])]
            if rows and not isinstance(ws.cell(row=rows[-1], column=rcol), MergedCell):
                ws.cell(row=rows[-1], column=rcol).value = name
                done.append(f"{want} {it['store_name']}")
        out = io.BytesIO()
        wb.save(out)
        if not _save_brand_master_to_drive(master["id"], out.getvalue()):
            return "❌ 통합관리리스트 저장에 실패했어요."
        return f"✅ 통합관리리스트 접수자 칸에 {name} 이름을 적었어요({', '.join(done) or '해당 줄 없음'})."
    except Exception as e:
        logger.exception("접수자 기재 중 오류")
        return f"❌ 접수자 기재 실패: {type(e).__name__}"


async def _sj_handle_receiver_reply(update: Update, user_text: str) -> bool:
    """텔레그램으로 올린 신청서의 담당자 이름을 물어본 상태에서 받은 답을 처리."""
    waiting = _sj_load(SEJUNG_AWAIT_RECEIVER_FILE, [])
    text = (user_text or "").strip()
    if not waiting or not text or text.startswith("/") or len(text) > 30:
        return False
    reply_to = getattr(update.message, "reply_to_message", None)
    reply_text = getattr(reply_to, "text", "") or "" if reply_to is not None else ""
    if reply_to is not None and "접수자" not in reply_text:
        return False  # 다른 질문(증명서 보낼 담당자 등)에 대한 답이면 그쪽 처리로 넘김
    if reply_to is None and _sj_load(SEJUNG_AWAIT_FILE, []):
        # 두 종류 질문이 동시에 대기 중인데 답장 지정이 없으면, 문장 속 매장명으로 판단
        n = _sj_norm(text)
        if not any(_sj_norm(it["store_name"]) in n for w in waiting for it in w.get("items", [])):
            return False
    idx = _sj_pick_waiting(update, waiting, "items")
    if _SJ_CANCEL_RE.search(text):
        # '직접 보냈어' 같은 답은 이름이 아님 - 접수자 질문은 접고, 증명서 발송 대기 취소는
        # 아래 증명서 담당자 처리에서 이어서 하도록 넘김(2026-11-16: '장유롯데'를 이름으로 적던 문제)
        dropped = waiting.pop(idx)
        _sj_save(SEJUNG_AWAIT_RECEIVER_FILE, waiting)
        if _sj_load(SEJUNG_AWAIT_FILE, []):
            return False
        names = ", ".join(i["store_name"] for i in dropped.get("items", []))
        await update.message.reply_text(f"🗑 '{names}' 접수자 질문은 넘어갈게요(통합관리리스트는 그대로 둬요).")
        return True
    entry = waiting.pop(idx)
    _sj_save(SEJUNG_AWAIT_RECEIVER_FILE, waiting)
    name = _sj_name_from_reply(text)
    email_addr = _load_contacts().get(_norm_contact_name(name))
    if email_addr and _is_own_address(email_addr):
        email_addr = None
    # 나중에 처브 증명서가 오면 이 담당자에게 보내도록 기억해둔 신청 기록도 갱신
    requests = _sj_load(SEJUNG_REQUESTS_FILE, {})
    for it in entry["items"]:
        key = f"{entry.get('brand') or ''}|{_sj_norm(it['store_name'])}"
        if key in requests:
            requests[key]["sender_name"] = name
            if email_addr:
                requests[key]["sender_email"] = email_addr
    _sj_save(SEJUNG_REQUESTS_FILE, requests)
    msg = await asyncio.to_thread(_sj_set_receiver_in_master, entry.get("brand"), entry["items"], name)
    tail = (f"증명서가 오면 {name}({email_addr})님께 보낼게요." if email_addr else
            f"{name}님 이메일이 등록돼 있지 않아요. /setcontact {name} 이메일 로 등록해두시면 증명서가 왔을 때 바로 보낼 수 있어요.")
    await update.message.reply_text(f"{msg}\n{tail}")
    if waiting:
        await update.message.reply_text("다음 신청서의 담당자도 알려주세요.")
    return True
_SJ_DOC_EXTS = (".pdf", ".jpg", ".jpeg", ".png")
_SJ_CLOSURE_DOC_TYPES = ("사업자등록증", "통장사본")


def _sj_upload_closure_docs(brand: str, named_docs: list) -> str:
    """[(저장할 이름, bytes)]를 정산관리/브랜드/해지변경신청서/YY년 M월 폴더에 저장."""
    service = _get_gdrive_service()
    if service is None:
        return "❌ 구글 드라이브에 연결되지 않아 해지 서류를 저장하지 못했어요."
    root = _sj_find_root(service)
    if root is None:
        return "❌ 드라이브에서 '정산관리' 폴더를 못 찾아 해지 서류를 저장하지 못했어요."
    try:
        month_id = _sj_month_folder(service, root, brand, "해지변경신청서")
        for name, data in named_docs:
            low = name.lower()
            mime = ("application/pdf" if low.endswith(".pdf") else
                    "image/png" if low.endswith(".png") else "image/jpeg")
            service.files().create(
                body={"name": name, "parents": [month_id]},
                media_body=_gdrive_media(data, mimetype=mime), fields="id",
            ).execute()
        today = dt.datetime.now(KST).date()
        return (f"💾 해지 서류 {len(named_docs)}개({', '.join(n for n, _ in named_docs)})를 "
                f"정산관리/{brand}/해지변경신청서/{today:%y}년 {today.month}월 폴더에 저장했어요.")
    except Exception as e:
        logger.exception("해지 서류 드라이브 저장 중 오류")
        return f"❌ 해지 서류 드라이브 저장 실패: {type(e).__name__}"


async def _sj_save_closure_docs(brand: str | None, stores: list, docs: list) -> str | None:
    """해지 관련 서류(사업자등록증/통장사본)를 '통장사본_OO점.pdf'처럼
    이름 붙여 해지변경신청서 폴더에 저장(2026-11-13 유진님 요청). 서류가 없으면 None."""
    if not docs or not stores or not brand:
        return None
    named, skipped = [], []
    used = set()
    for fn, data in docs:
        if not fn.lower().endswith(_SJ_DOC_EXTS):
            continue
        category = await _classify_closure_doc_deep(fn, data)
        if category not in _SJ_CLOSURE_DOC_TYPES:
            skipped.append(fn)
            continue
        norm_fn = _sj_norm(fn)
        store = next((st for st in stores if _sj_norm(st) and _sj_norm(st) in norm_fn), stores[0])
        ext = os.path.splitext(fn)[1].lower() or ".pdf"
        name = f"{category}_{store}{ext}"
        n = 2
        while name in used:
            name = f"{category}_{store}_{n}{ext}"
            n += 1
        used.add(name)
        named.append((name, data))
    return await asyncio.to_thread(_sj_upload_closure_docs, brand, named) if named else None


def _sj_request_file_kind(filename: str, data: bytes) -> str:
    """신청서 파일 하나가 추가인지 해지인지: PDF는 '소재지 제외/Decreased'면 해지, 엑셀은 매장명이
    적힌 시트로 판단. 애매하면 파일명('해지')을 보고, 그래도 모르면 추가."""
    if filename.lower().endswith(".xlsx"):
        items = _sj_parse_request(data, filename) or []
        if items and all(i["kind"] == "해지" for i in items):
            return "해지"
        if items:
            return "추가"
    else:
        t = (_extract_pdf_text(data) or "").replace(" ", "")
        if "소재지제외" in t or "DecreasedTSI" in t or "ReturnPremium" in t:
            return "해지"
        if "소재지추가" in t or "IncreasedTSI" in t:
            return "추가"
    return "해지" if "해지" in filename and "가입" not in filename else "추가"


# ---------- 직영점 신청서 계좌 확인 (2026-11-24 유진님 요청) ----------
# ㈜세정 직영점 추가가입/해지 신청서에 적힌 계좌가 브랜드별로 정해진 계좌와 같으면 그냥 넘어가고,
# 다르거나 못 찾으면 텔레그램으로 알림. (비교는 숫자만: 하이픈·공백 무시)
_SJ_DIRECT_ACCOUNTS = {
    "올리비아로렌": ("신한은행", "040-015-088964"),
    "오뷔엘알": ("신한은행", "040-015-088964"),
    "웰메이드": ("신한은행", "140-001-277300"),
    "트레몰로": ("우리은행", "043-132564-13-515"),
}
_SJ_ACCOUNT_WORDS_RX = re.compile(r"계좌|은행|입금|예금주|환급|위임")


def _sj_expected_account(brand: str | None, store: str | None):
    t = f"{brand or ''} {store or ''}"
    if "오뷔엘알" in t or "ovlr" in t.lower():
        return _SJ_DIRECT_ACCOUNTS["오뷔엘알"]
    for b, acc in _SJ_DIRECT_ACCOUNTS.items():
        if b in t:
            return acc
    return None


def _sj_xlsx_text(data: bytes) -> str:
    """엑셀 신청서의 모든 칸 글자를 이어 붙임(계좌 찾기용)."""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        parts = []
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                parts.extend(str(v) for v in row if v not in (None, ""))
        return " ".join(parts)
    except Exception:
        logger.exception("엑셀 신청서 글자 읽기 실패(계좌 확인)")
        return ""


def _sj_find_accounts(text: str) -> list:
    """'계좌/은행/입금/예금주/환급' 근처에 적힌 계좌번호 모양(숫자-숫자, 10~14자리)을 찾음."""
    found = []
    for m in _SJ_ACCOUNT_WORDS_RX.finditer(text or ""):
        window = text[max(0, m.start() - 40): m.end() + 100]
        for a in re.findall(r"\d[\d\-]{7,22}\d", window):
            digits = re.sub(r"\D", "", a)
            if not (10 <= len(digits) <= 14) or digits.startswith(("010", "070")):
                continue
            if a not in found:
                found.append(a)
    return found


def _sj_wrong_accounts(text: str, correct_acc: str) -> list:
    """'계좌' 글자 바로 뒤(60자 안)에 적힌 계좌번호 중 정해진 계좌가 아닌 것.
    사업자등록번호(000-00-00000)·전화번호 모양은 빼서, 계좌가 아닌 번호를 고치는 일이 없게 함."""
    good = re.sub(r"\D", "", correct_acc)
    out = []
    for m in re.finditer(r"계좌", text or ""):
        window = text[m.end(): m.end() + 60]
        for a in re.findall(r"\d[\d\- ]{8,22}\d", window):
            a = a.strip()
            digits = re.sub(r"\D", "", a)
            if not (10 <= len(digits) <= 14) or digits == good:
                continue
            if (re.fullmatch(r"\d{3}-\d{2}-\d{5}", a)  # 사업자등록번호
                    or re.fullmatch(r"0(?:1[016789]|2|[3-6][1-5]|70|303|505)[- ]\d{3,4}[- ]\d{4}", a)  # 전화
                    or re.fullmatch(r"01[016789]\d{7,8}", digits)):  # 휴대폰(하이픈 없음)
                continue
            if a not in out:
                out.append(a)
    return out


def _sj_check_direct_accounts(items: list, texts: list) -> list:
    """직영점 신청 건의 계좌를 정해진 계좌와 비교(2026-11-24).
    반환: [{"item", "bank", "acc", "status": "wrong"|"missing", "wrong": [틀린 계좌], "found": [찾은 번호]}]
    계좌가 맞는 건은 목록에 넣지 않음(그냥 넘어감)."""
    results = []
    joined = "\n".join(t for t in texts if t)
    compact = re.sub(r"[\s\-]", "", joined)
    for it in items:
        store = it.get("store_name") or ""
        if "직영" not in store or it.get("kind") not in ("추가", "해지"):
            continue
        exp = _sj_expected_account(it.get("brand"), store)
        if not exp:
            continue
        bank, acc = exp
        if re.sub(r"\D", "", acc) in compact:
            continue  # 정해진 계좌가 신청서에 있음 -> 통과
        extra = str((it.get("fields") or {}).get("비고") or "")
        wrong = _sj_wrong_accounts(joined + " " + extra, acc)
        found = wrong or _sj_find_accounts(joined + " " + extra)
        results.append({"item": it, "bank": bank, "acc": acc,
                        "status": "wrong" if found else "missing", "wrong": wrong, "found": found})
    return results


# ----- 틀린 계좌를 신청서 파일 안에서 정해진 계좌로 고치기 (2026-11-24 유진님 요청) -----
_SJ_FIX_FONT_NAME = None


def _sj_fix_font() -> str:
    """신청서 글자와 비슷한 바탕 글꼴(batang.ttc)을 reportlab에 등록."""
    global _SJ_FIX_FONT_NAME
    if _SJ_FIX_FONT_NAME is None:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        pdfmetrics.registerFont(TTFont("SJBatang", CERT_FONT_PATH, subfontIndex=0))
        _SJ_FIX_FONT_NAME = "SJBatang"
    return _SJ_FIX_FONT_NAME


def _sj_acc_rx(acc: str):
    """계좌 숫자 사이에 하이픈·공백이 있어도 찾는 정규식."""
    d = re.sub(r"\D", "", acc)
    return re.compile(r"[\s\-]*".join(re.escape(c) for c in d))


def _sj_fix_account_text(text: str, wrong: list, bank: str, acc: str) -> str:
    new = text
    for w in wrong:
        new = _sj_acc_rx(w).sub(acc, new)
    if new != text and bank not in new:
        new = re.sub(r"[가-힣A-Za-z]*은행", bank, new, count=1)
    return new


def _sj_fix_pdf_account(pdf_bytes: bytes, wrong: list, bank: str, acc: str):
    """PDF 신청서에서 틀린 계좌가 적힌 줄을 흰색으로 덮고 정해진 은행·계좌로 다시 써넣음.
    도장(날인)·다른 글자는 건드리지 않음. 반환 (고친 PDF 또는 None, [(원래 글자, 고친 글자)])."""
    from reportlab.pdfbase import pdfmetrics
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()
    changed = []
    font = _sj_fix_font()
    for page in reader.pages:
        frags = []

        def cb(text, cm, tm, fd, fs):
            if text and text.strip():
                frags.append((text, list(cm), list(tm), fs))

        page.extract_text(visitor_text=cb)
        hits = []
        for text, cm, tm, fs in frags:
            new = _sj_fix_account_text(text, wrong, bank, acc)
            if new == text:
                continue
            x = cm[0] * tm[4] + cm[2] * tm[5] + cm[4]
            y = cm[1] * tm[4] + cm[3] * tm[5] + cm[5]
            size = (fs or 10) * (abs(tm[0]) or 1) * (abs(cm[0]) or 1)
            hits.append((x, y, size, text, new))
        if hits:
            w, h = float(page.mediabox.width), float(page.mediabox.height)
            buf = io.BytesIO()
            c = pdf_canvas.Canvas(buf, pagesize=(w, h))
            for x, y, size, old, new in hits:
                wid = max(pdfmetrics.stringWidth(old, font, size), pdfmetrics.stringWidth(new, font, size)) * 1.1 + 4
                c.setFillColorRGB(1, 1, 1)
                c.rect(x - 1, y - size * 0.35, wid, size * 1.4, stroke=0, fill=1)
                c.setFillColorRGB(0, 0, 0)
                c.setFont(font, size)
                c.drawString(x, y, new)
                changed.append((old.strip(), new.strip()))
            c.save()
            buf.seek(0)
            page.merge_page(PdfReader(buf).pages[0])
        writer.add_page(page)
    if not changed:
        return None, []
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue(), changed


def _sj_fix_xlsx_account(data: bytes, wrong: list, bank: str, acc: str):
    """엑셀 신청서 안의 틀린 계좌를 정해진 은행·계좌로 바꿈. 파일 속 글자(XML)만 바꿔서
    도장 그림·서식은 그대로 둠(openpyxl로 다시 저장하면 그림이 빠질 수 있어서)."""
    import html as _html
    changed = []
    wrong_digits = {re.sub(r"\D", "", w) for w in wrong}
    zin = zipfile.ZipFile(io.BytesIO(data))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            b = zin.read(info.filename)
            name = info.filename
            is_sheet = name.startswith("xl/worksheets/") and name.endswith(".xml")
            if name == "xl/sharedStrings.xml" or is_sheet:
                s = b.decode("utf-8")

                def fix_t(m):
                    txt = _html.unescape(m.group(2))
                    new = _sj_fix_account_text(txt, wrong, bank, acc)
                    if new == txt:
                        return m.group(0)
                    changed.append((txt.strip(), new.strip()))
                    return m.group(1) + _html.escape(new, quote=False) + m.group(3)

                s = re.sub(r"(<t(?:\s[^>]*)?>)([^<]*)(</t>)", fix_t, s)
                if is_sheet:
                    def fix_num(m):
                        attrs, v = m.group(1), m.group(2)
                        # 숫자 칸(t 없음 또는 t="n")에 계좌가 숫자로 들어 있는 경우 -> 글자 칸으로 바꿔 넣음
                        if v in wrong_digits and (' t="' not in attrs or ' t="n"' in attrs):
                            changed.append((v, acc))
                            attrs = attrs.replace(' t="n"', "")
                            return f'<c{attrs} t="inlineStr"><is><t>{acc}</t></is></c>'
                        return m.group(0)
                    s = re.sub(r"<c([^>]*)><v>(\d+)</v></c>", fix_num, s)
                b = s.encode("utf-8")
            zout.writestr(info, b)
    if not changed:
        return None, []
    return out.getvalue(), changed


def _sj_fix_request_accounts(request_files: list, fixes: list):
    """틀린 계좌가 있는 신청서 파일들을 고침. fixes: _sj_check_direct_accounts의 'wrong' 결과들.
    반환 (바뀐 파일 목록 [(이름, 내용)], {원래이름: 새이름}, 알림 문구 목록)"""
    new_files, renamed, msgs = [], {}, []
    for fn, data in request_files:
        cur, changes = data, []
        for fx in fixes:
            if not fx["wrong"]:
                continue
            if fn.lower().endswith(".pdf"):
                fixed, ch = _sj_fix_pdf_account(cur, fx["wrong"], fx["bank"], fx["acc"])
            elif fn.lower().endswith((".xlsx", ".xlsm")):
                fixed, ch = _sj_fix_xlsx_account(cur, fx["wrong"], fx["bank"], fx["acc"])
            else:
                fixed, ch = None, []
            if fixed:
                cur, changes = fixed, changes + ch
        if changes:
            stem, ext = os.path.splitext(fn)
            new_fn = f"{stem}_계좌수정{ext}"
            renamed[fn] = new_fn
            new_files.append((new_fn, cur))
            msgs.extend(f"  - '{o}' → '{n}'" for o, n in dict.fromkeys(changes))
        else:
            new_files.append((fn, data))
    return new_files, renamed, msgs


def _sj_is_excel_name(fn: str) -> bool:
    return (fn or "").lower().endswith((".xlsx", ".xlsm", ".xls"))


def _sj_excel_to_pdf(service, fn: str, data: bytes) -> bytes | None:
    """엑셀 신청서를 PDF로 변환(2026-11-24). Render에는 엑셀/리브레오피스가 없어서
    구글 드라이브에 임시 구글시트로 올려 PDF로 내려받은 뒤 임시 파일은 바로 지움.
    A4 세로, 가로폭을 한 페이지에 맞춤, 눈금선 없음."""
    src_mime = "application/vnd.ms-excel" if fn.lower().endswith(".xls") else _XLSX_MIMETYPE
    tmp = service.files().create(
        body={"name": f"_PDF변환임시_{fn}", "mimeType": "application/vnd.google-apps.spreadsheet"},
        media_body=_gdrive_media(data, mimetype=src_mime), fields="id",
    ).execute()
    tmp_id = tmp["id"]
    try:
        pdf = None
        try:
            url = (f"https://docs.google.com/spreadsheets/d/{tmp_id}/export?format=pdf&size=A4"
                   "&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false"
                   "&pagenum=UNDEFINED&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5")
            resp, content = service._http.request(url)
            if getattr(resp, "status", 0) == 200 and content[:4] == b"%PDF":
                pdf = content
        except Exception:
            logger.warning("시트 PDF 내보내기(인쇄설정) 실패, 기본 내보내기로 재시도: %s", fn)
        if pdf is None:
            content = service.files().export(fileId=tmp_id, mimeType="application/pdf").execute()
            if content and content[:4] == b"%PDF":
                pdf = content
        return pdf
    finally:
        try:
            service.files().delete(fileId=tmp_id).execute()
        except Exception:
            logger.warning("PDF 변환 임시 시트 삭제 실패: %s", tmp_id)


def _sj_save_request_files(brand: str | None, request_files: list) -> str:
    """신청서 파일(엑셀/PDF)을 드라이브 정산관리/브랜드/월 폴더에 원래 파일명으로 저장(2026-11-13 요청)."""
    if not request_files:
        return "ℹ️ 저장할 신청서 파일이 없어요."
    if not brand:
        return "⚠️ 브랜드를 몰라서 신청서를 드라이브에 저장하지 못했어요."
    service = _get_gdrive_service()
    if service is None:
        return "❌ 구글 드라이브에 연결되지 않아 신청서를 저장하지 못했어요."
    root = _sj_find_root(service)
    if root is None:
        return "❌ 드라이브에서 '정산관리' 폴더를 못 찾아 신청서를 저장하지 못했어요."
    try:
        # 브랜드/추가변경신청서, 브랜드/해지변경신청서 로 나눠 월별 저장(2026-11-13 유진님 요청)
        today = dt.datetime.now(KST).date()
        saved, same, converted, convert_failed = [], [], [], []
        import hashlib
        for fn, data in request_files:
            kind_folder = "해지변경신청서" if _sj_request_file_kind(fn, data) == "해지" else "추가변경신청서"
            month_id = _sj_month_folder(service, root, brand, kind_folder)
            # 같은 내용의 파일이 이미 폴더에 있으면 다시 올리지 않음(같은 신청서 재발송, 2026-11-23)
            # 엑셀을 PDF로 바꿔 저장한 파일은 PDF 내용이 매번 달라지므로 원본 엑셀의 md5(src_md5)로 비교
            md5 = hashlib.md5(data).hexdigest()
            existing = service.files().list(
                q=f"'{month_id}' in parents and trashed=false",
                fields="files(id,name,md5Checksum,appProperties)", pageSize=500,
            ).execute().get("files", [])
            if any(f.get("md5Checksum") == md5 or (f.get("appProperties") or {}).get("src_md5") == md5
                   for f in existing):
                same.append(fn)
                continue
            save_name, save_data, mime, props = fn, data, "application/pdf", None
            if _sj_is_excel_name(fn):
                # 엑셀 신청서는 PDF로 바꿔서 저장(2026-11-24 유진님 요청). 처브로 보내는 첨부는 원본 엑셀 그대로.
                pdf = None
                try:
                    pdf = _sj_excel_to_pdf(service, fn, data)
                except Exception:
                    logger.exception("신청서 엑셀 -> PDF 변환 중 오류: %s", fn)
                if pdf:
                    save_name = re.sub(r"\.(xlsx|xlsm|xls)$", "", fn, flags=re.I) + ".pdf"
                    save_data, props = pdf, {"src_md5": md5}
                    converted.append(save_name)
                else:
                    mime = _XLSX_MIMETYPE if fn.lower().endswith((".xlsx", ".xlsm")) else "application/vnd.ms-excel"
                    convert_failed.append(fn)
            body_meta = {"name": save_name, "parents": [month_id]}
            if props:
                body_meta["appProperties"] = props
            service.files().create(
                body=body_meta, media_body=_gdrive_media(save_data, mimetype=mime), fields="id",
            ).execute()
            saved.append(f"{kind_folder}/{today:%y}년 {today.month}월")
        parts = []
        if saved:
            parts.append(f"💾 신청서 {len(saved)}개를 드라이브 정산관리/{brand}/({', '.join(sorted(set(saved)))}) 폴더에 저장했어요.")
        if converted:
            parts.append(f"📄 엑셀 신청서는 PDF로 바꿔 저장했어요({', '.join(converted)}).")
        if convert_failed:
            parts.append(f"⚠️ PDF 변환에 실패해 엑셀 원본으로 저장했어요({', '.join(convert_failed)}).")
        if same:
            parts.append(f"💾 같은 내용의 신청서가 드라이브에 이미 있어 다시 저장하지 않았어요({', '.join(same)}).")
        return "\n".join(parts)
    except Exception as e:
        logger.exception("신청서 드라이브 저장 중 오류")
        return f"❌ 신청서 드라이브 저장 실패: {type(e).__name__}"


async def _sj_process_request(bot, attachments: list, subject: str, body: str,
                              sender_email: str | None, sender_name: str | None, source: str,
                              message_id: str | None = None, ack_to: str | None = None) -> bool:
    """변경신청서 처리 공통 흐름(메일로 온 경우, 텔레그램에 직접 올린 경우 둘 다):
    매장/신청내용 읽기 -> 보낸 사람 기억 -> 통합관리리스트 기재 -> 신청서를 드라이브에 저장 ->
    처브 4명에게 보낼지 확인 버튼. 신청서가 아니면 False.
    2026-11-23 중복 방지: 같은 메일/첨부는 건너뛰고, 이미 기재된 매장은 후속 작업도 하지 않음."""
    if not attachments:
        return False
    xlsx = [(fn, d) for fn, d in attachments if fn.lower().endswith(".xlsx")]
    others = [(fn, d) for fn, d in attachments if not fn.lower().endswith(".xlsx")]

    items, request_files = [], []
    for fn, data in xlsx:
        parsed = _sj_parse_request(data, fn)
        if parsed is not None:
            request_files.append((fn, data))
        if parsed:
            items.extend(parsed)
    pdf_texts, has_pdf_request = [], False
    pdf_request_texts, pdf_rule_items = [], []
    for fn, data in others:
        if fn.lower().endswith(".pdf"):
            t = _extract_pdf_text(data)
            if _sj_is_chubb_cert_text(t):
                return False  # 처브 증명서는 증명서 처리로 넘김
            if _sj_pdf_is_request(t):
                has_pdf_request = True
                request_files.append((fn, data))
                pdf_request_texts.append(t)
                pdf_rule_items.extend(_sj_parse_request_pdf(data, fn))
            pdf_texts.append(t)
    text_all = f"{subject} {body}"
    # 제목/본문만으로 신청 메일로 볼 때는 반드시 ㈜세정 건이라는 표시가 있어야 함(세정 메일 주소,
    # 브랜드명/증권번호, '세정' 글자). 예전엔 '요청'+'가입' 같은 단어만 보고 다른 거래처 메일
    # (예: 파블로항공 드론보험 견적요청)까지 세정 신청으로 잘못 처리했음(2026-11-16).
    is_sejung = (
        any(d in (sender_email or "").lower() for d in ("sejung", "ovlr"))
        or "세정" in text_all
        or _sj_brand_from_text(text_all) is not None
    )
    looks_like_request = is_sejung and (
        any(w in text_all for w in _SJ_REQUEST_WORDS) and any(k in text_all for k in _SJ_REQUEST_KINDS)
    )
    if not items and not has_pdf_request and not looks_like_request:
        return False

    # 1단계(2026-11-23): 같은 메일(Message-ID)이나 같은 첨부(SHA-256)가 다시 오면 처리하지 않음.
    # 이력은 드라이브 '가입증명서및정산' 폴더의 세정_처리이력.json(재배포해도 남게)
    hash_files = request_files or [(fn, d) for fn, d in attachments if fn.lower().endswith((".xlsx", ".pdf"))]
    file_hashes = [(_sj_sha256(d), fn) for fn, d in hash_files]
    prev = await asyncio.to_thread(_sj_history_check, message_id, [h for h, _ in file_hashes])
    if prev:
        await _send_message_retrying(
            bot, ALLOWED_USER_ID,
            f"⏭ 이미 처리한 정산파일이라 건너뜀 (최초 처리: {prev.get('at') or '?'}/"
            f"{prev.get('sender') or '보낸사람 미확인'})\n- {subject}",
        )
        return True

    brand = (next((i.get("brand") for i in items if i.get("brand")), None)
             or _sj_brand_from_text(subject, body, *pdf_texts, *[fn for fn, _ in attachments]))
    if not items:
        found = await asyncio.to_thread(_sj_extract_request_from_text, subject, body, pdf_texts)
        if pdf_rule_items:
            found = _sj_apply_pdf_rules(found, pdf_rule_items)
        items = [{"kind": f["kind"], "brand": brand, "store_name": f["store_name"], "fields": f.get("fields") or {}}
                 for f in found]
    elif has_pdf_request:
        # 엑셀 신청서와 PDF 신청서가 같이 온 메일: PDF에만 있는 매장도 빠뜨리지 않고 기재
        found = await asyncio.to_thread(_sj_extract_request_from_text, subject, "", pdf_request_texts)
        if pdf_rule_items:
            found = _sj_apply_pdf_rules(found, pdf_rule_items)
        for f in found:
            if any(i["kind"] == f["kind"] and (_sj_norm(i["store_name"]) == _sj_norm(f["store_name"])
                                               or _sj_store_similar(i["store_name"], f["store_name"]))
                   for i in items):
                continue  # 엑셀에 이미 있는 매장은 엑셀 값 그대로
            items.append({"kind": f["kind"], "brand": brand, "store_name": f["store_name"],
                          "fields": f.get("fields") or {}})
    for item in items:
        item["brand"] = item.get("brand") or brand
    # 매장명을 제대로 못 읽은 건('~ 일간' 같은 양식 글자 등)은 리스트에 적지 않고 알림(2026-11-24)
    bad_names = [i for i in items if not _sj_valid_store_name(i.get("store_name"))]
    items = [i for i in items if _sj_valid_store_name(i.get("store_name"))]
    # 매장코드가 비어 있으면 메일 제목 '(OLM7201)'이나 첨부 파일명에서 채움(2026-11-26)
    # 코드가 딱 1개이고 코드 없는 매장도 1곳일 때만(여러 매장이면 엉뚱한 코드가 붙을 수 있어서)
    no_code = [i for i in items if i.get("kind") in ("추가", "해지")
               and not str((i.get("fields") or {}).get("매장코드") or "").strip()]
    if len(no_code) == 1:
        codes = set(re.findall(r"\((" + _SJ_CODE_RX + r")\)", subject or ""))
        codes |= {c for c in (_sj_code_from_filename(fn) for fn, _ in request_files) if c}
        if len(codes) == 1:
            no_code[0].setdefault("fields", {})["매장코드"] = codes.pop()
    # 2단계(2026-11-23): 한 파일 안에 같은 매장이 두 번 적혀 온 경우 한 번만 처리
    items, dropped = _sj_dedupe_items(items)
    # 통합관리리스트 기재(같은 건이 있으면 ①건너뜀 ②확인 버튼 ③새로 기재 - item['_dup']에 결과)
    list_msgs = await asyncio.to_thread(_sj_write_requests_to_master, items, sender_name)
    for d in dropped:
        list_msgs.append(f"⏭ {d['kind']} {d['store_name']}: 같은 파일 안에 두 번 적혀 있어 한 번만 처리했어요.")
    if bad_names:
        list_msgs.append(f"⚠️ 매장명을 읽지 못한 신청 {len(bad_names)}건은 통합관리리스트에 적지 않았어요. "
                         "신청서를 직접 확인해 주세요.")
    new_n = sum(1 for i in items if i.get("_dup") == "new")
    same_n = sum(1 for i in items if i.get("_dup") == "same") + len(dropped)
    confirm_items = [i for i in items if i.get("_dup") == "confirm"]
    summary = f"📊 신규 기재 {new_n}건 / 중복 건너뜀 {same_n}건 / 확인 필요 {len(confirm_items)}건"
    who = f"{sender_name or ''} {sender_email or ''}".strip() or "담당자 미확인"
    sender_label = sender_name or sender_email or source
    # 3단계: 이미 기재돼 건너뛴 매장(①)은 보낸 사람 기억·드라이브 저장·처브 발송·접수자 질문을 다시 하지 않음
    live = [i for i in items if i.get("_dup") != "same"]
    for item in live:
        _sj_remember_request(item, sender_email, sender_name, subject)
    stores_all = [f"{i['kind']} {i['store_name']}" for i in items]

    if items and not live:
        await _send_message_retrying(
            bot, ALLOWED_USER_ID,
            f"📨 ㈜세정 변경신청서({source}): {brand or '브랜드 미확인'}, 보낸 사람: {who}\n"
            + "\n".join(list_msgs)
            + "\n모두 이미 기재된 매장이라 드라이브 저장·처브 발송은 하지 않았어요.\n\n" + summary,
        )
        await asyncio.to_thread(_sj_history_record, message_id, file_hashes, sender_label, stores_all)
        return True

    # 직영점 신청서 계좌 확인(2026-11-24): 맞으면 그냥 넘어가고, 틀리면 신청서 파일 안의 계좌를
    # 정해진 계좌로 고친 '_계좌수정' 파일로 저장·처브 발송하고 알림. 계좌를 못 찾으면 알림만.
    orig_request_files = list(request_files)
    acc_alerts, acc_fixed_note = [], ""
    try:
        acc_texts = list(pdf_request_texts) + [
            _sj_xlsx_text(d) for fn, d in request_files if fn.lower().endswith((".xlsx", ".xlsm"))
        ] + [body or ""]
        acc_results = _sj_check_direct_accounts(live, acc_texts)
        fixes = [r for r in acc_results if r["status"] == "wrong" and r["wrong"]]
        fix_msgs, renamed = [], {}
        if fixes and request_files:
            request_files, renamed, fix_msgs = await asyncio.to_thread(_sj_fix_request_accounts, request_files, fixes)
            if renamed:
                fixed_map = dict(request_files)
                attachments = [(renamed.get(fn, fn), fixed_map.get(renamed.get(fn, fn), d)) for fn, d in attachments]
        for r in acc_results:
            it = r["item"]
            head = f"{it.get('brand') or ''} {it['kind']} {it['store_name']}"
            if r["status"] == "missing":
                acc_alerts.append(f"⚠️ 직영점 계좌 확인 필요: {head}\n  - 신청서 계좌: 신청서에서 계좌를 찾지 못함\n"
                                  f"  - 정해진 계좌: {r['bank']} {r['acc']}")
            elif renamed:
                acc_alerts.append(f"🔧 직영점 계좌가 달라서 신청서를 고쳤어요: {head}\n"
                                  f"  - 정해진 계좌: {r['bank']} {r['acc']}\n" + "\n".join(fix_msgs) +
                                  f"\n  - 고친 파일({', '.join(renamed.values())})로 드라이브 저장·처브 발송해요.")
                acc_fixed_note = (f"\n\n※ {it['store_name']} 환급계좌를 ㈜세정 지정 계좌({r['bank']} {r['acc']})로 "
                                  "정정한 신청서입니다.")
            else:
                acc_alerts.append(f"⚠️ 직영점 계좌가 정해진 계좌와 달라요(신청서 파일을 자동으로 고치지 못함): {head}\n"
                                  f"  - 신청서 계좌: {', '.join(r['found'])}\n"
                                  f"  - 정해진 계좌: {r['bank']} {r['acc']}\n  - 직접 확인해 주세요.")
    except Exception as e:
        logger.exception("직영점 계좌 확인/수정 중 오류")
        request_files = orig_request_files
        acc_alerts = [f"⚠️ 직영점 계좌 확인 중 오류: {type(e).__name__} ({_sj_err_where()}) - 계좌는 직접 확인해 주세요."]

    save_msg = await asyncio.to_thread(_sj_save_request_files, brand, request_files)
    closing_stores = [i["store_name"] for i in live if i.get("kind") == "해지"]
    if closing_stores and brand:
        _sj_save(SEJUNG_LAST_CLOSE_FILE, {"brand": brand, "stores": closing_stores,
                                          "at": dt.datetime.now(KST).isoformat()})
        extra_docs = [(fn, d) for fn, d in others if (fn, d) not in orig_request_files]
        doc_msg = await _sj_save_closure_docs(brand, closing_stores, extra_docs)
        if doc_msg:
            save_msg += "\n" + doc_msg

    desc = ", ".join(f"{i['kind']} {i['store_name']}" for i in live) or "매장 미확인"
    skipped_note = ""
    if same_n:
        skipped = [f"{i['kind']} {i['store_name']}" for i in items if i.get("_dup") == "same"]
        skipped_note = (f"\n(이미 기재돼 건너뛴 매장{': ' + ', '.join(skipped) if skipped else ''}은 메일 본문에서 뺐어요. "
                        "첨부 파일은 원본 그대로예요.)")
    cert_id = _queue_pending_certs(
        [{"store_name": desc, "out_name": fn, "pdf_bytes": data} for fn, data in attachments],
        target_email=", ".join(CHUBB_RECIPIENTS),
        subject=_chubb_request_subject(live, brand, desc),
        body=_chubb_request_body(live, brand),  # 계좌를 고쳐도 본문은 평소 문구 그대로(2026-11-24 요청)
        label="변경신청서",
        sent_keys=[_sj_sent_key("처브", i.get("brand") or brand, i["kind"], i["store_name"])
                   for i in live if i.get("kind") in ("추가", "해지")],
    )
    files = ", ".join(fn for fn, _ in attachments)
    if not sender_email and not sender_name and live:
        # 텔레그램으로 직접 올린 신청서는 담당자(접수자)를 모르니 유진님께 물어봄(2026-11-16 요청)
        waiting = _sj_load(SEJUNG_AWAIT_RECEIVER_FILE, [])
        waiting.append({"brand": brand, "items": [{"kind": i["kind"], "store_name": i["store_name"]} for i in live]})
        _sj_save(SEJUNG_AWAIT_RECEIVER_FILE, waiting)
    await _send_message_retrying(
        bot, ALLOWED_USER_ID,
        f"📨 ㈜세정 변경신청서({source}): {brand or '브랜드 미확인'}, {desc}, 보낸 사람: {who}\n"
        + "\n".join(list_msgs) + "\n" + save_msg + skipped_note +
        f"\n\n{summary}"
        f"\n\n첨부 {len(attachments)}개({files})를 그대로 붙여서 처브 담당자 {len(CHUBB_RECIPIENTS)}명에게 보낼까요?",
        reply_markup=_queue_confirm_keyboard(cert_id),
    )
    # ② 일자·금액이 다른 매장은 하나씩 [덮어쓰기][새로 추가][무시] 버튼으로 확인
    for i in confirm_items:
        c = i["_confirm"]
        await _send_message_retrying(
            bot, ALLOWED_USER_ID,
            f"❓ {i['brand']} {i['kind']} {i['store_name']}: 이미 적힌 줄({c['sheet']} {c['row']}행)과 다른 점이 있어요.\n"
            + "\n".join(f"  - {d}" for d in c["diffs"])
            + "\n정정이면 [덮어쓰기], 다른 건이면 [새로 추가], 필요 없으면 [무시]를 눌러주세요.",
            reply_markup=_sj_dup_keyboard(c["id"]),
        )
    if not sender_email and not sender_name and live:
        await _send_message_retrying(
            bot, ALLOWED_USER_ID,
            f"❓ '{desc}' 신청서의 ㈜세정 담당자(접수자)가 누구예요? 이름을 보내주세요(예: 김가진).",
        )
    # 직영점 계좌 확인 결과 알림(맞으면 알림 없음, 2026-11-24)
    for alert in acc_alerts:
        await _send_message_retrying(bot, ALLOWED_USER_ID, alert)
    # 세정 담당자에게 "잘 받았고 처리 후 증명서 보내드리겠다" 답장 + 텔레그램 알림(2026-11-23)
    # 이미 처리한 메일·이미 기재된 매장만 있는 메일은 위에서 끝나므로 답장이 두 번 가지 않음
    if ack_to and SEJUNG_AUTO_ACK:
        await _sj_send_ack(bot, ack_to, sender_name, subject, message_id, live, brand)
    await asyncio.to_thread(_sj_history_record, message_id, file_hashes, sender_label, stores_all)
    return True


# ---------- 처브 증명서(PDF) 읽기 ----------

def _sj_is_chubb_cert_text(text: str) -> bool:
    t = text.replace(" ", "")
    return ("보험가입증명서" in t or "배서증명서" in t or "CertificationofInsurance" in t
            or "CertificationofEndorsement" in t) and ("처브" in t or "Chubb" in t or "CHUBB" in t or "ACE" in t)


_SJ_DATE_RX = re.compile(r"(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일")


def _sj_regex_cert_fields(text: str, filename: str = "") -> dict:
    """AI 없이 PDF 글자/파일명에서 핵심 값(매장명·증권번호·보험기간 등)을 직접 찾음.
    처브가 '원계약 양식'으로 다시 보낸 증명서(2026-11-18)는 AI가 매장명을 못 뽑아서 생긴 예비 장치.
    금액(재고/시설/건물)은 칸 순서가 뒤섞여 확실하지 않아 여기서는 안 뽑음(None이면 신청서 값을 씀)."""
    out: dict = {}
    t = text or ""
    nospace = t.replace(" ", "")
    out["doc_type"] = "배서증명서" if "배서증명서" in nospace else "가입증명서"
    m = re.search(r"배서구분\s*(\S+)", t)
    if m:
        out["endorsement_type"] = m.group(1)
    m = re.search(r"기준일\s*(\d{4}-\d{1,2}-\d{1,2})", t)
    if m:
        out["base_date"] = m.group(1)
    m = re.search(r"\b\d{3}[A-Z]\d{6}\b", t)
    if m:
        out["policy_no"] = m.group(0)
    m = re.search(r"각\s*대리점주\s*\(([^)]*)\)?", t.replace("\n", ""))
    if m:
        out["insured"] = "㈜세정, ㈜세정글로벌 및 각 대리점주 (" + m.group(1).strip() + ")"
    m = re.search(r"면적\s*\(PY\)\s*([\d.]+)", t)
    if m:
        try:
            area = float(m.group(1))
            if area > 0:  # 0평으로 적힌 증명서(배상 미가입 매장 등)는 리스트 평수를 덮어쓰지 않음
                out["area_py"] = int(area) if area.is_integer() else area
        except ValueError:
            pass
    m = re.search(r"소재지[ \t]*주소[ \t]*(\S.+)", t)
    if m and re.search(r"[시군구]", m.group(1)):
        out["address"] = m.group(1).strip()
    # 매장명: 예전 양식 '수신인 (점포명) TR장유롯데마트 배서보험료', 원계약 양식 'TR장유롯데마트 총보험료'
    store = None
    m = re.search(r"수신인\s*(?:\(점포명\))?\s*(\S[^\n]*?)\s*(?:배서보험료|총보험료|환급보험료)", t)
    if m:
        store = m.group(1).strip()
    if not store:
        for line in t.split("\n"):
            m = re.match(r"^\s*(\S[^\n]*?)\s+(?:총보험료|배서보험료|환급보험료)\s*[\d,]*\s*$", line)
            if m and "수신인" not in m.group(1):
                store = m.group(1).strip()
                break
    # 파일명(예: 'WMJ530A_롯데동부산아울렛점_2026-11-13.pdf')에서 매장코드/매장명 보강
    base = os.path.splitext(os.path.basename(filename or ""))[0]
    base = re.sub(r"^\[(?:가입|해지)\]", "", base)
    parts = [p for p in base.split("_") if p]
    # 매장코드 모양: WMJ530A, OLK0791, WMO0171, 특판 코드 WMO01A1(숫자 사이 영문)까지(2026-11-24)
    if parts and re.fullmatch(r"[A-Z]{2,4}\d[A-Z0-9]{2,5}", parts[0]):
        out["store_code"] = parts[0]
        parts = parts[1:]
    # '◆ 오뷔엘알_OLM7201_매장명_날짜.pdf'처럼 코드가 맨 앞이 아닌 파일명도(2026-11-26)
    if not out.get("store_code"):
        c = _sj_code_from_filename(filename)
        if c:
            out["store_code"] = c
    if not store:
        store = _sj_store_from_filename(filename)
    if not store and parts and not re.fullmatch(r"[\d\-]+", parts[0]):
        store = parts[0]
    if store:
        out["store_name"] = store
    # 보험기간: 발행일자를 뺀 날짜 중 가장 이른 날=시작, 가장 늦은 날=종기
    dates = []
    for line in t.split("\n"):
        if "발행일자" in line:
            continue
        for y, mo, d in _SJ_DATE_RX.findall(line):
            try:
                dates.append(dt.date(int(y), int(mo), int(d)))
            except ValueError:
                pass
    if dates:
        out["start_date"] = min(dates).isoformat()
        if max(dates) != min(dates):
            out["end_date"] = max(dates).isoformat()
    return out


def _sj_cert_amounts(pdf_bytes: bytes) -> dict:
    """처브 가입증명서의 재물 가입금액을 칸 이름 그대로 읽음(2026-11-24).
    증명서는 재고를 '동산'(재고동산)으로 적는데, AI가 이 금액을 시설/비품에 넣는 경우가 있었음
    (웰메이드 특판 9건 중 7건). 글자 배치를 유지하는 layout 모드로 읽어서
    '동산 70,000,000원' -> 재고, '시설/비품 - 원' -> 0, '건물 - 원' -> 0 으로 확정함."""
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        t = "\n".join((p.extract_text(extraction_mode="layout") or "") for p in reader.pages)
    except Exception:
        logger.exception("증명서 금액 칸 읽기 실패(layout)")
        return {}
    t = unicodedata.normalize("NFC", t)

    def amt(label):
        m = re.search(label + r"[ \t]+([\d,]+|-)[ \t]*원", t)
        if not m:
            return None
        return 0 if m.group(1) == "-" else int(m.group(1).replace(",", ""))

    out = {}
    stock = amt(r"(?:재고[ \t]*)?동산")
    if stock is None:
        stock = amt(r"재고(?:자산)?")
    fac = amt(r"시설[ \t]*/[ \t]*비품")
    bld = amt(r"(?<![가-힣])건물")
    if stock is not None:
        out["stock"] = stock
    if fac is not None:
        out["facility"] = fac
    if bld is not None:
        out["building"] = bld
    return out


def _sj_extract_cert_fields(text: str, filename: str = "") -> dict | None:
    """PDF에서 뽑은 글자(순서가 뒤섞여 있음)를 Claude로 정리해서 필요한 값만 뽑음.
    AI 답이 비거나 깨지면 _sj_regex_cert_fields 값으로 채움(2026-11-18)."""
    system = (
        "아래는 처브(에이스아메리칸화재해상보험)의 보험가입증명서 또는 배서증명서 PDF에서 뽑은 텍스트예요. "
        "글자 순서가 뒤섞여 있으니 의미를 보고 값을 찾아서, 아래 JSON 형식으로만 답하세요(설명 없이).\n"
        '{"doc_type": "가입증명서" 또는 "배서증명서", "endorsement_type": 배서구분 글자 그대로 또는 null, '
        '"store_name": 보험가입증명서 수신인(매장명), "policy_no": 증권번호, "insured": 보험계약자 문구, '
        '"address": 소재지 주소, "area_py": 면적 숫자, "stock": 재고 금액 숫자, "facility": 시설/비품 금액 숫자, '
        '"building": 건물 금액 숫자, "total_insured": 총가입금액 숫자, "start_date": "YYYY-MM-DD", '
        '"end_date": "YYYY-MM-DD", "base_date": 배서 기준일 "YYYY-MM-DD" 또는 null, '
        '"annual_property_premium": 제1부문 연간보험료 숫자, "annual_liability_premium": 제4부문 연간보험료 숫자, '
        '"total_premium": 총보험료 또는 배서보험료 숫자, "property_refund": 재물 환급보험료 숫자 또는 null, '
        '"liability_refund": 배상 환급보험료 숫자 또는 null, "warehouse": 창고위치}\n'
        "금액에 '-'만 있으면 0, 모르면 null로 두세요. 숫자는 쉼표 없이 쓰고, JSON 앞뒤에 다른 글은 쓰지 마세요.\n"
        "'동산'(재고동산) 금액은 반드시 stock(재고)이에요. facility(시설/비품)에 넣지 마세요."
    )
    data: dict = {}
    raw = ""
    try:
        resp = client.messages.create(
            model=MODEL_NAME, max_tokens=2000, system=system,
            messages=[{"role": "user", "content": text[:8000]}],
        )
        raw = "".join(b.text for b in resp.content if b.type == "text").strip()
        m = re.search(r"\{.*\}", raw, re.S)
        if m:
            data = json.loads(m.group(0))
        else:
            logger.warning("처브 증명서: AI 답에서 JSON을 못 찾음(stop=%s): %s",
                           getattr(resp, "stop_reason", None), raw[:500])
    except Exception:
        logger.exception("처브 증명서 내용 읽기 중 오류. AI 답 앞부분: %s", raw[:500])
        data = {}
    if not isinstance(data, dict):
        data = {}
    # AI가 못 채운 값은 PDF 글자/파일명에서 직접 찾은 값으로 보강
    rx = _sj_regex_cert_fields(text, filename)
    for k, v in rx.items():
        if data.get(k) in (None, "", "null"):
            data[k] = v
    if not data.get("store_name"):
        logger.warning("처브 증명서: 매장명 못 찾음. AI 답 앞부분: %s", raw[:300])
        return None
    et = str(data.get("endorsement_type") or "")
    if data.get("doc_type") == "가입증명서" or "추가" in et:
        data["kind"] = "추가"
    else:
        data["kind"] = "해지"
    return data


def _sj_to_date(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return v if isinstance(v, dt.datetime) else dt.datetime(v.year, v.month, v.day)
    if isinstance(v, str):
        m = re.search(r"(\d{4})[-./년\s]+(\d{1,2})[-./월\s]+(\d{1,2})", v)
        if m:
            return dt.datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


# ---------- 드라이브: 정산관리 폴더 / 통합관리리스트 ----------

_sejung_root_cache = None


def _sj_find_root(service) -> str | None:
    global _sejung_root_cache
    if SEJUNG_FOLDER_ID:
        return SEJUNG_FOLDER_ID
    if _sejung_root_cache:
        return _sejung_root_cache
    resp = service.files().list(
        q="name='정산관리' and mimeType='application/vnd.google-apps.folder' and trashed=false",
        fields="files(id,name,parents)", pageSize=20,
    ).execute()
    files = resp.get("files", [])
    chosen = None
    for f in files:
        for p in f.get("parents") or []:
            try:
                parent = service.files().get(fileId=p, fields="name").execute()
                if unicodedata.normalize("NFC", parent.get("name", "")).strip() == "26년":
                    chosen = f["id"]
            except Exception:
                pass
    if chosen is None and len(files) == 1:
        chosen = files[0]["id"]
    _sejung_root_cache = chosen
    return chosen


def _sj_find_master(service, root_id: str) -> dict | None:
    """정산관리 폴더(없으면 그 상위 폴더)에서 이름에 '통합관리'가 들어간 가장 최근 엑셀을 찾음."""
    parents = [root_id]
    try:
        meta = service.files().get(fileId=root_id, fields="parents").execute()
        parents += meta.get("parents") or []
    except Exception:
        pass
    for pid in parents:
        resp = service.files().list(
            q=f"'{pid}' in parents and trashed=false and mimeType='{_XLSX_MIMETYPE}'",
            fields="files(id,name,modifiedTime)", orderBy="modifiedTime desc", pageSize=50,
        ).execute()
        for f in resp.get("files", []):
            name = unicodedata.normalize("NFC", f["name"])
            if not name.startswith("~$") and "통합관리" in name:
                return f
    return None


def _sj_drive_save_pdf_once(service, folder_id: str, name: str, pdf_bytes: bytes) -> str:
    """같은 증명서가 폴더에 두 번 쌓이지 않게 저장(2026-11-18 요청).
    - 내용이 똑같은 파일이 이미 있으면: 새로 만들지 않음(이름이 다르면 정해진 이름으로 바꿈)
    - 이름이 같은 파일이 있으면(처브 재송부 등 내용만 다름): 그 파일 내용을 새 것으로 교체
    - 둘 다 없으면 새로 만듦
    반환: 'same' / 'renamed:<옛 이름>' / 'replaced' / 'created'"""
    import hashlib
    md5 = hashlib.md5(pdf_bytes).hexdigest()
    files, token = [], None
    while True:
        resp = service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            fields="nextPageToken, files(id,name,md5Checksum)", pageSize=200, pageToken=token,
        ).execute()
        files += resp.get("files", [])
        token = resp.get("nextPageToken")
        if not token:
            break
    want = unicodedata.normalize("NFC", name)
    same_name = [f for f in files if unicodedata.normalize("NFC", f.get("name", "")) == want]
    same_content = [f for f in files if f.get("md5Checksum") == md5]
    if same_content:
        keep = next((f for f in same_content if f in same_name), same_content[0])
        if unicodedata.normalize("NFC", keep.get("name", "")) == want:
            return "same"
        if same_name:
            # 정해진 이름의 파일이 따로 이미 있으면 이름을 바꾸면 이름이 겹치니, 그 파일을 새 내용으로 교체만 함
            service.files().update(
                fileId=same_name[0]["id"], media_body=_gdrive_media(pdf_bytes, mimetype="application/pdf"),
            ).execute()
            return "replaced"
        service.files().update(fileId=keep["id"], body={"name": name}).execute()
        return "renamed:" + keep.get("name", "")
    if same_name:
        service.files().update(
            fileId=same_name[0]["id"], media_body=_gdrive_media(pdf_bytes, mimetype="application/pdf"),
        ).execute()
        return "replaced"
    service.files().create(
        body={"name": name, "parents": [folder_id]},
        media_body=_gdrive_media(pdf_bytes, mimetype="application/pdf"), fields="id",
    ).execute()
    return "created"


def _sj_month_folder(service, root_id: str, brand: str, sub: str | None = None) -> str:
    """sub는 '변경신청서/추가가입신청서'처럼 '/'로 여러 단계를 줄 수 있음."""
    """정산관리/브랜드/(종류 폴더)/YY년 M월 폴더 id. 종류 폴더는 유진님이 정한 구조에 맞춤:
    '추가변경신청서', '해지변경신청서', '추가가입증명서', '해지변경증명서'(2026-11-13). 없으면 만듦."""
    def _same(target):
        t = _sj_norm(target)
        return lambda name: _sj_norm(name) == t
    brand_id = _gdrive_find_child_folder(service, root_id, _same(brand))
    if brand_id is None:
        brand_id = _gdrive_create_folder(service, root_id, brand)
    for part in [p for p in (sub or "").split("/") if p]:
        sub_id = _gdrive_find_child_folder(service, brand_id, _same(part))
        if sub_id is None:
            sub_id = _gdrive_create_folder(service, brand_id, part)
        brand_id = sub_id
    today = dt.datetime.now(KST).date()
    month_name = f"{today:%y}년 {today.month}월"
    month_id = _gdrive_find_child_folder(service, brand_id, _same(month_name))
    if month_id is None:
        month_id = _gdrive_create_folder(service, brand_id, month_name)
    return month_id


_SJ_LOOKUP_RANGE_RE = re.compile(r"(최초가입전체리스트'?!\$[A-Z]+\$\d+:\$[A-Z]+\$)(\d+)")


def _sj_extend_lookup_ranges(ws) -> int:
    """해지매장 시트의 연간총보험료 수식(VLOOKUP)이 최초가입전체리스트의 일부 행까지만 보고 있으면
    (예: $E$4:$V$293인데 리스트는 866행까지), 리스트 끝까지 보도록 범위를 자동으로 늘림
    (2026-11-12 유진님 지시: 모자라면 봇이 알아서). 고친 칸 수를 반환."""
    wb = ws.parent
    last = 0
    for n in wb.sheetnames:
        if "최초가입" in unicodedata.normalize("NFC", n):
            last = wb[n].max_row
            break
    if not last:
        return 0
    fixed = 0
    for row in ws.iter_rows(min_row=4, max_row=ws.max_row):
        for cell in row:
            v = cell.value
            if not (isinstance(v, str) and v.startswith("=") and "최초가입전체리스트" in v):
                continue
            new_v = _SJ_LOOKUP_RANGE_RE.sub(
                lambda m: m.group(1) + str(max(int(m.group(2)), last)), v
            )
            if new_v != v:
                cell.value = new_v
                fixed += 1
    return fixed


# 처브 증명서가 오면 신청서 때 적은 값과 비교해서, 다르면 증명서 기준으로 고치는 칸들
# (주소는 표기 방식만 달라도 다르게 보여서 비교 대상에서 뺌 - 비어 있을 때만 채움)
_SJ_CERT_OVERWRITE_KEYS = ("재고", "시설/비품", "건물", "평수", "보험시작일", "보험종기일", "폐점일",
                           "재물부문환급", "영업부문환급")
_SJ_STATUS_WAIT = "[증명서 대기]"


def _sj_same_value(a, b) -> bool:
    if a in (None, "") and b in (None, ""):
        return True
    da, db = _sj_to_date(a), _sj_to_date(b)
    if da and db:
        return da.date() == db.date()
    try:
        return abs(float(a or 0) - float(b or 0)) <= 1
    except (TypeError, ValueError):
        return _sj_norm(a) == _sj_norm(b)


def _sj_fmt(v) -> str:
    d = _sj_to_date(v)
    if d:
        return d.strftime("%Y-%m-%d")
    if isinstance(v, (int, float)):
        return f"{v:,.0f}"
    return str(v) if v not in (None, "") else "(빈칸)"


def _sj_rebuild_month_column(ws, header: dict, first_row: int = 4) -> None:
    """'월별' 칸(A열)에 접수일자 기준 'YY년 M월'을 적고, 같은 달 줄끼리 세로로 병합함
    (2026-11-16 유진님 요청: 그달 들어온 건은 그달로 묶어서 월 표시). 매번 다시 계산함."""
    mcol = header.get("월별")
    ncol = header.get("매장명")
    dcol = header.get("접수일자")
    if not (mcol and ncol):
        return
    for m in list(ws.merged_cells.ranges):
        if m.min_col == mcol and m.max_col == mcol and m.min_row >= first_row:
            ws.unmerge_cells(str(m))
    last = first_row - 1
    for r in range(first_row, ws.max_row + 1):
        if ws.cell(row=r, column=ncol).value not in (None, ""):
            last = r
    labels = []
    for r in range(first_row, last + 1):
        d = _sj_to_date(ws.cell(row=r, column=dcol).value) if dcol else None
        has = ws.cell(row=r, column=ncol).value not in (None, "")
        labels.append(f"{d:%y}년 {d.month}월" if (has and d) else None)
        ws.cell(row=r, column=mcol).value = None
    r = first_row
    while r <= last:
        lab = labels[r - first_row]
        end = r
        while end + 1 <= last and labels[end + 1 - first_row] == lab:
            end += 1
        if lab:
            ws.cell(row=r, column=mcol).value = lab
            if end > r:
                ws.merge_cells(start_row=r, start_column=mcol, end_row=end, end_column=mcol)
        r = end + 1


def _sj_write_master(wb, brand: str, kind: str, values: dict, update_match: bool = False,
                     cert_review: bool = False, changes: list | None = None,
                     force_row: int | None = None) -> str:
    """통합관리리스트의 '{브랜드}_추가매장' 또는 '{브랜드}_해지매장' 시트 빈 줄에 기재.
    수식 칸은 건드리지 않고(맨 첫 데이터 줄의 수식을 복사), 이미 값이 들어 있는 칸(종기일 등
    미리 넣어둔 값)도 덮어쓰지 않음. 기재한 위치 설명을 반환.
    force_row: 중복 확인에서 [덮어쓰기]를 누른 경우 그 줄에 새 값을 모두 덮어씀(2026-11-23)."""
    want = "추가" if kind == "추가" else "해지"
    ws = None
    for n in wb.sheetnames:
        nn = unicodedata.normalize("NFC", n)
        if brand in nn and want in nn:
            ws = wb[n]
            break
    if ws is None:
        raise ValueError(f"통합관리리스트에 '{brand}_{want}매장' 시트가 없어요")

    header = {}
    for col in range(1, ws.max_column + 1):
        label = ws.cell(row=3, column=col).value or ws.cell(row=2, column=col).value
        if label not in (None, ""):
            header.setdefault(_norm_header(label), col)
    name_col = header.get("매장명")
    if name_col is None:
        raise ValueError(f"'{ws.title}' 시트에서 매장명 칸을 못 찾았어요")

    _sj_extend_lookup_ranges(ws)

    template_row = 4
    target = None
    if force_row:
        target = force_row
    elif update_match:
        # 신청서 때 이미 적어둔 같은 매장 줄이 있으면 그 줄을 갱신(중복 줄 방지). 날짜가 둘 다
        # 있는데 다르면(같은 매장의 다른 건) 새 줄로 봄.
        # 2026-11-17 보강: 날짜가 달라도 같은 매장 줄이 아직 '증명서 대기'(처리 중)면 같은 건으로 보고
        # 그 줄을 고쳐 씀. (롯데김해아울렛직영특판점이 1차로 잘못 읽은 날짜 11/12로 한 줄,
        # 2차로 바른 날짜 12/09로 또 한 줄 들어가 중복된 사고 방지)
        date_key = "보험시작일" if want == "추가" else "폐점일"
        dcol = header.get(date_key)
        want_d = _sj_to_date(values.get(date_key))
        name_target = _sj_norm(values.get("매장명"))
        st_col = header.get("업무상태")
        nt_col = header.get("비고")

        def _sj_row_pending(row: int) -> bool:
            if st_col:
                sv = str(ws.cell(row=row, column=st_col).value or "")
                if sv.strip():
                    return "대기" in sv
            if nt_col:
                return _SJ_STATUS_WAIT in str(ws.cell(row=row, column=nt_col).value or "")
            return False

        addr_col = header.get("매장주소")
        addr_target = _sj_norm(values.get("매장주소"))
        code_col = header.get("매장코드")
        code_target = str(values.get("매장코드") or "").strip().upper()

        def _addr_key(v) -> str:
            return re.sub(r"[\s,.\-()·]", "", unicodedata.normalize("NFC", str(v or ""))).lower()

        addr_key_target = _addr_key(values.get("매장주소"))
        exact, pending, similar = None, None, None
        addr_hits = []
        for r in range(template_row, ws.max_row + 1):
            row_name = ws.cell(row=r, column=name_col).value
            if row_name in (None, ""):
                continue
            cd = _sj_to_date(ws.cell(row=r, column=dcol).value) if dcol else None
            same_date = want_d is None or cd is None or cd.date() == want_d.date()
            # 2026-11-24: 매장코드가 같으면 매장명이 달라도 같은 매장(예: 'WMC안산고잔' = 'WMCNC고잔점(특판)')
            row_code = str(ws.cell(row=r, column=code_col).value or "").strip().upper() if code_col else ""
            if code_target and row_code and row_code == code_target and same_date:
                exact = r
                continue
            # 2026-11-24: 증명서 매장명이 리스트와 많이 달라도(유진님이 직접 고친 이름 등), 아직 '증명서 대기'인
            # 줄 중 주소가 같고 날짜가 같은 줄은 같은 건 후보로 둠(1곳일 때만 씀)
            if cert_review and addr_key_target and len(addr_key_target) >= 10 and _sj_row_pending(r) and same_date:
                rk = _addr_key(ws.cell(row=r, column=addr_col).value) if addr_col else ""
                if rk and len(rk) >= 10 and (rk == addr_key_target or rk in addr_key_target or addr_key_target in rk):
                    addr_hits.append(r)
            if _sj_norm(row_name) != name_target:
                # 2026-11-18 보강: 증명서 매장명이 신청서와 조금 달라도(예: '직영' 빠짐), 아직 '증명서 대기'인
                # 줄 중 매장명이 사실상 같고 날짜나 주소가 같은 줄이면 같은 건으로 보고 그 줄을 고침.
                # (롯데김해아울렛직영특판점 / 롯데김해아울렛특판점 두 줄 중복 사고 방지)
                if (cert_review and similar is None and _sj_row_pending(r)
                        and _sj_store_similar(row_name, values.get("매장명"))):
                    row_addr = _sj_norm(ws.cell(row=r, column=addr_col).value) if addr_col else ""
                    if want_d is not None and cd is not None:
                        ok = same_date  # 날짜가 둘 다 있으면 날짜가 같아야 같은 건
                    else:
                        ok = bool(addr_target and row_addr and (addr_target in row_addr or row_addr in addr_target))
                    if ok:
                        similar = r
                continue
            if same_date:
                exact = r
            elif _sj_row_pending(r):
                pending = r
        target = exact if exact is not None else (pending if pending is not None else similar)
        if target is None and len(addr_hits) == 1:
            target = addr_hits[0]
    updating = target is not None
    if target is None:
        target = template_row
        while ws.cell(row=target, column=name_col).value not in (None, ""):
            target += 1

    if target != template_row and not updating:
        # 새로 늘어나는 줄도 표 서식(테두리·정렬·숫자서식·행 높이)을 첫 데이터 줄과 똑같이 맞춤
        if ws.row_dimensions[template_row].height:
            ws.row_dimensions[target].height = ws.row_dimensions[template_row].height
        for col in range(1, ws.max_column + 1):
            tcell, cell = ws.cell(row=template_row, column=col), ws.cell(row=target, column=col)
            if not isinstance(cell, MergedCell) and not isinstance(tcell, MergedCell) and tcell.has_style:
                cell._style = copy.copy(tcell._style)

    for col in range(1, ws.max_column + 1):
        tcell = ws.cell(row=template_row, column=col)
        cell = ws.cell(row=target, column=col)
        if isinstance(cell, MergedCell):
            continue
        tv = tcell.value
        if isinstance(tv, str) and tv.startswith("="):
            if cell.value in (None, ""):
                cell.value = Translator(tv, origin=tcell.coordinate).translate_formula(cell.coordinate) \
                    if target != template_row else tv
            continue
        if target != template_row and cell.value in (None, "") and tcell.number_format:
            cell.number_format = tcell.number_format

    for key, col in header.items():
        if key not in values or values[key] in (None, ""):
            continue
        cell = ws.cell(row=target, column=col)
        if isinstance(cell, MergedCell):
            continue
        if isinstance(cell.value, str) and cell.value.startswith("="):
            continue
        # 보험 시작/종기/폐점일은 증명서 날짜가 정답이라, 양식에 미리 적힌 날짜가 있어도 덮어씀
        overwrite = ("보험시작일", "보험종기일", "폐점일") + (_SJ_CERT_OVERWRITE_KEYS if updating else ())
        if cell.value not in (None, "") and key not in overwrite and not force_row:
            continue
        if updating and cert_review and cell.value not in (None, "") and key in _SJ_CERT_OVERWRITE_KEYS:
            if _sj_same_value(cell.value, values[key]):
                continue
            if changes is not None:
                changes.append(f"{key}: {_sj_fmt(cell.value)} → {_sj_fmt(values[key])}")
        v = values[key]
        if key in ("접수일자", "보험시작일", "보험종기일", "폐점일"):
            v = _sj_to_date(v) or v
            if isinstance(v, dt.datetime):
                cell.number_format = "yyyy-mm-dd"
        elif isinstance(v, (int, float)) and cell.is_date:
            # 양식에 날짜 서식이 걸려 있는 금액 칸(예: 해지매장 재물/영업부문 환급)은 숫자 서식으로 바꿈
            cell.number_format = "#,##0"
        cell.value = v
    # 비고 끝에 덧붙일 내용(예: 처브 메일의 '환급날짜: 11/18')
    extra_note = values.get("_비고추가")
    note_col_x = header.get("비고")
    if extra_note and note_col_x and not isinstance(ws.cell(row=target, column=note_col_x), MergedCell):
        nc = ws.cell(row=target, column=note_col_x)
        cur = str(nc.value or "")
        if extra_note not in cur:
            nc.value = f"{cur} / {extra_note}" if cur else extra_note

    # 순번: 수식이 없는 시트(해지매장)는 줄 번호를 직접 적음(추가매장은 =ROW() 수식이 알아서 계산)
    seq_col = header.get("순번")
    if seq_col:
        seq_cell = ws.cell(row=target, column=seq_col)
        if not isinstance(seq_cell, MergedCell) and not (isinstance(seq_cell.value, str) and str(seq_cell.value).startswith("=")):
            seq_cell.value = target - template_row + 1
    _sj_rebuild_month_column(ws, header, template_row)

    # 진행 상태 표시: '업무상태' 칸(2026-11-12 비고 앞에 추가)이 있으면 거기에, 없으면 비고 끝에 붙임.
    # 신청서 때 '증명서 대기' -> 증명서 검토 후 '증명서 확인 MM/DD'(고친 게 있으면 '(수정)' 표시)
    today_md = f"{dt.datetime.now(KST):%m/%d}"
    status_col = header.get("업무상태")
    if status_col and not isinstance(ws.cell(row=target, column=status_col), MergedCell):
        st_cell = ws.cell(row=target, column=status_col)
        if cert_review:
            st_cell.value = f"증명서 확인 {today_md}" + (" (수정)" if changes else "")
        elif not str(st_cell.value or "").startswith("증명서 확인"):
            st_cell.value = "증명서 대기"
    else:
        note_col = header.get("비고")
        if note_col and not isinstance(ws.cell(row=target, column=note_col), MergedCell):
            note_cell = ws.cell(row=target, column=note_col)
            note = str(note_cell.value or "")
            if cert_review:
                done = f"[증명서 확인 {today_md}]"
                note = note.replace(_SJ_STATUS_WAIT, done) if _SJ_STATUS_WAIT in note else f"{note} {done}".strip()
            elif _SJ_STATUS_WAIT not in note and "[증명서 확인" not in note:
                note = f"{note} {_SJ_STATUS_WAIT}".strip()
            note_cell.value = note
    return f"'{ws.title}' 시트 {target}행" + (" (기존 줄 갱신)" if updating else "")


def _sj_build_values(kind: str, cert: dict, req: dict | None) -> dict:
    f = dict((req or {}).get("fields") or {})
    v = {
        "접수자": f.get("접수자") or (req or {}).get("sender_name"),
        "접수일자": f.get("접수일자") or (req or {}).get("received") or dt.date.today().isoformat(),
        "매장명": (req or {}).get("store_name") or cert.get("store_name"),
        "대표자명": f.get("대표자명"),
        "사업자번호": f.get("사업자번호"),
        "전화번호": f.get("전화번호"),
        "매장주소": f.get("매장주소") or cert.get("address"),
        "비고": f.get("비고"),
    }
    if kind == "추가":
        v.update({
            "매장코드": f.get("매장코드"),
            "재고": cert.get("stock") if cert.get("stock") is not None else f.get("재고"),
            "시설/비품": cert.get("facility") if cert.get("facility") is not None else f.get("시설/비품"),
            "건물": cert.get("building") if cert.get("building") is not None else f.get("건물"),
            "평수": cert.get("area_py") or f.get("평수"),
            "보험시작일": cert.get("start_date") or f.get("보험시작일"),
            "보험종기일": cert.get("end_date"),
            "창고주소": cert.get("warehouse") or f.get("창고주소"),
        })
    else:
        v.update({
            "폐점일": cert.get("base_date") or cert.get("start_date") or f.get("폐점일"),
            "재물부문환급": cert.get("property_refund") if cert.get("property_refund") is not None else f.get("재물부문환급"),
            "영업부문환급": cert.get("liability_refund") if cert.get("liability_refund") is not None else f.get("영업부문환급"),
        })
    return v


def _sj_lookup_store_code(wb, store_name: str) -> str | None:
    """통합관리리스트의 최초가입전체리스트(C=매장코드, E=매장명)와 추가매장 시트에서 매장코드를 찾음.
    2026-11-24: 매장명이 정확히 같을 때만(공백·괄호·끝 '점' 차이는 같게 봄). 이름 일부만 겹치는 다른 매장
    코드를 가져오던 문제 수정('NC고잔점' 코드가 'WMCNC고잔점(특판)'에 붙음)."""
    target = _sj_store_key(store_name)
    raw_target = _sj_norm(store_name)
    if not target:
        return None
    for n in wb.sheetnames:
        ws = wb[n]
        nn = unicodedata.normalize("NFC", n)
        if "최초가입" in nn:
            code_col, name_col = 3, 5
        elif "추가" in nn:
            code_col, name_col = 6, 5
        else:
            continue
        for r in range(4, ws.max_row + 1):
            name_v = ws.cell(row=r, column=name_col).value
            code = ws.cell(row=r, column=code_col).value
            if name_v and code and (_sj_norm(name_v) == raw_target or _sj_store_key(name_v) == target):
                return str(code).strip()
    return None


def _sj_cert_filename(code: str | None, store: str, kind: str, cert: dict, req: dict | None) -> str:
    """'매장코드_매장명_개시일(추가) 또는 해지일(해지).pdf' (유진님 지정 형식, 예: OLC0011_OL대전중리점_2026-11-12.pdf)"""
    f = (req or {}).get("fields") or {}
    if kind == "추가":
        d = cert.get("start_date") or f.get("보험시작일")
    else:
        d = cert.get("base_date") or cert.get("start_date") or f.get("폐점일")
    d = _sj_to_date(d)
    date_text = d.strftime("%Y-%m-%d") if d else dt.date.today().isoformat()
    safe_store = re.sub(r'[\\/:*?"<>|]', "", store).strip()
    # 가입증명서와 해지 변경증명서가 같은 매장·비슷한 날짜면 이름이 같아져서 앞에 구분 표시를 붙임(2026-11-16 요청)
    prefix = "[가입]" if kind == "추가" else "[해지]"
    return f"{prefix}{code or '코드미확인'}_{safe_store}_{date_text}.pdf"


def _sj_record_cert_to_drive(brand: str, kind: str, cert: dict, req: dict | None,
                              filename: str, pdf_bytes: bytes, refund_date: str | None = None) -> tuple[str, str, str]:
    """(통합리스트 기재 결과 문구, PDF 저장 결과 문구, 저장 파일명) 반환. 실패하면 오류 문구를 담음."""
    store = (req or {}).get("store_name") or str(cert.get("store_name") or "").strip()
    req_code = ((req or {}).get("fields") or {}).get("매장코드")
    service = _get_gdrive_service()
    if service is None:
        name = _sj_cert_filename(req_code, store, kind, cert, req)
        return ("❌ 구글 드라이브에 연결되지 않아 통합리스트에 기재하지 못했어요.",
                "❌ 드라이브 저장 실패(연결 안 됨)", name)
    root = _sj_find_root(service)
    if root is None:
        name = _sj_cert_filename(req_code, store, kind, cert, req)
        return ("❌ 드라이브에서 '정산관리' 폴더를 못 찾았어요.",
                "❌ 드라이브 저장 실패('정산관리' 폴더 없음)", name)

    master, wb = None, None
    try:
        master = _sj_find_master(service, root)
        if master is not None:
            content = _gdrive_download_file(service, master["id"])
            wb = openpyxl.load_workbook(io.BytesIO(content), data_only=False)
    except Exception:
        logger.exception("통합관리리스트 열기 중 오류")

    # 매장코드 우선순위(2026-11-24 수정): 신청서 → 처브 증명서 파일명(처브는 신청서 파일명 코드를 그대로 씀)
    # → 통합관리리스트에서 매장명이 '정확히' 같은 줄. 예전엔 리스트에서 이름이 '포함'만 돼도 가져와서
    # 'WMCNC고잔점(특판)'에 정규매장 'NC고잔점'의 코드 WMJ0582가 붙었음.
    code = (req_code or cert.get("store_code")
            or (_sj_lookup_store_code(wb, store) if wb is not None else None))
    out_name = _sj_cert_filename(code, store, kind, cert, req)

    # PDF 저장
    try:
        sub = "추가가입증명서" if kind == "추가" else "해지변경증명서"
        month_id = _sj_month_folder(service, root, brand, sub)
        action = _sj_drive_save_pdf_once(service, month_id, out_name, pdf_bytes)
        today = dt.datetime.now(KST).date()
        where_txt = f"드라이브 정산관리/{brand}/{sub}/{today:%y}년 {today.month}월 폴더"
        if action == "same":
            saved_msg = f"💾 {where_txt}에 같은 파일 '{out_name}'이 이미 있어서 새로 저장하지 않았어요."
        elif action.startswith("renamed:"):
            saved_msg = (f"💾 {where_txt}에 같은 내용의 파일('{action[8:]}')이 이미 있어서, "
                         f"새로 저장하지 않고 이름만 '{out_name}'로 바꿨어요.")
        elif action == "replaced":
            saved_msg = f"💾 {where_txt}의 '{out_name}'을 새로 받은 증명서로 바꿔 저장했어요(중복 없이)."
        else:
            saved_msg = f"💾 {where_txt}에 '{out_name}'로 저장했어요."
    except Exception as e:
        logger.exception("처브 증명서 드라이브 저장 중 오류")
        saved_msg = f"❌ 드라이브 저장 실패: {type(e).__name__}"

    # 통합관리리스트 기재
    try:
        if master is None or wb is None:
            list_msg = ("❌ 정산관리 폴더에서 '통합관리'가 들어간 엑셀을 못 찾아 리스트에 기재하지 못했어요. "
                        "통합관리리스트 파일을 정산관리 폴더에 넣어주세요.")
        else:
            values = _sj_build_values(kind, cert, req)
            if code and not values.get("매장코드"):
                values["매장코드"] = code
            changes = []
            if refund_date:
                values["_비고추가"] = f"환급날짜: {refund_date}"
            where = _sj_write_master(wb, brand, kind, values, update_match=True,
                                     cert_review=True, changes=changes)
            out = io.BytesIO()
            wb.save(out)
            ok = _save_brand_master_to_drive(master["id"], out.getvalue())
            if not ok:
                list_msg = "❌ 통합관리리스트 저장에 실패했어요."
            elif "기존 줄" not in where:
                list_msg = (f"⚠️ 신청서로 적어둔 줄이 없어서 증명서 내용으로 새로 적었어요({where}). "
                            "신청서가 봇을 거치지 않은 건인지 확인해주세요.")
            elif changes:
                list_msg = ("⚠️ 증명서를 검토해보니 신청서 때 적은 내용과 달라서, 증명서 기준으로 고쳤어요"
                            f"({where}):\n" + "\n".join(f"  - {c}" for c in changes))
            else:
                list_msg = f"✅ 증명서를 검토했어요. 신청서 때 적은 내용과 일치해요({where})."
    except Exception as e:
        logger.exception("통합관리리스트 기재 중 오류")
        list_msg = f"❌ 통합관리리스트 기재 실패: {type(e).__name__}: {str(e)[:150]}"
    return list_msg, saved_msg, out_name


async def _sj_queue_to_sejung(bot, target_email: str, store: str, kind: str, brand: str,
                              filename: str, pdf_bytes: bytes, who: str, refund_date: str | None = None) -> None:
    label = "보험가입증명서" if kind == "추가" else "변경증서(환급)"
    cert_id = _queue_pending_certs(
        [{"store_name": store, "out_name": filename, "pdf_bytes": pdf_bytes}],
        target_email=target_email,
        subject=f"[㈜세정 {brand}] {store} {label}",
        body=_sejung_cert_body(kind, refund_date, store),
        label=label,
        sent_keys=[_sj_sent_key("세정", brand, kind, store)],
    )
    await bot.send_message(
        chat_id=ALLOWED_USER_ID,
        text=f"세정 담당자 {who}에게 '{store}' {label}를 보낼까요?",
        reply_markup=_queue_confirm_keyboard(cert_id),
    )


# ---------- 증명서 여러 개를 담당자 1명에게 메일 1통으로 묶어 보내기 (2026-11-24 유진님 요청) ----------

def _sejung_cert_bundle_body(entries: list) -> str:
    adds = [e["store"] for e in entries if e["kind"] == "추가"]
    dels = [e for e in entries if e["kind"] != "추가"]
    if adds and dels:
        what = "추가매장의 보험가입증명서와 해지매장의 변경증서(환급)를"
    elif adds:
        what = "추가매장의 보험가입증명서를"
    else:
        what = "해지매장의 변경증서(환급)를"
    lines = []
    if adds:
        lines.append("- 추가(보험가입증명서): " + ", ".join(adds))
    for e in dels:
        extra = ""
        if e.get("refund_date"):
            mo, d = e["refund_date"].split("/")
            extra = f" (환급보험료 {int(mo)}월 {int(d)}일 지급 예정)"
        lines.append(f"- 해지(변경증명서): {e['store']}{extra}")
    return (
        "안녕하십니까~\n"
        f"보내주신 {what} 보내드립니다.\n\n"
        + "\n".join(lines) + "\n\n"
        "업무에 참고하여 주시기 바랍니다.\n\n"
        "즐거운 하루 보내십시요\n"
        "감사합니다.\n"
        "수고하세요\n\n"
        "--\n" + _CERT_EMAIL_SIGNATURE
    )


async def _sj_queue_bundle(bot, target_email: str, who: str, entries: list) -> None:
    """증명서 여러 개를 메일 1통으로 묶어 발송 대기에 올리고 [보내기] 버튼을 띄움.
    entries: [{store, kind, brand, filename, pdf_bytes, refund_date}]"""
    if len(entries) == 1:
        e = entries[0]
        await _sj_queue_to_sejung(bot, target_email, e["store"], e["kind"], e["brand"], e["filename"],
                                  e["pdf_bytes"], who, e.get("refund_date"))
        return
    kinds = {e["kind"] for e in entries}
    label = ("보험가입증명서·변경증서" if len(kinds) > 1
             else "보험가입증명서" if "추가" in kinds else "변경증서(환급)")
    brands = "/".join(dict.fromkeys(e["brand"] for e in entries if e.get("brand")))
    stores = [e["store"] for e in entries]
    cert_id = _queue_pending_certs(
        [{"store_name": e["store"], "out_name": e["filename"], "pdf_bytes": e["pdf_bytes"]} for e in entries],
        target_email=target_email,
        subject=f"[㈜세정 {brands}] {label} {len(entries)}건 ({stores[0]} 외 {len(entries) - 1}곳)",
        body=_sejung_cert_bundle_body(entries),
        label=label,
        sent_keys=[_sj_sent_key("세정", e.get("brand"), e["kind"], e["store"]) for e in entries],
    )
    await _send_message_retrying(
        bot, ALLOWED_USER_ID,
        f"📨 세정 담당자 {who}({target_email})에게 {label} {len(entries)}건을 메일 1통으로 보낼까요?\n"
        + "\n".join(f"  - {e['kind']} {e['store']}" for e in entries),
        reply_markup=_queue_confirm_keyboard(cert_id),
    )


async def _sj_flush_cert_batch(bot, batch: list) -> None:
    """한 메일에서 모은 증명서를 담당자(받는 사람)별로 묶어 발송 대기에 올림."""
    groups: dict = {}
    for e in batch:
        groups.setdefault(e["target"].lower(), []).append(e)
    for items in groups.values():
        await _sj_queue_bundle(bot, items[0]["target"], items[0]["who"], items)


_SJ_SEND_RX = re.compile(r"보내|발송|전달")


def _sj_info_from_cert_pdf(filename: str, data: bytes) -> dict:
    """텔레그램으로 올린 증명서 PDF에서 매장명/구분/브랜드만 가볍게 읽음(묶음 발송용, AI 안 씀)."""
    text = _extract_pdf_text(data)
    rx = _sj_regex_cert_fields(text, filename)
    base = os.path.splitext(os.path.basename(filename or ""))[0]
    if base.startswith("[해지]") or rx.get("doc_type") == "배서증명서":
        kind = "해지"
    else:
        kind = "추가"
    store = rx.get("store_name") or re.sub(r"^\[(?:가입|해지)\]", "", base)
    brand = _sj_brand_from_text(rx.get("policy_no"), rx.get("insured"), filename, text[:3000])
    return {"store": str(store).strip(), "kind": kind, "brand": brand}


async def _sj_bundle_send_from_caption(bot, caption: str, files: list) -> None:
    """'9개 추가가입증명서를 세정 서석란 과장한테 한번에 보내줘' 같은 지시로 올린 파일들을
    담당자에게 메일 1통으로 묶어 발송 대기에 올림. 이미 처리한 증명서여도 발송만 함(리스트 재기재 없음)."""
    m = re.search(r"[\w.+-]+@[\w-]+\.[\w.]+", caption)
    target, who = (m.group(0), m.group(0)) if m else (None, None)
    if not target:
        name = _sj_name_from_reply(caption)
        target = _load_contacts().get(_norm_contact_name(name)) if name else None
        who = name
    if not target:
        await _send_message_retrying(
            bot, ALLOWED_USER_ID,
            f"❓ '{caption}'에서 받는 사람을 찾지 못했어요. 등록된 담당자 이름(/contacts)이나 이메일을 "
            "같이 적어서 다시 올려주세요.",
        )
        return
    entries = []
    for fn, data in files:
        info = await asyncio.to_thread(_sj_info_from_cert_pdf, fn, data)
        entries.append(dict(info, filename=fn, pdf_bytes=data, refund_date=None))
    if len(entries) == 1:
        # 1건이어도 지시대로 보낼 수 있게 확인 버튼을 띄움
        e = entries[0]
        await _sj_queue_to_sejung(bot, target, e["store"], e["kind"], e["brand"] or "", e["filename"],
                                  e["pdf_bytes"], who, None)
        return
    await _sj_queue_bundle(bot, target, who, entries)


async def _sj_flush_upload_group(context: ContextTypes.DEFAULT_TYPE) -> None:
    """텔레그램에 여러 파일을 한 번에 올린 묶음(앨범)을 다 받은 뒤 처리.
    설명(캡션)에 '보내줘' 같은 지시가 있으면 묶음 발송, 없으면 파일마다 평소처럼 처리."""
    key = context.job.data
    groups = context.application.bot_data.get("sj_upload_groups", {})
    grp = groups.get(key)
    if not grp:
        return
    if time.monotonic() - grp.get("last", 0) < 3:
        # 아직 묶음의 파일이 들어오는 중이면 조금 더 기다림
        context.job_queue.run_once(_sj_flush_upload_group, 3, data=key, name=f"sjgrp-{key}")
        return
    groups.pop(key, None)
    caption = " ".join(c for c in grp["captions"] if c).strip()
    try:
        if caption and _SJ_SEND_RX.search(caption):
            await _sj_bundle_send_from_caption(context.bot, caption, grp["files"])
        else:
            for fn, data in grp["files"]:
                await _sj_process_uploaded_file(context.bot, fn, data)
    except Exception as e:
        logger.exception("텔레그램 묶음 파일 처리 중 오류")
        await _send_message_retrying(
            context.bot, ALLOWED_USER_ID,
            f"⚠️ 올린 파일 묶음 처리 중 오류: {type(e).__name__} ({_sj_err_where()}): {str(e)[:150]}",
        )


def _sj_refund_date_from_text(text: str) -> str | None:
    """처브 메일 본문에서 환급 지급 예정일을 찾아 'MM/DD'로 반환(예: '환급은 11/18 오후 6시경 지급'
    -> '11/18', '11월 18일 환급' -> '11/18'). 못 찾으면 None. (2026-11-16 유진님 요청)"""
    if not text or "환급" not in text:
        return None
    for line in re.split(r"[\n.]", text):
        if "환급" not in line:
            continue
        m = re.search(r"(\d{1,2})\s*/\s*(\d{1,2})", line) or re.search(r"(\d{1,2})\s*월\s*(\d{1,2})\s*일", line)
        if m:
            mo, d = int(m.group(1)), int(m.group(2))
            if 1 <= mo <= 12 and 1 <= d <= 31:
                return f"{mo:02d}/{d:02d}"
    return None


async def _sj_handle_cert_file(bot, filename: str, pdf_bytes: bytes, mail_body: str = "",
                               batch: list | None = None) -> bool:
    """처브 보험가입증명서/배서증명서 PDF면 리스트 기재 + 드라이브 저장 + 세정 담당자 발송 준비.
    처리했으면 True(증명서가 아니면 False).
    batch: 한 메일에 증명서가 여러 개 올 때 담당자별로 묶어 메일 1통으로 보내려고 모아두는 목록
    (2026-11-24). 주면 여기서 바로 발송 대기에 올리지 않고 batch에 담기만 함."""
    text = _extract_pdf_text(pdf_bytes)
    if not text or not _sj_is_chubb_cert_text(text):
        return False

    import hashlib
    digest = hashlib.sha1(pdf_bytes).hexdigest()
    processed = _sj_load(SEJUNG_PROCESSED_FILE, {})
    if digest in processed:
        await _send_message_retrying(
            bot, ALLOWED_USER_ID,
            f"ℹ️ '{filename}'은 이미 처리한 증명서라 다시 처리하지 않았어요({processed[digest]}).",
        )
        return True

    cert = await asyncio.to_thread(_sj_extract_cert_fields, text, filename)
    if cert:
        # 재고(동산)/시설/건물 금액은 증명서 칸 이름으로 직접 읽은 값을 우선(2026-11-24 동산→비품 오기재 수정)
        cert.update(await asyncio.to_thread(_sj_cert_amounts, pdf_bytes))
    if not cert or not cert.get("store_name"):
        await _send_message_retrying(
            bot, ALLOWED_USER_ID,
            f"⚠️ '{filename}'은 처브 증명서 같은데 매장명을 읽지 못했어요. 직접 확인해주세요.",
        )
        return True
    brand = _sj_brand_from_text(cert.get("policy_no"), cert.get("insured"), filename)
    if brand is None:
        await _send_message_retrying(
            bot, ALLOWED_USER_ID,
            f"⚠️ '{filename}'({cert.get('store_name')})의 브랜드를 알아보지 못했어요. 직접 확인해주세요.",
        )
        return True
    kind = cert["kind"]
    store = str(cert["store_name"]).strip()
    refund_date = _sj_refund_date_from_text(mail_body) if kind == "해지" else None
    req = _sj_find_request(brand, store)

    list_msg, saved_msg, out_name = await asyncio.to_thread(
        _sj_record_cert_to_drive, brand, kind, cert, req, filename, pdf_bytes, refund_date,
    )
    filename = out_name  # 세정 담당자에게 보내는 첨부파일 이름도 같은 형식으로
    kind_label = "추가매장 보험가입증명서" if kind == "추가" else "해지매장 변경증서(환급)"
    await _send_message_retrying(
        bot, ALLOWED_USER_ID,
        f"📄 처브 {kind_label} 도착: {brand} {store}\n{list_msg}\n{saved_msg}",
    )
    processed[digest] = f"{dt.date.today().isoformat()} {brand} {store}"
    _sj_save(SEJUNG_PROCESSED_FILE, processed)

    target = (req or {}).get("sender_email")
    if target and batch is not None:
        batch.append({"target": target, "who": (req or {}).get("sender_name") or target, "store": store,
                      "kind": kind, "brand": brand, "filename": filename, "pdf_bytes": pdf_bytes,
                      "refund_date": refund_date})
    elif target:
        await _sj_queue_to_sejung(bot, target, store, kind, brand, filename, pdf_bytes,
                                  (req or {}).get("sender_name") or target, refund_date)
    else:
        # 담당자를 모르면 유진님께 물어봄 - 답장(이름 또는 이메일)을 받으면 그때 발송 준비
        os.makedirs(SEJUNG_AWAIT_DIR, exist_ok=True)
        path = os.path.join(SEJUNG_AWAIT_DIR, f"{digest}.pdf")
        with open(path, "wb") as f:
            f.write(pdf_bytes)
        awaiting = _sj_load(SEJUNG_AWAIT_FILE, [])
        # 같은 매장·같은 종류가 이미 대기 중이면 새 것으로 바꿈(같은 매장이 두 번 쌓이던 문제, 2026-11-17)
        for w in [w for w in awaiting if w.get("brand") == brand and w.get("kind") == kind
                  and _sj_norm(w.get("store")) == _sj_norm(store)]:
            awaiting.remove(w)
            if w.get("path") != path:
                try:
                    os.remove(w.get("path") or "")
                except OSError:
                    pass
        awaiting.append({"store": store, "kind": kind, "brand": brand, "filename": filename, "path": path,
                         "refund_date": refund_date, "asked_at": dt.datetime.now().timestamp()})
        _sj_save(SEJUNG_AWAIT_FILE, awaiting)
        await _send_message_retrying(
            bot, ALLOWED_USER_ID,
            f"❓ '{store}' {kind_label}를 보낼 ㈜세정 담당자를 모르겠어요(신청서가 메일로 온 기록이 없어요). "
            "담당자 이름이나 이메일을 답장으로 보내주세요.",
        )
    return True


async def sejung_waiting_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/waiting - 담당자 답을 기다리는 ㈜세정 건 목록."""
    if not is_allowed(update):
        return
    cert_wait = _sj_load(SEJUNG_AWAIT_FILE, [])
    recv_wait = _sj_load(SEJUNG_AWAIT_RECEIVER_FILE, [])
    lines = [f"- [증명서 발송 담당자] {w['brand']} {w['store']}" for w in cert_wait]
    lines += [f"- [신청서 접수자] {w.get('brand') or ''} " + ", ".join(i["store_name"] for i in w.get("items", []))
              for w in recv_wait]
    await update.message.reply_text(
        ("⏳ 담당자 답을 기다리는 건:\n" + "\n".join(lines) +
         "\n\n필요 없는 건은 /cancelwait 매장명 으로 취소하세요.") if lines else "대기 중인 건이 없어요."
    )


async def sejung_cancel_wait_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/cancelwait <매장명> - 담당자 답 대기 건을 취소(직접 메일을 보냈을 때 등). 2026-11-16 요청."""
    if not is_allowed(update):
        return
    target = _sj_norm(" ".join(context.args or []))
    if not target:
        await update.message.reply_text("사용법: /cancelwait 매장명 (대기 목록은 /waiting)")
        return
    removed = []
    cert_wait = _sj_load(SEJUNG_AWAIT_FILE, [])
    keep = []
    for w in cert_wait:
        if target in _sj_norm(w.get("store")) or _sj_norm(w.get("store")) in target:
            removed.append(f"[증명서 발송] {w['store']}")
            try:
                os.remove(w["path"])
            except OSError:
                pass
        else:
            keep.append(w)
    _sj_save(SEJUNG_AWAIT_FILE, keep)
    recv_wait = _sj_load(SEJUNG_AWAIT_RECEIVER_FILE, [])
    keep2 = []
    for w in recv_wait:
        names = [i["store_name"] for i in w.get("items", [])]
        if any(target in _sj_norm(n) or _sj_norm(n) in target for n in names):
            removed.append(f"[신청서 접수자] {', '.join(names)}")
        else:
            keep2.append(w)
    _sj_save(SEJUNG_AWAIT_RECEIVER_FILE, keep2)
    await update.message.reply_text(
        f"🗑 대기 취소했어요: {', '.join(removed)}" if removed else "해당 매장의 대기 건을 찾지 못했어요. /waiting 으로 확인해주세요."
    )


_SJ_CANCEL_RE = re.compile(r"직접\s*보냈|이미\s*보냈|보냈어|보냈습니다|보냈음|안\s*보내|보내지\s*마|취소|필요\s*없")


def _sj_name_from_reply(text: str) -> str:
    """'세정 장혜진 님한테 메일 보내줘' 같은 문장에서 담당자 이름만 뽑음. 등록된 연락처 이름이
    문장 안에 있으면 그걸 우선 씀(2026-11-16: 문장 전체를 이름으로 착각하던 문제 수정)."""
    t = (text or "").strip()
    for name in sorted(_load_contacts().keys(), key=len, reverse=True):
        if name and name in t.replace(" ", ""):
            return name
    waiting_names = [w.get("store") for w in _sj_load(SEJUNG_AWAIT_FILE, [])] + [
        i.get("store_name") for w in _sj_load(SEJUNG_AWAIT_RECEIVER_FILE, []) for i in w.get("items", [])]
    for st in sorted([x for x in waiting_names if x], key=len, reverse=True):
        t = t.replace(st, " ")
    t = re.sub(r"(㈜|\(주\))?\s*세정", " ", t)
    t = re.sub(r"(담당자|대리|과장|차장|부장|팀장|주임|사원|대표|님|씨)", " ", t)
    t = re.sub(r"(한테|에게|께|으로|로)?\s*(이?메일|메일로)?\s*(보내\S*|전달\S*|부탁\S*)", " ", t)
    words = re.findall(r"[가-힣]{2,4}", t)
    return words[0] if words else (text or "").strip()


def _sj_waiting_stores(w: dict, key: str) -> list:
    return [w.get(key)] if isinstance(w.get(key), str) else [x.get("store_name") for x in w.get("items", [])]


def _sj_latest_waiting(waiting: list) -> int:
    """봇이 가장 마지막에 물어본 건의 위치(질문 시각이 없으면 목록의 맨 뒤)."""
    return max(range(len(waiting)), key=lambda i: (waiting[i].get("asked_at") or 0, i))


def _sj_pick_waiting(update: Update, waiting: list, key: str) -> int:
    """대기 목록 중 이번 답이 어느 건에 대한 것인지 고름: 유진님이 '답장'한 봇 메시지에 적힌 매장명
    -> 답 문장에 있는 매장명(일부만 적어도 됨) -> 봇이 가장 마지막에 물어본 건 순.
    (2026-11-17: 매장명 없이 이름만 답하면 가장 오래된 건으로 처리돼, 방금 물어본 OL이천직영점
    대신 예전에 남아 있던 롯데동부산아울렛점으로 발송 준비되던 문제 수정)"""
    reply_to = getattr(update.message, "reply_to_message", None)
    sources = [getattr(reply_to, "text", "") or "", getattr(update.message, "text", "") or ""]
    for src in sources:
        n = _sj_norm(src)
        for i, w in enumerate(waiting):
            if any(st and _sj_norm(st) in n for st in _sj_waiting_stores(w, key)):
                return i
    # '동부산아울렛점 직접 보냈어'처럼 매장명을 일부만 적은 경우: 딱 한 매장에만 맞을 때만 인정
    tokens = [_sj_norm(t) for t in re.findall(r"[0-9A-Za-z가-힣]+점", getattr(update.message, "text", "") or "")]
    tokens = [t for t in tokens if len(t) >= 3]
    hits = [i for i, w in enumerate(waiting)
            if any(st and t in _sj_norm(st) for st in _sj_waiting_stores(w, key) for t in tokens)]
    if hits and len({_sj_norm(str(_sj_waiting_stores(waiting[i], key))) for i in hits}) == 1:
        return hits[0]
    return _sj_latest_waiting(waiting)


async def _sj_handle_await_contact_reply(update: Update, user_text: str) -> bool:
    """담당자를 물어본 상태에서 유진님이 보낸 답(이름 또는 이메일)을 처리. 처리했으면 True."""
    awaiting = _sj_load(SEJUNG_AWAIT_FILE, [])
    if not awaiting:
        return False
    text = user_text.strip()
    if text.startswith("/"):
        return False
    reply_to = getattr(update.message, "reply_to_message", None)
    if reply_to is not None and "접수자" in (getattr(reply_to, "text", "") or ""):
        return False  # 신청서 접수자 질문에 대한 답이면 그쪽 처리로 넘김
    idx = _sj_pick_waiting(update, awaiting, "store")
    item = awaiting[idx]
    # '직접 보냈어', '안 보내도 돼', '취소' 같은 답이면 발송 대기를 취소(2026-11-16 유진님 요청)
    if _SJ_CANCEL_RE.search(text):
        awaiting.pop(idx)
        _sj_save(SEJUNG_AWAIT_FILE, awaiting)
        try:
            os.remove(item["path"])
        except OSError:
            pass
        await update.message.reply_text(f"🗑 '{item['store']}' 증명서 발송 대기를 취소했어요(직접 보내신 걸로 처리).")
        if awaiting:
            await update.message.reply_text(f"남은 대기 건: {', '.join(w['store'] for w in awaiting)}")
        return True
    email_match = re.search(r"[\w.+-]+@[\w-]+\.[\w.]+", text)
    target, who = None, text
    if email_match:
        target = email_match.group(0)
    else:
        who = _sj_name_from_reply(text)
        contacts = _load_contacts()
        target = contacts.get(_norm_contact_name(who))
        if target and _is_own_address(target):
            target = None
    if not target:
        await update.message.reply_text(
            f"'{who}' 담당자의 이메일이 등록돼 있지 않아요. 이메일 주소를 그대로 보내주시거나, "
            "/setcontact 이름 이메일 로 등록한 뒤 이름을 다시 보내주세요. "
            f"(대기 중인 건: {item['store']})"
        )
        return True
    try:
        with open(item["path"], "rb") as f:
            pdf_bytes = f.read()
    except OSError:
        awaiting.pop(idx)
        _sj_save(SEJUNG_AWAIT_FILE, awaiting)
        await update.message.reply_text("⚠️ 대기 중이던 증명서 파일을 찾지 못했어요. 처브 메일에서 직접 보내주세요.")
        return True
    awaiting.pop(idx)
    _sj_save(SEJUNG_AWAIT_FILE, awaiting)
    try:
        os.remove(item["path"])
    except OSError:
        pass
    await _sj_queue_to_sejung(update.get_bot(), target, item["store"], item["kind"], item["brand"],
                              item["filename"], pdf_bytes, who, item.get("refund_date"))
    if awaiting:
        # 다음에 실제로 처리될 건(가장 최근 건)을 묻고, 그 건을 '마지막으로 물어본 건'으로 표시
        nxt = awaiting[_sj_latest_waiting(awaiting)]
        nxt["asked_at"] = dt.datetime.now().timestamp()
        _sj_save(SEJUNG_AWAIT_FILE, awaiting)
        others = [w["store"] for w in awaiting if w is not nxt]
        await update.message.reply_text(
            f"다음으로 '{nxt['store']}' 담당자도 알려주세요. 직접 보내셨으면 '{nxt['store']} 직접 보냈어'라고 답해주세요."
            + (f"\n(그 밖의 대기 건: {', '.join(others)} - 매장명과 이름을 같이 보내주시면 그 건으로 처리해요)" if others else "")
        )
    return True


async def check_new_mail(context: ContextTypes.DEFAULT_TYPE) -> None:
    global last_uid_seen

    if not GMAIL_ADDRESS or not GMAIL_APP_PASSWORD:
        return

    try:
        imap = imaplib.IMAP4_SSL("imap.gmail.com")
        imap.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        imap.select("INBOX")

        status, data = imap.uid("search", None, "ALL")
        uids = data[0].split()
        if not uids:
            imap.logout()
            return

        latest_uid = int(uids[-1])

        if last_uid_seen is None:
            # 처음 실행될 때는 지금 시점만 기준으로 잡고, 과거 메일은 알리지 않음
            last_uid_seen = latest_uid
            imap.logout()
            return

        new_uids = [uid for uid in uids if int(uid) > last_uid_seen]

        for uid in new_uids:
            status, msg_data = imap.uid("fetch", uid, "(RFC822)")
            if not msg_data or msg_data[0] is None:
                continue
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            subject = _decode_mime_words(msg.get("Subject", "(제목 없음)"))
            sender = _decode_mime_words(msg.get("From", "(발신자 알 수 없음)"))
            body = _get_email_body(msg)[:2000]

            # 듀오링고 등 광고/구독성 메일은 '새 메일 도착' 알림을 안 보냄(2026-09-29 유진님
            # 요청). 등록된 정산 담당자 연락처로 온 메일은 절대 무시하지 않음. 알림만 건너뛰고
            # 첨부파일 처리(정산양식 동기화 등)는 그대로 진행함 - 혹시라도 무시 대상으로
            # 잘못 걸려도 실제 업무 처리 자체는 놓치지 않도록.
            muted = _is_muted_sender(sender, subject, msg.get("List-Unsubscribe", ""))
            if not muted:
                try:
                    response = client.messages.create(
                        model=MODEL_NAME,
                        max_tokens=300,
                        system="이메일 내용을 한국어로 3줄 이내로 간결하게 요약해줘. 핵심만 전달해.",
                        messages=[
                            {
                                "role": "user",
                                "content": f"보낸사람: {sender}\n제목: {subject}\n본문:\n{body}",
                            }
                        ],
                    )
                    summary = response.content[0].text
                except Exception:
                    logger.exception("메일 요약 중 오류")
                    summary = "(요약 생성 실패)"

                text = f"📬 새 메일 도착\n\n보낸사람: {sender}\n제목: {subject}\n\n요약:\n{summary}"
                await context.bot.send_message(chat_id=ALLOWED_USER_ID, text=text)

            # ㈜세정 x 처브 업무(2026-11-12~): 예전 정산양식/가입증명서 생성/폐점서류 로직은 끔.
            #  - 엑셀 첨부가 변경신청서면 -> 처브 4명에게 보낼지 확인 버튼
            #  - PDF 첨부가 처브 보험가입증명서/배서증명서면 -> 리스트 기재 + 드라이브 저장 + 세정 담당자 발송 준비
            sender_email = email.utils.parseaddr(sender)[1] or None
            sender_name = email.utils.parseaddr(sender)[0] or None
            # 1) 세정 담당자의 변경신청 메일부터 확인(엑셀/PDF 신청서 + 첨부서류 전부를 처브로)
            try:
                handled_request = await _sj_handle_request_mail(
                    context.bot, msg, subject, _get_email_body(msg), sender_email, sender_name
                )
            except Exception as e:
                handled_request = True
                logger.exception("변경신청 메일 처리 중 오류")
                await _send_message_retrying(
                    context.bot, ALLOWED_USER_ID,
                    f"⚠️ 변경신청 메일('{subject}') 처리 중 오류: {type(e).__name__}: {str(e)[:150]}",
                )
            # 2) 신청 메일이 아니면 처브 증명서(PDF)인지 확인
            if handled_request:
                continue
            pdf_atts = [(fn, b) for fn, b in _extract_non_xlsx_attachments(msg) if fn.lower().endswith(".pdf")]
            # 처브가 원계약 증명서 전체(수백 장)를 한 메일에 보내면 하나씩 '추가매장'으로 기재하면 안 됨
            # (2026-11-18 OL 332장 건). 증명서가 많으면 처리하지 않고 알림만 보냄.
            # 단, 우리가 보낸 신청 메일('정산 양식 송부')에 대한 처브 답장이면 개수와 상관없이 처리함
            # (2026-11-24 웰메이드 특판 9개 증명서가 '대량 메일'로 막혀 저장이 안 된 건)
            is_reply_to_request = "정산 양식 송부" in (subject or "").replace("  ", " ")
            if len(pdf_atts) > SEJUNG_BULK_CERT_LIMIT and not is_reply_to_request:
                await _send_message_retrying(
                    context.bot, ALLOWED_USER_ID,
                    f"📦 '{subject}' 메일에 PDF가 {len(pdf_atts)}개 있어요. 원계약 전체 증명서 같은 대량 메일이라 "
                    "통합관리리스트에 자동 기재하지 않았어요. 추가/해지 매장 건이면 그 PDF만 텔레그램으로 올려주세요.",
                )
                continue
            # 한 메일의 증명서들은 담당자별로 묶어서 메일 1통으로 보내도록 모아둠(2026-11-24)
            cert_batch: list = []
            for att_filename, att_bytes in pdf_atts:
                try:
                    await _sj_handle_cert_file(context.bot, att_filename, att_bytes, _get_email_body(msg),
                                               batch=cert_batch)
                except Exception as e:
                    logger.exception("처브 증명서 처리 중 오류")
                    await _send_message_retrying(
                        context.bot, ALLOWED_USER_ID,
                        f"⚠️ 메일 첨부 '{att_filename}' 처리 중 오류: {type(e).__name__}: {str(e)[:150]}",
                    )
            if cert_batch:
                try:
                    await _sj_flush_cert_batch(context.bot, cert_batch)
                except Exception as e:
                    logger.exception("증명서 묶음 발송 준비 중 오류")
                    await _send_message_retrying(
                        context.bot, ALLOWED_USER_ID,
                        f"⚠️ 증명서 묶음 발송 준비 중 오류: {type(e).__name__} ({_sj_err_where()}): {str(e)[:150]}",
                    )

        last_uid_seen = latest_uid
        imap.logout()
    except Exception:
        logger.exception("메일 확인 중 오류")


def main() -> None:
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("reset", reset))
    app.add_handler(CommandHandler("remind", remind))
    app.add_handler(CommandHandler("mail", mail_command))
    app.add_handler(CommandHandler("route", route_start))
    app.add_handler(CommandHandler("sethome", set_home))
    app.add_handler(CommandHandler("setwork", set_work))
    app.add_handler(CommandHandler("towork", commute_to_work))
    app.add_handler(CommandHandler("tohome", commute_to_home))
    app.add_handler(CommandHandler("setcontact", set_contact_command))
    app.add_handler(CommandHandler("contacts", list_contacts_command))
    app.add_handler(CommandHandler("mutesender", mute_sender_command))
    app.add_handler(CommandHandler("unmutesender", unmute_sender_command))
    app.add_handler(CommandHandler("mutedsenders", list_muted_senders_command))
    app.add_handler(CommandHandler("pending", list_pending_command))
    app.add_handler(CommandHandler("waiting", sejung_waiting_command))
    app.add_handler(CommandHandler("cancelwait", sejung_cancel_wait_command))
    app.add_handler(CallbackQueryHandler(handle_cert_confirmation))
    app.add_handler(MessageHandler(filters.LOCATION, handle_location))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    if GMAIL_ADDRESS and GMAIL_APP_PASSWORD:
        app.job_queue.run_repeating(
            check_new_mail, interval=MAIL_CHECK_INTERVAL, first=10
        )
        logger.info("이메일 확인 작업이 등록되었습니다 (%d초 간격).", MAIL_CHECK_INTERVAL)
    else:
        logger.info("GMAIL_ADDRESS/GMAIL_APP_PASSWORD가 없어 이메일 확인 기능은 꺼져 있습니다.")

    if GMAIL_ADDRESS and GMAIL_APP_PASSWORD:
        app.job_queue.run_repeating(
            check_livestock_claims, interval=CLAIMS_CHECK_INTERVAL, first=60
        )
        logger.info("소 사고접수 자동 기재 작업이 등록되었습니다 (%d초 간격).", CLAIMS_CHECK_INTERVAL)

    # 6일 주기 토큰 갱신 알림은 껐음(2026-10-21): 구글 앱을 '프로덕션'으로 게시해서 refresh
    # token이 더 이상 7일마다 만료되지 않음. 대신 매일 아침 드라이브 연결만 점검해서, 끊겼을
    # 때에만 알림을 보냄(check_gdrive_health).
    app.job_queue.run_daily(check_gdrive_health, time=dt.time(hour=8, minute=30, tzinfo=KST))
    logger.info("구글 드라이브 연결 점검 작업이 등록되었습니다(매일 08:30 KST, 끊겼을 때만 알림).")

    logger.info("봇을 시작합니다...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
