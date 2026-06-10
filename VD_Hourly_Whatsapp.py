#!/usr/bin/env python3

import os
import time
import io
import logging
import tempfile
import pytz
from datetime import datetime, timedelta
from typing import List
import json

import requests
from PIL import Image, ImageEnhance, ImageChops
from pdf2image import convert_from_bytes
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


# =========================
# ENV VARIABLES & SETTINGS
# =========================
SHEET_ID = os.getenv("SHEET_ID")
SHEET_NAME = "VD Top Batch Day View 1st April Onwards"
VD_REPORT_SHEET_NAME = "VD Report"
IST = pytz.timezone("Asia/Kolkata")
EVENT_START_DATE = datetime(2026, 6, 1, tzinfo=IST).date()

# =========================
# VD REPORT TIME-BASED RANGES
# =========================
# Between 6:00 AM and 9:00 AM IST → morning range
VD_REPORT_MORNING_RANGE = "X33:AB42"   # 6 AM – 9 AM IST
# All other times → default range
VD_REPORT_DEFAULT_RANGE = "X22:AC32"  # Outside 6–9 AM IST

# =========================
# SHEET RANGES
# =========================
DAY_RANGES = [
    [f"{SHEET_NAME}!A1281:F1303", f"{SHEET_NAME}!K1283:R1303"],     # Day 0
    [f"{SHEET_NAME}!A1304:F1326", f"{SHEET_NAME}!K1306:R1326"],     # Day 1
    [f"{SHEET_NAME}!A1327:F1349", f"{SHEET_NAME}!K1329:R1349"],     # Day 2
    [f"{SHEET_NAME}!A1350:F1372", f"{SHEET_NAME}!K1352:R1372"],     # Day 3
    [f"{SHEET_NAME}!A1373:F1395", f"{SHEET_NAME}!K1375:R1395"],     # Day 4
    [f"{SHEET_NAME}!A1396:F1418", f"{SHEET_NAME}!K1398:R1418"],     # Day 5
    [f"{SHEET_NAME}!A1419:F1441", f"{SHEET_NAME}!K1421:R1441"],     # Day 6
    [f"{SHEET_NAME}!A1442:F1464", f"{SHEET_NAME}!K1444:R1464"],     # Day 7
    [f"{SHEET_NAME}!A1465:F1487", f"{SHEET_NAME}!K1467:R1487"],     # Day 8
    [f"{SHEET_NAME}!A1488:F1510", f"{SHEET_NAME}!K1490:R1510"],     # Day 9
    [f"{SHEET_NAME}!A1511:F1533", f"{SHEET_NAME}!K1513:R1533"],     # Day 10
    [f"{SHEET_NAME}!A1534:F1556", f"{SHEET_NAME}!K1536:R1556"],     # Day 11
    [f"{SHEET_NAME}!A1557:F1579", f"{SHEET_NAME}!K1559:R1579"],     # Day 12
    [f"{SHEET_NAME}!A1580:F1602", f"{SHEET_NAME}!K1582:R1602"],     # Day 13
    [f"{SHEET_NAME}!A1603:F1625", f"{SHEET_NAME}!K1605:R1625"],     # Day 14
    [f"{SHEET_NAME}!A1626:F1648", f"{SHEET_NAME}!K1628:R1648"],     # Day 15
    [f"{SHEET_NAME}!A1649:F1671", f"{SHEET_NAME}!K1651:R1671"],     # Day 16
    [f"{SHEET_NAME}!A1672:F1694", f"{SHEET_NAME}!K1674:R1694"],     # Day 17
    [f"{SHEET_NAME}!A1695:F1717", f"{SHEET_NAME}!K1697:R1717"],     # Day 18
    [f"{SHEET_NAME}!A1718:F1740", f"{SHEET_NAME}!K1720:R1740"],   # Day 19
    [f"{SHEET_NAME}!A1741:F1763", f"{SHEET_NAME}!K1743:R1763"], # Day 20
    [f"{SHEET_NAME}!A1764:F1786", f"{SHEET_NAME}!K1766:R1786"], # Day 21
    [f"{SHEET_NAME}!A1787:F1809", f"{SHEET_NAME}!K1789:R1809"], # Day 22
    [f"{SHEET_NAME}!A1810:F1832", f"{SHEET_NAME}!K1812:R1832"], # Day 23
    [f"{SHEET_NAME}!A1833:F1855", f"{SHEET_NAME}!K1835:R1855"], # Day 24
    [f"{SHEET_NAME}!A1856:F1878", f"{SHEET_NAME}!K1858:R1878"], # Day 25
    [f"{SHEET_NAME}!A1879:F1901", f"{SHEET_NAME}!K1881:R1901"], # Day 26
    [f"{SHEET_NAME}!A1902:F1924", f"{SHEET_NAME}!K1904:R1924"], # Day 27
    [f"{SHEET_NAME}!A1925:F1947", f"{SHEET_NAME}!K1927:R1947"], # Day 28
    [f"{SHEET_NAME}!A1948:F1970", f"{SHEET_NAME}!K1950:R1970"], # Day 29
    [f"{SHEET_NAME}!A1971:F1993", f"{SHEET_NAME}!K1973:R1993"], # Day 30
]


def get_current_ranges():
    now_ist = datetime.now(IST)
    
    # Rollover logic: Before 5:00 AM IST, still consider it the previous reporting day
    cutoff_today = now_ist.replace(hour=5, minute=0, second=0, microsecond=0)
    if now_ist < cutoff_today:
        effective_date = (now_ist - timedelta(days=1)).date()
    else:
        effective_date = now_ist.date()
        
    day_diff = (effective_date - EVENT_START_DATE).days
    # The list has 31 items (index 0 to 30)
    day_index = min(max(0, day_diff), 30)
    
    logger.info("Reporting Day Index: %s (Effective Date: %s)", day_index, effective_date)
    return DAY_RANGES[day_index]


# =========================
# CLOUDINARY
# =========================
CLOUD_NAME = os.getenv("CLOUD_NAME")
UPLOAD_PRESET = os.getenv("UPLOAD_PRESET")
UPLOAD_URL = f"https://api.cloudinary.com/v1_1/{CLOUD_NAME}/image/upload"

# =========================
# AISENSY
# =========================
AISENSY_API_KEY = os.getenv("AISENSY_API_KEY")
CAMPAIGN_NAME = os.getenv("AISENSY_CAMPAIGN_NAME")
DESTINATIONS = [
    d.strip() for d in os.getenv("DESTINATIONS", "").split(",") if d.strip()
]

# =========================
# IMAGE SETTINGS
# =========================
TARGET_SIZE_BYTES = 4 * 1024 * 1024
JPEG_QUALITIES = [95, 85, 75, 65, 55]

# =========================
# LOGGING
# =========================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)
logger = logging.getLogger("bizcat")


# =========================
# HELPERS
# =========================
def refresh_creds(creds: Credentials):
    if not creds.valid:
        creds.refresh(Request())
        logger.info("google token refreshed")


def get_sheet_gid(creds: Credentials, sheet_name: str) -> str:
    service = build("sheets", "v4", credentials=creds)
    meta = service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()

    for sheet in meta["sheets"]:
        props = sheet["properties"]
        if props["title"] == sheet_name:
            return str(props["sheetId"])

    raise RuntimeError(f"sheet {sheet_name} not found")


def jpg_bytes(img: Image.Image, quality: int) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=quality, optimize=True, progressive=True)
    return buf.getvalue()


def optimize_image(img: Image.Image) -> bytes:
    if img.mode != "RGB":
        img = img.convert("RGB")

    # Try quality reduction
    for q in JPEG_QUALITIES:
        data = jpg_bytes(img, q)
        logger.info("jpeg quality %s size %.2f MB", q, len(data) / 1024 / 1024)
        if len(data) <= TARGET_SIZE_BYTES:
            return data

    # Resize fallback
    w, h = img.size
    for _ in range(3):
        w = int(w * 0.96)
        h = int(h * 0.96)
        img = img.resize((w, h), Image.LANCZOS)

        data = jpg_bytes(img, 65)
        if len(data) <= TARGET_SIZE_BYTES:
            return data

    return data


def crop_white_space(img: Image.Image) -> Image.Image:
    bg = Image.new(img.mode, img.size, img.getpixel((0, 0)))
    diff = ImageChops.difference(img, bg)
    diff = ImageEnhance.Contrast(diff).enhance(3.0)
    bbox = diff.getbbox()
    return img.crop(bbox) if bbox else img


# =========================
# MAIN LOGIC
# =========================
def get_vd_report_range(now_ist: datetime) -> str:
    """Return the VD Report sheet range based on current IST time.

    6:00 AM <= time < 9:00 AM  →  morning range (X33:AB42)
    Any other time             →  default range  (X22:AC32)
    """
    hour = now_ist.hour
    if 6 <= hour < 9:
        logger.info("VD Report: morning window (06:00–09:00 IST) → %s", VD_REPORT_MORNING_RANGE)
        return VD_REPORT_MORNING_RANGE
    else:
        logger.info("VD Report: outside morning window → %s", VD_REPORT_DEFAULT_RANGE)
        return VD_REPORT_DEFAULT_RANGE


def export_and_upload_images() -> List[str]:
    creds_info = json.loads(os.getenv("GOOGLE_CREDENTIALS_JSON"))

    creds = Credentials.from_service_account_info(
        creds_info,
        scopes=[
            "https://www.googleapis.com/auth/drive.readonly",
            "https://www.googleapis.com/auth/spreadsheets.readonly",
        ],
    )

    refresh_creds(creds)
    
    now_ist = datetime.now(IST)
    now_mins = now_ist.hour * 60 + now_ist.minute
    
    # Execution Window: 8:30 AM to 1:30 AM (t+1)
    # If between 1:31 AM and 8:29 AM, skip.
    if 1 * 60 + 30 < now_mins < 8 * 60 + 30:
        logger.info("Current time %s is outside the scheduled window (08:30 - 01:30). Skipping.", now_ist.strftime("%H:%M"))
        return []

    # Get current day ranges (existing sheet)
    day_ranges = get_current_ranges()
    sheet_gid = get_sheet_gid(creds, SHEET_NAME)
    
    tasks = []
    for r in day_ranges:
        tasks.append((SHEET_NAME, sheet_gid, r))

    # ── VD Report: time-based range ──────────────────────────────────────
    vd_report_range_str = get_vd_report_range(now_ist)
    vd_report_full_range = f"{VD_REPORT_SHEET_NAME}!{vd_report_range_str}"
    vd_report_gid = get_sheet_gid(creds, VD_REPORT_SHEET_NAME)
    tasks.append((VD_REPORT_SHEET_NAME, vd_report_gid, vd_report_full_range))
    logger.info("Added VD Report range: %s", vd_report_full_range)

    uploaded_urls = []

    for i, (s_name, gid, s_range) in enumerate(tasks, start=1):
        range_only = s_range.split("!")[1]

        export_url = (
            f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export"
            f"?format=pdf&portrait=false&gid={gid}&range={range_only}"
            f"&size=A2&scale=5&top_margin=0.25&bottom_margin=0.25"
            f"&left_margin=0.25&right_margin=0.25&fzr=false"
            f"&gridlines=false&printtitle=false"
        )

        logger.info("Exporting task %s: %s (%s)", i, s_range, s_name)

        response = requests.get(
            export_url,
            headers={"Authorization": f"Bearer {creds.token}"},
            timeout=90,
        )
        response.raise_for_status()

        pages = convert_from_bytes(response.content, dpi=300, first_page=1, last_page=1)
        img = pages[0].convert("RGB")
        img = ImageEnhance.Sharpness(img).enhance(2.0)
        img = crop_white_space(img)

        jpg_data = optimize_image(img)

        with tempfile.NamedTemporaryFile(delete=False, suffix=f"_table_{i}.jpg") as tmp:
            tmp.write(jpg_data)
            filename = tmp.name

        try:
            with open(filename, "rb") as f:
                upload = requests.post(
                    UPLOAD_URL,
                    files={"file": f},
                    data={
                        "upload_preset": UPLOAD_PRESET,
                        "folder": f"BizCat_Exports/{now_ist.strftime('%Y-%m-%d')}",
                    },
                    timeout=60,
                )
                upload.raise_for_status()

            url = upload.json().get("secure_url")
            if url:
                uploaded_urls.append(url)
                logger.info("Uploaded %s: %s", s_range, url)

        finally:
            os.remove(filename)

        time.sleep(2)

    return uploaded_urls


def send_via_aisensy(urls: List[str]):
    if not urls:
        logger.warning("no images generated")
        return

    for dest in DESTINATIONS:
        for i, url in enumerate(urls, start=1):
            payload = {
                "apiKey": AISENSY_API_KEY,
                "campaignName": CAMPAIGN_NAME,
                "destination": dest,
                "userName": "PW Online- Analytics",
                "templateParams": [datetime.now(IST).strftime("%d %B %Y")],
                "source": "automation-script",
                "media": {
                    "url": url,
                    "filename": f"table_{i}.jpg"
                },
            }

            r = requests.post(
                "https://backend.aisensy.com/campaign/t1/api",
                json=payload,
                timeout=30,
            )

            logger.info("sent to %s image %s status %s", dest, i, r.status_code)
            time.sleep(5)


# =========================
# ENTRY POINT
# =========================
if __name__ == "__main__":
    required = [
        "SHEET_ID",
        "CLOUD_NAME",
        "UPLOAD_PRESET",
        "AISENSY_API_KEY",
        "DESTINATIONS",
    ]

    missing = [v for v in required if not os.getenv(v)]
    if missing:
        raise EnvironmentError(f"missing secrets: {', '.join(missing)}")

    Image.MAX_IMAGE_PIXELS = 300_000_000

    logger.info("Automation run started (IST time: %s)", datetime.now(IST).strftime("%Y-%m-%d %H:%M"))
    
    try:
        urls = export_and_upload_images()
        if urls:
            send_via_aisensy(urls)
            logger.info("Automation run completed successfully")
        else:
            logger.info("No images to send. Automation run finished.")
    except Exception as e:
        logger.error("Error during automation run: %s", e, exc_info=True)
        import sys
        sys.exit(1)
