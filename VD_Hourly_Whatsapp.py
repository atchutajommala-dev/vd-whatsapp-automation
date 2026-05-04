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
# ENV VARIABLES
# =========================
SHEET_ID = os.getenv("SHEET_ID")
REPORT_SHEET_NAME = "VD Report"

IST = pytz.timezone("Asia/Kolkata")
SCHEDULE_SLOTS = ["08:30", "11:30", "15:30", "18:30", "22:30", "00:30"]

# =========================
# SHEET RANGES (REMOVED STANDARD DAY VIEWS)
# =========================
# Standard day-based ranges have been removed as per request.
# Only VD Report ranges are now active.



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

# TODAY string will be generated inside the loop

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
    
    # GID for Report sheet
    sheet_report_gid = get_sheet_gid(creds, REPORT_SHEET_NAME)

    logger.info("Report GID: %s", sheet_report_gid)
    
    # 2. Bracket-based Algorithm (Handles delays automatically)
    now_ist = datetime.now(IST)
    now_mins = now_ist.hour * 60 + now_ist.minute
    
    # Define timing points in minutes desde midnight
    M0830 = 8 * 60 + 30  # 510
    M1130 = 11 * 60 + 30 # 690
    M0030 = 0 * 60 + 30  # 30
    
    extra_range = None
    bracket_label = ""
    
    if M0830 <= now_mins < M1130:
        # Bracket: 8:30 AM to 11:30 AM
        extra_range = f"{REPORT_SHEET_NAME}!X33:AA41"
        bracket_label = "08:30 AM Report"
    elif (now_mins >= M1130) or (now_mins < M0030):
        # Bracket: 11:30 AM until 00:30 AM (Wraps around midnight)
        # Covers 11:30, 15:30, 18:30, and 22:30 slots
        extra_range = f"{REPORT_SHEET_NAME}!X22:AF31"
        bracket_label = "Midday/Evening Report"
    else:
        # Bracket: 00:30 AM to 08:29 AM
        extra_range = None
        bracket_label = "00:30 AM (Standard Only)"
            
    # Prepare final list of tasks: [(sheet_name, gid, range_string), ...]
    tasks = []

    # Add the extra range if identified by the bracket algorithm
    if extra_range:
        tasks.append((REPORT_SHEET_NAME, sheet_report_gid, extra_range))
    else:
        logger.info("No report range scheduled for this time bracket (%s)", bracket_label)

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
        send_via_aisensy(urls)
        logger.info("Automation run completed successfully")
    except Exception as e:
        logger.error("Error during automation run: %s", e, exc_info=True)
        # We might want to exit with non-zero if it's a cron job for monitoring
        import sys
        sys.exit(1)
