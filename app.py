"""
Data Entry Automation Dashboard & Multi-Step Portal Automation
=============================================================
Streamlit UI + Playwright engine for:
1. Multi-Step Portal Workflow (Techkriti'27 Campus Ambassador & similar multi-step portals)
2. Generic Single-Page Form Automation (CSV column to field mapping)

Run: streamlit run app.py
"""

from __future__ import annotations

import asyncio
import csv
import io
import os
import random
import re
import subprocess
import sys
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any

import openpyxl
import pandas as pd
import streamlit as st
from playwright.async_api import Browser, BrowserContext, Page, async_playwright

# ---------------------------------------------------------------------------
# Cloud Deployment & Environment Helpers
# ---------------------------------------------------------------------------
@st.cache_resource
def ensure_playwright_installed() -> None:
    """Ensure Playwright's Chromium binary is installed once per container boot."""
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            browser.close()
    except Exception:
        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=False)


ensure_playwright_installed()


def resolve_headless_setting(requested_headless: bool) -> bool:
    """Automatically force headless=True if running on Linux without a display server."""
    if sys.platform.startswith("linux") and not os.environ.get("DISPLAY"):
        return True
    return requested_headless


# ---------------------------------------------------------------------------
# Constants & Defaults
# ---------------------------------------------------------------------------
DEFAULT_PORTAL_URL = "https://ca.techkriti.org/?ref=33e9cb9c541e9fa7"
DEFAULT_REFERRER_TOKEN = "CA-2027-5293"
DEFAULT_OUTPUT_CSV = "completed_records.csv"

ANANTNAG_ADDRESSES = [
    "Govt Degree College Road, Khanabal, Anantnag, J&K - 192101",
    "Khanabal, Anantnag, Jammu & Kashmir - 192101",
    "KP Road, Anantnag, Jammu & Kashmir - 192101",
    "Lal Chowk, Anantnag, J&K - 192101",
]

GENDER_OPTIONS = ["Male", "Female"]
YEAR_OF_STUDY_OPTIONS = ["1st Year", "2nd Year", "3rd Year"]
PREFERRED_CONTACT_OPTIONS = ["Email", "LinkedIn", "Other"]

# ---------------------------------------------------------------------------
# Page configuration & global styling
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Automation Hub | Multi-Step Portal & Form Engine",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    .main-header { font-size: 2.1rem; font-weight: 800; margin-bottom: 0.2rem; color: #f8fafc; }
    .sub-header { color: #94a3b8; margin-bottom: 1.5rem; font-size: 1rem; }
    .log-box {
        background: #090d16;
        color: #38bdf8;
        font-family: 'SFMono-Regular', Consolas, 'Liberation Mono', Menlo, monospace;
        font-size: 0.84rem;
        padding: 1.1rem;
        border-radius: 10px;
        max-height: 420px;
        overflow-y: auto;
        white-space: pre-wrap;
        border: 1px solid #1e293b;
        box-shadow: inset 0 2px 8px rgba(0,0,0,0.4);
    }
    .badge-card {
        background: #0f172a;
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 0.85rem 1.1rem;
        margin-bottom: 0.8rem;
    }
    .success-banner {
        background: linear-gradient(135deg, #059669 0%, #10b981 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 12px;
        text-align: center;
        font-size: 1.3rem;
        font-weight: 700;
        margin: 1rem 0;
        box-shadow: 0 4px 20px rgba(16, 185, 129, 0.35);
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Data Models
# ---------------------------------------------------------------------------
@dataclass
class FieldMapping:
    """Maps one CSV column to a generic form field."""

    csv_column: str
    selector: str = ""
    label_name: str = ""
    is_optional: bool = False


@dataclass
class GenericAutomationState:
    """State for generic form automation tab."""

    logs: list[str] = field(default_factory=list)
    running: bool = False
    completed: bool = False
    current_row: int = 0
    total_rows: int = 0
    wait_seconds_remaining: int = 0
    wait_total_seconds: int = 0
    last_error: str | None = None
    lock: threading.Lock = field(default_factory=threading.Lock)

    def append_log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        with self.lock:
            self.logs.append(f"[{timestamp}] {message}")

    def get_logs(self) -> str:
        with self.lock:
            return "\n".join(self.logs)

    def set_wait(self, remaining: int, total: int) -> None:
        with self.lock:
            self.wait_seconds_remaining = remaining
            self.wait_total_seconds = total


@dataclass
class PortalAutomationState:
    """State for the dedicated multi-step portal automation."""

    logs: list[str] = field(default_factory=list)
    running: bool = False
    stop_requested: bool = False
    completed: bool = False
    current_row_idx: int = 0
    start_row_num: int = 1
    end_row_num: int = 0
    total_rows_to_process: int = 0
    wait_seconds_remaining: int = 0
    wait_total_seconds: int = 0
    last_error: str | None = None
    results: list[dict[str, Any]] = field(default_factory=list)
    excel_bytes: bytes | None = None
    lock: threading.Lock = field(default_factory=threading.Lock)

    def append_log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        with self.lock:
            self.logs.append(f"[{timestamp}] {message}")

    def get_logs(self) -> str:
        with self.lock:
            return "\n".join(self.logs)

    def set_wait(self, remaining: int, total: int) -> None:
        with self.lock:
            self.wait_seconds_remaining = remaining
            self.wait_total_seconds = total

    def add_result(self, record: dict[str, Any]) -> None:
        with self.lock:
            self.results.append(record)

    def get_results(self) -> list[dict[str, Any]]:
        with self.lock:
            return list(self.results)

    def set_excel_bytes(self, b: bytes | None) -> None:
        with self.lock:
            self.excel_bytes = b

    def get_excel_bytes(self) -> bytes | None:
        with self.lock:
            return self.excel_bytes

    def reset_for_run(self, start_row: int, end_row: int, total_to_process: int) -> None:
        with self.lock:
            self.logs.clear()
            self.running = True
            self.stop_requested = False
            self.completed = False
            self.current_row_idx = 0
            self.start_row_num = start_row
            self.end_row_num = end_row
            self.total_rows_to_process = total_to_process
            self.wait_seconds_remaining = 0
            self.wait_total_seconds = 0
            self.last_error = None


# ---------------------------------------------------------------------------
# Session State Initialization
# ---------------------------------------------------------------------------
def init_session_state() -> None:
    defaults: dict[str, Any] = {
        # Generic form state
        "df": None,
        "field_mappings": [],
        "automation": GenericAutomationState(),
        "form_url": "",
        "headless": True,
        "submit_selector": "",
        "preferred_contact_label": "Preferred contact",
        # Portal workflow state
        "portal_df": None,
        "portal_state": PortalAutomationState(),
        "portal_target_url": DEFAULT_PORTAL_URL,
        "portal_referrer_token": DEFAULT_REFERRER_TOKEN,
        "portal_headless": False,  # Default unchecked so user can watch live
        "portal_delay_secs": 3,
        "portal_excel_raw": None,
        "portal_excel_filename": "Techkriti_CA_Registration_Tracker(1).xlsx",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


init_session_state()


# ---------------------------------------------------------------------------
# Utility Helpers
# ---------------------------------------------------------------------------
def clean_contact_number(val: Any) -> str:
    """Extract a clean 10-digit mobile string from any raw cell value."""
    if pd.isna(val):
        return ""
    raw = str(val).strip()
    if raw.endswith(".0"):
        raw = raw[:-2]
    # Filter only digits
    digits = "".join(c for c in raw if c.isdigit())
    # If 12 digits starting with 91, extract 10 digits
    if len(digits) == 12 and digits.startswith("91"):
        return digits[2:]
    return digits


def detect_portal_columns(columns: list[str]) -> dict[str, str | None]:
    """
    Flexibly match the 4 required student columns case-insensitively:
    1. Student Full Name
    2. Email Address
    3. Contact Number
    4. College / Institution
    """
    detected: dict[str, str | None] = {
        "name": None,
        "email": None,
        "contact": None,
        "college": None,
    }
    col_clean = {c: str(c).strip().lower().replace("_", " ").replace("-", " ") for c in columns}

    # Name matching
    for c, norm in col_clean.items():
        if norm in ["student full name", "full name", "student name", "name", "candidate name"]:
            detected["name"] = c
            break
    if not detected["name"]:
        for c, norm in col_clean.items():
            if "name" in norm and "college" not in norm and "inst" not in norm:
                detected["name"] = c
                break

    # Email matching
    for c, norm in col_clean.items():
        if norm in ["email address", "email id", "email", "e-mail", "mail"]:
            detected["email"] = c
            break
    if not detected["email"]:
        for c, norm in col_clean.items():
            if "email" in norm or "mail" in norm:
                detected["email"] = c
                break

    # Contact / Mobile matching
    for c, norm in col_clean.items():
        if norm in [
            "contact number",
            "contact",
            "mobile number",
            "mobile",
            "phone number",
            "phone",
            "whatsapp number",
        ]:
            detected["contact"] = c
            break
    if not detected["contact"]:
        for c, norm in col_clean.items():
            if any(k in norm for k in ["contact", "mobile", "phone"]):
                detected["contact"] = c
                break

    # College / Institution matching
    for c, norm in col_clean.items():
        if any(
            norm == k
            for k in [
                "college / institution",
                "college/institution",
                "college",
                "institution",
                "college name",
                "university",
                "institute",
            ]
        ):
            detected["college"] = c
            break
    if not detected["college"]:
        for c, norm in col_clean.items():
            if any(k in norm for k in ["college", "institution", "university", "institute", "campus"]):
                detected["college"] = c
                break

    return detected


def append_to_completed_csv(filepath: str, record: dict[str, Any]) -> None:
    """Append a completed row safely to CSV with header if missing."""
    file_exists = os.path.exists(filepath) and os.path.getsize(filepath) > 0
    fieldnames = [
        "S.No.",
        "Student Full Name",
        "Email Address",
        "Contact Number",
        "College / Institution",
        "Generated Link",
        "Status",
    ]
    with open(filepath, mode="a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        if not file_exists:
            writer.writeheader()
        writer.writerow(record)


def write_to_openpyxl_workbook(
    wb: openpyxl.Workbook,
    target_row_num: int,
    s_no: int,
    student_name: str,
    email: str,
    contact: str,
    college: str,
    generated_link: str,
) -> None:
    """
    Write one student record starting from Row 9 in Columns A through F:
      A: S.No (starting from 2 at Row 9)
      B: Student Full Name
      C: Email Address
      D: Contact Number
      E: College / Institution
      F: Generated Link
    Preserves Rows 1–8 and Column G completely untouched.
    """
    ws = wb.active
    ws.cell(row=target_row_num, column=1, value=s_no)
    ws.cell(row=target_row_num, column=2, value=student_name)
    ws.cell(row=target_row_num, column=3, value=email)
    ws.cell(row=target_row_num, column=4, value=contact)
    ws.cell(row=target_row_num, column=5, value=college)
    ws.cell(row=target_row_num, column=6, value=generated_link)


# ---------------------------------------------------------------------------
# Multi-Step Portal Playwright Execution
# ---------------------------------------------------------------------------
async def _react_fill(page: Page, locator_str: str, value: str) -> None:
    """
    Fill a React-controlled input/textarea correctly:
    Uses Playwright fill() (which triggers proper input + change events in React)
    followed by explicit dispatchEvent('input') and 'change' to ensure React state updates.
    """
    loc = page.locator(locator_str).first
    await loc.wait_for(state="visible", timeout=8_000)
    await loc.click()
    await loc.fill(value)
    await loc.dispatch_event("input")
    await loc.dispatch_event("change")


async def _react_select(page: Page, locator_str: str, value: str) -> None:
    """
    Select an option from a React-controlled <select> and dispatch change event.
    """
    loc = page.locator(locator_str).first
    await loc.wait_for(state="visible", timeout=8_000)
    await loc.select_option(value)
    await loc.dispatch_event("change")


async def process_portal_student_session(
    browser: Browser,
    student_data: dict[str, str],
    target_url: str,
    referrer_token: str,
    state: PortalAutomationState,
    row_display_idx: int,
    total_in_batch: int,
) -> tuple[str, str]:
    """
    Executes an isolated session for a single student row on the portal:
    1. Creates brand-new browser context (isolated cookies/storage).
    2. Navigates directly to the target referral URL (e.g. ?ref=33e9cb9c541e9fa7).
    3. Step 1: Clicks JOIN NOW / BECOME AN AMBASSADOR, selects manual email entry,
               fills initial Name/Email, clicks Proceed to Registration Form.
    4. Step 2: Fills the profile details form with React-aware fill + dispatchEvent
               (Name, Email, Mobile, WhatsApp, Anantnag Address, Gender, College,
               City, Year of Study, Referrer CA Id = CA-2027-5293).
    5. Step 3: Scrolls to and clicks "Submit & Open CA Dashboard →", retries if
               the Register form is still visible after 2s, then verifies Dashboard loaded.
    6. Step 4: Mandatory 3-second visual pause on Dashboard for confirmation.
    7. Step 5: Scrapes generated referral link / CA ID from Dashboard.
    8. Step 6: Clicks Logout and closes context.
    """
    student_name = student_data["name"]
    student_email = student_data["email"]
    contact_number = student_data["contact"]
    college_name = student_data["college"]

    state.append_log(
        f"👤 [{row_display_idx}/{total_in_batch}] Starting session for: {student_name} ({student_email})"
    )

    # 1. Fresh isolated context per student — no shared cookies or localStorage
    context: BrowserContext = await browser.new_context(
        viewport={"width": 1366, "height": 768},
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
    )
    page: Page = await context.new_page()

    unique_link = ""
    status = "Failed"

    try:
        # ----------------------------------------------------------------
        # 2. Navigate to target referral URL (always the full ?ref= link)
        # ----------------------------------------------------------------
        state.append_log(f"🌐 Navigating to: {target_url}")
        await page.goto(target_url, wait_until="domcontentloaded", timeout=60_000)
        await page.wait_for_timeout(1500)

        # ----------------------------------------------------------------
        # 3. Step 1: Land on portal and open manual registration
        # ----------------------------------------------------------------
        # Check if the full profile registration form is already open
        is_profile_open = await page.locator(
            'input[name="mobile"], input[placeholder*="phone number" i], textarea[name="postalAddress"]'
        ).count() > 0

        if not is_profile_open:
            state.append_log("🔘 Clicking CTA button (JOIN NOW / BECOME AN AMBASSADOR)...")
            cta_clicked = False
            cta_selectors = [
                'button:has-text("JOIN NOW")',
                'a:has-text("JOIN NOW")',
                'button:has-text("BECOME AN AMBASSADOR NOW")',
                'a:has-text("BECOME AN AMBASSADOR NOW")',
                'button:has-text("REGISTER NOW")',
                'a:has-text("REGISTER NOW")',
                'button:has-text("Become an Ambassador")',
                'button:has-text("Register")',
            ]
            for sel in cta_selectors:
                loc = page.locator(sel).first
                if await loc.count() > 0 and await loc.is_visible():
                    await loc.click()
                    cta_clicked = True
                    break
            if not cta_clicked:
                loc = page.get_by_role(
                    "button", name=re.compile(r"join now|register now|become an ambassador", re.I)
                ).first
                if await loc.count() > 0:
                    await loc.click()

            await page.wait_for_timeout(1000)

            # Click the "Or Register with Email Directly" manual toggle
            manual_selectors = [
                'button:has-text("Or Register with Email Directly")',
                'button:has-text("Register with Email Directly")',
                'button:has-text("Register with Email")',
                'button:has-text("Email Directly")',
                'button:has-text("Manual")',
                'a:has-text("Manual")',
            ]
            for m_sel in manual_selectors:
                m_loc = page.locator(m_sel).first
                if await m_loc.count() > 0 and await m_loc.is_visible():
                    state.append_log("📧 Clicked: manual email registration toggle.")
                    await m_loc.click()
                    await page.wait_for_timeout(600)
                    break

            # Fill Name in the initial auth modal
            for n_sel in [
                'input[placeholder="Full Name"]',
                'input[placeholder*="Full Name" i]',
            ]:
                n_loc = page.locator(n_sel).first
                if await n_loc.count() > 0 and await n_loc.is_visible():
                    await n_loc.fill(student_name)
                    await n_loc.dispatch_event("input")
                    await n_loc.dispatch_event("change")
                    break

            # Fill Email in the initial auth modal
            for e_sel in [
                'input[placeholder="Email Address*"]',
                'input[placeholder*="Email" i]',
                'input[type="email"]',
                'input[name="email"]',
            ]:
                e_loc = page.locator(e_sel).first
                if await e_loc.count() > 0 and await e_loc.is_visible():
                    await e_loc.fill(student_email)
                    await e_loc.dispatch_event("input")
                    await e_loc.dispatch_event("change")
                    break

            # Click "Proceed to Registration Form →"
            for p_sel in [
                'button:has-text("Proceed to Registration Form")',
                'button:has-text("Proceed")',
                'button:has-text("Continue")',
                'form button[type="submit"]',
            ]:
                p_loc = page.locator(p_sel).first
                if await p_loc.count() > 0 and await p_loc.is_visible():
                    state.append_log("➡️ Clicking: Proceed to Registration Form...")
                    await p_loc.click()
                    await page.wait_for_timeout(1500)
                    break

        # ----------------------------------------------------------------
        # 4. Step 2: Complete the Profile Details Form (React-aware)
        # ----------------------------------------------------------------
        state.append_log("📝 Waiting for the full registration form to appear...")
        await page.locator(
            'input[name="mobile"], input[placeholder*="phone number" i], textarea[name="postalAddress"]'
        ).first.wait_for(state="visible", timeout=18_000)

        # 4a. Name* — fill only if the field is empty / editable
        name_field = page.locator('input[name="name"]').first
        if await name_field.count() > 0:
            try:
                if not (await name_field.input_value()).strip():
                    await name_field.fill(student_name)
                    await name_field.dispatch_event("input")
                    await name_field.dispatch_event("change")
            except Exception:
                pass

        # 4b. Email Id* — fill only if empty / editable
        email_field = page.locator('input[name="email"], input[type="email"]').first
        if await email_field.count() > 0:
            try:
                if not (await email_field.input_value()).strip():
                    await email_field.fill(student_email)
                    await email_field.dispatch_event("input")
                    await email_field.dispatch_event("change")
            except Exception:
                pass

        # 4c. Mobile*
        await _react_fill(page, 'input[name="mobile"], input[placeholder*="phone number" i]', contact_number)
        state.append_log(f"  ✓ Mobile: {contact_number}")

        # 4d. WhatsApp Number* — same as Contact Number
        await _react_fill(page, 'input[name="whatsapp"], input[placeholder*="WhatsApp number" i]', contact_number)
        state.append_log(f"  ✓ WhatsApp: {contact_number}")

        # 4e. Postal Address* — randomly picked Anantnag local address
        chosen_address = random.choice(ANANTNAG_ADDRESSES)
        await _react_fill(
            page,
            'textarea[name="postalAddress"], textarea[placeholder*="postal address" i]',
            chosen_address,
        )
        state.append_log(f"  ✓ Address: {chosen_address}")

        # 4f. Gender* — random Male/Female via React-aware select
        chosen_gender = random.choice(GENDER_OPTIONS)
        gender_sel_str = 'select[name="gender"]'
        if await page.locator(gender_sel_str).count() > 0:
            await _react_select(page, gender_sel_str, chosen_gender)
        else:
            g_lbl = page.get_by_label(re.compile(r"gender", re.I)).first
            if await g_lbl.count() > 0:
                await g_lbl.select_option(chosen_gender)
                await g_lbl.dispatch_event("change")
        state.append_log(f"  ✓ Gender: {chosen_gender}")

        # 4g. College Name*
        await _react_fill(page, 'input[name="collegeName"], input[placeholder*="college name" i]', college_name)
        state.append_log(f"  ✓ College: {college_name}")

        # 4h. College City* — always "Anantnag"
        await _react_fill(page, 'input[name="collegeCity"], input[placeholder*="college city" i]', "Anantnag")
        state.append_log("  ✓ City: Anantnag")

        # 4i. Year of Study* — random 1st/2nd/3rd Year via React-aware select
        chosen_year = random.choice(YEAR_OF_STUDY_OPTIONS)
        year_sel_str = 'select[name="yearOfStudy"]'
        if await page.locator(year_sel_str).count() > 0:
            await _react_select(page, year_sel_str, chosen_year)
        else:
            y_lbl = page.get_by_label(re.compile(r"year of study", re.I)).first
            if await y_lbl.count() > 0:
                await y_lbl.select_option(chosen_year)
                await y_lbl.dispatch_event("change")
        state.append_log(f"  ✓ Year of Study: {chosen_year}")

        # 4j. Referrer CA Id / Token — ALWAYS clear and fill with CA-2027-5293
        #     The form field expects a CA ID like "CA-2027-5293", NOT the full URL.
        #     Force-override: if referrer_token is empty or accidentally contains a URL, reset it.
        ca_id_to_fill = referrer_token.strip()
        if not ca_id_to_fill or "http" in ca_id_to_fill or "ref=" in ca_id_to_fill:
            ca_id_to_fill = "CA-2027-5293"

        ref_field = page.locator(
            'input[placeholder*="CA-2027" i], '
            'input[placeholder*="referral token" i], '
            'input[name*="referrer" i], '
            'input[name*="referral" i]'
        ).first

        # Fallback: find the label "Referrer CA Id" and locate the input inside its parent
        if await ref_field.count() == 0:
            try:
                ref_field = page.locator(
                    'label:has-text("Referrer CA Id") ~ input, '
                    'label:has-text("Referrer CA Id") + input'
                ).first
            except Exception:
                pass

        if await ref_field.count() > 0:
            try:
                await ref_field.wait_for(state="visible", timeout=5_000)
                await ref_field.scroll_into_view_if_needed()
                await ref_field.click()
                await page.keyboard.press("Control+A")
                await page.keyboard.press("Backspace")
                await ref_field.fill("")
                await ref_field.fill(ca_id_to_fill)
                await ref_field.dispatch_event("input")
                await ref_field.dispatch_event("change")
                await page.wait_for_timeout(1000)      # pause so CA-2027-5293 is visibly shown
                # Confirm what was actually written — log it for debugging
                actual_val = await ref_field.input_value()
                state.append_log(f"  ✓ Referrer CA Id field value confirmed: '{actual_val}'")
            except Exception as ref_err:
                state.append_log(f"  ⚠️ Referrer field warning: {ref_err}")
        else:
            state.append_log("  ⚠️ Referrer CA Id input not found on page — skipping.")

        state.append_log(
            f"📋 All fields filled — Address='{chosen_address}', Gender={chosen_gender}, "
            f"Year={chosen_year}, City=Anantnag, Referrer={referrer_token.strip()}"
        )

        # ----------------------------------------------------------------
        # 5. Step 3: Scroll to and click "Submit & Open CA Dashboard →"
        # ----------------------------------------------------------------
        state.append_log("🚀 Locating and clicking the Submit button...")
        submit_btn = page.locator(
            'button:has-text("Submit & Open CA Dashboard"), '
            'button:has-text("Submit & Open CA Dashboard →"), '
            'button:has-text("Saving Profile"), '
            'button[type="submit"]'
        ).first
        await submit_btn.wait_for(state="visible", timeout=10_000)
        await submit_btn.scroll_into_view_if_needed()
        await page.wait_for_timeout(400)
        await submit_btn.click()
        state.append_log("  → Submit button clicked. Waiting 2s to check outcome...")

        # Wait 2 seconds and check if the Register form is STILL visible
        await page.wait_for_timeout(2000)
        register_form_still_up = await page.locator(
            'input[name="mobile"], textarea[name="postalAddress"]'
        ).count() > 0

        if register_form_still_up:
            # Retry submit — dispatchEvent('click') as extra guarantee
            state.append_log("  ⚠️ Form still visible after 2s — retrying submit via dispatchEvent...")
            try:
                await submit_btn.dispatch_event("click")
            except Exception:
                # Last resort: find any visible submit button and click it
                for btn_sel in [
                    'button:has-text("Submit & Open CA Dashboard")',
                    'button[type="submit"]',
                ]:
                    btn = page.locator(btn_sel).last
                    if await btn.count() > 0 and await btn.is_visible():
                        await btn.click()
                        break
            await page.wait_for_timeout(2500)

        # ----------------------------------------------------------------
        # 6. Step 4: Verify CA Dashboard actually loaded
        # ----------------------------------------------------------------
        state.append_log("🔍 Verifying CA Dashboard has loaded...")
        dashboard_loaded = False

        # Dashboard indicators: "Welcome back", "Points Earned", "Balance: 10 Pts",
        # "Logout" button, or any element containing "CA-2027-"
        dashboard_checks = [
            'button:has-text("Logout")',
            'button:has-text("Log out")',
            '*:has-text("Points Earned")',
            '*:has-text("Balance:")',
            '*:has-text("Welcome back")',
            '*:has-text("CA Dashboard")',
            '*:has-text("My Dashboard")',
        ]
        for _ in range(15):  # poll up to ~15 seconds
            for dash_sel in dashboard_checks:
                if await page.locator(dash_sel).count() > 0:
                    dashboard_loaded = True
                    break
            if dashboard_loaded:
                break
            # Also check that the Register form is gone
            reg_form_gone = await page.locator(
                'input[name="mobile"], textarea[name="postalAddress"]'
            ).count() == 0
            if reg_form_gone:
                # Form gone but no dashboard text yet — give a moment
                await page.wait_for_timeout(500)
                for dash_sel in dashboard_checks:
                    if await page.locator(dash_sel).count() > 0:
                        dashboard_loaded = True
                        break
            if dashboard_loaded:
                break
            await page.wait_for_timeout(1000)

        if not dashboard_loaded:
            # Final check by page content
            content = await page.content()
            if any(kw in content for kw in ["Points Earned", "Balance:", "Welcome back", "CA Dashboard", "Logout"]):
                dashboard_loaded = True

        if not dashboard_loaded:
            # Mark as Failed with reason
            state.append_log(
                f"❌ Dashboard did NOT load for {student_name}. "
                "Registration may have failed (validation error or duplicate email)."
            )
            status = "Failed"
            unique_link = "Failed: Dashboard never loaded after submission"
        else:
            state.append_log(f"✅ CA Dashboard loaded successfully for {student_name}!")

            # ----------------------------------------------------------------
            # 7. MANDATORY 3-second visual pause on Dashboard
            # ----------------------------------------------------------------
            state.append_log("⏳ Pausing 3 seconds on Dashboard for visual confirmation...")
            await page.wait_for_timeout(3000)

            # ----------------------------------------------------------------
            # 8. Step 5: Scrape the newly generated referral link / CA ID
            # ----------------------------------------------------------------
            state.append_log("🔎 Scraping generated referral link and CA ID from Dashboard...")
            ca_id_found = ""

            # Check all inputs on the dashboard page
            for inp in await page.locator("input").all():
                try:
                    val = (await inp.input_value()).strip()
                    if "?ref=" in val and "techkriti" in val:
                        unique_link = val
                        break
                    elif re.match(r"CA-2027-\d+", val):
                        ca_id_found = val
                except Exception:
                    pass

            # Search page content
            if not unique_link:
                content = await page.content()
                m = re.search(
                    r"https?://[a-zA-Z0-9.-]*techkriti\.org/\?ref=[a-zA-Z0-9_-]+",
                    content,
                )
                if m:
                    unique_link = m.group(0)

            if not ca_id_found:
                content = await page.content()
                m_ca = re.search(r"\bCA-2027-\d+\b", content)
                if m_ca:
                    ca_id_found = m_ca.group(0)

            # Search text nodes displaying "?ref="
            if not unique_link:
                for el in await page.locator('*:has-text("?ref=")').all():
                    try:
                        txt = (await el.text_content() or "").strip()
                        m_txt = re.search(
                            r"https?://[^\s\"']+\?ref=[a-zA-Z0-9_-]+", txt
                        )
                        if m_txt:
                            unique_link = m_txt.group(0)
                            break
                    except Exception:
                        pass

            # Construct link from CA ID if direct link wasn't found
            if not unique_link and ca_id_found:
                unique_link = f"https://ca.techkriti.org/?ref={ca_id_found}"

            if unique_link:
                status = "Success"
                state.append_log(f"🎉 Generated Referral Link: {unique_link}")
            else:
                # Dashboard loaded but we couldn't scrape the link yet
                unique_link = ca_id_found or "Dashboard loaded — link not extracted"
                status = "Success"
                state.append_log(
                    f"✅ Registration success for {student_name} (link not scraped, CA ID: {ca_id_found or 'pending'})."
                )

        # ----------------------------------------------------------------
        # 9. Step 6: Logout and close context
        # ----------------------------------------------------------------
        try:
            logout_btn = page.locator(
                'button:has-text("Logout"), button:has-text("Log out")'
            ).first
            if await logout_btn.count() > 0 and await logout_btn.is_visible():
                state.append_log("🚪 Clicking Logout...")
                await logout_btn.click(timeout=4000)
                await page.wait_for_timeout(800)
        except Exception as logout_err:
            state.append_log(f"  ℹ️ Logout skipped: {logout_err}")

    except Exception as exc:
        state.append_log(f"❌ Unhandled error for {student_name}: {exc}")
        unique_link = f"Failed: {exc}"
        status = "Failed"
    finally:
        await context.close()

    return unique_link, status


async def run_portal_automation_async(
    df: pd.DataFrame,
    start_row: int,
    end_row: int,
    col_mapping: dict[str, str],
    target_url: str,
    referrer_token: str,
    delay_secs: int,
    headless: bool,
    excel_raw_bytes: bytes | None,
    state: PortalAutomationState,
) -> None:
    """Orchestrates multi-row portal automation with isolated browser sessions."""
    total_to_process = (end_row - start_row) + 1
    state.reset_for_run(start_row, end_row, total_to_process)
    state.append_log(
        f"🚀 Starting Multi-Step Portal Engine. Processing rows {start_row} to {end_row} ({total_to_process} total)."
    )

    # Initialize openpyxl workbook if template was provided
    wb: openpyxl.Workbook | None = None
    if excel_raw_bytes:
        try:
            wb = openpyxl.load_workbook(io.BytesIO(excel_raw_bytes))
            state.append_log("📊 Loaded existing Excel tracker template.")
        except Exception as exc:
            state.append_log(f"⚠️ Could not load Excel template: {exc}. Creating fresh workbook.")
            wb = openpyxl.Workbook()
    else:
        # Create fresh workbook for live Excel updates
        wb = openpyxl.Workbook()
        ws = wb.active
        # Setup headers in row 8
        ws.cell(row=8, column=1, value="S.No.")
        ws.cell(row=8, column=2, value="Student Full Name")
        ws.cell(row=8, column=3, value="Email Address")
        ws.cell(row=8, column=4, value="Contact Number")
        ws.cell(row=8, column=5, value="College / Institution")
        ws.cell(row=8, column=6, value="Generated Link")

    # Ensure headless mode on headless Linux (Streamlit Community Cloud)
    effective_headless = resolve_headless_setting(headless)
    if effective_headless and not headless:
        state.append_log("ℹ️ No DISPLAY server detected on Linux. Forcing headless=True for cloud stability.")

    async with async_playwright() as playwright:
        browser = await playwright.chromium.launch(
            headless=effective_headless,
            args=["--disable-blink-features=AutomationControlled"],
        )

        try:
            for loop_idx, csv_row_idx in enumerate(range(start_row, end_row + 1)):
                if state.stop_requested:
                    state.append_log("⏹️ User clicked Stop. Halting automation gracefully.")
                    break

                state.current_row_idx = loop_idx + 1
                row_data = df.iloc[csv_row_idx - 1]

                # Extract and clean values
                name_val = str(row_data[col_mapping["name"]]).strip() if pd.notna(row_data[col_mapping["name"]]) else ""
                email_val = str(row_data[col_mapping["email"]]).strip() if pd.notna(row_data[col_mapping["email"]]) else ""
                contact_val = clean_contact_number(row_data[col_mapping["contact"]])
                college_val = str(row_data[col_mapping["college"]]).strip() if pd.notna(row_data[col_mapping["college"]]) else ""

                student_dict = {
                    "name": name_val,
                    "email": email_val,
                    "contact": contact_val,
                    "college": college_val,
                }

                # S.No: row 9 corresponds to S.No 2 per tracker template specs
                s_no_val = 2 + (csv_row_idx - 1)
                excel_row_num = 9 + (csv_row_idx - 1)

                # Execute registration
                unique_link, status = await process_portal_student_session(
                    browser=browser,
                    student_data=student_dict,
                    target_url=target_url,
                    referrer_token=referrer_token,
                    state=state,
                    row_display_idx=state.current_row_idx,
                    total_in_batch=total_to_process,
                )

                # Record result
                result_record = {
                    "S.No.": s_no_val,
                    "Student Full Name": name_val,
                    "Email Address": email_val,
                    "Contact Number": contact_val,
                    "College / Institution": college_val,
                    "Generated Link": unique_link,
                    "Status": status,
                }
                state.add_result(result_record)

                # Append live to CSV
                append_to_completed_csv(DEFAULT_OUTPUT_CSV, result_record)

                # Update Excel tracker
                if wb is not None:
                    try:
                        write_to_openpyxl_workbook(
                            wb=wb,
                            target_row_num=excel_row_num,
                            s_no=s_no_val,
                            student_name=name_val,
                            email=email_val,
                            contact=contact_val,
                            college=college_val,
                            generated_link=unique_link,
                        )
                        out_stream = io.BytesIO()
                        wb.save(out_stream)
                        state.set_excel_bytes(out_stream.getvalue())
                    except Exception as xl_err:
                        state.append_log(f"⚠️ Excel write warning: {xl_err}")

                # Inter-row delay
                if loop_idx < total_to_process - 1 and not state.stop_requested:
                    state.append_log(f"⏸️ Waiting {delay_secs}s before next student submission...")
                    for rem in range(delay_secs, 0, -1):
                        if state.stop_requested:
                            break
                        state.set_wait(rem, delay_secs)
                        await asyncio.sleep(1)
                    state.set_wait(0, 0)

            state.append_log("🎉 Multi-Step Portal batch processing finished!")
            state.completed = True
        finally:
            await browser.close()
            state.running = False


def portal_thread_target(
    df: pd.DataFrame,
    start_row: int,
    end_row: int,
    col_mapping: dict[str, str],
    target_url: str,
    referrer_token: str,
    delay_secs: int,
    headless: bool,
    excel_raw_bytes: bytes | None,
    state: PortalAutomationState,
) -> None:
    """Background worker thread runner."""
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    try:
        loop.run_until_complete(
            run_portal_automation_async(
                df=df,
                start_row=start_row,
                end_row=end_row,
                col_mapping=col_mapping,
                target_url=target_url,
                referrer_token=referrer_token,
                delay_secs=delay_secs,
                headless=headless,
                excel_raw_bytes=excel_raw_bytes,
                state=state,
            )
        )
    except Exception as exc:
        state.append_log(f"💥 Portal worker crashed: {exc}")
        state.last_error = str(exc)
        state.running = False
    finally:
        loop.close()


def start_portal_automation(
    df: pd.DataFrame,
    start_row: int,
    end_row: int,
    col_mapping: dict[str, str],
    target_url: str,
    referrer_token: str,
    delay_secs: int,
    headless: bool,
    excel_raw_bytes: bytes | None,
) -> None:
    """Spawns background thread for the portal workflow."""
    state: PortalAutomationState = st.session_state.portal_state
    if state.running:
        st.warning("Automation is already in progress.")
        return

    thread = threading.Thread(
        target=portal_thread_target,
        args=(
            df.copy(),
            start_row,
            end_row,
            col_mapping,
            target_url,
            referrer_token,
            delay_secs,
            headless,
            excel_raw_bytes,
            state,
        ),
        daemon=True,
    )
    thread.start()
    st.session_state["_portal_thread"] = thread


# ---------------------------------------------------------------------------
# Generic Single-Page Form Automation (Preserved)
# ---------------------------------------------------------------------------
async def fill_field(page: Page, mapping: FieldMapping, value: str) -> bool:
    raw = str(value).strip() if pd.notna(value) else ""
    if not raw:
        return mapping.is_optional

    locator = None
    if mapping.selector.strip():
        locator = page.locator(mapping.selector.strip()).first
    elif mapping.label_name.strip():
        locator = page.get_by_label(mapping.label_name.strip(), exact=False).first
    else:
        return mapping.is_optional

    try:
        await locator.wait_for(state="visible", timeout=10_000)
        tag = await locator.evaluate("el => el.tagName.toLowerCase()")
        if tag == "select":
            await locator.select_option(label=raw)
        elif tag == "input":
            input_type = await locator.get_attribute("type") or "text"
            if input_type in ("checkbox", "radio"):
                await locator.check()
            else:
                await locator.fill(raw)
        elif tag == "textarea":
            await locator.fill(raw)
        else:
            await locator.fill(raw)
        return True
    except Exception:
        if mapping.is_optional:
            return True
        raise


async def select_preferred_contact(page: Page, label: str) -> str:
    choice = random.choice(PREFERRED_CONTACT_OPTIONS)
    try:
        radio = page.get_by_label(label, exact=False).get_by_role("radio", name=choice)
        await radio.check(timeout=5_000)
    except Exception:
        await page.get_by_role("radio", name=choice, exact=False).first.check(timeout=5_000)
    return choice


async def submit_generic_form(page: Page, submit_selector: str) -> None:
    if submit_selector.strip():
        await page.locator(submit_selector.strip()).first.click()
    else:
        for candidate in [
            'input[type="submit"]',
            'button[type="submit"]',
            'button:has-text("Submit")',
            'input[value="Submit"]',
        ]:
            loc = page.locator(candidate).first
            if await loc.count() > 0:
                await loc.click()
                return
        raise RuntimeError("Could not find submit button. Configure submit CSS selector.")


async def process_generic_single_row(
    page: Page,
    form_url: str,
    row: pd.Series,
    mappings: list[FieldMapping],
    submit_selector: str,
    preferred_contact_label: str,
) -> None:
    await page.goto(form_url, wait_until="domcontentloaded", timeout=60_000)
    await page.wait_for_timeout(500)
    for mapping in mappings:
        if mapping.csv_column not in row.index:
            continue
        value = row[mapping.csv_column]
        await fill_field(page, mapping, value)

    await select_preferred_contact(page, preferred_contact_label)
    await submit_generic_form(page, submit_selector)
    await page.wait_for_timeout(2_000)


async def run_generic_automation_async(
    df: pd.DataFrame,
    form_url: str,
    mappings: list[FieldMapping],
    state: GenericAutomationState,
    headless: bool,
    submit_selector: str,
    preferred_contact_label: str,
) -> None:
    state.total_rows = len(df)
    state.current_row = 0
    state.completed = False
    state.last_error = None

    effective_headless = resolve_headless_setting(headless)
    if effective_headless and not headless:
        state.append_log("ℹ️ No DISPLAY server detected on Linux. Forcing headless=True for cloud stability.")

    async with async_playwright() as playwright:
        browser = await playwright.chromium.launch(headless=effective_headless)
        context = await browser.new_context(
            viewport={"width": 1280, "height": 720},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0 Safari/537.36",
        )
        page = await context.new_page()

        try:
            for idx, (_, row) in enumerate(df.iterrows()):
                state.current_row = idx + 1
                name = str(row.get("Full Name", row.get("Name", f"Row {idx+1}")))
                state.append_log(f"⏳ Starting submission for: {name}...")
                try:
                    await process_generic_single_row(
                        page,
                        form_url,
                        row,
                        mappings,
                        submit_selector,
                        preferred_contact_label,
                    )
                    state.append_log(f"✅ Successfully submitted for {name}.")
                except Exception as exc:
                    state.append_log(f"❌ Failed for {name}: {exc}")
                    state.last_error = str(exc)

                if idx < len(df) - 1:
                    wait_secs = random.randint(15, 30)
                    state.append_log(f"⏸️ Waiting {wait_secs}s before next submission...")
                    for remaining in range(wait_secs, 0, -1):
                        state.set_wait(remaining, wait_secs)
                        await asyncio.sleep(1)
                    state.set_wait(0, 0)

            state.append_log("🎉 Generic form submissions complete!")
            state.completed = True
        finally:
            await browser.close()
            state.running = False


def generic_thread_target(
    df: pd.DataFrame,
    form_url: str,
    mappings: list[FieldMapping],
    state: GenericAutomationState,
    headless: bool,
    submit_selector: str,
    preferred_contact_label: str,
) -> None:
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    try:
        loop.run_until_complete(
            run_generic_automation_async(
                df,
                form_url,
                mappings,
                state,
                headless,
                submit_selector,
                preferred_contact_label,
            )
        )
    except Exception as exc:
        state.append_log(f"💥 Generic automation crashed: {exc}")
        state.last_error = str(exc)
        state.running = False
    finally:
        loop.close()


def start_generic_automation(
    df: pd.DataFrame,
    form_url: str,
    mappings: list[FieldMapping],
    headless: bool,
    submit_selector: str,
    preferred_contact_label: str,
) -> None:
    state: GenericAutomationState = st.session_state.automation
    if state.running:
        st.warning("Automation is already running.")
        return

    state.logs.clear()
    state.running = True
    state.completed = False
    state.wait_seconds_remaining = 0
    state.append_log("🚀 Generic automation engine started.")

    thread = threading.Thread(
        target=generic_thread_target,
        args=(
            df.copy(),
            form_url,
            mappings,
            state,
            headless,
            submit_selector,
            preferred_contact_label,
        ),
        daemon=True,
    )
    thread.start()
    st.session_state["_generic_thread"] = thread


# ---------------------------------------------------------------------------
# UI: Multi-Step Portal Workflow Tab
# ---------------------------------------------------------------------------
def render_portal_workflow() -> None:
    state: PortalAutomationState = st.session_state.portal_state

    st.markdown("### 🏛️ Techkriti'27 Campus Ambassador Multi-Step Workflow")
    st.caption(
        "Automate isolated student registrations, form auto-fill with Anantnag demographics, "
        "and real-time extraction of unique referral links into CSV & Excel tracker."
    )

    # 1. Configuration Cards
    col_c1, col_c2 = st.columns([1.2, 1])
    with col_c1:
        target_url = st.text_input(
            "Target Referral / Portal URL",
            value=st.session_state.portal_target_url,
            help="The referral URL opened for each student in a clean browser context.",
        )
        st.session_state.portal_target_url = target_url

    with col_c2:
        ref_token = st.text_input(
            "Referrer CA Id / Token",
            value="CA-2027-5293",
            key="referrer_ca_id_token_fixed",
            help="The CA ID filled into the 'Referrer CA Id / Token (Optional)' field on the registration form. Must be a CA ID like CA-2027-5293, NOT the portal URL.",
        )
        # Strict guard: if the value somehow contains a URL, force-reset it
        if not ref_token.strip() or "http" in ref_token or "ref=" in ref_token:
            ref_token = "CA-2027-5293"
        st.session_state.portal_referrer_token = ref_token

    st.divider()

    # 2. File Uploads (CSV & Excel)
    st.subheader("Step 1 — Upload Data & Tracker")
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        csv_file = st.file_uploader(
            "Upload Students CSV (4 columns: Name, Email, Contact, College)",
            type=["csv"],
            key="portal_csv_uploader",
        )
        if csv_file is not None:
            try:
                st.session_state.portal_df = pd.read_csv(csv_file)
                st.success(
                    f"Loaded **{len(st.session_state.portal_df)}** rows from `{csv_file.name}`"
                )
            except Exception as e:
                st.error(f"Error parsing CSV: {e}")

    with col_u2:
        excel_file = st.file_uploader(
            "Optional: Upload Existing Excel Tracker (`.xlsx`) to update from Row 9",
            type=["xlsx"],
            key="portal_excel_uploader",
        )
        if excel_file is not None:
            st.session_state.portal_excel_raw = excel_file.getvalue()
            st.session_state.portal_excel_filename = excel_file.name
            st.success(f"Attached tracker `{excel_file.name}` (Updates will start at Row 9).")

    df = st.session_state.portal_df
    if df is None:
        st.info("👆 Please upload your student CSV file above to proceed.")
        return

    # 3. Column Detection & Mapping
    st.subheader("Step 2 — Column Mapping & Auto-Fill Settings")
    detected = detect_portal_columns(list(df.columns))
    cols_list = list(df.columns)

    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    with col_m1:
        name_idx = cols_list.index(detected["name"]) if detected["name"] in cols_list else 0
        name_col = st.selectbox("Student Full Name Column", options=cols_list, index=name_idx)

    with col_m2:
        email_idx = cols_list.index(detected["email"]) if detected["email"] in cols_list else 0
        email_col = st.selectbox("Email Address Column", options=cols_list, index=email_idx)

    with col_m3:
        contact_idx = cols_list.index(detected["contact"]) if detected["contact"] in cols_list else 0
        contact_col = st.selectbox("Contact / Mobile Column", options=cols_list, index=contact_idx)

    with col_m4:
        college_idx = cols_list.index(detected["college"]) if detected["college"] in cols_list else 0
        college_col = st.selectbox("College / Institution Column", options=cols_list, index=college_idx)

    active_mapping = {
        "name": name_col,
        "email": email_col,
        "contact": contact_col,
        "college": college_col,
    }

    with st.expander("ℹ️ Demographics & Generated Default Field Values", expanded=False):
        st.markdown(
            """
            - **WhatsApp Number**: Exactly mirrors `Contact Number` from CSV
            - **Postal Address**: Randomly assigned from verified local Anantnag postal routes:
              - *Govt Degree College Road, Khanabal, Anantnag, J&K - 192101*
              - *Khanabal, Anantnag, Jammu & Kashmir - 192101*
              - *KP Road, Anantnag, Jammu & Kashmir - 192101*
              - *Lal Chowk, Anantnag, J&K - 192101*
            - **Gender**: Randomly chosen between `Male` and `Female`
            - **College City**: Always set to `Anantnag`
            - **Year of Study**: Randomly selected from `1st Year`, `2nd Year`, or `3rd Year`
            - **Referrer Token**: CA ID `CA-2027-5293` filled into the Referrer CA Id / Token field
            """
        )

    st.divider()

    # 4. Controls & Execution Parameters
    st.subheader("Step 3 — Execution Controls")
    col_ctrl1, col_ctrl2, col_ctrl3, col_ctrl4 = st.columns([1, 1, 1, 1.2])

    with col_ctrl1:
        headless_mode = st.checkbox(
            "Headless browser",
            value=st.session_state.portal_headless,
            help="Uncheck to open a visible Chrome window and watch form filling live.",
        )
        st.session_state.portal_headless = headless_mode

    with col_ctrl2:
        delay_val = st.slider(
            "Delay between rows (sec)",
            min_value=1,
            max_value=10,
            value=st.session_state.portal_delay_secs,
        )
        st.session_state.portal_delay_secs = delay_val

    with col_ctrl3:
        start_row_input = st.number_input(
            "Start Row (1-based)",
            min_value=1,
            max_value=len(df),
            value=1,
            help="Resume from any row number if stopped midway.",
        )

    with col_ctrl4:
        end_row_input = st.number_input(
            "End Row (inclusive)",
            min_value=int(start_row_input),
            max_value=len(df),
            value=len(df),
            help="Process up to this row number in the CSV.",
        )

    # Action Buttons
    col_btn1, col_btn2, _ = st.columns([1.2, 1.2, 3])
    with col_btn1:
        if st.button(
            "▶️ Start Portal Automation",
            type="primary",
            disabled=state.running,
            use_container_width=True,
        ):
            start_portal_automation(
                df=df,
                start_row=int(start_row_input),
                end_row=int(end_row_input),
                col_mapping=active_mapping,
                target_url=target_url.strip(),
                referrer_token=ref_token.strip(),
                delay_secs=int(delay_val),
                headless=headless_mode,
                excel_raw_bytes=st.session_state.portal_excel_raw,
            )
            st.rerun()

    with col_btn2:
        if st.button("⏹️ Stop Automation", disabled=not state.running, use_container_width=True):
            state.stop_requested = True
            st.warning("Stop requested. Finishing current student session...")

    # 5. Live Progress & Status
    if state.running:
        total = max(state.total_rows_to_process, 1)
        curr = state.current_row_idx
        st.progress(curr / total, text=f"Processing Student {curr} of {total} (Rows {state.start_row_num} to {state.end_row_num})")

    if state.running and state.wait_seconds_remaining > 0:
        st.info(
            f"⏳ **Cooldown between submissions:** {state.wait_seconds_remaining}s remaining before next session..."
        )

    # 6. Live Log Display
    st.markdown("**Live Activity Log**")
    st.markdown(
        f'<div class="log-box">{state.get_logs() or "Engine ready. Click \'Start Portal Automation\' to begin."}</div>',
        unsafe_allow_html=True,
    )

    # 7. Success Banner
    if state.completed and not state.running:
        st.markdown(
            """
            <div class="success-banner">
                🎉 MULTI-STEP PORTAL AUTOMATION COMPLETE! 🎉<br>
                <span style="font-size:1rem;font-weight:normal;">
                    All requested student registrations have been submitted and logged.
                </span>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # 8. Live Results Table & File Downloads
    st.subheader("Step 4 — Live Output & Downloads")
    results = state.get_results()
    if results:
        results_df = pd.DataFrame(results)
        st.dataframe(results_df, use_container_width=True)

    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        if os.path.exists(DEFAULT_OUTPUT_CSV) and os.path.getsize(DEFAULT_OUTPUT_CSV) > 0:
            with open(DEFAULT_OUTPUT_CSV, "rb") as f_csv:
                st.download_button(
                    label="📥 Download completed_records.csv",
                    data=f_csv.read(),
                    file_name="completed_records.csv",
                    mime="text/csv",
                    use_container_width=True,
                )
        else:
            st.button("📥 Download completed_records.csv", disabled=True, use_container_width=True)

    with col_dl2:
        excel_data = state.get_excel_bytes()
        if excel_data:
            st.download_button(
                label=f"📥 Download Updated Excel Tracker ({st.session_state.portal_excel_filename})",
                data=excel_data,
                file_name=st.session_state.portal_excel_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.button("📥 Download Updated Excel Tracker", disabled=True, use_container_width=True)

    # Auto-rerun loop while background worker is running
    if state.running:
        time.sleep(1)
        st.rerun()


# ---------------------------------------------------------------------------
# UI: Generic Single-Page Form Workflow Tab
# ---------------------------------------------------------------------------
def render_generic_workflow() -> None:
    st.markdown("### ⚙️ Generic Form Automation Dashboard")
    st.caption("Upload any CSV, map arbitrary form fields/selectors, and submit sequentially.")

    with st.sidebar:
        st.header("Generic Form Settings")
        submit_sel = st.text_input(
            "Submit button CSS selector (optional)",
            value=st.session_state.submit_selector,
            placeholder='button[type="submit"]',
        )
        st.session_state.submit_selector = submit_sel

        pref_label = st.text_input(
            "Preferred contact field label",
            value=st.session_state.preferred_contact_label,
            help="Label text for the Email / LinkedIn / Twitter radio group.",
        )
        st.session_state.preferred_contact_label = pref_label

        st.divider()
        st.markdown("**Preferred contact randomizer**")
        st.write(f"Each submission randomly picks: {', '.join(PREFERRED_CONTACT_OPTIONS)}")

    # Step 1: Upload
    col1, col2 = st.columns([2, 1])
    with col1:
        form_url = st.text_input(
            "Target Form URL",
            value=st.session_state.form_url,
            placeholder="https://example.com/register",
            help="Full URL of the form page to automate.",
            key="generic_form_url",
        )
        st.session_state.form_url = form_url

    with col2:
        headless = st.checkbox("Headless browser", value=st.session_state.headless, key="generic_headless")
        st.session_state.headless = headless

    uploaded = st.file_uploader("Upload CSV file", type=["csv"], key="generic_csv_uploader")
    if uploaded is None:
        st.info("Upload a `.csv` file to configure column mappings.")
        return

    try:
        df = pd.read_csv(uploaded)
    except Exception as exc:
        st.error(f"Could not parse CSV: {exc}")
        return

    st.session_state.df = df
    st.success(f"Loaded **{len(df)}** rows × **{len(df.columns)}** columns.")
    st.dataframe(df.head(5), use_container_width=True)

    # Step 2: Mapping
    st.subheader("Step 2 — Column Mapping")
    columns = list(df.columns)
    if len(st.session_state.field_mappings) != len(columns):
        st.session_state.field_mappings = [FieldMapping(csv_column=col) for col in columns]

    mappings: list[FieldMapping] = []
    header_cols = st.columns([1.2, 1.5, 1.5, 0.6])
    header_cols[0].markdown("**CSV Column**")
    header_cols[1].markdown("**CSS Selector**")
    header_cols[2].markdown("**Label Name**")
    header_cols[3].markdown("**Optional?**")

    for i, col in enumerate(columns):
        existing = st.session_state.field_mappings[i]
        row_cols = st.columns([1.2, 1.5, 1.5, 0.6])
        row_cols[0].markdown(f"`{col}`")
        selector = row_cols[1].text_input(
            f"sel_{i}",
            value=existing.selector,
            label_visibility="collapsed",
            placeholder="#field-id",
            key=f"map_sel_{i}",
        )
        label = row_cols[2].text_input(
            f"lbl_{i}",
            value=existing.label_name,
            label_visibility="collapsed",
            placeholder="Field label on form",
            key=f"map_lbl_{i}",
        )
        optional = row_cols[3].checkbox(
            "opt",
            value=existing.is_optional,
            label_visibility="collapsed",
            key=f"map_opt_{i}",
        )
        mappings.append(
            FieldMapping(
                csv_column=col,
                selector=selector,
                label_name=label,
                is_optional=optional,
            )
        )
    st.session_state.field_mappings = mappings

    # Step 3: Run
    st.subheader("Step 3 — Run Automation")
    state: GenericAutomationState = st.session_state.automation
    can_start = bool(form_url.strip()) and df is not None and len(df) > 0
    configured = any(m.selector.strip() or m.label_name.strip() for m in mappings)

    col_start, col_stop = st.columns([1, 3])
    with col_start:
        if st.button(
            "▶️ Start Generic Automation",
            type="primary",
            disabled=state.running or not can_start or not configured,
            use_container_width=True,
            key="start_generic_btn",
        ):
            start_generic_automation(
                df,
                form_url.strip(),
                mappings,
                st.session_state.headless,
                submit_sel,
                pref_label,
            )
            st.rerun()

    with col_stop:
        if state.running:
            prog = state.current_row / max(state.total_rows, 1)
            st.progress(prog, text=f"Processing row {state.current_row} / {state.total_rows}")

    if state.running and state.wait_seconds_remaining > 0:
        st.info(
            f"⏳ Cooldown: {state.wait_seconds_remaining}s (of {state.wait_total_seconds}s) remaining..."
        )

    st.markdown("**Live Log**")
    st.markdown(
        f'<div class="log-box">{state.get_logs() or "Waiting for logs..."}</div>',
        unsafe_allow_html=True,
    )

    if state.completed and not state.running:
        st.markdown(
            """
            <div class="success-banner">
                🎊 All Generic Form Submissions Completed! 🎊
            </div>
            """,
            unsafe_allow_html=True,
        )

    if state.running:
        time.sleep(1)
        st.rerun()


# ---------------------------------------------------------------------------
# Main Application Entrypoint
# ---------------------------------------------------------------------------
def main() -> None:
    st.markdown('<p class="main-header">🤖 Form & Portal Automation Dashboard</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="sub-header">Streamlit + Playwright engine for multi-step portal workflows & bulk data entry</p>',
        unsafe_allow_html=True,
    )

    tab_portal, tab_generic = st.tabs(
        ["🚀 Multi-Step Portal Workflow", "⚙️ Generic Form Automation"]
    )

    with tab_portal:
        render_portal_workflow()

    with tab_generic:
        render_generic_workflow()


if __name__ == "__main__":
    main()
