import base64
import io
import os
import time
from typing import Dict, List, Optional, Set, Tuple

import pandas as pd
import requests
import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI

# ==========================================
# 기본 설정
# ==========================================
load_dotenv()

API_BASE_URL = os.environ.get(
    "API_BASE_URL",
    "https://k-beauty-api.onrender.com",
).rstrip("/")
COMPLIANCE_API_URL = f"{API_BASE_URL}/api/v1/compliance-report"
DATABASE_STATUS_API_URL = f"{API_BASE_URL}/api/v1/database-status"

OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")
MASTER_KEY = os.environ.get("MASTER_KEY")
client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None

APP_VERSION = "v8.0.0 (Structured Regulation Status)"
BANNED_SUBSTANCES_STATUS = "Loaded from server data/*.csv"
KEYWORD_MATCHING_STATUS = "Explicit API status mapping"

REQUEST_TIMEOUT_SECONDS = 60
LICENSE_TIMEOUT_SECONDS = 10

st.set_page_config(
    page_title="Global K-Beauty Compliance",
    page_icon="💄",
    layout="wide",
)

# ==========================================
# 세션 상태
# ==========================================
def initialize_session_state() -> None:
    defaults = {
        "free_uses_left": 3,
        "api_result": None,
        "current_auth_msg": None,
        "current_auth_tier": "INVALID",
        "test_notice_shown": False,
        "disclaimer_agreed": False,
        "last_entered_key": None,
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


initialize_session_state()


def reset_results() -> None:
    st.session_state.api_result = None


# ==========================================
# 테스트 서버 안내
# ==========================================
@st.dialog("🚧 System Notice: Test Server")
def show_test_server_popup() -> None:
    st.warning(
        "현재 이 시스템은 **테스트 중(Testing in progress)** 입니다. "
        "일부 기능이 임시적으로 제한되거나 불안정할 수 있습니다.\n\n"
        "This system is currently running in a **test environment**. "
        "Temporary disruptions may occur."
    )

    if st.button("확인 (Acknowledge)", use_container_width=True):
        st.session_state.test_notice_shown = True
        st.rerun()


if not st.session_state.test_notice_shown:
    show_test_server_popup()


# ==========================================
# 이미지 성분 추출
# ==========================================
@st.cache_data(show_spinner=False, ttl=3600)
def extract_ingredients_from_image(file_bytes: bytes) -> Optional[List[str]]:
    if client is None:
        return None

    base64_image = base64.b64encode(file_bytes).decode("utf-8")
    vision_prompt = (
        "Extract cosmetic ingredient names in their original order. "
        "Return ingredient names separated only by a vertical bar (|). "
        "Do not use commas as separators because commas may be part of an ingredient name. "
        "Example: 정제수|글리세린|1,2-헥산디올"
    )

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": vision_prompt},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/jpeg;base64,{base64_image}"
                            },
                        },
                    ],
                }
            ],
            temperature=0.0,
        )

        content = response.choices[0].message.content or ""
        ingredients = [item.strip() for item in content.split("|") if item.strip()]
        return ingredients or None
    except Exception:
        return None


# ==========================================
# Gumroad 라이선스 검증
# ==========================================
def check_license_status(key: str, increment: bool = False) -> Tuple[str, str]:
    clean_key = key.strip()
    if not clean_key:
        return "INVALID", "❌ License Key를 입력해 주세요."

    if MASTER_KEY and clean_key == MASTER_KEY:
        return "PRO", "👑 CEO Master Key Active! (Unlimited)"

    verify_url = "https://api.gumroad.com/v2/licenses/verify"

    try:
        pro_response = requests.post(
            verify_url,
            data={
                "product_permalink": "pkoph",
                "license_key": clean_key,
                "increment_uses_count": "false",
            },
            timeout=LICENSE_TIMEOUT_SECONDS,
        )
        pro_response.raise_for_status()
        pro_data = pro_response.json()

        if pro_data.get("success") and not pro_data.get("purchase", {}).get(
            "refunded"
        ):
            return "PRO", "🏆 PRO Bulk Access Granted! (Unlimited)"
    except (requests.RequestException, ValueError):
        pass

    try:
        standard_response = requests.post(
            verify_url,
            data={
                "product_permalink": "lyibre",
                "license_key": clean_key,
                "increment_uses_count": "true" if increment else "false",
            },
            timeout=LICENSE_TIMEOUT_SECONDS,
        )
        standard_response.raise_for_status()
        standard_data = standard_response.json()

        if standard_data.get("success") and not standard_data.get(
            "purchase", {}
        ).get("refunded"):
            uses = int(standard_data.get("uses", 0) or 0)

            if uses >= 50:
                return (
                    "EXPIRED",
                    "🚫 Monthly limit (50/50) reached. Please upgrade to PRO.",
                )

            return (
                "STANDARD",
                f"🔓 Standard Access Granted! (Remaining: {50 - uses}/50)",
            )
    except requests.RequestException:
        return (
            "ERROR",
            "📡 Connection to License Server timed out. Please try again.",
        )
    except (ValueError, TypeError):
        return "ERROR", "📡 Invalid response from License Server."

    return "INVALID", "❌ Invalid or Refunded License Key."


# ==========================================
# API 상태 및 응답 유틸리티
# ==========================================
@st.cache_data(show_spinner=False, ttl=300)
def fetch_database_status() -> Optional[dict]:
    try:
        response = requests.get(
            DATABASE_STATUS_API_URL,
            timeout=10,
        )
        response.raise_for_status()
        return response.json()
    except (requests.RequestException, ValueError):
        return None


def normalize_api_status(value: object) -> str:
    status = str(value or "").strip().upper()

    allowed = {
        "PASS",
        "BANNED",
        "RESTRICTED",
        "WARNING_REQUIRED",
        "REGULATED",
        "VERIFICATION_REQUIRED",
    }

    return status if status in allowed else "VERIFICATION_REQUIRED"


def status_label(status: str) -> str:
    labels = {
        "PASS": "🟢 PASS",
        "BANNED": "🔴 BANNED",
        "RESTRICTED": "🟠 RESTRICTED",
        "WARNING_REQUIRED": "🟡 WARNING REQUIRED",
        "REGULATED": "🟡 REGULATED",
        "VERIFICATION_REQUIRED": "⚪ VERIFICATION REQUIRED",
    }
    return labels.get(status, "⚪ VERIFICATION REQUIRED")


def build_result_dataframe(
    report_details: List[dict],
    ingredient_source_map: Dict[str, Set[str]],
) -> pd.DataFrame:
    rows = []

    for index, item in enumerate(report_details, start=1):
        original_ingredient = str(item.get("original_ingredient", "")).strip()
        explicit_status = item.get("compliance_status") or item.get(
            "restriction_type"
        )
        normalized_status = normalize_api_status(explicit_status)

        source_files = sorted(ingredient_source_map.get(original_ingredient, set()))

        rows.append(
            {
                "No.": index,
                "Original Ingredient": original_ingredient,
                "INCI Name": str(item.get("inci_name", "")).strip(),
                "CAS Number": str(item.get("cas_number", "N/A")).strip() or "N/A",
                "Compliance Status": status_label(normalized_status),
                "Status Code": normalized_status,
                "Source File": ", ".join(source_files),
                "Regulation Reason": str(
                    item.get("regulation_reason", "")
                ).strip(),
                "Regulation Notice": str(
                    item.get("regulation_notice", "")
                ).strip(),
                "Translation Source": str(item.get("source", "")).strip(),
            }
        )

    return pd.DataFrame(rows)


def build_excel_report(
    result_df: pd.DataFrame,
    summary_data: dict,
) -> bytes:
    report_df = result_df.copy()

    if not report_df.empty:
        report_df["Source File"] = report_df["Source File"].apply(
            lambda value: value.split(", ") if value else [""]
        )
        report_df = report_df.explode("Source File").reset_index(drop=True)
        report_df["No."] = range(1, len(report_df) + 1)

    status_counts = summary_data.get("status_counts", {})
    summary_rows = [
        {"Item": "Target Market", "Value": summary_data.get("target_market", "")},
        {
            "Item": "Overall Status",
            "Value": summary_data.get("compliance_status", ""),
        },
        {
            "Item": "Total Checked",
            "Value": summary_data.get("total_checked", 0),
        },
        {
            "Item": "Database File",
            "Value": summary_data.get("database_file", ""),
        },
        {
            "Item": "Database Last Updated",
            "Value": summary_data.get("database_last_updated", ""),
        },
        {
            "Item": "Report Generated At",
            "Value": summary_data.get("report_generated_at", ""),
        },
    ]

    for status_name, count in status_counts.items():
        summary_rows.append(
            {"Item": f"Count - {status_name}", "Value": count}
        )

    summary_df = pd.DataFrame(summary_rows)
    excel_buffer = io.BytesIO()

    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        report_df.to_excel(
            writer,
            index=False,
            sheet_name="Compliance_Report",
        )
        summary_df.to_excel(
            writer,
            index=False,
            sheet_name="Summary",
        )

    return excel_buffer.getvalue()


# ==========================================
# 사이드바 상태
# ==========================================
st.sidebar.error("🚧 STATUS: TEST SERVER")
st.sidebar.success(f"🟢 SYSTEM ONLINE (Ver: {APP_VERSION})")
st.sidebar.markdown(f"🗂️ Regulatory DB:  \n{BANNED_SUBSTANCES_STATUS}")
st.sidebar.markdown(f"🔍 Matching Engine:  \n{KEYWORD_MATCHING_STATUS}")

server_database_status = fetch_database_status()
if server_database_status:
    loaded_total = sum(
        int(item.get("loaded_records", 0) or 0)
        for item in server_database_status.values()
        if isinstance(item, dict)
    )
    st.sidebar.success(f"📚 Loaded regulatory records: {loaded_total:,}")
else:
    st.sidebar.warning("📡 Regulatory DB status unavailable")

st.sidebar.markdown("---")

# ==========================================
# Premium Access
# ==========================================
st.sidebar.subheader("🔐 Premium Access")
entered_password = st.sidebar.text_input(
    "Enter your License Key:",
    type="password",
)

is_vip = False
is_pro_or_ceo = False

if entered_password:
    if st.session_state.last_entered_key != entered_password:
        auth_tier, auth_msg = check_license_status(
            entered_password,
            increment=False,
        )
        st.session_state.current_auth_tier = auth_tier
        st.session_state.current_auth_msg = auth_msg
        st.session_state.last_entered_key = entered_password

    auth_tier = st.session_state.current_auth_tier
    auth_msg = st.session_state.current_auth_msg

    if auth_tier == "PRO":
        st.sidebar.success(auth_msg)
        is_vip = True
        is_pro_or_ceo = True
    elif auth_tier == "STANDARD":
        st.sidebar.success(auth_msg)
        is_vip = True
    elif auth_tier == "EXPIRED":
        st.sidebar.error(auth_msg)
    elif auth_tier == "ERROR":
        st.sidebar.warning(auth_msg)
    else:
        st.sidebar.error(auth_msg)
else:
    st.sidebar.warning(
        f"🎁 Free Trial Active: {st.session_state.free_uses_left} checks remaining."
    )

st.sidebar.markdown("---")

# ==========================================
# 결제 메뉴
# ==========================================
st.sidebar.subheader("💎 Choose Your Plan")


@st.dialog("⚠️ License Key Information & Payment")
def show_payment_popup(plan_name: str, link: str) -> None:
    st.warning(
        "Upon successful payment, your **License Key (Access Code)** will be "
        "included in the **Gumroad Receipt Email**. Please check your email "
        "after payment."
    )
    st.link_button(
        f"👉 Go to {plan_name} Payment Page",
        link,
        use_container_width=True,
    )


if st.sidebar.button(
    "💳 Subscribe Standard ($299/mo)",
    use_container_width=True,
):
    show_payment_popup(
        "Standard Plan",
        "https://dahee5.gumroad.com/l/lyibre",
    )

st.sidebar.caption(
    "✔️ 50 Scans per month\n\n✔️ Single File Scan & Excel Download"
)
st.sidebar.markdown("<br>", unsafe_allow_html=True)

if st.sidebar.button(
    "🚀 Subscribe PRO Bulk ($499/mo)",
    use_container_width=True,
):
    show_payment_popup(
        "PRO Bulk Plan",
        "https://dahee5.gumroad.com/l/pkoph",
    )

st.sidebar.caption(
    "✔️ Unlimited Scans\n\n✔️ Multiple Image/File Uploads & Batch Extraction"
)
st.sidebar.markdown("---")

# ==========================================
# 메인 화면
# ==========================================
st.title("🌍 Global K-Beauty Compliance Master")
st.markdown(
    "##### AI-powered cosmetic ingredient screening for core global markets"
)

st.markdown("### 🌟 Core Compliance Engine")
col_adv1, col_adv2 = st.columns(2)

with col_adv1:
    st.info(
        "#### 🛡️ Structured Status\n"
        "The system uses explicit regulatory status fields returned by the API. "
        "It does not infer BANNED or RESTRICTED from warning text."
    )
    st.success(
        "#### 🌐 Country-specific CSV Matching\n"
        "The selected market is checked against its server-side data/*.csv file."
    )

with col_adv2:
    st.warning(
        "#### 🚫 Manual Verification\n"
        "Unknown or uncertain ingredients are marked Verification Required "
        "instead of being treated as safe."
    )
    st.error(
        "#### 🇰🇷 Korean-INCI Matching\n"
        "The Korean master database is checked before AI-assisted normalization."
    )

st.markdown("---")

if not is_vip and st.session_state.free_uses_left <= 0:
    st.error(
        "🔒 Free Trial Expired. Please subscribe and enter your License Key."
    )
    st.stop()

# ==========================================
# 법적 고지
# ==========================================
if not st.session_state.disclaimer_agreed:
    st.markdown("### ⚠️ REQUIRED ACTION: Review Legal Disclaimer")
    st.warning(
        "⚖️ LEGAL DISCLAIMER: Informational screening only. INCI names, "
        "regulatory statuses, concentration limits, product categories and "
        "current official requirements must be independently verified before use."
    )

    if st.checkbox(
        "I HAVE READ THE LEGAL DISCLAIMER AND AGREE TO THE TERMS OF USE."
    ):
        st.session_state.disclaimer_agreed = True
        st.rerun()

    st.stop()
else:
    if st.button("⚖️ Review Terms"):
        st.session_state.disclaimer_agreed = False
        st.rerun()

# ==========================================
# 분석 작업공간
# ==========================================
st.subheader("🚀 Compliance Analysis Workspace")

country_options = {
    "US": "US (FDA / California regulatory screening)",
    "EU": "EU (Cosmetics Regulation Annex screening)",
    "CN": "CN (China NMPA)",
    "ASEAN": "ASEAN (ASEAN Cosmetic Directive)",
    "UK": "UK (UK Cosmetics Regulation)",
    "SFDA": "SFDA (Saudi Food and Drug Authority)",
    "HALAL": "HALAL (ingredient-origin and certification screening)",
    "EAC": "EAC (Eurasian Economic Union)",
}

selected_display_name = st.selectbox(
    "1️⃣ Select Target Market",
    options=list(country_options.values()),
    on_change=reset_results,
)

target_country = next(
    key
    for key, value in country_options.items()
    if value == selected_display_name
)

trial_active = not is_vip and st.session_state.free_uses_left > 0

if is_pro_or_ceo or trial_active:
    uploaded_files = st.file_uploader(
        "2️⃣ Upload Files (Single or Multiple) - [PRO / FREE TRIAL]",
        type=["csv", "xlsx", "jpg", "jpeg", "png"],
        accept_multiple_files=True,
        on_change=reset_results,
    )
else:
    uploaded_file = st.file_uploader(
        "2️⃣ Upload A Single File - [Standard Mode]",
        type=["csv", "xlsx", "jpg", "jpeg", "png"],
        accept_multiple_files=False,
        on_change=reset_results,
    )
    uploaded_files = [uploaded_file] if uploaded_file is not None else []

ingredient_source_map: Dict[str, Set[str]] = {}

for uploaded_file in uploaded_files:
    file_extension = uploaded_file.name.rsplit(".", 1)[-1].lower()

    if file_extension in {"jpg", "jpeg", "png"}:
        st.image(
            uploaded_file,
            caption=f"Uploaded: {uploaded_file.name}",
            width=180,
        )
        st.warning(f"🤖 AI Scanning [{uploaded_file.name}]...")

        if client is None:
            st.error("OPENAI_API_KEY가 없어 이미지 성분 추출을 실행할 수 없습니다.")
            continue

        extracted_list = extract_ingredients_from_image(uploaded_file.getvalue())

        if not extracted_list:
            st.error(f"Image scan failed: {uploaded_file.name}")
            continue

        for ingredient in extracted_list:
            ingredient_source_map.setdefault(ingredient, set()).add(
                uploaded_file.name
            )

        st.markdown(f"##### 📝 Extracted List from {uploaded_file.name}")
        st.dataframe(
            pd.DataFrame(
                {
                    "No.": range(1, len(extracted_list) + 1),
                    "Ingredient": extracted_list,
                }
            ),
            hide_index=True,
            use_container_width=True,
        )
        continue

    try:
        if file_extension == "csv":
            uploaded_df = pd.read_csv(uploaded_file)
        else:
            uploaded_df = pd.read_excel(uploaded_file)
    except Exception as error:
        st.error(f"파일을 읽을 수 없습니다: {uploaded_file.name} ({error})")
        continue

    if uploaded_df.empty or len(uploaded_df.columns) == 0:
        st.warning(f"빈 파일입니다: {uploaded_file.name}")
        continue

    st.markdown(f"📊 Preview: {uploaded_file.name}")
    st.dataframe(uploaded_df.head(3), use_container_width=True)

    ingredients = (
        uploaded_df.iloc[:, 0].dropna().astype(str).map(str.strip).tolist()
    )

    for ingredient in ingredients:
        if ingredient:
            ingredient_source_map.setdefault(ingredient, set()).add(
                uploaded_file.name
            )

unique_ingredients = list(ingredient_source_map.keys())

if unique_ingredients:
    st.info(f"총 {len(unique_ingredients):,}개의 고유 성분을 확인했습니다.")

    if st.button(
        "🚀 Run Global Compliance Check!",
        use_container_width=True,
    ):
        if is_vip and not is_pro_or_ceo:
            run_tier, run_msg = check_license_status(
                entered_password,
                increment=True,
            )

            if run_tier == "EXPIRED":
                st.error(
                    "🚫 Monthly limit reached. Cannot proceed. Please upgrade to PRO."
                )
                st.stop()

            if run_tier != "STANDARD":
                st.error(run_msg)
                st.stop()

            st.session_state.current_auth_msg = run_msg

        status_text = st.empty()
        progress_bar = st.progress(0)

        status_text.warning("⏳ Step 1/4: Preparing ingredient list...")
        progress_bar.progress(25)
        time.sleep(0.3)

        status_text.warning(
            f"🕵️ Step 2/4: Loading [{target_country}] regulatory database..."
        )
        progress_bar.progress(50)
        time.sleep(0.3)

        status_text.warning("🧬 Step 3/4: Matching explicit regulatory status...")
        progress_bar.progress(75)
        time.sleep(0.3)

        status_text.info("📊 Step 4/4: Generating compliance report...")
        progress_bar.progress(90)

        try:
            api_response = requests.post(
                COMPLIANCE_API_URL,
                json={
                    "ingredients": unique_ingredients,
                    "target": target_country,
                },
                timeout=REQUEST_TIMEOUT_SECONDS,
            )

            if api_response.status_code != 200:
                try:
                    error_detail = api_response.json().get(
                        "detail",
                        api_response.text,
                    )
                except ValueError:
                    error_detail = api_response.text

                raise RuntimeError(
                    f"API returned HTTP {api_response.status_code}: {error_detail}"
                )

            result_data = api_response.json()
            report_details = result_data.get("report_details", [])

            if not isinstance(report_details, list):
                raise ValueError("API report_details 형식이 올바르지 않습니다.")

            result_df = build_result_dataframe(
                report_details,
                ingredient_source_map,
            )

            excel_data = build_excel_report(result_df, result_data)

            st.session_state.api_result = {
                "df": result_df,
                "excel_data": excel_data,
                "overall_status": result_data.get(
                    "compliance_status",
                    "REVIEW_REQUIRED",
                ),
                "status_counts": result_data.get("status_counts", {}),
                "failed_count": result_data.get("failed_count", 0),
                "target": result_data.get("target_market", target_country),
                "database_file": result_data.get("database_file", ""),
                "database_last_updated": result_data.get(
                    "database_last_updated"
                ),
                "report_generated_at": result_data.get("report_generated_at"),
            }

            progress_bar.progress(100)
            status_text.success("✅ Analysis Complete!")

            if not is_vip:
                st.session_state.free_uses_left -= 1

            st.rerun()

        except (requests.RequestException, RuntimeError, ValueError) as error:
            status_text.empty()
            progress_bar.empty()
            st.error(f"API Connection or Response Error: {error}")

# ==========================================
# 결과 및 다운로드
# ==========================================
if st.session_state.api_result is not None:
    result = st.session_state.api_result
    result_df = result["df"]

    st.markdown("---")
    st.subheader("📋 Compliance Result")

    metadata_col1, metadata_col2, metadata_col3 = st.columns(3)
    metadata_col1.metric("Target Market", result["target"])
    metadata_col2.metric("Database File", result["database_file"] or "N/A")
    metadata_col3.metric(
        "DB Last Updated",
        result["database_last_updated"] or "Unknown",
    )

    status_counts = result.get("status_counts", {})
    summary_columns = st.columns(4)
    summary_columns[0].metric("PASS", status_counts.get("PASS", 0))
    summary_columns[1].metric("BANNED", status_counts.get("BANNED", 0))
    summary_columns[2].metric(
        "RESTRICTED",
        status_counts.get("RESTRICTED", 0),
    )
    summary_columns[3].metric(
        "VERIFY / OTHER",
        status_counts.get("VERIFICATION_REQUIRED", 0)
        + status_counts.get("WARNING_REQUIRED", 0)
        + status_counts.get("REGULATED", 0),
    )

    overall_status = str(result.get("overall_status", "")).upper()

    if overall_status == "PASS":
        st.success(
            f"🎉 No matched regulated ingredients were found for {result['target']}."
        )
    elif overall_status == "FAIL":
        st.error(
            f"🚨 One or more BANNED ingredients were found for {result['target']}."
        )
    else:
        st.warning(
            f"⚠️ Review required for one or more ingredients in {result['target']}."
        )

    available_sources = sorted(
        {
            source
            for source_text in result_df.get("Source File", pd.Series(dtype=str))
            for source in str(source_text).split(", ")
            if source
        }
    )

    selected_files = st.multiselect(
        "📂 Filter by Source File",
        options=available_sources,
    )

    selected_statuses = st.multiselect(
        "🚦 Filter by Compliance Status",
        options=[
            "🟢 PASS",
            "🔴 BANNED",
            "🟠 RESTRICTED",
            "🟡 WARNING REQUIRED",
            "🟡 REGULATED",
            "⚪ VERIFICATION REQUIRED",
        ],
    )

    display_df = result_df.copy()

    if selected_files:
        display_df = display_df[
            display_df["Source File"].apply(
                lambda value: any(
                    selected_file in str(value).split(", ")
                    for selected_file in selected_files
                )
            )
        ]

    if selected_statuses:
        display_df = display_df[
            display_df["Compliance Status"].isin(selected_statuses)
        ]

    display_df = display_df.copy()
    display_df["No."] = range(1, len(display_df) + 1)

    visible_columns = [
        "No.",
        "Original Ingredient",
        "INCI Name",
        "CAS Number",
        "Compliance Status",
        "Source File",
        "Regulation Reason",
        "Regulation Notice",
    ]

    st.dataframe(
        display_df[visible_columns],
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        "📥 Download Merged Excel Report",
        data=result["excel_data"],
        file_name=f"Merged_Report_{result['target']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
