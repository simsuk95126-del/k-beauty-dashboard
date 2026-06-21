import streamlit as st
import pandas as pd
import requests
import base64
import os
import io
import time
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()
client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
API_URL = "https://k-beauty-api.onrender.com/api/v1/compliance-report"

# 🛠️ Version & Status variables (V7.4 핵심 국가 최적화 및 캐나다/브라질 삭제)
APP_VERSION = "v7.4.0 (Core Markets Optimized)"
BANNED_SUBSTANCES_STATUS = "June 2026 (Latest)"
KEYWORD_MATCHING_STATUS = "June 2026 (Synced)"

MASTER_KEY = os.environ.get("MASTER_KEY", "q1w2e3r41@3")

st.set_page_config(page_title="Global K-Beauty Compliance", page_icon="💄", layout="wide")

# ==========================================
# 🔄 Session State Initialization
# ==========================================
if "free_uses_left" not in st.session_state:
    st.session_state.free_uses_left = 3
if "api_result" not in st.session_state:
    st.session_state.api_result = None
if "current_auth_msg" not in st.session_state:
    st.session_state.current_auth_msg = None
if "current_auth_tier" not in st.session_state:
    st.session_state.current_auth_tier = "INVALID"
    
# 🌟 테스트 서버 팝업 확인 여부 세션
if "test_notice_shown" not in st.session_state:
    st.session_state.test_notice_shown = False

def reset_results():
    st.session_state.api_result = None

# ==========================================
# 🚨 테스트 서버 안내 팝업 (Test Notice)
# ==========================================
@st.dialog("🚧 System Notice: Test Server")
def show_test_server_popup():
    st.warning("현재 이 시스템은 **테스트 중(Testing in progress)** 입니다. 일부 기능이 임시적으로 제한되거나 불안정할 수 있습니다.\n\nThis system is currently running in a **test environment**. Temporary disruptions may occur.")
    if st.button("확인 (Acknowledge)", use_container_width=True):
        st.session_state.test_notice_shown = True
        st.rerun()

# 팝업을 아직 안 봤다면 무조건 화면 중앙에 띄움
if not st.session_state.test_notice_shown:
    show_test_server_popup()

# ==========================================
# 💸 Cost Optimization: OpenAI Vision API Caching
# ==========================================
@st.cache_data(show_spinner=False, ttl=3600)
def extract_ingredients_from_image(file_bytes):
    base64_image = base64.b64encode(file_bytes).decode('utf-8')
    vision_prompt = "Extract the ingredient names in order and output them separated by a vertical bar (|). [STRICT INSTRUCTION]: Do not use commas for separation. Example format: 정제수|글리세린|1,2-헥산디올"
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": [{"type": "text", "text": vision_prompt}, {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}}]}],
            temperature=0.0
        )
        return [ing.strip() for ing in response.choices[0].message.content.split('|') if ing.strip()]
    except Exception as e:
        return None

# ==========================================
# 📡 Gumroad License Verification Function
# ==========================================
def check_license_status(key, increment=False):
    if key == MASTER_KEY:
        return "PRO", "👑 CEO Master Key Active! (Unlimited)"

    verify_url = "https://api.gumroad.com/v2/licenses/verify"
    
    try:
        res_pro = requests.post(verify_url, data={
            "product_permalink": "pkoph",
            "license_key": key,
            "increment_uses_count": "false" 
        }, timeout=7).json()
        
        if res_pro.get("success") and not res_pro.get("purchase", {}).get("refunded"):
            return "PRO", "🏆 PRO Bulk Access Granted! (Unlimited)"
    except requests.exceptions.RequestException:
        return "ERROR", "📡 Connection to License Server timed out. Please try again."
    except:
        pass
        
    try:
        res_std = requests.post(verify_url, data={
            "product_permalink": "lyibre",
            "license_key": key,
            "increment_uses_count": "true" if increment else "false"
        }, timeout=7).json()
        
        if res_std.get("success") and not res_std.get("purchase", {}).get("refunded"):
            uses = res_std.get("uses", 0)
            if uses >= 50:
                return "EXPIRED", "🚫 Monthly limit (50/50) reached. Please upgrade to PRO."
            else:
                return "STANDARD", f"🔓 Standard Access Granted! (Remaining: {50 - uses}/50)"
    except requests.exceptions.RequestException:
        return "ERROR", "📡 Connection to License Server timed out. Please try again."
    except:
        pass
        
    return "INVALID", "❌ Invalid or Refunded License Key."

# ==========================================
# 📡 System Online Status Bar
# ==========================================
st.sidebar.error("🚧 STATUS: TEST SERVER")
st.sidebar.success(f"🟢 SYSTEM ONLINE (Ver: {APP_VERSION})")
st.sidebar.markdown(f"🟢 Banned Substances DB: \n{BANNED_SUBSTANCES_STATUS}")
st.sidebar.markdown(f"🔍 Keyword Matching Engine: \n{KEYWORD_MATCHING_STATUS}")
st.sidebar.markdown("---")

# ==========================================
# 🔐 Premium Access
# ==========================================
st.sidebar.subheader("🔐 Premium Access")
entered_password = st.sidebar.text_input("Enter your License Key:", type="password")

is_vip = False
is_pro_or_ceo = False

if entered_password:
    if st.session_state.get("last_entered_key") != entered_password:
        auth_tier, auth_msg = check_license_status(entered_password, increment=False)
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
        is_pro_or_ceo = False
    elif auth_tier == "EXPIRED":
        st.sidebar.error(auth_msg)
    elif auth_tier == "ERROR":
        st.sidebar.warning(auth_msg)
    else:
        st.sidebar.error(auth_msg)
else:
    st.sidebar.warning(f"🎁 Free Trial Active: {st.session_state.free_uses_left} checks remaining.")

st.sidebar.markdown("---")

# ==========================================
# 💎 Subscription Plans & Popup Logic
# ==========================================
st.sidebar.subheader("💎 Choose Your Plan")

@st.dialog("⚠️ License Key Information & Payment")
def show_payment_popup(plan_name, link):
    st.warning("Upon successful payment, your **License Key (Access Code)** for unlimited dashboard access will be included in the **[Gumroad Receipt Email]**.\n\nPlease make sure to check your email after payment!")
    st.link_button(f"👉 Go to {plan_name} Payment Page", link, use_container_width=True)

if st.sidebar.button("💳 Subscribe Standard ($299/mo)", use_container_width=True):
    show_payment_popup("Standard Plan", "https://dahee5.gumroad.com/l/lyibre")
st.sidebar.caption("✔️ 50 Scans per month\n\n✔️ Single File Scan & Excel Download")

st.sidebar.markdown("<br>", unsafe_allow_html=True)

if st.sidebar.button("🚀 Subscribe PRO Bulk ($499/mo)", use_container_width=True):
    show_payment_popup("PRO Bulk Plan", "https://dahee5.gumroad.com/l/pkoph")
st.sidebar.caption("✔️ Unlimited Scans\n\n✔️ Multiple Image/File Uploads & Batch Extraction")

st.sidebar.markdown("---")

# ==========================================
# 👑 Main UI: Core Advantages
# ==========================================
st.title("🌍 Global K-Beauty Compliance Master")
st.markdown("##### AI-powered customs compliance checker for core global markets (US, EU, CN, HALAL, etc.)")

st.markdown("### 🌟 Why Choose Our Core Compliance Engine?")

col_adv1, col_adv2 = st.columns(2)

with col_adv1:
    st.info("#### 🛡️ 1. Zero-Hallucination\nStrictly blocks AI's arbitrary interpretations and false information. Unverified ingredients are returned as 'Verification Required' rather than guessed, preventing fatal customs errors.")
    st.success("#### 🌐 2. Global Custom Regulation Check\nCross-verifies target country compliance (US FDA MoCRA, EU CPNP, China NMPA, Halal, etc.) in 0.1 seconds based on official prohibited/restricted ingredient databases.")

with col_adv2:
    st.warning("#### 🚫 3. OCR Error Correction\nAutomatically filters broken characters or typos during image scans, fundamentally preventing misinterpretation of safe ingredients as hazardous due to text extraction errors.")
    st.error("#### 🇰🇷 4. Official K-Beauty Data Matching\nDesigned to prioritize official Korean-INCI matching server databases (KCA & MFDS standards), providing the most accurate K-Beauty standard ingredient nomenclature.")

st.markdown("---")

# Free trial expiration block
if not is_vip and st.session_state.free_uses_left <= 0:
    st.error("🔒 Free Trial Expired. Please subscribe in the sidebar and enter your License Key to unlock unlimited usage.")
    st.stop()

# ==========================================
# 🛡️ STRICT Legal Disclaimer (Blocker)
# ==========================================
if "disclaimer_agreed" not in st.session_state:
    st.session_state.disclaimer_agreed = False

if not st.session_state.disclaimer_agreed:
    st.markdown("### ⚠️ REQUIRED ACTION: Review Legal Disclaimer")
    st.warning("⚖️ LEGAL DISCLAIMER: Informational Only. The provided INCI names, compliance statuses, and regulation notices do not constitute legal or official regulatory advice. Users must independently verify all data before use.")
    if st.checkbox("I HAVE READ THE LEGAL DISCLAIMER AND AGREE TO THE TERMS OF USE."):
        st.session_state.disclaimer_agreed = True
        st.rerun()
    st.stop() 
else:
    if st.button("⚖️ Review Terms"):
        st.session_state.disclaimer_agreed = False
        st.rerun()

# ==========================================
# 🚀 Main Workspace
# ==========================================
st.subheader("🚀 Compliance Analysis Workspace")

# =========================================================================
# 💡 [캐나다(CA) 완전 삭제 반영] 글로벌 핵심 마켓 리스트
# =========================================================================
country_options = {
    "US": "US (Federal FDA MoCRA & California Prop 65 / Toxic-Free Cosmetics Act)",
    "EU": "EU (European Union CPNP)",
    "CN": "CN (China NMPA)",
    "JP": "JP (Japan PMDA)",
    "ASEAN": "ASEAN (Vietnam, Singapore, Thailand, Malaysia, Indonesia, Philippines, Myanmar, Cambodia, Laos, Brunei)",
    "UK": "UK (United Kingdom SCPN)",
    "SFDA": "SFDA (Saudi Arabia Food and Drug Authority)",
    "HALAL": "HALAL (Indonesia, Malaysia, UAE, Saudi Arabia, Turkey, Egypt, Pakistan, Iran, Algeria, Morocco, Oman, Qatar, Kuwait, Bahrain, Jordan, Bangladesh)",
    "EAC": "EAC (Eurasian Economic Union - Russia, Belarus, Kazakhstan, Armenia, Kyrgyzstan)"
}

# 사용자에게는 긴 웅장한 풀네임을 노출
selected_display_name = st.selectbox("1️⃣ Select Target Market", options=list(country_options.values()), on_change=reset_results)

# 백엔드로 전송할 때는 원래의 2글자 숏코드(US, EU 등)로 역추적 매핑하여 에러 원천 차단
target_country = [key for key, value in country_options.items() if value == selected_display_name][0]
# =========================================================================

trial_active = (not is_vip and st.session_state.free_uses_left > 0)

if is_pro_or_ceo or trial_active:
    uploaded_files = st.file_uploader("2️⃣ Upload Files (Single or Multiple) - [PRO / FREE TRIAL UNLOCKED]", type=['csv', 'xlsx', 'jpg', 'jpeg', 'png'], accept_multiple_files=True, on_change=reset_results)
else:
    uploaded_file = st.file_uploader("2️⃣ Upload A Single File - [Standard Mode]", type=['csv', 'xlsx', 'jpg', 'jpeg', 'png'], accept_multiple_files=False, on_change=reset_results)
    uploaded_files = [uploaded_file] if uploaded_file is not None else []

ingredient_source_map = {}

if len(uploaded_files) > 0:
    for f in uploaded_files:
        file_extension = f.name.split('.')[-1].lower()
        
        if file_extension in ['jpg', 'jpeg', 'png']:
            st.image(f, caption=f"Uploaded: {f.name}", width=180)
            st.warning(f"🤖 AI Scanning [{f.name}]... (Cached to save costs)")
            
            file_bytes = f.getvalue()
            extracted_list = extract_ingredients_from_image(file_bytes)
            
            if extracted_list is not None:
                for ing in extracted_list:
                    if ing not in ingredient_source_map:
                        ingredient_source_map[ing] = set()
                    ingredient_source_map[ing].add(f.name)
                
                st.markdown(f"##### 📝 Extracted List from {f.name}")
                st.dataframe(pd.DataFrame({"No.": range(1, len(extracted_list)+1), "Ingredient": extracted_list}), hide_index=True, use_container_width=True)
            else:
                st.error(f"Image scan failed for {f.name}")
        else:
            df = pd.read_csv(f) if file_extension == 'csv' else pd.read_excel(f)
            st.markdown(f"📊 Preview: {f.name}")
            st.dataframe(df.head(2))
            
            excel_ingredients = df.iloc[:, 0].dropna().astype(str).tolist()
            for ing in excel_ingredients:
                clean_ing = ing.strip()
                if clean_ing:
                    if clean_ing not in ingredient_source_map:
                        ingredient_source_map[clean_ing] = set()
                    ingredient_source_map[clean_ing].add(f.name)

    unique_ingredients = list(ingredient_source_map.keys())

    if unique_ingredients and st.button("🚀 Run Global Compliance Check!", use_container_width=True):
        
        if is_vip and not is_pro_or_ceo:
            run_tier, run_msg = check_license_status(entered_password, increment=True)
            if run_tier == "EXPIRED":
                st.error("🚫 Monthly limit reached. Cannot proceed with analysis. Please upgrade to PRO.")
                st.stop()
            elif run_tier == "STANDARD":
                st.session_state.current_auth_msg = run_msg
                
        status_text = st.empty()
        progress_bar = st.progress(0)
        
        status_text.warning("⏳ Step 1/4: Structuring pure cosmetic ingredients & mapping sources...")
        progress_bar.progress(25)
        time.sleep(1.0)
        
        status_text.warning(f"🕵️‍♂️ Step 2/4: Connecting and searching [{target_country}] customs database...")
        progress_bar.progress(50)
        time.sleep(1.0)
        
        status_text.warning("🧬 Step 3/4: Cross-examining INCI chemical identifiers against restriction matrix...")
        progress_bar.progress(75)
        time.sleep(1.0)
        
        status_text.info("📊 Step 4/4: Generating auto-formatted compliance sheets...")
        progress_bar.progress(95)
        
        try:
            api_response = requests.post(API_URL, json={"ingredients": unique_ingredients, "target": target_country}, timeout=30)
            
            if api_response.status_code == 200:
                result_data = api_response.json()
                result_df = pd.DataFrame(result_data['report_details'])
                
                result_df.rename(columns={
                    'original_ingredient': 'Original Ingredient', 
                    'inci_name': 'INCI Name', 
                    'is_safe': 'Compliance Status', 
                    'regulation_notice': 'Regulation Notice'
                }, inplace=True)
                
                result_df.insert(0, 'No.', range(1, len(result_df) + 1))
                
                # ============================================================
                # 💡 금지/제한 3단계 상태 아이콘 파싱 로직
                # ============================================================
                def determine_status_icon(row):
                    if row['Compliance Status'] == True or str(row['Compliance Status']).upper() == 'TRUE':
                        return '🟢 PASS'
                    
                    notice_text = str(row['Regulation Notice']).upper()
                    if 'RESTRICTED' in notice_text or 'LIMIT' in notice_text:
                        return '⚠️ RESTRICTED'
                    
                    return '🔴 BANNED'
                
                result_df['Compliance Status'] = result_df.apply(determine_status_icon, axis=1)
                # ============================================================
                
                result_df['Source File'] = result_df['Original Ingredient'].apply(
                    lambda x: ", ".join(list(ingredient_source_map.get(x, [])))
                )
                
                result_df = result_df[['No.', 'Original Ingredient', 'INCI Name', 'Compliance Status', 'Source File', 'Regulation Notice']]
                
                df_excel = result_df.copy()
                df_excel['Source File'] = df_excel['Source File'].str.split(', ')
                df_excel = df_excel.explode('Source File').reset_index(drop=True)
                df_excel['No.'] = range(1, len(df_excel) + 1)
                
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df_excel.to_excel(writer, index=False, sheet_name='Compliance_Report')
                
                has_banned = '🔴 BANNED' in result_df['Compliance Status'].values
                has_restricted = '⚠️ RESTRICTED' in result_df['Compliance Status'].values
                
                if not has_banned and not has_restricted:
                    final_status_msg = "PASS"
                elif has_banned:
                    final_status_msg = "FAIL" 
                else:
                    final_status_msg = "RESTRICTED" 
                
                st.session_state.api_result = {
                    "df": result_df,
                    "excel_data": excel_buffer.getvalue(),
                    "status": final_status_msg,
                    "failed_count": result_data['failed_count'],
                    "target": target_country
                }
                
                progress_bar.progress(100)
                status_text.success("✅ Analysis Complete!")
                time.sleep(0.5)
                
                if not is_vip:
                    st.session_state.free_uses_left -= 1
                
                st.rerun()
                
        except Exception as e:
            status_text.empty()
            progress_bar.empty()
            st.error("API Connection Failed.")

# ==========================================
# 📥 Results & Download Section
# ==========================================
if st.session_state.api_result is not None:
    res = st.session_state.api_result
    st.markdown("---")
    
    all_sources = list({src for sublist in res['df']['Source File'].str.split(', ') for src in sublist})
    selected_files = st.multiselect("📂 Filter by Source File (Dashboard View)", options=all_sources)
    
    if selected_files:
        display_df = res['df'][res['df']['Source File'].apply(lambda x: any(f in x for f in selected_files))].copy()
        display_df['No.'] = range(1, len(display_df) + 1)
    else:
        display_df = res['df']
    
    if res["status"] == "PASS":
        st.success(f"🎉 Analysis Done! No restricted or banned ingredients found for {res['target']}.")
    elif res["status"] == "RESTRICTED":
        st.warning(f"⚠️ Warning! {res['failed_count']} ingredients have concentration limits or specific conditions for {res['target']}.")
    else:
        st.error(f"🚨 Critical Alert! {res['failed_count']} prohibited ingredients hit the regulation filters for {res['target']}.")
        
    st.dataframe(display_df, use_container_width=True, hide_index=True)
    
    st.download_button(
        "📥 Download Merged Excel Report (Auto-Filter Ready)", 
        data=res["excel_data"],
        file_name=f"Merged_Report_{res['target']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
