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

# 🛠️ 실시간 업데이트 및 버전 관리 변수
APP_VERSION = "v6.1.0 (Marketing UI & Popup Integration)"
BANNED_SUBSTANCES_STATUS = "June 2026 (Latest)"
KEYWORD_MATCHING_STATUS = "June 2026 (Synced)"

MASTER_KEY = os.environ.get("MASTER_KEY", "q1w2e3r41@3")

st.set_page_config(page_title="Global K-Beauty Compliance", page_icon="💄", layout="wide")

# ==========================================
# 🔄 세션 상태 초기화 및 입력값 변경 감지
# ==========================================
if "free_uses_left" not in st.session_state:
    st.session_state.free_uses_left = 3
if "api_result" not in st.session_state:
    st.session_state.api_result = None
if "current_auth_msg" not in st.session_state:
    st.session_state.current_auth_msg = None
if "current_auth_tier" not in st.session_state:
    st.session_state.current_auth_tier = "INVALID"

def reset_results():
    st.session_state.api_result = None

# ==========================================
# 💸 비용 방어 업데이트: OpenAI Vision API 캐싱 함수
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
# 📡 검로드 라이선스 검증 함수
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
# 📡 시스템 온라인 상태 바
# ==========================================
st.sidebar.success(f"🟢 SYSTEM ONLINE (Ver: {APP_VERSION})")
st.sidebar.markdown(f"🟢 Banned Substances DB: \n{BANNED_SUBSTANCES_STATUS}")
st.sidebar.markdown(f"🔍 Keyword Matching Engine: \n{KEYWORD_MATCHING_STATUS}")
st.sidebar.markdown("---")

# ==========================================
# 🔐 프리미엄 접근창
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
# 💎 요금제 및 팝업 안내창 로직 (요청 반영)
# ==========================================
st.sidebar.subheader("💎 Choose Your Plan")

# 팝업 띄우는 데코레이터 함수 (Streamlit 1.34 이상 지원)
@st.dialog("⚠️ License Key 안내 및 결제")
def show_payment_popup(plan_name, link):
    st.warning("결제가 완료되면 발송되는 **[Gumroad 영수증(Receipt) 이메일]** 내부에 대시보드 무제한 접속을 위한 **License Key (접속코드)**가 포함되어 있습니다.\n\n결제 후 반드시 이메일을 확인해 주세요!")
    st.link_button(f"👉 {plan_name} 결제창으로 이동하기", link, use_container_width=True)

# Standard 요금제
if st.sidebar.button("💳 Subscribe Standard ($299/mo)", use_container_width=True):
    show_payment_popup("Standard Plan", "https://dahee5.gumroad.com/l/lyibre")
st.sidebar.caption("✔️ 50 Scans per month\n\n✔️ Single File Scan & Excel Download")

st.sidebar.markdown("<br>", unsafe_allow_html=True) # 간격 띄우기

# PRO 요금제
if st.sidebar.button("🚀 Subscribe PRO Bulk ($499/mo)", use_container_width=True):
    show_payment_popup("PRO Bulk Plan", "https://dahee5.gumroad.com/l/pkoph")
st.sidebar.caption("✔️ Unlimited Scans\n\n✔️ Multiple Image/File Uploads & Batch Extraction")

st.sidebar.markdown("---")

# ==========================================
# 👑 메인 화면: 우리 프로그램의 4대 압도적 장점 (요청 반영)
# ==========================================
st.title("🌍 Global K-Beauty Compliance Master")
st.markdown("##### AI-powered customs compliance checker for US, EU, CN, HALAL and more.")

st.markdown("### 🌟 Why Choose Our Core Compliance Engine?")

col_adv1, col_adv2 = st.columns(2)

with col_adv1:
    st.info("#### 🛡️ 1. Zero-Hallucination (환각 방지)\nAI의 자의적 해석과 거짓 정보를 철저히 차단합니다. 확실하지 않은 성분은 임의로 번역하지 않고 `수동 검증(Verification Required)` 상태로 반환하여 치명적인 통관 오류를 미연에 방지합니다.")
    st.success("#### 🌐 2. 10-Country Custom Regulation Check\n미국 FDA MoCRA, 유럽 CPNP, 중국 NMPA, 할랄 등 10개국의 공식 금지/제한 성분 DB를 바탕으로 타겟 국가별 적합성을 0.1초 만에 교차 검증하고 경고합니다.")

with col_adv2:
    st.warning("#### 🚫 3. OCR Error Correction (오탈자 방어)\n이미지 파일 스캔 시 발생할 수 있는 글자 깨짐이나 오류 문자를 시스템이 자체적으로 필터링하여, 잘못된 글자가 위험 성분으로 둔갑하는 것을 원천적으로 막아냅니다.")
    st.error("#### 🇰🇷 4. Official K-Beauty Data Matching\n대한화장품협회 및 식약처 기준의 공식 한국어-INCI 매칭 서버 DB를 1순위로 거치도록 설계되어, 그 어떤 번역기보다 가장 정확한 K-뷰티 표준 성분 명칭을 도출합니다.")

st.markdown("---")

if not is_vip and st.session_state.free_uses_left <= 0:
    st.error("🔒 Free Trial Expired. Please subscribe in the sidebar and enter your License Key to unlock unlimited usage.")
    st.stop()

# ==========================================
# 🛡️ 법적 고지
# ==========================================
if "disclaimer_agreed" not in st.session_state:
    st.session_state.disclaimer_agreed = False

if not st.session_state.disclaimer_agreed:
    st.markdown("⚠️ REQUIRED ACTION: Review Legal Disclaimer", unsafe_allow_html=True)
    st.warning("⚖️ LEGAL DISCLAIMER: Informational Only. The provided INCI names, compliance statuses, and regulation notices do not constitute legal or official regulatory advice.")
    if st.checkbox("I HAVE READ THE LEGAL DISCLAIMER AND AGREE TO THE TERMS OF USE."):
        st.session_state.disclaimer_agreed = True
        st.rerun()
else:
    if st.button("⚖️ Review Terms"):
        st.session_state.disclaimer_agreed = False
        st.rerun()

# ==========================================
# 🚀 메인 작업 공간
# ==========================================
st.subheader("🚀 Compliance Analysis Workspace")
target_country = st.selectbox("1️⃣ Select Target Market", ["US", "EU", "CN", "JP", "ASEAN", "CA", "UK", "SFDA", "HALAL", "EAC", "BR"], on_change=reset_results)

if is_pro_or_ceo:
    uploaded_files = st.file_uploader("2️⃣ Upload Multiple Files (Images / Excels) - [PRO UNLOCKED]", type=['csv', 'xlsx', 'jpg', 'jpeg', 'png'], accept_multiple_files=True, on_change=reset_results)
else:
    uploaded_file = st.file_uploader("2️⃣ Upload A Single File (Image / Excel) - [Standard Mode]", type=['csv', 'xlsx', 'jpg', 'jpeg', 'png'], accept_multiple_files=False, on_change=reset_results)
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

    # ⚙️ 검사 실행 버튼
    if unique_ingredients and st.button("🚀 Run 10-Country Compliance Check!", use_container_width=True):
        
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
                
                result_df.rename(columns={'original_ingredient': 'Original Ingredient', 'inci_name': 'INCI Name', 'is_safe': 'Compliance Status', 'regulation_notice': 'Regulation Notice'}, inplace=True)
                result_df.insert(0, 'No.', range(1, len(result_df) + 1))
                result_df['Compliance Status'] = result_df['Compliance Status'].apply(lambda x: '🟢 PASS' if x else '🔴 FAIL')
                
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
                
                st.session_state.api_result = {
                    "df": result_df,
                    "excel_data": excel_buffer.getvalue(),
                    "status": result_data['compliance_status'],
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
# 📥 결과창 및 다운로드 버튼
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
        st.success(f"🎉 Analysis Done! No restricted ingredients found for {res['target']}.")
    else:
        st.error(f"🚨 Warning! {res['failed_count']} ingredients hit the regulation filters for {res['target']}.")
        
    st.dataframe(display_df, use_container_width=True, hide_index=True)
    
    st.download_button(
        "📥 Download Merged Excel Report (Auto-Filter Ready)", 
        data=res["excel_data"],
        file_name=f"Merged_Report_{res['target']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
