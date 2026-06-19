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
APP_VERSION = "v6.0.0 (License API Integration)"
BANNED_SUBSTANCES_STATUS = "June 2026 (Latest)"
KEYWORD_MATCHING_STATUS = "June 2026 (Synced)"

st.set_page_config(page_title="Global K-Beauty Compliance", page_icon="💄", layout="wide")

# ==========================================
# 🔄 세션 상태 초기화
# ==========================================
if "free_uses_left" not in st.session_state:
    st.session_state.free_uses_left = 3
if "api_result" not in st.session_state:
    st.session_state.api_result = None

# ==========================================
# 📡 검로드 라이선스 검증 함수 (핵심 추가 포인트)
# ==========================================
def check_license_status(key, increment=False):
    # 1. 대표님 마스터 키 (제한 없음)
    if key == "q1w2e3r41@3":
        return "PRO", "👑 CEO Master Key Active! (Unlimited)"

    verify_url = "https://api.gumroad.com/v2/licenses/verify"
    
    # 2. PRO 요금제 확인 (무제한)
    try:
        res_pro = requests.post(verify_url, data={
            "product_permalink": "pkoph",  # PRO Bulk 요금제 고유 링크
            "license_key": key,
            "increment_uses_count": "false" # PRO는 무제한이므로 횟수 카운트 생략
        }).json()
        
        # 결제 완료 상태이고 환불된 적이 없는지 검증
        if res_pro.get("success") and not res_pro.get("purchase", {}).get("refunded"):
            return "PRO", "🏆 PRO Bulk Access Granted! (Unlimited)"
    except:
        pass
        
    # 3. Standard 요금제 확인 (월 50회 제한)
    try:
        res_std = requests.post(verify_url, data={
            "product_permalink": "lyibre", # Standard 요금제 고유 링크
            "license_key": key,
            "increment_uses_count": "true" if increment else "false" # 검사 버튼 누를 때만 카운트!
        }).json()
        
        if res_std.get("success") and not res_std.get("purchase", {}).get("refunded"):
            uses = res_std.get("uses", 0)
            if uses >= 50:
                return "EXPIRED", "🚫 Monthly limit (50) reached. Please upgrade to PRO."
            else:
                return "STANDARD", f"🔓 Standard Access Granted! (Remaining: {50 - uses}/50)"
    except:
        pass
        
    return "INVALID", "❌ Invalid or Refunded License Key."

# ==========================================
# 📡 시스템 온라인 상태 바
# ==========================================
st.sidebar.success(f"🟢 SYSTEM ONLINE (Ver: {APP_VERSION})")
st.sidebar.markdown(f"❌ Banned Substances DB: \n{BANNED_SUBSTANCES_STATUS}")
st.sidebar.markdown(f"🔍 Keyword Matching Engine: \n{KEYWORD_MATCHING_STATUS}")
st.sidebar.markdown("---")

# ==========================================
# 🔐 프리미엄 접근창
# ==========================================
st.sidebar.subheader("🔐 Premium Access")
entered_password = st.sidebar.text_input("Enter your License Key:", type="password")

is_vip = False
is_pro_or_ceo = False

# 키워드 입력 시 화면단 인증 (횟수 차감 안 됨)
if entered_password:
    auth_tier, auth_msg = check_license_status(entered_password, increment=False)
    
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
    else:
        st.sidebar.error(auth_msg)
else:
    st.sidebar.warning(f"🎁 Free Trial Active: {st.session_state.free_uses_left} checks remaining.")

# ==========================================
# 💎 요금제 마케팅 배너
# ==========================================
st.sidebar.markdown("---")
st.sidebar.subheader("💎 Choose Your Plan")
st.sidebar.markdown(" Standard Plan ($299/mo) \n* 50 Scans per month \n* Single File Scan")
st.sidebar.link_button("💳 Subscribe Standard", "https://dahee5.gumroad.com/l/lyibre", use_container_width=True)

st.sidebar.markdown(" 🏆 PRO Bulk Plan ($499/mo) \n* Unlimited Scans \n* Multiple Image/File Uploads")
st.sidebar.link_button("🚀 Subscribe PRO Bulk", "https://dahee5.gumroad.com/l/pkoph", use_container_width=True)
st.sidebar.markdown("---")
st.sidebar.markdown("💬 Need Help or Found a Bug? ")
st.sidebar.markdown("📧 Contact Support")

# ==========================================
# 👑 메인 화면 소개
# ==========================================
st.title("🌍 Global K-Beauty Compliance Master")
st.markdown("##### AI-powered customs compliance checker for US, EU, CN, HALAL and more.")

st.markdown("### 💎 Multi-Tier Pricing & Plan Benefits")
col_plan1, col_plan2 = st.columns(2)
with col_plan1:
    st.info("🔹 Standard Plan ($299/mo)\n* 50 Compliance Scans per month\n* Single image or file upload at a time\n* Instant translation & cross-check\n* Excel report download")
with col_plan2:
    st.success("🏆 PRO Bulk Plan ($499/mo) \n* Unlimited Scans & multiple uploads at once\n* Automated batch extraction & matrix analyzer\n* Merged Master Excel Report for all files\n* Priority developer support")
st.markdown("---")

# ==========================================
# 🛑 철통 방어선 (무료 차단)
# ==========================================
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
    st.warning("⚖️ LEGAL DISCLAIMER: Informational Only. The provided INCI names, compliance statuses, and regulation notices do not constitute legal or official regulatory advice. Users must independently verify all data.")
    if st.checkbox("I HAVE READ THE LEGAL DISCLAIMER AND AGREE TO THE TERMS OF USE."):
        st.session_state.disclaimer_agreed = True
        st.rerun()
else:
    if st.button("⚖️ Review Terms"):
        st.session_state.disclaimer_agreed = False
        st.rerun()
        st.markdown("", unsafe_allow_html=True)

# ==========================================
# 🚀 메인 작업 공간 (진짜 검사 엔진)
# ==========================================
st.subheader("🚀 Compliance Analysis Workspace")
target_country = st.selectbox("1️⃣ Select Target Market", ["US", "EU", "CN", "JP", "ASEAN", "CA", "UK", "SFDA", "HALAL", "EAC", "BR"])

if is_pro_or_ceo:
    uploaded_files = st.file_uploader("2️⃣ Upload Multiple Files (Images / Excels) - [PRO UNLOCKED]", type=['csv', 'xlsx', 'jpg', 'jpeg', 'png'], accept_multiple_files=True)
else:
    uploaded_file = st.file_uploader("2️⃣ Upload A Single File (Image / Excel) - [Standard Mode]", type=['csv', 'xlsx', 'jpg', 'jpeg', 'png'], accept_multiple_files=False)
    uploaded_files = [uploaded_file] if uploaded_file is not None else []

ingredient_source_map = {}

if len(uploaded_files) > 0:
    for f in uploaded_files:
        file_extension = f.name.split('.')[-1].lower()
        
        if file_extension in ['jpg', 'jpeg', 'png']:
            st.image(f, caption=f"Uploaded: {f.name}", width=180)
            st.warning(f"🤖 AI Scanning [{f.name}]... (5~10 sec)")
            base64_image = base64.b64encode(f.getvalue()).decode('utf-8')
            
            vision_prompt = "Extract the ingredient names in order and output them separated by a vertical bar (|). [STRICT INSTRUCTION]: Do not use commas for separation. Example format: 정제수|글리세린|1,2-헥산디올"
            try:
                response = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[{"role": "user", "content": [{"type": "text", "text": vision_prompt}, {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}}]}],
                    temperature=0.0
                )
                extracted_list = [ing.strip() for ing in response.choices[0].message.content.split('|') if ing.strip()]
                
                for ing in extracted_list:
                    if ing not in ingredient_source_map:
                        ingredient_source_map[ing] = set()
                    ingredient_source_map[ing].add(f.name)
                
                st.markdown(f"##### 📝 Extracted List from {f.name}")
                st.dataframe(pd.DataFrame({"No.": range(1, len(extracted_list)+1), "Ingredient": extracted_list}), hide_index=True, use_container_width=True)
            except Exception as e:
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
        
        # 🌟 [요금제 횟수 차감 타이밍] 진짜 검사 버튼을 눌렀을 때만 검로드 API에 차감 승인 요청!
        if is_vip and not is_pro_or_ceo:
            run_tier, run_msg = check_license_status(entered_password, increment=True)
            if run_tier == "EXPIRED":
                st.error("🚫 Monthly limit reached. Cannot proceed with analysis. Please upgrade to PRO.")
                st.stop()
                
        status_text = st.empty()
        progress_bar = st.progress(0)
        
        status_text.warning("⏳ Step 1/4: Structuring pure cosmetic ingredients & mapping sources...")
        progress_bar.progress(25)
        time.sleep(1.5)
        
        status_text.warning(f"🕵️‍♂️ Step 2/4: Connecting and searching [{target_country}] customs database...")
        progress_bar.progress(50)
        time.sleep(2.0)
        
        status_text.warning("🧬 Step 3/4: Cross-examining INCI chemical identifiers against restriction matrix...")
        progress_bar.progress(75)
        time.sleep(2.0)
        
        status_text.info("📊 Step 4/4: Generating auto-formatted compliance sheets...")
        progress_bar.progress(95)
        
        try:
            api_response = requests.post(API_URL, json={"ingredients": unique_ingredients, "target": target_country})
            
            if api_response.status_code == 200:
                result_data = api_response.json()
                result_df = pd.DataFrame(result_data['report_details'])
                
                # 기본 컬럼 정리
                result_df.rename(columns={'original_ingredient': 'Original Ingredient', 'inci_name': 'INCI Name', 'is_safe': 'Compliance Status', 'regulation_notice': 'Regulation Notice'}, inplace=True)
                result_df.insert(0, 'No.', range(1, len(result_df) + 1))
                result_df['Compliance Status'] = result_df['Compliance Status'].apply(lambda x: '🟢 PASS' if x else '🔴 FAIL')
                
                # 출처 파일(Source File) 컬럼 추가 매핑 
                result_df['Source File'] = result_df['Original Ingredient'].apply(
                    lambda x: ", ".join(list(ingredient_source_map.get(x, [])))
                )
                
                result_df = result_df[['No.', 'Original Ingredient', 'INCI Name', 'Compliance Status', 'Source File', 'Regulation Notice']]
                
                # 🌟 [V5.6 핵심 기능] 엑셀 다운로드용 데이터 행 분리(explode) 최적화
                df_excel = result_df.copy()
                df_excel['Source File'] = df_excel['Source File'].str.split(', ')
                df_excel = df_excel.explode('Source File').reset_index(drop=True)
                df_excel['No.'] = range(1, len(df_excel) + 1)
                
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df_excel.to_excel(writer, index=False, sheet_name='Compliance_Report')
                
                # 메모리 저장
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
                
                # 무료 버전 차감
                if not is_vip:
                    st.session_state.free_uses_left -= 1
                
                st.rerun()
                
        except Exception as e:
            status_text.empty()
            progress_bar.empty()
            st.error("API Connection Failed.")

# ==========================================
# 📥 결과창 및 다운로드 버튼 현출 (화면 다중 필터 포함)
# ==========================================
if st.session_state.api_result is not None:
    res = st.session_state.api_result
    st.markdown("---")
    
    # 🌟 대시보드 화면용 파일별 조회 필터
    all_sources = list({src for sublist in res['df']['Source File'].str.split(', ') for src in sublist})
    selected_files = st.multiselect("📂 Filter by Source File (Dashboard View)", options=all_sources)
    
    if selected_files:
        display_df = res['df'][res['df']['Source File'].apply(lambda x: any(f in x for f in selected_files))]
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
