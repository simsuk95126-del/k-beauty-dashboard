import streamlit as st
import pandas as pd
import requests
import base64
import os
import io
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()
client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
API_URL = "https://k-beauty-api.onrender.com/api/v1/compliance-report"

# 🛠️ 실시간 업데이트 및 버전 관리 변수
APP_VERSION = "v5.2.0 (PRO Link Fully Integrated)"
BANNED_SUBSTANCES_STATUS = "June 2026 (Latest)"
KEYWORD_MATCHING_STATUS = "June 2026 (Synced)"

st.set_page_config(page_title="Global K-Beauty Compliance", page_icon="💄", layout="wide")

# ==========================================
# 🔄 세션 상태 초기화 (무료 3회 맛보기)
# ==========================================
if "free_uses_left" not in st.session_state:
    st.session_state.free_uses_left = 3

# ==========================================
# 📡 시스템 온라인 상태 바 (사이드바 최상단 고정)
# ==========================================
st.sidebar.success(f"🟢 **SYSTEM ONLINE** (Ver: {APP_VERSION})")
st.sidebar.markdown(f"❌ **Banned Substances DB:**\n{BANNED_SUBSTANCES_STATUS}")
st.sidebar.markdown(f"🔍 **Keyword Matching Engine:**\n{KEYWORD_MATCHING_STATUS}")
st.sidebar.markdown("---")

# ==========================================
# 🔐 프리미엄 접근창 (CEO 마스터키 및 플랜 비밀번호)
# ==========================================
# VIP-KBEAUTY-2026: 스탠다드 플랜 ($299)
# PRO-BULK-9988: 프로 대량스캔 플랜 ($499)
# q1w2e3r41@3: 대표님 전용 평생 무제한 마스터키 🔥
VALID_PASSWORDS = ["VIP-KBEAUTY-2026", "PRO-BULK-9988", "q1w2e3r41@3"]

st.sidebar.subheader("🔐 Premium Access")
entered_password = st.sidebar.text_input("Enter your Access Code:", type="password")

# 등급 판정 로직
is_vip = entered_password in VALID_PASSWORDS
is_pro_or_ceo = entered_password in ["PRO-BULK-9988", "q1w2e3r41@3"]

st.sidebar.markdown("---")
st.sidebar.subheader("💎 Choose Your Plan")

# 💳 스탠다드 요금제 버튼 ($299)
st.sidebar.markdown("**Standard Plan ($299/mo)**\nSingle File Scan")
st.sidebar.link_button("💳 Subscribe Standard", "https://dahee5.gumroad.com/l/lyibre", use_container_width=True)

st.sidebar.markdown("<br>", unsafe_allow_html=True)

# 💳 🏆 PRO 대량스캔 요금제 버튼 ($499) - 대표님이 주신 새 주소 완벽 연동! 🔥
st.sidebar.markdown("**🏆 PRO Bulk Plan ($499/mo)**\nUnlimited Multiple Image/File Uploads")
st.sidebar.link_button("🚀 Subscribe PRO Bulk", "https://dahee5.gumroad.com/l/pkoph", use_container_width=True)

# ==========================================
# 📞 고객 지원 창구
# ==========================================
st.sidebar.markdown("---")
st.sidebar.markdown("💬 **Need Help or Found a Bug?**")
st.sidebar.markdown("📧 [Contact Support](mailto:simsuk95126@gmail.com)")


# ==========================================
# 👑 플랫폼 마케팅 문구
# ==========================================
st.title("🌍 Global K-Beauty Compliance Master")
st.markdown("##### AI-powered customs compliance checker for US, EU, CN, HALAL and more.")

st.markdown("### 💎 Multi-Tier Pricing & Plan Benefits")
col_plan1, col_plan2 = st.columns(2)
with col_plan1:
    st.info("🔹 **Standard Plan ($299/mo)**\n* 3 Free Trial Uses\n* Single image or file upload at a time\n* Instant translation & cross-check\n* Excel report download")
with col_plan2:
    st.success("🏆 **PRO Bulk Plan ($499/mo)**\n* **Unlimited multiple image/file uploads** at once\n* Automated batch extraction & matrix analyzer\n* **Merged Master Excel Report** for all files\n* Priority developer support")
st.markdown("---")


# ==========================================
# 🛑 철통 방어선 (비밀번호 없는 일반 유저가 3회 다 썼을 때 차단)
# ==========================================
if not is_vip and st.session_state.free_uses_left <= 0:
    st.error("🔒 **Free Trial Expired.** You have used all 3 free compliance checks. Please subscribe in the sidebar and enter your VIP/PRO Access Code to unlock unlimited usage.")
    st.stop()  # 웹사이트 가동 중단 및 차단!

if is_vip:
    if entered_password == "q1w2e3r41@3":
        st.sidebar.success("👑 **CEO Master Key Active! (Unlimited PRO Access)**")
    elif entered_password == "PRO-BULK-9988":
        st.sidebar.success("🏆 **PRO Bulk Member Access Granted!**")
    else:
        st.sidebar.success("🔓 Standard VIP Access Granted!")
else:
    st.sidebar.warning(f"🎁 Free Trial Active: {st.session_state.free_uses_left} checks remaining.")


# ==========================================
# 🛡️ 법적 고지 (Disclaimer)
# ==========================================
if "disclaimer_agreed" not in st.session_state:
    st.session_state.disclaimer_agreed = False

if not st.session_state.disclaimer_agreed:
    st.markdown("<h2 style='color: #d9534f;'>⚠️ REQUIRED ACTION: Review Legal Disclaimer</h2>", unsafe_allow_html=True)
    st.warning("⚖️ LEGAL DISCLAIMER: Informational Only. The provided INCI names, compliance statuses, and regulation notices do not constitute legal or official regulatory advice. Users must independently verify all data.")
    if st.checkbox("I HAVE READ THE LEGAL DISCLAIMER AND AGREE TO THE TERMS OF USE."):
        st.session_state.disclaimer_agreed = True
        st.rerun()
else:
    if st.button("⚖️ Review Terms"):
        st.session_state.disclaimer_agreed = False
        st.rerun()
            
    st.markdown("<br>", unsafe_allow_html=True)

    # ==========================================
    # 🚀 메인 작업 공간 (진짜 검사 엔진)
    # ==========================================
    st.subheader("🚀 Compliance Analysis Workspace")
    target_country = st.selectbox("1️⃣ Select Target Market", ["US", "EU", "CN", "JP", "ASEAN", "CA", "UK", "SFDA", "HALAL", "EAC", "BR"])
    
    # ⚙️ 프로 등급 및 관리자만 다중 파일 업로드 허용 활성화
    if is_pro_or_ceo:
        uploaded_files = st.file_uploader("2️⃣ Upload Multiple Files (Images / Excels) - [PRO UNLOCKED]", type=['csv', 'xlsx', 'jpg', 'jpeg', 'png'], accept_multiple_files=True)
    else:
        uploaded_file = st.file_uploader("2️⃣ Upload A Single File (Image / Excel) - [Standard Mode]", type=['csv', 'xlsx', 'jpg', 'jpeg', 'png'], accept_multiple_files=False)
        uploaded_files = [uploaded_file] if uploaded_file is not None else []

    all_ingredients = []
    
    if len(uploaded_files) > 0:
        for f in uploaded_files:
            file_extension = f.name.split('.')[-1].lower()
            
            if file_extension in ['jpg', 'jpeg', 'png']:
                st.image(f, caption=f"Uploaded: {f.name}", width=180)
                st.warning(f"🤖 AI Scanning [{f.name}]... (5~10 sec)")
                base64_image = base64.b64encode(f.getvalue()).decode('utf-8')
                vision_prompt = "Extract the ingredient names in order and output them separated by commas (,). Do not guess."
                
                try:
                    response = client.chat.completions.create(
                        model="gpt-4o",
                        messages=[{"role": "user", "content": [{"type": "text", "text": vision_prompt}, {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}}]}],
                        temperature=0.0
                    )
                    extracted_list = [ing.strip() for ing in response.choices[0].message.content.split(',') if ing.strip()]
                    all_ingredients.extend(extracted_list)
                    
                    # 📝 추출 성분 표 형태 리스트업 복구본 작동
                    st.markdown(f"##### 📝 Extracted List from {f.name}")
                    st.dataframe(pd.DataFrame({"No.": range(1, len(extracted_list)+1), "Ingredient": extracted_list}), hide_index=True, use_container_width=True)
                except Exception as e:
                    st.error(f"Image scan failed for {f.name}")
            else:
                df = pd.read_csv(f) if file_extension == 'csv' else pd.read_excel(f)
                st.markdown(f"📊 Preview: {f.name}")
                st.dataframe(df.head(2))
                all_ingredients.extend(df.iloc[:, 0].dropna().astype(str).tolist())

        all_ingredients = list(set(all_ingredients))

        if all_ingredients and st.button("🚀 Run 10-Country Compliance Check!", use_container_width=True):
            with st.spinner(f"Searching [{target_country}] customs database... 🕵️‍♂️"):
                try:
                    api_response = requests.post(API_URL, json={"ingredients": all_ingredients, "target": target_country})
                    if api_response.status_code == 200:
                        result_data = api_response.json()
                        result_df = pd.DataFrame(result_data['report_details'])
                        result_df.rename(columns={'original_ingredient': 'Original Ingredient', 'inci_name': 'INCI Name', 'is_safe': 'Compliance Status', 'regulation_notice': 'Regulation Notice'}, inplace=True)
                        result_df.insert(0, 'No.', range(1, len(result_df) + 1))
                        result_df['Compliance Status'] = result_df['Compliance Status'].apply(lambda x: '🟢 PASS' if x else '🔴 FAIL')
                        
                        st.markdown("---")
                        if result_data['compliance_status'] == "PASS":
                            st.success(f"🎉 Analysis Done! No restricted ingredients found for {target_country}.")
                        else:
                            st.error(f"🚨 Warning! {result_data['failed_count']} ingredients hit the regulation filters.")
                        
                        st.dataframe(result_df, use_container_width=True, hide_index=True)

                        excel_buffer = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            result_df.to_excel(writer, index=False, sheet_name='Compliance_Report')
                        st.download_button("📥 Download Merged Excel Report", data=excel_buffer.getvalue(), file_name=f"Merged_Report_{target_country}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                        
                        # 무료 횟수 소진 처리 (마스터키 및 구독자는 제외)
                        if not is_vip:
                            st.session_state.free_uses_left -= 1
                            st.rerun()
                            
                except Exception as e:
                    st.error("API Connection Failed.")
