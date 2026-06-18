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
APP_VERSION = "v4.8.0 (3 Free Trials)"
BANNED_SUBSTANCES_STATUS = "June 2026 (Latest)"
KEYWORD_MATCHING_STATUS = "June 2026 (Synced)"

st.set_page_config(page_title="Global K-Beauty Compliance", page_icon="💄", layout="wide")

# ==========================================
# 🔄 세션 상태 초기화 (무료 이용 횟수 추적용 - 대표님 지시사항 반영: 3회)
# ==========================================
if "free_uses_left" not in st.session_state:
    st.session_state.free_uses_left = 3  # 👉 무료 3회 제공!

# ==========================================
# 📡 시스템 온라인 상태 바 (사이드바 최상단 고정)
# ==========================================
st.sidebar.success(f"🟢 **SYSTEM ONLINE** (Ver: {APP_VERSION})")
st.sidebar.markdown(f"❌ **Banned Substances DB:**\n{BANNED_SUBSTANCES_STATUS}")
st.sidebar.markdown(f"🔍 **Keyword Matching Engine:**\n{KEYWORD_MATCHING_STATUS}")
st.sidebar.markdown("---")

# ==========================================
# 🔐 프리미엄 접근창 (결제 연동 완료)
# ==========================================
VALID_PASSWORDS = ["VIP-KBEAUTY-2026", "TEST-CEO-1234"]
st.sidebar.subheader("🔐 Premium Access")
entered_password = st.sidebar.text_input("Enter your Access Code:", type="password")

# VIP 인증 여부 판정
is_vip = entered_password in VALID_PASSWORDS

st.sidebar.markdown("---")
st.sidebar.subheader("🔥 Premium Features")
st.sidebar.write("✔️ AI Photo Scanner (Vision OCR)")
st.sidebar.write("✔️ Auto-Formatted Excel Reports")

# 👉 대표님의 검로드 결제 링크
st.sidebar.link_button("💳 Subscribe Now ($299/mo)", "https://dahee5.gumroad.com/l/lyibre", use_container_width=True)

# ==========================================
# 📞 고객 센터 / 피드백 창구 (사이드바 하단)
# ==========================================
st.sidebar.markdown("---")
st.sidebar.markdown("💬 **Need Help or Found a Bug?**")
st.sidebar.markdown("📧 [Contact Support](mailto:simsuk95126@gmail.com)")
st.sidebar.markdown("💡 [Request a New Feature](https://forms.google.com/)")


# ==========================================
# 👑 플랫폼 마케팅 문구
# ==========================================
st.title("🌍 Global K-Beauty Compliance Master")
st.markdown("##### AI-powered customs compliance checker for US, EU, CN, HALAL and more.")

st.markdown("### 💎 Why Choose Our Platform?")
st.markdown(f"• 🔄 **Real-Time Regulatory DB:** Directly synced with official global prohibited items ({BANNED_SUBSTANCES_STATUS}) to pre-emptively block custom clearance rejection.")
st.markdown("• 📸 **AI Vision OCR (No Typing):** Just upload a product back-photo. Our strict Vision AI reads and extracts Korean ingredient text instantly with zero hallucinations.")
st.markdown("• 🌍 **10-Country Cross Check:** Simultaneously cross-examine ingredients against US FDA (MoCRA), EU CPNP, China NMPA, and strict HALAL standards in 1 second.")
st.markdown("---")


# ==========================================
# 🛑 철통 방어선 (무료 기회 3회를 다 썼을 때 차단)
# ==========================================
if not is_vip and st.session_state.free_uses_left <= 0:
    st.error("🔒 **Free Trial Expired.** You have used all 3 free compliance checks. Please subscribe in the sidebar and enter your VIP Access Code to unlock unlimited usage.")
    st.stop()  # 👉 여기서 코드 렌더링을 완전히 멈춥니다! 얄짤없이 결제 유도.

# 상태 알림창 표시 (사이드바)
if is_vip:
    st.sidebar.success("🔓 VIP Access Granted! Unlimited Checks Enrolled.")
else:
    st.sidebar.warning(f"🎁 Free Trial Active: {st.session_state.free_uses_left} checks remaining.")


# ==========================================
# 🛡️ 법적 고지 (동의해야만 밑에 화면이 열림)
# ==========================================
if "disclaimer_agreed" not in st.session_state:
    st.session_state.disclaimer_agreed = False

if not st.session_state.disclaimer_agreed:
    st.markdown("<h2 style='color: #d9534f;'>⚠️ REQUIRED ACTION: Review Legal Disclaimer</h2>", unsafe_allow_html=True)
    
    st.warning("""
    ⚖️ LEGAL DISCLAIMER & LIMITATION OF LIABILITY
    1. Informational Only: The provided INCI names, compliance statuses, and regulation notices do not constitute legal, medical, or official regulatory advice.
    2. No Liability: Under no circumstances shall the API provider be liable for any direct, indirect, incidental, or consequential damages arising from the use of this tool.
    3. User Responsibility: Users must independently verify all data with certified regulatory professionals before commercial application.
    """)
    
    st.markdown("""
        <style>
        .big-font { font-size:24px !important; font-weight: bold !important; color: #111111 !important; }
        </style>
        """, unsafe_allow_html=True)
    
    st.markdown('<p class="big-font">👇 Check the box below to agree and unlock the system:</p>', unsafe_allow_html=True)
    
    if st.checkbox("I HAVE READ THE LEGAL DISCLAIMER AND AGREE TO THE TERMS OF USE."):
        st.session_state.disclaimer_agreed = True
        st.rerun()

else:
    col_status, col_reset = st.columns([5, 1])
    with col_status:
        st.write("✅ *Legal Disclaimer Accepted. Analysis Engine Unlocked.*")
    with col_reset:
        if st.button("⚖️ Review Terms", use_container_width=True):
            st.session_state.disclaimer_agreed = False
            st.rerun()
            
    st.markdown("<br>", unsafe_allow_html=True)

    # ==========================================
    # 🚀 메인 작업 공간 (진짜 검사 엔진)
    # ==========================================
    st.subheader("🚀 Compliance Analysis Workspace")
    target_country = st.selectbox("1️⃣ Select Target Market", ["US", "EU", "CN", "JP", "ASEAN", "CA", "UK", "SFDA", "HALAL", "EAC", "BR"])
    uploaded_file = st.file_uploader("2️⃣ Upload File (Image / Excel / CSV)", type=['csv', 'xlsx', 'jpg', 'jpeg', 'png'])

    if uploaded_file is not None:
        ingredient_list = []
        file_extension = uploaded_file.name.split('.')[-1].lower()

        if file_extension in ['jpg', 'jpeg', 'png']:
            st.image(uploaded_file, caption="Uploaded Image", width=220)
            st.warning("🤖 AI Scanner is extracting ingredients... (5~10 sec)")
            base64_image = base64.b64encode(uploaded_file.getvalue()).decode('utf-8')
            vision_prompt = "Extract the ingredient names in order and output them separated by commas (,). [STRICT INSTRUCTION]: Do not guess. If blurry, output '[Unreadable]'."
            
            try:
                response = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[{"role": "user", "content": [{"type": "text", "text": vision_prompt}, {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}}]}],
                    temperature=0.0
                )
                ingredient_list = [ing.strip() for ing in response.choices[0].message.content.split(',')]
                st.success("✅ Extraction Complete!")
            except Exception as e:
                st.error("Image scan failed.")
        
        else:
            df = pd.read_csv(uploaded_file) if file_extension == 'csv' else pd.read_excel(uploaded_file)
            st.dataframe(df.head(2))
            ingredient_list = df.iloc[:, 0].dropna().astype(str).tolist()

        if ingredient_list and st.button("🚀 Run 10-Country Compliance Check!", use_container_width=True):
            with st.spinner(f"Searching [{target_country}] customs database... 🕵️‍♂️"):
                try:
                    api_response = requests.post(API_URL, json={"ingredients": ingredient_list, "target": target_country})
                    if api_response.status_code == 200:
                        result_data = api_response.json()
                        result_df = pd.DataFrame(result_data['report_details'])
                        result_df.rename(columns={'original_ingredient': 'Original Ingredient', 'inci_name': 'INCI Name', 'is_safe': 'Compliance Status', 'regulation_notice': 'Regulation Notice'}, inplace=True)
                        result_df.insert(0, 'No.', range(1, len(result_df) + 1))
                        result_df['Compliance Status'] = result_df['Compliance Status'].apply(lambda x: '🟢 PASS' if x else '🔴 FAIL')
                        
                        st.markdown("---")
                        if result_data['compliance_status'] == "PASS":
                            st.success(f"🎉 Perfect! No restricted ingredients found for {target_country}.")
                        else:
                            st.error(f"🚨 Warning! {result_data['failed_count']} ingredients hit the regulation filters for {target_country}.")
                        
                        st.dataframe(result_df, use_container_width=True, hide_index=True)

                        excel_buffer = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            result_df.to_excel(writer, index=False, sheet_name='Compliance_Report')
                        st.download_button("📥 Download Excel Report", data=excel_buffer.getvalue(), file_name=f"Report_{target_country}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                        
                        # ⚠️ [핵심 자동화] 유효한 비밀번호가 없다면 무료 횟수 차감!
                        if not is_vip:
                            st.session_state.free_uses_left -= 1
                            st.rerun()  # 검사 횟수가 차감되면 화면을 바로 새로고침하여 횟수를 깎아버림!
                            
                except Exception as e:
                    st.error("API Connection Failed.")
