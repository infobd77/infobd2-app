import streamlit as st
import requests
import xml.etree.ElementTree as ET
import pandas as pd
from io import BytesIO
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
import xlsxwriter
from urllib.parse import quote_plus
import time
import urllib3
import datetime
import random 
import folium
from streamlit_folium import st_folium

# SSL 경고 비활성화
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# =========================================================
# [설정] UI 및 스타일
# =========================================================
st.set_page_config(page_title="부동산 원클릭 분석 Pro", page_icon="🏢", layout="centered")

st.markdown("""
    <style>
        .block-container {
            max-width: 1000px; 
            padding-top: 3rem; 
            padding-bottom: 2rem;
            padding-left: 2rem;
            padding-right: 2rem;
        }
        
        button[data-testid="stNumberInputStepDown"],
        button[data-testid="stNumberInputStepUp"] { display: none !important; }
        .stNumberInput label { display: none; }
        
        input[type="text"] { 
            text-align: right !important; 
            font-size: 24px !important; 
            font-weight: 800 !important;
            font-family: 'Pretendard', sans-serif;
            color: #333 !important;
            padding-right: 10px !important;
        }

        div[data-testid="stTextInput"] input[aria-label="주소 입력"] {
            text-align: left !important;
            font-size: 18px !important;
            font-weight: 600 !important;
        }
        
        div[data-testid="stTextInput"] input[aria-label="공시지가"] {
            text-align: center !important;
            font-size: 20px !important;
            color: #1a237e !important;
        }
        div[data-testid="stTextInput"] input[aria-label="용도지역"] {
            text-align: center !important;
            font-size: 20px !important;
            color: #1a237e !important;
        }

        input[aria-label="매매금액"] {
             color: #D32F2F !important; 
             font-size: 32px !important; 
        }

        .stButton > button {
            width: 100%;
            background-color: #1a237e;
            color: white;
            font-size: 18px;
            font-weight: bold;
            padding: 14px;
            border-radius: 8px;
            border: none;
            box-shadow: 0 4px 6px rgba(0,0,0,0.2);
            transition: all 0.3s;
        }
        .stButton > button:hover {
            background-color: #0d47a1;
            transform: translateY(-2px);
        }
        
        .unit-price-box {
            background-color: #f5f5f5;
            border: 1px solid #e0e0e0;
            padding: 8px;
            border-radius: 8px;
            margin-top: 10px;
            text-align: center;
        }
        .unit-price-value {
            font-size: 22px; 
            font-weight: 800; 
            color: #111;
        }
        
        .ai-summary-box {
            background-color: #fff;
            border: 1px solid #ddd;
            border-top: 4px solid #1a237e;
            padding: 30px;
            border-radius: 5px;
            margin-top: 20px;
            text-align: left;
            box-shadow: 0 10px 25px rgba(0,0,0,0.08);
        }
        .ai-title {
            font-size: 24px;
            font-weight: 800;
            color: #1a237e;
            margin-bottom: 25px;
            border-bottom: 2px solid #eee;
            padding-bottom: 15px;
            letter-spacing: -0.5px;
        }
        .insight-item {
            margin-bottom: 18px;
            font-size: 17px;
            line-height: 1.7;
            color: #424242;
        }
        
        .link-btn {
            display: inline-block;
            width: 100%;
            padding: 10px;
            margin: 5px 0;
            text-align: center;
            border-radius: 5px;
            text-decoration: none;
            font-weight: bold;
            color: white !important;
            transition: 0.3s;
        }
        .naver-btn { background-color: #03C75A; }
        .eum-btn { background-color: #1a237e; }
        .naver-btn:hover, .eum-btn:hover { opacity: 0.8; }
        
        .selected-tags {
            background-color: #e3f2fd;
            color: #1565c0;
            padding: 6px 12px;
            border-radius: 20px;
            font-size: 14px;
            font-weight: 700;
            margin-right: 6px;
            display: inline-block;
            margin-bottom: 6px;
            border: 1px solid #bbdefb;
        }
    </style>
    """, unsafe_allow_html=True)

# =========================================================
# [설정] 인증키 및 전역 변수
# =========================================================
USER_KEY = "Xl5W1ALUkfEhomDR8CBUoqBMRXphLTIB7CuTto0mjsg0CQQspd7oUEmAwmw724YtkjnV05tdEx6y4yQJCe3W0g=="
VWORLD_KEY = "47B30ADD-AECB-38F3-B5B4-DD92CCA756C5"

if 'zoning' not in st.session_state: st.session_state['zoning'] = ""
if 'generated_insights' not in st.session_state: st.session_state['generated_insights'] = [] # 생성된 후보군
if 'final_selected_insights' not in st.session_state: st.session_state['final_selected_insights'] = [] # 사용자가 최종 선택한 것
if 'price' not in st.session_state: st.session_state['price'] = 0
if 'addr' not in st.session_state: st.session_state['addr'] = "" 
if 'last_click_lat' not in st.session_state: st.session_state['last_click_lat'] = 0.0
if 'fetched_lp' not in st.session_state: st.session_state['fetched_lp'] = 0
if 'fetched_zoning' not in st.session_state: st.session_state['fetched_zoning'] = ""

def reset_analysis():
    st.session_state['generated_insights'] = []
    st.session_state['final_selected_insights'] = []
    st.session_state['fetched_lp'] = 0
    st.session_state['fetched_zoning'] = ""

# --- [API 및 보조 함수] ---
def get_address_from_coords(lat, lng):
    url = "https://api.vworld.kr/req/address" 
    params = {
        "service": "address", "request": "getaddress", "version": "2.0", "crs": "EPSG:4326",
        "point": f"{lng},{lat}", "type": "PARCEL", "format": "json", "errorformat": "json", "key": VWORLD_KEY
    }
    try:
        response = requests.get(url, params=params, timeout=5, verify=False)
        data = response.json()
        if data.get('response', {}).get('status') == 'OK':
            return data['response']['result'][0]['text']
    except: return None
    return None

def render_styled_block(label, value, is_area=False):
    st.markdown(f"""
    <div style="margin-bottom: 10px;">
        <div style="font-size: 16px; color: #666; font-weight: 600; margin-bottom: 2px;">{label}</div>
        <div style="font-size: 24px; font-weight: 800; color: #111; line-height: 1.2;">{value}</div>
    </div>
    """, unsafe_allow_html=True)

def comma_input(label, unit, key, default_val, help_text=""):
    st.markdown(f"""
        <div style='font-size: 16px; font-weight: 700; color: #333; margin-bottom: 4px;'>
            {label} <span style='font-size:12px; color:#888; font-weight:400;'>{help_text}</span>
        </div>
    """, unsafe_allow_html=True)
    c_in, c_unit = st.columns([3, 1]) 
    with c_in:
        if key not in st.session_state: st.session_state[key] = default_val
        current_val = st.session_state[key]
        formatted_val = f"{current_val:,}" if current_val != 0 else ""
        val_input = st.text_input(label, value=formatted_val, key=f"{key}_widget", label_visibility="hidden")
        try:
            if val_input.strip() == "": new_val = 0
            else: new_val = int(str(val_input).replace(',', '').strip())
            st.session_state[key] = new_val
        except: new_val = 0
    with c_unit:
        st.markdown(f"<div style='margin-top: 15px; font-size: 18px; font-weight: 600; color: #555;'>{unit}</div>", unsafe_allow_html=True)
    return new_val

def format_date_dot(date_str):
    if not date_str or len(date_str) != 8: return date_str
    return f"{date_str[:4]}.{date_str[4:6]}.{date_str[6:]}"

def format_area_html(val_str):
    try:
        val = float(val_str)
        if val == 0: return "-"
        pyung = val * 0.3025
        return f"{val:,.2f}㎡<br><span style='color: #E53935;'>({pyung:,.1f}평)</span>"
    except: return "-"

def format_area_ppt(val_str):
    try:
        val = float(val_str)
        if val == 0: return "-"
        pyung = val * 0.3025
        return f"{val:,.2f}㎡ ({pyung:,.1f}평)"
    except: return "-"

# --- [AI 인사이트 생성 (전문가/마케팅)] ---
def generate_insight_candidates(info, finance, zoning, env_features, user_comment, comp_df=None, target_dong=""):
    points = []
    
    # [마케팅 DB: 이모티콘 제거, 전문 용어 사용]
    marketing_db = {
        "역세권": [
            "■ [초역세권 입지] 지하철역 도보권 내 위치하여 풍부한 유동인구와 직장인 출퇴근 수요를 독점하는 핵심 입지",
            "☑ [교통 허브 프리미엄] 대중교통 접근성이 탁월하여 사옥 및 임대용으로 공실 리스크가 극히 낮은 안전 자산",
            "■ [환금성 우수] 경기 변동에도 흔들리지 않는 역세권 불패 신화를 입증하는 탄탄한 수요층 보유"
        ],
        "더블역세권": [
            "■ [더블 역세권의 희소성] 2개 노선이 교차하는 환승 거점으로 광역 수요까지 흡수하는 최상급 교통 요지",
            "☑ [황금 노선 크로스] 서울 주요 업무지구로의 이동이 자유로워 기업체 사옥 임차 수요가 끊이지 않는 곳",
            "■ [광역 접근성] 가시성과 접근성을 동시에 만족시키는 더블 역세권 프리미엄으로 자산 가치 상승 주도"
        ],
        "대로변": [
            "■ [가시성 최상급] 차량 및 보행자 노출도가 압도적인 대로변에 위치하여 기업 홍보 및 옥외 광고 효과 극대화",
            "☑ [Trophy Asset] 웅장한 전면 효과로 기업의 브랜드 가치를 높여줄 수 있는 랜드마크급 사옥 부지",
            "■ [상징적 입지] 누구나 쉽게 찾아올 수 있는 대로변 입지로 병의원 및 대형 프랜차이즈 입점 최적화"
        ],
        "코너입지": [
            "■ [3면 개방 코너] 가시성과 접근성이 탁월한 코너 건물로 전 층 채광 및 자연 광고 효과 우수",
            "☑ [S급 상권 요지] 양방향 도로에 접해 있어 차량 진입과 보행자 유입이 수월한 코너 입지 프리미엄",
            "■ [개방형 설계 유리] 코너 입지의 장점을 살린 건축 설계로 임차인 선호도가 매우 높은 희소 매물"
        ],
        "이면코너": [
            "■ [알짜배기 이면 코너] 메인 도로의 소음은 피하고 접근성은 확보한 실속형 사옥 및 고급 F&B 상권",
            "☑ [가성비 투자처] 대로변 대비 합리적인 평단가로 매입하여 높은 임대 수익률을 기대할 수 있는 코너 입지",
            "■ [특화 상권 형성] 아늑하고 프라이빗한 분위기를 선호하는 트렌디한 리테일 및 카페 입점 유리"
        ],
        "학군지": [
            "■ [명문 학군 수요] 대치/목동급 학원가 수요를 배후에 둔 안정적인 교육 특화 상권으로 공실 우려 없음",
            "☑ [항시 유동인구] 학생 및 학부모 유동인구가 365일 끊이지 않는 항아리 상권 내 알짜 빌딩",
            "■ [소비력 갖춘 배후] 우수한 학군을 찾아 유입되는 고소득 배후 세대를 바탕으로 탄탄한 임대 수익 확보"
        ],
        "먹자상권": [
            "■ [24시간 활성 상권] 점심부터 저녁 회식까지 유동인구가 끊이지 않는 메인 먹자골목 핵심 입지",
            "☑ [권리금 형성 지역] 매출이 검증된 바닥 권리금 형성 지역으로 임차 대기 수요가 풍부한 안정적 투자처",
            "■ [복합 소비 상권] 직장인 및 거주민이 어우러진 소비 상권으로 경기 불황에도 강한 면모 과시"
        ],
        "항아리상권": [
            "■ [항아리 상권의 강점] 외부 유출 없이 내부 대단지 배후 수요가 꽉 갇혀 있는 독점적 상권 입지",
            "☑ [생활 밀착형 수익] 병원, 학원, 마트 등 필수 근생 시설 입점에 최적화되어 안정적인 장기 임대 가능",
            "■ [충성 고객 확보] 한번 유입되면 단골이 되는 항아리 상권 특성상 매출 변동성이 적어 임차인 선호도 높음"
        ],
        "오피스상권": [
            "■ [탄탄한 직장인 수요] 오피스 밀집 지역으로 구매력 높은 직장인 수요가 365일 뒷받침되는 오피스 상권",
            "☑ [B2B 임차 수요] 주변 대기업 및 협력사들의 사무실 수요로 공실 걱정 없는 비즈니스 핵심 입지",
            "■ [업무 인프라 완비] 은행, 관공서, 비즈니스 센터 등 업무 지원 시설이 풍부하여 법인 사옥으로 최적"
        ],
        "신축/리모델링": [
            "■ [신축급 컨디션] 최신 트렌드를 반영한 수려한 내외관으로 추가 비용 없이 즉시 임대 수익 실현 가능",
            "☑ [유지보수 비용 절감] 신축 또는 리모델링 완료되어 시설물 관리가 용이하고 운영 비용이 최소화된 매물",
            "■ [우량 임차인 유치] 깔끔한 내외관과 최신 설비(엘리베이터 등)를 갖추어 병원, IT 기업 등 우량 임차 유리"
        ],
        "신축빌딩": [
            "■ [지역 내 랜드마크] 최신 건축 공법과 세련된 디자인으로 지역 내 독보적인 존재감을 뽐내는 신축 빌딩",
            "☑ [희소성 높은 신축] 노후 건물이 많은 지역 내 단비 같은 신축 매물로 임대 경쟁력 절대 우위 선점",
            "■ [자산 가치 프리미엄] 신축 프리미엄으로 향후 매각 시 높은 시세 차익(Capital Gain) 기대 가능"
        ],
        "급매물": [
            "■ [시세 파괴 초급매] 소유주 사정으로 시세 대비 현저히 저렴하게 진행되는 다시 없을 기회의 매물",
            "☑ [확실한 안전 마진] 주변 시세 대비 낮은 평단가로 매입 즉시 시세 차익 효과를 누릴 수 있는 알짜 자산",
            "■ [즉시 매수 추천] 가격 메리트가 확실하여 시장에 나오자마자 거래될 확률이 높은 A급 급매물"
        ],
        "사옥추천": [
            "■ [최적의 사옥 스펙] 쾌적한 업무 환경과 넉넉한 주차 공간, 효율적인 레이아웃을 갖춘 사옥 맞춤형 건물",
            "☑ [기업 브랜딩 효과] 세련된 외관과 우수한 가시성으로 기업의 아이덴티티를 높여줄 수 있는 사옥 입지",
            "■ [임직원 만족도] 대중교통 접근성이 좋고 주변 편의시설이 풍부하여 임직원 근무 만족도가 높은 곳"
        ],
        "메디컬입지": [
            "■ [메디컬 전문 빌딩] 병의원 개원에 필요한 엘리베이터, 전력, 주차 등 하드웨어 스펙이 완벽한 곳",
            "☑ [처방전 수요 독점] 약국 입점이 가능한 입지로 고수익 임대료 창출 및 건물의 가치를 높이는 매물",
            "■ [닥터 선호도 1위] 배후 세대가 탄탄하고 가시성이 좋아 피부과, 내과, 치과 등 개원 문의가 많은 입지"
        ],
        "밸류업유망": [
            "■ [밸류업(Value-up) 최적] 리모델링 또는 신축 시 용적률 이득과 임대료 상승 여력이 확실한 원석 매물",
            "☑ [자산 가치 상승] 현재 저평가되어 있으나 적극적인 MD 구성과 리노베이션으로 가치를 극대화할 수 있는 곳",
            "■ [디벨로퍼 추천] 명도가 용이하고 대지 형상이 우수하여 신축 부지로 활용 시 개발 이익 극대화 예상"
        ],
        "주차편리": [
            "■ [주차 스트레스 Zero] 강남권에서 보기 드문 넉넉한 주차 공간 확보로 임차인 및 내방객 만족도 최상",
            "☑ [자주식 주차] 기계식 주차의 불편함 없는 편리한 자주식 주차 공간으로 대형 차량 진입도 수월"
        ],
        "명도협의가능": [
            "■ [즉시 명도 가능] 매수 후 즉시 리모델링이나 신축이 가능하도록 명도 협의가 완료된 깔끔한 매물",
            "☑ [실사용자 추천] 복잡한 명도 과정 없이 바로 입주하여 사용할 수 있어 실사용 목적 매수자에게 최적"
        ],
        "수익형": [
            "■ [고수익 수익형] 탄탄한 임차 구성을 바탕으로 매월 안정적인 현금 흐름이 발생하는 알짜 수익형 부동산",
            "☑ [공실률 제로 도전] 우수한 입지와 합리적인 임대료 책정으로 공실 위험 없이 꾸준한 수익 창출 가능"
        ],
        "관리상태최상": [
            "■ [주인 직영 관리] 건물주가 직접 거주하며 꼼꼼하게 관리하여 내외관 컨디션이 신축급으로 유지된 건물",
            "☑ [손볼 곳 없음] 누수나 하자 없이 완벽하게 관리되어 매수 후 추가적인 유지보수 비용이 들지 않는 매물"
        ],
        "숲세권": [
            "■ [쾌적한 숲세권] 도심 속에서 자연을 느낄 수 있는 쾌적한 업무 환경으로 창의적인 업무 능률 향상",
            "☑ [힐링 오피스] 인근 공원 및 녹지와 인접하여 산책과 휴식이 가능한 워라밸 최적화 입지"
        ]
    }
    
    # 1. 사용자 코멘트
    if user_comment:
        points.append(f"📌 {user_comment.strip()}")

    # 2. 키워드 기반 (랜덤)
    if env_features:
        # 키워드별로 하나씩 다 가져와서 풍부하게 만듦
        random.shuffle(env_features)
        for feat in env_features:
            if feat in marketing_db:
                # 해당 키워드의 문구 중 하나 랜덤 선택
                points.append(random.choice(marketing_db[feat]))

    # 3. 가격 경쟁력 (전문 용어 사용)
    if comp_df is not None and not comp_df.empty:
        try:
            sold_df = comp_df[comp_df['구분'].astype(str).str.contains('매각|완료|매매', na=False)]
            if not sold_df.empty:
                avg_price = sold_df['평당가'].mean()
                my_price = finance['land_pyeong_price_val']
                diff = my_price - avg_price
                diff_pct = abs(diff / avg_price) * 100
                loc_text = target_dong if target_dong else "인근"
                if diff < 0:
                    msgs = [
                        f"■ [가격 경쟁력 우위] {loc_text} 실거래 평균(평 {avg_price:,.0f}만) 대비 {diff_pct:.1f}% 저렴한 확실한 저평가 매물",
                        f"☑ [안전 마진 확보] 주변 시세보다 합리적인 가격으로 진입하여 매입 즉시 시세 차익(Capital Gain) 기대"
                    ]
                    points.append(random.choice(msgs))
                else:
                    msgs = [
                        f"■ [프리미엄 입증] {loc_text} 평균을 상회하나, 독보적인 입지와 {zoning} 용적률 이점을 감안하면 합리적 가치",
                        f"☑ [대장주 입지] 단순 시세 비교 불가! 압도적인 컨디션과 입지로 지역 내 시세를 리딩하는 Trophy Asset"
                    ]
                    points.append(random.choice(msgs))
        except: pass

    # 4. 수익률 (Cash Flow 강조)
    yield_val = finance['yield']
    if yield_val >= 4.0:
        msgs = [
            f"■ [고수익 Cash Flow] 연 {yield_val:.1f}%의 놀라운 수익률로 고금리 기조에도 이자 비용을 상회하는 효자 상품",
            f"☑ [수익률 최상급] 강남권에서 보기 드문 {yield_val:.1f}%대 고수익 매물로 안정적인 현금 흐름(Cash Flow) 창출"
        ]
        points.append(random.choice(msgs))
    elif yield_val >= 3.0:
        msgs = [
            f"■ [안정적 포트폴리오] 연 {yield_val:.1f}%의 꾸준한 임대 수익(Income Gain)과 지가 상승(Capital Gain)을 동시에 추구",
            f"☑ [리스크 헷지] 공실 걱정 없는 탄탄한 입지에서 나오는 연 {yield_val:.1f}%의 안정적인 운용 수익률 확보"
        ]
        points.append(random.choice(msgs))
    else:
        msgs = [
            f"■ [미래 가치 투자] 당장의 수익률보다는 향후 개발 호재와 지가 상승에 베팅하는 공격적인 SI(Strategic Investor) 투자처",
            f"☑ [시세 차익형 자산] 보유할수록 가치가 오르는 토지 가치에 집중할 수 있는 인플레이션 헷지 수단"
        ]
        points.append(random.choice(msgs))

    # 부족하면 채워넣기
    fallback_msgs = [
        "■ [희소 가치] 매물 잠김이 심한 지역 내 오랜만에 등장한 A급 매물로 빠른 소진 예상",
        "☑ [성공 투자 요건] 입지, 가격, 상권 3박자를 모두 갖춘 보기 드문 육각형 매물",
        "■ [불패 입지] 한번 들어오면 나가지 않는 임차인들이 선호하는 검증된 명당 자리"
    ]
    
    random.shuffle(fallback_msgs)
    for msg in fallback_msgs:
        points.append(msg)
        
    # 중복 제거 및 리스트 반환
    return list(dict.fromkeys(points))

# --- [API 조회 함수들] ---
@st.cache_data(show_spinner=False)
def get_pnu_and_coords(address):
    url = "http://api.vworld.kr/req/search"
    search_type = 'road' if '로' in address or '길' in address else 'parcel'
    params = {"service": "search", "request": "search", "version": "2.0", "crs": "EPSG:4326", "size": "1", "page": "1", "query": address, "type": "address", "category": search_type, "format": "json", "errorformat": "json", "key": VWORLD_KEY}
    try:
        res = requests.get(url, params=params, timeout=3)
        data = res.json()
        if data['response']['status'] == 'NOT_FOUND':
            params['query'] = "서울특별시 " + address
            res = requests.get(url, params=params, timeout=3)
            data = res.json()
        if data['response']['status'] == 'NOT_FOUND': return None
        item = data['response']['result']['items'][0]
        pnu = item.get('address', {}).get('pnu') or item.get('id')
        lng = float(item['point']['x']); lat = float(item['point']['y'])
        full_address = item.get('address', {}).get('parcel', '') 
        if not full_address: full_address = item.get('address', {}).get('road', '') 
        if not full_address: full_address = address
        return {"pnu": pnu, "lat": lat, "lng": lng, "full_addr": full_address}
    except: return None

@st.cache_data(show_spinner=False)
def get_zoning_smart(lat, lng):
    url = "http://api.vworld.kr/req/data"
    delta = 0.0005
    min_x, min_y = lng - delta, lat - delta
    max_x, max_y = lng + delta, lat + delta
    params = {"service": "data", "request": "GetFeature", "data": "LT_C_UQ111", "key": VWORLD_KEY, "format": "json", "size": "10", "geomFilter": f"BOX({min_x},{min_y},{max_x},{max_y})", "domain": "localhost"}
    try:
        res = requests.get(url, params=params, timeout=3, verify=False)
        if res.status_code == 200:
            data = res.json()
            features = data.get('response', {}).get('result', {}).get('featureCollection', {}).get('features', [])
            if features:
                zonings = [f['properties']['UNAME'] for f in features]
                return ", ".join(sorted(list(set(zonings))))
    except: pass
    return ""

@st.cache_data(show_spinner=False)
def get_land_price(pnu):
    url = "http://apis.data.go.kr/1611000/NsdiIndvdLandPriceService/getIndvdLandPriceAttr"
    current_year = datetime.datetime.now().year
    years_to_check = range(current_year, current_year - 7, -1) 
    for year in years_to_check:
        params = {"serviceKey": USER_KEY, "pnu": pnu, "format": "xml", "numOfRows": "1", "pageNo": "1", "stdrYear": str(year)}
        try:
            res = requests.get(url, params=params, timeout=4)
            if res.status_code == 200:
                root = ET.fromstring(res.content)
                if root.findtext('.//resultCode') == '00':
                    price_node = root.find('.//indvdLandPrice')
                    if price_node is not None and price_node.text: return int(price_node.text)
        except: continue
        time.sleep(0.05)
    return 0

@st.cache_data(show_spinner=False)
def get_building_info_smart(pnu):
    base_url = "https://apis.data.go.kr/1613000/BldRgstHubService/getBrTitleInfo"
    sigungu = pnu[0:5]; bjdong = pnu[5:10]; bun = pnu[11:15]; ji = pnu[15:19]
    plat_code = '1' if pnu[10] == '2' else '0'
    params = {"serviceKey": USER_KEY, "sigunguCd": sigungu, "bjdongCd": bjdong, "platGbCd": plat_code, "bun": bun, "ji": ji, "numOfRows": "1", "pageNo": "1"}
    try:
        res = requests.get(base_url, params=params, timeout=5, verify=False)
        if res.status_code == 200: return parse_xml_response(res.content)
        return {"error": f"서버 상태: {res.status_code}"}
    except Exception as e: return {"error": str(e)}

def parse_xml_response(content):
    try:
        root = ET.fromstring(content)
        item = root.find('.//item')
        if item is None: return None
        
        indr_mech = int(item.findtext('indrMechUtcnt', '0') or 0)
        indr_auto = int(item.findtext('indrAutoUtcnt', '0') or 0)
        total_indoor = indr_mech + indr_auto
        oudr_mech = int(item.findtext('oudrMechUtcnt', '0') or 0)
        oudr_auto = int(item.findtext('oudrAutoUtcnt', '0') or 0)
        total_outdoor = oudr_mech + oudr_auto
        total_parking = total_indoor + total_outdoor
        parking_str = f"{total_parking}대(옥내{total_indoor}/옥외{total_outdoor})"
        ride_elvt = int(item.findtext('rideUseElvtCnt', '0') or 0)
        emgen_elvt = int(item.findtext('emgenUseElvtCnt', '0') or 0)
        total_elvt = ride_elvt + emgen_elvt
        elvt_str = f"{total_elvt}대"
        
        return {
            "bldNm": item.findtext('bldNm', '-'),
            "mainPurpsCdNm": item.findtext('mainPurpsCdNm', '정보없음'),
            "strctCdNm": item.findtext('strctCdNm', '정보없음'),
            "platArea": float(item.findtext('platArea', '0') or 0),
            "totArea": float(item.findtext('totArea', '0') or 0),
            "platArea_html": format_area_html(item.findtext('platArea', '0')),
            "totArea_html": format_area_html(item.findtext('totArea', '0')),
            "archArea_html": format_area_html(item.findtext('archArea', '0')),
            "groundArea_html": format_area_html(item.findtext('vlRatEstmTotArea', '0')),
            "platArea_ppt": format_area_ppt(item.findtext('platArea', '0')),
            "totArea_ppt": format_area_ppt(item.findtext('totArea', '0')),
            "archArea_ppt": format_area_ppt(item.findtext('archArea', '0')),
            "archArea_val": float(item.findtext('archArea', '0') or 0),
            "groundArea": float(item.findtext('vlRatEstmTotArea', '0') or 0),
            "groundArea_ppt": format_area_ppt(item.findtext('vlRatEstmTotArea', '0')),
            "ugrndFlrCnt": item.findtext('ugrndFlrCnt', '0'),
            "grndFlrCnt": item.findtext('grndFlrCnt', '0'),
            "useAprDay": format_date_dot(item.findtext('useAprDay', '')),
            "bcRat": float(item.findtext('bcRat', '0') or 0),
            "vlRat": float(item.findtext('vlRat', '0') or 0),
            "rideUseElvtCnt": elvt_str,
            "parking": parking_str
        }
    except Exception as e: return {"error": str(e)}

@st.cache_data(show_spinner=False)
def get_cadastral_map_image(lat, lng):
    delta = 0.0015 
    minx, miny = lng - delta, lat - delta
    maxx, maxy = lng + delta, lat + delta
    bbox = f"{minx},{miny},{maxx},{maxy}"
    layer = "LP_PA_CBND_BUBUN"
    url = f"https://api.vworld.kr/req/wms?SERVICE=WMS&REQUEST=GetMap&VERSION=1.3.0&LAYERS={layer}&STYLES={layer}&CRS=EPSG:4326&BBOX={bbox}&WIDTH=400&HEIGHT=300&FORMAT=image/png&TRANSPARENT=FALSE&BGCOLOR=0xFFFFFF&EXCEPTIONS=text/xml&KEY={VWORLD_KEY}"
    headers = {"User-Agent": "Mozilla/5.0", "Referer": "http://localhost:8501"}
    try:
        res = requests.get(url, headers=headers, timeout=5, verify=False)
        if res.status_code == 200 and 'image' in res.headers.get('Content-Type', ''): return BytesIO(res.content)
    except: pass
    return None

@st.cache_data(show_spinner=False)
def get_static_map_image(lat, lng):
    url = f"http://api.vworld.kr/req/image?service=image&request=getmap&key={VWORLD_KEY}&center={lng},{lat}&crs=EPSG:4326&zoom=17&size=600,400&format=png&basemap=GRAPHIC"
    try:
        res = requests.get(url, timeout=3)
        if res.status_code == 200 and 'image' in res.headers.get('Content-Type', ''): return BytesIO(res.content)
    except: pass
    return None

# [PPT 생성 함수]
def create_pptx(info, full_addr, finance, zoning, lat, lng, land_price, selling_points, images_dict, template_binary=None):
    bld_name = info.get('bldNm')
    if not bld_name or bld_name == '-':
        dong = full_addr.split(' ')[2] if len(full_addr.split(' ')) > 2 else ""
        bld_name = f"{dong} 빌딩" if dong else "사옥용 빌딩"
        
    # [요청 1] 공시지가: 웹 계산 후 "만원/평" 표기
    lp_py_val = (land_price / 10000) / 0.3025 if land_price > 0 else 0
    lp_str_final = f"{lp_py_val:,.0f}만원/평"
    
    # [요청 2] 공시지가 총액: 소수점 제외, 앞에 "합 " 붙이기
    total_lp_val = land_price * info['platArea'] if land_price and info['platArea'] else 0
    total_lp_num = int(total_lp_val / 100000000) if total_lp_val > 0 else 0
    total_lp_str_final = f"합 {total_lp_num:,}억" if total_lp_num > 0 else "-"

    ai_points_str = "\n".join(selling_points[:5]) if selling_points else "분석된 특징이 없습니다."

    plat_m2 = f"{info['platArea']:,}" if info['platArea'] else "-"
    plat_py = f"{info['platArea'] * 0.3025:,.1f}" if info['platArea'] else "-"
    tot_m2 = f"{info['totArea']:,}" if info['totArea'] else "-"
    tot_py = f"{info['totArea'] * 0.3025:,.1f}" if info['totArea'] else "-"
    
    arch_val = info.get('archArea_val', 0)
    if arch_val == 0 and info['platArea'] > 0 and info['bcRat'] > 0:
        arch_val = info['platArea'] * (info['bcRat'] / 100)
    arch_m2 = f"{arch_val:,.1f}"
    arch_py = f"{arch_val * 0.3025:,.1f}"
    
    ground_val = info.get('groundArea', 0)
    if ground_val == 0 and info['totArea'] > 0:
            ground_val = info['totArea']
    ground_m2 = f"{ground_val:,}"
    ground_py = f"{ground_val * 0.3025:,.1f}"
    
    use_date = info.get('useAprDay', '-')

    # [요청 3] 매매평단가 앞에 "평 " 붙이기
    market_price_py_val = finance.get('land_pyeong_price_val', 0)
    market_price_str = f"평 {market_price_py_val:,.0f}만원"

    ctx_vals = {
        'plat_m2': plat_m2, 'plat_py': plat_py,
        'tot_m2': tot_m2, 'tot_py': tot_py,
        'arch_m2': arch_m2, 'arch_py': arch_py,
        'ground_m2': ground_m2, 'ground_py': ground_py,
        'use_date': use_date
    }

    # 1. 템플릿이 있는 경우
    if template_binary:
        prs = Presentation(template_binary)
        
        deep_blue = RGBColor(0, 51, 153) 
        deep_red = RGBColor(204, 0, 0)   
        black = RGBColor(0, 0, 0)
        gray_border = RGBColor(128, 128, 128)
        dark_gray_border = RGBColor(80, 80, 80)

        data_map = {
            "{{빌딩이름}}": bld_name,
            "{{소재지}}": full_addr,
            "{{용도지역}}": zoning,
            "{{AI물건분석내용 4가지 }}": ai_points_str,
            "{{공시지가}}": lp_str_final,
            "{{공시지가 총액}}": total_lp_str_final,
            "{{준공년도}}": use_date,
            "{{건물규모}}": f"B{info.get('ugrndFlrCnt')} / {info.get('grndFlrCnt')}F",
            "{{건폐율}}": f"{info.get('bcRat', 0)}%",
            "{{용적률}}": f"{info.get('vlRat', 0)}%",
            "{{승강기}}": info.get('rideUseElvtCnt', '-'),
            "{{주차대수}}": info.get('parking', '-'),
            "{{건물주구조}}": info.get('strctCdNm', '-'),
            "{{건물용도}}": info.get('mainPurpsCdNm', '-'),
            "{{보증금}}": f"{finance['deposit']:,} 만원" if finance['deposit'] else "-",
            "{{월임대료}}": f"{finance['rent']:,} 만원" if finance['rent'] else "-",
            "{{관리비}}": f"{finance['maintenance']:,} 만원" if finance['maintenance'] else "-",
            "{{수익률}}": f"년 {finance['yield']:.1f}%" if finance['yield'] else "-",
            "{{융자금}}": f"{finance['loan']:,} 억원" if finance['loan'] else "-",
            "{{매매금액}}": f"{finance['price']:,} 억원" if finance['price'] else "-",
            "{{대지평단가}}": market_price_str,
            "{{건물미래가치 활용도}}": "사옥 및 수익용 리모델링 추천",
            "{{위치도}}": "", 
            "{{지적도}}": "",
            "{{건축물대장}}": "",
            "{{건물사진}}": ""
        }

        def replace_text_in_frame(text_frame, mapper, ctx):
            for p in text_frame.paragraphs:
                p_text = p.text
                
                # [요청 1] 공시지가: 검정, 굵게
                if "{{공시지가}}" in p_text:
                    p.text = str(mapper["{{공시지가}}"])
                    for r in p.runs: r.font.color.rgb = black; r.font.bold = True; r.font.size = Pt(12)
                    return
                # [요청 2] 공시지가 총액: 빨강, 굵게
                if "{{공시지가 총액}}" in p_text:
                    p.text = str(mapper["{{공시지가 총액}}"])
                    for r in p.runs: r.font.color.rgb = deep_red; r.font.bold = True; r.font.size = Pt(12)
                    return

                financial_keys = ["{{보증금}}", "{{월임대료}}", "{{관리비}}", "{{융자금}}"]
                found_fin_key = None
                for k in financial_keys:
                    if k in p_text: found_fin_key = k; break
                
                if found_fin_key:
                    val_str = str(mapper[found_fin_key])
                    if " " in val_str:
                        num_part, unit_part = val_str.split(' ', 1)
                        p.text = "" 
                        run_num = p.add_run(); run_num.text = num_part + " "; run_num.font.size = Pt(12); run_num.font.bold = True; run_num.font.color.rgb = black
                        run_unit = p.add_run(); run_unit.text = unit_part; run_unit.font.size = Pt(10); run_unit.font.bold = True; run_unit.font.color.rgb = black
                    else:
                        p.text = val_str
                        for r in p.runs: r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = black
                    return 

                if "{{매매금액}}" in p_text:
                    val_str = str(mapper["{{매매금액}}"])
                    if " " in val_str:
                        num, unit = val_str.split(' ', 1)
                        p.text = ""
                        run_n = p.add_run(); run_n.text = num + " "; run_n.font.size = Pt(16); run_n.font.bold = True; run_n.font.color.rgb = deep_blue
                        run_u = p.add_run(); run_u.text = unit; run_u.font.size = Pt(10); run_u.font.bold = True; run_u.font.color.rgb = deep_blue
                    else:
                        p.text = val_str
                        for r in p.runs: r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = deep_blue
                    continue

                if "{{대지면적}}" in p_text:
                    if "평" in p_text:
                        p.text = p_text.replace("{{대지면적}}", ctx['plat_py'])
                        for r in p.runs: r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = deep_blue
                    else:
                        p.text = p_text.replace("{{대지면적}}", ctx['plat_m2'])
                        for r in p.runs: r.font.size = Pt(10)
                elif "{{연면적}}" in p_text:
                    if "평" in p_text:
                        p.text = p_text.replace("{{연면적}}", ctx['tot_py'])
                        for r in p.runs: r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = deep_blue
                    else:
                        p.text = p_text.replace("{{연면적}}", ctx['tot_m2'])
                        for r in p.runs: r.font.size = Pt(10)
                elif "{{건축면적}}" in p_text:
                    if "평" in p_text:
                        p.text = p_text.replace("{{건축면적}}", ctx['arch_py'])
                        for r in p.runs: r.font.size = Pt(10)
                    else:
                        p.text = p_text.replace("{{건축면적}}", ctx['arch_m2'])
                        for r in p.runs: r.font.size = Pt(10)
                elif "{{지상면적}}" in p_text:
                    if "평" in p_text:
                        p.text = p_text.replace("{{지상면적}}", ctx['ground_py'])
                        for r in p.runs: r.font.size = Pt(10)
                    else:
                        p.text = p_text.replace("{{지상면적}}", ctx['ground_m2'])
                        for r in p.runs: r.font.size = Pt(10)
                elif "{{준공년도}}" in p_text:
                    new_text = p_text.replace("{{준공년도}}", ctx['use_date'])
                    if ctx['use_date'] + "㎡" in new_text: new_text = new_text.replace("㎡", "")
                    p.text = new_text
                    for r in p.runs: r.font.size = Pt(10)
                else:
                    found_key = None
                    for k in mapper.keys():
                        if k in p_text: found_key = k; break
                    if found_key:
                        val = str(mapper[found_key])
                        p.text = p_text.replace(found_key, val)
                        for r in p.runs:
                            r.font.size = Pt(10)
                            if found_key == "{{빌딩이름}}": r.font.size = Pt(25); r.font.bold = True
                            elif found_key == "{{수익률}}": r.font.size = Pt(12); r.font.color.rgb = deep_red; r.font.bold = True
                            elif found_key == "{{대지평단가}}": r.font.size = Pt(10); r.font.color.rgb = deep_blue; r.font.bold = True

        def replace_text_in_shape(shape, mapper, ctx):
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                for child_shape in shape.shapes:
                    replace_text_in_shape(child_shape, mapper, ctx)
                return
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if cell.text_frame:
                            replace_text_in_frame(cell.text_frame, mapper, ctx)
                return
            if shape.has_text_frame:
                replace_text_in_frame(shape.text_frame, mapper, ctx)
        
        for slide in prs.slides:
            for shape in slide.shapes: replace_text_in_shape(shape, data_map, ctx_vals)

        # [이미지 삽입]
        img_insert_map = {
            1: ('u1', Cm(2.55), Cm(3.5), Cm(24.59), Cm(15.74)), 
            2: ('u2', Cm(1.0), Cm(3.5), Cm(13.91), Cm(10.97)), 
            4: ('u3', Cm(4.65), Cm(3.5), Cm(20.4), Cm(15.74)), 
            5: ('u4', Cm(3.36), Cm(3.5), Cm(22.97), Cm(15.74)) 
        }
        for s_idx, (key, l, t, w, h) in img_insert_map.items():
            if s_idx < len(prs.slides) and key in images_dict and images_dict[key]:
                img_file = images_dict[key]; img_file.seek(0)
                pic = prs.slides[s_idx].shapes.add_picture(img_file, l, t, width=w, height=h)
                pic.line.visible = True; pic.line.width = Pt(1.5)
                pic.line.color.rgb = dark_gray_border if s_idx == 2 else gray_border

        if 6 < len(prs.slides):
            slide7 = prs.slides[6]
            u5_keys = ['u5_1', 'u5_2', 'u5_3', 'u5_4']
            positions = [(Cm(1.0), Cm(3.5)), (Cm(15.1), Cm(3.5)), (Cm(1.0), Cm(11.75)), (Cm(15.1), Cm(11.75))]
            w_s7, h_s7 = Cm(13.6), Cm(7.75)
            for idx, u_key in enumerate(u5_keys):
                if u_key in images_dict and images_dict[u_key]:
                    f = images_dict[u_key]; f.seek(0)
                    pic = slide7.shapes.add_picture(f, positions[idx][0], positions[idx][1], width=w_s7, height=h_s7)
                    pic.line.visible = True; pic.line.width = Pt(1.0); pic.line.color.rgb = gray_border

        output = BytesIO()
        prs.save(output)
        return output.getvalue()
    
    # 2. 템플릿 없는 경우 (기본 PPT)
    else:
        prs = Presentation(); prs.slide_width = Cm(21.0); prs.slide_height = Cm(29.7)
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        title_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(1.0), Cm(1.0), Cm(19.0), Cm(2.0))
        title_box.fill.background(); title_box.line.color.rgb = RGBColor(200, 200, 200); title_box.line.width = Pt(1)
        tf = title_box.text_frame
        tf.text = bld_name; p = tf.paragraphs[0]; p.font.size = Pt(28); p.font.bold = True; p.font.name = "맑은 고딕"; p.alignment = PP_ALIGN.CENTER

        lbl_img = slide.shapes.add_textbox(Cm(1.0), Cm(2.9), Cm(9.2), Cm(0.6)); lbl_img.text_frame.text = "건물사진"; lbl_img.text_frame.paragraphs[0].font.bold=True
        main_img = images_dict.get('u2')
        if main_img:
            main_img.seek(0)
            pic = slide.shapes.add_picture(main_img, Cm(1.0), Cm(3.5), width=Cm(9.2), height=Cm(11.5))
            pic.line.visible = True; pic.line.color.rgb = RGBColor(200, 200, 200); pic.line.width = Pt(1)
        else:
            rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(1.0), Cm(3.5), Cm(9.2), Cm(11.5)); rect.fill.background(); rect.line.color.rgb = RGBColor(200, 200, 200)

        lbl_map = slide.shapes.add_textbox(Cm(1.0), Cm(15.2), Cm(9.2), Cm(0.6)); lbl_map.text_frame.text = "위치도"; lbl_map.text_frame.paragraphs[0].font.bold=True
        loc_img = images_dict.get('u1')
        if loc_img:
            loc_img.seek(0)
            pic_map = slide.shapes.add_picture(loc_img, Cm(1.0), Cm(15.8), width=Cm(9.2), height=Cm(12.0))
            pic_map.line.visible = True; pic_map.line.color.rgb = RGBColor(200, 200, 200); pic_map.line.width = Pt(1)
        else:
            map_img = get_static_map_image(lat, lng)
            if map_img: 
                pic_map = slide.shapes.add_picture(map_img, Cm(1.0), Cm(15.8), width=Cm(9.2), height=Cm(12.0))
                pic_map.line.visible = True; pic_map.line.color.rgb = RGBColor(200, 200, 200)

        lbl_tbl = slide.shapes.add_textbox(Cm(10.8), Cm(2.9), Cm(9.2), Cm(0.6)); lbl_tbl.text_frame.text = "건물개요"; lbl_tbl.text_frame.paragraphs[0].font.bold=True
        table = slide.shapes.add_table(11, 4, Cm(10.8), Cm(3.5), Cm(9.2), Cm(11.5)).table
        table.columns[0].width = Cm(2.3); table.columns[1].width = Cm(2.3); table.columns[2].width = Cm(2.3); table.columns[3].width = Cm(2.3)
        
        lp_py = (land_price / 10000) / 0.3025 if land_price > 0 else 0
        data = [
            ["소재지", full_addr, "", ""], ["용도", zoning, "공시지가", f"{lp_py:,.0f}만/평"],
            ["대지", info['platArea_ppt'], "도로", "M"], ["연면적", info['totArea_ppt'], "준공", info['useAprDay']],
            ["지상", info['totArea_ppt'], "규모", f"B{info.get('ugrndFlrCnt')}/{info.get('grndFlrCnt')}F"], ["건축", info['archArea_ppt'], "승강기", info['rideUseElvtCnt']],
            ["건/용", f"{info.get('bcRat')}%/{info.get('vlRat')}%", "주차", info['parking'].split('(')[0]], ["주용도", info.get('mainPurpsCdNm','-'), "주구조", info.get('strctCdNm','-')],
            ["보증금", f"{finance['deposit']:,.0f}만", "융자", f"{finance['loan']:,}억"], ["임대료", f"{finance['rent']:,}만", "수익률", f"{finance['yield']:.1f}%"],
            ["관리비", f"{finance['maintenance']:,}만", "매도가", f"{finance['price']:,}억"]
        ]
        for r in range(11):
            for c in range(4):
                cell = table.cell(r, c); cell.text = str(data[r][c]); cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                p = cell.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(9); p.font.name = "맑은 고딕"
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(240, 248, 255) if c % 2 == 0 else RGBColor(255, 255, 255)

        lbl_ai = slide.shapes.add_textbox(Cm(10.8), Cm(23.9), Cm(9.2), Cm(0.6)); lbl_ai.text_frame.text = "투자포인트 내용"; lbl_ai.text_frame.paragraphs[0].font.bold=True
        rect_ai = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(10.8), Cm(24.5), Cm(9.2), Cm(3.5))
        rect_ai.fill.background(); rect_ai.line.color.rgb = RGBColor(200, 200, 200)
        tx_ai = slide.shapes.add_textbox(Cm(10.9), Cm(24.6), Cm(9.0), Cm(3.3)); tf_ai = tx_ai.text_frame; tf_ai.word_wrap = True
        summary_text = ""
        if selling_points:
            for pt in selling_points[:5]: summary_text += f"{pt.strip()}\n"
        else: summary_text = "■ 역세권 입지로 투자가치 우수\n■ 안정적인 임대 수익 기대"
        tf_ai.text = summary_text; 
        for p in tf_ai.paragraphs: p.font.size = Pt(10)
        
        output = BytesIO()
        prs.save(output)
        return output.getvalue()

def create_excel(info, full_addr, finance, zoning, lat, lng, land_price, selling_points, uploaded_img):
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('부동산분석')
    
    fmt_title = workbook.add_format({'bold': True, 'font_size': 20, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#EAEAEA'})
    fmt_label = workbook.add_format({'bold': True, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#F0F8FF'}) 
    fmt_val = workbook.add_format({'bold': True, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
    fmt_val_red = workbook.add_format({'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_color': 'red'})
    fmt_box = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'top', 'text_wrap': True})
    fmt_header = workbook.add_format({'bold': True, 'font_size': 12, 'align': 'left'})

    worksheet.set_column('A:A', 2); worksheet.set_column('B:E', 12); worksheet.set_column('F:F', 2); worksheet.set_column('G:J', 12) 

    bld_name = info.get('bldNm')
    if not bld_name or bld_name == '-': bld_name = f"{full_addr.split(' ')[2] if len(full_addr.split(' ')) > 2 else ''} 빌딩"
    worksheet.merge_range('B2:J3', bld_name, fmt_title)

    worksheet.write('B5', '건물사진', fmt_header); worksheet.merge_range('B6:E20', '', fmt_box) 
    if uploaded_img: uploaded_img.seek(0); worksheet.insert_image('B6', 'building.png', {'image_data': uploaded_img, 'x_scale': 0.5, 'y_scale': 0.5, 'object_position': 2})

    worksheet.write('B22', '위치도', fmt_header); worksheet.merge_range('B23:E35', '', fmt_box)
    map_img_xls = f"http://api.vworld.kr/req/image?service=image&request=getmap&key={VWORLD_KEY}&center={lng},{lat}&crs=EPSG:4326&zoom=17&size=600,400&format=png&basemap=GRAPHIC"
    try:
        res = requests.get(map_img_xls, timeout=3)
        if res.status_code == 200: worksheet.insert_image('B23', 'map.png', {'image_data': BytesIO(res.content), 'x_scale': 0.7, 'y_scale': 0.7})
    except: pass

    worksheet.write('G5', '건물개요', fmt_header)
    lp_py = (land_price / 10000) / 0.3025 if land_price > 0 else 0
    bcvl_text = f"{info['bcRat']:.2f}%\n{info['vlRat']:.2f}%"
    
    table_data_xls = [
        ["소재지", full_addr, "용도", zoning], ["공시지가", f"{lp_py:,.0f}만/평", "대지", info['platArea_ppt']], 
        ["도로", "6M", "연면적", info['totArea_ppt']], ["준공", info['useAprDay'], "지상", info['totArea_ppt']],
        ["규모", f"B{info['ugrndFlrCnt']}/{info['grndFlrCnt']}F", "건축", info['archArea_ppt']], ["승강기", info['rideUseElvtCnt'], "건/용", bcvl_text],
        ["주차", info['parking'].split('(')[0], "주용도", info.get('mainPurpsCdNm','-')], ["주구조", info.get('strctCdNm','-'), "보증금", f"{finance['deposit']:,.0f}만"],
        ["융자", f"{finance['loan']:,}억", "임대료", f"{finance['rent']:,}만"], ["수익률", f"{finance['yield']:.1f}%", "관리비", f"{finance['maintenance']:,}만"],
        ["매도가", f"{finance['price']:,}억", "", ""] 
    ]
    for i, row in enumerate(table_data_xls):
        worksheet.write(5 + i, 6, row[0], fmt_label) 
        if row[0] == "매도가": worksheet.merge_range(5 + i, 7, 5 + i, 9, row[1], fmt_val_red)
        else: worksheet.write(5 + i, 7, row[1], fmt_val) 
        if row[0] != "매도가": worksheet.write(5 + i, 8, row[2], fmt_label); worksheet.write(5 + i, 9, row[3], fmt_val) 

    worksheet.write('G17', '지적도', fmt_header); worksheet.merge_range('G18:J26', '', fmt_box)
    cad_img = get_cadastral_map_image(lat, lng)
    if cad_img: worksheet.insert_image('G18', 'cad.png', {'image_data': cad_img, 'x_scale': 0.6, 'y_scale': 0.6})

    worksheet.write('G28', '투자포인트 내용', fmt_header); worksheet.merge_range('G29:J35', '', fmt_box)
    summary_text = ""
    if selling_points:
        for pt in selling_points[:5]: summary_text += f"{pt.strip()}\n"
    else: summary_text = "■ 역세권 입지로 투자가치 우수\n■ 안정적인 임대 수익 기대"
    worksheet.write('G29', summary_text, fmt_box)
    worksheet.merge_range('B37:J37', "JS 제이에스부동산(주) 김창익 이사 010-6595-5700", fmt_title)
    workbook.close()
    return output.getvalue()

# [메인 실행]
st.title("🏢 부동산 매입 분석기 Pro")
st.markdown("---")

with st.expander("🗺 지도에서 직접 클릭하여 찾기 (Click)", expanded=False):
    m = folium.Map(location=[37.5172, 127.0473], zoom_start=14)
    output = st_folium(m, width=700, height=400)
    if output and output.get("last_clicked"):
        lat = output["last_clicked"]["lat"]; lng = output["last_clicked"]["lng"]
        if "last_click_lat" not in st.session_state or st.session_state["last_click_lat"] != lat:
            st.session_state["last_click_lat"] = lat
            found_addr = get_address_from_coords(lat, lng)
            if found_addr:
                st.success(f"📍 지도 클릭 확인! 변환된 주소: {found_addr}")
                st.session_state['addr'] = found_addr; reset_analysis(); st.rerun()
            else: st.warning("⚠️ 주소를 찾을 수 없는 위치입니다.")

link_container = st.container()
addr_input = st.text_input("주소 입력", placeholder="예: 강남구 논현동 254-4", key="addr", on_change=reset_analysis)

if addr_input:
    with st.spinner("데이터 분석 중..."):
        location = get_pnu_and_coords(addr_input)
        if not location: st.error("❌ 주소를 찾을 수 없습니다.")
        else:
            with link_container:
                col_l1, col_l2 = st.columns(2)
                with col_l1: st.markdown(f"<a href='https://map.naver.com/v5/search/{quote_plus(location['full_addr'])}' target='_blank' class='link-btn naver-btn'>📍 네이버지도 위치확인</a>", unsafe_allow_html=True)
                with col_l2: 
                    if location.get('pnu'): st.markdown(f"<a href='https://www.eum.go.kr/web/ar/lu/luLandDet.jsp?pnu={location['pnu']}&mode=search&isNoScr=script' target='_blank' class='link-btn eum-btn'>📑 토지이음 규제정보 확인</a>", unsafe_allow_html=True)
            
            if not st.session_state['zoning']: st.session_state['zoning'] = get_zoning_smart(location['lat'], location['lng'])
            if not st.session_state['fetched_zoning']: st.session_state['fetched_zoning'] = st.session_state['zoning']

            info = get_building_info_smart(location['pnu'])
            land_price = get_land_price(location['pnu'])
            if land_price > 0 and st.session_state['fetched_lp'] == 0: st.session_state['fetched_lp'] = land_price
            
            if not info or "error" in info: st.error(f"조회 실패: {info.get('error')}")
            else:
                st.success("✅ 분석 완료!")
                
                # [요청 4] 사진 업로드 박스 4열 배치 (슬라이드 7처럼)
                st.write("##### 📸 PPT 삽입용 사진 업로드")
                
                st.write("▼ 기본 사진 (위치도/메인/지적도/대장)")
                col_u1, col_u2, col_u3, col_u4 = st.columns(4)
                with col_u1: u1 = st.file_uploader("Slide 2: 위치도", type=['png', 'jpg', 'jpeg'], key="u1")
                with col_u2: u2 = st.file_uploader("Slide 3: 건물메인", type=['png', 'jpg', 'jpeg'], key="u2")
                with col_u3: u3 = st.file_uploader("Slide 5: 지적도", type=['png', 'jpg', 'jpeg'], key="u3")
                with col_u4: u4 = st.file_uploader("Slide 6: 대장", type=['png', 'jpg', 'jpeg'], key="u4")
                
                st.write("▼ 추가 사진 (Slide 7)")
                c_u5_1, c_u5_2, c_u5_3, c_u5_4 = st.columns(4)
                with c_u5_1: u5_1 = st.file_uploader("추가1", type=['png','jpg'], key="u5_1")
                with c_u5_2: u5_2 = st.file_uploader("추가2", type=['png','jpg'], key="u5_2")
                with c_u5_3: u5_3 = st.file_uploader("추가3", type=['png','jpg'], key="u5_3")
                with c_u5_4: u5_4 = st.file_uploader("추가4", type=['png','jpg'], key="u5_4")
                
                images_map = {'u1': u1, 'u2': u2, 'u3': u3, 'u4': u4, 'u5_1': u5_1, 'u5_2': u5_2, 'u5_3': u5_3, 'u5_4': u5_4}

                st.markdown("---")
                st.markdown("""<div style="background-color: #f8f9fa; padding: 50px; border-radius: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">""", unsafe_allow_html=True)
                c1, c2 = st.columns([2, 1])
                with c1: render_styled_block("소재지", addr_input)
                with c2: render_styled_block("건물명", info.get('bldNm'))
                st.write("") 
                
                c_lp1, c_lp2, c_lp3 = st.columns(3)
                with c_lp1:
                    lp_val = st.text_input("공시지가(원/㎡)", value=f"{st.session_state['fetched_lp']:,}")
                    try: land_price = int(lp_val.replace(',', ''))
                    except: land_price = 0
                with c_lp2:
                    if land_price > 0 and info['platArea'] > 0: render_styled_block("공시지가 총액(추정)", f"{land_price * info['platArea'] / 100000000:,.2f}억")
                    else: render_styled_block("공시지가 총액", "-")
                with c_lp3: st.empty()
                st.write("")
                st.markdown("<hr style='margin: 10px 0; border-top: 1px dashed #ddd;'>", unsafe_allow_html=True)
                
                c2_1, c2_2, c2_3 = st.columns(3)
                with c2_1:
                    zoning_val = st.text_input("용도지역", value=st.session_state['fetched_zoning'])
                    st.session_state['zoning'] = zoning_val
                with c2_2: render_styled_block("대지면적", info['platArea_html'], is_area=True)
                with c2_3: render_styled_block("연면적", info['totArea_html'], is_area=True)
                st.write("")
                c3_1, c3_2, c3_3 = st.columns(3)
                with c3_1: render_styled_block("준공년도", info['useAprDay'])
                with c3_2: render_styled_block("건축면적", info['archArea_html'], is_area=True)
                with c3_3: render_styled_block("지상면적", info['groundArea_html'], is_area=True)
                st.write("")
                c4_1, c4_2, c4_3 = st.columns(3)
                with c4_1: render_styled_block("건물규모", f"B{info['ugrndFlrCnt']} / {info['grndFlrCnt']}F")
                with c4_2: render_styled_block("승강기/주차", f"{info.get('rideUseElvtCnt')} / {info.get('parking')}")
                with c4_3: render_styled_block("건폐/용적", f"{info.get('bcRat')}% / {info.get('vlRat')}%")
                st.write("")
                c5_1, c5_2, c5_3 = st.columns(3)
                with c5_1: render_styled_block("건물용도", info.get('mainPurpsCdNm'))
                with c5_2: render_styled_block("건물주구조", info.get('strctCdNm'))
                with c5_3: st.empty()
                st.markdown("</div>", unsafe_allow_html=True)
                st.markdown("---")

                st.subheader("💰 금액 정보")
                st.markdown("""<div style="background-color: #f8f9fa; padding: 20px; border-radius: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">""", unsafe_allow_html=True)
                st.write("") 
                row1_1, row1_2, row1_3 = st.columns(3)
                with row1_1: deposit_val = comma_input("보증금", "만원", "deposit", 0)
                with row1_2: rent_val = comma_input("월임대료", "만원", "rent", 0)
                with row1_3: maint_val = comma_input("관리비", "만원", "maint", 0)
                st.write("") 
                row2_1, row2_2, row2_3 = st.columns(3)
                with row2_1: loan_val = comma_input("융자금", "억원", "loan", 0)
                with row2_2: 
                    st.markdown(f"""<div style='font-size: 16px; font-weight: 700; color: #D32F2F; margin-bottom: 4px;'>매매금액</div>""", unsafe_allow_html=True)
                    c_in_p, c_unit_p = st.columns([3, 1]) 
                    with c_in_p:
                        if "price" not in st.session_state: st.session_state["price"] = 0
                        current_p = st.session_state["price"]; fmt_price = f"{current_p:,}" if current_p != 0 else ""
                        p_input = st.text_input("매매금액", value=fmt_price, key="price_input", label_visibility="hidden")
                        try: st.session_state["price"] = 0 if p_input.strip() == "" else int(str(p_input).replace(',', '').strip())
                        except: st.session_state["price"] = 0
                    with c_unit_p: st.markdown(f"<div style='margin-top: 15px; font-size: 18px; font-weight: 600; color: #555;'>억원</div>", unsafe_allow_html=True)
                price_val = st.session_state["price"]
                try:
                    real_invest_won = (price_val * 10000) - deposit_val
                    yield_rate = ((rent_val * 12) / real_invest_won) * 100 if real_invest_won > 0 else 0
                except: yield_rate = 0
                with row2_3:
                    st.markdown(f"""<div style='font-size: 16px; font-weight: 700; color: #1e88e5; margin-bottom: 4px;'>수익률</div><div style='background-color: #fff; border: 1px solid #ddd; border-radius: 5px; padding: 10px; text-align: center;'><span style='font-size: 28px; font-weight: 900; color: #111;'>{yield_rate:.2f}</span><span style='font-size: 18px; font-weight: 600; color: #555;'>%</span></div>""", unsafe_allow_html=True)
                st.markdown("<hr style='margin: 15px 0; border-top: 1px dashed #ddd;'>", unsafe_allow_html=True)
                
                land_py = info['platArea'] * 0.3025; tot_py = info['totArea'] * 0.3025; price_won = price_val * 100000000
                land_price_per_py = (price_won / land_py) / 10000 if land_py > 0 else 0
                tot_price_per_py = (price_won / tot_py) / 10000 if tot_py > 0 else 0
                cp1, cp2 = st.columns(2)
                with cp1: st.markdown(f"""<div class="unit-price-box"><div style="font-size:14px; color:#666;">대지 평당가</div><div class="unit-price-value">{land_price_per_py:,.0f} 만원</div></div>""", unsafe_allow_html=True)
                with cp2: st.markdown(f"""<div class="unit-price-box"><div style="font-size:14px; color:#666;">연면적 평당가</div><div class="unit-price-value">{tot_price_per_py:,.0f} 만원</div></div>""", unsafe_allow_html=True)
                st.markdown("</div>", unsafe_allow_html=True)
                st.markdown("---")

                st.subheader("🔍 AI 물건분석 (Key Insights)")
                # [요청 7] 키워드 추가 및 5열 배치 (체크박스)
                st.write("###### 👇 해당되는 키워드를 선택하세요 (다중선택)")
                env_options = [
                    "역세권", "더블역세권", "대로변", "코너입지", "이면코너", 
                    "학군지", "먹자상권", "항아리상권", "오피스상권", "신축/리모델링", 
                    "신축빌딩", "급매물", "사옥추천", "메디컬입지", "밸류업유망",
                    "주차편리", "명도협의가능", "수익형", "관리상태최상", "숲세권"
                ]
                cols_check = st.columns(5); selected_envs = []
                for i, opt in enumerate(env_options):
                    if cols_check[i % 5].checkbox(opt): selected_envs.append(opt)
                
                # [요청 10] 선택된 키워드 목록 하단 표시
                if selected_envs:
                    st.write("")
                    st.write("✅ **선택된 키워드:**")
                    tags_html = "".join([f"<span class='selected-tags'>{tag}</span>" for tag in selected_envs])
                    st.markdown(tags_html, unsafe_allow_html=True)

                st.write("")
                
                with st.expander("📂 비교 분석용 엑셀 데이터 업로드 (선택사항)", expanded=True):
                    st.info("💡 엑셀 필수 컬럼: 구분, 소재지, 대지면적, 매매금액")
                    comp_file = st.file_uploader("주변 매매사례/매물 엑셀 업로드", type=['xlsx', 'xls'], key=f"excel_{addr_input}")
                    filtered_comp_df = None; target_dong = ""
                    if comp_file:
                        try:
                            addr_parts = location['full_addr'].split(' '); 
                            for part in addr_parts: 
                                if part.endswith('동'): target_dong = part; break
                            raw_df = pd.read_excel(comp_file); raw_df.columns = [c.strip() for c in raw_df.columns]
                            required_cols = ['구분', '소재지', '대지면적', '매매금액']
                            if all(col in raw_df.columns for col in required_cols):
                                filtered_df = raw_df[raw_df['소재지'].astype(str).str.contains(target_dong, na=False)].copy() if target_dong else raw_df.copy()
                                if not filtered_df.empty:
                                    filtered_df['대지면적_숫자'] = pd.to_numeric(filtered_df['대지면적'], errors='coerce').fillna(0)
                                    filtered_df['매매금액_숫자'] = pd.to_numeric(filtered_df['매매금액'], errors='coerce').fillna(0)
                                    filtered_df['환산면적(평)'] = filtered_df['대지면적_숫자'].apply(lambda x: x * 0.3025 if x > 1000 else x)
                                    filtered_df['평당가'] = filtered_df.apply(lambda r: r['매매금액_숫자'] / r['환산면적(평)'] if r['환산면적(평)'] > 0 else 0, axis=1)
                                    filtered_comp_df = filtered_df[filtered_df['평당가'] > 0].copy()
                                    if not filtered_comp_df.empty:
                                        st.success(f"✅ '{target_dong}' 관련 데이터 {len(filtered_comp_df)}건을 찾아 분석합니다.")
                                        col_res1, col_res2 = st.columns(2)
                                        sold_cases = filtered_comp_df[filtered_comp_df['구분'].astype(str).str.contains('매각|완료|매매', na=False)]
                                        with col_res1:
                                            if not sold_cases.empty: st.markdown(f"<div style='padding:10px; background-color:#e8f5e9; border-radius:5px;'><div style='font-weight:bold; color:#2e7d32;'>📉 {target_dong} 매각 평균</div><div style='font-size:14px;'>평당 <b>{sold_cases['평당가'].mean():,.0f} 만원</b></div></div>", unsafe_allow_html=True)
                                            else: st.info(f"{target_dong} 매각 사례 없음")
                                        with col_res2:
                                            ongoing_cases = filtered_comp_df[~filtered_comp_df.index.isin(sold_cases.index)]
                                            if not ongoing_cases.empty: st.markdown(f"<div style='padding:10px; background-color:#e3f2fd; border-radius:5px;'><div style='font-weight:bold; color:#1565c0;'>📢 {target_dong} 진행 매물</div><div style='font-size:14px;'>평당 <b>{ongoing_cases['평당가'].mean():,.0f} 만원</b></div></div>", unsafe_allow_html=True)
                                            else: st.warning(f"⚠️ 엑셀 파일에 '{target_dong}' 관련 데이터가 없습니다.")
                                    else: st.warning(f"⚠️ 엑셀 파일에 '{target_dong}'이 포함된 주소가 없습니다.")
                            else: st.error(f"엑셀 컬럼 확인 필요! (필수: {required_cols})")
                        except Exception as e: st.error(f"엑셀 처리 오류: {e}")

                user_comment = st.text_area("📝 추가 특징 입력 (예: 1층 스타벅스 입점, 주인세대 명도 가능 등)", height=80)
                
                # [요청 5] 버튼 이름 변경 ("전문가" 제거 -> "인사이트요약")
                if st.button("🤖 인사이트요약 (Click)"):
                    with st.spinner("빅데이터 분석 및 리포트 생성 중..."):
                        finance_data_for_ai = {"yield": yield_rate, "price": price_val, "land_pyeong_price_val": land_price_per_py}
                        # [요청 8, 9] 후보군 생성
                        generated_candidates = generate_insight_candidates(info, finance_data_for_ai, st.session_state['zoning'], selected_envs, user_comment, filtered_comp_df, target_dong)
                        st.session_state['generated_insights'] = generated_candidates
                        st.session_state['final_selected_insights'] = [] # 초기화

                # [요청 11] 생성된 인사이트 중 사용자가 선택
                if st.session_state['generated_insights']:
                    st.write("###### 💡 생성된 투자포인트 중 사용할 내용을 선택하세요 (최대 5개)")
                    selected_indices = []
                    
                    # 체크박스로 리스트 출력
                    for idx, point in enumerate(st.session_state['generated_insights']):
                        if st.checkbox(point, key=f"insight_{idx}"):
                            selected_indices.append(point)
                    
                    st.session_state['final_selected_insights'] = selected_indices

                    if len(st.session_state['final_selected_insights']) > 5:
                        st.warning("⚠️ 최대 5개까지만 선택 가능합니다. 상위 5개만 반영됩니다.")
                        st.session_state['final_selected_insights'] = st.session_state['final_selected_insights'][:5]

                # [요청 6, 12] 선택된 내용 보여주기
                if st.session_state['final_selected_insights']:
                    st.markdown(f"""<div class="ai-summary-box"><div class="ai-title">🌟 투자포인트 내용</div>""", unsafe_allow_html=True)
                    for point in st.session_state['final_selected_insights']: 
                        st.markdown(f"<div class='insight-item'>{point}</div>", unsafe_allow_html=True)
                    st.markdown("</div>", unsafe_allow_html=True)

                st.markdown("---")
                
                finance_data = {
                    "price": price_val, "deposit": deposit_val, "rent": rent_val, 
                    "maintenance": maint_val, "loan": loan_val, "yield": yield_rate, 
                    "real_invest_eok": (price_val * 10000 - deposit_val) / 10000,
                    "land_pyeong_price_val": land_price_per_py, 
                    "tot_pyeong_price": f"{tot_price_per_py:,.0f} 만원"
                }
                z_val = st.session_state.get('zoning', '') if isinstance(st.session_state.get('zoning', ''), str) else ""
                
                # 최종 선택된 포인트만 전달
                final_summary = st.session_state.get('final_selected_insights', [])
                file_for_excel = u2 if 'u2' in locals() else None

                c_ppt, c_xls = st.columns([1, 1])
                with c_ppt:
                    st.write("##### 📥 PPT 저장")
                    ppt_template = st.file_uploader("9장짜리 샘플 PPT 템플릿 업로드 (선택)", type=['pptx'], key=f"tpl_{addr_input}")
                    if ppt_template: st.success("✅ 템플릿 적용됨")
                    pptx_file = create_pptx(info, location['full_addr'], finance_data, z_val, location['lat'], location['lng'], land_price, final_summary, images_map, template_binary=ppt_template)
                    addr_parts = location['full_addr'].split()
                    short_addr = " ".join(addr_parts[1:]) if len(addr_parts) > 1 else location['full_addr']
                    pptx_name = f"{price_val}억-{short_addr} {info.get('bldNm').replace('-','').strip()}.pptx"
                    
                    if pptx_file:
                        st.download_button(label="PPT 다운로드", data=pptx_file, file_name=pptx_name, mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", use_container_width=True)
                    else:
                        st.error("PPT 생성 중 오류가 발생했습니다. 잠시 후 다시 시도해주세요.")
                with c_xls:
                    st.write("##### 📥 엑셀 저장")
                    xlsx_file = create_excel(info, location['full_addr'], finance_data, z_val, location['lat'], location['lng'], land_price, final_summary, file_for_excel)
                    xlsx_name = f"{price_val}억-{short_addr} {info.get('bldNm').replace('-','').strip()}.xlsx"
                    st.download_button(label="엑셀 다운로드", data=xlsx_file, file_name=xlsx_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
