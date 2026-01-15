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
from pptx.enum.dml import MSO_LINE
import xlsxwriter
from urllib.parse import quote_plus
import time
import urllib3
import datetime
# [라이브러리]
import folium
from streamlit_folium import st_folium
import streamlit.components.v1 as components

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
            padding: 15px;
            border-radius: 8px;
            margin-top: 10px;
            text-align: center;
        }
        .unit-price-value {
            font-size: 28px; 
            font-weight: 900; 
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
        .insight-label {
            font-weight: 700;
            color: #1565C0;
            margin-right: 8px;
        }
    </style>
    """, unsafe_allow_html=True)

# =========================================================
# [설정] 인증키 및 전역 변수 초기화
# =========================================================
USER_KEY = "Xl5W1ALUkfEhomDR8CBUoqBMRXphLTIB7CuTto0mjsg0CQQspd7oUEmAwmw724YtkjnV05tdEx6y4yQJCe3W0g=="
VWORLD_KEY = "47B30ADD-AECB-38F3-B5B4-DD92CCA756C5"
KAKAO_API_KEY = "2a3330b822a5933035eacec86061ee41"

if 'zoning' not in st.session_state: st.session_state['zoning'] = ""
if 'selling_summary' not in st.session_state: st.session_state['selling_summary'] = []
if 'price' not in st.session_state: st.session_state['price'] = 0
if 'addr' not in st.session_state: st.session_state['addr'] = "" 
if 'last_click_lat' not in st.session_state: st.session_state['last_click_lat'] = 0.0

def reset_analysis():
    st.session_state['selling_summary'] = []

# --- [좌표 -> 주소 변환 함수] ---
def get_address_from_coords(lat, lng):
    url = "https://api.vworld.kr/req/address" 
    params = {
        "service": "address",
        "request": "getaddress",
        "version": "2.0",
        "crs": "EPSG:4326",
        "point": f"{lng},{lat}", 
        "type": "PARCEL", 
        "format": "json",
        "errorformat": "json",
        "key": VWORLD_KEY
    }
    try:
        response = requests.get(url, params=params, timeout=5, verify=False)
        data = response.json()
        if data.get('response', {}).get('status') == 'OK':
            return data['response']['result'][0]['text']
    except:
        return None
    return None

# --- [디자인 함수] ---
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
        if key not in st.session_state:
            st.session_state[key] = default_val
        current_val = st.session_state[key]
        
        formatted_val = f"{current_val:,}" if current_val != 0 else ""
        
        val_input = st.text_input(label, value=formatted_val, key=f"{key}_widget", label_visibility="hidden")
        try:
            if val_input.strip() == "":
                new_val = 0
            else:
                new_val = int(str(val_input).replace(',', '').strip())
            st.session_state[key] = new_val
        except:
            new_val = 0
            
    with c_unit:
        st.markdown(f"<div style='margin-top: 15px; font-size: 18px; font-weight: 600; color: #555;'>{unit}</div>", unsafe_allow_html=True)
    return new_val

# --- [보조 함수] ---
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

# --- [AI 인사이트 생성] ---
def generate_insight_summary(info, finance, zoning, env_features, user_comment, comp_df=None, target_dong=""):
    points = []
    
    if user_comment:
        clean_comment = user_comment.replace("\n", " ").strip()
        points.append(clean_comment)

    if comp_df is not None and not comp_df.empty:
        try:
            sold_df = comp_df[comp_df['구분'].astype(str).str.contains('매각|완료|매매', na=False)]
            if not sold_df.empty:
                avg_price = sold_df['평당가'].mean() 
                my_price = finance['land_pyeong_price_val'] 
                diff = my_price - avg_price
                diff_pct = abs(diff / avg_price) * 100
                max_price = sold_df['평당가'].max()
                loc_prefix = f"{target_dong} " if target_dong else "인근 "

                if diff < 0:
                    points.append(f"✅ {loc_prefix}매각 사례 평균(평당 {avg_price:,.0f}만) 대비 {diff_pct:.1f}% 저렴")
                    points.append(f"{loc_prefix}최고 실거래가(평당 {max_price:,.0f}만) 대비 확실한 가격 메리트")
                elif diff == 0:
                     points.append(f"{loc_prefix}실거래 시세(평당 {avg_price:,.0f}만)와 동일한 적정 시세")
                else:
                    points.append(f"{loc_prefix}평균 시세 상회하나, {zoning} 및 신축급 가치 반영 필요")
                
                points.append(f"📊 {loc_prefix}유사 입지 {len(sold_df)}건의 실거래 데이터 정밀 분석 결과")
            else:
                points.append(f"{target_dong} 인근 매각 완료 사례 없음 (진행 중 매물만 존재)")
        except Exception as e:
            pass
    elif comp_df is not None and comp_df.empty:
        points.append(f"⚠️ 업로드된 데이터에서 '{target_dong}' 관련 매매 사례를 찾을 수 없습니다.")

    if env_features:
        env_short = "/".join(env_features[:2])
        points.append(f"{env_short} 등 유동인구 풍부한 핵심 입지")
    else:
        points.append("역세권 및 대로변을 낀 탁월한 접근성과 가시성")

    yield_val = finance['yield']
    if yield_val >= 3.0:
        points.append(f"연 {yield_val:.1f}%의 안정적인 고수익과 탄탄한 임차 구성")
    else:
        points.append(f"공실 걱정 없는 안정적 임대 수익 및 높은 환금성")

    year = int(info['useAprDay'][:4]) if info.get('useAprDay') else 0
    age = datetime.datetime.now().year - year
    if age < 5:
        points.append("신축급 최상의 내외관 컨디션으로 즉시 수익 창출")
    elif age > 20:
        points.append("향후 리모델링 및 신축 개발 시 시세 차익 극대화")
    else:
        points.append("우수한 관리 상태로 추가 비용 없는 효율적 운영")
        
    return points[:6]

# --- [데이터 조회 함수] ---
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
    return "직접입력"

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
            # [추가] 값 자체를 저장
            "archArea_val": float(item.findtext('archArea', '0') or 0),
            "groundArea": float(item.findtext('vlRatEstmTotArea', '0') or 0), # 지상면적(용적률산정연면적)
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
        if res.status_code == 200 and 'image' in res.headers.get('Content-Type', ''): 
            return BytesIO(res.content)
    except: pass
    return None

# [PPT 생성 함수 - 꽉 채우기 모드 및 좌표 수정]
def create_pptx(info, full_addr, finance, zoning, lat, lng, land_price, selling_points, images_dict, template_binary=None):
    if template_binary:
        prs = Presentation(template_binary)
        
        deep_blue = RGBColor(0, 51, 153) 
        deep_red = RGBColor(204, 0, 0)   
        black = RGBColor(0, 0, 0)
        gray_border = RGBColor(128, 128, 128)
        dark_gray_border = RGBColor(80, 80, 80)

        bld_name = info.get('bldNm')
        if not bld_name or bld_name == '-':
            dong = full_addr.split(' ')[2] if len(full_addr.split(' ')) > 2 else ""
            bld_name = f"{dong} 빌딩" if dong else "사옥용 빌딩"
            
        lp_py = (land_price / 10000) / 0.3025 if land_price > 0 else 0
        total_lp_val = land_price * info['platArea'] if land_price and info['platArea'] else 0
        total_lp_str = f"{total_lp_val/100000000:,.1f}억" if total_lp_val > 0 else "-"
        ai_points_str = "\n".join(selling_points[:4]) if selling_points else "분석된 특징이 없습니다."

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

        ctx_vals = {
            'plat_m2': plat_m2, 'plat_py': plat_py,
            'tot_m2': tot_m2, 'tot_py': tot_py,
            'arch_m2': arch_m2, 'arch_py': arch_py,
            'ground_m2': ground_m2, 'ground_py': ground_py,
            'use_date': use_date
        }

        data_map = {
            "{{빌딩이름}}": bld_name,
            "{{소재지}}": full_addr,
            "{{용도지역}}": zoning,
            "{{AI물건분석내용 4가지 }}": ai_points_str,
            "{{공시지가}}": f"{land_price:,}" if land_price else "-",
            "{{공시지가 총액}}": total_lp_str,
            "{{준공년도}}": use_date,
            "{{건물규모}}": f"B{info.get('ugrndFlrCnt')} / {info.get('grndFlrCnt')}F",
            "{{건폐율}}": f"{info.get('bcRat', 0)}%",
            "{{용적률}}": f"{info.get('vlRat', 0)}%",
            "{{승강기}}": info.get('rideUseElvtCnt', '-'),
            "{{주차대수}}": info.get('parking', '-'),
            "{{건물주구조}}": info.get('strctCdNm', '-'),
            "{{건물용도}}": info.get('mainPurpsCdNm', '-'),
            "{{보증금}}": f"{finance['deposit']:,}만원" if finance['deposit'] else "-",
            "{{월임대료}}": f"{finance['rent']:,}만원" if finance['rent'] else "-",
            "{{관리비}}": f"{finance['maintenance']:,}만원" if finance['maintenance'] else "-",
            "{{수익률}}": f"{finance['yield']:.1f}%" if finance['yield'] else "-",
            "{{융자금}}": f"{finance['loan']:,}억원" if finance['loan'] else "-",
            "{{매매금액}}": f"{finance['price']:,}억원" if finance['price'] else "-",
            "{{대지평단가}}": finance.get('land_pyeong_price', '-'),
            "{{건물미래가치 활용도}}": "사옥 및 수익용 리모델링 추천",
            "{{위치도}}": "", 
            "{{지적도}}": "",
            "{{건축물대장}}": "",
            "{{건물사진}}": ""
        }

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

        def replace_text_in_frame(text_frame, mapper, ctx):
            for p in text_frame.paragraphs:
                p_text = p.text
                if "{{대지면적}}" in p_text:
                    if "평" in p_text:
                        p.text = p_text.replace("{{대지면적}}", ctx['plat_py'])
                        for r in p.runs: 
                            r.font.size = Pt(10)
                            r.font.color.rgb = deep_blue
                    else:
                        p.text = p_text.replace("{{대지면적}}", ctx['plat_m2'])
                        for r in p.runs: r.font.size = Pt(10)
                elif "{{연면적}}" in p_text:
                    if "평" in p_text:
                        p.text = p_text.replace("{{연면적}}", ctx['tot_py'])
                        for r in p.runs: 
                            r.font.size = Pt(10)
                            r.font.color.rgb = deep_blue
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
                    if ctx['use_date'] + "㎡" in new_text:
                        new_text = new_text.replace("㎡", "")
                    p.text = new_text
                    for r in p.runs: r.font.size = Pt(10)
                else:
                    found_key = None
                    for k in mapper.keys():
                        if k in p_text:
                            found_key = k
                            break
                    if found_key:
                        val = str(mapper[found_key])
                        p.text = p_text.replace(found_key, val)
                        for r in p.runs:
                            r.font.size = Pt(10)
                            if found_key == "{{빌딩이름}}":
                                r.font.size = Pt(25)
                                r.font.bold = True
                            elif found_key in ["{{보증금}}", "{{월임대료}}", "{{관리비}}", "{{융자금}}"]:
                                r.font.size = Pt(12)
                            elif found_key == "{{수익률}}":
                                r.font.size = Pt(12)
                                r.font.color.rgb = deep_red
                                r.font.bold = True
                            elif found_key == "{{매매금액}}":
                                r.font.size = Pt(16)
                                r.font.color.rgb = deep_blue
                                r.font.bold = True
                            elif found_key == "{{대지평단가}}":
                                r.font.size = Pt(10)
                                r.font.color.rgb = deep_blue
                                r.font.bold = True
        
        for slide in prs.slides:
            for shape in slide.shapes:
                replace_text_in_shape(shape, data_map, ctx_vals)

        # [이미지 삽입 - 꽉 채우기 좌표]
        img_insert_map = {
            1: ('u1', Cm(0.5), Cm(3.5), Cm(20.0), Cm(16.0)), 
            2: ('u2', Cm(0.5), Cm(3.5), Cm(10.2), Cm(14.0)), 
            4: ('u3', Cm(0.5), Cm(3.5), Cm(20.0), Cm(16.0)), 
            5: ('u4', Cm(0.5), Cm(3.5), Cm(20.0), Cm(16.0)), 
            6: ('u5', Cm(0.5), Cm(3.5), Cm(20.0), Cm(16.0))  
        }

        for s_idx, (key, l, t, w, h) in img_insert_map.items():
            if s_idx < len(prs.slides) and key in images_dict and images_dict[key] is not None:
                img_file = images_dict[key]
                img_file.seek(0)
                slide = prs.slides[s_idx]
                pic = slide.shapes.add_picture(img_file, l, t, width=w, height=h)
                
                line = pic.line
                line.visible = True
                line.width = Pt(1.5)
                if s_idx == 2:
                    line.color.rgb = dark_gray_border
                else:
                    line.color.rgb = gray_border

        output = BytesIO()
        prs.save(output)
        return output.getvalue()

    prs = Presentation()
    prs.slide_width = Cm(21.0)
    prs.slide_height = Cm(29.7)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    title_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(1.0), Cm(1.0), Cm(19.0), Cm(2.0))
    title_box.fill.background()
    title_box.line.color.rgb = RGBColor(200, 200, 200)
    title_box.line.width = Pt(1)
    
    tf = title_box.text_frame
    bld_name = info.get('bldNm')
    if not bld_name or bld_name == '-':
        dong = full_addr.split(' ')[2] if len(full_addr.split(' ')) > 2 else ""
        bld_name = f"{dong} 빌딩" if dong else "사옥용 빌딩"
        
    tf.text = bld_name
    p = tf.paragraphs[0]
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.name = "맑은 고딕"
    p.font.color.rgb = RGBColor(0, 0, 0)
    p.alignment = PP_ALIGN.CENTER

    img_y = Cm(3.5)
    img_h = Cm(11.5)
    left_x = Cm(1.0)
    col_w = Cm(9.2)
    
    lbl_img = slide.shapes.add_textbox(left_x, img_y - Cm(0.6), col_w, Cm(0.6))
    lbl_img.text_frame.text = "건물사진"
    lbl_img.text_frame.paragraphs[0].font.size = Pt(12) 
    lbl_img.text_frame.paragraphs[0].font.bold = True
    lbl_img.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

    main_img = images_dict.get('u2')
    if main_img:
        main_img.seek(0)
        slide.shapes.add_picture(main_img, left_x, img_y, width=col_w, height=img_h)
    else:
        box = slide.shapes.add_textbox(left_x, img_y, col_w, img_h)
        box.text_frame.text = "" 
    
    rect_img = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_x, img_y, col_w, img_h)
    rect_img.fill.background()
    rect_img.line.color.rgb = RGBColor(200, 200, 200)
    rect_img.line.width = Pt(1)

    map_y = Cm(15.8)
    map_h = Cm(12.0)

    lbl_map = slide.shapes.add_textbox(left_x, map_y - Cm(0.6), col_w, Cm(0.6))
    lbl_map.text_frame.text = "위치도"
    lbl_map.text_frame.paragraphs[0].font.size = Pt(12)
    lbl_map.text_frame.paragraphs[0].font.bold = True
    lbl_map.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

    map_img = get_static_map_image(lat, lng)
    if map_img: slide.shapes.add_picture(map_img, left_x, map_y, width=col_w, height=map_h)
    
    rect_map = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_x, map_y, col_w, map_h)
    rect_map.fill.background()
    rect_map.line.color.rgb = RGBColor(200, 200, 200)
    rect_map.line.width = Pt(1)

    right_x = Cm(10.8)
    
    tbl_y = Cm(3.5)
    tbl_h = Cm(11.5)
    
    lbl_tbl = slide.shapes.add_textbox(right_x, tbl_y - Cm(0.6), col_w, Cm(0.6))
    lbl_tbl.text_frame.text = "건물개요"
    lbl_tbl.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
    lbl_tbl.text_frame.paragraphs[0].font.size = Pt(12)
    lbl_tbl.text_frame.paragraphs[0].font.bold = True
    lbl_tbl.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 0, 0)

    shape = slide.shapes.add_table(11, 4, right_x, tbl_y, col_w, tbl_h)
    table = shape.table
    
    table.columns[0].width = Cm(2.3)
    table.columns[1].width = Cm(2.3)
    table.columns[2].width = Cm(2.3)
    table.columns[3].width = Cm(2.3)

    lp_py = (land_price / 10000) / 0.3025 if land_price > 0 else 0
    bcvl_text = f"{info['bcRat']:.2f}%\n{info['vlRat']:.2f}%"

    data = [
        ["소재지", full_addr, "", ""], 
        ["용도", zoning, "공시지가", f"{lp_py:,.0f}만/평"],
        ["대지", info['platArea_ppt'], "도로", "M"],
        ["연면적", info['totArea_ppt'], "준공", info['useAprDay']],
        ["지상", info['totArea_ppt'], "규모", f"B{info['ugrndFlrCnt']}/ {info['grndFlrCnt']}F"],
        ["건축", info['archArea_ppt'], "승강기", info['rideUseElvtCnt']],
        ["건/용", bcvl_text, "주차", info['parking'].split('(')[0]], 
        ["주용도", info.get('mainPurpsCdNm','-'), "주구조", info.get('strctCdNm','-')],
        ["보증금", f"{finance['deposit']:,.0f}만", "융자", f"{finance['loan']:,}억"],
        ["임대료", f"{finance['rent']:,}만", "수익률", f"{finance['yield']:.1f}%"],
        ["관리비", f"{finance['maintenance']:,}만", "매도가", f"{finance['price']:,}억"]
    ]

    for r in range(11):
        for c in range(4):
            cell = table.cell(r, c)
            cell.text = str(data[r][c])
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            
            for paragraph in cell.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.CENTER
                paragraph.font.name = "맑은 고딕"
                paragraph.font.bold = True
                paragraph.font.size = Pt(9) 
                paragraph.font.color.rgb = RGBColor(0, 0, 0) 

            if c % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(240, 248, 255)
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(255, 255, 255)

            if (r == 2 and c == 1) or (r == 3 and c == 1):
                text_val = str(data[r][c])
                if '(' in text_val:
                    area_part, pyung_part = text_val.split(' (')
                    pyung_part = '(' + pyung_part
                    cell.text_frame.clear()
                    p = cell.text_frame.paragraphs[0]
                    run1 = p.add_run(); run1.text = area_part + " "; run1.font.bold=True; run1.font.size=Pt(9); run1.font.color.rgb=RGBColor(0,0,0)
                    run2 = p.add_run(); run2.text = pyung_part; run2.font.bold=True; run2.font.size=Pt(9); run2.font.color.rgb=RGBColor(255,0,0)
                    p.alignment = PP_ALIGN.CENTER

            if r == 1: 
                for p in cell.text_frame.paragraphs: p.font.size = Pt(8)

            if r == 10 and c == 3: 
                for p in cell.text_frame.paragraphs:
                    p.font.color.rgb = RGBColor(255, 0, 0)
                    p.font.size = Pt(16)
    
    cell_addr = table.cell(0, 1)
    cell_addr.merge(table.cell(0, 3))

    cad_y = Cm(15.5) 
    cad_h = Cm(8.0) 

    lbl_cad = slide.shapes.add_textbox(right_x, cad_y - Cm(0.6), col_w, Cm(0.6))
    lbl_cad.text_frame.text = "지적도"
    lbl_cad.text_frame.paragraphs[0].font.size = Pt(12)
    lbl_cad.text_frame.paragraphs[0].font.bold = True
    lbl_cad.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

    cad_img = get_cadastral_map_image(lat, lng)
    if cad_img: slide.shapes.add_picture(cad_img, right_x, cad_y, width=col_w, height=cad_h)
    
    rect_cad = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, right_x, cad_y, col_w, cad_h)
    rect_cad.fill.background()
    rect_cad.line.color.rgb = RGBColor(200, 200, 200)
    rect_cad.line.width = Pt(1)

    ai_y = Cm(24.5) 
    ai_h = Cm(3.5)

    lbl_ai = slide.shapes.add_textbox(right_x, ai_y - Cm(0.6), col_w, Cm(0.6))
    lbl_ai.text_frame.text = "건물특징"
    lbl_ai.text_frame.paragraphs[0].font.size = Pt(12)
    lbl_ai.text_frame.paragraphs[0].font.bold = True
    lbl_ai.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

    rect_ai = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, right_x, ai_y, col_w, ai_h)
    rect_ai.fill.background()
    rect_ai.line.color.rgb = RGBColor(200, 200, 200)
    rect_ai.line.width = Pt(1)
    
    tx_ai = slide.shapes.add_textbox(right_x + Cm(0.1), ai_y + Cm(0.1), col_w - Cm(0.2), ai_h - Cm(0.2))
    tf_ai = tx_ai.text_frame
    tf_ai.word_wrap = True
    
    if selling_points:
        summary_text = ""
        for idx, pt in enumerate(selling_points[:5]):
            clean = pt.replace("</span>", "").replace("**", "").strip()
            summary_text += f"• {clean}\n"
        tf_ai.text = summary_text
    else:
        tf_ai.text = "• 역세권 입지로 투자가치 우수\n• 안정적인 임대 수익 기대"
    
    for p in tf_ai.paragraphs: 
        p.font.size = Pt(10)
        p.space_after = Pt(5)

    foot = slide.shapes.add_textbox(Cm(0), Cm(28.5), Cm(21.0), Cm(0.7))
    foot.text_frame.text = "제이에스부동산중개(주) "
    foot.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    foot.text_frame.paragraphs[0].font.bold = True
    foot.text_frame.paragraphs[0].font.size = Pt(12)

    output = BytesIO()
    prs.save(output)
    return output.getvalue()

# [엑셀 생성 - 복구됨]
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

    worksheet.set_column('A:A', 2) 
    worksheet.set_column('B:E', 12) 
    worksheet.set_column('F:F', 2) 
    worksheet.set_column('G:J', 12) 

    bld_name = info.get('bldNm')
    if not bld_name or bld_name == '-':
        dong = full_addr.split(' ')[2] if len(full_addr.split(' ')) > 2 else ""
        bld_name = f"{dong} 빌딩" if dong else "사옥용 빌딩"
    worksheet.merge_range('B2:J3', bld_name, fmt_title)

    worksheet.write('B5', '건물사진', fmt_header)
    worksheet.merge_range('B6:E20', '', fmt_box) 
    if uploaded_img:
        uploaded_img.seek(0)
        worksheet.insert_image('B6', 'building.png', {'image_data': uploaded_img, 'x_scale': 0.5, 'y_scale': 0.5, 'object_position': 2})

    worksheet.write('B22', '위치도', fmt_header)
    worksheet.merge_range('B23:E35', '', fmt_box)
    
    map_img_xls = f"http://api.vworld.kr/req/image?service=image&request=getmap&key={VWORLD_KEY}&center={lng},{lat}&crs=EPSG:4326&zoom=17&size=600,400&format=png&basemap=GRAPHIC"
    try:
        res = requests.get(map_img_xls, timeout=3)
        if res.status_code == 200:
            worksheet.insert_image('B23', 'map.png', {'image_data': BytesIO(res.content), 'x_scale': 0.7, 'y_scale': 0.7})
    except: pass

    worksheet.write('G5', '건물개요', fmt_header)
    
    lp_py = (land_price / 10000) / 0.3025 if land_price > 0 else 0
    bcvl_text = f"{info['bcRat']:.2f}%\n{info['vlRat']:.2f}%"
    
    table_data_xls = [
        ["소재지", full_addr, "용도", zoning],
        ["공시지가", f"{lp_py:,.0f}만/평", "대지", info['platArea_ppt']], 
        ["도로", "6M", "연면적", info['totArea_ppt']],
        ["준공", info['useAprDay'], "지상", info['totArea_ppt']],
        ["규모", f"B{info['ugrndFlrCnt']}/ {info['grndFlrCnt']}F", "건축", info['archArea_ppt']],
        ["승강기", info['rideUseElvtCnt'], "건/용", bcvl_text],
        ["주차", info['parking'].split('(')[0], "주용도", info.get('mainPurpsCdNm','-')],
        ["주구조", info.get('strctCdNm','-'), "보증금", f"{finance['deposit']:,.0f}만"],
        ["융자", f"{finance['loan']:,}억", "임대료", f"{finance['rent']:,}만"],
        ["수익률", f"{finance['yield']:.1f}%", "관리비", f"{finance['maintenance']:,}만"],
        ["매도가", f"{finance['price']:,}억", "", ""] 
    ]

    start_row = 5
    for i, row in enumerate(table_data_xls):
        worksheet.write(start_row + i, 6, row[0], fmt_label) 
        if row[0] == "매도가":
             worksheet.merge_range(start_row + i, 7, start_row + i, 9, row[1], fmt_val_red)
        else:
             worksheet.write(start_row + i, 7, row[1], fmt_val) 
        
        if row[0] != "매도가":
            worksheet.write(start_row + i, 8, row[2], fmt_label) 
            worksheet.write(start_row + i, 9, row[3], fmt_val) 

    worksheet.write('G17', '지적도', fmt_header) 
    worksheet.merge_range('G18:J26', '', fmt_box)
    cad_img = get_cadastral_map_image(lat, lng)
    if cad_img:
        worksheet.insert_image('G18', 'cad.png', {'image_data': cad_img, 'x_scale': 0.6, 'y_scale': 0.6})

    worksheet.write('G28', '건물특징', fmt_header)
    worksheet.merge_range('G29:J35', '', fmt_box)
    
    summary_text = ""
    if selling_points:
        for idx, pt in enumerate(selling_points[:5]):
            clean = pt.replace("</span>", "").replace("**", "").strip()
            summary_text += f"• {clean}\n"
    else:
        summary_text = "• 역세권 입지로 투자가치 우수\n• 안정적인 임대 수익 기대"
        
    worksheet.write('G29', summary_text, fmt_box)

    worksheet.merge_range('B37:J37', "JS 제이에스부동산(주) 김창익 이사 010-6595-5700", fmt_title)

    workbook.close()
    return output.getvalue()

# [메인 실행]
st.title("🏢 부동산 매입 분석기 Pro")
st.markdown("---")

# --- [통합된 부분] 지도에서 클릭하여 찾기 ---
with st.expander("🗺 지도에서 직접 클릭하여 찾기 (Click)", expanded=False):
    m = folium.Map(location=[37.5172, 127.0473], zoom_start=14)
    output = st_folium(m, width=700, height=400)

    if output and output.get("last_clicked"):
        lat = output["last_clicked"]["lat"]
        lng = output["last_clicked"]["lng"]
        
        if "last_click_lat" not in st.session_state or st.session_state["last_click_lat"] != lat:
            st.session_state["last_click_lat"] = lat
            
            found_addr = get_address_from_coords(lat, lng)
            if found_addr:
                st.success(f"📍 지도 클릭 확인! 변환된 주소: {found_addr}")
                st.session_state['addr'] = found_addr
                reset_analysis()
                st.rerun()
            else:
                st.warning("⚠️ 주소를 찾을 수 없는 위치입니다.")

# --- [주소 입력창] ---
addr_input = st.text_input("주소 입력", placeholder="예: 강남구 논현동 254-4", key="addr", on_change=reset_analysis)

if addr_input:
    with st.spinner("데이터 분석 중..."):
        location = get_pnu_and_coords(addr_input)
        
        if not location:
            st.error("❌ 주소를 찾을 수 없습니다.")
        else:
            if not st.session_state['zoning']:
                fetched_zoning = get_zoning_smart(location['lat'], location['lng'])
                st.session_state['zoning'] = fetched_zoning

            info = get_building_info_smart(location['pnu'])
            land_price = get_land_price(location['pnu'])
            
            if not info or "error" in info:
                st.error(f"조회 실패: {info.get('error')}")
            else:
                st.success("✅ 분석 완료!")
                
                # [위치 이동 및 UI 변경] 5개의 파일 업로더 생성 - 밖으로 꺼냄
                st.write("##### 📸 PPT 삽입용 사진 업로드 (박스 안으로 드래그 하세요)")
                
                col_u1, col_u2, col_u3 = st.columns(3)
                with col_u1: u1 = st.file_uploader("Slide 2: 위치도", type=['png', 'jpg', 'jpeg'], key="u1")
                with col_u2: u2 = st.file_uploader("Slide 3: 건물메인", type=['png', 'jpg', 'jpeg'], key="u2")
                with col_u3: u3 = st.file_uploader("Slide 5: 지적도", type=['png', 'jpg', 'jpeg'], key="u3")
                
                col_u4, col_u5 = st.columns(2)
                with col_u4: u4 = st.file_uploader("Slide 6: 건축물대장", type=['png', 'jpg', 'jpeg'], key="u4")
                with col_u5: u5 = st.file_uploader("Slide 7: 추가사진", type=['png', 'jpg', 'jpeg'], key="u5")
                
                images_map = {'u1': u1, 'u2': u2, 'u3': u3, 'u4': u4, 'u5': u5}

                st.markdown("---")

                # [건물 및 토지 정보]
                st.markdown("""<div style="background-color: #f8f9fa; padding: 50px; border-radius: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">""", unsafe_allow_html=True)
                
                c1, c2 = st.columns([2, 1])
                with c1: render_styled_block("소재지", addr_input)
                with c2: render_styled_block("건물명", info.get('bldNm'))
                st.write("") 

                # 공시지가
                c_lp1, c_lp2, c_lp3 = st.columns(3)
                with c_lp1:
                    if land_price > 0:
                        render_styled_block("개별공시지가(㎡)", f"{land_price:,} 원")
                    else:
                        st.warning("⚠️ 공시지가 조회 불가")
                        manual_lp = st.number_input("공시지가 직접입력(원)", value=0, step=1000)
                        land_price = manual_lp 

                with c_lp2:
                    if land_price > 0 and info['platArea'] > 0:
                        total_lp = land_price * info['platArea']
                        total_lp_eok = total_lp / 100000000
                        render_styled_block("공시지가 총액(추정)", f"{total_lp_eok:,.2f}억")
                    else:
                         render_styled_block("공시지가 총액", "-")
                with c_lp3: st.empty()
                st.write("")
                st.markdown("<hr style='margin: 10px 0; border-top: 1px dashed #ddd;'>", unsafe_allow_html=True)

                c2_1, c2_2, c2_3 = st.columns(3)
                with c2_1:
                    if st.session_state['zoning'] == "직접입력":
                        st.write("용도지역")
                        zoning_manual = st.text_input("zoning", value="", label_visibility="collapsed")
                        st.session_state['zoning'] = zoning_manual 
                    else:
                        render_styled_block("용도지역", st.session_state['zoning'])
                        
                with c2_2: 
                    render_styled_block("대지면적", info['platArea_html'], is_area=True)
                        
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

                # [금액 정보]
                st.subheader("💰 금액 정보")
                st.markdown("""<div style="background-color: #f8f9fa; padding: 20px; border-radius: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">""", unsafe_allow_html=True)
                st.write("") 

                row1_1, row1_2, row1_3 = st.columns(3)
                with row1_1: deposit_val = comma_input("보증금", "만원", "deposit", 0, help_text="")
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
                        current_p = st.session_state["price"]
                        fmt_price = f"{current_p:,}" if current_p != 0 else ""
                        p_input = st.text_input("매매금액", value=fmt_price, key="price_input", label_visibility="hidden")
                        try:
                            if p_input.strip() == "": price_val = 0
                            else: price_val = int(p_input.replace(',', '').strip())
                            st.session_state["price"] = price_val
                        except: price_val = 0
                    with c_unit_p:
                        st.markdown(f"<div style='margin-top: 15px; font-size: 18px; font-weight: 600; color: #555;'>억원</div>", unsafe_allow_html=True)

                try:
                    real_invest_won = (price_val * 10000) - deposit_val
                    real_invest_eok = real_invest_won / 10000
                    if real_invest_won > 0: yield_rate = ((rent_val * 12) / real_invest_won) * 100
                    else: yield_rate = 0
                except: 
                    yield_rate = 0
                    real_invest_eok = 0

                with row2_3:
                    st.markdown(f"""
                        <div style='font-size: 16px; font-weight: 700; color: #1e88e5; margin-bottom: 4px;'>수익률</div>
                        <div style='background-color: #fff; border: 1px solid #ddd; border-radius: 5px; padding: 10px; text-align: center;'>
                            <span style='font-size: 28px; font-weight: 900; color: #111;'>{yield_rate:.2f}</span>
                            <span style='font-size: 18px; font-weight: 600; color: #555;'>%</span>
                        </div>
                    """, unsafe_allow_html=True)

                st.markdown("<hr style='margin: 15px 0; border-top: 1px dashed #ddd;'>", unsafe_allow_html=True)
                
                land_py = info['platArea'] * 0.3025
                tot_py = info['totArea'] * 0.3025
                price_won = price_val * 100000000

                land_price_per_py = 0
                tot_price_per_py = 0
                
                if land_py > 0: land_price_per_py = (price_won / land_py) / 10000 
                if tot_py > 0: tot_price_per_py = (price_won / tot_py) / 10000        

                cp1, cp2 = st.columns(2)
                with cp1:
                    st.markdown(f"""<div class="unit-price-box"><div style="font-size:14px; color:#666;">대지 평당가</div><div class="unit-price-value">{land_price_per_py:,.0f} 만원</div></div>""", unsafe_allow_html=True)
                with cp2:
                    st.markdown(f"""<div class="unit-price-box"><div style="font-size:14px; color:#666;">연면적 평당가</div><div class="unit-price-value">{tot_price_per_py:,.0f} 만원</div></div>""", unsafe_allow_html=True)

                st.markdown("</div>", unsafe_allow_html=True)
                st.markdown("---")

                # [AI 인사이트 요약]
                st.subheader("🔍 AI 물건분석 (Key Insights)")
                
                st.write("###### 👇 해당되는 키워드를 선택하세요 (다중선택)")
                env_options = [
                    "역세권", "대로변", "코너입지", "학군지", 
                    "먹자상권", "오피스상권", "숲세권", "신축/리모델링",
                    "급매물", "사옥추천", "메디컬입지", "주차편리", 
                    "명도협의가능", "수익형", "밸류업유망", "관리상태최상"
                ]
                
                cols_check = st.columns(4)
                selected_envs = []
                for i, opt in enumerate(env_options):
                    if cols_check[i % 4].checkbox(opt):
                        selected_envs.append(opt)

                st.write("")
                
                with st.expander("📂 비교 분석용 엑셀 데이터 업로드 (선택사항)", expanded=True):
                    st.info("💡 엑셀 필수 컬럼: 구분, 소재지, 대지면적, 매매금액")
                    comp_file = st.file_uploader("주변 매매사례/매물 엑셀 업로드", type=['xlsx', 'xls'], key=f"excel_{addr_input}")
                    filtered_comp_df = None
                    target_dong = ""
                    
                    if comp_file:
                        try:
                            addr_parts = location['full_addr'].split(' ')
                            for part in addr_parts:
                                if part.endswith('동'):
                                    target_dong = part
                                    break
                            
                            raw_df = pd.read_excel(comp_file)
                            raw_df.columns = [c.strip() for c in raw_df.columns]
                            
                            required_cols = ['구분', '소재지', '대지면적', '매매금액']
                            if all(col in raw_df.columns for col in required_cols):
                                if target_dong:
                                    filtered_df = raw_df[raw_df['소재지'].astype(str).str.contains(target_dong, na=False)].copy()
                                else:
                                    filtered_df = raw_df.copy()

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
                                            if not sold_cases.empty:
                                                avg_sold = sold_cases['평당가'].mean()
                                                st.markdown(f"""
                                                <div style="padding:10px; background-color:#e8f5e9; border-radius:5px;">
                                                    <div style="font-weight:bold; color:#2e7d32;">📉 {target_dong} 매각 평균</div>
                                                    <div style="font-size:14px;">평당 <b>{avg_sold:,.0f} 만원</b></div>
                                                </div>
                                                """, unsafe_allow_html=True)
                                            else:
                                                st.info(f"{target_dong} 매각 사례 없음")

                                        with col_res2:
                                            ongoing_cases = filtered_comp_df[~filtered_comp_df.index.isin(sold_cases.index)]
                                            if not ongoing_cases.empty:
                                                avg_ongoing = ongoing_cases['평당가'].mean()
                                                st.markdown(f"""
                                                <div style="padding:10px; background-color:#e3f2fd; border-radius:5px;">
                                                    <div style="font-weight:bold; color:#1565c0;">📢 {target_dong} 진행 매물</div>
                                                    <div style="font-size:14px;">평당 <b>{avg_ongoing:,.0f} 만원</b></div>
                                                </div>
                                                """, unsafe_allow_html=True)
                                            else:
                                                st.warning(f"⚠️ 엑셀 파일에 '{target_dong}' 관련 데이터가 없습니다.")
                                    else:
                                        st.warning(f"⚠️ 엑셀 파일에 '{target_dong}'이 포함된 주소가 없습니다.")
                            else:
                                st.error(f"엑셀 컬럼 확인 필요! (필수: {required_cols})")
                        except Exception as e:
                            st.error(f"엑셀 처리 오류: {e}")

                user_comment = st.text_area("📝 추가 특징 입력 (예: 1층 스타벅스 입점, 주인세대 명도 가능 등)", height=80)

                if st.button("🤖 전문가 인사이트 요약 생성 (Click)"):
                    with st.spinner("빅데이터 분석 및 리포트 작성 중..."):
                        finance_data_for_ai = {
                            "yield": yield_rate, 
                            "price": price_val,
                            "land_pyeong_price_val": land_price_per_py
                        }
                        
                        summary_points = generate_insight_summary(
                            info, finance_data_for_ai, st.session_state['zoning'], 
                            selected_envs, user_comment, filtered_comp_df, target_dong
                        )
                        st.session_state['selling_summary'] = summary_points
                
                if st.session_state['selling_summary']:
                    st.markdown(f"""<div class="ai-summary-box"><div class="ai-title">🌟 전문가 투자 포인트 (Key Insights)</div>""", unsafe_allow_html=True)
                    for point in st.session_state['selling_summary']:
                        st.markdown(f"<div class='insight-item'>{point}</div>", unsafe_allow_html=True)
                    st.markdown("</div>", unsafe_allow_html=True)

                st.markdown("---")

                st.subheader("🗺 지도 및 다운로드")
                
                naver_map_url = f"https://map.naver.com/v5/search/{quote_plus(location['full_addr'])}"
                st.markdown(f"**[📍 네이버 지도에서 위치 확인하기 (Click)]({naver_map_url})**")
                
                finance_data = {
                    "price": price_val, "deposit": deposit_val, "rent": rent_val, 
                    "maintenance": maint_val, "loan": loan_val, "yield": yield_rate, 
                    "real_invest_eok": real_invest_eok,
                    "land_pyeong_price": f"{land_price_per_py:,.0f} 만원",
                    "tot_pyeong_price": f"{tot_price_per_py:,.0f} 만원"
                }
                z_val = st.session_state.get('zoning', '') if isinstance(st.session_state.get('zoning', ''), str) else ""
                current_summary = st.session_state.get('selling_summary', [])

                file_for_excel = u2 if 'u2' in locals() else None

                st.markdown("---")
                
                c_ppt, c_xls = st.columns([1, 1])
                
                with c_ppt:
                    st.write("##### 📥 PPT 저장")
                    
                    ppt_template = st.file_uploader("9장짜리 샘플 PPT 템플릿 업로드 (선택)", type=['pptx'], key=f"tpl_{addr_input}")
                    
                    if ppt_template:
                        st.success("✅ 템플릿 적용됨 (9장 생성 모드 + 자동 사진 삽입)")
                        
                    pptx_file = create_pptx(info, location['full_addr'], finance_data, z_val, location['lat'], location['lng'], land_price, current_summary, images_map, template_binary=ppt_template)
                    
                    st.download_button(label="PPT 다운로드", data=pptx_file, file_name=f"부동산분석_{addr_input}.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", use_container_width=True)
                
                with c_xls:
                    st.write("##### 📥 엑셀 저장")
                    # create_excel 함수가 여기서 정상적으로 호출됩니다.
                    xlsx_file = create_excel(info, location['full_addr'], finance_data, z_val, location['lat'], location['lng'], land_price, current_summary, file_for_excel)

                    st.download_button(label="엑셀 다운로드", data=xlsx_file, file_name=f"부동산분석_{addr_input}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
