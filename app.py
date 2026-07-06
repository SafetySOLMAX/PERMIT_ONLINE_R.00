import streamlit as st
import pandas as pd
import datetime
import random
import os
import time
from PIL import Image

# ==============================================================================
# [ส่วนเพิ่มชุดที่ 1 - หัว]: ระบบเชื่อมต่อฐานข้อมูล SQLite กับ SharePoint องค์กร
# ==============================================================================
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

SHAREPOINT_URL = "https://solmax.sharepoint.com/sites/SafetyTHRayong"
SHAREPOINT_FOLDER_URL = "/sites/SafetyTHRayong/Shared Documents" 
DB_FILE_NAME = "hse_permits.db"

# 🛑 ข้อแนะนำ: ตรงนี้ให้ใส่อีเมลและรหัสผ่านองค์กรของคุณ 
USERNAME = "bulans@solmax.com" 
PASSWORD = "ใส่รหัสผ่านคลาวด์องค์กรของคุณตรงนี้"

def get_sharepoint_context():
    return ClientContext(SHAREPOINT_URL).with_credentials(UserCredential(USERNAME, PASSWORD))

# ฟังก์ชันดึงไฟล์ฐานข้อมูลล่าสุดจากคลาวด์บริษัทลงมาทำงาน
def download_db_from_sharepoint():
    try:
        ctx = get_sharepoint_context()
        download_path = os.path.join(os.getcwd(), DB_FILE_NAME)
        remotefile_url = f"{SHAREPOINT_FOLDER_URL}/{DB_FILE_NAME}"
        
        with open(download_path, "wb") as local_file:
            ctx.web.get_file_by_server_relative_url(remotefile_url).download(local_file).execute_query()
        print("ดาวน์โหลดฐานข้อมูลสำเร็จ")
    except Exception as e:
        st.warning(f"⚠️ ยังไม่มีไฟล์ฐานข้อมูลบนคลาวด์ หรือสิทธิ์การเข้าถึงมีปัญหา ระบบกำลังเริ่มใช้ DB ชั่วคราว: {e}")

# ฟังก์ชันส่งไฟล์ฐานข้อมูลที่อัปเดตแล้วกลับขึ้นไปจัดเก็บแบบถาวร
def upload_db_to_sharepoint():
    try:
        ctx = get_sharepoint_context()
        target_folder = ctx.web.get_folder_by_server_relative_url(SHAREPOINT_FOLDER_URL)
        local_path = os.path.join(os.getcwd(), DB_FILE_NAME)
        
        with open(local_path, 'rb') as content_file:
            file_content = content_file.read()
            target_folder.upload_file(DB_FILE_NAME, file_content).execute_query()
        print("อัปโหลดฐานข้อมูลขึ้น SharePoint สำเร็จ")
    except Exception as e:
        st.error(f"❌ ไม่สามารถส่งไฟล์กลับขึ้น SharePoint ได้: {e}")

# สั่งให้ระบบวิ่งไปเอาไฟล์ล่าสุดมาจากคลาวด์ก่อนเปิดหน้าเว็บขึ้นมา (ถ้ายังไม่มีไฟล์อยู่ในแอป)
if not os.path.exists(DB_FILE_NAME):
    download_db_from_sharepoint()
# ==============================================================================

# 1. ตั้งค่าหน้าจอระบบ
st.set_page_config(
    page_title="Solmax Work Permit Online System",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. โหลดโลโก้บริษัท
def load_solmax_logo():
    logo_path = "Solmax_Digital_Logo_green_black (1).jpg"
    if os.path.exists(logo_path):
        return Image.open(logo_path)
    return None

logo_img = load_solmax_logo()

# 3. ปรับปรุง CSS สไตล์ Solmax Premium ให้สแกนง่าย สวยงาม
st.markdown(
    """
    <style>
    .stApp { background-color: #F8FAFC; }
    [data-testid="stSidebar"] { background-color: #0B1B3D !important; }
    [data-testid="stSidebar"] label, [data-testid="stSidebar"] label p, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span, [data-testid="stSidebar"] .stRadio div label { 
        color: #FFFFFF !important; font-size: 18px !important; font-weight: 600 !important; 
    }
    .sidebar-sub { text-align: center; color: #CBD5E1 !important; font-size: 14px; margin-bottom: 20px; }
    .sidebar-title { text-align: center; font-weight: 700; color: #00D084 !important; font-size: 22px; margin-top: 15px; }
    h1 { color: #0B1B3D !important; font-size: 36px !important; font-weight: 700 !important; margin-bottom: 10px !important; }
    .stWidget label p, label, .stMarkdown p, .stCheckbox p, .stRadio p { font-size: 18px !important; font-weight: 700 !important; color: #0F172A !important; }
    input, select, textarea, div[data-baseweb="select"] span, .stFileUploader p { font-size: 17px !important; color: #0F172A !important; font-weight: 500 !important; }
    div.stAlert, div[data-testid="stForm"], .custom-card { background-color: #FFFFFF !important; border-radius: 14px !important; box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05) !important; border: 1px solid #E2E8F0 !important; padding: 32px !important; margin-bottom: 30px !important; }
    div.stButton > button { background-color: #00D084 !important; color: #0B1B3D !important; font-weight: 800 !important; font-size: 18px !important; border-radius: 8px !important; border: none !important; padding: 12px 30px !important; transition: all 0.3s ease; }
    div.stButton > button:hover { background-color: #05B574 !important; color: #FFFFFF !important; box-shadow: 0 4px 12px rgba(5, 181, 116, 0.3); }
    </style>
    """, unsafe_allow_html=True
)

# 4. โครงสร้างข้อมูลมาตรการตรวจสอบความปลอดภัย
permit_master_data = {
    "FM-HSE-001: ใบอนุญาตทำงานที่ก่อให้เกิดความร้อน (Hot Work Permit)": {
        "checklist": [
            "ปั๊มน้ำระบบดับเพลิงพร้อมใช้งานและอยู่ในโหมดอัตโนมัติ / Fire pump ready & Auto mode",
            "วาล์วควบคุมการจ่ายน้ำสำหรับระบบสปริงเกอร์เปิดอยู่ / Sprinkler valves open",
            "มีเครื่องดับเพลิงเคมีประจำจุดปฏิบัติงานและสามารถใช้งานได้ / Fire extinguishers available and ready",
            "เคลื่อนย้ายสิ่งที่ติดไฟออกจากระยะ 35 ฟุต (10 เมตร) หรือคลุมด้วยผ้ากันไฟที่ได้รับอนุมัติมาตรฐาน",
            "แยกแหล่งก๊าซไวไฟ ของเหลวที่ติดไฟได้ หรือฝุ่นผงออกจากพื้นที่ปฏิบัติงาน",
            "อุปกรณ์สำหรับการทำงานที่ก่อให้เกิดความร้อน (สายเชื่อม, หัวตัด, เครื่องขัด) อยู่ในสภาพสมบูรณ์เรียบร้อย",
            "มีผู้เฝ้าระวังไฟ (Fire Watcher) ประจำจุดปฏิบัติงานตลอดเวลาและหลังเลิกงานอย่างน้อย 1 ชั่วโมง"
        ],
        "ppe": ["หน้ากากเชื่อมไฟฟ้า / แว่นตานิรภัยเซฟตี้", "ถุงมือหนังสำหรับงานเชื่อม", "ชุดผ้ากันไฟ / ปลอกแขนกันความร้อน", "รองเท้านิรภัยหัวเหล็ก"]
    },
    "FM-HSE-002: ใบอนุญาตทำงานขุดเจาะพื้นดิน (Excavation Work Permit)": {
        "checklist": [
            "ตรวจสอบและยืนยันว่าบริเวณใต้พื้นดินไม่มีสายไฟ ท่อประปา หรือท่อส่งลำเลียงสารเคมีใต้ดิน",
            "มีการปิดกั้นพื้นที่ขุดเจาะด้วยเชือกหรือรั้วกั้นมาตรฐานรอบแนวเขตขุดตลอดเวลา",
            "ตัด / ล็อก / แขวนป้ายเตือนอันตราย (LOTO) ทางเครื่องจักรกลและระบบไฟฟ้าเรียบร้อยแล้ว",
            "ติดตั้งป้ายเตือนอันตรายเขตก่อสร้าง/งานขุดเจาะอย่างชัดเจน",
            "ตรวจสอบแล้วว่าบริเวณที่ขุดเจาะไม่มีการรั่วซึมของสารเคมีหรือสารกัดกร่อน",
            "กรณีหลุมขุดลึกเกิน 1.2 เมตร ต้องติดตั้งบันไดหนีภัยและทำผนังโครงสร้างกันดินพังทลายอย่างมั่นคง",
            "ทำเครื่องหมายชี้บ่งตำแหน่งที่มีระบบสาธารณูปโภคหรืออุปกรณ์ใต้ดินให้เห็นเด่นชัด"
        ],
        "ppe": ["หมวกนิรภัย (Safety Helmet)", "เสื้อสะท้อนแสงเขตก่อสร้าง", "ถุงมือผ้า/ถุงมือยางหนา", "รองเท้านิรภัยหัวเหล็กหรือรองเท้าบู๊ทยางกันสิ่งมีคม"]
    },
    "FM-HSE-003: ใบอนุญาตทำงานบนที่สูง (High Area Work Permit)": {
        "checklist": [
            "จัดทำ Barricade แผงกั้นและป้ายเตือนเครื่องหมายความปลอดภัยบริเวณพื้นที่ด้านล่างที่ปฏิบัติงาน",
            "งานที่สูงเกิน 2 เมตร ติดตั้งนั่งร้านที่มีราวกันตก (Hand Rail) โครงค้ำยันแน่นหนา และแผ่นทางเดินแข็งแรง",
            "นั่งร้านได้รับการตรวจสอบความปลอดภัยและติดป้ายสัญลักษณ์ Tag สีเขียวพร้อมใช้งานแล้ว",
            "ห้ามโยนสิ่งของหรือเครื่องมือขึ้น-ลงโดยเด็ดขาด ให้ใช้เชือกผูกดึงขึ้น-ลงเท่านั้น",
            "กรณีใช้บันไดสำเร็จรูป ต้องมีสภาพสมบูรณ์ แข็งแรง และมีการผูกยึดล็อกฐานป้องกันการล้มเอียงขณะปีน",
            "เครื่องมือชิ้นเล็กสำหรับการพกพา ต้องมีการผูกเชือกคล้องไว้กับตัวผู้ปฏิบัติงานขณะใช้งานบนที่สูง",
            "ผู้ปฏิบัติงานผ่านการตรวจวัดปริมาณแอลกอฮอล์เป็น 0% และไม่มีโรคกลัวความสูงหรือโรคประจำตัวอันตราย"
        ],
        "ppe": ["เข็มขัดนิรภัยเต็มตัวพร้อมสายช่วยชีวิต (Full Body Safety Harness with Lanyard)", "หมวกนิรภัยพร้อมสายรัดคอ", "ถุงมือกระชับมือ", "รองเท้านิรภัยชนิดกันลื่น"]
    },
    "FM-HSE-004: ใบอนุญาตทำงานเกี่ยวกับไฟฟ้า (Electrical Work Permit)": {
        "checklist": [
            "ทำการตัดแยกกระแสไฟฟ้า ล็อกเบรกเกอร์ และแขวนป้ายเตือนอันตราย (LOTO) ในจุดที่เกี่ยวข้องเสร็จสิ้น",
            "ตรวจสอบสภาพสายไฟฟ้า โครงสร้างตู้ไฟ และตัวนำไฟฟ้าให้อยู่ในสภาพสมบูรณ์ ไม่มีรอยฉีกขาดหรือชำรุด",
            "จัดวางสายเคเบิ้ลและสายไฟชั่วคราวให้อยู่ในจุดปลอดภัย ไม่กีดขวางหรือพาดผ่านทางสัญจร",
            "ตรวจเช็กระบบโครงสร้างแล้ว มั่นใจว่าไม่มีผลกระทบด้านพลังงานย้อนกลับไปยังอุปกรณ์ตัวอื่น",
            "ตรวจสอบจุดยึดต่อสายไฟทั้งหมดอย่างละเอียด ต้องไม่มีสภาพเปลือยหรือไม่มีฉนวนหุ้ม",
            "ตรวจสอบพื้นที่ทำงานโดยรอบ ต้องแห้งสนิท ไม่เปียกชื้น หรือไม่มีสารเคมีไวไฟรั่วไหล",
            "เครื่องมือช่างไฟฟ้า (คีม, ไขควง, มิเตอร์) มีฉนวนกันไฟฟ้าหุ้มที่ด้ามจับอย่างแน่นหนา ไม่ชำรุด"
        ],
        "ppe": ["หมวกนิรภัยสำหรับงานไฟฟ้า", "ถุงมือยางป้องกันกระแสไฟฟ้า (Insulated Rubber Gloves)", "แว่นตานิรภัยกันประกายไฟความร้อน", "รองเท้าหนังพื้นยางหุ้มส้นชนิดกันไฟดูด"]
    },
    "FM-HSE-021: ใบอนุญาตเข้ามาปฏิบัติงานทั่วไป (General Work Permit)": {
        "checklist": [
            "ทำการปิดกั้นและตัดแยกพลังงานที่เป็นอันตราย (LOTO) เรียบร้อยแล้ว (กรณีเครื่องจักร)",
            "ปิดกั้นพื้นที่ปฏิบัติงานด้วยเทปขาว-แดงชั่วคราวและติดตั้งป้ายชี้บ่งระวังอันตรายอย่างชัดเจน",
            "ผู้ปฏิบัติงานทุกคนผ่านการอบรมกฎความปลอดภัย มีบัตรประจำตัวผู้รับเหมาและติดสแดงให้เห็นเด่นชัด",
            "เครื่องมือกล หรือเครื่องมือไฟฟ้าผ่านการตรวจสภาพจากฝ่ายซ่อมบำรุงและติดสติกเกอร์สีเขียวผ่านการตรวจ",
            "จัดทำกิจกรรมคุยความปลอดภัยหน้างาน (Safety Toolbox Talk) ชี้แจงขั้นตอนและความเสี่ยงก่อนเริ่มงาน"
        ],
        "ppe": ["หมวกนิรภัย (Safety Helmet)", "แว่นตานิรภัยป้องกันเศษวัสดุ", "ถุงมือผ้าป้องกันการบาดเจ็บ", "รองเท้านิรภัยมาตรฐาน (Safety Shoes)"]
    },
    "FM-HSE-037: ใบอนุญาตปฏิบัติงานในสถานที่อับอากาศ (Confined Space Permit)": {
        "checklist": [
            "ทำความสะอาดบริเวณพื้นที่อับอากาศจนปราศจากคราบน้ำมัน สารเคมีตกค้าง หรือวัสดุที่ติดไฟได้",
            "ทำการตัดแยกพลังงาน ล็อกท่อส่ง ปิดวาล์ว และติดป้ายเตือนระวังอันตรายอย่างสมบูรณ์",
            "ติดตั้งพัดลม/เครื่องดูดระบายอากาศเปิดใช้งานต่อเนื่องเพื่อหมุนเวียนอากาศภายใน",
            "จัดตั้งผู้เฝ้าระวังภัยอับอากาศ (Stand-by Man) อยู่ประจำตำแหน่งทางเข้า-ออกตลอดเวลาปฏิบัติงาน",
            "ติดตั้งป้ายเตือนพื้นที่อับอากาศอันตราย และบอร์ดกระดานลงชื่อตรวจสอบคนเข้า-ออกอย่างเข้มงวด",
            "จัดเตรียมอุปกรณ์ช่วยชีวิต ฉุกเฉิน (Life Line, รอกกู้ภัย หรือชุดปฐมพยาบาล) พร้อมใช้งานในพื้นที่",
            "ผ่านการตรวจวัดระดับแก๊ส (แก๊สไวไฟ %LEL ต้องเป็น 0, ออกซิเจนอยู่ระหว่าง 19.5% - 23.5%, แก๊สพิษ %CO ปลอดภัย)"
        ],
        "ppe": ["ชุดเชือกช่วยชีวิตและสายรัดลำตัวนิรภัย (Harness & Life Line)", "เครื่องตรวจวัดก๊าซพกพา (Portable Gas Detector)", "หน้ากากกรองฝุ่นก๊าซ หรือชุดถังอัดอากาศ SCBA (กรณีอากาศไม่พอ)", "อุปกรณ์ลดเสียงดัง (Ear Plug/Muffs)"]
    }
}

# 5. รายชื่อผู้ควบคุมงาน (ครบทั้ง 16 รายชื่ออ้างอิงตามไฟล์แนบ Sheet ผู้ควบคุมงาน)
list_supervisors = [
    {"name": "นายธนพล เดชนนท์ธนวัฒน์", "email": "dthanaphon@solmax.com"},
    {"name": "นายชัยกร ชารีแก้ว", "email": "ChaiyakornC@solmax.com"},
    {"name": "นายสัมฤทธิ์ ก้องโสตร", "email": "ksamrit@solmax.com"},
    {"name": "นายประภาส แก้วอรสาร", "email": "praphask@solmax.com"},
    {"name": "นายไพฑูรย์ บางกุ้ง", "email": "paitoonb@solmax.com"},
    {"name": "นายสุรรัช เสถียรุจิกานนท์", "email": "ssurarat@solmax.com"},
    {"name": "นายมนพัทธ์ ผลกานต์ดี", "email": "Monphatp@solmax.com"},
    {"name": "นายชาญชัย กะทิศาสตร์", "email": "Chanchaik@solmax.com"},
    {"name": "นายพิชิตชัย นันท์ขุนทด", "email": "pichitchain@solmax.com"},
    {"name": "นายกิตติ เรียงไข", "email": "Kittir@solmax.com"},
    {"name": "น.ส.มนิสา กาญจณา", "email": "manisak@solmax.com"},
    {"name": "นายนริน บุญช่วยเหลือ", "email": "narinb@solmax.com"},
    {"name": "นายจักรพงษ์ สว่างวงศ์", "email": "jakkapongs@solmax.com"},
    {"name": "น.ส.ณฐมน บุญมีรอด", "email": "nathamonb@solmax.com"},
    {"name": "น.ส.ชุตินันท์ แซ่แต้", "email": "chutinuns@solmax.com"},
    {"name": "น.ส.เนตรนภา ดำบรรพ์", "email": "netnapad@solmax.com"}
]

# รายชื่อเจ้าของพื้นที่และผู้อนุญาต - ช่วงเวลาทำการปกติ (ครบ 9 รายชื่ออ้างอิงตามไฟล์แนบ)
list_owners_normal = [
    {"name": "นายวิจักษณ์ ยุ่นชัย", "email": "ywichak@solmax.com"},
    {"name": "นายวีระศักดิ์ พูลเกษม", "email": "pwerasak@solmax.com"},
    {"name": "นายปรม ศรีวิสุทธิ์", "email": "paroms@solmax.com"},
    {"name": "นายสรนัย นาภรณ์", "email": "SorranaiN@solmax.com"},
    {"name": "นางสิริพร ฉายาพรเลิศ", "email": "siripornc@solmax.com"},
    {"name": "นายโกศล บัวลา", "email": "kosolb@solmax.com"},
    {"name": "น.ส.บุหลัน แสนทวีสุข", "email": "bulans@solmax.com"},
    {"name": "นายธนพล เดชนนท์ธนวัฒน์", "email": "dthanaphon@solmax.com"},
    {"name": "นายชัยกร ชารีแก้ว", "email": "ChaiyakornC@solmax.com"}
]

# รายชื่อเจ้าของพื้นที่และผู้อนุญาต - ช่วงนอกเวลาทำการปกติ (ครบ 4 รายชื่ออ้างอิงตามไฟล์แนบ)
list_owners_offhours = [
    {"name": "นายสัมฤทธิ์ ก้องโสตร", "email": "ksamrit@solmax.com"},
    {"name": "นายประภาส แก้วอรสาร", "email": "praphask@solmax.com"},
    {"name": "นายไพฑูรย์ บางกุ้ง", "email": "paitoonb@solmax.com"},
    {"name": "นายสุรรัช เสถียรุจิกานนท์", "email": "ssurarat@solmax.com"}
]

if "permits" not in st.session_state:
    st.session_state.permits = []
if "worker_count" not in st.session_state:
    st.session_state.worker_count = 1

# --- ส่วนจัดวางโครงสร้างเมนูแถบข้าง (Sidebar Menu) ---
if logo_img:
    st.sidebar.image(logo_img, use_container_width=True)
st.sidebar.markdown(f"<div class='sidebar-title'>SOLMAX GEOSYNTHETICS</div>", unsafe_allow_html=True)
st.sidebar.markdown(f"<div class='sidebar-sub'>Work Permit Central Digital Hub</div>", unsafe_allow_html=True)
st.sidebar.write("---")

page = st.sidebar.radio(
    "เมนูระบบ / Menu Navigation", 
    ["📊 แดชบอร์ดภาพรวม / Dashboard Overview", 
     "📝 ยื่นเปิดใบขออนุญาต / Apply New Permit", 
     "🔍 ตรวจสอบและอนุมัติเปิดงาน / Permit Opening Approval",
     "🔒 แจ้งอนุมัติปิดงาน / Permit Closure Panel"]
)

# ================= PAGE 1: DASHBOARD =================
if "แดชบอร์ด" in page:
    st.title("📊 HSE System Real-time Dashboard")
    st.markdown("##### บริษัท โซลแมกซ์ จีโอซินเทติคส์ จำกัด (Solmax Geosynthetics Co., Ltd.)")
    st.write("---")
    
    total = len(st.session_state.permits)
    fully_approved = len([p for p in st.session_state.permits if p["status"] == "Fully Approved 🟢"])
    pending = len([p for p in st.session_state.permits if "Pending" in p["status"] or "HSE" in p["status"]])
    closed = len([p for p in st.session_state.permits if "Closed" in p["status"]])
    
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("ใบงานทั้งหมด / Total Permits", total)
    m2.metric("อนุมัติสมบูรณ์ / Fully Approved 🟢", fully_approved)
    m3.metric("อยู่ระหว่างรออนุมัติ / Pending Logs ⏳", pending)
    m4.metric("ปิดงานสำเร็จ / Closed 🔒", closed)
    
    st.write("---")
    st.subheader("📋 ตารางติดตามสถานะใบอนุญาตทำงาน / Permit Track Board")
    if st.session_state.permits:
        df_show = pd.DataFrame(st.session_state.permits)
        st.dataframe(df_show[["id", "date", "type", "applicant", "shift", "status"]], use_container_width=True, hide_index=True)
    else:
        st.info("💡 ไม่มีประวัติข้อมูลการยื่นใบอนุญาตทำงานในระบบขณะนี้")

# ================= PAGE 2: APPLY REQUEST PERMIT =================
elif "ยื่นเปิดใบขออนุญาต" in page:
    st.title("📝 ยื่นเปิดใบขออนุญาตทำงาน / Apply For Work Permit")
    st.markdown("##### บริษัท โซลแมกซ์ จีโอซินเทติคส์ จำกัด (Solmax Geosynthetics Co., Ltd.)")
    
    st.markdown("### 🔄 ต่ออายุใบอนุญาตเดิม (Permit Extension)")
    existing_ids = [p["id"] for p in st.session_state.permits if p["status"] in ["Fully Approved 🟢", "Pending HSE Approval 👮"]]
    selected_ext_id = st.selectbox("เลือกเลขที่ใบงานเดิมเพื่อต่ออายุทำงานรอบถัดไป (ไม่เกิน 24 ชม.)", ["-- สร้างใบงานใหม่ / Create New --"] + existing_ids)
    
    base_data = {}
    if selected_ext_id != "-- สร้างใบงานใหม่ / Create New --":
        base_data = next((p for p in st.session_state.permits if p["id"] == selected_ext_id), {})
        st.success(f"🔗 ดึงข้อมูลจากใบงานเดิม {selected_ext_id} เรียบร้อยแล้ว")

    now = datetime.datetime.now()
    generated_id = f"WP-{now.strftime('%Y%m%d')}-{random.randint(100, 999)}"
    current_time_str = now.strftime("%Y-%m-%d %H:%M:%S")
    
    col_id1, col_id2 = st.columns(2)
    with col_id1:
        st.info(f"**🔢 เลขที่ใบงานประจำกะ / Permit ID:** {generated_id}")
    with col_id2:
        st.info(f"**⏰ วันที่-เวลาลงทะเบียน / Applied Time:** {current_time_str}")
        
    st.write("---")
    
    permit_type = st.selectbox("🎯 ประเภทใบอนุญาตทำงาน / Select Permit Type", list(permit_master_data.keys()), 
                               index=list(permit_master_data.keys()).index(base_data["type"]) if "type" in base_data else 0)
    
    st.markdown("### ⏱️ รอบกะเวลาปฏิบัติงาน (จำกัดเปิด-ปิดวันต่อวัน) / Shift Control")
    shift_time = st.radio("เลือกรอบเวลาทำงาน / Select Work Shift", 
                          ["รอบกลางวัน (08:00 - 20:00 น.) / Day Shift", "รอบกลางคืน (20:00 - 08:00 น. ของวันถัดไป) / Night Shift"])
    
    manual_off_hours = st.checkbox("งานนี้นอกเวลาทำการปกติ หรือตรงกับวันหยุดเสาร์-อาทิตย์", key="manual_off_hours_check")
    
    # 🔥 แก้ไข LOGIC ตรงนี้: ยึดตามการเลือกของ User เป็นหลัก ไม่โดนเวลาเครื่องคอมพิวเตอร์บังคับครอบงำ
    is_off_hours = False
    if "Night Shift" in shift_time or manual_off_hours:
        is_off_hours = True

    if is_off_hours:
        st.error("⚠️ Status ปัจจุบัน: [นอกเวลาทำการปกติ / กะกลางคืน หรือ วันหยุด] ระบบจะดึงรายชื่อผู้อนุมัติเฉพาะกะนอกเวลาทำการ")
    else:
        st.success("☀️ Status ปัจจุบัน: [เวลาทำการปกติ 08:00 - 17:00 น.]")

    with st.form("premium_solmax_form_v12"):
        st.markdown("#### 🏢 ข้อมูลทั่วไป / General Details")
        c1, c2 = st.columns(2)
        with c1:
            applicant = st.text_input("ชื่อผู้ขออนุญาตทำงาน / Applicant Name", value=base_data.get("applicant", ""))
            company = st.text_input("บริษัทผู้รับเหมา / Contractor Company", value=base_data.get("company", ""))
        with c2:
            area = st.text_input("พื้นที่ปฏิบัติงาน / Work Location Area", value=base_data.get("area", ""))
            description = st.text_input("ลักษณะงานโดยละเอียด / Scope of Work", value=base_data.get("description", ""))
            
        st.markdown("#### 👤 ข้อมูลพนักงานและผลการประเมินความปลอดภัยประจำตัว")
        btn_w1, btn_w2, _ = st.columns([1.5, 1.5, 5])
        with btn_w1:
            if st.form_submit_button("➕ เพิ่มรายชื่อพนักงาน"):
                st.session_state.worker_count += 1
                st.rerun()
        with btn_w2:
            if st.form_submit_button("➖ ลดรายชื่อล่าสุด") and st.session_state.worker_count > 1:
                st.session_state.worker_count -= 1
                st.rerun()
                
        workers_list = []
        for i in range(st.session_state.worker_count):
            w_col1, w_col2 = st.columns([2, 2])
            with w_col1:
                w_name = st.text_input(f"ชื่อ-นามสกุล คนที่ {i+1} / Worker Name {i+1}", key=f"wname_v12_{i}")
            with w_col2:
                w_status = st.selectbox(f"สถานะการอบรมคนที่ {i+1} / Safety Status", 
                                        ["ผ่านการอบรมมีบัตรประจำตัว", "อบรมแล้วรอบัตร", "อบรมชั่วคราว"], key=f"wstat_v12_{i}")
            if w_name:
                workers_list.append({"name": w_name, "training": w_status})

        st.markdown("#### 📁 แนบเอกสารวิเคราะห์ความปลอดภัยและภาพถ่าย")
        cc1, cc2 = st.columns(2)
        with cc1:
            uploaded_jsa = st.file_uploader("แนบเอกสาร JSA / SDS หรือเอกสารที่เกี่ยวข้อง (PDF, PNG, JPG)", type=["pdf", "png", "jpg", "jpeg"], key="jsa_v12")
        with cc2:
            uploaded_site = st.file_uploader("อัปโหลดรูปภาพสภาพหน้างานปัจจุบัน / Upload Site Photo", type=["png", "jpg", "jpeg"], key="site_v12")
            
        st.markdown("#### 🛡️ รายการตรวจสอบความปลอดภัยประจำฟอร์ม / Safety Checklist")
        target_data = permit_master_data[permit_type]
        
        st.markdown(f"**📋 มาตรการป้องกันอันตรายก่อนเริ่มงาน ({permit_type.split(':')[0]}):**")
        checklist_responses = {}
        for idx, item in enumerate(target_data["checklist"]):
            checklist_responses[f"check_{idx}"] = st.checkbox(item, key=f"chk_v12_{idx}")
            
        st.markdown("**🦺 อุปกรณ์ป้องกันอันตรายส่วนบุคคลภาคบังคับ (Required PPE):**")
        ppe_responses = {}
        for idx, ppe_item in enumerate(target_data["ppe"]):
            ppe_responses[f"ppe_{idx}"] = st.checkbox(f"เตรียมสวมใส่: {ppe_item}", key=f"ppe_v12_{idx}", value=True)
            
        st.markdown("#### 👤 ลำดับผู้พิจารณาอนุมัติเปิดงาน / Step Approval Routing")
        selected_sup = st.selectbox("1. ผู้ควบคุมงาน / Supervisor In Charge", [f"{s['name']} [{s['email']}]" for s in list_supervisors])
        
        if is_off_hours:
            owner_options = [f"{o['name']} [{o['email']}]" for o in list_owners_offhours]
            all_issuers_options = [f"{o['name']} [{o['email']}]" for o in list_owners_offhours]
        else:
            owner_options = [f"{o['name']} [{o['email']}]" for o in list_owners_normal]
            all_issuers_options = [f"{o['name']} [{o['email']}]" for o in list_owners_normal]
            
        selected_owner = st.selectbox("2. เจ้าของพื้นที่ / Area Owner", owner_options)
        selected_issuer = st.selectbox("3. ผู้อนุญาต / Permit Issuer", all_issuers_options, index=min(1, len(all_issuers_options)-1))
        
        if is_off_hours:
            selected_hse_name = st.selectbox("4. เจ้าหน้าที่ความปลอดภัย (จป.) / Safety Officer", all_issuers_options, key="hse_offhours_selectbox")
            hse_chosen = selected_hse_name.split(" [")[0]
        else:
            hse_chosen = "น.ส.บุหลัน แสนทวีสุข"
            st.info(f"ℹ️ **4. เจ้าหน้าที่ความปลอดภัย (จป.) ประจำกะ:** {hse_chosen}")
            
        submit_btn = st.form_submit_button("🚀 ยื่นคำขอเปิดใบอนุญาตเข้าระบบ / Submit Permit Request")
        
        if submit_btn:
            if not applicant or not company or not area:
                st.error("❌ กรุณากรอกข้อมูลฟิลด์หลักให้ครบถ้วนก่อนส่งข้อมูล")
            elif not all(checklist_responses.values()):
                st.error("❌ ต้องกดยืนยันตรวจสอบมาตรการความปลอดภัย Safety Checklist ให้ครบทุกข้อเพื่อยืนยันความปลอดภัยหน้างาน")
            elif not all(ppe_responses.values()):
                st.error("❌ ต้องกดรับรองการสวมใส่อุปกรณ์คุ้มครองความปลอดภัยส่วนบุคคล (PPE) ครบถ้วน")
            else:
                new_permit = {
                    "id": generated_id,
                    "date": current_time_str,
                    "timestamp": None,
                    "type": permit_type,
                    "applicant": applicant,
                    "company": company,
                    "area": area,
                    "description": description,
                    "shift": shift_time,
                    "workers": workers_list,
                    "supervisor": selected_sup.split(" [")[0],
                    "owner": selected_owner.split(" [")[0],
                    "issuer": selected_issuer.split(" [")[0],
                    "hse": hse_chosen,
                    "is_off_hours": is_off_hours,
                    "sup_approved": False,
                    "owner_approved": False,
                    "issuer_approved": False,
                    "hse_approved": False,
                    "status": "Pending Step 1: Supervisor Approval ⏳",
                    "closed_sup_approved": False,
                    "closed_owner_approved": False,
                    "closed_issuer_approved": False
                }
                st.session_state.permits.append(new_permit)
                
                upload_db_to_sharepoint()
                st.success(f"✅ ยื่นคำขอเปิดใบอนุญาตสำเร็จ! ใบงานประเภท {permit_type.split(':')[0]} เลขที่ {generated_id} ถูกส่งเข้าระบบและอัปเดตไปที่ SharePoint เรียบร้อยแล้ว")

# ================= PAGE 3: APPROVAL PANEL =================
elif "ตรวจสอบและอนุมัติเปิดงาน" in page:
    st.title("🔍 แผงควบคุมตรวจอนุมัติเปิดงานตามลำดับขั้น / Permit Opening Workflow")
    st.write("---")
    
    active_permits = [p for p in st.session_state.permits if "Pending" in p["status"] or p["status"] == "Pending HSE Approval 👮"]
    
    if not active_permits:
        st.success("🎉 ปัจจุบันไม่มีใบอนุญาตทำงานค้างพิจารณาเปิดงาน")
    else:
        for idx, permit in enumerate(active_permits):
            is_p_off = permit.get("is_off_hours", False)
            time_tag = " [นอกเวลาทำการ 🌙]" if is_p_off else " [เวลาปกติ ☀️]"
            
            with st.expander(f"📄 ใบงาน เลขที่: {permit['id']} | ประเภท: {permit['type'].split(':')[0]} | Status: {permit['status']}", expanded=True):
                st.markdown(f"**🎯 ชนิดงานชี้เฉพาะ:** {permit['type']} | **⏱️ รอบเวลา:** {permit['shift']}")
                st.markdown(f"**🏢 บริษัทผู้รับเหมา:** {permit['company']} | **📍 พื้นที่รับผิดชอบ:** {permit['area']}")
                
                st.write("---")
                st.markdown("##### 🛡️ 4. ตรวจสอบสิทธิ์ขั้นตอนสุดท้าย (เจ้าหน้าที่ความปลอดภัย จป.)")
                
                if is_p_off:
                    st.info(f"👤 ผู้มีสิทธิ์อนุมัติเปิดงานในฐานะ จป. นอกเวลา: **{permit['hse']}**")
                else:
                    if permit['issuer_approved'] and permit.get("timestamp") is not None:
                        elapsed_time = time.time() - permit["timestamp"]
                        time_remaining = 600 - elapsed_time
                        
                        if time_remaining > 0 and not permit['hse_approved']:
                            mins, secs = divmod(int(time_remaining), 60)
                            st.warning(f"⏳ ส่งต่อถึง จป. แล้ว! อยู่ในระหว่างเวลาล็อกสิทธิ์ของ ({permit['hse']}) เหลือเวลานับถอยหลัง: {mins:02d}:{secs:02d} นาที")
                        elif time_remaining <= 0 and not permit['hse_approved']:
                            st.error(f"⏰ เกิน 10 นาทีหลังจากขั้นที่ 3 อนุมัติแล้ว และ จป. ({permit['hse']}) ยังไม่ได้ดำเนินการในเวลา")
                            
                            sub_options = [f"{o['name']} [{o['email']}]" for o in list_owners_normal]
                            substitute_hse = st.selectbox(
                                "🔄 สิทธิ์การผูกขาดถูกเปิดออก โปรดเลือกผู้อนุมัติร่วมประจำพื้นที่เพื่อพิจารณาลงนามแทน จป.:", 
                                sub_options, 
                                key=f"sub_hse_select_{permit['id']}"
                            )
                            permit['hse'] = substitute_hse.split(" [")[0]
                    else:
                        st.info("⏳ Status: รอให้ขั้นตอนที่ 1, 2 และ 3 ลงนามเสร็จสิ้นก่อน ระบบจึงจะเริ่มนับเวลา 10 นาทีของ จป.")
                
                st.write("---")
                st.markdown("##### 🚥 ลำดับสถานะการเซ็นเปิดงานปัจจุบัน:")
                st.write(f"- 1. ผู้ควบคุมงาน ({permit['supervisor']}): {'✅ อนุมัติแล้ว' if permit['sup_approved'] else '⏳ รอการลงนาม'}")
                st.write(f"- 2. เจ้าของพื้นที่ ({permit['owner']}): {'✅ อนุมัติแล้ว' if permit['owner_approved'] else '⏳ รอการลงนาม'}")
                st.write(f"- 3. ผู้อนุญาต ({permit['issuer']}): {'✅ อนุมัติแล้ว' if permit['issuer_approved'] else '⏳ รอการลงนาม'}")
                st.write(f"- 4. ผู้อนุมัติสิทธิ์ จป./ผู้แทนความปลอดภัย ({permit['hse']}): {'✅ อนุมัติเปิดงานสมบูรณ์' if permit['hse_approved'] else '⏳ รอขั้นตอนก่อนหน้าเสร็จสิ้น'}")
                
                st.write("---")
                b_col1, b_col2, b_col3, b_col4 = st.columns(4)
                
                with b_col1:
                    if not permit['sup_approved']:
                        if st.button(f"✍️ 1. อนุมัติ ({permit['supervisor']})", key=f"btn_sup_{permit['id']}"):
                            permit['sup_approved'] = True
                            permit['status'] = "Pending Step 2: Area Owner Approval ⏳"
                            upload_db_to_sharepoint() 
                            st.rerun()
                with b_col2:
                    if permit['sup_approved'] and not permit['owner_approved']:
                        if st.button(f"✍️ 2. อนุมัติ ({permit['owner']})", key=f"btn_own_{permit['id']}"):
                            permit['owner_approved'] = True
                            permit['status'] = "Pending Step 3: Permit Issuer Approval ⏳"
                            upload_db_to_sharepoint()
                            st.rerun()
                with b_col3:
                    if permit['sup_approved'] and permit['owner_approved'] and not permit['issuer_approved']:
                        if st.button(f"✍️ 3. อนุมัติ ({permit['issuer']})", key=f"btn_iss_{permit['id']}"):
                            permit['issuer_approved'] = True
                            permit['status'] = "Pending HSE Approval 👮"
                            permit['timestamp'] = time.time()
                            upload_db_to_sharepoint()
                            st.rerun()
                with b_col4:
                    if permit['sup_approved'] and permit['owner_approved'] and permit['issuer_approved'] and not permit['hse_approved']:
                        if st.button(f"✍️ 4. อนุมัติปิดท้าย ({permit['hse']})", key=f"btn_hse_{permit['id']}"):
                            permit['hse_approved'] = True
                            permit['status'] = "Fully Approved 🟢"
                            upload_db_to_sharepoint()
                            st.success("เปิดใบอนุญาตทำงานเสร็จสิ้นสมบูรณ์!")
                            st.rerun()
                            
        if any(p['issuer_approved'] and not p['hse_approved'] and not p['is_off_hours'] for p in active_permits):
            st.button("🔄 อัปเดตเวลาตัวนับถอยหลังหน้างาน (Refresh Timer)")

# ================= PAGE 4: PERMIT CLOSURE PANEL =================
elif "แจ้งอนุมัติปิดงาน" in page:
    st.title("🔒 ระบบแจ้งและอนุมัติปิดงานความปลอดภัย / Permit Closure Panel")
    st.write("---")
    
    open_permits = [p for p in st.session_state.permits if p["status"] in ["Fully Approved 🟢", "Pending Closure ⏳"]]
    
    if not open_permits:
        st.info("💡 ไม่มีใบงานที่อยู่ระหว่างดำเนินการเปิดหน้างานในระบบขณะนี้")
    else:
        for idx, permit in enumerate(open_permits):
            with st.expander(f"📦 ใบงานแจ้งปิดงาน เลขที่: {permit['id']} | ผู้ขอ: {permit['applicant']}", expanded=True):
                st.markdown(f"**📍 พื้นที่ปฏิบัติงาน:** {permit['area']} | **🎯 ประเภทใบงาน:** {permit['type']}")
                st.write("---")
                
                st.markdown("##### 🚥 Status การเซ็นปิดหน้างาน:")
                st.write(f"- 1. ผู้ควบคุมงาน ({permit['supervisor']}): {'✅ เซ็นปิดแล้ว' if permit['closed_sup_approved'] else '⏳ รอเซ็นปิดงาน'}")
                st.write(f"- 2. เจ้าของพื้นที่ ({permit['owner']}): {'✅ เซ็นปิดแล้ว' if permit['closed_owner_approved'] else '⏳ รอเซ็นปิดงาน'}")
                st.write(f"- 3. ผู้อนุญาต ({permit['issuer']}): {'✅ ปิดสมบูรณ์แล้ว' if permit['closed_issuer_approved'] else '⏳ รอเซ็นปิดงาน'}")
                
                st.write("---")
                bc1, bc2, bc3 = st.columns(3)
                with bc1:
                    if not permit['closed_sup_approved']:
                        if st.button(f"🔒 1. ผู้ควบคุมงาน ปิดงาน", key=f"cl_sup_{permit['id']}"):
                            permit['closed_sup_approved'] = True
                            permit['status'] = "Pending Closure ⏳"
                            upload_db_to_sharepoint() 
                            st.rerun()
                with bc2:
                    if permit['closed_sup_approved'] and not permit['closed_owner_approved']:
                        if st.button(f"🔒 2. เจ้าของพื้นที่ ปิดงาน", key=f"cl_own_{permit['id']}"):
                            permit['closed_owner_approved'] = True
                            upload_db_to_sharepoint()
                            st.rerun()
                with bc3:
                    if permit['closed_sup_approved'] and permit['closed_owner_approved'] and not permit['closed_issuer_approved']:
                        if st.button(f"🔒 3. ผู้อนุญาต ปิดงานสมบูรณ์", key=f"cl_iss_{permit['id']}"):
                            permit['closed_issuer_approved'] = True
                            permit['status'] = "Closed Completed 🔒"
                            upload_db_to_sharepoint()
                            st.success("ปิดใบงานความปลอดภัยเสร็จสิ้นอย่างสมบูรณ์!")
                            st.rerun()