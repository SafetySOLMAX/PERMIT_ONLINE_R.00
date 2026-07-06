
import streamlit as st
import pandas as pd
import sqlite3, os, json, uuid, socket, urllib.parse
from pathlib import Path
from datetime import datetime, date, time
from PIL import Image

APP_TITLE = "SOLMAX Online Permit to Work"
DB_PATH = "data/ptw_streamlit.db"
UPLOAD_DIR = Path("uploads")
UPLOAD_DIR.mkdir(exist_ok=True)
Path("data").mkdir(exist_ok=True)

APPROVER_ROLES = ["ผู้ควบคุมงาน", "ผู้อนุญาต", "เจ้าของพื้นที่", "เจ้าหน้าที่ความปลอดภัย"]
WORK_TYPES = ["General Work", "Hot Work", "Confined Space", "Work at Height", "Excavation", "Electrical Work", "Lifting Work", "Other"]
SAFETY_CHECKS = [
    "ได้ปิดกั้น/ตัดแยกพลังงาน (LOTO) เรียบร้อยแล้ว",
    "ได้ปิดกั้นพื้นที่ด้วยเทปขาว-แดงและป้ายชี้บ่งอย่างชัดเจน",
    "ผู้ปฏิบัติงานผ่านการอบรมและมีบัตรประจำตัวผู้รับเหมา",
    "เครื่องมือกล/อุปกรณ์ไฟฟ้าผ่านการตรวจสอบและติดสติกเกอร์สีเขียว",
    "ขออนุญาตใช้กระแสไฟ/ต่อสายไฟกับช่างไฟฟ้าแล้ว",
    "มีการประชุม Toolbox Talk ก่อนเริ่มงาน",
    "มีแผนฉุกเฉินและอุปกรณ์ระงับเหตุเหมาะสม",
]
PPE_LIST = ["หมวกนิรภัย", "แว่นตานิรภัย/กระบังหน้า", "รองเท้านิรภัย", "ถุงมือ", "อุปกรณ์ป้องกันหู", "หน้ากากกรองสารเคมี/ฝุ่น", "Full body harness", "ชุดกันไฟ/FR clothing", "อื่น ๆ"]
HAZARDS = ["ไฟ/ประกายไฟ", "ไฟฟ้า", "สารเคมี", "ตกจากที่สูง", "อับอากาศ", "งานขุด", "เครื่องจักรเคลื่อนที่", "งานยก", "ฝุ่น/เสียง/ความร้อน", "ลื่น สะดุด หกล้ม"]

st.set_page_config(page_title=APP_TITLE, page_icon="✅", layout="wide")

CUSTOM_CSS = """
<style>
.block-container {padding-top: 1.2rem;}
.main-title {font-size: 1.8rem; font-weight: 800; color:#0b6b50;}
.sub {color:#6b7280;}
.kpi {padding:18px; border-radius:18px; background:#ffffff; border-left:7px solid #0b6b50; box-shadow:0 3px 12px rgba(0,0,0,.06)}
.kpi h2 {margin:0; font-size:2rem; color:#1f2937}
.kpi span {color:#6b7280; font-weight:600}
.status-approved {color:white;background:#0b6b50;padding:4px 10px;border-radius:999px;}
.status-submitted {color:white;background:#2563eb;padding:4px 10px;border-radius:999px;}
.status-review {color:white;background:#d97706;padding:4px 10px;border-radius:999px;}
.status-rejected {color:white;background:#dc2626;padding:4px 10px;border-radius:999px;}
.status-closed {color:white;background:#374151;padding:4px 10px;border-radius:999px;}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with db() as c:
        c.execute("""CREATE TABLE IF NOT EXISTS permits(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            doc_no TEXT UNIQUE,
            created_at TEXT,
            requester TEXT,
            requester_phone TEXT,
            company TEXT,
            department TEXT,
            area TEXT,
            persons_count INTEGER,
            work_type TEXT,
            work_description TEXT,
            tools TEXT,
            start_datetime TEXT,
            end_datetime TEXT,
            electrical_related TEXT,
            hazards TEXT,
            safety_checks TEXT,
            ppe TEXT,
            controls TEXT,
            emergency_plan TEXT,
            photos TEXT,
            status TEXT DEFAULT 'Submitted'
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS approvals(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            permit_id INTEGER,
            role TEXT,
            approver_name TEXT,
            approver_email TEXT,
            token TEXT UNIQUE,
            decision TEXT DEFAULT 'Pending',
            comment TEXT,
            decided_at TEXT
        )""")
init_db()


def local_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = "127.0.0.1"
    finally:
        s.close()
    return ip


def base_url():
    # ใช้ IP LAN เพื่อให้โทรศัพท์ใน Wi-Fi เดียวกันเปิดได้
    return f"http://{local_ip()}:8501"


def next_doc_no():
    prefix = datetime.now().strftime("PTW-%Y%m-")
    with db() as c:
        row = c.execute("SELECT doc_no FROM permits WHERE doc_no LIKE ? ORDER BY doc_no DESC LIMIT 1", (prefix+"%",)).fetchone()
    next_no = int(row["doc_no"].split("-")[-1]) + 1 if row else 1
    return f"{prefix}{next_no:04d}"


def recalc_status(permit_id):
    with db() as c:
        rows = c.execute("SELECT decision FROM approvals WHERE permit_id=?", (permit_id,)).fetchall()
        decisions = [r["decision"] for r in rows]
        if any(d == "Rejected" for d in decisions):
            status = "Rejected"
        elif decisions and all(d == "Approved" for d in decisions):
            status = "Approved"
        elif any(d == "Approved" for d in decisions):
            status = "In Review"
        else:
            status = "Submitted"
        c.execute("UPDATE permits SET status=? WHERE id=?", (status, permit_id))


def get_permit(permit_id):
    with db() as c:
        p = c.execute("SELECT * FROM permits WHERE id=?", (permit_id,)).fetchone()
        a = c.execute("SELECT * FROM approvals WHERE permit_id=? ORDER BY id", (permit_id,)).fetchall()
    return p, a


def all_permits_df():
    with db() as c:
        rows = c.execute("SELECT * FROM permits ORDER BY id DESC").fetchall()
    return pd.DataFrame([dict(r) for r in rows])


def save_uploaded_photos(files, doc_no):
    saved=[]
    folder = UPLOAD_DIR / doc_no
    folder.mkdir(exist_ok=True)
    for f in files:
        if f is None: continue
        suffix = Path(f.name).suffix.lower()
        safe_name = f"{datetime.now():%Y%m%d%H%M%S}_{uuid.uuid4().hex[:8]}{suffix}"
        path = folder / safe_name
        path.write_bytes(f.getbuffer())
        saved.append(str(path))
    return saved


def status_badge(status):
    cls = {
        "Approved":"status-approved", "Submitted":"status-submitted", "In Review":"status-review",
        "Rejected":"status-rejected", "Closed":"status-closed"
    }.get(status, "status-submitted")
    return f"<span class='{cls}'>{status}</span>"


def mailto_link(email, subject, body):
    return "mailto:" + urllib.parse.quote(email) + "?subject=" + urllib.parse.quote(subject) + "&body=" + urllib.parse.quote(body)


def approval_url(token):
    return f"{base_url()}/?token={token}"


def dashboard():
    st.markdown(f"<div class='main-title'>{APP_TITLE}</div><div class='sub'>Dashboard สำหรับติดตามใบขออนุญาตทำงานแบบ Real-time</div>", unsafe_allow_html=True)
    df = all_permits_df()
    if df.empty:
        st.info("ยังไม่มีข้อมูล Permit — ไปที่เมนู 'สร้าง Permit ใหม่' เพื่อเริ่มใช้งาน")
        return
    c1,c2,c3,c4,c5 = st.columns(5)
    total=len(df)
    pending=int(df['status'].isin(['Submitted','In Review']).sum())
    approved=int((df['status']=='Approved').sum())
    rejected=int((df['status']=='Rejected').sum())
    today=int(pd.to_datetime(df['created_at']).dt.date.eq(date.today()).sum())
    for col,label,value in [(c1,'ทั้งหมด',total),(c2,'วันนี้',today),(c3,'รออนุมัติ',pending),(c4,'อนุมัติแล้ว',approved),(c5,'ไม่อนุมัติ',rejected)]:
        col.markdown(f"<div class='kpi'><span>{label}</span><h2>{value}</h2></div>", unsafe_allow_html=True)
    st.divider()
    left,right = st.columns([1,1])
    with left:
        st.subheader("สถานะ Permit")
        st.bar_chart(df['status'].value_counts())
    with right:
        st.subheader("ประเภทงาน")
        st.bar_chart(df['work_type'].value_counts())
    st.subheader("รายการ Permit ล่าสุด")
    show = df[['doc_no','created_at','requester','company','area','work_type','start_datetime','end_datetime','status']].copy()
    st.dataframe(show, use_container_width=True, hide_index=True)


def new_permit_form():
    st.markdown("## 📝 สร้างใบขออนุญาตทำงาน / New Permit to Work")
    doc_no = next_doc_no()
    created_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with st.form("ptw_form", clear_on_submit=False):
        st.markdown("### ส่วนที่ 1: ข้อมูลทั่วไป")
        c1,c2 = st.columns(2)
        c1.text_input("เลขที่เอกสาร / Document No.", value=doc_no, disabled=True)
        c2.text_input("วันที่และเวลายื่นคำขอ", value=created_at, disabled=True)
        c1,c2,c3 = st.columns(3)
        requester = c1.text_input("ชื่อผู้ขออนุญาตทำงาน", placeholder="ชื่อ-สกุล")
        requester_phone = c2.text_input("เบอร์ติดต่อ", placeholder="08x-xxx-xxxx")
        company = c3.text_input("บริษัท/ผู้รับเหมา")
        c1,c2,c3 = st.columns(3)
        department = c1.text_input("หน่วยงาน/แผนก")
        area = c2.text_input("พื้นที่ปฏิบัติงาน")
        persons_count = c3.number_input("จำนวนผู้เข้าปฏิบัติงาน", min_value=1, max_value=100, value=1)
        work_type = st.selectbox("ประเภทงาน", WORK_TYPES)
        work_description = st.text_area("ลักษณะงานโดยสังเขป / Work Description", height=120)
        tools = st.text_area("อุปกรณ์/เครื่องมือที่ใช้", height=80)
        c1,c2,c3 = st.columns(3)
        start_d = c1.date_input("วันที่เริ่มงาน", value=date.today())
        start_t = c2.time_input("เวลาเริ่มงาน", value=time(8,0))
        end_t = c3.time_input("เวลาสิ้นสุดโดยประมาณ", value=time(17,0))
        end_d = start_d
        electrical_related = st.radio("งานที่ต้องเข้าระบบไฟฟ้าหรืออุปกรณ์ควบคุม", ["ไม่เกี่ยวข้อง", "เกี่ยวข้อง"], horizontal=True)
        st.markdown("### ส่วนที่ 2: การประเมินความเสี่ยงและมาตรการควบคุม")
        hazards = st.multiselect("อันตรายที่เกี่ยวข้อง", HAZARDS)
        safety_checks = st.multiselect("การตรวจสอบความปลอดภัยก่อนเริ่มงาน", SAFETY_CHECKS)
        ppe = st.multiselect("PPE ที่ต้องใช้", PPE_LIST)
        controls = st.text_area("มาตรการควบคุมเพิ่มเติม / Controls", height=100)
        emergency_plan = st.text_area("แผนฉุกเฉิน / Emergency Plan", height=80)
        photos = st.file_uploader("แนบรูปหน้างานเพื่อประกอบการอนุมัติ", type=['png','jpg','jpeg'], accept_multiple_files=True)
        st.markdown("### ส่วนที่ 3: ผู้อนุมัติ")
        approvers = []
        for role in APPROVER_ROLES:
            c1,c2 = st.columns(2)
            name = c1.text_input(f"{role} - ชื่อ", key=f"{role}_name")
            email = c2.text_input(f"{role} - อีเมล", key=f"{role}_email")
            approvers.append((role,name,email))
        submitted = st.form_submit_button("ส่งคำขออนุมัติ", type="primary")
    if submitted:
        if not requester or not company or not area or not work_description:
            st.error("กรุณากรอกข้อมูลสำคัญ: ผู้ขอ บริษัท พื้นที่ และรายละเอียดงาน")
            return
        saved_photos = save_uploaded_photos(photos, doc_no)
        start_dt = datetime.combine(start_d, start_t).strftime('%Y-%m-%d %H:%M')
        end_dt = datetime.combine(end_d, end_t).strftime('%Y-%m-%d %H:%M')
        with db() as c:
            cur = c.execute("""INSERT INTO permits(doc_no,created_at,requester,requester_phone,company,department,area,persons_count,work_type,work_description,tools,start_datetime,end_datetime,electrical_related,hazards,safety_checks,ppe,controls,emergency_plan,photos,status)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", (doc_no,created_at,requester,requester_phone,company,department,area,persons_count,work_type,work_description,tools,start_dt,end_dt,electrical_related,json.dumps(hazards,ensure_ascii=False),json.dumps(safety_checks,ensure_ascii=False),json.dumps(ppe,ensure_ascii=False),controls,emergency_plan,json.dumps(saved_photos,ensure_ascii=False),'Submitted'))
            pid = cur.lastrowid
            for role,name,email in approvers:
                c.execute("INSERT INTO approvals(permit_id,role,approver_name,approver_email,token) VALUES(?,?,?,?,?)", (pid,role,name,email,str(uuid.uuid4())))
        st.success(f"สร้างคำขอ {doc_no} สำเร็จ")
        st.info("ไปที่เมนู 'รายละเอียด/Approval Links' เพื่อ copy ลิงก์หรือ mailto สำหรับส่งอีเมลอนุมัติ")
        st.session_state['last_pid']=pid


def detail_and_links():
    st.markdown("## 🔗 รายละเอียด Permit / Approval Links")
    df = all_permits_df()
    if df.empty:
        st.info("ยังไม่มี Permit")
        return
    options = [f"{r.doc_no} | {r.requester} | {r.area} | {r.status}" for r in df.itertuples()]
    selected = st.selectbox("เลือก Permit", options)
    doc_no = selected.split(" | ")[0]
    permit_id = int(df.loc[df['doc_no']==doc_no,'id'].iloc[0])
    p,a = get_permit(permit_id)
    st.markdown(f"### {p['doc_no']} {status_badge(p['status'])}", unsafe_allow_html=True)
    c1,c2,c3 = st.columns(3)
    c1.write(f"**ผู้ขอ:** {p['requester']}")
    c2.write(f"**บริษัท:** {p['company']}")
    c3.write(f"**พื้นที่:** {p['area']}")
    st.write(f"**รายละเอียดงาน:** {p['work_description']}")
    st.write(f"**มาตรการควบคุม:** {p['controls']}")
    st.write(f"**ช่วงเวลา:** {p['start_datetime']} ถึง {p['end_datetime']}")
    photos=json.loads(p['photos'] or '[]')
    if photos:
        st.markdown("#### รูปหน้างาน")
        cols=st.columns(3)
        for i,ph in enumerate(photos):
            if Path(ph).exists(): cols[i%3].image(ph, use_container_width=True)
    st.markdown("### ลิงก์อนุมัติสำหรับส่งอีเมล")
    for x in a:
        url = approval_url(x['token'])
        subject = f"Approval Required: {p['doc_no']} - Permit to Work"
        body = f"เรียน {x['approver_name']}\n\nกรุณาตรวจสอบและอนุมัติใบขออนุญาตทำงาน {p['doc_no']} ก่อนเริ่มงาน\n\nลิงก์อนุมัติ: {url}\n\nขอบคุณครับ/ค่ะ"
        st.write(f"**{x['role']}** | {x['approver_name']} | {x['approver_email']} | สถานะ: {x['decision']}")
        st.code(url)
        if x['approver_email']:
            st.markdown(f"[เปิดอีเมลเพื่อส่งลิงก์อนุมัติ]({mailto_link(x['approver_email'],subject,body)})")
    st.markdown("### QR Code สำหรับเปิดระบบ")
    try:
        import qrcode
        qr_path = Path('data/ptw_qr.png')
        qrcode.make(base_url()).save(qr_path)
        st.image(str(qr_path), width=220)
        st.code(base_url())
    except Exception:
        st.warning("ถ้าต้องการ QR Code ให้ติดตั้ง package: py -m pip install qrcode pillow")


def approval_page(token):
    with db() as c:
        appr = c.execute("SELECT * FROM approvals WHERE token=?", (token,)).fetchone()
    if not appr:
        st.error("ลิงก์อนุมัติไม่ถูกต้อง")
        return
    p,a = get_permit(appr['permit_id'])
    st.markdown(f"## ✅ Approval: {p['doc_no']}")
    st.write(f"**บทบาท:** {appr['role']}")
    st.write(f"**ผู้อนุมัติ:** {appr['approver_name']}")
    st.write(f"**พื้นที่:** {p['area']} | **ประเภทงาน:** {p['work_type']}")
    st.write(f"**รายละเอียดงาน:** {p['work_description']}")
    st.write(f"**มาตรการควบคุม:** {p['controls']}")
    photos=json.loads(p['photos'] or '[]')
    if photos:
        cols=st.columns(3)
        for i,ph in enumerate(photos):
            if Path(ph).exists(): cols[i%3].image(ph, use_container_width=True)
    if appr['decision'] != 'Pending':
        st.info(f"รายการนี้ถูกดำเนินการแล้ว: {appr['decision']} เมื่อ {appr['decided_at']}")
        return
    comment = st.text_area("ความคิดเห็นประกอบการอนุมัติ")
    c1,c2 = st.columns(2)
    if c1.button("อนุมัติ", type="primary"):
        with db() as c:
            c.execute("UPDATE approvals SET decision=?, comment=?, decided_at=? WHERE token=?", ('Approved',comment,datetime.now().strftime('%Y-%m-%d %H:%M:%S'),token))
        recalc_status(appr['permit_id'])
        st.success("บันทึกผลอนุมัติแล้ว")
        st.rerun()
    if c2.button("ไม่อนุมัติ"):
        with db() as c:
            c.execute("UPDATE approvals SET decision=?, comment=?, decided_at=? WHERE token=?", ('Rejected',comment,datetime.now().strftime('%Y-%m-%d %H:%M:%S'),token))
        recalc_status(appr['permit_id'])
        st.error("บันทึกผลไม่อนุมัติแล้ว")
        st.rerun()


def template_explorer():
    st.markdown("## 📄 อ่าน Excel ต้นแบบ")
    st.write("อัปโหลด Excel ต้นแบบเพื่อให้ระบบดึงข้อความ/หัวข้อคำถามออกมาเป็นตาราง สำหรับใช้ปรับฟอร์ม Streamlit ต่อ")
    f = st.file_uploader("อัปโหลดไฟล์ Excel", type=['xlsx','xls'])
    if f:
        try:
            xls = pd.ExcelFile(f)
            sheet = st.selectbox("เลือก Sheet", xls.sheet_names)
            raw = pd.read_excel(f, sheet_name=sheet, header=None, dtype=str).fillna("")
            rows=[]
            for r in range(raw.shape[0]):
                vals=[str(v).strip() for v in raw.iloc[r].tolist() if str(v).strip()]
                if vals:
                    rows.append({"row":r+1,"text":" | ".join(vals)})
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
        except Exception as e:
            st.error(f"อ่านไฟล์ไม่สำเร็จ: {e}")


def export_data():
    st.markdown("## ⬇️ Export ข้อมูล")
    df=all_permits_df()
    if df.empty:
        st.info("ยังไม่มีข้อมูล")
        return
    st.download_button("ดาวน์โหลดข้อมูล Permit เป็น CSV", df.to_csv(index=False).encode('utf-8-sig'), file_name="ptw_permits.csv", mime="text/csv")
    st.dataframe(df, use_container_width=True, hide_index=True)


def main():
    qp = st.query_params
    token = qp.get("token", None)
    if token:
        approval_page(token)
        return
    with st.sidebar:
        st.image("https://dummyimage.com/600x120/0b6b50/ffffff&text=SOLMAX+PTW", use_container_width=True)
        menu = st.radio("เมนู", ["Dashboard", "สร้าง Permit ใหม่", "รายละเอียด/Approval Links", "อ่าน Excel ต้นแบบ", "Export ข้อมูล"])
        st.caption("หมายเหตุ: หากให้มือถือสแกนได้ ต้องรันด้วยคำสั่ง --server.address 0.0.0.0 และมือถืออยู่ Network เดียวกัน")
    if menu == "Dashboard": dashboard()
    elif menu == "สร้าง Permit ใหม่": new_permit_form()
    elif menu == "รายละเอียด/Approval Links": detail_and_links()
    elif menu == "อ่าน Excel ต้นแบบ": template_explorer()
    elif menu == "Export ข้อมูล": export_data()

if __name__ == "__main__":
    main()
