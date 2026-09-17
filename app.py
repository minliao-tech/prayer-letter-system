from flask import Flask, request, redirect, url_for, render_template, send_file, flash
from pathlib import Path
from datetime import date, datetime, timedelta
import sqlite3, io
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)
app.secret_key = "change-this-before-production"
BASE = Path(__file__).parent
DB = BASE / "prayer_letters.db"

GROUPS = {
    "A": {
        "主任牧師辦公室": [], "聖工聯席會": [], "靈糧全球使徒性網絡": [],
        "創意藝術媒體處": [], "宣教植堂處": [], "啟示性事奉處": [],
        "人資財會處": [], "行政資訊處": []
    },
    "B": {
        "創新育成中心": [], "事業處": [], "社會服務處": [],
        "愛鄰協會": [], "白絲帶網安關懷協會": []
    },
    "C": {
        "全人關顧中心": [],
        "牧養支援處": ["牧養企劃部", "教育訓練部"],
        "基層福音事工處": [],
        "總部牧區": ["成人牧區", "北區會堂", "兒童牧區", "天使心牧區", "聯合崇拜",
                    "學生牧區", "職場牧區", "喜樂家族", "台語牧區", "客語牧區"],
        "國際事工中心": ["英語牧區", "印尼牧區", "越南牧區", "菲律賓牧區"],
        "靈糧國度領袖學院": ["神學院", "生命培訓學院", "巴拿巴宣教學院", "職場轉化學院"]
    },
    "D": {
        "福音中心": ["金門", "興隆", "萬金", "古亭", "澎湖", "竹圍", "三樹", "五股", "萬華", "東石"]
    }
}

# 以使用者提供的 2026/09/11 D組文件作為輪值基準：
# 09/11 D → 09/18 A → 09/25 B → 10/02 C → 10/09 D...
ANCHOR_DATE = date(2026, 9, 11)
ORDER = ["D", "A", "B", "C"]

def db():
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    return con

def init_db():
    con = db()
    con.executescript("""
    CREATE TABLE IF NOT EXISTS submissions(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        week_date TEXT NOT NULL,
        group_code TEXT NOT NULL,
        parent_unit TEXT NOT NULL,
        unit_name TEXT NOT NULL,
        contact_name TEXT,
        contact_email TEXT,
        status TEXT NOT NULL DEFAULT 'submitted',
        created_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS prayer_items(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        submission_id INTEGER NOT NULL,
        sort_order INTEGER NOT NULL,
        content TEXT NOT NULL,
        FOREIGN KEY(submission_id) REFERENCES submissions(id) ON DELETE CASCADE
    );
    """)
    con.commit()
    con.close()

def friday_of_week(d=None):
    d = d or date.today()
    # Friday = 4
    return d + timedelta(days=(4 - d.weekday()) % 7) if d.weekday() <= 4 else d - timedelta(days=d.weekday()-4)

def group_for_week(week_date):
    delta_weeks = (week_date - ANCHOR_DATE).days // 7
    return ORDER[delta_weeks % 4]

def current_cycle():
    wd = friday_of_week()
    return wd, group_for_week(wd)

def expected_units(group_code):
    out = []
    for parent, children in GROUPS[group_code].items():
        if children:
            for child in children:
                out.append((parent, child))
        else:
            out.append((parent, parent))
    return out

@app.route("/")
def index():
    wd, group_code = current_cycle()
    return render_template("index.html", week_date=wd.isoformat(), group_code=group_code, groups=GROUPS)

@app.route("/submit", methods=["GET","POST"])
def submit():
    default_week, default_group = current_cycle()
    if request.method == "POST":
        week_date = request.form["week_date"]
        group_code = request.form["group_code"]
        parent_unit = request.form["parent_unit"]
        unit_name = request.form["unit_name"]
        contact_name = request.form.get("contact_name","").strip()
        contact_email = request.form.get("contact_email","").strip()
        items = [x.strip() for x in request.form.getlist("prayer_item") if x.strip()]
        if not items:
            flash("請至少填寫一項代禱事項。")
            return redirect(url_for("submit"))
        con = db()
        cur = con.execute("""INSERT INTO submissions
            (week_date,group_code,parent_unit,unit_name,contact_name,contact_email,status,created_at)
            VALUES(?,?,?,?,?,?,?,?)""",
            (week_date,group_code,parent_unit,unit_name,contact_name,contact_email,"submitted",datetime.now().isoformat(timespec="seconds")))
        sid = cur.lastrowid
        for i, item in enumerate(items, 1):
            con.execute("INSERT INTO prayer_items(submission_id,sort_order,content) VALUES(?,?,?)",(sid,i,item))
        con.commit(); con.close()
        return render_template("thanks.html", unit_name=unit_name, count=len(items))
    return render_template("submit.html", week_date=default_week.isoformat(), group_code=default_group, groups=GROUPS)

@app.route("/api/units/<group_code>")
def units(group_code):
    return GROUPS.get(group_code, {})

@app.route("/admin")
def admin():
    week = request.args.get("week")
    if week:
        wd = date.fromisoformat(week)
    else:
        wd, _ = current_cycle()
    group_code = request.args.get("group") or group_for_week(wd)
    expected = expected_units(group_code)
    con = db()
    rows = con.execute("""SELECT s.*,
        (SELECT COUNT(*) FROM prayer_items p WHERE p.submission_id=s.id) item_count
        FROM submissions s WHERE week_date=? AND group_code=? ORDER BY created_at DESC""",
        (wd.isoformat(),group_code)).fetchall()
    latest = {}
    for r in rows:
        key=(r["parent_unit"],r["unit_name"])
        if key not in latest: latest[key]=r
    status_rows=[]
    for parent, unit in expected:
        status_rows.append({"parent":parent,"unit":unit,"submission":latest.get((parent,unit))})
    con.close()
    return render_template("admin.html", week_date=wd.isoformat(), group_code=group_code,
                           status_rows=status_rows, submitted=sum(1 for x in status_rows if x["submission"]),
                           total=len(status_rows))

@app.route("/admin/edit/<int:sid>", methods=["GET","POST"])
def edit_submission(sid):
    con=db()
    sub=con.execute("SELECT * FROM submissions WHERE id=?",(sid,)).fetchone()
    if not sub:
        con.close(); return "Not found",404
    if request.method=="POST":
        items=[x.strip() for x in request.form.getlist("prayer_item") if x.strip()]
        con.execute("UPDATE submissions SET status=? WHERE id=?",(request.form.get("status","approved"),sid))
        con.execute("DELETE FROM prayer_items WHERE submission_id=?",(sid,))
        for i,item in enumerate(items,1):
            con.execute("INSERT INTO prayer_items(submission_id,sort_order,content) VALUES(?,?,?)",(sid,i,item))
        con.commit(); con.close()
        return redirect(url_for("admin",week=sub["week_date"],group=sub["group_code"]))
    items=con.execute("SELECT * FROM prayer_items WHERE submission_id=? ORDER BY sort_order",(sid,)).fetchall()
    con.close()
    return render_template("edit.html", sub=sub, items=items)

def set_run_font(run, size=12, bold=False, color=None):
    run.font.name = "Microsoft JhengHei"
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)

@app.route("/export")
def export_word():
    week = request.args.get("week")
    group_code = request.args.get("group")
    if not week or not group_code:
        return "Missing week/group",400
    con=db()
    rows=con.execute("""SELECT * FROM submissions
        WHERE week_date=? AND group_code=? AND status IN ('submitted','approved')
        ORDER BY id""",(week,group_code)).fetchall()
    latest={}
    for r in rows:
        latest[(r["parent_unit"],r["unit_name"])]=r

    doc=Document()
    sec=doc.sections[0]
    title=doc.add_paragraph()
    title.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=title.add_run(f"{week.replace('-','')} 教會代禱信 {group_code}組")
    set_run_font(r,16,True)

    p=doc.add_paragraph()
    r=p.add_run("【為所提單位的同工禱告】")
    set_run_font(r,12,True,(192,0,0))

    for parent, children in GROUPS[group_code].items():
        units=children if children else [parent]
        if children:
            p=doc.add_paragraph()
            r=p.add_run(parent)
            set_run_font(r,13,True,(0,102,204))
        for unit in units:
            sub=latest.get((parent,unit))
            if not sub: 
                continue
            display = f"{unit}{parent}" if group_code=="D" and parent=="福音中心" else unit
            p=doc.add_paragraph()
            r=p.add_run(f"【{display}】")
            set_run_font(r,12,True,(0,102,204))
            items=con.execute("SELECT * FROM prayer_items WHERE submission_id=? ORDER BY sort_order",(sub["id"],)).fetchall()
            for i,item in enumerate(items,1):
                p=doc.add_paragraph()
                p.paragraph_format.space_after=Pt(4)
                r=p.add_run(f"{i}. {item['content']}")
                set_run_font(r,12,False,(0,128,0))
    con.close()
    bio=io.BytesIO()
    doc.save(bio); bio.seek(0)
    filename=f"{week.replace('-','')} 教會代禱信 {group_code}組.docx"
    return send_file(bio,as_attachment=True,download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

if __name__ == "__main__":
    init_db()
    app.run(debug=True, port=5000)
