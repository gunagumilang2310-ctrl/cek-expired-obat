from datetime import date, datetime
from io import BytesIO
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import or_
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

app = Flask(__name__)
app.config["SECRET_KEY"] = "ganti-dengan-secret-key-anda"
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///obat.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db = SQLAlchemy(app)

class Obat(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    kode = db.Column(db.String(40), unique=True, nullable=False)
    nama = db.Column(db.String(150), nullable=False)
    kategori = db.Column(db.String(100), default="-")
    batch = db.Column(db.String(80), default="-")
    stok = db.Column(db.Integer, nullable=False, default=0)
    satuan = db.Column(db.String(30), default="pcs")
    tanggal_masuk = db.Column(db.Date, nullable=True)
    tanggal_expired = db.Column(db.Date, nullable=False)
    lokasi = db.Column(db.String(100), default="-")
    keterangan = db.Column(db.Text, default="")

def status_obat(exp):
    sisa = (exp - date.today()).days
    if sisa < 0: return "Expired", sisa
    if sisa <= 30: return "Kritis", sisa
    if sisa <= 90: return "Segera", sisa
    return "Aman", sisa

@app.context_processor
def inject_today():
    return {"today": date.today()}

@app.route("/")
def dashboard():
    items = Obat.query.order_by(Obat.tanggal_expired.asc()).all()
    stats = {"total":len(items), "expired":0, "kritis":0, "segera":0, "aman":0}
    alerts=[]
    for o in items:
        st, days = status_obat(o.tanggal_expired)
        if st == "Expired": stats["expired"] += 1
        elif st == "Kritis": stats["kritis"] += 1
        elif st == "Segera": stats["segera"] += 1
        else: stats["aman"] += 1
        if st != "Aman": alerts.append((o,st,days))
    return render_template("dashboard.html", stats=stats, alerts=alerts[:8])

@app.route("/obat")
def daftar_obat():
    q=request.args.get("q","").strip()
    status=request.args.get("status","")
    query=Obat.query
    if q:
        query=query.filter(or_(Obat.nama.ilike(f"%{q}%"), Obat.kode.ilike(f"%{q}%"), Obat.batch.ilike(f"%{q}%")))
    items=query.order_by(Obat.tanggal_expired.asc()).all()
    if status:
        items=[o for o in items if status_obat(o.tanggal_expired)[0]==status]
    return render_template("obat.html", items=items, q=q, status=status)

@app.route("/obat/tambah", methods=["GET","POST"])
@app.route("/obat/<int:oid>/edit", methods=["GET","POST"])
def form_obat(oid=None):
    item=Obat.query.get_or_404(oid) if oid else None
    if request.method=="POST":
        try:
            kode=request.form["kode"].strip()
            dupe=Obat.query.filter_by(kode=kode).first()
            if dupe and (not item or dupe.id != item.id):
                raise ValueError("Kode obat sudah digunakan.")
            if not item: item=Obat()
            item.kode=kode
            item.nama=request.form["nama"].strip()
            item.kategori=request.form.get("kategori","-").strip() or "-"
            item.batch=request.form.get("batch","-").strip() or "-"
            item.stok=int(request.form["stok"])
            if item.stok < 0: raise ValueError("Stok tidak boleh negatif.")
            item.satuan=request.form.get("satuan","pcs").strip() or "pcs"
            item.tanggal_masuk=datetime.strptime(request.form["tanggal_masuk"],"%Y-%m-%d").date() if request.form.get("tanggal_masuk") else None
            item.tanggal_expired=datetime.strptime(request.form["tanggal_expired"],"%Y-%m-%d").date()
            item.lokasi=request.form.get("lokasi","-").strip() or "-"
            item.keterangan=request.form.get("keterangan","").strip()
            db.session.add(item); db.session.commit()
            flash("Data obat berhasil disimpan.","success")
            return redirect(url_for("daftar_obat"))
        except Exception as e:
            db.session.rollback()
            flash(str(e) if isinstance(e,ValueError) else "Periksa kembali data yang dimasukkan.","danger")
    return render_template("form_obat.html", item=item)

@app.post("/obat/<int:oid>/hapus")
def hapus_obat(oid):
    item=Obat.query.get_or_404(oid)
    db.session.delete(item); db.session.commit()
    flash("Data obat dihapus.","success")
    return redirect(url_for("daftar_obat"))

def filtered_items():
    q=request.args.get("q","").strip()
    status=request.args.get("status","")
    query=Obat.query
    if q: query=query.filter(or_(Obat.nama.ilike(f"%{q}%"),Obat.kode.ilike(f"%{q}%"),Obat.batch.ilike(f"%{q}%")))
    items=query.order_by(Obat.tanggal_expired.asc()).all()
    return [o for o in items if not status or status_obat(o.tanggal_expired)[0]==status]

@app.route("/export/excel")
def export_excel():
    items=filtered_items()
    wb=Workbook(); ws=wb.active; ws.title="Data Obat"
    headers=["Kode","Nama Obat","Kategori","Batch","Stok","Satuan","Tanggal Masuk","Tanggal Expired","Sisa Hari","Status","Lokasi","Keterangan"]
    ws.append(headers)
    for o in items:
        st,days=status_obat(o.tanggal_expired)
        ws.append([o.kode,o.nama,o.kategori,o.batch,o.stok,o.satuan,o.tanggal_masuk.isoformat() if o.tanggal_masuk else "",o.tanggal_expired.isoformat(),days,st,o.lokasi,o.keterangan])
    for c in ws[1]:
        c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="173B35")
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
    for col in ws.columns:
        letter=col[0].column_letter
        ws.column_dimensions[letter].width=min(max(max(len(str(c.value or "")) for c in col)+2,12),32)
    out=BytesIO(); wb.save(out); out.seek(0)
    return send_file(out,as_attachment=True,download_name="laporan_obat.xlsx",mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/export/pdf")
def export_pdf():
    items=filtered_items()
    out=BytesIO(); doc=SimpleDocTemplate(out,pagesize=landscape(A4),rightMargin=24,leftMargin=24,topMargin=24,bottomMargin=24)
    styles=getSampleStyleSheet()
    story=[Paragraph("LAPORAN MONITORING EXPIRED OBAT",styles["Title"]),Paragraph("Dicetak: "+date.today().strftime("%d-%m-%Y"),styles["Normal"]),Spacer(1,12)]
    data=[["Kode","Nama Obat","Batch","Stok","Expired","Sisa Hari","Status","Lokasi"]]
    for o in items:
        st,days=status_obat(o.tanggal_expired)
        data.append([o.kode,o.nama,o.batch,str(o.stok)+" "+o.satuan,o.tanggal_expired.strftime("%d-%m-%Y"),str(days),st,o.lokasi])
    table=Table(data,repeatRows=1,colWidths=[58,150,70,55,75,60,65,100])
    table.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#173B35")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("GRID",(0,0),(-1,-1),.4,colors.HexColor("#D5DED9")),("FONTSIZE",(0,0),(-1,-1),8),("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#F3F7F5")]),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("PADDING",(0,0),(-1,-1),6)]))
    story.append(table); doc.build(story); out.seek(0)
    return send_file(out,as_attachment=True,download_name="laporan_obat.pdf",mimetype="application/pdf")

@app.route("/tentang")
def tentang(): return render_template("tentang.html")

with app.app_context():
    db.create_all()

if __name__=="__main__":
    app.run(debug=True)
