from flask import Flask, render_template, request, redirect, url_for, flash, send_file, session
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime
from openpyxl import Workbook
import io, os

os.makedirs("/var/data", exist_ok=True)

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "acar-patron-2025-gizli-anahtar")
# Yeni Patron Sistemi Veritabanı (Eski verileri etkilemez)
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:////var/data/acar_patron.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db = SQLAlchemy(app)

# ============ VERİTABANI MODELLERİ ============
class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    full_name = db.Column(db.String(120), default="")

class Kasa(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    tarih = db.Column(db.DateTime, default=datetime.now)
    tip = db.Column(db.String(30), nullable=False)
    tutar = db.Column(db.Float, nullable=False)
    aciklama = db.Column(db.String(255), default="")

class Gider(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    tarih = db.Column(db.DateTime, default=datetime.now)
    aciklama = db.Column(db.String(255), nullable=False)
    tutar = db.Column(db.Float, nullable=False)

class Borc(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    isim = db.Column(db.String(120), nullable=False)
    aciklama = db.Column(db.String(255), default="")
    toplam_borc = db.Column(db.Float, nullable=False)
    odenen = db.Column(db.Float, default=0)
    tarih = db.Column(db.DateTime, default=datetime.now)

    @property
    def kalan(self):
        return round(self.toplam_borc - self.odenen, 2)

class Personel(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    ad_soyad = db.Column(db.String(120), nullable=False)
    gorev = db.Column(db.String(50), default="")
    maas = db.Column(db.Float, default=0)
    ise_giris = db.Column(db.DateTime)
    maas_gunu = db.Column(db.Integer, default=1)
    avanslar = db.relationship('Avans', backref='personel', lazy=True)
    izinler = db.relationship('Izin', backref='personel', lazy=True)

class Avans(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    personel_id = db.Column(db.Integer, db.ForeignKey('personel.id'), nullable=False)
    tutar = db.Column(db.Float, nullable=False)
    tarih = db.Column(db.DateTime, default=datetime.now)
    aciklama = db.Column(db.String(255), default="")

class Izin(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    personel_id = db.Column(db.Integer, db.ForeignKey('personel.id'), nullable=False)
    gun_sayisi = db.Column(db.Float, nullable=False)
    tarih = db.Column(db.DateTime, default=datetime.now)
    aciklama = db.Column(db.String(255), default="")

class Log(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    tarih = db.Column(db.DateTime, default=datetime.now)
    islem = db.Column(db.String(255))
    detay = db.Column(db.Text, default="")

# ============ YARDIMCILAR ============
def giris_yapmis(): return "user_id" in session
def giris_gerek(f):
    from functools import wraps
    @wraps(f)
    def wrapper(*a, **kw):
        if not giris_yapmis(): return redirect(url_for("login"))
        return f(*a, **kw)
    return wrapper
def log_yaz(islem, detay=""):
    db.session.add(Log(islem=islem, detay=detay))
    db.session.commit()

# ============ ROTALAR ============
@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        u = User.query.filter_by(username=request.form["username"]).first()
        if u and check_password_hash(u.password_hash, request.form["password"]):
            session["user_id"] = u.id
            log_yaz("GİRİŞ", u.username)
            return redirect(url_for("kasa"))
        flash("Kullanıcı adı veya şifre yanlış!", "danger")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/")
@giris_gerek
def index(): return redirect(url_for("kasa"))

# ============ GÜNLÜK KASA ============
@app.route("/kasa", methods=["GET", "POST"])
@giris_gerek
def kasa():
    secili_tarih = request.args.get("tarih", datetime.now().strftime("%Y-%m-%d"))
    
    if request.method == "POST":
        try: tutar = float(request.form["tutar"].replace(",", "."))
        except: flash("Geçerli tutar girin!", "danger"); return redirect(url_for("kasa"))
        
        islem_tarihi = datetime.now()
        if request.form.get("tarih"):
            try: islem_tarihi = datetime.strptime(request.form["tarih"], "%Y-%m-%d")
            except: pass
            
        db.session.add(Kasa(tarih=islem_tarihi, tip=request.form["tip"], tutar=tutar, aciklama=request.form.get("aciklama", "")))
        db.session.commit()
        log_yaz("KASA EKLE", f"{tutar} TL - {request.form['tip']}")
        flash("Kasa kaydı başarıyla eklendi.", "success")
        return redirect(url_for("kasa", tarih=secili_tarih))

    kayitlar = Kasa.query.filter(Kasa.tarih.like(f"{secili_tarih}%")).order_by(Kasa.tarih.desc()).all()
    
    # Günlük Toplamlar
    toplamlar = {"Nakit":0, "Kredi Kartı":0, "Havale":0, "Euro":0, "Dolar":0, "ETS":0, "Tatilbudur":0, "OtelFiyat":0, "Jolly Tur":0}
    for k in kayitlar:
        if k.tip in toplamlar: toplamlar[k.tip] += k.tutar

    return render_template("kasa.html", kayitlar=kayitlar, secili_tarih=secili_tarih, toplamlar=toplamlar)

@app.route("/kasa_sil/<int:id>")
@giris_gerek
def kasa_sil(id):
    k = Kasa.query.get_or_404(id); tarih = k.tarih.strftime("%Y-%m-%d")
    db.session.delete(k); db.session.commit()
    log_yaz("KASA SİL", f"{k.tutar} TL silindi.")
    flash("Kasa kaydı silindi.", "success")
    return redirect(url_for("kasa", tarih=tarih))

# ============ GİDERLER ============
@app.route("/giderler", methods=["GET", "POST"])
@giris_gerek
def giderler():
    secili_tarih = request.args.get("tarih", datetime.now().strftime("%Y-%m-%d"))
    
    if request.method == "POST":
        try: tutar = float(request.form["tutar"].replace(",", "."))
        except: flash("Geçerli tutar girin!", "danger"); return redirect(url_for("giderler"))
            
        islem_tarihi = datetime.now()
        if request.form.get("tarih"):
            try: islem_tarihi = datetime.strptime(request.form["tarih"], "%Y-%m-%d")
            except: pass
            
        db.session.add(Gider(tarih=islem_tarihi, aciklama=request.form["aciklama"], tutar=tutar))
        db.session.commit()
        log_yaz("GİDER EKLE", f"{request.form['aciklama']} - {tutar} TL")
        flash("Gider başarıyla eklendi.", "success")
        return redirect(url_for("giderler", tarih=secili_tarih))

    gider_list = Gider.query.filter(Gider.tarih.like(f"{secili_tarih}%")).order_by(Gider.tarih.desc()).all()
    toplam_gider = sum(g.tutar for g in gider_list)
    return render_template("giderler.html", giderler=gider_list, secili_tarih=secili_tarih, toplam_gider=toplam_gider)

@app.route("/gider_sil/<int:id>")
@giris_gerek
def gider_sil(id):
    g = Gider.query.get_or_404(id); tarih = g.tarih.strftime("%Y-%m-%d")
    db.session.delete(g); db.session.commit()
    log_yaz("GİDER SİL", f"ID {id} silindi.")
    flash("Gider silindi.", "success")
    return redirect(url_for("giderler", tarih=tarih))

# ============ BORÇLULAR (MANUEL) ============
@app.route("/borclular", methods=["GET", "POST"])
@giris_gerek
def borclular():
    if request.method == "POST":
        try: tutar = float(request.form["toplam_borc"].replace(",", "."))
        except: flash("Geçerli tutar girin!", "danger"); return redirect(url_for("borclular"))
            
        db.session.add(Borc(isim=request.form["isim"], aciklama=request.form.get("aciklama", ""), toplam_borc=tutar))
        db.session.commit()
        log_yaz("BORÇ EKLE", f"{request.form['isim']} - {tutar} TL")
        flash("Borç kaydı oluşturuldu.", "success")
        return redirect(url_for("borclular"))

    borc_list = Borc.query.filter(Borc.odenen < Borc.toplam_borc).all()
    return render_template("borclular.html", borclular=borc_list)

@app.route("/borc_ode/<int:id>", methods=["POST"])
@giris_gerek
def borc_ode(id):
    b = Borc.query.get_or_404(id)
    try: tutar = float(request.form["tutar"].replace(",", "."))
    except: flash("Geçerli tutar girin!", "danger"); return redirect(url_for("borclular"))
    if tutar > b.kalan: flash("Ödeme kalan borçtan büyük olamaz!", "danger"); return redirect(url_for("borclular"))
    
    b.odenen += tutar
    db.session.commit()
    log_yaz("BORÇ ÖDEME", f"{b.isim} - {tutar} TL ödedi.")
    flash("Ödeme alındı.", "success")
    return redirect(url_for("borclular"))

@app.route("/borc_sil/<int:id>")
@giris_gerek
def borc_sil(id):
    b = Borc.query.get_or_404(id)
    db.session.delete(b); db.session.commit()
    flash("Borç tamamen silindi.", "success")
    return redirect(url_for("borclular"))

# ============ PERSONEL, MAAŞ VE İZİN ============
@app.route("/personel", methods=["GET", "POST"])
@giris_gerek
def personel():
    if request.method == "POST":
        try: maas = float(request.form["maas"].replace(",", "."))
        except: maas = 0
        try: ise_giris = datetime.strptime(request.form["ise_giris"], "%Y-%m-%d")
        except: ise_giris = datetime.now()
        try: maas_gunu = int(request.form["maas_gunu"])
        except: maas_gunu = 1
            
        db.session.add(Personel(ad_soyad=request.form["ad_soyad"], gorev=request.form["gorev"], maas=maas, ise_giris=ise_giris, maas_gunu=maas_gunu))
        db.session.commit()
        flash("Personel eklendi.", "success")
        return redirect(url_for("personel"))
    return render_template("personel.html", personeller=Personel.query.all())

@app.route("/avans_ver/<int:pid>", methods=["POST"])
@giris_gerek
def avans_ver(pid):
    p = Personel.query.get_or_404(pid)
    try: tutar = float(request.form["tutar"].replace(",", "."))
    except: flash("Geçerli tutar girin!", "danger"); return redirect(url_for("personel"))
    db.session.add(Avans(personel_id=p.id, tutar=tutar, aciklama=request.form.get("aciklama", ""))); db.session.commit()
    flash("Avans verildi.", "success")
    return redirect(url_for("personel"))

@app.route("/izin_ekle/<int:pid>", methods=["POST"])
@giris_gerek
def izin_ekle(pid):
    p = Personel.query.get_or_404(pid)
    try: gun = float(request.form["gun_sayisi"].replace(",", "."))
    except: flash("Geçerli gün sayısı girin!", "danger"); return redirect(url_for("personel"))
    db.session.add(Izin(personel_id=p.id, gun_sayisi=gun, aciklama=request.form.get("aciklama", ""))); db.session.commit()
    flash(f"{gun} gün izin eklendi.", "success")
    return redirect(url_for("personel"))

@app.route("/maas_ode/<int:pid>")
@giris_gerek
def maas_ode(pid):
    p = Personel.query.get_or_404(pid)
    kalan = p.maas - sum(a.tutar for a in p.avanslar)
    if kalan > 0:
        db.session.add(Gider(aciklama=f"{p.ad_soyad} Maaş Ödemesi", tutar=kalan))
        for a in p.avanslar: db.session.delete(a)
        db.session.commit(); flash(f"{kalan} TL maaş ödendi ve gidere eklendi.", "success")
    else: flash("Ödenecek maaş kalmamış.", "danger")
    return redirect(url_for("personel"))

@app.route("/personel_sil/<int:pid>")
@giris_gerek
def personel_sil(pid):
    p = Personel.query.get_or_404(pid); Avans.query.filter_by(personel_id=p.id).delete(); Izin.query.filter_by(personel_id=p.id).delete(); db.session.delete(p); db.session.commit(); return redirect(url_for("personel"))

# ============ BAŞLANGIÇ ============
def init_db():
    with app.app_context():
        db.create_all()
        if not User.query.filter_by(username="admin").first():
            db.session.add(User(username="admin", password_hash=generate_password_hash("1234"), full_name="Patron"))
            db.session.commit(); print("✅ Patron sistemi kuruldu.")

init_db()
if __name__ == "__main__": app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
