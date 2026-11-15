from flask import Flask, render_template, request, redirect, url_for, flash, send_file, session
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime
from openpyxl import Workbook
import io, os

os.makedirs("/var/data", exist_ok=True)

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "acar-patron-2025-gizli-anahtar")
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:////var/data/acar_patron.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db = SQLAlchemy(app)

# ============ VERİTABANI MODELLERİ ============
class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    full_name = db.Column(db.String(120), default="")
    role = db.Column(db.String(20), default="personel") 

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
    avanslar = db.relationship('Avans', backref='personel', lazy=True, cascade="all, delete-orphan")
    izinler = db.relationship('Izin', backref='personel', lazy=True, cascade="all, delete-orphan")

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

class AylikKur(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    ay = db.Column(db.String(7), unique=True, nullable=False) 
    usd = db.Column(db.Float, default=34.0)
    eur = db.Column(db.Float, default=37.0)
    gbp = db.Column(db.Float, default=43.0) 

class Log(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    tarih = db.Column(db.DateTime, default=datetime.now)
    islem = db.Column(db.String(255))
    detay = db.Column(db.Text, default="")

# ============ YARDIMCILAR ============
def giris_yapmis(): return "user_id" in session
def admin_mi(): return session.get("role") == "admin"
def giris_gerek(f):
    from functools import wraps
    @wraps(f)
    def wrapper(*a, **kw):
        if not giris_yapmis(): return redirect(url_for("login"))
        return f(*a, **kw)
    return wrapper

def islem_yetkisi_gerek(f):
    from functools import wraps
    @wraps(f)
    def wrapper(*a, **kw):
        if not giris_yapmis(): return redirect(url_for("login"))
        if session.get("role") == "ortak":
            flash("Ortak hesapları sadece görüntüleme yapabilir, işlem yapamaz!", "danger")
            return redirect(url_for("kasa"))
        return f(*a, **kw)
    return wrapper

def log_yaz(islem, detay=""):
    db.session.add(Log(islem=islem, detay=detay))
    db.session.commit()

@app.context_processor
def inject_user():
    return dict(current_user=User.query.get(session.get("user_id")) if giris_yapmis() else None, 
                admin_mi=admin_mi(), ortak_mi=(session.get("role") == "ortak"))

# ============ ROTALAR ============
@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        u = User.query.filter_by(username=request.form["username"]).first()
        if u and check_password_hash(u.password_hash, request.form["password"]):
            session["user_id"] = u.id
            session["role"] = u.role
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
    secili_ay = secili_tarih[:7] # Seçili tarihin ait olduğu ay (Örn: 2026-11)
    
    if request.method == "POST":
        if session.get("role") == "ortak": return redirect(url_for("kasa"))
        
        form_tipi = request.form.get("form_tipi")
        islem_tarihi = datetime.now()
        if request.form.get("tarih"):
            try: islem_tarihi = datetime.strptime(request.form["tarih"], "%Y-%m-%d")
            except: pass
            
        if form_tipi == "gelir":
            try: tutar = float(request.form["tutar"].replace(",", "."))
            except: flash("Geçerli tutar girin!", "danger"); return redirect(url_for("kasa"))
            db.session.add(Kasa(tarih=islem_tarihi, tip=request.form["tip"], tutar=tutar, aciklama=request.form.get("aciklama", "")))
            db.session.commit()
            log_yaz("KASA EKLE", f"{tutar} {request.form['tip']}")
            flash("Gelir kaydı başarıyla eklendi.", "success")
            
        elif form_tipi == "gider":
            try: tutar = float(request.form["tutar"].replace(",", "."))
            except: flash("Geçerli tutar girin!", "danger"); return redirect(url_for("kasa"))
            db.session.add(Gider(tarih=islem_tarihi, aciklama=request.form["aciklama"], tutar=tutar))
            db.session.commit()
            log_yaz("GİDER EKLE", f"{request.form['aciklama']} - {tutar} TL")
            flash("Gider kaydı başarıyla eklendi.", "success")
            
        return redirect(url_for("kasa", tarih=secili_tarih))

    kayitlar = Kasa.query.filter(Kasa.tarih.like(f"{secili_tarih}%")).order_by(Kasa.tarih.desc()).all()
    giderler = Gider.query.filter(Gider.tarih.like(f"{secili_tarih}%")).order_by(Gider.tarih.desc()).all()
    
    toplamlar = {"Nakit":0, "Kredi Kartı":0, "Havale":0, "Euro":0, "Dolar":0, "Pound":0, "ETS":0, "Tatilbudur":0, "OtelFiyat":0, "Jolly Tur":0}
    for k in kayitlar:
        if k.tip in toplamlar: toplamlar[k.tip] += k.tutar

    # DÖVİZ KURLARINI ÇEK (Seçili ayın kurları)
    kur = AylikKur.query.filter_by(ay=secili_ay).first()
    usd_kur = kur.usd if kur else 34.0
    eur_kur = kur.eur if kur else 37.0
    gbp_kur = kur.gbp if kur else 43.0

    # TOPLAM GELİR HESAPLAMA (TL + Döviz Karşılıkları)
    tl_gelir = sum(v for k, v in toplamlar.items() if k not in ["Euro", "Dolar", "Pound"])
    dolar_tl = toplamlar["Dolar"] * usd_kur
    euro_tl = toplamlar["Euro"] * eur_kur
    pound_tl = toplamlar["Pound"] * gbp_kur

    toplam_gelir = tl_gelir + dolar_tl + euro_tl + pound_tl
    toplam_gider = sum(g.tutar for g in giderler)
    kalan = toplam_gelir - toplam_gider

    return render_template("kasa.html", kayitlar=kayitlar, giderler=giderler, secili_tarih=secili_tarih, 
                           toplamlar=toplamlar, toplam_gelir=toplam_gelir, toplam_gider=toplam_gider, kalan=kalan,
                           usd_kur=usd_kur, eur_kur=eur_kur, gbp_kur=gbp_kur, 
                           dolar_tl=dolar_tl, euro_tl=euro_tl, pound_tl=pound_tl)

@app.route("/kasa_sil/<int:id>")
@islem_yetkisi_gerek
def kasa_sil(id):
    k = Kasa.query.get_or_404(id); tarih = k.tarih.strftime("%Y-%m-%d")
    db.session.delete(k); db.session.commit()
    log_yaz("KASA SİL", f"{k.tutar} silindi.")
    flash("Kasa kaydı silindi.", "success")
    return redirect(url_for("kasa", tarih=tarih))

@app.route("/gider_sil/<int:id>")
@islem_yetkisi_gerek
def gider_sil(id):
    g = Gider.query.get_or_404(id); tarih = g.tarih.strftime("%Y-%m-%d")
    db.session.delete(g); db.session.commit()
    log_yaz("GİDER SİL", f"ID {id} silindi.")
    flash("Gider silindi.", "success")
    return redirect(url_for("kasa", tarih=tarih))

# ============ AYLIK KASA ============
@app.route("/aylik_kasa", methods=["GET", "POST"])
@giris_gerek
def aylik_kasa():
    secili_ay = request.values.get("ay", datetime.now().strftime("%Y-%m"))
    
    if request.method == "POST":
        if session.get("role") == "ortak": return redirect(url_for("aylik_kasa"))
        try:
            usd = float(request.form["usd"].replace(",", "."))
            eur = float(request.form["eur"].replace(",", "."))
            gbp = float(request.form["gbp"].replace(",", "."))
            kur = AylikKur.query.filter_by(ay=secili_ay).first()
            if kur:
                kur.usd = usd; kur.eur = eur; kur.gbp = gbp
            else:
                db.session.add(AylikKur(ay=secili_ay, usd=usd, eur=eur, gbp=gbp))
            db.session.commit()
            flash("Döviz kurları başarıyla güncellendi.", "success")
        except:
            flash("Geçersiz kur değeri girdiniz!", "danger")
        return redirect(url_for("aylik_kasa", ay=secili_ay))

    kur = AylikKur.query.filter_by(ay=secili_ay).first()
    usd_kur = kur.usd if kur else 34.0
    eur_kur = kur.eur if kur else 37.0
    gbp_kur = kur.gbp if kur else 43.0

    kayitlar = Kasa.query.filter(Kasa.tarih.like(f"{secili_ay}%")).all()
    toplamlar = {"Nakit":0, "Kredi Kartı":0, "Havale":0, "Euro":0, "Dolar":0, "Pound":0, "ETS":0, "Tatilbudur":0, "OtelFiyat":0, "Jolly Tur":0}
    for k in kayitlar:
        if k.tip in toplamlar: toplamlar[k.tip] += k.tutar
        
    tl_gelir = sum(v for k, v in toplamlar.items() if k not in ["Euro", "Dolar", "Pound"])
    dolar_tl = toplamlar["Dolar"] * usd_kur
    euro_tl = toplamlar["Euro"] * eur_kur
    pound_tl = toplamlar["Pound"] * gbp_kur
    
    toplam_gelir_tl = tl_gelir + dolar_tl + euro_tl + pound_tl

    giderler = Gider.query.filter(Gider.tarih.like(f"{secili_ay}%")).all()
    toplam_gider = sum(g.tutar for g in giderler)
    kalan = toplam_gelir_tl - toplam_gider

    return render_template("aylik_kasa.html", secili_ay=secili_ay, toplamlar=toplamlar, 
                           toplam_gider=toplam_gider, usd_kur=usd_kur, eur_kur=eur_kur, gbp_kur=gbp_kur,
                           dolar_tl=dolar_tl, euro_tl=euro_tl, pound_tl=pound_tl, toplam_gelir_tl=toplam_gelir_tl, kalan=kalan)

# ============ BORÇLULAR ============
@app.route("/borclular", methods=["GET", "POST"])
@giris_gerek
def borclular():
    if request.method == "POST":
        if session.get("role") == "ortak": return redirect(url_for("borclular"))
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
@islem_yetkisi_gerek
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
@islem_yetkisi_gerek
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
        if not admin_mi(): return redirect(url_for("personel"))
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

@app.route("/personel_duzenle/<int:pid>", methods=["POST"])
@islem_yetkisi_gerek
def personel_duzenle(pid):
    p = Personel.query.get_or_404(pid)
    try: p.maas = float(request.form["maas"].replace(",", "."))
    except: pass
    p.ad_soyad = request.form["ad_soyad"]
    p.gorev = request.form["gorev"]
    try: p.maas_gunu = int(request.form["maas_gunu"])
    except: pass
    try: p.ise_giris = datetime.strptime(request.form["ise_giris"], "%Y-%m-%d")
    except: pass
    
    db.session.commit()
    log_yaz("PERSONEL DÜZENLE", p.ad_soyad)
    flash("Personel bilgileri başarıyla güncellendi.", "success")
    return redirect(url_for("personel"))

@app.route("/avans_ver/<int:pid>", methods=["POST"])
@islem_yetkisi_gerek
def avans_ver(pid):
    p = Personel.query.get_or_404(pid)
    try: tutar = float(request.form["tutar"].replace(",", "."))
    except: flash("Geçerli tutar girin!", "danger"); return redirect(url_for("personel"))
    db.session.add(Avans(personel_id=p.id, tutar=tutar, aciklama=request.form.get("aciklama", ""))); db.session.commit()
    flash("Avans verildi.", "success")
    return redirect(url_for("personel"))

@app.route("/izin_ekle/<int:pid>", methods=["POST"])
@islem_yetkisi_gerek
def izin_ekle(pid):
    p = Personel.query.get_or_404(pid)
    try: gun = float(request.form["gun_sayisi"].replace(",", "."))
    except: flash("Geçerli gün sayısı girin!", "danger"); return redirect(url_for("personel"))
        
    izin_tarihi = datetime.now()
    if request.form.get("tarih"):
        try: izin_tarihi = datetime.strptime(request.form["tarih"], "%Y-%m-%d")
        except: pass
        
    db.session.add(Izin(personel_id=p.id, gun_sayisi=gun, tarih=izin_tarihi, aciklama=request.form.get("aciklama", ""))); db.session.commit()
    flash(f"{gun} gün izin eklendi.", "success")
    return redirect(url_for("personel"))

@app.route("/izin_sil/<int:id>")
@islem_yetkisi_gerek
def izin_sil(id):
    i = Izin.query.get_or_404(id)
    db.session.delete(i); db.session.commit()
    flash("İzin kaydı silindi.", "success")
    return redirect(url_for("personel"))

@app.route("/izin_duzenle/<int:id>", methods=["POST"])
@islem_yetkisi_gerek
def izin_duzenle(id):
    i = Izin.query.get_or_404(id)
    try: i.gun_sayisi = float(request.form["gun_sayisi"].replace(",", "."))
    except: pass
    i.aciklama = request.form.get("aciklama", "")
    if request.form.get("tarih"):
        try: i.tarih = datetime.strptime(request.form["tarih"], "%Y-%m-%d")
        except: pass
    db.session.commit()
    flash("İzin kaydı güncellendi.", "success")
    return redirect(url_for("personel"))

@app.route("/maas_ode/<int:pid>")
@islem_yetkisi_gerek
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
@islem_yetkisi_gerek
def personel_sil(pid):
    p = Personel.query.get_or_404(pid)
    db.session.delete(p); db.session.commit()
    flash("Personel ve tüm geçmişi silindi.", "success")
    return redirect(url_for("personel"))

# ============ KULLANICI (ORTAK VE PERSONEL) EKLEME ============
@app.route("/kullanicilar", methods=["GET", "POST"])
@giris_gerek
def kullanicilar():
    if not admin_mi():
        flash("Bu sayfaya sadece Patron erişebilir!", "danger")
        return redirect(url_for("kasa"))

    if request.method == "POST":
        kullanici_adi = request.form["username"]
        sifre = request.form["password"]
        ad_soyad = request.form["full_name"]
        rol = request.form.get("role", "personel")
        
        if User.query.filter_by(username=kullanici_adi).first():
            flash("Bu kullanıcı adı zaten alınmış!", "danger")
        else:
            yeni_user = User(username=kullanici_adi, password_hash=generate_password_hash(sifre), full_name=ad_soyad, role=rol)
            db.session.add(yeni_user)
            db.session.commit()
            log_yaz("KULLANICI EKLE", f"{kullanici_adi} - Rol: {rol}")
            flash(f"✅ {ad_soyad} adlı kullanıcı sisteme eklendi.", "success")
        return redirect(url_for("kullanicilar"))
        
    return render_template("kullanicilar.html", kullanicilar=User.query.all())

@app.route("/kullanici_sil/<int:id>")
@giris_gerek
def kullanici_sil(id):
    if not admin_mi(): return redirect(url_for("kasa"))
    if id == session.get("user_id"):
        flash("Kendi hesabınızı silemezsiniz!", "danger")
        return redirect(url_for("kullanicilar"))
    u = User.query.get_or_404(id)
    db.session.delete(u); db.session.commit()
    flash("Kullanıcı silindi.", "success")
    return redirect(url_for("kullanicilar"))

@app.route("/sifre_degistir", methods=["GET", "POST"])
@giris_gerek
def sifre_degistir():
    if request.method == "POST":
        eski = request.form["eski_sifre"]
        yeni = request.form["yeni_sifre"]
        tekrar = request.form["yeni_sifre_tekrar"]
        
        u = User.query.get(session["user_id"])
        if not check_password_hash(u.password_hash, eski):
            flash("Mevcut şifreniz yanlış!", "danger")
        elif yeni != tekrar:
            flash("Yeni şifreler birbiriyle uyuşmuyor!", "danger")
        elif len(yeni) < 4:
            flash("Yeni şifre en az 4 karakter olmalıdır!", "danger")
        else:
            u.password_hash = generate_password_hash(yeni)
            db.session.commit()
            flash("✅ Şifreniz başarıyla güncellendi!", "success")
            return redirect(url_for("kasa"))
    return render_template("sifre_degistir.html")

# ============ BAŞLANGIÇ ============
def init_db():
    with app.app_context():
        db.create_all()
        
        from sqlalchemy import text
        try:
            db.session.execute(text("ALTER TABLE aylik_kur ADD COLUMN gbp FLOAT DEFAULT 43.0"))
            db.session.commit()
        except: db.session.rollback()

        try:
            db.session.execute(text("ALTER TABLE user ADD COLUMN role VARCHAR(20) DEFAULT 'admin'"))
            db.session.commit()
        except: db.session.rollback()

        admin_user = User.query.filter_by(username="admin").first()
        if admin_user:
            admin_user.role = "admin"
            db.session.commit()
        else:
            db.session.add(User(username="admin", password_hash=generate_password_hash("1234"), full_name="Baş Patron", role="admin"))
            db.session.commit()
            
        print("✅ Patron sistemi kuruldu ve yetkiler doğrulandı.")

init_db()
if __name__ == "__main__": app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
