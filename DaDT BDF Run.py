#!/usr/bin/env python3
"""
DaDT Structure Definition Tool v16.0
=====================================
Tab 1: BDF Merge Preparation
Tab 2: Understanding Structure Type (apply 'Bar Property'/'Skin Property' + offsets, solve base model)
Tab 3: Post Processing (OP2 -> H5 batch conversion, Load Extraction folder prep)
"""

# ==================== KÜTÜPHANE/MODÜL İMPORTLARI ====================
# Tkinter: Python'un standart GUI (grafik arayüz) kütüphanesi. Bu araç
# tamamen Tkinter tabanlı masaüstü bir uygulamadır.
import tkinter as tk
# ttk: Tkinter'ın "themed" (temalı, daha modern görünümlü) widget seti.
# filedialog: dosya/klasör seçme pencereleri (Browse butonları için).
# messagebox: hata/uyarı/bilgi pop-up pencereleri (messagebox.showerror vb.).
# scrolledtext: otomatik kaydırma çubuklu metin kutusu (Log panelleri için).
from tkinter import ttk, filedialog, messagebox, scrolledtext
# os: dosya/klasör yolu işlemleri (os.path.join, os.makedirs, os.listdir, os.walk vb.).
import os
# re: düzenli ifadeler (regex) - bazı metin ayrıştırma işlemlerinde kullanılıyor.
import re
# threading: arka planda (GUI'yi dondurmadan) uzun süren işleri çalıştırmak için
# ayrı thread (iş parçacığı) açmaya yarar. Her "RUN" butonu işi bir thread'e atar.
import threading
# csv: CSV (virgülle ayrılmış değerler) dosyalarını okuma/yazma.
import csv
# shutil: dosya/klasör kopyalama (copy2), taşıma (move), klasör ağacı kopyalama (copytree).
import shutil
# tempfile: geçici dosya/klasör oluşturma (bu dosyada şu an aktif kullanılmıyor,
# ileride geçici dosya ihtiyacı olursa hazır dursun diye import edilmiş olabilir).
import tempfile
# subprocess: harici programları (Nastran, dönüştürücü exe'ler, .bat dosyaları vb.)
# başlatıp çalıştırmak için. subprocess.run (bekleyerek) ve subprocess.Popen
# (beklemeden, arka planda başlatmak için) ikisi de kullanılıyor.
import subprocess
# time: zaman ölçümü (time.time() ile süre hesaplama) ve bekleme (time.sleep()).
import time
# concurrent.futures: birden fazla işi (örn. birden fazla Nastran çalıştırması,
# birden fazla exe çalıştırması) aynı anda, sınırlı sayıda paralel thread ile
# çalıştırmak için (ThreadPoolExecutor).
import concurrent.futures
# datetime: tarih/saat damgası üretmek için (örn. "BaseModel_20260101_120000" gibi
# klasör isimlerinde kullanılan zaman damgası).
from datetime import datetime
# pandas: Excel/CSV dosyalarını tablo (DataFrame) olarak okuyup işlemek için.
import pandas as pd
# pyNastran: NASTRAN'ın BDF (Bulk Data File) formatındaki dosyalarını ayrıştırmak
# (parse etmek) için kullanılan üçüncü parti kütüphane. Sadece BDF sınıfı
# import ediliyor (tüm pyNastran kütüphanesi değil).
from pyNastran.bdf.bdf import BDF
# numpy: sayısal/vektör hesaplamalar için (örn. offset hesaplarında normal vektör
# işlemleri - np.cross, np.linalg.norm vb.).
import numpy as np
# openpyxl: Excel (.xlsx/.xlsm) dosyalarını hücre hücre okuyup YAZMAK için.
# pandas sadece okuma/toplu analiz için kullanılırken, var olan bir Excel şablonunu
# (örneğin makrolu bir .xlsm dosyasını) KORUYARAK belirli hücrelerini güncellemek
# gerektiğinde openpyxl kullanılıyor.
import openpyxl

# ==================== GÖRSELLEŞTİRME (VISUALIZATION) BAĞIMLILIKLARI ====================
# matplotlib: SADECE "Visualization Results" sekmesindeki 3B FEM mesh görünümü
# ve 2B grafikler için kullanılıyor. Uygulamanın geri kalanı (diğer 7 sekme)
# matplotlib'e hiç ihtiyaç duymuyor - bu yüzden import opsiyonel: kurulu
# değilse uygulamanın tamamı yine de normal çalışır, sadece Visualization
# Results sekmesi kullanıcıya "matplotlib kurulu değil" uyarısı gösterir.
try:
    import matplotlib
    matplotlib.use('TkAgg')  # Tkinter penceresi içine gömülebilen çizim motoru
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
    from mpl_toolkits.mplot3d.art3d import Line3DCollection, Poly3DCollection
    import matplotlib.cm as mpl_cm
    import matplotlib.colors as mpl_colors
    import matplotlib.ticker as mpl_ticker
    from matplotlib.patches import Patch
    _MATPLOTLIB_AVAILABLE = True
except Exception:
    _MATPLOTLIB_AVAILABLE = False

# _APP_COLORS: uygulamanın renk paletini tutan global (modül seviyesi) sözlük.
# Başlangıçta boş, _configure_app_style() çağrıldığında içi dolduruluyor.
# Tüm sekmeler (tab) bu sözlüğü widget renklerini ayarlarken okuyor
# (örn. _APP_COLORS['surface'], _APP_COLORS['primary'] gibi).
_APP_COLORS = {}  # populated by _configure_app_style(); shared palette for both tabs


def _configure_app_style(root):
    """One-time visual theme for the whole app. Both tabs share this root's
    ttk.Style singleton, so configuring it once here restyles every ttk
    widget in both tabs without touching how any of them behave. Pure
    look-and-feel - no widget's command/callback is affected."""
    # Uygulamanın tüm renk paleti burada tanımlanıyor. Bu sadece GÖRSEL bir
    # ayardır - hiçbir butonun/alanın İŞLEVİNİ (ne yaptığını) etkilemez.
    colors = {
        'bg': '#f1f5f9',           # genel arka plan rengi (açık gri-mavi)
        'surface': '#ffffff',      # kutu/alan (entry, listbox vb.) arka planı (beyaz)
        'text': '#0f172a',         # normal yazı rengi (koyu lacivert/siyah)
        'muted': '#64748b',        # soluk/ikincil yazı rengi (açıklama metinleri için)
        'border': '#cbd5e1',       # kenarlık rengi (açık gri)
        'primary': '#2563eb',      # ana vurgu rengi (mavi) - örn. RUN butonu
        'primary_dark': '#1d4ed8', # ana vurgu renginin koyu tonu (hover/başlık rengi)
        'success': '#16a34a',      # başarı rengi (yeşil)
        'success_dark': '#15803d', # başarı renginin koyu tonu
        'danger': '#dc2626',       # hata/tehlike rengi (kırmızı)
        'danger_dark': '#b91c1c',  # hata renginin koyu tonu
        'log_bg': '#0f172a',       # Log panelinin arka planı (koyu - terminal gibi görünsün diye)
        'log_fg': '#e2e8f0',       # Log panelindeki yazı rengi (açık renk, koyu zemin üzerinde okunsun diye)
    }

    # Ana pencerenin (root) arka plan rengini ayarla.
    root.configure(bg=colors['bg'])

    # ttk.Style nesnesi: Tkinter'ın temalı widget'larının (ttk.Button, ttk.Label
    # vb.) görünümünü merkezi olarak kontrol eden nesne.
    style = ttk.Style(root)
    try:
        # 'clam' teması: platformdan bağımsız, daha modern/düz görünümlü bir tema.
        # Bazı sistemlerde bu tema bulunmayabilir, o yüzden try/except içinde.
        style.theme_use('clam')
    except tk.TclError:
        # Tema bulunamazsa sessizce devam et - varsayılan tema kullanılır,
        # uygulama yine de çalışır, sadece görünüm biraz farklı olur.
        pass

    # Normal yazı tipi (etiketler, butonlar vb. için).
    base_font = ('Segoe UI', 10)
    # Bölüm başlıkları için kalın yazı tipi (örn. "1. Input Files" gibi başlıklar).
    section_font = ('Segoe UI', 10, 'bold')

    # Aşağıdaki style.configure() çağrıları, her bir ttk widget TÜRÜ için
    # (ör. tüm TFrame'ler, tüm TLabel'lar) ortak varsayılan görünüm ayarlıyor.
    # '.' = tüm widget'lar için en genel/kök ayar.
    style.configure('.', background=colors['bg'], foreground=colors['text'], font=base_font)
    style.configure('TFrame', background=colors['bg'])
    style.configure('TLabel', background=colors['bg'], foreground=colors['text'])
    style.configure('TCheckbutton', background=colors['bg'])

    # 'Header.TLabel' stili: sekme başlıklarında kullanılan büyük, kalın,
    # vurgulu başlık etiketi (örn. "Post Processing", "Joint Run" gibi başlıklar).
    style.configure('Header.TLabel', font=('Segoe UI', 15, 'bold'),
                    foreground=colors['primary_dark'], background=colors['bg'])
    # 'Subtitle.TLabel' stili: başlığın altındaki küçük açıklama satırı.
    style.configure('Subtitle.TLabel', font=('Segoe UI', 9),
                    foreground=colors['muted'], background=colors['bg'])

    # LabelFrame (örn. "1. Input Files" çerçeveli kutular) ve onun başlık yazısı.
    style.configure('TLabelframe', background=colors['bg'], bordercolor=colors['border'])
    style.configure('TLabelframe.Label', background=colors['bg'],
                    foreground=colors['primary_dark'], font=section_font)

    # Notebook = sekme (tab) çubuğunun kendisi.
    style.configure('TNotebook', background=colors['bg'], borderwidth=0)
    # Her bir sekme sekmesi (örn. "Post Processing" yazan tıklanabilir kulakçık).
    style.configure('TNotebook.Tab', font=('Segoe UI', 10, 'bold'), padding=(16, 8),
                    background=colors['border'], foreground=colors['text'])
    # style.map(): widget'ın DURUMUNA göre (örn. seçili/seçili değil) farklı
    # renk uygulamak için kullanılır. Burada: seçili (aktif) sekme beyaz zeminli
    # ve koyu mavi yazılı olsun diye ayarlanıyor.
    style.map('TNotebook.Tab',
              background=[('selected', colors['surface'])],
              foreground=[('selected', colors['primary_dark'])])

    # Normal (vurgusuz) buton görünümü - örn. "Browse", "Clear Log" butonları.
    style.configure('TButton', font=base_font, padding=(10, 5),
                    background=colors['surface'], foreground=colors['text'])
    # Butona tıklanıp basılı tutulduğunda (active) arka planı değişsin.
    style.map('TButton', background=[('active', colors['border'])])

    # 'Accent.TButton' stili: vurgulu/önemli butonlar için (örn. RUN butonları,
    # SOLVE BASE MODEL butonu). Mavi zemin, beyaz yazı - dikkat çeksin diye.
    style.configure('Accent.TButton', font=('Segoe UI', 10, 'bold'), padding=(14, 7),
                    background=colors['primary'], foreground='white')
    # active (tıklanırken) daha koyu mavi, disabled (pasif/gri) iken kenarlık rengi.
    style.map('Accent.TButton',
              background=[('active', colors['primary_dark']), ('disabled', colors['border'])],
              foreground=[('disabled', colors['muted'])])

    # 'Danger.TButton' stili: tehlikeli/geri alınamaz işlemler için kırmızı buton
    # (bu dosyada şu an için tanımlı ama hiçbir yerde kullanılmıyor olabilir -
    # ileride silme/sıfırlama gibi bir buton eklenirse hazır dursun diye).
    style.configure('Danger.TButton', font=('Segoe UI', 10, 'bold'), padding=(14, 7),
                    background=colors['danger'], foreground='white')
    style.map('Danger.TButton',
              background=[('active', colors['danger_dark']), ('disabled', colors['border'])],
              foreground=[('disabled', colors['muted'])])

    # Metin giriş kutuları (Entry) ve açılır liste kutuları (Combobox) görünümü.
    style.configure('TEntry', padding=4, fieldbackground=colors['surface'], bordercolor=colors['border'])
    style.configure('TCombobox', padding=4)

    # İlerleme çubuğu (Progressbar) görünümü - her sekmenin altındaki mavi
    # ilerleme çubuklarının rengi/kalınlığı burada ayarlanıyor.
    style.configure('Horizontal.TProgressbar', troughcolor=colors['border'],
                    background=colors['primary'], bordercolor=colors['border'], thickness=14)

    # Yukarıda tanımlanan renk sözlüğünü global _APP_COLORS içine kopyala,
    # böylece diğer tüm sınıflar (PostProcessingTab, JointRunTab vb.)
    # bu renklere _APP_COLORS['...'] şeklinde erişebilsin.
    _APP_COLORS.update(colors)
    return colors


# ==================== ANA UYGULAMA SINIFI ====================
# IntegratedBDFRFTool: tüm uygulamanın "çatı" sınıfı. Ana pencereyi (root),
# üstteki başlık şeridini (banner) ve sekme (Notebook) çubuğunu oluşturur;
# her bir sekmenin kendi sınıfını (BarPropertySolverTab, PostProcessingTab,
# JointRunTab, FullRunTab, IterativeFullRunTab) örnekleyip Notebook'a ekler.
# "BDF Merge Preparation" sekmesinin arayüzü ise ayrı bir sınıfa çıkarılmamış,
# doğrudan bu sınıfın içinde (setup_tab1 metodunda) tanımlı.
class IntegratedBDFRFTool:
    def __init__(self, root):
        # root: Tkinter'ın ana pencere nesnesi (tk.Tk() ile main() fonksiyonunda
        # oluşturulup buraya parametre olarak geçiriliyor).
        self.root = root
        # Pencere başlığı (Windows'ta pencerenin üst çubuğunda görünen yazı).
        self.root.title("DaDT Joint Optimization Tool")
        # Pencerenin açılış boyutu (genişlik x yükseklik, piksel).
        self.root.geometry("1150x950")
        # Pencerenin küçültülebileceği en küçük boyut sınırı.
        self.root.minsize(1000, 700)
        # Yukarıda tanımlanan _configure_app_style() fonksiyonunu çağırıp
        # tüm uygulamanın renk/tema ayarlarını bir kere, en başta uyguluyoruz.
        # Dönen renk sözlüğü self.colors içinde de saklanıyor (ayrıca global
        # _APP_COLORS içine de yazılmıştı).
        self.colors = _configure_app_style(self.root)

        # ---- Tab 1 (BDF Merge Preparation) değişkenleri ----
        # Bu sekmeye özel veriler; BarPropertySolverTab/PostProcessingTab gibi
        # ayrı sınıflara çıkarılmadığı için doğrudan bu ana sınıfın (self)
        # özellikleri (attribute) olarak tutuluyor.
        self.thermal_bdfs = []    # kullanıcının eklediği "Thermal" BDF dosya yolları listesi
        self.maneuver_bdfs = []   # kullanıcının eklediği "Maneuver" BDF dosya yolları listesi
        self.excel_path = tk.StringVar()   # birleştirme ayarlarını içeren Excel dosyasının yolu
        self.output_folder = tk.StringVar()   # birleştirilmiş BDF'lerin yazılacağı çıktı klasörü
        self.output_thermal_name = tk.StringVar(value="merged_thermal.bdf")   # çıktı thermal BDF'in varsayılan dosya adı
        self.output_maneuver_name = tk.StringVar(value="merged_maneuver.bdf")  # çıktı maneuver BDF'in varsayılan dosya adı
        self.set_id = tk.StringVar(value="99")         # NASTRAN SET kart ID'si (varsayılan 99)
        self.node_set_id = tk.StringVar(value="199")   # NASTRAN düğüm (node) SET kart ID'si (varsayılan 199)
        self.temp_initial = tk.StringVar(value="10")   # başlangıç sıcaklığı değeri (varsayılan 10)

        # Sekme kulakçıklarına eklenen renkli kare ikonların (PhotoImage)
        # Python referanslarını canlı tutmak için liste - referans kalmazsa
        # Tkinter bu imgeleri çöp toplayıcıya (garbage collector) kaptırıp
        # sekmelerdeki renkleri görünmez (boş) yapar.
        self._tab_color_icons = []

        # Arayüzü (tüm sekmeleri) kuran metodu çağır - bu metod aşağıda tanımlı.
        self.setup_ui()

    # _make_tab_icon: verilen hex renk koduyla dolu, küçük kare bir
    # PhotoImage üretir. Sekme (tab) kulakçıklarının arka plan rengini
    # ttk ile doğrudan değiştirmek desteklenmediği için, bu renkli kare
    # ikon her sekmenin metninin hemen önüne (image + compound='left')
    # eklenerek sekmeler grup grup renk kodlanmış gibi gösterilir.
    def _make_tab_icon(self, hex_color, size=14):
        icon = tk.PhotoImage(width=size, height=size)
        icon.put(hex_color, to=(0, 0, size, size))
        self._tab_color_icons.append(icon)  # GC'ye karşı referansı sakla
        return icon

    def setup_ui(self):
        # ---- Üst başlık şeridi (banner) ----
        # Pencerenin en üstünde görünen "🔧 DaDT Joint Optimization Tool"
        # yazısını içeren çerçeve.
        banner = ttk.Frame(self.root, padding=(16, 12))
        banner.pack(fill=tk.X)
        ttk.Label(banner, text="\U0001F527 DaDT Joint Optimization Tool", style='Header.TLabel').pack(side=tk.LEFT)
        ttk.Label(banner, text="  v16.0", style='Subtitle.TLabel').pack(side=tk.LEFT, pady=(6, 0))
        # Banner ile sekme çubuğu arasına yatay bir ayraç (ince çizgi) çiz.
        ttk.Separator(self.root, orient='horizontal').pack(fill=tk.X)

        # ---- Sekme (Notebook) çubuğu ----
        # Notebook: üstte tıklanabilir sekme kulakçıkları (tabs) gösteren,
        # her sekmenin kendi içeriğini barındıran Tkinter widget'ı.
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        # ---- Sekme renk grupları ----
        # BDF Merge Preparation: mavi (tek başına kendi rengi).
        # Nastran Solver / Post Processing / Joint Run: aynı grup, yeşil.
        # Full Run: sarı. Iterative Full Run: bordo. Visualization Results: mor.
        _TAB_COLOR_BLUE = '#2563eb'
        _TAB_COLOR_GREEN = '#16a34a'
        _TAB_COLOR_YELLOW = '#eab308'
        _TAB_COLOR_MAROON = '#800000'
        _TAB_COLOR_PURPLE = '#9333ea'

        # ---- Sekme 1: BDF Merge Preparation (mavi) ----
        # Bu sekmenin içeriği ayrı bir sınıfa çıkarılmamış; setup_tab1()
        # metodu doğrudan bu notebook'un ilk sekmesinin (self.tab1) içini dolduruyor.
        self.tab1 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="\U0001F4C4  BDF Merge Preparation")
        self.notebook.tab(self.tab1, image=self._make_tab_icon(_TAB_COLOR_BLUE), compound='left')
        self.setup_tab1()

        # ---- Sekme 2: Nastran Solver (yeşil) ----
        # BarPropertySolverTab sınıfının bir örneği (self.tab3_tool) oluşturulup
        # self.tab3 çerçevesinin içine yerleştiriliyor. Bu örnek referansı daha
        # sonra Full Run / Iterative Full Run sekmeleri tarafından da (dışarıdan,
        # bu sınıfın kodu hiç değiştirilmeden) kullanılacak.
        self.tab3 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab3, text="\U0001F9EE  Nastran Solver")
        self.notebook.tab(self.tab3, image=self._make_tab_icon(_TAB_COLOR_GREEN), compound='left')
        self.tab3_tool = BarPropertySolverTab(self.tab3, self.root)

        # ---- Sekme 3: Post Processing (yeşil) ----
        # OP2 -> H5 dönüşümü ve Load Extraction klasör hazırlığını yapan sekme.
        self.tab4 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab4, text="\U0001F4E6  Post Processing")
        self.notebook.tab(self.tab4, image=self._make_tab_icon(_TAB_COLOR_GREEN), compound='left')
        self.tab4_tool = PostProcessingTab(self.tab4, self.root)

        # ---- Sekme 4: Joint Run (yeşil) ----
        # Joint Load H5'lerini Set bazında gruplayıp RSS Case Finding / Load
        # Modifying / Joint Run exe'lerini çalıştıran sekme.
        self.tab5 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab5, text="\U0001F517  Joint Run")
        self.notebook.tab(self.tab5, image=self._make_tab_icon(_TAB_COLOR_GREEN), compound='left')
        self.tab5_tool = JointRunTab(self.tab5, self.root)

        # ---- Sekme 5: Full Run (sarı) ----
        # Nastran Solver + Post Processing + Joint Run'ı uç uca
        # zincirleyen sekme. Dikkat: tab3_tool/tab4_tool/tab5_tool nesne
        # referansları FullRunTab'e PARAMETRE olarak geçiriliyor - yani Full Run
        # kendi kopyasını oluşturmuyor, doğrudan yukarıdaki aynı 3 sekme
        # nesnesini (onların StringVar'larını set edip metodlarını çağırarak)
        # dışarıdan yönetiyor.
        self.tab6 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab6, text="⚡  Full Run")
        self.notebook.tab(self.tab6, image=self._make_tab_icon(_TAB_COLOR_YELLOW), compound='left')
        self.tab6_tool = FullRunTab(self.tab6, self.root, self.tab3_tool, self.tab4_tool, self.tab5_tool)

        # ---- Sekme 6: Iterative Full Run (bordo) ----
        # Full Run'ın aynı mantığının N kez döngüye alınmış hali (her
        # iterasyonun sonucu bir sonraki iterasyonun girdisine besleniyor).
        # Aynı şekilde tab3_tool/tab4_tool/tab5_tool referansları paylaşılıyor.
        self.tab7 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab7, text="\U0001F501  Iterative Full Run")
        self.notebook.tab(self.tab7, image=self._make_tab_icon(_TAB_COLOR_MAROON), compound='left')
        self.tab7_tool = IterativeFullRunTab(self.tab7, self.root, self.tab3_tool, self.tab4_tool, self.tab5_tool)

        # ---- Sekme 7: Visualization Results (mor) ----
        # Diğer sekmelerden bağımsız, kendi kendine yeten yeni bir sekme -
        # kullanıcının verdiği bir BDF'i 3B olarak çizer ve Iterative Full
        # Run'ın ürettiği DT_SIZING_BAR_SUMMARY.csv / DT_SIZING_PROP_SUMMARY.csv
        # sonuçlarını (t, d, Material, Life) bar elemanlarının üzerine
        # renk kodlayarak gösterir.
        self.tab8 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab8, text="\U0001F4CA  Visualization Results")
        self.notebook.tab(self.tab8, image=self._make_tab_icon(_TAB_COLOR_PURPLE), compound='left')
        self.tab8_tool = VisualizationResultsTab(self.tab8, self.root)


    # setup_tab1: "BDF Merge Preparation" sekmesinin tüm arayüzünü (Thermal
    # ve Maneuver BDF listeleri, ayarlar, PROCESS butonu, Log paneli) kurar.
    # Bu sekmenin iş mantığı (gerçek BDF birleştirme işlemi) ayrı metodlarda
    # (start_processing, collect_all_lines_from_masters vb., aşağıda).
    def setup_tab1(self):
        # Sekmenin içini dolduracak ana (en dış) çerçeve.
        main = ttk.Frame(self.tab1, padding="10")
        main.pack(fill=tk.BOTH, expand=True)

        # Sekme başlığı.
        ttk.Label(main, text="BDF Merge Preparation v8", style='Header.TLabel').pack(pady=(0,10))

        # === THERMAL BÖLÜMÜ ===
        # "Thermal" tipi master BDF dosyalarının eklenip listelendiği kutu.
        thermal_main = ttk.LabelFrame(main, text="\U0001F321  THERMAL", padding="8")
        thermal_main.pack(fill=tk.X, pady=5)

        # "MASTER BDFs:" etiketi + "Add..."/"Clear" butonları + kaç dosya
        # seçildiğini gösteren sayaç etiketini içeren tek satırlık çerçeve.
        thm_master_f = ttk.Frame(thermal_main)
        thm_master_f.pack(fill=tk.X, pady=2)
        ttk.Label(thm_master_f, text="MASTER BDFs:", width=12).pack(side=tk.LEFT)
        # "Add..." butonu: dosya seçme penceresi açıp self.thermal_bdfs listesine ekler.
        ttk.Button(thm_master_f, text="Add...", command=self.add_thermal_bdfs).pack(side=tk.LEFT, padx=2)
        # "Clear" butonu: self.thermal_bdfs listesini ve listbox'ı temizler.
        ttk.Button(thm_master_f, text="Clear", command=self.clear_thermal_bdfs).pack(side=tk.LEFT, padx=2)
        # Kaç dosya seçildiğini gösteren metin (örn. "3 files") - dinamik güncellenir.
        self.thermal_count = tk.StringVar(value="0 files")
        ttk.Label(thm_master_f, textvariable=self.thermal_count).pack(side=tk.LEFT, padx=5)

        # Seçilen Thermal BDF dosyalarının isimlerini listeleyen kutu (Listbox).
        self.thermal_listbox = tk.Listbox(thermal_main, height=3, width=100, font=('Segoe UI', 9),
                                           bg=self.colors['surface'], fg=self.colors['text'],
                                           selectbackground=self.colors['primary'], selectforeground='white',
                                           relief='flat', highlightthickness=1,
                                           highlightbackground=self.colors['border'],
                                           highlightcolor=self.colors['primary'])
        self.thermal_listbox.pack(fill=tk.X, pady=2)

        # === MANEUVER BÖLÜMÜ ===
        # Yukarıdaki Thermal bölümüyle BİREBİR AYNI yapı, sadece "Maneuver"
        # tipi master BDF dosyaları için (ayrı liste, ayrı listbox, ayrı butonlar).
        maneuver_main = ttk.LabelFrame(main, text="\U0001F6EB  MANEUVER", padding="8")
        maneuver_main.pack(fill=tk.X, pady=5)

        man_master_f = ttk.Frame(maneuver_main)
        man_master_f.pack(fill=tk.X, pady=2)
        ttk.Label(man_master_f, text="MASTER BDFs:", width=12).pack(side=tk.LEFT)
        ttk.Button(man_master_f, text="Add...", command=self.add_maneuver_bdfs).pack(side=tk.LEFT, padx=2)
        ttk.Button(man_master_f, text="Clear", command=self.clear_maneuver_bdfs).pack(side=tk.LEFT, padx=2)
        self.maneuver_count = tk.StringVar(value="0 files")
        ttk.Label(man_master_f, textvariable=self.maneuver_count).pack(side=tk.LEFT, padx=5)

        self.maneuver_listbox = tk.Listbox(maneuver_main, height=3, width=100, font=('Segoe UI', 9),
                                            bg=self.colors['surface'], fg=self.colors['text'],
                                            selectbackground=self.colors['primary'], selectforeground='white',
                                            relief='flat', highlightthickness=1,
                                            highlightbackground=self.colors['border'],
                                            highlightcolor=self.colors['primary'])
        self.maneuver_listbox.pack(fill=tk.X, pady=2)

        # === AYARLAR (SETTINGS) BÖLÜMÜ ===
        # Birleştirme işlemi için gereken Excel ayar dosyası, çıktı klasörü ve
        # NASTRAN kart ID'leri (SET ID, NODE SET ID) ile başlangıç sıcaklığı.
        # grid() ile satır/sütun tablosu şeklinde hizalanmış bir form oluşturuluyor.
        sf = ttk.LabelFrame(main, text="⚙  Settings", padding="10")
        sf.pack(fill=tk.X, pady=5)
        # Satır 0: birleştirme ayarlarını içeren Excel dosyasının yolu.
        ttk.Label(sf, text="Excel:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(sf, textvariable=self.excel_path, width=70).grid(row=0, column=1, padx=5)
        ttk.Button(sf, text="Browse", command=self.browse_excel).grid(row=0, column=2)
        # Satır 1: birleştirilmiş BDF'lerin yazılacağı çıktı klasörü.
        ttk.Label(sf, text="Output:").grid(row=1, column=0, sticky=tk.W)
        ttk.Entry(sf, textvariable=self.output_folder, width=70).grid(row=1, column=1, padx=5)
        ttk.Button(sf, text="Browse", command=self.browse_output).grid(row=1, column=2)
        # Satır 2: oluşturulacak NASTRAN SET kartının ID numarası.
        ttk.Label(sf, text="SET ID:").grid(row=2, column=0, sticky=tk.W)
        ttk.Entry(sf, textvariable=self.set_id, width=10).grid(row=2, column=1, sticky=tk.W, padx=5)
        # Satır 3: oluşturulacak NASTRAN düğüm (node) SET kartının ID numarası.
        ttk.Label(sf, text="NODE SET ID:").grid(row=3, column=0, sticky=tk.W)
        ttk.Entry(sf, textvariable=self.node_set_id, width=10).grid(row=3, column=1, sticky=tk.W, padx=5)
        # Satır 4: başlangıç sıcaklığı (TEMP(INIT) kartı için).
        ttk.Label(sf, text="TEMP(INIT):").grid(row=4, column=0, sticky=tk.W)
        ttk.Entry(sf, textvariable=self.temp_initial, width=10).grid(row=4, column=1, sticky=tk.W, padx=5)

        # === İŞLEM BUTONLARI ===
        bf = ttk.Frame(main)
        bf.pack(fill=tk.X, pady=10)
        # Ana işlem butonu: tıklanınca start_processing() çağrılır (aşağıda
        # tanımlı - arka planda thread açıp gerçek birleştirme işlemini başlatır).
        self.process_btn = ttk.Button(bf, text="▶  PROCESS & MERGE", command=self.start_processing,
                                       style='Accent.TButton')
        self.process_btn.pack(side=tk.LEFT, padx=5)
        # Log panelini temizleyen buton.
        ttk.Button(bf, text="Clear Log", command=self.clear_log1).pack(side=tk.LEFT)

        # İlerleme çubuğu: mode='indeterminate' -> belirsiz süreli işlemler için
        # (belirli bir yüzde göstermek yerine sürekli ileri-geri hareket eden
        # animasyonlu bir çubuk). Diğer sekmelerdeki progress bar'lardan farklı
        # olarak burada belirli bir ilerleme yüzdesi hesaplanmıyor.
        self.progress1 = ttk.Progressbar(main, mode='indeterminate')
        self.progress1.pack(fill=tk.X, pady=5)

        # === LOG PANELİ ===
        # İşlem sırasındaki tüm log mesajlarının yazıldığı kaydırılabilir metin kutusu.
        lf = ttk.LabelFrame(main, text="\U0001F4DC  Log", padding="10")
        lf.pack(fill=tk.BOTH, expand=True)
        self.log_text1 = scrolledtext.ScrolledText(lf, height=15, font=('Consolas', 10),
                                                    bg=self.colors['log_bg'], fg=self.colors['log_fg'],
                                                    insertbackground=self.colors['log_fg'], relief='flat',
                                                    borderwidth=0)
        self.log_text1.pack(fill=tk.BOTH, expand=True)

    # ============= TAB 1 HELPERS =============
    # add_thermal_bdfs: "Add..." butonuna basılınca çağrılır. Dosya seçme
    # penceresi açar, seçilen her BDF dosyasını (zaten listede yoksa)
    # self.thermal_bdfs listesine ve görünen listbox'a ekler, sayaç etiketini günceller.
    def add_thermal_bdfs(self):
        # Çoklu dosya seçimine izin veren (askopenfilenames - sonunda 's' var)
        # standart Windows "Aç" penceresi. Sadece .bdf/.dat/.nas uzantılı
        # dosyalar ya da "Tümü" filtresiyle her dosya seçilebilir.
        files = filedialog.askopenfilenames(filetypes=[("BDF","*.bdf *.dat *.nas"),("All","*.*")])
        for f in files:
            # Aynı dosya iki kere eklenmesin diye kontrol.
            if f not in self.thermal_bdfs:
                self.thermal_bdfs.append(f)
                self.thermal_listbox.insert(tk.END, f)
        # "X files" sayaç etiketini güncel dosya sayısına göre yeniden yaz.
        self.thermal_count.set(f"{len(self.thermal_bdfs)} files")

    # clear_thermal_bdfs: "Clear" butonu - Thermal listesini ve listbox'ı tamamen boşaltır.
    def clear_thermal_bdfs(self):
        self.thermal_bdfs.clear()
        self.thermal_listbox.delete(0, tk.END)
        self.thermal_count.set("0 files")

    # add_maneuver_bdfs: add_thermal_bdfs ile BİREBİR AYNI mantık, sadece
    # Maneuver listesi/listbox'ı için.
    def add_maneuver_bdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("BDF","*.bdf *.dat *.nas"),("All","*.*")])
        for f in files:
            if f not in self.maneuver_bdfs:
                self.maneuver_bdfs.append(f)
                self.maneuver_listbox.insert(tk.END, f)
        self.maneuver_count.set(f"{len(self.maneuver_bdfs)} files")

    # clear_maneuver_bdfs: Maneuver listesini temizler.
    def clear_maneuver_bdfs(self):
        self.maneuver_bdfs.clear()
        self.maneuver_listbox.delete(0, tk.END)
        self.maneuver_count.set("0 files")

    # browse_excel: tek bir Excel dosyası seçtirip self.excel_path'e yazar.
    def browse_excel(self):
        f = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xls")])
        if f: self.excel_path.set(f)

    # browse_output: bir klasör seçtirip self.output_folder'a yazar.
    def browse_output(self):
        f = filedialog.askdirectory()
        if f: self.output_folder.set(f)

    # log1: Tab 1'in Log paneline bir satır mesaj ekler, otomatik en alta
    # kaydırır (see) ve update_idletasks() ile pencerenin HEMEN (bekletmeden)
    # yeniden çizilmesini sağlar - böylece uzun süren işlem sırasında log
    # mesajları anlık olarak ekranda görünür.
    # NOT: Diğer sekmelerdeki log(self, msg) metodları self.root.after(0, ...)
    # kullanarak thread-safe çalışır; bu eski (Tab 1'e özel) log1 metodu ise
    # doğrudan widget'a yazıyor - Tab 1'in işlemi zaten doğrudan (thread
    # içinden) bu şekilde çağırdığı için bugüne kadar sorun çıkarmamış.
    def log1(self, msg):
        self.log_text1.insert(tk.END, msg + "\n")
        self.log_text1.see(tk.END)
        self.root.update_idletasks()

    # clear_log1: "Clear Log" butonu - Tab 1'in log panelini boşaltır.
    def clear_log1(self):
        self.log_text1.delete(1.0, tk.END)

    # format_include_nastran: NASTRAN'ın BDF formatında her satırın en fazla
    # 72 karakter olması gerektiği kuralına uyacak şekilde, uzun bir dosya
    # yolunu içeren INCLUDE kartını birden fazla satıra böler.
    def format_include_nastran(self, abs_path):
        """Uzun INCLUDE path'lerini Nastran formatına uygun böler."""
        # Önce tek satıra sığıyor mu diye dene - sığıyorsa bölmeye gerek yok.
        include_line = f"INCLUDE '{abs_path}'"
        if len(include_line) <= 72:
            return [include_line]
        # Sığmıyorsa: yolu '/' karakterinden parçalara ayırıp, 72 karakter
        # sınırını aşmayacak şekilde parça parça yeni satırlara dağıt.
        parts = abs_path.split('/')
        lines = []
        current_line = "INCLUDE '"
        for i, part in enumerate(parts):
            # Son parça mı (dosya adının kendisi), yoksa ara bir klasör mü?
            # Ara klasörlerin sonuna tekrar '/' eklenir (yol bütünlüğü için).
            is_last = (i == len(parts) - 1)
            segment = part if is_last else part + '/'
            if len(current_line + segment) <= 72:
                # Bu parça mevcut satıra sığıyorsa satıra ekle.
                current_line += segment
            else:
                # Sığmıyorsa mevcut satırı listeye kaydet (boş değilse) ve
                # yeni bir satır başlat.
                if current_line != "INCLUDE '":
                    lines.append(current_line)
                current_line = segment
        # Son kalan satırı da (tırnak kapatarak) listeye ekle.
        if current_line:
            current_line += "'"
            lines.append(current_line)
        return lines

    # read_file_safe: bir dosyayı okurken karşılaşılabilecek farklı karakter
    # kodlamalarına (encoding) karşı dayanıklı okuma yapar. BDF dosyaları
    # farklı programlarla/işletim sistemleriyle üretilmiş olabileceğinden
    # kodlaması garanti değildir; bu yüzden sırayla birkaç kodlama denenir.
    def read_file_safe(self, fpath):
        """Dosyayı güvenli şekilde oku."""
        # Sırasıyla utf-8, latin-1, cp1252 (Windows Batı Avrupa), iso-8859-1
        # kodlamalarını dene - ilk BAŞARILI olan (hata fırlatmayan) kodlamayla
        # dosyanın tamamı okunup döndürülür.
        for enc in ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']:
            try:
                # errors='replace': o kodlamada geçersiz bir bayt dizisi
                # bulunursa hata fırlatmak yerine onu bir "?" benzeri
                # yer tutucu karakterle değiştirip okumaya devam eder.
                with open(fpath, 'r', encoding=enc, errors='replace') as f:
                    return f.read()
            except:
                # Bu kodlama açma/okuma sırasında hata verdiyse bir sonraki
                # kodlamayı dene.
                continue
        # Hiçbir kodlama işe yaramadıysa (çok nadir bir durum) boş string döndür
        # - çağıran kod bunu "dosya okunamadı" gibi değerlendirebilir.
        return ""

    # extract_subcase_load_info: bir BDF dosyasının İÇERİĞİNİ (metin olarak)
    # tarayıp, içindeki HER BİR SUBCASE bloğu için SUBCASE numarası, o
    # subcase'in kullandığı LOAD kart ID'si, TEMPERATURE(LOAD) kart ID'si ve
    # varsa SUBTITLE metnini çıkarır. Bu bilgi daha sonra farklı master
    # BDF'lerdeki subcase'leri birbiriyle eşleştirip birleştirmek için kullanılır.
    def extract_subcase_load_info(self, bdf_content):
        """
        BDF içeriğinden TÜM SUBCASE bilgilerini çıkarır.
        Bir dosyada birden fazla SUBCASE olabilir (MASTER_GUST.BDF gibi).

        Returns: list of dicts, her biri:
            - subcase_id
            - load_id
            - temp_load_id
            - subtitle
        """
        # Bulunan her subcase'in bilgisini biriktireceğimiz sonuç listesi.
        results = []
        # Dosya içeriğini satır satır işlemek için '\n' karakterinden böl.
        lines = bdf_content.split('\n')

        # "Şu an işlenmekte olan" subcase'in bilgilerini tutan geçici
        # değişkenler. Yeni bir SUBCASE satırına rastlandığında bir önceki
        # subcase'in bilgisi results listesine kaydedilip bu değişkenler sıfırlanır.
        current_subcase = None
        current_load = None
        current_temp_load = None
        current_subtitle = None

        # Dosyanın her satırını tek tek (yukarıdan aşağıya) tara.
        for line in lines:
            # Anahtar kelime (SUBCASE, LOAD vb.) karşılaştırması büyük/küçük
            # harfe duyarsız yapılabilsin diye satırın büyük harfli hali.
            upper = line.upper().strip()
            # SUBTITLE gibi büyük/küçük harfin KORUNMASI gereken metinler
            # için satırın orijinal (değiştirilmemiş) hali.
            original = line.strip()

            # --- SUBCASE satırı (örn. "SUBCASE 10") ---
            if upper.startswith('SUBCASE'):
                # Yeni bir SUBCASE'e geçtik demek ki bir önceki subcase'in
                # toplanan bilgisi tamamlandı - onu sonuç listesine kaydet.
                if current_subcase is not None:
                    results.append({
                        'subcase_id': current_subcase,
                        'load_id': current_load,
                        'temp_load_id': current_temp_load,
                        'subtitle': current_subtitle
                    })

                # Şimdi yeni subcase'i başlat: "SUBCASE 10" satırını boşluğa
                # göre parçalara ayırıp 2. kelimeyi (parts[1], subcase ID'si)
                # tam sayıya çevir. Bu yeni subcase için LOAD/TEMP/SUBTITLE
                # bilgileri henüz bilinmediğinden None'a sıfırlanır.
                parts = upper.split()
                if len(parts) >= 2:
                    try:
                        current_subcase = int(parts[1])
                        current_load = None
                        current_temp_load = None
                        current_subtitle = None
                    except:
                        # Sayıya çevrilemiyorsa (beklenmedik format) sessizce atla.
                        pass

            # --- LOAD = satırı (örn. "LOAD = 2001") ---
            # current_subcase doluyken (yani bir SUBCASE bloğunun içindeyken) geçerli.
            elif current_subcase and upper.startswith('LOAD') and '=' in upper:
                # Regex ile "LOAD" kelimesinden sonraki '=' işaretini ve
                # ardından gelen sayıyı yakala.
                match = re.search(r'LOAD\s*=\s*(\d+)', upper)
                if match:
                    current_load = int(match.group(1))

            # --- TEMPERATURE(LOAD) = satırı (örn. "TEMPERATURE(LOAD) = 3001") ---
            elif current_subcase and 'TEMPERATURE' in upper and 'LOAD' in upper and '=' in upper:
                # Parantez içindeki boşluklara da tolerans gösteren regex.
                match = re.search(r'TEMPERATURE\s*\(\s*LOAD\s*\)\s*=\s*(\d+)', upper)
                if match:
                    current_temp_load = int(match.group(1))

            # --- SUBTITLE satırı (örn. "SUBTITLE Gust Load Case 1") ---
            elif current_subcase and upper.startswith('SUBTITLE'):
                # original (büyük/küçük harfi korunmuş) satırdan "SUBTITLE"
                # kelimesinden sonraki metni yakala - bu, subcase'in okunabilir
                # başlığıdır.
                m = re.search(r'SUBTITLE\s+(.+)', original, re.IGNORECASE)
                if m:
                    current_subtitle = m.group(1).strip()

            # --- "BEGIN BULK" satırına ulaşıldıysa ---
            # Bu, dosyanın "Case Control" (SUBCASE tanımları) bölümünün
            # bittiği, "Bulk Data" (gerçek kartlar: GRID, CQUAD4 vb.)
            # bölümünün başladığı anlamına gelir. Artık SUBCASE aranmayacağı
            # için döngüden tamamen çık (break).
            elif upper.startswith('BEGIN') and 'BULK' in upper:
                break

        # Döngü bittiğinde elde kalan EN SON subcase'in bilgisi henüz
        # results listesine eklenmemiş olabilir (çünkü bir sonraki SUBCASE
        # satırı gelmediği için o "kaydet" adımı hiç tetiklenmedi) - onu da
        # burada manuel olarak ekle.
        if current_subcase is not None:
            results.append({
                'subcase_id': current_subcase,
                'load_id': current_load,
                'temp_load_id': current_temp_load,
                'subtitle': current_subtitle
            })

        return results

    # parse_multiline_includes: bir BDF dosyasının içeriğinde geçen tüm
    # INCLUDE kartlarını bulur. INCLUDE kartı, başka bir BDF dosyasının
    # içeriğinin bu dosyaya "dahil edilmesini" söyleyen bir NASTRAN kartıdır
    # (örn. ortak kullanılan GRID/element tanımlarını ayrı bir dosyada tutup
    # birden fazla ana dosyadan INCLUDE ile çağırmak için kullanılır).
    # Zorluk: dosya yolu çok uzunsa (format_include_nastran'ın yaptığının
    # tersi) INCLUDE kartı BİRDEN FAZLA SATIRA yayılabilir - bu fonksiyon
    # böyle çok satırlı INCLUDE'ları da doğru şekilde tek parça halinde toplar.
    def parse_multiline_includes(self, content, bdf_dir):
        """
        Çok satırlı INCLUDE'ları parse eder.
        Nastran formatında INCLUDE şöyle olabilir:

        INCLUDE '../../../_COMMON_/INTERFACE/STRUCTURAL/INTER_AF_HT/
        INTER_AF_HT_STRU.BDF'

        Returns: list of dict with keys: lines, full_text, abs_path, start_idx, end_idx
        """
        # İçeriği satır satır işlemek için böl.
        lines = content.split('\n')
        # Bulunan her INCLUDE kartının bilgisini tutacağımız sonuç listesi.
        includes = []
        i = 0

        # Satırları baştan sona tek tek tara (while döngüsü kullanılıyor
        # çünkü bir INCLUDE bulununca birden fazla satır birden "atlanıyor",
        # bu yüzden normal for yerine elle kontrol edilen indeks (i) gerekiyor).
        while i < len(lines):
            line = lines[i]
            upper = line.upper().strip()

            if upper.startswith('INCLUDE'):
                # Bir INCLUDE kartının başlangıcını bulduk. Şimdi bu kartın
                # TAMAMINI (tek satır mı çok satırlı mı olduğuna bakmaksızın)
                # toplamamız gerekiyor.
                include_lines = [line]
                full_text = line

                # NASTRAN'da dosya yolu tek tırnak (') veya çift tırnak (")
                # içinde yazılabilir - hangisi kullanılmışsa onu tespit et.
                quote_char = None
                if "'" in line:
                    quote_char = "'"
                elif '"' in line:
                    quote_char = '"'

                if quote_char:
                    # Şu ana kadar toplanan metinde kaç tane tırnak karakteri var?
                    # Tırnak SAYISI ÇİFT ise (0, 2, 4...) tırnaklar açılıp
                    # kapanmış demektir - INCLUDE kartı tek satırda tamamlanmış.
                    # Tırnak sayısı TEK ise (1, 3...) tırnak hâlâ açık demektir
                    # - yani path bir sonraki satıra devam ediyor.
                    quote_count = full_text.count(quote_char)

                    # Tırnak kapanana kadar (quote_count çift olana kadar)
                    # bir sonraki satırları da bu INCLUDE kartına dahil et.
                    j = i + 1
                    while quote_count % 2 != 0 and j < len(lines):
                        next_line = lines[j]
                        include_lines.append(next_line)
                        full_text += '\n' + next_line
                        quote_count = full_text.count(quote_char)
                        j += 1

                    # Artık INCLUDE kartının TAM metni elimizde. İçindeki
                    # satır sonlarını (newline) silip tek bir satırmış gibi
                    # birleştir, sonra regex ile tırnaklar ARASINDAKİ
                    # (gerçek dosya yolunu içeren) kısmı yakala.
                    clean_text = full_text.replace('\n', '')
                    match = re.search(rf"INCLUDE\s*{quote_char}([^{quote_char}]*){quote_char}",
                                     clean_text, re.IGNORECASE)
                    if match:
                        inc_path = match.group(1).strip()
                        # Yol GÖRECELİ (relative, örn. "../ortak/dosya.bdf")
                        # ise, bu BDF dosyasının bulunduğu klasöre (bdf_dir)
                        # göre MUTLAK (absolute) yola çevir. Yol zaten mutlaksa
                        # (örn. "C:\..." veya "/...") olduğu gibi normalize et.
                        if not os.path.isabs(inc_path):
                            abs_path = os.path.normpath(os.path.join(bdf_dir, inc_path))
                        else:
                            abs_path = os.path.normpath(inc_path)
                        # Windows'un ters eğik çizgisini (\) NASTRAN/Unix tarzı
                        # düz eğik çizgiye (/) çevir - tutarlı/taşınabilir
                        # bir yol formatı için.
                        abs_path = abs_path.replace('\\', '/')

                        # Bu INCLUDE kartına ait tüm bilgiyi (orijinal
                        # satırlar, birleşik metin, çözümlenmiş mutlak yol,
                        # dosyadaki başlangıç/bitiş satır indeksleri) sonuç
                        # listesine ekle.
                        includes.append({
                            'lines': include_lines,
                            'full_text': full_text,
                            'abs_path': abs_path,
                            'start_idx': i,
                            'end_idx': j - 1 if j > i + 1 else i
                        })

                    # Tarama indeksini, bu INCLUDE kartının bitişinden
                    # (j) devam edecek şekilde ileri al - aynı satırları
                    # tekrar işlememek için.
                    i = j
                else:
                    # "INCLUDE" ile başlıyor ama tırnak yoksa (beklenmedik/
                    # bozuk format) - sadece bir sonraki satıra geç.
                    i += 1
            else:
                # Bu satır INCLUDE ile ilgili değil - bir sonraki satıra geç.
                i += 1

        return includes

    # collect_all_lines_from_masters: birleştirme işleminin KALBİ olan metod.
    # Kullanıcının seçtiği tüm "Master BDF" dosyalarını tek tek açar, Excel'de
    # istenen SUBCASE ID'lerle eşleşen dosyaları bulur, o dosyaların TÜM
    # satırlarını (case-control komutları hariç) tek bir büyük satır listesinde
    # toplar, INCLUDE kartlarını COMMON (ortak, ayrı tutulacak) ve diğer
    # (STRUCTURE/INTERFACE gibi, doğrudan satıra eklenecek) olarak ikiye
    # ayırır, ve en sonunda tekrar eden (duplicate) satırları temizler.
    def collect_all_lines_from_masters(self, bdf_files, load_case_set, common_type):
        """
        Tüm MASTER BDF'lerden TÜM SATIRLARI toplar.

        1. Excel'deki SUBCASE ID'lere uyan MASTER BDF'leri bulur
        2. Her birinden TÜM satırları alır ve alt alta yapıştırır
        3. INCLUDE'ları kategorize et:
           - COMMON LOAD/THERMAL → Ayrı tut (sonra INCLUDE olarak eklenecek)
           - Diğerleri (STRUCTURE, INTERFACE, vs.) → Satır olarak ekle (pyNastran açacak)
        4. Duplicate satırları çıkar

        NOT: Bir BDF birden fazla SUBCASE içerebilir (MASTER_GUST.BDF gibi)

        common_type: 'LOAD' veya 'THERMAL'

        Returns: (all_lines, common_includes, subcase_info_map)
        """
        all_lines_raw = []  # Tüm satırlar (INCLUDE satırları dahil - COMMON hariç)
        common_includes_raw = []  # COMMON INCLUDE path'leri
        # subcase_info_map: her SUBCASE ID'si için (load_id/temp_load_id,
        # subtitle) bilgisini tutan sözlük - process_thermal_bdf/
        # process_maneuver_bdf tarafından yeni SUBCASE kartları yazarken kullanılacak.
        subcase_info_map = {}
        # Aynı dosyayı (birden fazla subcase'i Excel listesiyle eşleştiği
        # için) iki kere tam olarak işlememek amacıyla "zaten işlendi" takibi.
        processed_files = set()

        self.log1(f"    Reading ALL lines from {len(bdf_files)} MASTER BDFs...")
        matched_count = 0
        matched_subcases = 0

        # ---- Her bir Master BDF dosyasını tek tek işle ----
        for bdf_path in bdf_files:
            # INCLUDE yollarını bu dosyanın bulunduğu klasöre göre
            # mutlaklaştırabilmek için dosyanın klasörünü al.
            bdf_dir = os.path.dirname(os.path.abspath(bdf_path))
            # Dosyayı (kodlama sorunlarına dayanıklı şekilde) oku.
            content = self.read_file_safe(bdf_path)

            # Dosya adında "INREL" geçiyorsa bu dosyayı tamamen atla -
            # (muhtemelen bir ara/yardımcı dosya türü, birleştirmeye dahil edilmiyor).
            bdf_basename = os.path.basename(bdf_path).upper()
            if 'INREL' in bdf_basename:
                self.log1(f"      SKIP (INREL): {os.path.basename(bdf_path)}")
                continue

            # Bu dosyadaki TÜM SUBCASE bloklarının bilgisini çıkar (bir
            # dosyada birden fazla subcase olabileceğini unutmayın).
            all_subcases = self.extract_subcase_load_info(content)

            # Bu dosyadaki subcase'lerden hangileri, kullanıcının Excel'de
            # istediği SUBCASE ID kümesiyle (load_case_set) eşleşiyor?
            matching_subcases = []
            for sc_info in all_subcases:
                sc_id = sc_info['subcase_id']
                if sc_id and sc_id in load_case_set:
                    matching_subcases.append(sc_info)

            # En az bir eşleşen subcase varsa bu dosyayı işlemeye devam et.
            if matching_subcases:
                # Bu dosya DAHA ÖNCE (başka bir subcase eşleşmesi yüzünden)
                # tam olarak işlenmişse, satırlarını TEKRAR eklemeye gerek yok
                # (zaten all_lines_raw içinde) - sadece yeni eşleşen
                # subcase'lerin bilgi kaydını (subcase_info_map) tamamla ve
                # bir sonraki dosyaya geç (continue).
                if bdf_path in processed_files:
                    for sc_info in matching_subcases:
                        sc_id = sc_info['subcase_id']
                        if sc_id not in subcase_info_map:
                            if common_type == 'THERMAL':
                                subcase_info_map[sc_id] = {
                                    'temp_load_id': sc_info['temp_load_id'] if sc_info['temp_load_id'] else sc_id,
                                    'subtitle': sc_info['subtitle'] if sc_info['subtitle'] else f"Thermal Case {sc_id}"
                                }
                            else:
                                subcase_info_map[sc_id] = {
                                    'load_id': sc_info['load_id'] if sc_info['load_id'] else sc_id,
                                    'subtitle': sc_info['subtitle'] if sc_info['subtitle'] else f"Manoeuvre {sc_id}"
                                }
                            matched_subcases += 1
                    continue

                # Bu dosyayı "işlendi" olarak işaretle (bir daha tam işlenmesin).
                processed_files.add(bdf_path)
                matched_count += 1

                # Hangi subcase ID'lerinin eşleştiğini log'a yaz (en fazla 5
                # tanesi gösterilsin, fazlası "..." ile kısaltılsın).
                sc_ids = [str(sc['subcase_id']) for sc in matching_subcases]
                self.log1(f"      MATCH: {os.path.basename(bdf_path)} ({len(matching_subcases)} subcases: {', '.join(sc_ids[:5])}{'...' if len(sc_ids) > 5 else ''})")

                # Eşleşen her subcase için bilgi kaydı oluştur (LOAD tipindeyse
                # load_id, THERMAL tipindeyse temp_load_id tutulur; subtitle
                # yoksa otomatik bir varsayılan metin üretilir).
                for sc_info in matching_subcases:
                    sc_id = sc_info['subcase_id']
                    matched_subcases += 1
                    if common_type == 'THERMAL':
                        subcase_info_map[sc_id] = {
                            'temp_load_id': sc_info['temp_load_id'] if sc_info['temp_load_id'] else sc_id,
                            'subtitle': sc_info['subtitle'] if sc_info['subtitle'] else f"Thermal Case {sc_id}"
                        }
                    else:
                        subcase_info_map[sc_id] = {
                            'load_id': sc_info['load_id'] if sc_info['load_id'] else sc_id,
                            'subtitle': sc_info['subtitle'] if sc_info['subtitle'] else f"Manoeuvre {sc_id}"
                        }

                # Bu dosyadaki TÜM INCLUDE kartlarını (çok satırlı olanlar
                # dahil) önceden ayrıştırılmış halde al.
                all_includes = self.parse_multiline_includes(content, bdf_dir)

                # INCLUDE'ları iki kategoriye ayıracağız:
                # - COMMON (LOAD ya da THERMAL): ayrı tutulup, path olarak
                #   fonksiyon çıktısında (common_includes_raw) döndürülecek.
                # - Diğerleri (STRUCTURE/INTERFACE vb.): doğrudan satır
                #   listesine YENİ bir INCLUDE kartı olarak eklenecek (mutlak
                #   yol ile, tek satır halinde - pyNastran bu dosyayı kendisi açacak).
                common_include_indices = set()  # COMMON INCLUDE'a ait, ATLANACAK satır indeksleri
                structure_include_count = 0
                common_include_count = 0

                for inc in all_includes:
                    abs_path_upper = inc['abs_path'].upper()

                    # Bu INCLUDE'un hedef yolu "_COMMON_/LOAD" veya
                    # "_COMMON_/THERMAL" gibi bir klasör içeriyor mu? (hem alt
                    # çizgili hem çizgisiz klasör adı varyasyonu kontrol
                    # ediliyor). common_type parametresine göre hangisine
                    # bakılacağı belirleniyor.
                    is_common = False
                    if common_type == 'LOAD':
                        if '_COMMON_/LOAD' in abs_path_upper or '/COMMON/LOAD' in abs_path_upper:
                            is_common = True
                    elif common_type == 'THERMAL':
                        if '_COMMON_/THERMAL' in abs_path_upper or '/COMMON/THERMAL' in abs_path_upper:
                            is_common = True

                    if is_common:
                        # COMMON INCLUDE: yolunu ayrı listeye kaydet, bu
                        # kartın kapladığı satırları (start_idx..end_idx)
                        # "atlanacaklar" kümesine ekle (aşağıdaki asıl satır
                        # toplama döngüsünde bu satırlar es geçilecek).
                        common_includes_raw.append(inc['abs_path'])
                        for idx in range(inc['start_idx'], inc['end_idx'] + 1):
                            common_include_indices.add(idx)
                        common_include_count += 1
                    else:
                        # COMMON olmayan INCLUDE (örn. STRUCTURE/INTERFACE):
                        # mutlak yol ile TEK SATIRLIK yeni bir INCLUDE kartı
                        # oluşturup doğrudan all_lines_raw'a ekle - böylece
                        # çok satırlı/göreceli orijinal hali yerine tek
                        # satırlık, mutlak yollu, temiz bir INCLUDE kartı kullanılmış olur.
                        include_line = f"INCLUDE '{inc['abs_path']}'"
                        all_lines_raw.append(include_line)
                        # Bu INCLUDE'un orijinal (eski) satırlarını da atlanacaklar kümesine ekle.
                        for idx in range(inc['start_idx'], inc['end_idx'] + 1):
                            common_include_indices.add(idx)
                        structure_include_count += 1

                # Şimdi dosyanın TÜM satırlarını tara ve (INCLUDE satırları
                # hariç - onları yukarıda zaten ayrıca işledik) gerçek veri
                # satırlarını topla.
                lines = content.split('\n')
                line_count = 0

                for idx, line in enumerate(lines):
                    # Bu satır, yukarıda işlediğimiz bir INCLUDE kartının
                    # parçasıysa (COMMON ya da structure fark etmez) atla -
                    # zaten ya common_includes_raw'a ya da yeni tek satırlık
                    # haliyle all_lines_raw'a eklendi.
                    if idx in common_include_indices:
                        continue

                    stripped = line.strip()
                    if not stripped:
                        # Boş satırları atla.
                        continue

                    # "Case Control" bölümüne ait komut satırlarını (SOL,
                    # CEND, TITLE, SUBCASE, LOAD=, SPC=, çıktı istek kartları
                    # DISPLACEMENT/FORCE/vb., BEGIN BULK, ENDDATA gibi)
                    # ATLA - bunlar birleştirilmiş dosyada YENİDEN,
                    # process_thermal_bdf/process_maneuver_bdf tarafından
                    # TEK SEFERDE ve doğru SUBCASE numaralarıyla yazılacak;
                    # buradan ham haliyle kopyalanırsa birden fazla kez
                    # tekrarlanıp bozuk bir dosya oluşur.
                    upper = stripped.upper()
                    skip_keywords = ['SOL ', 'SOL\t', 'CEND', 'TITLE', 'SUBTITLE', 'ECHO',
                                    'SUBCASE', 'LOAD =', 'LOAD=', 'SPC =', 'SPC=',
                                    'TEMPERATURE', 'DISPLACEMENT', 'FORCE', 'GPFORCE',
                                    'OLOAD', 'SPCFORCE', 'SET ', 'BEGIN BULK',
                                    'BEGIN,BULK', 'ENDDATA']
                    is_skip = any(upper.startswith(kw) for kw in skip_keywords)

                    if not is_skip:
                        # Case-control/atlanacak bir satır değilse (yani
                        # gerçek bir Bulk Data kartıysa, örn. GRID, CQUAD4,
                        # PARAM...) - orijinal satırı OLDUĞU GİBİ (formatı
                        # bozmadan) topla.
                        all_lines_raw.append(line)  # Orijinal satırı koru
                        line_count += 1

                self.log1(f"        -> {line_count} data lines, {structure_include_count} structure includes, {common_include_count} common includes")

        # Tüm dosyalar işlendikten sonra özet istatistikleri logla.
        self.log1(f"    Matched {matched_count} MASTER BDFs with {matched_subcases} total subcases")
        self.log1(f"    Total raw lines (including structure includes): {len(all_lines_raw)}")
        self.log1(f"    Total raw COMMON includes: {len(common_includes_raw)}")

        # ---- Tekrar eden (duplicate) satırları temizle ----
        # Birden fazla master BDF'den aynı kart (örn. aynı GRID tanımı) iki
        # farklı dosyada da bulunabilir - bu satırlar birleşik çıktıda SADECE
        # BİR KERE görünmeli, yoksa NASTRAN "tanım tekrarlandı" hatası verir.
        self.log1("    Removing duplicate lines...")
        seen_lines = set()
        unique_lines = []
        for line in all_lines_raw:
            # Karşılaştırmayı boşluk farklarına duyarsız yapmak için satırı
            # "normalize" et: baştaki/sondaki/aradaki fazla boşlukları tek
            # boşluğa indirir (split() + join(' ', ...) kombinasyonu).
            normalized = ' '.join(line.split())
            if normalized not in seen_lines:
                seen_lines.add(normalized)
                unique_lines.append(line)  # orijinal (normalize EDİLMEMİŞ) hali saklanır

        # Aynı mantıkla COMMON INCLUDE yollarındaki tekrarları da temizle
        # (büyük/küçük harf ve \ / ayrımına duyarsız karşılaştırma ile).
        seen_includes = set()
        unique_common_includes = []
        for inc_path in common_includes_raw:
            normalized = inc_path.lower().replace('\\', '/')
            if normalized not in seen_includes:
                seen_includes.add(normalized)
                unique_common_includes.append(inc_path)

        self.log1(f"    After removing duplicates:")
        self.log1(f"      Unique lines: {len(unique_lines)} (removed {len(all_lines_raw) - len(unique_lines)} duplicates)")
        self.log1(f"      Unique COMMON includes: {len(unique_common_includes)}")

        return unique_lines, unique_common_includes, subcase_info_map

    # extract_param_cards: Bulk Data içindeki tüm PARAM kartlarını (örn.
    # "PARAM,POST,-1" gibi NASTRAN çözüm parametreleri) bulur, aynı isimli
    # PARAM'ın BİRDEN FAZLA KEZ geçmesini (duplicate) engeller, ve PARAM
    # kartlarını diğer kartlardan AYRI bir listeye çıkarır (çağıran kod
    # bunları muhtemelen dosyanın başka bir yerine, örn. en üste, topluca
    # yazmak istiyor).
    def extract_param_cards(self, bulk_data):
        """
        Bulk data'dan PARAM kartlarını ayırır.
        Returns: (param_lines, remaining_bulk_data)
        """
        lines = bulk_data.split('\n')
        param_lines = []    # Bulunan PARAM kartları (ve devam satırları) burada toplanır
        other_lines = []    # PARAM olmayan tüm diğer satırlar burada toplanır

        seen_params = set()  # Duplicate PARAM kontrolü - daha önce eklenen PARAM isimleri

        i = 0
        while i < len(lines):
            line = lines[i]
            stripped = line.strip()
            upper = stripped.upper()

            if upper.startswith('PARAM'):
                # Bir PARAM kartı bulduk. Bu kart birden fazla satıra
                # yayılmış (devam/continuation satırları olan) olabilir -
                # bu bloğu toplayacağımız liste.
                param_block = [line]

                # PARAM'ın İSMİNİ çıkar (örn. "PARAM,POST,-1" -> isim "POST")
                # - bu isim, aynı parametrenin tekrar eklenip eklenmediğini
                # kontrol etmek için kullanılacak.
                param_name = None
                try:
                    if ',' in line:
                        # Virgülle ayrılmış (free-field) format: "PARAM,POST,-1"
                        # -> parts[0]="PARAM", parts[1]="POST", isim parts[1].
                        parts = line.split(',')
                        if len(parts) > 1:
                            param_name = parts[1].strip().upper()
                    else:
                        # Sabit genişlikli (fixed-field) format: isim 8-16.
                        # karakterler arasında.
                        if len(line) >= 16:
                            param_name = line[8:16].strip().upper()
                except:
                    pass

                i += 1
                # Bu PARAM kartının devamı niteliğindeki satırları da topla.
                # Bir satır şu durumlarda "devam satırı" sayılır:
                #  - '+' veya '*' ile başlıyorsa (klasik continuation işareti), VEYA
                #  - 8 boşlukla başlıyor, boş değil, '$' (yorum) ile
                #    başlamıyor VE bilinen bir kart ismiyle (GRID, CBAR vb.)
                #    başlamıyorsa (yani yeni bir kart değil, önceki kartın
                #    devamı gibi görünüyorsa).
                while i < len(lines):
                    next_line = lines[i]
                    next_stripped = next_line.strip()
                    is_cont = (next_line.startswith('+') or
                              next_line.startswith('*') or
                              (next_line.startswith('        ') and next_stripped and
                               not next_stripped.startswith('$') and
                               not any(next_stripped.upper().startswith(card) for card in
                                      ['PARAM', 'GRID', 'CBAR', 'CBEAM', 'CQUAD', 'CTRIA',
                                       'MAT', 'PBAR', 'PSHELL', 'FORCE', 'MOMENT', 'RBE',
                                       'CORD', 'SPC', 'MPC', 'INCLUDE', 'ENDDATA'])))
                    if is_cont and next_stripped:
                        param_block.append(next_line)
                        i += 1
                    else:
                        # Devam satırı değilse (yeni bir kart başlıyorsa)
                        # döngüden çık - bu PARAM bloğu tamamlandı.
                        break

                # Bu isimde bir PARAM DAHA ÖNCE eklenmemişse, bloğu sonuç
                # listesine ekle ve ismini "görüldü" kümesine kaydet. Daha
                # önce eklenmiş bir isimse (duplicate) bu blok tamamen ATLANIR.
                if param_name and param_name not in seen_params:
                    seen_params.add(param_name)
                    param_lines.extend(param_block)
            else:
                # PARAM kartı değilse olduğu gibi "diğer satırlar" listesine ekle.
                other_lines.append(line)
                i += 1

        # (PARAM kartları, geri kalan bulk data) ikilisini döndür.
        return param_lines, '\n'.join(other_lines)

    # check_and_remove_duplicates: collect_all_lines_from_masters'tan gelen
    # devasa, birden fazla dosyadan birleştirilmiş satır listesinde kalan
    # (orada normalize edilmiş METİN bazlı duplicate temizliğinden kaçabilmiş)
    # tekrar eden NASTRAN kartlarını, NASTRAN'ın kendi mantığına göre daha
    # akıllıca tespit edip ayıklar. İki farklı duplicate kontrol stratejisi
    # kullanılıyor:
    #   1. "ID bazlı" kartlar (GRID, element'ler, property'ler, materyaller,
    #      koordinat sistemleri): bu kartlarda asıl kimlik onların ID
    #      numarasıdır - aynı ID iki kez tanımlanırsa (içerikleri farklı bile
    #      olsa) NASTRAN hata verir, bu yüzden SADECE ID'ye bakılarak
    #      ikincisi atılır.
    #   2. "Satır bazlı" kartlar (SPC, FORCE, MOMENT, MPC, LOAD, TEMP, TEMPD):
    #      bunlarda aynı ID birden fazla kez meşru olarak görülebilir (örn.
    #      aynı SET ID'sine farklı düğümler için ayrı SPC kartları), bu
    #      yüzden sadece kartın TAMAMI (tüm alanlarıyla) birebir AYNI ise
    #      duplicate sayılır.
    def check_and_remove_duplicates(self, bulk_data):
        """
        Bulk data içindeki duplicate kartları tespit edip kaldırır.

        - Element/Property/Material kartları: ID bazlı kontrol (aynı ID → duplicate)
        - SPC/FORCE/MOMENT kartları: Tüm satır bazlı kontrol (birebir aynıysa → duplicate)
        """
        self.log1("    Checking for duplicate entries...")

        lines = bulk_data.split('\n')

        # ID bazlı kontrol yapılacak kart türleri. Her kart tipi için ayrı
        # bir set() tutuluyor - "bu tipte, şu ID'ler şu ana kadar görüldü"
        # bilgisini saklamak için. (GRID düğümleri, çubuk/kabuk/katı
        # elemanlar, rijit bağlantılar RBE2/RBE3, kesit özellikleri PBAR/
        # PSHELL/PCOMP vb., malzeme kartları MAT1/MAT2/..., koordinat
        # sistemleri CORD1x/CORD2x.)
        id_based_cards = {
            'GRID': set(),
            'CBAR': set(),
            'CBEAM': set(),
            'CROD': set(),
            'CONROD': set(),
            'CQUAD4': set(),
            'CQUAD8': set(),
            'CTRIA3': set(),
            'CTRIA6': set(),
            'CHEXA': set(),
            'CPENTA': set(),
            'CTETRA': set(),
            'CBUSH': set(),
            'CELAS1': set(),
            'CELAS2': set(),
            'CDAMP1': set(),
            'CDAMP2': set(),
            'CMASS1': set(),
            'CMASS2': set(),
            'RBE2': set(),
            'RBE3': set(),
            'PBAR': set(),
            'PBARL': set(),
            'PBEAM': set(),
            'PBEAML': set(),
            'PROD': set(),
            'PSHELL': set(),
            'PCOMP': set(),
            'PCOMPG': set(),
            'PSOLID': set(),
            'PBUSH': set(),
            'PELAS': set(),
            'PDAMP': set(),
            'PMASS': set(),
            'PTUBE': set(),
            'PVISC': set(),
            'PGAP': set(),
            'PWELD': set(),
            'MAT1': set(),
            'MAT2': set(),
            'MAT8': set(),
            'MAT9': set(),
            'MATS1': set(),
            'CORD1R': set(),
            'CORD2R': set(),
            'CORD1C': set(),
            'CORD2C': set(),
            'CORD1S': set(),
            'CORD2S': set(),
        }

        # Satır bazlı (ID değil, kartın TAMAMI) kontrol yapılacak kart
        # türleri: yük/sınır şartı kartları.
        line_based_cards = ['SPC', 'SPC1', 'FORCE', 'MOMENT', 'MPC', 'LOAD', 'TEMP', 'TEMPD']
        seen_full_lines = set()  # Satır-bazlı kartlar için "şu ana kadar görülen" normalize edilmiş tam kart metinleri

        # Her kart tipinden kaç tane duplicate (tekrar) bulunup atıldığını
        # saymak için - en sonda kullanıcıya özet rapor göstermek amacıyla.
        duplicate_counts = {}

        # Sonuçta (duplicate'ler çıkarılmış hali) kalacak satırlar.
        result_lines = []
        i = 0

        # Tüm satırları, continuation (devam) satırlarını blok halinde
        # atlayabilmek için while döngüsüyle (elle ilerletilen indeksle) tara.
        while i < len(lines):
            line = lines[i]
            stripped = line.strip()

            # Boş satır veya '$' ile başlayan yorum satırı - duplicate
            # kontrolüne tabi değil, direkt sonuç listesine ekle ve geç.
            if not stripped or stripped.startswith('$'):
                result_lines.append(line)
                i += 1
                continue

            upper = stripped.upper()

            # Bu satırın hangi kart TÜRÜNE ait olduğunu (varsa) tespit et.
            card_type = None
            is_line_based = False

            # Önce satır-bazlı (line_based_cards) listesinde bir eşleşme var
            # mı diye bak. Kart ismi tam eşleşmeli VEYA ismin hemen
            # ardından boşluk/virgül/tab/yıldız gelmeli (örn. "SPC1" satırını
            # yanlışlıkla "SPC" ile eşleştirmemek için bu kontrol önemli).
            for ctype in line_based_cards:
                if upper.startswith(ctype) and (len(upper) == len(ctype) or
                    upper[len(ctype)] in ' ,\t*'):
                    card_type = ctype
                    is_line_based = True
                    break

            # Satır-bazlı bir kart değilse, ID-bazlı kart listesinde
            # (id_based_cards sözlüğünün anahtarları) bir eşleşme ara.
            if not card_type:
                for ctype in id_based_cards.keys():
                    if upper.startswith(ctype) and (len(upper) == len(ctype) or
                        upper[len(ctype)] in ' ,\t*'):
                        card_type = ctype
                        break

            # Bilinen bir kart türüyle eşleşti mi?
            if card_type:
                if is_line_based:
                    # ---- SATIR-BAZLI KONTROL (SPC/FORCE/MOMENT/MPC/LOAD/TEMP/TEMPD) ----
                    # Bu kartın (varsa devam satırlarıyla birlikte) TÜM
                    # metnini topla - karşılaştırma TÜM kart üzerinden yapılacak.
                    full_card = [line]
                    j = i + 1

                    # Aynı "devam satırı mı?" testi (PARAM/check_and_remove
                    # fonksiyonlarındaki ile aynı mantık): '+' / '*' ile
                    # başlıyor, YA DA 8 boşlukla başlayıp boş/yorum olmayan
                    # ve bilinen bir yeni kart ismiyle başlamayan bir satır.
                    while j < len(lines):
                        next_line = lines[j]
                        next_stripped = next_line.strip()
                        is_cont = (next_line.startswith('+') or
                                  next_line.startswith('*') or
                                  (next_line.startswith('        ') and next_stripped and
                                   not next_stripped.startswith('$') and
                                   not any(next_stripped.upper().startswith(ct) for ct in
                                          list(id_based_cards.keys()) + line_based_cards)))
                        if is_cont and next_stripped:
                            full_card.append(next_line)
                            j += 1
                        else:
                            break

                    # Kartın TÜM satırlarını (boşluk farklarına duyarsız
                    # şekilde normalize edip) '|' karakteriyle birleştirerek
                    # TEK bir karşılaştırma anahtarı (normalized_card) oluştur.
                    normalized_card = '|'.join(' '.join(l.split()) for l in full_card)

                    if normalized_card in seen_full_lines:
                        # Bu TAM kart daha önce görülmüş (birebir aynı) -
                        # duplicate sayılır, sonuç listesine EKLENMEDEN atlanır.
                        if card_type not in duplicate_counts:
                            duplicate_counts[card_type] = 0
                        duplicate_counts[card_type] += 1
                        i = j
                        continue
                    else:
                        # İlk kez görülüyor - "görüldü" kümesine ekle ve
                        # kartın TÜM satırlarını sonuç listesine ekle.
                        seen_full_lines.add(normalized_card)
                        result_lines.extend(full_card)
                        i = j
                        continue
                else:
                    # ---- ID-BAZLI KONTROL (GRID, elemanlar, property'ler, materyaller, koordinatlar) ----
                    # Bu kart tipinde sadece ID numarasına bakılır - kartın
                    # geri kalan içeriği (koordinatlar, malzeme değerleri
                    # vb.) karşılaştırmaya dahil edilmez.
                    card_id = None
                    try:
                        if ',' in line:
                            # Virgülle ayrılmış format: 2. alan (parts[1]) ID'dir.
                            parts = line.split(',')
                            if len(parts) > 1:
                                id_str = parts[1].strip()
                                if id_str:
                                    card_id = int(float(id_str))
                        else:
                            # Sabit genişlikli format: ID, 8-16. karakterler arası.
                            if len(line) >= 16:
                                id_str = line[8:16].strip()
                                if id_str:
                                    card_id = int(float(id_str))
                    except:
                        # ID okunamazsa (beklenmedik format) sessizce geç -
                        # bu durumda card_id None kalır ve aşağıdaki
                        # "if card_id is not None" bloğuna hiç girilmez,
                        # satır normal şekilde (duplicate kontrolü yapılmadan) eklenir.
                        pass

                    if card_id is not None:
                        if card_id in id_based_cards[card_type]:
                            # Bu ID, bu kart tipinde DAHA ÖNCE görülmüş -
                            # duplicate. Sayacı artır.
                            if card_type not in duplicate_counts:
                                duplicate_counts[card_type] = 0
                            duplicate_counts[card_type] += 1

                            # Bu kartın devam satırlarını da (aynı "devam
                            # satırı mı?" testiyle) bul ve hepsini ATLA -
                            # sadece ilk satırı değil, kartın TAMAMINI
                            # sonuçtan çıkarmamız gerekiyor.
                            i += 1
                            while i < len(lines):
                                next_line = lines[i]
                                next_stripped = next_line.strip()
                                is_cont = (next_line.startswith('+') or
                                          next_line.startswith('*') or
                                          (next_line.startswith('        ') and next_stripped and
                                           not next_stripped.startswith('$') and
                                           not any(next_stripped.upper().startswith(ct) for ct in
                                                  list(id_based_cards.keys()) + line_based_cards)))
                                if is_cont and next_stripped:
                                    i += 1
                                else:
                                    break
                            continue
                        else:
                            # Bu ID ilk kez görülüyor - "görüldü" kümesine ekle
                            # (devam satırları ayrıca işlenmeden, normal akışla
                            # - yani aşağıdaki result_lines.append(line) ile -
                            # sonuç listesine eklenecek, çünkü burada sadece
                            # İLK satır/ID kontrol edildi, continue YOK).
                            id_based_cards[card_type].add(card_id)

            # Buraya gelindiyse: ya bilinen bir kart tipi değildi, ya da
            # ID-bazlı bir karttı ve ilk kez görülüyordu (duplicate değildi)
            # - her iki durumda da bu TEK satırı sonuç listesine ekle ve
            # bir sonraki satıra geç. (ID-bazlı kartların devam satırları da
            # döngünün bir sonraki turlarında aynı şekilde - "bilinen kart
            # tipiyle eşleşmedikleri için" - buradan teker teker eklenmiş olur.)
            result_lines.append(line)
            i += 1

        # ---- Özet rapor ----
        # Hangi kart tipinden kaç tane duplicate silindiğini kullanıcıya logla.
        if duplicate_counts:
            self.log1("    Removed duplicates:")
            for ctype, count in sorted(duplicate_counts.items()):
                self.log1(f"      {ctype}: {count}")
            total = sum(duplicate_counts.values())
            self.log1(f"      TOTAL: {total} duplicate entries removed")
        else:
            self.log1("    No duplicates found")

        # Duplicate'leri temizlenmiş satırları tekrar tek bir metne
        # (newline ile birleştirerek) döndür.
        return '\n'.join(result_lines)

    # merge_lines_with_pynastran: collect_all_lines_from_masters'tan gelen
    # (ve check_and_remove_duplicates ile temizlenmiş) satır listesini GERÇEK
    # bir BDF dosyası olarak geçici (temp) bir dosyaya yazar, sonra bu dosyayı
    # pyNastran ile OKUTARAK (içindeki INCLUDE kartlarını da takip ederek)
    # tam/düzgün bir model haline getirir ve TEKRAR BDF formatında geri yazdırır
    # - böylece pyNastran'ın kendi iç format düzenlemesinden/doğrulamasından
    # geçmiş, tutarlı bir bulk data elde edilir. pyNastran bu işlemi
    # yapamazsa (örn. hâlâ bir duplicate ID hatası varsa), expand_includes_
    # manually() ile daha "kaba" ama garantili bir yönteme geri düşülür.
    def merge_lines_with_pynastran(self, lines):
        """
        Satırları temp BDF'e yazıp pyNastran ile merge eder.
        Duplicate ID hatası olursa, satırları direkt yazar.
        Returns: merged bulk data string
        """
        if not lines:
            self.log1("    WARNING: No lines to merge!")
            return ""

        # Sistemin geçici dosya klasöründe (örn. Windows'ta %TEMP%) çalışma dosyası yolu.
        temp_bdf_path = os.path.join(tempfile.gettempdir(), "_temp_lines_to_merge.bdf")

        self.log1(f"    Writing {len(lines)} lines to temp BDF...")

        try:
            # Geçici, GEÇERLİ bir BDF dosyası oluştur: minimal bir case-control
            # bölümü (SOL 101 / CEND / BEGIN BULK) ekleyip ardından TÜM
            # toplanan satırları yaz, en sona ENDDATA koy. pyNastran, geçerli
            # bir BDF dosyası beklediği için bu "iskelet" gerekli.
            with open(temp_bdf_path, 'w', encoding='utf-8') as f:
                f.write("$ Temporary BDF for merging\n")
                f.write("SOL 101\n")
                f.write("CEND\n")
                f.write("BEGIN BULK\n")
                for line in lines:
                    f.write(line + "\n")
                f.write("ENDDATA\n")

            self.log1(f"    Reading temp BDF with pyNastran (following includes)...")

            try:
                # pyNastran'ın BDF sınıfıyla bu geçici dosyayı oku.
                # xref=False: cross-reference (referans çözümleme) YAPMA -
                # sadece ham kartları oku, daha hızlı.
                # read_includes=True: içindeki INCLUDE kartlarının
                # gösterdiği dosyaları da OTOMATİK olarak açıp içeriklerini
                # modele dahil et (asıl istediğimiz bu).
                bdf = BDF(debug=False)
                # allow_duplicate_ids ile dene
                bdf.read_bdf(temp_bdf_path, validate=False, xref=False,
                            read_includes=True, save_file_structure=False)

                # Okuma başarılıysa modeldeki eleman/özellik sayılarını logla
                # (kullanıcıya "şu kadar düğüm/eleman okundu" bilgisi vermek için).
                self.log1(f"      Loaded: {len(bdf.nodes)} nodes, {len(bdf.elements)} elements")
                self.log1(f"      Properties: {len(bdf.properties)}, Materials: {len(bdf.materials)}")
                self.log1(f"      Coords: {len(bdf.coords)}, MPCs: {len(bdf.mpcs)}, SPCs: {len(bdf.spcs)}")

                # pyNastran'ın kendi write_bdf() metoduyla modeli TEKRAR bir
                # BDF dosyasına yaz - bu, pyNastran'ın kartları kendi
                # standart formatında (tutarlı sütun genişlikleri vb. ile)
                # yeniden ürettiği, "temizlenmiş" bir çıktı dosyasıdır.
                # size=8: sabit 8 karakterlik (small field) format kullan.
                merged_temp_path = os.path.join(tempfile.gettempdir(), "_temp_merged.bdf")
                bdf.write_bdf(merged_temp_path, size=8, is_double=False)

                # pyNastran'ın yazdığı dosyayı tekrar metin olarak oku.
                with open(merged_temp_path, 'r', errors='ignore') as f:
                    merged_content = f.read()

                # Artık işimiz bitti, geçici "merged" dosyasını sil.
                if os.path.exists(merged_temp_path): os.remove(merged_temp_path)

            except Exception as e:
                # pyNastran okuma/yazma sırasında bir hata verdi (örn. hâlâ
                # çözülmemiş bir duplicate ID ya da format sorunu olabilir).
                # Bu durumda pyNastran'ı tamamen devre dışı bırakıp, daha
                # basit ama her zaman çalışan bir "manuel INCLUDE açma"
                # yöntemine (expand_includes_manually) geri dönülüyor.
                self.log1(f"    pyNastran failed: {str(e)[:100]}")
                self.log1(f"    Falling back to direct file reading...")

                # pyNastran başarısız oldu - INCLUDE'ları manuel aç
                merged_content = self.expand_includes_manually(temp_bdf_path)

            # Başta oluşturduğumuz geçici giriş dosyasını (artık gerek
            # kalmadığı için) temizle.
            if os.path.exists(temp_bdf_path): os.remove(temp_bdf_path)

            # Elimizdeki içerik hem case-control hem bulk data bölümünü
            # içeriyor olabilir - sadece "BEGIN BULK"tan SONRAKİ kısmı (gerçek
            # bulk data kartlarını) almak istiyoruz, bu yüzden regex ile
            # "BEGIN BULK" (ya da "BEGIN,BULK") metnini bulup ondan sonrasını kes.
            bulk_match = re.search(r'BEGIN\s*,?\s*BULK', merged_content, re.IGNORECASE)
            if bulk_match:
                bulk_data = merged_content[bulk_match.end():]
            else:
                # "BEGIN BULK" hiç bulunamadıysa (beklenmedik durum) tüm
                # içeriği bulk data olarak kabul et.
                bulk_data = merged_content

            # pyNastran'ın kendi eklediği yorum satırlarını ($pyNastran ile
            # başlayanlar) ve artık gereksiz olan ENDDATA/INCLUDE/SOL/CEND/
            # BEGIN gibi "iskelet" satırlarını (bunlar zaten üst seviyede,
            # process_thermal_bdf/process_maneuver_bdf tarafından YENİDEN
            # oluşturulacağı için) temizle.
            result_lines = []
            for l in bulk_data.split('\n'):
                if l.startswith('$pyNastran'): continue
                if l.strip().upper().startswith('ENDDATA'): continue
                if l.strip().upper().startswith('INCLUDE'): continue
                if l.strip().upper().startswith('SOL '): continue
                if l.strip().upper().startswith('CEND'): continue
                if l.strip().upper().startswith('BEGIN'): continue
                result_lines.append(l)

            result = '\n'.join(result_lines)
            self.log1(f"      Merged bulk data: {len(result)} characters")
            return result

        except Exception as e:
            # Beklenmedik herhangi bir hata (dosya yazma izni, disk dolu
            # vb.) olursa logla, hata izini (traceback) göster, geçici
            # dosyayı temizlemeye çalış ve boş string döndür (işlem
            # başarısız oldu anlamına gelir, çağıran kod bunu kontrol eder).
            self.log1(f"    ERROR merging: {e}")
            import traceback
            self.log1(traceback.format_exc())
            if os.path.exists(temp_bdf_path): os.remove(temp_bdf_path)
            return ""

    # expand_includes_manually: pyNastran'ın BDF okuma/INCLUDE-takip etme
    # işlemi BAŞARISIZ olduğunda devreye giren "yedek plan" (fallback)
    # yöntem. pyNastran'ın karmaşık ayrıştırma mantığına güvenmeden, sadece
    # INCLUDE kartlarını BULUP, gösterdikleri dosyanın içeriğini OLDUĞU GİBİ
    # (metin olarak) o INCLUDE satırının yerine yapıştırır. Daha basit/kaba
    # bir yöntem ama pyNastran'ın takıldığı durumlarda bile çalışır.
    def expand_includes_manually(self, bdf_path):
        """
        INCLUDE'ları manuel olarak açar (pyNastran başarısız olduğunda).
        """
        self.log1("    Expanding includes manually...")

        content = self.read_file_safe(bdf_path)
        bdf_dir = os.path.dirname(os.path.abspath(bdf_path))

        # Bu dosyadaki tüm INCLUDE kartlarını (çok satırlı olanlar dahil) bul.
        all_includes = self.parse_multiline_includes(content, bdf_dir)

        # Satırları işle
        lines = content.split('\n')
        result_lines = []

        # Her INCLUDE kartının kapladığı satır indekslerini (başlangıç-bitiş
        # arası HER satır indeksi), o INCLUDE'un bilgisine eşleyen bir
        # sözlük oluştur - aşağıdaki ana döngüde "bu satır bir INCLUDE'un
        # parçası mı, öyleyse hangisinin?" sorusuna hızlıca cevap vermek için.
        include_indices = {}
        for inc in all_includes:
            for idx in range(inc['start_idx'], inc['end_idx'] + 1):
                include_indices[idx] = inc

        # Aynı dosyanın birden fazla INCLUDE kartıyla (nadiren) iki kez
        # çağrılması durumunda onu sadece BİR KEZ açmak için takip kümesi.
        processed_includes = set()

        for idx, line in enumerate(lines):
            if idx in include_indices:
                inc = include_indices[idx]
                # Bu INCLUDE'un (olası çok satırlı bloğunun) sadece İLK
                # satırında (start_idx) gerçek açma işlemini yap - devam
                # satırlarında (start_idx'ten sonraki satırlarda) hiçbir
                # şey eklenmez, yani o satırlar sessizce atlanmış olur.
                if idx == inc['start_idx']:
                    inc_path = inc['abs_path']
                    if inc_path not in processed_includes:
                        processed_includes.add(inc_path)
                        # INCLUDE'un gösterdiği hedef dosyayı oku ve
                        # içeriğini, nereden geldiğini belirten bir yorum
                        # satırıyla birlikte sonuç listesine ekle.
                        if os.path.exists(inc_path):
                            try:
                                inc_content = self.read_file_safe(inc_path)
                                result_lines.append(f"$ === EXPANDED: {os.path.basename(inc_path)} ===")
                                for inc_line in inc_content.split('\n'):
                                    # Bu dosyanın İÇİNDE başka bir INCLUDE
                                    # kartı varsa (iç içe/recursive INCLUDE),
                                    # onu AÇMAYA ÇALIŞMADAN sadece atla -
                                    # bu fonksiyon sadece TEK SEVİYE açma
                                    # yapıyor, basit tutuluyor.
                                    if not inc_line.strip().upper().startswith('INCLUDE'):
                                        result_lines.append(inc_line)
                            except Exception as e:
                                # Hedef dosya okunamadıysa (bozuk/erişim
                                # hatası vb.) bunu bir yorum satırı olarak
                                # çıktıya yaz - en azından sorun görünür olsun.
                                result_lines.append(f"$ ERROR reading {inc_path}: {e}")
                        else:
                            # Hedef dosya hiç yoksa (silinmiş/yanlış yol)
                            # bunu da bir yorum satırıyla belirt.
                            result_lines.append(f"$ FILE NOT FOUND: {inc_path}")
            else:
                # INCLUDE ile ilgisi olmayan normal bir satır - olduğu gibi aktar.
                result_lines.append(line)

        self.log1(f"      Expanded {len(processed_includes)} includes")
        return '\n'.join(result_lines)

    # start_processing: "PROCESS & MERGE" butonuna basılınca çağrılır.
    # Önce temel alanların (en az bir BDF listesi, Excel, çıktı klasörü)
    # doldurulduğunu kontrol eder, sonra butonu kilitler, ilerleme çubuğunu
    # (belirsiz/animasyonlu modda) başlatır ve asıl işi (process_merge)
    # arka planda bir thread'de çalıştırır - GUI dondurulmasın diye.
    def start_processing(self):
        if not self.thermal_bdfs and not self.maneuver_bdfs:
            messagebox.showerror("Error","Add BDF files"); return
        if not self.excel_path.get():
            messagebox.showerror("Error","Select Excel"); return
        if not self.output_folder.get():
            messagebox.showerror("Error","Select output folder"); return
        self.process_btn.config(state=tk.DISABLED)
        self.progress1.start()
        threading.Thread(target=self.process_merge, daemon=True).start()

    # process_merge: Tab 1'in asıl iş akışı. Excel'den gerekli bilgileri
    # (hangi SUBCASE'lerin isteneceği, hangi element/node ID'lerinin dahil
    # edileceği) okur, sonra THERMAL ve MANEUVER birleştirme işlemlerini
    # (varsa, kullanıcı o türden BDF eklemişse) sırayla başlatır.
    def process_merge(self):
        try:
            self.log1("="*70)
            self.log1("BDF Merger Tool v8")
            self.log1("="*70)

            self.log1("\n[1] Reading Excel...")
            # Excel dosyasını pandas ile aç ve içindeki sheet (sayfa) isimlerini al.
            xl = pd.ExcelFile(self.excel_path.get())
            sheets = xl.sheet_names
            self.log1(f"    Available sheets: {sheets}")

            thermal_sh = maneuver_sh = element_sh = node_sh = None

            # ---- Birinci geçiş: isim bazlı TAM/özel eşleşmeleri ara ----
            # Excel sheet isimleri kullanıcıdan kullanıcıya değişebileceği
            # için, önce bilinen anahtar kelimelere göre akıllı bir eşleştirme
            # deneniyor (sabit sheet SIRASINA güvenmek yerine).
            for s in sheets:
                sl = s.lower()
                # "Element_Set" ismiyle TAM eşleşme (öncelikli).
                if sl == 'element_set' or sl == 'elementset':
                    element_sh = s
                # "Node_Set" ismiyle TAM eşleşme (öncelikli).
                elif sl == 'node_set' or sl == 'nodeset':
                    node_sh = s
                # İçinde "thermal" geçen ilk sheet.
                elif 'thermal' in sl and not thermal_sh:
                    thermal_sh = s
                # İçinde "maneuver" (İngilizce) veya "manevra" (Türkçe) geçen ilk sheet.
                elif ('maneuver' in sl or 'manevra' in sl) and not maneuver_sh:
                    maneuver_sh = s

            # ---- İkinci geçiş: element_sh/node_sh hâlâ bulunamadıysa, daha
            # gevşek (kısmi) eşleşme dene ----
            # (örn. "Set_Element" gibi kelime sırası farklı bir isim de
            # yakalanabilsin diye "element" VE "set" kelimelerinin İKİSİNİN
            # DE geçmesi yeterli sayılıyor, sırası önemli değil.)
            if not element_sh:
                for s in sheets:
                    sl = s.lower()
                    if 'element' in sl and 'set' in sl:
                        element_sh = s
                        break
            if not node_sh:
                for s in sheets:
                    sl = s.lower()
                    if 'node' in sl and 'set' in sl:
                        node_sh = s
                        break

            # ---- Son çare: hâlâ bulunamayan sheet'ler için SIRAYA göre
            # (0=Thermal, 1=Maneuver, 2=Element_Set, 3=Node_Set) varsay ----
            # Bu, eski/isimlendirme kuralına uymayan Excel dosyalarıyla bile
            # bir şekilde çalışabilmek için bir geriye dönük uyumluluk tedbiri.
            if not thermal_sh and len(sheets) > 0: thermal_sh = sheets[0]
            if not maneuver_sh and len(sheets) > 1: maneuver_sh = sheets[1]
            if not element_sh and len(sheets) > 2: element_sh = sheets[2]
            if not node_sh and len(sheets) > 3: node_sh = sheets[3]

            self.log1(f"    Using sheets -> Thermal: '{thermal_sh}', Maneuver: '{maneuver_sh}', Element_Set: '{element_sh}', Node_Set: '{node_sh}'")

            # Her sheet'in İLK sütununu oku, boş hücreleri at (dropna),
            # tam sayıya çevir (astype(int)) ve Python listesine dönüştür.
            # Bu listeler: hangi SUBCASE ID'lerinin isteneceği (thermal/
            # maneuver_cases) ve hangi element/node ID'lerinin dahil
            # edileceği (element_ids/node_ids, ayrıca sıralı hale getiriliyor).
            thermal_cases = pd.read_excel(xl, sheet_name=thermal_sh).iloc[:,0].dropna().astype(int).tolist() if thermal_sh else []
            maneuver_cases = pd.read_excel(xl, sheet_name=maneuver_sh).iloc[:,0].dropna().astype(int).tolist() if maneuver_sh else []
            element_ids = sorted(pd.read_excel(xl, sheet_name=element_sh).iloc[:,0].dropna().astype(int).tolist()) if element_sh else []
            node_ids = sorted(pd.read_excel(xl, sheet_name=node_sh).iloc[:,0].dropna().astype(int).tolist()) if node_sh else []

            self.log1(f"    Thermal cases: {len(thermal_cases)}")
            self.log1(f"    Maneuver cases: {len(maneuver_cases)}")
            self.log1(f"    Element IDs: {len(element_ids)}")
            self.log1(f"    Node IDs: {len(node_ids)}")

            # GUI'deki Settings bölümünden girilen sabit değerleri oku.
            set_id = int(self.set_id.get())
            node_set_id = int(self.node_set_id.get())
            temp_initial = self.temp_initial.get()
            out_dir = self.output_folder.get()
            # Çıktı klasörü yoksa oluştur (varsa dokunma - exist_ok=True).
            os.makedirs(out_dir, exist_ok=True)

            # THERMAL BDF'ler eklendiyse, THERMAL birleştirme işlemini başlat.
            if self.thermal_bdfs:
                self.log1("\n" + "="*70)
                self.log1("[2] Processing THERMAL...")
                self.log1("="*70)
                self.process_thermal_bdf(self.thermal_bdfs, thermal_cases, element_ids, set_id,
                    node_ids, node_set_id, temp_initial, os.path.join(out_dir, self.output_thermal_name.get()))

            # MANEUVER BDF'ler eklendiyse, MANEUVER birleştirme işlemini başlat.
            if self.maneuver_bdfs:
                self.log1("\n" + "="*70)
                self.log1("[3] Processing MANEUVER...")
                self.log1("="*70)
                self.process_maneuver_bdf(self.maneuver_bdfs, maneuver_cases, element_ids, set_id,
                    node_ids, node_set_id, os.path.join(out_dir, self.output_maneuver_name.get()))

            self.log1("\n" + "="*70)
            self.log1("COMPLETED!")
            self.log1("="*70)
            # İşlem tamamlandığında kullanıcıya bir "Done" pop-up'ı göster.
            # self.root.after(0, ...) kullanılıyor çünkü bu kod bir arka plan
            # thread'inde çalışıyor ve Tkinter widget'larına (messagebox
            # dahil) sadece ANA thread'den dokunulmalı - after(0, ...) bu
            # çağrıyı ana thread'in event döngüsüne güvenli şekilde iletir.
            self.root.after(0, lambda: messagebox.showinfo("Done","Merge completed!"))
        except Exception as e:
            # İşlem sırasında herhangi bir hata olursa logla, hata izini
            # göster ve kullanıcıya bir hata pop-up'ı göster.
            self.log1(f"\nERROR: {e}")
            import traceback
            self.log1(traceback.format_exc())
            self.root.after(0, lambda: messagebox.showerror("Error",str(e)))
        finally:
            # Hata olsun olmasın (finally): ilerleme çubuğunu durdur ve
            # "PROCESS & MERGE" butonunu tekrar tıklanabilir hale getir.
            self.root.after(0, lambda: [self.progress1.stop(), self.process_btn.config(state=tk.NORMAL)])

    # process_thermal_bdf: THERMAL BDF birleştirme işleminin tamamı. Önceki
    # yardımcı metodları (collect_all_lines_from_masters, merge_lines_with_
    # pynastran, check_and_remove_duplicates, extract_param_cards,
    # format_include_nastran) sırayla çağırıp, sonunda EKSİKSİZ, GEÇERLİ bir
    # NASTRAN BDF dosyası (kendi SOL/CEND/SET/SUBCASE başlıkları + birleşmiş
    # bulk data + COMMON INCLUDE'lar + ENDDATA) yazıp diske kaydeder.
    def process_thermal_bdf(self, bdf_files, load_cases, element_ids, set_id, node_ids, node_set_id, temp_initial, output_path):
        self.log1(f"    MASTER BDFs: {len(bdf_files)}")
        self.log1(f"    Load cases to match: {len(load_cases)}")
        # Hızlı üyelik kontrolü (in) için listeyi set'e çevir.
        load_case_set = set(load_cases)

        # ---- Adım 1: istenen SUBCASE'lere ait tüm satırları topla ----
        self.log1("\n    === Step 1: Collecting all lines from MASTER BDFs ===")
        all_lines, common_includes, subcase_info_map = self.collect_all_lines_from_masters(
            bdf_files, load_case_set, 'THERMAL'
        )

        # ---- Adım 2: toplanan satırları pyNastran ile "temiz" bir bulk data'ya dönüştür ----
        self.log1("\n    === Step 2: Merging with pyNastran ===")
        merged_bulk_data = self.merge_lines_with_pynastran(all_lines)

        # ---- Adım 2.5: pyNastran'dan sonra bile kalmış olabilecek
        # duplicate kartları (özellikle farklı master dosyalardan gelen,
        # pyNastran'ın göz ardı etmiş olabileceği tekrarları) temizle ----
        if merged_bulk_data:
            self.log1("\n    === Step 2.5: Checking duplicates ===")
            merged_bulk_data = self.check_and_remove_duplicates(merged_bulk_data)

        # ---- Adım 3: nihai çıktı BDF dosyasının TÜM satırlarını sırayla oluştur ----
        self.log1("\n    === Step 3: Writing output BDF ===")
        out = []
        # Dosya başlığı/yorumu.
        out.append(f"$ {'='*60}")
        out.append(f"$ THERMAL - MERGED BDF (v8)")
        out.append(f"$ {'='*60}")
        # NASTRAN çözüm tipi (SOL 101 = lineer statik analiz), Case Control
        # bölümünün başlangıcı (CEND), çıktı ekosunu kapat, başlık,
        # başlangıç sıcaklığı kartı.
        out.append("SOL 101")
        out.append("CEND")
        out.append("ECHO=NONE")
        out.append(f"TITLE = THERMAL ANALYSIS")
        out.append(f"TEMPERATURE(INITIAL) = {temp_initial}")
        out.append("$")

        # ---- SET tanımı (eleman ID'leri) ----
        # NASTRAN'da bir SET kartı tek satıra sığmayabilir (çok fazla ID
        # varsa) - bu yüzden element_ids listesi, her satır en fazla ~60
        # karakter olacak şekilde "chunk"lara (parçalara) bölünüyor.
        chunks = []
        current = ""
        for eid in element_ids:
            test = f"{current},{eid}" if current else str(eid)
            if len(test) > 60 and current:
                # Bu ID eklenirse satır çok uzayacak - mevcut parçayı
                # kapat, yeni bir parçaya bu ID ile başla.
                chunks.append(current)
                current = str(eid)
            else:
                current = test
        if current: chunks.append(current)

        # Her parçayı ayrı bir satır olarak yaz. İlk satır "SET <id> = ..."
        # ile başlar, ARA satırlar devam niteliğinde girintili yazılır, SON
        # satırın sonunda virgül OLMAZ (NASTRAN'da kart bittiği anlamına gelir).
        for i, chunk in enumerate(chunks):
            if i == 0:
                out.append(f"SET {set_id} = {chunk}" + ("," if len(chunks) > 1 else ""))
            elif i == len(chunks) - 1:
                out.append(f"         {chunk}")
            else:
                out.append(f"         {chunk},")

        # ---- SET tanımı (düğüm/node ID'leri) ----
        # Yukarıdakiyle BİREBİR AYNI mantık, bu sefer node_ids listesi için.
        node_chunks = []
        current = ""
        for nid in node_ids:
            test = f"{current},{nid}" if current else str(nid)
            if len(test) > 60 and current:
                node_chunks.append(current)
                current = str(nid)
            else:
                current = test
        if current: node_chunks.append(current)

        for i, chunk in enumerate(node_chunks):
            if i == 0:
                out.append(f"SET {node_set_id} = {chunk}" + ("," if len(node_chunks) > 1 else ""))
            elif i == len(node_chunks) - 1:
                out.append(f"         {chunk}")
            else:
                out.append(f"         {chunk},")

        # ---- Çıktı istek kartları (hangi sonuçların hesaplanıp raporlanacağı) ----
        # DISPLACEMENT: tüm düğümlerin yer değiştirmesi.
        # FORCE: yukarıda tanımlanan eleman SET'i için iç kuvvetler.
        # GPFORCE: yukarıda tanımlanan düğüm SET'i için düğüm kuvvetleri.
        # OLOAD/SPCFORCE: uygulanan yükler ve sınır şartı (SPC) tepki kuvvetleri.
        out.append("$")
        out.append("DISPLACEMENT(SORT1,PLOT,REAL)=ALL")
        out.append(f"FORCE(SORT1,PLOT,REAL,CENTER)={set_id}")
        out.append(f"GPFORCE(PLOT)={node_set_id}")
        out.append("OLOAD(PLOT)=ALL")
        out.append("SPCFORCE(SORT1,PLOT)=ALL")
        out.append("$")

        # ---- SUBCASE tanımları ----
        # Excel'den istenen HER load case (lc) için bir SUBCASE bloğu yaz.
        # Eğer bu subcase'in bilgisi (temp_load_id, subtitle)
        # collect_all_lines_from_masters tarafından önceden tespit
        # edilmişse (subcase_info_map içinde varsa) o bilgi kullanılır;
        # yoksa (eşleşme bulunamadıysa) makul bir varsayılan (subcase ID'sinin
        # kendisini TEMPERATURE(LOAD) ID'si olarak kullanmak, otomatik
        # bir "Thermal Case N" başlığı üretmek) uygulanır.
        for lc in load_cases:
            if lc in subcase_info_map:
                info = subcase_info_map[lc]
                temp_load = info['temp_load_id']
                subtitle = info['subtitle']
            else:
                temp_load = lc
                subtitle = f"Thermal Case {lc}"
            out.append(f"SUBCASE {lc}")
            out.append(f"SUBTITLE {subtitle}")
            out.append("SPC = 1")
            out.append(f"TEMPERATURE(LOAD) = {temp_load}")
            out.append("$")

        # Case Control bölümü bitti, Bulk Data bölümü başlıyor.
        out.append("BEGIN BULK")

        # ---- PARAM kartlarını ayır ve BEGIN BULK'tan HEMEN SONRA yaz ----
        # (NASTRAN'da PARAM kartlarının bulk data'nın başında olması
        # alışılagelmiş/tercih edilen bir konumdur, bu yüzden merged_bulk_
        # data'nın geri kalanından ayrı tutulup en öne yazılıyor.)
        param_lines = []
        if merged_bulk_data:
            param_lines, merged_bulk_data = self.extract_param_cards(merged_bulk_data)

        if param_lines:
            out.append("$ --- PARAM CARDS ---")
            out.extend(param_lines)
            out.append("$")

        # ---- Birleştirilmiş yapısal/model verisi (geri kalan tüm bulk data) ----
        if merged_bulk_data:
            out.append(f"$ {'='*60}")
            out.append(f"$ MERGED STRUCTURE DATA")
            out.append(f"$ {'='*60}")
            out.append(merged_bulk_data)

        # ---- Ortak (COMMON) Thermal INCLUDE kartları ----
        # collect_all_lines_from_masters tarafından "COMMON" olarak
        # ayrılmış (satır içine kopyalanmayıp sadece yolu saklanmış)
        # INCLUDE'lar burada, dosyanın EN SONUNDA, tek tek INCLUDE kartları
        # olarak (gerekirse çok satıra bölünerek, format_include_nastran
        # ile) yazılıyor. Alfabetik sıraya (sorted) konularak çıktı tutarlı/
        # öngörülebilir oluyor.
        out.append("$")
        out.append(f"$ {'='*60}")
        out.append(f"$ COMMON THERMAL INCLUDES ({len(common_includes)} files)")
        out.append(f"$ {'='*60}")

        for abs_path in sorted(common_includes):
            include_lines = self.format_include_nastran(abs_path)
            out.extend(include_lines)

        out.append("$")
        out.append("ENDDATA")

        # Oluşturulan tüm satırları tek bir metin olarak birleştirip çıktı
        # dosyasına yaz.
        with open(output_path, 'w') as f:
            f.write('\n'.join(out))

        self.log1(f"    Output: {os.path.basename(output_path)}")
        self.log1(f"    COMMON THERMAL INCLUDES: {len(common_includes)}")

    # process_maneuver_bdf: process_thermal_bdf ile BİREBİR AYNI akışı
    # (satırları topla -> pyNastran ile birleştir -> duplicate temizle ->
    # SET/SUBCASE başlıklarıyla birlikte çıktı BDF'ini yaz) izler, aradaki
    # TEK farkı şu: THERMAL'da "TEMPERATURE(LOAD)" kartı kullanılırken
    # MANEUVER'da normal "LOAD" kartı kullanılır (ve TEMPERATURE(INITIAL)
    # satırı hiç yoktur - maneuver analizinde başlangıç sıcaklığı kavramı
    # yok). Aşağıdaki yorumlar bu yüzden daha kısa tutuldu; adım adım aynı
    # mantık için process_thermal_bdf'in yorumlarına bakılabilir.
    def process_maneuver_bdf(self, bdf_files, load_cases, element_ids, set_id, node_ids, node_set_id, output_path):
        self.log1(f"    MASTER BDFs: {len(bdf_files)}")
        self.log1(f"    Load cases to match: {len(load_cases)}")
        load_case_set = set(load_cases)

        # Adım 1: istenen SUBCASE'lere ait tüm satırları topla (common_type='LOAD').
        self.log1("\n    === Step 1: Collecting all lines from MASTER BDFs ===")
        all_lines, common_includes, subcase_info_map = self.collect_all_lines_from_masters(
            bdf_files, load_case_set, 'LOAD'
        )

        # Adım 2: pyNastran ile birleştir.
        self.log1("\n    === Step 2: Merging with pyNastran ===")
        merged_bulk_data = self.merge_lines_with_pynastran(all_lines)

        # Adım 2.5: kalan duplicate kartları temizle.
        if merged_bulk_data:
            self.log1("\n    === Step 2.5: Checking duplicates ===")
            merged_bulk_data = self.check_and_remove_duplicates(merged_bulk_data)

        # Adım 3: çıktı BDF dosyasını satır satır oluştur.
        self.log1("\n    === Step 3: Writing output BDF ===")
        out = []
        out.append(f"$ {'='*60}")
        out.append(f"$ MANEUVER - MERGED BDF (v8)")
        out.append(f"$ {'='*60}")
        # THERMAL'dan farklı olarak burada TEMPERATURE(INITIAL) kartı YOK
        # - maneuver (manevra) analizinde başlangıç sıcaklığı tanımlanmaz.
        out.append("SOL 101")
        out.append("CEND")
        out.append("ECHO=NONE")
        out.append(f"TITLE = MANEUVER ANALYSIS")
        out.append("$")

        # ---- SET tanımı (eleman ID'leri) - process_thermal_bdf ile aynı
        # "60 karakteri aşma, yeni satıra böl" mantığı. ----
        chunks = []
        current = ""
        for eid in element_ids:
            test = f"{current},{eid}" if current else str(eid)
            if len(test) > 60 and current:
                chunks.append(current)
                current = str(eid)
            else:
                current = test
        if current: chunks.append(current)

        for i, chunk in enumerate(chunks):
            if i == 0:
                out.append(f"SET {set_id} = {chunk}" + ("," if len(chunks) > 1 else ""))
            elif i == len(chunks) - 1:
                out.append(f"         {chunk}")
            else:
                out.append(f"         {chunk},")

        # ---- SET tanımı (düğüm ID'leri) - yine aynı mantık. ----
        node_chunks = []
        current = ""
        for nid in node_ids:
            test = f"{current},{nid}" if current else str(nid)
            if len(test) > 60 and current:
                node_chunks.append(current)
                current = str(nid)
            else:
                current = test
        if current: node_chunks.append(current)

        for i, chunk in enumerate(node_chunks):
            if i == 0:
                out.append(f"SET {node_set_id} = {chunk}" + ("," if len(node_chunks) > 1 else ""))
            elif i == len(node_chunks) - 1:
                out.append(f"         {chunk}")
            else:
                out.append(f"         {chunk},")

        # ---- Çıktı istek kartları (THERMAL ile aynı). ----
        out.append("$")
        out.append("DISPLACEMENT(SORT1,PLOT,REAL)=ALL")
        out.append(f"FORCE(SORT1,PLOT,REAL,CENTER)={set_id}")
        out.append(f"GPFORCE(PLOT)={node_set_id}")
        out.append("OLOAD(PLOT)=ALL")
        out.append("SPCFORCE(SORT1,PLOT)=ALL")
        out.append("$")

        # ---- SUBCASE tanımları ----
        # THERMAL'daki "TEMPERATURE(LOAD) = temp_load" yerine burada
        # "LOAD = load_id" kartı yazılıyor - manevra analizinde uygulanan
        # yük, TEMPERATURE değil LOAD kartıyla referans gösterilir.
        for lc in load_cases:
            if lc in subcase_info_map:
                info = subcase_info_map[lc]
                load_id = info['load_id']
                subtitle = info['subtitle']
            else:
                load_id = lc
                subtitle = f"Manoeuvre {lc}"
            out.append(f"SUBCASE {lc}")
            out.append(f"SUBTITLE {subtitle}")
            out.append("SPC = 1")
            out.append(f"LOAD = {load_id}")
            out.append("$")

        out.append("BEGIN BULK")

        # PARAM kartlarını ayır ve BEGIN BULK'tan hemen sonra yaz (THERMAL ile aynı).
        param_lines = []
        if merged_bulk_data:
            param_lines, merged_bulk_data = self.extract_param_cards(merged_bulk_data)

        if param_lines:
            out.append("$ --- PARAM CARDS ---")
            out.extend(param_lines)
            out.append("$")

        # Birleştirilmiş yapısal veri.
        if merged_bulk_data:
            out.append(f"$ {'='*60}")
            out.append(f"$ MERGED STRUCTURE DATA")
            out.append(f"$ {'='*60}")
            out.append(merged_bulk_data)

        # ---- Ortak (COMMON) LOAD INCLUDE kartları ----
        # (THERMAL'daki "COMMON THERMAL" bölümünün LOAD karşılığı.)
        out.append("$")
        out.append(f"$ {'='*60}")
        out.append(f"$ COMMON LOAD INCLUDES ({len(common_includes)} files)")
        out.append(f"$ {'='*60}")

        for abs_path in sorted(common_includes):
            include_lines = self.format_include_nastran(abs_path)
            out.extend(include_lines)

        out.append("$")
        out.append("ENDDATA")

        # Oluşturulan tüm satırları çıktı dosyasına yaz.
        with open(output_path, 'w') as f:
            f.write('\n'.join(out))

        self.log1(f"    Output: {os.path.basename(output_path)}")
        self.log1(f"    COMMON LOAD INCLUDES: {len(common_includes)}")


# ============================================================================
# BarPropertySolverTab  --  "Understanding Structure Type" sekmesi (Tab 3)
# ============================================================================
# Bu sınıf kullanıcıdan BDF dosyalarını ve Property/Element Excel'lerini alır,
# Excel'de verilen Bar (PBARL dim1/dim2) ve Skin (PSHELL t) değerlerini BDF'e
# uygular, offset (göçme/landing) elemanlarını işler ve ortaya çıkan
# "base model"i gerçek Nastran çözücüsüyle (tek parça veya SUBCASE'lere göre
# paralel parçalara bölünmüş olarak) çalıştırır. Diğer sekmeler (Full Run,
# Iterative Full Run) bu sınıfın nesnesini DOĞRUDAN paylaşır (kopyasını değil)
# - yani bu sınıftaki self.* alanlarına yazılan her şey diğer sekmelerden de
# görülür.
class BarPropertySolverTab:
    def __init__(self, parent_frame, root):
        self.root = root
        self.parent_frame = parent_frame
        self.maneuver_bdfs = []  # "Maneuver BDF (Offset Source)" listesi - offset hesaplamasının referans aldığı BDF'ler
        # self.root.title("Joint Structure Type Understanding")
        # self.root.geometry("1300x950")

        # ---- Girdi yolları (Input paths) ----
        self.bdf_paths = []                    # "All BDF Files" listbox'ındaki tüm BDF dosya yolları
        # Master Input Excel: Bar/Skin Property VE Offset Element ID
        # sayfalarının ikisini de içeren TEK Excel dosyası. Arayüzde artık
        # tek bir "Master Input Excel" alanı olarak gösterildiği için
        # element_excel_path, property_excel_path ile AYNI StringVar
        # nesnesini paylaşır (ikisi de aynı dosya yolunu gösterir) -
        # load_properties()/load_element_ids() metodlarının kendisi hiç
        # değişmedi, sadece ikisi de artık aynı dosyayı okuyor.
        self.property_excel_path = tk.StringVar()   # Master Input Excel dosya yolu
        self.element_excel_path = self.property_excel_path    # aynı Master Input Excel
        self.nastran_path = tk.StringVar()     # Nastran çalıştırılabilir dosyasının (.exe) yolu
        self.nastran_memory = tk.StringVar(value="")  # opsiyonel bellek sınırı (örn. "4gb") - boşsa Nastran varsayılanı kullanılır
        self.scratch_folder = tk.StringVar(value="")  # optional; blank = each run's own output folder (still collision-free)
        self.output_folder = tk.StringVar()    # çözüm sonuçlarının yazılacağı klasör
        self.cases_per_split = tk.StringVar(value="")   # optional; blank/0 = solve as a single BDF (no split)
        self.max_parallel_runs = tk.StringVar(value="1")  # 1 = sequential

        # Hiçbir Excel/BDF değeri bulunamazsa kullanılacak varsayılan bar kalınlığı
        self.default_bar_thickness = 2.0

        # ---- Veri saklama alanları (Data storage) ----
        self.bdf_model = None
        self.bdf_models = []
        self.bar_properties = {}          # PID -> {'dim1': val, 'dim2': val} (base model values)
        self.skin_properties = {}         # PID -> {'thickness': val} from 'Skin Property' sheet
        self.pbarl_dims = {}              # PID -> {'dim1': val, 'dim2': val} from BDF
        self.current_bar_thicknesses = {} # PID -> base thickness (from load_properties)
        self.current_skin_thicknesses = {}# PID -> effective thickness (Excel 'Skin Property' if given, else BDF), for base model + offsets
        self.original_skin_thicknesses = {} # PID -> raw PSHELL thickness from BDF (fallback only)
        self.original_bar_thicknesses = {} # PID -> original dim1 from BDF
        self.maneuver_base_path = None    # Excel-properties-applied maneuver (offset source) BDF, set by _build_base_model
        self.landing_elem_ids = []        # offset uygulanacak "landing" (göçme) eleman ID'leri
        self.bar_offset_elem_ids = []     # offset uygulanacak bar eleman ID'leri

        # ---- Çalışma (run) durumu ----
        self.is_running = False       # bir Nastran çözümü şu an sürüyor mu?
        self.properties_loaded = False  # Property Excel'i başarıyla yüklendi mi? (SOLVE butonunu aktive etmek için kontrol edilir)

        self.setup_ui()

    # ==================== GUI ====================
    # setup_ui: bu sekmenin tüm arayüzünü kurar. İçerik dikeyde kaydırılabilir
    # (Canvas + Scrollbar + içteki scrollable_frame) bir alana yerleştirilir,
    # çünkü sekmede 4 ayrı bölüm (Input Files / Parallel Run Splitting /
    # Actions / Log) alt alta sığdırılır.
    def setup_ui(self):
        canvas = tk.Canvas(self.parent_frame)
        scrollbar = ttk.Scrollbar(self.parent_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        main = scrollable_frame

        # Başlık ve alt başlık.
        ttk.Label(main, text="Nastran Solver", style='Header.TLabel').pack(pady=10)
        ttk.Label(main, text="Apply Property Excel + offsets and solve the base model",
                 style='Subtitle.TLabel').pack()

        # ---- Section 1: Input Files ----
        f1 = ttk.LabelFrame(main, text="\U0001F4C1  1. Input Files", padding=10)
        f1.pack(fill=tk.X, pady=5, padx=10)

        # "All BDF Files" listesi: solve edilecek tüm BDF dosyalarının
        # (örn. ana yapı + maneuver + thermal) seçildiği, Add/Remove
        # butonlu bir Listbox.
        bdf_frame = ttk.Frame(f1)
        bdf_frame.pack(fill=tk.X, pady=2)
        ttk.Label(bdf_frame, text="All BDF Files:", width=18).pack(side=tk.LEFT, anchor=tk.N)

        bdf_list_frame = ttk.Frame(bdf_frame)
        bdf_list_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.bdf_listbox = tk.Listbox(bdf_list_frame, height=3, width=55, selectmode=tk.SINGLE,
                                       font=('Segoe UI', 9), bg=_APP_COLORS['surface'], fg=_APP_COLORS['text'],
                                       selectbackground=_APP_COLORS['primary'], selectforeground='white',
                                       relief='flat', highlightthickness=1,
                                       highlightbackground=_APP_COLORS['border'],
                                       highlightcolor=_APP_COLORS['primary'])
        self.bdf_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        bdf_scroll = ttk.Scrollbar(bdf_list_frame, orient="vertical", command=self.bdf_listbox.yview)
        bdf_scroll.pack(side=tk.LEFT, fill=tk.Y)
        self.bdf_listbox.configure(yscrollcommand=bdf_scroll.set)

        bdf_btn_frame = ttk.Frame(bdf_frame)
        bdf_btn_frame.pack(side=tk.LEFT, padx=5)
        ttk.Button(bdf_btn_frame, text="Add", command=self.add_bdf).pack(fill=tk.X, pady=1)
        ttk.Button(bdf_btn_frame, text="Remove", command=self.remove_bdf).pack(fill=tk.X, pady=1)

        self.bdf_status = ttk.Label(f1, text="No BDF files loaded", foreground="gray")
        self.bdf_status.pack(anchor=tk.W)

        # "Maneuver BDF (Offset Source)" listesi: offset hesabının referans
        # alacağı, SADECE maneuver BDF'lerinin seçildiği ayrı bir liste
        # (All BDF Files'tan bağımsız - aynı dosya her ikisine de eklenebilir).
        man_frame = ttk.Frame(f1)
        man_frame.pack(fill=tk.X, pady=2)
        man_lbl = ttk.Frame(man_frame)
        man_lbl.pack(side=tk.LEFT, anchor=tk.N)
        ttk.Label(man_lbl, text="Maneuver BDF:", width=18).pack(anchor=tk.W)
        ttk.Label(man_lbl, text="  (Offset Source)", foreground="blue").pack(anchor=tk.W)

        man_list_frame = ttk.Frame(man_frame)
        man_list_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.maneuver_listbox = tk.Listbox(man_list_frame, height=3, width=55, selectmode=tk.SINGLE,
                                            font=('Segoe UI', 9), bg=_APP_COLORS['surface'], fg=_APP_COLORS['text'],
                                            selectbackground=_APP_COLORS['primary'], selectforeground='white',
                                            relief='flat', highlightthickness=1,
                                            highlightbackground=_APP_COLORS['border'],
                                            highlightcolor=_APP_COLORS['primary'])
        self.maneuver_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        man_scroll = ttk.Scrollbar(man_list_frame, orient="vertical", command=self.maneuver_listbox.yview)
        man_scroll.pack(side=tk.LEFT, fill=tk.Y)
        self.maneuver_listbox.configure(yscrollcommand=man_scroll.set)

        man_btn_frame = ttk.Frame(man_frame)
        man_btn_frame.pack(side=tk.LEFT, padx=5)
        ttk.Button(man_btn_frame, text="Add", command=self._add_maneuver_bdf).pack(fill=tk.X, pady=1)
        ttk.Button(man_btn_frame, text="Remove", command=self._remove_maneuver_bdf).pack(fill=tk.X, pady=1)


        # Master Input Excel satırı (yol kutusu + Browse + Load butonu, bkz.
        # _add_file_row). Bar/Skin Property ve Offset Element ID sayfalarının
        # İKİSİ de bu TEK Excel'de bulunur; "Load" butonu her ikisini de
        # yükler (bkz. _load_master_input).
        self._add_file_row(f1, "Master Input Excel:", self.property_excel_path, self._load_master_input)
        self.prop_status = ttk.Label(f1, text="Not loaded", foreground="gray")
        self.prop_status.pack(anchor=tk.W)
        self.elem_status = ttk.Label(f1, text="Not loaded", foreground="gray")
        self.elem_status.pack(anchor=tk.W)

        # Nastran .exe yolu
        nast_frame = ttk.Frame(f1)
        nast_frame.pack(fill=tk.X, pady=2)
        ttk.Label(nast_frame, text="Nastran Exe:", width=18).pack(side=tk.LEFT)
        ttk.Entry(nast_frame, textvariable=self.nastran_path, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(nast_frame, text="Browse", command=lambda: self.nastran_path.set(
            filedialog.askopenfilename(filetypes=[("Executable", "*.exe"), ("All", "*.*")])
        )).pack(side=tk.LEFT)

        # NX Nastran Requested Memory (optional user-imposed cap, e.g. "4gb", "500mb")
        mem_frame = ttk.Frame(f1)
        mem_frame.pack(fill=tk.X, pady=2)
        ttk.Label(mem_frame, text="Requested Memory:", width=18).pack(side=tk.LEFT)
        ttk.Entry(mem_frame, textvariable=self.nastran_memory, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Label(mem_frame, text="(optional, e.g. 4gb / 500mb - blank = Nastran default)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        # NX Nastran Scratch Folder (optional; each run gets its own unique
        # subfolder under it, so parallel runs never collide on scratch files)
        scr_frame = ttk.Frame(f1)
        scr_frame.pack(fill=tk.X, pady=2)
        ttk.Label(scr_frame, text="Scratch Folder:", width=18).pack(side=tk.LEFT)
        ttk.Entry(scr_frame, textvariable=self.scratch_folder, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(scr_frame, text="Browse", command=lambda: self.scratch_folder.set(
            filedialog.askdirectory()
        )).pack(side=tk.LEFT)
        scr_hint = ttk.Frame(f1)
        scr_hint.pack(fill=tk.X, pady=(0, 2))
        ttk.Label(scr_hint, text="(recommended if C: is low on space - blank = Nastran uses its own default "
                                  "temp location. If set, each run gets its own subfolder here, set as that "
                                  "run's TEMP/TMP, so scratch never lands on C: or collides between runs)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        # Çözüm sonuçlarının yazılacağı çıktı klasörü.
        out_frame = ttk.Frame(f1)
        out_frame.pack(fill=tk.X, pady=2)
        ttk.Label(out_frame, text="Output Folder:", width=18).pack(side=tk.LEFT)
        ttk.Entry(out_frame, textvariable=self.output_folder, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(out_frame, text="Browse", command=lambda: self.output_folder.set(
            filedialog.askdirectory()
        )).pack(side=tk.LEFT)

        # ---- Section 2: Parallel Run Splitting ----
        # Büyük SUBCASE sayılı BDF'leri küçük parçalara bölüp paralel
        # Nastran süreçleriyle çalıştırmak için ayarlar.
        f2 = ttk.LabelFrame(main, text="\U0001F500  2. Parallel Run Splitting", padding=10)
        f2.pack(fill=tk.X, pady=5, padx=10)

        split_frame = ttk.Frame(f2)
        split_frame.pack(fill=tk.X, pady=2)
        ttk.Label(split_frame, text="Cases per Split:", width=18).pack(side=tk.LEFT)
        ttk.Entry(split_frame, textvariable=self.cases_per_split, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(split_frame, text="(blank/0 = solve as one BDF; e.g. 100 with 500 SUBCASEs -> "
                                     "run1..run5, same header/bulk data, 100 cases each)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        parallel_frame = ttk.Frame(f2)
        parallel_frame.pack(fill=tk.X, pady=2)
        ttk.Label(parallel_frame, text="Max Parallel Runs:", width=18).pack(side=tk.LEFT)
        ttk.Entry(parallel_frame, textvariable=self.max_parallel_runs, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(parallel_frame, text="(1 = one Nastran run at a time; e.g. 5 = up to 5 runN files at once - "
                                        "check available Nastran licenses/CPU/RAM first)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        # ---- Section 3: Actions ----
        f3 = ttk.LabelFrame(main, text="▶  3. Actions", padding=10)
        f3.pack(fill=tk.X, pady=5, padx=10)

        btn_row = ttk.Frame(f3)
        btn_row.pack(fill=tk.X)
        # SOLVE butonu başlangıçta DISABLED - sadece load_properties() başarıyla
        # çalışıp properties_loaded=True olunca NORMAL'e çevrilir (bkz. load_properties).
        self.btn_start = ttk.Button(btn_row, text="▶  SOLVE BASE MODEL", command=self.start_solve,
                                     state=tk.DISABLED, style='Accent.TButton')
        self.btn_start.pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row, text="Clear Log", command=self.clear_log).pack(side=tk.LEFT, padx=5)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(f3, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=5)

        # ---- Section 4: Log ----
        f5 = ttk.LabelFrame(main, text="\U0001F4DC  4. Log", padding=10)
        f5.pack(fill=tk.BOTH, expand=True, pady=5, padx=10)

        self.log_text = scrolledtext.ScrolledText(f5, height=20, width=100, font=('Consolas', 10),
                                                    bg=_APP_COLORS['log_bg'], fg=_APP_COLORS['log_fg'],
                                                    insertbackground=_APP_COLORS['log_fg'], relief='flat',
                                                    borderwidth=0)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    # _add_file_row: "Label + Entry + Browse + Load" düzenindeki tekrarlayan
    # satırları (Property Excel, Offset Element IDs) tek yerden üretmek
    # için küçük bir yardımcı. var: ilgili StringVar; load_cmd: "Load"
    # butonuna basınca çağrılacak fonksiyon.
    def _add_file_row(self, parent, label, var, load_cmd):
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=2)
        ttk.Label(frame, text=label, width=18).pack(side=tk.LEFT)
        ttk.Entry(frame, textvariable=var, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame, text="Browse", command=lambda: var.set(
            filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")])
        )).pack(side=tk.LEFT)
        ttk.Button(frame, text="Load", command=load_cmd).pack(side=tk.LEFT, padx=2)

    # log: arka plan thread'inden güvenli şekilde log paneline mesaj eklemek
    # için - gerçek widget güncellemesi root.after(0, ...) ile ana (GUI) thread'e devredilir.
    def log(self, msg):
        def _do():
            self.log_text.insert(tk.END, msg + "\n")
            self.log_text.see(tk.END)
        self.root.after(0, _do)

    def clear_log(self):
        self.log_text.delete(1.0, tk.END)

    # ==================== BDF FILE OPS ====================

    # _add_maneuver_bdf / _remove_maneuver_bdf: "Maneuver BDF (Offset Source)"
    # listesine dosya ekleme/çıkarma. self.maneuver_bdfs listesi ve
    # maneuver_listbox widget'ı her zaman aynı index sırasında senkron tutulur.
    def _add_maneuver_bdf(self):
        files = filedialog.askopenfilenames(filetypes=[("BDF", "*.bdf *.dat *.nas"), ("All", "*.*")])
        for f in files:
            if f not in self.maneuver_bdfs:
                self.maneuver_bdfs.append(f)
                self.maneuver_listbox.insert(tk.END, os.path.basename(f))

    def _remove_maneuver_bdf(self):
        sel = self.maneuver_listbox.curselection()
        if sel:
            idx = sel[0]
            self.maneuver_listbox.delete(idx)
            del self.maneuver_bdfs[idx]

    # add_bdf / remove_bdf: "All BDF Files" listesine dosya ekleme/çıkarma -
    # yukarıdaki maneuver listesiyle aynı mantık, fakat self.bdf_paths/
    # bdf_listbox çifti üzerinde çalışır ve durum etiketini günceller.
    def add_bdf(self):
        paths = filedialog.askopenfilenames(filetypes=[("BDF", "*.bdf *.dat *.nas"), ("All", "*.*")])
        for p in paths:
            if p not in self.bdf_paths:
                self.bdf_paths.append(p)
                self.bdf_listbox.insert(tk.END, os.path.basename(p))
        self.bdf_status.config(text=f"{len(self.bdf_paths)} BDF files selected", foreground="blue")

    def remove_bdf(self):
        selection = self.bdf_listbox.curselection()
        if selection:
            idx = selection[0]
            self.bdf_listbox.delete(idx)
            del self.bdf_paths[idx]
        self.bdf_status.config(text=f"{len(self.bdf_paths)} BDF files selected", foreground="blue")

    # ==================== BDF LOADING ====================
    # load_bdf: "All BDF Files" listesindeki her dosyayı pyNastran BDF() ile
    # okur (xref=False - cross-reference adımı atlanır, çünkü aşağıda sadece
    # ham alanlar (.pid, .type, .dim, .t) okunuyor, *_ref nesneleri veya
    # .Area()/.Centroid() gibi geometri metodları hiç kullanılmıyor), PBARL
    # dim1/dim2 ve PSHELL kalınlıklarını çıkarır ve "orijinal BDF değeri"
    # olarak saklar (Excel'den gelen değerler bunların üzerine yazılacaktır).
    def load_bdf(self):
        if not self.bdf_paths:
            messagebox.showerror("Error", "Add at least one BDF file")
            return

        self.log("\n" + "=" * 70)
        self.log("LOADING BDF FILES")
        self.log("=" * 70)
        self.log(f"  BDF files to load: {len(self.bdf_paths)}")
        for i, p in enumerate(self.bdf_paths):
            self.log(f"    {i + 1}. {os.path.basename(p)}")

        t_load_start = time.time()
        try:
            path = self.bdf_paths[0]  # ilk dosya "ana" BDF olarak kabul edilir (self.bdf_model bu olacak)

            # Parse every file in All BDF Files exactly once. self.bdf_model
            # below is just an alias for the first one - it used to be parsed
            # a SECOND time here on its own before this same loop parsed it
            # again as bdf_models[0], doubling the (already slow, pure-Python
            # card-by-card) parse time for that file.
            #
            # xref=False: everything read from these models below (elem.pid,
            # elem.type, prop.dim, prop.t) is a raw field pyNastran fills in
            # at parse time regardless of xref - xref only additionally
            # resolves *_ref object references (pid_ref, mid_ref, ...) and
            # geometry methods like .Area()/.Centroid(), none of which this
            # tool calls anymore. Skipping that whole extra cross-reference
            # pass cuts a real chunk off every BDF load.
            # Her BDF dosyası için: boş bir BDF() nesnesi oluştur, dosyayı oku
            # (validate=False: pyNastran'ın kart doğrulamasını atla - zaten
            # bu dosyalar gerçek Nastran tarafından önceden doğrulandı;
            # read_includes=False: INCLUDE kartlarının içeriğini bu aşamada
            # genişletme - burada sadece ana dosyanın property/eleman
            # bilgisi gerekiyor), ve sonucu bdf_models listesine ekle.
            self.bdf_models = []
            for bdf_path in self.bdf_paths:
                self.log(f"\n  Loading: {os.path.basename(bdf_path)}")
                t_file_start = time.time()
                model = BDF(debug=False)
                model.read_bdf(bdf_path, validate=False, xref=False, read_includes=False, encoding='latin-1')
                self.log(f"    ...parsed in {time.time() - t_file_start:.1f}s "
                          f"({len(model.nodes)} nodes, {len(model.elements)} elements)")
                self.bdf_models.append({'path': bdf_path, 'model': model, 'name': os.path.basename(bdf_path)})

            # self.bdf_model: geriye dönük uyumluluk için - ilk dosyanın
            # modeline kısayol (alias), ayrı bir parse değil.
            self.bdf_model = self.bdf_models[0]['model']

            self.log(f"\n  Main BDF: {os.path.basename(path)}")
            self.log(f"  Nodes: {len(self.bdf_model.nodes)}")
            self.log(f"  Elements: {len(self.bdf_model.elements)}")
            self.log(f"  Properties: {len(self.bdf_model.properties)}")
            self.log(f"  Materials: {len(self.bdf_model.materials)}")

            # Just a shell/bar count for the log - nothing here (element ->
            # property mapping, material densities, Centroid()/Area()/
            # Length()) is read by anything downstream anymore now that
            # there's no stress extraction/reporting left, only Nastran runs.
            shell_count = bar_count = 0
            for elem in self.bdf_model.elements.values():
                if elem.type in ['CQUAD4', 'CTRIA3', 'CQUAD8', 'CTRIA6']:
                    shell_count += 1
                elif elem.type in ['CBAR', 'CBEAM']:
                    bar_count += 1

            # PBARL kartlarından dim1/dim2 boyutlarını çıkar. dims[0]=dim1,
            # dims[1]=dim2; sadece bir boyut varsa ikisi de ona eşitlenir.
            self.pbarl_dims = {}
            for pid, prop in self.bdf_model.properties.items():
                if prop.type == 'PBARL':
                    dims = prop.dim if hasattr(prop, 'dim') else []
                    if len(dims) >= 2:
                        self.pbarl_dims[pid] = {'dim1': dims[0], 'dim2': dims[1]}
                    elif len(dims) == 1:
                        self.pbarl_dims[pid] = {'dim1': dims[0], 'dim2': dims[0]}

            # BDF'teki orijinal dim1 değerlerini "original_bar_thicknesses"
            # olarak sakla - bunlar Excel'de bu PID için değer YOKSA
            # kullanılacak referans değerlerdir.
            self.original_bar_thicknesses = {}
            if self.pbarl_dims:
                self.log(f"  PBARL dimensions extracted: {len(self.pbarl_dims)} properties")
                for pid, d in self.pbarl_dims.items():
                    self.original_bar_thicknesses[pid] = d['dim1']
                    self.log(f"    PID {pid}: dim1={d['dim1']}, dim2={d['dim2']}")

            # Extract PSHELL thicknesses from BDF (fallback for offset calc when
            # a PID isn't in the Excel 'Skin Property' sheet). Stored separately
            # from self.current_skin_thicknesses (the effective value actually
            # used) and merged via _refresh_effective_skin_thicknesses, so
            # whichever of load_bdf()/load_properties() the user happens to run
            # second never silently discards the other's data - there is no
            # "Load BDF" button, so this runs lazily from start_solve() and can
            # easily land AFTER load_properties() in the normal click order.
            self.original_skin_thicknesses = {}
            for pid, prop in self.bdf_model.properties.items():
                if prop.type == 'PSHELL':
                    t = prop.t if hasattr(prop, 't') and prop.t is not None else None
                    if t is not None:
                        self.original_skin_thicknesses[pid] = t
            if self.original_skin_thicknesses:
                self.log(f"  PSHELL thicknesses extracted: {len(self.original_skin_thicknesses)} properties (fallback for offset calc)")
            # Excel değerleri varsa onları, yoksa BDF'ten okunan bu
            # orijinal PSHELL kalınlıklarını "etkin" (current_skin_thicknesses)
            # değer olarak birleştir.
            self._refresh_effective_skin_thicknesses()

            self.log(f"  Shells: {shell_count}, Bars: {bar_count}")
            self.log(f"\n  Total BDF models loaded: {len(self.bdf_models)} "
                      f"(took {time.time() - t_load_start:.1f}s total)")

            self.root.after(0, lambda: self.bdf_status.config(
                text=f"Loaded: {len(self.bdf_models)} BDFs, {len(self.bdf_model.elements)} elements",
                foreground="green"
            ))

            # Çıktı klasörü henüz boşsa, ilk BDF'in bulunduğu klasörü varsayılan olarak öner.
            if not self.output_folder.get():
                out_dir = os.path.dirname(path)
                self.root.after(0, lambda: self.output_folder.set(out_dir))

        except Exception as e:
            self.log(f"ERROR after {time.time() - t_load_start:.1f}s: {e}")
            import traceback
            self.log(traceback.format_exc())
            self.root.after(0, lambda: self.bdf_status.config(text="Error", foreground="red"))

    # _load_master_input: Master Input Excel satırının tek "Load" butonuna
    # bağlı arayüz fonksiyonu - property_excel_path ve element_excel_path
    # zaten AYNI StringVar olduğu için, burada sadece load_properties() ve
    # load_element_ids()'i SIRAYLA çağırmak yeterli; o iki metodun kendi
    # içeriği hiç değişmedi.
    def _load_master_input(self):
        self.load_properties()
        self.load_element_ids()

    # ==================== PROPERTY LOADING ====================
    def load_properties(self):
        """Load the Excel-defined bar Dim1/Dim2 ('Bar Property' sheet) and skin
        thickness ('Skin Property' sheet). These are the values the base model
        gets written with (see _build_base_model); any PID without an entry in
        either sheet keeps its original BDF value untouched."""
        path = self.property_excel_path.get()
        if not path:
            messagebox.showerror("Error", "Select Property Excel")
            return

        self.log("\n" + "=" * 70)
        self.log("LOADING PROPERTIES")
        self.log("=" * 70)

        t_start = time.time()
        try:
            xl = pd.ExcelFile(path)
            self.log(f"  Sheets: {xl.sheet_names}")  # kullanıcıya hangi sayfaların bulunduğunu göster (teşhis amaçlı)

            # ---- Locate the Bar/Skin Property sheets. No structure-grouping
            # sheet is read anymore - there is no sweep to group by, only a
            # single base-model solve, so every PID that has an entry in
            # 'Bar Property'/'Skin Property' gets that value applied; any PID
            # without one keeps its original BDF value untouched. ----
            bar_dim_sheet = skin_sheet = None
            # 1. geçiş: sayfa adını normalize et (alt çizgi/boşluk/tire sil,
            # küçük harfe çevir) ve TAM eşleşme ara ("Bar_Property", "bar property" vb. hepsi "barproperty" olur).
            for sheet in xl.sheet_names:
                sl = sheet.lower().replace('_', '').replace(' ', '').replace('-', '')
                if sl == 'barproperty':
                    bar_dim_sheet = sheet
                elif sl == 'skinproperty':
                    skin_sheet = sheet
            # 2. geçiş: tam eşleşme bulunamadıysa, normalize edilmiş adda
            # "bar"+"prop" veya "skin"+"prop" alt dizisi arayan daha gevşek eşleşme.
            for sheet in xl.sheet_names:
                sl = sheet.lower().replace('_', '').replace(' ', '').replace('-', '')
                if bar_dim_sheet is None and 'bar' in sl and 'prop' in sl:
                    bar_dim_sheet = sheet
                elif skin_sheet is None and 'skin' in sl and 'prop' in sl:
                    skin_sheet = sheet

            # ---- 1. Excel bar Dim1/Dim2 ('Bar Property' sheet: PID, Dim1, Dim2) ----
            self.bar_properties = {}
            self.current_bar_thicknesses = {}
            if bar_dim_sheet:
                self.log(f"\n  Reading bar Dim1/Dim2 from '{bar_dim_sheet}'...")
                df_bar = pd.read_excel(xl, sheet_name=bar_dim_sheet)
                # Sütunlar sırasıyla: PID, Dim1, (opsiyonel) Dim2. Dim2 yoksa Dim1 ile aynı kabul edilir.
                for _, row in df_bar.iterrows():
                    try:
                        pid = int(row.iloc[0])
                        d1 = float(row.iloc[1])
                        d2 = float(row.iloc[2]) if len(df_bar.columns) > 2 else d1
                        self.bar_properties[pid] = {'dim1': d1, 'dim2': d2}
                        self.current_bar_thicknesses[pid] = d1
                    except Exception:
                        pass  # sayıya çevrilemeyen/eksik satırı sessizce atla
                self.log(f"    Bar dimensions loaded: {len(self.bar_properties)} properties")
            else:
                self.log("  No 'Bar Property' sheet found - bar dimensions will be left as-is (BDF values).")

            # ---- 2. Excel skin thickness ('Skin Property' sheet: PID, Thickness) ----
            self.skin_properties = {}
            if skin_sheet:
                self.log(f"\n  Reading skin thickness from '{skin_sheet}'...")
                df_skin = pd.read_excel(xl, sheet_name=skin_sheet)
                for _, row in df_skin.iterrows():
                    try:
                        pid = int(row.iloc[0])
                        t = float(row.iloc[1])
                        self.skin_properties[pid] = {'thickness': t}
                    except Exception:
                        pass
                self.log(f"    Skin thicknesses loaded: {len(self.skin_properties)} properties")
            else:
                self.log("  No 'Skin Property' sheet found - skin thickness will be left as-is (BDF values).")
            # Etkin kalınlığı (önce Excel, yoksa BDF) şimdi yeniden hesapla -
            # load_bdf() bundan önce çalışmış olsun ya da olmasın bu çağrı güvenlidir.
            self._refresh_effective_skin_thicknesses()

            total = len(self.bar_properties)
            self.log(f"\n  Properties loaded in {time.time() - t_start:.1f}s")
            self.prop_status.config(
                text=f"Loaded: {total} bar props, {len(self.skin_properties)} skin props",
                foreground="green"
            )
            self.properties_loaded = True
            self.btn_start.config(state=tk.NORMAL)  # artık SOLVE butonu tıklanabilir

        except Exception as e:
            self.log(f"ERROR after {time.time() - t_start:.1f}s: {e}")
            import traceback
            self.log(traceback.format_exc())
            self.prop_status.config(text="Error", foreground="red")
            self.properties_loaded = False
            self.btn_start.config(state=tk.DISABLED)

    def _refresh_effective_skin_thicknesses(self):
        """Recompute self.current_skin_thicknesses - the effective value used
        for offset calc - as: Excel 'Skin Property' sheet value if given,
        else the BDF's own PSHELL thickness (self.original_skin_thicknesses).
        There is no dedicated "Load BDF" button - load_bdf() only runs lazily
        from start_solve() - so load_bdf() and load_properties() can each run
        in either order relative to the other. Calling this from both after
        they update their own half of the data means neither one's data ever
        gets silently discarded by the other running afterward."""
        merged = dict(self.original_skin_thicknesses)
        for pid, d in self.skin_properties.items():
            merged[pid] = d['thickness']
        self.current_skin_thicknesses = merged

    # ==================== ELEMENT IDs (OFFSET) ====================
    # load_element_ids: "Offset Element IDs" Excel'ini oku. Bu Excel'de iki
    # tür sayfa olabilir - adında "landing" geçen (göçme/landing elemanları)
    # ve adında hem "bar" hem "offset" geçen (bar offset elemanları). Her
    # sayfanın ilk sütunu eleman ID listesi olarak okunur.
    def load_element_ids(self):
        path = self.element_excel_path.get()
        if not path:
            return

        self.log("\n" + "=" * 70)
        self.log("LOADING ELEMENT IDs FOR OFFSET")
        self.log("=" * 70)

        try:
            xl = pd.ExcelFile(path)
            self.landing_elem_ids = []
            self.bar_offset_elem_ids = []

            for s in xl.sheet_names:
                sl = s.lower().replace('_', '').replace(' ', '')
                df = pd.read_excel(xl, sheet_name=s)
                if 'landing' in sl:
                    self.landing_elem_ids = df.iloc[:, 0].dropna().astype(int).tolist()
                    self.log(f"  Landing: {len(self.landing_elem_ids)}")
                elif 'bar' in sl and 'offset' in sl:
                    self.bar_offset_elem_ids = df.iloc[:, 0].dropna().astype(int).tolist()
                    self.log(f"  Bar offset: {len(self.bar_offset_elem_ids)}")

            self.elem_status.config(
                text=f"Landing: {len(self.landing_elem_ids)}, Bar offset: {len(self.bar_offset_elem_ids)}",
                foreground="green"
            )

        except Exception as e:
            self.log(f"ERROR: {e}")
            self.elem_status.config(text="Error", foreground="red")

    # ==================== BDF WRITING ====================
    def _write_bdf_for_model(self, folder, model, original_path, thickness_overrides):
        """thickness_overrides: {pid: new_dim1}. Only these PBARL PIDs are
        rewritten; every other card in original_path (the base model) is left
        untouched. Taking the overrides as an explicit argument - instead of
        reading a shared self.current_bar_thicknesses/active_group_pids state -
        is what makes this safe to call from multiple threads at once."""
        output_bdf = os.path.join(folder, os.path.basename(original_path))

        with open(original_path, 'r', encoding='latin-1') as f:
            lines = f.readlines()

        new_lines = []
        i = 0
        # Dosyayı satır satır tara. Her "PBARL" ana kartı için PID'i sabit
        # genişlikli (8-karakter) alan olarak oku (line[8:16]); bu PID
        # thickness_overrides içindeyse kartın DEVAM satırını (continuation -
        # "+" veya "*" ile başlayan ya da boşlukla başlayıp içerik taşıyan
        # satır) bulup dim1/dim2 alanlarını yeni kalınlıkla değiştir.
        while i < len(lines):
            line = lines[i]

            if line.startswith('PBARL'):
                try:
                    pid = int(line[8:16].strip())
                    if pid in thickness_overrides:
                        t = max(0.1, thickness_overrides[pid])  # kalınlık 0'ın altına/sıfıra düşmesin
                        new_lines.append(line)  # PBARL ana satırı değişmeden kalır (PID/tip burada)
                        i += 1
                        while i < len(lines) and (
                            lines[i].startswith('+') or lines[i].startswith('*') or
                            (lines[i][0] == ' ' and lines[i].strip() and not lines[i].strip().startswith('$'))
                        ):
                            cont = lines[i]
                            if cont.strip() and not cont.strip().startswith('$'):
                                try:
                                    # dim2'yi devam satırından oku (16:24) - eksikse yeni dim1 (t) ile aynı kabul et.
                                    original_dim2 = cont[16:24].strip()
                                    dim2_val = float(original_dim2) if original_dim2 else t
                                except:
                                    dim2_val = t
                                cont_name = cont[:8]  # devam satırının ilk 8 karakteri ("+" süreklilik işareti vb.) korunur
                                rest = cont[24:] if len(cont) > 24 else '\n'  # dim2'den sonraki alanlar (varsa) değişmeden kalır
                                new_cont = f"{cont_name}{t:8.4f}{dim2_val:8.4f}{rest}"
                                new_lines.append(new_cont)
                            else:
                                new_lines.append(cont)
                            i += 1
                        continue
                except:
                    pass

            new_lines.append(line)
            i += 1

        with open(output_bdf, 'w', encoding='latin-1') as f:
            f.writelines(new_lines)

        self.log(f"    BDF written: {os.path.basename(output_bdf)}")
        return output_bdf

    # ==================== BASE MODEL (Excel Bar/Skin properties) ====================
    def _apply_excel_properties(self, source_path, folder):
        """Write a copy of source_path with ALL bar Dim1/Dim2 ('Bar Property' sheet)
        and ALL skin thicknesses ('Skin Property' sheet) applied - i.e. every PID,
        not just the active sweep group. This is the base model that the sweep's
        per-group thickness changes are layered on top of.

        Named "..._prop_updated" (rather than keeping source_path's own name)
        so it's never confused with the later, actually-offset file - a
        source file already named "..._offseted.bdf" (e.g. carried over from
        an earlier processing stage) would otherwise make this property-only
        step look like it had already applied offsets, and the real offset
        step's own "_offseted" suffix on top of that produced a misleading
        "..._offseted_offseted" name. Any existing trailing "_offseted" on
        the source's own name is stripped first, so e.g.
        "merged_maneuver_offseted.bdf" becomes "merged_maneuver_prop_updated.bdf"
        here and "merged_maneuver_prop_updated_offseted.bdf" once actually
        offset - not "merged_maneuver_offseted_prop_updated...", which just
        carries the source's stale suffix along for no reason."""
        base, ext = os.path.splitext(os.path.basename(source_path))
        if base.endswith('_offseted'):
            base = base[:-len('_offseted')]
        output_bdf = os.path.join(folder, base + "_prop_updated" + ext)

        with open(source_path, 'r', encoding='latin-1') as f:
            lines = f.readlines()

        new_lines = []
        i = 0
        bar_updated = 0
        skin_updated = 0
        # _write_bdf_for_model'deki PBARL mantığıyla aynı sabit-genişlik
        # alan okuma/yazma yaklaşımı, ama burada TÜM self.bar_properties
        # PID'leri (sweep grubu kısıtlaması yok) ve ayrıca PSHELL kalınlığı
        # da işleniyor.
        while i < len(lines):
            line = lines[i]

            if line.startswith('PBARL'):
                try:
                    pid = int(line[8:16].strip())
                    if pid in self.bar_properties:
                        d1 = self.bar_properties[pid]['dim1']
                        d2 = self.bar_properties[pid]['dim2']
                        new_lines.append(line)
                        i += 1
                        if i < len(lines) and (
                            lines[i].startswith('+') or lines[i].startswith('*') or
                            (lines[i][0] == ' ' and lines[i].strip() and not lines[i].strip().startswith('$'))
                        ):
                            cont = lines[i]
                            cont_name = cont[:8]
                            rest = cont[24:] if len(cont) > 24 else '\n'
                            new_lines.append(f"{cont_name}{d1:8.4f}{d2:8.4f}{rest}")
                            i += 1
                        bar_updated += 1
                        continue
                except:
                    pass

            elif line.startswith('PSHELL'):
                try:
                    # PSHELL kart formatı: PID alanı [8:16], ardından MID1 [16:24],
                    # kalınlık (t) alanı [24:32] - bu yüzden [:24] + yeni t + [32:] sonrası.
                    pid = int(line[8:16].strip())
                    if pid in self.skin_properties:
                        t = self.skin_properties[pid]['thickness']
                        rest = line[32:] if len(line) > 32 else '\n'
                        new_lines.append(line[:24] + f"{t:8.4f}" + rest)
                        skin_updated += 1
                        i += 1
                        continue
                except:
                    pass

            new_lines.append(line)
            i += 1

        with open(output_bdf, 'w', encoding='latin-1') as f:
            f.writelines(new_lines)

        self.log(f"    {os.path.basename(source_path)}: {bar_updated} bar props, "
                  f"{skin_updated} skin props set to Excel base values")
        return output_bdf

    def _build_base_model(self, run_folder):
        """Apply the Excel base Bar/Skin properties to every loaded BDF and to
        the maneuver (offset source) BDF. These base files carry ONLY the
        Excel properties - no offsets. Offsets are computed and applied fresh
        for every iteration (baseline included) by _prepare_iteration_bdfs,
        since they depend on the maneuver BDF's geometry at that iteration's
        thickness, not on some value frozen once up front."""
        base_folder = os.path.join(run_folder, "_base_model")
        os.makedirs(base_folder, exist_ok=True)

        self.log("\n" + "=" * 70)
        self.log("BUILDING BASE MODEL (Excel Bar/Skin Property thicknesses)")
        self.log("=" * 70)

        t_build_start = time.time()
        t_props_start = time.time()

        # Maneuver (offset kaynağı) BDF'ine de Excel Bar/Skin değerlerini uygula.
        # Listede birden fazla maneuver BDF olsa da sadece İLKİ (maneuver_bdfs[0])
        # offset kaynağı olarak kullanılır.
        self.maneuver_base_path = None
        if self.maneuver_bdfs:
            man_folder = os.path.join(base_folder, "maneuver")
            os.makedirs(man_folder, exist_ok=True)
            self.maneuver_base_path = self._apply_excel_properties(self.maneuver_bdfs[0], man_folder)
            self.log(f"  Base maneuver (offset source) BDF ready: {os.path.basename(self.maneuver_base_path)}")
        else:
            self.log("  No maneuver BDF provided - no offsets will be computed.")

        # "All BDF Files" listesindeki HER BDF için kendi alt klasörüne
        # Excel özelliklerini uygulanmış bir kopya yaz; sonucu bdf_info
        # sözlüğüne 'base_path' olarak geri kaydet (sonraki adımlar bunu kullanacak).
        for bdf_idx, bdf_info in enumerate(self.bdf_models):
            bdf_folder = os.path.join(base_folder, f"bdf_{bdf_idx + 1}_{os.path.splitext(bdf_info['name'])[0]}")
            os.makedirs(bdf_folder, exist_ok=True)
            base_path = self._apply_excel_properties(bdf_info['path'], bdf_folder)
            bdf_info['base_path'] = base_path
            self.log(f"  Base model ready for {bdf_info['name']}: {os.path.basename(base_path)}")

        self.log(f"  Properties applied in {time.time() - t_props_start:.1f}s")

        # Base model üzerinde Nastran'ı çalıştır. thickness_overrides={} demek,
        # her PID'in self.bar_properties'teki temel (Excel) değerini kullanacağı
        # anlamına gelir - offsetler bu aynı temel kalınlıktaki maneuver BDF'inden
        # yeniden hesaplanır.
        self.log("\n  Running Nastran on the BASE MODEL...")
        t_nastran_start = time.time()
        n_ok, n_total = self._run_single_iteration(base_folder, {}, label="[BASE] ")
        nastran_elapsed = time.time() - t_nastran_start
        if n_total == 0:
            self.log(f"  WARNING: no Nastran runs were produced (check Nastran path/BDF files). "
                      f"({nastran_elapsed:.1f}s)")
        elif n_ok == n_total:
            self.log(f"  All {n_total} Nastran run(s) completed successfully in {nastran_elapsed:.1f}s.")
        else:
            self.log(f"  WARNING: {n_total - n_ok}/{n_total} Nastran run(s) failed after "
                      f"{nastran_elapsed:.1f}s - check the log above.")
        self.log(f"  Base model results saved under: {base_folder}")

        self.log(f"BASE MODEL SOLVE COMPLETE (total {time.time() - t_build_start:.1f}s).\n")

    # ==================== OFFSET CALCULATION / APPLICATION ====================
    def _calculate_offset_csv(self, thickness_overrides, folder, label=""):
        """Write the maneuver (offset source) BDF at THIS iteration's
        thickness_overrides (layered on its Excel-properties base, same as any
        other BDF), then compute landing/bar offsets from that freshly-written
        geometry and save them to offsets.csv in `folder`. Called once per
        iteration (sweep step or baseline) - not once for the whole sweep -
        since the maneuver geometry (and therefore the offsets) changes with
        whichever group's thickness is being tested. Returns the CSV path, or
        None if there's no maneuver BDF or no offset elements configured."""
        if not self.maneuver_base_path:
            return None
        if not self.landing_elem_ids and not self.bar_offset_elem_ids:
            return None

        t_offset_start = time.time()
        man_folder = os.path.join(folder, "maneuver_for_offset")
        os.makedirs(man_folder, exist_ok=True)
        # Bu iterasyonun thickness_overrides'ını maneuver BDF'ine uygulayarak,
        # offset hesabının HER ZAMAN o anki (iterasyona özgü) geometriyi
        # kullanmasını sağla - geometri (dolayısıyla offsetler) kalınlıkla değişir.
        maneuver_path = self._write_bdf_for_model(man_folder, None, self.maneuver_base_path, thickness_overrides)

        try:
            t_parse_start = time.time()
            bdf = BDF(debug=False)
            # xref=True burada GEREKLİ - aşağıda node.get_position() ve
            # elem.node_ids üzerinden gerçek 3B koordinatlar okunacak.
            bdf.read_bdf(maneuver_path, validate=False, xref=True, read_includes=False, encoding='latin-1')
            self.log(f"    {label}Maneuver BDF parsed for offsets in {time.time() - t_parse_start:.1f}s")
        except Exception as e:
            self.log(f"    {label}Offset calc error reading maneuver BDF: {e}")
            return None

        try:
            # ---- Landing (shell) offsetleri: zoffset = -t/2 ----
            # "Landing" (göçme) elemanı, kalınlığının yarısı kadar kendi yüzey
            # normali yönünün TERSİNE (-normal) kaydırılır - bu klasik bir
            # "shell offset to the surface" yaklaşımıdır (shell orta düzlemi
            # yerine gerçek yüzeyde modellensin diye).
            landing_normals = {}
            rows = []

            for eid in self.landing_elem_ids:
                if eid not in bdf.elements:
                    continue
                elem = bdf.elements[eid]
                pid = elem.pid if hasattr(elem, 'pid') else None
                if pid is None:
                    continue

                # Bu PID'in etkin skin kalınlığını al (Excel öncelikli, yoksa BDF); yoksa 3.0 varsayılan.
                t = self.current_skin_thicknesses.get(pid, 3.0)
                t = max(0.1, t)
                zoffset = -t / 2.0

                # Üçgen/quad'ın ilk 3 düğümünden iki kenar vektörü oluşturup
                # çapraz çarpımla (cross product) birim yüzey normalini hesapla.
                normal = None
                if elem.type in ['CQUAD4', 'CTRIA3']:
                    try:
                        nids = elem.node_ids[:3]
                        nodes = [bdf.nodes[n] for n in nids if n in bdf.nodes]
                        if len(nodes) >= 3:
                            p1, p2, p3 = [np.array(n.get_position()) for n in nodes]
                            nvec = np.cross(p2 - p1, p3 - p1)
                            nl = np.linalg.norm(nvec)
                            if nl > 1e-10:
                                normal = nvec / nl
                    except Exception:
                        pass
                landing_normals[eid] = normal  # bar offsetleri için sonra tekrar kullanılacak
                rows.append({'Type': 'Landing', 'Element': eid, 'ZOffset': zoffset,
                             'WA_X': '', 'WA_Y': '', 'WA_Z': ''})

            # Her düğümün hangi shell (CQUAD4/CTRIA3) elemanlarına ait olduğunun
            # haritasını çıkar - bir bar elemanının "hangi landing shell'in
            # üzerinde oturduğunu" bulmak için kullanılacak.
            node_to_shells = {}
            for eid, elem in bdf.elements.items():
                if elem.type in ['CQUAD4', 'CTRIA3']:
                    for nid in elem.node_ids:
                        node_to_shells.setdefault(nid, []).append(eid)

            # ---- Bar offsetleri: WA = WB = -normal * (landing_t + bar_t/2) ----
            # Bir bar elemanı, iki ucunun oturduğu landing shell'in üstünde
            # duruyor kabul edilir; bu yüzden offset vektörü, o shell'in
            # normali yönünün tersine, (shell kalınlığı + bar kalınlığının
            # yarısı) kadar kaydırılır.
            bar_min = self.default_bar_thickness
            for eid in self.bar_offset_elem_ids:
                if eid not in bdf.elements:
                    continue
                elem = bdf.elements[eid]
                if elem.type not in ['CBAR', 'CBEAM']:
                    continue
                pid = elem.pid if hasattr(elem, 'pid') else None
                if pid is None:
                    continue

                # Bu iterasyonda uygulanan override varsa onu, yoksa Excel
                # temel değerini, o da yoksa varsayılan bar kalınlığını kullan.
                bar_t = thickness_overrides.get(pid, self.bar_properties.get(pid, {}).get('dim1', bar_min))
                bar_t = max(0.1, bar_t)

                # Bar'ın iki ucunun (node_ids[:2]) ortak oturduğu shell
                # eleman(lar)ını bul (iki düğümün de ait olduğu shell kümelerinin kesişimi).
                bar_nodes = elem.node_ids[:2]
                if bar_nodes[0] in node_to_shells and bar_nodes[1] in node_to_shells:
                    common = set(node_to_shells[bar_nodes[0]]) & set(node_to_shells[bar_nodes[1]])
                    max_t = 0
                    best_normal = None
                    # Ortak shell'ler arasından EN KALIN olanı seç (en kalın
                    # landing kalınlığı, bar'ın gerçekte hangi yüzeyde
                    # oturduğunu en iyi temsil eden varsayım).
                    for shell_eid in common:
                        if shell_eid in landing_normals:
                            shell_elem = bdf.elements[shell_eid]
                            shell_pid = shell_elem.pid if hasattr(shell_elem, 'pid') else None
                            if shell_pid:
                                shell_t = self.current_skin_thicknesses.get(shell_pid, 0)
                                if shell_t > max_t:
                                    max_t = shell_t
                                    best_normal = landing_normals.get(shell_eid)

                    if best_normal is not None and max_t > 0:
                        offset_mag = max_t + bar_t / 2.0
                        vec = -best_normal * offset_mag
                        rows.append({'Type': 'Bar', 'Element': eid, 'ZOffset': '',
                                     'WA_X': vec[0], 'WA_Y': vec[1], 'WA_Z': vec[2]})

            if not rows:
                return None

            # Hesaplanan tüm offsetleri (landing + bar) offsets.csv'ye yaz.
            n_landing = sum(1 for r in rows if r['Type'] == 'Landing')
            n_bar = sum(1 for r in rows if r['Type'] == 'Bar')
            csv_path = os.path.join(folder, 'offsets.csv')
            with open(csv_path, 'w', newline='') as f:
                w = csv.DictWriter(f, fieldnames=['Type', 'Element', 'ZOffset', 'WA_X', 'WA_Y', 'WA_Z'])
                w.writeheader()
                w.writerows(rows)
            self.log(f"    {label}Offset CSV: {n_landing} landing, {n_bar} bar -> {csv_path} "
                      f"(total {time.time() - t_offset_start:.1f}s)")
            return csv_path

        except Exception as e:
            self.log(f"    {label}Offset calc error: {e}")
            return None

    def _apply_offsets_from_csv(self, bdf_path, csv_path, folder, label=""):
        """Read an offsets.csv (from _calculate_offset_csv) and write a copy of
        bdf_path with those ZOFFSET/WA-WB values applied, saved as
        <name>_offseted.bdf in `folder`."""
        landing_offsets = {}
        bar_offsets = {}
        # offsets.csv'yi geri oku: Landing satırları {eid: zoffset}, Bar
        # satırları {eid: (WA_X, WA_Y, WA_Z)} sözlüklerine dönüştürülür.
        try:
            with open(csv_path, 'r', newline='') as f:
                for row in csv.DictReader(f):
                    eid = int(row['Element'])
                    if row['Type'] == 'Landing':
                        landing_offsets[eid] = float(row['ZOffset'])
                    elif row['Type'] == 'Bar':
                        bar_offsets[eid] = (float(row['WA_X']), float(row['WA_Y']), float(row['WA_Z']))
        except Exception as e:
            self.log(f"    {label}Could not read offsets CSV: {e}")
            return bdf_path

        try:
            with open(bdf_path, 'r', encoding='latin-1') as f:
                lines = f.readlines()

            # fmt: bir sayıyı Nastran'ın 8 karakterlik sabit alan genişliğine
            # sığdır - normal ondalık gösterim (.4f) 8 karaktere sığmazsa
            # (çok büyük/küçük sayı), bilimsel gösterime (.2E) düş.
            def fmt(v, w=8):
                s = f"{v:.4f}"
                return s[:w].ljust(w) if len(s) <= w else f"{v:.2E}"[:w].ljust(w)

            new_lines = []
            i = 0
            applied_landing = 0
            applied_bar = 0

            while i < len(lines):
                line = lines[i]

                # CQUAD4 kartında ZOFFSET alanı 65-72. karakterler (0-index [64:72])
                # - satırı önce 72 karaktere kadar boşlukla doldur, sonra o
                # alana yeni zoffset değerini yaz.
                if line.startswith('CQUAD4'):
                    try:
                        eid = int(line[8:16].strip())
                        if eid in landing_offsets:
                            padded = line.rstrip().ljust(72)
                            new_lines.append(padded[:64] + fmt(landing_offsets[eid]) + '\n')
                            applied_landing += 1
                            i += 1
                            continue
                    except:
                        pass

                # CTRIA3'te ZOFFSET alanı CQUAD4'e göre bir alan daha erken
                # başlar (49-56. karakter, [48:56]) çünkü CTRIA3'ün 4. düğüm
                # alanı (G4) yoktur.
                elif line.startswith('CTRIA3'):
                    try:
                        eid = int(line[8:16].strip())
                        if eid in landing_offsets:
                            padded = line.rstrip().ljust(56)
                            new_lines.append(padded[:48] + fmt(landing_offsets[eid]) + '\n')
                            applied_landing += 1
                            i += 1
                            continue
                    except:
                        pass

                # CBAR'da WA (uç A ofseti) ve WB (uç B ofseti) devam satırında
                # yer alır - bu tool WA=WB=offset_vec uyguluyor (bar'ın her iki
                # ucu da aynı miktarda kaydırılıyor).
                elif line.startswith('CBAR'):
                    try:
                        eid = int(line[8:16].strip())
                        if eid in bar_offsets:
                            offset_vec = bar_offsets[eid]

                            # Kartın zaten bir devam satırı varsa (PA/PB
                            # alanlarını [:24] olarak koruyup) WA/WB alanlarını
                            # (24'ten sonraki 6x8 karakter) üzerine yaz.
                            if i + 1 < len(lines) and (
                                lines[i + 1].startswith('+') or lines[i + 1].startswith('*') or
                                (lines[i + 1][0] == ' ' and lines[i + 1].strip())
                            ):
                                cont_line = lines[i + 1]
                                if len(cont_line) < 24:
                                    cont_line = cont_line.rstrip().ljust(24)
                                new_cont = cont_line[:24]
                                new_cont += fmt(offset_vec[0])
                                new_cont += fmt(offset_vec[1])
                                new_cont += fmt(offset_vec[2])
                                new_cont += fmt(offset_vec[0])
                                new_cont += fmt(offset_vec[1])
                                new_cont += fmt(offset_vec[2])
                                new_cont += '\n'

                                new_lines.append(line)
                                new_lines.append(new_cont)
                                applied_bar += 1
                                i += 2
                                continue
                            else:
                                # Devam satırı yoksa SIFIRDAN bir devam satırı
                                # üret: benzersiz bir süreklilik etiketi
                                # ("+CB" + eleman ID'sinin son 4 hanesi) ile
                                # ana satıra ekle, PA/PB alanlarını boş
                                # (varsayılan pin bağlantısı) bırakıp WA/WB'yi doldur.
                                cont_name = '+CB' + str(eid)[-4:]
                                new_lines.append(line.rstrip() + cont_name + '\n')

                                new_cont = cont_name.ljust(8)
                                new_cont += '        '   # PA
                                new_cont += '        '   # PB
                                new_cont += fmt(offset_vec[0])
                                new_cont += fmt(offset_vec[1])
                                new_cont += fmt(offset_vec[2])
                                new_cont += fmt(offset_vec[0])
                                new_cont += fmt(offset_vec[1])
                                new_cont += fmt(offset_vec[2])
                                new_cont += '\n'
                                new_lines.append(new_cont)

                                applied_bar += 1
                                i += 1
                                continue
                    except:
                        pass

                new_lines.append(line)
                i += 1

            base, ext = os.path.splitext(os.path.basename(bdf_path))
            output_bdf = os.path.join(folder, base + "_offseted" + ext)
            with open(output_bdf, 'w', encoding='latin-1') as f:
                f.writelines(new_lines)

            self.log(f"    {label}Offsets applied from CSV: {applied_landing} landing, {applied_bar} bar")
            return output_bdf

        except Exception as e:
            self.log(f"    {label}Offset apply error: {e}")
            return bdf_path

    def _prepare_iteration_bdfs(self, iter_folder, thickness_overrides, label=""):
        """Write every bdf_model's BDF at thickness_overrides (layered on its
        Excel-properties base), compute a FRESH offsets.csv from the maneuver
        (source) BDF at this same thickness_overrides, and apply those offsets
        to every written BDF. If Cases per Split is set, each prop-updated +
        offseted BDF is then split into run1/run2/... files by SUBCASE count
        (see _split_bdf_by_cases). Returns a list of {'name', 'subfolder',
        'run_paths', 'run_ok'} dicts, one per bdf_model - 'run_paths' has one
        entry unless splitting produced more."""
        n_bdfs = len(self.bdf_models) if self.bdf_models else 1

        # Bu iterasyonun kalınlıklarına göre offsets.csv'yi YENİDEN hesapla
        # (maneuver BDF'inin geometrisi her iterasyonda değişebileceği için).
        offset_csv = self._calculate_offset_csv(thickness_overrides, iter_folder, label=label)

        try:
            cases_per_split = int(float(self.cases_per_split.get())) if self.cases_per_split.get().strip() else 0
        except (ValueError, TypeError):
            cases_per_split = 0

        per_bdf = []
        # Her yüklü BDF modeli için: önce kalınlık override'larını uygula,
        # sonra (varsa) hesaplanan offsetleri uygula, sonra referans amaçlı
        # "structure-only" (INCLUDE'suz) bir kopya yaz, son olarak
        # Cases per Split ayarına göre run1/run2/... dosyalarına böl.
        for bdf_idx, bdf_info in enumerate(self.bdf_models):
            bdf_name = bdf_info['name']
            bdf_subfolder = os.path.join(iter_folder, f"bdf_{bdf_idx + 1}_{os.path.splitext(bdf_name)[0]}") if n_bdfs > 1 else iter_folder
            os.makedirs(bdf_subfolder, exist_ok=True)
            base_source = bdf_info.get('base_path', bdf_info['path'])
            bdf_path = self._write_bdf_for_model(bdf_subfolder, bdf_info['model'], base_source, thickness_overrides)
            if offset_csv:
                bdf_path = self._apply_offsets_from_csv(bdf_path, offset_csv, bdf_subfolder, label=label)
            self._write_structure_only_bdf(bdf_path, bdf_subfolder, label=label)

            run_paths = self._split_bdf_by_cases(bdf_path, cases_per_split, label=label)
            per_bdf.append({'name': bdf_name, 'subfolder': bdf_subfolder, 'run_paths': run_paths, 'run_ok': False})
        return per_bdf

    # ==================== CASE SPLITTING (FOR PARALLEL RUNS) ====================
    def _split_bdf_by_cases(self, bdf_path, cases_per_split, label=""):
        """Split bdf_path's SUBCASEs into separate run1.bdf/run2.bdf/... files
        of up to cases_per_split SUBCASEs each, written next to bdf_path. Every
        run file keeps the EXACT same format: the same header (everything
        before the first SUBCASE - SOL/CEND/TITLE/SET/output-request lines)
        and the exact same BEGIN BULK...ENDDATA section (bulk data + includes)
        as the source BDF - only which SUBCASE blocks are present differs.
        Returns [bdf_path] unchanged if cases_per_split is 0/blank, there are
        no SUBCASEs, or there are already <= cases_per_split of them (nothing
        to split)."""
        if not cases_per_split or cases_per_split <= 0:
            return [bdf_path]

        t_split_start = time.time()
        try:
            with open(bdf_path, 'r', encoding='latin-1') as f:
                lines = f.readlines()
        except Exception as e:
            self.log(f"    {label}Split error reading {os.path.basename(bdf_path)}: {e}")
            return [bdf_path]

        # Dosyadaki her "SUBCASE" satırının index'ini bul - bölme sadece
        # gerçekten cases_per_split'ten FAZLA SUBCASE varsa yapılır.
        subcase_idxs = [i for i, l in enumerate(lines) if l.strip().upper().startswith('SUBCASE')]
        if len(subcase_idxs) <= cases_per_split:
            return [bdf_path]

        bulk_idx = next((i for i, l in enumerate(lines) if l.strip().upper().startswith('BEGIN BULK')), len(lines))

        # header_lines: ilk SUBCASE'den önceki her şey (SOL/CEND/TITLE/SET/
        # çıktı istek kartları) - her run dosyasında DEĞİŞMEDEN tekrarlanır.
        # bulk_lines: BEGIN BULK'tan sonraki her şey - o da değişmeden tekrarlanır.
        header_lines = lines[:subcase_idxs[0]]
        bulk_lines = lines[bulk_idx:]
        block_bounds = subcase_idxs + [bulk_idx]

        base, ext = os.path.splitext(bdf_path)
        n_subcases = len(subcase_idxs)
        n_runs = (n_subcases + cases_per_split - 1) // cases_per_split  # yukarı yuvarlama (ceiling division)
        run_paths = []

        # Her run dosyası: aynı header + o run'a ait SUBCASE blokları + aynı
        # bulk data (ve INCLUDE'lar) - sadece hangi SUBCASE'lerin dahil
        # olduğu run'dan run'a değişir.
        for run_idx in range(n_runs):
            start_sc = run_idx * cases_per_split
            end_sc = min(start_sc + cases_per_split, n_subcases)
            chunk_lines = []
            for sc_i in range(start_sc, end_sc):
                chunk_lines.extend(lines[block_bounds[sc_i]:block_bounds[sc_i + 1]])

            out_path = f"{base}_run{run_idx + 1}{ext}"
            with open(out_path, 'w', encoding='latin-1') as f:
                f.writelines(header_lines)
                f.writelines(chunk_lines)
                f.writelines(bulk_lines)
            run_paths.append(out_path)

        self.log(f"    {label}Split {n_subcases} cases into {n_runs} run file(s) "
                  f"of up to {cases_per_split} cases each (same header/bulk data in every run) "
                  f"in {time.time() - t_split_start:.1f}s")
        return run_paths

    # ==================== STRUCTURE-ONLY BDF (NO INCLUDES) ====================
    def _write_structure_only_bdf(self, bdf_path, folder, label=""):
        """Write a copy of the given (prop-updated + offseted) BDF with INCLUDE
        cards removed, so only the structure itself remains, without needing
        the include files. Reference-only output - never solved."""
        try:
            with open(bdf_path, 'r', encoding='latin-1') as f:
                lines = f.readlines()

            # "INCLUDE" ile başlayan tüm satırları at - geriye sadece
            # dosyanın kendi içindeki yapısal kartlar kalır (harici dosyalara
            # bağımlılık olmadan tek bakışta okunabilir referans çıktısı).
            new_lines = [l for l in lines if not l.strip().upper().startswith('INCLUDE')]
            removed = len(lines) - len(new_lines)

            base, ext = os.path.splitext(bdf_path)
            output_bdf = base + "_stru" + ext
            with open(output_bdf, 'w', encoding='latin-1') as f:
                f.writelines(new_lines)

            self.log(f"    {label}Structure-only BDF written: {os.path.basename(output_bdf)} ({removed} include lines removed)")
            return output_bdf

        except Exception as e:
            self.log(f"    {label}Structure-only BDF error: {e}")
            return None

    # ==================== NASTRAN EXECUTION ====================
    def _run_nastran(self, bdf_path, folder, label=""):
        nastran = self.nastran_path.get()
        if not nastran or not os.path.exists(nastran):
            self.log(f"    {label}WARNING: Nastran exe not found!")
            return False

        t_run_start = time.time()

        # Neither of the two earlier approaches to scratch isolation actually
        # redirected where Nastran puts its scratch files. Confirmed on real
        # hardware: with a Scratch Folder configured, a run FATAL-errored with
        # "There is not enough space on the disk", and the offending scratch
        # file was sitting in C:\Users\<user>\AppData\Local\Temp\<n>\ - the
        # plain Windows per-process temp folder, completely ignoring both the
        # sdirectory= argument and cwd= that were tried before. That path
        # shape is exactly what Windows' GetTempPath() returns, which reads
        # the TMP/TEMP environment variables - so Nastran's scratch placement
        # here follows those variables, not argv or the process cwd. Override
        # them for this subprocess only (not the whole session) so its
        # scratch lands in the chosen Scratch Folder instead of the system
        # C: temp drive, without touching any other process on the machine.
        scratch_base = self.scratch_folder.get().strip()
        run_scratch_dir = None
        made_temp_scratch = False
        try:
            env = None
            if scratch_base:
                # Scratch Folder verilmişse: o klasörün altında bu run'a özel,
                # benzersiz bir alt klasör oluştur ve TEMP/TMP ortam
                # değişkenlerini SADECE bu alt process için bu klasöre yönlendir
                # (os.environ.copy() ile - gerçek ortam değişkenleri/diğer
                # süreçler etkilenmez).
                os.makedirs(scratch_base, exist_ok=True)
                run_scratch_dir = tempfile.mkdtemp(prefix="nastran_scr_", dir=scratch_base)
                made_temp_scratch = True
                env = os.environ.copy()
                env['TEMP'] = run_scratch_dir
                env['TMP'] = run_scratch_dir
                self.log(f"    {label}Scratch dir (via TEMP/TMP): {run_scratch_dir}")

            cwd_dir = run_scratch_dir or folder

            # Nastran komut satırı: out= sonuç çıktı klasörü, scratch=yes
            # geçici dosyalar kullanılsın, batch=no etkileşimsiz/arka plan modu.
            cmd = [nastran, bdf_path, f"out={folder}", "scratch=yes", "batch=no"]
            mem = self.nastran_memory.get().strip()
            if mem:
                cmd.append(f"mem={mem}")

            result = subprocess.run(cmd, cwd=cwd_dir, env=env, capture_output=True, text=True)
            if result.returncode != 0:
                self.log(f"    {label}WARNING: Nastran exited with code {result.returncode}")
                if result.stderr:
                    self.log(f"    {label}stderr: {result.stderr.strip()[:500]}")

            if made_temp_scratch:
                # Nastran gerçekten bu özel scratch klasörünü mü kullandı,
                # yoksa kendi varsayılan konumuna mı yazdı - klasörün boş
                # olup olmamasına bakarak kullanıcıyı uyar.
                scratch_files = os.listdir(run_scratch_dir)
                if scratch_files:
                    self.log(f"    {label}Scratch folder in use: {len(scratch_files)} file(s) in {run_scratch_dir}")
                else:
                    self.log(f"    {label}WARNING: Scratch folder is still empty - Nastran likely used its "
                              f"own default temp location instead of {run_scratch_dir}")

            # subprocess.run() (no shell=True) blocks until Nastran itself has
            # actually exited, unlike the earlier shell=True + Popen().wait()
            # combination. Still confirm the .op2 as a safety net in case some
            # install's launcher detaches a child process anyway.
            if not self._wait_for_op2(bdf_path, folder, scratch_dir=run_scratch_dir, label=label):
                self.log(f"    {label}WARNING: .op2 never appeared/finished writing - "
                          f"treating run as failed (after {time.time() - t_run_start:.1f}s)")
                return False
            self.log(f"    {label}Nastran run finished in {time.time() - t_run_start:.1f}s")
            return True
        except Exception as e:
            self.log(f"    {label}Nastran run error after {time.time() - t_run_start:.1f}s: {e}")
            return False
        finally:
            # Scratch files are pure intermediates, never needed once the run is
            # done - clean up our own temp scratch dir so a long sweep doesn't
            # fill the scratch disk.
            if made_temp_scratch and run_scratch_dir and os.path.isdir(run_scratch_dir):
                shutil.rmtree(run_scratch_dir, ignore_errors=True)

    def _describe_run_progress(self, dirs, base):
        """Diagnostic string for the heartbeat below: which of the usual Nastran
        output files exist yet in the watched folder(s), and their size."""
        parts = []
        seen = set()
        # Her izlenen klasörde (çıktı klasörü + varsa scratch klasörü) olası
        # Nastran çıktı uzantılarını (.f04/.f06/.log/.op2) kontrol et ve
        # var olanların boyutunu raporla - bu sadece teşhis/log amaçlı,
        # gerçek "bitti mi?" kararı _wait_for_op2 içinde verilir.
        for d in dirs:
            if not d or d in seen or not os.path.isdir(d):
                continue
            seen.add(d)
            for ext in ('.f04', '.f06', '.log', '.op2'):
                p = os.path.join(d, base + ext)
                if os.path.exists(p):
                    try:
                        parts.append(f"{os.path.basename(p)}={os.path.getsize(p)}B")
                    except OSError:
                        pass
        return ", ".join(parts) if parts else "no output files found yet in any watched folder"

    def _wait_for_op2(self, bdf_path, folder, scratch_dir=None, timeout=3600, poll_interval=5,
                       heartbeat_every=30, label=""):
        """Poll until Nastran has genuinely finished this run.

        This used to declare "done" once the .op2 file's size stopped
        changing for 3 polls (15s). On a large model that's wrong: Nastran
        can go minutes with no bytes written to the .op2 during assembly/
        decomposition, long before the run is actually done - so that
        heuristic returned True while the real solve (and the detached
        "analysis" process behind it) kept running in the background. The
        tool then read a still-incomplete .op2 (empty/garbage results) AND
        freed its worker slot early, letting more jobs start than Max
        Parallel Runs was supposed to allow.

        The reliable signal is Nastran's own .f06 log: it contains "FATAL"
        on failure and "END OF JOB" once a run has genuinely completed,
        written only when the solve itself is done - not when some
        launcher process happens to exit. Only if no .f06 ever appears at
        all (a different install's behavior) does this fall back to a much
        longer op2-size-stability window (minutes, not seconds) as a last
        resort.
        """
        base = os.path.splitext(os.path.basename(bdf_path))[0]
        f06_path = os.path.join(folder, base + '.f06')
        op2_search_dirs = [folder] + ([scratch_dir] if scratch_dir and scratch_dir != folder else [])
        start = time.time()
        last_heartbeat = 0.0
        last_size = -1
        stable_since = None
        fallback_stable_seconds = 180

        # _find_op2: .op2 dosyasını önce çıktı klasöründe, yoksa scratch
        # klasöründe ara (bazı Nastran kurulumları .op2'yi out= yerine
        # scratch klasörüne yazabiliyor); sadece boyutu > 0 olanı say.
        def _find_op2():
            for d in op2_search_dirs:
                candidate = os.path.join(d, base + '.op2')
                if os.path.exists(candidate) and os.path.getsize(candidate) > 0:
                    return d, candidate
            return None, None

        while time.time() - start < timeout:
            elapsed = time.time() - start

            if os.path.exists(f06_path):
                # .f06'nın sadece son 8000 karakterini oku (performans) -
                # "FATAL"/"END OF JOB" metinleri log'un en sonunda belirir.
                try:
                    with open(f06_path, 'r', errors='ignore') as f:
                        tail = f.read()[-8000:].upper()
                except OSError:
                    tail = ''

                if 'FATAL' in tail:
                    self.log(f"    {label}Nastran reported a FATAL error - see {f06_path}")
                    return False

                if 'END OF JOB' in tail:
                    op2_dir, op2_path = _find_op2()
                    if op2_path:
                        # .op2 beklenen çıktı klasöründe değilse (scratch
                        # klasörüne yazılmışsa), beklenen yere kopyala -
                        # sonraki aşamalar (OP2->H5 dönüşümü vb.) onu hep
                        # aynı, tahmin edilebilir klasörde bulsun.
                        if op2_dir != folder:
                            self.log(f"    {label}NOTE: .op2 appeared in {op2_dir} (not the out= folder) - copying it to {folder}")
                            try:
                                shutil.copy2(op2_path, os.path.join(folder, base + '.op2'))
                            except Exception as e:
                                self.log(f"    {label}WARNING: could not copy .op2 into {folder}: {e}")
                        return True
                    # F06 says the job is over but the .op2 hasn't shown up
                    # yet - keep polling a little longer rather than failing.
            else:
                # Bu kurulumda hiç .f06 üretilmiyor - son çare olarak .op2
                # boyutunun uzun süre (fallback_stable_seconds) değişmediğini
                # görünce "bitti" kabul et.
                _, op2_path = _find_op2()
                if op2_path:
                    size = os.path.getsize(op2_path)
                    if size == last_size:
                        if stable_since is None:
                            stable_since = time.time()
                        elif time.time() - stable_since >= fallback_stable_seconds:
                            return True
                    else:
                        stable_since = None
                    last_size = size

            if elapsed - last_heartbeat >= heartbeat_every:
                last_heartbeat = elapsed
                self.log(f"    {label}...still waiting for Nastran ({int(elapsed)}s elapsed): "
                          f"{self._describe_run_progress([folder] + op2_search_dirs, base)}")

            time.sleep(poll_interval)

        self.log(f"    {label}WARNING: timed out after {timeout}s waiting for {base} to finish")
        _, op2_path = _find_op2()
        return op2_path is not None

    # ==================== SOLVE BASE MODEL ====================
    def _run_single_iteration(self, folder, thickness_overrides, label=""):
        """Write every BDF + fresh offsets for thickness_overrides (splitting
        into run1/run2/... files per Cases per Split), and run Nastran on
        every resulting run file - up to Max Parallel Runs at a time. No
        stress extraction/reporting happens anymore - this tool's job ends
        at producing the .op2/.f06 results themselves. thickness_overrides
        is {} for the base model's own solve. Offsets are recomputed here
        every call (see _prepare_iteration_bdfs / _calculate_offset_csv)
        from the maneuver BDF at this same thickness_overrides. Returns
        (n_ok, n_total) Nastran job counts."""
        try:
            t_prep_start = time.time()
            per_bdf = self._prepare_iteration_bdfs(folder, thickness_overrides, label=label)
            self.log(f"    {label}BDF prep (write + offsets + split) took {time.time() - t_prep_start:.1f}s")

            # Flatten every bdf_model's run file(s) into one job list, so a
            # single Max Parallel Runs cap governs all of them together -
            # whether they came from splitting one bdf_model or from having
            # several bdf_models.
            jobs = []
            for pb in per_bdf:
                multi = len(pb['run_paths']) > 1
                for run_idx, run_path in enumerate(pb['run_paths']):
                    run_folder = os.path.join(pb['subfolder'], f"run{run_idx + 1}") if multi else pb['subfolder']
                    os.makedirs(run_folder, exist_ok=True)
                    jobs.append({
                        'name': pb['name'], 'bdf_path': run_path, 'folder': run_folder,
                        'label': f"{label}[{pb['name']} run{run_idx + 1}] " if multi else label,
                        'run_ok': False,
                    })

            try:
                max_parallel = max(1, int(float(self.max_parallel_runs.get())))
            except (ValueError, TypeError):
                max_parallel = 1

            self.log(f"    {label}Running {len(jobs)} Nastran job(s), max {max_parallel} at a time...")
            t_jobs_start = time.time()

            # Nastran'ı çalıştır ve HER işin gerçekten bitmesini BEKLE
            # (.op2'nin tam yazılmış olması dahil). ThreadPoolExecutor,
            # max_parallel kadar işi aynı anda çalıştırıp fazlasını kuyrukta tutar.
            with concurrent.futures.ThreadPoolExecutor(max_workers=max_parallel) as executor:
                future_to_job = {
                    executor.submit(self._run_nastran, j['bdf_path'], j['folder'], j['label']): j
                    for j in jobs
                }
                done = 0
                for future in concurrent.futures.as_completed(future_to_job):
                    j = future_to_job[future]
                    done += 1
                    try:
                        j['run_ok'] = future.result()
                    except Exception as e:
                        self.log(f"    {label}Nastran run raised an exception for "
                                  f"{os.path.basename(j['bdf_path'])}: {e}")
                        j['run_ok'] = False
                    self.log(f"    {label}{done}/{len(jobs)} Nastran run(s) finished "
                              f"({'OK' if j['run_ok'] else 'FAILED'}: {os.path.basename(j['bdf_path'])})")

            self.log(f"    {label}All {len(jobs)} Nastran job(s) done - wall time "
                      f"{time.time() - t_jobs_start:.1f}s (max {max_parallel} at a time)")
            n_ok = sum(1 for j in jobs if j['run_ok'])
            return n_ok, len(jobs)

        except Exception as e:
            self.log(f"  {label}Iteration ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())
            return 0, 0

    # start_solve: "SOLVE BASE MODEL" butonunun komutu - girdileri doğrular,
    # butonu kilitler ve gerçek işi (_run_solve) GUI'yi dondurmamak için
    # ayrı bir daemon thread'de başlatır.
    def start_solve(self):
        if not self.bdf_paths:
            messagebox.showerror("Error", "Add at least one BDF file")
            return
        if not self.properties_loaded:
            messagebox.showerror("Error", "Load Property Excel first")
            return

        self.is_running = True
        self.btn_start.config(state=tk.DISABLED)

        threading.Thread(target=self._run_solve, daemon=True).start()

    def _run_solve(self):
        """Process the property excel, apply offsets, and solve the base model once."""
        t_solve_start = time.time()
        try:
            # BDF loading (pyNastran parsing every node/element/property) is
            # the slow part and used to run synchronously in start_solve(),
            # freezing the whole GUI until it finished. There is no separate
            # "Load BDF" button, so it happens here instead - on this
            # background thread - the first time SOLVE BASE MODEL is pressed.
            if not self.bdf_model:
                self.load_bdf()
            if not self.bdf_model:
                self.log("\nSOLVE ERROR: BDF loading failed - see log above.")
                return

            # Her çalıştırma zaman damgalı kendi klasörüne gider
            # ("BaseModel_YYYYMMDD_HHMMSS") - böylece ardışık SOLVE
            # tıklamaları birbirinin üzerine yazmaz.
            output_base = self.output_folder.get()
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            run_folder = os.path.join(output_base, f"BaseModel_{ts}")
            os.makedirs(run_folder, exist_ok=True)

            self.log("\n" + "=" * 70)
            self.log("SOLVING BASE MODEL")
            self.log("=" * 70)
            self.log(f"  Bar properties: {len(self.bar_properties)}, Skin properties: {len(self.skin_properties)}")
            split = self.cases_per_split.get().strip()
            self.log(f"  Cases per split: {split if split else 'off (single BDF)'}, Max parallel runs: {self.max_parallel_runs.get()}")
            self.log(f"  Output: {run_folder}")

            self.root.after(0, lambda: self.progress_var.set(10))
            self._build_base_model(run_folder)
            self.root.after(0, lambda: self.progress_var.set(100))

            self.log(f"\n{'=' * 70}")
            self.log("BASE MODEL SOLVE COMPLETE")
            self.log(f"  Output folder: {run_folder}")
            self.log(f"  Total time (incl. BDF load): {time.time() - t_solve_start:.1f}s")
            self.log(f"{'=' * 70}")

        except Exception as e:
            self.log(f"\nSOLVE ERROR after {time.time() - t_solve_start:.1f}s: {e}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            # Hata olsa da olmasa da: koşum bayrağını indir ve SOLVE
            # butonunu yeniden tıklanabilir yap.
            self.is_running = False
            self.root.after(0, lambda: self.btn_start.config(state=tk.NORMAL))


# ============================================================================
# PostProcessingTab  --  "Post Processing" sekmesi (Tab 4)
# ============================================================================
# ============================================================================
# PostProcessingTab'ın görevi iki aşamalıdır: (1) BarPropertySolverTab'ın
# ürettiği .op2 dosyalarını toplu şekilde .h5'e dönüştürmek (harici bir
# converter exe'yi bir Batch_Run.bat üretip çalıştırarak), (2) ardından
# Load Extraction için gereken klasör yapısını (Example LE şablonunun N
# kopyası + her birine özgü TF-X Excel'i) hazırlamak. Tek bir RUN butonu bu
# iki aşamayı art arda çalıştırır.
# ============================================================================
class PostProcessingTab:
    """OP2 -> H5 batch conversion, then Load Extraction folder prep - one
    RUN button runs both stages back to back. Lists every .op2 under a
    chosen folder (recursively, since the Understanding Structure Type
    tab's split/parallel runs nest them under per-run subfolders), writes
    a Batch_Run.bat next to the chosen converter exe - one line per OP2,
    "<exe> "<op2_path>" "<h5_path>"" - and runs it exactly as if it had
    been pasted into that folder and double-clicked by hand. Only once
    every expected H5 file is confirmed on disk does it move on to
    splitting the Load_Extraction sheet and writing the TF-X workbooks."""

    def __init__(self, parent_frame, root):
        self.root = root
        self.parent_frame = parent_frame

        self.op2_folder = tk.StringVar()          # OP2 dosyalarının aranacağı kök klasör
        self.converter_exe_path = tk.StringVar()  # OP2->H5 dönüştürücü .exe
        self.output_folder = tk.StringVar()       # üretilen .h5 dosyalarının yazılacağı klasör
        self.op2_files = []                        # List butonuyla bulunan .op2 dosya yolları (önizleme listesi)

        # ---- Load Extraction prep (2. aşama girdileri) ----
        self.le_input_excel_path = tk.StringVar()    # settings excel with the 'Load_Extraction' sheet
        self.le_example_folder = tk.StringVar()      # template folder (Load_Ext_H5 + TF-X_Load_Extraction.xlsx)
        self.le_maneuver_bdf_path = tk.StringVar()   # maneuver yüklerinin kaynağı BDF
        self.le_thermal_bdf_path = tk.StringVar()    # thermal yüklerinin kaynağı BDF
        self.le_pickle_path = tk.StringVar()         # Load_Ext_H5.exe'nin kullandığı önbellek (.pickle) dosyası
        self.le_split_count = tk.StringVar(value="10")  # Load_Extraction sayfası kaç LoadExtraction_NN grubuna bölünecek
        self.le_output_folder = tk.StringVar()       # where the split LoadExtraction_NN folders are created
        self.le_exe_path = tk.StringVar()             # Load_Ext_H5.exe
        self.le_max_memory = tk.StringVar(value="50")  # --max-memory, GB

        self.is_running = False

        self.setup_ui()

    # ==================== GUI ====================
    # setup_ui: "OP2->H5 dönüşümü" (Section 1) ve "Load Extraction hazırlığı"
    # (Section 2) girdilerini tek bir kaydırılabilir sekmede toplar.
    def setup_ui(self):
        canvas = tk.Canvas(self.parent_frame)
        scrollbar = ttk.Scrollbar(self.parent_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        main = scrollable_frame

        ttk.Label(main, text="Post Processing", style='Header.TLabel').pack(pady=10)
        ttk.Label(main, text="OP2 -> H5 conversion, then Load Extraction prep - one RUN, start to finish",
                 style='Subtitle.TLabel').pack()

        # ---- Section 1: Input Files ----
        f1 = ttk.LabelFrame(main, text="\U0001F4C1  1. Input Files", padding=10)
        f1.pack(fill=tk.X, pady=5, padx=10)

        # OP2 Folder - dönüştürülecek .op2 dosyalarının aranacağı kök klasör
        op2_frame = ttk.Frame(f1)
        op2_frame.pack(fill=tk.X, pady=2)
        ttk.Label(op2_frame, text="OP2 Folder:", width=18).pack(side=tk.LEFT)
        ttk.Entry(op2_frame, textvariable=self.op2_folder, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(op2_frame, text="Browse", command=lambda: self.op2_folder.set(
            filedialog.askdirectory()
        )).pack(side=tk.LEFT, padx=2)
        ttk.Button(op2_frame, text="List", command=self.list_op2_files).pack(side=tk.LEFT, padx=2)

        # OP2 files found (read-only listing, searched recursively). This is
        # just a preview - RUN always re-scans the folder fresh on its own.
        list_frame = ttk.Frame(f1)
        list_frame.pack(fill=tk.X, pady=2)
        ttk.Label(list_frame, text="OP2 Files Found:", width=18).pack(side=tk.LEFT, anchor=tk.N)
        op2_list_frame = ttk.Frame(list_frame)
        op2_list_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.op2_listbox = tk.Listbox(op2_list_frame, height=8, width=70, selectmode=tk.EXTENDED,
                                       font=('Consolas', 9), bg=_APP_COLORS['surface'], fg=_APP_COLORS['text'],
                                       selectbackground=_APP_COLORS['primary'], selectforeground='white',
                                       relief='flat', highlightthickness=1,
                                       highlightbackground=_APP_COLORS['border'],
                                       highlightcolor=_APP_COLORS['primary'])
        self.op2_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        op2_scroll = ttk.Scrollbar(op2_list_frame, orient="vertical", command=self.op2_listbox.yview)
        op2_scroll.pack(side=tk.LEFT, fill=tk.Y)
        self.op2_listbox.configure(yscrollcommand=op2_scroll.set)

        self.op2_status = ttk.Label(f1, text="No OP2 files listed", foreground="gray")
        self.op2_status.pack(anchor=tk.W)

        # Converter Exe - OP2'leri H5'e çeviren harici program
        exe_frame = ttk.Frame(f1)
        exe_frame.pack(fill=tk.X, pady=2)
        ttk.Label(exe_frame, text="Converter Exe:", width=18).pack(side=tk.LEFT)
        ttk.Entry(exe_frame, textvariable=self.converter_exe_path, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(exe_frame, text="Browse", command=lambda: self.converter_exe_path.set(
            filedialog.askopenfilename(filetypes=[("Executable", "*.exe"), ("All", "*.*")])
        )).pack(side=tk.LEFT)
        exe_hint = ttk.Frame(f1)
        exe_hint.pack(fill=tk.X, pady=(0, 2))
        ttk.Label(exe_hint, text="(e.g. op2hdf5_standalone.exe - Batch_Run.bat is written next to this "
                                  "exe and run from there, same as pasting/running it by hand)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        # Output Folder - üretilecek .h5 dosyalarının yazılacağı klasör
        out_frame = ttk.Frame(f1)
        out_frame.pack(fill=tk.X, pady=2)
        ttk.Label(out_frame, text="Output Folder (H5):", width=18).pack(side=tk.LEFT)
        ttk.Entry(out_frame, textvariable=self.output_folder, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(out_frame, text="Browse", command=lambda: self.output_folder.set(
            filedialog.askdirectory()
        )).pack(side=tk.LEFT)

        # ---- Section 2: Load Extraction Input Files ----
        f2 = ttk.LabelFrame(main, text="\U0001F4C1  2. Load Extraction Input Files", padding=10)
        f2.pack(fill=tk.X, pady=5, padx=10)

        # Settings Excel ('Load_Extraction' sayfasını okur)
        le_xl_frame = ttk.Frame(f2)
        le_xl_frame.pack(fill=tk.X, pady=2)
        ttk.Label(le_xl_frame, text="Master Input Excel:", width=18).pack(side=tk.LEFT)
        ttk.Entry(le_xl_frame, textvariable=self.le_input_excel_path, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(le_xl_frame, text="Browse", command=lambda: self.le_input_excel_path.set(
            filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")])
        )).pack(side=tk.LEFT)
        le_xl_hint = ttk.Frame(f2)
        le_xl_hint.pack(fill=tk.X, pady=(0, 2))
        ttk.Label(le_xl_hint, text="(the settings excel with the 'Load_Extraction' sheet: Bar Element ID, "
                                    "..., Set)", font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        # Example Load Extraction template folder
        le_ex_frame = ttk.Frame(f2)
        le_ex_frame.pack(fill=tk.X, pady=2)
        ttk.Label(le_ex_frame, text="Example LE Folder:", width=18).pack(side=tk.LEFT)
        ttk.Entry(le_ex_frame, textvariable=self.le_example_folder, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(le_ex_frame, text="Browse", command=lambda: self.le_example_folder.set(
            filedialog.askdirectory()
        )).pack(side=tk.LEFT)
        le_ex_hint = ttk.Frame(f2)
        le_ex_hint.pack(fill=tk.X, pady=(0, 2))
        ttk.Label(le_ex_hint, text="(template folder, e.g. Example_LE - contains Load_Ext_H5 and "
                                    "TF-X_Load_Extraction.xlsx; copied once per split)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        # Maneuver BDF
        le_man_frame = ttk.Frame(f2)
        le_man_frame.pack(fill=tk.X, pady=2)
        ttk.Label(le_man_frame, text="BDF (Maneuver):", width=18).pack(side=tk.LEFT)
        ttk.Entry(le_man_frame, textvariable=self.le_maneuver_bdf_path, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(le_man_frame, text="Browse", command=lambda: self.le_maneuver_bdf_path.set(
            filedialog.askopenfilename(filetypes=[("BDF", "*.bdf *.dat *.nas"), ("All", "*.*")])
        )).pack(side=tk.LEFT)

        # Thermal BDF
        le_thm_frame = ttk.Frame(f2)
        le_thm_frame.pack(fill=tk.X, pady=2)
        ttk.Label(le_thm_frame, text="BDF (Thermal):", width=18).pack(side=tk.LEFT)
        ttk.Entry(le_thm_frame, textvariable=self.le_thermal_bdf_path, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(le_thm_frame, text="Browse", command=lambda: self.le_thermal_bdf_path.set(
            filedialog.askopenfilename(filetypes=[("BDF", "*.bdf *.dat *.nas"), ("All", "*.*")])
        )).pack(side=tk.LEFT)

        # Pickle file
        le_pkl_frame = ttk.Frame(f2)
        le_pkl_frame.pack(fill=tk.X, pady=2)
        ttk.Label(le_pkl_frame, text="Pickle File:", width=18).pack(side=tk.LEFT)
        ttk.Entry(le_pkl_frame, textvariable=self.le_pickle_path, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(le_pkl_frame, text="Browse", command=lambda: self.le_pickle_path.set(
            filedialog.askopenfilename(filetypes=[("Pickle", "*.pickle *.pkl"), ("All", "*.*")])
        )).pack(side=tk.LEFT)

        # Split count
        le_split_frame = ttk.Frame(f2)
        le_split_frame.pack(fill=tk.X, pady=2)
        ttk.Label(le_split_frame, text="Split Into:", width=18).pack(side=tk.LEFT)
        ttk.Entry(le_split_frame, textvariable=self.le_split_count, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(le_split_frame, text="(number of structure groups - rows are grouped by the 'Set' "
                                        "column and split across this many folders)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        # Output folder for the split LoadExtraction_NN folders
        le_out_frame = ttk.Frame(f2)
        le_out_frame.pack(fill=tk.X, pady=2)
        ttk.Label(le_out_frame, text="Output Folder:", width=18).pack(side=tk.LEFT)
        ttk.Entry(le_out_frame, textvariable=self.le_output_folder, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(le_out_frame, text="Browse", command=lambda: self.le_output_folder.set(
            filedialog.askdirectory()
        )).pack(side=tk.LEFT)

        # Load_Ext_H5.exe
        le_exe_frame = ttk.Frame(f2)
        le_exe_frame.pack(fill=tk.X, pady=2)
        ttk.Label(le_exe_frame, text="Load Extraction Exe:", width=18).pack(side=tk.LEFT)
        ttk.Entry(le_exe_frame, textvariable=self.le_exe_path, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(le_exe_frame, text="Browse", command=lambda: self.le_exe_path.set(
            filedialog.askopenfilename(filetypes=[("Executable", "*.exe"), ("All", "*.*")])
        )).pack(side=tk.LEFT)
        le_exe_hint = ttk.Frame(f2)
        le_exe_hint.pack(fill=tk.X, pady=(0, 2))
        ttk.Label(le_exe_hint, text="(e.g. Load_Ext_H5.exe - a loadext.bat is written into every "
                                     "LoadExtraction_NN folder and run from this exe's folder)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        # Max memory (GB)
        le_mem_frame = ttk.Frame(f2)
        le_mem_frame.pack(fill=tk.X, pady=2)
        ttk.Label(le_mem_frame, text="Max Memory (GB):", width=18).pack(side=tk.LEFT)
        ttk.Entry(le_mem_frame, textvariable=self.le_max_memory, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(le_mem_frame, text="(passed as --max-memory to Load_Ext_H5.exe)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        # ---- Section 3: Actions (RUN sits at the very bottom, after every
        # input above - both stages run back to back from this one button) ----
        f3 = ttk.LabelFrame(main, text="▶  3. Actions", padding=10)
        f3.pack(fill=tk.X, pady=5, padx=10)

        btn_row = ttk.Frame(f3)
        btn_row.pack(fill=tk.X)
        self.btn_run = ttk.Button(btn_row, text="▶  RUN", command=self.start_run, style='Accent.TButton')
        self.btn_run.pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row, text="Clear Log", command=self.clear_log).pack(side=tk.LEFT, padx=5)
        ttk.Label(btn_row, text="(1. OP2 -> H5 conversion  2. Load Extraction prep (only once every "
                                 "H5 is confirmed on disk)  3. loadext.bat launched for every TF-X "
                                 "group via Load_Ext_H5.exe, 20s apart)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(f3, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=5)

        # ---- Section 4: Log ----
        f4 = ttk.LabelFrame(main, text="\U0001F4DC  4. Log", padding=10)
        f4.pack(fill=tk.BOTH, expand=True, pady=5, padx=10)

        self.log_text = scrolledtext.ScrolledText(f4, height=20, width=100, font=('Consolas', 10),
                                                    bg=_APP_COLORS['log_bg'], fg=_APP_COLORS['log_fg'],
                                                    insertbackground=_APP_COLORS['log_fg'], relief='flat',
                                                    borderwidth=0)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def log(self, msg):
        def _do():
            self.log_text.insert(tk.END, msg + "\n")
            self.log_text.see(tk.END)
        self.root.after(0, _do)

    def clear_log(self):
        self.log_text.delete(1.0, tk.END)

    # ==================== OP2 LISTING ====================
    # _scan_op2_files: verilen klasörü (ve TÜM alt klasörlerini - os.walk ile
    # rekürsif) tara, ".op2" ile biten her dosyayı bul. Rekürsif olması
    # gerekli çünkü Understanding Structure Type sekmesinin paralel/split
    # çalıştırmaları .op2'leri run1/run2/... alt klasörlerine yerleştiriyor.
    @staticmethod
    def _scan_op2_files(folder):
        files = []
        for dirpath, _, filenames in os.walk(folder):
            for fn in filenames:
                if fn.lower().endswith('.op2'):
                    files.append(os.path.join(dirpath, fn))
        files.sort()
        return files

    def _refresh_op2_listbox(self):
        folder = self.op2_folder.get().strip()
        self.op2_listbox.delete(0, tk.END)
        for p in self.op2_files:
            self.op2_listbox.insert(tk.END, os.path.relpath(p, folder) if folder else p)
        self.op2_status.config(
            text=f"{len(self.op2_files)} OP2 file(s) found" if self.op2_files else "No OP2 files found",
            foreground="green" if self.op2_files else "gray"
        )

    def list_op2_files(self):
        """Manual preview only - just fills the listbox above. RUN below
        always re-scans the OP2 Folder fresh on its own, so this is not a
        required step before pressing RUN."""
        folder = self.op2_folder.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Error", "Select a valid OP2 folder")
            return

        self.log("\n" + "=" * 70)
        self.log("LISTING OP2 FILES")
        self.log("=" * 70)
        self.log(f"  Searching (recursively): {folder}")

        t_start = time.time()
        self.op2_files = self._scan_op2_files(folder)
        self._refresh_op2_listbox()
        self.log(f"  Found {len(self.op2_files)} OP2 file(s) in {time.time() - t_start:.1f}s")

    # ==================== OP2 -> H5 CONVERSION ====================
    def _convert_op2_to_h5(self):
        """Scan OP2 Folder fresh, write Batch_Run.bat, run it, and verify
        every expected H5 file is on disk. Returns (n_ok, n_total)."""
        t_start = time.time()
        folder = self.op2_folder.get().strip()

        self.log("\n" + "=" * 70)
        self.log("LISTING OP2 FILES")
        self.log("=" * 70)
        self.log(f"  Searching (recursively): {folder}")
        self.op2_files = self._scan_op2_files(folder)
        self.root.after(0, self._refresh_op2_listbox)
        self.log(f"  Found {len(self.op2_files)} OP2 file(s) in {time.time() - t_start:.1f}s")

        if not self.op2_files:
            self.log("  No OP2 files found - nothing to convert.")
            return 0, 0

        exe_path = self.converter_exe_path.get().strip()
        exe_dir = os.path.dirname(exe_path) or None  # bat dosyasının yazılıp çalıştırılacağı klasör
        exe_name = os.path.basename(exe_path)          # bat içinde SADECE dosya adıyla referans verilecek
        out_dir = self.output_folder.get().strip()
        os.makedirs(out_dir, exist_ok=True)

        self.log("\n" + "=" * 70)
        self.log("OP2 -> H5 BATCH CONVERSION")
        self.log("=" * 70)
        self.log(f"  OP2 files: {len(self.op2_files)}")
        self.log(f"  Converter: {exe_path}")
        self.log(f"  Output folder: {out_dir}")

        # One line per OP2: "<exe_name>" "<op2_path>" "<h5_path>" - written
        # into the exe's own folder and run from there (cwd=exe_dir), so
        # referencing the exe by its bare name resolves exactly the way it
        # does when this bat file is pasted there and run by hand.
        bat_folder = exe_dir or os.getcwd()
        bat_path = os.path.join(bat_folder, "Batch_Run.bat")
        expected_h5 = []
        lines = []
        for op2_path in self.op2_files:
            base = os.path.splitext(os.path.basename(op2_path))[0]
            h5_path = os.path.join(out_dir, base + ".h5")
            lines.append(f'{exe_name} "{op2_path}" "{h5_path}"')
            expected_h5.append(h5_path)

        with open(bat_path, 'w', encoding='latin-1') as f:
            f.write('\n'.join(lines) + '\n')
        self.log(f"  Batch file written: {bat_path}")  # Batch_Run.bat - her satırda bir OP2->H5 dönüşüm komutu

        self.root.after(0, lambda: self.progress_var.set(2))
        self.log(f"\n  Running batch ({len(lines)} conversion(s))...")
        # Bat dosyasını cmd.exe üzerinden, converter exe'nin KENDİ klasöründen
        # (cwd=exe_dir) çalıştır - tam olarak kullanıcı o klasörde çift
        # tıklamış gibi davranır.
        result = subprocess.run(["cmd", "/c", bat_path], cwd=exe_dir, capture_output=True, text=True)
        self.root.after(0, lambda: self.progress_var.set(30))

        if result.returncode != 0:
            self.log(f"  WARNING: batch exited with code {result.returncode}")
        if result.stdout:
            self.log(f"  stdout: {result.stdout.strip()[:2000]}")
        if result.stderr:
            self.log(f"  stderr: {result.stderr.strip()[:2000]}")

        # Beklenen her .h5 dosyasının gerçekten diskte oluştuğunu doğrula -
        # bu, "converter sessizce başarısız oldu ama bat dosyası 0 ile çıktı"
        # durumunu yakalayan asıl kontrol.
        n_ok = sum(1 for p in expected_h5 if os.path.exists(p))
        self.log(f"\n  {n_ok}/{len(expected_h5)} H5 file(s) produced (took {time.time() - t_start:.1f}s)")
        if n_ok < len(expected_h5):
            missing = [os.path.basename(p) for p in expected_h5 if not os.path.exists(p)]
            shown = ', '.join(missing[:20]) + (" ..." if len(missing) > 20 else "")
            self.log(f"  Missing: {shown}")

        self.root.after(0, lambda: self.progress_var.set(33))
        return n_ok, len(expected_h5)

    # ==================== LOAD EXTRACTION PREP ====================
    # _chunk_list: items listesini n parçaya böl, parçalar mümkün olduğunca
    # eşit boyutlu olsun (k=taban boyut, m=kalan - ilk m parça bir eleman
    # daha fazla alır) - numpy.array_split ile aynı dengeli bölme mantığı.
    @staticmethod
    def _chunk_list(items, n):
        """Split items into n contiguous, as-equal-as-possible chunks (same
        balanced partition as numpy.array_split), preserving order."""
        n = max(1, n)
        k, m = divmod(len(items), n)
        chunks = []
        start = 0
        for i in range(n):
            size = k + (1 if i < m else 0)
            chunks.append(items[start:start + size])
            start += size
        return chunks

    @staticmethod
    def _find_data_sheet(wb):
        """The worksheet named 'Joint_Input' - that is where the Load_
        Extraction sheet's rows (Bar Element ID, ...) get pasted in the
        TF-X workbook. Falls back to searching for a sheet whose header
        row contains 'Bar Element ID' if no 'Joint_Input' sheet exists."""
        for ws in wb.worksheets:
            if ws.title.strip().lower().replace(' ', '_') == 'joint_input':
                return ws
        for ws in wb.worksheets:
            try:
                row1 = next(ws.iter_rows(min_row=1, max_row=1))
            except StopIteration:
                continue
            for cell in row1:
                if cell.value and str(cell.value).strip().lower() == 'bar element id':
                    return ws
        return None

    @staticmethod
    def _find_tfx_workbook_path(dest_folder):
        """The '...Load_Extraction....xls*' workbook inside a copied
        LoadExtraction_NN folder, or None if none is found."""
        for dirpath, _, filenames in os.walk(dest_folder):
            for fn in filenames:
                fnl = fn.lower()
                if 'load_extraction' in fnl.replace(' ', '_') and fnl.endswith(('.xlsx', '.xlsm', '.xls', '.xlsb')):
                    return os.path.join(dirpath, fn)
        return None

    @staticmethod
    def _find_control_sheet(wb, marker_text):
        """The worksheet containing marker_text anywhere in its first ~40
        rows - the control panel sheet with the BDF/Pickle/HDF file fields."""
        marker_lower = marker_text.lower()
        for ws in wb.worksheets:
            for row in ws.iter_rows(min_row=1, max_row=40, max_col=10):
                for cell in row:
                    if cell.value and marker_lower in str(cell.value).lower():
                        return ws
        return None

    def _prepare_load_extraction_core(self):
        """Split the Load_Extraction sheet and write every TF-X workbook.
        Assumes validation already happened in start_run() and that the H5
        Output Folder is now complete (only called after a fully-successful
        _convert_op2_to_h5()). Returns (n_done, n_total_groups)."""
        t_start = time.time()
        input_excel = self.le_input_excel_path.get().strip()
        example_folder = self.le_example_folder.get().strip()
        out_base = self.le_output_folder.get().strip()
        n = int(float(self.le_split_count.get()))
        os.makedirs(out_base, exist_ok=True)

        self.log("\n" + "=" * 70)
        self.log("PREPARING LOAD EXTRACTION")
        self.log("=" * 70)

        # ---- 1. 'Load_Extraction' sayfasını oku ----
        xl = pd.ExcelFile(input_excel)
        le_sheet = None
        for sheet in xl.sheet_names:
            sl = sheet.lower().replace('_', '').replace(' ', '')
            if 'load' in sl and 'extraction' in sl:
                le_sheet = sheet
                break
        if not le_sheet:
            self.log("  ERROR: No 'Load_Extraction' sheet found in the Input Excel.")
            return 0, n

        df = pd.read_excel(xl, sheet_name=le_sheet)
        self.log(f"  Read '{le_sheet}': {len(df)} rows, columns: {list(df.columns)}")

        set_col = next((c for c in df.columns if str(c).strip().lower() == 'set'), None)
        if not set_col:
            self.log("  ERROR: No 'Set' column found in the Load_Extraction sheet.")
            return 0, n

        # ---- 2. 'Set' sütununa göre grupla (ilk görülme sırasıyla), n parçaya böl ----
        unique_sets = list(dict.fromkeys(df[set_col].tolist()))
        set_chunks = self._chunk_list(unique_sets, n)
        self.log(f"  {len(unique_sets)} unique Set(s) -> split into {n} group(s)")
        empty_chunks = sum(1 for c in set_chunks if not c)
        if empty_chunks:
            self.log(f"  NOTE: {empty_chunks} group(s) have no Sets (fewer unique Sets than the split count)")

        # ---- 3. H5 dosyalarını, OP2->H5 dönüşümünün Output Folder'ından topla ----
        h5_folder = self.output_folder.get().strip()
        h5_files = []
        if h5_folder and os.path.isdir(h5_folder):
            for dirpath, _, filenames in os.walk(h5_folder):
                for fn in filenames:
                    if fn.lower().endswith('.h5'):
                        h5_files.append(os.path.join(dirpath, fn))
            h5_files.sort()
        self.log(f"  H5 files found under Output Folder ({h5_folder or 'not set'}): {len(h5_files)}")

        man_bdf = self.le_maneuver_bdf_path.get().strip()
        thm_bdf = self.le_thermal_bdf_path.get().strip()
        pickle_path = self.le_pickle_path.get().strip()

        # ---- 4. Her gruba (chunk) kendi "LoadExtraction_NN" klasörü ----
        n_done = 0
        for i, chunk_sets in enumerate(set_chunks):
            idx = i + 1
            self.log(f"\n  --- Group {idx}/{n}: {len(chunk_sets)} Set(s) ---")
            dest_folder = os.path.join(out_base, f"LoadExtraction_{idx:02d}")
            # Şablon klasörü (Example LE Folder) bu gruba özel klasöre
            # kopyala - her grup kendi bağımsız, tam çalışır kopyasına sahip olur.
            shutil.copytree(example_folder, dest_folder, dirs_exist_ok=True)
            self.log(f"    Template copied to: {dest_folder}")

            tfx_path = self._find_tfx_workbook_path(dest_folder)
            if not tfx_path:
                self.log("    ERROR: No '...Load_Extraction....xls*' file found in the copied template - skipping.")
                continue
            if tfx_path.lower().endswith('.xlsb'):
                self.log(f"    ERROR: {os.path.basename(tfx_path)} is a binary .xlsb workbook - "
                          f"openpyxl can't edit it. Skipping.")
                continue

            # keep_vba: .xlsm/.xltm/.xlam uzantılı dosyalarda makro (VBA) kodu
            # korunmalı - aksi halde openpyxl kaydederken makroları siler.
            keep_vba = tfx_path.lower().endswith(('.xlsm', '.xltm', '.xlam'))
            wb = openpyxl.load_workbook(tfx_path, keep_vba=keep_vba)

            data_ws = self._find_data_sheet(wb)
            if data_ws is None:
                self.log(f"    WARNING: no 'Joint_Input' sheet (or 'Bar Element ID' header) found in "
                          f"{os.path.basename(tfx_path)} - structure rows not written.")
            else:
                # Başlık satırından (1. satır) sütun adı -> sütun numarası haritası çıkar.
                header_map = {}
                for cell in next(data_ws.iter_rows(min_row=1, max_row=1)):
                    if cell.value:
                        header_map[str(cell.value).strip().lower()] = cell.column

                subset_df = df[df[set_col].isin(chunk_sets)]

                # Yapıştırmadan önce başlığın altındaki eski örnek satırları temizle.
                for col_idx in header_map.values():
                    for r in range(2, 2002):
                        data_ws.cell(row=r, column=col_idx).value = None

                # Bu gruba ait satırları (subset_df) sütun adı eşleştirmesiyle yapıştır.
                for r_off, (_, row) in enumerate(subset_df.iterrows()):
                    r = r_off + 2
                    for col_name in subset_df.columns:
                        col_idx = header_map.get(str(col_name).strip().lower())
                        if col_idx:
                            val = row[col_name]
                            data_ws.cell(row=r, column=col_idx).value = None if pd.isna(val) else val
                self.log(f"    Wrote {len(subset_df)} row(s) into '{data_ws.title}'")

            # "Control" sayfasını, içinde "BDF File for Sprung System" metni
            # geçen hücreye bakarak bul (sayfa adı sabit değil, bu yüzden
            # içerik ipucuyla aranıyor).
            control_ws = self._find_control_sheet(wb, "BDF File for Sprung System")
            if control_ws is None:
                self.log(f"    WARNING: no control sheet (with 'BDF File for Sprung System') found in "
                          f"{os.path.basename(tfx_path)} - C9/C10/C11/C17+ not written.")
            else:
                # C9/C10/C11: maneuver/thermal BDF ve pickle dosya yolları.
                control_ws['C9'] = man_bdf
                control_ws['C10'] = thm_bdf
                control_ws['C11'] = pickle_path
                # C17'den başlayarak tüm H5 dosya yollarını, önce eski
                # değerleri temizleyip, listele.
                for r in range(17, 2017):
                    control_ws.cell(row=r, column=3).value = None
                for r_off, p in enumerate(h5_files):
                    control_ws.cell(row=17 + r_off, column=3).value = p
                self.log(f"    Wrote C9/C10/C11 (maneuver/thermal/pickle) and {len(h5_files)} "
                          f"H5 path(s) from C17 into '{control_ws.title}'")

            wb.save(tfx_path)
            n_done += 1
            self.root.after(0, lambda p=33 + (idx / n) * 33: self.progress_var.set(p))

        self.log(f"\n  LOAD EXTRACTION PREP: {n_done}/{n} group(s) (took {time.time() - t_start:.1f}s)")
        return n_done, n

    # ==================== RUN TF-X ANALYSES (Load_Ext_H5.exe BATCH FILES) ====================
    # _find_le_groups_folders: LE Output Folder altındaki tüm "loadextraction_"
    # ile başlayan alt klasörleri bul. İsimler sıfır doldurmalı (01, 02, ...)
    # olduğu için düz alfabetik sıralama 99 gruba kadar doğru sırayı verir.
    def _find_le_groups_folders(self):
        """Every LoadExtraction_NN folder under the LE Output Folder,
        sorted by name (zero-padded, so plain alphabetical sort is correct
        for up to 99 groups)."""
        base = self.le_output_folder.get().strip()
        if not base or not os.path.isdir(base):
            return []
        folders = [
            os.path.join(base, name) for name in os.listdir(base)
            if os.path.isdir(os.path.join(base, name)) and name.lower().startswith('loadextraction_')
        ]
        folders.sort(key=lambda p: os.path.basename(p).lower())
        return folders

    def _run_tfx_batch_core(self):
        """Write a loadext.bat into every LoadExtraction_NN folder (cd's
        into Load_Ext_H5.exe's own folder, then calls it with that group's
        TF-X excel and the shared Maneuver/Thermal/Pickle paths) and launch
        each one, staggering the start of each by 20s - fire and forget,
        not waiting for a run to finish before starting the next. Returns
        (n_started, n_total)."""
        t_start = time.time()
        self.log("\n" + "=" * 70)
        self.log("RUNNING TF-X ANALYSES (Load_Ext_H5.exe)")
        self.log("=" * 70)

        group_folders = self._find_le_groups_folders()
        if not group_folders:
            self.log("  No LoadExtraction_NN folders found - nothing to run.")
            return 0, 0
        n = len(group_folders)
        self.log(f"  {n} Load Extraction group(s) found")

        exe_path = self.le_exe_path.get().strip()
        exe_dir = os.path.dirname(exe_path) or None
        exe_name = os.path.basename(exe_path)
        maneuver_bdf = self.le_maneuver_bdf_path.get().strip()
        thermal_bdf = self.le_thermal_bdf_path.get().strip()
        pickle_path = self.le_pickle_path.get().strip()
        max_memory = self.le_max_memory.get().strip()

        n_started = 0
        for i, dest_folder in enumerate(group_folders):
            idx = i + 1
            self.log(f"\n  --- Group {idx}/{n}: {os.path.basename(dest_folder)} ---")
            tfx_path = self._find_tfx_workbook_path(dest_folder)
            if not tfx_path:
                self.log("    ERROR: No '...Load_Extraction....xls*' file found in this folder - skipping.")
                self.root.after(0, lambda p=66 + (idx / n) * 34: self.progress_var.set(p))
                continue

            # Load_Ext_H5.exe'ye verilecek sabit komut satırı argümanları -
            # bu gruba özgü olan tek şey --xls (bu grubun TF-X Excel'i).
            cmd_line = (
                f'{exe_name} --mode "Joint" --bdf "{maneuver_bdf}" --sprung-bdf "{thermal_bdf}" '
                f'--xls "{tfx_path}" --cid "Material" --nsigma "1.28" --pickle "{pickle_path}" '
                f'--preproc "HM" --stress_output "NO" --output_coord "Buckling" --skin-angle-limit "30" '
                f'--max-memory {max_memory} --machine "Local"'
            )
            bat_path = os.path.join(dest_folder, "loadext.bat")
            with open(bat_path, 'w', encoding='latin-1') as f:
                f.write(f'cd /d "{exe_dir}"\n{cmd_line}\n')
            self.log(f"    Batch file written: {bat_path}")

            # Popen (subprocess.run DEĞİL) ile "fırlat ve unut" (fire and
            # forget) şeklinde başlat - bu TF-X analizi dakikalarca/saatlerce
            # sürebilir ve bu tool onun bitmesini BEKLEMEZ, sadece başlatır.
            # Çıktısı o grubun kendi loadext_log.txt'sine yönlendirilir.
            log_path = os.path.join(dest_folder, "loadext_log.txt")
            try:
                log_file = open(log_path, 'w')
                subprocess.Popen(["cmd", "/c", bat_path], cwd=exe_dir, stdout=log_file, stderr=subprocess.STDOUT)
                self.log(f"    Launched (output -> {log_path})")
                n_started += 1
            except Exception as e:
                self.log(f"    ERROR launching {bat_path}: {e}")

            self.root.after(0, lambda p=66 + (idx / n) * 34: self.progress_var.set(p))

            # Her başlatma arasında 20s bekle - aynı anda çok fazla ağır
            # analiz sürecinin birden fırlatılıp sistemi/lisans sunucusunu
            # boğmasını önlemek için kademeli (staggered) başlatma.
            if idx < n:
                self.log("    Waiting 20s before starting the next one...")
                time.sleep(20)

        self.log(f"\n  TF-X ANALYSES: {n_started}/{n} launched (took {time.time() - t_start:.1f}s). "
                  f"Each runs independently - check each group's loadext_log.txt for progress.")
        return n_started, n

    # ==================== RUN (FULL PIPELINE) ====================
    # start_run: "RUN" butonunun komutu. Her iki aşamanın (OP2->H5 ve Load
    # Extraction hazırlığı) TÜM girdilerini önceden doğrular - gerçek
    # çalışma (arka plan thread'inde _run_pipeline) başlamadan hiçbir
    # eksik/geçersiz alan yüzünden yarı yolda hata vermesin diye.
    def start_run(self):
        # OP2 -> H5 girdileri
        if not self.op2_folder.get().strip() or not os.path.isdir(self.op2_folder.get().strip()):
            messagebox.showerror("Error", "Select a valid OP2 Folder")
            return
        exe_path = self.converter_exe_path.get().strip()
        if not exe_path or not os.path.exists(exe_path):
            messagebox.showerror("Error", "Select a valid Converter Exe")
            return
        if not self.output_folder.get().strip():
            messagebox.showerror("Error", "Select an Output Folder (H5)")
            return

        # Load Extraction girdileri
        if not self.le_input_excel_path.get().strip() or not os.path.exists(self.le_input_excel_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Input Excel")
            return
        if not self.le_example_folder.get().strip() or not os.path.isdir(self.le_example_folder.get().strip()):
            messagebox.showerror("Error", "Select a valid Example LE Folder")
            return
        if not self.le_output_folder.get().strip():
            messagebox.showerror("Error", "Select an Output Folder for the split")
            return
        try:
            n = int(float(self.le_split_count.get()))
            if n <= 0:
                raise ValueError
        except (ValueError, TypeError):
            messagebox.showerror("Error", "Split Into must be a positive number")
            return
        if not self.le_exe_path.get().strip() or not os.path.exists(self.le_exe_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Load Extraction Exe (Load_Ext_H5.exe)")
            return
        try:
            if float(self.le_max_memory.get()) <= 0:
                raise ValueError
        except (ValueError, TypeError):
            messagebox.showerror("Error", "Max Memory (GB) must be a positive number")
            return

        self.is_running = True
        self.btn_run.config(state=tk.DISABLED)
        self.root.after(0, lambda: self.progress_var.set(0))
        threading.Thread(target=self._run_pipeline, daemon=True).start()

    # _run_pipeline: tüm sekmenin 3 aşamasını SIRAYLA çalıştırır ve her
    # aşama SADECE öncekinin TAM başarılı olmasıyla devam eder - yarım
    # kalmış/eksik bir ara adım üzerine bina edip yanlış/eksik sonuç
    # üretmek yerine, hatayı erkenden durdurup kullanıcıya bildirir.
    def _run_pipeline(self):
        """OP2 -> H5 conversion, then (only if every H5 came out OK) Load
        Extraction prep, then (only if at least one group was prepared)
        launching every group's loadext.bat (Load_Ext_H5.exe) - one stage
        after another, end to end, from this one button."""
        t_start = time.time()
        try:
            self.log("\n" + "#" * 70)
            self.log("RUN: OP2 -> H5 CONVERSION -> LOAD EXTRACTION PREP -> TF-X ANALYSES")
            self.log("#" * 70)

            # Aşama 1: OP2 -> H5 dönüşümü.
            n_ok, n_total = self._convert_op2_to_h5()

            if n_total == 0:
                self.log("\nSTOPPED: no OP2 files were converted - Load Extraction prep not started.")
                return
            if n_ok < n_total:
                self.log(f"\nSTOPPED: only {n_ok}/{n_total} H5 file(s) were produced - Load "
                          f"Extraction prep not started. Fix the issue above, then press RUN again.")
                return

            # Aşama 2: Load Extraction klasör/Excel hazırlığı - SADECE tüm H5'ler başarılı olduysa çalışır.
            n_done, n_groups = self._prepare_load_extraction_core()

            if n_done == 0:
                self.log("\nSTOPPED: no Load Extraction group was prepared - TF-X analyses not started.")
                return

            # Aşama 3: her gruba ait TF-X analizini (Load_Ext_H5.exe) başlat - SADECE en az bir grup hazırsa çalışır.
            n_started, n_analyses = self._run_tfx_batch_core()

            self.root.after(0, lambda: self.progress_var.set(100))
            self.log(f"\n{'#' * 70}")
            self.log(f"RUN COMPLETE: {n_total} H5 file(s) converted, {n_done}/{n_groups} Load "
                      f"Extraction group(s) prepared, {n_started}/{n_analyses} TF-X analysis "
                      f"run(s) launched (total {time.time() - t_start:.1f}s)")
            self.log(f"{'#' * 70}")

        except Exception as e:
            self.log(f"\nERROR after {time.time() - t_start:.1f}s: {e}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            self.is_running = False
            self.root.after(0, lambda: self.btn_run.config(state=tk.NORMAL))


# ============================================================================
# JointRunTab  --  "Joint Run" sekmesi (Tab 5)
# ============================================================================
class JointRunTab:
    """Joint Run: groups every Joint_Loads_<SetName>.h5 into its own Set
    folder, copies the Example Joint Input Excel template and the Util
    Folder's contents into each one, and fills the excel's 'Spectrum Joint
    Input' sheet from the Main Input Excel's DTCG_Input_Excel_Data rows
    whose Structure Name matches that Set. Runs the License Checker exe
    once, then per Set folder: writes a User_Log.csv, copies in and runs
    the Load Modifying exe (which produces Joint_Loads_Arranged.h5 -
    the original Set-named H5 is then deleted), runs the RSS Case Finding
    exe (finds the H5 on its own, cwd'd into that folder), writes the
    matching Cap/Skin Joint Loads CSV rows (scaled by a user factor) into
    the 'Limit Stress Check' sheet, and finally runs the Joint Run exe."""
    # Türkçe özet: Her Joint_Loads_<SetName>.h5 dosyası kendi "Set" klasörüne
    # taşınır; o klasöre şablon Excel + Util Folder kopyalanır, Excel'in
    # "Spectrum Joint Input" sayfası Main Input Excel'den o Set'e ait
    # satırlarla doldurulur. Sırasıyla: License Checker (bir kez, başta),
    # her Set için Load Modifying (H5'i "Arranged" hale getirir), RSS Case
    # Finding, Limit Stress Check sayfasının doldurulması, ve son olarak
    # Joint Run exe çalıştırılır - hepsi bu sekmenin tek RUN butonundan.

    def __init__(self, parent_frame, root):
        self.root = root
        self.parent_frame = parent_frame

        self.h5_folder = tk.StringVar()                  # Joint_Loads_<Set>.h5 dosyalarının bulunduğu klasör
        self.main_input_excel_path = tk.StringVar()      # DTCG_Input_Excel_Data sayfasını içeren ana Excel
        self.example_joint_excel_path = tk.StringVar()   # her Set klasörüne kopyalanacak şablon Excel
        self.util_folder = tk.StringVar()                # her Set klasörüne kopyalanacak yardımcı dosyalar klasörü
        self.license_checker_exe_path = tk.StringVar()   # bir kez, en başta çalıştırılan lisans kontrol programı
        self.load_modifying_exe_path = tk.StringVar()    # her Set klasöründe çalıştırılan yük düzenleme programı
        self.rss_exe_path = tk.StringVar()                # RSS Case Finding programı
        self.joint_run_exe_path = tk.StringVar()          # son adımda çalıştırılan Joint Run programı
        self.output_folder = tk.StringVar()               # Set klasörlerinin oluşturulacağı kök çıktı klasörü
        self.max_parallel_runs = tk.StringVar(value="4")  # aynı anda kaç Set'in paralel işleneceği
        self.scale_factor = tk.StringVar(value="1")       # Limit Stress Check'e yazılan yük değerlerine uygulanan çarpan
        self.h5_files = []                                 # List butonuyla bulunan H5 dosyaları (önizleme)

        self.is_running = False

        self.setup_ui()

    # ==================== GUI ====================
    def setup_ui(self):
        canvas = tk.Canvas(self.parent_frame)
        scrollbar = ttk.Scrollbar(self.parent_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        main = scrollable_frame

        ttk.Label(main, text="Joint Run", style='Header.TLabel').pack(pady=10)
        ttk.Label(main, text="Group Joint Loads H5s by Set, fill each Joint Input Excel, run RSS "
                              "Case Finding, then write Limit Stress Check",
                 style='Subtitle.TLabel').pack()

        # ---- Section 1: Input Files ----
        f1 = ttk.LabelFrame(main, text="\U0001F4C1  1. Input Files", padding=10)
        f1.pack(fill=tk.X, pady=5, padx=10)

        # Joint Load H5 Folder
        h5_frame = ttk.Frame(f1)
        h5_frame.pack(fill=tk.X, pady=2)
        ttk.Label(h5_frame, text="Joint Load H5 Folder:", width=20).pack(side=tk.LEFT)
        ttk.Entry(h5_frame, textvariable=self.h5_folder, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(h5_frame, text="Browse", command=lambda: self.h5_folder.set(
            filedialog.askdirectory()
        )).pack(side=tk.LEFT, padx=2)
        ttk.Button(h5_frame, text="List", command=self.list_joint_h5_files).pack(side=tk.LEFT, padx=2)

        list_frame = ttk.Frame(f1)
        list_frame.pack(fill=tk.X, pady=2)
        ttk.Label(list_frame, text="H5 Files Found:", width=20).pack(side=tk.LEFT, anchor=tk.N)
        h5_list_frame = ttk.Frame(list_frame)
        h5_list_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.h5_listbox = tk.Listbox(h5_list_frame, height=8, width=70, selectmode=tk.EXTENDED,
                                      font=('Consolas', 9), bg=_APP_COLORS['surface'], fg=_APP_COLORS['text'],
                                      selectbackground=_APP_COLORS['primary'], selectforeground='white',
                                      relief='flat', highlightthickness=1,
                                      highlightbackground=_APP_COLORS['border'],
                                      highlightcolor=_APP_COLORS['primary'])
        self.h5_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        h5_scroll = ttk.Scrollbar(h5_list_frame, orient="vertical", command=self.h5_listbox.yview)
        h5_scroll.pack(side=tk.LEFT, fill=tk.Y)
        self.h5_listbox.configure(yscrollcommand=h5_scroll.set)

        self.h5_status = ttk.Label(f1, text="No H5 files listed", foreground="gray")
        self.h5_status.pack(anchor=tk.W)
        ttk.Label(f1, text="(Joint_Loads_<SetName>.h5 - every match is processed; this list is just "
                            "a preview, RUN always re-scans)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(anchor=tk.W, pady=(0, 2))

        # Master Input Excel
        main_xl_frame = ttk.Frame(f1)
        main_xl_frame.pack(fill=tk.X, pady=2)
        ttk.Label(main_xl_frame, text="Master Input Excel:", width=20).pack(side=tk.LEFT)
        ttk.Entry(main_xl_frame, textvariable=self.main_input_excel_path, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(main_xl_frame, text="Browse", command=lambda: self.main_input_excel_path.set(
            filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")])
        )).pack(side=tk.LEFT)
        ttk.Label(f1, text="(the 'DTCG_Input_Excel_Data' sheet - filtered by 'Structure Name' per Set)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(anchor=tk.W, pady=(0, 2))

        # Example Joint Input Excel
        ex_xl_frame = ttk.Frame(f1)
        ex_xl_frame.pack(fill=tk.X, pady=2)
        ttk.Label(ex_xl_frame, text="Example Joint Input Excel:", width=20).pack(side=tk.LEFT)
        ttk.Entry(ex_xl_frame, textvariable=self.example_joint_excel_path, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(ex_xl_frame, text="Browse", command=lambda: self.example_joint_excel_path.set(
            filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")])
        )).pack(side=tk.LEFT)
        ttk.Label(f1, text="(template with 'Spectrum Joint Input' and 'Limit Stress Check' sheets - "
                            "copied into every Set folder, renamed to <SetName>)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(anchor=tk.W, pady=(0, 2))

        # Util Folder
        util_frame = ttk.Frame(f1)
        util_frame.pack(fill=tk.X, pady=2)
        ttk.Label(util_frame, text="Util Folder:", width=20).pack(side=tk.LEFT)
        ttk.Entry(util_frame, textvariable=self.util_folder, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(util_frame, text="Browse", command=lambda: self.util_folder.set(
            filedialog.askdirectory()
        )).pack(side=tk.LEFT)
        ttk.Label(f1, text="(the whole folder, e.g. 'util', is copied into every Set folder)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(anchor=tk.W, pady=(0, 2))

        # License Checker Exe
        lic_frame = ttk.Frame(f1)
        lic_frame.pack(fill=tk.X, pady=2)
        ttk.Label(lic_frame, text="License Checker Exe:", width=20).pack(side=tk.LEFT)
        ttk.Entry(lic_frame, textvariable=self.license_checker_exe_path, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(lic_frame, text="Browse", command=lambda: self.license_checker_exe_path.set(
            filedialog.askopenfilename(filetypes=[("Executable", "*.exe"), ("All", "*.*")])
        )).pack(side=tk.LEFT)
        ttk.Label(f1, text="(run once at the very start of RUN, no arguments)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(anchor=tk.W, pady=(0, 2))

        # Load Modifying Exe
        lm_frame = ttk.Frame(f1)
        lm_frame.pack(fill=tk.X, pady=2)
        ttk.Label(lm_frame, text="Load Modifying Exe:", width=20).pack(side=tk.LEFT)
        ttk.Entry(lm_frame, textvariable=self.load_modifying_exe_path, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(lm_frame, text="Browse", command=lambda: self.load_modifying_exe_path.set(
            filedialog.askopenfilename(filetypes=[("Executable", "*.exe"), ("All", "*.*")])
        )).pack(side=tk.LEFT)
        ttk.Label(f1, text="(copied into every Set folder and run there, before RSS Case Finding - "
                            "writes Joint_Loads_Arranged.h5, after which every other .h5 in the "
                            "folder is deleted)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(anchor=tk.W, pady=(0, 2))

        # RSS Case Finding Exe
        rss_frame = ttk.Frame(f1)
        rss_frame.pack(fill=tk.X, pady=2)
        ttk.Label(rss_frame, text="RSS Case Finding Exe:", width=20).pack(side=tk.LEFT)
        ttk.Entry(rss_frame, textvariable=self.rss_exe_path, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(rss_frame, text="Browse", command=lambda: self.rss_exe_path.set(
            filedialog.askopenfilename(filetypes=[("Executable", "*.exe"), ("All", "*.*")])
        )).pack(side=tk.LEFT)
        ttk.Label(f1, text="(copied into every Set folder and run there with no arguments - it "
                            "finds the H5 on its own and writes the Cap/Skin Joint Loads CSVs; "
                            "the copy is removed again before Joint Run runs)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(anchor=tk.W, pady=(0, 2))

        # Joint Run Exe
        jr_frame = ttk.Frame(f1)
        jr_frame.pack(fill=tk.X, pady=2)
        ttk.Label(jr_frame, text="Joint Run Exe:", width=20).pack(side=tk.LEFT)
        ttk.Entry(jr_frame, textvariable=self.joint_run_exe_path, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(jr_frame, text="Browse", command=lambda: self.joint_run_exe_path.set(
            filedialog.askopenfilename(filetypes=[("Executable", "*.exe"), ("All", "*.*")])
        )).pack(side=tk.LEFT)
        ttk.Label(f1, text="(run last: the RSS Case Finding / Load Modifying exe copies are "
                            "deleted, then this is copied into every Set folder and run there, "
                            "up to Max Parallel Runs at a time - refills as soon as a slot frees up)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(anchor=tk.W, pady=(0, 2))

        # Output Folder
        out_frame = ttk.Frame(f1)
        out_frame.pack(fill=tk.X, pady=2)
        ttk.Label(out_frame, text="Output Folder:", width=20).pack(side=tk.LEFT)
        ttk.Entry(out_frame, textvariable=self.output_folder, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(out_frame, text="Browse", command=lambda: self.output_folder.set(
            filedialog.askdirectory()
        )).pack(side=tk.LEFT)
        ttk.Label(f1, text="(every Set gets its own <Output Folder>/<SetName> folder)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(anchor=tk.W, pady=(0, 2))

        # Max Parallel Runs
        par_frame = ttk.Frame(f1)
        par_frame.pack(fill=tk.X, pady=2)
        ttk.Label(par_frame, text="Max Parallel Runs:", width=20).pack(side=tk.LEFT)
        ttk.Entry(par_frame, textvariable=self.max_parallel_runs, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(par_frame, text="(how many Load Modifying / RSS Case Finding / Joint Run runs "
                                   "execute at the same time)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        # Limit Stress Check scale factor
        scale_frame = ttk.Frame(f1)
        scale_frame.pack(fill=tk.X, pady=2)
        ttk.Label(scale_frame, text="LSC Scale Factor:", width=20).pack(side=tk.LEFT)
        ttk.Entry(scale_frame, textvariable=self.scale_factor, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(scale_frame, text="(multiplies every F/N Bearing-Bypass value written into Limit "
                                     "Stress Check)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        # ---- Section 2: Actions (RUN sits at the very bottom) ----
        f2 = ttk.LabelFrame(main, text="▶  2. Actions", padding=10)
        f2.pack(fill=tk.X, pady=5, padx=10)

        btn_row = ttk.Frame(f2)
        btn_row.pack(fill=tk.X)
        self.btn_run = ttk.Button(btn_row, text="▶  RUN", command=self.start_run, style='Accent.TButton')
        self.btn_run.pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row, text="Clear Log", command=self.clear_log).pack(side=tk.LEFT, padx=5)
        ttk.Label(btn_row, text="(1. License Checker  2. Group H5s + Util Folder + fill Spectrum "
                                 "Joint Input  3. Load Modifying  4. RSS Case Finding  5. Limit "
                                 "Stress Check  6. Joint Run - parallel stages wait for all)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(f2, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=5)

        # ---- Section 3: Log ----
        f3 = ttk.LabelFrame(main, text="\U0001F4DC  3. Log", padding=10)
        f3.pack(fill=tk.BOTH, expand=True, pady=5, padx=10)

        self.log_text = scrolledtext.ScrolledText(f3, height=20, width=100, font=('Consolas', 10),
                                                    bg=_APP_COLORS['log_bg'], fg=_APP_COLORS['log_fg'],
                                                    insertbackground=_APP_COLORS['log_fg'], relief='flat',
                                                    borderwidth=0)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def log(self, msg):
        def _do():
            self.log_text.insert(tk.END, msg + "\n")
            self.log_text.see(tk.END)
        self.root.after(0, _do)

    def clear_log(self):
        self.log_text.delete(1.0, tk.END)

    # ==================== H5 LISTING ====================
    # _scan_joint_h5_files: H5 Folder altında (rekürsif), adı "joint_loads_"
    # ile başlayıp ".h5" ile biten tüm dosyaları bul - bunlar işlenecek
    # "Set" dosyalarıdır.
    @staticmethod
    def _scan_joint_h5_files(folder):
        files = []
        for dirpath, _, filenames in os.walk(folder):
            for fn in filenames:
                fnl = fn.lower()
                if fnl.startswith('joint_loads_') and fnl.endswith('.h5'):
                    files.append(os.path.join(dirpath, fn))
        files.sort()
        return files

    # _set_name_from_h5: "Joint_Loads_<SetName>.h5" dosya adından SetName
    # kısmını çıkar - bu ad hem Set klasörünün adı hem de şablon Excel'in
    # yeniden adlandırılacağı ad olarak kullanılır.
    @staticmethod
    def _set_name_from_h5(h5_path):
        base = os.path.splitext(os.path.basename(h5_path))[0]
        return base[len('Joint_Loads_'):] if base.lower().startswith('joint_loads_') else base

    def _refresh_h5_listbox(self):
        folder = self.h5_folder.get().strip()
        self.h5_listbox.delete(0, tk.END)
        for p in self.h5_files:
            self.h5_listbox.insert(tk.END, os.path.relpath(p, folder) if folder else p)
        self.h5_status.config(
            text=f"{len(self.h5_files)} H5 file(s) found" if self.h5_files else "No H5 files found",
            foreground="green" if self.h5_files else "gray"
        )

    def list_joint_h5_files(self):
        """Manual preview only - RUN below always re-scans fresh."""
        folder = self.h5_folder.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Error", "Select a valid Joint Load H5 Folder")
            return
        self.log("\n" + "=" * 70)
        self.log("LISTING JOINT LOAD H5 FILES")
        self.log("=" * 70)
        self.h5_files = self._scan_joint_h5_files(folder)
        self._refresh_h5_listbox()
        self.log(f"  Found {len(self.h5_files)} Joint_Loads_*.h5 file(s)")

    # ==================== SHARED HELPERS ====================
    # _normalize_header: sütun/sayfa adlarını karşılaştırmak için boşluk ve
    # alt çizgiyi silip küçük harfe çevirir ("Bar Element ID" == "bar_element_id").
    @staticmethod
    def _normalize_header(h):
        return str(h).strip().lower().replace(' ', '').replace('_', '')

    # _find_sheet_by_name: normalize edilmiş sayfa adına göre bir worksheet bul.
    @staticmethod
    def _find_sheet_by_name(wb, normalized_target):
        for ws in wb.worksheets:
            if JointRunTab._normalize_header(ws.title) == normalized_target:
                return ws
        return None

    def _find_joint_set_folders(self):
        """Every immediate subfolder of Output Folder that contains at
        least one .h5 file - i.e. every Set folder this tab created."""
        base = self.output_folder.get().strip()
        if not base or not os.path.isdir(base):
            return []
        folders = []
        # Sadece BİRİNCİ seviye alt klasörlere bak (os.walk gibi rekürsif
        # değil) - her Set klasörü doğrudan Output Folder'ın altındadır.
        for name in os.listdir(base):
            full = os.path.join(base, name)
            if os.path.isdir(full) and any(fn.lower().endswith('.h5') for fn in os.listdir(full)):
                folders.append(full)
        folders.sort(key=lambda p: os.path.basename(p).lower())
        return folders

    @staticmethod
    def _find_joint_input_workbook_path(dest_folder, set_name):
        """The '<SetName>.xls*' workbook inside a Set folder, or None."""
        set_name_l = set_name.strip().lower()
        for fn in os.listdir(dest_folder):
            full = os.path.join(dest_folder, fn)
            if not os.path.isfile(full):
                continue
            base, ext = os.path.splitext(fn)
            if base.strip().lower() == set_name_l and ext.lower() in ('.xlsx', '.xlsm', '.xls', '.xlsb'):
                return full
        return None

    # _read_csv_keyed: RSS Case Finding'in ürettiği Cap/Skin Joint Loads
    # CSV'sini oku, başlıkları normalize et, "Bar Element ID" sütununu
    # index yap - böylece sonradan belirli bir eleman ID'si için satır
    # aramak tek bir .loc[] erişimi olur.
    @classmethod
    def _read_csv_keyed(cls, path):
        """Read a Cap/Skin Joint Loads CSV, normalize its headers, and key
        it by Bar Element ID for lookup. Returns None if the file doesn't
        exist or has no Bar Element ID column."""
        if not os.path.exists(path):
            return None
        df = pd.read_csv(path)
        df.columns = [cls._normalize_header(c) for c in df.columns]
        if 'barelementid' not in df.columns:
            return None
        df['barelementid'] = df['barelementid'].astype(str).str.strip()
        return df.set_index('barelementid')

    # ==================== STAGE 1: GROUP H5s + FILL SPECTRUM JOINT INPUT ====================
    # _prepare_joint_folders_core: her Joint_Loads_<Set>.h5 için kendi Set
    # klasörünü oluşturur (H5'i kopyalar, şablon Excel'i <Set>.xlsx olarak
    # kopyalar, Util Folder'ı kopyalar) ve o Excel'in "Spectrum Joint Input"
    # sayfasını Main Input Excel'in DTCG_Input_Excel_Data sayfasından o Set'e
    # (Structure Name) ait satırlarla doldurur.
    def _prepare_joint_folders_core(self):
        """Returns (n_done, n_total)."""
        t_start = time.time()
        h5_folder = self.h5_folder.get().strip()
        main_excel = self.main_input_excel_path.get().strip()
        example_excel = self.example_joint_excel_path.get().strip()
        out_base = self.output_folder.get().strip()
        os.makedirs(out_base, exist_ok=True)

        self.log("\n" + "=" * 70)
        self.log("PREPARING JOINT RUN FOLDERS")
        self.log("=" * 70)

        self.h5_files = self._scan_joint_h5_files(h5_folder)
        self.root.after(0, self._refresh_h5_listbox)
        self.log(f"  Found {len(self.h5_files)} Joint_Loads_*.h5 file(s) under {h5_folder}")
        if not self.h5_files:
            return 0, 0
        n = len(self.h5_files)

        # Main Input Excel'den "DTCG_Input_Excel_Data" sayfasını bul.
        xl = pd.ExcelFile(main_excel)
        data_sheet = None
        for sheet in xl.sheet_names:
            sl = sheet.lower().replace('_', '').replace(' ', '')
            if sl == 'dtcginputexceldata':
                data_sheet = sheet
                break
        if not data_sheet:
            self.log("  ERROR: No 'DTCG_Input_Excel_Data' sheet found in the Main Input Excel.")
            return 0, n
        main_df = pd.read_excel(xl, sheet_name=data_sheet)
        self.log(f"  Read '{data_sheet}': {len(main_df)} rows")

        struct_col = next((c for c in main_df.columns if str(c).strip().lower() == 'structure name'), None)
        if not struct_col:
            self.log("  ERROR: No 'Structure Name' column found in the DTCG_Input_Excel_Data sheet.")
            return 0, n

        example_ext = os.path.splitext(example_excel)[1]

        n_done = 0
        for i, h5_path in enumerate(self.h5_files):
            idx = i + 1
            set_name = self._set_name_from_h5(h5_path)
            self.log(f"\n  --- {idx}/{n}: {set_name} ---")
            dest_folder = os.path.join(out_base, set_name)
            os.makedirs(dest_folder, exist_ok=True)

            # 1. H5 dosyasını kendi Set klasörüne kopyala.
            shutil.copy2(h5_path, os.path.join(dest_folder, os.path.basename(h5_path)))
            self.log(f"    H5 copied to: {dest_folder}")

            # 2. Şablon Excel'i Set'in adıyla (örn. "Spar1_IB.xlsx") kopyala.
            joint_xl_path = os.path.join(dest_folder, set_name + example_ext)
            shutil.copy2(example_excel, joint_xl_path)

            # 3. Util Folder'ın TÜM içeriğini de bu Set klasörüne kopyala.
            util_folder = self.util_folder.get().strip()
            if util_folder:
                util_dest = os.path.join(dest_folder, os.path.basename(os.path.normpath(util_folder)))
                shutil.copytree(util_folder, util_dest, dirs_exist_ok=True)
                self.log(f"    Util Folder copied to: {util_dest}")

            # 4. Main Input Excel'den SADECE bu Set'in (Structure Name) satırlarını filtrele.
            subset_df = main_df[main_df[struct_col].astype(str).str.strip() == set_name]
            self.log(f"    {len(subset_df)} row(s) match Structure Name '{set_name}'")

            keep_vba = joint_xl_path.lower().endswith(('.xlsm', '.xltm', '.xlam'))
            wb = openpyxl.load_workbook(joint_xl_path, keep_vba=keep_vba)
            # 5. "Spectrum Joint Input" sayfasını bul ve filtrelenmiş satırları yapıştır.
            spectrum_ws = self._find_sheet_by_name(wb, 'spectrumjointinput')
            if spectrum_ws is None:
                self.log("    WARNING: no 'Spectrum Joint Input' sheet found in the copied "
                          "template - rows not written.")
            else:
                header_map = {}
                for cell in next(spectrum_ws.iter_rows(min_row=1, max_row=1)):
                    if cell.value:
                        header_map[str(cell.value).strip().lower()] = cell.column

                for col_idx in header_map.values():
                    for r in range(2, 2002):
                        spectrum_ws.cell(row=r, column=col_idx).value = None

                for r_off, (_, row) in enumerate(subset_df.iterrows()):
                    r = r_off + 2
                    for col_name in subset_df.columns:
                        col_idx = header_map.get(str(col_name).strip().lower())
                        if col_idx:
                            val = row[col_name]
                            spectrum_ws.cell(row=r, column=col_idx).value = None if pd.isna(val) else val
                self.log(f"    Wrote {len(subset_df)} row(s) into '{spectrum_ws.title}'")

            wb.save(joint_xl_path)
            n_done += 1
            self.root.after(0, lambda p=5 + (idx / n) * 20: self.progress_var.set(p))

        self.log(f"\n  JOINT RUN FOLDER PREP: {n_done}/{n} Set folder(s) prepared "
                  f"(took {time.time() - t_start:.1f}s)")
        return n_done, n

    # ==================== STAGE 2: LICENSE CHECKER (ONCE) ====================
    # _run_license_checker_core: License Checker exe'sini TEK SEFERLİK,
    # argümansız başlat - bir pencere açıp açık kalabileceği için Popen ile
    # başlatılıp beklenmez, sadece fırlatılır.
    def _run_license_checker_core(self):
        """Launch the License Checker exe exactly once, no arguments - the
        same as double-clicking it. It is not waited on (it may open a
        window and stay open), so this just fires it and moves on."""
        self.log("\n" + "=" * 70)
        self.log("RUNNING LICENSE CHECKER")
        self.log("=" * 70)
        exe_path = self.license_checker_exe_path.get().strip()
        exe_dir = os.path.dirname(exe_path) or None
        try:
            subprocess.Popen([exe_path], cwd=exe_dir)
            self.log(f"  License Checker launched: {exe_path}")
            return True
        except Exception as e:
            self.log(f"  ERROR launching License Checker: {e}")
            return False

    # ==================== STAGE 3: LOAD MODIFYING (PARALLEL) ====================
    def _run_load_modifying_core(self):
        """For every Set folder: write User_Log.csv, copy in the Load
        Modifying exe, and run it (up to Max Parallel Runs at a time,
        waiting for all to finish). Once a run produces
        Joint_Loads_Arranged.h5, the original Joint_Loads_<Set>.h5 is
        deleted so only Joint_Loads_Arranged.h5 remains. Returns
        (n_ok, n_total)."""
        t_start = time.time()
        self.log("\n" + "=" * 70)
        self.log("RUNNING LOAD MODIFYING")
        self.log("=" * 70)

        set_folders = self._find_joint_set_folders()
        if not set_folders:
            self.log("  No Set folders found - nothing to run.")
            return 0, 0
        n = len(set_folders)

        exe_path = self.load_modifying_exe_path.get().strip()
        main_excel = self.main_input_excel_path.get().strip()
        example_ext = os.path.splitext(self.example_joint_excel_path.get().strip())[1]
        try:
            max_workers = max(1, int(float(self.max_parallel_runs.get())))
        except (ValueError, TypeError):
            max_workers = 1

        # Main Input Excel'in "Load_Modify" sayfasından, her Set için
        # "Load Arrangement Direction" değerini oku - bu değer aşağıda
        # yazılacak User_Log.csv'nin bir sütunu olacak.
        direction_map = {}
        try:
            xl = pd.ExcelFile(main_excel)
            lm_sheet = next((s for s in xl.sheet_names if s.strip().lower() == 'load_modify'), None)
            if lm_sheet:
                lm_df = pd.read_excel(xl, sheet_name=lm_sheet)
                set_col = next((c for c in lm_df.columns if str(c).strip().lower() == 'set'), None)
                dir_col = next((c for c in lm_df.columns
                                if str(c).strip().lower() == 'load arrangement direction'), None)
                if set_col and dir_col:
                    for _, row in lm_df.iterrows():
                        direction_map[str(row[set_col]).strip().lower()] = row[dir_col]
                else:
                    self.log("  WARNING: 'Load_Modify' sheet found but missing 'Set' / "
                              "'Load Arrangement Direction' columns.")
            else:
                self.log("  WARNING: No 'Load_Modify' sheet found in the Main Input Excel.")
        except Exception as e:
            self.log(f"  WARNING: could not read 'Load_Modify' sheet: {e}")

        # ---- Her Set klasörüne User_Log.csv yaz ve exe'yi kopyala ----
        for dest_folder in set_folders:
            set_name = os.path.basename(dest_folder)
            joint_xl_path = os.path.join(dest_folder, set_name + example_ext)
            direction = direction_map.get(set_name.strip().lower(), "")
            if not direction:
                self.log(f"  WARNING: no Load Arrangement Direction found for '{set_name}' "
                          f"in the Load_Modify sheet.")

            # Load Modifying exe'sinin bekledği formatta User_Log.csv yaz -
            # bu exe çalıştığında kendi çalışma klasöründe bu dosyayı okuyup
            # hangi H5'i hangi Excel'e göre nasıl düzenleyeceğini öğrenir.
            user_log_path = os.path.join(dest_folder, "User_Log.csv")
            with open(user_log_path, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(["Process Type", "Load Extraction Path", "User Working Path",
                                  "User Working Folder Name", "Input Excel Path", "Output Location",
                                  "Load Arrangement Direction"])
                writer.writerow([
                    "Auto",
                    os.path.join(dest_folder, "Joint_Loads_Arranged.h5"),
                    dest_folder,
                    f"{set_name}_Working Folder",
                    joint_xl_path,
                    dest_folder,
                    direction,
                ])

            shutil.copy2(exe_path, os.path.join(dest_folder, os.path.basename(exe_path)))

        self.log(f"  User_Log.csv written and Load Modifying exe copied into {n} Set folder(s)")
        self.log(f"  Running up to {max_workers} at a time")

        # _run_one: Load Modifying exe'sini bu Set klasöründe (cwd=dest_folder)
        # çalıştır, başarılı olup "Joint_Loads_Arranged.h5" üretmişse,
        # klasördeki DİĞER tüm .h5 dosyalarını (orijinal Joint_Loads_<Set>.h5
        # dahil) sil - geriye sadece "Arranged" (düzenlenmiş) H5 kalsın.
        def _run_one(dest_folder):
            set_name = os.path.basename(dest_folder)
            local_exe = os.path.join(dest_folder, os.path.basename(exe_path))
            log_path = os.path.join(dest_folder, "load_modifying_log.txt")
            try:
                with open(log_path, 'w') as log_file:
                    result = subprocess.run([local_exe], cwd=dest_folder, stdout=log_file,
                                             stderr=subprocess.STDOUT)
                ok = result.returncode == 0
            except Exception as e:
                return set_name, False, str(e)

            arrangement_path = os.path.join(dest_folder, "Joint_Loads_Arranged.h5")
            if os.path.exists(arrangement_path):
                for fn in os.listdir(dest_folder):
                    if fn.lower().endswith('.h5') and fn.lower() != 'joint_loads_arranged.h5':
                        os.remove(os.path.join(dest_folder, fn))
                return set_name, ok, None
            return set_name, False, "Joint_Loads_Arranged.h5 was not produced"

        # Tüm Set klasörlerini, en fazla max_workers kadarı aynı anda
        # çalışacak şekilde paralel işle; hepsi bitene kadar bekle.
        n_ok = 0
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(_run_one, f) for f in set_folders]
            done = 0
            for future in concurrent.futures.as_completed(futures):
                set_name, ok, err = future.result()
                done += 1
                if ok:
                    n_ok += 1
                    self.log(f"    [{done}/{n}] {set_name}: OK")
                else:
                    self.log(f"    [{done}/{n}] {set_name}: FAILED{(' - ' + err) if err else ''}")
                self.root.after(0, lambda p=25 + (done / n) * 20: self.progress_var.set(p))

        self.log(f"\n  LOAD MODIFYING: {n_ok}/{n} completed OK (took {time.time() - t_start:.1f}s)")
        return n_ok, n

    # ==================== STAGE 4: RSS CASE FINDING (PARALLEL) ====================
    # _run_rss_case_finding_core: Load Modifying ile aynı "exe'yi her Set
    # klasörüne kopyala, cwd=dest_folder ile argümansız çalıştır, paralel
    # işle" deseninin daha basit hali - burada argüman/CSV üretimi yok,
    # exe kendi klasöründeki (Joint_Loads_Arranged.h5) dosyayı kendisi bulur.
    def _run_rss_case_finding_core(self):
        """Copy the RSS Case Finding exe into every Set folder and run it
        there (no arguments - it finds the H5 in its own cwd), up to Max
        Parallel Runs at a time, waiting for all of them to finish.
        Returns (n_ok, n_total)."""
        t_start = time.time()
        self.log("\n" + "=" * 70)
        self.log("RUNNING RSS CASE FINDING")
        self.log("=" * 70)

        set_folders = self._find_joint_set_folders()
        if not set_folders:
            self.log("  No Set folders found - nothing to run.")
            return 0, 0
        n = len(set_folders)

        exe_path = self.rss_exe_path.get().strip()
        try:
            max_workers = max(1, int(float(self.max_parallel_runs.get())))
        except (ValueError, TypeError):
            max_workers = 1

        for dest_folder in set_folders:
            shutil.copy2(exe_path, os.path.join(dest_folder, os.path.basename(exe_path)))
        self.log(f"  RSS Case Finding exe copied into {n} Set folder(s)")
        self.log(f"  Running up to {max_workers} at a time")

        def _run_one(dest_folder):
            set_name = os.path.basename(dest_folder)
            local_exe = os.path.join(dest_folder, os.path.basename(exe_path))
            log_path = os.path.join(dest_folder, "rss_case_finding_log.txt")
            try:
                with open(log_path, 'w') as log_file:
                    result = subprocess.run([local_exe], cwd=dest_folder, stdout=log_file,
                                             stderr=subprocess.STDOUT)
                return set_name, result.returncode == 0, None
            except Exception as e:
                return set_name, False, str(e)

        n_ok = 0
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(_run_one, f) for f in set_folders]
            done = 0
            for future in concurrent.futures.as_completed(futures):
                set_name, ok, err = future.result()
                done += 1
                if ok:
                    n_ok += 1
                    self.log(f"    [{done}/{n}] {set_name}: OK")
                else:
                    self.log(f"    [{done}/{n}] {set_name}: FAILED{(' - ' + err) if err else ''}")
                self.root.after(0, lambda p=45 + (done / n) * 20: self.progress_var.set(p))

        self.log(f"\n  RSS CASE FINDING: {n_ok}/{n} completed OK (took {time.time() - t_start:.1f}s)")
        return n_ok, n

    # ==================== STAGE 5: WRITE LIMIT STRESS CHECK ====================
    def _write_limit_stress_check_core(self):
        """For every Set folder, read Cap_Joint_Loads_Arranged.csv and
        Skin_Joint_Loads_Arranged.csv (RSS Case Finding runs against
        Joint_Loads_Arranged.h5, so that's what it names its output after),
        and for every Spectrum Joint Input row
        (keyed by Bar Element ID, routed to Cap/Skin by Type of Element)
        write the matching row into that workbook's own Limit Stress Check
        sheet. Returns (n_done, n_total)."""
        t_start = time.time()
        self.log("\n" + "=" * 70)
        self.log("WRITING LIMIT STRESS CHECK")
        self.log("=" * 70)

        set_folders = self._find_joint_set_folders()
        if not set_folders:
            self.log("  No Set folders found - nothing to write.")
            return 0, 0
        n = len(set_folders)

        try:
            scale = float(self.scale_factor.get())
        except (ValueError, TypeError):
            scale = 1.0
        self.log(f"  Scale factor: {scale}")

        # CSV başlığı (normalize) -> Limit Stress Check sayfasındaki TAM
        # başlık adı (şablonda yazdığı gibi, boşluklu/büyük-küçük harfli).
        col_map = {
            'fbearingx': 'FX Bearing',
            'fbearingy': 'FY Bearing',
            'nxbypass': 'Nx Bypass',
            'nybypass': 'Ny Bypass',
            'nxybypass': 'Nxy Bypass',
        }

        n_done = 0
        for i, dest_folder in enumerate(set_folders):
            idx = i + 1
            set_name = os.path.basename(dest_folder)
            self.log(f"\n  --- {idx}/{n}: {set_name} ---")

            joint_xl_path = self._find_joint_input_workbook_path(dest_folder, set_name)
            if not joint_xl_path:
                self.log(f"    ERROR: No '{set_name}.xls*' workbook found in this folder - skipping.")
                self.root.after(0, lambda p=65 + (idx / n) * 15: self.progress_var.set(p))
                continue

            # RSS Case Finding'in ürettiği Cap/Skin CSV'lerini Bar Element ID
            # ile indexlenmiş halde oku.
            cap_df = self._read_csv_keyed(os.path.join(dest_folder, "Cap_Joint_Loads_Arranged.csv"))
            skin_df = self._read_csv_keyed(os.path.join(dest_folder, "Skin_Joint_Loads_Arranged.csv"))
            if cap_df is None:
                self.log("    WARNING: Cap_Joint_Loads_Arranged.csv not found")
            if skin_df is None:
                self.log("    WARNING: Skin_Joint_Loads_Arranged.csv not found")

            keep_vba = joint_xl_path.lower().endswith(('.xlsm', '.xltm', '.xlam'))
            wb = openpyxl.load_workbook(joint_xl_path, keep_vba=keep_vba)
            spectrum_ws = self._find_sheet_by_name(wb, 'spectrumjointinput')
            lsc_ws = self._find_sheet_by_name(wb, 'limitstresscheck')
            if spectrum_ws is None or lsc_ws is None:
                self.log("    ERROR: missing 'Spectrum Joint Input' or 'Limit Stress Check' "
                          "sheet - skipping.")
                self.root.after(0, lambda p=65 + (idx / n) * 15: self.progress_var.set(p))
                continue

            # "Spectrum Joint Input" sayfasının başlığından Bar Element ID
            # ve Type of Element sütun numaralarını bul - bu sayfa HANGİ
            # elemanların (ve Cap mı Skin mi olduklarının) işleneceğini belirler.
            sp_header = {}
            for cell in next(spectrum_ws.iter_rows(min_row=1, max_row=1)):
                if cell.value:
                    sp_header[self._normalize_header(cell.value)] = cell.column
            bar_id_col = sp_header.get('barelementid')
            type_col = sp_header.get('typeofelement')
            if not bar_id_col or not type_col:
                self.log("    ERROR: 'Bar Element ID' / 'Type of Element' not found in "
                          "Spectrum Joint Input - skipping.")
                self.root.after(0, lambda p=65 + (idx / n) * 15: self.progress_var.set(p))
                continue

            lsc_header = {}
            for cell in next(lsc_ws.iter_rows(min_row=1, max_row=1)):
                if cell.value:
                    lsc_header[self._normalize_header(cell.value)] = cell.column
            lsc_bar_id_col = lsc_header.get('barelementid')

            # Yazmadan önce "Limit Stress Check" sayfasındaki eski satırları temizle.
            for col_idx in lsc_header.values():
                for r in range(2, 2002):
                    lsc_ws.cell(row=r, column=col_idx).value = None

            # "Spectrum Joint Input" sayfasının HER satırı için: eleman
            # tipine (Cap/Skin) göre doğru CSV'den (cap_df/skin_df) o Bar
            # Element ID'nin satırını bul, bulunan F/N Bearing-Bypass
            # değerlerini scale çarpanıyla çarpıp "Limit Stress Check"
            # sayfasına aynı satır sırasıyla yaz. Eşleşme bulunamazsa satır
            # sırası korunur (out_row yine ilerler) ama boş bırakılır - böylece
            # Spectrum Joint Input ile Limit Stress Check satırları HER ZAMAN
            # birbirine karşılık gelir.
            last_row = spectrum_ws.max_row
            n_written = 0
            n_missing = 0
            out_row = 2
            for r in range(2, last_row + 1):
                bar_id = spectrum_ws.cell(row=r, column=bar_id_col).value
                elem_type = spectrum_ws.cell(row=r, column=type_col).value
                if bar_id is None:
                    continue
                key = str(bar_id).strip()
                type_l = str(elem_type).strip().lower() if elem_type else ''
                if 'cap' in type_l:
                    src_df = cap_df
                elif 'skin' in type_l:
                    src_df = skin_df
                else:
                    src_df = None

                row_data = src_df.loc[key] if (src_df is not None and key in src_df.index) else None
                if row_data is None:
                    n_missing += 1
                    out_row += 1
                    continue

                if lsc_bar_id_col:
                    lsc_ws.cell(row=out_row, column=lsc_bar_id_col).value = bar_id
                # col_map'teki her CSV sütununu (normalize) ilgili Limit
                # Stress Check başlığına eşleyip değeri scale ile çarparak yaz.
                for norm_name, lsc_name in col_map.items():
                    col_idx = lsc_header.get(self._normalize_header(lsc_name))
                    if not col_idx:
                        continue
                    val = row_data.get(norm_name)
                    if val is None or (isinstance(val, float) and pd.isna(val)):
                        lsc_ws.cell(row=out_row, column=col_idx).value = None
                    else:
                        try:
                            lsc_ws.cell(row=out_row, column=col_idx).value = float(val) * scale
                        except (ValueError, TypeError):
                            lsc_ws.cell(row=out_row, column=col_idx).value = val
                n_written += 1
                out_row += 1

            wb.save(joint_xl_path)
            extra = f", {n_missing} Bar Element ID(s) not found in CSV" if n_missing else ""
            self.log(f"    Wrote {n_written} row(s) into 'Limit Stress Check'{extra}")
            n_done += 1
            self.root.after(0, lambda p=65 + (idx / n) * 15: self.progress_var.set(p))

        self.log(f"\n  LIMIT STRESS CHECK: {n_done}/{n} workbook(s) updated "
                  f"(took {time.time() - t_start:.1f}s)")
        return n_done, n

    # ==================== STAGE 6: JOINT RUN (ROLLING PARALLEL POOL) ====================
    def _run_joint_run_exe_core(self):
        """Delete the RSS Case Finding / Load Modifying exe copies left in
        every Set folder, then copy in and run the Joint Run exe there (no
        arguments), keeping up to Max Parallel Runs running at the same
        time - as soon as one finishes, the next queued Set starts
        immediately (ThreadPoolExecutor's own scheduling already does
        this: submitting every task up front means a free worker always
        pulls the next one, rather than waiting for a whole batch of N to
        finish). Waits for every run to finish. Returns (n_ok, n_total)."""
        t_start = time.time()
        self.log("\n" + "=" * 70)
        self.log("RUNNING JOINT RUN")
        self.log("=" * 70)

        set_folders = self._find_joint_set_folders()
        if not set_folders:
            self.log("  No Set folders found - nothing to run.")
            return 0, 0
        n = len(set_folders)

        exe_path = self.joint_run_exe_path.get().strip()
        rss_exe_name = os.path.basename(self.rss_exe_path.get().strip())
        lm_exe_name = os.path.basename(self.load_modifying_exe_path.get().strip())
        try:
            max_workers = max(1, int(float(self.max_parallel_runs.get())))
        except (ValueError, TypeError):
            max_workers = 1

        # Önceki aşamalardan kalan RSS Case Finding / Load Modifying exe
        # kopyalarını her Set klasöründen temizle (artık gerekli değiller,
        # klasörü karıştırmasınlar), sonra Joint Run exe'sini kopyala.
        for dest_folder in set_folders:
            for stale_name in (rss_exe_name, lm_exe_name):
                stale_path = os.path.join(dest_folder, stale_name)
                if os.path.exists(stale_path):
                    os.remove(stale_path)
            shutil.copy2(exe_path, os.path.join(dest_folder, os.path.basename(exe_path)))
        self.log(f"  RSS Case Finding / Load Modifying exe copies removed, Joint Run exe "
                  f"copied into {n} Set folder(s)")
        self.log(f"  Keeping up to {max_workers} running at a time")

        def _run_one(dest_folder):
            set_name = os.path.basename(dest_folder)
            local_exe = os.path.join(dest_folder, os.path.basename(exe_path))
            log_path = os.path.join(dest_folder, "joint_run_log.txt")
            try:
                with open(log_path, 'w') as log_file:
                    result = subprocess.run([local_exe], cwd=dest_folder, stdout=log_file,
                                             stderr=subprocess.STDOUT)
                return set_name, result.returncode == 0, None
            except Exception as e:
                return set_name, False, str(e)

        n_ok = 0
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(_run_one, f) for f in set_folders]
            done = 0
            for future in concurrent.futures.as_completed(futures):
                set_name, ok, err = future.result()
                done += 1
                if ok:
                    n_ok += 1
                    self.log(f"    [{done}/{n}] {set_name}: OK")
                else:
                    self.log(f"    [{done}/{n}] {set_name}: FAILED{(' - ' + err) if err else ''}")
                self.root.after(0, lambda p=80 + (done / n) * 20: self.progress_var.set(p))

        self.log(f"\n  JOINT RUN: {n_ok}/{n} completed OK (took {time.time() - t_start:.1f}s)")
        return n_ok, n

    # ==================== RUN (FULL PIPELINE) ====================
    # start_run: "RUN" butonunun komutu - 6 aşamanın (License Checker,
    # klasör hazırlığı, Load Modifying, RSS Case Finding, Limit Stress
    # Check, Joint Run) TÜM girdilerini önceden doğrular, sonra gerçek işi
    # (_run_pipeline) arka plan thread'inde başlatır.
    def start_run(self):
        if not self.h5_folder.get().strip() or not os.path.isdir(self.h5_folder.get().strip()):
            messagebox.showerror("Error", "Select a valid Joint Load H5 Folder")
            return
        if not self.main_input_excel_path.get().strip() or not os.path.exists(self.main_input_excel_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Main Input Excel")
            return
        if not self.example_joint_excel_path.get().strip() or not os.path.exists(self.example_joint_excel_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Example Joint Input Excel")
            return
        if not self.util_folder.get().strip() or not os.path.isdir(self.util_folder.get().strip()):
            messagebox.showerror("Error", "Select a valid Util Folder")
            return
        if not self.license_checker_exe_path.get().strip() or not os.path.exists(self.license_checker_exe_path.get().strip()):
            messagebox.showerror("Error", "Select a valid License Checker Exe")
            return
        if not self.load_modifying_exe_path.get().strip() or not os.path.exists(self.load_modifying_exe_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Load Modifying Exe")
            return
        if not self.rss_exe_path.get().strip() or not os.path.exists(self.rss_exe_path.get().strip()):
            messagebox.showerror("Error", "Select a valid RSS Case Finding Exe")
            return
        if not self.joint_run_exe_path.get().strip() or not os.path.exists(self.joint_run_exe_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Joint Run Exe")
            return
        if not self.output_folder.get().strip():
            messagebox.showerror("Error", "Select an Output Folder")
            return
        try:
            if int(float(self.max_parallel_runs.get())) <= 0:
                raise ValueError
        except (ValueError, TypeError):
            messagebox.showerror("Error", "Max Parallel Runs must be a positive number")
            return
        try:
            float(self.scale_factor.get())
        except (ValueError, TypeError):
            messagebox.showerror("Error", "LSC Scale Factor must be a number")
            return

        self.is_running = True
        self.btn_run.config(state=tk.DISABLED)
        self.root.after(0, lambda: self.progress_var.set(0))
        threading.Thread(target=self._run_pipeline, daemon=True).start()

    def _run_pipeline(self):
        """License Checker once, then group H5s + Util Folder + fill
        Spectrum Joint Input, then Load Modifying, then RSS Case Finding
        (each waiting for all of its Set folders to finish), then write
        every Limit Stress Check sheet from the CSVs, then Joint Run -
        one stage after another, end to end, from this one button."""
        t_start = time.time()
        try:
            self.log("\n" + "#" * 70)
            self.log("RUN: LICENSE CHECKER -> PREPARE FOLDERS -> LOAD MODIFYING -> "
                      "RSS CASE FINDING -> LIMIT STRESS CHECK -> JOINT RUN")
            self.log("#" * 70)

            # Aşama 1: Lisans kontrolü (bir kez, beklenmeden).
            self._run_license_checker_core()

            # Aşama 2: her Set için klasör hazırlığı.
            n_done, n_total = self._prepare_joint_folders_core()
            if n_total == 0:
                self.log("\nSTOPPED: no Joint_Loads_*.h5 files found - nothing to run.")
                return
            if n_done == 0:
                self.log("\nSTOPPED: no Set folder was prepared - Load Modifying not started.")
                return

            # Aşama 3: Load Modifying (paralel).
            n_lm_ok, n_lm = self._run_load_modifying_core()
            if n_lm == 0:
                self.log("\nSTOPPED: no Set folder found for Load Modifying.")
                return

            # Aşama 4: RSS Case Finding (paralel).
            n_ok, n_rss = self._run_rss_case_finding_core()
            if n_rss == 0:
                self.log("\nSTOPPED: no Set folder found for RSS Case Finding.")
                return

            # Aşama 5: Limit Stress Check sayfalarının doldurulması.
            n_written, n_lsc = self._write_limit_stress_check_core()

            # Aşama 6: Joint Run exe (paralel havuz).
            n_jr_ok, n_jr = self._run_joint_run_exe_core()

            self.root.after(0, lambda: self.progress_var.set(100))
            self.log(f"\n{'#' * 70}")
            self.log(f"RUN COMPLETE: {n_done}/{n_total} Set folder(s) prepared, {n_lm_ok}/{n_lm} "
                      f"Load Modifying run(s) OK, {n_ok}/{n_rss} RSS Case Finding run(s) OK, "
                      f"{n_written}/{n_lsc} Limit Stress Check sheet(s) updated, {n_jr_ok}/{n_jr} "
                      f"Joint Run run(s) OK (total {time.time() - t_start:.1f}s)")
            self.log(f"{'#' * 70}")

        except Exception as e:
            self.log(f"\nERROR after {time.time() - t_start:.1f}s: {e}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            self.is_running = False
            self.root.after(0, lambda: self.btn_run.config(state=tk.NORMAL))


# ============================================================================
# FullRunTab  --  "Full Run" sekmesi (Tab 6)
# ============================================================================
class FullRunTab:
    """Chains Understanding Structure Type -> Post Processing -> Joint Run
    end to end from one button, by driving the three existing tab objects
    directly (their own classes are never modified): push this tab's own
    field values into each tool's real StringVars/lists, then call that
    tool's own blocking pipeline method - _run_solve() /
    _run_pipeline() / _run_pipeline() - the exact same methods each tab's
    own RUN button already uses in its background thread, just called
    here in sequence on FullRunTab's own background thread instead.

    What's auto-wired between stages (no separate input for these):
      - OP2 Folder (Post Processing) = every .op2 produced by the solve,
        moved into <run_folder>/_base_model/OP2_Results
      - Output Folder H5 (Post Processing) = <run_folder>/_base_model/H5_Results
      - BDF (Maneuver) / BDF (Thermal) (Post Processing) = the
        '..._prop_updated_offseted_stru.bdf' files the solve wrote,
        matched by 'maneuver'/'thermal' in their filename
      - Output Folder (Post Processing LE split) =
        <run_folder>/_base_model/Load_Extraction_Results
      - Joint Load H5 Folder (Joint Run) = that same Load_Extraction_Results
        folder (Load_Ext_H5.exe writes each group's Joint_Loads_*.h5 into
        its own LoadExtraction_NN/Load_Ext_H5 subfolder, and Joint Run's
        own H5 scan is already recursive, so no copying is needed)
      - Output Folder (Joint Run) = <run_folder>/_base_model/Joint_Run

    Post Processing's own loadext.bat runs are fire-and-forget (by
    design - see PostProcessingTab), so after _run_pipeline() returns
    this waits for every Load_Ext_H5.exe process to actually exit
    (polling tasklist by image name) before starting Joint Run, since
    Joint Run's H5 scan depends on them having finished writing."""
    # Türkçe özet: Bu sınıf KENDİ hesaplama mantığını içermez - sadece
    # Understanding Structure Type (tab3_tool), Post Processing (tab4_tool)
    # ve Joint Run (tab5_tool) sekmelerinin GERÇEK nesnelerini (constructor'da
    # referans olarak alınmış, kopyalanmamış) kullanıp, bu sekmenin kendi
    # girdi alanlarını o nesnelerin StringVar'larına yazıp, o nesnelerin
    # kendi (RUN butonlarının zaten çağırdığı) pipeline metodlarını SIRAYLA
    # çağırır. Böylece 3 sekmenin kodu hiç değiştirilmeden, tek bir "Full
    # Run" düğmesinden uçtan uca çalıştırılabilir.

    def __init__(self, parent_frame, root, tab3_tool, tab4_tool, tab5_tool):
        self.root = root
        self.parent_frame = parent_frame
        # Bu üç referans KOPYA değil - gerçek tab nesneleri. Bu sekmenin
        # yazdığı her şey (StringVar değerleri, bdf_paths listesi vb.)
        # doğrudan o sekmelerin kendi state'ine yazılır.
        self.tab3_tool = tab3_tool
        self.tab4_tool = tab4_tool
        self.tab5_tool = tab5_tool

        # Master Input Excel: Bar/Skin Property, Offset Element ID,
        # Load_Extraction ve DTCG_Input_Excel_Data sayfalarının HEPSİNİ
        # içeren TEK Excel dosyası (Iterative Full Run sekmesiyle aynı
        # mantık). Arayüzde tek bir "Master Input Excel" alanı gösterildiği
        # için, aşağıdaki 4 StringVar (property_excel_path, element_excel_path,
        # le_input_excel_path, jr_main_input_excel_path) hepsi AYNI StringVar
        # nesnesini paylaşır - _run_pipeline'daki t3/t4/t5'e aktarma kodu
        # hiç değişmedi, sadece hepsi artık aynı dosyayı okuyor.
        self.master_input_excel_path = tk.StringVar()

        # ---- Section 1: Understanding Structure Type ----
        self.bdf_paths = []
        self.maneuver_bdfs = []
        self.property_excel_path = self.master_input_excel_path
        self.element_excel_path = self.master_input_excel_path
        self.nastran_path = tk.StringVar()
        self.nastran_memory = tk.StringVar()
        self.scratch_folder = tk.StringVar()
        self.isu_output_folder = tk.StringVar()
        self.cases_per_split = tk.StringVar()
        self.isu_max_parallel_runs = tk.StringVar(value="1")

        # ---- Section 2: Post Processing ----
        self.converter_exe_path = tk.StringVar()
        self.le_input_excel_path = self.master_input_excel_path
        self.le_example_folder = tk.StringVar()
        self.le_pickle_path = tk.StringVar()
        self.le_split_count = tk.StringVar(value="10")
        self.le_exe_path = tk.StringVar()
        self.le_max_memory = tk.StringVar(value="50")

        # ---- Section 3: Joint Run ----
        self.jr_main_input_excel_path = self.master_input_excel_path
        self.jr_example_joint_excel_path = tk.StringVar()
        self.jr_util_folder = tk.StringVar()
        self.jr_license_checker_exe_path = tk.StringVar()
        self.jr_load_modifying_exe_path = tk.StringVar()
        self.jr_rss_exe_path = tk.StringVar()
        self.jr_joint_run_exe_path = tk.StringVar()
        self.jr_max_parallel_runs = tk.StringVar(value="4")
        self.jr_scale_factor = tk.StringVar(value="1")

        self.is_running = False

        self.setup_ui()

    # ==================== GUI ====================
    def setup_ui(self):
        canvas = tk.Canvas(self.parent_frame)
        scrollbar = ttk.Scrollbar(self.parent_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        main = scrollable_frame

        ttk.Label(main, text="Full Run", style='Header.TLabel').pack(pady=10)
        ttk.Label(main, text="Nastran Solver -> Post Processing -> Joint Run, "
                              "chained end to end from one button",
                 style='Subtitle.TLabel').pack()

        # ---- Section 1: Nastran Solver ----
        f1 = ttk.LabelFrame(main, text="\U0001F9EE  1. Nastran Solver", padding=10)
        f1.pack(fill=tk.X, pady=5, padx=10)

        bdf_frame = ttk.Frame(f1)
        bdf_frame.pack(fill=tk.X, pady=2)
        ttk.Label(bdf_frame, text="All BDF Files:", width=18).pack(side=tk.LEFT, anchor=tk.N)
        bdf_list_frame = ttk.Frame(bdf_frame)
        bdf_list_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.bdf_listbox = tk.Listbox(bdf_list_frame, height=3, width=55, selectmode=tk.SINGLE,
                                       font=('Segoe UI', 9), bg=_APP_COLORS['surface'], fg=_APP_COLORS['text'],
                                       selectbackground=_APP_COLORS['primary'], selectforeground='white',
                                       relief='flat', highlightthickness=1,
                                       highlightbackground=_APP_COLORS['border'],
                                       highlightcolor=_APP_COLORS['primary'])
        self.bdf_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        bdf_scroll = ttk.Scrollbar(bdf_list_frame, orient="vertical", command=self.bdf_listbox.yview)
        bdf_scroll.pack(side=tk.LEFT, fill=tk.Y)
        self.bdf_listbox.configure(yscrollcommand=bdf_scroll.set)
        bdf_btn_frame = ttk.Frame(bdf_frame)
        bdf_btn_frame.pack(side=tk.LEFT, padx=5)
        ttk.Button(bdf_btn_frame, text="Add", command=self._add_bdf).pack(fill=tk.X, pady=1)
        ttk.Button(bdf_btn_frame, text="Remove", command=self._remove_bdf).pack(fill=tk.X, pady=1)

        man_frame = ttk.Frame(f1)
        man_frame.pack(fill=tk.X, pady=2)
        man_lbl = ttk.Frame(man_frame)
        man_lbl.pack(side=tk.LEFT, anchor=tk.N)
        ttk.Label(man_lbl, text="Maneuver BDF:", width=18).pack(anchor=tk.W)
        ttk.Label(man_lbl, text="  (Offset Source)", foreground="blue").pack(anchor=tk.W)
        man_list_frame = ttk.Frame(man_frame)
        man_list_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.maneuver_listbox = tk.Listbox(man_list_frame, height=3, width=55, selectmode=tk.SINGLE,
                                            font=('Segoe UI', 9), bg=_APP_COLORS['surface'], fg=_APP_COLORS['text'],
                                            selectbackground=_APP_COLORS['primary'], selectforeground='white',
                                            relief='flat', highlightthickness=1,
                                            highlightbackground=_APP_COLORS['border'],
                                            highlightcolor=_APP_COLORS['primary'])
        self.maneuver_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        man_scroll = ttk.Scrollbar(man_list_frame, orient="vertical", command=self.maneuver_listbox.yview)
        man_scroll.pack(side=tk.LEFT, fill=tk.Y)
        self.maneuver_listbox.configure(yscrollcommand=man_scroll.set)
        man_btn_frame = ttk.Frame(man_frame)
        man_btn_frame.pack(side=tk.LEFT, padx=5)
        ttk.Button(man_btn_frame, text="Add", command=self._add_maneuver_bdf).pack(fill=tk.X, pady=1)
        ttk.Button(man_btn_frame, text="Remove", command=self._remove_maneuver_bdf).pack(fill=tk.X, pady=1)
        ttk.Label(f1, text="(this BDF must also be included in 'All BDF Files' above with "
                            "'maneuver' in its filename, so Post Processing can auto-find its "
                            "solved, structure-only copy)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(anchor=tk.W, pady=(0, 2))
        ttk.Label(f1, text="(include a BDF with 'thermal' in its filename in 'All BDF Files' too, "
                            "for Post Processing's Thermal BDF to be auto-found)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(anchor=tk.W, pady=(0, 2))

        self._add_file_row(f1, "Master Input Excel:", self.master_input_excel_path,
                            [("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")])
        ttk.Label(f1, text="(one excel with every sheet needed: Bar Property, Skin Property, "
                            "Landing_Offset, Bar_Offset, Load_Extraction, DTCG_Input_Excel_Data)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(anchor=tk.W, pady=(0, 2))
        self._add_file_row(f1, "Nastran Exe:", self.nastran_path,
                            [("Executable", "*.exe"), ("All", "*.*")])

        mem_frame = ttk.Frame(f1)
        mem_frame.pack(fill=tk.X, pady=2)
        ttk.Label(mem_frame, text="Requested Memory:", width=18).pack(side=tk.LEFT)
        ttk.Entry(mem_frame, textvariable=self.nastran_memory, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Label(mem_frame, text="(optional, e.g. 4gb / 500mb - blank = Nastran default)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        self._add_folder_row(f1, "Scratch Folder:", self.scratch_folder)
        self._add_folder_row(f1, "Output Folder:", self.isu_output_folder)
        ttk.Label(f1, text="(every run creates its own BaseModel_<timestamp> folder here)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(anchor=tk.W, pady=(0, 2))

        split_frame = ttk.Frame(f1)
        split_frame.pack(fill=tk.X, pady=2)
        ttk.Label(split_frame, text="Cases per Split:", width=18).pack(side=tk.LEFT)
        ttk.Entry(split_frame, textvariable=self.cases_per_split, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(split_frame, text="(blank/0 = solve as one BDF)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        par1_frame = ttk.Frame(f1)
        par1_frame.pack(fill=tk.X, pady=2)
        ttk.Label(par1_frame, text="Max Parallel Runs:", width=18).pack(side=tk.LEFT)
        ttk.Entry(par1_frame, textvariable=self.isu_max_parallel_runs, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(par1_frame, text="(Nastran runs at the same time)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        # ---- Section 2: Post Processing ----
        f2 = ttk.LabelFrame(main, text="\U0001F4E6  2. Post Processing", padding=10)
        f2.pack(fill=tk.X, pady=5, padx=10)

        self._add_file_row(f2, "Converter Exe:", self.converter_exe_path,
                            [("Executable", "*.exe"), ("All", "*.*")])
        self._add_folder_row(f2, "Example LE Folder:", self.le_example_folder)
        self._add_file_row(f2, "Pickle File:", self.le_pickle_path,
                            [("Pickle", "*.pickle *.pkl"), ("All", "*.*")])
        self._add_file_row(f2, "Load Extraction Exe:", self.le_exe_path,
                            [("Executable", "*.exe"), ("All", "*.*")])

        split2_frame = ttk.Frame(f2)
        split2_frame.pack(fill=tk.X, pady=2)
        ttk.Label(split2_frame, text="Split Into:", width=18).pack(side=tk.LEFT)
        ttk.Entry(split2_frame, textvariable=self.le_split_count, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(split2_frame, text="(number of structure groups)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        mem2_frame = ttk.Frame(f2)
        mem2_frame.pack(fill=tk.X, pady=2)
        ttk.Label(mem2_frame, text="Max Memory (GB):", width=18).pack(side=tk.LEFT)
        ttk.Entry(mem2_frame, textvariable=self.le_max_memory, width=10).pack(side=tk.LEFT, padx=5)

        ttk.Label(f2, text="(Input Excel, OP2 Folder, Output Folder H5, BDF Maneuver/Thermal and "
                            "the LE split Output Folder are all auto-wired from stage 1's results)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(anchor=tk.W, pady=(4, 2))

        # ---- Section 3: Joint Run ----
        f3 = ttk.LabelFrame(main, text="\U0001F517  3. Joint Run", padding=10)
        f3.pack(fill=tk.X, pady=5, padx=10)

        self._add_file_row(f3, "Example Joint Input Excel:", self.jr_example_joint_excel_path,
                            [("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")])
        self._add_folder_row(f3, "Util Folder:", self.jr_util_folder)
        self._add_file_row(f3, "License Checker Exe:", self.jr_license_checker_exe_path,
                            [("Executable", "*.exe"), ("All", "*.*")])
        self._add_file_row(f3, "Load Modifying Exe:", self.jr_load_modifying_exe_path,
                            [("Executable", "*.exe"), ("All", "*.*")])
        self._add_file_row(f3, "RSS Case Finding Exe:", self.jr_rss_exe_path,
                            [("Executable", "*.exe"), ("All", "*.*")])
        self._add_file_row(f3, "Joint Run Exe:", self.jr_joint_run_exe_path,
                            [("Executable", "*.exe"), ("All", "*.*")])

        par2_frame = ttk.Frame(f3)
        par2_frame.pack(fill=tk.X, pady=2)
        ttk.Label(par2_frame, text="Max Parallel Runs:", width=18).pack(side=tk.LEFT)
        ttk.Entry(par2_frame, textvariable=self.jr_max_parallel_runs, width=10).pack(side=tk.LEFT, padx=5)

        scale_frame = ttk.Frame(f3)
        scale_frame.pack(fill=tk.X, pady=2)
        ttk.Label(scale_frame, text="LSC Scale Factor:", width=18).pack(side=tk.LEFT)
        ttk.Entry(scale_frame, textvariable=self.jr_scale_factor, width=10).pack(side=tk.LEFT, padx=5)

        ttk.Label(f3, text="(Main Input Excel, Joint Load H5 Folder and Output Folder are all "
                            "auto-wired from stage 1/2's results)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(anchor=tk.W, pady=(4, 2))

        # ---- Section 4: Actions (RUN sits at the very bottom) ----
        f4 = ttk.LabelFrame(main, text="▶  4. Actions", padding=10)
        f4.pack(fill=tk.X, pady=5, padx=10)

        btn_row = ttk.Frame(f4)
        btn_row.pack(fill=tk.X)
        self.btn_run = ttk.Button(btn_row, text="▶  RUN", command=self.start_run, style='Accent.TButton')
        self.btn_run.pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row, text="Clear Log", command=self.clear_log).pack(side=tk.LEFT, padx=5)
        ttk.Label(btn_row, text="(1. Solve base model  2. Post Processing (H5 + Load Extraction + "
                                 "Load_Ext_H5.exe runs)  3. wait for those to finish  4. Joint Run)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(f4, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=5)

        # ---- Section 5: Log ----
        f5 = ttk.LabelFrame(main, text="\U0001F4DC  5. Log", padding=10)
        f5.pack(fill=tk.BOTH, expand=True, pady=5, padx=10)

        self.log_text = scrolledtext.ScrolledText(f5, height=20, width=100, font=('Consolas', 10),
                                                    bg=_APP_COLORS['log_bg'], fg=_APP_COLORS['log_fg'],
                                                    insertbackground=_APP_COLORS['log_fg'], relief='flat',
                                                    borderwidth=0)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def _add_file_row(self, parent, label, var, filetypes):
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=2)
        ttk.Label(frame, text=label, width=18).pack(side=tk.LEFT)
        ttk.Entry(frame, textvariable=var, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame, text="Browse", command=lambda: var.set(
            filedialog.askopenfilename(filetypes=filetypes)
        )).pack(side=tk.LEFT)

    def _add_folder_row(self, parent, label, var):
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=2)
        ttk.Label(frame, text=label, width=18).pack(side=tk.LEFT)
        ttk.Entry(frame, textvariable=var, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame, text="Browse", command=lambda: var.set(
            filedialog.askdirectory()
        )).pack(side=tk.LEFT)

    def log(self, msg):
        def _do():
            self.log_text.insert(tk.END, msg + "\n")
            self.log_text.see(tk.END)
        self.root.after(0, _do)

    def clear_log(self):
        self.log_text.delete(1.0, tk.END)

    # ==================== BDF LIST OPS ====================
    def _add_bdf(self):
        paths = filedialog.askopenfilenames(filetypes=[("BDF", "*.bdf *.dat *.nas"), ("All", "*.*")])
        for p in paths:
            if p not in self.bdf_paths:
                self.bdf_paths.append(p)
                self.bdf_listbox.insert(tk.END, os.path.basename(p))

    def _remove_bdf(self):
        sel = self.bdf_listbox.curselection()
        if sel:
            idx = sel[0]
            self.bdf_listbox.delete(idx)
            del self.bdf_paths[idx]

    def _add_maneuver_bdf(self):
        files = filedialog.askopenfilenames(filetypes=[("BDF", "*.bdf *.dat *.nas"), ("All", "*.*")])
        for f in files:
            if f not in self.maneuver_bdfs:
                self.maneuver_bdfs.append(f)
                self.maneuver_listbox.insert(tk.END, os.path.basename(f))

    def _remove_maneuver_bdf(self):
        sel = self.maneuver_listbox.curselection()
        if sel:
            idx = sel[0]
            self.maneuver_listbox.delete(idx)
            del self.maneuver_bdfs[idx]

    # ==================== AUTO-WIRING HELPERS ====================
    # _list_base_model_folders: Understanding Structure Type sekmesinin
    # ("BaseModel_YYYYMMDD_HHMMSS" adıyla, zaman damgalı) ürettiği çıktı
    # klasörlerinin adlarını listele - _run_pipeline, solve'den ÖNCE ve
    # SONRA bu kümeyi alıp farkına bakarak hangi klasörün YENİ üretildiğini bulur.
    @staticmethod
    def _list_base_model_folders(output_base):
        if not output_base or not os.path.isdir(output_base):
            return set()
        return {
            name for name in os.listdir(output_base)
            if name.startswith('BaseModel_') and os.path.isdir(os.path.join(output_base, name))
        }

    def _move_op2_files(self, search_root, dest_folder):
        """Collect every .op2 under search_root (excluding dest_folder
        itself) first, then move them all into dest_folder - two passes
        so moving files mid-walk never disturbs os.walk's own traversal.
        Returns how many were moved."""
        dest_abs = os.path.abspath(dest_folder)
        found = []
        # 1. geçiş: hedef klasörün KENDİSİNİ hariç tutarak (dirnames[:] = []
        # ile os.walk'ın o alt klasöre inmesini durdurarak) tüm .op2'leri topla.
        for dirpath, dirnames, filenames in os.walk(search_root):
            if os.path.abspath(dirpath) == dest_abs:
                dirnames[:] = []
                continue
            for fn in filenames:
                if fn.lower().endswith('.op2'):
                    found.append(os.path.join(dirpath, fn))

        # 2. geçiş: şimdi hepsini taşı - os.walk tamamen bittiği için
        # taşıma işlemi gezinmeyi bozmaz. Aynı isimde dosya zaten varsa
        # "_2", "_3", ... ekleyerek üzerine yazmadan benzersiz bir ad bul.
        moved = 0
        for src in found:
            fn = os.path.basename(src)
            dst = os.path.join(dest_folder, fn)
            if os.path.exists(dst):
                base, ext = os.path.splitext(fn)
                i = 2
                while os.path.exists(dst):
                    dst = os.path.join(dest_folder, f"{base}_{i}{ext}")
                    i += 1
                self.log(f"    NOTE: name collision, {fn} -> {os.path.basename(dst)}")
            shutil.move(src, dst)
            moved += 1
        return moved

    # _find_stru_bdf: BarPropertySolverTab'ın ürettiği "..._stru.bdf"
    # (structure-only, Excel özellikleri + offset uygulanmış) dosyalardan
    # adında keyword ('maneuver' ya da 'thermal') geçeni bul - Post
    # Processing'in "BDF (Maneuver)"/"BDF (Thermal)" alanlarına otomatik
    # doldurulacak dosya budur.
    @staticmethod
    def _find_stru_bdf(base_model_folder, keyword):
        """The '..._stru.bdf' (structure-only, prop-updated + offseted)
        file whose name contains keyword ('maneuver'/'thermal'), or None."""
        matches = []
        for dirpath, _, filenames in os.walk(base_model_folder):
            for fn in filenames:
                fnl = fn.lower()
                if fnl.endswith('_stru.bdf') and keyword in fnl:
                    matches.append(os.path.join(dirpath, fn))
        return matches[0] if matches else None

    # _wait_for_exe_to_finish: Post Processing'in loadext.bat çalıştırmaları
    # "fırlat ve unut" (Popen, beklenmeden) olduğu için elde BEKLENECEK bir
    # process handle'ı yok. Bu yüzden Windows'un `tasklist` komutuyla, verilen
    # .exe adının hâlâ çalışıp çalışmadığı periyodik olarak sorgulanır -
    # PostProcessingTab'ın kodunu değiştirmeden bitişini anlamanın tek yolu bu.
    def _wait_for_exe_to_finish(self, exe_name, poll_seconds=30, max_wait_seconds=12 * 3600):
        """Post Processing's loadext.bat runs are fire-and-forget (see
        PostProcessingTab._run_tfx_batch_core), so there's no handle to
        wait on - poll `tasklist` by image name instead, since that's the
        only way to know they're done without touching that tab's code."""
        if not exe_name:
            return
        self.log(f"\n  Waiting for every '{exe_name}' process to finish before starting Joint Run...")
        waited = 0
        last_log = 0
        while waited < max_wait_seconds:
            try:
                result = subprocess.run(
                    ["tasklist", "/FI", f"IMAGENAME eq {exe_name}"],
                    capture_output=True, text=True
                )
                running = exe_name.lower() in result.stdout.lower()
            except Exception as e:
                self.log(f"  WARNING: could not check running processes ({e}) - proceeding without waiting.")
                return
            if not running:
                self.log(f"  No more '{exe_name}' processes running (waited {waited}s).")
                return
            time.sleep(poll_seconds)
            waited += poll_seconds
            if waited - last_log >= 300:
                self.log(f"    ...still waiting on '{exe_name}' ({waited}s elapsed)")
                last_log = waited
        self.log(f"  WARNING: gave up waiting for '{exe_name}' after {waited}s - proceeding anyway.")

    # ==================== RUN (FULL PIPELINE) ====================
    # start_run: 3 aşamanın (Understanding Structure Type, Post Processing,
    # Joint Run) TÜM girdilerini önceden doğrular, sonra gerçek zincirlemeyi
    # (_run_pipeline) arka plan thread'inde başlatır.
    def start_run(self):
        if not self.bdf_paths:
            messagebox.showerror("Error", "Add at least one All BDF File")
            return
        if not self.master_input_excel_path.get().strip() or not os.path.exists(self.master_input_excel_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Master Input Excel")
            return
        if not self.nastran_path.get().strip() or not os.path.exists(self.nastran_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Nastran Exe")
            return
        if not self.isu_output_folder.get().strip():
            messagebox.showerror("Error", "Select an Output Folder for Understanding Structure Type")
            return
        if not self.converter_exe_path.get().strip() or not os.path.exists(self.converter_exe_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Converter Exe")
            return
        if not self.le_example_folder.get().strip() or not os.path.isdir(self.le_example_folder.get().strip()):
            messagebox.showerror("Error", "Select a valid Example LE Folder")
            return
        if not self.le_exe_path.get().strip() or not os.path.exists(self.le_exe_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Load Extraction Exe")
            return
        if not self.jr_example_joint_excel_path.get().strip() or not os.path.exists(self.jr_example_joint_excel_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Example Joint Input Excel")
            return
        if not self.jr_util_folder.get().strip() or not os.path.isdir(self.jr_util_folder.get().strip()):
            messagebox.showerror("Error", "Select a valid Util Folder")
            return
        if not self.jr_license_checker_exe_path.get().strip() or not os.path.exists(self.jr_license_checker_exe_path.get().strip()):
            messagebox.showerror("Error", "Select a valid License Checker Exe")
            return
        if not self.jr_load_modifying_exe_path.get().strip() or not os.path.exists(self.jr_load_modifying_exe_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Load Modifying Exe")
            return
        if not self.jr_rss_exe_path.get().strip() or not os.path.exists(self.jr_rss_exe_path.get().strip()):
            messagebox.showerror("Error", "Select a valid RSS Case Finding Exe")
            return
        if not self.jr_joint_run_exe_path.get().strip() or not os.path.exists(self.jr_joint_run_exe_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Joint Run Exe")
            return
        try:
            if int(float(self.le_split_count.get())) <= 0:
                raise ValueError
        except (ValueError, TypeError):
            messagebox.showerror("Error", "Split Into must be a positive number")
            return
        try:
            if float(self.le_max_memory.get()) <= 0:
                raise ValueError
        except (ValueError, TypeError):
            messagebox.showerror("Error", "Post Processing Max Memory must be a positive number")
            return
        try:
            if int(float(self.jr_max_parallel_runs.get())) <= 0:
                raise ValueError
        except (ValueError, TypeError):
            messagebox.showerror("Error", "Joint Run Max Parallel Runs must be a positive number")
            return
        try:
            float(self.jr_scale_factor.get())
        except (ValueError, TypeError):
            messagebox.showerror("Error", "LSC Scale Factor must be a number")
            return

        self.is_running = True
        self.btn_run.config(state=tk.DISABLED)
        self.root.after(0, lambda: self.progress_var.set(0))
        threading.Thread(target=self._run_pipeline, daemon=True).start()

    def _run_pipeline(self):
        t_start = time.time()
        try:
            self.log("\n" + "#" * 70)
            self.log("FULL RUN: UNDERSTANDING STRUCTURE TYPE -> POST PROCESSING -> JOINT RUN")
            self.log("#" * 70)

            # ==================== STAGE A: UNDERSTANDING STRUCTURE TYPE ====================
            self.log("\n" + "#" * 70)
            self.log("STAGE 1/3: SOLVING BASE MODEL")
            self.log("#" * 70)

            # Bu sekmenin kendi girdi alanlarını DOĞRUDAN BarPropertySolverTab
            # (tab3_tool) nesnesinin gerçek state'ine yaz - sanki kullanıcı o
            # sekmeye geçip aynı değerleri kendisi girmiş gibi.
            t3 = self.tab3_tool
            t3.bdf_paths = list(self.bdf_paths)
            t3.maneuver_bdfs = list(self.maneuver_bdfs)
            t3.property_excel_path.set(self.property_excel_path.get())
            t3.element_excel_path.set(self.element_excel_path.get())
            t3.nastran_path.set(self.nastran_path.get())
            t3.nastran_memory.set(self.nastran_memory.get())
            t3.scratch_folder.set(self.scratch_folder.get())
            t3.output_folder.set(self.isu_output_folder.get())
            t3.cases_per_split.set(self.cases_per_split.get())
            t3.max_parallel_runs.set(self.isu_max_parallel_runs.get())
            self.root.after(0, self._refresh_tab3_listboxes)

            t3.load_properties()
            if not t3.properties_loaded:
                self.log("\nSTOPPED: Property Excel failed to load - see the Understanding "
                          "Structure Type log/tab above for details.")
                return
            if self.element_excel_path.get().strip():
                t3.load_element_ids()

            # Solve'den ÖNCEKİ BaseModel_* klasörlerini kaydet - solve'den
            # SONRA aradaki farka (set farkı) bakarak YENİ oluşan klasörü bul.
            output_base = self.isu_output_folder.get().strip()
            before_folders = self._list_base_model_folders(output_base)

            self.root.after(0, lambda: self.progress_var.set(5))
            # t3'ün kendi RUN butonunun arka planda çağırdığı AYNI metod -
            # burada senkron (bloklayarak) çağrılıyor, çünkü zaten biz kendi
            # arka plan thread'imizdeyiz.
            t3._run_solve()
            self.root.after(0, lambda: self.progress_var.set(33))

            after_folders = self._list_base_model_folders(output_base)
            new_folders = after_folders - before_folders
            if len(new_folders) == 1:
                run_folder = os.path.join(output_base, new_folders.pop())
            elif new_folders:
                # Birden fazla yeni klasör varsa (beklenmez ama olası), en
                # son değiştirilmiş olanı (en güncel) seç.
                candidates = [os.path.join(output_base, n) for n in new_folders]
                run_folder = max(candidates, key=os.path.getmtime)
            else:
                run_folder = None

            if not run_folder:
                self.log("\nSTOPPED: no new BaseModel_<timestamp> folder appeared - the solve "
                          "likely failed before writing any output. See the Understanding "
                          "Structure Type log/tab above.")
                return
            self.log(f"\n  Base model run folder: {run_folder}")

            base_model_folder = os.path.join(run_folder, "_base_model")
            op2_count_check = sum(
                1 for dirpath, _, filenames in os.walk(base_model_folder) for fn in filenames
                if fn.lower().endswith('.op2')
            )
            if op2_count_check == 0:
                self.log("\nSTOPPED: no .op2 files were produced - Post Processing not started. "
                          "See the Understanding Structure Type log/tab above.")
                return
            self.log(f"  {op2_count_check} .op2 file(s) found")

            # ==================== STAGE B: POST PROCESSING ====================
            self.log("\n" + "#" * 70)
            self.log("STAGE 2/3: POST PROCESSING")
            self.log("#" * 70)

            # Bu Full Run'a özel 3 alt klasör: OP2'lerin toplanacağı yer,
            # H5 dönüşüm sonuçları, ve Load Extraction grup klasörleri.
            op2_results_dir = os.path.join(base_model_folder, "OP2_Results")
            h5_results_dir = os.path.join(base_model_folder, "H5_Results")
            le_results_dir = os.path.join(base_model_folder, "Load_Extraction_Results")
            os.makedirs(op2_results_dir, exist_ok=True)
            os.makedirs(h5_results_dir, exist_ok=True)
            os.makedirs(le_results_dir, exist_ok=True)

            # Solve'ün her alt klasöre (run1/run2/... veya bdf_N_.../) dağıtık
            # şekilde yazdığı .op2'leri TEK bir klasörde topla - Post
            # Processing'in OP2 Folder taraması zaten rekürsif olduğundan bu
            # adım zorunlu değil ama sonuçların tek yerde olmasını sağlıyor.
            n_moved = self._move_op2_files(base_model_folder, op2_results_dir)
            self.log(f"  Moved {n_moved} .op2 file(s) into {op2_results_dir}")

            # Post Processing'in "BDF (Maneuver)"/"BDF (Thermal)" alanlarını
            # otomatik doldurmak için, solve'ün ürettiği structure-only
            # BDF'lerden adında "maneuver"/"thermal" geçenleri bul.
            maneuver_stru = self._find_stru_bdf(base_model_folder, "maneuver")
            thermal_stru = self._find_stru_bdf(base_model_folder, "thermal")
            if not maneuver_stru:
                self.log("  WARNING: no '..._stru.bdf' file with 'maneuver' in its name was found "
                          "under _base_model - make sure a maneuver BDF is included in 'All BDF "
                          "Files' (not just as the separate Maneuver/Offset Source entry).")
            if not thermal_stru:
                self.log("  WARNING: no '..._stru.bdf' file with 'thermal' in its name was found "
                          "under _base_model - make sure a thermal BDF is included in 'All BDF Files'.")
            self.log(f"  BDF (Maneuver) -> {maneuver_stru}")
            self.log(f"  BDF (Thermal) -> {thermal_stru}")

            # PostProcessingTab (tab4_tool) nesnesinin alanlarını, bu
            # aşamada hesapladığımız otomatik yollarla (op2_results_dir,
            # h5_results_dir, maneuver_stru/thermal_stru, le_results_dir)
            # ve kullanıcının bu sekmeye girdiği değerlerle doldur.
            t4 = self.tab4_tool
            t4.op2_folder.set(op2_results_dir)
            t4.converter_exe_path.set(self.converter_exe_path.get())
            t4.output_folder.set(h5_results_dir)
            t4.le_input_excel_path.set(self.le_input_excel_path.get())
            t4.le_example_folder.set(self.le_example_folder.get())
            t4.le_maneuver_bdf_path.set(maneuver_stru or "")
            t4.le_thermal_bdf_path.set(thermal_stru or "")
            t4.le_pickle_path.set(self.le_pickle_path.get())
            t4.le_split_count.set(self.le_split_count.get())
            t4.le_output_folder.set(le_results_dir)
            t4.le_exe_path.set(self.le_exe_path.get())
            t4.le_max_memory.set(self.le_max_memory.get())

            self.root.after(0, lambda: self.progress_var.set(40))
            t4._run_pipeline()
            self.root.after(0, lambda: self.progress_var.set(66))

            group_folders_found = [
                name for name in (os.listdir(le_results_dir) if os.path.isdir(le_results_dir) else [])
                if os.path.isdir(os.path.join(le_results_dir, name)) and name.lower().startswith('loadextraction_')
            ]
            if not group_folders_found:
                self.log("\nSTOPPED: no LoadExtraction_NN folders were created - Joint Run not "
                          "started. See the Post Processing log/tab above.")
                return
            self.log(f"  {len(group_folders_found)} LoadExtraction_NN group folder(s) found")

            # t4._run_pipeline() döndüğünde loadext.bat'ların hepsi SADECE
            # BAŞLATILMIŞTIR (fire-and-forget) - gerçekten bitmelerini
            # tasklist ile bekle, çünkü Joint Run'ın H5 taraması onların
            # Joint_Loads_*.h5 çıktısına bağımlı.
            self._wait_for_exe_to_finish(os.path.basename(self.le_exe_path.get().strip()))
            self.root.after(0, lambda: self.progress_var.set(70))

            # ==================== STAGE C: JOINT RUN ====================
            self.log("\n" + "#" * 70)
            self.log("STAGE 3/3: JOINT RUN")
            self.log("#" * 70)

            joint_run_dir = os.path.join(base_model_folder, "Joint_Run")
            os.makedirs(joint_run_dir, exist_ok=True)

            # JointRunTab (tab5_tool) nesnesinin alanlarını doldur - H5
            # Folder = Post Processing'in ÜRETTİĞİ le_results_dir (JointRunTab
            # kendi H5 taraması rekürsif olduğundan ayrıca kopyalamaya gerek yok).
            t5 = self.tab5_tool
            t5.h5_folder.set(le_results_dir)
            t5.main_input_excel_path.set(self.jr_main_input_excel_path.get())
            t5.example_joint_excel_path.set(self.jr_example_joint_excel_path.get())
            t5.util_folder.set(self.jr_util_folder.get())
            t5.license_checker_exe_path.set(self.jr_license_checker_exe_path.get())
            t5.load_modifying_exe_path.set(self.jr_load_modifying_exe_path.get())
            t5.rss_exe_path.set(self.jr_rss_exe_path.get())
            t5.joint_run_exe_path.set(self.jr_joint_run_exe_path.get())
            t5.output_folder.set(joint_run_dir)
            t5.max_parallel_runs.set(self.jr_max_parallel_runs.get())
            t5.scale_factor.set(self.jr_scale_factor.get())

            self.root.after(0, lambda: self.progress_var.set(75))
            t5._run_pipeline()
            self.root.after(0, lambda: self.progress_var.set(100))

            self.log(f"\n{'#' * 70}")
            self.log(f"FULL RUN COMPLETE (total {time.time() - t_start:.1f}s)")
            self.log(f"  Base model run folder: {run_folder}")
            self.log(f"{'#' * 70}")

        except Exception as e:
            self.log(f"\nERROR after {time.time() - t_start:.1f}s: {e}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            self.is_running = False
            self.root.after(0, lambda: self.btn_run.config(state=tk.NORMAL))

    # _refresh_tab3_listboxes: t3.bdf_paths/maneuver_bdfs listelerine
    # yukarıda doğrudan yazdığımız için, o sekmenin GUI listbox'ları
    # otomatik güncellenmez - kullanıcı o sekmeye geçip baktığında doğru
    # listeyi görsün diye burada manuel senkronize ediliyor.
    def _refresh_tab3_listboxes(self):
        t3 = self.tab3_tool
        t3.bdf_listbox.delete(0, tk.END)
        for p in t3.bdf_paths:
            t3.bdf_listbox.insert(tk.END, os.path.basename(p))
        t3.maneuver_listbox.delete(0, tk.END)
        for p in t3.maneuver_bdfs:
            t3.maneuver_listbox.insert(tk.END, os.path.basename(p))


# ============================================================================
# IterativeFullRunTab  --  "Iterative Full Run" sekmesi (Tab 7)
# ============================================================================
class IterativeFullRunTab:
    """Same chaining as Full Run (Understanding Structure Type -> Post
    Processing -> Joint Run, driving the three existing tab objects
    directly - their own classes are never modified), looped a fixed
    number of times, feeding each iteration's sizing result back into the
    next iteration's Bar Property thicknesses:

      1. (every iteration) Solve the base model, then fill
         DTCG_Input_Excel_Data's blank 't1(mm)'/'tfem' columns (keyed by
         Bar Property ID) from the PBARL dim1 values the solve actually
         wrote into its '..._prop_updated_offseted*.bdf' files - so the
         joint input excels always reflect what was really solved.
      2. Run Post Processing, then Joint Run (which already waits for
         every Joint Run exe to finish).
      3. Build a flight-hours sizing summary from every 'Nasgro Final
         Summary' csv Joint Run produced (DT_FULL_SUMMARY /
         DT_SIZING_BAR_SUMMARY / DT_SIZING_PROP_SUMMARY /
         DT_LIFE_CHECK_BAR_SUMMARY / DT_LIFE_CHECK_PROP_SUMMARY.csv) -
         a Python port of the reference Summary_Creating.py script.
      4. (iterations before the last) delete every .h5 under that
         iteration's Load Extraction Results / Joint Run folders to save
         disk space, then update the working excel's 'Bar Property'
         sheet Dim1 (only Dim1 - Dim2/t2(mm) are never touched) from
         DT_SIZING_PROP_SUMMARY's governing thickness per Bar Property
         ID, ready for the next iteration's solve.

    A single Master Input Excel supplies every sheet all three tabs need
    (Bar Property, Skin Property, Offset element ID sheets, Load_
    Extraction, DTCG_Input_Excel_Data, Load_Modify) - the user's own file
    is never modified; everything happens on a private working copy."""
    # Türkçe özet: FullRunTab'ın yaptığı 3 aşamalı zincirlemenin (Understanding
    # Structure Type -> Post Processing -> Joint Run) AYNISI, N kere
    # TEKRARLANIR. Her iterasyonda: (1) base model çözülür ve DTCG_Input_
    # Excel_Data'nın t1(mm)/tfem sütunları, solve'ün yazdığı gerçek PBARL
    # dim1 değerleriyle doldurulur; (2) Post Processing + Joint Run koşulur;
    # (3) Joint Run'ın ürettiği "Nasgro Final Summary" CSV'lerinden bir uçuş
    # saati (flight hours) boyutlandırma özeti çıkarılır (referans
    # Summary_Creating.py script'inin Python'a taşınmış hali); (4) SON
    # iterasyon DIŞINDAKİ her iterasyonda: disk yer açmak için o iterasyonun
    # H5 dosyaları silinir, ve boyutlandırma özetindeki yönetici (governing)
    # kalınlık değerleri bir sonraki iterasyonun Bar Property Dim1 (SADECE
    # Dim1) değerleri olarak "working excel"e yazılır - kullanıcının asıl
    # Master Input Excel'i HİÇBİR ZAMAN değiştirilmez.

    def __init__(self, parent_frame, root, tab3_tool, tab4_tool, tab5_tool):
        self.root = root
        self.parent_frame = parent_frame
        self.tab3_tool = tab3_tool
        self.tab4_tool = tab4_tool
        self.tab5_tool = tab5_tool

        # ---- Section 1: Understanding Structure Type ----
        self.bdf_paths = []
        self.maneuver_bdfs = []
        self.master_input_excel_path = tk.StringVar()   # Bar/Skin Property + DTCG_Input_Excel_Data + Load_Modify vb. HER sayfayı içeren tek Excel
        self.nastran_path = tk.StringVar()
        self.nastran_memory = tk.StringVar()
        self.scratch_folder = tk.StringVar()
        self.isu_output_folder = tk.StringVar()
        self.cases_per_split = tk.StringVar()
        self.isu_max_parallel_runs = tk.StringVar(value="1")

        # ---- Section 2: Post Processing ----
        self.converter_exe_path = tk.StringVar()
        self.le_example_folder = tk.StringVar()
        self.le_pickle_path = tk.StringVar()
        self.le_split_count = tk.StringVar(value="10")
        self.le_exe_path = tk.StringVar()
        self.le_max_memory = tk.StringVar(value="50")

        # ---- Section 3: Joint Run ----
        self.jr_example_joint_excel_path = tk.StringVar()
        self.jr_util_folder = tk.StringVar()
        self.jr_license_checker_exe_path = tk.StringVar()
        self.jr_load_modifying_exe_path = tk.StringVar()
        self.jr_rss_exe_path = tk.StringVar()
        self.jr_joint_run_exe_path = tk.StringVar()
        self.jr_max_parallel_runs = tk.StringVar(value="4")
        self.jr_scale_factor = tk.StringVar(value="1")

        # ---- Section 4: Iterasyon ayarları ----
        self.target_flight_hours = tk.StringVar()   # hedef uçuş saati - boyutlandırmanın "yeterli" sayılma eşiği
        self.n_iterations = tk.StringVar(value="3")  # toplam kaç kere tüm zincirin tekrarlanacağı

        self.is_running = False

        self.setup_ui()

    # ==================== GUI ====================
    def setup_ui(self):
        canvas = tk.Canvas(self.parent_frame)
        scrollbar = ttk.Scrollbar(self.parent_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        main = scrollable_frame

        ttk.Label(main, text="Iterative Full Run", style='Header.TLabel').pack(pady=10)
        ttk.Label(main, text="Full Run, looped - each iteration's sizing summary feeds the next "
                              "iteration's Bar Property thicknesses",
                 style='Subtitle.TLabel').pack()

        # ---- Section 1: Nastran Solver ----
        f1 = ttk.LabelFrame(main, text="\U0001F9EE  1. Nastran Solver", padding=10)
        f1.pack(fill=tk.X, pady=5, padx=10)

        bdf_frame = ttk.Frame(f1)
        bdf_frame.pack(fill=tk.X, pady=2)
        ttk.Label(bdf_frame, text="All BDF Files:", width=18).pack(side=tk.LEFT, anchor=tk.N)
        bdf_list_frame = ttk.Frame(bdf_frame)
        bdf_list_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.bdf_listbox = tk.Listbox(bdf_list_frame, height=3, width=55, selectmode=tk.SINGLE,
                                       font=('Segoe UI', 9), bg=_APP_COLORS['surface'], fg=_APP_COLORS['text'],
                                       selectbackground=_APP_COLORS['primary'], selectforeground='white',
                                       relief='flat', highlightthickness=1,
                                       highlightbackground=_APP_COLORS['border'],
                                       highlightcolor=_APP_COLORS['primary'])
        self.bdf_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        bdf_scroll = ttk.Scrollbar(bdf_list_frame, orient="vertical", command=self.bdf_listbox.yview)
        bdf_scroll.pack(side=tk.LEFT, fill=tk.Y)
        self.bdf_listbox.configure(yscrollcommand=bdf_scroll.set)
        bdf_btn_frame = ttk.Frame(bdf_frame)
        bdf_btn_frame.pack(side=tk.LEFT, padx=5)
        ttk.Button(bdf_btn_frame, text="Add", command=self._add_bdf).pack(fill=tk.X, pady=1)
        ttk.Button(bdf_btn_frame, text="Remove", command=self._remove_bdf).pack(fill=tk.X, pady=1)

        man_frame = ttk.Frame(f1)
        man_frame.pack(fill=tk.X, pady=2)
        man_lbl = ttk.Frame(man_frame)
        man_lbl.pack(side=tk.LEFT, anchor=tk.N)
        ttk.Label(man_lbl, text="Maneuver BDF:", width=18).pack(anchor=tk.W)
        ttk.Label(man_lbl, text="  (Offset Source)", foreground="blue").pack(anchor=tk.W)
        man_list_frame = ttk.Frame(man_frame)
        man_list_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.maneuver_listbox = tk.Listbox(man_list_frame, height=3, width=55, selectmode=tk.SINGLE,
                                            font=('Segoe UI', 9), bg=_APP_COLORS['surface'], fg=_APP_COLORS['text'],
                                            selectbackground=_APP_COLORS['primary'], selectforeground='white',
                                            relief='flat', highlightthickness=1,
                                            highlightbackground=_APP_COLORS['border'],
                                            highlightcolor=_APP_COLORS['primary'])
        self.maneuver_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        man_scroll = ttk.Scrollbar(man_list_frame, orient="vertical", command=self.maneuver_listbox.yview)
        man_scroll.pack(side=tk.LEFT, fill=tk.Y)
        self.maneuver_listbox.configure(yscrollcommand=man_scroll.set)
        man_btn_frame = ttk.Frame(man_frame)
        man_btn_frame.pack(side=tk.LEFT, padx=5)
        ttk.Button(man_btn_frame, text="Add", command=self._add_maneuver_bdf).pack(fill=tk.X, pady=1)
        ttk.Button(man_btn_frame, text="Remove", command=self._remove_maneuver_bdf).pack(fill=tk.X, pady=1)
        ttk.Label(f1, text="(also include this BDF, and one with 'thermal' in its filename, in "
                            "'All BDF Files' above, so Post Processing can auto-find their solved "
                            "structure-only copies)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(anchor=tk.W, pady=(0, 2))

        self._add_file_row(f1, "Master Input Excel:", self.master_input_excel_path,
                            [("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")])
        ttk.Label(f1, text="(one excel with every sheet needed: Bar Property, Skin Property, "
                            "Landing_Offset, Bar_Offset, Load_Extraction, DTCG_Input_Excel_Data, "
                            "Load_Modify - never modified in place, only a private working copy is)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(anchor=tk.W, pady=(0, 2))

        self._add_file_row(f1, "Nastran Exe:", self.nastran_path,
                            [("Executable", "*.exe"), ("All", "*.*")])

        mem_frame = ttk.Frame(f1)
        mem_frame.pack(fill=tk.X, pady=2)
        ttk.Label(mem_frame, text="Requested Memory:", width=18).pack(side=tk.LEFT)
        ttk.Entry(mem_frame, textvariable=self.nastran_memory, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Label(mem_frame, text="(optional, e.g. 4gb / 500mb - blank = Nastran default)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        self._add_folder_row(f1, "Scratch Folder:", self.scratch_folder)
        self._add_folder_row(f1, "Output Folder:", self.isu_output_folder)
        ttk.Label(f1, text="(every solve creates its own BaseModel_<timestamp> folder here)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(anchor=tk.W, pady=(0, 2))

        split_frame = ttk.Frame(f1)
        split_frame.pack(fill=tk.X, pady=2)
        ttk.Label(split_frame, text="Cases per Split:", width=18).pack(side=tk.LEFT)
        ttk.Entry(split_frame, textvariable=self.cases_per_split, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(split_frame, text="(blank/0 = solve as one BDF)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        par1_frame = ttk.Frame(f1)
        par1_frame.pack(fill=tk.X, pady=2)
        ttk.Label(par1_frame, text="Max Parallel Runs:", width=18).pack(side=tk.LEFT)
        ttk.Entry(par1_frame, textvariable=self.isu_max_parallel_runs, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(par1_frame, text="(Nastran runs at the same time)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        # ---- Section 2: Post Processing ----
        f2 = ttk.LabelFrame(main, text="\U0001F4E6  2. Post Processing", padding=10)
        f2.pack(fill=tk.X, pady=5, padx=10)

        self._add_file_row(f2, "Converter Exe:", self.converter_exe_path,
                            [("Executable", "*.exe"), ("All", "*.*")])
        self._add_folder_row(f2, "Example LE Folder:", self.le_example_folder)
        self._add_file_row(f2, "Pickle File:", self.le_pickle_path,
                            [("Pickle", "*.pickle *.pkl"), ("All", "*.*")])
        self._add_file_row(f2, "Load Extraction Exe:", self.le_exe_path,
                            [("Executable", "*.exe"), ("All", "*.*")])

        split2_frame = ttk.Frame(f2)
        split2_frame.pack(fill=tk.X, pady=2)
        ttk.Label(split2_frame, text="Split Into:", width=18).pack(side=tk.LEFT)
        ttk.Entry(split2_frame, textvariable=self.le_split_count, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(split2_frame, text="(number of structure groups)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        mem2_frame = ttk.Frame(f2)
        mem2_frame.pack(fill=tk.X, pady=2)
        ttk.Label(mem2_frame, text="Max Memory (GB):", width=18).pack(side=tk.LEFT)
        ttk.Entry(mem2_frame, textvariable=self.le_max_memory, width=10).pack(side=tk.LEFT, padx=5)

        ttk.Label(f2, text="(Input Excel, OP2 Folder, Output Folder H5, BDF Maneuver/Thermal and "
                            "the LE split Output Folder are all auto-wired every iteration)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(anchor=tk.W, pady=(4, 2))

        # ---- Section 3: Joint Run ----
        f3 = ttk.LabelFrame(main, text="\U0001F517  3. Joint Run", padding=10)
        f3.pack(fill=tk.X, pady=5, padx=10)

        self._add_file_row(f3, "Example Joint Input Excel:", self.jr_example_joint_excel_path,
                            [("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")])
        self._add_folder_row(f3, "Util Folder:", self.jr_util_folder)
        self._add_file_row(f3, "License Checker Exe:", self.jr_license_checker_exe_path,
                            [("Executable", "*.exe"), ("All", "*.*")])
        self._add_file_row(f3, "Load Modifying Exe:", self.jr_load_modifying_exe_path,
                            [("Executable", "*.exe"), ("All", "*.*")])
        self._add_file_row(f3, "RSS Case Finding Exe:", self.jr_rss_exe_path,
                            [("Executable", "*.exe"), ("All", "*.*")])
        self._add_file_row(f3, "Joint Run Exe:", self.jr_joint_run_exe_path,
                            [("Executable", "*.exe"), ("All", "*.*")])

        par2_frame = ttk.Frame(f3)
        par2_frame.pack(fill=tk.X, pady=2)
        ttk.Label(par2_frame, text="Max Parallel Runs:", width=18).pack(side=tk.LEFT)
        ttk.Entry(par2_frame, textvariable=self.jr_max_parallel_runs, width=10).pack(side=tk.LEFT, padx=5)

        scale_frame = ttk.Frame(f3)
        scale_frame.pack(fill=tk.X, pady=2)
        ttk.Label(scale_frame, text="LSC Scale Factor:", width=18).pack(side=tk.LEFT)
        ttk.Entry(scale_frame, textvariable=self.jr_scale_factor, width=10).pack(side=tk.LEFT, padx=5)

        ttk.Label(f3, text="(Main Input Excel, Joint Load H5 Folder and Output Folder are all "
                            "auto-wired every iteration)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(anchor=tk.W, pady=(4, 2))

        # ---- Section 4: Iteration Settings ----
        f4 = ttk.LabelFrame(main, text="\U0001F501  4. Iteration Settings", padding=10)
        f4.pack(fill=tk.X, pady=5, padx=10)

        tfh_frame = ttk.Frame(f4)
        tfh_frame.pack(fill=tk.X, pady=2)
        ttk.Label(tfh_frame, text="Target Flight Hours:", width=18).pack(side=tk.LEFT)
        ttk.Entry(tfh_frame, textvariable=self.target_flight_hours, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Label(tfh_frame, text="(a Bar Element only counts as sized once its Life >= this)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        niter_frame = ttk.Frame(f4)
        niter_frame.pack(fill=tk.X, pady=2)
        ttk.Label(niter_frame, text="Number of Iterations:", width=18).pack(side=tk.LEFT)
        ttk.Entry(niter_frame, textvariable=self.n_iterations, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Label(niter_frame, text="(H5s are deleted after every iteration except the last one)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        # ---- Section 5: Actions (RUN sits at the very bottom) ----
        f5 = ttk.LabelFrame(main, text="▶  5. Actions", padding=10)
        f5.pack(fill=tk.X, pady=5, padx=10)

        btn_row = ttk.Frame(f5)
        btn_row.pack(fill=tk.X)
        self.btn_run = ttk.Button(btn_row, text="▶  RUN", command=self.start_run, style='Accent.TButton')
        self.btn_run.pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_row, text="Clear Log", command=self.clear_log).pack(side=tk.LEFT, padx=5)
        ttk.Label(btn_row, text="(each iteration: solve -> fill t1(mm)/tfem -> Post Processing -> "
                                 "Joint Run -> sizing summary -> update Bar Property Dim1)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=5)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(f5, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=5)

        # ---- Section 6: Log ----
        f6 = ttk.LabelFrame(main, text="\U0001F4DC  6. Log", padding=10)
        f6.pack(fill=tk.BOTH, expand=True, pady=5, padx=10)

        self.log_text = scrolledtext.ScrolledText(f6, height=20, width=100, font=('Consolas', 10),
                                                    bg=_APP_COLORS['log_bg'], fg=_APP_COLORS['log_fg'],
                                                    insertbackground=_APP_COLORS['log_fg'], relief='flat',
                                                    borderwidth=0)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def _add_file_row(self, parent, label, var, filetypes):
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=2)
        ttk.Label(frame, text=label, width=18).pack(side=tk.LEFT)
        ttk.Entry(frame, textvariable=var, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame, text="Browse", command=lambda: var.set(
            filedialog.askopenfilename(filetypes=filetypes)
        )).pack(side=tk.LEFT)

    def _add_folder_row(self, parent, label, var):
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=2)
        ttk.Label(frame, text=label, width=18).pack(side=tk.LEFT)
        ttk.Entry(frame, textvariable=var, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame, text="Browse", command=lambda: var.set(
            filedialog.askdirectory()
        )).pack(side=tk.LEFT)

    def log(self, msg):
        def _do():
            self.log_text.insert(tk.END, msg + "\n")
            self.log_text.see(tk.END)
        self.root.after(0, _do)

    def clear_log(self):
        self.log_text.delete(1.0, tk.END)

    # ==================== BDF LIST OPS ====================
    def _add_bdf(self):
        paths = filedialog.askopenfilenames(filetypes=[("BDF", "*.bdf *.dat *.nas"), ("All", "*.*")])
        for p in paths:
            if p not in self.bdf_paths:
                self.bdf_paths.append(p)
                self.bdf_listbox.insert(tk.END, os.path.basename(p))

    def _remove_bdf(self):
        sel = self.bdf_listbox.curselection()
        if sel:
            idx = sel[0]
            self.bdf_listbox.delete(idx)
            del self.bdf_paths[idx]

    def _add_maneuver_bdf(self):
        files = filedialog.askopenfilenames(filetypes=[("BDF", "*.bdf *.dat *.nas"), ("All", "*.*")])
        for f in files:
            if f not in self.maneuver_bdfs:
                self.maneuver_bdfs.append(f)
                self.maneuver_listbox.insert(tk.END, os.path.basename(f))

    def _remove_maneuver_bdf(self):
        sel = self.maneuver_listbox.curselection()
        if sel:
            idx = sel[0]
            self.maneuver_listbox.delete(idx)
            del self.maneuver_bdfs[idx]

    # ==================== AUTO-WIRING / SIZING HELPERS ====================
    @staticmethod
    def _list_base_model_folders(output_base):
        if not output_base or not os.path.isdir(output_base):
            return set()
        return {
            name for name in os.listdir(output_base)
            if name.startswith('BaseModel_') and os.path.isdir(os.path.join(output_base, name))
        }

    def _move_op2_files(self, search_root, dest_folder):
        dest_abs = os.path.abspath(dest_folder)
        found = []
        for dirpath, dirnames, filenames in os.walk(search_root):
            if os.path.abspath(dirpath) == dest_abs:
                dirnames[:] = []
                continue
            for fn in filenames:
                if fn.lower().endswith('.op2'):
                    found.append(os.path.join(dirpath, fn))

        moved = 0
        for src in found:
            fn = os.path.basename(src)
            dst = os.path.join(dest_folder, fn)
            if os.path.exists(dst):
                base, ext = os.path.splitext(fn)
                i = 2
                while os.path.exists(dst):
                    dst = os.path.join(dest_folder, f"{base}_{i}{ext}")
                    i += 1
                self.log(f"    NOTE: name collision, {fn} -> {os.path.basename(dst)}")
            shutil.move(src, dst)
            moved += 1
        return moved

    @staticmethod
    def _find_stru_bdf(base_model_folder, keyword):
        matches = []
        for dirpath, _, filenames in os.walk(base_model_folder):
            for fn in filenames:
                fnl = fn.lower()
                if fnl.endswith('_stru.bdf') and keyword in fnl:
                    matches.append(os.path.join(dirpath, fn))
        return matches[0] if matches else None

    # _find_all_stru_bdfs: _find_stru_bdf'in "keyword" ile tek dosya arayan
    # halinden farklı olarak, base_model klasörü altındaki TÜM "..._stru.bdf"
    # dosyalarını (maneuver VE thermal, hepsi) döndürür - PBARL dim1
    # değerlerini DTCG_Input_Excel_Data'ya geri yazarken HER BDF'teki
    # değerler taranmalı (hangi BDF'te olduğu önemli değil, PID aynı PID'dir).
    @staticmethod
    def _find_all_stru_bdfs(base_model_folder):
        matches = []
        for dirpath, _, filenames in os.walk(base_model_folder):
            for fn in filenames:
                if fn.lower().endswith('_stru.bdf'):
                    matches.append(os.path.join(dirpath, fn))
        return matches

    # _extract_pbarl_dim1: BarPropertySolverTab'ın PBARL yazarken kullandığı
    # AYNI sabit-genişlik alan okuma yaklaşımı (PID [8:16], devam satırında
    # dim1 [8:16]) - burada sadece OKUMA için, tam bir pyNastran parse
    # gerekmediğinden hafif bir metin taraması olarak yapılıyor.
    @staticmethod
    def _extract_pbarl_dim1(bdf_path):
        """Lightweight text-scan for PBARL dim1 per PID - the same
        fixed-width field reading BarPropertySolverTab's own BDF-writing
        code uses, so a full pyNastran parse isn't needed just to read
        these back out of an already-written BDF."""
        dim1_by_pid = {}
        try:
            with open(bdf_path, 'r', encoding='latin-1') as f:
                lines = f.readlines()
        except Exception:
            return dim1_by_pid
        i = 0
        while i < len(lines):
            line = lines[i]
            if line.startswith('PBARL'):
                try:
                    pid = int(line[8:16].strip())
                except ValueError:
                    pid = None
                i += 1
                if pid is not None and i < len(lines) and (
                    lines[i].startswith('+') or lines[i].startswith('*') or
                    (lines[i][0] == ' ' and lines[i].strip() and not lines[i].strip().startswith('$'))
                ):
                    cont = lines[i]
                    try:
                        dim1_by_pid[pid] = float(cont[8:16].strip())
                    except ValueError:
                        pass
                continue
            i += 1
        return dim1_by_pid

    # _update_dtcg_t1_tfem: DTCG_Input_Excel_Data sayfasının 't1(mm)' ve
    # 'tfem' sütunlarını (Bar Property ID'ye göre eşleştirerek), solve'ün
    # gerçekten yazdığı PBARL dim1 değerleriyle doldur - böylece Joint
    # Run'a giden Excel, "istenen" değil "gerçekten çözülen" kalınlığı yansıtır.
    def _update_dtcg_t1_tfem(self, excel_path, dim1_by_pid):
        """Fill DTCG_Input_Excel_Data's 't1(mm)'/'tfem' columns (by 'Bar
        Property ID') from dim1_by_pid. t2(mm) and every other column are
        never touched."""
        keep_vba = excel_path.lower().endswith(('.xlsm', '.xltm', '.xlam'))
        wb = openpyxl.load_workbook(excel_path, keep_vba=keep_vba)
        ws = None
        for s in wb.worksheets:
            sl = s.title.strip().lower().replace('_', '').replace(' ', '')
            if sl == 'dtcginputexceldata':
                ws = s
                break
        if ws is None:
            self.log("  WARNING: no 'DTCG_Input_Excel_Data' sheet found - t1(mm)/tfem not updated.")
            return 0

        header = {}
        for cell in next(ws.iter_rows(min_row=1, max_row=1)):
            if cell.value:
                norm = str(cell.value).strip().lower().replace('\n', '').replace(' ', '')
                header[norm] = cell.column
        pid_col = header.get('barpropertyid')
        t1_col = header.get('t1(mm)')
        tfem_col = header.get('tfem')
        if not pid_col or (not t1_col and not tfem_col):
            self.log("  WARNING: 'Bar Property ID' / 't1(mm)' / 'tfem' columns not found in "
                      "DTCG_Input_Excel_Data - t1(mm)/tfem not updated.")
            return 0

        n_updated = 0
        for row in range(2, ws.max_row + 1):
            pid_cell = ws.cell(row=row, column=pid_col)
            if pid_cell.value is None:
                continue
            try:
                pid = int(pid_cell.value)
            except (ValueError, TypeError):
                continue
            if pid in dim1_by_pid:
                val = dim1_by_pid[pid]
                if t1_col:
                    ws.cell(row=row, column=t1_col).value = val
                if tfem_col:
                    ws.cell(row=row, column=tfem_col).value = val
                n_updated += 1

        wb.save(excel_path)
        self.log(f"  Updated t1(mm)/tfem for {n_updated} DTCG_Input_Excel_Data row(s) "
                  f"from the solved BDF")
        return n_updated

    # _update_bar_property_dim1: bir sonraki iterasyon için "Bar Property"
    # sayfasının Dim1 sütununu (B sütunu, PID A sütununda - load_properties()
    # bunu başlığa göre değil POZİSYONA göre okuduğu için buradaki yazma da
    # pozisyonel) prop_map'teki ({pid: yeni_dim1}) değerlerle güncelle.
    # Dim2'ye HİÇ dokunulmaz.
    def _update_bar_property_dim1(self, excel_path, prop_map):
        """Overwrite the 'Bar Property' sheet's Dim1 column (column B,
        PID in column A - load_properties() reads it positionally, not by
        header) from prop_map ({pid: new_dim1}). Dim2 is never touched."""
        keep_vba = excel_path.lower().endswith(('.xlsm', '.xltm', '.xlam'))
        wb = openpyxl.load_workbook(excel_path, keep_vba=keep_vba)
        ws = None
        for s in wb.worksheets:
            sl = s.title.strip().lower().replace('_', '').replace(' ', '').replace('-', '')
            if sl == 'barproperty':
                ws = s
                break
        if ws is None:
            self.log("  WARNING: no 'Bar Property' sheet found - thicknesses not updated "
                      "for the next iteration.")
            return 0

        n_updated = 0
        for row in range(2, ws.max_row + 1):
            pid_cell = ws.cell(row=row, column=1)
            if pid_cell.value is None:
                continue
            try:
                pid = int(pid_cell.value)
            except (ValueError, TypeError):
                continue
            if pid in prop_map:
                ws.cell(row=row, column=2).value = prop_map[pid]
                n_updated += 1

        wb.save(excel_path)
        self.log(f"  Updated Dim1 for {n_updated} Bar Property row(s) from the previous "
                  f"iteration's sizing summary")
        return n_updated

    # _build_flight_hours_summary: referans Summary_Creating.py script'inin
    # flight_hours_summary() fonksiyonunun Python/pandas'a taşınmış hali.
    # Amaç: tüm Set'lerin "Nasgro Final Summary" CSV'lerini birleştirip,
    # her bar elemanı/property için "HEDEF uçuş saatini karşılayan, mümkün
    # olan EN İNCE (en hafif) kalınlığı" bulmak - bu, bir sonraki
    # iterasyonun Bar Property Dim1 değeri olarak kullanılacak.
    def _build_flight_hours_summary(self, joint_run_folder, wanted_flight_hours):
        """Port of the reference Summary_Creating.py's flight_hours_summary():
        every '...Nasgro Final Summary...' csv under joint_run_folder is
        concatenated (Structure_Name/Fastened_Parts taken from the csv's
        path relative to joint_run_folder - 1st/2nd path component), then
        DT_FULL_SUMMARY / DT_SIZING_BAR_SUMMARY / DT_SIZING_PROP_SUMMARY /
        DT_LIFE_CHECK_BAR_SUMMARY / DT_LIFE_CHECK_PROP_SUMMARY.csv are
        written into joint_run_folder. Returns (df_prop_sum, t_col) or
        (None, None) if no summary csvs were found or the expected
        columns are missing."""
        self.log("\n" + "=" * 70)
        self.log("BUILDING FLIGHT HOURS SUMMARY")
        self.log("=" * 70)

        # Joint Run klasörü altındaki (her Set'in kendi alt klasöründe)
        # TÜM "Nasgro Final Summary" CSV'lerini bul.
        csv_files = []
        for dirpath, _, filenames in os.walk(joint_run_folder):
            for fn in filenames:
                if 'nasgro final summary' in fn.lower():
                    csv_files.append(os.path.join(dirpath, fn))
        self.log(f"  Found {len(csv_files)} 'Nasgro Final Summary' csv file(s) under {joint_run_folder}")
        if not csv_files:
            return None, None

        # Her CSV'nin joint_run_folder'a göre RELATİF yolundan Structure_Name
        # (1. klasör - yani Set adı) ve Fastened_Parts (2. klasör, varsa)
        # çıkarılıp CSV'ye ek sütun olarak eklenir - böylece birleştirilmiş
        # tabloda her satırın hangi Set/parçadan geldiği izlenebilir.
        frames = []
        for each_csv in csv_files:
            rel = os.path.relpath(each_csv, joint_run_folder)
            parts = rel.split(os.sep)
            structure_name = parts[0] if parts else ''
            fastened_part = parts[1] if len(parts) >= 3 else ''
            try:
                df_now = pd.read_csv(each_csv, sep=',')
            except Exception as e:
                self.log(f"    WARNING: could not read {each_csv}: {e}")
                continue
            df_now['Structure_Name'] = structure_name
            df_now['Fastened_Parts'] = fastened_part
            frames.append(df_now)

        if not frames:
            self.log("  WARNING: none of the 'Nasgro Final Summary' csvs could be read.")
            return None, None

        # TÜM Set'lerin satırlarını tek bir "DT_FULL_SUMMARY" tablosunda birleştir.
        full_summary = pd.concat(frames, ignore_index=True)
        full_summary.to_csv(os.path.join(joint_run_folder, 'DT_FULL_SUMMARY.csv'),
                             encoding='utf-8', index=False)

        t_col = next((c for c in full_summary.columns
                      if str(c).strip().replace('\n', '').replace(' ', '').lower() == 't(mm)'), None)
        if t_col is None or 'Life' not in full_summary.columns or \
                'Bar Element ID' not in full_summary.columns or \
                'Bar Property ID' not in full_summary.columns:
            self.log("  WARNING: expected 'Life' / 't(mm)' / 'Bar Element ID' / 'Bar Property ID' "
                      "columns not found in the Nasgro Final Summary csvs - cannot build "
                      "sizing/life-check summaries.")
            return None, None

        # ---- Boyutlandırma (sizing) mantığı ----
        # 1. Önce SADECE hedef uçuş saatini (Life >= wanted_flight_hours)
        #    KARŞILAYAN satırları filtrele.
        # 2. Bu satırları kalınlığa (t_col) göre KÜÇÜKTEN BÜYÜĞE sırala,
        #    her eleman için (Bar Element ID [+ Type]) İLK (yani EN İNCE,
        #    hedefi hâlâ karşılayan) satırı tut - drop_duplicates varsayılan
        #    olarak ilk görüleni tutar.
        filtered_for_life = full_summary[full_summary['Life'] >= wanted_flight_hours]
        filtered_sorted = filtered_for_life.sort_values(by=t_col, ascending=True)
        dedup_cols = ['Bar Element ID', 'Bar Element Type'] if 'Bar Element Type' in full_summary.columns \
            else ['Bar Element ID']
        filtered_sorted_rmd = filtered_sorted.drop_duplicates(subset=dedup_cols)
        filtered_sorted_rmd = filtered_sorted_rmd.sort_values(by=t_col, ascending=False)

        # 3. Hedefi karşılayan (en ince) satırları, TÜM satırların kalınlığa
        #    göre büyükten küçüğe sıralanmış haliyle birleştir, sonra aynı
        #    elemanlar için yine drop_duplicates uygula. Bu sırada
        #    filtered_sorted_rmd (hedefi karşılayan, zaten ince) satırlar
        #    ÖNCE geldiği için, bir eleman hedefi karşılıyorsa o SEÇİLİR;
        #    hiçbir satır hedefi karşılamıyorsa (filtered_sorted_rmd'de o
        #    eleman için satır yoksa) full_summary_th_sort'taki EN KALIN
        #    (en güvenli/muhafazakâr) satır kullanılır - yani "hedefi
        #    karşılayamayan elemanlar için en kalın (en güvenli) değere geri dön" davranışı.
        full_summary_th_sort = full_summary.sort_values(by=t_col, ascending=False)
        df_bar_sum = pd.concat([filtered_sorted_rmd, full_summary_th_sort], ignore_index=True)
        df_bar_sum = df_bar_sum.drop_duplicates(subset=dedup_cols)
        df_bar_sum = df_bar_sum.sort_values(by=t_col, ascending=False)
        df_bar_sum = df_bar_sum.drop_duplicates(subset=['Bar Element ID'])
        df_bar_sum.to_csv(os.path.join(joint_run_folder, 'DT_SIZING_BAR_SUMMARY.csv'),
                           encoding='utf-8', index=False)

        # Eleman bazlı özeti (df_bar_sum), her PROPERTY için en kalın (en
        # muhafazakâr/yönetici) satırı tutarak Property bazlı özete indir -
        # bir Property'ye bağlı birden fazla eleman varsa, hepsini güvenle
        # kapsayan en kalın değer seçilir. Bu df_prop_sum, bir sonraki
        # iterasyonun Bar Property Dim1 değerlerinin KAYNAĞIDIR.
        df_prop_sum = df_bar_sum.sort_values(by=t_col, ascending=False)
        df_prop_sum = df_prop_sum.drop_duplicates(subset=['Bar Property ID'])
        df_prop_sum.to_csv(os.path.join(joint_run_folder, 'DT_SIZING_PROP_SUMMARY.csv'),
                            encoding='utf-8', index=False)

        # ---- Ömür kontrolü (life check) özetleri - sadece raporlama,
        # sizing kararı için kullanılmaz: en KISA ömürlü satırı öne çıkarır. ----
        life_check_summary = full_summary.sort_values(by='Life', ascending=True)
        life_check_summary.drop_duplicates(subset=['Bar Element ID']).to_csv(
            os.path.join(joint_run_folder, 'DT_LIFE_CHECK_BAR_SUMMARY.csv'), encoding='utf-8', index=False)
        life_check_summary.drop_duplicates(subset=['Bar Property ID']).to_csv(
            os.path.join(joint_run_folder, 'DT_LIFE_CHECK_PROP_SUMMARY.csv'), encoding='utf-8', index=False)

        self.log(f"  DT_FULL_SUMMARY: {len(full_summary)} rows, DT_SIZING_BAR_SUMMARY: "
                  f"{len(df_bar_sum)} rows, DT_SIZING_PROP_SUMMARY: {len(df_prop_sum)} rows")
        return df_prop_sum, t_col

    # _delete_h5_files: verilen klasör altındaki (rekürsif) TÜM .h5
    # dosyalarını sil - son iterasyon DIŞINDAKİ her iterasyonun Load
    # Extraction/Joint Run sonuçlarından disk yer açmak için çağrılır
    # (H5'ler büyük olabilir ve ara iterasyonlarda saklanmasına gerek yok).
    @staticmethod
    def _delete_h5_files(folder):
        n = 0
        for dirpath, _, filenames in os.walk(folder):
            for fn in filenames:
                if fn.lower().endswith('.h5'):
                    try:
                        os.remove(os.path.join(dirpath, fn))
                        n += 1
                    except OSError:
                        pass
        return n

    def _wait_for_exe_to_finish(self, exe_name, poll_seconds=30, max_wait_seconds=12 * 3600):
        """Post Processing's loadext.bat runs are fire-and-forget (see
        PostProcessingTab._run_tfx_batch_core), so there's no handle to
        wait on - poll `tasklist` by image name instead."""
        if not exe_name:
            return
        self.log(f"\n  Waiting for every '{exe_name}' process to finish before starting Joint Run...")
        waited = 0
        last_log = 0
        while waited < max_wait_seconds:
            try:
                result = subprocess.run(
                    ["tasklist", "/FI", f"IMAGENAME eq {exe_name}"],
                    capture_output=True, text=True
                )
                running = exe_name.lower() in result.stdout.lower()
            except Exception as e:
                self.log(f"  WARNING: could not check running processes ({e}) - proceeding without waiting.")
                return
            if not running:
                self.log(f"  No more '{exe_name}' processes running (waited {waited}s).")
                return
            time.sleep(poll_seconds)
            waited += poll_seconds
            if waited - last_log >= 300:
                self.log(f"    ...still waiting on '{exe_name}' ({waited}s elapsed)")
                last_log = waited
        self.log(f"  WARNING: gave up waiting for '{exe_name}' after {waited}s - proceeding anyway.")

    # ==================== RUN (ITERATIVE FULL PIPELINE) ====================
    # start_run: tüm iterasyon zincirinin (N kere tekrarlanan Understanding
    # Structure Type -> Post Processing -> Joint Run) girdilerini önceden
    # doğrular - target_flight_hours ve n_iterations dahil - sonra
    # gerçek döngüyü (_run_pipeline) arka plan thread'inde başlatır.
    def start_run(self):
        if not self.bdf_paths:
            messagebox.showerror("Error", "Add at least one All BDF File")
            return
        if not self.master_input_excel_path.get().strip() or not os.path.exists(self.master_input_excel_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Master Input Excel")
            return
        if not self.nastran_path.get().strip() or not os.path.exists(self.nastran_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Nastran Exe")
            return
        if not self.isu_output_folder.get().strip():
            messagebox.showerror("Error", "Select an Output Folder for Understanding Structure Type")
            return
        if not self.converter_exe_path.get().strip() or not os.path.exists(self.converter_exe_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Converter Exe")
            return
        if not self.le_example_folder.get().strip() or not os.path.isdir(self.le_example_folder.get().strip()):
            messagebox.showerror("Error", "Select a valid Example LE Folder")
            return
        if not self.le_exe_path.get().strip() or not os.path.exists(self.le_exe_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Load Extraction Exe")
            return
        if not self.jr_example_joint_excel_path.get().strip() or not os.path.exists(self.jr_example_joint_excel_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Example Joint Input Excel")
            return
        if not self.jr_util_folder.get().strip() or not os.path.isdir(self.jr_util_folder.get().strip()):
            messagebox.showerror("Error", "Select a valid Util Folder")
            return
        if not self.jr_license_checker_exe_path.get().strip() or not os.path.exists(self.jr_license_checker_exe_path.get().strip()):
            messagebox.showerror("Error", "Select a valid License Checker Exe")
            return
        if not self.jr_load_modifying_exe_path.get().strip() or not os.path.exists(self.jr_load_modifying_exe_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Load Modifying Exe")
            return
        if not self.jr_rss_exe_path.get().strip() or not os.path.exists(self.jr_rss_exe_path.get().strip()):
            messagebox.showerror("Error", "Select a valid RSS Case Finding Exe")
            return
        if not self.jr_joint_run_exe_path.get().strip() or not os.path.exists(self.jr_joint_run_exe_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Joint Run Exe")
            return
        try:
            if int(float(self.le_split_count.get())) <= 0:
                raise ValueError
        except (ValueError, TypeError):
            messagebox.showerror("Error", "Split Into must be a positive number")
            return
        try:
            if float(self.le_max_memory.get()) <= 0:
                raise ValueError
        except (ValueError, TypeError):
            messagebox.showerror("Error", "Post Processing Max Memory must be a positive number")
            return
        try:
            if int(float(self.jr_max_parallel_runs.get())) <= 0:
                raise ValueError
        except (ValueError, TypeError):
            messagebox.showerror("Error", "Joint Run Max Parallel Runs must be a positive number")
            return
        try:
            float(self.jr_scale_factor.get())
        except (ValueError, TypeError):
            messagebox.showerror("Error", "LSC Scale Factor must be a number")
            return
        try:
            float(self.target_flight_hours.get())
        except (ValueError, TypeError):
            messagebox.showerror("Error", "Target Flight Hours must be a number")
            return
        try:
            if int(float(self.n_iterations.get())) <= 0:
                raise ValueError
        except (ValueError, TypeError):
            messagebox.showerror("Error", "Number of Iterations must be a positive number")
            return

        self.is_running = True
        self.btn_run.config(state=tk.DISABLED)
        self.root.after(0, lambda: self.progress_var.set(0))
        threading.Thread(target=self._run_pipeline, daemon=True).start()

    def _run_pipeline(self):
        t_start = time.time()
        try:
            self.log("\n" + "#" * 70)
            self.log("ITERATIVE FULL RUN")
            self.log("#" * 70)

            n_iterations = int(float(self.n_iterations.get()))
            target_life = float(self.target_flight_hours.get())
            output_base = self.isu_output_folder.get().strip()
            os.makedirs(output_base, exist_ok=True)

            # Kullanıcının Master Input Excel'i HİÇ DOKUNULMADAN kalır -
            # tüm iterasyonlar, bu özel "working excel" ÇALIŞMA KOPYASI
            # üzerinde okuma/yazma yapar.
            working_excel = os.path.join(output_base, "_IterativeFullRun_Working_Input.xlsx")
            shutil.copy2(self.master_input_excel_path.get().strip(), working_excel)
            self.log(f"  Working copy of the Master Input Excel: {working_excel}")
            self.log(f"  Target flight hours: {target_life}, Iterations: {n_iterations}")

            prop_summary_df = None  # önceki iterasyonun boyutlandırma özeti (ilk iterasyonda yok)
            t_col = None

            for iteration in range(1, n_iterations + 1):
                iter_t_start = time.time()
                self.log("\n" + "#" * 70)
                self.log(f"ITERATION {iteration}/{n_iterations}")
                self.log("#" * 70)
                self.root.after(0, lambda p=(iteration - 1) / n_iterations * 100: self.progress_var.set(p))

                # Bu, DÖNGÜNÜN GERİ BESLEME (feedback) ADIMI: ilk iterasyon
                # HARİÇ, bir önceki iterasyonun boyutlandırma özetindeki
                # (prop_summary_df) kalınlık değerleri, bu iterasyonun
                # çözümünden ÖNCE working_excel'in Bar Property Dim1'ine yazılır.
                if iteration > 1 and prop_summary_df is not None and t_col is not None:
                    prop_map = dict(zip(prop_summary_df['Bar Property ID'], prop_summary_df[t_col]))
                    self._update_bar_property_dim1(working_excel, prop_map)

                # ---- STAGE A: UNDERSTANDING STRUCTURE TYPE ----
                # FullRunTab'daki aynı mantık: t3 nesnesinin alanlarını
                # doldur ve _run_solve()'u çağır - TEK FARK, Property/Element
                # Excel'in HER İKİSİNİN de aynı working_excel'i göstermesi
                # (Master Input Excel tek dosyada hepsini içerdiği için).
                self.log(f"\n--- Iteration {iteration}: Understanding Structure Type ---")
                t3 = self.tab3_tool
                t3.bdf_paths = list(self.bdf_paths)
                t3.maneuver_bdfs = list(self.maneuver_bdfs)
                t3.property_excel_path.set(working_excel)
                t3.element_excel_path.set(working_excel)
                t3.nastran_path.set(self.nastran_path.get())
                t3.nastran_memory.set(self.nastran_memory.get())
                t3.scratch_folder.set(self.scratch_folder.get())
                t3.output_folder.set(output_base)
                t3.cases_per_split.set(self.cases_per_split.get())
                t3.max_parallel_runs.set(self.isu_max_parallel_runs.get())
                self.root.after(0, self._refresh_tab3_listboxes)

                t3.load_properties()
                if not t3.properties_loaded:
                    self.log("\nSTOPPED: Property Excel failed to load - see the Understanding "
                              "Structure Type log/tab above for details.")
                    return
                t3.load_element_ids()

                before_folders = self._list_base_model_folders(output_base)
                t3._run_solve()
                after_folders = self._list_base_model_folders(output_base)
                new_folders = after_folders - before_folders
                if len(new_folders) == 1:
                    run_folder = os.path.join(output_base, new_folders.pop())
                elif new_folders:
                    candidates = [os.path.join(output_base, n) for n in new_folders]
                    run_folder = max(candidates, key=os.path.getmtime)
                else:
                    run_folder = None

                if not run_folder:
                    self.log(f"\nSTOPPED (iteration {iteration}): no new BaseModel_<timestamp> "
                              "folder appeared - the solve likely failed. See the Understanding "
                              "Structure Type log/tab above.")
                    return
                self.log(f"  Base model run folder: {run_folder}")

                base_model_folder = os.path.join(run_folder, "_base_model")
                op2_count_check = sum(
                    1 for dirpath, _, filenames in os.walk(base_model_folder) for fn in filenames
                    if fn.lower().endswith('.op2')
                )
                if op2_count_check == 0:
                    self.log(f"\nSTOPPED (iteration {iteration}): no .op2 files were produced. "
                              "See the Understanding Structure Type log/tab above.")
                    return
                self.log(f"  {op2_count_check} .op2 file(s) found")

                # Bu iterasyonda ÇÖZÜLEN her BDF'teki (maneuver + thermal,
                # hepsi) PBARL dim1 değerlerini oku ve DTCG_Input_Excel_
                # Data'nın t1(mm)/tfem sütunlarına geri yaz - Joint Run'a
                # giden Excel, "istenen" değeri değil GERÇEKTEN ÇÖZÜLEN
                # kalınlığı yansıtsın.
                stru_bdfs = self._find_all_stru_bdfs(base_model_folder)
                dim1_by_pid = {}
                for p in stru_bdfs:
                    dim1_by_pid.update(self._extract_pbarl_dim1(p))
                self._update_dtcg_t1_tfem(working_excel, dim1_by_pid)

                # Bu iterasyonun working_excel durumunu, o iterasyonun kendi
                # _base_model klasörüne anlık görüntü (snapshot) olarak
                # kaydet - sonradan "bu iterasyonda hangi kalınlıklar
                # kullanıldı" diye geri bakılabilsin.
                iter_snapshot = os.path.join(base_model_folder, f"Main_Input_Iteration_{iteration}.xlsx")
                shutil.copy2(working_excel, iter_snapshot)
                self.log(f"  Working excel snapshot saved: {iter_snapshot}")

                # ---- STAGE B: POST PROCESSING ----
                self.log(f"\n--- Iteration {iteration}: Post Processing ---")
                op2_results_dir = os.path.join(base_model_folder, "OP2_Results")
                h5_results_dir = os.path.join(base_model_folder, "H5_Results")
                le_results_dir = os.path.join(base_model_folder, "Load_Extraction_Results")
                os.makedirs(op2_results_dir, exist_ok=True)
                os.makedirs(h5_results_dir, exist_ok=True)
                os.makedirs(le_results_dir, exist_ok=True)

                n_moved = self._move_op2_files(base_model_folder, op2_results_dir)
                self.log(f"  Moved {n_moved} .op2 file(s) into {op2_results_dir}")

                maneuver_stru = self._find_stru_bdf(base_model_folder, "maneuver")
                thermal_stru = self._find_stru_bdf(base_model_folder, "thermal")
                if not maneuver_stru:
                    self.log("  WARNING: no '..._stru.bdf' with 'maneuver' in its name found.")
                if not thermal_stru:
                    self.log("  WARNING: no '..._stru.bdf' with 'thermal' in its name found.")

                t4 = self.tab4_tool
                t4.op2_folder.set(op2_results_dir)
                t4.converter_exe_path.set(self.converter_exe_path.get())
                t4.output_folder.set(h5_results_dir)
                t4.le_input_excel_path.set(working_excel)
                t4.le_example_folder.set(self.le_example_folder.get())
                t4.le_maneuver_bdf_path.set(maneuver_stru or "")
                t4.le_thermal_bdf_path.set(thermal_stru or "")
                t4.le_pickle_path.set(self.le_pickle_path.get())
                t4.le_split_count.set(self.le_split_count.get())
                t4.le_output_folder.set(le_results_dir)
                t4.le_exe_path.set(self.le_exe_path.get())
                t4.le_max_memory.set(self.le_max_memory.get())

                t4._run_pipeline()

                group_folders_found = [
                    name for name in (os.listdir(le_results_dir) if os.path.isdir(le_results_dir) else [])
                    if os.path.isdir(os.path.join(le_results_dir, name)) and name.lower().startswith('loadextraction_')
                ]
                if not group_folders_found:
                    self.log(f"\nSTOPPED (iteration {iteration}): no LoadExtraction_NN folders "
                              "were created. See the Post Processing log/tab above.")
                    return
                self.log(f"  {len(group_folders_found)} LoadExtraction_NN group folder(s) found")

                # t4._run_pipeline() döndüğünde loadext.bat'lar henüz
                # BAŞLATILMIŞ olabilir (fire-and-forget) - Joint Run'ın H5
                # taraması bunların çıktısına bağımlı olduğundan, hepsinin
                # gerçekten bitmesini bekle.
                self._wait_for_exe_to_finish(os.path.basename(self.le_exe_path.get().strip()))

                # ---- STAGE C: JOINT RUN ----
                # FullRunTab ile aynı mantık: t5 nesnesinin alanlarını doldur
                # (Main Input Excel = bu iterasyonun working_excel'i) ve
                # t5._run_pipeline()'ı çağır - bu metod zaten Joint Run exe
                # havuzunun tamamının bitmesini bekler.
                self.log(f"\n--- Iteration {iteration}: Joint Run ---")
                joint_run_dir = os.path.join(base_model_folder, "Joint_Run")
                os.makedirs(joint_run_dir, exist_ok=True)

                t5 = self.tab5_tool
                t5.h5_folder.set(le_results_dir)
                t5.main_input_excel_path.set(working_excel)
                t5.example_joint_excel_path.set(self.jr_example_joint_excel_path.get())
                t5.util_folder.set(self.jr_util_folder.get())
                t5.license_checker_exe_path.set(self.jr_license_checker_exe_path.get())
                t5.load_modifying_exe_path.set(self.jr_load_modifying_exe_path.get())
                t5.rss_exe_path.set(self.jr_rss_exe_path.get())
                t5.joint_run_exe_path.set(self.jr_joint_run_exe_path.get())
                t5.output_folder.set(joint_run_dir)
                t5.max_parallel_runs.set(self.jr_max_parallel_runs.get())
                t5.scale_factor.set(self.jr_scale_factor.get())

                t5._run_pipeline()

                # ---- STAGE D: BOYUTLANDIRMA ÖZETİ ----
                # Bu iterasyonun Joint Run sonuçlarından (Nasgro Final
                # Summary CSV'leri) prop_summary_df'i üret - bu değişken bir
                # SONRAKİ iterasyonun başındaki geri besleme adımında
                # (yukarıda) kullanılacak.
                prop_summary_df, t_col = self._build_flight_hours_summary(joint_run_dir, target_life)
                if prop_summary_df is None:
                    self.log(f"\nSTOPPED (iteration {iteration}): no sizing summary could be built "
                              "- not continuing to further iterations. See the log above.")
                    return

                # SON iterasyon DIŞINDAKİ her iterasyonda: bu iterasyonun H5
                # dosyaları (hem Load Extraction hem Joint Run) artık
                # gerekli değil (sonraki iterasyon kendi H5'lerini üretecek)
                # - disk yer açmak için sil. SON iterasyonun H5'leri
                # kullanıcı inceleyebilsin diye KORUNUR.
                if iteration < n_iterations:
                    n_del_le = self._delete_h5_files(le_results_dir)
                    n_del_jr = self._delete_h5_files(joint_run_dir)
                    self.log(f"  Deleted {n_del_le} H5 file(s) from Load Extraction Results and "
                              f"{n_del_jr} from Joint Run (not the last iteration)")

                self.log(f"\nIteration {iteration}/{n_iterations} complete "
                          f"({time.time() - iter_t_start:.1f}s)")

            self.root.after(0, lambda: self.progress_var.set(100))
            self.log(f"\n{'#' * 70}")
            self.log(f"ITERATIVE FULL RUN COMPLETE (total {time.time() - t_start:.1f}s)")
            self.log(f"{'#' * 70}")

        except Exception as e:
            self.log(f"\nERROR after {time.time() - t_start:.1f}s: {e}")
            import traceback
            self.log(traceback.format_exc())

        finally:
            self.is_running = False
            self.root.after(0, lambda: self.btn_run.config(state=tk.NORMAL))

    # _refresh_tab3_listboxes: t3.bdf_paths/maneuver_bdfs listelerine
    # yukarıda doğrudan yazıldığı için, o sekmenin GUI listbox'ları
    # otomatik güncellenmez - burada manuel senkronize edilir.
    def _refresh_tab3_listboxes(self):
        t3 = self.tab3_tool
        t3.bdf_listbox.delete(0, tk.END)
        for p in t3.bdf_paths:
            t3.bdf_listbox.insert(tk.END, os.path.basename(p))
        t3.maneuver_listbox.delete(0, tk.END)
        for p in t3.maneuver_bdfs:
            t3.maneuver_listbox.insert(tk.END, os.path.basename(p))


# ============================================================================
# VisualizationResultsTab  --  "Visualization Results" sekmesi (Tab 8)
# ============================================================================
# Kullanıcının verdiği bir BDF dosyasını (HyperMesh benzeri) 3B olarak çizer
# ve Iterative Full Run'ın ürettiği DT_SIZING_BAR_SUMMARY.csv /
# DT_SIZING_PROP_SUMMARY.csv dosyalarındaki sonuçları (t, d, Material, Life)
# bar elemanlarının üzerine renk kodlayarak gösterir - "Bar Property ID" veya
# "Bar Element ID" bazında gruplama seçilebilir. Diğer hiçbir sekmenin kodunu
# değiştirmez, bağımsız ve kendi kendine yeten yeni bir sekmedir.
class VisualizationResultsTab:
    def __init__(self, parent_frame, root):
        self.root = root
        self.parent_frame = parent_frame

        # ---- Girdi yolları ----
        self.bdf_path = tk.StringVar()        # görselleştirilecek BDF dosyası
        self.summary_folder = tk.StringVar()  # DT_SIZING_*_SUMMARY.csv'lerin aranacağı klasör

        # ---- Çizim ayarları ----
        self.group_by = tk.StringVar(value="Bar Property ID")  # "Bar Property ID" | "Bar Element ID"
        self.color_by = tk.StringVar(value="Life")              # yüklenen CSV'nin sütunlarından biri - sabit bir liste değil, bkz. _refresh_metric_options

        # ---- Renk skalası (contour) ayarları ----
        # scale_auto açıkken Min/Max veriden otomatik hesaplanır (eski
        # davranış); kullanıcı kapatıp kendi Min/Max/Levels değerlerini
        # girerse, renk skalası o aralığa ve o kadar banda (kademeye) göre çizilir.
        self.scale_auto = tk.BooleanVar(value=True)
        self.scale_min = tk.StringVar(value="")
        self.scale_max = tk.StringVar(value="")
        self.scale_levels = tk.StringVar(value="10")  # kaç ayrı renk bandına (kademeye) bölüneceği

        # ---- Görünüm ayarı ----
        # True iken 2B yan grafik gizlenir ve 3B FEM görünümü tüm çizim
        # alanını kaplar - kullanıcı modeli daha büyük görüp rahatça
        # döndürüp/yakınlaştırıp inceleyebilsin diye.
        self.maximize_3d = tk.BooleanVar(value=False)

        # ---- BDF'ten çıkarılan mesh verisi ----
        self.bdf_model = None
        self.node_coords = {}      # nid -> (x, y, z)
        self.bar_elements = []     # [{'eid', 'pid', 'n1', 'n2'}, ...]  (CBAR/CBEAM)
        self.shell_elements = []   # [{'eid', 'pid', 'nodes': [...]}, ...]  (CQUAD4/CTRIA3/...)

        # ---- CSV'den okunan sonuç verisi ----
        self.bar_summary_df = None   # index: Bar Element ID
        self.prop_summary_df = None  # index: Bar Property ID

        self.setup_ui()

    # ==================== GUI ====================
    def setup_ui(self):
        main = ttk.Frame(self.parent_frame, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main, text="Visualization Results", style='Header.TLabel').pack(pady=(0, 4))
        ttk.Label(main, text="Load a BDF + a sizing summary folder, then color the FEM mesh by any "
                              "column from that CSV, grouped by Bar Property ID or Bar Element ID",
                 style='Subtitle.TLabel').pack(pady=(0, 8))

        # matplotlib kurulu değilse, sekmenin geri kalanını hiç kurmadan
        # sadece bir uyarı göster - uygulamanın diğer 7 sekmesi yine de
        # tamamen normal çalışmaya devam eder.
        if not _MATPLOTLIB_AVAILABLE:
            ttk.Label(main, text="matplotlib bulunamadı. Bu sekmeyi kullanmak için önce şunu çalıştırın:\n"
                                  "    pip install matplotlib",
                     foreground=_APP_COLORS.get('danger', 'red'),
                     font=('Segoe UI', 10, 'bold'), justify=tk.LEFT).pack(pady=30)
            return

        # ---- Section 1: Model & Sizing Data ----
        f1 = ttk.LabelFrame(main, text="\U0001F4C1  1. Model & Sizing Data", padding=10)
        f1.pack(fill=tk.X, pady=5)

        bdf_frame = ttk.Frame(f1)
        bdf_frame.pack(fill=tk.X, pady=2)
        ttk.Label(bdf_frame, text="BDF File:", width=16).pack(side=tk.LEFT)
        ttk.Entry(bdf_frame, textvariable=self.bdf_path, width=55).pack(side=tk.LEFT, padx=5)
        ttk.Button(bdf_frame, text="Browse", command=self._browse_bdf).pack(side=tk.LEFT, padx=2)
        ttk.Button(bdf_frame, text="Load BDF", command=self.load_bdf).pack(side=tk.LEFT, padx=2)
        self.bdf_status = ttk.Label(f1, text="No BDF loaded", foreground="gray")
        self.bdf_status.pack(anchor=tk.W, pady=(0, 6))

        sum_frame = ttk.Frame(f1)
        sum_frame.pack(fill=tk.X, pady=2)
        ttk.Label(sum_frame, text="Summary Folder:", width=16).pack(side=tk.LEFT)
        ttk.Entry(sum_frame, textvariable=self.summary_folder, width=55).pack(side=tk.LEFT, padx=5)
        ttk.Button(sum_frame, text="Browse", command=self._browse_summary_folder).pack(side=tk.LEFT, padx=2)
        ttk.Button(sum_frame, text="Load Summary", command=self.load_summary).pack(side=tk.LEFT, padx=2)
        ttk.Label(f1, text="(recursively searches for DT_SIZING_BAR_SUMMARY.csv / DT_SIZING_PROP_SUMMARY.csv "
                            "- written by Iterative Full Run's sizing step into the Joint Run output folder; "
                            "falls back to DT_FULL_SUMMARY.csv if those aren't found)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(anchor=tk.W, pady=(0, 2))
        self.summary_status = ttk.Label(f1, text="No summary loaded", foreground="gray")
        self.summary_status.pack(anchor=tk.W)

        # ---- Section 2: Plot Settings ----
        f2 = ttk.LabelFrame(main, text="\U0001F3A8  2. Plot Settings", padding=10)
        f2.pack(fill=tk.X, pady=5)

        gb_frame = ttk.Frame(f2)
        gb_frame.pack(fill=tk.X, pady=2)
        ttk.Label(gb_frame, text="Group By:", width=16).pack(side=tk.LEFT)
        ttk.Radiobutton(gb_frame, text="Bar Property ID", variable=self.group_by,
                        value="Bar Property ID", command=self._on_group_by_changed).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(gb_frame, text="Bar Element ID", variable=self.group_by,
                        value="Bar Element ID", command=self._on_group_by_changed).pack(side=tk.LEFT, padx=5)

        cb_frame = ttk.Frame(f2)
        cb_frame.pack(fill=tk.X, pady=2)
        ttk.Label(cb_frame, text="Color By:", width=16).pack(side=tk.LEFT)
        # Bu a\u00E7\u0131l\u0131r listenin i\u00E7eri\u011Fi SAB\u0130T de\u011Fil - "Bar Element ID"/"Bar
        # Property ID" d\u0131\u015F\u0131ndaki T\u00DCM s\u00FCtunlar, y\u00FCklenen CSV neyse ona g\u00F6re
        # _refresh_metric_options() taraf\u0131ndan doldurulur (bkz. load_summary
        # ve _on_group_by_changed). B\u00F6ylece farkl\u0131 bir CSV format\u0131
        # verildi\u011Finde (s\u00FCtun adlar\u0131 farkl\u0131 olsa bile) otomatik uyum sa\u011Flan\u0131r.
        self.color_combo = ttk.Combobox(cb_frame, textvariable=self.color_by, state='readonly', width=22, values=[])
        self.color_combo.pack(side=tk.LEFT, padx=5)
        self.color_combo.bind('<<ComboboxSelected>>', lambda e: self._on_setting_changed())
        ttk.Button(cb_frame, text="\u25B6  Plot", command=self.plot, style='Accent.TButton').pack(side=tk.LEFT, padx=15)
        # ---- Renk skalası (contour) ayar satırı ----
        # "Auto" işaretliyken Min/Max veriden hesaplanır (eski davranış);
        # kapatılırsa kullanıcı kendi Min/Max/Levels (kademe sayısı)
        # değerlerini girip "Apply" ile (veya Enter'a basarak) uygulayabilir.
        scale_frame = ttk.Frame(f2)
        scale_frame.pack(fill=tk.X, pady=2)
        ttk.Label(scale_frame, text="Color Scale:", width=16).pack(side=tk.LEFT)
        ttk.Checkbutton(scale_frame, text="Auto", variable=self.scale_auto,
                        command=self._on_scale_auto_toggled).pack(side=tk.LEFT)
        ttk.Label(scale_frame, text="Min:").pack(side=tk.LEFT, padx=(10, 2))
        self.scale_min_entry = ttk.Entry(scale_frame, textvariable=self.scale_min, width=9, state='disabled')
        self.scale_min_entry.pack(side=tk.LEFT)
        ttk.Label(scale_frame, text="Max:").pack(side=tk.LEFT, padx=(8, 2))
        self.scale_max_entry = ttk.Entry(scale_frame, textvariable=self.scale_max, width=9, state='disabled')
        self.scale_max_entry.pack(side=tk.LEFT)
        ttk.Label(scale_frame, text="Levels:").pack(side=tk.LEFT, padx=(8, 2))
        self.scale_levels_entry = ttk.Entry(scale_frame, textvariable=self.scale_levels, width=5)
        self.scale_levels_entry.pack(side=tk.LEFT)
        for entry in (self.scale_min_entry, self.scale_max_entry, self.scale_levels_entry):
            entry.bind('<Return>', lambda e: self._on_setting_changed())
        ttk.Button(scale_frame, text="Apply", command=self._on_setting_changed).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Label(scale_frame, text="(uncheck Auto to set your own min/max and how many color bands "
                                     "the legend is split into - only applies when the selected Color "
                                     "By column is numeric; text columns use a legend instead)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(side=tk.LEFT, padx=8)

        # ---- Görünüm satırı ----
        view_frame = ttk.Frame(f2)
        view_frame.pack(fill=tk.X, pady=2)
        ttk.Checkbutton(view_frame, text="Maximize 3D View (hide the side chart for a bigger FEM view)",
                        variable=self.maximize_3d, command=self._on_setting_changed).pack(side=tk.LEFT)

        self.stats_label = ttk.Label(f2, text="", foreground=_APP_COLORS.get('primary_dark', '#1d4ed8'),
                                     font=('Segoe UI', 9, 'bold'))
        self.stats_label.pack(anchor=tk.W, pady=(4, 0))

        # ---- Section 3: FEM Visualization ----
        f3 = ttk.LabelFrame(main, text="\U0001F5FA  3. FEM Visualization", padding=5)
        f3.pack(fill=tk.BOTH, expand=True, pady=5)

        self.fig = Figure(figsize=(12, 6.5), dpi=95)
        self.canvas = FigureCanvasTkAgg(self.fig, master=f3)
        toolbar_frame = ttk.Frame(f3)
        toolbar_frame.pack(fill=tk.X)
        # NavigationToolbar2Tk: matplotlib'in standart araç çubuğu - döndür/kaydır/
        # yakınlaştır ve çizimi PNG olarak kaydetme butonlarını hazır olarak verir.
        self.toolbar = NavigationToolbar2Tk(self.canvas, toolbar_frame)
        self.toolbar.update()
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        # Fare tekerleği ile 3B görünümde yakınlaştırma/uzaklaştırma -
        # NavigationToolbar2Tk'nin "zoom to rectangle" aracı 3B eksenlerde
        # pek işe yaramadığından, modeli yakından incelemenin en pratik
        # yolu bu - bkz. _on_scroll_zoom. Bir kere bağlanması yeterli,
        # çünkü self.canvas her plot()'ta yeniden oluşturulmuyor.
        self.canvas.mpl_connect('scroll_event', self._on_scroll_zoom)
        self._draw_placeholder()

        # ---- Durum/log kutusu ----
        self.log_text = scrolledtext.ScrolledText(main, height=4, font=('Consolas', 9),
                                                    bg=_APP_COLORS['log_bg'], fg=_APP_COLORS['log_fg'],
                                                    insertbackground=_APP_COLORS['log_fg'], relief='flat',
                                                    borderwidth=0)
        self.log_text.pack(fill=tk.X, pady=(5, 0))

    def _log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)

    def _browse_bdf(self):
        path = filedialog.askopenfilename(filetypes=[("BDF", "*.bdf *.dat *.nas"), ("All", "*.*")])
        if path:
            self.bdf_path.set(path)

    def _browse_summary_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.summary_folder.set(folder)

    # Group By / Color By değiştiğinde, zaten bir çizim varsa otomatik
    # olarak yeniden çiz - kullanıcının her seferinde "Plot" butonuna
    # tıklamasına gerek kalmasın.
    def _on_setting_changed(self):
        if self.bdf_model is not None:
            self.plot()

    # _on_scale_auto_toggled: "Auto" kutusu işaretlenip/kaldırıldığında
    # Min/Max giriş kutularını etkinleştir/devre dışı bırak (Levels her
    # zaman düzenlenebilir kalır - bant sayısı Auto modda da geçerli),
    # sonra bir çizim varsa yeniden çiz.
    def _on_scale_auto_toggled(self):
        state = 'disabled' if self.scale_auto.get() else 'normal'
        self.scale_min_entry.config(state=state)
        self.scale_max_entry.config(state=state)
        self._on_setting_changed()

    # _on_scroll_zoom: fare tekerleği ile SADECE 3B eksenler üzerindeyken
    # yakınlaştır/uzaklaştır - mevcut görünümün merkezini sabit tutarak
    # X/Y/Z sınırlarını aynı oranda küçültüp/büyütür.
    def _on_scroll_zoom(self, event):
        ax = event.inaxes
        if ax is None or not hasattr(ax, 'get_proj'):
            return  # 2B panel üzerindeyse veya eksen dışındaysa bir şey yapma
        scale = 0.9 if event.button == 'up' else 1.1
        try:
            x_lim = ax.get_xlim3d(); y_lim = ax.get_ylim3d(); z_lim = ax.get_zlim3d()
        except Exception:
            return
        xc, yc, zc = sum(x_lim) / 2, sum(y_lim) / 2, sum(z_lim) / 2
        xr = (x_lim[1] - x_lim[0]) / 2 * scale
        yr = (y_lim[1] - y_lim[0]) / 2 * scale
        zr = (z_lim[1] - z_lim[0]) / 2 * scale
        ax.set_xlim3d(xc - xr, xc + xr)
        ax.set_ylim3d(yc - yr, yc + yr)
        ax.set_zlim3d(zc - zr, zc + zr)
        self.canvas.draw_idle()

    # ==================== BDF YÜKLEME ====================
    def load_bdf(self):
        path = self.bdf_path.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror("Error", "Select a valid BDF file")
            return

        self._log(f"Loading BDF: {os.path.basename(path)} ...")
        try:
            # xref=True: node.get_position() ve elem.node_ids/elem.pid'in
            # düzgün çözümlenmesi (özellikle farklı koordinat sistemlerine
            # bağlı GRID'ler) için gerekli - _calculate_offset_csv'nin de
            # aynı sebeple xref=True kullandığı gibi.
            model = BDF(debug=False)
            try:
                model.read_bdf(path, validate=False, xref=True, read_includes=True, encoding='latin-1')
            except Exception:
                # INCLUDE edilen dosyalar diskte bulunamazsa (örn. sadece
                # ana BDF kullanıcıda varsa), INCLUDE'ları genişletmeden
                # tekrar dene - en azından ana dosyanın geometrisi görülebilsin.
                model = BDF(debug=False)
                model.read_bdf(path, validate=False, xref=True, read_includes=False, encoding='latin-1')
                self._log("  NOTE: some INCLUDE files could not be resolved - loaded the main file only.")

            self.bdf_model = model
            self._extract_mesh()

            self.bdf_status.config(
                text=f"Loaded: {len(self.node_coords)} nodes, {len(self.bar_elements)} bar element(s), "
                     f"{len(self.shell_elements)} shell element(s)",
                foreground="green")
            self._log(f"BDF loaded OK: {len(self.node_coords)} nodes, {len(self.bar_elements)} bar "
                      f"element(s) (CBAR/CBEAM), {len(self.shell_elements)} shell element(s) "
                      f"(CQUAD4/CTRIA3/...)")
            self.plot()
        except Exception as e:
            self.bdf_status.config(text="Error", foreground="red")
            self._log(f"ERROR loading BDF: {e}")
            messagebox.showerror("Error", f"Could not load BDF:\n{e}")

    # _extract_mesh: xref'lenmiş pyNastran modelinden, çizim için gereken
    # düğüm koordinatlarını ve CBAR/CBEAM (bar) ile CQUAD4/CTRIA3/... (shell)
    # elemanlarının düğüm bağlantılarını çıkarır.
    def _extract_mesh(self):
        model = self.bdf_model
        self.node_coords = {}
        for nid, node in model.nodes.items():
            try:
                self.node_coords[nid] = tuple(node.get_position())
            except Exception:
                pass  # koordinat sistemi çözülemeyen tekil düğümü atla

        self.bar_elements = []
        self.shell_elements = []
        for eid, elem in model.elements.items():
            if elem.type in ('CBAR', 'CBEAM'):
                try:
                    nids = elem.node_ids
                except Exception:
                    continue
                if len(nids) >= 2 and nids[0] in self.node_coords and nids[1] in self.node_coords:
                    pid = elem.pid if not hasattr(elem, 'Pid') else elem.Pid()
                    self.bar_elements.append({'eid': eid, 'pid': pid, 'n1': nids[0], 'n2': nids[1]})
            elif elem.type in ('CQUAD4', 'CTRIA3', 'CQUAD8', 'CTRIA6'):
                try:
                    nids = [n for n in elem.node_ids if n in self.node_coords]
                except Exception:
                    continue
                if len(nids) >= 3:
                    pid = elem.pid if not hasattr(elem, 'Pid') else elem.Pid()
                    self.shell_elements.append({'eid': eid, 'pid': pid, 'nodes': nids})

    # ==================== SUMMARY CSV YÜKLEME ====================
    @staticmethod
    def _normalize_header(h):
        return str(h).strip().lower().replace('\n', '').replace(' ', '').replace('(', '').replace(')', '')

    # _standardize_summary_df: bir özet CSV'sinin SADECE iki ID sütununu
    # ("Bar Element ID"/"Bar Property ID" - gerçek dosyalarda farklı
    # yazılabiliyor, örn. "BarElementID") tanıyıp sabit isimlere çevirir;
    # bu ikisi gruplama/indeksleme için özel anlam taşıdığından sabit
    # kalmaları gerekiyor. DİĞER TÜM sütunlar (t, d, Material, Life, ya da
    # kullanıcının getirdiği BAŞKA herhangi bir sütun - örn. "Stress",
    # "Margin of Safety") OLDUĞU GİBİ bırakılır, sadece görünüm için
    # (satır sonu/boşluk fazlalığı) temizlenir - böylece "Color By" listesi
    # hangi CSV verilirse verilsin o CSV'nin GERÇEK sütunlarını gösterir.
    def _standardize_summary_df(self, df):
        rename = {}
        beid_col = bpid_col = None
        for c in df.columns:
            norm = self._normalize_header(c)
            if beid_col is None and norm == 'barelementid':
                beid_col = c
            elif bpid_col is None and norm == 'barpropertyid':
                bpid_col = c
        if beid_col:
            rename[beid_col] = 'Bar Element ID'
        if bpid_col:
            rename[bpid_col] = 'Bar Property ID'
        for c in df.columns:
            if c in rename:
                continue
            clean = ' '.join(str(c).strip().replace('\n', ' ').split())
            if clean and clean != c:
                rename[c] = clean
        return df.rename(columns=rename)

    # _metric_columns: bir özet DataFrame'inde, iki ID sütunu DIŞINDA kalan
    # - yani "Color By" açılır listesinde gösterilecek - tüm sütun adlarını
    # döndürür. Bu, uygulamanın sütun isimlerini SABİT olarak bilmesini
    # gerektirmez; kullanıcı farklı bir CSV formatı getirse bile (Bar
    # Element ID / Bar Property ID sütunları olduğu sürece) otomatik uyum sağlar.
    @staticmethod
    def _metric_columns(df):
        if df is None:
            return []
        id_cols = {'Bar Element ID', 'Bar Property ID'}
        return [c for c in df.columns if c not in id_cols]

    # _is_numeric_column: verilen sütunun değerlerinin (en az yarısının)
    # sayıya çevrilebilir olup olmadığına bakarak, bu metriğin sürekli bir
    # renk skalasıyla mı (contour), yoksa kategorik bir lejantla mı
    # ("Material" sütunu gibi) gösterilmesi gerektiğine karar verir.
    @staticmethod
    def _is_numeric_column(series):
        try:
            if len(series) == 0:
                return False
            coerced = pd.to_numeric(series, errors='coerce')
            return coerced.notna().mean() >= 0.5
        except Exception:
            return False

    # _fmt_fixed: bir sayıyı HER ZAMAN sabit ondalıklı (fixed-point)
    # gösterimle döndürür - Python'un varsayılan/genel (":.3g" gibi)
    # biçimlendirmesi büyük/küçük sayılarda "2.25e+04" gibi bilimsel
    # gösterime kayabiliyor; Min/Max kutuları ve renk skalası (contour)
    # etiketleri kullanıcı için her zaman okunaklı sabit ondalık göstersin diye.
    @staticmethod
    def _fmt_fixed(x):
        try:
            x = float(x)
        except (TypeError, ValueError):
            return str(x)
        if x == 0:
            return "0"
        a = abs(x)
        if a >= 1000:
            return f"{x:.0f}"
        elif a >= 1:
            return f"{x:.2f}"
        elif a >= 0.001:
            return f"{x:.4f}"
        else:
            return f"{x:.6f}"

    # _scan_summary_files: Summary Folder altında (rekürsif) DT_SIZING_BAR_
    # SUMMARY.csv / DT_SIZING_PROP_SUMMARY.csv / DT_FULL_SUMMARY.csv dosyalarını ara.
    @staticmethod
    def _scan_summary_files(folder):
        bar_path = prop_path = full_path = None
        for dirpath, _, filenames in os.walk(folder):
            for fn in filenames:
                fnl = fn.lower()
                if fnl == 'dt_sizing_bar_summary.csv' and bar_path is None:
                    bar_path = os.path.join(dirpath, fn)
                elif fnl == 'dt_sizing_prop_summary.csv' and prop_path is None:
                    prop_path = os.path.join(dirpath, fn)
                elif fnl == 'dt_full_summary.csv' and full_path is None:
                    full_path = os.path.join(dirpath, fn)
        return bar_path, prop_path, full_path

    def load_summary(self):
        folder = self.summary_folder.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Error", "Select a valid Summary Folder")
            return

        self._log(f"Scanning for summary CSVs under: {folder}")
        bar_path, prop_path, full_path = self._scan_summary_files(folder)
        self.bar_summary_df = None
        self.prop_summary_df = None
        msgs = []

        try:
            if bar_path:
                df = self._standardize_summary_df(pd.read_csv(bar_path))
                if 'Bar Element ID' in df.columns:
                    df = df.drop_duplicates(subset=['Bar Element ID']).set_index('Bar Element ID')
                    self.bar_summary_df = df
                    msgs.append(f"Bar Element ID: {len(df)} row(s) (DT_SIZING_BAR_SUMMARY.csv)")
            if prop_path:
                df = self._standardize_summary_df(pd.read_csv(prop_path))
                if 'Bar Property ID' in df.columns:
                    df = df.drop_duplicates(subset=['Bar Property ID']).set_index('Bar Property ID')
                    self.prop_summary_df = df
                    msgs.append(f"Bar Property ID: {len(df)} row(s) (DT_SIZING_PROP_SUMMARY.csv)")

            # Sizing özetleri hiç bulunamadıysa, DT_FULL_SUMMARY.csv'den basit
            # bir geri dönüş (fallback) üret - her ID için İLK görülen satırı
            # kullanır (Iterative Full Run'ın kendi "governing/en kalın"
            # seçim mantığını burada TEKRARLAMAZ, sadece hızlı bir önizleme sağlar).
            if self.bar_summary_df is None and self.prop_summary_df is None and full_path:
                full_df = self._standardize_summary_df(pd.read_csv(full_path))
                if 'Bar Element ID' in full_df.columns:
                    self.bar_summary_df = full_df.drop_duplicates(subset=['Bar Element ID']).set_index('Bar Element ID')
                    msgs.append(f"Bar Element ID: {len(self.bar_summary_df)} row(s) "
                                f"(DT_FULL_SUMMARY.csv fallback, first row per ID)")
                if 'Bar Property ID' in full_df.columns:
                    self.prop_summary_df = full_df.drop_duplicates(subset=['Bar Property ID']).set_index('Bar Property ID')
                    msgs.append(f"Bar Property ID: {len(self.prop_summary_df)} row(s) "
                                f"(DT_FULL_SUMMARY.csv fallback, first row per ID)")

            if not msgs:
                self.summary_status.config(text="No DT_SIZING_BAR_SUMMARY.csv / DT_SIZING_PROP_SUMMARY.csv / "
                                                 "DT_FULL_SUMMARY.csv found", foreground="red")
                self._log("  No usable summary CSV found under that folder.")
                return

            self.summary_status.config(text="; ".join(msgs), foreground="green")
            for m in msgs:
                self._log(f"  {m}")
            # Yeni yüklenen CSV'nin GERÇEK sütunlarına göre "Color By"
            # listesini yeniden doldur - farklı bir CSV formatı (farklı
            # sütun adları) verilmiş olsa bile otomatik uyum sağlanır.
            self._refresh_metric_options()
            if self.bdf_model is not None:
                self.plot()
        except Exception as e:
            self.summary_status.config(text="Error", foreground="red")
            self._log(f"ERROR loading summary: {e}")
            messagebox.showerror("Error", f"Could not load summary CSV(s):\n{e}")

    # _refresh_metric_options: şu an seçili Group By'a (Bar Element ID /
    # Bar Property ID) karşılık gelen DataFrame'in sütunlarından "Color By"
    # açılır listesini yeniden kurar. Mevcut seçim artık mevcut değilse
    # ("life" içeren bir sütun varsa onu, yoksa ilk sütunu) makul bir
    # varsayılana geçer.
    def _refresh_metric_options(self):
        df = self.bar_summary_df if self.group_by.get() == 'Bar Element ID' else self.prop_summary_df
        cols = self._metric_columns(df)
        self.color_combo['values'] = cols
        if not cols:
            self.color_by.set("")
        elif self.color_by.get() not in cols:
            preferred = next((c for c in cols if 'life' in c.lower()), cols[0])
            self.color_by.set(preferred)

    # _on_group_by_changed: Group By değiştiğinde - iki özet CSV'nin
    # (DT_SIZING_BAR_SUMMARY.csv / DT_SIZING_PROP_SUMMARY.csv) sütunları
    # birbirinden farklı olabileceğinden - önce Color By listesini o
    # grubun gerçek sütunlarına göre tazele, sonra (bir model yüklüyse) yeniden çiz.
    def _on_group_by_changed(self):
        self._refresh_metric_options()
        self._on_setting_changed()

    # ==================== ÇİZİM (PLOT) ====================
    def _draw_placeholder(self):
        self.fig.clear()
        ax = self.fig.add_subplot(1, 1, 1)
        ax.text(0.5, 0.5, "Load a BDF file (and optionally a sizing summary)\nto see the FEM visualization here",
                ha='center', va='center', fontsize=11, color='gray')
        ax.set_xticks([]); ax.set_yticks([])
        for spine in ax.spines.values():
            spine.set_visible(False)
        self.canvas.draw()

    @staticmethod
    def _get_cmap(name, n=None):
        # matplotlib sürümleri arasında colormap API'si değişti
        # (eski: cm.get_cmap(name, n), yeni: matplotlib.colormaps[name]) -
        # ikisini de deneyerek her iki sürümde de çalışmasını sağla.
        try:
            base = matplotlib.colormaps[name]
        except Exception:
            base = mpl_cm.get_cmap(name)
        if n is not None and n > 0:
            try:
                return base.resampled(n)
            except AttributeError:
                try:
                    return mpl_cm.get_cmap(name, n)
                except Exception:
                    return base
        return base

    @staticmethod
    def _set_axes_equal(ax):
        """3B eksenlerin X/Y/Z oranlarını eşitler (matplotlib'in 3D
        projeksiyonu varsayılan olarak set_aspect('equal') desteklemez) -
        aksi halde BDF geometrisi gerçek oranlarından saptırılmış görünür."""
        x_lim = ax.get_xlim3d(); y_lim = ax.get_ylim3d(); z_lim = ax.get_zlim3d()
        spans = [abs(x_lim[1] - x_lim[0]), abs(y_lim[1] - y_lim[0]), abs(z_lim[1] - z_lim[0])]
        centers = [sum(x_lim) / 2, sum(y_lim) / 2, sum(z_lim) / 2]
        radius = max(spans) / 2 if max(spans) > 0 else 1
        ax.set_xlim3d([centers[0] - radius, centers[0] + radius])
        ax.set_ylim3d([centers[1] - radius, centers[1] + radius])
        ax.set_zlim3d([centers[2] - radius, centers[2] + radius])

    # plot: ana çizim fonksiyonu - solda 3B renklendirilmiş FEM mesh'i,
    # sağda aynı veriyi (ID'ye göre sıralanmış) 2B grafik olarak gösterir.
    def plot(self):
        if not _MATPLOTLIB_AVAILABLE:
            return
        if self.bdf_model is None:
            messagebox.showerror("Error", "Load a BDF file first")
            return

        try:
            group_by = self.group_by.get()
            metric = self.color_by.get()
            df = self.bar_summary_df if group_by == 'Bar Element ID' else self.prop_summary_df
            if not metric:
                messagebox.showerror("Error", "Load a Summary Folder first (no Color By column available)")
                return
            # Seçili metrik sütunu SAYISAL mı (örn. Life, t, d, ya da
            # kullanıcının CSV'sindeki herhangi bir sayısal sütun) yoksa
            # METİN/KATEGORİK mi (örn. Material) - bu, sütun adına göre
            # DEĞİL, DEĞERLERE bakarak otomatik belirlenir, böylece farklı
            # bir CSV formatı verildiğinde de doğru görselleştirme seçilir.
            if df is not None and metric in df.columns:
                is_material = not self._is_numeric_column(df[metric])
            else:
                is_material = False  # bu grup için hiç veri yok - sayısal yoldaki "veri yok" mesajı daha açıklayıcı

            self.fig.clear()
            # "Maximize 3D View" işaretliyse yan 2B grafiği hiç oluşturma -
            # 3B FEM görünümü tüm çizim alanını kaplasın (büyütülmüş, daha
            # rahat döndürülüp/incelenebilen görünüm). İşaretli değilse,
            # 3B tarafa 2B tarafından daha fazla pay ver (3:2 oranı) - yine
            # de varsayılan olarak eski 50/50'den daha büyük bir FEM görünümü.
            if self.maximize_3d.get():
                ax3d = self.fig.add_subplot(1, 1, 1, projection='3d')
                ax2d = None
            else:
                gs = self.fig.add_gridspec(1, 2, width_ratios=[3, 2])
                ax3d = self.fig.add_subplot(gs[0], projection='3d')
                ax2d = self.fig.add_subplot(gs[1])

            # ---- Her bar elemanının (eid/pid'ine göre) bu metrikteki değerini topla ----
            records = []  # (bar_element_dict, key, value)
            for be in self.bar_elements:
                key = be['eid'] if group_by == 'Bar Element ID' else be['pid']
                val = None
                if df is not None and metric in df.columns and key in df.index:
                    v = df.loc[key, metric]
                    if isinstance(v, pd.Series):  # aynı ID'den birden fazla satır kalmışsa
                        v = v.iloc[0]
                    val = v
                records.append((be, key, val))

            # ---- Bağlam (context) için shell mesh'i soluk gri çiz ----
            if self.shell_elements:
                verts = [[self.node_coords[n] for n in se['nodes']] for se in self.shell_elements]
                shell_poly = Poly3DCollection(verts, facecolor=(0.6, 0.6, 0.6, 0.12),
                                              edgecolor=(0.5, 0.5, 0.5, 0.25), linewidths=0.3)
                ax3d.add_collection3d(shell_poly)

            cmap = None
            norm = None
            numeric_vals = []
            mat_color = {}
            materials = []

            if is_material:
                materials = sorted({str(v) for _, _, v in records
                                    if v is not None and not (isinstance(v, float) and pd.isna(v))})
                palette = self._get_cmap('tab20', max(len(materials), 1))
                mat_color = {m: palette(i % 20) for i, m in enumerate(materials)}
            else:
                for _, _, v in records:
                    try:
                        if v is not None and not (isinstance(v, float) and pd.isna(v)):
                            numeric_vals.append(float(v))
                    except (TypeError, ValueError):
                        pass
                # "Life" (ya da "life" geçen başka bir metrik adı) için
                # kırmızı(düşük)-yeşil(yüksek) renk skalası anlamlı - diğer
                # TÜM sayısal sütunlar (sabit bir liste değil, kullanıcının
                # getirdiği herhangi bir metrik) için genel amaçlı viridis.
                cmap_name = 'RdYlGn' if 'life' in metric.lower() else 'viridis'

                # ---- Renk skalası aralığı: Auto ise veriden, değilse
                # kullanıcının Min/Max kutularından al. Kademe (Levels)
                # sayısı HER ZAMAN kullanıcıdan okunur (Auto modda da
                # geçerlidir) - HyperMesh'teki gibi bantlı (discrete)
                # bir contour skalası elde etmek için. ----
                if numeric_vals:
                    auto_vmin, auto_vmax = min(numeric_vals), max(numeric_vals)
                else:
                    auto_vmin, auto_vmax = 0.0, 1.0

                if self.scale_auto.get():
                    vmin, vmax = auto_vmin, auto_vmax
                    # Kullanıcı Auto'yu kapatırsa düzenlemeye başlayacağı
                    # makul bir başlangıç değeri görsün diye kutuları güncelle.
                    self.scale_min.set(self._fmt_fixed(vmin))
                    self.scale_max.set(self._fmt_fixed(vmax))
                else:
                    try:
                        vmin = float(self.scale_min.get())
                        vmax = float(self.scale_max.get())
                        if vmin >= vmax:
                            raise ValueError
                    except (TypeError, ValueError):
                        self._log("  WARNING: invalid Min/Max - falling back to Auto range for this plot.")
                        vmin, vmax = auto_vmin, auto_vmax

                if vmin == vmax:
                    vmin, vmax = vmin - 1, vmax + 1

                try:
                    n_levels = int(float(self.scale_levels.get()))
                    if n_levels < 1:
                        raise ValueError
                except (TypeError, ValueError):
                    n_levels = 10
                    self.scale_levels.set("10")

                # BoundaryNorm + ayrık (discrete) renk haritası: değer
                # aralığını n_levels eşit banda bölüp her banda sabit bir
                # renk atar - HyperMesh'teki "contour legend" görünümü.
                boundaries = [vmin + (vmax - vmin) * i / n_levels for i in range(n_levels + 1)]
                cmap = self._get_cmap(cmap_name, n_levels)
                norm = mpl_colors.BoundaryNorm(boundaries, cmap.N, clip=True)

            # ---- Bar elemanlarını (CBAR/CBEAM) renklendirerek çiz ----
            segments = []
            seg_colors = []
            for be, key, val in records:
                segments.append([self.node_coords[be['n1']], self.node_coords[be['n2']]])
                if is_material:
                    seg_colors.append(mat_color.get(str(val), (0.75, 0.75, 0.75, 1.0))
                                      if val is not None and not (isinstance(val, float) and pd.isna(val))
                                      else (0.75, 0.75, 0.75, 1.0))
                else:
                    try:
                        if val is None or (isinstance(val, float) and pd.isna(val)):
                            raise ValueError
                        seg_colors.append(cmap(norm(float(val))))
                    except (TypeError, ValueError):
                        seg_colors.append((0.75, 0.75, 0.75, 1.0))  # veri yok -> nötr gri
            if segments:
                line_coll = Line3DCollection(segments, colors=seg_colors, linewidths=2.5)
                ax3d.add_collection3d(line_coll)

            # ---- Eksen sınırları: tüm düğümlerin kapladığı kutuya göre ----
            if self.node_coords:
                xs, ys, zs = zip(*self.node_coords.values())
                ax3d.set_xlim3d(min(xs), max(xs))
                ax3d.set_ylim3d(min(ys), max(ys))
                ax3d.set_zlim3d(min(zs), max(zs))
                self._set_axes_equal(ax3d)
            # ---- 3B görünümü büyüt (zoom) ----
            # matplotlib'in Axes3D'si, kendisine ayrılan dikdörtgen alanı
            # (özellikle TEK başına tüm figürü kapladığında, yani Maximize
            # 3D View açıkken) kübik görünüm korumak için otomatik olarak
            # KÜÇÜLTÜP ortalar - bu da "büyütülmüş" görünümün aslında daha
            # KÜÇÜK görünmesine yol açar. set_box_aspect'in zoom parametresi
            # (matplotlib 3.6+), ayrılan alanı değiştirmeden İÇERİĞİ o alan
            # içinde büyütür - model gerçekten daha büyük ve incelemeye
            # daha elverişli görünsün diye. Maximize açıkken daha agresif,
            # kapalıyken (2B panelle paylaşılan dar alanda) daha ölçülü
            # bir zoom uygulanıyor. Eski matplotlib sürümlerinde zoom
            # parametresi yoksa (TypeError), sessizce aspect-only'e düş.
            try:
                ax3d.set_box_aspect((1, 1, 1), zoom=1.5 if ax2d is None else 1.2)
            except TypeError:
                try:
                    ax3d.set_box_aspect((1, 1, 1))
                except Exception:
                    pass
            ax3d.set_xlabel('X'); ax3d.set_ylabel('Y'); ax3d.set_zlabel('Z')
            ax3d.set_title(f"{group_by} \u2192 {metric}", fontsize=10)

            if is_material and materials:
                handles = [Patch(color=mat_color[m], label=m) for m in materials]
                # Lejant, ax3d'nin KENDİ köşesine değil, FİGÜRÜN sol üst
                # köşesine sabitleniyor - zoom ile büyütülen 3B içerik
                # genelde kutusunun ortasına/sağına doğru render edildiği
                # için (bkz. set_box_aspect zoom notu), eksen-bazlı bir
                # lejant onun üstüne biner; figür-bazlı lejant bundan bağımsız.
                self.fig.legend(handles=handles, loc='upper left', fontsize=7, bbox_to_anchor=(0.01, 0.98))
            elif not is_material and numeric_vals:
                mappable = mpl_cm.ScalarMappable(norm=norm, cmap=cmap)
                mappable.set_array([])
                # boundaries/ticks=boundaries + spacing='uniform': renk
                # çubuğunu HyperMesh'teki gibi n_levels adet eşit genişlikte
                # bant olarak çizer, her bandın sınır değerini etiket yapar.
                # format yerine FuncFormatter: "%.3g" büyük/küçük sayılarda
                # bilimsel gösterime (2.25e+04) kayabiliyor - kullanıcı her
                # zaman sabit ondalıklı (fixed-point) etiket görsün diye.
                self.fig.colorbar(mappable, ax=ax3d, boundaries=boundaries, ticks=boundaries,
                                  spacing='uniform', shrink=0.7, pad=0.08, label=metric,
                                  format=mpl_ticker.FuncFormatter(lambda v, pos: self._fmt_fixed(v)))

            # ---- İstatistik satırı: ax2d olsun olmasın (Maximize 3D View
            # açıkken bile) HER ZAMAN güncellenir. ----
            if is_material:
                self.stats_label.config(
                    text=f"{len(records)} bar element(s), {len(materials)} distinct '{metric}' value(s)")
            else:
                ys_vals_all = []
                for _, _, val in records:
                    try:
                        if val is None or (isinstance(val, float) and pd.isna(val)):
                            continue
                        ys_vals_all.append(float(val))
                    except (TypeError, ValueError):
                        pass
                if ys_vals_all:
                    self.stats_label.config(
                        text=f"{len(ys_vals_all)}/{len(records)} bar element(s) have {metric} data - "
                             f"min={self._fmt_fixed(min(ys_vals_all))}  max={self._fmt_fixed(max(ys_vals_all))}  "
                             f"avg={self._fmt_fixed(sum(ys_vals_all)/len(ys_vals_all))}  "
                             f"(scale: {self._fmt_fixed(vmin)} to {self._fmt_fixed(vmax)}, {n_levels} band(s))")
                else:
                    self.stats_label.config(text=f"{len(records)} bar element(s) - no {metric} data loaded")

            # ---- Sağdaki 2B panel: "Maximize 3D View" kapalıyken çizilir -
            # sayısal değerler için ID'ye göre sıralı dağılım, Material için
            # eleman sayısı çubuk grafiği. ----
            if ax2d is not None and is_material:
                counts = {}
                for _, _, val in records:
                    if val is None or (isinstance(val, float) and pd.isna(val)):
                        continue
                    key_m = str(val)
                    counts[key_m] = counts.get(key_m, 0) + 1
                if counts:
                    names = list(counts.keys())
                    vals = [counts[n] for n in names]
                    bar_colors = [mat_color[n] for n in names]
                    ax2d.bar(range(len(names)), vals, color=bar_colors)
                    ax2d.set_xticks(range(len(names)))
                    ax2d.set_xticklabels(names, rotation=45, ha='right', fontsize=7)
                    ax2d.set_ylabel('Element count')
                    ax2d.set_title(f"'{metric}' distribution", fontsize=10)
                else:
                    ax2d.text(0.5, 0.5, f'No {metric} data', ha='center', va='center')
            elif ax2d is not None:
                xs_ids, ys_vals = [], []
                for be, key, val in records:
                    try:
                        if val is None or (isinstance(val, float) and pd.isna(val)):
                            continue
                        ys_vals.append(float(val))
                        xs_ids.append(key)
                    except (TypeError, ValueError):
                        pass
                if ys_vals:
                    order = sorted(range(len(xs_ids)), key=lambda i: ys_vals[i])
                    xs_sorted = [xs_ids[i] for i in order]
                    ys_sorted = [ys_vals[i] for i in order]
                    pt_colors = [cmap(norm(v)) for v in ys_sorted]
                    ax2d.scatter(range(len(xs_sorted)), ys_sorted, c=pt_colors, s=16)
                    step = max(1, len(xs_sorted) // 15)
                    tick_idx = list(range(0, len(xs_sorted), step))
                    ax2d.set_xticks(tick_idx)
                    ax2d.set_xticklabels([str(xs_sorted[i]) for i in tick_idx], rotation=60, fontsize=6)
                    ax2d.set_xlabel(group_by)
                    ax2d.set_ylabel(metric)
                    ax2d.set_title(f"{metric} by {group_by} (sorted)", fontsize=10)
                    ax2d.grid(alpha=0.3)
                else:
                    ax2d.text(0.5, 0.5, f'No {metric} data loaded for this grouping\n'
                                        f'(Load Summary first)', ha='center', va='center')

            # tight_layout() 2B paneli varken kenar boşluklarını iyi
            # ayarlıyor, ama TEK BAŞINA 3B eksen (Maximize açıkken) için
            # onu tekrar ortalayıp küçültüyor - bu yüzden o durumda sabit,
            # geniş kenar boşluklarıyla subplots_adjust kullanılıyor.
            if ax2d is not None:
                self.fig.tight_layout()
            else:
                self.fig.subplots_adjust(left=0.02, right=0.96, top=0.90, bottom=0.08)
            self.canvas.draw()
        except Exception as e:
            self._log(f"ERROR plotting: {e}")
            messagebox.showerror("Error", f"Could not plot:\n{e}")


# ============================================================================
# main()  --  programın giriş noktası
# ============================================================================
# Tkinter ana penceresini (root) oluşturur, tüm sekmeleri barındıran
# IntegratedBDFRFTool uygulamasını bu pencere üzerine kurar, ve Tkinter'ın
# olay döngüsünü (mainloop) başlatır - bu çağrı, pencere kapatılana kadar
# bloklar (GUI olaylarını işler).
def main():
    root = tk.Tk()
    app = IntegratedBDFRFTool(root)
    root.mainloop()


# Bu dosya doğrudan çalıştırıldığında (başka bir modül tarafından import
# edilmediğinde) main()'i çağır - standart Python giriş noktası deseni.
if __name__ == "__main__":
    main()
