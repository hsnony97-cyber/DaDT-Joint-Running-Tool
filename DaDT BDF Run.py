#!/usr/bin/env python3
"""
DaDT Structure Definition Tool v16.0
=====================================
Tab 1: BDF Merge Preparation
Tab 2: Understanding Structure Type (apply 'Bar Property'/'Skin Property' + offsets, solve base model)
Tab 3: Post Processing (OP2 -> H5 batch conversion, Load Extraction folder prep)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import re
import threading
import csv
import shutil
import tempfile
import subprocess
import time
import concurrent.futures
from datetime import datetime
import pandas as pd
from pyNastran.bdf.bdf import BDF
import numpy as np
import openpyxl

_APP_COLORS = {}  # populated by _configure_app_style(); shared palette for both tabs


def _configure_app_style(root):
    """One-time visual theme for the whole app. Both tabs share this root's
    ttk.Style singleton, so configuring it once here restyles every ttk
    widget in both tabs without touching how any of them behave. Pure
    look-and-feel - no widget's command/callback is affected."""
    colors = {
        'bg': '#f1f5f9',
        'surface': '#ffffff',
        'text': '#0f172a',
        'muted': '#64748b',
        'border': '#cbd5e1',
        'primary': '#2563eb',
        'primary_dark': '#1d4ed8',
        'success': '#16a34a',
        'success_dark': '#15803d',
        'danger': '#dc2626',
        'danger_dark': '#b91c1c',
        'log_bg': '#0f172a',
        'log_fg': '#e2e8f0',
    }

    root.configure(bg=colors['bg'])

    style = ttk.Style(root)
    try:
        style.theme_use('clam')
    except tk.TclError:
        pass

    base_font = ('Segoe UI', 10)
    section_font = ('Segoe UI', 10, 'bold')

    style.configure('.', background=colors['bg'], foreground=colors['text'], font=base_font)
    style.configure('TFrame', background=colors['bg'])
    style.configure('TLabel', background=colors['bg'], foreground=colors['text'])
    style.configure('TCheckbutton', background=colors['bg'])

    style.configure('Header.TLabel', font=('Segoe UI', 15, 'bold'),
                    foreground=colors['primary_dark'], background=colors['bg'])
    style.configure('Subtitle.TLabel', font=('Segoe UI', 9),
                    foreground=colors['muted'], background=colors['bg'])

    style.configure('TLabelframe', background=colors['bg'], bordercolor=colors['border'])
    style.configure('TLabelframe.Label', background=colors['bg'],
                    foreground=colors['primary_dark'], font=section_font)

    style.configure('TNotebook', background=colors['bg'], borderwidth=0)
    style.configure('TNotebook.Tab', font=('Segoe UI', 10, 'bold'), padding=(16, 8),
                    background=colors['border'], foreground=colors['text'])
    style.map('TNotebook.Tab',
              background=[('selected', colors['surface'])],
              foreground=[('selected', colors['primary_dark'])])

    style.configure('TButton', font=base_font, padding=(10, 5),
                    background=colors['surface'], foreground=colors['text'])
    style.map('TButton', background=[('active', colors['border'])])

    style.configure('Accent.TButton', font=('Segoe UI', 10, 'bold'), padding=(14, 7),
                    background=colors['primary'], foreground='white')
    style.map('Accent.TButton',
              background=[('active', colors['primary_dark']), ('disabled', colors['border'])],
              foreground=[('disabled', colors['muted'])])

    style.configure('Danger.TButton', font=('Segoe UI', 10, 'bold'), padding=(14, 7),
                    background=colors['danger'], foreground='white')
    style.map('Danger.TButton',
              background=[('active', colors['danger_dark']), ('disabled', colors['border'])],
              foreground=[('disabled', colors['muted'])])

    style.configure('TEntry', padding=4, fieldbackground=colors['surface'], bordercolor=colors['border'])
    style.configure('TCombobox', padding=4)

    style.configure('Horizontal.TProgressbar', troughcolor=colors['border'],
                    background=colors['primary'], bordercolor=colors['border'], thickness=14)

    _APP_COLORS.update(colors)
    return colors


class IntegratedBDFRFTool:
    def __init__(self, root):
        self.root = root
        self.root.title("DaDT Structure Definition Tool v16.0")
        self.root.geometry("1150x950")
        self.root.minsize(1000, 700)
        self.colors = _configure_app_style(self.root)

        # Tab 1 variables
        self.thermal_bdfs = []
        self.maneuver_bdfs = []
        self.excel_path = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.output_thermal_name = tk.StringVar(value="merged_thermal.bdf")
        self.output_maneuver_name = tk.StringVar(value="merged_maneuver.bdf")
        self.set_id = tk.StringVar(value="99")
        self.node_set_id = tk.StringVar(value="199")
        self.temp_initial = tk.StringVar(value="10")
        

        self.setup_ui()
    
    def setup_ui(self):
        banner = ttk.Frame(self.root, padding=(16, 12))
        banner.pack(fill=tk.X)
        ttk.Label(banner, text="\U0001F527 DaDT Structure Definition Tool", style='Header.TLabel').pack(side=tk.LEFT)
        ttk.Label(banner, text="  v16.0", style='Subtitle.TLabel').pack(side=tk.LEFT, pady=(6, 0))
        ttk.Separator(self.root, orient='horizontal').pack(fill=tk.X)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        self.tab1 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="\U0001F4C4  BDF Merge Preparation")
        self.setup_tab1()

        self.tab3 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab3, text="\U0001F9EE  Understanding Structure Type")
        self.tab3_tool = BarPropertySolverTab(self.tab3, self.root)

        self.tab4 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab4, text="\U0001F4E6  Post Processing")
        self.tab4_tool = PostProcessingTab(self.tab4, self.root)

        self.tab5 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab5, text="\U0001F517  Joint Run")
        self.tab5_tool = JointRunTab(self.tab5, self.root)

        self.tab6 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab6, text="⚡  Full Run")
        self.tab6_tool = FullRunTab(self.tab6, self.root, self.tab3_tool, self.tab4_tool, self.tab5_tool)

        self.tab7 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab7, text="\U0001F501  Iterative Full Run")
        self.tab7_tool = IterativeFullRunTab(self.tab7, self.root, self.tab3_tool, self.tab4_tool, self.tab5_tool)


    def setup_tab1(self):
        main = ttk.Frame(self.tab1, padding="10")
        main.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main, text="BDF Merge Preparation v8", style='Header.TLabel').pack(pady=(0,10))

        # === THERMAL SECTION ===
        thermal_main = ttk.LabelFrame(main, text="\U0001F321  THERMAL", padding="8")
        thermal_main.pack(fill=tk.X, pady=5)
        
        thm_master_f = ttk.Frame(thermal_main)
        thm_master_f.pack(fill=tk.X, pady=2)
        ttk.Label(thm_master_f, text="MASTER BDFs:", width=12).pack(side=tk.LEFT)
        ttk.Button(thm_master_f, text="Add...", command=self.add_thermal_bdfs).pack(side=tk.LEFT, padx=2)
        ttk.Button(thm_master_f, text="Clear", command=self.clear_thermal_bdfs).pack(side=tk.LEFT, padx=2)
        self.thermal_count = tk.StringVar(value="0 files")
        ttk.Label(thm_master_f, textvariable=self.thermal_count).pack(side=tk.LEFT, padx=5)
        
        self.thermal_listbox = tk.Listbox(thermal_main, height=3, width=100, font=('Segoe UI', 9),
                                           bg=self.colors['surface'], fg=self.colors['text'],
                                           selectbackground=self.colors['primary'], selectforeground='white',
                                           relief='flat', highlightthickness=1,
                                           highlightbackground=self.colors['border'],
                                           highlightcolor=self.colors['primary'])
        self.thermal_listbox.pack(fill=tk.X, pady=2)
        
        # === MANEUVER SECTION ===
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
        
        # === SETTINGS ===
        sf = ttk.LabelFrame(main, text="⚙  Settings", padding="10")
        sf.pack(fill=tk.X, pady=5)
        ttk.Label(sf, text="Excel:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(sf, textvariable=self.excel_path, width=70).grid(row=0, column=1, padx=5)
        ttk.Button(sf, text="Browse", command=self.browse_excel).grid(row=0, column=2)
        ttk.Label(sf, text="Output:").grid(row=1, column=0, sticky=tk.W)
        ttk.Entry(sf, textvariable=self.output_folder, width=70).grid(row=1, column=1, padx=5)
        ttk.Button(sf, text="Browse", command=self.browse_output).grid(row=1, column=2)
        ttk.Label(sf, text="SET ID:").grid(row=2, column=0, sticky=tk.W)
        ttk.Entry(sf, textvariable=self.set_id, width=10).grid(row=2, column=1, sticky=tk.W, padx=5)
        ttk.Label(sf, text="NODE SET ID:").grid(row=3, column=0, sticky=tk.W)
        ttk.Entry(sf, textvariable=self.node_set_id, width=10).grid(row=3, column=1, sticky=tk.W, padx=5)
        ttk.Label(sf, text="TEMP(INIT):").grid(row=4, column=0, sticky=tk.W)
        ttk.Entry(sf, textvariable=self.temp_initial, width=10).grid(row=4, column=1, sticky=tk.W, padx=5)
        
        bf = ttk.Frame(main)
        bf.pack(fill=tk.X, pady=10)
        self.process_btn = ttk.Button(bf, text="▶  PROCESS & MERGE", command=self.start_processing,
                                       style='Accent.TButton')
        self.process_btn.pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Clear Log", command=self.clear_log1).pack(side=tk.LEFT)

        self.progress1 = ttk.Progressbar(main, mode='indeterminate')
        self.progress1.pack(fill=tk.X, pady=5)

        lf = ttk.LabelFrame(main, text="\U0001F4DC  Log", padding="10")
        lf.pack(fill=tk.BOTH, expand=True)
        self.log_text1 = scrolledtext.ScrolledText(lf, height=15, font=('Consolas', 10),
                                                    bg=self.colors['log_bg'], fg=self.colors['log_fg'],
                                                    insertbackground=self.colors['log_fg'], relief='flat',
                                                    borderwidth=0)
        self.log_text1.pack(fill=tk.BOTH, expand=True)
    
    # ============= TAB 1 HELPERS =============
    def add_thermal_bdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("BDF","*.bdf *.dat *.nas"),("All","*.*")])
        for f in files:
            if f not in self.thermal_bdfs:
                self.thermal_bdfs.append(f)
                self.thermal_listbox.insert(tk.END, f)
        self.thermal_count.set(f"{len(self.thermal_bdfs)} files")
    
    def clear_thermal_bdfs(self):
        self.thermal_bdfs.clear()
        self.thermal_listbox.delete(0, tk.END)
        self.thermal_count.set("0 files")
    
    def add_maneuver_bdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("BDF","*.bdf *.dat *.nas"),("All","*.*")])
        for f in files:
            if f not in self.maneuver_bdfs:
                self.maneuver_bdfs.append(f)
                self.maneuver_listbox.insert(tk.END, f)
        self.maneuver_count.set(f"{len(self.maneuver_bdfs)} files")
    
    def clear_maneuver_bdfs(self):
        self.maneuver_bdfs.clear()
        self.maneuver_listbox.delete(0, tk.END)
        self.maneuver_count.set("0 files")
    
    def browse_excel(self):
        f = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xls")])
        if f: self.excel_path.set(f)
    
    def browse_output(self):
        f = filedialog.askdirectory()
        if f: self.output_folder.set(f)
    
    def log1(self, msg):
        self.log_text1.insert(tk.END, msg + "\n")
        self.log_text1.see(tk.END)
        self.root.update_idletasks()
    
    def clear_log1(self):
        self.log_text1.delete(1.0, tk.END)
    
    def format_include_nastran(self, abs_path):
        """Uzun INCLUDE path'lerini Nastran formatına uygun böler."""
        include_line = f"INCLUDE '{abs_path}'"
        if len(include_line) <= 72:
            return [include_line]
        parts = abs_path.split('/')
        lines = []
        current_line = "INCLUDE '"
        for i, part in enumerate(parts):
            is_last = (i == len(parts) - 1)
            segment = part if is_last else part + '/'
            if len(current_line + segment) <= 72:
                current_line += segment
            else:
                if current_line != "INCLUDE '":
                    lines.append(current_line)
                current_line = segment
        if current_line:
            current_line += "'"
            lines.append(current_line)
        return lines
    
    def read_file_safe(self, fpath):
        """Dosyayı güvenli şekilde oku."""
        for enc in ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']:
            try:
                with open(fpath, 'r', encoding=enc, errors='replace') as f:
                    return f.read()
            except:
                continue
        return ""
    
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
        results = []
        lines = bdf_content.split('\n')
        
        current_subcase = None
        current_load = None
        current_temp_load = None
        current_subtitle = None
        
        for line in lines:
            upper = line.upper().strip()
            original = line.strip()
            
            # SUBCASE satırı
            if upper.startswith('SUBCASE'):
                # Önceki subcase'i kaydet
                if current_subcase is not None:
                    results.append({
                        'subcase_id': current_subcase,
                        'load_id': current_load,
                        'temp_load_id': current_temp_load,
                        'subtitle': current_subtitle
                    })
                
                # Yeni subcase başlat
                parts = upper.split()
                if len(parts) >= 2:
                    try:
                        current_subcase = int(parts[1])
                        current_load = None
                        current_temp_load = None
                        current_subtitle = None
                    except:
                        pass
            
            # LOAD = satırı
            elif current_subcase and upper.startswith('LOAD') and '=' in upper:
                match = re.search(r'LOAD\s*=\s*(\d+)', upper)
                if match:
                    current_load = int(match.group(1))
            
            # TEMPERATURE(LOAD) = satırı
            elif current_subcase and 'TEMPERATURE' in upper and 'LOAD' in upper and '=' in upper:
                match = re.search(r'TEMPERATURE\s*\(\s*LOAD\s*\)\s*=\s*(\d+)', upper)
                if match:
                    current_temp_load = int(match.group(1))
            
            # SUBTITLE satırı
            elif current_subcase and upper.startswith('SUBTITLE'):
                m = re.search(r'SUBTITLE\s+(.+)', original, re.IGNORECASE)
                if m:
                    current_subtitle = m.group(1).strip()
            
            # BEGIN BULK'a ulaştıysak case control bitti
            elif upper.startswith('BEGIN') and 'BULK' in upper:
                break
        
        # Son subcase'i kaydet
        if current_subcase is not None:
            results.append({
                'subcase_id': current_subcase,
                'load_id': current_load,
                'temp_load_id': current_temp_load,
                'subtitle': current_subtitle
            })
        
        return results
    
    def parse_multiline_includes(self, content, bdf_dir):
        """
        Çok satırlı INCLUDE'ları parse eder.
        Nastran formatında INCLUDE şöyle olabilir:
        
        INCLUDE '../../../_COMMON_/INTERFACE/STRUCTURAL/INTER_AF_HT/
        INTER_AF_HT_STRU.BDF'
        
        Returns: list of dict with keys: lines, full_text, abs_path, start_idx, end_idx
        """
        lines = content.split('\n')
        includes = []
        i = 0
        
        while i < len(lines):
            line = lines[i]
            upper = line.upper().strip()
            
            if upper.startswith('INCLUDE'):
                # INCLUDE başladı - tırnak içindeki path'i bul
                include_lines = [line]
                full_text = line
                
                # Tırnak sayısını kontrol et - tek tırnak veya çift tırnak
                quote_char = None
                if "'" in line:
                    quote_char = "'"
                elif '"' in line:
                    quote_char = '"'
                
                if quote_char:
                    # Tırnak sayısını say
                    quote_count = full_text.count(quote_char)
                    
                    # Eğer tek tırnak varsa (açılmış ama kapanmamış), devam satırlarını oku
                    j = i + 1
                    while quote_count % 2 != 0 and j < len(lines):
                        next_line = lines[j]
                        include_lines.append(next_line)
                        full_text += '\n' + next_line
                        quote_count = full_text.count(quote_char)
                        j += 1
                    
                    # Path'i çıkar - newline'ları temizle
                    clean_text = full_text.replace('\n', '')
                    match = re.search(rf"INCLUDE\s*{quote_char}([^{quote_char}]*){quote_char}", 
                                     clean_text, re.IGNORECASE)
                    if match:
                        inc_path = match.group(1).strip()
                        # Absolute path'e çevir
                        if not os.path.isabs(inc_path):
                            abs_path = os.path.normpath(os.path.join(bdf_dir, inc_path))
                        else:
                            abs_path = os.path.normpath(inc_path)
                        abs_path = abs_path.replace('\\', '/')
                        
                        includes.append({
                            'lines': include_lines,
                            'full_text': full_text,
                            'abs_path': abs_path,
                            'start_idx': i,
                            'end_idx': j - 1 if j > i + 1 else i
                        })
                    
                    i = j
                else:
                    i += 1
            else:
                i += 1
        
        return includes
    
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
        subcase_info_map = {}
        processed_files = set()  # Aynı dosyayı birden fazla kez işlememek için
        
        self.log1(f"    Reading ALL lines from {len(bdf_files)} MASTER BDFs...")
        matched_count = 0
        matched_subcases = 0
        
        for bdf_path in bdf_files:
            bdf_dir = os.path.dirname(os.path.abspath(bdf_path))
            content = self.read_file_safe(bdf_path)
            
            # INREL dosyalarını atla
            bdf_basename = os.path.basename(bdf_path).upper()
            if 'INREL' in bdf_basename:
                self.log1(f"      SKIP (INREL): {os.path.basename(bdf_path)}")
                continue
            
            # TÜM SUBCASE bilgilerini al (birden fazla olabilir)
            all_subcases = self.extract_subcase_load_info(content)
            
            # Bu dosyadaki hangi subcase'ler Excel listesinde?
            matching_subcases = []
            for sc_info in all_subcases:
                sc_id = sc_info['subcase_id']
                if sc_id and sc_id in load_case_set:
                    matching_subcases.append(sc_info)
            
            # Eşleşen subcase varsa bu dosyayı işle
            if matching_subcases:
                # Dosya daha önce işlendiyse sadece subcase info'ları ekle
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
                
                processed_files.add(bdf_path)
                matched_count += 1
                
                # Log - kaç subcase eşleşti
                sc_ids = [str(sc['subcase_id']) for sc in matching_subcases]
                self.log1(f"      MATCH: {os.path.basename(bdf_path)} ({len(matching_subcases)} subcases: {', '.join(sc_ids[:5])}{'...' if len(sc_ids) > 5 else ''})")
                
                # Subcase info'ları kaydet
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
                
                # Önce tüm INCLUDE'ları parse et (çok satırlı dahil)
                all_includes = self.parse_multiline_includes(content, bdf_dir)
                
                # INCLUDE'ları kategorize et
                common_include_indices = set()  # COMMON INCLUDE satır indeksleri
                structure_include_count = 0
                common_include_count = 0
                
                for inc in all_includes:
                    abs_path_upper = inc['abs_path'].upper()
                    
                    # Bu INCLUDE COMMON LOAD/THERMAL mı?
                    is_common = False
                    if common_type == 'LOAD':
                        if '_COMMON_/LOAD' in abs_path_upper or '/COMMON/LOAD' in abs_path_upper:
                            is_common = True
                    elif common_type == 'THERMAL':
                        if '_COMMON_/THERMAL' in abs_path_upper or '/COMMON/THERMAL' in abs_path_upper:
                            is_common = True
                    
                    if is_common:
                        # COMMON INCLUDE - path'i kaydet, satırları atla
                        common_includes_raw.append(inc['abs_path'])
                        for idx in range(inc['start_idx'], inc['end_idx'] + 1):
                            common_include_indices.add(idx)
                        common_include_count += 1
                    else:
                        # Structure/Interface/vs INCLUDE - tek satır INCLUDE olarak ekle (pyNastran açacak)
                        # Absolute path ile yeni INCLUDE satırı oluştur
                        include_line = f"INCLUDE '{inc['abs_path']}'"
                        all_lines_raw.append(include_line)
                        # Orijinal satırları atla (common_include_indices'e ekle)
                        for idx in range(inc['start_idx'], inc['end_idx'] + 1):
                            common_include_indices.add(idx)
                        structure_include_count += 1
                
                # TÜM SATIRLARI oku (INCLUDE satırları hariç - hem COMMON hem diğerleri)
                lines = content.split('\n')
                line_count = 0
                
                for idx, line in enumerate(lines):
                    # INCLUDE satırı mı? Atla (zaten yukarıda işledik)
                    if idx in common_include_indices:
                        continue
                    
                    stripped = line.strip()
                    if not stripped:
                        continue
                    
                    # Case control satırlarını atla
                    upper = stripped.upper()
                    skip_keywords = ['SOL ', 'SOL\t', 'CEND', 'TITLE', 'SUBTITLE', 'ECHO', 
                                    'SUBCASE', 'LOAD =', 'LOAD=', 'SPC =', 'SPC=', 
                                    'TEMPERATURE', 'DISPLACEMENT', 'FORCE', 'GPFORCE', 
                                    'OLOAD', 'SPCFORCE', 'SET ', 'BEGIN BULK', 
                                    'BEGIN,BULK', 'ENDDATA']
                    is_skip = any(upper.startswith(kw) for kw in skip_keywords)
                    
                    if not is_skip:
                        all_lines_raw.append(line)  # Orijinal satırı koru
                        line_count += 1
                
                self.log1(f"        -> {line_count} data lines, {structure_include_count} structure includes, {common_include_count} common includes")
        
        self.log1(f"    Matched {matched_count} MASTER BDFs with {matched_subcases} total subcases")
        self.log1(f"    Total raw lines (including structure includes): {len(all_lines_raw)}")
        self.log1(f"    Total raw COMMON includes: {len(common_includes_raw)}")
        
        # Duplicate satırları çıkar
        self.log1("    Removing duplicate lines...")
        seen_lines = set()
        unique_lines = []
        for line in all_lines_raw:
            # Normalize et (boşlukları düzenle)
            normalized = ' '.join(line.split())
            if normalized not in seen_lines:
                seen_lines.add(normalized)
                unique_lines.append(line)
        
        # Duplicate COMMON INCLUDE'ları çıkar
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
    
    def extract_param_cards(self, bulk_data):
        """
        Bulk data'dan PARAM kartlarını ayırır.
        Returns: (param_lines, remaining_bulk_data)
        """
        lines = bulk_data.split('\n')
        param_lines = []
        other_lines = []
        
        seen_params = set()  # Duplicate PARAM kontrolü
        
        i = 0
        while i < len(lines):
            line = lines[i]
            stripped = line.strip()
            upper = stripped.upper()
            
            if upper.startswith('PARAM'):
                # PARAM kartı - continuation satırlarını da al
                param_block = [line]
                
                # PARAM name'ini çıkar (duplicate kontrolü için)
                param_name = None
                try:
                    if ',' in line:
                        parts = line.split(',')
                        if len(parts) > 1:
                            param_name = parts[1].strip().upper()
                    else:
                        if len(line) >= 16:
                            param_name = line[8:16].strip().upper()
                except:
                    pass
                
                i += 1
                # Continuation satırlarını kontrol et
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
                        break
                
                # Duplicate kontrolü
                if param_name and param_name not in seen_params:
                    seen_params.add(param_name)
                    param_lines.extend(param_block)
            else:
                other_lines.append(line)
                i += 1
        
        return param_lines, '\n'.join(other_lines)
    
    def check_and_remove_duplicates(self, bulk_data):
        """
        Bulk data içindeki duplicate kartları tespit edip kaldırır.
        
        - Element/Property/Material kartları: ID bazlı kontrol (aynı ID → duplicate)
        - SPC/FORCE/MOMENT kartları: Tüm satır bazlı kontrol (birebir aynıysa → duplicate)
        """
        self.log1("    Checking for duplicate entries...")
        
        lines = bulk_data.split('\n')
        
        # ID bazlı kontrol yapılacak kartlar
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
        
        # Tüm satır bazlı kontrol yapılacak kartlar (SPC, FORCE, MOMENT, MPC)
        line_based_cards = ['SPC', 'SPC1', 'FORCE', 'MOMENT', 'MPC', 'LOAD', 'TEMP', 'TEMPD']
        seen_full_lines = set()  # Tüm satır için
        
        # İstatistikler
        duplicate_counts = {}
        
        result_lines = []
        i = 0
        
        while i < len(lines):
            line = lines[i]
            stripped = line.strip()
            
            # Boş veya comment satırı
            if not stripped or stripped.startswith('$'):
                result_lines.append(line)
                i += 1
                continue
            
            upper = stripped.upper()
            
            # Hangi tip kart?
            card_type = None
            is_line_based = False
            
            # Önce line-based kartları kontrol et
            for ctype in line_based_cards:
                if upper.startswith(ctype) and (len(upper) == len(ctype) or 
                    upper[len(ctype)] in ' ,\t*'):
                    card_type = ctype
                    is_line_based = True
                    break
            
            # Sonra ID-based kartları kontrol et
            if not card_type:
                for ctype in id_based_cards.keys():
                    if upper.startswith(ctype) and (len(upper) == len(ctype) or 
                        upper[len(ctype)] in ' ,\t*'):
                        card_type = ctype
                        break
            
            if card_type:
                if is_line_based:
                    # LINE-BASED: Tüm satırı (ve continuation'ları) karşılaştır
                    full_card = [line]
                    j = i + 1
                    
                    # Continuation satırlarını topla
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
                    
                    # Normalize edilmiş tam kart
                    normalized_card = '|'.join(' '.join(l.split()) for l in full_card)
                    
                    if normalized_card in seen_full_lines:
                        # Duplicate - atla
                        if card_type not in duplicate_counts:
                            duplicate_counts[card_type] = 0
                        duplicate_counts[card_type] += 1
                        i = j
                        continue
                    else:
                        seen_full_lines.add(normalized_card)
                        result_lines.extend(full_card)
                        i = j
                        continue
                else:
                    # ID-BASED: Sadece ID'ye bak
                    card_id = None
                    try:
                        if ',' in line:
                            parts = line.split(',')
                            if len(parts) > 1:
                                id_str = parts[1].strip()
                                if id_str:
                                    card_id = int(float(id_str))
                        else:
                            if len(line) >= 16:
                                id_str = line[8:16].strip()
                                if id_str:
                                    card_id = int(float(id_str))
                    except:
                        pass
                    
                    if card_id is not None:
                        if card_id in id_based_cards[card_type]:
                            # Duplicate - atla
                            if card_type not in duplicate_counts:
                                duplicate_counts[card_type] = 0
                            duplicate_counts[card_type] += 1
                            
                            # Continuation satırlarını da atla
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
                            id_based_cards[card_type].add(card_id)
            
            result_lines.append(line)
            i += 1
        
        # Rapor
        if duplicate_counts:
            self.log1("    Removed duplicates:")
            for ctype, count in sorted(duplicate_counts.items()):
                self.log1(f"      {ctype}: {count}")
            total = sum(duplicate_counts.values())
            self.log1(f"      TOTAL: {total} duplicate entries removed")
        else:
            self.log1("    No duplicates found")
        
        return '\n'.join(result_lines)
    
    def merge_lines_with_pynastran(self, lines):
        """
        Satırları temp BDF'e yazıp pyNastran ile merge eder.
        Duplicate ID hatası olursa, satırları direkt yazar.
        Returns: merged bulk data string
        """
        if not lines:
            self.log1("    WARNING: No lines to merge!")
            return ""
        
        temp_bdf_path = os.path.join(tempfile.gettempdir(), "_temp_lines_to_merge.bdf")
        
        self.log1(f"    Writing {len(lines)} lines to temp BDF...")
        
        try:
            # Temp BDF oluştur
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
                bdf = BDF(debug=False)
                # allow_duplicate_ids ile dene
                bdf.read_bdf(temp_bdf_path, validate=False, xref=False, 
                            read_includes=True, save_file_structure=False)
                
                self.log1(f"      Loaded: {len(bdf.nodes)} nodes, {len(bdf.elements)} elements")
                self.log1(f"      Properties: {len(bdf.properties)}, Materials: {len(bdf.materials)}")
                self.log1(f"      Coords: {len(bdf.coords)}, MPCs: {len(bdf.mpcs)}, SPCs: {len(bdf.spcs)}")
                
                # Merge edilmiş BDF'i yaz
                merged_temp_path = os.path.join(tempfile.gettempdir(), "_temp_merged.bdf")
                bdf.write_bdf(merged_temp_path, size=8, is_double=False)
                
                with open(merged_temp_path, 'r', errors='ignore') as f:
                    merged_content = f.read()
                
                if os.path.exists(merged_temp_path): os.remove(merged_temp_path)
                
            except Exception as e:
                self.log1(f"    pyNastran failed: {str(e)[:100]}")
                self.log1(f"    Falling back to direct file reading...")
                
                # pyNastran başarısız oldu - INCLUDE'ları manuel aç
                merged_content = self.expand_includes_manually(temp_bdf_path)
            
            # Temizlik
            if os.path.exists(temp_bdf_path): os.remove(temp_bdf_path)
            
            # BEGIN BULK sonrasını al
            bulk_match = re.search(r'BEGIN\s*,?\s*BULK', merged_content, re.IGNORECASE)
            if bulk_match:
                bulk_data = merged_content[bulk_match.end():]
            else:
                bulk_data = merged_content
            
            # Gereksiz satırları temizle
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
            self.log1(f"    ERROR merging: {e}")
            import traceback
            self.log1(traceback.format_exc())
            if os.path.exists(temp_bdf_path): os.remove(temp_bdf_path)
            return ""
    
    def expand_includes_manually(self, bdf_path):
        """
        INCLUDE'ları manuel olarak açar (pyNastran başarısız olduğunda).
        """
        self.log1("    Expanding includes manually...")
        
        content = self.read_file_safe(bdf_path)
        bdf_dir = os.path.dirname(os.path.abspath(bdf_path))
        
        # INCLUDE'ları bul ve aç
        all_includes = self.parse_multiline_includes(content, bdf_dir)
        
        # Satırları işle
        lines = content.split('\n')
        result_lines = []
        
        # INCLUDE satır indekslerini topla
        include_indices = {}
        for inc in all_includes:
            for idx in range(inc['start_idx'], inc['end_idx'] + 1):
                include_indices[idx] = inc
        
        processed_includes = set()
        
        for idx, line in enumerate(lines):
            if idx in include_indices:
                inc = include_indices[idx]
                # Sadece start_idx'te işle (continuation satırlarını atla)
                if idx == inc['start_idx']:
                    inc_path = inc['abs_path']
                    if inc_path not in processed_includes:
                        processed_includes.add(inc_path)
                        # INCLUDE dosyasını oku
                        if os.path.exists(inc_path):
                            try:
                                inc_content = self.read_file_safe(inc_path)
                                result_lines.append(f"$ === EXPANDED: {os.path.basename(inc_path)} ===")
                                for inc_line in inc_content.split('\n'):
                                    # Recursive INCLUDE'ları atla (basit tutuyoruz)
                                    if not inc_line.strip().upper().startswith('INCLUDE'):
                                        result_lines.append(inc_line)
                            except Exception as e:
                                result_lines.append(f"$ ERROR reading {inc_path}: {e}")
                        else:
                            result_lines.append(f"$ FILE NOT FOUND: {inc_path}")
            else:
                result_lines.append(line)
        
        self.log1(f"      Expanded {len(processed_includes)} includes")
        return '\n'.join(result_lines)
    
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
    
    def process_merge(self):
        try:
            self.log1("="*70)
            self.log1("BDF Merger Tool v8")
            self.log1("="*70)
            
            self.log1("\n[1] Reading Excel...")
            xl = pd.ExcelFile(self.excel_path.get())
            sheets = xl.sheet_names
            self.log1(f"    Available sheets: {sheets}")

            thermal_sh = maneuver_sh = element_sh = node_sh = None

            # First pass: look for exact or specific matches
            for s in sheets:
                sl = s.lower()
                # Element_Set exact match (priority)
                if sl == 'element_set' or sl == 'elementset':
                    element_sh = s
                # Node_Set exact match (priority)
                elif sl == 'node_set' or sl == 'nodeset':
                    node_sh = s
                # Thermal
                elif 'thermal' in sl and not thermal_sh:
                    thermal_sh = s
                # Maneuver
                elif ('maneuver' in sl or 'manevra' in sl) and not maneuver_sh:
                    maneuver_sh = s

            # Second pass: if element_sh/node_sh not found, look for partial matches
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

            # Fallback to index-based if still not found
            if not thermal_sh and len(sheets) > 0: thermal_sh = sheets[0]
            if not maneuver_sh and len(sheets) > 1: maneuver_sh = sheets[1]
            if not element_sh and len(sheets) > 2: element_sh = sheets[2]
            if not node_sh and len(sheets) > 3: node_sh = sheets[3]

            self.log1(f"    Using sheets -> Thermal: '{thermal_sh}', Maneuver: '{maneuver_sh}', Element_Set: '{element_sh}', Node_Set: '{node_sh}'")

            thermal_cases = pd.read_excel(xl, sheet_name=thermal_sh).iloc[:,0].dropna().astype(int).tolist() if thermal_sh else []
            maneuver_cases = pd.read_excel(xl, sheet_name=maneuver_sh).iloc[:,0].dropna().astype(int).tolist() if maneuver_sh else []
            element_ids = sorted(pd.read_excel(xl, sheet_name=element_sh).iloc[:,0].dropna().astype(int).tolist()) if element_sh else []
            node_ids = sorted(pd.read_excel(xl, sheet_name=node_sh).iloc[:,0].dropna().astype(int).tolist()) if node_sh else []

            self.log1(f"    Thermal cases: {len(thermal_cases)}")
            self.log1(f"    Maneuver cases: {len(maneuver_cases)}")
            self.log1(f"    Element IDs: {len(element_ids)}")
            self.log1(f"    Node IDs: {len(node_ids)}")

            set_id = int(self.set_id.get())
            node_set_id = int(self.node_set_id.get())
            temp_initial = self.temp_initial.get()
            out_dir = self.output_folder.get()
            os.makedirs(out_dir, exist_ok=True)

            if self.thermal_bdfs:
                self.log1("\n" + "="*70)
                self.log1("[2] Processing THERMAL...")
                self.log1("="*70)
                self.process_thermal_bdf(self.thermal_bdfs, thermal_cases, element_ids, set_id,
                    node_ids, node_set_id, temp_initial, os.path.join(out_dir, self.output_thermal_name.get()))

            if self.maneuver_bdfs:
                self.log1("\n" + "="*70)
                self.log1("[3] Processing MANEUVER...")
                self.log1("="*70)
                self.process_maneuver_bdf(self.maneuver_bdfs, maneuver_cases, element_ids, set_id,
                    node_ids, node_set_id, os.path.join(out_dir, self.output_maneuver_name.get()))
            
            self.log1("\n" + "="*70)
            self.log1("COMPLETED!")
            self.log1("="*70)
            self.root.after(0, lambda: messagebox.showinfo("Done","Merge completed!"))
        except Exception as e:
            self.log1(f"\nERROR: {e}")
            import traceback
            self.log1(traceback.format_exc())
            self.root.after(0, lambda: messagebox.showerror("Error",str(e)))
        finally:
            self.root.after(0, lambda: [self.progress1.stop(), self.process_btn.config(state=tk.NORMAL)])
    
    def process_thermal_bdf(self, bdf_files, load_cases, element_ids, set_id, node_ids, node_set_id, temp_initial, output_path):
        self.log1(f"    MASTER BDFs: {len(bdf_files)}")
        self.log1(f"    Load cases to match: {len(load_cases)}")
        load_case_set = set(load_cases)
        
        # Step 1: Tüm satırları topla
        self.log1("\n    === Step 1: Collecting all lines from MASTER BDFs ===")
        all_lines, common_includes, subcase_info_map = self.collect_all_lines_from_masters(
            bdf_files, load_case_set, 'THERMAL'
        )
        
        # Step 2: pyNastran ile merge et
        self.log1("\n    === Step 2: Merging with pyNastran ===")
        merged_bulk_data = self.merge_lines_with_pynastran(all_lines)
        
        # Step 2.5: Duplicate kontrolü
        if merged_bulk_data:
            self.log1("\n    === Step 2.5: Checking duplicates ===")
            merged_bulk_data = self.check_and_remove_duplicates(merged_bulk_data)
        
        # Step 3: Output dosyası oluştur
        self.log1("\n    === Step 3: Writing output BDF ===")
        out = []
        out.append(f"$ {'='*60}")
        out.append(f"$ THERMAL - MERGED BDF (v8)")
        out.append(f"$ {'='*60}")
        out.append("SOL 101")
        out.append("CEND")
        out.append("ECHO=NONE")
        out.append(f"TITLE = THERMAL ANALYSIS")
        out.append(f"TEMPERATURE(INITIAL) = {temp_initial}")
        out.append("$")
        
        # SET definition
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

        # SET definition (nodes)
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

        out.append("$")
        out.append("DISPLACEMENT(SORT1,PLOT,REAL)=ALL")
        out.append(f"FORCE(SORT1,PLOT,REAL,CENTER)={set_id}")
        out.append(f"GPFORCE(PLOT)={node_set_id}")
        out.append("OLOAD(PLOT)=ALL")
        out.append("SPCFORCE(SORT1,PLOT)=ALL")
        out.append("$")

        # SUBCASE definitions
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
        
        out.append("BEGIN BULK")
        
        # PARAM kartlarını ayır ve BEGIN BULK'tan hemen sonra yaz
        param_lines = []
        if merged_bulk_data:
            param_lines, merged_bulk_data = self.extract_param_cards(merged_bulk_data)
        
        if param_lines:
            out.append("$ --- PARAM CARDS ---")
            out.extend(param_lines)
            out.append("$")
        
        # Merged bulk data
        if merged_bulk_data:
            out.append(f"$ {'='*60}")
            out.append(f"$ MERGED STRUCTURE DATA")
            out.append(f"$ {'='*60}")
            out.append(merged_bulk_data)
        
        # Common Thermal INCLUDE'ları
        out.append("$")
        out.append(f"$ {'='*60}")
        out.append(f"$ COMMON THERMAL INCLUDES ({len(common_includes)} files)")
        out.append(f"$ {'='*60}")
        
        for abs_path in sorted(common_includes):
            include_lines = self.format_include_nastran(abs_path)
            out.extend(include_lines)
        
        out.append("$")
        out.append("ENDDATA")
        
        with open(output_path, 'w') as f:
            f.write('\n'.join(out))
        
        self.log1(f"    Output: {os.path.basename(output_path)}")
        self.log1(f"    COMMON THERMAL INCLUDES: {len(common_includes)}")
    
    def process_maneuver_bdf(self, bdf_files, load_cases, element_ids, set_id, node_ids, node_set_id, output_path):
        self.log1(f"    MASTER BDFs: {len(bdf_files)}")
        self.log1(f"    Load cases to match: {len(load_cases)}")
        load_case_set = set(load_cases)
        
        # Step 1: Tüm satırları topla
        self.log1("\n    === Step 1: Collecting all lines from MASTER BDFs ===")
        all_lines, common_includes, subcase_info_map = self.collect_all_lines_from_masters(
            bdf_files, load_case_set, 'LOAD'
        )
        
        # Step 2: pyNastran ile merge et
        self.log1("\n    === Step 2: Merging with pyNastran ===")
        merged_bulk_data = self.merge_lines_with_pynastran(all_lines)
        
        # Step 2.5: Duplicate kontrolü
        if merged_bulk_data:
            self.log1("\n    === Step 2.5: Checking duplicates ===")
            merged_bulk_data = self.check_and_remove_duplicates(merged_bulk_data)
        
        # Step 3: Output dosyası oluştur
        self.log1("\n    === Step 3: Writing output BDF ===")
        out = []
        out.append(f"$ {'='*60}")
        out.append(f"$ MANEUVER - MERGED BDF (v8)")
        out.append(f"$ {'='*60}")
        out.append("SOL 101")
        out.append("CEND")
        out.append("ECHO=NONE")
        out.append(f"TITLE = MANEUVER ANALYSIS")
        out.append("$")
        
        # SET definition
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

        # SET definition (nodes)
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

        out.append("$")
        out.append("DISPLACEMENT(SORT1,PLOT,REAL)=ALL")
        out.append(f"FORCE(SORT1,PLOT,REAL,CENTER)={set_id}")
        out.append(f"GPFORCE(PLOT)={node_set_id}")
        out.append("OLOAD(PLOT)=ALL")
        out.append("SPCFORCE(SORT1,PLOT)=ALL")
        out.append("$")

        # SUBCASE definitions
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
        
        # PARAM kartlarını ayır ve BEGIN BULK'tan hemen sonra yaz
        param_lines = []
        if merged_bulk_data:
            param_lines, merged_bulk_data = self.extract_param_cards(merged_bulk_data)
        
        if param_lines:
            out.append("$ --- PARAM CARDS ---")
            out.extend(param_lines)
            out.append("$")
        
        # Merged bulk data
        if merged_bulk_data:
            out.append(f"$ {'='*60}")
            out.append(f"$ MERGED STRUCTURE DATA")
            out.append(f"$ {'='*60}")
            out.append(merged_bulk_data)
        
        # Common Load INCLUDE'ları
        out.append("$")
        out.append(f"$ {'='*60}")
        out.append(f"$ COMMON LOAD INCLUDES ({len(common_includes)} files)")
        out.append(f"$ {'='*60}")
        
        for abs_path in sorted(common_includes):
            include_lines = self.format_include_nastran(abs_path)
            out.extend(include_lines)
        
        out.append("$")
        out.append("ENDDATA")
        
        with open(output_path, 'w') as f:
            f.write('\n'.join(out))
        
        self.log1(f"    Output: {os.path.basename(output_path)}")
        self.log1(f"    COMMON LOAD INCLUDES: {len(common_includes)}")


class BarPropertySolverTab:
    def __init__(self, parent_frame, root):
        self.root = root
        self.parent_frame = parent_frame
        self.maneuver_bdfs = []
        # self.root.title("Joint Structure Type Understanding")
        # self.root.geometry("1300x950")

        # Input paths
        self.bdf_paths = []
        self.property_excel_path = tk.StringVar()
        self.element_excel_path = tk.StringVar()
        self.nastran_path = tk.StringVar()
        self.nastran_memory = tk.StringVar(value="")
        self.scratch_folder = tk.StringVar(value="")  # optional; blank = each run's own output folder (still collision-free)
        self.output_folder = tk.StringVar()
        self.cases_per_split = tk.StringVar(value="")   # optional; blank/0 = solve as a single BDF (no split)
        self.max_parallel_runs = tk.StringVar(value="1")  # 1 = sequential

        # Default bar thickness used only when a property has no Excel/BDF value
        self.default_bar_thickness = 2.0

        # Data storage
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
        self.landing_elem_ids = []
        self.bar_offset_elem_ids = []

        # Run state
        self.is_running = False
        self.properties_loaded = False

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

        ttk.Label(main, text="Joint Structure Type Understanding", style='Header.TLabel').pack(pady=10)
        ttk.Label(main, text="Apply Property Excel + offsets and solve the base model",
                 style='Subtitle.TLabel').pack()

        # ---- Section 1: Input Files ----
        f1 = ttk.LabelFrame(main, text="\U0001F4C1  1. Input Files", padding=10)
        f1.pack(fill=tk.X, pady=5, padx=10)

        # All BDF Files
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

        # Maneuver BDF (Offset Source)
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


        # Property Excel
        self._add_file_row(f1, "Property Excel:", self.property_excel_path, self.load_properties)
        self.prop_status = ttk.Label(f1, text="Not loaded", foreground="gray")
        self.prop_status.pack(anchor=tk.W)

        # Offset Element IDs
        self._add_file_row(f1, "Offset Element IDs:", self.element_excel_path, self.load_element_ids)
        self.elem_status = ttk.Label(f1, text="Not loaded", foreground="gray")
        self.elem_status.pack(anchor=tk.W)

        # Nastran Exe
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

        # Output Folder
        out_frame = ttk.Frame(f1)
        out_frame.pack(fill=tk.X, pady=2)
        ttk.Label(out_frame, text="Output Folder:", width=18).pack(side=tk.LEFT)
        ttk.Entry(out_frame, textvariable=self.output_folder, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(out_frame, text="Browse", command=lambda: self.output_folder.set(
            filedialog.askdirectory()
        )).pack(side=tk.LEFT)

        # ---- Section 2: Parallel Run Splitting ----
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

    def _add_file_row(self, parent, label, var, load_cmd):
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=2)
        ttk.Label(frame, text=label, width=18).pack(side=tk.LEFT)
        ttk.Entry(frame, textvariable=var, width=45).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame, text="Browse", command=lambda: var.set(
            filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")])
        )).pack(side=tk.LEFT)
        ttk.Button(frame, text="Load", command=load_cmd).pack(side=tk.LEFT, padx=2)

    def log(self, msg):
        def _do():
            self.log_text.insert(tk.END, msg + "\n")
            self.log_text.see(tk.END)
        self.root.after(0, _do)

    def clear_log(self):
        self.log_text.delete(1.0, tk.END)

    # ==================== BDF FILE OPS ====================

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
            path = self.bdf_paths[0]

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
            self.bdf_models = []
            for bdf_path in self.bdf_paths:
                self.log(f"\n  Loading: {os.path.basename(bdf_path)}")
                t_file_start = time.time()
                model = BDF(debug=False)
                model.read_bdf(bdf_path, validate=False, xref=False, read_includes=False, encoding='latin-1')
                self.log(f"    ...parsed in {time.time() - t_file_start:.1f}s "
                          f"({len(model.nodes)} nodes, {len(model.elements)} elements)")
                self.bdf_models.append({'path': bdf_path, 'model': model, 'name': os.path.basename(bdf_path)})

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

            # Extract PBARL dimensions from BDF
            self.pbarl_dims = {}
            for pid, prop in self.bdf_model.properties.items():
                if prop.type == 'PBARL':
                    dims = prop.dim if hasattr(prop, 'dim') else []
                    if len(dims) >= 2:
                        self.pbarl_dims[pid] = {'dim1': dims[0], 'dim2': dims[1]}
                    elif len(dims) == 1:
                        self.pbarl_dims[pid] = {'dim1': dims[0], 'dim2': dims[0]}

            # Store original dim1 from BDF as the original thicknesses
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
            self._refresh_effective_skin_thicknesses()

            self.log(f"  Shells: {shell_count}, Bars: {bar_count}")
            self.log(f"\n  Total BDF models loaded: {len(self.bdf_models)} "
                      f"(took {time.time() - t_load_start:.1f}s total)")

            self.root.after(0, lambda: self.bdf_status.config(
                text=f"Loaded: {len(self.bdf_models)} BDFs, {len(self.bdf_model.elements)} elements",
                foreground="green"
            ))

            if not self.output_folder.get():
                out_dir = os.path.dirname(path)
                self.root.after(0, lambda: self.output_folder.set(out_dir))

        except Exception as e:
            self.log(f"ERROR after {time.time() - t_load_start:.1f}s: {e}")
            import traceback
            self.log(traceback.format_exc())
            self.root.after(0, lambda: self.bdf_status.config(text="Error", foreground="red"))

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
            self.log(f"  Sheets: {xl.sheet_names}")

            # ---- Locate the Bar/Skin Property sheets. No structure-grouping
            # sheet is read anymore - there is no sweep to group by, only a
            # single base-model solve, so every PID that has an entry in
            # 'Bar Property'/'Skin Property' gets that value applied; any PID
            # without one keeps its original BDF value untouched. ----
            bar_dim_sheet = skin_sheet = None
            for sheet in xl.sheet_names:
                sl = sheet.lower().replace('_', '').replace(' ', '').replace('-', '')
                if sl == 'barproperty':
                    bar_dim_sheet = sheet
                elif sl == 'skinproperty':
                    skin_sheet = sheet
            # Fallback partial match for Bar/Skin Property sheets
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
                for _, row in df_bar.iterrows():
                    try:
                        pid = int(row.iloc[0])
                        d1 = float(row.iloc[1])
                        d2 = float(row.iloc[2]) if len(df_bar.columns) > 2 else d1
                        self.bar_properties[pid] = {'dim1': d1, 'dim2': d2}
                        self.current_bar_thicknesses[pid] = d1
                    except Exception:
                        pass
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
            # Recompute the effective thickness (Excel first, BDF fallback) now,
            # regardless of whether load_bdf() has run yet or not.
            self._refresh_effective_skin_thicknesses()

            total = len(self.bar_properties)
            self.log(f"\n  Properties loaded in {time.time() - t_start:.1f}s")
            self.prop_status.config(
                text=f"Loaded: {total} bar props, {len(self.skin_properties)} skin props",
                foreground="green"
            )
            self.properties_loaded = True
            self.btn_start.config(state=tk.NORMAL)

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
        while i < len(lines):
            line = lines[i]

            if line.startswith('PBARL'):
                try:
                    pid = int(line[8:16].strip())
                    if pid in thickness_overrides:
                        t = max(0.1, thickness_overrides[pid])
                        new_lines.append(line)
                        i += 1
                        while i < len(lines) and (
                            lines[i].startswith('+') or lines[i].startswith('*') or
                            (lines[i][0] == ' ' and lines[i].strip() and not lines[i].strip().startswith('$'))
                        ):
                            cont = lines[i]
                            if cont.strip() and not cont.strip().startswith('$'):
                                try:
                                    original_dim2 = cont[16:24].strip()
                                    dim2_val = float(original_dim2) if original_dim2 else t
                                except:
                                    dim2_val = t
                                cont_name = cont[:8]
                                rest = cont[24:] if len(cont) > 24 else '\n'
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

        self.maneuver_base_path = None
        if self.maneuver_bdfs:
            man_folder = os.path.join(base_folder, "maneuver")
            os.makedirs(man_folder, exist_ok=True)
            self.maneuver_base_path = self._apply_excel_properties(self.maneuver_bdfs[0], man_folder)
            self.log(f"  Base maneuver (offset source) BDF ready: {os.path.basename(self.maneuver_base_path)}")
        else:
            self.log("  No maneuver BDF provided - no offsets will be computed.")

        for bdf_idx, bdf_info in enumerate(self.bdf_models):
            bdf_folder = os.path.join(base_folder, f"bdf_{bdf_idx + 1}_{os.path.splitext(bdf_info['name'])[0]}")
            os.makedirs(bdf_folder, exist_ok=True)
            base_path = self._apply_excel_properties(bdf_info['path'], bdf_folder)
            bdf_info['base_path'] = base_path
            self.log(f"  Base model ready for {bdf_info['name']}: {os.path.basename(base_path)}")

        self.log(f"  Properties applied in {time.time() - t_props_start:.1f}s")

        # Run Nastran on the base model. thickness_overrides={} means every
        # PID uses its self.bar_properties base value - offsets get computed
        # fresh from the maneuver BDF at that same base thickness.
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
        maneuver_path = self._write_bdf_for_model(man_folder, None, self.maneuver_base_path, thickness_overrides)

        try:
            t_parse_start = time.time()
            bdf = BDF(debug=False)
            bdf.read_bdf(maneuver_path, validate=False, xref=True, read_includes=False, encoding='latin-1')
            self.log(f"    {label}Maneuver BDF parsed for offsets in {time.time() - t_parse_start:.1f}s")
        except Exception as e:
            self.log(f"    {label}Offset calc error reading maneuver BDF: {e}")
            return None

        try:
            # Landing (shell) offsets: zoffset = -t/2
            landing_normals = {}
            rows = []

            for eid in self.landing_elem_ids:
                if eid not in bdf.elements:
                    continue
                elem = bdf.elements[eid]
                pid = elem.pid if hasattr(elem, 'pid') else None
                if pid is None:
                    continue

                t = self.current_skin_thicknesses.get(pid, 3.0)
                t = max(0.1, t)
                zoffset = -t / 2.0

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
                landing_normals[eid] = normal
                rows.append({'Type': 'Landing', 'Element': eid, 'ZOffset': zoffset,
                             'WA_X': '', 'WA_Y': '', 'WA_Z': ''})

            # Node -> shell mapping for bar offset
            node_to_shells = {}
            for eid, elem in bdf.elements.items():
                if elem.type in ['CQUAD4', 'CTRIA3']:
                    for nid in elem.node_ids:
                        node_to_shells.setdefault(nid, []).append(eid)

            # Bar offsets: WA = WB = -normal * (landing_t + bar_t/2)
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

                bar_t = thickness_overrides.get(pid, self.bar_properties.get(pid, {}).get('dim1', bar_min))
                bar_t = max(0.1, bar_t)

                bar_nodes = elem.node_ids[:2]
                if bar_nodes[0] in node_to_shells and bar_nodes[1] in node_to_shells:
                    common = set(node_to_shells[bar_nodes[0]]) & set(node_to_shells[bar_nodes[1]])
                    max_t = 0
                    best_normal = None
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

            def fmt(v, w=8):
                s = f"{v:.4f}"
                return s[:w].ljust(w) if len(s) <= w else f"{v:.2E}"[:w].ljust(w)

            new_lines = []
            i = 0
            applied_landing = 0
            applied_bar = 0

            while i < len(lines):
                line = lines[i]

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

                elif line.startswith('CBAR'):
                    try:
                        eid = int(line[8:16].strip())
                        if eid in bar_offsets:
                            offset_vec = bar_offsets[eid]

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

        offset_csv = self._calculate_offset_csv(thickness_overrides, iter_folder, label=label)

        try:
            cases_per_split = int(float(self.cases_per_split.get())) if self.cases_per_split.get().strip() else 0
        except (ValueError, TypeError):
            cases_per_split = 0

        per_bdf = []
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

        subcase_idxs = [i for i, l in enumerate(lines) if l.strip().upper().startswith('SUBCASE')]
        if len(subcase_idxs) <= cases_per_split:
            return [bdf_path]

        bulk_idx = next((i for i, l in enumerate(lines) if l.strip().upper().startswith('BEGIN BULK')), len(lines))

        header_lines = lines[:subcase_idxs[0]]
        bulk_lines = lines[bulk_idx:]
        block_bounds = subcase_idxs + [bulk_idx]

        base, ext = os.path.splitext(bdf_path)
        n_subcases = len(subcase_idxs)
        n_runs = (n_subcases + cases_per_split - 1) // cases_per_split
        run_paths = []

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
                os.makedirs(scratch_base, exist_ok=True)
                run_scratch_dir = tempfile.mkdtemp(prefix="nastran_scr_", dir=scratch_base)
                made_temp_scratch = True
                env = os.environ.copy()
                env['TEMP'] = run_scratch_dir
                env['TMP'] = run_scratch_dir
                self.log(f"    {label}Scratch dir (via TEMP/TMP): {run_scratch_dir}")

            cwd_dir = run_scratch_dir or folder

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

        def _find_op2():
            for d in op2_search_dirs:
                candidate = os.path.join(d, base + '.op2')
                if os.path.exists(candidate) and os.path.getsize(candidate) > 0:
                    return d, candidate
            return None, None

        while time.time() - start < timeout:
            elapsed = time.time() - start

            if os.path.exists(f06_path):
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
                # No .f06 at all for this install/config - fall back to a
                # long op2 size-stability window instead of failing outright.
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

            # Run Nastran and WAIT for every job to actually finish (incl.
            # each .op2 being fully written).
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
            self.is_running = False
            self.root.after(0, lambda: self.btn_start.config(state=tk.NORMAL))


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

        self.op2_folder = tk.StringVar()
        self.converter_exe_path = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.op2_files = []

        # ---- Load Extraction prep ----
        self.le_input_excel_path = tk.StringVar()    # settings excel with the 'Load_Extraction' sheet
        self.le_example_folder = tk.StringVar()      # template folder (Load_Ext_H5 + TF-X_Load_Extraction.xlsx)
        self.le_maneuver_bdf_path = tk.StringVar()
        self.le_thermal_bdf_path = tk.StringVar()
        self.le_pickle_path = tk.StringVar()
        self.le_split_count = tk.StringVar(value="10")
        self.le_output_folder = tk.StringVar()       # where the split LoadExtraction_NN folders are created
        self.le_exe_path = tk.StringVar()             # Load_Ext_H5.exe
        self.le_max_memory = tk.StringVar(value="50")  # --max-memory, GB

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

        ttk.Label(main, text="Post Processing", style='Header.TLabel').pack(pady=10)
        ttk.Label(main, text="OP2 -> H5 conversion, then Load Extraction prep - one RUN, start to finish",
                 style='Subtitle.TLabel').pack()

        # ---- Section 1: Input Files ----
        f1 = ttk.LabelFrame(main, text="\U0001F4C1  1. Input Files", padding=10)
        f1.pack(fill=tk.X, pady=5, padx=10)

        # OP2 Folder
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

        # Converter Exe
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

        # Output Folder
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

        # Settings Excel (reads the 'Load_Extraction' sheet)
        le_xl_frame = ttk.Frame(f2)
        le_xl_frame.pack(fill=tk.X, pady=2)
        ttk.Label(le_xl_frame, text="Input Excel:", width=18).pack(side=tk.LEFT)
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
        exe_dir = os.path.dirname(exe_path) or None
        exe_name = os.path.basename(exe_path)
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
        self.log(f"  Batch file written: {bat_path}")

        self.root.after(0, lambda: self.progress_var.set(2))
        self.log(f"\n  Running batch ({len(lines)} conversion(s))...")
        result = subprocess.run(["cmd", "/c", bat_path], cwd=exe_dir, capture_output=True, text=True)
        self.root.after(0, lambda: self.progress_var.set(30))

        if result.returncode != 0:
            self.log(f"  WARNING: batch exited with code {result.returncode}")
        if result.stdout:
            self.log(f"  stdout: {result.stdout.strip()[:2000]}")
        if result.stderr:
            self.log(f"  stderr: {result.stderr.strip()[:2000]}")

        n_ok = sum(1 for p in expected_h5 if os.path.exists(p))
        self.log(f"\n  {n_ok}/{len(expected_h5)} H5 file(s) produced (took {time.time() - t_start:.1f}s)")
        if n_ok < len(expected_h5):
            missing = [os.path.basename(p) for p in expected_h5 if not os.path.exists(p)]
            shown = ', '.join(missing[:20]) + (" ..." if len(missing) > 20 else "")
            self.log(f"  Missing: {shown}")

        self.root.after(0, lambda: self.progress_var.set(33))
        return n_ok, len(expected_h5)

    # ==================== LOAD EXTRACTION PREP ====================
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

        # ---- 1. Read the 'Load_Extraction' sheet ----
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

        # ---- 2. Group by Set (first-seen order), split into n chunks ----
        unique_sets = list(dict.fromkeys(df[set_col].tolist()))
        set_chunks = self._chunk_list(unique_sets, n)
        self.log(f"  {len(unique_sets)} unique Set(s) -> split into {n} group(s)")
        empty_chunks = sum(1 for c in set_chunks if not c)
        if empty_chunks:
            self.log(f"  NOTE: {empty_chunks} group(s) have no Sets (fewer unique Sets than the split count)")

        # ---- 3. H5 files from the OP2->H5 conversion's own Output Folder ----
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

        # ---- 4. One split folder per chunk ----
        n_done = 0
        for i, chunk_sets in enumerate(set_chunks):
            idx = i + 1
            self.log(f"\n  --- Group {idx}/{n}: {len(chunk_sets)} Set(s) ---")
            dest_folder = os.path.join(out_base, f"LoadExtraction_{idx:02d}")
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

            keep_vba = tfx_path.lower().endswith(('.xlsm', '.xltm', '.xlam'))
            wb = openpyxl.load_workbook(tfx_path, keep_vba=keep_vba)

            data_ws = self._find_data_sheet(wb)
            if data_ws is None:
                self.log(f"    WARNING: no 'Joint_Input' sheet (or 'Bar Element ID' header) found in "
                          f"{os.path.basename(tfx_path)} - structure rows not written.")
            else:
                header_map = {}
                for cell in next(data_ws.iter_rows(min_row=1, max_row=1)):
                    if cell.value:
                        header_map[str(cell.value).strip().lower()] = cell.column

                subset_df = df[df[set_col].isin(chunk_sets)]

                # Clear stale example rows under the header before pasting
                for col_idx in header_map.values():
                    for r in range(2, 2002):
                        data_ws.cell(row=r, column=col_idx).value = None

                for r_off, (_, row) in enumerate(subset_df.iterrows()):
                    r = r_off + 2
                    for col_name in subset_df.columns:
                        col_idx = header_map.get(str(col_name).strip().lower())
                        if col_idx:
                            val = row[col_name]
                            data_ws.cell(row=r, column=col_idx).value = None if pd.isna(val) else val
                self.log(f"    Wrote {len(subset_df)} row(s) into '{data_ws.title}'")

            control_ws = self._find_control_sheet(wb, "BDF File for Sprung System")
            if control_ws is None:
                self.log(f"    WARNING: no control sheet (with 'BDF File for Sprung System') found in "
                          f"{os.path.basename(tfx_path)} - C9/C10/C11/C17+ not written.")
            else:
                control_ws['C9'] = man_bdf
                control_ws['C10'] = thm_bdf
                control_ws['C11'] = pickle_path
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

            log_path = os.path.join(dest_folder, "loadext_log.txt")
            try:
                log_file = open(log_path, 'w')
                subprocess.Popen(["cmd", "/c", bat_path], cwd=exe_dir, stdout=log_file, stderr=subprocess.STDOUT)
                self.log(f"    Launched (output -> {log_path})")
                n_started += 1
            except Exception as e:
                self.log(f"    ERROR launching {bat_path}: {e}")

            self.root.after(0, lambda p=66 + (idx / n) * 34: self.progress_var.set(p))

            if idx < n:
                self.log("    Waiting 20s before starting the next one...")
                time.sleep(20)

        self.log(f"\n  TF-X ANALYSES: {n_started}/{n} launched (took {time.time() - t_start:.1f}s). "
                  f"Each runs independently - check each group's loadext_log.txt for progress.")
        return n_started, n

    # ==================== RUN (FULL PIPELINE) ====================
    def start_run(self):
        # OP2 -> H5 inputs
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

        # Load Extraction inputs
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

            n_ok, n_total = self._convert_op2_to_h5()

            if n_total == 0:
                self.log("\nSTOPPED: no OP2 files were converted - Load Extraction prep not started.")
                return
            if n_ok < n_total:
                self.log(f"\nSTOPPED: only {n_ok}/{n_total} H5 file(s) were produced - Load "
                          f"Extraction prep not started. Fix the issue above, then press RUN again.")
                return

            n_done, n_groups = self._prepare_load_extraction_core()

            if n_done == 0:
                self.log("\nSTOPPED: no Load Extraction group was prepared - TF-X analyses not started.")
                return

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

    def __init__(self, parent_frame, root):
        self.root = root
        self.parent_frame = parent_frame

        self.h5_folder = tk.StringVar()
        self.main_input_excel_path = tk.StringVar()
        self.example_joint_excel_path = tk.StringVar()
        self.util_folder = tk.StringVar()
        self.license_checker_exe_path = tk.StringVar()
        self.load_modifying_exe_path = tk.StringVar()
        self.rss_exe_path = tk.StringVar()
        self.joint_run_exe_path = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.max_parallel_runs = tk.StringVar(value="4")
        self.scale_factor = tk.StringVar(value="1")
        self.h5_files = []

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

        # Main Input Excel
        main_xl_frame = ttk.Frame(f1)
        main_xl_frame.pack(fill=tk.X, pady=2)
        ttk.Label(main_xl_frame, text="Main Input Excel:", width=20).pack(side=tk.LEFT)
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
    @staticmethod
    def _normalize_header(h):
        return str(h).strip().lower().replace(' ', '').replace('_', '')

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

            shutil.copy2(h5_path, os.path.join(dest_folder, os.path.basename(h5_path)))
            self.log(f"    H5 copied to: {dest_folder}")

            joint_xl_path = os.path.join(dest_folder, set_name + example_ext)
            shutil.copy2(example_excel, joint_xl_path)

            util_folder = self.util_folder.get().strip()
            if util_folder:
                util_dest = os.path.join(dest_folder, os.path.basename(os.path.normpath(util_folder)))
                shutil.copytree(util_folder, util_dest, dirs_exist_ok=True)
                self.log(f"    Util Folder copied to: {util_dest}")

            subset_df = main_df[main_df[struct_col].astype(str).str.strip() == set_name]
            self.log(f"    {len(subset_df)} row(s) match Structure Name '{set_name}'")

            keep_vba = joint_xl_path.lower().endswith(('.xlsm', '.xltm', '.xlam'))
            wb = openpyxl.load_workbook(joint_xl_path, keep_vba=keep_vba)
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

        # ---- Write User_Log.csv + copy the exe into every Set folder ----
        for dest_folder in set_folders:
            set_name = os.path.basename(dest_folder)
            joint_xl_path = os.path.join(dest_folder, set_name + example_ext)
            direction = direction_map.get(set_name.strip().lower(), "")
            if not direction:
                self.log(f"  WARNING: no Load Arrangement Direction found for '{set_name}' "
                          f"in the Load_Modify sheet.")

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

        # CSV header (normalized) -> Limit Stress Check header (exact, as in the template)
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

            for col_idx in lsc_header.values():
                for r in range(2, 2002):
                    lsc_ws.cell(row=r, column=col_idx).value = None

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

            self._run_license_checker_core()

            n_done, n_total = self._prepare_joint_folders_core()
            if n_total == 0:
                self.log("\nSTOPPED: no Joint_Loads_*.h5 files found - nothing to run.")
                return
            if n_done == 0:
                self.log("\nSTOPPED: no Set folder was prepared - Load Modifying not started.")
                return

            n_lm_ok, n_lm = self._run_load_modifying_core()
            if n_lm == 0:
                self.log("\nSTOPPED: no Set folder found for Load Modifying.")
                return

            n_ok, n_rss = self._run_rss_case_finding_core()
            if n_rss == 0:
                self.log("\nSTOPPED: no Set folder found for RSS Case Finding.")
                return

            n_written, n_lsc = self._write_limit_stress_check_core()

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

    def __init__(self, parent_frame, root, tab3_tool, tab4_tool, tab5_tool):
        self.root = root
        self.parent_frame = parent_frame
        self.tab3_tool = tab3_tool
        self.tab4_tool = tab4_tool
        self.tab5_tool = tab5_tool

        # ---- Section 1: Understanding Structure Type ----
        self.bdf_paths = []
        self.maneuver_bdfs = []
        self.property_excel_path = tk.StringVar()
        self.element_excel_path = tk.StringVar()
        self.nastran_path = tk.StringVar()
        self.nastran_memory = tk.StringVar()
        self.scratch_folder = tk.StringVar()
        self.isu_output_folder = tk.StringVar()
        self.cases_per_split = tk.StringVar()
        self.isu_max_parallel_runs = tk.StringVar(value="1")

        # ---- Section 2: Post Processing ----
        self.converter_exe_path = tk.StringVar()
        self.le_input_excel_path = tk.StringVar()
        self.le_example_folder = tk.StringVar()
        self.le_pickle_path = tk.StringVar()
        self.le_split_count = tk.StringVar(value="10")
        self.le_exe_path = tk.StringVar()
        self.le_max_memory = tk.StringVar(value="50")

        # ---- Section 3: Joint Run ----
        self.jr_main_input_excel_path = tk.StringVar()
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
        ttk.Label(main, text="Understanding Structure Type -> Post Processing -> Joint Run, "
                              "chained end to end from one button",
                 style='Subtitle.TLabel').pack()

        # ---- Section 1: Understanding Structure Type ----
        f1 = ttk.LabelFrame(main, text="\U0001F9EE  1. Understanding Structure Type", padding=10)
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

        self._add_file_row(f1, "Property Excel:", self.property_excel_path,
                            [("Excel", "*.xlsx *.xls"), ("All", "*.*")])
        self._add_file_row(f1, "Offset Element IDs:", self.element_excel_path,
                            [("Excel", "*.xlsx *.xls"), ("All", "*.*")])
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
        self._add_file_row(f2, "Input Excel:", self.le_input_excel_path,
                            [("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")])
        ttk.Label(f2, text="(the settings excel with the 'Load_Extraction' sheet)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(anchor=tk.W, pady=(0, 2))
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

        ttk.Label(f2, text="(OP2 Folder, Output Folder H5, BDF Maneuver/Thermal and the LE split "
                            "Output Folder are all auto-wired from stage 1's results)",
                 font=('Helvetica', 8, 'italic'), foreground="gray").pack(anchor=tk.W, pady=(4, 2))

        # ---- Section 3: Joint Run ----
        f3 = ttk.LabelFrame(main, text="\U0001F517  3. Joint Run", padding=10)
        f3.pack(fill=tk.X, pady=5, padx=10)

        self._add_file_row(f3, "Main Input Excel:", self.jr_main_input_excel_path,
                            [("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")])
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

        ttk.Label(f3, text="(Joint Load H5 Folder and Output Folder are auto-wired from stage 2's "
                            "Load Extraction Results)",
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
        """The '..._stru.bdf' (structure-only, prop-updated + offseted)
        file whose name contains keyword ('maneuver'/'thermal'), or None."""
        matches = []
        for dirpath, _, filenames in os.walk(base_model_folder):
            for fn in filenames:
                fnl = fn.lower()
                if fnl.endswith('_stru.bdf') and keyword in fnl:
                    matches.append(os.path.join(dirpath, fn))
        return matches[0] if matches else None

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
    def start_run(self):
        if not self.bdf_paths:
            messagebox.showerror("Error", "Add at least one All BDF File")
            return
        if not self.property_excel_path.get().strip() or not os.path.exists(self.property_excel_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Property Excel")
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
        if not self.le_input_excel_path.get().strip() or not os.path.exists(self.le_input_excel_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Post Processing Input Excel")
            return
        if not self.le_example_folder.get().strip() or not os.path.isdir(self.le_example_folder.get().strip()):
            messagebox.showerror("Error", "Select a valid Example LE Folder")
            return
        if not self.le_exe_path.get().strip() or not os.path.exists(self.le_exe_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Load Extraction Exe")
            return
        if not self.jr_main_input_excel_path.get().strip() or not os.path.exists(self.jr_main_input_excel_path.get().strip()):
            messagebox.showerror("Error", "Select a valid Joint Run Main Input Excel")
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

            output_base = self.isu_output_folder.get().strip()
            before_folders = self._list_base_model_folders(output_base)

            self.root.after(0, lambda: self.progress_var.set(5))
            t3._run_solve()
            self.root.after(0, lambda: self.progress_var.set(33))

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
                self.log("  WARNING: no '..._stru.bdf' file with 'maneuver' in its name was found "
                          "under _base_model - make sure a maneuver BDF is included in 'All BDF "
                          "Files' (not just as the separate Maneuver/Offset Source entry).")
            if not thermal_stru:
                self.log("  WARNING: no '..._stru.bdf' file with 'thermal' in its name was found "
                          "under _base_model - make sure a thermal BDF is included in 'All BDF Files'.")
            self.log(f"  BDF (Maneuver) -> {maneuver_stru}")
            self.log(f"  BDF (Thermal) -> {thermal_stru}")

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

            self._wait_for_exe_to_finish(os.path.basename(self.le_exe_path.get().strip()))
            self.root.after(0, lambda: self.progress_var.set(70))

            # ==================== STAGE C: JOINT RUN ====================
            self.log("\n" + "#" * 70)
            self.log("STAGE 3/3: JOINT RUN")
            self.log("#" * 70)

            joint_run_dir = os.path.join(base_model_folder, "Joint_Run")
            os.makedirs(joint_run_dir, exist_ok=True)

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

    def _refresh_tab3_listboxes(self):
        t3 = self.tab3_tool
        t3.bdf_listbox.delete(0, tk.END)
        for p in t3.bdf_paths:
            t3.bdf_listbox.insert(tk.END, os.path.basename(p))
        t3.maneuver_listbox.delete(0, tk.END)
        for p in t3.maneuver_bdfs:
            t3.maneuver_listbox.insert(tk.END, os.path.basename(p))


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

    def __init__(self, parent_frame, root, tab3_tool, tab4_tool, tab5_tool):
        self.root = root
        self.parent_frame = parent_frame
        self.tab3_tool = tab3_tool
        self.tab4_tool = tab4_tool
        self.tab5_tool = tab5_tool

        # ---- Section 1: Understanding Structure Type ----
        self.bdf_paths = []
        self.maneuver_bdfs = []
        self.master_input_excel_path = tk.StringVar()
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

        # ---- Section 4: Iteration settings ----
        self.target_flight_hours = tk.StringVar()
        self.n_iterations = tk.StringVar(value="3")

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

        # ---- Section 1: Understanding Structure Type ----
        f1 = ttk.LabelFrame(main, text="\U0001F9EE  1. Understanding Structure Type", padding=10)
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

    @staticmethod
    def _find_all_stru_bdfs(base_model_folder):
        matches = []
        for dirpath, _, filenames in os.walk(base_model_folder):
            for fn in filenames:
                if fn.lower().endswith('_stru.bdf'):
                    matches.append(os.path.join(dirpath, fn))
        return matches

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

        csv_files = []
        for dirpath, _, filenames in os.walk(joint_run_folder):
            for fn in filenames:
                if 'nasgro final summary' in fn.lower():
                    csv_files.append(os.path.join(dirpath, fn))
        self.log(f"  Found {len(csv_files)} 'Nasgro Final Summary' csv file(s) under {joint_run_folder}")
        if not csv_files:
            return None, None

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

        filtered_for_life = full_summary[full_summary['Life'] >= wanted_flight_hours]
        filtered_sorted = filtered_for_life.sort_values(by=t_col, ascending=True)
        dedup_cols = ['Bar Element ID', 'Bar Element Type'] if 'Bar Element Type' in full_summary.columns \
            else ['Bar Element ID']
        filtered_sorted_rmd = filtered_sorted.drop_duplicates(subset=dedup_cols)
        filtered_sorted_rmd = filtered_sorted_rmd.sort_values(by=t_col, ascending=False)

        full_summary_th_sort = full_summary.sort_values(by=t_col, ascending=False)
        df_bar_sum = pd.concat([filtered_sorted_rmd, full_summary_th_sort], ignore_index=True)
        df_bar_sum = df_bar_sum.drop_duplicates(subset=dedup_cols)
        df_bar_sum = df_bar_sum.sort_values(by=t_col, ascending=False)
        df_bar_sum = df_bar_sum.drop_duplicates(subset=['Bar Element ID'])
        df_bar_sum.to_csv(os.path.join(joint_run_folder, 'DT_SIZING_BAR_SUMMARY.csv'),
                           encoding='utf-8', index=False)

        df_prop_sum = df_bar_sum.sort_values(by=t_col, ascending=False)
        df_prop_sum = df_prop_sum.drop_duplicates(subset=['Bar Property ID'])
        df_prop_sum.to_csv(os.path.join(joint_run_folder, 'DT_SIZING_PROP_SUMMARY.csv'),
                            encoding='utf-8', index=False)

        life_check_summary = full_summary.sort_values(by='Life', ascending=True)
        life_check_summary.drop_duplicates(subset=['Bar Element ID']).to_csv(
            os.path.join(joint_run_folder, 'DT_LIFE_CHECK_BAR_SUMMARY.csv'), encoding='utf-8', index=False)
        life_check_summary.drop_duplicates(subset=['Bar Property ID']).to_csv(
            os.path.join(joint_run_folder, 'DT_LIFE_CHECK_PROP_SUMMARY.csv'), encoding='utf-8', index=False)

        self.log(f"  DT_FULL_SUMMARY: {len(full_summary)} rows, DT_SIZING_BAR_SUMMARY: "
                  f"{len(df_bar_sum)} rows, DT_SIZING_PROP_SUMMARY: {len(df_prop_sum)} rows")
        return df_prop_sum, t_col

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

            working_excel = os.path.join(output_base, "_IterativeFullRun_Working_Input.xlsx")
            shutil.copy2(self.master_input_excel_path.get().strip(), working_excel)
            self.log(f"  Working copy of the Master Input Excel: {working_excel}")
            self.log(f"  Target flight hours: {target_life}, Iterations: {n_iterations}")

            prop_summary_df = None
            t_col = None

            for iteration in range(1, n_iterations + 1):
                iter_t_start = time.time()
                self.log("\n" + "#" * 70)
                self.log(f"ITERATION {iteration}/{n_iterations}")
                self.log("#" * 70)
                self.root.after(0, lambda p=(iteration - 1) / n_iterations * 100: self.progress_var.set(p))

                if iteration > 1 and prop_summary_df is not None and t_col is not None:
                    prop_map = dict(zip(prop_summary_df['Bar Property ID'], prop_summary_df[t_col]))
                    self._update_bar_property_dim1(working_excel, prop_map)

                # ---- STAGE A: UNDERSTANDING STRUCTURE TYPE ----
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

                stru_bdfs = self._find_all_stru_bdfs(base_model_folder)
                dim1_by_pid = {}
                for p in stru_bdfs:
                    dim1_by_pid.update(self._extract_pbarl_dim1(p))
                self._update_dtcg_t1_tfem(working_excel, dim1_by_pid)

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

                self._wait_for_exe_to_finish(os.path.basename(self.le_exe_path.get().strip()))

                # ---- STAGE C: JOINT RUN ----
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

                # ---- STAGE D: SIZING SUMMARY ----
                prop_summary_df, t_col = self._build_flight_hours_summary(joint_run_dir, target_life)
                if prop_summary_df is None:
                    self.log(f"\nSTOPPED (iteration {iteration}): no sizing summary could be built "
                              "- not continuing to further iterations. See the log above.")
                    return

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

    def _refresh_tab3_listboxes(self):
        t3 = self.tab3_tool
        t3.bdf_listbox.delete(0, tk.END)
        for p in t3.bdf_paths:
            t3.bdf_listbox.insert(tk.END, os.path.basename(p))
        t3.maneuver_listbox.delete(0, tk.END)
        for p in t3.maneuver_bdfs:
            t3.maneuver_listbox.insert(tk.END, os.path.basename(p))


def main():
    root = tk.Tk()
    app = IntegratedBDFRFTool(root)
    root.mainloop()


if __name__ == "__main__":
    main()
