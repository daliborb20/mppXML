import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
import pandas as pd
import xml.etree.ElementTree as ET
from decimal import Decimal, InvalidOperation
from datetime import datetime
import os, csv, traceback

try:
    import pyodbc
except Exception:
    pyodbc = None
try:
    import pymssql
except Exception:
    pymssql = None

EMBEDDED_KONTA_MAP = {}
EMBEDDED_KONTA_META = {}

TIP_MAP = {
    'Tekući promet': 0,
    'Otvaranje p. knjiga': 1,
    'Zatvaranje p. knjiga': 2,
    'Izvod': 20,
    'Ulazni racuni': 21,
    'Uvoz': 22,
    'Maloprodaja': 23,
    'Izlazni racuni': 24,
    'Zarade': 24,
    'Nivelacije': 26,
    'Kursne razlike': 27,
    'Pdv nalog': 28,
    'Amortizacija': 29,
    'Vremenska razgranicenja (AVR & PVR)': 30,
    'Putni nalog': 31,
    'Izvoz': 37,
    'Asignacije, Kompenzacije, Cesije': 38,
}
TIP_OPCIJE = list(TIP_MAP.keys())

MAIN_REQUIRED = ['konto','duguje','potražuje','poslovni partner','dokument','datum promene','opis']

def normalize_header(h): return (h or '').strip().lower()

def norm_konto(s):
    if s is None:
        return ''
    s = str(s)
    if s.endswith('.0'):
        s = s[:-2]
    s = s.strip().replace(' ', '')
    for ch in ['.', '-', '/', '\\']:
        s = s.replace(ch, '')
    return s

def find_columns(df, required):
    normalized = {normalize_header(c): c for c in df.columns}
    mapping, missing = {}, []
    alt_map = {
        'potražuje': ['potrazuje','potrazue','potrazuj'],
        'duguje': ['dug','duznik'],
        'datum promene': ['datum','datum_promene','datum promjena','datum prom','datumpromene','datprom'],
        'poslovni partner': ['partner','poslovni_partner','poslovnipartner']
    }
    for col in required:
        if col in normalized:
            mapping[col] = normalized[col]
        else:
            found = None
            for a in alt_map.get(col, []):
                a_norm = normalize_header(a)
                if a_norm in normalized:
                    found = normalized[a_norm]; break
            if found: mapping[col]=found
            else: missing.append(col)
    return mapping, missing

def parse_amount(x):
    if x is None or (isinstance(x,float) and pd.isna(x)) or str(x).strip() == '':
        return None
    s = str(x).strip()
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    else:
        if ',' in s:
            s = s.replace('.', '').replace(',', '.')
    try:
        q = Decimal(s); return f'{q:.4f}'
    except InvalidOperation:
        return None

def parse_date_to_iso_tz(x):
    dt = pd.to_datetime(x, dayfirst=True, errors='coerce')
    if pd.isna(dt): return None
    return dt.strftime('%Y-%m-%dT00:00:00+02:00')

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Dalibor Bogicevic: PSIT MPP XLSX u XML generator')
        self.geometry('1400x1000')
        self.minsize(1100, 800)

        # --- Stilovi ---
        self.setup_styles()

        self.configure(bg='#e0e8f0')

        # --- Promenljive ---
        self.xlsx_path = tk.StringVar()
        self.out_path = tk.StringVar()
        self.napomena = tk.StringVar(value='Generisano iz XLSX')
        self.tip_naloga_var = tk.StringVar(value='Tekući promet')
        self.preduzece_var = tk.StringVar(value='(Nije učitano)')
        self.sifra_preduzeca = tk.StringVar(value='')
        self.sql_server = tk.StringVar(value='GTRS24MPP')
        self.sql_instance = tk.StringVar(value='')
        self.sql_port = tk.StringVar(value='1433')
        self.sql_database = tk.StringVar(value='mAS2')
        self.sql_windows_auth = tk.BooleanVar(value=True)
        self.sql_username = tk.StringVar(value='sa')
        self.sql_password = tk.StringVar(value='')
        self.status = tk.StringVar(value=f'Spremno. Fallback mapa: {len(EMBEDDED_KONTA_MAP)} konta.')
        
        self.preduzeca = []
        self.df = None
        self._debug_rows = []
        self._sql_konta_map = None
        self._sql_konta_meta = None

        # --- Glavni okvir ---
        main_frame = ttk.Frame(self, padding="15", style='App.TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(3, weight=3) # Preview
        main_frame.grid_rowconfigure(4, weight=2) # Log

        # --- Gornji deo: Podaci i SQL ---
        top_section = ttk.Frame(main_frame, style='App.TFrame')
        top_section.grid(row=0, column=0, sticky='ew', pady=(0, 10))
        top_section.grid_columnconfigure(1, weight=1)
        
        # --- Okvir za osnovne podatke (levo) ---
        input_frame = ttk.LabelFrame(top_section, text="Osnovni Podaci", padding="10")
        input_frame.grid(row=0, column=0, sticky='ns', padx=(0, 10))
        pad = {'padx': 5, 'pady': 6}

        ttk.Label(input_frame, text='Ulazni XLSX:').grid(row=0, column=0, sticky='w', **pad)
        ttk.Entry(input_frame, textvariable=self.xlsx_path, width=50).grid(row=1, column=0, sticky='ew', **pad)
        ttk.Button(input_frame, text='Odaberi…', command=self.choose_xlsx).grid(row=1, column=1, sticky='ew', **pad)

        ttk.Label(input_frame, text='Tip naloga:').grid(row=2, column=0, sticky='w', **pad)
        self.tip_combo = ttk.Combobox(input_frame, textvariable=self.tip_naloga_var, state='readonly', values=TIP_OPCIJE, width=48)
        self.tip_combo.grid(row=3, column=0, sticky='ew', **pad)

        ttk.Label(input_frame, text='Preduzeće:').grid(row=4, column=0, sticky='w', **pad)
        self.preduzece_combo = ttk.Combobox(input_frame, textvariable=self.preduzece_var, state='readonly', width=48)
        self.preduzece_combo.grid(row=5, column=0, sticky='ew', **pad)
        ttk.Button(input_frame, text='Učitaj iz SQL', command=self.load_preduzeca_sql).grid(row=5, column=1, sticky='ew', **pad)
        
        # --- SQL okvir (desno) ---
        sql_frame = ttk.LabelFrame(top_section, text="SQL Server Konekcija", padding="10")
        sql_frame.grid(row=0, column=1, sticky='nsew')
        sql_frame.grid_columnconfigure(1, weight=1)
        sql_frame.grid_columnconfigure(3, weight=1)

        ttk.Label(sql_frame, text='Server:').grid(row=0, column=0, sticky='e', **pad)
        ttk.Entry(sql_frame, textvariable=self.sql_server).grid(row=0, column=1, sticky='ew', **pad)
        ttk.Label(sql_frame, text='Instance:').grid(row=0, column=2, sticky='e', **pad)
        ttk.Entry(sql_frame, textvariable=self.sql_instance).grid(row=0, column=3, sticky='ew', **pad)
        ttk.Label(sql_frame, text='Port:').grid(row=0, column=4, sticky='e', **pad)
        ttk.Entry(sql_frame, textvariable=self.sql_port, width=8).grid(row=0, column=5, sticky='w', **pad)

        ttk.Checkbutton(sql_frame, text='Windows autentikacija', variable=self.sql_windows_auth).grid(row=1, column=0, columnspan=2, sticky='w', **pad)
        ttk.Label(sql_frame, text='Korisnik:').grid(row=2, column=0, sticky='e', **pad)
        ttk.Entry(sql_frame, textvariable=self.sql_username).grid(row=2, column=1, sticky='ew', **pad)
        ttk.Label(sql_frame, text='Lozinka:').grid(row=2, column=2, sticky='e', **pad)
        ttk.Entry(sql_frame, textvariable=self.sql_password, show='*').grid(row=2, column=3, sticky='ew', **pad)
        ttk.Label(sql_frame, text='Baza:').grid(row=2, column=4, sticky='e', **pad)
        ttk.Entry(sql_frame, textvariable=self.sql_database, width=15).grid(row=2, column=5, sticky='w', **pad)
        
        btn_frame = ttk.Frame(sql_frame)
        btn_frame.grid(row=3, column=0, columnspan=6, sticky='e', pady=(10,0))
        ttk.Button(btn_frame, text='Test konekcije', command=self.test_sql).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text='Učitaj konta iz SQL', command=self.load_konta_sql).pack(side=tk.LEFT)

        # --- Akcije i izlaz ---
        action_frame = ttk.LabelFrame(main_frame, text="Izlaz i Generisanje", padding="10")
        action_frame.grid(row=1, column=0, sticky='ew', pady=10)
        action_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(action_frame, text='Napomena / Spoljni broj:').grid(row=0, column=0, sticky='w', **pad)
        self.napomena_entry = ttk.Entry(action_frame, textvariable=self.napomena)
        self.napomena_entry.grid(row=0, column=1, columnspan=2, sticky='ew', **pad)

        ttk.Label(action_frame, text='Izlazni XML:').grid(row=1, column=0, sticky='w', **pad)
        ttk.Entry(action_frame, textvariable=self.out_path).grid(row=1, column=1, sticky='ew', **pad)
        ttk.Button(action_frame, text='Sačuvaj kao…', command=self.choose_xml).grid(row=1, column=2, padx=5)
        
        ttk.Button(action_frame, text='Generiši XML', command=self.generate, style='Accent.TButton').grid(row=2, column=1, columnspan=2, pady=(10,0), sticky='e')

        # --- Statusna linija ---
        ttk.Label(main_frame, textvariable=self.status, style='Status.TLabel').grid(row=2, column=0, sticky='ew', pady=5)

        # --- Pregled (Preview) ---
        preview_frame = ttk.LabelFrame(main_frame, text="Pregled XLSX (prvih 20 redova)", padding="10")
        preview_frame.grid(row=3, column=0, sticky='nsew', pady=5)
        self.tree = ttk.Treeview(preview_frame, show='headings')
        yscroll = ttk.Scrollbar(preview_frame, orient='vertical', command=self.tree.yview)
        xscroll = ttk.Scrollbar(preview_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        
        yscroll.pack(side='right', fill='y')
        xscroll.pack(side='bottom', fill='x')
        self.tree.pack(fill='both', expand=True)

        # --- Log ---
        log_frame = ttk.LabelFrame(main_frame, text="Log (uživo)", padding="10")
        log_frame.grid(row=4, column=0, sticky='nsew', pady=5)
        self.log = ScrolledText(log_frame, height=10, wrap='word', state='disabled', relief='flat',
                                bg='#ffffff', fg='#333333', font=('Consolas', 10))
        self.log.pack(fill='both', expand=True)

    def setup_styles(self):
        """Konfiguriše stilove za moderan izgled aplikacije."""
        BG_COLOR = '#e0e8f0'
        FRAME_BG = '#f7f9fc'
        TEXT_COLOR = '#333333'
        BUTTON_BG = '#0078d7'
        BUTTON_FG = '#ffffff'
        BUTTON_HOVER = '#005a9e'
        ACCENT_BG = '#28a745'
        ACCENT_HOVER = '#218838'
        ENTRY_BG = '#ffffff'
        FONT_NORMAL = ('Segoe UI', 10)
        FONT_BOLD = ('Segoe UI', 10, 'bold')
        FONT_TITLE = ('Segoe UI', 11, 'bold')
        
        style = ttk.Style(self)
        style.theme_use('clam')

        # Okviri
        style.configure('App.TFrame', background=BG_COLOR)
        style.configure('TFrame', background=FRAME_BG)
        
        # LabelFrame
        style.configure('TLabelframe', background=FRAME_BG, bordercolor='#c0c8d0', relief='solid', borderwidth=1)
        style.configure('TLabelframe.Label', background=FRAME_BG, foreground=TEXT_COLOR, font=FONT_TITLE, padding=(10, 5))
        
        # Label
        style.configure('TLabel', background=FRAME_BG, foreground=TEXT_COLOR, font=FONT_NORMAL, padding=5)
        style.configure('Status.TLabel', background=BG_COLOR, foreground='#555555', font=('Segoe UI', 9))

        # Entry i Combobox
        style.configure('TEntry', font=FONT_NORMAL, fieldbackground=ENTRY_BG, bordercolor='#999999', lightcolor=FRAME_BG, darkcolor=FRAME_BG, padding=5)
        style.map('TEntry', bordercolor=[('focus', BUTTON_BG)])
        style.configure('TCombobox', font=FONT_NORMAL, fieldbackground=ENTRY_BG, selectbackground=BUTTON_BG, selectforeground=BUTTON_FG, bordercolor='#999999', padding=5)
        self.option_add('*TCombobox*Listbox.font', FONT_NORMAL)

        # Dugmad
        style.configure('TButton', font=FONT_BOLD, background=BUTTON_BG, foreground=BUTTON_FG, padding=(10, 5), borderwidth=0, relief='flat')
        style.map('TButton', background=[('active', BUTTON_HOVER), ('!disabled', BUTTON_BG)])
        
        style.configure('Accent.TButton', font=FONT_BOLD, background=ACCENT_BG, foreground=BUTTON_FG, padding=(15, 8))
        style.map('Accent.TButton', background=[('active', ACCENT_HOVER), ('!disabled', ACCENT_BG)])
        
        # Checkbutton
        style.configure('TCheckbutton', background=FRAME_BG, font=FONT_NORMAL, foreground=TEXT_COLOR)
        style.map('TCheckbutton', indicatorcolor=[('selected', BUTTON_BG), ('!selected', '#cccccc')], background=[('active', '#e5e5e5')])
        
        # Treeview
        style.configure("Treeview", background="#ffffff", foreground=TEXT_COLOR, rowheight=28, fieldbackground="#ffffff", font=FONT_NORMAL, borderwidth=1, relief='solid')
        style.map('Treeview', background=[('selected', '#0078d7')], foreground=[('selected', '#ffffff')])
        style.configure("Treeview.Heading", background="#d0d8e0", foreground=TEXT_COLOR, font=FONT_BOLD, padding=8, relief='flat')
        style.map("Treeview.Heading", background=[('active', '#c0c8d0')])
        
    def _log(self, msg):
        try:
            self.log.configure(state='normal')
            self.log.insert('end', f"{datetime.now().strftime('%H:%M:%S')}  {msg}\n")
            self.log.see('end')
            self.log.configure(state='disabled')
        except Exception:
            pass

    def choose_xlsx(self):
        path = filedialog.askopenfilename(title='Odaberite XLSX', filetypes=[('Excel','*.xlsx')])
        if path:
            self.xlsx_path.set(path)
            self._log(f'Odabran XLSX: {path}')
            self.load_preview()

    def choose_xml(self):
        path = filedialog.asksaveasfilename(title='Sačuvaj XML kao', defaultextension='.xml', filetypes=[('XML','*.xml')])
        if path:
            self.out_path.set(path)
            self._log(f'Odabrana izlazna putanja: {path}')

    def _open_sql(self):
        server = self.sql_server.get().strip()
        instance = self.sql_instance.get().strip()
        port = self.sql_port.get().strip()
        database = self.sql_database.get().strip()
        self._log(f"Priprema konekcije: server='{server}', instance='{instance}', port='{port}', baza='{database}'")
        if self.sql_windows_auth.get():
            if not pyodbc:
                raise RuntimeError("Windows auth tražen, ali pyodbc nije instaliran. 'pip install pyodbc'")
            target = server
            if instance: target = f"{server}\\{instance}"
            elif port: target = f"{server},{port}"
            drivers_pref = ['ODBC Driver 18 for SQL Server','ODBC Driver 17 for SQL Server','SQL Server Native Client 11.0','SQL Server']
            try: available = list(pyodbc.drivers())
            except Exception: available = []
            self._log(f'Dostupni ODBC driveri: {available}')
            for drv in drivers_pref:
                if (not available) or (drv in available):
                    cs = f"DRIVER={{{drv}}};SERVER={target};DATABASE={database};Trusted_Connection=Yes;TrustServerCertificate=Yes;Encrypt=No;"
                    self._log(f"Pokušavam ODBC driver: '{drv}' → SERVER={target}; DATABASE={database} (Windows auth)")
                    try:
                        cn = pyodbc.connect(cs, timeout=5)
                        self._log(f"ODBC uspeh sa driverom: '{drv}'")
                        return cn
                    except Exception as e:
                        self._log(f"ODBC neuspeh sa '{drv}': {e}")
                        continue
            raise RuntimeError('ODBC Windows auth nije dostupan ili konekcija odbijena.')
        if not pymssql:
            raise RuntimeError("pymssql nije instaliran. 'pip install pymssql'")
        user = self.sql_username.get().strip()
        pwd = self.sql_password.get()
        port_i = int(port or '1433')
        self._log(f'Pokušavam pymssql (SQL auth) → server={server}, port={port_i}, baza={database}, user={user}')
        cn = pymssql.connect(server=server, user=user, password=pwd, database=database, port=port_i, tds_version='7.4', login_timeout=5)
        self._log('pymssql konekcija uspešna.')
        return cn

    def test_sql(self):
        try:
            cn = self._open_sql()
            cn.close()
            self._log('Test konekcije: USPEH')
            messagebox.showinfo('Uspeh', 'Konekcija uspešna! (HYBRID)')
        except Exception as e:
            self._log('Test konekcije: NEUSPEH\n' + ''.join(traceback.format_exception_only(type(e), e)).strip())
            messagebox.showerror('Greška', f'Neuspešna konekcija:\n{e}')

    def load_preduzeca_sql(self):
        try:
            cn = self._open_sql()
            cur = cn.cursor()
            q = 'SELECT cp_preduzece_id, CAST(sifra AS varchar(64)) AS sifra, CAST(naziv AS varchar(255)) AS naziv FROM dbo.cp_preduzece ORDER BY sifra'
            self._log(f'SQL upit: {q}')
            cur.execute(q)
            rows = cur.fetchall()
            cn.close()
            self.preduzeca = [{'id': int(r[0]), 'sifra': str(r[1] or ''), 'naziv': str(r[2] or '')} for r in rows]
            if not self.preduzeca:
                self._log('cp_preduzece: 0 redova')
                messagebox.showwarning('Info', 'Nije pronađeno nijedno preduzeće u cp_preduzece.'); return
            disp = [f"{p['sifra']} — {p['naziv']}" for p in self.preduzeca]
            self.preduzece_combo.configure(values=disp)
            self.preduzece_combo.current(0)
            self.preduzece_var.set(disp[0])
            self.sifra_preduzeca.set(self.preduzeca[0]['sifra'])
            self._log(f'Učitano preduzeća: {len(disp)}')
            messagebox.showinfo('OK', f'Učitano preduzeća: {len(disp)}')
        except Exception as e:
            self._log('Greška pri učitavanju preduzeća\n' + ''.join(traceback.format_exception_only(type(e), e)).strip())
            messagebox.showerror('Greška', f'Neuspešno učitavanje preduzeća:\n{e}')

    def load_konta_sql(self):
        try:
            cn = self._open_sql()
            cur = cn.cursor()
            q = ('SELECT fk_kp_konto_id, CAST(Broj AS varchar(64)) AS Broj, '
                 'CAST(Naziv AS varchar(255)) AS Naziv FROM dbo.fk_kp_konto')
            self._log('SQL upit (konta): ' + q)
            cur.execute(q)
            rows = cur.fetchall()
            cn.close()
            m = {}
            meta = {}
            for r in rows:
                kid = int(r[0]); broj = str(r[1] or '').strip(); broj_norm = norm_konto(broj)
                if not broj_norm: continue
                m[broj_norm] = kid
                meta[kid] = {'Broj': broj_norm, 'Naziv': str(r[2] or '')}
            self._sql_konta_map = m
            self._sql_konta_meta = meta
            self.status.set(f'Učitano iz SQL: {len(m)} konta')
            self._log(f'Mapa konta iz SQL: {len(m)} unosa')
            messagebox.showinfo('OK', f'Učitano iz SQL: {len(m)} konta')
        except Exception as e:
            self._sql_konta_map = None
            self._sql_konta_meta = None
            self._log('Greška pri učitavanju konta\n' + ''.join(traceback.format_exception_only(type(e), e)).strip())
            messagebox.showerror('Greška', f'Neuspelo učitavanje konta iz SQL:\n{e}')

    def load_preview(self):
        try:
            df = pd.read_excel(self.xlsx_path.get(), dtype=str)
            df = df.dropna(how='all')
            self.df = df
            self.show_preview(df)
            n = len(self._sql_konta_map) if self._sql_konta_map else len(EMBEDDED_KONTA_MAP)
            src = 'SQL' if self._sql_konta_map else 'EMBEDDED'
            self.status.set(f'Učitan XLSX. Konta dostupno: {n} (izvor: {src})')
            self._log(f'XLSX učitan. Kolone: {list(df.columns)}. Redova (ukupno): {len(df)}')
            missing = [c for c in MAIN_REQUIRED if normalize_header(c) not in [normalize_header(x) for x in df.columns]]
            if missing: self._log(f'UPOZORENJE: Moguće nedostaju kolone: {missing}')
        except Exception as e:
            self._log('Ne može da učita XLSX\n' + ''.join(traceback.format_exception_only(type(e), e)).strip())
            messagebox.showerror('Greška', f'Ne može da učita XLSX:\n{e}')

    def show_preview(self, df):
        for c in self.tree.get_children(): self.tree.delete(c)
        self.tree['columns'] = list(df.columns)
        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=140, stretch=True)
        for _, row in df.head(20).iterrows():
            self.tree.insert('', 'end', values=[str(v) for v in row])

    def _write_debug_csv(self, path):
        try:
            with open(path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=['row','status','reason','konto_raw','konto_norm'])
                writer.writeheader()
                for r in self._debug_rows: writer.writerow(r)
            self._log(f'Debug log zapisan: {path}')
        except Exception as e:
            self._log(f'Ne mogu da upišem debug CSV: {e}')

    def _current_konta_map(self):
        return self._sql_konta_map if self._sql_konta_map else EMBEDDED_KONTA_MAP

    def _current_konta_meta(self):
        return self._sql_konta_meta if self._sql_konta_meta else EMBEDDED_KONTA_META

    def generate(self):
        self._debug_rows = [] # Clear debug rows for each run
        try:
            if self.df is None:
                messagebox.showwarning('Upozorenje', 'Prvo učitaj XLSX.')
                self._log('Generate: nema učitanog XLSX-a')
                return
            if not self.sifra_preduzeca.get().strip():
                messagebox.showwarning('Upozorenje', 'Prvo učitaj i odaberi preduzeće iz SQL-a.')
                self._log('Generate: nije odabrano preduzeće')
                return
            mapping, missing = find_columns(self.df, MAIN_REQUIRED)
            if missing:
                self._log(f'Generate: nedostaju kolone: {missing}')
                messagebox.showerror('Greška', 'Nedostaju kolone: ' + ', '.join(missing)); return
            tip_name = self.tip_naloga_var.get().strip()
            tip_id = TIP_MAP.get(tip_name, 24)
            dates_iso = [d for d in (parse_date_to_iso_tz(v) for v in self.df[mapping['datum promene']]) if d]
            header_date = dates_iso[0] if dates_iso else datetime.now().strftime('%Y-%m-%dT00:00:00+02:00')
            nalog_id = 900000
            root = ET.Element('Dokumenti')
            nalog = ET.SubElement(root, 'Nalog_za_knjiženje')
            ET.SubElement(nalog, 'Šifra_x0020_preduzeca').text = self.sifra_preduzeca.get().strip()
            ET.SubElement(nalog, 'fk_nk_nalog_za_knjizenje_id').text = str(nalog_id)
            ET.SubElement(nalog, 'Status').text = '2'
            ET.SubElement(nalog, 'tip_x0020_id').text = str(tip_id)
            ET.SubElement(nalog, 'Tip').text = tip_name
            ET.SubElement(nalog, 'Broj').text = f'<{nalog_id}>'
            ET.SubElement(nalog, 'Org_x0020_broj').text = f'<{nalog_id}>'
            ET.SubElement(nalog, 'Datum').text = header_date
            note = self.napomena.get().strip() or 'Generisano iz XLSX'
            ET.SubElement(nalog, 'Napomena').text = note
            ET.SubElement(nalog, 'Spoljni_x0020_broj').text = note
            used_kids = set()
            konta_map = self._current_konta_map()
            konta_meta = self._current_konta_meta()
            rb = 1
            not_in_map = 0
            for idx, row in self.df.iterrows():
                row_num = int(idx) + 2 # Excel row number (1-based index + header)
                konto_raw = row.get(mapping['konto'], '')
                konto_broj = norm_konto(konto_raw)
                if not konto_broj:
                    self._debug_rows.append({'row': row_num, 'status':'SKIP','reason':'konto empty','konto_raw':str(konto_raw),'konto_norm':konto_broj})
                    continue
                if konto_broj not in konta_map:
                    not_in_map += 1
                    self._debug_rows.append({'row': row_num, 'status':'SKIP','reason':'konto not in map','konto_raw':str(konto_raw),'konto_norm':konto_broj})
                    continue
                konto_id = konta_map[konto_broj]
                used_kids.add(konto_id)
                stavka = ET.SubElement(root, 'Stavka_naloga_za_knjizenje')
                ET.SubElement(stavka, 'fk_nk_stavka_naloga_za_knjizenje_id').text = str(nalog_id + rb)
                ET.SubElement(stavka, 'fk_nk_nalog_za_knjizenje_id').text = str(nalog_id)
                ET.SubElement(stavka, 'fk_kp_konto_id').text = str(konto_id)
                ET.SubElement(stavka, 'Redni_x0020_broj').text = str(rb)
                d_prom = parse_date_to_iso_tz(row.get(mapping['datum promene']))
                if d_prom: ET.SubElement(stavka, 'Datum_x0020_promene').text = d_prom
                doc = '' if pd.isna(row.get(mapping['dokument'])) else str(row.get(mapping['dokument'])).strip()
                if doc: ET.SubElement(stavka, 'Broj_x0020_dokumenta').text = doc
                duguje = parse_amount(row.get(mapping['duguje']))
                potrazuje = parse_amount(row.get(mapping['potražuje']))
                added = False
                if duguje and float(duguje) != 0.0:
                    ET.SubElement(stavka, 'Duguje').text = duguje; added = True
                if potrazuje and float(potrazuje) != 0.0:
                    ET.SubElement(stavka, 'Potrazuje').text = potrazuje; added = True
                if not added:
                    root.remove(stavka)
                    self._debug_rows.append({'row': row_num, 'status':'SKIP','reason':'zero amounts','konto_raw':str(konto_raw),'konto_norm':konto_broj})
                    continue
                opis = '' if pd.isna(row.get(mapping['opis'])) else str(row.get(mapping['opis'])).strip()
                if opis: ET.SubElement(stavka, 'Opis').text = opis
                ET.SubElement(stavka, 'Subanalitika').text = ''
                ET.SubElement(stavka, 'Valuta_x0020_ID').text = '1'
                ET.SubElement(stavka, 'Kurs').text = '0'
                rb += 1
            for kid in sorted(used_kids):
                meta = konta_meta.get(kid, {'Broj':'', 'Naziv':''})
                kb = ET.SubElement(root, 'Konto')
                ET.SubElement(kb, 'fk_kp_konto_id').text = str(kid)
                ET.SubElement(kb, 'Broj').text = str(meta.get('Broj',''))
                ET.SubElement(kb, 'Naziv').text = str(meta.get('Naziv',''))
                ET.SubElement(kb, 'Dozvoljeno_x0020_knjiženje').text = '1'
                ET.SubElement(kb, 'Devizni').text = '0'
            out_path = self.out_path.get() or (os.path.splitext(self.xlsx_path.get())[0] + '_HYBRID_v4c_FIXED.xml')
            self.out_path.set(out_path)
            debug_csv = os.path.join(os.path.dirname(out_path) or '.', 'xml_import_debug.csv')
            ET.ElementTree(root).write(out_path, encoding='utf-8', xml_declaration=True)
            self._write_debug_csv(debug_csv)
            if rb == 1:
                try: os.remove(out_path)
                except Exception: pass
                self._log('Nijedna stavka nije generisana — verovatno neprepoznata konta ili nula iznosi.')
                messagebox.showerror('Greška', 'Nijedna stavka nije generisana. Pogledaj \'xml_import_debug.csv\'.')
                return
            self._log(f'GENERISANO OK: {out_path}. Stavki: {rb-1}. Konto not-in-map: {not_in_map}.')
            src = ('SQL' if self._sql_konta_map else 'EMBEDDED')
            messagebox.showinfo('Gotovo', 'XML generisan (' + src + ' mapa):\n' + out_path + '\n\nDebug log:\n' + debug_csv)
        except Exception as e:
            self._log('Greška u generate()\n' + traceback.format_exc())
            messagebox.showerror('Greška', f'Neuspeh generisanja XML-a:\n{e}')

if __name__ == '__main__':
    App().mainloop()