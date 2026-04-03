import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime, date
import xlwt
from io import BytesIO

# ─── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Analisi 44020 – Fatture da Ricevere",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}

/* Dark professional theme */
.stApp {
    background: #0f1117;
    color: #e8eaf0;
}

.main .block-container {
    padding-top: 2rem;
    max-width: 1400px;
}

/* Header */
.header-band {
    background: linear-gradient(135deg, #1a1f2e 0%, #0d1520 100%);
    border: 1px solid #2a3a5c;
    border-radius: 12px;
    padding: 2rem 2.5rem;
    margin-bottom: 2rem;
    position: relative;
    overflow: hidden;
}
.header-band::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
    background: linear-gradient(90deg, #3b82f6, #06b6d4, #3b82f6);
}
.header-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.6rem;
    font-weight: 600;
    color: #e8eaf0;
    margin: 0;
    letter-spacing: -0.02em;
}
.header-sub {
    font-size: 0.85rem;
    color: #64748b;
    margin-top: 0.3rem;
    font-family: 'IBM Plex Mono', monospace;
}

/* Cards */
.card {
    background: #1a1f2e;
    border: 1px solid #2a3a5c;
    border-radius: 10px;
    padding: 1.5rem;
    margin-bottom: 1.2rem;
}
.card-ok {
    border-left: 4px solid #22c55e;
    background: linear-gradient(90deg, #0f2918 0%, #1a1f2e 30%);
}
.card-warn {
    border-left: 4px solid #f59e0b;
    background: linear-gradient(90deg, #1f1a08 0%, #1a1f2e 30%);
}
.card-error {
    border-left: 4px solid #ef4444;
    background: linear-gradient(90deg, #1f0a0a 0%, #1a1f2e 30%);
}
.card-info {
    border-left: 4px solid #3b82f6;
    background: linear-gradient(90deg, #0a1220 0%, #1a1f2e 30%);
}

.card-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.8rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: #64748b;
    margin-bottom: 0.5rem;
}
.card-value {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.8rem;
    font-weight: 600;
    color: #e8eaf0;
}
.card-label {
    font-size: 0.82rem;
    color: #94a3b8;
    margin-top: 0.2rem;
}

/* Badge */
.badge-ok    { background:#14532d; color:#4ade80; padding:2px 10px; border-radius:20px; font-size:.75rem; font-weight:600; font-family:'IBM Plex Mono',monospace; }
.badge-err   { background:#450a0a; color:#f87171; padding:2px 10px; border-radius:20px; font-size:.75rem; font-weight:600; font-family:'IBM Plex Mono',monospace; }
.badge-warn  { background:#451a03; color:#fb923c; padding:2px 10px; border-radius:20px; font-size:.75rem; font-weight:600; font-family:'IBM Plex Mono',monospace; }
.badge-info  { background:#0c1a3a; color:#60a5fa; padding:2px 10px; border-radius:20px; font-size:.75rem; font-weight:600; font-family:'IBM Plex Mono',monospace; }

/* Section title */
.section-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.72rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.12em;
    color: #3b82f6;
    margin: 2rem 0 1rem 0;
    padding-bottom: 0.5rem;
    border-bottom: 1px solid #1e3a5f;
}

/* Check rows */
.check-row {
    display: flex;
    align-items: flex-start;
    gap: 1rem;
    padding: 1rem 1.2rem;
    border-radius: 8px;
    margin-bottom: 0.6rem;
}
.check-icon { font-size: 1.2rem; flex-shrink: 0; margin-top: 0.1rem; }
.check-body { flex: 1; }
.check-title { font-weight: 600; font-size: 0.95rem; color: #e8eaf0; }
.check-detail { font-size: 0.83rem; color: #94a3b8; margin-top: 0.2rem; }
.check-amount { font-family: 'IBM Plex Mono', monospace; font-size: 0.85rem; }

/* Table overrides */
[data-testid="stDataFrame"] {
    border: 1px solid #2a3a5c;
    border-radius: 8px;
    overflow: hidden;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: #0d1117;
    border-right: 1px solid #1e3a5f;
}
[data-testid="stSidebar"] .stMarkdown { color: #94a3b8; }

/* Metric */
div[data-testid="metric-container"] {
    background: #1a1f2e;
    border: 1px solid #2a3a5c;
    border-radius: 8px;
    padding: 0.8rem 1rem;
}
div[data-testid="metric-container"] label { color: #64748b !important; font-family: 'IBM Plex Mono', monospace; font-size: 0.75rem !important; }
div[data-testid="metric-container"] [data-testid="stMetricValue"] { font-family: 'IBM Plex Mono', monospace; color: #e8eaf0 !important; }

/* Tabs */
.stTabs [data-baseweb="tab-list"] { background: #1a1f2e; border-radius: 8px; border: 1px solid #2a3a5c; padding: 4px; gap: 4px; }
.stTabs [data-baseweb="tab"] { background: transparent; color: #64748b; border-radius: 6px; font-family: 'IBM Plex Mono', monospace; font-size: 0.8rem; font-weight: 600; }
.stTabs [aria-selected="true"] { background: #3b82f6 !important; color: white !important; }

/* Alerts */
.stAlert { border-radius: 8px; }
</style>
""", unsafe_allow_html=True)


# ─── PARSER TXT ────────────────────────────────────────────────────────────────
def parse_amount(s):
    """Converti stringa importo italiano → float. '' o nan → None."""
    if s is None:
        return None
    s = str(s).strip().replace('\xa0', '').replace(' ', '')
    if s in ('', 'nan', 'None', '-'):
        return None
    s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except Exception:
        return None


def parse_date(s):
    """Converti stringa data gg/mm/aa → datetime.date. None se fallisce."""
    if s is None:
        return None
    s = str(s).strip()
    for fmt in ('%d/%m/%y', '%d/%m/%Y'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def parse_txt(content: str) -> pd.DataFrame:
    """Parser per il TXT tab-separated esportato da GB Software."""
    rows = []
    lines = content.splitlines()
    for line in lines:
        parts = line.split('\t')
        # Riga dati: la colonna 1 (indice 1) contiene '44020' o 'SAP'
        if len(parts) < 12:
            continue
        conto = parts[1].strip() if len(parts) > 1 else ''
        if conto not in ('44020', 'SAP'):
            continue
        # Struttura: [0]=empty [1]=conto [2]=DataReg [3]=NDoc [4]=DataDoc
        #            [5]=PrIva [6]=CO [7]=IV [8]=Causale [9]=Descrizione
        #            [10]=Dare [11]=Avere [12]=Saldo
        data_reg_s = parts[2].strip() if len(parts) > 2 else ''
        n_doc      = parts[3].strip() if len(parts) > 3 else ''
        data_doc_s = parts[4].strip() if len(parts) > 4 else ''
        causale    = parts[8].strip() if len(parts) > 8 else ''
        descrizione= parts[9].strip() if len(parts) > 9 else ''
        dare_s     = parts[10].strip() if len(parts) > 10 else ''
        avere_s    = parts[11].strip() if len(parts) > 11 else ''
        saldo_s    = parts[12].strip() if len(parts) > 12 else ''

        data_reg = parse_date(data_reg_s)
        data_doc = parse_date(data_doc_s)
        dare  = parse_amount(dare_s)
        avere = parse_amount(avere_s)
        saldo = parse_amount(saldo_s)

        if data_reg is None:
            continue
        # Saldo iniziale
        if conto == 'SAP':
            causale = 'SALDO_INIZIALE'

        rows.append({
            'conto': conto,
            'data_reg': data_reg,
            'n_doc': n_doc,
            'data_doc': data_doc,
            'causale': causale,
            'descrizione': descrizione,
            'dare': dare,
            'avere': avere,
            'saldo': saldo,
        })

    df = pd.DataFrame(rows)
    return df


def parse_xls(file_bytes: bytes) -> pd.DataFrame:
    """Parser per il file XLS (BIFF) con struttura GB Software."""
    import xlrd
    wb = xlrd.open_workbook(file_contents=file_bytes)
    ws = wb.sheet_by_index(0)
    rows = []
    for r in range(ws.nrows):
        row_vals = [ws.cell_value(r, c) for c in range(ws.ncols)]
        conto = str(row_vals[1]).strip() if len(row_vals) > 1 else ''
        if conto not in ('44020', 'SAP'):
            continue
        # Date: in XLS sono float (Excel serial)
        def xls_date(val):
            if val is None or val == '':
                return None
            if isinstance(val, float) and val > 0:
                try:
                    return datetime(*xlrd.xldate_as_tuple(val, wb.datemode)).date()
                except Exception:
                    return None
            if isinstance(val, str) and val.strip():
                return parse_date(val)
            return None

        data_reg  = xls_date(row_vals[2]) if len(row_vals) > 2 else None
        n_doc     = str(row_vals[3]).strip() if len(row_vals) > 3 else ''
        data_doc  = xls_date(row_vals[4]) if len(row_vals) > 4 else None
        causale   = str(row_vals[8]).strip() if len(row_vals) > 8 else ''
        descr     = str(row_vals[9]).strip() if len(row_vals) > 9 else ''

        def num(v):
            if isinstance(v, (int, float)) and not (isinstance(v, float) and np.isnan(v)):
                return float(v) if v != 0 else None
            return parse_amount(str(v))

        dare  = num(row_vals[10]) if len(row_vals) > 10 else None
        avere = num(row_vals[11]) if len(row_vals) > 11 else None
        saldo = num(row_vals[12]) if len(row_vals) > 12 else None

        if data_reg is None:
            continue
        if conto == 'SAP':
            causale = 'SALDO_INIZIALE'

        rows.append({
            'conto': conto,
            'data_reg': data_reg,
            'n_doc': n_doc,
            'data_doc': data_doc,
            'causale': causale,
            'descrizione': descr,
            'dare': dare,
            'avere': avere,
            'saldo': saldo,
        })
    return pd.DataFrame(rows)


def parse_pdf(file_bytes: bytes) -> pd.DataFrame:
    """Estrae righe dal PDF usando pdfplumber e le parsa."""
    import pdfplumber
    lines_all = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                lines_all.extend(text.splitlines())

    # Pattern date gg/mm/aa o gg/mm/aaaa
    date_pat = re.compile(r'\d{2}/\d{2}/\d{2,4}')
    amount_pat = re.compile(r'-?\d{1,3}(?:\.\d{3})*,\d{2}')

    rows = []
    for line in lines_all:
        line = line.strip()
        if not line:
            continue
        # Cerca almeno una data e almeno un importo
        dates = date_pat.findall(line)
        amounts = amount_pat.findall(line)
        if not dates or not amounts:
            continue
        # Prima data = data_reg
        data_reg = parse_date(dates[0])
        if data_reg is None:
            continue

        # Determina dare / avere / saldo dagli importi trovati
        # Logica: ultimo importo = saldo, gli altri = dare o avere
        def am(s): return parse_amount(s)

        dare_v, avere_v, saldo_v = None, None, None
        if len(amounts) == 1:
            saldo_v = am(amounts[0])
        elif len(amounts) == 2:
            saldo_v = am(amounts[-1])
            val = am(amounts[0])
            # Saldo negativo → avere; positivo → dare
            if saldo_v is not None and saldo_v < 0:
                avere_v = val
            else:
                dare_v = val
        elif len(amounts) >= 3:
            dare_v  = am(amounts[-3]) if am(amounts[-3]) else None
            avere_v = am(amounts[-2]) if am(amounts[-2]) else None
            saldo_v = am(amounts[-1])
            if dare_v == 0:
                dare_v = None
            if avere_v == 0:
                avere_v = None

        # Estrai descrizione (testo tra prima data e primi importi)
        # Prendi tutto il testo, rimuovi date e importi
        desc = line
        for d in dates:
            desc = desc.replace(d, ' ')
        for a in amounts:
            desc = desc.replace(a, ' ')
        desc = re.sub(r'\s+', ' ', desc).strip()
        # Rimuovi caratteri numerici residui (N.Doc ecc.)
        desc = re.sub(r'\b\d+\b', ' ', desc).strip()
        desc = re.sub(r'\s+', ' ', desc).strip()

        causale = ''
        if 'Rilevaz' in desc or 'AL03' in desc:
            causale = 'AL03'
        elif 'Fattura' in desc or 'FA' in desc:
            causale = 'FA'
        elif 'Giroconto' in desc or 'AL01' in desc:
            causale = 'AL01'
        elif 'Saldo anno' in desc or 'SAP' in desc:
            causale = 'SALDO_INIZIALE'

        rows.append({
            'conto': '44020',
            'data_reg': data_reg,
            'n_doc': '',
            'data_doc': parse_date(dates[1]) if len(dates) > 1 else None,
            'causale': causale,
            'descrizione': desc[:80],
            'dare': dare_v,
            'avere': avere_v,
            'saldo': saldo_v,
        })
    return pd.DataFrame(rows) if rows else pd.DataFrame()


# ─── LOGICA DI ANALISI ─────────────────────────────────────────────────────────
def fmt_eur(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return '–'
    return f"€ {v:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')


def analyze(df: pd.DataFrame, anno_bilancio: int, anno_prec: int, anno_succ: int):
    """
    Esegue tutte le verifiche richieste.
    Restituisce un dict con i risultati strutturati.
    """
    result = {}
    TOLL = 0.005  # tolleranza centesimi floating point

    # Filtra per anno
    def by_year(y):
        return df[df['data_reg'].apply(lambda d: d.year if d else 0) == y].copy()

    df_prec = by_year(anno_prec)
    df_bil  = by_year(anno_bilancio)
    df_succ = by_year(anno_succ)

    # ── Calcola saldi a fine anno ────────────────────────────────────────────
    def saldo_fine_anno(df_anno):
        """Ultimo saldo nell'anno."""
        if df_anno.empty:
            return None
        valid = df_anno[df_anno['saldo'].notna()]
        if valid.empty:
            return None
        return float(valid.iloc[-1]['saldo'])

    def saldo_inizio_anno(df_anno):
        """Primo saldo (saldo iniziale o prima riga)."""
        if df_anno.empty:
            return None
        row0 = df_anno[df_anno['causale'] == 'SALDO_INIZIALE']
        if not row0.empty:
            return float(row0.iloc[0]['avere']) if row0.iloc[0]['avere'] else 0.0
        valid = df_anno[df_anno['saldo'].notna()]
        return float(valid.iloc[0]['saldo']) if not valid.empty else None

    saldo_31_prec = saldo_fine_anno(df_prec)
    saldo_31_bil  = saldo_fine_anno(df_bil)
    saldo_31_succ = saldo_fine_anno(df_succ)

    # Saldo a zero durante anno bilancio: verifica che a inizio anno (dopo storni)
    # il saldo vada a zero, poi torni in avere
    # "Va a zero" = dopo le registrazioni dare di inizio anno il saldo è ~0
    def saldo_minimo_anno(df_anno):
        """Valore assoluto minimo del saldo durante l'anno."""
        valid = df_anno[df_anno['saldo'].notna()]
        if valid.empty:
            return None
        return float(valid['saldo'].abs().min())

    # Riga con saldo più vicino a zero nell'anno bilancio
    def riga_saldo_zero(df_anno):
        valid = df_anno[df_anno['saldo'].notna()].copy()
        if valid.empty:
            return None
        idx = valid['saldo'].abs().idxmin()
        return valid.loc[idx]

    saldo_min_bil = saldo_minimo_anno(df_bil)
    riga_zero = riga_saldo_zero(df_bil)

    # CHECK 1: Il saldo va a zero nell'anno bilancio
    va_a_zero = saldo_min_bil is not None and saldo_min_bil < TOLL
    saldo_chiude_avere_bil = saldo_31_bil is not None and saldo_31_bil < -TOLL

    result['check1'] = {
        'va_a_zero': va_a_zero,
        'saldo_min_bil': saldo_min_bil,
        'riga_zero': riga_zero,
        'chiude_avere': saldo_chiude_avere_bil,
        'saldo_31_bil': saldo_31_bil,
        'ok': va_a_zero and saldo_chiude_avere_bil,
    }

    # ── Righe avere fine anno bilancio e dare inizio anno successivo ─────────
    def righe_avere_31dic(df_anno, y):
        """Righe in avere al 31.12 dell'anno y."""
        d = date(y, 12, 31)
        mask = (
            df_anno['data_reg'].apply(lambda x: x == d) &
            df_anno['avere'].notna() &
            (df_anno['causale'] != 'SALDO_INIZIALE')
        )
        return df_anno[mask][['data_reg', 'descrizione', 'causale', 'dare', 'avere', 'saldo']].copy()

    def righe_dare_inizio_anno(df_anno, y):
        """Righe in dare nei primi mesi dell'anno y (storni)."""
        mask = (
            df_anno['data_reg'].apply(lambda x: x.year == y if x else False) &
            df_anno['dare'].notna() &
            (df_anno['causale'] != 'SALDO_INIZIALE') &
            (df_anno['causale'] != 'AL01')   # esclude giroconto apertura
        )
        return df_anno[mask][['data_reg', 'descrizione', 'causale', 'dare', 'avere', 'saldo']].copy()

    avere_bil   = righe_avere_31dic(df_bil, anno_bilancio)
    dare_succ   = righe_dare_inizio_anno(df_succ, anno_succ)

    sum_avere_bil = float(avere_bil['avere'].sum()) if not avere_bil.empty else 0.0
    sum_dare_succ = float(dare_succ['dare'].sum()) if not dare_succ.empty else 0.0

    # CHECK 2: Cifra per cifra (matching puntuale)
    def match_cifre(avere_rows, dare_rows):
        """
        Tenta di abbinare ogni riga avere con una riga dare di pari importo.
        Restituisce (matched_pairs, unmatched_avere, unmatched_dare).
        """
        avere_list = list(avere_rows.itertuples(index=True))
        dare_list  = list(dare_rows.itertuples(index=True))
        matched = []
        used_dare = set()

        for a in avere_list:
            a_val = round(float(a.avere), 2)
            found = False
            for d in dare_list:
                if d.Index in used_dare:
                    continue
                d_val = round(float(d.dare), 2)
                if abs(a_val - d_val) < TOLL:
                    matched.append((a, d))
                    used_dare.add(d.Index)
                    found = True
                    break
            if not found:
                pass  # rimarrà come unmatched

        unmatched_avere = [a for a in avere_list if not any(a.Index == m[0].Index for m in matched)]
        unmatched_dare  = [d for d in dare_list  if d.Index not in used_dare]
        return matched, unmatched_avere, unmatched_dare

    matched, unm_avere, unm_dare = match_cifre(avere_bil, dare_succ)
    check2_ok = len(unm_avere) == 0 and len(unm_dare) == 0

    result['check2'] = {
        'ok': check2_ok,
        'matched': matched,
        'unmatched_avere': unm_avere,
        'unmatched_dare': unm_dare,
        'sum_avere': sum_avere_bil,
        'sum_dare': sum_dare_succ,
        'avere_rows': avere_bil,
        'dare_rows': dare_succ,
    }

    # CHECK 3: Totali corrispondono (anche se cifre non abbinate 1:1)
    diff_totali = abs(sum_avere_bil - sum_dare_succ)
    result['check3'] = {
        'ok': diff_totali < TOLL,
        'sum_avere': sum_avere_bil,
        'sum_dare': sum_dare_succ,
        'diff': sum_avere_bil - sum_dare_succ,
    }

    # CHECK 4: Righe senza corrispondente
    result['check4'] = {
        'unmatched_avere': unm_avere,
        'unmatched_dare': unm_dare,
    }

    # ── Saldi di riepilogo ───────────────────────────────────────────────────
    result['saldi'] = {
        'saldo_31_prec': saldo_31_prec,
        'saldo_31_bil': saldo_31_bil,
        'saldo_31_succ': saldo_31_succ,
        'sum_avere_bil': sum_avere_bil,
        'sum_dare_succ': sum_dare_succ,
    }

    result['df_prec'] = df_prec
    result['df_bil']  = df_bil
    result['df_succ'] = df_succ
    result['avere_bil'] = avere_bil
    result['dare_succ'] = dare_succ

    return result


# ─── EXPORT EXCEL ──────────────────────────────────────────────────────────────
def build_excel_report(df: pd.DataFrame, res: dict, anno_bilancio: int, anno_prec: int, anno_succ: int) -> bytes:
    """Crea un file XLS con il report completo."""
    wb = xlwt.Workbook(encoding='utf-8')

    def add_sheet_df(name, df_in):
        ws = wb.add_sheet(name[:31])
        if df_in is None or df_in.empty:
            ws.write(0, 0, 'Nessun dato')
            return
        for c, col in enumerate(df_in.columns):
            ws.write(0, c, str(col))
        for r, row in enumerate(df_in.itertuples(index=False), 1):
            for c, val in enumerate(row):
                ws.write(r, c, str(val) if val is not None else '')

    # Foglio riepilogo
    ws_sum = wb.add_sheet('Riepilogo')
    ws_sum.write(0, 0, f'ANALISI CONTO 44020 – Anno bilancio {anno_bilancio}')
    ws_sum.write(2, 0, 'CHECK 1 – Saldo a zero e chiusura in avere')
    ws_sum.write(3, 0, 'Va a zero:');   ws_sum.write(3, 1, 'SÌ' if res['check1']['va_a_zero'] else 'NO')
    ws_sum.write(4, 0, 'Chiude in avere:'); ws_sum.write(4, 1, 'SÌ' if res['check1']['chiude_avere'] else 'NO')
    ws_sum.write(5, 0, f"Saldo 31/12/{anno_bilancio}:"); ws_sum.write(5, 1, res['check1']['saldo_31_bil'] or 0)

    ws_sum.write(7, 0, 'CHECK 2 – Matching cifra per cifra')
    ws_sum.write(8, 0, 'Risultato:'); ws_sum.write(8, 1, 'OK' if res['check2']['ok'] else 'ANOMALIA')
    ws_sum.write(9, 0, f"Totale avere {anno_bilancio}:"); ws_sum.write(9, 1, res['check3']['sum_avere'])
    ws_sum.write(10, 0, f"Totale dare {anno_succ}:"); ws_sum.write(10, 1, res['check3']['sum_dare'])
    ws_sum.write(11, 0, 'Differenza:'); ws_sum.write(11, 1, round(res['check3']['diff'], 2))

    # Fogli dati
    add_sheet_df(f'Avere_{anno_bilancio}', res['avere_bil'])
    add_sheet_df(f'Dare_{anno_succ}', res['dare_succ'])

    # Foglio non abbinati
    ws_unm = wb.add_sheet('Non_abbinati')
    ws_unm.write(0, 0, f'AVERE {anno_bilancio} SENZA CORRISPONDENTE IN DARE {anno_succ}')
    ws_unm.write(1, 0, 'Data'); ws_unm.write(1, 1, 'Descrizione'); ws_unm.write(1, 2, 'Avere')
    r = 2
    for row in res['check4']['unmatched_avere']:
        ws_unm.write(r, 0, str(row.data_reg))
        ws_unm.write(r, 1, str(row.descrizione))
        ws_unm.write(r, 2, float(row.avere))
        r += 1

    r += 2
    ws_unm.write(r, 0, f'DARE {anno_succ} SENZA CORRISPONDENTE IN AVERE {anno_bilancio}')
    r += 1
    ws_unm.write(r, 0, 'Data'); ws_unm.write(r, 1, 'Descrizione'); ws_unm.write(r, 2, 'Dare')
    r += 1
    for row in res['check4']['unmatched_dare']:
        ws_unm.write(r, 0, str(row.data_reg))
        ws_unm.write(r, 1, str(row.descrizione))
        ws_unm.write(r, 2, float(row.dare))
        r += 1

    # Mastrino completo
    add_sheet_df('Mastrino_completo', df)

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─── HELPERS UI ────────────────────────────────────────────────────────────────
def check_card(ok, title, detail, extra=''):
    cls = 'card-ok' if ok else 'card-error'
    icon = '✅' if ok else '❌'
    st.markdown(f"""
    <div class="card {cls}">
      <div style="display:flex;align-items:flex-start;gap:.8rem;">
        <span style="font-size:1.3rem">{icon}</span>
        <div>
          <div class="check-title">{title}</div>
          <div class="check-detail">{detail}</div>
          {'<div class="check-amount" style="margin-top:.4rem;color:#94a3b8;">'+extra+'</div>' if extra else ''}
        </div>
      </div>
    </div>""", unsafe_allow_html=True)


def df_table(df_in, height=300):
    """Mostra dataframe con stile scuro."""
    if df_in is None or df_in.empty:
        st.info('Nessuna riga.')
        return
    # Rinomina e formatta
    display = df_in.copy()
    for col in ['dare', 'avere', 'saldo']:
        if col in display.columns:
            display[col] = display[col].apply(lambda x: fmt_eur(x) if pd.notna(x) else '–')
    st.dataframe(display, use_container_width=True, height=height)


# ─── APP PRINCIPALE ────────────────────────────────────────────────────────────
def main():
    # ── Header ──────────────────────────────────────────────────────────────
    st.markdown("""
    <div class="header-band">
      <div class="header-title">Analisi Conto 44020</div>
      <div class="header-sub">GB Software · Fatture da Ricevere · Verifica scritture di assestamento</div>
    </div>
    """, unsafe_allow_html=True)

    # ── Sidebar ─────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("### ⚙️ Parametri")
        anno_bilancio = st.number_input("Anno oggetto del bilancio", value=2025, min_value=2000, max_value=2100, step=1)
        anno_prec = st.number_input("Anno precedente", value=anno_bilancio - 1, min_value=2000, max_value=2100, step=1)
        anno_succ = st.number_input("Anno successivo", value=anno_bilancio + 1, min_value=2000, max_value=2100, step=1)

        st.markdown("---")
        st.markdown("### 📂 Carica Mastrino")
        uploaded = st.file_uploader(
            "File mastrino conto 44020",
            type=['txt', 'xls', 'pdf'],
            help="Esporta da GB Software il mastrino conto 44020 su più anni. Formati accettati: TXT (tab-sep), XLS, PDF."
        )

        st.markdown("---")
        st.markdown("""
        <div style="font-size:.75rem;color:#475569;line-height:1.6;">
        <b style="color:#64748b">Legenda verifiche</b><br>
        ① Saldo a zero + chiusura avere<br>
        ② Matching cifra per cifra<br>
        ③ Totali corrispondenti<br>
        ④ Righe senza corrispondente
        </div>
        """, unsafe_allow_html=True)

    # ── Caricamento e parsing ────────────────────────────────────────────────
    if uploaded is None:
        st.markdown("""
        <div class="card card-info" style="text-align:center;padding:3rem;">
          <div style="font-size:2.5rem;margin-bottom:1rem;">📁</div>
          <div style="font-size:1.1rem;color:#94a3b8;margin-bottom:.5rem;">Carica il mastrino nella barra laterale</div>
          <div style="font-size:.85rem;color:#475569;">Formati accettati: TXT (tab-separato), XLS (vecchio formato), PDF</div>
        </div>
        """, unsafe_allow_html=True)
        return

    # Parsing
    with st.spinner("Parsing del file in corso…"):
        fname = uploaded.name.lower()
        raw = uploaded.read()
        df = pd.DataFrame()
        parse_error = None

        try:
            if fname.endswith('.txt'):
                text = raw.decode('utf-8', errors='replace')
                df = parse_txt(text)
            elif fname.endswith('.xls'):
                df = parse_xls(raw)
            elif fname.endswith('.pdf'):
                df = parse_pdf(raw)
            else:
                parse_error = "Formato non riconosciuto."
        except Exception as e:
            parse_error = str(e)

    if parse_error:
        st.error(f"Errore nel parsing: {parse_error}")
        return

    if df.empty:
        st.error("Nessuna riga trovata nel file. Verificare il formato.")
        return

    # Analisi
    with st.spinner("Analisi in corso…"):
        try:
            res = analyze(df, anno_bilancio, anno_prec, anno_succ)
        except Exception as e:
            st.error(f"Errore nell'analisi: {e}")
            import traceback; st.code(traceback.format_exc())
            return

    # ── KPI bar ─────────────────────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    saldi = res['saldi']
    with c1:
        v = saldi['saldo_31_prec']
        st.metric(f"Saldo 31/12/{anno_prec}", fmt_eur(v),
                  delta="Avere" if v and v < 0 else ("Dare" if v and v > 0 else "Zero"))
    with c2:
        v = saldi['saldo_31_bil']
        st.metric(f"Saldo 31/12/{anno_bilancio}", fmt_eur(v),
                  delta="Avere ✓" if v and v < 0 else ("Dare ⚠" if v and v > 0 else "Zero"))
    with c3:
        v = saldi['saldo_31_succ']
        st.metric(f"Saldo ultimo {anno_succ}", fmt_eur(v))
    with c4:
        n_anomalie = sum([
            not res['check1']['ok'],
            not res['check2']['ok'],
            not res['check3']['ok'],
            len(res['check4']['unmatched_avere']) + len(res['check4']['unmatched_dare']) > 0,
        ])
        st.metric("Anomalie rilevate", n_anomalie,
                  delta="✓ Tutto ok" if n_anomalie == 0 else f"⚠ {n_anomalie} verifica/e",
                  delta_color="normal" if n_anomalie == 0 else "inverse")

    # ── Tabs ─────────────────────────────────────────────────────────────────
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "① Saldo a zero",
        "② Matching cifre",
        "③ Totali",
        "④ Non abbinati",
        "📋 Mastrino",
    ])

    with tab1:
        st.markdown('<div class="section-title">Check 1 · Il saldo va a zero durante l\'anno di bilancio</div>', unsafe_allow_html=True)

        ok1 = res['check1']['ok']
        va_zero = res['check1']['va_a_zero']
        chiude_avere = res['check1']['chiude_avere']
        saldo_min = res['check1']['saldo_min_bil']
        saldo_fine = res['check1']['saldo_31_bil']

        check_card(
            va_zero,
            f"Saldo a zero nell'anno {anno_bilancio}",
            f"Valore assoluto minimo raggiunto: {fmt_eur(saldo_min)}",
            "Lo storno delle fatture da ricevere ha portato il saldo a zero." if va_zero else
            "ANOMALIA: il saldo non è mai andato a zero. Verificare se gli storni {anno_succ} sono stati registrati."
        )
        check_card(
            chiude_avere,
            f"Saldo chiude in avere al 31/12/{anno_bilancio}",
            f"Saldo finale: {fmt_eur(saldo_fine)}",
            "Le nuove fatture da ricevere hanno riportato il conto in avere." if chiude_avere else
            "ANOMALIA: il saldo non chiude in avere. Verificare le scritture di fine anno."
        )

        if res['check1']['riga_zero'] is not None:
            riga = res['check1']['riga_zero']
            st.markdown(f"""
            <div class="card card-info" style="margin-top:1rem;">
              <div class="card-title">Riga con saldo più vicino a zero</div>
              <div style="font-family:'IBM Plex Mono',monospace;font-size:.85rem;color:#e8eaf0;">
                {riga['data_reg']} &nbsp;|&nbsp; {riga['descrizione']} &nbsp;|&nbsp; Saldo: {fmt_eur(riga['saldo'])}
              </div>
            </div>
            """, unsafe_allow_html=True)

    with tab2:
        st.markdown('<div class="section-title">Check 2 · Matching cifra per cifra: Avere {anno_bilancio} ↔ Dare {anno_succ}</div>'.format(anno_bilancio=anno_bilancio, anno_succ=anno_succ), unsafe_allow_html=True)

        c2ok = res['check2']['ok']
        matched = res['check2']['matched']
        unm_a = res['check2']['unmatched_avere']
        unm_d = res['check2']['unmatched_dare']

        check_card(
            c2ok,
            f"Matching puntuale importi: {len(matched)} coppie abbinate",
            f"Non abbinati in avere {anno_bilancio}: {len(unm_a)} | Non abbinati in dare {anno_succ}: {len(unm_d)}",
        )

        if matched:
            st.markdown(f"**✅ {len(matched)} coppie abbinate:**")
            rows_m = []
            for a, d in matched:
                rows_m.append({
                    f'Data avere ({anno_bilancio})': str(a.data_reg),
                    f'Descrizione avere': str(a.descrizione),
                    'Importo': fmt_eur(float(a.avere)),
                    f'Data dare ({anno_succ})': str(d.data_reg),
                    f'Descrizione dare': str(d.descrizione),
                })
            df_table(pd.DataFrame(rows_m))

        if unm_a or unm_d:
            st.markdown(f"**❌ Non abbinati:**")
            if unm_a:
                st.markdown(f"*Avere {anno_bilancio} senza corrispondente dare {anno_succ}:*")
                df_table(pd.DataFrame([{
                    'Data': str(r.data_reg), 'Descrizione': r.descrizione,
                    'Avere': fmt_eur(float(r.avere))
                } for r in unm_a]))
            if unm_d:
                st.markdown(f"*Dare {anno_succ} senza corrispondente avere {anno_bilancio}:*")
                df_table(pd.DataFrame([{
                    'Data': str(r.data_reg), 'Descrizione': r.descrizione,
                    'Dare': fmt_eur(float(r.dare))
                } for r in unm_d]))

    with tab3:
        st.markdown('<div class="section-title">Check 3 · Totali corrispondenti</div>', unsafe_allow_html=True)

        c3 = res['check3']
        check_card(
            c3['ok'],
            "Totale avere anno bilancio = Totale dare anno successivo",
            f"Somma avere {anno_bilancio}: {fmt_eur(c3['sum_avere'])} | Somma dare {anno_succ}: {fmt_eur(c3['sum_dare'])}",
            f"Differenza: {fmt_eur(c3['diff'])}" if not c3['ok'] else "Totali coincidono ✓"
        )

        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown(f"**Righe avere 31/12/{anno_bilancio}**")
            df_table(res['avere_bil'])
        with col_b:
            st.markdown(f"**Righe dare {anno_succ} (storni)**")
            df_table(res['dare_succ'])

    with tab4:
        st.markdown('<div class="section-title">Check 4 · Importi senza corrispondente</div>', unsafe_allow_html=True)

        unm_a = res['check4']['unmatched_avere']
        unm_d = res['check4']['unmatched_dare']

        if not unm_a and not unm_d:
            st.markdown("""
            <div class="card card-ok">
              <div style="font-size:1.1rem;color:#4ade80;">✅ Tutti gli importi hanno corrispondenza.</div>
            </div>
            """, unsafe_allow_html=True)
        else:
            if unm_a:
                st.markdown(f"### ⚠️ Avere {anno_bilancio} non stornati in {anno_succ}")
                rows = []
                for r in unm_a:
                    rows.append({
                        'Data': str(r.data_reg),
                        'Descrizione': r.descrizione,
                        'Avere': fmt_eur(float(r.avere)),
                        'Segno': 'AVERE',
                        'Anomalia': f'Manca storno in dare {anno_succ}'
                    })
                df_table(pd.DataFrame(rows))

            if unm_d:
                st.markdown(f"### ⚠️ Dare {anno_succ} senza avere corrispondente in {anno_bilancio}")
                rows = []
                for r in unm_d:
                    rows.append({
                        'Data': str(r.data_reg),
                        'Descrizione': r.descrizione,
                        'Dare': fmt_eur(float(r.dare)),
                        'Segno': 'DARE',
                        'Anomalia': f'Manca scrittura avere in {anno_bilancio}'
                    })
                df_table(pd.DataFrame(rows))

    with tab5:
        st.markdown('<div class="section-title">Mastrino completo caricato</div>', unsafe_allow_html=True)

        # Filtro per anno
        anni_disponibili = sorted(set(d.year for d in df['data_reg'] if d is not None))
        sel_anni = st.multiselect("Filtra per anno", anni_disponibili, default=anni_disponibili)

        df_view = df[df['data_reg'].apply(lambda d: d.year if d else 0).isin(sel_anni)].copy()
        df_view['data_reg'] = df_view['data_reg'].astype(str)
        df_view['data_doc'] = df_view['data_doc'].astype(str)
        for col in ['dare', 'avere', 'saldo']:
            if col in df_view.columns:
                df_view[col] = df_view[col].apply(lambda x: fmt_eur(x) if pd.notna(x) else '–')
        st.dataframe(df_view.drop(columns=['conto', 'n_doc'], errors='ignore'), use_container_width=True, height=500)

    # ── Export Excel ─────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown('<div class="section-title">Export</div>', unsafe_allow_html=True)

    col_exp1, col_exp2 = st.columns([1, 3])
    with col_exp1:
        try:
            xls_bytes = build_excel_report(df, res, anno_bilancio, anno_prec, anno_succ)
            st.download_button(
                label="⬇️ Scarica Report Excel",
                data=xls_bytes,
                file_name=f"analisi_44020_{anno_bilancio}.xls",
                mime="application/vnd.ms-excel",
            )
        except Exception as e:
            st.error(f"Errore export Excel: {e}")


if __name__ == "__main__":
    main()
