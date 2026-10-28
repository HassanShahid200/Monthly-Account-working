from html import escape
import gzip
import hashlib
import json
import secrets
import uuid
from urllib.parse import urlsplit, urlunsplit, urlencode
from io import BytesIO
import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode, DataReturnMode

ORDER = [
    'Sales - net', 'Cost of sales', 'Selling and distribution expenses',
    'Administrative expenses', 'Other operating incomes',
    'Other operating expenses', 'Finance Costs', 'Taxation'
]
MAIN, SUB, CODE = 'Current Main Class', 'Current Sub Class', 'Acct Code'
REQUIRED_COLUMNS = {"Series", "Due Date", "Remarks", "Doc. No."}

def normalize_code(series):
    return series.astype('string').str.strip().str.replace(r'\.0$', '', regex=True).replace('', pd.NA)

def parse_ledger_dates(series):
    """Read SAP dates as day/month/year; keep saved Excel/ISO dates intact.

    Explicit formats prevent ambiguous dates such as 01.04.25 being read
    as January 4. Unsupported values remain NaT and are reported by the UI.
    """
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, errors='coerce')
    text = series.astype('string').str.strip()
    result = pd.Series(pd.NaT, index=series.index, dtype='datetime64[ns]')
    for fmt in ('%d.%m.%y', '%d.%m.%Y', '%d/%m/%Y', '%d/%m/%y',
                '%d-%m-%Y', '%d-%m-%y', '%Y-%m-%d',
                '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M:%S.%f',
                '%Y-%m-%dT%H:%M:%S', '%Y-%m-%dT%H:%M:%S.%f'):
        missing = result.isna() & text.notna()
        if not missing.any():
            break
        result.loc[missing] = pd.to_datetime(text.loc[missing], format=fmt, errors='coerce')
    return result


def prepare_report_dates(df, date_col):
    df = df.copy()
    df[date_col] = parse_ledger_dates(df[date_col])
    invalid = int(df[date_col].isna().sum())
    if invalid:
        st.warning(f"{invalid:,} transactions have missing or unsupported dates in "
                   f"{date_col} and will be excluded. Expected SAP format: DD.MM.YY.")
    return df


@st.cache_data
def load_data(content, filename):
    ext = filename.split('.')[-1].lower()
    if ext == 'txt':
        df = pd.read_csv(BytesIO(content), sep="	", encoding="utf-16", low_memory=False)
    elif ext == 'csv':
        df = pd.read_csv(BytesIO(content), low_memory=False)
    else:
        df = pd.read_excel(BytesIO(content))
        
    df.columns = df.columns.astype(str).str.strip()
    missing = sorted(REQUIRED_COLUMNS - set(df.columns))
    
    if missing:
        st.warning("Missing required SAP columns for row-wise cleaning: " + ", ".join(missing))
    else:
        df["Acct Code"] = df["Due Date"].where(df["Series"].isna()).ffill()
        df["Acct Name"] = df["Remarks"].where(df["Series"].isna()).ffill()
        df = df.loc[~df["Series"].isna()].copy()
        
        parts = df["Doc. No."].astype("string").str.split(n=1, expand=True).reindex(columns=[0, 1])
        df["Origin"] = parts[0]
        df["Doc No"] = parts[1]

    df[CODE] = normalize_code(df.get("Acct Code", pd.Series()))
    if df[CODE].isna().any():
        df = df.dropna(subset=[CODE])

    return df

@st.cache_data
def load_mapping(content, filename):
    df = pd.read_excel(BytesIO(content))
    df.columns = df.columns.astype(str).str.strip()
    
    if 'Account code' not in df.columns:
        if 'Acct Code' in df.columns:
            df = df.rename(columns={'Acct Code': 'Account code'})
        else:
            st.error("Mapping file must contain an 'Account code' column.")
            return pd.DataFrame()
            
    df['Account code'] = normalize_code(df['Account code'])
    
    if MAIN not in df.columns: df[MAIN] = 'Unclassified'
    if SUB not in df.columns: df[SUB] = 'Unclassified'
    if 'Title' not in df.columns: df['Title'] = df['Account code']
        
    df = df[['Account code', MAIN, SUB, 'Title']].dropna(subset=['Account code']).drop_duplicates()
    duplicates = df.loc[df['Account code'].duplicated(keep=False), 'Account code'].unique()
    if len(duplicates):
        st.error('Conflicting mapping rows for account codes: ' + ', '.join(duplicates.astype(str)))
        return pd.DataFrame()
    return df

def filter_by_date(df, date_col, start_date, end_date):
    mask = df[date_col].notna()
    if start_date:
        mask &= (df[date_col].dt.date >= pd.to_datetime(start_date).date())
    if end_date:
        mask &= (df[date_col].dt.date <= pd.to_datetime(end_date).date())
    return df[mask]

def percentage_change(col1, col2, formula_choice):
    """Column 1 is latest/first; Column 2 is earlier/second for new reports.

    Explicit older formulas remain supported for saved reports.
    """
    if formula_choice == "(Col 1 - Col 2) / Col 2":
        return (col1 - col2) / col2 if col2 else None
    if formula_choice == "(Col 1 - Col 2) / Col 1":
        return (col1 - col2) / col1 if col1 else None
    if formula_choice == "(Col 2 - Col 1) / Col 2":
        return (col2 - col1) / col2 if col2 else None
    return (col2 - col1) / col1 if col1 else None


def visible_rows(summary, expanded, col1_name, col2_name, formula_choice):
    rows, paths = [], []
    def append_row(part, path, label):
        old, new = float(part[col1_name].sum()), float(part[col2_name].sum())
        
        change = percentage_change(old, new, formula_choice)

        if len(path) < 3:
            count = part[CODE].nunique()
            label = f'{label} ({count} {"account" if count == 1 else "accounts"})'
        
        rows.append({'Hierarchy': label,
                     col1_name: old, col2_name: new, '% Change': change,
                     '_level': len(path), '_path': json.dumps(path, ensure_ascii=False, separators=(',', ':')),
                     '_expanded': '', '_action': '', '_id': len(paths)})
        paths.append(path)

    main_values = sorted(summary[MAIN].unique(), key=lambda x: (ORDER.index(x) if x in ORDER else len(ORDER), x))
    for main in main_values:
        part = summary.loc[summary[MAIN].eq(main)]
        append_row(part, (main,), main)
        if (main,) not in expanded:
            continue
        for sub, subgroup in part.groupby(SUB, sort=True):
            append_row(subgroup, (main, sub), sub)
            if (main, sub) in expanded:
                for _, account in subgroup.sort_values(['Title', CODE]).iterrows():
                    append_row(account.to_frame().T, (main, sub, account[CODE]),
                               f"{account['Title']} [{account[CODE]}]")
    return pd.DataFrame(rows), paths

def accounting_amount(value, decimals=0):
    if value is None or pd.isna(value):
        return '—'
    value = round(float(value), decimals)
    number = f"{abs(value):,.{decimals}f}"
    return f"({number})" if value < 0 else number


def pnl_statement_rows(summary, col1_name, col2_name):
    totals = summary.groupby(MAIN)[[col1_name, col2_name]].sum()
    def values(names):
        return tuple(float(totals.reindex(names).fillna(0)[c].sum())
                     for c in (col1_name, col2_name))
    rows = []
    def add(label, names, style='normal'):
        rows.append((label, *values(names), style))
    add('Sales - net', [ORDER[0]])
    add('Cost of sales', [ORDER[1]])
    add('Gross profit', ORDER[:2], 'subtotal')
    for main in ORDER[2:7]:
        add('Finance costs' if main == 'Finance Costs' else main, [main])
    # Keep additional mapped classes visible and included in the reconciliation.
    extra = [name for name in totals.index if name not in ORDER]
    for main in extra:
        add(str(main), [main])
    add('Profit before taxation', ORDER[:7] + extra, 'subtotal')
    add('Taxation', ['Taxation'])
    add('Profit/(Loss) for the period', ORDER + extra, 'final')
    return rows


def render_pnl_statement(summary, col1_config, col2_config, formula_choice="(Col 1 - Col 2) / Col 2"):
    def heading(config):
        return escape(str(config.get('name') or ''))
    def percent(first, second):
        value = percentage_change(first, second, formula_choice)
        return 'N/A' if value is None else f'{value:.2%}'
    body = []
    for label, first, second, style in pnl_statement_rows(
            summary, col1_config['name'], col2_config['name']):
        body.append(f'<tr class="{style}"><td>{escape(label)}</td>'
                    f'<td>{accounting_amount(first)}</td><td>{accounting_amount(second)}</td>'
                    f'<td>{percent(first, second)}</td></tr>')
    body.append('<tr class="eps"><td>Earnings per share - basic and diluted</td>'
                f'<td>{accounting_amount(col1_config.get("eps"), 2)}</td>'
                f'<td>{accounting_amount(col2_config.get("eps"), 2)}</td>'
                f'<td>{percent(col1_config["eps"], col2_config["eps"]) if col1_config.get("eps") is not None and col2_config.get("eps") is not None else "N/A"}</td></tr>')
    st.markdown('''
<style>
.pnl-wrap {overflow-x:auto; background:white; color:#111; padding:24px 16px;}
.pnl {width:100%; border-collapse:separate; border-spacing:12px 0; font-family:Arial,sans-serif; font-size:14px;}
.pnl th {text-align:center; padding:5px 8px 26px; vertical-align:top;}
.pnl td {padding:5px 8px;}
.pnl td:first-child {width:43%; text-align:left;}
.pnl td:not(:first-child) {width:19%; text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap;}
.pnl .currency th {padding:0 8px 32px; font-size:13px;}
.pnl .subtotal td {font-weight:bold; padding-top:10px; padding-bottom:24px;}
.pnl .subtotal td:not(:first-child) {border-top:1px solid #111;}
.pnl .final td {font-weight:bold; padding-top:10px; padding-bottom:5px;}
.pnl .final td:not(:first-child) {border-top:1px solid #111; border-bottom:3px double #111;}
.pnl .eps td {padding-top:28px; font-weight:bold;}
.pnl .eps td:not(:first-child) {border-bottom:3px double #111;}
</style>
''' + '<div class="pnl-wrap"><table class="pnl"><thead><tr><th></th><th>'
        + heading(col1_config) + '</th><th>' + heading(col2_config)
        + '</th><th>% Change</th></tr><tr class="currency"><th></th><th colspan="2">(Rupees)</th><th></th></tr></thead><tbody>'
        + ''.join(body) + '</tbody></table></div>', unsafe_allow_html=True)


def cloud_client():
    from supabase import create_client
    url = str(st.secrets.get('SUPABASE_URL', '')).rstrip('/')
    key = str(st.secrets.get('SUPABASE_SECRET_KEY', ''))
    if not url or not key:
        return None
    return create_client(url, key)

def token_digest(token):
    return hashlib.sha256(token.encode('utf-8')).hexdigest()

def fetch_shared_report(token):
    client = cloud_client()
    if not client: raise ValueError("Database connection failed. Check SUPABASE_URL and SUPABASE_SECRET_KEY.")
    records = client.table('ledger_reports').select('*').eq('share_token_hash', token_digest(token)).limit(1).execute().data
    if not records:
        raise ValueError('This report link is invalid or has expired.')
    report = records[0]
    bucket = str(st.secrets.get('SUPABASE_BUCKET', 'Ledger_Reports'))
    content = client.storage.from_(bucket).download(report['file_path'])
    return report, content

@st.cache_data(max_entries=2, show_spinner=False)
def serialize_snapshot(df):
    # Excel cell-by-cell export is expensive for large ledgers. CSV + gzip
    # retains all columns and full amount precision, with no new dependency.
    return gzip.compress(df.to_csv(index=False).encode('utf-8'), compresslevel=1, mtime=0)


def read_report_snapshot(content, filename):
    if filename.lower().endswith('.csv.gz') or content[:2] == b'\x1f\x8b':
        # Preserve text identifiers, including leading zeroes.
        identifiers = {c: 'string' for c in
                       (CODE, 'Account code', 'Doc No', 'Doc. No.', 'Origin',
                        'Title', 'Acct Name', MAIN, SUB, 'Offset Acct')}
        return pd.read_csv(BytesIO(content), compression='gzip', dtype=identifiers,
                           float_precision='round_trip', low_memory=False)
    # Existing cloud reports remain readable.
    return pd.read_excel(BytesIO(content))


def save_report_snapshot(df, title, comments, col1_config, col2_config, sel_date, sel_amt,
                         formula_choice, progress=None):
    def update(message):
        if progress:
            progress(message)
    update('1/3 · Preparing and compressing ledger...')
    content = serialize_snapshot(df)
    client = cloud_client()
    if not client:
        raise ValueError("Supabase credentials not found in secrets.")
    bucket = str(st.secrets.get('SUPABASE_BUCKET', 'Ledger_Reports'))
    report_id = str(uuid.uuid4())
    path = report_id + '/ledger.csv.gz'
    token = secrets.token_urlsafe(32)
    payload = {
        'id': report_id, 'title': title.strip() or 'Financial Report',
        'file_path': path, 'original_filename': 'mapped_ledger.csv.gz',
        'comments': dict(comments),
        'reporting_period': json.dumps({
            'col1': col1_config, 'col2': col2_config,
            'sel_date': sel_date, 'sel_amt': sel_amt, 'formula': formula_choice
        }),
        'share_token_hash': token_digest(token)
    }
    update(f'2/3 · Uploading compressed ledger ({len(content) / 1024 / 1024:.2f} MB)...')
    try:
        client.storage.from_(bucket).upload(
            path=path, file=content, file_options={'content-type': 'application/gzip'})
    except Exception as exc:
        raise RuntimeError(f'Ledger upload failed: {exc}') from exc
    update('3/3 · Saving report settings and comments...')
    try:
        client.table('ledger_reports').insert(payload).execute()
    except Exception as exc:
        # A timeout may happen after the server committed the insert.
        # Confirm absence before removing the file, so a saved report stays valid.
        try:
            records = client.table('ledger_reports').select('id').eq('id', report_id).limit(1).execute().data
        except Exception:
            raise RuntimeError('Could not confirm report save. Check Open Saved Report before retrying. '
                               f'Details: {exc}') from exc
        if not records:
            try:
                client.storage.from_(bucket).remove([path])
            except Exception:
                pass
            raise RuntimeError(f'Report settings could not be saved: {exc}') from exc
    update('Report saved successfully.')
    return token

def fetch_recent_reports():
    client = cloud_client()
    if not client: return []
    res = client.table('ledger_reports').select('id, title, created_at').order('created_at', desc=True).limit(50).execute()
    return res.data

def load_saved_report(report_id):
    client = cloud_client()
    bucket = str(st.secrets.get('SUPABASE_BUCKET', 'Ledger_Reports'))
    report = client.table('ledger_reports').select('*').eq('id', report_id).limit(1).execute().data[0]
    content = client.storage.from_(bucket).download(report['file_path'])
    return report, content

def report_link(base_url, token):
    parsed = urlsplit(base_url.strip())
    if parsed.scheme != 'https' or not parsed.netloc:
        raise ValueError("Invalid APP_URL. Make sure it starts with https://")
    return urlunsplit((parsed.scheme, parsed.netloc, parsed.path or '/', urlencode({'report': token}), ''))

def run():
    st.set_page_config(layout='wide', page_title='Financial Dashboard')
    
    # ========================================================
    # GITHUB ICON AUR STREAMLIT MENU HIDE KARNE KA CODE
    # ========================================================
    hide_streamlit_style = """
    <style>
    /* GitHub icon, Main Menu aur Footer hide karne ke liye */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .viewerBadge_container__1QSob {display: none !important;}
    .viewerBadge_link__1S137 {display: none !important;}
    .viewerBadge_text__1JaDK {display: none !important;}
    
    /* Naye Streamlit versions ke liye */
    [data-testid="stToolbar"] {visibility: hidden;}
    [data-testid="stDecoration"] {display: none;}
    [data-testid="stStatusWidget"] {visibility: hidden;}
    </style>
    """
    st.markdown(hide_streamlit_style, unsafe_allow_html=True)
    # ========================================================
    
    for key, default in [('expanded', set()), ('draft_comments', {}), ('ledger_path', None), ('ledger_col', None), ('handled_action', None)]:
        st.session_state.setdefault(key, default)

    # ========================================================
    # CHECK FOR SHARED REPORT LINK (READ ONLY MODE)
    # ========================================================
    report_token = st.query_params.get('report', '')
    read_only_mode = False
    
    if report_token:
        try:
            report_meta, content = fetch_shared_report(report_token)
            st.title('SAP Financial Ledger Dashboard')
            st.subheader(report_meta['title'])
            st.caption('Shared report - View only mode (Comments cannot be edited)')
            
            df = read_report_snapshot(content, report_meta.get('original_filename', 'workbook.xlsx'))
            
            st.session_state.draft_comments = report_meta.get('comments', {})
            configs = json.loads(report_meta['reporting_period'])
            
            col1_config = configs.get('col1', {'name': '', 'start': None, 'end': None})
            col2_config = configs.get('col2', {'name': '', 'start': None, 'end': None})
            sel_date = configs.get('sel_date', 'Posting Date')
            sel_amt = configs.get('sel_amt', 'Deb./Cred. (LC)')
            formula_choice = configs.get('formula', '(Col 1 - Col 2) / Col 2')
            
            if sel_date in df.columns:
                df = prepare_report_dates(df, sel_date)
                
            mode = "Shared View"
            read_only_mode = True
        except Exception as e:
            st.error(f"Could not load shared report: {e}")
            return
    else:
        st.title('P&L Comparison Report')
        # ========================================================
        # PASSWORD PROTECTION FOR EDITORS
        # ========================================================
        expected_password = str(st.secrets.get('EDITOR_PASSWORD', ''))
        if expected_password:
            if not st.session_state.get('authenticated'):
                st.info("This dashboard is protected. Please log in.")
                password = st.text_input("Enter Editor Password", type="password")
                if st.button("Login"):
                    if password == expected_password:
                        st.session_state.authenticated = True
                        st.rerun()
                    else:
                        st.error("Incorrect password!")
                return
                
        if st.sidebar.button("Logout"):
            st.session_state.authenticated = False
            st.rerun()
            
        mode = st.radio("Choose Action", ["Create New Report (Upload Data)", "Open Saved Report (Cloud)"], horizontal=True)
        
        df = pd.DataFrame()
        col1_config = {'name': '', 'start': None, 'end': None}
        col2_config = {'name': '', 'start': None, 'end': None}
        sel_date = 'Posting Date'
        sel_amt = 'Deb./Cred. (LC)'
        formula_choice = '(Col 1 - Col 2) / Col 2'
        
        if mode == "Create New Report (Upload Data)":
            st.header("1. Upload Data")
            col1, col2 = st.columns(2)
            with col1:
                upload_ledger = st.file_uploader('Upload Raw SAP Export (TXT, CSV, XLSX)', type=['txt', 'csv', 'xlsx'])
            with col2:
                upload_mapping = st.file_uploader('Upload Mapping File (XLSX)', type=['xlsx'])
                
            if not upload_ledger or not upload_mapping:
                st.info('Please upload both files to proceed.')
                return
                
            raw_df = load_data(upload_ledger.getvalue(), upload_ledger.name)
            mapping_df = load_mapping(upload_mapping.getvalue(), upload_mapping.name)
            
            if raw_df.empty or mapping_df.empty:
                st.warning('Could not load data. Ensure files are not empty.')
                return
                
            st.header("2. Column Settings")
            date_options = [c for c in raw_df.columns if 'date' in c.lower()]
            amt_options = [c for c in raw_df.columns if 'LC' in c or 'FC' in c or pd.api.types.is_numeric_dtype(raw_df[c])]
            
            c_date, c_amt = st.columns(2)
            sel_date = c_date.selectbox("Select Date Column", raw_df.columns, index=raw_df.columns.get_loc(date_options[0]) if date_options else 0)
            sel_amt = c_amt.selectbox("Select Amount Column", raw_df.columns, index=raw_df.columns.get_loc(amt_options[0]) if amt_options else 0)
                
            df = pd.merge(raw_df, mapping_df, left_on=CODE, right_on='Account code', how='inner', validate='many_to_one')
            if df.empty:
                st.error('No matching accounts found between the raw ledger and your mapping file.')
                return
            
            df = prepare_report_dates(df, sel_date)
            if df[sel_amt].dtype == 'object':
                df[sel_amt] = pd.to_numeric(df[sel_amt].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            else:
                df[sel_amt] = pd.to_numeric(df[sel_amt], errors='coerce').fillna(0)
            
            st.success(f"Matched {df[CODE].nunique()} unique accounts. Total transactions: {len(df)}")
            
            st.header("3. Calculation Settings")
            c1, c2, c3 = st.columns(3)
            with c1:
                st.subheader("Column 1 Setup — Latest Period (First)")
                c1_name = st.text_input("Column 1 Name", value="")
                c1_start = st.date_input("Start Date (Col 1)", value=None, key='c1s')
                c1_end = st.date_input("End Date (Col 1)", value=None, key='c1e')
            with c2:
                st.subheader("Column 2 Setup — Earlier Period (Second)")
                c2_name = st.text_input("Column 2 Name", value="")
                c2_start = st.date_input("Start Date (Col 2)", value=None, key='c2s')
                c2_end = st.date_input("End Date (Col 2)", value=None, key='c2e')
            with c3:
                st.subheader("Formula Setup")
                formula_choice = st.selectbox("Select % Change Formula", [
                    "(Col 1 - Col 2) / Col 2",
                    "(Col 1 - Col 2) / Col 1",
                    "(Col 2 - Col 1) / Col 1",
                    "(Col 2 - Col 1) / Col 2"
                ])
                
            col1_config = {'name': c1_name, 'start': c1_start, 'end': c1_end}
            col2_config = {'name': c2_name, 'start': c2_start, 'end': c2_end}

            st.info(f"Selected Formula: **{formula_choice}**")
            st.caption("SAP dates use DD.MM.YY. For the screenshot comparison, select Posting Date, "
                       "Deb./Cred. (LC), Column 1 = 2026, Column 2 = 2025, "
                       "and (Col 1 - Col 2) / Col 2.")
            if not st.button("Generate Report", type='primary') and not st.session_state.get('report_generated'):
                return
            st.session_state.report_generated = True

        elif mode == "Open Saved Report (Cloud)":
            st.header("1. Open Saved Report")
            reports = fetch_recent_reports()
            if not reports:
                st.info("No saved reports found in Supabase.")
                return
                
            report_choices = {r['id']: f"{r['title']} (Saved on {r['created_at'][:10]})" for r in reports}
            selected_id = st.selectbox("Select a report to open", list(report_choices.keys()), format_func=lambda x: report_choices[x])
            
            if st.button("Load Selected Report", type="primary"):
                with st.spinner("Downloading report from cloud..."):
                    report_meta, content = load_saved_report(selected_id)
                    st.session_state.loaded_df = read_report_snapshot(content, report_meta.get('original_filename', 'workbook.xlsx'))
                    
                    st.session_state.draft_comments = report_meta.get('comments', {})
                    if report_meta.get('reporting_period'):
                        try:
                            configs = json.loads(report_meta['reporting_period'])
                            st.session_state.loaded_col1_config = configs.get('col1')
                            st.session_state.loaded_col2_config = configs.get('col2')
                            st.session_state.loaded_sel_date = configs.get('sel_date', 'Posting Date')
                            st.session_state.loaded_sel_amt = configs.get('sel_amt', 'Deb./Cred. (LC)')
                            st.session_state.loaded_formula = configs.get('formula', '(Col 1 - Col 2) / Col 2')
                        except:
                            pass
                            
            if 'loaded_df' in st.session_state:
                df = st.session_state.loaded_df
                sel_date = st.session_state.get('loaded_sel_date', sel_date)
                sel_amt = st.session_state.get('loaded_sel_amt', sel_amt)
                
                if sel_date in df.columns:
                    df = prepare_report_dates(df, sel_date)
                st.success("Report loaded successfully!")
                
                col1_config = st.session_state.get('loaded_col1_config', col1_config)
                col2_config = st.session_state.get('loaded_col2_config', col2_config)
                formula_choice = st.session_state.get('loaded_formula', formula_choice)
            else:
                return

    # ========================================================
    # DASHBOARD LOGIC (Runs for New, Saved, and Shared Reports)
    # ========================================================
    col1_name = col1_config['name']
    col2_name = col2_config['name']
    
    reserved = {MAIN, SUB, CODE, 'Title', 'Hierarchy', '% Change', 'Comments',
                '_level', '_path', '_expanded', '_action', '_id'}
    if (not col1_name.strip() or not col2_name.strip() or col1_name == col2_name
            or col1_name in reserved or col2_name in reserved):
        st.error('Choose two different, non-empty period names that are not ledger column names.')
        return
    for config in (col1_config, col2_config):
        if config['start'] and config['end'] and pd.Timestamp(config['start']) > pd.Timestamp(config['end']):
            st.error(f"Start date must be on or before end date for {config['name']}.")
            return

    df_col1 = filter_by_date(df, sel_date, col1_config['start'], col1_config['end'])
    df_col2 = filter_by_date(df, sel_date, col2_config['start'], col2_config['end'])
    
    # We multiply by -1 ONLY for the P&L summary totals, exactly as requested
    col1_totals = df_col1.groupby(CODE)[sel_amt].sum() * -1
    col2_totals = df_col2.groupby(CODE)[sel_amt].sum() * -1
    
    summary = df[[MAIN, SUB, CODE, 'Title']].drop_duplicates()
    summary[col1_name] = summary[CODE].map(col1_totals).fillna(0)
    summary[col2_name] = summary[CODE].map(col2_totals).fillna(0)
    
    st.divider()
    if not read_only_mode:
        st.subheader('3. Profit and Loss Summary')
    
    if not read_only_mode:
        with st.expander('Earnings per share (optional)'):
            eps1, eps2 = st.columns(2)
            col1_config['eps'] = eps1.number_input(
                f"EPS · {col1_name}", value=col1_config.get('eps'), step=0.01, format='%.2f')
            col2_config['eps'] = eps2.number_input(
                f"EPS · {col2_name}", value=col2_config.get('eps'), step=0.01, format='%.2f')
            st.caption('Enter the approved EPS for each period. Blank values display as a dash.')
    pnl_tab, level_tab = st.tabs(['P&L Format', 'Level-wise Report'])
    with pnl_tab:
        render_pnl_statement(summary, col1_config, col2_config, formula_choice)
        st.caption('Rupees rounded to whole Rupees. Negative amounts appear in parentheses. '
                   'Subtotals use underlying unrounded ledger values.')
    with level_tab:
        st.caption('Level 1: main category (dark blue) · Level 2: subclass (light blue) · Level 3: account (white). ' 
                   'Use + / − to expand or collapse; click a name or amount for the ledger.')

        all_branches = {(m,) for m in summary[MAIN].unique()}
        all_branches.update(tuple(r) for r in summary[[MAIN, SUB]].drop_duplicates().itertuples(index=False, name=None))
        table, paths = visible_rows(summary, all_branches, col1_name, col2_name, formula_choice)

        comments = st.session_state.draft_comments
        table['Comments'] = table['_path'].map(comments).fillna('')

        gb = GridOptionsBuilder.from_dataframe(table)
        gb.configure_default_column(sortable=False, filter=False, resizable=True, editable=False)
        for col in ['_id', '_level', '_action', '_path', '_expanded']:
            gb.configure_column(col, hide=True)
    
        gb.configure_column('Hierarchy', header_name='Main Class / Sub Class / Account', minWidth=450,
                            cellStyle={'whiteSpace': 'pre'},
                            headerComponent=JsCode('''class LevelHeader {
                init(p) {
                    this.el = document.createElement('div');
                    this.el.style.cssText = 'width:100%;display:flex;flex-direction:column;gap:8px;';
                    const title = document.createElement('span');
                    title.textContent = p.displayName;
                    this.el.appendChild(title);
                    const controls = document.createElement('div');
                    controls.style.cssText = 'display:flex;gap:6px;';
                    ['Level 1', 'Level 2', 'Level 3'].forEach((label, index) => {
                        const button = document.createElement('button');
                        button.type = 'button';
                        button.textContent = label;
                        button.style.cssText = 'cursor:pointer;padding:4px 12px;border:1px solid #a7bbce;border-radius:5px;background:#ffffff;color:#17324d;font-size:13px;font-weight:bold;';
                        button.addEventListener('click', (event) => {
                            event.stopPropagation();
                            const opened = [];
                            p.api.forEachNode(node => {
                                if (node.data && node.data._level <= index) opened.push(node.data._path);
                            });
                            window.__ledgerExpanded = opened;
                            p.api.onFilterChanged();
                            p.api.refreshCells({columns: ['Hierarchy'], force: true});
                        });
                        controls.appendChild(button);
                    });
                    this.el.appendChild(controls);
                }
                getGui() { return this.el; }
            }'''),
                            cellRenderer=JsCode('''class HierarchyRenderer {
                init(p) {
                    this.el = document.createElement('div');
                    if (p.node.rowPinned) {
                        this.el.textContent = p.value;
                        return;
                    }
                    this.el.style.display = 'flex';
                    this.el.style.alignItems = 'center';
                    this.el.style.paddingLeft = ((p.data._level - 1) * 22) + 'px';
                    const label = String(p.value || '').trim();
                    if (p.data._level < 3) {
                        const button = document.createElement('button');
                        button.type = 'button';
                        const isOpen = (window.__ledgerExpanded || []).includes(p.data._path);
                        button.textContent = isOpen ? '−' : '+';
                
                        const btnColor = p.data._level === 1 ? '#ffffff' : '#17324d';
                        button.style.cssText = 'cursor:pointer;width:24px;height:24px;margin-right:9px;border:1.5px solid ' + btnColor + ';border-radius:4px;background:transparent;color:' + btnColor + ';font-weight:bold;font-size:16px;line-height:1;display:flex;align-items:center;justify-content:center;';
                
                        button.addEventListener('click', function(e) {
                            e.stopPropagation();
                            const key = p.data._path;
                            const opened = window.__ledgerExpanded || [];
                            if (opened.includes(key)) {
                                const prefix = JSON.parse(key);
                                window.__ledgerExpanded = opened.filter(function(item) {
                                    const path = JSON.parse(item);
                                    return !prefix.every((value, i) => path[i] === value);
                                });
                            } else {
                                window.__ledgerExpanded = opened.concat([key]);
                            }
                            p.api.onFilterChanged();
                            p.api.refreshCells({columns: ['Hierarchy'], force: true});
                        });
                        this.el.appendChild(button);
                    }
                    const text = document.createElement('span');
                    text.textContent = label;
                    this.el.appendChild(text);
                }
                getGui() { return this.el; }
            }'''))
    
        money = JsCode("function(p) { if (p.value == null) return ''; const v = Number(p.value); const n = Math.abs(v).toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2}); return v < 0 ? '(' + n + ')' : n; }")
        percent = JsCode("function(p) { return p.value == null ? 'N/A' : Number(p.value).toLocaleString('en-US', {style:'percent', minimumFractionDigits:2, maximumFractionDigits:2}); }")

        gb.configure_column(col2_name, type=['numericColumn'], minWidth=150, valueFormatter=money)
        gb.configure_column(col1_name, type=['numericColumn'], minWidth=150, valueFormatter=money)
        gb.configure_column('% Change', type=['numericColumn'], valueFormatter=percent)

        gb.configure_column('Comments', minWidth=280, width=360, wrapText=True, autoHeight=True,
                            editable=JsCode('function(p) { return !p.node.rowPinned; }') if not read_only_mode else False,
                            cellEditor='agLargeTextCellEditor', cellEditorPopup=True,
                            cellEditorParams={'maxLength': 4000, 'rows': 5, 'cols': 45})
                    
        total_new, total_old = float(summary[col2_name].sum()), float(summary[col1_name].sum())

        total_change = percentage_change(total_old, total_new, formula_choice)

        gb.configure_grid_options(
            headerHeight=78,
            stopEditingWhenCellsLoseFocus=True,
            getRowId=JsCode("function(p) { return p.data._path; }"),
            pinnedBottomRowData=[{'Hierarchy': 'Grand Total', col1_name: total_old,
                                  col2_name: total_new, '% Change': total_change}],
            context={'expanded': [json.dumps(path, ensure_ascii=False, separators=(',', ':'))
                                  for path in sorted(st.session_state.expanded)]},
            isExternalFilterPresent=JsCode('function() { return true; }'),
            doesExternalFilterPass=JsCode('''function(p) {
                const path = JSON.parse(p.data._path);
                for (let depth = 1; depth < path.length; depth++) {
                    if (!(window.__ledgerExpanded || []).includes(JSON.stringify(path.slice(0, depth)))) return false;
                }
                return true;
            }'''),
            onGridReady=JsCode('function(p) { window.__ledgerExpanded = ((p.context && p.context.expanded) || []).slice(); p.api.onFilterChanged(); }'),
            onFirstDataRendered=JsCode('function(p) { p.api.onFilterChanged(); }'),
            onCellClicked=JsCode('''function(p) {
                if (!p.data || p.node.rowPinned || ['_action', 'Comments', '% Change'].includes(p.colDef.field)) return;
                p.api.forEachNode(function(node) { node.data._action = ''; });
                p.data._expanded = JSON.stringify(window.__ledgerExpanded || []);
                p.data._action = 'ledger:' + p.colDef.field + ':' + Date.now();
            }'''),
            onCellValueChanged=JsCode('''function(p) {
                if (p.colDef.field === 'Comments') {
                    p.api.forEachNode(function(node) { node.data._action = ''; });
                    p.data._expanded = JSON.stringify(window.__ledgerExpanded || []);
                }
            }'''),
            getRowStyle=JsCode('''function(p) {
                if (!p.data) return {};
                if (p.node.rowPinned) return { backgroundColor: '#17324d', color: '#ffffff', fontWeight: '700', borderTop: '2px solid #a7bbce'};
                if (p.data._level === 1) return { backgroundColor: '#17324d', color: '#ffffff', fontWeight: '700', borderTop: '2px solid #ffffff'};
                if (p.data._level === 2) return { backgroundColor: '#e7eff7', color: '#17324d', fontWeight: '600', borderBottom: '1px solid #cedcea'};
                return { backgroundColor: '#ffffff', color: '#374151', borderBottom: '1px solid #eef2f6'};
            }''')
        )

        st.caption('Click a name or amount to open its ledger below.')
        response = AgGrid(table, gridOptions=gb.build(), height=500, theme='streamlit',
                          allow_unsafe_jscode=True, update_on=['cellClicked', 'cellValueChanged'],
                          data_return_mode=DataReturnMode.AS_INPUT,
                          should_grid_return=JsCode('''function(p) {
                              const event = p.eventData;
                              if (!event || !event.data || (event.node && event.node.rowPinned)) return false;
                              const field = event.colDef && event.colDef.field;
                              if (p.streamlitRerunEventTriggerName === 'cellClicked') {
                                  return ['Hierarchy', "''' + col2_name + '''", "''' + col1_name + '''"].includes(field);
                              }
                              return p.streamlitRerunEventTriggerName === 'cellValueChanged' && field === 'Comments';
                          }'''))

        returned = response.data

        if not read_only_mode:
            if isinstance(returned, pd.DataFrame) and {'_path', 'Comments'}.issubset(returned.columns):
                valid = set(table['_path'])
                for _, row in returned.iterrows():
                    path = row['_path']
                    if path in valid:
                        val = row['Comments']
                        text = '' if pd.isna(val) else str(val)[:4000]
                        if text:
                            st.session_state.draft_comments[path] = text
                        else:
                            st.session_state.draft_comments.pop(path, None)

        # Handle Action / Drill-down
        if isinstance(returned, pd.DataFrame) and '_action' in returned.columns:
            actions = returned.loc[returned['_action'].fillna('').ne('')]
            if not actions.empty:
                action_row = actions.iloc[-1]
                action_path = paths[int(action_row['_id'])]
                action = str(action_row['_action'])
        
                token = (action_row['_id'], action)
                if st.session_state.get('handled_action') != token:
                    st.session_state.handled_action = token
                    st.session_state.ledger_path = action_path
                    opened = action_row.get('_expanded', '')
                    if opened:
                        st.session_state.expanded = {tuple(json.loads(item)) for item in json.loads(opened)}
            
                    clicked_column = action.split(':')[1]
                    if clicked_column in [col1_name, col2_name]:
                        st.session_state.ledger_col = clicked_column
                    else:
                        st.session_state.ledger_col = col1_name


    # ========================================================
    # LEDGER DRILL DOWN SECTION
    # ========================================================
    path = st.session_state.ledger_path
    if path is not None:
        st.divider()
        st.subheader('Ledger Details')
        if st.button('Close ledger details', key='close_pnl_ledger'):
            st.session_state.ledger_path = None
            st.rerun()
    
        st.write(f"**Selected scope:** {' → '.join(path)}")
    
        if st.session_state.ledger_col not in [col1_name, col2_name]:
            st.session_state.ledger_col = col1_name
        
        selected_col = st.radio('Select the ledger period to view:', [col1_name, col2_name], horizontal=True, key='ledger_col')
    
        if selected_col == col1_name:
            detail = filter_by_date(df, sel_date, col1_config['start'], col1_config['end'])
            period_label = col1_name
        else:
            detail = filter_by_date(df, sel_date, col2_config['start'], col2_config['end'])
            period_label = col2_name
    
        mask = detail[MAIN].eq(path[0])
        if len(path) >= 2: mask &= detail[SUB].eq(path[1])
        if len(path) == 3: mask &= detail[CODE].eq(path[2])
    
        detail = detail.loc[mask].copy()
    
        if sel_date in detail.columns:
            detail = detail.sort_values(sel_date)
    
        a, b = st.columns(2)
        a.metric(f'Transactions · {period_label}', len(detail))
        b.metric('Original Ledger Net Amount', f'{detail[sel_amt].sum():,.2f}')
    
        # Keep unmodified original amounts as requested. Plus is Debit, Minus is Credit.
        detail['Debit'] = detail[sel_amt].clip(lower=0)
        detail['Credit'] = -detail[sel_amt].clip(upper=0)
    
        st.caption('Net = Debit − Credit.')
    
        columns = [CODE, 'Acct Name', MAIN, SUB, 'Title', sel_date, 'Origin', 'Doc No',
                   'Remarks', 'Offset Acct', 'Offset Acct Name', 'Debit', 'Credit', sel_amt]
        detail = detail[[col for col in columns if col in detail.columns]].copy()
    
        if sel_date in detail.columns:
            detail[sel_date] = detail[sel_date].dt.strftime('%Y-%m-%d')
        
        if detail.empty:
            st.info('No transactions were found for this selection and period.')
        else:
            total_row = {col: '' for col in detail.columns}
            total_row[CODE] = 'Grand Total'
            for col in ['Debit', 'Credit', sel_amt]:
                total_row[col] = float(detail[col].sum())
            
            detail = pd.concat([detail, pd.DataFrame([total_row])], ignore_index=True)
            st.dataframe(detail, use_container_width=True, hide_index=True,
                         column_config={col: st.column_config.NumberColumn(format='%.2f')
                                        for col in ['Debit', 'Credit', sel_amt]})
        
            st.download_button('Download Selected Ledger (CSV)', detail.to_csv(index=False).encode('utf-8-sig'),
                               file_name=f'ledger_{period_label}.csv', mime='text/csv')

    # ========================================================
    # SAVE TO CLOUD (ONLY FOR EDITORS)
    # ========================================================
       # ========================================================
    # SAVE TO CLOUD (ONLY FOR EDITORS)
    # ========================================================
    if not read_only_mode:
        st.divider()
        st.subheader("5. Save to Cloud")
        title = st.text_input('Report title', value=f'{col1_name} vs {col2_name} Financial Report')
        
        default_url = str(st.secrets.get('APP_URL', ''))
        base_url_input = st.text_input('Verify your App URL (Must start with https://):', value=default_url)
        
        if st.button('Save & Share Report', type='primary'):
            try:
                if not base_url_input or urlsplit(base_url_input.strip()).scheme != "https" or not urlsplit(base_url_input.strip()).netloc:
                    st.error("Please enter a valid APP_URL starting with https://")
                else:
                    with st.status('Saving report...', expanded=True) as save_status:
                        share_token = save_report_snapshot(
                            df, title, st.session_state.draft_comments, 
                            {'name': col1_name, 'start': str(col1_config['start']) if col1_config['start'] else None, 'end': str(col1_config['end']) if col1_config['end'] else None, 'eps': col1_config.get('eps')}, 
                            {'name': col2_name, 'start': str(col2_config['start']) if col2_config['start'] else None, 'end': str(col2_config['end']) if col2_config['end'] else None, 'eps': col2_config.get('eps')},
                            sel_date, sel_amt, formula_choice, progress=save_status.write
                        )
                        save_status.update(label='Report saved successfully', state='complete', expanded=False)
                    st.session_state.last_saved_report_link = report_link(base_url_input, share_token)
                    st.success('Report saved successfully!')
            except Exception as exc:
                st.error(f'Save failed: {exc}')
        if st.session_state.get('last_saved_report_link'):
            st.caption('Last successfully saved report:')
            st.code(st.session_state.last_saved_report_link)

if __name__ == '__main__':
    run()
