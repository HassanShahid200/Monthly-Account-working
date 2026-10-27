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
AMOUNT = 'Deb./Cred. (LC)'
REQUIRED_COLUMNS = {"Series", "Due Date", "Remarks", "Doc. No."}
DEFAULT_COLUMNS = ["Acct Code", "Acct Name", "Posting Date", "Origin", "Doc No", "Remarks", "Offset Acct", "Offset Acct Name", "Deb./Cred. (LC)", "Deb./Cred. (FC)", "Balance Due (LC)"]

def normalize_code(series):
    return series.astype('string').str.strip().str.replace(r'\.0$', '', regex=True).replace('', pd.NA)

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

    ordered = [c for c in DEFAULT_COLUMNS if c in df.columns]
    ordered += [c for c in df.columns if c not in ordered]
    df = df.reindex(columns=ordered)
    
    df[CODE] = normalize_code(df[CODE])
    if df[CODE].isna().any():
        df = df.dropna(subset=[CODE])

    if "Posting Date" in df.columns:
        dates = pd.to_datetime(df['Posting Date'], errors='coerce')
        df['Posting Date'] = dates
        
    amounts = pd.to_numeric(df.get(AMOUNT, df.get('Deb./Cred. (FC)')), errors='coerce').fillna(0)
    df[AMOUNT] = amounts
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
        
    return df[['Account code', MAIN, SUB, 'Title']].dropna(subset=['Account code'])

def filter_by_date(df, start_date, end_date):
    mask = pd.Series(True, index=df.index)
    if start_date:
        mask &= (df['Posting Date'].dt.date >= pd.to_datetime(start_date).date())
    if end_date:
        mask &= (df['Posting Date'].dt.date <= pd.to_datetime(end_date).date())
    return df[mask]

def visible_rows(summary, expanded, col1_name, col2_name):
    rows, paths = [], []
    def append_row(part, path, label):
        old, new = float(part[col1_name].sum()), float(part[col2_name].sum())
        change = (new - old) / new if new != 0 else None
        
        marker = ('− ' if path in expanded else '+ ') if len(path) < 3 else ''
        if len(path) < 3:
            count = part[CODE].nunique()
            label = f'{label} ({count} {"account" if count == 1 else "accounts"})'
        
        rows.append({'Hierarchy': '    ' * (len(path) - 1) + marker + label,
                     col2_name: new, col1_name: old, '% Change': change,
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

def save_report_snapshot(df, title, comments, col1_config, col2_config):
    client = cloud_client()
    if not client:
        raise ValueError("Supabase credentials not found in secrets.")
    
    bucket = str(st.secrets.get('SUPABASE_BUCKET', 'Ledger_Reports'))
    report_id = str(uuid.uuid4())
    path = report_id + '/workbook.xlsx'
    token = secrets.token_urlsafe(32)
    
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    
    client.storage.from_(bucket).upload(path=path, file=output.getvalue())
    
    payload = {
        'id': report_id, 'title': title.strip() or 'Financial Report',
        'file_path': path, 'original_filename': 'mapped_ledger.xlsx',
        'comments': dict(comments), 
        'reporting_period': json.dumps({'col1': col1_config, 'col2': col2_config}),
        'share_token_hash': token_digest(token)
    }
    client.table('ledger_reports').insert(payload).execute()
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
    if not parsed.netloc:
        raise ValueError("Invalid APP_URL. Make sure it starts with https://")
    return urlunsplit((parsed.scheme, parsed.netloc, parsed.path or '/', urlencode({'report': token}), ''))

def run():
    st.set_page_config(layout='wide', page_title='Financial Dashboard')
    
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
            
            df = pd.read_excel(BytesIO(content))
            if 'Posting Date' in df.columns:
                df['Posting Date'] = pd.to_datetime(df['Posting Date'], errors='coerce')
                
            st.session_state.draft_comments = report_meta.get('comments', {})
            configs = json.loads(report_meta['reporting_period'])
            col1_config = configs.get('col1', {'name': '2025', 'start': None, 'end': None})
            col2_config = configs.get('col2', {'name': '2026', 'start': None, 'end': None})
            
            mode = "Shared View"
            read_only_mode = True
        except Exception as e:
            st.error(f"Could not load shared report: {e}")
            return
    else:
        st.title('SAP Financial Ledger Dashboard')
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
        col1_config = {'name': '2025', 'start': None, 'end': None}
        col2_config = {'name': '2026', 'start': None, 'end': None}
        
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
                
            df = pd.merge(raw_df, mapping_df, left_on=CODE, right_on='Account code', how='inner')
            if df.empty:
                st.error('No matching accounts found between the raw ledger and your mapping file.')
                return
                
            df[AMOUNT] = df[AMOUNT] * -1
            st.success(f"Matched {df[CODE].nunique()} unique accounts. Total transactions: {len(df)}")
            
            st.header("2. Calculation Settings")
            c1, c2 = st.columns(2)
            with c1:
                st.subheader("Column 1 Setup")
                c1_name = st.text_input("Column 1 Name", value="2025")
                c1_start = st.date_input("Start Date (Col 1)", value=None, key='c1s')
                c1_end = st.date_input("End Date (Col 1)", value=None, key='c1e')
            with c2:
                st.subheader("Column 2 Setup")
                c2_name = st.text_input("Column 2 Name", value="2026")
                c2_start = st.date_input("Start Date (Col 2)", value=None, key='c2s')
                c2_end = st.date_input("End Date (Col 2)", value=None, key='c2e')
                
            col1_config = {'name': c1_name, 'start': c1_start, 'end': c1_end}
            col2_config = {'name': c2_name, 'start': c2_start, 'end': c2_end}

            st.info(f"Percentage Change Formula: **({c2_name} - {c1_name}) / {c2_name}**")
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
                    st.session_state.loaded_df = pd.read_excel(BytesIO(content))
                    
                    st.session_state.draft_comments = report_meta.get('comments', {})
                    if report_meta.get('reporting_period'):
                        try:
                            configs = json.loads(report_meta['reporting_period'])
                            st.session_state.loaded_col1_config = configs.get('col1')
                            st.session_state.loaded_col2_config = configs.get('col2')
                        except:
                            pass
                            
            if 'loaded_df' in st.session_state:
                df = st.session_state.loaded_df
                df['Posting Date'] = pd.to_datetime(df['Posting Date'], errors='coerce')
                st.success("Report loaded successfully!")
                
                col1_config = st.session_state.get('loaded_col1_config', col1_config)
                col2_config = st.session_state.get('loaded_col2_config', col2_config)
            else:
                return

    # ========================================================
    # DASHBOARD LOGIC (Runs for New, Saved, and Shared Reports)
    # ========================================================
    col1_name = col1_config['name']
    col2_name = col2_config['name']
    
    df_col1 = filter_by_date(df, col1_config['start'], col1_config['end'])
    df_col2 = filter_by_date(df, col2_config['start'], col2_config['end'])
    
    col1_totals = df_col1.groupby(CODE)[AMOUNT].sum()
    col2_totals = df_col2.groupby(CODE)[AMOUNT].sum()
    
    summary = df[[MAIN, SUB, CODE, 'Title']].drop_duplicates()
    summary[col1_name] = summary[CODE].map(col1_totals).fillna(0)
    summary[col2_name] = summary[CODE].map(col2_totals).fillna(0)
    
    st.divider()
    if not read_only_mode:
        st.subheader('3. Profit and Loss Summary')
    
    all_branches = {(m,) for m in summary[MAIN].unique()}
    all_branches.update(tuple(r) for r in summary[[MAIN, SUB]].drop_duplicates().itertuples(index=False, name=None))
    table, paths = visible_rows(summary, all_branches, col1_name, col2_name)
    
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
                    button.style.cssText = 'cursor:pointer;padding:4px 12px;border:1px solid #a7bbce;border-radius:5px;background:#ffffff;color:#17324d;font-size:13px;';
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
                    button.style.cssText = 'cursor:pointer;width:26px;height:26px;margin-right:9px;border:1px solid currentColor;border-radius:5px;background:transparent;';
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
                text.textContent = p.data._level < 3 ? label.slice(2) : label;
                this.el.appendChild(text);
            }
            getGui() { return this.el; }
        }'''))
        
    money = JsCode("function(p) { return p.value == null ? '' : Number(p.value).toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2}); }")
    percent = JsCode("function(p) { return p.value == null ? 'N/A' : Number(p.value).toLocaleString('en-US', {style:'percent', minimumFractionDigits:2, maximumFractionDigits:2}); }")
    
    gb.configure_column(col2_name, type=['numericColumn'], minWidth=150, valueFormatter=money)
    gb.configure_column(col1_name, type=['numericColumn'], minWidth=150, valueFormatter=money)
    gb.configure_column('% Change', type=['numericColumn'], valueFormatter=percent)
    
    gb.configure_column('Comments', minWidth=280, width=360, wrapText=True, autoHeight=True,
                        editable=JsCode('function(p) { return !p.node.rowPinned; }') if not read_only_mode else False,
                        cellEditor='agLargeTextCellEditor', cellEditorPopup=True,
                        cellEditorParams={'maxLength': 4000, 'rows': 5, 'cols': 45})
                        
    total_new, total_old = float(summary[col2_name].sum()), float(summary[col1_name].sum())
    gb.configure_grid_options(
        headerHeight=78,
        stopEditingWhenCellsLoseFocus=True,
        getRowId=JsCode("function(p) { return p.data._path; }"),
        pinnedBottomRowData=[{'Hierarchy': 'Grand Total', col2_name: total_new,
                              col1_name: total_old,
                              '% Change': (total_new - total_old) / total_new if total_new else None}],
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
                    st.session_state.ledger_col = col2_name

    # ========================================================
    # LEDGER DRILL DOWN SECTION
    # ========================================================
    path = st.session_state.ledger_path
    if path is not None:
        st.divider()
        st.subheader('Ledger Details')
        
        st.write(f"**Selected scope:** {' → '.join(path)}")
        
        # Make sure the state is valid
        if st.session_state.ledger_col not in [col2_name, col1_name]:
            st.session_state.ledger_col = col2_name
            
        selected_col = st.radio('Select the ledger period to view:', [col2_name, col1_name], horizontal=True, key='ledger_col')
        
        if selected_col == col1_name:
            detail = filter_by_date(df, col1_config['start'], col1_config['end'])
            period_label = col1_name
        else:
            detail = filter_by_date(df, col2_config['start'], col2_config['end'])
            period_label = col2_name
        
        mask = detail[MAIN].eq(path[0])
        if len(path) >= 2: mask &= detail[SUB].eq(path[1])
        if len(path) == 3: mask &= detail[CODE].eq(path[2])
        
        detail = detail.loc[mask].copy()
        
        if 'Posting Date' in detail.columns:
            detail = detail.sort_values('Posting Date')
        
        a, b = st.columns(2)
        a.metric(f'Transactions · {period_label}', len(detail))
        b.metric('Mapped Net Amount (After -1)', f'{detail[AMOUNT].sum():,.2f}')
        
        detail['Debit'] = detail[AMOUNT].clip(lower=0)
        detail['Credit'] = -detail[AMOUNT].clip(upper=0)
        st.caption('Net = Debit − Credit.')
        
        columns = [CODE, 'Acct Name', MAIN, SUB, 'Title', 'Posting Date', 'Origin', 'Doc No',
                   'Remarks', 'Offset Acct', 'Offset Acct Name', 'Debit', 'Credit', AMOUNT]
        detail = detail[[col for col in columns if col in detail.columns]].copy()
        
        if 'Posting Date' in detail.columns:
            detail['Posting Date'] = detail['Posting Date'].dt.strftime('%Y-%m-%d')
            
        if detail.empty:
            st.info('No transactions were found for this selection and period.')
        else:
            total_row = {col: '' for col in detail.columns}
            total_row[CODE] = 'Grand Total'
            for col in ['Debit', 'Credit', AMOUNT]:
                total_row[col] = float(detail[col].sum())
                
            detail = pd.concat([detail, pd.DataFrame([total_row])], ignore_index=True)
            st.dataframe(detail, use_container_width=True, hide_index=True,
                         column_config={col: st.column_config.NumberColumn(format='%.2f')
                                        for col in ['Debit', 'Credit', AMOUNT]})
            
            st.download_button('Download Selected Ledger (CSV)', detail.to_csv(index=False).encode('utf-8-sig'),
                               file_name=f'ledger_{period_label}.csv', mime='text/csv')

    # ========================================================
    # SAVE TO CLOUD (ONLY FOR EDITORS)
    # ========================================================
    if not read_only_mode:
        st.divider()
        st.subheader("5. Save to Cloud")
        title = st.text_input('Report title', value=f'{col2_name} vs {col1_name} Financial Report')
        
        default_url = str(st.secrets.get('APP_URL', ''))
        base_url_input = st.text_input('Verify your App URL (Must start with https://):', value=default_url)
        
        if st.button('Save & Share Report', type='primary'):
            try:
                if not base_url_input or "https://" not in base_url_input:
                    st.error("Please enter a valid APP_URL starting with https://")
                else:
                    with st.spinner('Saving workbook and comments online...'):
                        share_token = save_report_snapshot(df, title, st.session_state.draft_comments, 
                                                           {'name': col1_name, 'start': str(col1_config['start']) if col1_config['start'] else None, 'end': str(col1_config['end']) if col1_config['end'] else None}, 
                                                           {'name': col2_name, 'start': str(col2_config['start']) if col2_config['start'] else None, 'end': str(col2_config['end']) if col2_config['end'] else None})
                    st.success('Report saved successfully!')
                    st.code(report_link(base_url_input, share_token))
            except Exception as exc:
                st.error(f'Save failed: {exc}')

if __name__ == '__main__':
    run()
