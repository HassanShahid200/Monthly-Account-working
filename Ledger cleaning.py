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
        df = pd.read_csv(BytesIO(content), sep="\t", encoding="utf-16", low_memory=False)
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
        df['Year'] = dates.dt.year

    amounts = pd.to_numeric(df.get(AMOUNT, df.get('Deb./Cred. (FC)')), errors='coerce').fillna(0)
    df[AMOUNT] = amounts
    df[MAIN] = 'Unclassified'
    df[SUB] = 'Unclassified'
    df['Title'] = df.get('Acct Name', df[CODE])
    return df

def make_summary(df):
    pivot = df.pivot_table(index=[MAIN, SUB, CODE, 'Title'], columns='Year', values=AMOUNT,
                           aggfunc='sum', fill_value=0, observed=True)
    pivot = pivot.reindex(columns=[2026, 2025], fill_value=0).mul(-1)
    pivot.columns = ['2026', '2025']
    return pivot.reset_index()

def visible_rows(summary, expanded):
    rows, paths = [], []
    def append_row(part, path, label):
        old, new = float(part['2025'].sum()), float(part['2026'].sum())
        change = (new - old) / old if old != 0 else None
        marker = ('− ' if path in expanded else '+ ') if len(path) < 3 else ''
        if len(path) < 3:
            count = part[CODE].nunique()
            label = f'{label} ({count} {"account" if count == 1 else "accounts"})'
        rows.append({'Hierarchy': '    ' * (len(path) - 1) + marker + label,
                     '2026': new, '2025': old, '% Change': change,
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

def save_report_snapshot(content, filename, title, comments, period):
    client = cloud_client()
    if not client:
        raise ValueError("Supabase credentials not found in secrets.")
    
    bucket = str(st.secrets.get('SUPABASE_BUCKET', 'Ledger_Reports'))
    report_id = str(uuid.uuid4())
    path = report_id + '/workbook.xlsx'
    token = secrets.token_urlsafe(32)
    
    client.storage.from_(bucket).upload(path=path, file=content)
    
    payload = {'id': report_id, 'title': title.strip() or 'Financial Report',
               'file_path': path, 'original_filename': filename,
               'comments': dict(comments), 'reporting_period': period,
               'share_token_hash': token_digest(token)}
    client.table('ledger_reports').insert(payload).execute()
    return token

def merge_comments(returned, valid_paths, comments):
    if not isinstance(returned, pd.DataFrame) or not {'_path', 'Comments'}.issubset(returned.columns):
        return
    valid = set(valid_paths)
    for _, row in returned.iterrows():
        path = row['_path']
        if path in valid:
            value = row['Comments']
            text = '' if pd.isna(value) else str(value)[:4000]
            if text:
                comments[path] = text
            else:
                comments.pop(path, None)

def run():
    st.set_page_config(layout='wide', page_title='Financial Dashboard')
    st.title('SAP Financial Ledger Dashboard')
    
    for key, default in [('expanded', set()), ('draft_comments', {}), ('ledger_path', None)]:
        st.session_state.setdefault(key, default)

    upload = st.file_uploader('Upload SAP Export (TXT, CSV, XLSX)', type=['txt', 'csv', 'xlsx'])
    if upload is None:
        st.info('Upload your export to start.')
        return
        
    content = upload.getvalue()
    filename = upload.name
    
    try:
        df = load_data(content, filename)
        if df.empty:
            st.warning('No valid transactions found.')
            return
            
        summary = make_summary(df)
        st.subheader('1. Profit and Loss Summary')
        
        all_branches = {(m,) for m in summary[MAIN].unique()}
        all_branches.update(tuple(r) for r in summary[[MAIN, SUB]].drop_duplicates().itertuples(index=False, name=None))
        table, paths = visible_rows(summary, all_branches)
        
        comments = st.session_state.draft_comments
        table['Comments'] = table['_path'].map(comments).fillna('')
        
        gb = GridOptionsBuilder.from_dataframe(table)
        gb.configure_default_column(sortable=False, filter=False, resizable=True, editable=False)
        for col in ['_id', '_level', '_action', '_path', '_expanded']:
            gb.configure_column(col, hide=True)
            
        gb.configure_column('Hierarchy', header_name='Main Class / Sub Class / Account', minWidth=450,
                            cellStyle={'whiteSpace': 'pre'},
                            headerComponent=JsCode("""class LevelHeader {
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
                                if (node.data && node.data._level <= index + 1) opened.push(node.data._path);
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
            }"""),
                            cellRenderer=JsCode("""class HierarchyRenderer {
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
            }"""))
            
        money = JsCode("function(p) { return p.value == null ? '' : Number(p.value).toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2}); }")
        percent = JsCode("function(p) { return p.value == null ? 'N/A' : Number(p.value).toLocaleString('en-US', {style:'percent', minimumFractionDigits:2, maximumFractionDigits:2}); }")
        
        gb.configure_column('2026', type=['numericColumn'], minWidth=150, valueFormatter=money)
        gb.configure_column('2025', type=['numericColumn'], minWidth=150, valueFormatter=money)
        gb.configure_column('% Change', type=['numericColumn'], valueFormatter=percent)
        
        gb.configure_column('Comments', minWidth=280, width=360, wrapText=True, autoHeight=True,
                            editable=JsCode('function(p) { return !p.node.rowPinned; }'),
                            cellEditor='agLargeTextCellEditor', cellEditorPopup=True,
                            cellEditorParams={'maxLength': 4000, 'rows': 5, 'cols': 45})
                            
        total_new, total_old = float(summary['2026'].sum()), float(summary['2025'].sum())
        gb.configure_grid_options(
            headerHeight=78,
            stopEditingWhenCellsLoseFocus=True,
            getRowId=JsCode("function(p) { return p.data._path; }"),
            pinnedBottomRowData=[{'Hierarchy': 'Grand Total', '2026': total_new,
                                  '2025': total_old,
                                  '% Change': (total_new - total_old) / total_old if total_old else None}],
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
                if (!p.data || p.node.rowPinned || ['_action', 'Comments'].includes(p.colDef.field)) return;
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
            getRowStyle=JsCode("""function(p) {
                if (!p.data) return {};
                if (p.node.rowPinned) return { backgroundColor: '#17324d', color: '#ffffff', fontWeight: '700'};
                if (p.data._level === 1) return { backgroundColor: '#17324d', color: '#ffffff', fontWeight: '700'};
                if (p.data._level === 2) return { backgroundColor: '#e7eff7', color: '#17324d', fontWeight: '600'};
                return { backgroundColor: '#ffffff', color: '#374151'};
            }""")
        )

        response = AgGrid(table, gridOptions=gb.build(), height=500, theme='streamlit',
                          allow_unsafe_jscode=True, update_on=['cellClicked', 'cellValueChanged'],
                          data_return_mode=DataReturnMode.AS_INPUT,
                          should_grid_return=JsCode('''function(p) {
                              const event = p.eventData;
                              if (!event || !event.data || (event.node && event.node.rowPinned)) return false;
                              const field = event.colDef && event.colDef.field;
                              if (p.streamlitRerunEventTriggerName === 'cellClicked') {
                                  return ['Hierarchy', '2026', '2025', '% Change'].includes(field);
                              }
                              return p.streamlitRerunEventTriggerName === 'cellValueChanged' && field === 'Comments';
                          }'''))

        returned = response.data
        merge_comments(returned, table['_path'], comments)
        if isinstance(returned, pd.DataFrame) and '_expanded' in returned.columns:
            states = returned['_expanded'].dropna()
            states = states.loc[states.ne('')]
            if not states.empty:
                st.session_state.expanded = {tuple(json.loads(item)) for item in json.loads(states.iloc[-1])}
                
        st.divider()
        st.subheader("3. Save to Cloud")
        title = st.text_input('Report title', value='2026 vs 2025 Financial Report')
        base_url = st.text_input('Your Streamlit app URL (optional)', placeholder='https://your-app.streamlit.app')
        
        if st.button('Save & Share Report', type='primary'):
            try:
                base_url = base_url or "https://example.streamlit.app"
                with st.spinner('Saving workbook and comments online...'):
                    share_token = save_report_snapshot(content, filename, title, comments, 'all')
                st.success('Report saved successfully!')
                st.code(f"{base_url}?report={share_token}")
            except Exception as exc:
                st.error(f'Save failed: Make sure Supabase secrets are set in Streamlit. Error: {exc}')

    except Exception as exc:
        st.error(f'Unable to process: {exc}')

if __name__ == '__main__':
    run()
