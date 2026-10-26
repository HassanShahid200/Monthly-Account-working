import pandas as pd
import streamlit as st
import json
import io
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode, DataReturnMode

# -----------------------------------------------------------------------------
# 1. DATA CLEANING & LOADING (From Desktop App)
# -----------------------------------------------------------------------------
REQUIRED_COLUMNS = {"Series", "Due Date", "Remarks", "Doc. No."}
DEFAULT_COLUMNS = ["Acct Code", "Acct Name", "Posting Date", "Origin", "Doc No", "Remarks", "Offset Acct", "Offset Acct Name", "Deb./Cred. (LC)", "Deb./Cred. (FC)", "Balance Due (LC)"]

@st.cache_data
def load_and_clean_ledger(uploaded_file):
    """Reads TXT, CSV, or XLSX and applies the row-wise cleaning logic."""
    ext = uploaded_file.name.split('.')[-1].lower()
    
    if ext == 'txt':
        df = pd.read_csv(uploaded_file, sep="\t", encoding="utf-16", low_memory=False)
    elif ext == 'csv':
        df = pd.read_csv(uploaded_file, low_memory=False)
    elif ext in ['xlsx', 'xls']:
        df = pd.read_excel(uploaded_file)
    else:
        raise ValueError(f"Unsupported file format: {ext}")

    df.columns = df.columns.astype(str).str.strip()
    missing = sorted(REQUIRED_COLUMNS - set(df.columns))
    if missing:
        raise ValueError("Missing required column(s): " + ", ".join(missing))

    # Row-wise cleaning: Forward fill account codes/names from header rows
    df["Acct Code"] = df["Due Date"].where(df["Series"].isna()).ffill()
    df["Acct Name"] = df["Remarks"].where(df["Series"].isna()).ffill()
    df = df.loc[~df["Series"].isna()].copy()

    # Split Document Number
    parts = df["Doc. No."].astype("string").str.split(n=1, expand=True).reindex(columns=[0, 1])
    df["Origin"] = parts[0]
    df["Doc No"] = parts[1]
    
    ordered = [c for c in DEFAULT_COLUMNS if c in df.columns]
    ordered += [c for c in df.columns if c not in ordered]
    df = df.reindex(columns=ordered)
    
    if "Posting Date" in df.columns:
        df["Posting Date"] = pd.to_datetime(df["Posting Date"], errors="coerce")
        
    return df

@st.cache_data
def load_mapping(uploaded_mapping):
    """Loads the account classification mapping if provided."""
    mapping = pd.read_excel(uploaded_mapping)
    mapping.columns = mapping.columns.astype(str).str.strip()
    mapping['Account code'] = mapping['Account code'].astype('string').str.strip().str.replace(r'\.0$', '', regex=True)
    return mapping[['Account code', 'Current Main Class', 'Current Sub Class', 'Title']].dropna(subset=['Account code'])

# -----------------------------------------------------------------------------
# 2. DYNAMIC SUMMARY & HIERARCHY 
# -----------------------------------------------------------------------------
def filter_and_sum(df, date_col, amount_col, start_date, end_date):
    """Filters data based on dates and returns summed amounts per account."""
    mask = pd.Series(True, index=df.index)
    if start_date:
        mask &= (df[date_col] >= pd.to_datetime(start_date))
    if end_date:
        mask &= (df[date_col] <= pd.to_datetime(end_date))
    
    filtered = df.loc[mask].copy()
    filtered[amount_col] = pd.to_numeric(filtered[amount_col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
    return filtered.groupby('Acct Code')[amount_col].sum()

def visible_rows(summary, expanded, p1_name, p2_name, show_variance):
    rows, paths = [], []

    def append_row(part, path, label):
        old, new = float(part[p1_name].sum()), float(part[p2_name].sum())
        change = (new - old) / old if old != 0 else None
        marker = ('− ' if path in expanded else '+ ') if len(path) < 3 else ''
        if len(path) < 3:
            count = part['Acct Code'].nunique()
            label = f'{label} ({count} {"account" if count == 1 else "accounts"})'
        
        row_data = {
            'Hierarchy': '    ' * (len(path) - 1) + marker + label,
            p2_name: new, 
            p1_name: old, 
            '_level': len(path), 
            '_path': json.dumps(path, ensure_ascii=False, separators=(',', ':')),
            '_expanded': '', '_action': '', '_id': len(paths)
        }
        if show_variance:
            row_data['% Change'] = change
            
        rows.append(row_data)
        paths.append(path)

    main_values = sorted(summary['Current Main Class'].unique())
    for main in main_values:
        part = summary.loc[summary['Current Main Class'].eq(main)]
        append_row(part, (main,), main)
        if (main,) not in expanded:
            continue
        for sub, subgroup in part.groupby('Current Sub Class', sort=True):
            append_row(subgroup, (main, sub), sub)
            if (main, sub) in expanded:
                for _, account in subgroup.sort_values(['Title', 'Acct Code']).iterrows():
                    append_row(account.to_frame().T, (main, sub, account['Acct Code']),
                               f"{account['Title']} [{account['Acct Code']}]")
    return pd.DataFrame(rows), paths

# -----------------------------------------------------------------------------
# 3. STREAMLIT UI
# -----------------------------------------------------------------------------
st.set_page_config(layout='wide', page_title='Automated Financial Dashboard')

st.title("Ledger Cleaning & Dynamic Dashboard")

if 'expanded' not in st.session_state:
    st.session_state.expanded = set()

with st.sidebar:
    st.header("1. Upload Files")
    raw_file = st.file_uploader("Raw SAP Ledger (TXT, CSV, XLSX)", type=['txt', 'csv', 'xlsx', 'xls'])
    mapping_file = st.file_uploader("Account Mapping (Optional XLSX)", type=['xlsx', 'xls'])
    
if raw_file:
    try:
        with st.spinner("Cleaning row-wise ledger data..."):
            df = load_and_clean_ledger(raw_file)
            st.success(f"Cleaned {len(df)} transactions successfully.")
        
        # Apply Mapping if provided, else use Unclassified
        df['Acct Code'] = df['Acct Code'].astype('string').str.strip().str.replace(r'\.0$', '', regex=True)
        if mapping_file:
            mapping = load_mapping(mapping_file)
            df = df.merge(mapping, left_on='Acct Code', right_on='Account code', how='left')
            df['Current Main Class'] = df['Current Main Class'].fillna('Unclassified')
            df['Current Sub Class'] = df['Current Sub Class'].fillna('Unclassified')
            df['Title'] = df['Title'].fillna(df['Acct Name']).fillna(df['Acct Code'])
        else:
            df['Current Main Class'] = 'Unclassified'
            df['Current Sub Class'] = 'Unclassified'
            df['Title'] = df['Acct Name']

        st.subheader("2. Configure Balance Columns")
        date_col = st.selectbox("Select Posting Date Column", [c for c in df.columns if 'Date' in c], index=0)
        amount_col = st.selectbox("Select Amount Column", [c for c in df.columns if 'LC' in c or 'FC' in c], index=0)
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("### First Column (e.g., 2025 or Base Period)")
            p1_name = st.text_input("Column 1 Name", value="2025")
            p1_start = st.date_input("Start Date (Column 1)", value=None)
            p1_end = st.date_input("End Date (Column 1)", value=None)
            
        with col2:
            st.markdown("### Second Column (e.g., 2026 or Current Period)")
            p2_name = st.text_input("Column 2 Name", value="2026")
            p2_start = st.date_input("Start Date (Column 2)", value=None)
            p2_end = st.date_input("End Date (Column 2)", value=None)

        show_variance = st.checkbox(f"Calculate % Change ({p2_name} vs {p1_name})", value=True)

        if st.button("Generate Dashboard", type="primary"):
            p1_totals = filter_and_sum(df, date_col, amount_col, p1_start, p1_end)
            p2_totals = filter_and_sum(df, date_col, amount_col, p2_start, p2_end)
            
            # Create unique account list with classes
            summary = df[['Acct Code', 'Title', 'Current Main Class', 'Current Sub Class']].drop_duplicates()
            summary[p1_name] = summary['Acct Code'].map(p1_totals).fillna(0) * -1  # Standard P&L flip
            summary[p2_name] = summary['Acct Code'].map(p2_totals).fillna(0) * -1
            
            st.subheader(f"{p2_name} vs {p1_name} Summary")
            
            all_branches = {(m,) for m in summary['Current Main Class'].unique()}
            all_branches.update(tuple(r) for r in summary[['Current Main Class', 'Current Sub Class']].drop_duplicates().itertuples(index=False, name=None))
            
            table, paths = visible_rows(summary, all_branches, p1_name, p2_name, show_variance)
            
            gb = GridOptionsBuilder.from_dataframe(table)
            gb.configure_default_column(sortable=False, filter=False, resizable=True)
            gb.configure_column('_id', hide=True)
            gb.configure_column('_level', hide=True)
            gb.configure_column('_action', hide=True)
            gb.configure_column('_path', hide=True)
            gb.configure_column('_expanded', hide=True)
            
            money_format = JsCode("function(p) { return p.value == null ? '' : Number(p.value).toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2}); }")
            gb.configure_column(p2_name, type=['numericColumn'], valueFormatter=money_format)
            gb.configure_column(p1_name, type=['numericColumn'], valueFormatter=money_format)
            
            if show_variance:
                percent_format = JsCode("function(p) { return p.value == null ? 'N/A' : Number(p.value).toLocaleString('en-US', {style:'percent', minimumFractionDigits:2}); }")
                gb.configure_column('% Change', type=['numericColumn'], valueFormatter=percent_format)
            
            # AgGrid configuration for hierarchy
            gb.configure_grid_options(
                getRowId=JsCode("function(p) { return p.data._path; }"),
                onCellClicked=JsCode('''function(p) {
                    if (!p.data || ['_action'].includes(p.colDef.field)) return;
                    p.api.forEachNode(function(node) { node.data._action = ''; });
                    p.data._expanded = JSON.stringify(window.__ledgerExpanded || []);
                    p.data._action = 'ledger:' + p.colDef.field + ':' + Date.now();
                }'''),
                getRowStyle=JsCode("""function(p) {
                    if (!p.data) return {};
                    if (p.data._level === 1) return {backgroundColor: '#17324d', color: '#ffffff', fontWeight: '700'};
                    if (p.data._level === 2) return {backgroundColor: '#e7eff7', color: '#17324d', fontWeight: '600'};
                    return {backgroundColor: '#ffffff', color: '#374151'};
                }""")
            )
            
            AgGrid(table, gridOptions=gb.build(), height=600, theme='streamlit', allow_unsafe_jscode=True, data_return_mode=DataReturnMode.AS_INPUT)
            
    except Exception as e:
        st.error(f"Error processing file: {e}")
else:
    st.info("Upload your raw SAP B1 Ledger file in the sidebar to begin.")
