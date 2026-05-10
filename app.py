import base64
import io
import random

import dash
from dash import dcc, html, dash_table, Input, Output, State, ALL, ctx
import dash_bootstrap_components as dbc
import pandas as pd
import plotly.graph_objects as go

from classifier import classify_dataframe

app = dash.Dash(
    __name__,
    external_stylesheets=[dbc.themes.FLATLY],
    suppress_callback_exceptions=True,
    title='Analysis Dashboard',
)

CATEGORY_COLORS = {
    'Exclusion List': '#e74c3c',
    'Obsolete Drug': '#e67e22',
    'Claims Substitution': '#3498db',
    'NDC Substituted (Drug Alt Service)': '#27ae60',
    'Qty Updated': '#9b59b6',
    'Unclassified': '#95a5a6',
}

BASE_CATS = [
    'Exclusion List',
    'Obsolete Drug',
    'Claims Substitution',
    'NDC Substituted (Drug Alt Service)',
]

TILE_LABELS = {
    'Exclusion List': 'Exclusion List',
    'Obsolete Drug': 'Obsolete Drug',
    'Claims Substitution': 'Claims Substitution',
    'NDC Substituted (Drug Alt Service)': 'NDC Substituted',
    'Qty Updated': 'Qty Updated',
}

# ── helpers ──────────────────────────────────────────────────────────────────

def make_tile(category, count):
    color = CATEGORY_COLORS.get(category, '#95a5a6')
    label = TILE_LABELS.get(category, category)
    return dbc.Col(
        html.Div(
            [
                html.P(label, className='mb-1 fw-semibold',
                       style={'fontSize': '13px', 'color': '#555', 'lineHeight': '1.3'}),
                html.H2(f'{count:,}', className='mb-0 fw-bold', style={'color': color}),
                html.Small('records', style={'color': '#aaa', 'fontSize': '11px'}),
            ],
            id={'type': 'cat-tile', 'index': category},
            n_clicks=0,
            style={
                'borderTop': f'5px solid {color}',
                'borderRadius': '6px',
                'cursor': 'pointer',
                'background': 'white',
                'boxShadow': '0 2px 8px rgba(0,0,0,0.08)',
                'padding': '20px 10px',
                'textAlign': 'center',
            },
        ),
        width=True,
    )

def _rnd_ndc():
    return ''.join([str(random.randint(0, 9)) for _ in range(11)])

def _msg(*descs):
    parts = ','.join(f'{{"MessageDesc":"{d}"}}' for d in descs)
    return f'[{parts}]'

def _qty_msg():
    q = random.randint(10, 90)
    return f'Qty has been changed from: {q} to: {q}'

def make_sample_excel() -> bytes:
    random.seed(42)
    rows = []
    n = 1

    def row(status, log):
        nonlocal n
        r = [f'REF{n:04d}', f'CASE{n:04d}', status, log]
        n += 1
        return r

    # 1. Exclusion List — 15 rows (alternating GCN 25200 / 94200)
    for i in range(15):
        gcn = 25200 if i % 2 == 0 else 94200
        rows.append(row('200', _msg(
            f'DrugandClaimsSubstitutionWrapper_ACB : GCN in Exclusion List {gcn}'
        )))

    # 2. Obsolete Drug — Not Substituted — 15 rows
    for _ in range(15):
        rows.append(row('1302', _msg(
            f'Requested NDC: {_rnd_ndc()} is Obsolete',
            'No drug Alternatives were found for the NDC',
        )))

    # 3. Obsolete Drug — Substituted — 15 rows (every 3rd has qty)
    for i in range(15):
        ndc, alt = _rnd_ndc(), _rnd_ndc()
        descs = [
            f'Requested NDC: {ndc} is Obsolete',
            f'Requested NDC: {ndc} Substituted with Alternate NDC: {alt} '
            f'For DAW Code: {random.randint(0,1)} Drug Source: {random.choice(["W","X"])} Substitution Indicator: B',
        ]
        if i % 3 == 0:
            descs.append(_qty_msg())
        rows.append(row('200', _msg(*descs)))

    # 4. Claims Substitution — 15 rows (every 3rd has qty)
    for i in range(15):
        orig, claim = _rnd_ndc(), _rnd_ndc()
        descs = [
            f'GetClaimsData_ACB STEP:8 Original Drug: {orig} has been switched to '
            f'CLAIM NDC: {claim}, Claim Drug Source: Y',
        ]
        if i % 3 == 0:
            descs.append(_qty_msg())
        rows.append(row('200', _msg(*descs)))

    # 5. NDC Substituted — 6 sub-types × 10 rows each = 60 rows (every 3rd has qty)
    for daw in ['0', '1']:
        for src in ['W', 'X', 'Y']:
            for i in range(10):
                ndc, alt = _rnd_ndc(), _rnd_ndc()
                descs = [
                    f'Requested NDC: {ndc} Substituted with Alternate NDC: {alt} '
                    f'For DAW Code: {daw} Drug Source: {src} Substitution Indicator: B',
                ]
                if i % 3 == 0:
                    descs.append(_qty_msg())
                rows.append(row('200', _msg(*descs)))

    # Total: 15+15+15+15+60 = 120 rows
    df = pd.DataFrame(rows, columns=['Reference ID', 'Case ID', 'StatusID', 'LogTXT'])
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf.getvalue()


def make_export_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        # Summary by category
        cat_sum = df.groupby('Category', dropna=False).agg(
            Count=('Category', 'count'),
            Qty_Updated=('Qty Updated', 'sum'),
        ).reset_index()
        cat_sum.rename(columns={'Qty_Updated': 'Qty Updated Count'}, inplace=True)
        cat_sum.to_excel(writer, index=False, sheet_name='Summary by Category')

        # Summary by sub-category
        sub_sum = df.groupby(['Category', 'Sub Category'], dropna=False).agg(
            Count=('Category', 'count'),
            Qty_Updated=('Qty Updated', 'sum'),
        ).reset_index()
        sub_sum.rename(columns={'Qty_Updated': 'Qty Updated Count'}, inplace=True)
        sub_sum.to_excel(writer, index=False, sheet_name='Summary by Sub-Category')

        # All records
        out = df.copy()
        out['Qty Updated'] = out['Qty Updated'].map(
            {True: 'Yes', False: 'No', 1: 'Yes', 0: 'No'}
        )
        out.to_excel(writer, index=False, sheet_name='All Records')

    buf.seek(0)
    return buf.getvalue()


def stat_card(label: str, value: str, color: str):
    return dbc.Card(
        dbc.CardBody([
            html.H3(value, className=f'text-{color} mb-0 fw-bold'),
            html.P(label, className='text-muted mb-0 small'),
        ]),
        className='text-center shadow-sm h-100',
    )


# ── layout ───────────────────────────────────────────────────────────────────

app.layout = dbc.Container([
    dcc.Store(id='stored-data'),
    dcc.Download(id='dl-sample'),
    dcc.Download(id='dl-results'),

    # Header
    dbc.Row([
        dbc.Col([
            html.H4('Analysis Dashboard', className='mb-0 text-white fw-bold'),
        ], width=8),
        dbc.Col(
            dbc.Button('⬇ Download Sample Excel', id='btn-sample', color='light',
                       outline=True, size='sm', className='mt-1'),
            width=4, className='text-end',
        ),
    ], className='bg-primary p-3 rounded mb-4'),

    # Upload
    dcc.Upload(
        id='upload',
        children=html.Div([
            'Drag & Drop or ', html.A('Select your Excel File', style={'fontWeight': '600'}),
        ]),
        style={
            'width': '100%', 'height': '80px', 'lineHeight': '40px',
            'border': '2px dashed #3498db', 'borderRadius': '8px',
            'textAlign': 'center', 'background': '#f8f9fa',
            'cursor': 'pointer', 'paddingTop': '10px',
        },
    ),
    html.Div(id='status', className='mt-2 mb-2'),

    # Dynamic sections
    html.Div(id='summary-cards', className='mb-4'),
    html.Div(id='main-chart', className='mb-4'),
    html.Div(id='drilldown-chart', className='mb-3'),
    html.Div(id='table-section', className='mb-5'),

], fluid=True, style={'fontFamily': 'Segoe UI, sans-serif', 'padding': '20px'})


# ── callbacks ─────────────────────────────────────────────────────────────────

@app.callback(
    Output('dl-sample', 'data'),
    Input('btn-sample', 'n_clicks'),
    prevent_initial_call=True,
)
def download_sample(_):
    return dcc.send_bytes(make_sample_excel(), 'sample_input.xlsx')


@app.callback(
    Output('stored-data', 'data'),
    Output('status', 'children'),
    Input('upload', 'contents'),
    State('upload', 'filename'),
    prevent_initial_call=True,
)
def process_upload(contents, filename):
    _, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)
    try:
        df = pd.read_excel(io.BytesIO(decoded))
        required = {'Reference ID', 'Case ID', 'StatusID', 'LogTXT'}
        missing = required - set(df.columns)
        if missing:
            return None, dbc.Alert(f"Missing columns: {', '.join(missing)}", color='danger')
        df = classify_dataframe(df)
        msg = dbc.Alert(
            f"Processed {len(df):,} records from '{filename}'",
            color='success', dismissable=True, duration=6000,
        )
        return df.to_json(orient='split', date_format='iso'), msg
    except Exception as exc:
        return None, dbc.Alert(f'Error: {exc}', color='danger')


@app.callback(
    Output('summary-cards', 'children'),
    Output('main-chart', 'children'),
    Input('stored-data', 'data'),
    prevent_initial_call=True,
)
def render_summary(data):
    if not data:
        return '', ''

    df = pd.read_json(io.StringIO(data), orient='split')
    total = len(df)
    classified = int((df['Category'] != 'Unclassified').sum())
    qty_total = int(df['Qty Updated'].sum())

    cards = dbc.Row([
        dbc.Col(stat_card('Total Records', f'{total:,}', 'primary'), width=3),
        dbc.Col(stat_card('Classified', f'{classified:,}', 'success'), width=3),
        dbc.Col(stat_card('Unclassified', f'{total - classified:,}', 'warning'), width=3),
        dbc.Col(stat_card('Qty Updated (all)', f'{qty_total:,}', 'info'), width=3),
    ], className='g-3')

    # Build 5-row chart: 4 base categories + Qty Updated
    base_df = (
        df[df['Category'].isin(BASE_CATS)]
        .groupby('Category')
        .size()
        .reset_index(name='Count')
    )
    qty_total = int(df['Qty Updated'].sum())
    qty_row = pd.DataFrame([{'Category': 'Qty Updated', 'Count': qty_total}])
    cat_df = pd.concat([base_df, qty_row], ignore_index=True)

    # Tiles row
    tiles = dbc.Row(
        [make_tile(cat, cnt) for cat, cnt in zip(cat_df['Category'], cat_df['Count'])],
        className='g-3 mb-4',
    )

    # Donut chart
    donut_fig = go.Figure(go.Pie(
        labels=[TILE_LABELS.get(c, c) for c in cat_df['Category']],
        values=cat_df['Count'],
        hole=0.55,
        marker=dict(
            colors=[CATEGORY_COLORS.get(c, '#95a5a6') for c in cat_df['Category']],
            line=dict(color='white', width=3),
        ),
        textinfo='label+percent',
        textfont=dict(size=12),
        hovertemplate='<b>%{label}</b><br>Count: %{value:,}<br>Share: %{percent}<extra></extra>',
        showlegend=False,
    ))
    donut_fig.update_layout(
        plot_bgcolor='white', paper_bgcolor='white',
        height=360,
        margin=dict(t=20, b=20, l=20, r=20),
        font=dict(family='Segoe UI, sans-serif', size=12),
        annotations=[dict(
            text=f'<b>{len(df):,}</b><br>Total', x=0.5, y=0.5,
            font=dict(size=16, color='#333'), showarrow=False,
        )],
    )

    donut_card = dbc.Card([
        dbc.CardBody(
            dcc.Graph(id='donut-chart', figure=donut_fig, config={'displayModeBar': False}),
        ),
        dbc.CardFooter(
            dbc.Button('Export ALL Results to Excel', id='btn-export-all',
                       color='success', size='sm', outline=True),
            className='text-end',
        ),
    ], className='shadow-sm')

    return cards, html.Div([tiles, donut_card])


@app.callback(
    Output('dl-results', 'data'),
    Input('btn-export-all', 'n_clicks'),
    State('stored-data', 'data'),
    prevent_initial_call=True,
)
def export_all(n_clicks, data):
    if not n_clicks or not data:
        return None
    df = pd.read_json(io.StringIO(data), orient='split')
    return dcc.send_bytes(make_export_excel(df), 'analysis_results_all.xlsx')


@app.callback(
    Output('drilldown-chart', 'children'),
    Output('table-section', 'children'),
    Input({'type': 'cat-tile', 'index': ALL}, 'n_clicks'),
    Input('donut-chart', 'clickData'),
    State('stored-data', 'data'),
    prevent_initial_call=True,
)
def drill_down(tile_clicks, donut_click, data):
    if not data:
        return '', ''

    trigger = ctx.triggered_id
    if trigger is None:
        return '', ''

    if isinstance(trigger, dict) and trigger.get('type') == 'cat-tile':
        category = trigger['index']
    elif trigger == 'donut-chart' and donut_click:
        point = donut_click['points'][0]
        label = point.get('label', '')
        # map short label back to full category key
        category = next((k for k, v in TILE_LABELS.items() if v == label), label)
    else:
        return '', ''

    if not category:
        return '', ''

    df = pd.read_json(io.StringIO(data), orient='split')

    # Qty Updated is a virtual category — filter on flag, not on Category column
    if category == 'Qty Updated':
        filtered = df[df['Qty Updated'].isin([True, 1])].copy()
    else:
        filtered = df[df['Category'] == category].copy()

    drilldown_section = ''

    if category in ('Obsolete Drug', 'NDC Substituted (Drug Alt Service)'):
        sub_df = (
            filtered.groupby('Sub Category', dropna=False)
            .agg(Count=('Sub Category', 'count'), Qty_Updated=('Qty Updated', 'sum'))
            .reset_index()
        )
        if category == 'Obsolete Drug':
            sub_df.loc[sub_df['Sub Category'] == 'Not Substituted', 'Qty_Updated'] = 0

        sub_sorted = sub_df.sort_values('Count', ascending=True)

        sub_fig = go.Figure()
        sub_fig.add_trace(go.Bar(
            y=sub_sorted['Sub Category'],
            x=sub_sorted['Count'],
            orientation='h',
            name='Total',
            marker=dict(color='#3498db', line=dict(color='white', width=1)),
            text=[f'  {v:,}' for v in sub_sorted['Count']],
            textposition='outside',
            hovertemplate='<b>%{y}</b><br>Total: %{x:,}<extra></extra>',
        ))
        sub_fig.add_trace(go.Bar(
            y=sub_sorted['Sub Category'],
            x=sub_sorted['Qty_Updated'],
            orientation='h',
            name='Qty Updated',
            marker=dict(color='#27ae60', line=dict(color='white', width=1)),
            hovertemplate='<b>%{y}</b><br>Qty Updated: %{x:,}<extra></extra>',
        ))
        sub_fig.update_layout(
            title=dict(text=f'{category} — Sub-category Breakdown', font=dict(size=14)),
            barmode='overlay',
            plot_bgcolor='white', paper_bgcolor='white',
            height=max(300, len(sub_df) * 55 + 80),
            legend=dict(orientation='h', x=0, y=1.12, font=dict(size=11)),
            margin=dict(t=70, b=10, l=10, r=80),
            font=dict(family='Segoe UI, sans-serif', size=12),
        )
        sub_fig.update_xaxes(showgrid=True, gridcolor='#f0f0f0', zeroline=False)
        sub_fig.update_yaxes(showgrid=False)
        drilldown_section = dbc.Card(
            dbc.CardBody(dcc.Graph(figure=sub_fig, config={'displayModeBar': False})),
            className='shadow-sm',
        )

    elif category == 'Qty Updated':
        qty_breakdown = (
            filtered.groupby('Category').size().reset_index(name='Count')
            .sort_values('Count', ascending=True)
        )
        qty_colors = [CATEGORY_COLORS.get(c, '#95a5a6') for c in qty_breakdown['Category']]
        qty_fig = go.Figure(go.Bar(
            y=qty_breakdown['Category'],
            x=qty_breakdown['Count'],
            orientation='h',
            marker=dict(color=qty_colors, line=dict(color='white', width=1)),
            text=[f'  {v:,}' for v in qty_breakdown['Count']],
            textposition='outside',
            hovertemplate='<b>%{y}</b><br>Qty Updated: %{x:,}<extra></extra>',
        ))
        qty_fig.update_layout(
            title=dict(text='Qty Updated — Breakdown by Category', font=dict(size=14)),
            plot_bgcolor='white', paper_bgcolor='white',
            height=max(300, len(qty_breakdown) * 55 + 80),
            showlegend=False,
            margin=dict(t=60, b=10, l=10, r=80),
            font=dict(family='Segoe UI, sans-serif', size=12),
        )
        qty_fig.update_xaxes(showgrid=True, gridcolor='#f0f0f0', zeroline=False)
        qty_fig.update_yaxes(showgrid=False)
        drilldown_section = dbc.Card(
            dbc.CardBody(dcc.Graph(figure=qty_fig, config={'displayModeBar': False})),
            className='shadow-sm',
        )

    # Data table
    show_cols = ['Reference ID', 'Case ID', 'StatusID', 'Category', 'Sub Category', 'Qty Updated']
    table_df = filtered[show_cols].copy()
    table_df['Qty Updated'] = table_df['Qty Updated'].map(
        {True: 'Yes', False: 'No', 1: 'Yes', 0: 'No'}
    )

    table_section = dbc.Card([
        dbc.CardHeader(html.B(f'{category} — {len(filtered):,} records')),
        dbc.CardBody(
            dash_table.DataTable(
                data=table_df.to_dict('records'),
                columns=[{'name': c, 'id': c} for c in show_cols],
                page_size=15,
                filter_action='native',
                sort_action='native',
                style_table={'overflowX': 'auto'},
                style_header={
                    'backgroundColor': '#2c3e50', 'color': 'white',
                    'fontWeight': 'bold', 'padding': '10px',
                },
                style_cell={'textAlign': 'left', 'padding': '8px', 'fontSize': '13px'},
                style_data_conditional=[
                    {'if': {'row_index': 'odd'}, 'backgroundColor': '#f8f9fa'},
                    {'if': {'filter_query': '{Qty Updated} = "Yes"'},
                     'backgroundColor': '#eafaf1'},
                ],
            )
        ),
    ], className='shadow-sm')

    return drilldown_section, table_section



if __name__ == '__main__':
    app.run(debug=True, port=8050)
