import base64
import io

import dash
from dash import dcc, html, dash_table, Input, Output, State
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
    'Unclassified': '#95a5a6',
}

SAMPLE_DATA = [
    ['REF001', 'CASE001', '200',
     '[{"MessageDesc":"DrugandClaimsSubstitutionWrapper_ACB : GCN in Exclusion List 25200"}]'],
    ['REF002', 'CASE002', '1302',
     '[{"MessageDesc":"Requested NDC: 12345678901 is Obsolete"},{"MessageDesc":"No drug Alternatives were found for the NDC"}]'],
    ['REF003', 'CASE003', '200',
     '[{"MessageDesc":"Requested NDC: 12345678901 is Obsolete"},{"MessageDesc":"Requested NDC: 12345678901 Substituted with Alternate NDC: 98765432101 For DAW Code: 1 Drug Source: W Substitution Indicator: B"},{"MessageDesc":"Qty has been changed from: 30 to: 30"}]'],
    ['REF004', 'CASE004', '200',
     '[{"MessageDesc":"GetClaimsData_ACB STEP:8 Original Drug: 47781085289 has been switched to CLAIM NDC: 00548232100, Claim Drug Source: Y"}]'],
    ['REF005', 'CASE005', '200',
     '[{"MessageDesc":"Requested NDC: 58406002104 Substituted with Alternate NDC: 58406002101 For DAW Code: 1 Drug Source: W Substitution Indicator: B"},{"MessageDesc":"Qty has been changed from: 26 to: 26"}]'],
    ['REF006', 'CASE006', '200',
     '[{"MessageDesc":"Requested NDC: 78206013802 Substituted with Alternate NDC: 78206013801 For DAW Code: 0 Drug Source: X Substitution Indicator: G"}]'],
    ['REF007', 'CASE007', '200',
     '[{"MessageDesc":"Requested NDC: 69654058030 Substituted with Alternate NDC: 00115173701 For DAW Code: 0 Drug Source: X Substitution Indicator: G"},{"MessageDesc":"Qty has been changed from: 10 to: 15"}]'],
]


# ── helpers ──────────────────────────────────────────────────────────────────

def make_sample_excel() -> bytes:
    df = pd.DataFrame(SAMPLE_DATA, columns=['Reference ID', 'Case ID', 'StatusID', 'LogTXT'])
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

    cat_df = (
        df.groupby('Category')
        .agg(Count=('Category', 'count'), Qty_Updated=('Qty Updated', 'sum'))
        .reset_index()
        .sort_values('Count', ascending=False)
    )

    fig = go.Figure(go.Bar(
        x=cat_df['Category'],
        y=cat_df['Count'],
        text=cat_df['Count'],
        textposition='outside',
        marker_color=[CATEGORY_COLORS.get(c, '#95a5a6') for c in cat_df['Category']],
        customdata=cat_df['Qty_Updated'],
        hovertemplate='<b>%{x}</b><br>Count: %{y:,}<br>Qty Updated: %{customdata:,}<extra></extra>',
    ))
    fig.update_layout(
        title='Click a category bar to drill down',
        yaxis_title='Count', xaxis_title='',
        plot_bgcolor='white', paper_bgcolor='white',
        height=400, showlegend=False,
        margin=dict(t=50, b=20, l=40, r=20),
    )
    fig.update_yaxes(showgrid=True, gridcolor='#f0f0f0')

    # Overall export button
    chart_card = dbc.Card([
        dbc.CardBody(dcc.Graph(id='cat-chart', figure=fig, config={'displayModeBar': False})),
        dbc.CardFooter(
            dbc.Button('Export ALL Results to Excel', id='btn-export-all',
                       color='success', size='sm', outline=True),
            className='text-end',
        ),
    ], className='shadow-sm')

    return cards, chart_card


@app.callback(
    Output('dl-results', 'data'),
    Input('btn-export-all', 'n_clicks'),
    State('stored-data', 'data'),
    prevent_initial_call=True,
)
def export_all(_, data):
    if not data:
        return None
    df = pd.read_json(io.StringIO(data), orient='split')
    return dcc.send_bytes(make_export_excel(df), 'analysis_results_all.xlsx')


@app.callback(
    Output('drilldown-chart', 'children'),
    Output('table-section', 'children'),
    Input('cat-chart', 'clickData'),
    State('stored-data', 'data'),
    prevent_initial_call=True,
)
def drill_down(click_data, data):
    if not click_data or not data:
        return '', ''

    df = pd.read_json(io.StringIO(data), orient='split')
    category = click_data['points'][0]['x']
    filtered = df[df['Category'] == category].copy()

    drilldown_section = ''

    # Sub-category chart for Obsolete Drug and NDC Substituted
    if category in ('Obsolete Drug', 'NDC Substituted (Drug Alt Service)'):
        sub_df = (
            filtered.groupby('Sub Category', dropna=False)
            .agg(Count=('Sub Category', 'count'), Qty_Updated=('Qty Updated', 'sum'))
            .reset_index()
            .sort_values('Count', ascending=False)
        )
        # Obsolete / Not Substituted — no qty tracking
        if category == 'Obsolete Drug':
            sub_df.loc[sub_df['Sub Category'] == 'Not Substituted', 'Qty_Updated'] = 0

        sub_fig = go.Figure(go.Bar(
            x=sub_df['Sub Category'],
            y=sub_df['Count'],
            text=sub_df['Count'],
            textposition='outside',
            marker_color='#3498db',
            customdata=sub_df['Qty_Updated'],
            hovertemplate='<b>%{x}</b><br>Count: %{y:,}<br>Qty Updated: %{customdata:,}<extra></extra>',
        ))
        sub_fig.update_layout(
            title=f'{category} — Sub-category Breakdown',
            yaxis_title='Count', xaxis_title='',
            plot_bgcolor='white', paper_bgcolor='white',
            height=350, showlegend=False,
            margin=dict(t=50, b=20, l=40, r=20),
        )
        sub_fig.update_yaxes(showgrid=True, gridcolor='#f0f0f0')

        drilldown_section = dbc.Card(
            dbc.CardBody(dcc.Graph(figure=sub_fig, config={'displayModeBar': False})),
            className='shadow-sm',
        )

    # Data table
    show_cols = ['Reference ID', 'Case ID', 'StatusID', 'Category', 'Sub Category', 'Qty Updated']
    table_df = filtered[show_cols].copy()
    table_df['Qty Updated'] = table_df['Qty Updated'].map(
        {True: 'Yes', False: 'No', 1: 'Yes', 0: 'No'}
    )

    table_section = dbc.Card([
        dbc.CardHeader(dbc.Row([
            dbc.Col(
                html.B(f'{category} — {len(filtered):,} records'),
                width=8, className='my-auto',
            ),
            dbc.Col(
                dbc.Button('Export this view to Excel', id='btn-export-view',
                           color='success', size='sm'),
                width=4, className='text-end',
            ),
        ])),
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


@app.callback(
    Output('dl-results', 'data', allow_duplicate=True),
    Input('btn-export-view', 'n_clicks'),
    State('stored-data', 'data'),
    State('cat-chart', 'clickData'),
    prevent_initial_call=True,
)
def export_view(_, data, click_data):
    if not data or not click_data:
        return None
    df = pd.read_json(io.StringIO(data), orient='split')
    category = click_data['points'][0]['x']
    filtered = df[df['Category'] == category]
    filename = f'analysis_{category.replace(" ", "_").replace("(", "").replace(")", "")}.xlsx'
    return dcc.send_bytes(make_export_excel(filtered), filename)


if __name__ == '__main__':
    app.run(debug=True, port=8050)
