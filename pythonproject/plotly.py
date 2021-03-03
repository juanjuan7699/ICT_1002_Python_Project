import pandas as pd #install pandas
import plotly
import plotly.express as px #install plotly

import dash #install dash
import dash_table
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output

df = pd.read_excel (r'C:\Users\cocov\Downloads\covidinasean.xlsx') #read excel file

app = dash.Dash(__name__)
server = app.server
app.title = 'COVID-19 Data Analysis'


dff = df.groupby('location', as_index=False)[['total_deaths','total_cases']].sum()
print (dff[:10])

app.layout = html.Div([ #creating dash layout
    html.Div([
        html.H1('COVID-19 Data Analysis'),
        html.Div([
            html.P('Analysis of COVID-19 in ASEAN countries'),

        ])
    ], style={'text-align': 'center'}),
    html.Div([
        dash_table.DataTable( #datatable
            id='datatable_id',
            data=dff.to_dict('records'),
            columns=[
                {"name": i, "id": i, "deletable": False, "selectable": False} for i in dff.columns
            ],
            editable=False,
            filter_action="native",
            sort_action="native",
            sort_mode="multi",
            row_selectable="multi",
            row_deletable=False,
            selected_rows=[],
            page_action="native",
            page_current= 0,
            page_size= 6,

            style_cell_conditional=[
                {'if': {'column_id': 'location'},
                 'width': '40%', 'textAlign': 'left'},
                {'if': {'column_id': 'total_deaths'},
                 'width': '30%', 'textAlign': 'left'},
                {'if': {'column_id': 'total_cases'},
                 'width': '30%', 'textAlign': 'left'},
            ],
        ),
    ],className='row'),


    html.Div([
        html.Div([
            dcc.Dropdown(id='linedropdown', #dropdown for line graph
                options=[
                        {'label': 'Total COVID Cases', 'value': 'total_cases'},
                         {'label': 'Total COVID Deaths', 'value': 'total_deaths'}

                ],
                value='total_cases',
                multi=False,
                clearable=False
            ),
        ],className='six columns')



    ],className='row'),

    html.Div([

        html.Div([
            dcc.Graph(id='linechart'), #line graph
        ],className='six columns'),
        html.Div([
            dcc.Dropdown(id='piedropdown', #dropdown for pie
                         options=[
                             {'label': 'Total COVID Cases', 'value': 'total_cases'},
                             {'label': 'Total COVID Deaths', 'value': 'total_deaths'}

                         ],
                         value='total_cases',
                         multi=False,
                         clearable=False
                         ),
        ], className='six columns'),
        html.Div([
            dcc.Graph(id='piechart'), #pie chart
        ],className='six columns'),

    ],className='row'),


])


@app.callback(
    [Output('piechart', 'figure'),
     Output('linechart', 'figure')],
    [Input('datatable_id', 'selected_rows'),
     Input('piedropdown', 'value'),
     Input('linedropdown', 'value')]
)
def update_data(chosen_rows,piedropval,linedropval): #function to filter data from datatable
    if len(chosen_rows)==0:
        df_filterd = dff[dff['location'].isin(['Brunei','Cambodia','Indonesia','Malaysia', 'Laos', 'Myanmar', 'Philippines', 'Singapore', 'Thailand', 'Timor'])]
    else:
        print(chosen_rows)
        df_filterd = dff[dff.index.isin(chosen_rows)]

    pie_chart=px.pie(
            data_frame=df_filterd,
            names='location',
            values=piedropval,
            hole=.3,
            labels={'Countries':'location'}
            )


    list_chosen_countries=df_filterd['location'].tolist()
    df_line = df[df['location'].isin(list_chosen_countries)]

    line_chart = px.line(
            data_frame=df_line,
            x='Week_Number',
            y=linedropval,
            color='location',
            labels={'location':'location', 'dateRep':'date'},
            )
    line_chart.update_layout(uirevision='foo')

    return (pie_chart,line_chart)



if __name__ == '__main__':
    app.run_server(debug=True)
