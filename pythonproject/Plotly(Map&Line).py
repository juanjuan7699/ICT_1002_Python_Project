import pandas as pd
import plotly.graph_objects as go
from plotly.express import choropleth
import dash
import dash_core_components as dcc
import dash_html_components as html

covid = pd.read_excel (r'C:\Users\jiaji\PycharmProjects\pythonProject1\covidinasean.xlsx')

#choropleth map
df = covid[['iso_code','location', 'Week_Number','total_cases','total_deaths','total_tests']]
chart1=choropleth(df, locations='iso_code',hover_name='location',hover_data=df.columns,color='total_cases',animation_frame="Week_Number",color_continuous_scale='rainbow',title='Total Covid Cases in ASEAN')


#line graph with drop down
df1 = covid[['location', 'Week_Number','total_cases','total_tests']]
total_cases = df1.pivot_table('total_cases', 'Week_Number', 'location')
total_tests = df1.pivot_table('total_tests', 'Week_Number', 'location')

# create figure
chart2 = go.Figure()

# set up traces
chart2.add_trace(go.Scatter(x=total_cases.index,
                         y=total_cases[total_cases.columns[0]],
                         visible=True)
             )

chart2.add_trace(go.Scatter(x=total_tests.index,
                         y=total_tests[total_tests.columns[0]],
                         visible=True)
             )

updatemenu = []
buttons = []

# button with one option for each country
for col in total_cases.columns:
    buttons.append(dict(method='restyle',
                        label=col,
                        visible=True,
                        args=[{'y':[total_cases[col]],
                              'x':[total_cases.index],
                               'type':'scatter'}],
                        )
                  )
buttons1 = []
for col in total_tests.columns:
    buttons1.append(dict(method='restyle',
                        label=col,
                        visible=True,
                        args=[{'y': [total_tests[col]],
                               'x': [total_tests.index],
                               'type': 'scatter'},[1]],
                        )
                   )

# some adjustments to the updatemenus
updatemenu=[]
your_menu=dict()
updatemenu.append(your_menu)
your_menu2=dict()
updatemenu.append(your_menu2)
updatemenu[0]['buttons']=buttons
updatemenu[0]['direction']='down'
updatemenu[0]['showactive']=True
updatemenu[1]['buttons']=buttons1
updatemenu[1]['showactive']=True
updatemenu[1]['y']=0.5

# add dropdown menus to the figure
chart2.update_layout(showlegend=False, updatemenus=updatemenu,title='ASEAN Confirmed Cases Weekly Trend')
# Add annotation
chart2.update_layout(
    annotations=[
        go.layout.Annotation(text="<b>Total Cases:</b>",
                             x=-0.35, xref="paper",
                             y=1.15, yref="paper",
                             align="left", showarrow=False),
        go.layout.Annotation(text="<b>Total Tests:</b>",
                             x=-0.35, xref="paper", y=0.6,
                             yref="paper", showarrow=False),
                  ]
)


#dashboard
external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']
app = dash.Dash(__name__, external_stylesheets=external_stylesheets)

graph1 = dcc.Graph(
        id='graph1',
        figure=chart1,
        className="six columns"
    )
graph2 = dcc.Graph(
        id='graph2',
        figure=chart2,
        className="six columns"
    )

header = html.H2(children="COVID Dataset Analysis")
row1 = html.Div(children=[graph1, graph2],)
layout = html.Div(children=[header, row1], style={"text-align": "center"})
app.layout = layout

if __name__ == "__main__":
    app.run_server(debug=True)
