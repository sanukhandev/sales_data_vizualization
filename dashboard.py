import dash
import dash_core_components as dcc
from dash import html
import plotly.express as px
import pandas as pd
import statsmodels.api as sm
import networkx as nx
from pyvis.network import Network
import dash_bootstrap_components as dbc
import dash_cytoscape as cyto

# Load the data from Excel
df = pd.read_excel('Financial Sample.xlsx')  # replace with your file path

# Strip leading and trailing spaces from column names
df.columns = df.columns.str.strip()
# Create a pivot table with Countries as rows, Month as columns, and Sales as values
sales_pivot = df.pivot_table(index='Country', columns='Month Name', values='Sales', aggfunc='sum')

# Create a pivot table with Countries as rows, Month as columns, and Profit as values
profit_pivot = df.pivot_table(index='Country', columns='Month Name', values='Profit', aggfunc='sum')

# Create heatmap for Sales pivot table
fig4 = px.imshow(sales_pivot, title='Monthly Sales Heatmap')

fig5 = px.imshow(profit_pivot, title='Monthly Profit Heatmap')
# Create bar plot for Sales by Country
fig1 = px.bar(df, x="Country", y="Sales", title="Sales by Country")

# Create bar plot for Profits by Country
fig2 = px.bar(df, x="Country", y="Profit", title="Profits by Country")
df['Profit Margin'] = (df['Profit'] / df['Sales']) * 100

# Create bar plot for Profit Margin by Country
fig6 = px.bar(df, x='Country', y='Profit Margin', title='Profit Margin by Country')
# Create a line plot for Sales and Profits over time
df['Date'] = pd.to_datetime(df['Date'])
df.sort_values('Date', inplace=True)
fig3 = px.line(df, x="Date", y="Sales", title="Sales Over Time")
fig3.add_scatter(x=df["Date"], y=df["Profit"], mode='lines', name='Profit')
df['Date'] = pd.to_datetime(df['Date'])
# monthly_sales = df.groupby(pd.Grouper(key='Date', freq='M')).sum()['Sales']
# res = sm.tsa.seasonal_decompose(monthly_sales)
# fig7 = px.line(res.trend.reset_index(), x='Date', y='Sales', title='Time-Series Decomposition')


G = nx.from_pandas_edgelist(df, 'Segment', 'Product')

# Create elements for Dash Cytoscape
elements = [{'data': {'id': node}} for node in G.nodes()] + \
           [{'data': {'source': edge[0], 'target': edge[1]}} for edge in G.edges()]

# Create a Cytoscape graph
cyto_graph = cyto.Cytoscape(
    id='cytoscape',
    elements=elements,
    layout={'name': 'circle'}
)

# Choropleth map to visualize total sales by country
country_codes = pd.read_csv('country_codes.csv')
df_country = df.merge(country_codes, how='left', left_on='Country', right_on='Country')
df_country_sales = df_country.groupby('Code')['Sales'].sum().reset_index()
choropleth_map = px.choropleth(df_country_sales, locations='Code', color='Sales',
                               title='Geospatial Analysis - Sales by Country',
                               color_continuous_scale='Reds',
                               scope='world')# Initiate the app
app = dash.Dash(__name__)

# Create the layout
app.layout = html.Div([
    dcc.Graph(figure=fig1),
    dcc.Graph(figure=fig2),
    dcc.Graph(figure=fig3),
    dcc.Graph(figure=fig4),
    dcc.Graph(figure=fig5),
    dcc.Graph(figure=fig6),
    dcc.Graph(figure=choropleth_map),
    cyto_graph
    # html.Iframe(srcDoc=open('network.html').read(), width='100%', height='600px'),
    # dcc.Graph(figure=fig8)
])

# Run the app
if __name__ == '__main__':
    app.run_server(debug=True)
