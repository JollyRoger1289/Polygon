from polygon import RESTClient
import datetime as dt
import pandas as pd
import plotly.graph_objects as go
from plotly.offline import plot

client = RESTClient("bWiIsbqK3qNWQoT1r7SbofNZDIQjpImu")

stockTicker = 'SPY'

# daily bars
dataRequest = client.get_aggs(ticker = stockTicker,
                              multiplier = 2,
                              timespan = 'minute',
                              from_ = '2023-11-01',
                              to = '2023-12-31',
                              limit = 50000)

# list of polygon agg objects to DataFrame
priceData = pd.DataFrame(dataRequest)

priceData['Date'] = priceData['timestamp'].apply(
                          lambda x: pd.to_datetime(x*1000000))

priceData = priceData.set_index('Date')

#generate plotly figure
fig = go.Figure(data=[go.Candlestick(x=priceData.index,
                open=priceData['open'],
                high=priceData['high'],
                low=priceData['low'],
                close=priceData['close'])])

#open figure in browser
plot(fig, auto_open=True)

file_name = 'MarksDataDec23.xlsx'

priceData.to_excel(file_name)
print('DataFrame is written to Excel File successfully.')

