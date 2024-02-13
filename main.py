import MetaTrader5 as mt5
from datetime import datetime
import pytz
import pandas as pd

name = "GOLD#"

# display data on the MetaTrader 5 package
print("MetaTrader5 package author: ",mt5.__author__)
print("MetaTrader5 package version: ",mt5.__version__)
print()
# establish connection to the MetaTrader 5 terminal
if not mt5.initialize():
    print("initialize() failed, error code =",mt5.last_error())
    quit()

timezone = pytz.timezone("Asia/Bangkok")
utc_from = datetime(2024, 2, 10, tzinfo=timezone)
rates = mt5.copy_rates_from(name, mt5.TIMEFRAME_D1, utc_from, 60)
rates_frame = pd.DataFrame(rates)
# convert time in seconds into the datetime format
rates_frame['time']=pd.to_datetime(rates_frame['time'], unit='s').dt.date
rates_frame = rates_frame.reset_index(drop=True)
rates_frame.sort_values(by='time', ascending=True)

rates_frame = rates_frame.iloc[::-1]
name = rates_frame.head(1)['time'].to_string(index=False)
rates_frame.to_excel(f"{name}.xlsx", sheet_name='Sheet_name_1', index=False) 
 
# shut down connection to the MetaTrader 5 terminal
mt5.shutdown()