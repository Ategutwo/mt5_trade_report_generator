import MetaTrader5 as mt5
import pandas as pd
from datetime import datetime, timedelta
def connect():
    if not mt5.initialize():
        raise RuntimeError("Failed to initialize MT5")

def disconnect():
    mt5.shutdown()

def fetch_orders(start_date, end_date):
     # Get current date (without time component)
    current_date = datetime.now().date()
    
    # Convert end_date to date object if it isn't already
    if isinstance(end_date, datetime):
        end_date_date = end_date.date()
    else:
        end_date_date = end_date
    
    # If end_date is not today, add one day
    if end_date_date < current_date:
        end_date = end_date + timedelta(days=1)
    orders = mt5.history_orders_get(start_date, end_date)
    if orders is None:
        raise RuntimeError("No orders retrieved")

    return pd.DataFrame(orders._asdict() for orders in orders)

def fetch_deals(start_date, end_date):
    # Get current date (without time component)
    current_date = datetime.now().date()
    
    # Convert end_date to date object if it isn't already
    if isinstance(end_date, datetime):
        end_date_date = end_date.date()
    else:
        end_date_date = end_date
    
    # If end_date is not today, add one day
    if end_date_date < current_date:
        end_date = end_date + timedelta(days=1)
    deals = mt5.history_deals_get(start_date, end_date)
    if deals is None:
        raise RuntimeError("No deals retrieved")

    return pd.DataFrame(deals._asdict() for deals in deals)
