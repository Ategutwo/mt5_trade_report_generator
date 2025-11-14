import pandas as pd
import MetaTrader5 as mt5
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
from data_fetcher import connect, disconnect, fetch_orders, fetch_deals
from typing import Optional,Dict , List
import re
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
import numpy as np

def safe_float(value):
    """
    Ultimate safe float converter that handles all MT5/pandas edge cases
    Returns: float or None (never NaN)
    """
    # Handle all empty/None cases first
    if value is None or value == '':
        return None
    
    # Handle pandas/numpy missing values
    try:
        if pd.isna(value):
            return None
    except:
        pass
    
    # Handle zero values (common MT5 default)
    try:
        if float(value) == 0:
            return None
    except:
        pass
    
    # Final conversion attempt
    try:
        return float(value)
    except (ValueError, TypeError):
        return None

def parse_sl_tp_from_comment(comment):
    """Enhanced comment parser that handles more formats"""
    if not comment or not isinstance(comment, str):
        return None, None
    
    # Clean the comment string
    comment = (comment.lower()
               .replace('[', '').replace(']', '')
               .replace('(', '').replace(')', '')
               .strip())
    
    sl, tp = None, None
    
    # More comprehensive pattern matching
    patterns = [
        (r'sl[:=]?\s*(\d+\.?\d*)', 'sl'),  # sl:1.1234 or sl=1.1234
        (r'stop[:=]?\s*(\d+\.?\d*)', 'sl'), # stop:1.1234
        (r'tp[:=]?\s*(\d+\.?\d*)', 'tp'),   # tp:1.1234
        (r'take[:=]?\s*(\d+\.?\d*)', 'tp'), # take:1.1234
        (r'\/sl\s*(\d+\.?\d*)', 'sl'),      # /sl 1.1234
        (r'\/tp\s*(\d+\.?\d*)', 'tp')       # /tp 1.1234
    ]
    
    for pattern, target in patterns:
        matches = re.findall(pattern, comment)
        if matches:
            try:
                value = float(matches[0])
                if target == 'sl':
                    sl = value
                else:
                    tp = value
            except (ValueError, IndexError):
                continue
    
    return sl, tp

def get_sl_tp_values(position_id, orders_df, deals_df):
    """Complete SL/TP detection checking all possible sources"""
    sl, tp = None, None
    
    # 1. First check orders (both entry and modifications)
    position_orders = orders_df[orders_df['position_id'] == position_id]
    
    if not position_orders.empty:
        # Check entry orders first
        entry_orders = position_orders[position_orders['type'].isin([0, 1])]
        if not entry_orders.empty:
            entry_order = entry_orders.iloc[0]
            sl = safe_float(entry_order.get('sl'))
            tp = safe_float(entry_order.get('tp'))
            
            # Parse from order comment if needed
            if sl is None or tp is None:
                order_comment_sl, order_comment_tp = parse_sl_tp_from_comment(
                    str(entry_order.get('comment', ''))
                )
                sl = order_comment_sl if sl is None else sl
                tp = order_comment_tp if tp is None else tp
        
        # Check modification orders
        mod_orders = position_orders[position_orders['type'].isin([2, 3, 4])]
        for _, order in mod_orders.iterrows():
            current_sl = safe_float(order.get('sl'))
            current_tp = safe_float(order.get('tp'))
            if current_sl is not None:
                sl = current_sl
            if current_tp is not None:
                tp = current_tp
    
    # 2. Check deals (both entry and exit)
    position_deals = deals_df[deals_df['position_id'] == position_id]
    
    for _, deal in position_deals.iterrows():
        # Check deal execution prices (SL/TP triggers)
        if deal['entry'] == 2:  # SL triggered
            sl = safe_float(deal.get('price'))
        elif deal['entry'] == 3:  # TP triggered
            tp = safe_float(deal.get('price'))
        
        # Parse from deal comments as final fallback
        if sl is None or tp is None:
            deal_comment_sl, deal_comment_tp = parse_sl_tp_from_comment(str(deal.get('comment', '')))
            sl = deal_comment_sl if sl is None else sl
            tp = deal_comment_tp if tp is None else tp
    
    return sl, tp

def get_conversion_rate(symbol, deposit_currency="USD"):
    """Safe currency conversion rate calculation"""
    if symbol.endswith(deposit_currency):
        return 1.0
    
    quote_currency = symbol[-3:]
    if quote_currency == deposit_currency:
        return 1.0
    
    # Try direct conversion
    conversion_symbol = f"{quote_currency}{deposit_currency}"
    tick = mt5.symbol_info_tick(conversion_symbol)
    if tick and safe_float(tick.bid):
        return safe_float(tick.bid)
    
    # Try inverse conversion
    inverse_symbol = f"{deposit_currency}{quote_currency}"
    tick = mt5.symbol_info_tick(inverse_symbol)
    if tick and safe_float(tick.bid):
        return 1.0 / safe_float(tick.bid)
    
    # Fallback to 1:1 if conversion fails
    return 1.0

def extract_channel_name(comment: str) -> Optional[str]:
    """
    Extract multi-word channel names while removing all SL/TP patterns
    Returns the longest remaining text segment with words >3 characters
    """
    if not comment or not isinstance(comment, str):
        return None

    # Remove ALL SL/TP patterns brokers use (case insensitive)
    cleaned = re.sub(
        r'(sl|tp|stop[-\s]?loss|take[-\s]?profit|stoploss|takeprofit)[:=]?\s*[\d\.]*|'
        r'\/sl\b|\/tp\b|'
        r'\[sl\]|\[tp\]|'
        r'\b\d+\.\d+\b|'  # Standalone numbers like 1.2345
        r'\b\d{4,}\b',    # Trade numbers like 6409
        ' ', 
        comment, 
        flags=re.IGNORECASE
    )

    # Remove special chars but preserve spaces between words
    cleaned = re.sub(r'[^\w\s]', ' ', cleaned)
    
    # Find all potential channel segments
    segments = []
    current_segment = []
    
    for word in cleaned.split():
        if len(word) > 0:
            current_segment.append(word)
        else:
            if current_segment:
                segments.append(' '.join(current_segment))
                current_segment = []
    
    if current_segment:  # Add the last segment
        segments.append(' '.join(current_segment))
    
    # Return the longest remaining valid segment
    if segments:
        return max(segments, key=len)
    return None




def get_channel_for_position(position_id: int, orders_df: pd.DataFrame, deals_df: pd.DataFrame) -> Optional[str]:
    """
    Checks all related orders and deals for a position to find channel mentions.
    Returns the first valid channel name found, or None.
    """
    # Check entry order comment
    entry_orders = orders_df[
        (orders_df['position_id'] == position_id) & 
        (orders_df['type'].isin([0, 1]))
    ]
    
    if not entry_orders.empty:
        channel = extract_channel_name(entry_orders.iloc[0].get('comment'))
        if channel:
            return channel
    
    # Check modification orders
    mod_orders = orders_df[
        (orders_df['position_id'] == position_id) & 
        (orders_df['type'].isin([2, 3, 4]))
    ]
    
    for _, order in mod_orders.iterrows():
        channel = extract_channel_name(order.get('comment'))
        if channel:
            return channel
    
    # Check all related deals
    position_deals = deals_df[deals_df['position_id'] == position_id]
    for _, deal in position_deals.iterrows():
        channel = extract_channel_name(deal.get('comment'))
        if channel:
            return channel
    
    return None

def analyze_trades_by_channel(df: pd.DataFrame) -> pd.DataFrame:
    """
    Groups trades by channel name, magic id, and symbol/pair with detailed trade comments
    """
    if df.empty:
        return pd.DataFrame()
    
    # Create copies of columns we'll modify
    df = df.copy()
    df['Channel name'] = df['Comment'].apply(lambda x: str(x).split('|')[0].strip() if pd.notna(x) and '|' in str(x) else (str(x) if pd.notna(x) else ''))
    df['Channel name'] = df['Channel name'].replace('', 'Unknown Channel')
    df['magic id'] = df['Magic'].astype(str).str.strip()
    df['pairs'] = df['Symbol']
    
    # Ensure datetime columns are properly formatted
    df['Open Time'] = pd.to_datetime(df['Open Time'])
    df['Close Time'] = pd.to_datetime(df['Close Time'])
    
    # Calculate points - handle JPY pairs differently
    def calculate_points(row):
        if 'JPY' in row['pairs']:
            multiplier = 10**2  # For JPY pairs
        else:
            multiplier = 10**4  # For other pairs
            
        points = (row['Close Price'] - row['Open Price']) * multiplier
        return points * (-1 if row['Type'] == 'Sell' else 1)
    
    df['points'] = df.apply(calculate_points, axis=1)
    
    # Determine trade result using np.select()
    conditions = [
        (df['Profit'] < -0.01),
        (df['Profit'].abs() <= 0.01),  # Breakeven
        (df['Profit'] > 0.01)
    ]
    choices = ['lose', 'breakeven', 'win']
    df['Trade Result'] = np.select(conditions, choices, default='unknown')
    
    # Group and aggregate
    grouped = df.groupby(['Channel name', 'magic id', 'pairs'])
    
    def format_trade_details(trades):
        details = []
        for _, trade in trades.iterrows():
            details.append(
                f"({trade['Position']}) {trade['Open Time'].strftime('%Y/%m/%d (%H:%M:%S)')} - "
                f"{trade['Close Time'].strftime('%Y/%m/%d (%H:%M:%S)')} "
                f"1 {trade['Trade Result']} = {trade['Profit']:.2f}\n"
                f"({trade['Volume']:.2f}) {trade['Type'].lower()} ({int(abs(trade['points']))}points)"
            )
        return "\n\n".join(details)
    
    analysis = grouped.agg({
        'Profit': [
            ('total_win_amount', lambda x: x[x > 0.01].sum()),
            ('total_lose_amount', lambda x: x[x < -0.01].sum()),
            ('total_breakeven_amount', lambda x: x[(x >= -0.01) & (x <= 0.01)].sum()),
            ('win_details', lambda x: format_trade_details(df.loc[x.index][df.loc[x.index]['Trade Result'] == 'win'])),
            ('lose_details', lambda x: format_trade_details(df.loc[x.index][df.loc[x.index]['Trade Result'] == 'lose']))
        ],
        'Trade Result': [
            ('win_count', lambda x: (x == 'win').sum()),
            ('lose_count', lambda x: (x == 'lose').sum()),
            ('breakeven_count', lambda x: (x == 'breakeven').sum())
        ],
        'Swap': [('total_swap', 'sum')],
        'Commission': [('total_commission', 'sum')],
        'Fee': [('total_fee', 'sum')],
        'Volume': ['first']  # We need this for lot size in comments
    }).reset_index()
    
    # Flatten multi-index columns
    analysis.columns = ['_'.join(col).strip('_') for col in analysis.columns.values]
    
    # Calculate final metrics with proper NaN handling
    analysis['total net amount'] = (
        analysis['Profit_total_win_amount'].fillna(0) +
        analysis['Profit_total_lose_amount'].fillna(0) +
        analysis['Swap_total_swap'].fillna(0) +
        analysis['Commission_total_commission'].fillna(0) +
        analysis['Fee_total_fee'].fillna(0)
    )
    
    # Format output strings
    analysis['total win'] = (
        analysis['Trade Result_win_count'].astype(str) + ' win' +
        analysis['Profit_win_details'].apply(lambda x: f"\n{x}" if pd.notna(x) and x != '' else '')
    )
    
    analysis['total lose'] = (
        analysis['Trade Result_lose_count'].astype(str) + ' lose' +
        analysis['Profit_lose_details'].apply(lambda x: f"\n{x}" if pd.notna(x) and x != '' else '')
    )
    
    analysis['total 0.00(Breakeven)'] = analysis['Trade Result_breakeven_count'].astype(str) + ' breakeven'
    
    analysis['total win amount'] = (
        analysis['Profit_total_win_amount'].round(2).astype(str) +
        analysis['Profit_win_details'].apply(lambda x: f"\n{x}" if pd.notna(x) and x != '' else '')
    )
    
    analysis['total lose amount'] = (
        analysis['Profit_total_lose_amount'].round(2).astype(str) +
        analysis['Profit_lose_details'].apply(lambda x: f"\n{x}" if pd.notna(x) and x != '' else '')
    )
    
    # Select and rename final columns
    result = analysis[[
        'Channel name', 'magic id', 'pairs', 'total win', 'total lose',
        'total 0.00(Breakeven)', 'total win amount', 'total lose amount',
        'Swap_total_swap', 'Commission_total_commission', 'Fee_total_fee',
        'total net amount'
    ]].rename(columns={
        'Swap_total_swap': 'total swap',
        'Commission_total_commission': 'total commissions',
        'Fee_total_fee': 'total fee'
    })
    
    result['remark'] = ''
    
    return result

from openpyxl.comments import Comment
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

def add_channel_analysis_sheet(wb: Workbook, analysis_df: pd.DataFrame) -> Workbook:
    """Adds the channel analysis sheet with modern Excel comments"""
    # Create the worksheet
    analysis_ws = wb.create_sheet("Channel Analysis")
    
    # Write header
    header = [
        'Channel name', 'magic id', 'pairs', 'total win', 'total lose',
        'total 0.00(Breakeven)', 'total win amount', 'total lose amount',
        'total swap', 'total commissions', 'remark', 'total net amount'
    ]
    analysis_ws.append(header)
    
    # Format header
    for cell in analysis_ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    
    # Write data and add comments
    for _, row in analysis_df.iterrows():
        # Create display values without notes
        display_values = [
            row['Channel name'],
            row['magic id'],
            row['pairs'],
            row['total win'].split('\n')[0] if '\n' in str(row['total win']) else row['total win'],
            row['total lose'].split('\n')[0] if '\n' in str(row['total lose']) else row['total lose'],
            row['total 0.00(Breakeven)'],
            row['total win amount'].split('\n')[0] if '\n' in str(row['total win amount']) else row['total win amount'],
            row['total lose amount'].split('\n')[0] if '\n' in str(row['total lose amount']) else row['total lose amount'],
            row['total swap'],
            row['total commissions'],
            row['remark'],
            row['total net amount']
        ]
        analysis_ws.append(display_values)
        
        current_row = analysis_ws.max_row
        
        # Add comments where needed (only if not already exists)
        for col_idx, col_name in [(4, 'total win'), (5, 'total lose'), 
                                (7, 'total win amount'), (8, 'total lose amount')]:
            cell_value = str(row[col_name])
            if '\n' in cell_value:
                main_text, comment_text = cell_value.split('\n', 1)
                cell = analysis_ws.cell(row=current_row, column=col_idx)
                
                # Only add comment if one doesn't already exist
                if not cell.comment:
                    comment = Comment(
                        text=comment_text,
                        author="Trade Details",
                        width=400,  # Width in pixels
                        height=300  # Height in pixels
                    )
                    cell.comment = comment
    
    # Set column widths
    column_widths = {
        'A': 25,  # Channel name
        'B': 12,  # magic id
        'C': 10,  # pairs
        'D': 15,  # total win
        'E': 15,  # total lose
        'F': 20,  # total 0.00(Breakeven)
        'G': 18,  # total win amount
        'H': 18,  # total lose amount
        'I': 12,  # total swap
        'J': 15,  # total commissions
        'K': 15,  # remark
        'L': 15   # total net amount
    }
    
    for col_letter, width in column_widths.items():
        analysis_ws.column_dimensions[col_letter].width = width
    
    # Freeze header row
    analysis_ws.freeze_panes = 'A2'
    
    # Add filters
    analysis_ws.auto_filter.ref = f"A1:{get_column_letter(analysis_ws.max_column)}{1}"
    
    return wb

def generate_excel_report(start_date, end_date, output_path):
    """Enhanced report generator with channel detection"""
    try:
        # Connect and fetch data
        connect()
        orders_df = fetch_orders(start_date, end_date)
        deals_df = fetch_deals(start_date, end_date)
        disconnect()
        print(f"Total deals fetched: {len(deals_df)}")
        print(f"Unique position IDs: {len(deals_df['position_id'].unique())}")
        print(f"Entry=0 deals: {len(deals_df[deals_df['entry'] == 0])}")
        print(f"Entry=1 deals: {len(deals_df[deals_df['entry'] == 1])}")   
        # Convert numeric columns safely
        numeric_cols = ['sl', 'tp', 'price', 'volume_initial', 'profit', 'commission', 'swap']
        for col in numeric_cols:
            if col in orders_df:
                orders_df[col] = orders_df[col].apply(safe_float)
            if col in deals_df:
                deals_df[col] = deals_df[col].apply(safe_float)

        # Process positions
        print(f"Processing position {deals_df['entry']}")
        in_deals = deals_df[deals_df['entry'] == 0].copy()
        out_deals = deals_df[deals_df['entry'] == 1].copy()
        final_rows = []

        for position_id in in_deals['position_id'].unique():
            try:
                # Get position deals
                pos_in_deals = in_deals[in_deals['position_id'] == position_id]
                pos_out_deals = out_deals[out_deals['position_id'] == position_id]
                
                if pos_in_deals.empty or pos_out_deals.empty:
                    continue
                
                deal_in = pos_in_deals.iloc[0]
                deal_out = pos_out_deals.iloc[-1]

                # Get SL/TP
                sl, tp = get_sl_tp_values(position_id, orders_df, deals_df)

                # Get entry order
                entry_orders = orders_df[
                    (orders_df['position_id'] == position_id) & 
                    (orders_df['type'].isin([0, 1]))
                ]
                
                if entry_orders.empty:
                    continue
                    
                entry_order = entry_orders.iloc[0]
                symbol = entry_order['symbol']
                
                # Calculate position metrics
                symbol_info = mt5.symbol_info(symbol)
                contract_size = safe_float(symbol_info.trade_contract_size) if symbol_info else 100000
                conversion_rate = get_conversion_rate(symbol)
                
                position_deals = deals_df[deals_df['position_id'] == position_id]
                total_profit = position_deals['profit'].sum()
                total_commission = position_deals['commission'].sum()
                total_swap = position_deals['swap'].sum()
                
                open_price = safe_float(deal_in['price'])
                close_price = safe_float(deal_out['price'])
                volume = safe_float(entry_order['volume_initial'])
                
                if None in (open_price, close_price, volume, contract_size, conversion_rate):
                    continue
                
                value = close_price * volume * contract_size * conversion_rate
                direction = 1 if deal_in['type'] == 0 else -1
                change_pct = ((close_price - open_price) / open_price) * direction * 100

             
                
                # Find channel in all related records
                channel = get_channel_for_position(position_id, orders_df, deals_df)
                enhanced_comment = ""
                if channel:
                    enhanced_comment =  f"{channel}"

                # Add to final report
                final_rows.append({
                    'Open Time': pd.to_datetime(entry_order['time_setup'], unit='s'),
                    'Symbol': symbol,
                    'Position': int(position_id),
                    'Type': 'Buy' if deal_in['type'] == 0 else 'Sell',
                    'Volume': volume,
                    'Open Price': open_price,
                    'Stop Loss': sl,
                    'Take Profit': tp,
                    'Close Time': pd.to_datetime(deal_out['time'], unit='s'),
                    'Close Price': close_price,
                    'Value': round(value, 2),
                    'Commission': total_commission,
                    'Fee': 0.0,
                    'Swap': total_swap,
                    'Profit': total_profit,
                    'Change %': f"{round(change_pct, 2)}%",
                    'Magic': str(entry_order.get('magic', '')),
                    'Comment': enhanced_comment  # Now includes channel info if found
                })

            except Exception as e:
                print(f"Skipping position {position_id} due to error: {str(e)}")
                continue
            
        #  Create Excel report
        if not final_rows:
            raise ValueError("No valid trade data found for the selected period")

        df = pd.DataFrame(final_rows)
        
        # Ensure we have all expected columns
        expected_columns = [
            'Open Time', 'Symbol', 'Position', 'Type', 'Volume', 'Open Price',
            'Stop Loss', 'Take Profit', 'Close Time', 'Close Price', 'Value',
            'Commission', 'Fee', 'Swap', 'Profit', 'Change %', 'Magic', 'Comment'
        ]
        
        for col in expected_columns:
            if col not in df.columns:
                df[col] = None
                
        df = df[expected_columns]

        # Write to Excel
        wb = Workbook()
        ws = wb.active
        ws.title = "MT5 Trade Report"

        # Write header
        ws.append(expected_columns)
        
        # Write data
        for _, row in df.iterrows():
            ws.append([row[col] for col in expected_columns])

        # Format header
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')

        # Auto-adjust columns
        for col in ws.columns:
            max_length = max(
                len(str(cell.value)) if cell.value is not None else 0
                for cell in col
            )
            ws.column_dimensions[col[0].column_letter].width = max_length + 2


            # In your generate_excel_report function, after creating df:

        # First create the analysis DataFrame
        analysis_df = analyze_trades_by_channel(df)

        # Then add the sheet to the workbook
        add_channel_analysis_sheet(wb, analysis_df)

        # Save the workbook
        wb.save(output_path)
    
        wb.save(output_path)
        print(f"Successfully generated report with {len(final_rows)} trades at {output_path}")

    except Exception as e:
        raise RuntimeError(f"Report generation failed: {str(e)}")