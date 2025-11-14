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
from fetch_history import run_history_export
from openpyxl.comments import Comment
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
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
        print(row['pairs'])
        symbol = row['pairs'].upper()

        # Define multipliers for points (pipettes)
        if 'JPY' in symbol:
            multiplier = 1000  # JPY pairs: 2 decimal places → 3 decimal points
        elif 'XAU' in symbol or 'GOLD' in symbol:
            multiplier = 100   # Gold: 1-2 decimal places → 2-3 decimal points
        elif 'BTC' in symbol or 'ETH' in symbol:
            multiplier = 100   # Cryptos: 0-2 decimal places → 2-3 decimal points
        else:
            multiplier = 100000  # Most forex: 4-5 decimal places → 5-6 decimal points

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
    """Enhanced report generator with channel detection and error filtering"""
    try:
        # Connect and fetch data
        connect()
        
        df = run_history_export()
        df['Open Time'] = pd.to_datetime(df['Open Time'])
        df['Close Time'] = pd.to_datetime(df['Close Time'])

        # Filter by date range if specified
        if start_date and end_date:
            mask = (df['Close Time'] >= pd.to_datetime(start_date)) & \
                   (df['Close Time'] <= pd.to_datetime(end_date))
            df = df.loc[mask]
        
        # Identify trades with SL/TP exactly equal to 1 OR NaN in Type column
        def has_invalid_trade(row):
            """Check if trade has SL/TP exactly equal to 1 OR NaN in Type"""
            try:
                # Check for SL/TP = 1
                sl = safe_float(row.get('Stop Loss'))
                tp = safe_float(row.get('Take Profit'))
                
                if sl == 1.0 or tp == 1.0:
                    return True, "SL/TP=1"
                
                # Check for NaN in Type column
                if pd.isna(row.get('Symbol')):
                    return True, "Missing Type"
                
                return False, None
            except:
                return False, None
        
        # Split data into valid trades and error trades
        error_results = df.apply(has_invalid_trade, axis=1, result_type='expand')
        error_mask = error_results[0]
        error_reasons = error_results[1]
        
        error_df = df[error_mask].copy()
        valid_df = df[~error_mask].copy()
        
        # Add reason to error_df
        error_df['Filter Reason'] = error_reasons[error_mask]
        
        # Extract channel names only for valid trades
        valid_df['Comment'] = valid_df['Comment'].apply(extract_channel_name)

        # Write to Excel
        wb = Workbook()
        
        # Create main trade report sheet with valid trades only
        ws = wb.active
        ws.title = "MT5 Trade Report"

        # Ensure we have all expected columns for valid trades
        expected_columns = [
            'Open Time', 'Symbol', 'Position', 'Type', 'Volume', 'Open Price',
            'Stop Loss', 'Take Profit', 'Close Time', 'Close Price', 'Value',
            'Commission', 'Fee', 'Swap', 'Profit', 'Change %', 'Magic', 'Comment'
        ]
        
        for col in expected_columns:
            if col not in valid_df.columns:
                valid_df[col] = None
                
        valid_df = valid_df[expected_columns]

        # Write header
        ws.append(expected_columns)
        
        # Write valid data
        for _, row in valid_df.iterrows():
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

        # Create errors sheet with invalid trades
        if not error_df.empty:
            error_ws = wb.create_sheet("Errors")
            
            error_columns = expected_columns + ['Filter Reason']
            
            # Write error header
            error_ws.append(error_columns)

            # Write error data
            for _, row in error_df.iterrows():
                row_data = [row.get(col) for col in expected_columns]
                row_data.append(row['Filter Reason'])
                error_ws.append(row_data)

            # Format error header
            for cell in error_ws[1]:
                cell.font = Font(bold=True, color="FF0000")
                cell.alignment = Alignment(horizontal='center')

            # Auto-adjust error columns
            for col in error_ws.columns:
                max_length = max(
                    len(str(cell.value)) if cell.value is not None else 0
                    for cell in col
                )
                error_ws.column_dimensions[col[0].column_letter].width = max_length + 2

            # Create error analysis sheet (excluding NaN pairs)
            error_df_for_analysis = error_df[error_df['Symbol'].notna()].copy()
            if not error_df_for_analysis.empty:
                error_df_for_analysis['Comment'] = error_df_for_analysis['Comment'].apply(extract_channel_name)
                error_analysis_df = analyze_trades_by_channel(error_df_for_analysis)
                # Create error analysis sheet with a different name
                error_analysis_ws = wb.create_sheet("Error Analysis")
                add_channel_analysis_sheet_data(error_analysis_ws, error_analysis_df)

        # Create analysis only for valid trades
        analysis_df = analyze_trades_by_channel(valid_df)
        add_channel_analysis_sheet(wb, analysis_df)

        # Save the workbook
        wb.save(output_path)
        print(f"Successfully generated report with {len(valid_df)} valid trades and {len(error_df)} error trades at {output_path}")

    except Exception as e:
        raise RuntimeError(f"Report generation failed: {str(e)}")

def add_channel_analysis_sheet_data(ws, analysis_df):
    """Helper function to add analysis data to any worksheet"""
    # Write header
    header = [
        'Channel name', 'magic id', 'pairs', 'total win', 'total lose',
        'total 0.00(Breakeven)', 'total win amount', 'total lose amount',
        'total swap', 'total commissions', 'remark', 'total net amount'
    ]
    ws.append(header)
    
    # Format header
    for cell in ws[1]:
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
        ws.append(display_values)
        
        current_row = ws.max_row
        
        # Add comments where needed
        for col_idx, col_name in [(4, 'total win'), (5, 'total lose'), 
                                (7, 'total win amount'), (8, 'total lose amount')]:
            cell_value = str(row[col_name])
            if '\n' in cell_value:
                main_text, comment_text = cell_value.split('\n', 1)
                cell = ws.cell(row=current_row, column=col_idx)
                
                if not cell.comment:
                    comment = Comment(
                        text=comment_text,
                        author="Trade Details",
                        width=400,
                        height=300
                    )
                    cell.comment = comment
    
    # Set column widths
    column_widths = {
        'A': 25, 'B': 12, 'C': 10, 'D': 15, 'E': 15, 'F': 20,
        'G': 18, 'H': 18, 'I': 12, 'J': 15, 'K': 15, 'L': 15
    }
    
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width
    
    # Freeze header row
    ws.freeze_panes = 'A2'
    
    # Add filters
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{1}"





import pandas as pd
import MetaTrader5 as mt5
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
from data_fetcher import connect, disconnect, fetch_orders, fetch_deals
from typing import Optional, Dict, List
import re
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
import numpy as np
from fetch_history import run_history_export
from openpyxl.comments import Comment
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import os

# ... (keep all your existing functions: safe_float, parse_sl_tp_from_comment, 
# get_sl_tp_values, get_conversion_rate, extract_channel_name, 
# get_channel_for_position, analyze_trades_by_channel, 
# add_channel_analysis_sheet, add_channel_analysis_sheet_data)

def select_input_file():
    """Open file dialog to select input Excel file"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    file_path = filedialog.askopenfilename(
        title="Select Excel file with trade data",
        filetypes=[
            ("Excel files", "*.xlsx *.xls"),
            ("All files", "*.*")
        ]
    )
    
    root.destroy()
    return file_path

def select_output_file(default_name="analyzed_trades.xlsx"):
    """Open file dialog to select output Excel file"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    file_path = filedialog.asksaveasfilename(
        title="Save analysis report as",
        defaultextension=".xlsx",
        filetypes=[
            ("Excel files", "*.xlsx"),
            ("All files", "*.*")
        ],
        initialfile=default_name
    )
    
    root.destroy()
    return file_path

def get_sheet_name():
    """Dialog to get sheet name from user"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    sheet_name = simpledialog.askstring(
        "Sheet Name", 
        "Enter the name of the sheet containing trade data:",
        initialvalue="MT5 Trade Report"
    )
    
    root.destroy()
    return sheet_name if sheet_name else "data"

def get_date_range():
    """Dialog to get optional date range from user"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    use_dates = messagebox.askyesno(
        "Date Filter", 
        "Do you want to filter by date range?"
    )
    
    if use_dates:
        start_date = simpledialog.askstring(
            "Start Date", 
            "Enter start date (YYYY-MM-DD):",
            initialvalue="2024-01-01"
        )
        
        end_date = simpledialog.askstring(
            "End Date", 
            "Enter end date (YYYY-MM-DD):",
            initialvalue="2024-12-31"
        )
    else:
        start_date = end_date = None
    
    root.destroy()
    return start_date, end_date

def generate_excel_report_from_excel(input_file_path, output_file_path, data_sheet_name="data", start_date=None, end_date=None):
    """
    Generate report from existing Excel file with trade data
    
    Args:
        input_file_path: Path to the input Excel file
        output_file_path: Path where the output Excel file will be saved
        data_sheet_name: Name of the sheet containing the trade data (default: "data")
        start_date: Optional start date for filtering (YYYY-MM-DD)
        end_date: Optional end date for filtering (YYYY-MM-DD)
    """
    try:
        # Read data from Excel file
        print(f"Reading data from {input_file_path}, sheet: {data_sheet_name}")
        df = pd.read_excel(input_file_path, sheet_name=data_sheet_name)
        
        # Ensure column names match your expected format
        expected_columns = [
            'Open Time', 'Symbol', 'Position', 'Type', 'Volume', 'Open Price',
            'Stop Loss', 'Take Profit', 'Close Time', 'Close Price', 'Value',
            'Commission', 'Fee', 'Swap', 'Profit', 'Change %', 'Magic', 'Comment'
        ]
        
        # Check if all expected columns exist
        missing_columns = set(expected_columns) - set(df.columns)
        if missing_columns:
            print(f"Warning: Missing columns in input data: {missing_columns}")
            # Add missing columns with None values
            for col in missing_columns:
                df[col] = None
        
        # Convert datetime columns
        df['Open Time'] = pd.to_datetime(df['Open Time'], errors='coerce')
        df['Close Time'] = pd.to_datetime(df['Close Time'], errors='coerce')

        # Filter by date range if specified
        if start_date and end_date:
            mask = (df['Close Time'] >= pd.to_datetime(start_date)) & \
                   (df['Close Time'] <= pd.to_datetime(end_date))
            df = df.loc[mask]
            print(f"Filtered to {len(df)} trades between {start_date} and {end_date}")
        
        # Identify trades with SL/TP exactly equal to 1 OR NaN in Type column
        def has_invalid_trade(row):
            """Check if trade has SL/TP exactly equal to 1 OR NaN in Type"""
            try:
                # Check for SL/TP = 1
                sl = safe_float(row.get('Stop Loss'))
                tp = safe_float(row.get('Take Profit'))
                
                if sl == 1.0 or tp == 1.0:
                    return True, "SL/TP=1"
                
                # Check for NaN in Type column
                if pd.isna(row.get('Symbol')):
                    return True, "Missing Symbol"
                
                return False, None
            except:
                return False, None
        
        # Split data into valid trades and error trades
        error_results = df.apply(has_invalid_trade, axis=1, result_type='expand')
        error_mask = error_results[0]
        error_reasons = error_results[1]
        
        error_df = df[error_mask].copy()
        valid_df = df[~error_mask].copy()
        
        # Add reason to error_df
        error_df['Filter Reason'] = error_reasons[error_mask]
        
        # Extract channel names only for valid trades
        valid_df['Comment'] = valid_df['Comment'].apply(extract_channel_name)

        # Write to Excel
        wb = Workbook()
        
        # Create main trade report sheet with valid trades only
        ws = wb.active
        ws.title = "MT5 Trade Report"

        # Ensure we have all expected columns for valid trades
        for col in expected_columns:
            if col not in valid_df.columns:
                valid_df[col] = None
                
        valid_df = valid_df[expected_columns]

        # Write header
        ws.append(expected_columns)
        
        # Write valid data
        for _, row in valid_df.iterrows():
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

        # Create errors sheet with invalid trades
        if not error_df.empty:
            error_ws = wb.create_sheet("Errors")
            
            error_columns = expected_columns + ['Filter Reason']
            
            # Write error header
            error_ws.append(error_columns)

            # Write error data
            for _, row in error_df.iterrows():
                row_data = [row.get(col) for col in expected_columns]
                row_data.append(row['Filter Reason'])
                error_ws.append(row_data)

            # Format error header
            for cell in error_ws[1]:
                cell.font = Font(bold=True, color="FF0000")
                cell.alignment = Alignment(horizontal='center')

            # Auto-adjust error columns
            for col in error_ws.columns:
                max_length = max(
                    len(str(cell.value)) if cell.value is not None else 0
                    for cell in col
                )
                error_ws.column_dimensions[col[0].column_letter].width = max_length + 2

            # Create error analysis sheet (excluding NaN pairs)
            error_df_for_analysis = error_df[error_df['Symbol'].notna()].copy()
            if not error_df_for_analysis.empty:
                error_df_for_analysis['Comment'] = error_df_for_analysis['Comment'].apply(extract_channel_name)
                error_analysis_df = analyze_trades_by_channel(error_df_for_analysis)
                # Create error analysis sheet with a different name
                error_analysis_ws = wb.create_sheet("Error Analysis")
                add_channel_analysis_sheet_data(error_analysis_ws, error_analysis_df)

        # Create analysis only for valid trades
        analysis_df = analyze_trades_by_channel(valid_df)
        add_channel_analysis_sheet(wb, analysis_df)

        # Save the workbook
        wb.save(output_file_path)
        print(f"Successfully generated report with {len(valid_df)} valid trades and {len(error_df)} error trades at {output_file_path}")

    except Exception as e:
        raise RuntimeError(f"Report generation from Excel failed: {str(e)}")

def process_excel_workbook_with_gui():
    """
    Main function with GUI dialogs to process Excel workbook
    """
    try:
        # Step 1: Select input file
        print("Please select the input Excel file...")
        input_file_path = select_input_file()
        
        if not input_file_path:
            print("No file selected. Exiting.")
            return None
        
        # Step 2: Get sheet name
        print("Please specify the sheet name...")
        sheet_name = get_sheet_name()
        
        # Step 3: Check if date filtering is needed
        print("Do you want to filter by date range?")
        start_date, end_date = get_date_range()
        
        # Step 4: Select output file
        print("Please select where to save the analysis report...")
        
        # Generate default output name based on input file
        base_name = os.path.splitext(os.path.basename(input_file_path))[0]
        default_output = f"{base_name}_analyzed.xlsx"
        
        output_file_path = select_output_file(default_output)
        
        if not output_file_path:
            print("No output file selected. Exiting.")
            return None
        
        # Step 5: Process the file
        print("Processing... This may take a moment.")
        
        generate_excel_report_from_excel(
            input_file_path=input_file_path,
            output_file_path=output_file_path,
            data_sheet_name=sheet_name,
            start_date=start_date,
            end_date=end_date
        )
        
        # Show success message
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(
            "Success", 
            f"Analysis complete!\n\n"
            f"Output file: {output_file_path}\n"
            f"Data sheet: {sheet_name}\n"
            f"Date range: {start_date} to {end_date if start_date else 'No filter'}"
        )
        root.destroy()
        
        return output_file_path
        
    except Exception as e:
        # Show error message
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Error", f"Processing failed:\n{str(e)}")
        root.destroy()
        return None

def process_excel_workbook(input_file_path=None, output_file_path=None, data_sheet_name="data", start_date=None, end_date=None):
    """
    Programmatic function to process Excel workbook (original functionality preserved)
    """
    # If no parameters provided, use GUI mode
    if input_file_path is None:
        return process_excel_workbook_with_gui()
    
    # If no output path specified, create one based on input
    if output_file_path is None:
        base_name = input_file_path.replace('.xlsx', '').replace('.xls', '')
        output_file_path = f"{base_name}_analyzed.xlsx"
    
    # Generate the report
    generate_excel_report_from_excel(
        input_file_path=input_file_path,
        output_file_path=output_file_path,
        data_sheet_name=data_sheet_name,
        start_date=start_date,
        end_date=end_date
    )
    
    return output_file_path

# Example usage
if __name__ == "__main__":
    print("Excel Trade Analysis Tool")
    print("=" * 50)
    
    # Option 1: Use GUI (recommended for interactive use)
    result_file = process_excel_workbook_with_gui()
    
    if result_file:
        print(f"Analysis complete! Output file: {result_file}")
    else:
        print("Process cancelled or failed.")
    
    # Option 2: Use programmatically (for scripts)
    # result_file = process_excel_workbook(
    #     input_file_path="trades.xlsx",
    #     output_file_path="analysis.xlsx",
    #     data_sheet_name="data",
    #     start_date="2024-01-01",
    #     end_date="2024-12-31"
    # )