import os
import time
import pandas as pd

def find_mt5_files_folder():
    base = os.path.expanduser("~")
    terminal_root = os.path.join(base, "AppData", "Roaming", "MetaQuotes", "Terminal","D0E8209F77C8CF37AD8BF550E51FF075")
    for root, dirs, _ in os.walk(terminal_root):
        if "MQL5" in dirs:
            path = os.path.join(root, "MQL5", "Files")
            if os.path.isdir(path):
                return path
    raise FileNotFoundError("Could not find MT5 MQL5\\Files folder")

FILES_FOLDER = find_mt5_files_folder()
TRIGGER_FILE = os.path.join(FILES_FOLDER, "run_positionhistory.trigger")
OUTPUT_FILE = os.path.join(FILES_FOLDER, "history_output.csv")

def trigger_export():
    with open(TRIGGER_FILE, "w") as f:
        f.write("run")

def wait_for_output(timeout=60):
    for _ in range(timeout * 10):
        if os.path.exists(OUTPUT_FILE):
            return True
        time.sleep(0.1)
    return False

def read_history_csv():
    df = pd.read_csv(OUTPUT_FILE)
    os.remove(OUTPUT_FILE)
    return df

def run_history_export():
    trigger_export()
    print("Triggered MT5 script. Waiting for output...")

    if wait_for_output():
        print("CSV output detected.")
        return read_history_csv()
    else:
        raise TimeoutError("CSV file not found within timeout.")

# Example usage
if __name__ == "__main__":
    df = run_history_export()
    print(df.head())
