import pandas as pd
from sdv_manager import SDVManager
import os

# Setup
manager = SDVManager()
files = os.listdir('verified')
stat_files = [os.path.join('verified', f) for f in files if "CrfStatusHistory" in f and f.endswith(".xlsx")]

if not stat_files:
    print("No status file found")
    exit(1)

stat_file = stat_files[0]
print(f"Loading: {stat_file}")
manager.load_crf_status_file(stat_file)

# Test Performance
print("\n--- Testing CRA Performance (All) ---")
perf = manager.get_cra_performance()
if not perf.empty:
    print(perf.head(20).to_string())
    print(f"\nTotal records: {len(perf)}")
else:
    print("No performance data found.")

# Test with filter
if not perf.empty:
    user = perf['User'].iloc[0]
    date = perf['Date'].iloc[0]
    print(f"\n--- Testing filter: User={user}, Date={date} ---")
    filtered = manager.get_cra_performance(start_date=date, end_date=date, user_filter=user)
    print(filtered.to_string())
