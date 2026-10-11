import os
import json
from datetime import datetime

# 1. SET YOUR HARDCODED EXPIRATION DEADLINE HERE (Currently: Nov 15, 2026)
EXPIRY_DATE = datetime(2026, 11, 11)  

TRACKING_FILE = "system_log.dat"

def verify_trial_status():
    """
    Validates the trial package status.
    Returns (True, message, days_left) if valid, or (False, error_message, 0) if expired.
    """
    current_time = datetime.now()
    
    # Calculate days remaining
    days_left = (EXPIRY_DATE - current_time).days
    
    # Check A: Has the hard deadline calendar date passed?
    if current_time > EXPIRY_DATE or days_left < 0:
        return False, "❌ Your 1-month evaluation trial period has expired. Please contact the administrator for a license.", 0

    # Check B: Anti-Clock Tampering Check
    if os.path.exists(TRACKING_FILE):
        try:
            with open(TRACKING_FILE, "r") as f:
                data = json.load(f)
            
            last_run_str = data.get("last_run", "")
            if last_run_str:
                last_run_time = datetime.strptime(last_run_str, "%Y-%m-%d %H:%M:%S")
                
                # If current system time is BEHIND the last run time, the user rolled back their clock
                if current_time < last_run_time:
                    return False, "⚠️ Security Exception: System clock manipulation detected. Application locked.", 0
        except Exception:
            return False, "❌ Critical Security File Corrupted. Application execution halted.", 0

    # Update the tracking file with the new latest time
    try:
        with open(TRACKING_FILE, "w") as f:
            json.dump({"last_run": current_time.strftime("%Y-%m-%d %H:%M:%S")}, f)
    except Exception:
        pass
        
    return True, "Active", days_left
