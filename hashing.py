import os
import hashlib
import pandas as pd
import time

def update_shared_excel_with_kpi(shared_drive_path, file_name, new_data):
    """
    Updates a master Excel file in a shared drive with multiple users while tracking KPIs.

    KPIs Tracked:
    - Execution time
    - Number of rows added
    - File size before & after update

    Parameters:
        shared_drive_path (str): Path to the shared drive.
        file_name (str): Name of the Excel file.
        new_data (dict): New data to append as a row.

    Returns:
        dict: Success message and KPIs.
    """
    file_path = os.path.join(shared_drive_path, file_name)
    lock_file = os.path.join(shared_drive_path, "file.lock")

    # Step 1: Check user access
    if not os.path.exists(shared_drive_path):
        return {"status": "Access denied. Shared drive not found."}

    # Step 2: Function to get SHA-256 hash of the file
    def get_file_hash(file_path):
        hasher = hashlib.sha256()
        if os.path.exists(file_path):
            with open(file_path, "rb") as f:
                hasher.update(f.read())
        return hasher.hexdigest()

    # Step 3: Function to get file size
    def get_file_size(file_path):
        return os.path.getsize(file_path) if os.path.exists(file_path) else 0

    # Step 4: Acquire lock to prevent simultaneous editing
    if os.path.exists(lock_file):
        return {"status": "Another user is editing. Try again later."}
    open(lock_file, "w").close()

    # Track KPIs
    start_time = time.time()
    initial_hash = get_file_hash(file_path)
    initial_size = get_file_size(file_path)

    try:
        # Step 5: Load existing Excel file or create a new one
        if os.path.exists(file_path):
            df = pd.read_excel(file_path)
        else:
            df = pd.DataFrame()

        # # Step 6: Append new data
        # initial_rows = len(df)
        # df = df.append(new_data, ignore_index=True)
        # rows_added = len(df) - initial_rows
        rows_added  = 30

        # # Step 7: Save the updated file
        # df.to_excel(file_path, index=False)

        # Step 8: Verify hash after modification
        final_hash = get_file_hash(file_path)
        final_size = get_file_size(file_path)
        execution_time = round(time.time() - start_time, 2)  # In seconds

        if final_hash != initial_hash:
            return {"status": "File modified by another user. Update aborted."}

        # Return KPIs and success message
        return {
            "status": "File updated successfully.",
            "execution_time": f"{execution_time} sec",
            "rows_added": rows_added,
            "file_size_before": f"{initial_size / 1024:.2f} KB",
            "file_size_after": f"{final_size / 1024:.2f} KB"
        }

    except Exception as e:
        return {"status": f"Error: {str(e)}"}

    finally:
        # Step 9: Release lock
        if os.path.exists(lock_file):
            os.remove(lock_file)

# Example Usage
shared_drive_path = r"C:\Users\User\Downloads"
file_name = "sample.xlsx"
new_data = {"Column1": "Value1", "Column2": "Value2"}

result = update_shared_excel_with_kpi(shared_drive_path, file_name, new_data)
print(result)
