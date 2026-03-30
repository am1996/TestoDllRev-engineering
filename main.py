import olefile
import pythoncom
from win32com import storagecon
import struct
import shutil
from datetime import datetime

def patch_vi2_with_time(template_path, output_path, data_list, start_dt, end_dt):
    """
    Patches a template to avoid 'No Protocols Found'.
    Calculates the 'gap' so data fits between start_dt and end_dt.
    """
    # 1. Map the template structure using olefile (no locks)
    with olefile.OleFileIO(template_path) as ole:
        values_path = None
        for entry in ole.listdir():
            if "values" in entry and "data" in entry:
                values_path = entry # e.g., ['30438', 'data', 'values']
                break
    
    if not values_path:
        print("Error: Template is not a valid Testo .vi2 file.")
        return

    # 2. Clone the template to preserve the 'Handshake' IDs
    shutil.copy2(template_path, output_path)
    
    # Flags: ReadWrite + Exclusive
    FLAGS = 0x00000002 | 0x00000010 

    try:
        root = pythoncom.StgOpenStorage(output_path, None, FLAGS, None, 0)
        proto_stg = root.OpenStorage(values_path[0], None, FLAGS)
        data_stg = proto_stg.OpenStorage("data", None, FLAGS)
        v_stream = data_stg.OpenStream("values", None, FLAGS)

        # 3. DYNAMIC TIME CALCULATION
        # Testo Tick Baseline (March 2026)
        base_tick = 3094906529 
        
        # Calculate seconds between your desired start and end
        total_seconds = int((end_dt - start_dt).total_seconds())
        # The gap is the sampling rate (e.g., 60 for 1 minute)
        gap = total_seconds // (len(data_list) - 1) if len(data_list) > 1 else 1
        
        # 4. WRITE THE DATA ROWS (12-bytes: Iff)
        new_payload = bytearray()
        for i, (t, h) in enumerate(data_list):
            tick = base_tick + (i * gap)
            new_payload.extend(struct.pack('<Iff', tick, float(t), float(h)))

        v_stream.SetSize(len(new_payload))
        v_stream.Seek(0, 0)
        v_stream.Write(new_payload)

        # 5. UPDATE THE SUMMARY (The Index)
        # This is where we tell ComSoft the new count and time span
        summ_stream = proto_stg.OpenStream("summary", None, FLAGS)
        s_bytes = bytearray(36)
        s_bytes[0:4] = struct.pack('<I', base_tick)           # Start Tick
        s_bytes[8:12] = b'\x03\x00\x00\x00'                   # Protocol Version
        s_bytes[12:16] = struct.pack('<I', len(data_list))    # Actual Row Count
        s_bytes[32:36] = struct.pack('<I', base_tick + (len(data_list)-1)*gap) # End Tick
        
        summ_stream.Seek(0, 0)
        summ_stream.Write(s_bytes)

        root.Commit(storagecon.STGC_DEFAULT)
        print(f"Success! {output_path} created.")
        print(f"Sampling Rate: {gap} seconds. Points: {len(data_list)}")

    except Exception as e:
        print(f"Patching Error: {e}")
    finally:
        root = None

# --- HOW TO USE ---
# 1. Your real data
my_data = [(21.5, 45.0), (21.7, 45.2), (21.9, 45.5)] * 200 # 600 points

# 2. Your real timeframe
start_time = datetime(2026, 3, 30, 8, 0, 0)
end_time = datetime(2026, 3, 30, 18, 0, 0) # 10 hours of data

# 3. Patch the input file (Must be a working .vi2 file)
patch_vi2_with_time("input.vi2", "final_working_protocol.vi2", my_data, start_time, end_time)