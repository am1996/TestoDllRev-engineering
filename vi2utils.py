import olefile
import pythoncom
from win32com import storagecon
import struct
import shutil
from datetime import datetime

# --- CONFIGURATION & HELPERS ---

def date_to_testo_tick(date_str):
    """
    Converts YYYY-MM-DD HH:MM:SS to the specific 1980-based tick.
    Discovery: 0 = 1980-01-01 02:00:00 AM
    """
    testo_zero = datetime(1980, 1, 1, 2, 0, 0)
    try:
        target_dt = datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
    except ValueError:
        # Fallback if someone includes AM/PM in the string
        target_dt = datetime.strptime(date_str.replace(" AM", "").replace(" PM", ""), '%Y-%m-%d %H:%M:%S')
    
    diff = target_dt - testo_zero
    return int(diff.total_seconds())

def write_vi2(template_path, output_path, data_list, new_serial="99999999", base_date="2024-01-01 00:00:00", gap_in_sec=60):
    """
    Patches a .vi2 file with custom data, serial number, and timestamps.
    """
    # 1. Map the template using olefile (Safe Read)
    with olefile.OleFileIO(template_path) as ole:
        values_path = None
        for entry in ole.listdir():
            if "values" in entry and "data" in entry:
                values_path = entry 
                break
    
    if not values_path:
        print("Error: Could not find valid protocol structure in template.")
        return

    # 2. Duplicate the template to preserve internal OLE IDs
    shutil.copy2(template_path, output_path)
    
    # Flags for Read/Write and Exclusive access
    FLAGS = storagecon.STGM_READWRITE | storagecon.STGM_SHARE_EXCLUSIVE

    try:
        root = pythoncom.StgOpenStorage(output_path, None, FLAGS, None, 0)
        proto_folder_name = values_path[0] # Usually '30438'
        proto_stg = root.OpenStorage(proto_folder_name, None, FLAGS)
        
        # --- PART A: EDIT SERIAL NUMBER & METADATA ---
        try:
            t17b_stream = proto_stg.OpenStream("t17b", None, FLAGS)
            # Reconstruct metadata to ensure 'ProgTime' matches our data gap
            # ComSoft needs the '\t' (Tab) character to parse correctly
            new_t17b_data = (
                f"SerialNumber\t{new_serial}\r\n"
                f"DeviceType\t4\r\n"
                f"ProgTime\t{gap_in_sec}\r\n"
            )
            t17b_stream.SetSize(len(new_t17b_data))
            t17b_stream.Seek(0, 0)
            t17b_stream.Write(new_t17b_data.encode('ascii'))
            print(f"Metadata updated: Serial {new_serial}, Gap {gap_in_sec}s")
        except Exception as e:
            print(f"Warning: Could not update t17b metadata: {e}")

        # --- PART B: EDIT DATA ROWS ---
        data_stg = proto_stg.OpenStorage("data", None, FLAGS)
        v_stream = data_stg.OpenStream("values", None, FLAGS)
        
        base_tick = 0
        new_payload = bytearray()
        
        last_tick = base_tick
        for i, (t, h) in enumerate(data_list):
            current_tick = base_tick + (i * gap_in_sec)
            # 12-byte row: <I (Unsigned Int Tick), ff (Two Floats)
            new_payload.extend(struct.pack('<Iff', current_tick, float(t), float(h)))
            last_tick = current_tick

        v_stream.SetSize(len(new_payload))
        v_stream.Seek(0, 0)
        v_stream.Write(new_payload)

        # --- PART C: UPDATE SUMMARY ---
        # This is the 'Master Index' ComSoft uses to draw the X-Axis
        summ_stream = proto_stg.OpenStream("summary", None, FLAGS)
        s_bytes = bytearray(36)
        s_bytes[0:4] = struct.pack('<I', base_tick)           # Start Tick
        s_bytes[8:12] = b'\x03\x00\x00\x00'                   # Protocol Version
        s_bytes[12:16] = struct.pack('<I', len(data_list))    # Point Count
        s_bytes[32:36] = struct.pack('<I', last_tick)         # End Tick
        
        summ_stream.Seek(0, 0)
        summ_stream.Write(s_bytes)

        # Finalize
        root.Commit(storagecon.STGC_DEFAULT)
        print(f"Successfully created: {output_path}")
        print(f"Ticks: {base_tick} to {last_tick}")

    except Exception as e:
        print(f"Critical Error during writing: {e}")
    finally:
        # Explicitly clear variables to release COM locks
        v_stream = None
        data_stg = None
        summ_stream = None
        proto_stg = None
        root = None