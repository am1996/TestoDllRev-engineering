import pythoncom
from win32com import storagecon
import struct
import shutil

def write_vi2_dynamic_tick(template_path, output_path, data_list):
    """
    Patches a .vi2 file by extracting the base_tick directly from 
    the template's first data row.
    """
    # 1. Clone the template
    shutil.copy2(template_path, output_path)
    flags = storagecon.STGM_READWRITE | storagecon.STGM_SHARE_EXCLUSIVE
    
    cf = None
    try:
        # Open the OLE file
        cf = pythoncom.StgOpenStorage(output_path, None, flags, None, 0)
        
        # 2. Find the main storage folder (e.g., "30438")
        storage_name = None
        enum = cf.EnumElements()
        for stat in enum:
            if stat[1] == storagecon.STGTY_STORAGE and stat[0] != "channels":
                storage_name = stat[0]
                break
        
        if not storage_name:
            print("Error: Could not find the data storage folder.")
            return

        root_inner = cf.OpenStorage(storage_name, None, flags, None, 0)
        data_stg = root_inner.OpenStorage("data", None, flags, None, 0)
        v_stream = data_stg.OpenStream("values", None, flags, 0)

        # 3. EXTRACT BASE TICK FROM FILE
        # Read the first 4 bytes of the existing 'values' stream
        header_bytes = v_stream.Read(4)
        if len(header_bytes) < 4:
            # Fallback if the template is empty
            base_tick = 3094906529 
            print("Warning: Template values stream empty. Using default base_tick.")
        else:
            base_tick = struct.unpack('<I', header_bytes)[0]
            print(f"Extracted Base Tick from template: {base_tick}")

        # 4. GENERATE NEW DATA
        # Gap is ~453 based on your 6407 row analysis
        gap = 453 
        new_values = bytearray()
        for i, (temp, hum) in enumerate(data_list):
            current_tick = base_tick + (i * gap)
            new_values.extend(struct.pack('<Iff', current_tick, float(temp), float(hum)))

        # 5. OVERWRITE VALUES
        v_stream.SetSize(len(new_values))
        v_stream.Seek(0, 0)
        v_stream.Write(new_values)

        # 6. UPDATE SUMMARY (The Protocol "Table of Contents")
        summ_stream = root_inner.OpenStream("summary", None, flags, 0)
        new_summary = bytearray(36)
        
        # Start Tick (Offset 0)
        new_summary[0:4] = struct.pack('<I', base_tick)
        # Magic/Version (Offset 8) - Standard Testo version marker
        new_summary[8:12] = bytes.fromhex("03 00 00 00") 
        # Count (Offset 12)
        new_summary[12:16] = struct.pack('<I', len(data_list))
        # End Tick (Offset 32)
        last_tick = base_tick + ((len(data_list) - 1) * gap)
        new_summary[32:36] = struct.pack('<I', last_tick)

        summ_stream.Seek(0, 0)
        summ_stream.Write(new_summary)

        # Commit and close
        cf.Commit(storagecon.STGC_DEFAULT)
        print(f"Success! Created {output_path} with {len(data_list)} points.")

    except Exception as e:
        print(f"Critical Error: {e}")
    finally:
        # Release COM objects
        v_stream = None
        data_stg = None
        summ_stream = None
        root_inner = None
        cf = None

# --- Run ---
# Example data: List of (temp, humidity)
test_data = [(21.5, 48.2), (21.6, 48.3), (21.7, 48.4)]
write_vi2_dynamic_tick("example.vi2", "dynamic_output.vi2", test_data)