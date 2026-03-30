import pythoncom
from win32com import storagecon
import struct
import shutil

def write_vi2_by_patching(template_path, output_path, data_list):
    """
    Exactly replicates your C# logic:
    1. Copies a working template.
    2. Overwrites 'values' and 'summary' with your own data.
    """
    # Step 1: Clone the template
    shutil.copy2(template_path, output_path)

    flags = storagecon.STGM_READWRITE | storagecon.STGM_SHARE_EXCLUSIVE
    
    try:
        # Step 2: Open the clone
        cf = pythoncom.StgOpenStorage(output_path, None, flags, None, 0)
        
        # Step 3: Find the protocol storage (usually '30438')
        storage_name = None
        enum = cf.EnumElements()
        for stat in enum:
            if stat[1] == storagecon.STGTY_STORAGE and stat[0] != "channels":
                storage_name = stat[0]
                break
        
        root_inner = cf.OpenStorage(storage_name, None, flags, None, 0)
        
        # Step 4: Access 'values' and calculate Ticks (The C# logic)
        data_stg = root_inner.OpenStorage("data", None, flags, None, 0)
        v_stream = data_stg.OpenStream("values", None, flags, 0)
        
        # We need the original bytes to get the Start/End ticks for the summary
        orig_bytes = v_stream.Read(v_stream.Stat()[2])
        
        # Using index 0 as the starting point for ticks
        new_start_tick = struct.unpack_from('<I', orig_bytes, 0)[0]
        # Use the tick from the end of the original file to keep the span realistic
        new_end_tick = struct.unpack_from('<I', orig_bytes, len(orig_bytes) - 12)[0]
        
        actual_gap = float(new_end_tick - new_start_tick) / (len(data_list) - 1)

        # Step 5: Build new 'values' buffer
        new_buffer = bytearray()
        for i, (temp, hum) in enumerate(data_list):
            curr_tick = int(round(new_start_tick + (i * actual_gap)))
            # Format: Tick (uint), Temp (float), Hum (float)
            new_buffer.extend(struct.pack('<Iff', curr_tick, float(temp), float(hum)))
        
        v_stream.SetSize(len(new_buffer))
        v_stream.Seek(0, 0)
        v_stream.Write(new_buffer)

        # Step 6: Update 'summary' stream (Crucial for Protocol Recognition)
        summ_stream = root_inner.OpenStream("summary", None, flags, 0)
        summ_bytes = bytearray(summ_stream.Read(36))

        # Offset 0: Start Tick
        summ_bytes[0:4] = struct.pack('<I', new_start_tick)
        # Offset 12: Point Count (Fixes Truncation)
        summ_bytes[12:16] = struct.pack('<I', len(data_list))
        # Offset 32: End Tick
        summ_bytes[32:36] = struct.pack('<I', new_end_tick)

        summ_stream.Seek(0, 0)
        summ_stream.Write(summ_bytes)

        cf.Commit(storagecon.STGC_DEFAULT)
        print(f"Patched protocol successfully: {output_path}")

    except Exception as e:
        print(f"Error during patch: {e}")
    finally:
        cf = None

# --- Usage ---
# Pass your list of (temp, hum) and the path to a WORKING .vi2 file
my_data = [(23.5, 55.2), (23.6, 55.3), (23.7, 55.4)]
write_vi2_by_patching("example.vi2", "final_output.vi2", my_data)