import olefile
import struct

def extract_all_measurements(filename):
    with olefile.OleFileIO(filename) as ole:
        # Path to the data
        stream_path = "30438/data/values"
        if not ole.exists(stream_path):
            print("Could not find values stream!")
            return

        data = ole.openstream(stream_path).read()
        size = len(data)
        
        # Each 'row' is 3 floats (Marker, Temp, Humidity) = 12 bytes
        row_size = 12 
        num_rows = size // row_size

        print(f"Reading {num_rows} measurements...")
        print(f"{'Index':<8} | {'Marker':<12} | {'Temp (°C)':<10} | {'Humidity (%)':<12}")
        print("-" * 50)

        # Iterate through the entire 76,884 bytes
        measurements = []
        for i in range(num_rows):
            offset = i * row_size
            chunk = data[offset : offset + row_size]
            
            # Unpack 3 Little-Endian Floats ('<3f')
            marker, temp, hum = struct.unpack('<3f', chunk)
            
            # Print first 20 and last 5 so we don't spam the terminal, 
            # but you can remove this 'if' to see everything.
            print(f"{i:<8} | {marker:<12.4e} | {temp:<10.2f} | {hum:<12.2f}")
            
            measurements.append((marker, temp, hum))

        return measurements

def read_vi2_metadata(file_path):
    if not olefile.isOleFile(file_path):
        print("Not a valid OLE file.")
        return

    with olefile.OleFileIO(file_path) as ole:
        print(f"--- Metadata for: {file_path} ---")
        
        # 1. Read Standard Summary Information
        # This usually contains the 'Title' we set in the writer (e.g., "t17b: 85640531")
        if ole.exists("\x05SummaryInformation"):
            meta = ole.get_metadata()
            print(f"Title:  {meta.title}")
            print(f"Author: {meta.author}")
            print(f"Created: {meta.create_time}")
        
        # 2. Read Testo-specific Device Info (t17b stream)
        # We need to find where 't17b' lives. It's usually inside the protocol storage.
        t17b_path = None
        for entry in ole.listdir():
            if "t17b" in entry:
                t17b_path = "/".join(entry)
                break
        
        if t17b_path:
            with ole.openstream(t17b_path) as s:
                raw_data = s.read().decode('ascii', errors='ignore')
                print("\n--- Device Details (t17b) ---")
                print(raw_data.strip())

        # 3. Read Protocol Summary (Ticks and Counts)
        # This mirrors the 36-byte logic used in the write function
        summary_path = t17b_path.replace("t17b", "summary") if t17b_path else None
        if summary_path and ole.exists(summary_path):
            with ole.openstream(summary_path) as s:
                data = s.read(36)
                start_tick, version, count, end_tick = struct.unpack_from('<IHHII', data, 0)[:4]
                # Note: count is at offset 12, end_tick at offset 32
                count = struct.unpack_from('<I', data, 12)[0]
                end_tick = struct.unpack_from('<I', data, 32)[0]
                
                print("\n--- Session Summary ---")
                print(f"Start Tick:  {start_tick}")
                print(f"End Tick:    {end_tick}")
                print(f"Data Points: {count}")

def debug_vi2_structure(file_path):
    with olefile.OleFileIO(file_path) as ole:
        # Find the values stream
        values_path = None
        for entry in ole.listdir():
            if "values" in entry:
                values_path = "/".join(entry)
                break
        
        if not values_path:
            print("Could not find 'values' stream.")
            return

        with ole.openstream(values_path) as s:
            data = s.read()
            total_rows = len(data) // 12
            print(f"Total Rows Found: {total_rows}")

            # Read First Row (12 bytes)
            first_row = struct.unpack('<Iff', data[:12])
            # Read Last Row (12 bytes)
            last_row = struct.unpack('<Iff', data[-12:])

            print(f"First Row Binary: {data[:12].hex(' ')}")
            print(f"First Row Decoded: Tick={first_row[0]}, T={first_row[1]:.2f}, H={first_row[2]:.2f}")
            
            print(f"Last Row Binary:  {data[-12:].hex(' ')}")
            print(f"Last Row Decoded:  Tick={last_row[0]}, T={last_row[1]:.2f}, H={last_row[2]:.2f}")

            # Check the Summary Stream (36 bytes)
            summary_path = values_path.replace("data/values", "summary")
            if ole.exists(summary_path):
                with ole.openstream(summary_path) as ss:
                    s_data = ss.read(36)
                    print(f"\nFull Summary Hex: {s_data.hex(' ')}")
                    # Testo often uses offsets 0 and 32 for the 'Active' start/end
                    s_start = struct.unpack_from('<I', s_data, 0)[0]
                    s_end = struct.unpack_from('<I', s_data, 32)[0]
                    print(f"Summary Start Tick (Offset 0):  {s_start}")
                    print(f"Summary End Tick   (Offset 32): {s_end}")

debug_vi2_structure("example.vi2")

