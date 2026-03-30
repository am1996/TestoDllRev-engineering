from vi2utils import *
import struct
import olefile

# --- EXECUTION EXAMPLE ---

def debug_gap(filename):
    with olefile.OleFileIO(filename) as ole:
        # read the contents of the 'values' stream for debugging
        values_path = None
        for entry in ole.listdir():
            if "values" in entry or "timezone" in entry:
                values_path = entry 
                break
            if values_path:
                stream = ole.openstream(values_path)
                data = stream.read()
                print(data)
                print(f"Raw data from {values_path}: {data[:50]}...")  # Print first 50 bytes
                # Interpret the first few records (assuming 12 bytes each)
                for i in range(0, min(5, len(data) // 12)):
                    tick, temp, hum = struct.unpack('<Iff', data[i*12:(i+1)*12])
                    print(f"Record {i}: Tick={tick}, Temp={temp}, Hum={hum}")
                
                else:
                        print("No 'values' stream found.")

debug_gap("input.vi2")