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

# Run it
data_list = extract_all_measurements("example.vi2")