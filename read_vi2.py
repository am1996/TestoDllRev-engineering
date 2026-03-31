import olefile
import struct

# get file guid
file = olefile.OleFileIO("original.vi2")
stream = file.openstream("30438/data/values")

# Data stream reading
print("{:=^13} | {:=^20} | {:=^21}".format("Tick", "Temperature", "Humidity"))
for i in range(100):
    stream.seek(i * 12)
    # get 4 floating point values
    tick,temp, hum = struct.unpack("<iff", stream.read(12))

    print(" {:12} | {:20} | {:20}".format(tick, temp, hum))

print("{:=^13} | {:=^20} | {:=^21}".format("=", "=", "="))