using System;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using MetadataGenerator;

namespace OLEWriter
{
    [SupportedOSPlatform("windows")]
    class Program
    {
        [DllImport("ole32.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
        static extern int StgCreateStorageEx(
            string pwcsName, uint grfMode, uint stgfmt, uint grfAttrs,
            IntPtr pStgOptions, IntPtr reserved2,
            ref Guid riid, out IStorage ppObjectOpen);

        static void Main(string[] args)
        {
            string serialNumber = args.Length > 0 ? args[0] : "85640531";
            string filePath = "SampleData.vi2";
            Guid iidStorage = OleConstants.IID_IStorage;

            int hr = StgCreateStorageEx(filePath, OleConstants.STGM_RW_EXCL_CREATE, 0, 0,
                IntPtr.Zero, IntPtr.Zero, ref iidStorage, out IStorage rootStorage);

            if (hr != 0) { Console.WriteLine($"StgCreateStorageEx failed: 0x{hr:X8}"); return; }

            try
            {
                rootStorage.CreateStorage("30438", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out IStorage ss_30438);
                rootStorage.CreateStream("org", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out IStream orgStream);

                ss_30438.CreateStream("t17b", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out var dataStream);
                ss_30438.CreateStream("summary", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out var summaryStream);

                ss_30438.CreateStorage("data", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out IStorage data);
                data.CreateStream("scheme", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out var schemeStream);
                data.CreateStream("timezone", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out var timezoneStream);
                data.CreateStream("values", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out var valuesStream);

                ss_30438.CreateStorage("channels", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out IStorage channelStorage);
                channelStorage.CreateStorage("1", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out IStorage channel1Storage);
                channelStorage.CreateStorage("2", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out IStorage channel2Storage);

                schemeStream.Write(new byte[] { 0x02, 0x00, 0x01, 0x00, 0x03, 0x00 }, 6, IntPtr.Zero);

                byte[] header = CreateOleHeader("30438", 30438);
                OLE_TZI_FORMAT tzi = CreateEgyptTzi(); //TODO: edit to match the byte stream
                byte [] summary = CreateSummaryInformation(DateTime.UtcNow, 42); //TODO: edit to match the byte stream
                summaryStream.Write(summary, summary.Length, IntPtr.Zero);  // This stream contains the summary information properties, which point to the data stream for their values
                orgStream.Write(header, header.Length, IntPtr.Zero);
                new PropertySetWriter().WriteTemperatureProperties(channel1Storage);
                new PropertySetWriter().WriteHumidityProperties(channel2Storage);
                timezoneStream.Write(OleHelper.StructureToByteArray(tzi), Marshal.SizeOf(tzi), IntPtr.Zero);
                WriteSummaryInformation(ss_30438, serialNumber);
                rootStorage.Commit(0);

                Marshal.ReleaseComObject(channel2Storage);
                Marshal.ReleaseComObject(channel1Storage);
                Marshal.ReleaseComObject(channelStorage);
                Marshal.ReleaseComObject(valuesStream);
                Marshal.ReleaseComObject(timezoneStream);
                Marshal.ReleaseComObject(schemeStream);
                Marshal.ReleaseComObject(data);
                Marshal.ReleaseComObject(summaryStream);
                Marshal.ReleaseComObject(dataStream);
                Marshal.ReleaseComObject(ss_30438);
                Marshal.ReleaseComObject(orgStream);

                Console.WriteLine("Success! File created with Summary Information.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
            }
            finally
            {
                if (Marshal.IsComObject(rootStorage)) Marshal.ReleaseComObject(rootStorage);
            }
        }
        // oaZ4uCEEpbgDAAAABxkAADAAAAAAAAAAAAAAAOCTBAChpni4 where startdatetime is in the beginning and the end and datalength in the middle
        public static byte[] CreateSummaryInformation(DateTime startDatetime, int dataOffset)
        {
            // Standard OLE headers for 2 properties are usually 48 bytes
            byte[] rawBytes = new byte[48];

            // 1. Convert Time to FILETIME
            long fileTime = startDatetime.ToFileTimeUtc();
            byte[] timeBytes = BitConverter.GetBytes(fileTime);

            // 2. Write Modified Time (Bytes 0-7)
            Array.Copy(timeBytes, 0, rawBytes, 0, 8);

            // 3. PROPERTY 1: Subject (PID 3)
            // Byte 8-11: The ID
            BitConverter.GetBytes((uint)3).CopyTo(rawBytes, 8);
            // Byte 12-15: The Offset (Your 6407)
            BitConverter.GetBytes(dataOffset).CopyTo(rawBytes, 12);

            // 4. PROPERTY 2: Content Status (PID 48)
            // Byte 16-19: The ID (0x30)
            BitConverter.GetBytes((uint)48).CopyTo(rawBytes, 16);
            // Byte 20-23: The Offset for the Status string
            // If you don't have a second string, point it to dataOffset + some padding
            BitConverter.GetBytes(dataOffset + 128).CopyTo(rawBytes, 20);

            // 5. Reserved / Padding (Bytes 24-39)
            // Leave as 00

            // 6. Write Created Time (Bytes 40-47)
            // In a 48-byte block, the second timestamp is at the very end
            Array.Copy(timeBytes, 0, rawBytes, 40, 8);

            return rawBytes;
        }   
        public static byte[] CreateOleHeader(string storageName, uint dataLength)
        {
            // 1. Ensure the string ends with a null terminator for OLE compatibility
            if (!storageName.EndsWith("\0"))
            {
                storageName += "\0";
            }

            // 2. Convert string to ASCII bytes
            byte[] nameBytes = System.Text.Encoding.ASCII.GetBytes(storageName);

            // 3. Prepare the header (12 fixed bytes + length of the name)
            byte[] fullHeader = new byte[12 + nameBytes.Length];

            // OLE Version (5.0) -> 00 05 00 00
            byte[] version = BitConverter.GetBytes((uint)0x00000500);
            
            // Format ID (Embedded) -> 01 00 00 00
            byte[] formatId = BitConverter.GetBytes((uint)0x00000001);
            
            // Data Length (The size of the actual data following this header)
            byte[] len = BitConverter.GetBytes(dataLength);

            // 4. Assemble the array
            Buffer.BlockCopy(version, 0, fullHeader, 0, 4);
            Buffer.BlockCopy(formatId, 0, fullHeader, 4, 4);
            Buffer.BlockCopy(len, 0, fullHeader, 8, 4);
            Buffer.BlockCopy(nameBytes, 0, fullHeader, 12, nameBytes.Length);

            return fullHeader;
        }
        public static OLE_TZI_FORMAT CreateEgyptTzi()
        {
            byte[] tzi = new byte[]{
                0xbc, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 
                0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 
                0x88, 0xff, 0xff, 0xff, 0x45, 0x00, 0x67, 0x00, 
                0x79, 0x00, 0x70, 0x00, 0x74, 0x00, 0x20, 0x00, 
                0x53, 0x00, 0x74, 0x00, 0x61, 0x00, 0x6e, 0x00, 
                0x64, 0x00, 0x61, 0x00, 0x72, 0x00, 0x64, 0x00, 
                0x20, 0x00, 0x54, 0x00, 0x69, 0x00, 0x6d, 0x00, 
                0x65, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 
                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 
                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 
                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x0a, 0x00, 
                0x04, 0x00, 0x05, 0x00, 0x17, 0x00, 0x3b, 0x00, 
                0x3b, 0x00, 0xe7, 0x03, 0x00, 0x00, 0x00, 0x00, 
                0x45, 0x00, 0x67, 0x00, 0x79, 0x00, 0x70, 0x00, 
                0x74, 0x00, 0x20, 0x00, 0x44, 0x00, 0x61, 0x00, 
                0x79, 0x00, 0x6c, 0x00, 0x69, 0x00, 0x67, 0x00, 
                0x68, 0x00, 0x74, 0x00, 0x20, 0x00, 0x54, 0x00, 
                0x69, 0x00, 0x6d, 0x00, 0x65, 0x00, 0x00, 0x00, 
                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 
                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 
                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 
                0x00, 0x00, 0x04, 0x00, 0x05, 0x00, 0x04, 0x00, 
                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 
                0xc4, 0xff, 0xff, 0xff, 
            };
             GCHandle handle = GCHandle.Alloc(tzi, GCHandleType.Pinned);
            return Marshal.PtrToStructure<OLE_TZI_FORMAT>(handle.AddrOfPinnedObject());
        }
        static void WriteSummaryInformation(IStorage storage, string serialNumber) =>
            OleHelper.WritePropertySet(storage, OleConstants.Fmtid_SummaryInformation, new[]
            {
                (2, PROPVARIANT.LpStr("")),                      // PIDSI_TITLE
                (4, PROPVARIANT.LpStr($"t17b:{serialNumber}")),  // PIDSI_AUTHOR
                (6, PROPVARIANT.LpStr("")),                      // PIDSI_COMMENTS
            });
    }
}
