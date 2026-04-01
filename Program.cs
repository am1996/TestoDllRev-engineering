using NPOI.HPSF; // Header Property Storage Format
using NPOI.POIFS.FileSystem;
using System.IO;
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
                rootStorage.CreateStorage("30438", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out IStorage ss_30438_Storage);
                rootStorage.CreateStream("org", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out var orgStream);

                ss_30438_Storage.CreateStream("t17b", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out var dataStream);
                ss_30438_Storage.CreateStream("summary", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out var summaryStream);

                ss_30438_Storage.CreateStorage("data", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out IStorage dataStorage);
                dataStorage.CreateStream("scheme", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out var schemeStream);
                dataStorage.CreateStream("timezone", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out var timezoneStream);
                dataStorage.CreateStream("values", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out var valuesStream);
                ss_30438_Storage.CreateStorage("channels", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out IStorage channelStorage);
                channelStorage.CreateStorage("1",OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out IStorage channel1Storage);
                channelStorage.CreateStorage("2",OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out IStorage channel2Storage);

                // Hex: 01 00 02 00 03 00
                // Use IntPtr.Zero to tell Windows "I don't care how many bytes were written"
                schemeStream.Write(new byte[] { 0x02, 0x00, 0x01, 0x00, 0x03, 0x00 }, 6, IntPtr.Zero);
                TIME_ZONE_INFORMATION tzi = new TIME_ZONE_INFORMATION
                {
                    // UTC = local time + bias. For UTC+2, bias is -120.
                    Bias = -120, 
                    StandardName = "Egypt Standard Time",
                    StandardDate = new SYSTEMTIME 
                    { 
                        wYear = 0, 
                        wMonth = 10,     // October
                        wDay = 5,       // 5 means the "last" occurrence of the day
                        wDayOfWeek = 5, // 5 = Friday
                        wHour = 0       // Transition at midnight
                    },
                    StandardBias = 0,
                    DaylightName = "Egypt Daylight Time",
                    DaylightDate = new SYSTEMTIME 
                    { 
                        wYear = 0, 
                        wMonth = 4,      // April
                        wDay = 4,       // The 4th occurrence (April 24 is the 4th Friday in 2026)
                        wDayOfWeek = 5, // 5 = Friday
                        wHour = 0       // Transition at midnight
                    },
                    DaylightBias = -60 // Subtract 60 mins from standard bias during DST
                };
                byte[] tziBytes = StructureToByteArray(tzi);
                WriteSummaryInformation(ss_30438_Storage, serialNumber);
                timezoneStream.Write(tziBytes, tziBytes.Length, IntPtr.Zero);
                MetadataGenerator.PropertySetWriter propertysetwriter = new MetadataGenerator.PropertySetWriter();
                propertysetwriter.WriteSpecificProperties(channel1Storage);
                rootStorage.Commit(0);

                Marshal.ReleaseComObject(channelStorage);
                Marshal.ReleaseComObject(dataStorage);
                Marshal.ReleaseComObject(ss_30438_Storage);
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
        static void WriteSummaryInformation(IStorage rootStorage, string serialNumber)
        {
            IPropertySetStorage propSetStorage = (IPropertySetStorage)rootStorage;
            Guid fmtid = OleConstants.Fmtid_SummaryInformation;
            Guid clsid = Guid.Empty;

            int createHr = propSetStorage.Create(ref fmtid, ref clsid, 0, OleConstants.STGM_RW_EXCL_CREATE, out IPropertyStorage propStorage);
            if (createHr != 0) throw new COMException($"IPropertySetStorage.Create failed: 0x{createHr:X8}", createHr);

            try
            {
                PROPSPEC[] specs = new PROPSPEC[3];
                // Cast the integer property IDs to IntPtr to satisfy the new struct alignment
                specs[0] = new PROPSPEC { ulKind = OleConstants.PRSPEC_PROPID, data = (IntPtr)2 };
                specs[1] = new PROPSPEC { ulKind = OleConstants.PRSPEC_PROPID, data = (IntPtr)4 };
                specs[2] = new PROPSPEC { ulKind = OleConstants.PRSPEC_PROPID, data = (IntPtr)6 };

                PROPVARIANT[] vars = new PROPVARIANT[3];
                vars[0] = MakeLpstr("");
                vars[1] = MakeLpstr($"t17b:{serialNumber}");
                vars[2] = MakeLpstr("");

                int writeHr = propStorage.WriteMultiple(3, specs, vars, 2);

                // Cleanup memory using the new unionmember1 pointer
                foreach (var v in vars)
                    if (v.unionmember1 != IntPtr.Zero) Marshal.FreeCoTaskMem(v.unionmember1);

                if (writeHr != 0) throw new COMException($"WriteMultiple failed: 0x{writeHr:X8}", writeHr);
            }
            finally
            {
                Marshal.ReleaseComObject(propStorage);
            }
        }
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct SYSTEMTIME
        {
            public ushort wYear, wMonth, wDayOfWeek, wDay, wHour, wMinute, wSecond, wMilliseconds;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct TIME_ZONE_INFORMATION
        {
            public int Bias;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string StandardName;
            public SYSTEMTIME StandardDate;
            public int StandardBias;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string DaylightName;
            public SYSTEMTIME DaylightDate;
            public int DaylightBias;
        }

        // Helper to convert the struct to bytes for the IStream
        public static byte[] StructureToByteArray(object obj)
        {
            int len = Marshal.SizeOf(obj);
            byte[] arr = new byte[len];
            IntPtr ptr = Marshal.AllocHGlobal(len);
            Marshal.StructureToPtr(obj, ptr, true);
            Marshal.Copy(ptr, arr, 0, len);
            Marshal.FreeHGlobal(ptr);
            return arr;
        }
        static PROPVARIANT MakeLpstr(string s) => new PROPVARIANT
        {
            vt = OleConstants.VT_LPSTR,
            unionmember1 = Marshal.StringToCoTaskMemAnsi(s ?? string.Empty),
            unionmember2 = IntPtr.Zero // Pads the struct to ensure C++ array alignment
        };
    }
}
