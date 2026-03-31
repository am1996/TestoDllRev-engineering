using System;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

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
                rootStorage.CreateStream("org", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out var orgStream);

                ss_30438.CreateStream("t17b", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out var dataStream);
                ss_30438.CreateStream("summary", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out var summaryStream);

                ss_30438.CreateStorage("data", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out IStorage data);
                data.CreateStream("scheme", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out var schemeStream);
                data.CreateStream("timezone", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out var timezoneStream);
                data.CreateStream("values", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out var valuesStream);
                ss_30438.CreateStorage("channels", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out IStorage channels);

                // Hex: 01 00 02 00 03 00
                byte[] scehmeData = new byte[] { 0x02, 0x00, 0x01, 0x00, 0x03, 0x00 };
                schemeStream.Write(scehmeData, 0, scehmeData.Length);
                WriteSummaryInformation(ss_30438, serialNumber);

                rootStorage.Commit(0);

                Marshal.ReleaseComObject(channels);
                Marshal.ReleaseComObject(data);
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

        static PROPVARIANT MakeLpstr(string s) => new PROPVARIANT
        {
            vt = OleConstants.VT_LPSTR,
            unionmember1 = Marshal.StringToCoTaskMemAnsi(s ?? string.Empty),
            unionmember2 = IntPtr.Zero // Pads the struct to ensure C++ array alignment
        };
    }
}
