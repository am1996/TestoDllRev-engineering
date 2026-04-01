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
                rootStorage.CreateStream("org", OleConstants.STGM_RW_EXCL_CREATE, 0, 0, out var orgStream);

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

                TIME_ZONE_INFORMATION tzi = new TIME_ZONE_INFORMATION
                {
                    Bias = -120,
                    StandardName = "Egypt Standard Time",
                    StandardDate = new SYSTEMTIME { wYear = 0, wMonth = 10, wDay = 5, wDayOfWeek = 5, wHour = 0 },
                    StandardBias = 0,
                    DaylightName = "Egypt Daylight Time",
                    DaylightDate = new SYSTEMTIME { wYear = 0, wMonth = 4, wDay = 4, wDayOfWeek = 5, wHour = 0 },
                    DaylightBias = -60
                };
                timezoneStream.Write(OleHelper.StructureToByteArray(tzi), Marshal.SizeOf(tzi), IntPtr.Zero);

                WriteSummaryInformation(ss_30438, serialNumber);
                new PropertySetWriter().WriteTemperatureProperties(channel1Storage);
                new PropertySetWriter().WriteHumidityProperties(channel2Storage);
                timezoneStream.Write(OleHelper.StructureToByteArray(tzi), Marshal.SizeOf(tzi), IntPtr.Zero);
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

        static void WriteSummaryInformation(IStorage storage, string serialNumber) =>
            OleHelper.WritePropertySet(storage, OleConstants.Fmtid_SummaryInformation, new[]
            {
                (2, PROPVARIANT.LpStr("")),                      // PIDSI_TITLE
                (4, PROPVARIANT.LpStr($"t17b:{serialNumber}")),  // PIDSI_AUTHOR
                (6, PROPVARIANT.LpStr("")),                      // PIDSI_COMMENTS
            });
    }
}
