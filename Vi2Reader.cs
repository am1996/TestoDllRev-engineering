using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace TestoParser
{
    // --- 1. COM INTERFACES (The "Map" to OLE) ---
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct STATSTG
    {
        [MarshalAs(UnmanagedType.LPWStr)] public string pwcsName;
        public int type; // 1 = Stream, 2 = Storage (Folder)
        public long cbSize;
        public System.Runtime.InteropServices.ComTypes.FILETIME mtime;
        public System.Runtime.InteropServices.ComTypes.FILETIME ctime;
        public System.Runtime.InteropServices.ComTypes.FILETIME atime;
        public int grfMode;
        public int grfLocksSupported;
        public Guid clsid;
        public int grfStateBits;
        public int reserved;
    }

    [ComImport, Guid("0000000d-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IEnumSTATSTG
    {
        [PreserveSig] int Next(uint celt, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] STATSTG[] rgelt, out uint pceltFetched);
        [PreserveSig] int Skip(uint celt);
        [PreserveSig] int Reset();
        [PreserveSig] int Clone(out IEnumSTATSTG ppenum);
    }

    [ComImport, Guid("0000000b-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IStorage
    {
        [PreserveSig] int CreateStream([In, MarshalAs(UnmanagedType.LPWStr)] string pwcsName, uint grfMode, uint res1, uint res2, out IStream ppstm);
        [PreserveSig] int OpenStream([In, MarshalAs(UnmanagedType.LPWStr)] string pwcsName, IntPtr res1, uint grfMode, uint res2, out IStream ppstm);
        [PreserveSig] int CreateStorage([In, MarshalAs(UnmanagedType.LPWStr)] string pwcsName, uint grfMode, uint res1, uint res2, out IStorage ppstg);
        [PreserveSig] int OpenStorage([In, MarshalAs(UnmanagedType.LPWStr)] string pwcsName, IntPtr res1, uint grfMode, IntPtr snbExclude, uint res2, out IStorage ppstg);
        [PreserveSig] int EnumElements(uint res1, IntPtr res2, uint res3, out IEnumSTATSTG ppenum);
        // Simplified for brevity, but contains the core navigation methods
    }

    // --- 2. NATIVE BRIDGE (Your Reversed Ghidra Functions) ---
    public static class TcoleBridge
    {
        [DllImport("TCOLE.DLL", EntryPoint = "?Tcsgets@@YAJPAUIStream@@PADK@Z", CallingConvention = CallingConvention.Cdecl)]
        public static extern int Tcsgets(IStream pStream, byte[] pBuffer, uint maxLength);

        [DllImport("ole32.dll")]
        public static extern int StgOpenStorage([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, IStorage pstgPriority, uint grfMode, IntPtr snbExclude, uint reserved, out IStorage ppstgOpen);
    }

    // --- 3. THE READER CLASS ---
    public class Vi2Reader
    {
        private const uint STGM_READ = 0x00000000;
        private const uint STGM_SHARE_EXCLUSIVE = 0x00000010;

        public void ProcessFile(string path)
        {
            IStorage root;
            int hr = TcoleBridge.StgOpenStorage(path, null, STGM_READ | STGM_SHARE_EXCLUSIVE, IntPtr.Zero, 0, out root);

            if (hr == 0)
            {
                Console.WriteLine("File Opened Successfully.");
                ExploreStorage(root, "");
            }
            else
            {
                Console.WriteLine($"Failed to open file. HRESULT: {hr:X}");
            }
        }

        private void ExploreStorage(IStorage storage, string indent)
        {
            IEnumSTATSTG enumerator;
            storage.EnumElements(0, IntPtr.Zero, 0, out enumerator);

            STATSTG[] stats = new STATSTG[1];
            uint fetched;

            while (enumerator.Next(1, stats, out fetched) == 0 && fetched > 0)
            {
                string name = stats[0].pwcsName;
                if (stats[0].type == 1) // STREAM (File)
                {
                    Console.WriteLine($"{indent}Stream Found: {name}");
                    ReadStreamContent(storage, name);
                }
                else if (stats[0].type == 2) // STORAGE (Folder)
                {
                    Console.WriteLine($"{indent}Entering Folder: {name}");
                    IStorage subFolder;
                    storage.OpenStorage(name, IntPtr.Zero, STGM_READ | STGM_SHARE_EXCLUSIVE, IntPtr.Zero, 0, out subFolder);
                    ExploreStorage(subFolder, indent + "  ");
                }
            }
        }

        private void ReadStreamContent(IStorage parent, string streamName)
        {
            IStream stream;
            parent.OpenStream(streamName, IntPtr.Zero, STGM_READ | STGM_SHARE_EXCLUSIVE, 0, out stream);

            if (stream != null)
            {
                byte[] buffer = new byte[260];
                // Use your reversed Tcsgets to try and find a string
                int result = TcoleBridge.Tcsgets(stream, buffer, 260);

                if (result >= 0)
                {
                    string content = Encoding.ASCII.GetString(buffer).TrimEnd('\0');
                    if (!string.IsNullOrWhiteSpace(content))
                    {
                        Console.WriteLine($"    Data Read: {content}");
                    }
                }
            }
        }
    }
}