using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace OLEWriter
{
    // Necessary COM interface for OLE Storages
    [ComImport]
    [Guid("0000000b-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IStorage
    {
        [PreserveSig]
        int CreateStream([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, uint grfMode, uint reserved1, uint reserved2, out IStream ppstm);
        [PreserveSig]
        int OpenStream([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, IntPtr reserved1, uint grfMode, uint reserved2, out IStream ppstm);
        [PreserveSig]
        int CreateStorage([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, uint grfMode, uint reserved1, uint reserved2, out IStorage ppstg);
        // ... other methods omitted for brevity
    }

    class Program
    {
        // OLE Modes
        const uint STGM_READWRITE = 0x00000002;
        const uint STGM_CREATE = 0x00001000;
        const uint STGM_SHARE_EXCLUSIVE = 0x00000010;

        [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
        static extern int StgCreateStorageEx(
            string pwcsName,
            uint grfMode,
            uint stgfmt,
            uint grfAttrs,
            IntPtr pStgOptions,
            IntPtr reserved2,
            ref Guid riid,
            out IStorage ppObjectOpen);

        static void Main(string[] args)
        {
            string filePath = "SampleData.ole";
            Guid IID_IStorage = new Guid("0000000b-0000-0000-C000-000000000046"); // GUID for IStorage

            int hr = StgCreateStorageEx(
                filePath,
                STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE,
                0,
                0,
                IntPtr.Zero,
                IntPtr.Zero,
                ref IID_IStorage,
                out IStorage rootStorage
            );

            if (hr != 0) throw new Exception($"Failed to create storage. HR: {hr}");

            Console.WriteLine("OLE File Created.");

            // 2. Create a Sub-Storage (A 'Folder' inside the file)
            rootStorage.CreateStorage("MySubFolder", STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, 0, out IStorage subStorage);

            // 3. Create a Stream (A 'File' inside the folder)
            subStorage.CreateStream("DataStream", STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, 0, out IStream stream);

            // 4. Write data to the stream
            string content = "This data is hidden inside an OLE container!";
            byte[] bytes = System.Text.Encoding.UTF8.GetBytes(content);
            stream.Write(bytes, bytes.Length, IntPtr.Zero);

            // if in windows do the following to ensure the file is properly closed and flushed to disk
            if(RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                // 5. Cleanup
                Marshal.ReleaseComObject(stream);
                Marshal.ReleaseComObject(subStorage);
                Marshal.ReleaseComObject(rootStorage);
            };

            Console.WriteLine($"Successfully wrote to {filePath}");
        }
    }
}