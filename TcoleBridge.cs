using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes; // For IStream

namespace Vi2Extractor
{
    public static class TcoleBridge
    {
        // Name: ?Tcsgets@@YAJPAUIStream@@PADK@Z
        // Action: Reads a null-terminated string from a stream
        [DllImport("TCOLE.DLL", EntryPoint = "?Tcsgets@@YAJPAUIStream@@PADK@Z", CallingConvention = CallingConvention.Cdecl)]
        public static extern int Tcsgets(IStream pStream, byte[] pBuffer, uint maxLength);

        // Name: ?Tcsputs@@YAJPAUIStream@@PBD@Z
        // Action: Writes a string into a stream (calculates length automatically)
        [DllImport("TCOLE.DLL", EntryPoint = "?Tcsputs@@YAJPAUIStream@@PBD@Z", CallingConvention = CallingConvention.Cdecl)]
        public static extern int Tcsputs(IStream pStream, [MarshalAs(UnmanagedType.LPStr)] string pData);
    }
}