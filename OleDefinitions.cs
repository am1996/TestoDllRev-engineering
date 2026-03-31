using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace OLEWriter
{
    [StructLayout(LayoutKind.Sequential)]
    public struct PROPVARIANT
    {
        public ushort vt;
        public ushort wReserved1;
        public ushort wReserved2;
        public ushort wReserved3;
        public IntPtr unionmember1; // Holds the string pointer
        public IntPtr unionmember2; // Pads out the union size (24 bytes on x64)
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct PROPSPEC
    {
        public uint ulKind;
        public IntPtr data; // Replaces 'propid' to handle x64 C++ union padding safely
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct STATPROPSTG
    {
        [MarshalAs(UnmanagedType.LPWStr)] public string lpwstrName;
        public uint propid;
        public ushort vt;
    }

    [ComImport, Guid("0000000b-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IStorage
    {
        void CreateStream([In, MarshalAs(UnmanagedType.LPWStr)] string name, uint mode, uint r1, uint r2, out IStream stm);
        void OpenStream([In, MarshalAs(UnmanagedType.LPWStr)] string name, IntPtr r1, uint mode, uint r2, out IStream stm);
        void CreateStorage([In, MarshalAs(UnmanagedType.LPWStr)] string name, uint mode, uint r1, uint r2, out IStorage stg);
        void OpenStorage([In, MarshalAs(UnmanagedType.LPWStr)] string name, IntPtr prio, uint mode, IntPtr snb, uint res, out IStorage stg);
        void CopyTo(uint ciid, [In, MarshalAs(UnmanagedType.LPArray)] Guid[] iids, IntPtr snb, IStorage dest);
        void MoveElementTo([In, MarshalAs(UnmanagedType.LPWStr)] string name, IStorage dest, [In, MarshalAs(UnmanagedType.LPWStr)] string newName, uint flags);
        void Commit(uint flags);
        void Revert();
        void EnumElements(uint r1, IntPtr r2, uint r3, out IntPtr ppenum);
        void DestroyElement([In, MarshalAs(UnmanagedType.LPWStr)] string name);
        void SetElementTimes([In, MarshalAs(UnmanagedType.LPWStr)] string name, FILETIME c, FILETIME a, FILETIME m);
        void SetClass(ref Guid clsid);
        void SetStateBits(uint bits, uint mask);
        void Stat(out STATSTG stat, uint flags);
        void OpenStorageEx([In, MarshalAs(UnmanagedType.LPWStr)] string name, uint mode, uint stgfmt, uint grfAttrs, IntPtr pStgOptions, IntPtr reserved2, ref Guid riid, out IStorage stg);
    }

    [ComImport, Guid("0000013A-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPropertySetStorage
    {
        [PreserveSig] int Create([In] ref Guid rfmtid, [In] ref Guid pclsid, [In] uint grfFlags, [In] uint grfMode, [Out] out IPropertyStorage ppprstg);
        [PreserveSig] int Open([In] ref Guid rfmtid, [In] uint grfMode, [Out] out IPropertyStorage ppprstg);
        [PreserveSig] int Delete([In] ref Guid rfmtid);
        [PreserveSig] int Enum([Out] out IntPtr ppenum);
    }

    [ComImport, Guid("00000138-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPropertyStorage
    {
        [PreserveSig] int ReadMultiple(uint cpspec, [In, MarshalAs(UnmanagedType.LPArray)] PROPSPEC[] rgpspec, [Out, MarshalAs(UnmanagedType.LPArray)] PROPVARIANT[] rgpropvar);
        [PreserveSig] int WriteMultiple(uint cpspec, [In, MarshalAs(UnmanagedType.LPArray)] PROPSPEC[] rgpspec, [In, MarshalAs(UnmanagedType.LPArray)] PROPVARIANT[] rgpropvar, uint propidNameFirst);
        [PreserveSig] int DeleteMultiple(uint cpspec, [In, MarshalAs(UnmanagedType.LPArray)] PROPSPEC[] rgpspec);
        [PreserveSig] int ReadPropertyNames(uint cpspec, [In, MarshalAs(UnmanagedType.LPArray)] uint[] rgpropid, [Out, MarshalAs(UnmanagedType.LPArray)] string[] rglpwstrName);
        [PreserveSig] int WritePropertyNames(uint cpspec, [In, MarshalAs(UnmanagedType.LPArray)] uint[] rgpropid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rglpwstrName);
        [PreserveSig] int Delete(uint propid);
        [PreserveSig] int Commit(uint grfCommitFlags);
        [PreserveSig] int Revert();
        [PreserveSig] int Enum(out IntPtr ppenum);
        [PreserveSig] int SetTimes(FILETIME ctime, FILETIME atime, FILETIME mtime);
        [PreserveSig] int SetClass(ref Guid clsid);
        [PreserveSig] int Stat(out STATPROPSTG stat);
        [PreserveSig] int SetLocale(uint locale);
    }

    public static class OleConstants
    {
        public const ushort VT_LPSTR = 30;
        public const uint STGM_READWRITE      = 0x00000002;
        public const uint STGM_SHARE_EXCLUSIVE = 0x00000010;
        public const uint STGM_CREATE         = 0x00001000;
        public const uint STGM_RW_EXCL_CREATE = STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE;

        public const uint PRSPEC_PROPID = 1;

        public const ushort VT_LPWSTR = 31;

        public static readonly Guid IID_IStorage            = new Guid("0000000b-0000-0000-C000-000000000046");
        public static readonly Guid Fmtid_SummaryInformation = new Guid("F29F85E0-4FF9-1068-AB91-08002B27B3D9");
    }
}
