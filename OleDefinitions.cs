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
        public IntPtr unionmember1;
        public IntPtr unionmember2; // Pads out the union size (24 bytes on x64)

        public static PROPVARIANT LpStr(string s) => new PROPVARIANT
        {
            vt = OleConstants.VT_LPSTR,
            unionmember1 = Marshal.StringToCoTaskMemAnsi(s ?? string.Empty)
        };

        public static unsafe PROPVARIANT R8(double d)
        {
            PROPVARIANT pv = new PROPVARIANT { vt = OleConstants.VT_R8 };
            *(double*)&pv.unionmember1 = d;
            return pv;
        }

        public static unsafe PROPVARIANT UI4(uint u)
        {
            PROPVARIANT pv = new PROPVARIANT { vt = OleConstants.VT_UI4 };
            *(uint*)&pv.unionmember1 = u;
            return pv;
        }

        public void Free()
        {
            if ((vt == OleConstants.VT_LPSTR || vt == OleConstants.VT_LPWSTR) && unionmember1 != IntPtr.Zero)
                Marshal.FreeCoTaskMem(unionmember1);
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct PROPSPEC
    {
        public uint ulKind;
        public IntPtr data; // Handles x64 C++ union padding safely

        public static PROPSPEC FromId(int id) => new PROPSPEC
        {
            ulKind = OleConstants.PRSPEC_PROPID,
            data = (IntPtr)id
        };
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct STATPROPSTG
    {
        [MarshalAs(UnmanagedType.LPWStr)] public string lpwstrName;
        public uint propid;
        public ushort vt;
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
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] public string StandardName;
        public SYSTEMTIME StandardDate;
        public int StandardBias;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] public string DaylightName;
        public SYSTEMTIME DaylightDate;
        public int DaylightBias;
    }

    [ComImport, Guid("0000000c-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IStream
    {
        [PreserveSig] int Read([Out] byte[] pv, int cb, IntPtr pcbRead);
        [PreserveSig] int Write([In] byte[] pv, int cb, IntPtr pcbWritten);
        [PreserveSig] int Seek(long dlibMove, uint dwOrigin, IntPtr plibNewPosition);
        [PreserveSig] int SetSize(long libNewSize);
        [PreserveSig] int CopyTo(IStream pstm, long cb, IntPtr pcbRead, IntPtr pcbWritten);
        [PreserveSig] int Commit(uint grfCommitFlags);
        [PreserveSig] int Revert();
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
        public const uint STGM_READWRITE       = 0x00000002;
        public const uint STGM_SHARE_EXCLUSIVE = 0x00000010;
        public const uint STGM_CREATE          = 0x00001000;
        public const uint STGM_RW_EXCL_CREATE  = STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE;

        public const uint   PRSPEC_PROPID        = 1;
        public const ushort VT_LPSTR             = 30;
        public const ushort VT_LPWSTR            = 31;
        public const ushort VT_UI4               = 19;
        public const ushort VT_R8                = 5;

        public static readonly Guid IID_IStorage             = new Guid("0000000b-0000-0000-C000-000000000046");
        public static readonly Guid Fmtid_SummaryInformation = new Guid("F29F85E0-4FF9-1068-AB91-08002B27B3D9");
    }

    public static class OleHelper
    {
        public static byte[] StructureToByteArray(object obj)
        {
            int len = Marshal.SizeOf(obj);
            byte[] arr = new byte[len];
            IntPtr ptr = Marshal.AllocHGlobal(len);
            Marshal.StructureToPtr(obj, ptr, false);
            Marshal.Copy(ptr, arr, 0, len);
            Marshal.FreeHGlobal(ptr);
            return arr;
        }

        // Writes a property set to any IStorage. props is an array of (propId, PROPVARIANT) pairs.
        public static void WritePropertySet(IStorage storage, Guid fmtid, (int id, PROPVARIANT value)[] props)
        {
            IPropertySetStorage propSetStorage = (IPropertySetStorage)storage;
            Guid clsid = Guid.Empty;

            int hr = propSetStorage.Create(ref fmtid, ref clsid, 0, OleConstants.STGM_RW_EXCL_CREATE, out IPropertyStorage propStorage);
            if (hr != 0) throw new COMException($"IPropertySetStorage.Create failed: 0x{hr:X8}", hr);

            try
            {
                PROPSPEC[]   specs  = new PROPSPEC[props.Length];
                PROPVARIANT[] vars  = new PROPVARIANT[props.Length];
                for (int i = 0; i < props.Length; i++)
                {
                    specs[i] = PROPSPEC.FromId(props[i].id);
                    vars[i]  = props[i].value;
                }

                int writeHr = propStorage.WriteMultiple((uint)props.Length, specs, vars, 2);

                foreach (var v in vars) v.Free();

                if (writeHr != 0) throw new COMException($"WriteMultiple failed: 0x{writeHr:X8}", writeHr);

                propStorage.Commit(0);
            }
            finally
            {
                Marshal.ReleaseComObject(propStorage);
            }
        }
    }
}
