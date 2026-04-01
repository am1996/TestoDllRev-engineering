using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.IO;
using NPOI.HPSF;
using OLEWriter;

namespace MetadataGenerator;


[StructLayout(LayoutKind.Sequential)]
public struct PROPSPEC2
{
    public uint ulKind;
    public int propid; 
}

[StructLayout(LayoutKind.Explicit)]
public struct PROPVARIANT2
{
    [FieldOffset(0)] public ushort vt;
    [FieldOffset(8)] public double dblVal;
    [FieldOffset(8)] public uint uintVal;
    [FieldOffset(8)] public IntPtr ptr;

    public static PROPVARIANT2 FromString(string s) => new PROPVARIANT2 { vt = 30, ptr = Marshal.StringToCoTaskMemAnsi(s) };
    public static PROPVARIANT2 FromDouble(double d) => new PROPVARIANT2 { vt = 5, dblVal = d };
    public static PROPVARIANT2 FromUInt(uint u) => new PROPVARIANT2 { vt = 19, uintVal = u };
}
public class MetadataGenerator
{
    public byte[] CreateOleMetadata(string key, float value)
    {
        // 1. Create the base Property Set
        DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
        
        // 2. Use the CustomProperties wrapper (this handles the Dictionary and ReadOnly issues)
        CustomProperties customProperties = dsi.CustomProperties;
        if (customProperties == null)
        {
            customProperties = new CustomProperties();
        }

        // 3. Add your data. 
        // This automatically handles creating the ID and the name mapping.
        customProperties.Put(key, value);

        // 4. Link the custom properties back to the DocumentSummaryInformation
        dsi.CustomProperties = customProperties;

        // 5. Serialize to bytes
        using (MemoryStream ms = new MemoryStream())
        {
            dsi.Write(ms);
            return ms.ToArray();
        }
    }
}

public class PropertySetWriter
{
    // DocumentSummaryInformation Guid
    private static Guid FMTID_DocSummary = new Guid("9C45A730-7826-11D4-A48C-0000C0403AD3");

    public void WriteSpecificProperties(object iStorageInstance)
    {
        IPropertySetStorage propSetStorage = (IPropertySetStorage)iStorageInstance;
        IPropertyStorage? propStorage = null;

        try
        {
            // 1. Create/Open the property stream
            propSetStorage.Create(ref FMTID_DocSummary, Guid.Empty, 0, 0x00000002 | 0x00000010 | 0x00001000, out propStorage);

            // 2. Prepare the Property Specifications (The IDs from your image)
            PROPSPEC2[] specs = new PROPSPEC2[5];
            PROPVARIANT2[] values = new PROPVARIANT2[5];

            // ID 2: "temperature" (String)
            specs[0] = new PROPSPEC2 { ulKind = 1, propid = 2 };
            values[0] = PROPVARIANT2.FromString("temperature");

            // ID 3: -20 (8 byte real / Double)
            specs[1] = new PROPSPEC2 { ulKind = 1, propid = 3 };
            values[1] = PROPVARIANT2.FromDouble(-20.0);

            // ID 4: 70 (8 byte real / Double)
            specs[2] = new PROPSPEC2 { ulKind = 1, propid = 4 };
            values[2] = PROPVARIANT2.FromDouble(70.0);

            // ID 8: 4294967295 (Unsigned Long / UI4)
            specs[3] = new PROPSPEC2 { ulKind = 1, propid = 8 };
            values[3] = PROPVARIANT2.FromUInt(4294967295);

            // ID 9: 34303 (Unsigned Long / UI4)
            specs[4] = new PROPSPEC2 { ulKind = 1, propid = 9 };
            values[4] = PROPVARIANT2.FromUInt(34303);

            // 3. Write all properties at once
            propStorage.WriteMultiple((uint)specs.Length, specs, values, 0);
            
            propStorage.Commit(0);
        }
        finally
        {
            if (propStorage != null) Marshal.ReleaseComObject(propStorage);
        }
    }
}