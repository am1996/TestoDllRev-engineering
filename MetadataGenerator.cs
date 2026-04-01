using System;
using OLEWriter;

namespace MetadataGenerator;

public class PropertySetWriter
{
    private static Guid FMTID_DocSummary = new Guid("9C45A730-7826-11D4-A48C-0000C0403AD3");

    public void WriteSpecificProperties(IStorage storage) =>
        OleHelper.WritePropertySet(storage, FMTID_DocSummary, new[]
        {
            (2, PROPVARIANT.LpStr("temperature")),  // name
            (3, PROPVARIANT.R8(-20.0)),              // min
            (4, PROPVARIANT.R8(70.0)),               // max
            (8, PROPVARIANT.UI4(4294967295)),         // flag
            (9, PROPVARIANT.UI4(34303)),              // id
        });
}
