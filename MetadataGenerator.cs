using System;
using System.IO;
using NPOI.HPSF;

public class MetadataGenerator
{
    public byte[] CreateOleMetadata()
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
        customProperties.Put("temperature", 23.5);

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