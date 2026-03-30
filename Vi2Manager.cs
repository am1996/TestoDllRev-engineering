using System;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace Vi2Extractor
{
    public class Vi2Manager
    {
        private const uint MAX_STR_LEN = 260;

        /// <summary>
        /// Reads a sensor name or metadata string using the native Tcsgets function.
        /// </summary>
        public string ReadString(IStream stream)
        {
            byte[] buffer = new byte[MAX_STR_LEN];
            
            // Call the native Testo function
            int hr = TcoleBridge.Tcsgets(stream, buffer, MAX_STR_LEN);

            if (hr < 0) // Check for -0x7ff8fff2 (Buffer Overflow) or other errors
            {
                Console.WriteLine($"Warning: Tcsgets returned HRESULT {hr:X}");
                return string.Empty;
            }

            // Convert the byte array back to a C# string, stopping at the first null
            string result = Encoding.ASCII.GetString(buffer);
            return result.Split('\0')[0];
        }

        /// <summary>
        /// Overwrites a string in the file using the native Tcsputs function.
        /// </summary>
        public bool WriteString(IStream stream, string newValue)
        {
            // Call the native Testo function
            int hr = TcoleBridge.Tcsputs(stream, newValue);

            if (hr != 0)
            {
                Console.WriteLine($"Error: Tcsputs failed with HRESULT {hr:X}");
                return false;
            }

            return true;
        }
    }
}