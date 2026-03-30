using System;
using System.Runtime.InteropServices;
using Vi2Extractor;

namespace TestoReverseEng
{
    class Program
    {
        // The CLSID from your registry screenshot
        private const string CLSID_STR = "135B31F3-115A-4142-9CBA-269948E2EDAE";


        // 1. Define delegates for guessing. 
        // COM always needs 'IntPtr self' as the first argument.
        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        delegate int Try0(IntPtr self);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        delegate int Try1(IntPtr self, int a);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        delegate int Try2(IntPtr self, int a, int b);

        static void Main(string[] args)
        {
            TestoParser.Vi2Reader parser = new TestoParser.Vi2Reader();

            parser.ProcessFile("G:\\projects\\Vi2Converter\\Edited__85640531_2026_03_24_16_21_54.vi2");
            Console.WriteLine("--- Testo tcddka.dll Live Reverse Engineering ---");

            try
            {
                // Instantiate the 32-bit COM object
                Type comType = Type.GetTypeFromCLSID(new Guid(CLSID_STR));
                object comInstance = Activator.CreateInstance(comType);
                
                // Get the pointers
                IntPtr pUnk = Marshal.GetIUnknownForObject(comInstance);
                IntPtr vtablePtr = Marshal.ReadIntPtr(pUnk);
                Console.WriteLine($"COM Object Created. IUnknown Pointer: 0x{pUnk.ToString("X")}");
                Console.WriteLine($"VTable Pointer: 0x{vtablePtr.ToString("X")}");
                IntPtr addr = Marshal.ReadIntPtr(vtablePtr, 3 * IntPtr.Size);
                var test = Marshal.GetDelegateForFunctionPointer<Try2>(addr); // Testing 2 args

                Console.WriteLine("Calling Index 03 with 2 test arguments...");
                int res = test(pUnk,0, 1); 
                Console.WriteLine($"Success! Result: {res}");

                // Cleanup
                Marshal.Release(pUnk);
                Marshal.ReleaseComObject(comInstance);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n[CRASH] The call was invalid: {ex.Message}");
                Console.WriteLine("This usually means the function expected more/fewer arguments.");
            }

            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
        }

        static void CallIndex(IntPtr pUnk, IntPtr vtablePtr, int index)
        {
            IntPtr funcAddr = Marshal.ReadIntPtr(vtablePtr, index * IntPtr.Size);
            Console.WriteLine($"\nProbing Index {index:D2} at Address 0x{funcAddr.ToString("X")}");

            // Map to our 0-argument delegate
            var method = Marshal.GetDelegateForFunctionPointer<Try0>(funcAddr);

            // Execute
            int hresult = method(pUnk);
            
            Console.WriteLine($"Result: 0x{hresult:X}");
            
            if (hresult == 0) 
                Console.WriteLine(">>> SUCCESS: Method returned S_OK!");
            else if (hresult == 1)
                Console.WriteLine(">>> SUCCESS: Method returned True/1");
        }
    }
}