using System;

namespace EsploroExcelToXML
{
    class MainClass
    {
        public static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("ERROR: You need to supply an input excel file and an output XML file.");
                Console.WriteLine("Usage: EsploroExcelToXML input.xlsx output.xml");
                Environment.Exit(-3);
            }

            Console.WriteLine("Input Excel file: " + args[0]);
            Console.WriteLine("Output XML file: " + args[1]);

            Converter.ConvertExceltoXML(args[0], args[1]);

            Console.WriteLine("Conversion finished!");
        }
    }
}
