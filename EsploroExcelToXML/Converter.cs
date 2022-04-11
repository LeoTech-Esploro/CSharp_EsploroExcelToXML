using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Schema;

namespace EsploroExcelToXML
{
    public partial class Converter
    {
        public static bool hasValidationIssues = false;

        /// <summary>
        /// Takes an Excel file and an output file path as input and produces and validates an XML file that can be imported into Esploro
        /// </summary>
        public static void ConvertExceltoXML(string excelPath, string xmlPath)
        {
            // Check that excel file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine("ERROR: Could not find Excel file: " + excelPath);
                Environment.Exit(-5);
            }

            // Create users
            List<User> users = CreateUsersFromExcel(excelPath);

            // Create XML
            SerializeUsersAsXMLToFile(users, xmlPath);

            if (!File.Exists(xmlPath))
            {
                Console.WriteLine("ERROR: Could not save XML file to path: " + xmlPath);
                Console.WriteLine("(Do you have permission?)");
                Environment.Exit(-6);
            }

            // Validate XML
            // TODO: Validate *before* the document gets written to disk

            // Make sure schema files exist
            if (!File.Exists("rest_user.xsd"))
            {
                Console.WriteLine("ERROR: Could not find Schema file: rest_user.xsd. Cannot validate output!");
                Environment.Exit(-7);
            }
            if (!File.Exists("rest_researcher.xsd"))
            {
                Console.WriteLine("ERROR: Could not find Schema file: rest_researcher.xsd. Cannot validate output!");
                Environment.Exit(-8);
            }

            XmlDocument xml = new XmlDocument();
            xml.Load(xmlPath);
            xml.Schemas.Add("", XmlReader.Create(File.OpenRead("rest_user.xsd")));

            // TODO: I had to cheat and add <users> to the schema.
            // I suspect this indicates deeper issues with this whole method. We'll see.
            xml.Validate(ValidationEventHandler);

            if (hasValidationIssues)
            {
                Console.WriteLine("ERROR: XML Output failed to validate! It will not import correctly!");
                Environment.Exit(-4);
            }
            else
            {
                Console.WriteLine("INFO: Passed validation without errors.");
            }
        }

        public static void ValidationEventHandler(object sender, ValidationEventArgs e)
        {
            hasValidationIssues = true;

            switch (e.Severity)
            {
                case XmlSeverityType.Error:
                    Console.WriteLine("VALIDATION ERROR: {0}", e.Message);
                    break;
                case XmlSeverityType.Warning:
                    Console.WriteLine("VALIDATION WARNING: {0}", e.Message);
                    break;
            }
        }
    }
}
