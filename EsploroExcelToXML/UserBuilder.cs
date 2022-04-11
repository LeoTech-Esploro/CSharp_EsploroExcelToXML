using System;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;

namespace EsploroExcelToXML
{
    public partial class Converter
    {
        private static List<User> CreateUsersFromExcel(string excelPath)
        {
            // NOTE: Using EPPlus version 4.x so we can use the LGPL license
            // ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            ExcelPackage package = new ExcelPackage(new FileInfo(excelPath));
            ExcelWorksheet sheet = package.Workbook.Worksheets[1];

            // Create list of headers
            List<string> headers = new List<string>();
            // worksheet cells start at 1,1 so add a dummy value for alignment
            headers.Add("__ALIGNER__");

            // Parse header cells
            for (int i = 1; i <= sheet.Dimension.Columns; i++)
            {
                string headerText = sheet.Cells[1, i].Value.ToString();

                // TODO: Maybe throw warning if condition is not met?
                if (Utility.StringHasUsefulData(headerText))
                {
                    headers.Add(headerText);
                }
            }

            List<User> users = new List<User>();

            // Parse user entries
            try
            {
                for (int i = 2; i <= sheet.Dimension.Rows; i++)
                {
                    try
                    {
                        users.Add(new User(sheet, headers, i));
                    }
                    catch (InvalidEntryException e)
                    {
                        Console.WriteLine("WARNING: Invalid Entry Exception occured parsing row: row " + i + ".");
                        Console.WriteLine(e.Message);
                        Console.WriteLine(e.StackTrace);
                        Console.WriteLine("Ignoring user specified on row " + i + ".");
                    }
                }
            }
            catch (InvalidSheetException e)
            {
                Console.WriteLine("ERROR: Tried to convert using an invalid Excel sheet as input!");
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                Console.WriteLine("Stopping conversion.");

                // TODO: Proper exit codes
                Environment.Exit(-2);
            }
            catch (Exception e)
            {
                Console.WriteLine("ERROR: Unknown exception occurred!");
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                Console.WriteLine("Stopping conversion.");

                // TODO: Proper exit codes
                Environment.Exit(-1);
            }

            return users;
        }
    }
}
