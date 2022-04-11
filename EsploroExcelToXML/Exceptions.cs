using System;
namespace EsploroExcelToXML
{
    public class InvalidSheetException : Exception
    {
        public InvalidSheetException(string message) : base(message) { }
    }

    public class InvalidEntryException : Exception
    {
        public InvalidEntryException(string message) : base(message) { }
    }
}
