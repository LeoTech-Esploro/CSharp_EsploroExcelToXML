using System;
namespace EsploroExcelToXML
{
    /// <summary>
    /// Contains various helper methods for use in the codebase.
    /// </summary>
    public class Utility
    {
        /// <summary>
        /// Tries to determine if the string has useful data, or if it only contains whitespace.
        /// </summary>
        public static bool StringHasUsefulData(string str)
        {
            // Quick test for empty strings
            if (str == null || str == "") return false;

            // Create a copy of the string we can modify
            string testStr = String.Copy(str);

            // Remove common whitespace characters from the test string
            testStr.Replace(" ", String.Empty);
            testStr.Replace("\t", String.Empty);
            testStr.Replace("\n", String.Empty);

            // If the test string is now empty, it didn't contain useful data
            if (testStr == "") return false;

            return true;
        }

        /// <summary>
        /// Tries to determine if the string is a valid email.
        /// </summary>
        public static bool StringIsValidEmail(string str)
        {
            // Quick test for empty/useless strings
            if (!StringHasUsefulData(str)) return false;

            // TODO: This is NOT a robust test for a valid email. Replace this.

            // For now, just test and make sure the string at least contains the characters
            // '@' and '.' in that order, and is at least 5 characters long (a@b.c => 5 characters).

            if (str.Length < 5) return false;

            if (!str.Contains("@")) return false;
            if (!str.Contains(".")) return false;

            if (!(str.IndexOf('.') > str.IndexOf('@'))) return false;

            return true;
        }
    }
}
