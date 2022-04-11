using System;
using System.Collections.Generic;
using System.Xml;
using OfficeOpenXml;

namespace EsploroExcelToXML
{
    public class Researcher
    {
        public string FirstName;
        public string MiddleName;
        public string LastName;

        public string ProfileURLIdentifier;

        public List<string> Keywords;

        public Researcher(ExcelWorksheet sheet, List<string> headers, int row)
        {
            // Check state assumptions
            if (sheet.Dimension.Columns < 1)
            {
                throw new InvalidSheetException("Sheet is empty");
            }

            if (row < 2 || row > sheet.Dimension.Rows)
            {
                throw new InvalidEntryException("Tried to create researcher from invalid row: row " + row);
            }

            // Parse properties if available
            SetSimpleStrPropertyFromCell(sheet, row, headers, ref this.FirstName,  @"Researcher First Name");
            SetSimpleStrPropertyFromCell(sheet, row, headers, ref this.MiddleName, @"Researcher Middle Name");
            SetSimpleStrPropertyFromCell(sheet, row, headers, ref this.LastName,   @"Researcher Last Name");

            SetSimpleStrPropertyFromCell(sheet, row, headers, ref this.ProfileURLIdentifier, @"Profile URL Identifier");

            if (headers.Contains(@"Keywords"))
            {
                // Get data from cell
                int colIndex = headers.IndexOf(@"Keywords");
                string data = sheet.Cells[row, colIndex].Value.ToString();

                // Verify cell had useful data
                if (Utility.StringHasUsefulData(data))
                {
                    // Split up the string by separator ';' and turn it into a list
                    string[] keywords = data.Split(';');
                    this.Keywords = new List<string>(keywords);
                }
            }
        }

        public void WriteXML(XmlWriter xml)
        {
            xml.WriteStartElement("researcher");

            if (Utility.StringHasUsefulData(this.FirstName))  xml.WriteElementString("researcher_first_name",  this.FirstName);
            if (Utility.StringHasUsefulData(this.MiddleName)) xml.WriteElementString("researcher_middle_name", this.MiddleName);
            if (Utility.StringHasUsefulData(this.LastName))   xml.WriteElementString("researcher_last_name",   this.LastName);

            if (this.Keywords.Count > 1)
            {
                xml.WriteStartElement("researcher_keywords");

                foreach (string keyword in this.Keywords)
                {
                    xml.WriteStartElement("researcher_keyword");
                    xml.WriteElementString("value", keyword);
                    xml.WriteEndElement(); // researcher_keyword
                }

                xml.WriteEndElement(); // researcher_keywords
            }

            if (Utility.StringHasUsefulData(this.ProfileURLIdentifier))
            {
                xml.WriteStartElement("user_identifiers");

                xml.WriteStartElement("user_identifier");
                xml.WriteAttributeString("segment_type", "External"); // TODO: Determine how we know internal vs external

                // TODO: Need to get ID types from code table using API and select appropriate one
                xml.WriteElementString("id_type", "Other");
                xml.WriteElementString("value", this.ProfileURLIdentifier);
                xml.WriteElementString("status", "ACTIVE");

                xml.WriteEndElement(); // user_identifier

                xml.WriteEndElement(); // user_identifiers
            }

            xml.WriteEndElement(); // researcher
        }

        /// <summary>
        /// Shortcut for setting simple string properties.
        /// </summary>
        private void SetSimpleStrPropertyFromCell(ExcelWorksheet sheet, int row, List<string> headers, ref string property, string header)
        {
            if (headers.Contains(header))
            {
                // Get data from cell
                int colIndex = headers.IndexOf(header);
                string data = sheet.Cells[row, colIndex].Value.ToString();

                // Verify cell had useful data
                if (Utility.StringHasUsefulData(data))
                {
                    // Assign to property
                    property = data;
                }
            }
        }
    }
}
