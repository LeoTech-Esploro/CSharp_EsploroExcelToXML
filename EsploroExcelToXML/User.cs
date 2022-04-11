using System;
using System.Collections.Generic;
using System.Xml;
using OfficeOpenXml;

namespace EsploroExcelToXML
{
    public class User
    {
        public bool Affiliated;
        public string Email;
        public Researcher Researcher;

        public User(ExcelWorksheet sheet, List<string> headers, int row)
        {
            // Check state assumptions
            if (sheet.Dimension.Columns < 1)
            {
                throw new InvalidSheetException("Sheet is empty");
            }

            if (!headers.Contains(@"Email"))
            {
                throw new InvalidSheetException("Sheet does not contain an 'Email' column");
            }

            if (row < 2 || row > sheet.Dimension.Rows)
            {
                throw new InvalidEntryException("Tried to create user from invalid row: row " + row);
            }

            // Get email from worksheet
            int emailColumn = headers.IndexOf(@"Email");
            string entryEmail = sheet.Cells[row, emailColumn].Value.ToString();

            // Check email is valid
            if (!Utility.StringIsValidEmail(entryEmail))
            {
                throw new InvalidEntryException("Attempted to modify user with invalid email: " + entryEmail);
            }

            if (entryEmail.Length > 255)
            {
                throw new InvalidEntryException("Email is too long: " + entryEmail + ". Limit is 255 characters.");
            }

            this.Email = entryEmail;

            // Check affiliated status
            if (headers.Contains(@"Affiliated/Non-Affiliated"))
            {
                // Get data from cell
                int colIndex = headers.IndexOf(@"Affiliated/Non-Affiliated");
                string data = sheet.Cells[row, colIndex].Value.ToString();

                // Check affiliated status from string text
                if (data == @"Affiliated") this.Affiliated = true;
                else this.Affiliated = false;
            }

            // Get researcher from worksheet
            this.Researcher = new Researcher(sheet, headers, row);
        }

        public void WriteXML(XmlWriter xml)
        {
            xml.WriteStartElement("user");

            xml.WriteElementString("is_researcher", "true");

            xml.WriteStartElement("user_identifiers");

            xml.WriteStartElement("user_identifier");
            xml.WriteAttributeString("segment_type", "External");

            // TODO: Need to get ID types from code table using API and select appropriate one
            xml.WriteElementString("id_type", "Other");
            // TODO: Make sure we have an email here
            xml.WriteElementString("value", this.Email);
            xml.WriteElementString("status", "ACTIVE");

            xml.WriteEndElement(); // user_identifier

            xml.WriteEndElement(); // user_identifiers

            this.Researcher.WriteXML(xml);

            xml.WriteEndElement(); // user
        }
    }
}
