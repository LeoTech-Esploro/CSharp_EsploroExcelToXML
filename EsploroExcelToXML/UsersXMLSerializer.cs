using System;
using System.Collections.Generic;
using System.Xml;

namespace EsploroExcelToXML
{
    public partial class Converter
    {
        public static void SerializeUsersAsXMLToFile(List<User> users, string filePath)
        {
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Async = false;
            settings.Indent = true; // make resulting xml human-readable

            XmlWriter xml = XmlWriter.Create(filePath, settings);

            xml.WriteStartElement("users");

            foreach (User user in users)
            {
                user.WriteXML(xml);
            }

            xml.WriteEndElement(); // users

            xml.Flush();
            xml.Close();
        }
    }
}
