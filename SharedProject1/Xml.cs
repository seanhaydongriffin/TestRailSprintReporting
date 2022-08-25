using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;

namespace SharedProject
{
    public static class Xml
    {

        public static XmlNode GetNode(String xml_filename, String xpath)
        {
            try
            {
                var XmlDoc = new XmlDocument();
                XmlDoc.Load(xml_filename);
                return XmlDoc.SelectSingleNode(xpath);
            }
            catch (Exception)
            {
            }

            return null;
        }

        public static XmlNodeList GetChildNodes(String xml_filename, String xpath)
        {
            try
            {
                var XmlDoc = new XmlDocument();
                XmlDoc.Load(xml_filename);
                return XmlDoc.SelectNodes(xpath);
            }
            catch (Exception)
            {
            }

            return null;
        }

        public static string GetChildNodes(this XmlNode Node, String AttributeName)
        {
            try
            {
                return Node.Attributes[AttributeName].Value;
            }
            catch (Exception)
            {
            }

            return null;
        }

        public static XmlNode GetChildNode(this XmlNode Node, String xpath)
        {
            try
            {
                return Node.SelectSingleNode(xpath);
            }
            catch (Exception)
            {
            }

            return null;
        }

        public static string GetAttributeValue(this XmlNode Node, String AttributeName)
        {
            try
            {
                return Node.Attributes[AttributeName].Value;
            }
            catch (Exception)
            {
            }

            return null;
        }

    }
}
