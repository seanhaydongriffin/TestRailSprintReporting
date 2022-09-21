
using System.Xml.Linq;

namespace SharedProject.Confluence
{
    public class Status
    {

        public static XNamespace ac = "http://someuri";
        private XElement _Status;

        public Status(string text)
        {
            _Status = new XElement(ac + "structured-macro", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                new XAttribute(ac + "name", "status"),
                new XElement(ac + "parameter", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                    new XAttribute(ac + "name", "title"),
                    text
                )
            );
        }

        public XElement GetXElement()
        {
            return _Status;
        }

        public string ToStringDisableFormatting()
        {
            return _Status.ToString(SaveOptions.DisableFormatting);
        }


    }
}
