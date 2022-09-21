
using System.Xml.Linq;

namespace SharedProject.Confluence
{
    public class Toc
    {

        public static XNamespace ac = "http://someuri";
        private XElement _Toc;

        public Toc(int minLevel, int maxLevel)
        {
            _Toc = new XElement(ac + "structured-macro", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                new XAttribute(ac + "name", "toc"),
                new XAttribute(ac + "schema-version", "1"),
                new XAttribute(ac + "data-layout", "default"),
                new XElement(ac + "parameter", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                    new XAttribute(ac + "name", "minLevel"),
                    minLevel
                ),
                new XElement(ac + "parameter", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                    new XAttribute(ac + "name", "maxLevel"),
                    maxLevel
                ),
                new XElement("hr")
            );
        }

        public XElement GetXElement()
        {
            return _Toc;
        }

        public string ToStringDisableFormatting()
        {
            return _Toc.ToString(SaveOptions.DisableFormatting);
        }

    }
}
