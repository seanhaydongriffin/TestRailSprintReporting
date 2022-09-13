
using System.Xml.Linq;

namespace SharedProject.Confluence
{
    public class Section
    {

        public static XNamespace ac = "http://someuri";
        private XElement SectionXelement;

        public Section(string alignment, string width, string border)
        {
            SectionXelement = new XElement(ac + "structured-macro", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                new XAttribute(ac + "name", "section"),
                new XAttribute(ac + "align", alignment),
                new XAttribute("data-layout", width),
                new XElement(ac + "parameter", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                    new XAttribute(ac + "name", "border"),
                    border
                )
            );
        }

        public Section Add(SharedProject.Confluence.SectionColumns columns)
        {
            SectionXelement.Add(new XElement(ac + "rich-text-body", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                columns.GetList()
            ));

            return this;
        }

        public string ToStringDisableFormatting()
        {
            return SectionXelement.ToString(SaveOptions.DisableFormatting);
        }

    }
}
