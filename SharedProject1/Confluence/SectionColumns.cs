
using System.Collections.Generic;
using System.Xml.Linq;

namespace SharedProject.Confluence
{
    public class SectionColumns
    {

        public static XNamespace ac = "http://someuri";
        private List<XElement> _SectionColumns;

        public SectionColumns()
        {
            _SectionColumns = new List<XElement>();
        }

        public List<XElement> GetList()
        {
            return _SectionColumns;
        }

        public List<XElement> Add(string alignment, SharedProject.Confluence.Chart chart)
        {
            _SectionColumns.Add(new XElement(ac + "structured-macro", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                new XAttribute(ac + "name", "column"),
                new XAttribute(ac + "align", alignment),
                new XElement(ac + "rich-text-body", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                    chart.GetXElement()
                )
            ));

            return _SectionColumns;
        }



    }
}
