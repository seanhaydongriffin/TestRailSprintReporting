
using System.Xml.Linq;

namespace SharedProject.Confluence
{
    public class ChartBody
    {

        public static XNamespace ac = "http://someuri";
        private XElement _ChartBody;

        public ChartBody(SharedProject.Confluence.TableBody chart_table_body)
        {
            _ChartBody = new XElement(ac + "rich-text-body", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                new XElement("table",
                    chart_table_body.GetXElement()
                )
            );
        }

        public XElement GetXElement()
        {
            return _ChartBody;
        }



    }
}
