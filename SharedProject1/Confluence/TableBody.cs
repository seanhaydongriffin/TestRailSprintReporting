
using System.Xml.Linq;

namespace SharedProject.Confluence
{
    public class TableBody
    {

        private XElement TableBodyXelement;

        public TableBody()
        {
            TableBodyXelement = new XElement("tbody");
        }

        public XElement GetXElement()
        {
            return TableBodyXelement;
        }

        public TableBody(SharedProject.Confluence.TableRows rows)
        {
            TableBodyXelement = new XElement("tbody", rows.GetList());
        }

        public XElement AddChartTableHeading(params string[] table_headings)
        {
            var table_heading_row = new XElement("tr");

            foreach (var heading in table_headings)

                table_heading_row.Add(new XElement("th", new XElement("p", heading)));

            TableBodyXelement.Add(table_heading_row);
            return TableBodyXelement;
        }


    }
}
