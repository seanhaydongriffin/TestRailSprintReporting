
using System.Xml.Linq;

namespace SharedProject.Confluence
{
    public class TableColumnGroup
    {

        private XElement _TableColumnGroup;

        public TableColumnGroup()
        {
            _TableColumnGroup = new XElement("thead");
        }

        public TableColumnGroup(params string[] fields)
        {
            var row = new XElement("tr");

            foreach (var field in fields)

                row.Add(new XElement("th", new XElement("p", field)));

            _TableColumnGroup.Add(new XElement("thead", row));
        }

        public TableColumnGroup(params int[] widths)
        {
            _TableColumnGroup = new XElement("colgroup");

            foreach (var width in widths)

                _TableColumnGroup.Add(new XElement("col", new XAttribute("style", "width:" + width + "px;")));
        }

        public XElement GetXElement()
        {
            return _TableColumnGroup;
        }


    }
}
