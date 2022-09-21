
using System.Collections.Generic;
using System.Xml.Linq;

namespace SharedProject.Confluence
{
    public class TableRow
    {

        private XElement _TableRow;

        public TableRow()
        {
            _TableRow = new XElement("tr");
        }

        public TableRow(params TableCell[] cells)
        {
            _TableRow = new XElement("tr");

            foreach (var cell in cells)
            
                Add(cell);
        }

        public XElement GetXElement()
        {
            return _TableRow;
        }

        public TableRow Add(TableCell cell)
        {
            _TableRow.Add(new XElement("td", new XElement("sub", cell.GetList())));
            return this;
        }



    }
}
