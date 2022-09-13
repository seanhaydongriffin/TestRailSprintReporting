
using System.Collections.Generic;
using System.Xml.Linq;

namespace SharedProject.Confluence
{
    public class TableRows
    {

        private List<XElement> _TableRows;

        public TableRows()
        {
            _TableRows = new List<XElement>();
        }

        public List<XElement> GetList()
        {
            return _TableRows;
        }

        public TableRows Add(XElement row)
        {
            _TableRows.Add(row);
            return this;
        }

        public TableRows AddHeader(params string[] fields)
        {
            var row = new XElement("tr");

            foreach (var field in fields)

                row.Add(new XElement("th", new XElement("p", field)));

            _TableRows.Add(row);
            return this;
        }

        public TableRows Add(params string[] fields)
        {
            var row = new XElement("tr");

            foreach (var field in fields)

                row.Add(new XElement("td", new XElement("p", field)));

            _TableRows.Add(row);
            return this;
        }

        public TableRows Add(params object[] fields)
        {
            var row = new XElement("tr");

            foreach (var field in fields)

                row.Add(new XElement("td", new XElement("sub", field)));

            _TableRows.Add(row);
            return this;
        }


    }
}
