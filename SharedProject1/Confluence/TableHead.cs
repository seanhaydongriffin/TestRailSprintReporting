
using System.Xml.Linq;

namespace SharedProject.Confluence
{
    public class TableHead
    {

        private XElement _TableHead;

        public TableHead()
        {
            _TableHead = new XElement("thead");
        }

        public TableHead(params string[] fields)
        {
            _TableHead = new XElement("thead");

            foreach (var field in fields)

                _TableHead.Add(new XElement("th", new XElement("sub", field)));
        }

        public XElement GetXElement()
        {
            return _TableHead;
        }


    }
}
