
using System.Xml.Linq;

namespace SharedProject.Confluence
{
    public class Bold
    {

        private XElement _Bold;

        public Bold(string text = "")
        {
            _Bold = new XElement("b", text);
        }

        public XElement GetXElement()
        {
            return _Bold;
        }


    }
}
