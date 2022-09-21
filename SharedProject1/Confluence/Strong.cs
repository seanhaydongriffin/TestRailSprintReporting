
using System.Xml.Linq;

namespace SharedProject.Confluence
{
    public class Strong
    {

        private XElement _Strong;

        public Strong(string text)
        {
            _Strong = new XElement("strong", text);
        }

        public XElement GetXElement()
        {
            return _Strong;
        }


    }
}
