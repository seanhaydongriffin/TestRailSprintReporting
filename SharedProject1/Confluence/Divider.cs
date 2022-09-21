
using System.Xml.Linq;

namespace SharedProject.Confluence
{
    public class Divider
    {

        private XElement _Divider;

        public Divider(string name)
        {
            _Divider = new XElement("hr");
        }

        public XElement GetXElement()
        {
            return _Divider;
        }


    }
}
