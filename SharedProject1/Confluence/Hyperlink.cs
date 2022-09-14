
using System.Xml.Linq;

namespace SharedProject.Confluence
{
    public class Hyperlink
    {

        private XElement _Hyperlink;

        public Hyperlink(string url, string text)
        {
            _Hyperlink = new XElement("a", new XAttribute("href", url), text);
        }

        public XElement GetXElement()
        {
            return _Hyperlink;
        }

        public string ToStringDisableFormatting()
        {
            return _Hyperlink.ToString(SaveOptions.DisableFormatting);
        }


    }
}
