
using System.Xml.Linq;

namespace SharedProject.Confluence
{
    public class Heading1
    {

        private XElement _Heading1;

        public Heading1(string HyperlinkUrl, string HyperlinkText)
        {
            _Heading1 = new XElement("h1",
                new XElement("u",
                    new XElement("a",
                        new XAttribute("href", HyperlinkUrl),
                        HyperlinkText
                    )
                )
            );
        }

        public XElement GetXElement()
        {
            return _Heading1;
        }


        public string ToStringDisableFormatting()
        {
            return _Heading1.ToString(SaveOptions.DisableFormatting);
        }

    }
}
