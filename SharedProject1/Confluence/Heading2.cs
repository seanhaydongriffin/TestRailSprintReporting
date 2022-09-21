
using System.Xml.Linq;

namespace SharedProject.Confluence
{
    public class Heading2
    {

        private XElement _Heading2;

        public Heading2(string HyperlinkUrl, string HyperlinkText)
        {
            _Heading2 = new XElement("h2",
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
            return _Heading2;
        }


        public string ToStringDisableFormatting()
        {
            return _Heading2.ToString(SaveOptions.DisableFormatting);
        }

    }
}
