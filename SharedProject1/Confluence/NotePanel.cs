
using System.Xml.Linq;

namespace SharedProject.Confluence
{
    public class NotePanel
    {

        public static XNamespace ac = "http://someuri";
        private XElement _NotePanel;

        public NotePanel(string text, string hyperlink_text = "", string hyperlink_url = "")
        {
            XElement sub;

            if (hyperlink_url.Equals(""))

                sub = new XElement("sub", text);
            else

                sub = new XElement("sub", text, " ", new XElement("a", new XAttribute("href", hyperlink_url), hyperlink_text));

            _NotePanel = new XElement(ac + "adf-extension", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                new XElement(ac + "adf-node", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                    new XAttribute("type", "panel"),
                    new XElement(ac + "adf-attribute", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                        new XAttribute("key", "panel-type"),
                        "note"
                    ),
                    new XElement(ac + "adf-content", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                        new XElement("p",
                            sub
                        )
                    )
                )
            );
        }

        public XElement GetXElement()
        {
            return _NotePanel;
        }

        public string ToStringDisableFormatting()
        {
            return _NotePanel.ToString(SaveOptions.DisableFormatting);
        }


    }
}
