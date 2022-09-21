
using System.Xml.Linq;

namespace SharedProject.Confluence
{
    public class CustomPanel
    {

        public static XNamespace ac = "http://someuri";
        private XElement _CustomPanel;

        public CustomPanel(string panel_icon, string panel_icon_id, string background_color = "#EAE6FF", string text = "", string hyperlink_text = "", string hyperlink_url = "")
        {
            XElement sub;

            if (hyperlink_url.Equals(""))

                sub = new XElement("sub", text);
            else

                sub = new XElement("sub", text, " ", new XElement("a", new XAttribute("href", hyperlink_url), hyperlink_text));

            _CustomPanel = new XElement(ac + "structured-macro", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                new XAttribute(ac + "name", "panel"),
                new XElement(ac + "parameter", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                    new XAttribute(ac + "name", "panelIcon"),
                    panel_icon
                ),
                new XElement(ac + "parameter", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                    new XAttribute(ac + "name", "panelIconId"),
                    panel_icon_id
                ),
                new XElement(ac + "parameter", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                    new XAttribute(ac + "name", "bgColor"),
                    background_color
                ),
                new XElement(ac + "rich-text-body", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                    new XElement("p",
                        sub
                    )
                )
            );
        }

        public XElement GetXElement()
        {
            return _CustomPanel;
        }

        public string ToStringDisableFormatting()
        {
            return _CustomPanel.ToString(SaveOptions.DisableFormatting);
        }


    }
}
