
using System.Xml.Linq;

namespace SharedProject.Confluence
{
    public class Emoticon
    {

        public static XNamespace ac = "http://someuri";
        private XElement _Emoticon;

        public Emoticon(string name, string emoji_shortname, string emoji_id, string emoji_fallback)
        {
            _Emoticon = new XElement(ac + "emoticon", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName));

            if (!name.Equals(""))

                _Emoticon.Add(new XAttribute(ac + "name", name));

            if (!emoji_shortname.Equals(""))

                _Emoticon.Add(new XAttribute(ac + "emoji-shortname", emoji_shortname));

            if (!emoji_id.Equals(""))

                _Emoticon.Add(new XAttribute(ac + "emoji-id", emoji_id));

            if (!emoji_fallback.Equals(""))

                _Emoticon.Add(new XAttribute(ac + "emoji-fallback", emoji_fallback));

        }

        public XElement GetXElement()
        {
            return _Emoticon;
        }


    }
}
