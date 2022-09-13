
using System.Xml.Linq;

namespace SharedProject.Confluence
{
    public class Table
    {

        private XElement _Table;

        public Table(string data_layout)
        {
            _Table = new XElement("table", new XAttribute("data-layout", data_layout));
        }

        public XElement GetXElement()
        {
            return _Table;
        }

        public Table Add(SharedProject.Confluence.TableColumnGroup column_group)
        {
            _Table.Add(column_group.GetXElement());
            return this;
        }

        public Table Add(SharedProject.Confluence.TableHead head)
        {
            _Table.Add(head.GetXElement());
            return this;
        }

        public Table Add(SharedProject.Confluence.TableBody body)
        {
            _Table.Add(body.GetXElement());
            return this;
        }

        public string ToStringDisableFormatting()
        {
            return _Table.ToString(SaveOptions.DisableFormatting);
        }


    }
}
