
using System.Collections.Generic;
using System.Xml.Linq;

namespace SharedProject.Confluence
{
    public class TableCell
    {

        private List<object> _TableCell;

        public TableCell()
        {
            _TableCell = new List<object>();
        }

        public TableCell(params object[] fields)
        {
            _TableCell = new List<object>();

            foreach (var field in fields)
            {
                var tmp_field = field;

                if (field.GetType() == typeof(Anchor))

                    tmp_field = ((Anchor)field).GetXElement();

                if (field.GetType() == typeof(Hyperlink))

                    tmp_field = ((Hyperlink)field).GetXElement();

                _TableCell.Add(tmp_field);
            }
        }

        public TableCell Add(object field)
        {
            if (field.GetType() == typeof(Anchor))

                field = ((Anchor)field).GetXElement();

            if (field.GetType() == typeof(Hyperlink))

                field = ((Hyperlink)field).GetXElement();

            _TableCell.Add(field);
            return this;
        }

        public List<object> GetList()
        {
            return _TableCell;
        }

        public int GetCount()
        {
            return _TableCell.Count;
        }


    }
}
