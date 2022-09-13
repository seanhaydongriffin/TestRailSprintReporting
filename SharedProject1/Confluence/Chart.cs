
using System;
using System.Xml.Linq;

namespace SharedProject.Confluence
{
    public class Chart
    {

        public static XNamespace ac = "http://someuri";
        private XElement _Chart;

        public Chart(SharedProject.Confluence.ChartBody chart_body, string subTitle, string type, string stacked, int width, int height, string legend, string rangeAxisTickUnit, string categoryLabelPosition, string dataOrientation, string colors, string pieSectionLabel, int? max_not_passed = 0)
        {
            if (rangeAxisTickUnit.Equals(""))

                rangeAxisTickUnit = Math.Round(Math.Pow(2, (double)(((double)max_not_passed / (double)10) - ((double)height / 100) + 1) * 0.8), 0, MidpointRounding.AwayFromZero).ToString();

            _Chart = new XElement(ac + "structured-macro", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                new XAttribute(ac + "name", "chart"),
                new XElement(ac + "parameter", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                    new XAttribute(ac + "name", "subTitle"), subTitle
                ),
                new XElement(ac + "parameter", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                    new XAttribute(ac + "name", "type"), type
                ),
                new XElement(ac + "parameter", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                    new XAttribute(ac + "name", "stacked"), stacked
                ),
                new XElement(ac + "parameter", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                    new XAttribute(ac + "name", "width"), width
                ),
                new XElement(ac + "parameter", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                    new XAttribute(ac + "name", "height"), height
                ),
                new XElement(ac + "parameter", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                    new XAttribute(ac + "name", "legend"), legend
                ),
                new XElement(ac + "parameter", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                    new XAttribute(ac + "name", "rangeAxisTickUnit"), rangeAxisTickUnit
                ),
                new XElement(ac + "parameter", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                    new XAttribute(ac + "name", "categoryLabelPosition"), categoryLabelPosition
                ),
                new XElement(ac + "parameter", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                    new XAttribute(ac + "name", "dataOrientation"), dataOrientation
                ),
                new XElement(ac + "parameter", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                    new XAttribute(ac + "name", "colors"), colors
                ),
                new XElement(ac + "parameter", new XAttribute(XNamespace.Xmlns + "ac", ac.NamespaceName),
                    new XAttribute(ac + "name", "pieSectionLabel"), pieSectionLabel
                ),
                chart_body.GetXElement()
            );
        }

        public XElement GetXElement()
        {
            return _Chart;
        }

        public string ToStringDisableFormatting()
        {
            return _Chart.ToString(SaveOptions.DisableFormatting);
        }

    }
}
