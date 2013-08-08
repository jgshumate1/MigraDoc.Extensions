using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MigraDoc.DocumentObjectModel;
using HtmlAgilityPack;

namespace MigraDoc.Extensions.Html.Example
{
    public interface INodeHandler
    {
        DocumentObject NodeHandler(HtmlNode node, DocumentObject parent, DocumentObject current = null);
    }

    public class ReportHandler : INodeHandler
    {
        public DocumentObject NodeHandler(HtmlNode node, DocumentObject parent, DocumentObject current = null)
        {
            var @class = node.Attributes["class"];

            if (!(parent is Section))
            {
                return parent;
            }

            var sec = (Section)parent;

            if (@class != null && @class.Value == "report")
            {
                var margins = node.Attributes["data-margin"];
                if (margins == null)
                {
                    sec.PageSetup.TopMargin = Unit.FromCentimeter(0.1);
                    sec.PageSetup.BottomMargin = Unit.FromCentimeter(0.1);
                    sec.PageSetup.LeftMargin = Unit.FromCentimeter(0.1);
                    sec.PageSetup.RightMargin = Unit.FromCentimeter(0.1);
                }
                else
                {
                    var marginValues = margins.Value.Split(' ');

                    sec.PageSetup.TopMargin = Unit.FromCentimeter(double.Parse(marginValues[0]));
                    sec.PageSetup.RightMargin = Unit.FromCentimeter(double.Parse(marginValues[1]));
                    sec.PageSetup.BottomMargin = Unit.FromCentimeter(double.Parse(marginValues[2]));
                    sec.PageSetup.LeftMargin = Unit.FromCentimeter(double.Parse(marginValues[3]));
                }
            }

            return parent;
        }
    }
}
