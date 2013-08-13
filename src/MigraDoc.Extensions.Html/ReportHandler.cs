using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExCSS.Model;
using MigraDoc.DocumentObjectModel;
using HtmlAgilityPack;
using Unit = MigraDoc.DocumentObjectModel.Unit;

namespace MigraDoc.Extensions.Html
{
    public interface INodeHandler
    {
        DocumentObject NodeHandler(HtmlNode node, ExCSS.Stylesheet sheet, DocumentObject parent, DocumentObject current = null);
    }

    public class ReportHandler : INodeHandler
    {
        public DocumentObject NodeHandler(HtmlNode node, ExCSS.Stylesheet sheet, DocumentObject parent, DocumentObject current = null)
        {
            if (!(parent is Section))
            {
                return parent;
            }

            var sec = (Section)parent;

            var @class = node.Attributes["class"];

            if (@class != null)
            {
                var css = (from t in sheet.RuleSets
                        where t.Selectors.Any(x => x.SimpleSelectors.Any(z => z.Class == @class.Value))
                        select t).FirstOrDefault();

                if (css != null)
                {
                    var topmar = (from top in css.Declarations
                               where top.Name == "margin-top"
                               select top).FirstOrDefault();

                    var tm = topmar.Expression.Terms.Select(x => x.Value != String.Empty);
                }
            }
            else
            {

            }

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
