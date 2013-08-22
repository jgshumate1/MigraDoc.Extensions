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
            
            Unit margin;
            SetMargin(node, sheet, out margin, "margin-top");
            sec.PageSetup.TopMargin = margin;
            SetMargin(node, sheet, out margin, "margin-bottom");
            sec.PageSetup.BottomMargin = margin;
            SetMargin(node, sheet, out margin, "margin-left");
            sec.PageSetup.LeftMargin = margin;
            SetMargin(node, sheet, out margin, "margin-right");
            sec.PageSetup.LeftMargin = margin;
            return parent;
        }

        private void SetMargin(HtmlNode node, ExCSS.Stylesheet sheet, out Unit margin, string cssMargin)
        {
            var @class = node.Attributes["class"];

            margin = Unit.FromCentimeter(0.2);
            
            if (@class != null)
            {
                var css = (from t in sheet.RuleSets
                           where t.Selectors.Any(x => x.SimpleSelectors.Any(z => z.Class == @class.Value))
                           select t).FirstOrDefault();

                if (css == null)
                {
                    
                }

                if (css != null)
                {
                    var mar = (from top in css.Declarations
                                  where top.Name == cssMargin
                                  select top).FirstOrDefault();

                    if (mar != null)
                    {
                        var tm = mar.Expression.Terms.Where(x => x.Value != string.Empty).FirstOrDefault();

                        if (tm != null)
                        {
                           margin = Unit.FromCentimeter(double.Parse(tm.Value));
                        }
                    }
                }
            }
        }
    }
}
