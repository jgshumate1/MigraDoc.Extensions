using HtmlAgilityPack;
using MigraDoc.DocumentObjectModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using MigraDoc.DocumentObjectModel.Tables;

namespace MigraDoc.Extensions.Html
{
    public class HtmlConverter : IConverter
    {
        private IDictionary<string, Type> _mapping = new Dictionary<string, Type>();  
        private IDictionary<string, Func<HtmlNode, ExCSS.Stylesheet, DocumentObject, DocumentObject>> nodeHandlers
            = new Dictionary<string, Func<HtmlNode, ExCSS.Stylesheet, DocumentObject, DocumentObject>>();

        private ExCSS.Stylesheet _sheet;

        public HtmlConverter()
        {
            AddDefaultNodeHandlers();
            MapTypes();
        }

        public IDictionary<string, Func<HtmlNode, ExCSS.Stylesheet, DocumentObject, DocumentObject>> NodeHandlers
        {
            get
            {
                return nodeHandlers;
            }
        }

        public Action<Section> Convert(ExCSS.Stylesheet sheet, string contents)
        {
            return section => ConvertHtml(sheet, contents, section);
        }

        private void MapTypes()
        {
            _mapping.Add("div", typeof(ReportHandler));
        }

        private void ConvertHtml(ExCSS.Stylesheet sheet, string html, Section section)
        {
            _sheet = sheet;
            if (string.IsNullOrEmpty(html))
            {
                throw new ArgumentNullException("html");
            }

            if (section == null)
            {
                throw new ArgumentNullException("section");
            }

            //section.PageSetup.HeaderDistance = "0.001cm";
            section.PageSetup.FooterDistance = Unit.FromCentimeter(0.01);

 // Create a paragraph with centered page number. See definition of style "Footer".
            var footer = section.Footers.Primary.AddParagraph();
            //section.Footers.Primary.
            footer.Format.Alignment = ParagraphAlignment.Right;
            footer.AddPageField();
            footer.AddText(" of ");
            footer.AddNumPagesField();
            
            var doc = new HtmlDocument();
            doc.LoadHtml(html);
            ConvertHtmlNodes(doc.DocumentNode.ChildNodes, sheet, section);
        }

        private void ConvertHtmlNodes(HtmlNodeCollection nodes, ExCSS.Stylesheet sheet, DocumentObject section, DocumentObject current = null)
        {
            foreach (var node in nodes)
            {
                Func<HtmlNode, ExCSS.Stylesheet, DocumentObject, DocumentObject> nodeHandler;

                Type t;
                if (_mapping.TryGetValue(node.Name, out t))
                {
                    var instance = (INodeHandler)Activator.CreateInstance(t);
                    var result = instance.NodeHandler(node, sheet, section);

                    if (node.HasChildNodes)
                    {
                        ConvertHtmlNodes(node.ChildNodes, sheet, section, result);
                    }
                }
                else
                {
                    if (nodeHandlers.TryGetValue(node.Name, out nodeHandler))
                    {
                        // pass the current container or section
                        var result = nodeHandler(node, sheet, current ?? section);

                        if (node.HasChildNodes)
                        {
                            ConvertHtmlNodes(node.ChildNodes, sheet, section, result);
                        }
                    }
                    else
                    {
                        if (node.HasChildNodes)
                        {
                            ConvertHtmlNodes(node.ChildNodes, sheet, section, current);
                        }
                    }
                }
            }
        }

        private void AddDefaultNodeHandlers()
        {
            var headerIndex = -1;
            var cellIndex = -1;
            // Block Elements

            //nodeHandlers.Add("div", (node, sheet, parent) =>
            //    {
            //        var handler = new ReportHandler();
            //        return handler.NodeHandler(node, sheet, parent);
            //    });

            var nsheet = _sheet;

            // could do with a predicate/regex matcher so we could just use one handler for all headings
            nodeHandlers.Add("h1", AddHeading);
            nodeHandlers.Add("h2", AddHeading);
            nodeHandlers.Add("h3", AddHeading);
            nodeHandlers.Add("h4", AddHeading);
            nodeHandlers.Add("h5", AddHeading);
            nodeHandlers.Add("h6", AddHeading);

            nodeHandlers.Add("p", (node, sheet, parent) =>
            {
                return ((Section)parent).AddParagraph();
            });

            // Inline Elements

            nodeHandlers.Add("strong", (node, sheet, parent) => AddFormattedText(node, sheet, parent, TextFormat.Bold));
            nodeHandlers.Add("i", (node, sheet, parent) => AddFormattedText(node, sheet, parent, TextFormat.Italic));
            nodeHandlers.Add("em", (node, sheet, parent) => AddFormattedText(node, sheet, parent, TextFormat.Italic));
            nodeHandlers.Add("u", (node, sheet, parent) => AddFormattedText(node, sheet, parent, TextFormat.Underline));
            nodeHandlers.Add("a", (node, sheet, parent) =>
            {
                return GetParagraph(parent).AddHyperlink(node.GetAttributeValue("href", ""), HyperlinkType.Web);
            });
            nodeHandlers.Add("hr", (node, sheet, parent) => GetParagraph(parent).SetStyle("HorizontalRule"));
            nodeHandlers.Add("br", (node, sheet, parent) =>
            {
                if (parent is FormattedText)
                {
                    // inline elements can contain line breaks
                    ((FormattedText)parent).AddLineBreak();
                    return parent;
                }

                var paragraph = GetParagraph(parent);
                paragraph.AddLineBreak();
                return paragraph;
            });


            nodeHandlers.Add("table", (node, sheet, parent) =>
            {
                Table table = null;

                if (parent is Section)
                {
                    var s = (Section)parent;
                    table = s.AddTable();
                }

                if (parent is Document)
                {
                    var d = (Document)parent;
                    var s = (Section)d.Sections.LastObject;
                    table = s.AddTable();
                }

                table.Borders.Width = Unit.FromCentimeter(0.05);
                var margins = node.Attributes["data-padding"];
                if (margins == null)
                {
                    table.TopPadding = Unit.FromCentimeter(0.1);
                    table.TopPadding = Unit.FromCentimeter(0.1);
                    table.TopPadding = Unit.FromCentimeter(0.1);
                    table.TopPadding = Unit.FromCentimeter(0.1);
                }
                else
                {
                    var marginValues = margins.Value.Split(' ');

                    table.TopPadding = Unit.FromCentimeter(double.Parse(marginValues[0]));
                    table.RightPadding = Unit.FromCentimeter(double.Parse(marginValues[1]));
                    table.BottomPadding = Unit.FromCentimeter(double.Parse(marginValues[2]));
                    table.LeftPadding = Unit.FromCentimeter(double.Parse(marginValues[3]));
                }

                return table;
            });

            nodeHandlers.Add("thead", (node, sheet, parent) =>
                {
                    var @class = node.Attributes["class"];
                    var size = node.Attributes["data-size"];
                // we need to create columns before the rows are created
                // so find out how many columns we have
                if (node.HasChildNodes)
                {
                    foreach (var childNode in node.ChildNodes)
                    {
                        if (childNode.Name == "tr")
                        {
                            foreach (var theads in childNode.ChildNodes)
                            {
                                if (theads.Name == "th")
                                {
                                    var t = (Table)parent;
                                    var width = theads.Attributes["data-width"];

                                    var widthUnit = Unit.FromCentimeter(2);
                                    if (width != null)
                                    {
                                        widthUnit = Unit.FromCentimeter(double.Parse(width.Value));
                                    }

                                    var column = t.AddColumn(widthUnit);
                                    column.Format.Alignment = ParagraphAlignment.Center;
                                }
                            }
                        }
                    }
                }

                return parent;
            });

            nodeHandlers.Add("tr", (node, sheet, parent) =>
            {
                if (parent is Table)
                {
                    // have to add columns before rows are added
                    if (node.ParentNode.Name == "thead")
                    {
                        return parent;
                    }

                    var t = (Table)parent;
                    Row row = t.AddRow();
                    //row.HeadingFormat = true;
                    cellIndex = -1;
                    return row;
                }

                return parent;
            });

            nodeHandlers.Add("tbody", (node, sheet, parent) =>
            {
                return parent;
            });

            nodeHandlers.Add("th", (node, sheet, parent) =>
            {
                if (parent is Table)
                {
                    var t = (Table)parent;

                    if (headerIndex == -1)
                    {
                        headerIndex = 0;
                    }
                    else
                    {
                        headerIndex++;
                    }

                    var r = (Row)t.Rows.LastObject ?? t.AddRow();
                    r.HeadingFormat = true;

                    var c = r.Cells[headerIndex];
                    return c;
                }

                return parent;
            });

            nodeHandlers.Add("td", (node, sheet, parent) =>
            {
                if (parent is Row)
                {
                    if (cellIndex == -1)
                    {
                        cellIndex = 0;
                    }

                    var tdRow = (Row)parent;
                    var count = tdRow.Cells.Count;
                    var c = tdRow.Cells[cellIndex];
                    cellIndex++;
                    return c;
                }

                return parent;
            });

            nodeHandlers.Add("li", (node, sheet, parent) =>
            {
                var listStyle = node.ParentNode.Name == "ul"
                    ? "UnorderedList"
                    : "OrderedList";

                var section = (Section)parent;
                var isFirst = node.ParentNode.Elements("li").First() == node;
                var isLast = node.ParentNode.Elements("li").Last() == node;

                // if this is the first item add the ListStart paragraph
                if (isFirst)
                {
                    section.AddParagraph().SetStyle("ListStart");
                }

                var listItem = section.AddParagraph().SetStyle(listStyle);

                // disable continuation if this is the first list item
                listItem.Format.ListInfo.ContinuePreviousList = !isFirst;

                // if the this is the last item add the ListEnd paragraph
                if (isLast)
                {
                    section.AddParagraph().SetStyle("ListEnd");
                }

                return listItem;
            });

            nodeHandlers.Add("#text", (node, sheet, parent) =>
            {
                // remove line breaks
                var innerText = node.InnerText.Replace("\r", "").Replace("\n", "");

                if (string.IsNullOrWhiteSpace(innerText))
                {
                    return parent;
                }

                // decode escaped HTML
                innerText = WebUtility.HtmlDecode(innerText);

                // text elements must be wrapped in a paragraph but this could also be FormattedText or a Hyperlink!!
                // this needs some work
                if (parent is FormattedText)
                {
                    return ((FormattedText)parent).AddText(innerText);
                }
                if (parent is Hyperlink)
                {
                    return ((Hyperlink)parent).AddText(innerText);
                }
                if (parent is Cell)
                {
                    return ((Cell)parent).AddParagraph(innerText);
                }
                if (parent is Column)
                {
                    var c = (Column)parent;
                    foreach (var column in c.Table.Columns)
                    {
                        if (c == column)
                        {
                            var t = c.Table;
                            var r = (Row)t.Rows.LastObject;
                            var cell = r.Cells[c.Index];
                            cell.AddParagraph(innerText);
                        }
                    }

                    return parent;
                }

                if (parent is Row)
                {
                    return parent;
                }

                // otherwise a section or paragraph
                return GetParagraph(parent).AddText(innerText);
            });
        }

        private static DocumentObject AddFormattedText(HtmlNode node, ExCSS.Stylesheet sheet, DocumentObject parent, TextFormat format)
        {
            var formattedText = parent as FormattedText;
            if (formattedText != null)
            {
                return formattedText.Format(format);
            }

            // otherwise parent is paragraph or section
            return GetParagraph(parent).AddFormattedText(format);
        }

        private static DocumentObject AddHeading(HtmlNode node, ExCSS.Stylesheet sheet, DocumentObject parent)
        {
            return ((Section)parent).AddParagraph().SetStyle("Heading" + node.Name[1]);
        }

        private static Paragraph GetParagraph(DocumentObject parent)
        {
            return parent as Paragraph ?? ((Section)parent).AddParagraph();
        }

        private static Paragraph AddParagraphWithStyle(DocumentObject parent, string style)
        {
            return ((Section)parent).AddParagraph().SetStyle(style);
        }
    }
}
