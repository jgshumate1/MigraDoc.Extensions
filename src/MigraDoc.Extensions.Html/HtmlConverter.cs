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
        private IDictionary<string, Func<HtmlNode, DocumentObject, DocumentObject>> nodeHandlers
            = new Dictionary<string, Func<HtmlNode, DocumentObject, DocumentObject>>();

        public HtmlConverter()
        {
            AddDefaultNodeHandlers();
        }

        public IDictionary<string, Func<HtmlNode, DocumentObject, DocumentObject>> NodeHandlers
        {
            get
            {
                return nodeHandlers;
            }
        }

        public Action<Section> Convert(string contents)
        {
            return section => ConvertHtml(contents, section);
        }

        private void ConvertHtml(string html, Section section)
        {
            if (string.IsNullOrEmpty(html))
            {
                throw new ArgumentNullException("html");
            }

            if (section == null)
            {
                throw new ArgumentNullException("section");
            }

            section.PageSetup.HeaderDistance = "0.002cm";
            section.PageSetup.FooterDistance = "0.002cm";

            // Create a paragraph with centered page number. See definition of style "Footer".
            var footer = section.Footers.Primary.AddParagraph();
            footer.Format.Alignment = ParagraphAlignment.Right;
            footer.AddPageField();
            var doc = new HtmlDocument();
            doc.LoadHtml(html);
            ConvertHtmlNodes(doc.DocumentNode.ChildNodes, section);
        }

        private void ConvertHtmlNodes(HtmlNodeCollection nodes, DocumentObject section, DocumentObject current = null)
        {
            foreach (var node in nodes)
            {
                Func<HtmlNode, DocumentObject, DocumentObject> nodeHandler;
                if (nodeHandlers.TryGetValue(node.Name, out nodeHandler))
                {
                    // pass the current container or section
                    var result = nodeHandler(node, current ?? section);

                    //var pageSize = current != null ? current.Section.Document.DefaultPageSetup.PageHeight :
                    //    section.Document.DefaultPageSetup.PageHeight;

                    //for (int i = 0; i < section.Document.Sections[0].Elements.Count; i++)
                    //{
                    //    var element = section.Document.Sections[0].Elements[i].GetType();
                    //    var x = element.FullName;
                    //}

                    if (node.HasChildNodes)
                    {
                        ConvertHtmlNodes(node.ChildNodes, section, result);
                    }
                }
                else
                {
                    if (node.HasChildNodes)
                    {
                        ConvertHtmlNodes(node.ChildNodes, section, current);
                    }
                }
            }
        }

        private void AddDefaultNodeHandlers()
        {
            var headerIndex = -1;
            var cellIndex = -1;
            // Block Elements

            // could do with a predicate/regex matcher so we could just use one handler for all headings
            nodeHandlers.Add("h1", AddHeading);
            nodeHandlers.Add("h2", AddHeading);
            nodeHandlers.Add("h3", AddHeading);
            nodeHandlers.Add("h4", AddHeading);
            nodeHandlers.Add("h5", AddHeading);
            nodeHandlers.Add("h6", AddHeading);

            nodeHandlers.Add("p", (node, parent) =>
            {
                return ((Section)parent).AddParagraph();
            });

            // Inline Elements

            nodeHandlers.Add("strong", (node, parent) => AddFormattedText(node, parent, TextFormat.Bold));
            nodeHandlers.Add("i", (node, parent) => AddFormattedText(node, parent, TextFormat.Italic));
            nodeHandlers.Add("em", (node, parent) => AddFormattedText(node, parent, TextFormat.Italic));
            nodeHandlers.Add("u", (node, parent) => AddFormattedText(node, parent, TextFormat.Underline));
            nodeHandlers.Add("a", (node, parent) =>
            {
                return GetParagraph(parent).AddHyperlink(node.GetAttributeValue("href", ""), HyperlinkType.Web);
            });
            nodeHandlers.Add("hr", (node, parent) => GetParagraph(parent).SetStyle("HorizontalRule"));
            nodeHandlers.Add("br", (node, parent) =>
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


            nodeHandlers.Add("table", (node, parent) =>
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

                table.Borders.Width = Unit.FromCentimeter(0.075);

                return table;
            });

            nodeHandlers.Add("thead", (node, parent) =>
            {
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
                                    var column = t.AddColumn(Unit.FromCentimeter(5));
                                    column.Format.Alignment = ParagraphAlignment.Center;
                                }
                            }
                        }
                    }
                }

                return parent;
            });

            nodeHandlers.Add("tr", (node, parent) =>
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

            nodeHandlers.Add("tbody", (node, parent) =>
            {
                return parent;
            });

            nodeHandlers.Add("th", (node, parent) =>
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

            nodeHandlers.Add("td", (node, parent) =>
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

            nodeHandlers.Add("li", (node, parent) =>
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

            nodeHandlers.Add("#text", (node, parent) =>
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

        private static DocumentObject AddFormattedText(HtmlNode node, DocumentObject parent, TextFormat format)
        {
            var formattedText = parent as FormattedText;
            if (formattedText != null)
            {
                return formattedText.Format(format);
            }

            // otherwise parent is paragraph or section
            return GetParagraph(parent).AddFormattedText(format);
        }

        private static DocumentObject AddHeading(HtmlNode node, DocumentObject parent)
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
