﻿using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Extensions.Markdown;
using MigraDoc.Rendering;
using System.Diagnostics;
using System.IO;

namespace MigraDoc.Extensions.Html.Example
{
    class Program
    {
        static void Main(string[] args)
        {
            new Program().Run();
        }

        private string outputName = "output.pdf";

        void Run()
        {
            if (File.Exists(outputName))
            {
                File.Delete(outputName);
            }

            var doc = new Document();
            doc.DefaultPageSetup.Orientation = Orientation.Portrait;
            doc.DefaultPageSetup.PageFormat = PageFormat.A4;
            //doc.DefaultPageSetup.LeftMargin = Unit.FromMillimeter(5.0);
            //doc.DefaultPageSetup.RightMargin = Unit.FromMillimeter(5.0);
            //doc.DefaultPageSetup.BottomMargin = Unit.FromMillimeter(5.0);
            //doc.DefaultPageSetup.TopMargin = Unit.FromMillimeter(5.0);
            StyleDoc(doc);
            var section = doc.AddSection();
            var footer = new TextFrame();

            section.Footers.Primary.Add(footer);
            var html = File.ReadAllText("example.html");
            section.AddHtml(html);

            var renderer = new PdfDocumentRenderer();
            renderer.Document = doc;
            renderer.RenderDocument();

            renderer.Save(outputName);
            Process.Start(outputName);
        }

        private void AddTable(Document doc)
        {
            doc.LastSection.AddParagraph("Simple Tables", "Heading2");

            Table table = new Table();
            table.Borders.Width = 0.75;

            Column column = table.AddColumn(Unit.FromCentimeter(2));
            column.Format.Alignment = ParagraphAlignment.Center;

            table.AddColumn(Unit.FromCentimeter(5));

            Row row = table.AddRow();
            row.Shading.Color = Colors.PaleGoldenrod;
            Cell cell = row.Cells[0];
            cell.AddParagraph("Itemus");
            cell = row.Cells[1];
            cell.AddParagraph("Descriptum");

            row = table.AddRow();
            cell = row.Cells[0];
            cell.AddParagraph("1");
            cell = row.Cells[1];
            cell.AddParagraph("Test");

            row = table.AddRow();
            cell = row.Cells[0];
            cell.AddParagraph("2");
            cell = row.Cells[1];
            cell.AddParagraph("test2");

            table.SetEdge(0, 0, 2, 3, Edge.Box, BorderStyle.Single, 1.5, Colors.Black);

            doc.LastSection.Add(table);
        }

        private void StyleDoc(Document doc)
        {
            Color green = new Color(108, 179, 63),
                  brown = new Color(88, 71, 76),
                  lightbrown = new Color(150, 132, 126);

            var body = doc.Styles["Normal"];

            body.Font.Size = Unit.FromInch(0.14);
            body.Font.Color = new Color(51, 51, 51);

            body.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
            body.ParagraphFormat.LineSpacing = 1.25;
            body.ParagraphFormat.SpaceAfter = 10;

            var footer = doc.Styles["Footer"];
            footer.Font.Size = Unit.FromInch(0.125);
            footer.Font.Color = green;

            var h1 = doc.Styles["Heading1"];
            h1.Font.Color = green;
            h1.Font.Bold = true;
            h1.Font.Size = Unit.FromPoint(15);

            var h2 = doc.Styles["Heading2"];
            h2.Font.Color = green;
            h2.Font.Bold = true;
            h2.Font.Size = Unit.FromPoint(13);

            var h3 = doc.Styles["Heading3"];
            h3.Font.Bold = true;
            h3.Font.Color = Colors.Black;
            h3.Font.Size = Unit.FromPoint(11);

            var links = doc.Styles["Hyperlink"];
            links.Font.Color = green;

            var unorderedlist = doc.AddStyle("UnorderedList", "Normal");
            var listInfo = new ListInfo();
            listInfo.ListType = ListType.BulletList1;
            unorderedlist.ParagraphFormat.ListInfo = listInfo;
            unorderedlist.ParagraphFormat.LeftIndent = "1cm";
            unorderedlist.ParagraphFormat.FirstLineIndent = "-0.5cm";
            unorderedlist.ParagraphFormat.SpaceAfter = 0;

            var orderedlist = doc.AddStyle("OrderedList", "UnorderedList");
            orderedlist.ParagraphFormat.ListInfo.ListType = ListType.NumberList1;

            // for list spacing (since MigraDoc doesn't provide a list object that we can target)
            var listStart = doc.AddStyle("ListStart", "Normal");
            listStart.ParagraphFormat.SpaceAfter = 0;
            listStart.ParagraphFormat.LineSpacing = 0.5;
            var listEnd = doc.AddStyle("ListEnd", "ListStart");
            listEnd.ParagraphFormat.LineSpacing = 1;

            var hr = doc.AddStyle("HorizontalRule", "Normal");
            var hrBorder = new Border();
            hrBorder.Width = "1pt";
            hrBorder.Color = Colors.DarkGray;
            hr.ParagraphFormat.Borders.Bottom = hrBorder;
            hr.ParagraphFormat.LineSpacing = 0;
            hr.ParagraphFormat.SpaceBefore = 15;
        }
    }
}
