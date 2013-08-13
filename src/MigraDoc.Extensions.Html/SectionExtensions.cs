using MigraDoc.DocumentObjectModel;
using MigraDoc.Extensions.Html;
using System;

namespace MigraDoc.Extensions.Html
{
    public static class SectionExtensions
    {
        public static Section AddHtml(this Section section, ExCSS.Stylesheet sheet, string html)
        {
            return section.Add(sheet, html, new HtmlConverter());
        }
    }
}
