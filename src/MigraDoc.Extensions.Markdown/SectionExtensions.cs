using MigraDoc.DocumentObjectModel;
using MigraDoc.Extensions.Html;
using System;

namespace MigraDoc.Extensions.Markdown
{
    public static class SectionExtensions
    {
        public static Section AddMarkdown(this Section section, string markdown)
        {
            return section; // ection.Add(markdown, new Object());
        }
    }
}
