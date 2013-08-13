using MigraDoc.DocumentObjectModel;
using System;

namespace MigraDoc.Extensions
{
    public interface IConverter
    {
        Action<Section> Convert(ExCSS.Stylesheet sheet, string contents);
    }
}
