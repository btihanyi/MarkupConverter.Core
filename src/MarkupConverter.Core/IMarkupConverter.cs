using System;

namespace MarkupConverter.Core
{
    public interface IMarkupConverter
    {
        string ConvertXamlToHtml(string xamlText);

        string ConvertHtmlToXaml(string htmlText);

        string ConvertRtfToHtml(string rtfText);

        string ConvertHtmlToRtf(string htmlText);
    }
}
