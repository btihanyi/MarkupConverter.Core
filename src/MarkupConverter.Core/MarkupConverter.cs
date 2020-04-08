using System;

namespace MarkupConverter.Core
{
    public class MarkupConverter : IMarkupConverter
    {
        public string ConvertXamlToHtml(string xamlText)
        {
            return HtmlFromXamlConverter.ConvertXamlToHtml(xamlText, asFullDocument: false);
        }

        public string ConvertHtmlToXaml(string htmlText)
        {
            return HtmlToXamlConverter.ConvertHtmlToXaml(htmlText, asFlowDocument: true);
        }

        public string ConvertRtfToHtml(string rtfText)
        {
            return RtfToHtmlConverter.ConvertRtfToHtml(rtfText);
        }

        public string ConvertHtmlToRtf(string htmlText)
        {
            return HtmlToRtfConverter.ConvertHtmlToRtf(htmlText);
        }
    }
}
