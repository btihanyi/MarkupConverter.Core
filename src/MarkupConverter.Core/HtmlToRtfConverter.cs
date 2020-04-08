using System;
using System.IO;
using System.Windows;
using System.Windows.Documents;

namespace MarkupConverter.Core
{
    public static class HtmlToRtfConverter
    {
        public static string ConvertHtmlToRtf(string htmlText)
        {
            string xamlText = HtmlToXamlConverter.ConvertHtmlToXaml(htmlText, false);

            return ConvertXamlToRtf(xamlText);
        }

        private static string ConvertXamlToRtf(string xamlText)
        {
            if (string.IsNullOrEmpty(xamlText))
            {
                return string.Empty;
            }

            var flowDocument = new FlowDocument();
            var textRange = new TextRange(flowDocument.ContentStart, flowDocument.ContentEnd);

            // Create a MemoryStream of the XAML content
            using (var xamlMemoryStream = new MemoryStream())
            using (var xamlStreamWriter = new StreamWriter(xamlMemoryStream))
            {
                xamlStreamWriter.Write(xamlText);
                xamlStreamWriter.Flush();
                xamlMemoryStream.Position = 0;

                // Load the MemoryStream into TextRange ranging from start to end of RichTextBox.
                textRange.Load(xamlMemoryStream, DataFormats.Xaml);
            }

            using (var rtfMemoryStream = new MemoryStream())
            using (var rtfStreamReader = new StreamReader(rtfMemoryStream))
            {
                textRange = new TextRange(flowDocument.ContentStart, flowDocument.ContentEnd);
                textRange.Save(rtfMemoryStream, DataFormats.Rtf);
                rtfMemoryStream.Position = 0;

                return rtfStreamReader.ReadToEnd();
            }
        }
    }
}
