using System;
using System.IO;
using System.Windows;
using System.Windows.Documents;

namespace MarkupConverter.Core
{
    public static class RtfToHtmlConverter
    {
        public static string ConvertRtfToHtml(string rtfText)
        {
            string xamlText = $"<FlowDocument>{ConvertRtfToXaml(rtfText)}</FlowDocument>";

            return HtmlFromXamlConverter.ConvertXamlToHtml(xamlText, asFullDocument: false);
        }

        private static string ConvertRtfToXaml(string rtfText)
        {
            if (string.IsNullOrEmpty(rtfText))
            {
                return string.Empty;
            }

            var flowDocument = new FlowDocument();
            var textRange = new TextRange(flowDocument.ContentStart, flowDocument.ContentEnd);

            // Create a MemoryStream of the RTF content
            using (var rtfMemoryStream = new MemoryStream())
            using (var rtfStreamWriter = new StreamWriter(rtfMemoryStream))
            {
                rtfStreamWriter.Write(rtfText);
                rtfStreamWriter.Flush();
                rtfMemoryStream.Position = 0;

                // Load the MemoryStream into TextRange ranging from start to end of RichTextBox.
                textRange.Load(rtfMemoryStream, DataFormats.Rtf);
            }

            using (var rtfMemoryStream = new MemoryStream())
            using (var rtfStreamReader = new StreamReader(rtfMemoryStream))
            {
                textRange = new TextRange(flowDocument.ContentStart, flowDocument.ContentEnd);
                textRange.Save(rtfMemoryStream, DataFormats.Xaml);
                rtfMemoryStream.Position = 0;

                return rtfStreamReader.ReadToEnd();
            }
        }
    }
}
