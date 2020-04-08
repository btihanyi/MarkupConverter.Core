//---------------------------------------------------------------------------
//
// File: HtmlFromXamlConverter.cs
//
// Copyright (C) Microsoft Corporation.  All rights reserved.
//
// Description: Prototype for Xaml - HTML conversion
//
//---------------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;

namespace MarkupConverter.Core
{
    /// <summary>
    /// HtmlToXamlConverter is a static class that takes an HTML string
    /// and converts it into XAML.
    /// </summary>
    public static class HtmlFromXamlConverter
    {
        /// <summary>
        /// Main entry point for XAML-to-HTML converter.
        /// Converts a XAML string into HTML string.
        /// </summary>
        /// <param name="xamlString">
        /// XAML string to convert.
        /// </param>
        /// <param name="asFullDocument"></param>
        /// <returns>
        /// HTML string produced from a source XAML.
        /// </returns>
        public static string ConvertXamlToHtml(string xamlString, bool asFullDocument)
        {
            using var xamlReader = new XmlTextReader(new StringReader(xamlString));
            var htmlStringBuilder = new StringBuilder(100);
            using var htmlWriter = new XmlTextWriter(new StringWriter(htmlStringBuilder));

            if (WriteFlowDocument(xamlReader, htmlWriter, asFullDocument))
            {
                return htmlStringBuilder.ToString();
            }
            else
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Processes a root level element of XAML (normally it's FlowDocument element).
        /// </summary>
        /// <param name="xamlReader">
        /// XmlTextReader for a source XAML.
        /// </param>
        /// <param name="htmlWriter">
        /// XmlTextWriter producing resulting HTML.
        /// </param>
        private static bool WriteFlowDocument(XmlTextReader xamlReader, XmlTextWriter htmlWriter, bool asFullDocument)
        {
            if (!ReadNextToken(xamlReader))
            {
                // XAML content is empty - nothing to convert
                return false;
            }

            if (xamlReader.NodeType != XmlNodeType.Element || xamlReader.Name != "FlowDocument")
            {
                // Root FlowDocument element is missing
                return false;
            }

            // Create a buffer StringBuilder for collecting CSS properties for inline STYLE attributes
            // on every element level (it will be re-initialized on every level).
            var inlineStyle = new StringBuilder();

            if (asFullDocument)
            {
                htmlWriter.WriteStartElement("HTML");
                htmlWriter.WriteStartElement("BODY");
            }

            WriteFormattingProperties(xamlReader, htmlWriter, inlineStyle);

            WriteElementContent(xamlReader, htmlWriter, inlineStyle);

            if (asFullDocument)
            {
                htmlWriter.WriteEndElement();
                htmlWriter.WriteEndElement();
            }

            return true;
        }

        /// <summary>
        /// Reads attributes of the current XAML element and converts
        /// them into appropriate HTML attributes or CSS styles.
        /// </summary>
        /// <param name="xamlReader">
        /// XmlTextReader which is expected to be at XmlNodeType.Element
        /// (opening element tag) position.
        /// The reader will remain at the same level after function complete.
        /// </param>
        /// <param name="htmlWriter">
        /// XmlTextWriter for output HTML, which is expected to be in
        /// after WriteStartElement state.
        /// </param>
        /// <param name="inlineStyle">
        /// String builder for collecting CSS properties for inline STYLE attribute.
        /// </param>
        private static void WriteFormattingProperties(XmlTextReader xamlReader, XmlTextWriter htmlWriter, StringBuilder inlineStyle)
        {
            Debug.Assert(xamlReader.NodeType == XmlNodeType.Element);

            // Clear string builder for the inline style
            inlineStyle.Remove(0, inlineStyle.Length);

            if (!xamlReader.HasAttributes)
            {
                return;
            }

            bool borderSet = false;

            while (xamlReader.MoveToNextAttribute())
            {
                string css = null;

                switch (xamlReader.Name)
                {
                    // Character formatting properties
                    case "Background":
                        css = "background-color:" + ParseXamlColor(xamlReader.Value) + ";";
                        break;
                    case "FontFamily":
                        css = "font-family:" + xamlReader.Value + ";";
                        break;
                    case "FontStyle":
                        css = "font-style:" + xamlReader.Value.ToLowerInvariant() + ";";
                        break;
                    case "FontWeight":
                        css = "font-weight:" + xamlReader.Value.ToLowerInvariant() + ";";
                        break;
                    case "FontStretch":
                        break;
                    case "FontSize":
                        css = "font-size:" + xamlReader.Value + ";";
                        break;
                    case "Foreground":
                        css = "color:" + ParseXamlColor(xamlReader.Value) + ";";
                        break;
                    case "TextDecorations":
                        if (string.Equals(xamlReader.Value, "strikethrough", StringComparison.OrdinalIgnoreCase))
                        {
                            css = "text-decoration:line-through;";
                        }
                        else
                        {
                            css = "text-decoration:underline;";
                        }
                        break;
                    case "TextEffects":
                        break;
                    case "Emphasis":
                        break;
                    case "StandardLigatures":
                        break;
                    case "Variants":
                        break;
                    case "Capitals":
                        break;
                    case "Fraction":
                        break;

                    // Paragraph formatting properties
                    case "Padding":
                        css = "padding:" + ParseXamlThickness(xamlReader.Value) + ";";
                        break;
                    case "Margin":
                        css = "margin:" + ParseXamlThickness(xamlReader.Value) + ";";
                        break;
                    case "BorderThickness":
                        css = "border-width:" + ParseXamlThickness(xamlReader.Value) + ";";
                        borderSet = true;
                        break;
                    case "BorderBrush":
                        css = "border-color:" + ParseXamlColor(xamlReader.Value) + ";";
                        borderSet = true;
                        break;
                    case "LineHeight":
                        break;
                    case "TextIndent":
                        css = "text-indent:" + xamlReader.Value + ";";
                        break;
                    case "TextAlignment":
                        css = "text-align:" + xamlReader.Value + ";";
                        break;
                    case "IsKeptTogether":
                        break;
                    case "IsKeptWithNext":
                        break;
                    case "ColumnBreakBefore":
                        break;
                    case "PageBreakBefore":
                        break;
                    case "FlowDirection":
                        break;

                    // Table attributes
                    // ----------------
                    case "Width":
                        css = "width:" + xamlReader.Value + ";";
                        break;
                    case "ColumnSpan":
                        htmlWriter.WriteAttributeString("COLSPAN", xamlReader.Value);
                        break;
                    case "RowSpan":
                        htmlWriter.WriteAttributeString("ROWSPAN", xamlReader.Value);
                        break;

                    // Hyperlink Attributes
                    case "NavigateUri":
                        htmlWriter.WriteAttributeString("HREF", xamlReader.Value);
                        break;

                    case "TargetName":
                        htmlWriter.WriteAttributeString("TARGET", xamlReader.Value);
                        break;
                }

                if (css != null)
                {
                    inlineStyle.Append(css);
                }
            }

            if (borderSet)
            {
                inlineStyle.Append("border-style:solid;mso-element:para-border-div;");
            }

            // Return the xamlReader back to element level
            xamlReader.MoveToElement();
            Debug.Assert(xamlReader.NodeType == XmlNodeType.Element);
        }

        private static string ParseXamlColor(string color)
        {
            if (color.StartsWith("#"))
            {
                // Remove transparency value
                color = "#" + color.Substring(3);
            }

            return color;
        }

        private static string ParseXamlThickness(string thickness)
        {
            string[] values = thickness.Split(',');

            for (int i = 0; i < values.Length; i++)
            {
                if (double.TryParse(values[i], NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
                {
                    values[i] = Math.Ceiling(value).ToString();
                }
                else
                {
                    values[i] = "1";
                }
            }

            return values.Length switch
            {
                1 => thickness,
                2 => values[1] + " " + values[0],
                4 => values[1] + " " + values[2] + " " + values[3] + " " + values[0],
                _ => values[0],
            };
        }

        /// <summary>
        /// Reads a content of current XAML element, converts it.
        /// </summary>
        /// <param name="xamlReader">
        /// XmlTextReader which is expected to be at XmlNodeType.Element
        /// (opening element tag) position.
        /// </param>
        /// <param name="htmlWriter">
        /// May be null, in which case we are skipping the XAML element;
        /// without producing any output to HTML.
        /// </param>
        /// <param name="inlineStyle">
        /// StringBuilder used for collecting CSS properties for inline STYLE attribute.
        /// </param>
        private static void WriteElementContent(XmlTextReader xamlReader, XmlTextWriter htmlWriter, StringBuilder inlineStyle)
        {
            Debug.Assert(xamlReader.NodeType == XmlNodeType.Element);

            bool elementContentStarted = false;

            if (xamlReader.IsEmptyElement)
            {
                if (htmlWriter != null && !elementContentStarted && inlineStyle.Length > 0)
                {
                    // Output STYLE attribute and clear inlineStyle buffer.
                    htmlWriter.WriteAttributeString("STYLE", inlineStyle.ToString());
                    inlineStyle.Remove(0, inlineStyle.Length);
                }
            }
            else
            {
                while (ReadNextToken(xamlReader) && xamlReader.NodeType != XmlNodeType.EndElement)
                {
                    switch (xamlReader.NodeType)
                    {
                        case XmlNodeType.Element:
                            if (xamlReader.Name.Contains("."))
                            {
                                AddComplexProperty(xamlReader, inlineStyle);
                            }
                            else
                            {
                                if (htmlWriter != null && !elementContentStarted && inlineStyle.Length > 0)
                                {
                                    // Output STYLE attribute and clear inlineStyle buffer.
                                    htmlWriter.WriteAttributeString("STYLE", inlineStyle.ToString());
                                    inlineStyle.Remove(0, inlineStyle.Length);
                                }

                                elementContentStarted = true;
                                WriteElement(xamlReader, htmlWriter, inlineStyle);
                            }

                            Debug.Assert(xamlReader.NodeType == XmlNodeType.EndElement ||
                                         (xamlReader.NodeType == XmlNodeType.Element && xamlReader.IsEmptyElement));
                            break;

                        case XmlNodeType.Comment:
                            if (htmlWriter != null)
                            {
                                if (!elementContentStarted && inlineStyle.Length > 0)
                                {
                                    htmlWriter.WriteAttributeString("STYLE", inlineStyle.ToString());
                                }

                                htmlWriter.WriteComment(xamlReader.Value);
                            }

                            elementContentStarted = true;
                            break;

                        case XmlNodeType.CDATA:
                        case XmlNodeType.Text:
                        case XmlNodeType.SignificantWhitespace:
                            if (htmlWriter != null)
                            {
                                if (!elementContentStarted && inlineStyle.Length > 0)
                                {
                                    htmlWriter.WriteAttributeString("STYLE", inlineStyle.ToString());
                                }

                                htmlWriter.WriteString(xamlReader.Value);
                            }

                            elementContentStarted = true;
                            break;
                    }
                }

                Debug.Assert(xamlReader.NodeType == XmlNodeType.EndElement);
            }
        }

        /// <summary>
        /// Converts an element notation of complex property into.
        /// </summary>
        /// <param name="xamlReader">
        /// On entry this XmlTextReader must be on Element start tag;
        /// on exit - on EndElement tag.
        /// </param>
        /// <param name="inlineStyle">
        /// StringBuilder containing a value for STYLE attribute.
        /// </param>
        private static void AddComplexProperty(XmlTextReader xamlReader, StringBuilder inlineStyle)
        {
            Debug.Assert(xamlReader.NodeType == XmlNodeType.Element);

            if (inlineStyle != null && xamlReader.Name.EndsWith(".TextDecorations"))
            {
                inlineStyle.Append("text-decoration:underline;");
            }

            // Skip the element representing the complex property
            WriteElementContent(xamlReader, htmlWriter: null, inlineStyle: null);
        }

        /// <summary>
        /// Converts a XAML element into an appropriate HTML element.
        /// </summary>
        /// <param name="xamlReader">
        /// On entry this XmlTextReader must be on Element start tag;
        /// on exit - on EndElement tag.
        /// </param>
        /// <param name="htmlWriter">
        /// May be null, in which case we are skipping XAML content
        /// without producing any HTML output.
        /// </param>
        /// <param name="inlineStyle">
        /// StringBuilder used for collecting CSS properties for inline STYLE attributes on every level.
        /// </param>
        private static void WriteElement(XmlTextReader xamlReader, XmlTextWriter htmlWriter, StringBuilder inlineStyle)
        {
            Debug.Assert(xamlReader.NodeType == XmlNodeType.Element);

            if (htmlWriter == null)
            {
                // Skipping mode; recurse into the XAML element without any output
                WriteElementContent(xamlReader, htmlWriter: null, null);
            }
            else
            {
                string htmlElementName;
                switch (xamlReader.Name)
                {
                    case "Run":
                    case "Span":
                    case "InlineUIContainer":
                        htmlElementName = "SPAN";
                        break;
                    case "Bold":
                        htmlElementName = "B";
                        break;
                    case "Italic":
                        htmlElementName = "I";
                        break;
                    case "Paragraph":
                        htmlElementName = "P";
                        break;
                    case "BlockUIContainer":
                    case "Section":
                        htmlElementName = "DIV";
                        break;
                    case "Table":
                        htmlElementName = "TABLE";
                        break;
                    case "TableColumn":
                        htmlElementName = "COL";
                        break;
                    case "TableRowGroup":
                        htmlElementName = "TBODY";
                        break;
                    case "TableRow":
                        htmlElementName = "TR";
                        break;
                    case "TableCell":
                        htmlElementName = "TD";
                        break;
                    case "List":
                        string marker = xamlReader.GetAttribute("MarkerStyle");
                        if (marker == null || marker == "None" || marker == "Disc" || marker == "Circle" || marker == "Square" || marker == "Box")
                        {
                            htmlElementName = "UL";
                        }
                        else
                        {
                            htmlElementName = "OL";
                        }
                        break;
                    case "ListItem":
                        htmlElementName = "LI";
                        break;
                    case "Hyperlink":
                        htmlElementName = "A";
                        break;
                    default:
                        htmlElementName = null; // Ignore the element
                        break;
                }

                if (htmlWriter != null && htmlElementName != null)
                {
                    htmlWriter.WriteStartElement(htmlElementName);

                    WriteFormattingProperties(xamlReader, htmlWriter, inlineStyle);

                    WriteElementContent(xamlReader, htmlWriter, inlineStyle);

                    htmlWriter.WriteEndElement();
                }
                else
                {
                    // Skip this unrecognized XAML element
                    WriteElementContent(xamlReader, htmlWriter: null, null);
                }
            }
        }

        // Reader advance helpers
        // ----------------------

        /// <summary>
        /// Reads several items from xamlReader skipping all non-significant stuff.
        /// </summary>
        /// <param name="xamlReader">
        /// XmlTextReader from tokens are being read.
        /// </param>
        /// <returns>
        /// True if new token is available; false if end of stream reached.
        /// </returns>
        private static bool ReadNextToken(XmlReader xamlReader)
        {
            while (xamlReader.Read())
            {
                Debug.Assert(xamlReader.ReadState == ReadState.Interactive, "Reader is expected to be in Interactive state (" + xamlReader.ReadState + ")");
                switch (xamlReader.NodeType)
                {
                    case XmlNodeType.Element:
                    case XmlNodeType.EndElement:
                    case XmlNodeType.None:
                    case XmlNodeType.CDATA:
                    case XmlNodeType.Text:
                    case XmlNodeType.SignificantWhitespace:
                        return true;

                    case XmlNodeType.Whitespace:
                        if (xamlReader.XmlSpace == XmlSpace.Preserve)
                        {
                            return true;
                        }

                        // ignore insignificant whitespace
                        break;

                    case XmlNodeType.EndEntity:
                    case XmlNodeType.EntityReference:
                        // Implement entity reading
                        // xamlReader.ResolveEntity();
                        // xamlReader.Read();
                        // ReadChildNodes( parent, parentBaseUri, xamlReader, positionInfo);
                        break; // for now we ignore entities as insignificant stuff

                    case XmlNodeType.Comment:
                        return true;
                    case XmlNodeType.ProcessingInstruction:
                    case XmlNodeType.DocumentType:
                    case XmlNodeType.XmlDeclaration:
                    default:
                        // Ignorable stuff
                        break;
                }
            }

            return false;
        }
    }
}
