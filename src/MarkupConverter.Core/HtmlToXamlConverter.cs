//---------------------------------------------------------------------------
//
// File: HtmlXamlConverter.cs
//
// Copyright (C) Microsoft Corporation.  All rights reserved.
//
// Description: Prototype for HTML - XAML conversion
//
//---------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Windows; // DependencyProperty
using System.Windows.Documents; // TextElement
using System.Xml;

namespace MarkupConverter.Core
{
    /// <summary>
    /// HtmlToXamlConverter is a static class that takes an HTML string and converts it into XAML.
    /// </summary>
    public static class HtmlToXamlConverter
    {
        // The constants represent all XAML names used in a conversion
        public const string XamlFlowDocument = "FlowDocument";

        public const string XamlRun = "Run";
        public const string XamlSpan = "Span";
        public const string XamlHyperlink = "Hyperlink";
        public const string XamlHyperlinkNavigateUri = "NavigateUri";
        public const string XamlHyperlinkTargetName = "TargetName";

        public const string XamlSection = "Section";

        public const string XamlList = "List";

        public const string XamlListMarkerStyle = "MarkerStyle";
        public const string XamlListMarkerStyleNone = "None";
        public const string XamlListMarkerStyleDecimal = "Decimal";
        public const string XamlListMarkerStyleDisc = "Disc";
        public const string XamlListMarkerStyleCircle = "Circle";
        public const string XamlListMarkerStyleSquare = "Square";
        public const string XamlListMarkerStyleBox = "Box";
        public const string XamlListMarkerStyleLowerLatin = "LowerLatin";
        public const string XamlListMarkerStyleUpperLatin = "UpperLatin";
        public const string XamlListMarkerStyleLowerRoman = "LowerRoman";
        public const string XamlListMarkerStyleUpperRoman = "UpperRoman";

        public const string XamlListItem = "ListItem";

        public const string XamlLineBreak = "LineBreak";

        public const string XamlParagraph = "Paragraph";

        public const string XamlMargin = "Margin";
        public const string XamlPadding = "Padding";
        public const string XamlBorderBrush = "BorderBrush";
        public const string XamlBorderThickness = "BorderThickness";

        public const string XamlTable = "Table";

        public const string XamlTableColumn = "TableColumn";
        public const string XamlTableRowGroup = "TableRowGroup";
        public const string XamlTableRow = "TableRow";

        public const string XamlTableCell = "TableCell";
        public const string XamlTableCellBorderThickness = "BorderThickness";
        public const string XamlTableCellBorderBrush = "BorderBrush";

        public const string XamlTableCellColumnSpan = "ColumnSpan";
        public const string XamlTableCellRowSpan = "RowSpan";

        public const string XamlWidth = "Width";
        public const string XamlBrushesBlack = "Black";
        public const string XamlFontFamily = "FontFamily";

        public const string XamlFontSize = "FontSize";
        public const string XamlFontSizeXXLarge = "22pt"; // "XXLarge";
        public const string XamlFontSizeXLarge = "20pt"; // "XLarge";
        public const string XamlFontSizeLarge = "18pt"; // "Large";
        public const string XamlFontSizeMedium = "16pt"; // "Medium";
        public const string XamlFontSizeSmall = "12pt"; // "Small";
        public const string XamlFontSizeXSmall = "10pt"; // "XSmall";
        public const string XamlFontSizeXXSmall = "8pt"; // "XXSmall";

        public const string XamlFontWeight = "FontWeight";
        public const string XamlFontWeightBold = "Bold";

        public const string XamlFontStyle = "FontStyle";

        public const string XamlForeground = "Foreground";
        public const string XamlBackground = "Background";
        public const string XamlTextDecorations = "TextDecorations";
        public const string XamlTextDecorationsUnderline = "Underline";

        public const string XamlTextIndent = "TextIndent";
        public const string XamlTextAlignment = "TextAlignment";

        private const string XamlNamespace = "http://schemas.microsoft.com/winfx/2006/xaml/presentation";

        // Stores a parent XAML element for the case when selected fragment is inline.
        private static XmlElement inlineFragmentParentElement;

        /// <summary>
        /// Converts an HTML string into XAML string.
        /// </summary>
        /// <param name="htmlString">
        /// Input HTML which may be badly formated XML.
        /// </param>
        /// <param name="asFlowDocument">
        /// <see langword="true"/> indicates that we need a FlowDocument as a root element;
        /// <see langword="false"/> means that Section or Span elements will be used
        /// depending on StartFragment/EndFragment comments locations.
        /// </param>
        /// <returns>
        /// Well-formed XML representing XAML equivalent for the input HTML string.
        /// </returns>
        public static string ConvertHtmlToXaml(string htmlString, bool asFlowDocument)
        {
            // Create well-formed XML from HTML string
            var htmlElement = HtmlParser.ParseHtml(htmlString);

            // Decide what name to use as a root
            string rootElementName = asFlowDocument ? XamlFlowDocument : XamlSection;

            // Create an XmlDocument for generated XAML
            var xamlTree = new XmlDocument();
            var xamlFlowDocumentElement = xamlTree.CreateElement(null, rootElementName, XamlNamespace);

            // Extract style definitions from all STYLE elements in the document
            var stylesheet = new CssStylesheet(htmlElement);

            // Source context is a stack of all elements - ancestors of a parentElement
            var sourceContext = new List<XmlElement>(10);

            // Clear fragment parent
            inlineFragmentParentElement = null;

            // convert root HTML element
            AddBlock(xamlFlowDocumentElement, htmlElement, new Dictionary<string, string>(), stylesheet, sourceContext);

            // In case if the selected fragment is inline, extract it into a separate Span wrapper
            if (!asFlowDocument)
            {
                xamlFlowDocumentElement = ExtractInlineFragment(xamlFlowDocumentElement);
            }

            // Return a string representing resulting XAML
            xamlFlowDocumentElement.SetAttribute("xml:space", "preserve");
            return xamlFlowDocumentElement.OuterXml;
        }

        /// <summary>
        /// Returns a value for an attribute by its name (ignoring casing).
        /// </summary>
        /// <param name="element">
        /// XmlElement in which we are trying to find the specified attribute.
        /// </param>
        /// <param name="attributeName">
        /// String representing the attribute name to be searched for.
        /// </param>
        public static string GetAttribute(XmlElement element, string attributeName)
        {
            attributeName = attributeName.ToLowerInvariant();

            foreach (XmlAttribute attribute in element.Attributes)
            {
                if (attribute.Name.ToLowerInvariant() == attributeName)
                {
                    return attribute.Value;
                }
            }

            return null;
        }

        /// <summary>
        /// Returns string extracted from quotation marks.
        /// </summary>
        /// <param name="value">
        /// String representing value enclosed in quotation marks.
        /// </param>
        internal static string UnQuote(string value)
        {
            if ((value.StartsWith("\"") && value.EndsWith("\"")) || (value.StartsWith("'") && value.EndsWith("'")))
            {
                value = value[1..^1].Trim();
            }

            return value;
        }

        /// <summary>
        /// <para>
        /// Analyzes the given htmlElement expecting it to be converted
        /// into some of XAML Block elements and adds the converted block
        /// to the children collection of xamlParentElement.
        /// </para>
        /// <para>
        /// Analyzes the given XmlElement htmlElement, recognizes it as some HTML element
        /// and adds it as a child to a xamlParentElement.
        /// In some cases several following siblings of the given htmlElement
        /// will be consumed too (e.g. LIs encountered without wrapping UL/OL,
        /// which must be collected together and wrapped into one implicit List element).
        /// </para>
        /// </summary>
        /// <param name="xamlParentElement">
        /// Parent XAML element, to which new converted element will be added.
        /// </param>
        /// <param name="htmlNode">
        /// Source HTML element subject to convert to XAML.
        /// </param>
        /// <param name="inheritedProperties">
        /// Properties inherited from an outer context.
        /// </param>
        /// <param name="stylesheet"></param>
        /// <param name="sourceContext"></param>
        /// <returns>
        /// Last processed HTML node. Normally it should be the same htmlElement
        /// as was passed as a parameter, but in some irregular cases
        /// it could one of its following siblings.
        /// The caller must use this node to get to next sibling from it.
        /// </returns>
        private static XmlNode AddBlock(XmlElement xamlParentElement, XmlNode htmlNode, Dictionary<string, string> inheritedProperties, CssStylesheet stylesheet, List<XmlElement> sourceContext)
        {
            if (htmlNode is XmlComment xmlComments)
            {
                DefineInlineFragmentParent(xmlComments, xamlParentElement: null);
            }
            else if (htmlNode is XmlText)
            {
                htmlNode = AddImplicitParagraph(xamlParentElement, htmlNode, inheritedProperties, stylesheet, sourceContext);
            }
            else if (htmlNode is XmlElement htmlElement) // Identify element name
            {
                string htmlElementName = htmlElement.LocalName; // Keep the name case-sensitive to check XML names
                string htmlElementNamespace = htmlElement.NamespaceURI;

                if (htmlElementNamespace != HtmlParser.XhtmlNamespace)
                {
                    // Non-HTML element. skip it
                    // Isn't it too aggressive? What if this is just an error in HTML tag name?
                    // TODO: Consider skipping just a wrapper in recursing into the element tree,
                    // which may produce some garbage though coming from XML fragments.
                    return htmlElement;
                }

                // Put source element to the stack
                sourceContext.Add(htmlElement);

                // Convert the name to lowercase, because HTML elements are case-insensitive
                htmlElementName = htmlElementName.ToLowerInvariant();

                // Switch to an appropriate kind of processing depending on HTML element name
                switch (htmlElementName)
                {
                    // Sections:
                    case "html":
                    case "body":
                    case "div":
                    case "form": // Not a block according to XHTML spec
                    case "pre": // Renders text in a fixed-width font
                    case "blockquote":
                    case "caption":
                    case "center":
                    case "cite":
                        AddSection(xamlParentElement, htmlElement, inheritedProperties, stylesheet, sourceContext);
                        break;

                    // Paragraphs:
                    case "p":
                    case "h1":
                    case "h2":
                    case "h3":
                    case "h4":
                    case "h5":
                    case "h6":
                    case "nsrtitle":
                    case "textarea":
                    case "dd": // ???
                    case "dl": // ???
                    case "dt": // ???
                    case "tt": // ???
                        AddParagraph(xamlParentElement, htmlElement, inheritedProperties, stylesheet, sourceContext);
                        break;

                    case "ol":
                    case "ul":
                    case "dir": // Treat as UL element
                    case "menu": // Treat as UL element
                        // List element conversion
                        AddList(xamlParentElement, htmlElement, inheritedProperties, stylesheet, sourceContext);
                        break;
                    case "li":
                        // LI outside of OL/UL
                        // Collect all sibling LIs, wrap them into a List and then proceed with the element following the last of LIs
                        htmlNode = AddOrphanListItems(xamlParentElement, htmlElement, inheritedProperties, stylesheet, sourceContext);
                        break;

                    case "img":
                        // TODO: Add image processing
                        AddImage(xamlParentElement, htmlElement, inheritedProperties, stylesheet, sourceContext);
                        break;

                    case "table":
                        // hand off to table parsing function which will perform special table syntax checks
                        AddTable(xamlParentElement, htmlElement, inheritedProperties, stylesheet, sourceContext);
                        break;

                    case "tbody":
                    case "tfoot":
                    case "thead":
                    case "tr":
                    case "td":
                    case "th":
                        // Table stuff without table wrapper
                        ////TODO: Add special-case processing here for elements that should be within tables when the
                        ////      parent element is NOT a table. If the parent element is a table they can be processed normally.
                        ////      we need to compare against the parent element here, we can't just break on a switch
                        goto default; // Thus we will skip this element as unknown, but still recurse into it.

                    case "style": // We already pre-processed all style elements. Ignore it now
                    case "meta":
                    case "head":
                    case "title":
                    case "script":
                        // Ignore these elements
                        break;

                    default:
                        // Wrap a sequence of inlines into an implicit paragraph
                        htmlNode = AddImplicitParagraph(xamlParentElement, htmlElement, inheritedProperties, stylesheet, sourceContext);
                        break;
                }

                // Remove the element from the stack
                Debug.Assert(sourceContext.Count > 0 && sourceContext[sourceContext.Count - 1] == htmlElement);
                sourceContext.RemoveAt(sourceContext.Count - 1);
            }

            // Return last processed node
            return htmlNode;
        }

        private static void AddBreak(XmlElement xamlParentElement, string htmlElementName)
        {
            // Create new XAML element corresponding to this HTML element
            var xamlLineBreak = xamlParentElement.OwnerDocument.CreateElement(prefix: null, localName: XamlLineBreak, XamlNamespace);
            xamlParentElement.AppendChild(xamlLineBreak);
            if (htmlElementName == "hr")
            {
                var xamlHorizontalLine = xamlParentElement.OwnerDocument.CreateTextNode("----------------------");
                xamlParentElement.AppendChild(xamlHorizontalLine);
                xamlLineBreak = xamlParentElement.OwnerDocument.CreateElement(prefix: null, localName: XamlLineBreak, XamlNamespace);
                xamlParentElement.AppendChild(xamlLineBreak);
            }
        }

        /// <summary>
        /// Generates Section or Paragraph element from DIV depending whether it contains any block elements or not.
        /// </summary>
        /// <param name="xamlParentElement">
        /// XmlElement representing XAML parent to which the converted element should be added.
        /// </param>
        /// <param name="htmlElement">
        /// XmlElement representing HTML element to be converted.
        /// </param>
        /// <param name="inheritedProperties">
        /// Properties inherited from parent context.
        /// </param>
        /// <param name="stylesheet"></param>
        /// <param name="sourceContext"></param>
        /// true indicates that a content added by this call contains at least one block element
        /// </param>
        private static void AddSection(XmlElement xamlParentElement, XmlElement htmlElement, Dictionary<string, string> inheritedProperties, CssStylesheet stylesheet, List<XmlElement> sourceContext)
        {
            // Analyze the content of htmlElement to decide what XAML element to choose - Section or Paragraph.
            // If this Div has at least one block child then we need to use Section, otherwise use Paragraph
            bool htmlElementContainsBlocks = false;
            for (var htmlChildNode = htmlElement.FirstChild; htmlChildNode != null; htmlChildNode = htmlChildNode.NextSibling)
            {
                if (htmlChildNode is XmlElement htmlChildElement)
                {
                    string htmlChildName = htmlChildElement.LocalName.ToLowerInvariant();
                    if (HtmlSchema.IsBlockElement(htmlChildName))
                    {
                        htmlElementContainsBlocks = true;
                        break;
                    }
                }
            }

            if (!htmlElementContainsBlocks)
            {
                // The Div does not contain any block elements, so we can treat it as a Paragraph
                AddParagraph(xamlParentElement, htmlElement, inheritedProperties, stylesheet, sourceContext);
            }
            else
            {
                // The Div has some nested blocks, so we treat it as a Section

                // Create currentProperties as a compilation of local and inheritedProperties, set localProperties
                var currentProperties = GetElementProperties(htmlElement, inheritedProperties, out var localProperties, stylesheet, sourceContext);

                // Create a XAML element corresponding to this HTML element
                var xamlElement = xamlParentElement.OwnerDocument.CreateElement(prefix: null, localName: XamlSection, XamlNamespace);
                ApplyLocalProperties(xamlElement, localProperties, isBlock: true);

                // Decide whether we can unwrap this element as not having any formatting significance.
                if (!xamlElement.HasAttributes)
                {
                    // This elements is a group of block elements without any additional formatting.
                    // We can add blocks directly to xamlParentElement and avoid
                    // creating unnecessary Sections nesting.
                    xamlElement = xamlParentElement;
                }

                // Recurse into element subtree
                for (var htmlChildNode = htmlElement.FirstChild; htmlChildNode != null; htmlChildNode = htmlChildNode?.NextSibling)
                {
                    htmlChildNode = AddBlock(xamlElement, htmlChildNode, currentProperties, stylesheet, sourceContext);
                }

                // Add the new element to the parent.
                if (xamlElement != xamlParentElement)
                {
                    xamlParentElement.AppendChild(xamlElement);
                }
            }
        }

        /// <summary>
        /// Generates Paragraph element from P, H1-H7, Center etc.
        /// </summary>
        /// <param name="xamlParentElement">
        /// XmlElement representing XAML parent to which the converted element should be added.
        /// </param>
        /// <param name="htmlElement">
        /// XmlElement representing HTML element to be converted.
        /// </param>
        /// <param name="inheritedProperties">
        /// properties inherited from parent context.
        /// </param>
        /// <param name="stylesheet"></param>
        /// <param name="sourceContext"></param>
        /// true indicates that a content added by this call contains at least one block element
        /// </param>
        private static void AddParagraph(XmlElement xamlParentElement, XmlElement htmlElement, Dictionary<string, string> inheritedProperties, CssStylesheet stylesheet, List<XmlElement> sourceContext)
        {
            // Create currentProperties as a compilation of local and inheritedProperties, set localProperties
            var currentProperties = GetElementProperties(htmlElement, inheritedProperties, out var localProperties, stylesheet, sourceContext);

            // Create a XAML element corresponding to this HTML element
            var xamlElement = xamlParentElement.OwnerDocument.CreateElement(prefix: null, localName: XamlParagraph, XamlNamespace);
            ApplyLocalProperties(xamlElement, localProperties, isBlock: true);

            // Recurse into element subtree
            for (var htmlChildNode = htmlElement.FirstChild; htmlChildNode != null; htmlChildNode = htmlChildNode.NextSibling)
            {
                AddInline(xamlElement, htmlChildNode, currentProperties, stylesheet, sourceContext);
            }

            // Add the new element to the parent.
            xamlParentElement.AppendChild(xamlElement);
        }

        /// <summary>
        /// Creates a Paragraph element and adds all nodes starting from htmlNode
        /// converted to appropriate Inlines.
        /// </summary>
        /// <param name="xamlParentElement">
        /// XmlElement representing XAML parent to which the converted element should be added.
        /// </param>
        /// <param name="htmlNode">
        /// XmlNode starting a collection of implicitly wrapped inlines.
        /// </param>
        /// <param name="inheritedProperties">
        /// properties inherited from parent context.
        /// </param>
        /// <param name="stylesheet"></param>
        /// <param name="sourceContext"></param>
        /// true indicates that a content added by this call contains at least one block element
        /// </param>
        /// <returns>
        /// The last htmlNode added to the implicit paragraph
        /// </returns>
        private static XmlNode AddImplicitParagraph(XmlElement xamlParentElement, XmlNode htmlNode, Dictionary<string, string> inheritedProperties, CssStylesheet stylesheet, List<XmlElement> sourceContext)
        {
            // Collect all non-block elements and wrap them into implicit Paragraph
            var xamlParagraph = xamlParentElement.OwnerDocument.CreateElement(prefix: null, localName: XamlParagraph, XamlNamespace);
            XmlNode lastNodeProcessed = null;
            while (htmlNode != null)
            {
                if (htmlNode is XmlComment htmlComment)
                {
                    DefineInlineFragmentParent(htmlComment, xamlParentElement: null);
                }
                else if (htmlNode is XmlText)
                {
                    if (htmlNode.Value.Trim().Length > 0)
                    {
                        AddTextRun(xamlParagraph, htmlNode.Value);
                    }
                }
                else if (htmlNode is XmlElement xmlElement)
                {
                    string htmlChildName = xmlElement.LocalName.ToLowerInvariant();
                    if (HtmlSchema.IsBlockElement(htmlChildName))
                    {
                        // The sequence of non-blocked inlines ended. Stop implicit loop here.
                        break;
                    }
                    else
                    {
                        AddInline(xamlParagraph, xmlElement, inheritedProperties, stylesheet, sourceContext);
                    }
                }

                // Store last processed node to return it at the end
                lastNodeProcessed = htmlNode;
                htmlNode = htmlNode.NextSibling;
            }

            // Add the Paragraph to the parent
            // If only whitespaces and comments have been encountered,
            // then we have nothing to add in implicit paragraph; forget it.
            if (xamlParagraph.FirstChild != null)
            {
                xamlParentElement.AppendChild(xamlParagraph);
            }

            // Need to return last processed node
            return lastNodeProcessed;
        }

        private static void AddInline(XmlElement xamlParentElement, XmlNode htmlNode, Dictionary<string, string> inheritedProperties, CssStylesheet stylesheet, List<XmlElement> sourceContext)
        {
            if (htmlNode is XmlComment htmlComments)
            {
                DefineInlineFragmentParent(htmlComments, xamlParentElement);
            }
            else if (htmlNode is XmlText)
            {
                AddTextRun(xamlParentElement, htmlNode.Value);
            }
            else if (htmlNode is XmlElement htmlElement)
            {
                // Check whether this is an HTML element
                if (htmlElement.NamespaceURI != HtmlParser.XhtmlNamespace)
                {
                    return; // Skip non-HTML elements
                }

                // Identify element name
                string htmlElementName = htmlElement.LocalName.ToLowerInvariant();

                // Put source element to the stack
                sourceContext.Add(htmlElement);

                switch (htmlElementName)
                {
                    case "a":
                        AddHyperlink(xamlParentElement, htmlElement, inheritedProperties, stylesheet, sourceContext);
                        break;

                    case "img":
                        AddImage(xamlParentElement, htmlElement, inheritedProperties, stylesheet, sourceContext);
                        break;

                    case "br":
                    case "hr":
                        AddBreak(xamlParentElement, htmlElementName);
                        break;

                    default:
                        if (HtmlSchema.IsInlineElement(htmlElementName) || HtmlSchema.IsBlockElement(htmlElementName))
                        {
                            // Note: actually we do not expect block elements here,
                            // but if it happens to be here, we will treat it as a Span.
                            AddSpanOrRun(xamlParentElement, htmlElement, inheritedProperties, stylesheet, sourceContext);
                        }
                        break;
                }

                // Ignore all other elements non-(block/inline/image)

                // Remove the element from the stack
                Debug.Assert(sourceContext.Count > 0 && sourceContext[^1] == htmlElement);
                sourceContext.RemoveAt(sourceContext.Count - 1);
            }
        }

        private static void AddSpanOrRun(XmlElement xamlParentElement, XmlElement htmlElement, Dictionary<string, string> inheritedProperties, CssStylesheet stylesheet, List<XmlElement> sourceContext)
        {
            // Decide what XAML element to use for this inline element.
            // Check whether it contains any nested inlines
            bool elementHasChildren = false;
            for (var htmlNode = htmlElement.FirstChild; htmlNode != null; htmlNode = htmlNode.NextSibling)
            {
                if (htmlNode is XmlElement htmlChildElement)
                {
                    string htmlChildName = htmlChildElement.LocalName.ToLowerInvariant();
                    if (HtmlSchema.IsInlineElement(htmlChildName) || HtmlSchema.IsBlockElement(htmlChildName) ||
                        htmlChildName == "img" || htmlChildName == "br" || htmlChildName == "hr")
                    {
                        elementHasChildren = true;
                        break;
                    }
                }
            }

            string xamlElementName = elementHasChildren ? XamlSpan : XamlRun;

            // Create currentProperties as a compilation of local and inheritedProperties, set localProperties
            var currentProperties = GetElementProperties(htmlElement, inheritedProperties, out var localProperties, stylesheet, sourceContext);

            // Create a XAML element corresponding to this HTML element
            var xamlElement = xamlParentElement.OwnerDocument.CreateElement(prefix: null, localName: xamlElementName, XamlNamespace);
            ApplyLocalProperties(xamlElement, localProperties, isBlock: false);

            // Recurse into element subtree
            for (var htmlChildNode = htmlElement.FirstChild; htmlChildNode != null; htmlChildNode = htmlChildNode.NextSibling)
            {
                AddInline(xamlElement, htmlChildNode, currentProperties, stylesheet, sourceContext);
            }

            // Add the new element to the parent
            xamlParentElement.AppendChild(xamlElement);
        }

        // Adds a text run to a XAML tree.
        private static void AddTextRun(XmlElement xamlElement, string textData)
        {
            // Remove control characters
            for (int i = 0; i < textData.Length; i++)
            {
                if (char.IsControl(textData[i]))
                {
                    textData = textData.Remove(i--, 1);  // decrement i to compensate for character removal
                }
            }

            // Replace No-Breaks by spaces (160 is a code of &nbsp; entity in HTML)
            // This is a work around since WPF/XAML does not support &nbsp.
            textData = textData.Replace((char) 160, ' ');

            if (textData.Length > 0)
            {
                xamlElement.AppendChild(xamlElement.OwnerDocument.CreateTextNode(textData));
            }
        }

        private static void AddHyperlink(XmlElement xamlParentElement, XmlElement htmlElement, Dictionary<string, string> inheritedProperties, CssStylesheet stylesheet, List<XmlElement> sourceContext)
        {
            // Convert href attribute into NavigateUri and TargetName
            string href = GetAttribute(htmlElement, "href");
            if (href == null)
            {
                // When href attribute is missing - ignore the hyperlink
                AddSpanOrRun(xamlParentElement, htmlElement, inheritedProperties, stylesheet, sourceContext);
            }
            else
            {
                // Create currentProperties as a compilation of local and inheritedProperties, set localProperties
                var currentProperties = GetElementProperties(htmlElement, inheritedProperties, out var localProperties, stylesheet, sourceContext);

                // Create a XAML element corresponding to this HTML element
                var xamlElement = xamlParentElement.OwnerDocument.CreateElement(prefix: null, localName: XamlHyperlink, XamlNamespace);
                ApplyLocalProperties(xamlElement, localProperties, isBlock: false);

                string[] hrefParts = href.Split(new char[] { '#' });
                if (hrefParts.Length > 0 && hrefParts[0].Trim().Length > 0)
                {
                    xamlElement.SetAttribute(XamlHyperlinkNavigateUri, hrefParts[0].Trim());
                }

                if (hrefParts.Length == 2 && hrefParts[1].Trim().Length > 0)
                {
                    xamlElement.SetAttribute(XamlHyperlinkTargetName, hrefParts[1].Trim());
                }

                // Recurse into element subtree
                for (var htmlChildNode = htmlElement.FirstChild; htmlChildNode != null; htmlChildNode = htmlChildNode.NextSibling)
                {
                    AddInline(xamlElement, htmlChildNode, currentProperties, stylesheet, sourceContext);
                }

                // Add the new element to the parent.
                xamlParentElement.AppendChild(xamlElement);
            }
        }

        // Called when HTML comment is encountered to store a parent element
        // for the case when the fragment is inline - to extract it to a separate
        // Span wrapper after the conversion.
        private static void DefineInlineFragmentParent(XmlComment htmlComment, XmlElement xamlParentElement)
        {
            if (htmlComment.Value == "StartFragment")
            {
                inlineFragmentParentElement = xamlParentElement;
            }
            else if (htmlComment.Value == "EndFragment")
            {
                if (inlineFragmentParentElement == null && xamlParentElement != null)
                {
                    // Normally this cannot happen if comments produced by correct copying code
                    // in Word or IE, but when it is produced manually then fragment boundary
                    // markers can be inconsistent. In this case StartFragment takes precedence,
                    // but if it is not set, then we get the value from EndFragment marker.
                    inlineFragmentParentElement = xamlParentElement;
                }
            }
        }

        // Extracts a content of an element stored as InlineFragmentParentElement
        // into a separate Span wrapper.
        // Note: when selected content does not cross paragraph boundaries,
        // the fragment is marked within
        private static XmlElement ExtractInlineFragment(XmlElement xamlFlowDocumentElement)
        {
            if (inlineFragmentParentElement != null)
            {
                if (inlineFragmentParentElement.LocalName == XamlSpan)
                {
                    xamlFlowDocumentElement = inlineFragmentParentElement;
                }
                else
                {
                    xamlFlowDocumentElement = xamlFlowDocumentElement.OwnerDocument.CreateElement(prefix: null, localName: XamlSpan, XamlNamespace);
                    while (inlineFragmentParentElement.FirstChild != null)
                    {
                        var copyNode = inlineFragmentParentElement.FirstChild;
                        inlineFragmentParentElement.RemoveChild(copyNode);
                        xamlFlowDocumentElement.AppendChild(copyNode);
                    }
                }
            }

            return xamlFlowDocumentElement;
        }

        private static void AddImage(XmlElement xamlParentElement, XmlElement htmlElement, Dictionary<string, string> inheritedProperties, CssStylesheet stylesheet, List<XmlElement> sourceContext)
        {
            //// TODO: Implement images
        }

        /// <summary>
        /// Converts HTML ul or ol element into XAML list element. During conversion if the ul/ol element has any children
        /// that are not li elements, they are ignored and not added to the list element.
        /// </summary>
        /// <param name="xamlParentElement">
        /// XmlElement representing XAML parent to which the converted element should be added.
        /// </param>
        /// <param name="htmlListElement">
        /// XmlElement representing HTML ul/ol element to be converted.
        /// </param>
        /// <param name="inheritedProperties">
        /// properties inherited from parent context.
        /// </param>
        /// <param name="stylesheet"></param>
        /// <param name="sourceContext"></param>
        private static void AddList(XmlElement xamlParentElement, XmlElement htmlListElement, Dictionary<string, string> inheritedProperties, CssStylesheet stylesheet, List<XmlElement> sourceContext)
        {
            string htmlListElementName = htmlListElement.LocalName.ToLowerInvariant();

            var currentProperties = GetElementProperties(htmlListElement, inheritedProperties, out var localProperties, stylesheet, sourceContext);

            // Create XAML List element
            var xamlListElement = xamlParentElement.OwnerDocument.CreateElement(null, XamlList, XamlNamespace);

            // Set default list markers
            if (htmlListElementName == "ol")
            {
                // Ordered list
                xamlListElement.SetAttribute(XamlListMarkerStyle, XamlListMarkerStyleDecimal);
            }
            else
            {
                // Unordered list - all elements other than OL treated as unordered lists
                xamlListElement.SetAttribute(XamlListMarkerStyle, XamlListMarkerStyleDisc);
            }

            // Apply local properties to list to set marker attribute if specified
            // TODO: Should we have separate list attribute processing function?
            ApplyLocalProperties(xamlListElement, localProperties, isBlock: true);

            // Recurse into list subtree
            for (var htmlChildNode = htmlListElement.FirstChild; htmlChildNode != null; htmlChildNode = htmlChildNode.NextSibling)
            {
                if (htmlChildNode is XmlElement && string.Equals(htmlChildNode.LocalName, "li", StringComparison.OrdinalIgnoreCase))
                {
                    sourceContext.Add((XmlElement) htmlChildNode);
                    AddListItem(xamlListElement, (XmlElement) htmlChildNode, currentProperties, stylesheet, sourceContext);
                    Debug.Assert(sourceContext.Count > 0 && sourceContext[sourceContext.Count - 1] == htmlChildNode);
                    sourceContext.RemoveAt(sourceContext.Count - 1);
                }
                else
                {
                    // Not an li element. Add it to previous ListBoxItem
                    //  We need to append the content to the end
                    // of a previous list item.
                }
            }

            // Add the List element to XAML tree - if it is not empty
            if (xamlListElement.HasChildNodes)
            {
                xamlParentElement.AppendChild(xamlListElement);
            }
        }

        /// <summary>
        /// If li items are found without a parent ul/ol element in HTML string, creates xamlListElement as their parent and adds
        /// them to it. If the previously added node to the same xamlParentElement was a List, adds the elements to that list.
        /// Otherwise, we create a new xamlListElement and add them to it. Elements are added as long as li elements appear sequentially.
        /// The first non-li or text node stops the addition.
        /// </summary>
        /// <param name="xamlParentElement">
        /// Parent element for the list.
        /// </param>
        /// <param name="htmlLIElement">
        /// Start HTML li element without parent list.
        /// </param>
        /// <param name="inheritedProperties">
        /// Properties inherited from parent context.
        /// </param>
        /// <returns>
        /// XmlNode representing the first non-li node in the input after one or more li's have been processed.
        /// </returns>
        private static XmlElement AddOrphanListItems(XmlElement xamlParentElement, XmlElement htmlLIElement, Dictionary<string, string> inheritedProperties, CssStylesheet stylesheet, List<XmlElement> sourceContext)
        {
            Debug.Assert(string.Equals(htmlLIElement.LocalName, "li", StringComparison.OrdinalIgnoreCase));

            XmlElement lastProcessedListItemElement = null;

            // Find out the last element attached to the xamlParentElement, which is the previous sibling of this node
            var xamlListItemElementPreviousSibling = xamlParentElement.LastChild;
            XmlElement xamlListElement;
            if (xamlListItemElementPreviousSibling?.LocalName == XamlList)
            {
                // Previously added XAML element was a list. We will add the new li to it
                xamlListElement = (XmlElement) xamlListItemElementPreviousSibling;
            }
            else
            {
                // No list element near. Create our own.
                xamlListElement = xamlParentElement.OwnerDocument.CreateElement(null, XamlList, XamlNamespace);
            }

            XmlNode htmlChildNode = htmlLIElement;
            string htmlChildNodeName = htmlChildNode?.LocalName.ToLowerInvariant();

            // Current element properties missed here.
            // currentProperties = GetElementProperties(htmlLIElement, inheritedProperties, out localProperties, style sheet);

            // Add li elements to the parent xamlListElement we created as long as they appear sequentially
            // Use properties inherited from xamlParentElement for context
            while (htmlChildNode != null && htmlChildNodeName == "li")
            {
                AddListItem(xamlListElement, (XmlElement) htmlChildNode, inheritedProperties, stylesheet, sourceContext);
                lastProcessedListItemElement = (XmlElement) htmlChildNode;
                htmlChildNode = htmlChildNode.NextSibling;
                htmlChildNodeName = htmlChildNode?.LocalName.ToLowerInvariant();
            }

            return lastProcessedListItemElement;
        }

        /// <summary>
        /// Converts htmlLIElement into XAML ListItem element, and appends it to the parent xamlListElement.
        /// </summary>
        /// <param name="xamlListElement">
        /// XmlElement representing XAML List element to which the converted td/th should be added.
        /// </param>
        /// <param name="htmlLIElement">
        /// XmlElement representing HTML li element to be converted.
        /// </param>
        /// <param name="inheritedProperties">
        /// Properties inherited from parent context.
        /// </param>
        private static void AddListItem(XmlElement xamlListElement, XmlElement htmlLIElement, Dictionary<string, string> inheritedProperties, CssStylesheet stylesheet, List<XmlElement> sourceContext)
        {
            // Parameter validation
            Debug.Assert(xamlListElement != null);
            Debug.Assert(xamlListElement.LocalName == XamlList);
            Debug.Assert(htmlLIElement != null);
            Debug.Assert(string.Equals(htmlLIElement.LocalName, "li", StringComparison.OrdinalIgnoreCase));
            Debug.Assert(inheritedProperties != null);

            var currentProperties = GetElementProperties(htmlLIElement, inheritedProperties, out var localProperties, stylesheet, sourceContext);

            var xamlListItemElement = xamlListElement.OwnerDocument.CreateElement(null, XamlListItem, XamlNamespace);

            // TODO: process local properties for li element

            // Process children of the ListItem
            for (var htmlChildNode = htmlLIElement.FirstChild; htmlChildNode != null; htmlChildNode = htmlChildNode?.NextSibling)
            {
                htmlChildNode = AddBlock(xamlListItemElement, htmlChildNode, currentProperties, stylesheet, sourceContext);
            }

            // Add resulting ListBoxItem to a XAML parent
            xamlListElement.AppendChild(xamlListItemElement);
        }

        /// <summary>
        /// Converts htmlTableElement to a XAML Table element. Adds tbody elements if they are missing so
        /// that a resulting XAML Table element is properly formed.
        /// </summary>
        /// <param name="xamlParentElement">
        /// Parent XAML element to which a converted table must be added.
        /// </param>
        /// <param name="htmlTableElement">
        /// XmlElement representing the HTML table element to be converted.
        /// </param>
        /// <param name="inheritedProperties">
        /// Dictionary representing properties inherited from parent context.
        /// </param>
        private static void AddTable(XmlElement xamlParentElement, XmlElement htmlTableElement, Dictionary<string, string> inheritedProperties, CssStylesheet stylesheet, List<XmlElement> sourceContext)
        {
            // Parameter validation
            Debug.Assert(string.Equals(htmlTableElement.LocalName, "table", StringComparison.OrdinalIgnoreCase));
            Debug.Assert(xamlParentElement != null);
            Debug.Assert(inheritedProperties != null);

            // Create current properties to be used by children as inherited properties, set local properties
            var currentProperties = GetElementProperties(htmlTableElement, inheritedProperties, out var localProperties, stylesheet, sourceContext);

            // TODO: process localProperties for tables to override defaults, decide cell spacing defaults

            // Check if the table contains only one cell - we want to take only its content
            var singleCell = GetCellFromSingleCellTable(htmlTableElement);

            if (singleCell != null)
            {
                // Need to push skipped table elements onto sourceContext
                sourceContext.Add(singleCell);

                // Add the cell's content directly to parent
                for (var htmlChildNode = singleCell.FirstChild; htmlChildNode != null; htmlChildNode = htmlChildNode?.NextSibling)
                {
                    htmlChildNode = AddBlock(xamlParentElement, htmlChildNode, currentProperties, stylesheet, sourceContext);
                }

                Debug.Assert(sourceContext.Count > 0 && sourceContext[sourceContext.Count - 1] == singleCell);
                sourceContext.RemoveAt(sourceContext.Count - 1);
            }
            else
            {
                // Create xamlTableElement
                var xamlTableElement = xamlParentElement.OwnerDocument.CreateElement(null, XamlTable, XamlNamespace);

                // Analyze table structure for column widths and rowspan attributes
                var columnStarts = AnalyzeTableStructure(htmlTableElement, stylesheet);

                // Process COLGROUP & COL elements
                AddColumnInformation(htmlTableElement, xamlTableElement, columnStarts, currentProperties, stylesheet, sourceContext);

                // Process table body - TBODY and TR elements
                var htmlChildNode = htmlTableElement.FirstChild;

                while (htmlChildNode != null)
                {
                    string htmlChildName = htmlChildNode.LocalName.ToLowerInvariant();

                    // Process the element
                    if (htmlChildName == "tbody" || htmlChildName == "thead" || htmlChildName == "tfoot")
                    {
                        // Add more special processing for TableHeader and TableFooter
                        var xamlTableBodyElement = xamlTableElement.OwnerDocument.CreateElement(null, XamlTableRowGroup, XamlNamespace);
                        xamlTableElement.AppendChild(xamlTableBodyElement);

                        sourceContext.Add((XmlElement) htmlChildNode);

                        // Get properties of HTML tbody element
                        var tbodyElementCurrentProperties = GetElementProperties((XmlElement) htmlChildNode, currentProperties, out var tbodyElementLocalProperties, stylesheet, sourceContext);
                        //// TODO: apply local properties for tbody

                        // Process children of htmlChildNode, which is tbody, for tr elements
                        AddTableRowsToTableBody(xamlTableBodyElement, htmlChildNode.FirstChild, tbodyElementCurrentProperties, columnStarts, stylesheet, sourceContext);
                        if (xamlTableBodyElement.HasChildNodes)
                        {
                            xamlTableElement.AppendChild(xamlTableBodyElement);
                            //// else: if there is no TRs in this TBody, we simply ignore it
                        }

                        Debug.Assert(sourceContext.Count > 0 && sourceContext[sourceContext.Count - 1] == htmlChildNode);
                        sourceContext.RemoveAt(sourceContext.Count - 1);

                        htmlChildNode = htmlChildNode.NextSibling;
                    }
                    else if (htmlChildName == "tr")
                    {
                        // Tbody is not present, but tr element is present. Tr is wrapped in tbody
                        var xamlTableBodyElement = xamlTableElement.OwnerDocument.CreateElement(null, XamlTableRowGroup, XamlNamespace);

                        // We use currentProperties of xamlTableElement when adding rows since the tbody element is artificially created and has
                        // no properties of its own
                        htmlChildNode = AddTableRowsToTableBody(xamlTableBodyElement, htmlChildNode, currentProperties, columnStarts, stylesheet, sourceContext);
                        if (xamlTableBodyElement.HasChildNodes)
                        {
                            xamlTableElement.AppendChild(xamlTableBodyElement);
                        }
                    }
                    else
                    {
                        // Element is not tbody or tr. Ignore it.
                        // TODO: add processing for thead, tfoot elements and recovery for td elements
                        htmlChildNode = htmlChildNode.NextSibling;
                    }
                }

                if (xamlTableElement.HasChildNodes)
                {
                    xamlParentElement.AppendChild(xamlTableElement);
                }
            }
        }

        private static XmlElement GetCellFromSingleCellTable(XmlElement htmlTableElement)
        {
            XmlElement singleCell = null;

            for (var tableChild = htmlTableElement.FirstChild; tableChild != null; tableChild = tableChild.NextSibling)
            {
                string elementName = tableChild.LocalName.ToLowerInvariant();
                if (elementName == "tbody" || elementName == "thead" || elementName == "tfoot")
                {
                    if (singleCell != null)
                    {
                        return null;
                    }

                    for (var tbodyChild = tableChild.FirstChild; tbodyChild != null; tbodyChild = tbodyChild.NextSibling)
                    {
                        if (string.Equals(tbodyChild.LocalName, "tr", StringComparison.OrdinalIgnoreCase))
                        {
                            if (singleCell != null)
                            {
                                return null;
                            }

                            for (var trChild = tbodyChild.FirstChild; trChild != null; trChild = trChild.NextSibling)
                            {
                                string cellName = trChild.LocalName.ToLowerInvariant();
                                if (cellName == "td" || cellName == "th")
                                {
                                    if (singleCell != null)
                                    {
                                        return null;
                                    }

                                    singleCell = (XmlElement) trChild;
                                }
                            }
                        }
                    }
                }
                else if (string.Equals(tableChild.LocalName, "tr", StringComparison.OrdinalIgnoreCase))
                {
                    if (singleCell != null)
                    {
                        return null;
                    }

                    for (var trChild = tableChild.FirstChild; trChild != null; trChild = trChild.NextSibling)
                    {
                        string cellName = trChild.LocalName.ToLowerInvariant();
                        if (cellName == "td" || cellName == "th")
                        {
                            if (singleCell != null)
                            {
                                return null;
                            }

                            singleCell = (XmlElement) trChild;
                        }
                    }
                }
            }

            return singleCell;
        }

        /// <summary>
        /// Processes the information about table columns - COLGROUP and COL HTML elements.
        /// </summary>
        /// <param name="htmlTableElement">
        /// XmlElement representing a source HTML table.
        /// </param>
        /// <param name="xamlTableElement">
        /// XmlElement representing a resulting XAML table.
        /// </param>
        /// <param name="columnStartsAllRows">
        /// Array of doubles - column start coordinates.
        /// Can be null, which means that column size information is not available
        /// and we must use source colgroup/col information.
        /// In case when it's not null, we will ignore source colgroup/col information.
        /// </param>
        /// <param name="currentProperties"></param>
        /// <param name="stylesheet"></param>
        /// <param name="sourceContext"></param>
        private static void AddColumnInformation(XmlElement htmlTableElement, XmlElement xamlTableElement, List<double> columnStartsAllRows, Dictionary<string, string> currentProperties, CssStylesheet stylesheet, List<XmlElement> sourceContext)
        {
            // Add column information
            if (columnStartsAllRows != null)
            {
                // We have consistent information derived from table cells; use it
                // The last element in columnStarts represents the end of the table
                for (int columnIndex = 0; columnIndex < columnStartsAllRows.Count - 1; columnIndex++)
                {
                    var xamlColumnElement = xamlTableElement.OwnerDocument.CreateElement(null, XamlTableColumn, XamlNamespace);
                    xamlColumnElement.SetAttribute(XamlWidth, (columnStartsAllRows[columnIndex + 1] - columnStartsAllRows[columnIndex]).ToString());
                    xamlTableElement.AppendChild(xamlColumnElement);
                }
            }
            else
            {
                // We do not have consistent information from table cells;
                // Translate blindly colgroups from HTML.
                for (var htmlChildNode = htmlTableElement.FirstChild; htmlChildNode != null; htmlChildNode = htmlChildNode.NextSibling)
                {
                    if (string.Equals(htmlChildNode.LocalName, "colgroup", StringComparison.OrdinalIgnoreCase))
                    {
                        // TODO: add column width information to this function as a parameter and process it
                        AddTableColumnGroup(xamlTableElement, (XmlElement) htmlChildNode, currentProperties, stylesheet, sourceContext);
                    }
                    else if (string.Equals(htmlChildNode.LocalName, "col", StringComparison.OrdinalIgnoreCase))
                    {
                        AddTableColumn(xamlTableElement, (XmlElement) htmlChildNode, currentProperties, stylesheet, sourceContext);
                    }
                    else if (htmlChildNode is XmlElement)
                    {
                        // Some element which belongs to table body. Stop column loop.
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Converts htmlColgroupElement into XAML TableColumnGroup element, and appends it to the parent xamlTableElement.
        /// </summary>
        /// <param name="xamlTableElement">
        /// XmlElement representing XAML Table element to which the converted column group should be added.
        /// </param>
        /// <param name="htmlColgroupElement">
        /// XmlElement representing HTML colgroup element to be converted.
        /// <param name="inheritedProperties">
        /// Properties inherited from parent context.
        /// </param>
        private static void AddTableColumnGroup(XmlElement xamlTableElement, XmlElement htmlColgroupElement, Dictionary<string, string> inheritedProperties, CssStylesheet stylesheet, List<XmlElement> sourceContext)
        {
            var currentProperties = GetElementProperties(htmlColgroupElement, inheritedProperties, out var localProperties, stylesheet, sourceContext);

            // TODO: process local properties for colgroup

            // Process children of colgroup. Colgroup may contain only col elements.
            for (var htmlNode = htmlColgroupElement.FirstChild; htmlNode != null; htmlNode = htmlNode.NextSibling)
            {
                if (htmlNode is XmlElement && string.Equals(htmlNode.LocalName, "col", StringComparison.OrdinalIgnoreCase))
                {
                    AddTableColumn(xamlTableElement, (XmlElement) htmlNode, currentProperties, stylesheet, sourceContext);
                }
            }
        }

        /// <summary>
        /// Converts htmlColElement into XAML TableColumn element, and appends it to the parent
        /// xamlTableColumnGroupElement.
        /// </summary>
        /// <param name="xamlTableElement"></param>
        /// <param name="htmlColElement">
        /// XmlElement representing HTML col element to be converted.
        /// </param>
        /// <param name="inheritedProperties">
        /// properties inherited from parent context.
        /// </param>
        /// <param name="stylesheet"></param>
        /// <param name="sourceContext"></param>
        private static void AddTableColumn(XmlElement xamlTableElement, XmlElement htmlColElement, Dictionary<string, string> inheritedProperties, CssStylesheet stylesheet, List<XmlElement> sourceContext)
        {
            var currentProperties = GetElementProperties(htmlColElement, inheritedProperties, out var localProperties, stylesheet, sourceContext);

            var xamlTableColumnElement = xamlTableElement.OwnerDocument.CreateElement(null, XamlTableColumn, XamlNamespace);

            //// TODO: process local properties for TableColumn element

            // Col is an empty element, with no subtree
            xamlTableElement.AppendChild(xamlTableColumnElement);
        }

        /// <summary>
        /// Adds TableRow elements to xamlTableBodyElement. The rows are converted from HTML tr elements that
        /// may be the children of an HTML tbody element or an HTML table element with tbody missing.
        /// </summary>
        /// <param name="xamlTableBodyElement">
        /// XmlElement representing XAML TableRowGroup element to which the converted rows should be added.
        /// </param>
        /// <param name="htmlTRStartNode">
        /// XmlElement representing the first tr child of the tbody element to be read.
        /// </param>
        /// <param name="currentProperties">
        /// Dictionary representing current properties of the tbody element that are generated and applied in the
        /// AddTable function; to be used as inheritedProperties when adding tr elements.
        /// </param>
        /// <param name="columnStarts"></param>
        /// <param name="stylesheet"></param>
        /// <param name="sourceContext"></param>
        /// <returns>
        /// XmlNode representing the current position of the iterator among tr elements.
        /// </returns>
        private static XmlNode AddTableRowsToTableBody(XmlElement xamlTableBodyElement, XmlNode htmlTRStartNode, Dictionary<string, string> currentProperties, List<double> columnStarts, CssStylesheet stylesheet, List<XmlElement> sourceContext)
        {
            // Parameter validation
            Debug.Assert(xamlTableBodyElement.LocalName == XamlTableRowGroup);
            Debug.Assert(currentProperties != null);

            // Initialize child node for iterating through children to the first tr element
            var htmlChildNode = htmlTRStartNode;
            List<int> activeRowSpans = null;
            if (columnStarts != null)
            {
                activeRowSpans = new List<int>();
                InitializeActiveRowSpans(activeRowSpans, columnStarts.Count);
            }

            while (htmlChildNode != null && !string.Equals(htmlChildNode.LocalName, "tbody", StringComparison.OrdinalIgnoreCase))
            {
                if (string.Equals(htmlChildNode.LocalName, "tr", StringComparison.OrdinalIgnoreCase))
                {
                    var xamlTableRowElement = xamlTableBodyElement.OwnerDocument.CreateElement(null, XamlTableRow, XamlNamespace);

                    sourceContext.Add((XmlElement) htmlChildNode);

                    // Get tr element properties
                    var trElementCurrentProperties = GetElementProperties((XmlElement) htmlChildNode, currentProperties, out var trElementLocalProperties, stylesheet, sourceContext);
                    //// TODO: apply local properties to tr element

                    AddTableCellsToTableRow(xamlTableRowElement, htmlChildNode.FirstChild, trElementCurrentProperties, columnStarts, activeRowSpans, stylesheet, sourceContext);
                    if (xamlTableRowElement.HasChildNodes)
                    {
                        xamlTableBodyElement.AppendChild(xamlTableRowElement);
                    }

                    Debug.Assert(sourceContext.Count > 0 && sourceContext[sourceContext.Count - 1] == htmlChildNode);
                    sourceContext.RemoveAt(sourceContext.Count - 1);

                    // Advance
                    htmlChildNode = htmlChildNode.NextSibling;
                }
                else if (string.Equals(htmlChildNode.LocalName, "td", StringComparison.OrdinalIgnoreCase))
                {
                    // Tr element is not present. We create one and add td elements to it
                    var xamlTableRowElement = xamlTableBodyElement.OwnerDocument.CreateElement(null, XamlTableRow, XamlNamespace);

                    // This is incorrect formatting and the column starts should not be set in this case
                    Debug.Assert(columnStarts == null);

                    htmlChildNode = AddTableCellsToTableRow(xamlTableRowElement, htmlChildNode, currentProperties, columnStarts, activeRowSpans, stylesheet, sourceContext);
                    if (xamlTableRowElement.HasChildNodes)
                    {
                        xamlTableBodyElement.AppendChild(xamlTableRowElement);
                    }
                }
                else
                {
                    // Not a tr or td  element. Ignore it.
                    // TODO: consider better recovery here
                    htmlChildNode = htmlChildNode.NextSibling;
                }
            }

            return htmlChildNode;
        }

        /// <summary>
        /// Adds TableCell elements to xamlTableRowElement.
        /// </summary>
        /// <param name="xamlTableRowElement">
        /// XmlElement representing XAML TableRow element to which the converted cells should be added.
        /// </param>
        /// <param name="htmlTDStartNode">
        /// XmlElement representing the child of tr or tbody element from which we should start adding td elements.
        /// </param>
        /// <param name="currentProperties">
        /// properties of the current HTML tr element to which cells are to be added.
        /// </param>
        /// <returns>
        /// XmlElement representing the current position of the iterator among the children of the parent HTML tbody/tr element.
        /// </returns>
        private static XmlNode AddTableCellsToTableRow(XmlElement xamlTableRowElement, XmlNode htmlTDStartNode, Dictionary<string, string> currentProperties, List<double> columnStarts, List<int> activeRowSpans, CssStylesheet stylesheet, List<XmlElement> sourceContext)
        {
            // parameter validation
            Debug.Assert(xamlTableRowElement.LocalName == XamlTableRow);
            Debug.Assert(currentProperties != null);
            if (columnStarts != null)
            {
                Debug.Assert(activeRowSpans.Count == columnStarts.Count);
            }

            var htmlChildNode = htmlTDStartNode;
            int columnIndex = 0;

            while (htmlChildNode != null && !string.Equals(htmlChildNode.LocalName, "tr", StringComparison.OrdinalIgnoreCase) && !string.Equals(htmlChildNode.LocalName, "tbody", StringComparison.OrdinalIgnoreCase) && !string.Equals(htmlChildNode.LocalName, "thead", StringComparison.OrdinalIgnoreCase) && !string.Equals(htmlChildNode.LocalName, "tfoot", StringComparison.OrdinalIgnoreCase))
            {
                if (string.Equals(htmlChildNode.LocalName, "td", StringComparison.OrdinalIgnoreCase) || string.Equals(htmlChildNode.LocalName, "th", StringComparison.OrdinalIgnoreCase))
                {
                    var xamlTableCellElement = xamlTableRowElement.OwnerDocument.CreateElement(null, XamlTableCell, XamlNamespace);

                    sourceContext.Add((XmlElement) htmlChildNode);

                    var tdElementCurrentProperties = GetElementProperties((XmlElement) htmlChildNode, currentProperties, out var tdElementLocalProperties, stylesheet, sourceContext);

                    // TODO: determine if localProperties can be used instead of htmlChildNode in this call, and if they can,
                    // make necessary changes and use them instead.
                    ApplyPropertiesToTableCellElement((XmlElement) htmlChildNode, xamlTableCellElement);

                    if (columnStarts != null)
                    {
                        Debug.Assert(columnIndex < columnStarts.Count - 1);
                        while (columnIndex < activeRowSpans.Count && activeRowSpans[columnIndex] > 0)
                        {
                            activeRowSpans[columnIndex]--;
                            Debug.Assert(activeRowSpans[columnIndex] >= 0);
                            columnIndex++;
                        }

                        Debug.Assert(columnIndex < columnStarts.Count - 1);
                        double columnWidth = GetColumnWidth((XmlElement) htmlChildNode);
                        int columnSpan = CalculateColumnSpan(columnIndex, columnWidth, columnStarts);
                        int rowSpan = GetRowSpan((XmlElement) htmlChildNode);

                        // Column cannot have no span
                        Debug.Assert(columnSpan > 0);
                        Debug.Assert(columnIndex + columnSpan < columnStarts.Count);

                        xamlTableCellElement.SetAttribute(XamlTableCellColumnSpan, columnSpan.ToString());

                        // Apply row span
                        for (int spannedColumnIndex = columnIndex; spannedColumnIndex < columnIndex + columnSpan; spannedColumnIndex++)
                        {
                            Debug.Assert(spannedColumnIndex < activeRowSpans.Count);
                            activeRowSpans[spannedColumnIndex] = rowSpan - 1;
                            Debug.Assert(activeRowSpans[spannedColumnIndex] >= 0);
                        }

                        columnIndex += columnSpan;
                    }

                    AddDataToTableCell(xamlTableCellElement, htmlChildNode.FirstChild, tdElementCurrentProperties, stylesheet, sourceContext);
                    if (xamlTableCellElement.HasChildNodes)
                    {
                        xamlTableRowElement.AppendChild(xamlTableCellElement);
                    }

                    Debug.Assert(sourceContext.Count > 0 && sourceContext[sourceContext.Count - 1] == htmlChildNode);
                    sourceContext.RemoveAt(sourceContext.Count - 1);

                    htmlChildNode = htmlChildNode.NextSibling;
                }
                else
                {
                    // Not td element. Ignore it.
                    // TODO: Consider better recovery
                    htmlChildNode = htmlChildNode.NextSibling;
                }
            }

            return htmlChildNode;
        }

        /// <summary>
        /// adds table cell data to xamlTableCellElement.
        /// </summary>
        /// <param name="xamlTableCellElement">
        /// XmlElement representing XAML TableCell element to which the converted data should be added.
        /// </param>
        /// <param name="htmlDataStartNode">
        /// XmlElement representing the start element of data to be added to xamlTableCellElement.
        /// </param>
        /// <param name="currentProperties">
        /// Current properties for the HTML td/th element corresponding to xamlTableCellElement.
        /// </param>
        private static void AddDataToTableCell(XmlElement xamlTableCellElement, XmlNode htmlDataStartNode, Dictionary<string, string> currentProperties, CssStylesheet stylesheet, List<XmlElement> sourceContext)
        {
            // Parameter validation
            Debug.Assert(xamlTableCellElement.LocalName == XamlTableCell);
            Debug.Assert(currentProperties != null);

            for (var htmlChildNode = htmlDataStartNode; htmlChildNode != null; htmlChildNode = htmlChildNode?.NextSibling)
            {
                // Process a new HTML element and add it to the td element
                htmlChildNode = AddBlock(xamlTableCellElement, htmlChildNode, currentProperties, stylesheet, sourceContext);
            }
        }

        /// <summary>
        /// Performs a parsing pass over a table to read information about column width and rowspan attributes. This information
        /// is used to determine the starting point of each column.
        /// </summary>
        /// <param name="htmlTableElement">
        /// XmlElement representing HTML table whose structure is to be analyzed.
        /// </param>
        /// <returns>
        /// ArrayList of type double which contains the function output. If analysis is successful, this ArrayList contains
        /// all the points which are the starting position of any column in the table, ordered from left to right.
        /// In case if analysis was impossible we return null.
        /// </returns>
        private static List<double> AnalyzeTableStructure(XmlElement htmlTableElement, CssStylesheet stylesheet)
        {
            // Parameter validation
            Debug.Assert(string.Equals(htmlTableElement.LocalName, "table", StringComparison.OrdinalIgnoreCase));
            if (!htmlTableElement.HasChildNodes)
            {
                return null;
            }

            bool columnWidthsAvailable = true;

            var columnStarts = new List<double>();
            var activeRowSpans = new List<int>();
            Debug.Assert(columnStarts.Count == activeRowSpans.Count);

            var htmlChildNode = htmlTableElement.FirstChild;
            double tableWidth = 0;  // Keep track of table width which is the width of its widest row

            // Analyze tbody and tr elements
            while (htmlChildNode != null && columnWidthsAvailable)
            {
                Debug.Assert(columnStarts.Count == activeRowSpans.Count);

                switch (htmlChildNode.LocalName.ToLowerInvariant())
                {
                    case "tbody":
                        // Tbody element, we should analyze its children for trows
                        double tbodyWidth = AnalyzeTbodyStructure((XmlElement) htmlChildNode, columnStarts, activeRowSpans, tableWidth, stylesheet);
                        if (tbodyWidth > tableWidth)
                        {
                            // Table width must be increased to supported newly added wide row
                            tableWidth = tbodyWidth;
                        }
                        else if (tbodyWidth == 0)
                        {
                            // Tbody analysis may return 0, probably due to unprocessable format.
                            // We should also fail.
                            columnWidthsAvailable = false; // interrupt the analysis
                        }
                        break;

                    case "tr":
                        // Table row. Analyze column structure within row directly
                        double trWidth = AnalyzeTRStructure((XmlElement) htmlChildNode, columnStarts, activeRowSpans, tableWidth, stylesheet);
                        if (trWidth > tableWidth)
                        {
                            tableWidth = trWidth;
                        }
                        else if (trWidth == 0)
                        {
                            columnWidthsAvailable = false; // interrupt the analysis
                        }
                        break;

                    case "td":
                        // Incorrect formatting, too deep to analyze at this level. Return null.
                        // TODO: implement analysis at this level, possibly by creating a new tr
                        columnWidthsAvailable = false; // interrupt the analysis
                        break;

                    default:
                        // Element should not occur directly in table. Ignore it.
                        break;
                }

                htmlChildNode = htmlChildNode.NextSibling;
            }

            if (columnWidthsAvailable)
            {
                // Add an item for whole table width
                columnStarts.Add(tableWidth);
                VerifyColumnStartsAscendingOrder(columnStarts);
            }
            else
            {
                columnStarts = null;
            }

            return columnStarts;
        }

        /// <summary>
        /// Performs a parsing pass over a tbody to read information about column width and rowspan attributes. Information read about width
        /// attributes is stored in the reference ArrayList parameter columnStarts, which contains a list of all starting
        /// positions of all columns in the table, ordered from left to right. Row spans are taken into consideration when
        /// computing column starts.
        /// </summary>
        /// <param name="htmlTbodyElement">
        /// XmlElement representing HTML tbody whose structure is to be analyzed.
        /// </param>
        /// <param name="columnStarts">
        /// ArrayList of type double which contains the function output. If analysis fails, this parameter is set to null.
        /// </param>
        /// <param name="tableWidth">
        /// Current width of the table. This is used to determine if a new column when added to the end of table should
        /// come after the last column in the table or is actually splitting the last column in two. If it is only splitting
        /// the last column it should inherit row span for that column.
        /// </param>
        /// <returns>
        /// Calculated width of a tbody.
        /// In case of non-analyzable column width structure return 0;.
        /// </returns>
        private static double AnalyzeTbodyStructure(XmlElement htmlTbodyElement, List<double> columnStarts, List<int> activeRowSpans, double tableWidth, CssStylesheet stylesheet)
        {
            // Parameter validation
            Debug.Assert(string.Equals(htmlTbodyElement.LocalName, "tbody", StringComparison.OrdinalIgnoreCase));
            Debug.Assert(columnStarts != null);

            double tbodyWidth = 0;
            bool columnWidthsAvailable = true;

            if (!htmlTbodyElement.HasChildNodes)
            {
                return tbodyWidth;
            }

            // Set active row spans to 0 - thus ignoring row spans crossing tbody boundaries
            ClearActiveRowSpans(activeRowSpans);

            var htmlChildNode = htmlTbodyElement.FirstChild;

            // Analyze tr elements
            while (htmlChildNode != null && columnWidthsAvailable)
            {
                switch (htmlChildNode.LocalName.ToLowerInvariant())
                {
                    case "tr":
                        double trWidth = AnalyzeTRStructure((XmlElement) htmlChildNode, columnStarts, activeRowSpans, tbodyWidth, stylesheet);
                        if (trWidth > tbodyWidth)
                        {
                            tbodyWidth = trWidth;
                        }
                        break;

                    case "td":
                        columnWidthsAvailable = false; // interrupt the analysis
                        break;
                }

                htmlChildNode = htmlChildNode.NextSibling;
            }

            // Set active row spans to 0 - thus ignoring row spans crossing tbody boundaries
            ClearActiveRowSpans(activeRowSpans);

            return columnWidthsAvailable ? tbodyWidth : 0;
        }

        /// <summary>
        /// Performs a parsing pass over a tr element to read information about column width and rowspan attributes.
        /// </summary>
        /// <param name="htmlTRElement">
        /// XmlElement representing HTML tr element whose structure is to be analyzed.
        /// </param>
        /// <param name="columnStarts">
        /// ArrayList of type double which contains the function output. If analysis is successful, this ArrayList contains
        /// all the points which are the starting position of any column in the tr, ordered from left to right. If analysis fails,
        /// the ArrayList is set to null.
        /// </param>
        /// <param name="activeRowSpans">
        /// ArrayList representing all columns currently spanned by an earlier row span attribute. These columns should
        /// not be used for data in this row. The ArrayList actually contains notation for all columns in the table, if the
        /// active row span is set to 0 that column is not presently spanned but if it is > 0 the column is presently spanned.
        /// </param>
        /// <param name="tableWidth">
        /// Double value representing the current width of the table.
        /// Return 0 if analysis was unsuccessful.
        /// </param>
        private static double AnalyzeTRStructure(XmlElement htmlTRElement, List<double> columnStarts, List<int> activeRowSpans, double tableWidth, CssStylesheet stylesheet)
        {
            double columnWidth;

            // Parameter validation
            Debug.Assert(string.Equals(htmlTRElement.LocalName, "tr", StringComparison.OrdinalIgnoreCase));
            Debug.Assert(columnStarts != null);
            Debug.Assert(activeRowSpans != null);
            Debug.Assert(columnStarts.Count == activeRowSpans.Count);

            if (!htmlTRElement.HasChildNodes)
            {
                return 0;
            }

            bool columnWidthsAvailable = true;

            double columnStart = 0; // starting position of current column
            var htmlChildNode = htmlTRElement.FirstChild;
            int columnIndex = 0;

            // Skip spanned columns to get to real column start
            if (columnIndex < activeRowSpans.Count)
            {
                Debug.Assert(columnStarts[columnIndex] >= columnStart);
                if (columnStarts[columnIndex] == columnStart)
                {
                    // The new column may be in a spanned area
                    while (columnIndex < activeRowSpans.Count && activeRowSpans[columnIndex] > 0)
                    {
                        activeRowSpans[columnIndex]--;
                        Debug.Assert(activeRowSpans[columnIndex] >= 0);
                        columnIndex++;
                        columnStart = columnStarts[columnIndex];
                    }
                }
            }

            while (htmlChildNode != null && columnWidthsAvailable)
            {
                Debug.Assert(columnStarts.Count == activeRowSpans.Count);

                VerifyColumnStartsAscendingOrder(columnStarts);

                switch (htmlChildNode.LocalName.ToLowerInvariant())
                {
                    case "td":
                        Debug.Assert(columnIndex <= columnStarts.Count);
                        if (columnIndex < columnStarts.Count)
                        {
                            Debug.Assert(columnStart <= columnStarts[columnIndex]);
                            if (columnStart < columnStarts[columnIndex])
                            {
                                columnStarts.Insert(columnIndex, columnStart);

                                // There can be no row spans now - the column data will appear here
                                // Row spans may appear only during the column analysis
                                activeRowSpans.Insert(columnIndex, 0);
                            }
                        }
                        else
                        {
                            // Column start is greater than all previous starts. Row span must still be 0 because
                            // we are either adding after another column of the same row, in which case it should not inherit
                            // the previous column's span. Otherwise we are adding after the last column of some previous
                            // row, and assuming the table widths line up, we should not be spanned by it. If there is
                            // an incorrect table structure where a columns starts in the middle of a row span, we do not
                            // guarantee correct output
                            columnStarts.Add(columnStart);
                            activeRowSpans.Add(0);
                        }

                        columnWidth = GetColumnWidth((XmlElement) htmlChildNode);
                        if (columnWidth != -1)
                        {
                            int nextColumnIndex;
                            int rowSpan = GetRowSpan((XmlElement) htmlChildNode);

                            nextColumnIndex = GetNextColumnIndex(columnIndex, columnWidth, columnStarts, activeRowSpans);
                            if (nextColumnIndex != -1)
                            {
                                // Entire column width can be processed without hitting conflicting row span. This means that
                                // column widths line up and we can process them
                                Debug.Assert(nextColumnIndex <= columnStarts.Count);

                                // Apply row span to affected columns
                                for (int spannedColumnIndex = columnIndex; spannedColumnIndex < nextColumnIndex; spannedColumnIndex++)
                                {
                                    activeRowSpans[spannedColumnIndex] = rowSpan - 1;
                                    Debug.Assert(activeRowSpans[spannedColumnIndex] >= 0);
                                }

                                columnIndex = nextColumnIndex;

                                // Calculate columnsStart for the next cell
                                columnStart += columnWidth;

                                if (columnIndex < activeRowSpans.Count)
                                {
                                    Debug.Assert(columnStarts[columnIndex] >= columnStart);
                                    if (columnStarts[columnIndex] == columnStart)
                                    {
                                        // The new column may be in a spanned area
                                        while (columnIndex < activeRowSpans.Count && activeRowSpans[columnIndex] > 0)
                                        {
                                            activeRowSpans[columnIndex]--;
                                            Debug.Assert(activeRowSpans[columnIndex] >= 0);
                                            columnIndex++;
                                            columnStart = columnStarts[columnIndex];
                                        }
                                    }

                                    // else: the new column does not start at the same time as a pre existing column
                                    // so we don't have to check it for active row spans, it starts in the middle
                                    // of another column which has been checked already by the GetNextColumnIndex function
                                }
                            }
                            else
                            {
                                // Full column width cannot be processed without a pre existing row span.
                                // We cannot analyze widths
                                columnWidthsAvailable = false;
                            }
                        }
                        else
                        {
                            // Incorrect column width, stop processing
                            columnWidthsAvailable = false;
                        }
                        break;
                }

                htmlChildNode = htmlChildNode.NextSibling;
            }

            // The width of the tr element is the position at which it's last td element ends, which is calculated in
            // the columnStart value after each td element is processed
            return (columnWidthsAvailable ? columnStart : 0);
        }

        /// <summary>
        /// Gets row span attribute from htmlTDElement. Returns an integer representing the value of the rowspan attribute.
        /// Default value if attribute is not specified or if it is invalid is 1.
        /// </summary>
        /// <param name="htmlTDElement">
        /// HTML td element to be searched for rowspan attribute.
        /// </param>
        private static int GetRowSpan(XmlElement htmlTDElement)
        {
            string rowSpanAsString;
            int rowSpan;

            rowSpanAsString = GetAttribute(htmlTDElement, "rowspan");
            if (rowSpanAsString != null)
            {
                if (!int.TryParse(rowSpanAsString, out rowSpan))
                {
                    // Ignore invalid value of rowspan; treat it as 1
                    rowSpan = 1;
                }
            }
            else
            {
                // No row span, default is 1
                rowSpan = 1;
            }

            return rowSpan;
        }

        /// <summary>
        /// Gets index at which a column should be inserted into the columnStarts ArrayList. This is
        /// decided by the value columnStart. The columnStarts ArrayList is ordered in ascending order.
        /// Returns an integer representing the index at which the column should be inserted.
        /// </summary>
        /// <param name="columnIndex">
        /// Int representing the current column index. This acts as a clue while finding the insertion index.
        /// If the value of columnStarts at columnIndex is the same as columnStart, then this position already exists
        /// in the array and we can just return columnIndex.
        /// </param>
        /// <param name="columnWidth"></param>
        /// <param name="columnStarts">
        /// Array list representing starting coordinates of all columns in the table.
        /// </param>
        /// <param name="activeRowSpans"></param>
        private static int GetNextColumnIndex(int columnIndex, double columnWidth, List<double> columnStarts, List<int> activeRowSpans)
        {
            double columnStart;
            int spannedColumnIndex;

            // Parameter validation
            Debug.Assert(columnStarts != null);
            Debug.Assert(columnIndex >= 0 && columnIndex <= columnStarts.Count);
            Debug.Assert(columnWidth > 0);

            columnStart = columnStarts[columnIndex];
            spannedColumnIndex = columnIndex + 1;

            while (spannedColumnIndex < columnStarts.Count && columnStarts[spannedColumnIndex] < columnStart + columnWidth && spannedColumnIndex != -1)
            {
                if (activeRowSpans[spannedColumnIndex] > 0)
                {
                    // The current column should span this area, but something else is already spanning it
                    // Not analyzable
                    spannedColumnIndex = -1;
                }
                else
                {
                    spannedColumnIndex++;
                }
            }

            return spannedColumnIndex;
        }

        /// <summary>
        /// Used for clearing activeRowSpans array in the beginning/end of each tbody.
        /// </summary>
        /// <param name="activeRowSpans">
        /// ArrayList representing currently active row spans.
        /// </param>
        private static void ClearActiveRowSpans(List<int> activeRowSpans)
        {
            for (int columnIndex = 0; columnIndex < activeRowSpans.Count; columnIndex++)
            {
                activeRowSpans[columnIndex] = 0;
            }
        }

        /// <summary>
        /// Used for initializing activeRowSpans array in the before adding rows to tbody element.
        /// </summary>
        /// <param name="activeRowSpans">
        /// ArrayList representing currently active row spans.
        /// </param>
        /// <param name="count">
        /// Size to be give to array list.
        /// </param>
        private static void InitializeActiveRowSpans(List<int> activeRowSpans, int count)
        {
            for (int columnIndex = 0; columnIndex < count; columnIndex++)
            {
                activeRowSpans.Add(0);
            }
        }

        /// <summary>
        /// Calculates width of next TD element based on starting position of current element and it's width, which
        /// is calculated by the function.
        /// </summary>
        /// <param name="htmlTDElement">
        /// XmlElement representing HTML td element whose width is to be read.
        /// </param>
        /// <param name="columnStart">
        /// Starting position of current column.
        /// </param>
        private static double GetNextColumnStart(XmlElement htmlTDElement, double columnStart)
        {
            double columnWidth;

            // Parameter validation
            Debug.Assert(string.Equals(htmlTDElement.LocalName, "td", StringComparison.OrdinalIgnoreCase) || string.Equals(htmlTDElement.LocalName, "th", StringComparison.OrdinalIgnoreCase));
            Debug.Assert(columnStart >= 0);

            columnWidth = GetColumnWidth(htmlTDElement);

            if (columnWidth == -1)
            {
                return -1;
            }
            else
            {
                return columnStart + columnWidth;
            }
        }

        private static double GetColumnWidth(XmlElement htmlTDElement)
        {
            // Get string value for the width
            string columnWidthAsString = GetAttribute(htmlTDElement, "width")
                                         ?? GetCssAttribute(GetAttribute(htmlTDElement, "style"), "width");

            // We do not allow column width to be 0, if specified as 0 we will fail to record it
            if (!TryGetLengthValue(columnWidthAsString, out double columnWidth) || columnWidth == 0)
            {
                columnWidth = -1;
            }

            return columnWidth;
        }

        /// <summary>
        /// Calculates column span based the column width and the widths of all other columns. Returns an integer representing
        /// the column span.
        /// </summary>
        /// <param name="columnIndex">
        /// Index of the current column.
        /// </param>
        /// <param name="columnWidth">
        /// Width of the current column.
        /// </param>
        /// <param name="columnStarts">
        /// A list representing starting coordinates of all columns.
        /// </param>
        private static int CalculateColumnSpan(int columnIndex, double columnWidth, List<double> columnStarts)
        {
            // Current status of column width. Indicates the amount of width that has been scanned already
            double columnSpanningValue;
            int columnSpanningIndex;
            int columnSpan;
            double subColumnWidth; // Width of the smallest-grain columns in the table

            Debug.Assert(columnStarts != null);
            Debug.Assert(columnIndex < columnStarts.Count - 1);
            Debug.Assert(columnStarts[columnIndex] >= 0);
            Debug.Assert(columnWidth > 0);

            columnSpanningIndex = columnIndex;
            columnSpanningValue = 0;

            while (columnSpanningValue < columnWidth && columnSpanningIndex < columnStarts.Count - 1)
            {
                subColumnWidth = columnStarts[columnSpanningIndex + 1] - columnStarts[columnSpanningIndex];
                Debug.Assert(subColumnWidth > 0);
                columnSpanningValue += subColumnWidth;
                columnSpanningIndex++;
            }

            // Now, we have either covered the width we needed to cover or reached the end of the table, in which
            // case the column spans all the columns until the end
            columnSpan = columnSpanningIndex - columnIndex;
            Debug.Assert(columnSpan > 0);

            return columnSpan;
        }

        /// <summary>
        /// Verifies that values in columnStart, which represent starting coordinates of all columns, are arranged
        /// in ascending order.
        /// </summary>
        /// <param name="columnStarts">
        /// ArrayList representing starting coordinates of all columns.
        /// </param>
        private static void VerifyColumnStartsAscendingOrder(List<double> columnStarts)
        {
            Debug.Assert(columnStarts != null);

            double columnStart = -0.01;

            for (int columnIndex = 0; columnIndex < columnStarts.Count; columnIndex++)
            {
                Debug.Assert(columnStart < columnStarts[columnIndex]);
                columnStart = columnStarts[columnIndex];
            }
        }

        /// <summary>
        /// Analyzes local properties of HTML element, converts them into XAML equivalents, and applies them to xamlElement.
        /// </summary>
        /// <param name="xamlElement">
        /// XmlElement representing XAML element to which properties are to be applied.
        /// </param>
        /// <param name="localProperties">
        /// Dictionary representing local properties of HTML element that is converted into xamlElement.
        /// </param>
        private static void ApplyLocalProperties(XmlElement xamlElement, Dictionary<string, string> localProperties, bool isBlock)
        {
            bool marginSet = false;
            string marginTop = "0";
            string marginBottom = "0";
            string marginLeft = "0";
            string marginRight = "0";

            bool paddingSet = false;
            string paddingTop = "0";
            string paddingBottom = "0";
            string paddingLeft = "0";
            string paddingRight = "0";

            string borderColor = null;

            bool borderThicknessSet = false;
            string borderThicknessTop = "0";
            string borderThicknessBottom = "0";
            string borderThicknessLeft = "0";
            string borderThicknessRight = "0";

            foreach (var item in localProperties)
            {
                switch (item.Key)
                {
                    case "font-family":
                        // Convert from font-family value list into XAML FontFamily value
                        xamlElement.SetAttribute(XamlFontFamily, item.Value);
                        break;
                    case "font-style":
                        xamlElement.SetAttribute(XamlFontStyle, item.Value);
                        break;
                    case "font-variant":
                        // Convert from font-variant into XAML property
                        break;
                    case "font-weight":
                        xamlElement.SetAttribute(XamlFontWeight, item.Value);
                        break;
                    case "font-size":
                        // Convert from CSS size into FontSize
                        xamlElement.SetAttribute(XamlFontSize, item.Value);
                        break;
                    case "color":
                        SetPropertyValue(xamlElement, TextElement.ForegroundProperty, item.Value);
                        break;
                    case "background-color":
                        SetPropertyValue(xamlElement, TextElement.BackgroundProperty, item.Value);
                        break;
                    case "text-decoration-underline":
                        if (!isBlock && item.Value == "true")
                        {
                            xamlElement.SetAttribute(XamlTextDecorations, XamlTextDecorationsUnderline);
                        }
                        break;
                    case "text-decoration-none":
                    case "text-decoration-overline":
                    case "text-decoration-line-through":
                    case "text-decoration-blink":
                        // Convert from all other text-decorations values
                        if (!isBlock)
                        {
                        }
                        break;
                    case "text-transform":
                        // Convert from text-transform into XAML property
                        break;

                    case "text-indent":
                        if (isBlock)
                        {
                            xamlElement.SetAttribute(XamlTextIndent, item.Value);
                        }
                        break;

                    case "text-align":
                        if (isBlock)
                        {
                            xamlElement.SetAttribute(XamlTextAlignment, item.Value);
                        }
                        break;

                    case "width":
                    case "height":
                        // Decide what to do with width and height properties
                        break;

                    case "margin-top":
                        marginSet = true;
                        marginTop = item.Value;
                        break;
                    case "margin-right":
                        marginSet = true;
                        marginRight = item.Value;
                        break;
                    case "margin-bottom":
                        marginSet = true;
                        marginBottom = item.Value;
                        break;
                    case "margin-left":
                        marginSet = true;
                        marginLeft = item.Value;
                        break;

                    case "padding-top":
                        paddingSet = true;
                        paddingTop = item.Value;
                        break;
                    case "padding-right":
                        paddingSet = true;
                        paddingRight = item.Value;
                        break;
                    case "padding-bottom":
                        paddingSet = true;
                        paddingBottom = item.Value;
                        break;
                    case "padding-left":
                        paddingSet = true;
                        paddingLeft = item.Value;
                        break;

                    // NOTE: CSS names for elementary border styles have side indications in the middle (top/bottom/left/right)
                    // In our internal notation we intentionally put them at the end - to unify processing in ParseCssRectangleProperty method
                    case "border-color-top":
                        borderColor = item.Value;
                        break;
                    case "border-color-right":
                        borderColor = item.Value;
                        break;
                    case "border-color-bottom":
                        borderColor = item.Value;
                        break;
                    case "border-color-left":
                        borderColor = item.Value;
                        break;
                    case "border-style-top":
                    case "border-style-right":
                    case "border-style-bottom":
                    case "border-style-left":
                        // Implement conversion from border style
                        break;
                    case "border-width-top":
                        borderThicknessSet = true;
                        borderThicknessTop = item.Value;
                        break;
                    case "border-width-right":
                        borderThicknessSet = true;
                        borderThicknessRight = item.Value;
                        break;
                    case "border-width-bottom":
                        borderThicknessSet = true;
                        borderThicknessBottom = item.Value;
                        break;
                    case "border-width-left":
                        borderThicknessSet = true;
                        borderThicknessLeft = item.Value;
                        break;

                    case "list-style-type":
                        if (xamlElement.LocalName == XamlList)
                        {
                            string markerStyle = (item.Value).ToLowerInvariant() switch
                            {
                                "disc" => XamlListMarkerStyleDisc,
                                "circle" => XamlListMarkerStyleCircle,
                                "none" => XamlListMarkerStyleNone,
                                "square" => XamlListMarkerStyleSquare,
                                "box" => XamlListMarkerStyleBox,
                                "lower-latin" => XamlListMarkerStyleLowerLatin,
                                "upper-latin" => XamlListMarkerStyleUpperLatin,
                                "lower-roman" => XamlListMarkerStyleLowerRoman,
                                "upper-roman" => XamlListMarkerStyleUpperRoman,
                                "decimal" => XamlListMarkerStyleDecimal,
                                _ => XamlListMarkerStyleDisc,
                            };
                            xamlElement.SetAttribute(XamlListMarkerStyle, markerStyle);
                        }
                        break;

                    case "float":
                    case "clear":
                        if (isBlock)
                        {
                            // Convert float and clear properties
                        }
                        break;

                    case "display":
                        break;
                }
            }

            if (isBlock)
            {
                if (marginSet)
                {
                    ComposeThicknessProperty(xamlElement, XamlMargin, marginLeft, marginRight, marginTop, marginBottom);
                }

                if (paddingSet)
                {
                    ComposeThicknessProperty(xamlElement, XamlPadding, paddingLeft, paddingRight, paddingTop, paddingBottom);
                }

                if (borderColor != null)
                {
                    // We currently ignore possible difference in brush colors on different border sides. Use the last colored side mentioned
                    xamlElement.SetAttribute(XamlBorderBrush, borderColor);
                }

                if (borderThicknessSet)
                {
                    ComposeThicknessProperty(xamlElement, XamlBorderThickness, borderThicknessLeft, borderThicknessRight, borderThicknessTop, borderThicknessBottom);
                }
            }
        }

        // Create syntactically optimized four-value Thickness
        private static void ComposeThicknessProperty(XmlElement xamlElement, string propertyName, string left, string right, string top, string bottom)
        {
            // XAML syntax:
            // We have a reasonable interpretation for one value (all four edges), two values (horizontal, vertical),
            // and four values (left, top, right, bottom).
            //  switch (i) {
            //    case 1: return new Thickness(lengths[0]);
            //    case 2: return new Thickness(lengths[0], lengths[1], lengths[0], lengths[1]);
            //    case 4: return new Thickness(lengths[0], lengths[1], lengths[2], lengths[3]);
            //  }
            string thickness;

            // We do not accept negative margins
            if (left[0] == '0' || left[0] == '-')
            {
                left = "0";
            }

            if (right[0] == '0' || right[0] == '-')
            {
                right = "0";
            }

            if (top[0] == '0' || top[0] == '-')
            {
                top = "0";
            }

            if (bottom[0] == '0' || bottom[0] == '-')
            {
                bottom = "0";
            }

            if (left == right && top == bottom)
            {
                if (left == top)
                {
                    thickness = left;
                }
                else
                {
                    thickness = left + "," + top;
                }
            }
            else
            {
                thickness = left + "," + top + "," + right + "," + bottom;
            }

            // Need safer processing for a thickness value
            xamlElement.SetAttribute(propertyName, thickness);
        }

        private static void SetPropertyValue(XmlElement xamlElement, DependencyProperty property, string stringValue)
        {
            var typeConverter = System.ComponentModel.TypeDescriptor.GetConverter(property.PropertyType);
            try
            {
                object convertedValue = typeConverter.ConvertFromInvariantString(stringValue);
                if (convertedValue != null)
                {
                    xamlElement.SetAttribute(property.Name, stringValue);
                }
            }
            catch
            {
            }
        }

        /// <summary>
        /// Analyzes the tag of the htmlElement and infers its associated formatted properties.
        /// After that parses style attribute and adds all inline CSS styles.
        /// The resulting style attributes are collected in output parameter localProperties.
        /// </summary>
        /// <param name="htmlElement"></param>
        /// <param name="inheritedProperties">
        /// set of properties inherited from ancestor elements. Currently not used in the code. Reserved for the future development.
        /// </param>
        /// <param name="localProperties">
        /// returns all formatting properties defined by this element - implied by its tag, its attributes, or its CSS inline style.
        /// </param>
        /// <param name="stylesheet"></param>
        /// <param name="sourceContext"></param>
        /// <returns>
        /// returns a combination of previous context with local set of properties.
        /// This value is not used in the current code - intended for the future development.
        /// </returns>
        private static Dictionary<string, string> GetElementProperties(XmlElement htmlElement, Dictionary<string, string> inheritedProperties, out Dictionary<string, string> localProperties, CssStylesheet stylesheet, List<XmlElement> sourceContext)
        {
            // Start with context formatting properties
            var currentProperties = new Dictionary<string, string>(inheritedProperties);

            // Identify element name
            string elementName = htmlElement.LocalName.ToLowerInvariant();

            // Update current formatting properties depending on element tag
            localProperties = new Dictionary<string, string>();
            switch (elementName)
            {
                // Character formatting
                case "i":
                case "italic":
                case "em":
                    localProperties["font-style"] = "italic";
                    break;
                case "b":
                case "bold":
                case "strong":
                case "dfn":
                    localProperties["font-weight"] = "bold";
                    break;
                case "u":
                case "underline":
                    localProperties["text-decoration-underline"] = "true";
                    break;
                case "font":
                    string attributeValue = GetAttribute(htmlElement, "face");
                    if (attributeValue != null)
                    {
                        localProperties["font-family"] = attributeValue;
                    }

                    attributeValue = GetAttribute(htmlElement, "size");
                    if (attributeValue != null)
                    {
                        double fontSize = double.Parse(attributeValue) * (12.0 / 3.0);
                        if (fontSize < 1.0)
                        {
                            fontSize = 1.0;
                        }
                        else if (fontSize > 1000.0)
                        {
                            fontSize = 1000.0;
                        }

                        localProperties["font-size"] = fontSize.ToString();
                    }

                    attributeValue = GetAttribute(htmlElement, "color");
                    if (attributeValue != null)
                    {
                        localProperties["color"] = attributeValue;
                    }
                    break;
                case "samp":
                    localProperties["font-family"] = "Courier New"; // code sample
                    localProperties["font-size"] = XamlFontSizeXXSmall;
                    localProperties["text-align"] = "Left";
                    break;
                case "sub":
                    break;
                case "sup":
                    break;

                // Hyperlinks
                case "a": // href, hreflang, urn, methods, rel, rev, title
                    //  Set default hyperlink properties
                    break;
                case "acronym":
                    break;

                // Paragraph formatting:
                case "p":
                    // Set default paragraph properties
                    break;
                case "div":
                    // Set default div properties
                    break;
                case "pre":
                    localProperties["font-family"] = "Courier New"; // renders text in a fixed-width font
                    localProperties["font-size"] = XamlFontSizeXXSmall;
                    localProperties["text-align"] = "Left";
                    break;
                case "blockquote":
                    localProperties["margin-left"] = "16";
                    break;

                case "h1":
                    localProperties["font-size"] = XamlFontSizeXXLarge;
                    break;
                case "h2":
                    localProperties["font-size"] = XamlFontSizeXLarge;
                    break;
                case "h3":
                    localProperties["font-size"] = XamlFontSizeLarge;
                    break;
                case "h4":
                    localProperties["font-size"] = XamlFontSizeMedium;
                    break;
                case "h5":
                    localProperties["font-size"] = XamlFontSizeSmall;
                    break;
                case "h6":
                    localProperties["font-size"] = XamlFontSizeXSmall;
                    break;

                // List properties
                case "ul":
                    localProperties["list-style-type"] = "disc";
                    break;
                case "ol":
                    localProperties["list-style-type"] = "decimal";
                    break;

                case "table":
                case "body":
                case "html":
                    break;
            }

            // Override HTML defaults by CSS attributes - from style sheets and inline settings
            HtmlCssParser.GetElementPropertiesFromCssAttributes(htmlElement, elementName, stylesheet, localProperties, sourceContext);

            // Combine local properties with context to create new current properties
            foreach (var item in localProperties)
            {
                currentProperties.Add(item.Key, item.Value);
            }

            return currentProperties;
        }

        /// <summary>
        /// Extracts a value of CSS attribute from CSS style definition.
        /// </summary>
        /// <param name="cssStyle">
        /// Source CSS style definition.
        /// </param>
        /// <param name="attributeName">
        /// A name of CSS attribute to extract.
        /// </param>
        /// <returns>
        /// A string representation of an attribute value if found;
        /// null if there is no such attribute in a given string.
        /// </returns>
        private static string GetCssAttribute(string cssStyle, string attributeName)
        {
            // This is poor man's attribute parsing. Replace it by real CSS parsing
            if (cssStyle != null)
            {
                string[] styleValues;

                attributeName = attributeName.ToLowerInvariant();

                // Check for width specification in style string
                styleValues = cssStyle.Split(';');

                foreach (string styleValue in styleValues)
                {
                    string[] styleNameValue = styleValue.Split(':');
                    if (styleNameValue.Length == 2 && styleNameValue[0].Trim().ToLowerInvariant() == attributeName)
                    {
                        return styleNameValue[1].Trim();
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Converts a length value from string representation to a double.
        /// </summary>
        /// <param name="lengthAsString">
        /// Source string value of a length.
        /// </param>
        /// <param name="length"></param>
        private static bool TryGetLengthValue(string lengthAsString, out double length)
        {
            if (lengthAsString != null)
            {
                lengthAsString = lengthAsString.Trim().ToLowerInvariant();

                // We try to convert currentColumnWidthAsString into a double. This will eliminate widths of type "50%", etc.
                if (lengthAsString.EndsWith("pt"))
                {
                    lengthAsString = lengthAsString[0..^2];
                    if (double.TryParse(lengthAsString, NumberStyles.Float, CultureInfo.InvariantCulture, out length))
                    {
                        length = length * 96.0 / 72.0; // convert from points to pixels
                    }
                    else
                    {
                        length = double.NaN;
                    }
                }
                else if (lengthAsString.EndsWith("px"))
                {
                    lengthAsString = lengthAsString[0..^2];
                    if (!double.TryParse(lengthAsString, NumberStyles.Float, CultureInfo.InvariantCulture, out length))
                    {
                        length = double.NaN;
                    }
                }
                else
                {
                    if (!double.TryParse(lengthAsString, NumberStyles.Float, CultureInfo.InvariantCulture, out length)) // Assuming pixels
                    {
                        length = double.NaN;
                    }
                }
            }
            else
            {
                length = double.NaN;
            }

            return !double.IsNaN(length);
        }

        private static string GetColorValue(string colorValue)
        {
            // TODO: Implement color conversion
            return colorValue;
        }

        /// <summary>
        /// Applies properties to xamlTableCellElement based on the HTML td element it is converted from.
        /// </summary>
        /// <param name="htmlChildNode">
        /// HTML td/th element to be converted to XAML.
        /// </param>
        /// <param name="xamlTableCellElement">
        /// XmlElement representing XAML element for which properties are to be processed.
        /// </param>
        /// <remarks>
        /// TODO: Use the processed properties for htmlChildNode instead of using the node itself.
        /// </remarks>
        private static void ApplyPropertiesToTableCellElement(XmlElement htmlChildNode, XmlElement xamlTableCellElement)
        {
            // Parameter validation
            Debug.Assert(string.Equals(htmlChildNode.LocalName, "td", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(htmlChildNode.LocalName, "th", StringComparison.OrdinalIgnoreCase));
            Debug.Assert(xamlTableCellElement.LocalName == XamlTableCell);

            // Set default border thickness for xamlTableCellElement to enable grid lines
            xamlTableCellElement.SetAttribute(XamlTableCellBorderThickness, "1,1,1,1");
            xamlTableCellElement.SetAttribute(XamlTableCellBorderBrush, XamlBrushesBlack);
            string rowSpanString = GetAttribute(htmlChildNode, "rowspan");
            if (rowSpanString != null)
            {
                xamlTableCellElement.SetAttribute(XamlTableCellRowSpan, rowSpanString);
            }
        }
    }
}
