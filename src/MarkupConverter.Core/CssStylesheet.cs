//---------------------------------------------------------------------------
//
// File: CssStylesheet.cs
//
// Copyright (C) Microsoft Corporation.  All rights reserved.
//
//---------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Xml;

namespace MarkupConverter.Core
{
    internal class CssStylesheet
    {
        private List<StyleDefinition> styleDefinitions;

        public CssStylesheet(XmlElement htmlElement)
        {
            if (htmlElement != null)
            {
                this.DiscoverStyleDefinitions(htmlElement);
            }
        }

        // Recursively traverses an HTML tree, discovers STYLE elements and creates a style definition table
        // for further cascading style application
        public void DiscoverStyleDefinitions(XmlElement htmlElement)
        {
            if (string.Equals(htmlElement.LocalName, "link", StringComparison.OrdinalIgnoreCase))
            {
                return;

                // Add link elements processing for included style sheets
                // <link href="http://sc.msn.com/global/CSS/ptnr/orange.css" type=text/css \r\nrel=stylesheet>
            }

            if (!string.Equals(htmlElement.LocalName, "style", StringComparison.OrdinalIgnoreCase))
            {
                // This is not a STYLE element. Recurse into it
                for (var htmlChildNode = htmlElement.FirstChild; htmlChildNode != null; htmlChildNode = htmlChildNode.NextSibling)
                {
                    if (htmlChildNode is XmlElement htmlChildElement)
                    {
                        this.DiscoverStyleDefinitions(htmlChildElement);
                    }
                }

                return;
            }

            // Add style definitions from this style.

            // Collect all text from this style definition
            var stylesheetBuffer = new StringBuilder();

            for (var htmlChildNode = htmlElement.FirstChild; htmlChildNode != null; htmlChildNode = htmlChildNode.NextSibling)
            {
                if (htmlChildNode is XmlText || htmlChildNode is XmlComment)
                {
                    stylesheetBuffer.Append(RemoveComments(htmlChildNode.Value));
                }
            }

            // CssStylesheet has the following syntactical structure:
            //     @import declaration;
            //     selector { definition }
            // where "selector" is one of: ".classname", "tag name"
            // It can contain comments in the following form: /*...*/
            int nextCharacterIndex = 0;
            while (nextCharacterIndex < stylesheetBuffer.Length)
            {
                // Extract selector
                int selectorStart = nextCharacterIndex;
                while (nextCharacterIndex < stylesheetBuffer.Length && stylesheetBuffer[nextCharacterIndex] != '{')
                {
                    // Skip declaration directive starting from @
                    if (stylesheetBuffer[nextCharacterIndex] == '@')
                    {
                        while (nextCharacterIndex < stylesheetBuffer.Length && stylesheetBuffer[nextCharacterIndex] != ';')
                        {
                            nextCharacterIndex++;
                        }

                        selectorStart = nextCharacterIndex + 1;
                    }

                    nextCharacterIndex++;
                }

                if (nextCharacterIndex < stylesheetBuffer.Length)
                {
                    // Extract definition
                    int definitionStart = nextCharacterIndex;
                    while (nextCharacterIndex < stylesheetBuffer.Length && stylesheetBuffer[nextCharacterIndex] != '}')
                    {
                        nextCharacterIndex++;
                    }

                    // Define a style
                    if (nextCharacterIndex - definitionStart > 2)
                    {
                        this.AddStyleDefinition(
                            stylesheetBuffer.ToString(selectorStart, definitionStart - selectorStart),
                            stylesheetBuffer.ToString(definitionStart + 1, nextCharacterIndex - definitionStart - 2));
                    }

                    // Skip closing brace
                    if (nextCharacterIndex < stylesheetBuffer.Length)
                    {
                        Debug.Assert(stylesheetBuffer[nextCharacterIndex] == '}');
                        nextCharacterIndex++;
                    }
                }
            }
        }

        // Returns a string with all c-style comments replaced by spaces
        private string RemoveComments(string text)
        {
            int commentStart = text.IndexOf("/*");
            if (commentStart < 0)
            {
                return text;
            }

            int commentEnd = text.IndexOf("*/", commentStart + 2);
            if (commentEnd < 0)
            {
                return text.Substring(0, commentStart);
            }

            return text.Substring(0, commentStart) + " " + RemoveComments(text.Substring(commentEnd + 2));
        }

        public void AddStyleDefinition(string selector, string definition)
        {
            // Normalize parameter values
            selector = selector.Trim().ToLowerInvariant();
            definition = definition.Trim().ToLowerInvariant();
            if (selector.Length == 0 || definition.Length == 0)
            {
                return;
            }

            if (styleDefinitions == null)
            {
                styleDefinitions = new List<StyleDefinition>();
            }

            string[] simpleSelectors = selector.Split(',');

            for (int i = 0; i < simpleSelectors.Length; i++)
            {
                string simpleSelector = simpleSelectors[i].Trim();
                if (simpleSelector.Length > 0)
                {
                    styleDefinitions.Add(new StyleDefinition(simpleSelector, definition));
                }
            }
        }

        public string GetStyle(string elementName, List<XmlElement> sourceContext)
        {
            Debug.Assert(sourceContext.Count > 0);
            Debug.Assert(elementName == sourceContext[sourceContext.Count - 1].LocalName);

            // Add id processing for style selectors
            if (styleDefinitions != null)
            {
                for (int i = styleDefinitions.Count - 1; i >= 0; i--)
                {
                    string selector = styleDefinitions[i].Selector;

                    string[] selectorLevels = selector.Split(' ');

                    int indexInSelector = selectorLevels.Length - 1;
                    string selectorLevel = selectorLevels[indexInSelector].Trim();

                    if (MatchSelectorLevel(selectorLevel, sourceContext[sourceContext.Count - 1]))
                    {
                        return styleDefinitions[i].Definition;
                    }
                }
            }

            return null;
        }

        private bool MatchSelectorLevel(string selectorLevel, XmlElement xmlElement)
        {
            if (selectorLevel.Length == 0)
            {
                return false;
            }

            int indexOfDot = selectorLevel.IndexOf('.');
            int indexOfPound = selectorLevel.IndexOf('#');

            string selectorClass = null;
            string selectorId = null;
            string selectorTag = null;
            if (indexOfDot >= 0)
            {
                if (indexOfDot > 0)
                {
                    selectorTag = selectorLevel.Substring(0, indexOfDot);
                }

                selectorClass = selectorLevel.Substring(indexOfDot + 1);
            }
            else if (indexOfPound >= 0)
            {
                if (indexOfPound > 0)
                {
                    selectorTag = selectorLevel.Substring(0, indexOfPound);
                }

                selectorId = selectorLevel.Substring(indexOfPound + 1);
            }
            else
            {
                selectorTag = selectorLevel;
            }

            if (selectorTag != null && selectorTag != xmlElement.LocalName)
            {
                return false;
            }

            if (selectorId != null && HtmlToXamlConverter.GetAttribute(xmlElement, "id") != selectorId)
            {
                return false;
            }

            if (selectorClass != null && HtmlToXamlConverter.GetAttribute(xmlElement, "class") != selectorClass)
            {
                return false;
            }

            return true;
        }

        private class StyleDefinition
        {
            public StyleDefinition(string selector, string definition)
            {
                this.Selector = selector;
                this.Definition = definition;
            }

            public string Selector { get; }

            public string Definition { get; }
        }
    }
}
