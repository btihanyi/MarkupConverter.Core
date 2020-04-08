//---------------------------------------------------------------------------
//
// File: HtmlTokenType.cs
//
// Copyright (C) Microsoft Corporation.  All rights reserved.
//
// Description: Definition of token types supported by HtmlLexicalAnalyzer
//
//---------------------------------------------------------------------------

namespace MarkupConverter.Core
{
    /// <summary>
    /// Types of lexical tokens for HTML-to-XAML converter.
    /// </summary>
    internal enum HtmlTokenType
    {
        OpeningTagStart,
        ClosingTagStart,
        TagEnd,
        EmptyTagEnd,
        EqualSign,
        Name,
        Atom, // Any attribute value not in quotes
        Text, // Text content when accepting text
        Comment,
        EOF,
    }
}
