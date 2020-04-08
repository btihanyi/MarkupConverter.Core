//---------------------------------------------------------------------------
//
// File: HtmlLexicalAnalyzer.cs
//
// Copyright (C) Microsoft Corporation.  All rights reserved.
//
// Description: Lexical analyzer for HTML-to-Xaml converter
//
//---------------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace MarkupConverter.Core
{
    /// <summary>
    /// lexical analyzer class
    /// recognizes tokens as groups of characters separated by arbitrary amounts of whitespace
    /// also classifies tokens according to type.
    /// </summary>
    internal class HtmlLexicalAnalyzer
    {
        private readonly StringReader inputStringReader; // String reader which will move over input text
        private readonly StringBuilder nextToken; // Store token and type in local variables before copying them to output parameters

        private int nextCharacterCode; // Next character code read from input that is not yet part of any token and the character it represents
        private char nextCharacter;
        private int lookAheadCharacterCode;
        private char lookAheadCharacter;
        private char previousCharacter;
        private bool ignoreNextWhitespace;
        private bool isNextCharacterEntity; // Check if next character is an entity

        /// <summary>
        /// Initializes a new instance of the <see cref="HtmlLexicalAnalyzer"/> class by
        /// initializing the <see cref="inputStringReader"/> member with the string to be read
        /// also sets initial values for <see cref="nextCharacterCode"/> and <see cref="NextTokenType"/>.
        /// </summary>
        /// <param name="inputTextString">
        /// Text string to be parsed for XML content.
        /// </param>
        internal HtmlLexicalAnalyzer(string inputTextString)
        {
            inputStringReader = new StringReader(inputTextString);
            nextCharacterCode = 0;
            nextCharacter = ' ';
            lookAheadCharacterCode = inputStringReader.Read();
            lookAheadCharacter = (char) lookAheadCharacterCode;
            previousCharacter = ' ';
            ignoreNextWhitespace = true;
            nextToken = new StringBuilder(100);
            NextTokenType = HtmlTokenType.Text;

            // Read the first character so we have some value for the nextCharacter field
            GetNextCharacter();
        }

        internal HtmlTokenType NextTokenType { get; private set; }

        internal string NextToken => nextToken.ToString();

        private bool IsAtEndOfStream => nextCharacterCode == -1;

        private bool IsAtTagStart
        {
            get
            {
                return nextCharacter == '<' && (lookAheadCharacter == '/' || IsGoodForNameStart(lookAheadCharacter)) && !isNextCharacterEntity;
            }
        }

        private bool IsAtDirectiveStart
        {
            get
            {
                return (nextCharacter == '<' && lookAheadCharacter == '!' && !isNextCharacterEntity);
            }
        }

        /// <summary>
        /// retrieves next recognizable token from input string
        /// and identifies its type
        /// if no valid token is found, the output parameters are set to null
        /// if end of stream is reached without matching any token, token type
        /// parameter is set to EOF.
        /// </summary>
        internal void GetNextContentToken()
        {
            Debug.Assert(NextTokenType != HtmlTokenType.EOF);
            nextToken.Length = 0;
            if (IsAtEndOfStream)
            {
                NextTokenType = HtmlTokenType.EOF;
                return;
            }

            if (IsAtTagStart)
            {
                GetNextCharacter();

                if (nextCharacter == '/')
                {
                    nextToken.Append("</");
                    NextTokenType = HtmlTokenType.ClosingTagStart;

                    // advance
                    GetNextCharacter();
                    ignoreNextWhitespace = false; // Whitespaces after closing tags are significant
                }
                else
                {
                    NextTokenType = HtmlTokenType.OpeningTagStart;
                    nextToken.Append("<");
                    ignoreNextWhitespace = true; // Whitespaces after opening tags are insignificant
                }
            }
            else if (IsAtDirectiveStart)
            {
                // either a comment or CDATA
                GetNextCharacter();
                if (lookAheadCharacter == '[')
                {
                    // cdata
                    ReadDynamicContent();
                }
                else if (lookAheadCharacter == '-')
                {
                    ReadComment();
                }
                else
                {
                    // neither a comment nor cdata, should be something like DOCTYPE
                    // skip till the next tag ender
                    ReadUnknownDirective();
                }
            }
            else
            {
                // read text content, unless you encounter a tag
                NextTokenType = HtmlTokenType.Text;
                while (!IsAtTagStart && !IsAtEndOfStream && !IsAtDirectiveStart)
                {
                    if (nextCharacter == '<' && !isNextCharacterEntity && lookAheadCharacter == '?')
                    {
                        // ignore processing directive
                        SkipProcessingDirective();
                    }
                    else
                    {
                        if (nextCharacter <= ' ')
                        {
                            // Respect xml:preserve or its equivalents for whitespace processing
                            if (ignoreNextWhitespace)
                            {
                                // Ignore repeated whitespaces
                            }
                            else
                            {
                                // Treat any control character sequence as one whitespace
                                nextToken.Append(' ');
                            }

                            ignoreNextWhitespace = true; // and keep ignoring the following whitespaces
                        }
                        else
                        {
                            nextToken.Append(nextCharacter);
                            ignoreNextWhitespace = false;
                        }

                        GetNextCharacter();
                    }
                }
            }
        }

        /// <summary>
        /// Unconditionally returns a token which is one of: TagEnd, EmptyTagEnd, Name, Atom or EndOfStream
        /// Does not guarantee token reader advancing.
        /// </summary>
        internal void GetNextTagToken()
        {
            nextToken.Length = 0;
            if (IsAtEndOfStream)
            {
                NextTokenType = HtmlTokenType.EOF;
                return;
            }

            SkipWhiteSpace();

            if (nextCharacter == '>' && !isNextCharacterEntity)
            {
                // &gt; should not end a tag, so make sure it's not an entity
                NextTokenType = HtmlTokenType.TagEnd;
                nextToken.Append('>');
                GetNextCharacter();

                // Note: ignoreNextWhitespace must be set appropriately on tag start processing
            }
            else if (nextCharacter == '/' && lookAheadCharacter == '>')
            {
                // could be start of closing of empty tag
                NextTokenType = HtmlTokenType.EmptyTagEnd;
                nextToken.Append("/>");
                GetNextCharacter();
                GetNextCharacter();
                ignoreNextWhitespace = false; // Whitespace after no-scope tags are significant
            }
            else if (IsGoodForNameStart(nextCharacter))
            {
                NextTokenType = HtmlTokenType.Name;

                // starts a name
                // we allow character entities here
                // we do not throw exceptions here if end of stream is encountered
                // just stop and return whatever is in the token
                // if the parser is not expecting end of file after this it will call
                // the get next token function and throw an exception
                while (IsGoodForName(nextCharacter) && !IsAtEndOfStream)
                {
                    nextToken.Append(nextCharacter);
                    GetNextCharacter();
                }
            }
            else
            {
                // Unexpected type of token for a tag. Reprot one character as Atom, expecting that HtmlParser will ignore it.
                NextTokenType = HtmlTokenType.Atom;
                nextToken.Append(nextCharacter);
                GetNextCharacter();
            }
        }

        /// <summary>
        /// Unconditionally returns equal sign token. Even if there is no
        /// real equal sign in the stream, it behaves as if it were there.
        /// Does not guarantee token reader advancing.
        /// </summary>
        internal void GetNextEqualSignToken()
        {
            Debug.Assert(NextTokenType != HtmlTokenType.EOF);
            nextToken.Length = 0;

            nextToken.Append('=');
            NextTokenType = HtmlTokenType.EqualSign;

            SkipWhiteSpace();

            if (nextCharacter == '=')
            {
                // '=' is not in the list of entities, so no need to check for entities here
                GetNextCharacter();
            }
        }

        /// <summary>
        /// Unconditionally returns an atomic value for an attribute
        /// Even if there is no appropriate token it returns Atom value
        /// Does not guarantee token reader advancing.
        /// </summary>
        internal void GetNextAtomToken()
        {
            Debug.Assert(NextTokenType != HtmlTokenType.EOF);
            nextToken.Length = 0;

            SkipWhiteSpace();

            NextTokenType = HtmlTokenType.Atom;

            if ((nextCharacter == '\'' || nextCharacter == '"') && !isNextCharacterEntity)
            {
                char startingQuote = nextCharacter;
                GetNextCharacter();

                // Consume all characters between quotes
                while (!(nextCharacter == startingQuote && !isNextCharacterEntity) && !IsAtEndOfStream)
                {
                    nextToken.Append(nextCharacter);
                    GetNextCharacter();
                }

                if (nextCharacter == startingQuote)
                {
                    GetNextCharacter();
                }

                // complete the quoted value
                // NOTE: our recovery here is different from IE's
                // IE keeps reading until it finds a closing quote or end of file
                // if end of file, it treats current value as text
                // if it finds a closing quote at any point within the text, it eats everything between the quotes
                // TODO: Suggestion:
                // however, we could stop when we encounter end of file or an angle bracket of any kind
                // and assume there was a quote there
                // so the attribute value may be meaningless but it is never treated as text
            }
            else
            {
                while (!IsAtEndOfStream && !char.IsWhiteSpace(nextCharacter) && nextCharacter != '>')
                {
                    nextToken.Append(nextCharacter);
                    GetNextCharacter();
                }
            }
        }

        /// <summary>
        /// Advances a reading position by one character code
        /// and reads the next available character from a stream.
        /// This character becomes available as nextCharacter property.
        /// </summary>
        /// <remarks>
        /// Throws InvalidOperationException if attempted to be called on EndOfStream
        /// condition.
        /// </remarks>
        private void GetNextCharacter()
        {
            if (nextCharacterCode == -1)
            {
                throw new InvalidOperationException("GetNextCharacter method called at the end of a stream");
            }

            previousCharacter = nextCharacter;

            nextCharacter = lookAheadCharacter;
            nextCharacterCode = lookAheadCharacterCode;
            isNextCharacterEntity = false; // Next character not an entity as of now

            ReadLookAheadCharacter();

            if (nextCharacter == '&')
            {
                if (lookAheadCharacter == '#')
                {
                    // numeric entity - parse digits - &#DDDDD;
                    int entityCode = 0;
                    ReadLookAheadCharacter();

                    // largest numeric entity is 7 characters
                    for (int i = 0; i < 7 && char.IsDigit(lookAheadCharacter); i++)
                    {
                        entityCode = (10 * entityCode) + (lookAheadCharacterCode - '0');
                        ReadLookAheadCharacter();
                    }

                    if (lookAheadCharacter == ';')
                    {
                        // correct format - advance
                        ReadLookAheadCharacter();
                        nextCharacterCode = entityCode;

                        // if this is out of range it will set the character to '?'
                        nextCharacter = (char) nextCharacterCode;

                        // as far as we are concerned, this is an entity
                        isNextCharacterEntity = true;
                    }
                    else
                    {
                        // not an entity, set next character to the current lookahead character
                        // we would have eaten up some digits
                        nextCharacter = lookAheadCharacter;
                        nextCharacterCode = lookAheadCharacterCode;
                        ReadLookAheadCharacter();
                        isNextCharacterEntity = false;
                    }
                }
                else if (char.IsLetter(lookAheadCharacter))
                {
                    // entity is written as a string
                    string entity = string.Empty;

                    // maximum length of string entities is 10 characters
                    for (int i = 0; i < 10 && (char.IsLetter(lookAheadCharacter) || char.IsDigit(lookAheadCharacter)); i++)
                    {
                        entity += lookAheadCharacter;
                        ReadLookAheadCharacter();
                    }

                    if (lookAheadCharacter == ';')
                    {
                        // advance
                        ReadLookAheadCharacter();

                        if (HtmlSchema.IsEntity(entity))
                        {
                            nextCharacter = HtmlSchema.EntityCharacterValue(entity);
                            nextCharacterCode = nextCharacter;
                            isNextCharacterEntity = true;
                        }
                        else
                        {
                            // just skip the whole thing - invalid entity
                            // move on to the next character
                            nextCharacter = lookAheadCharacter;
                            nextCharacterCode = lookAheadCharacterCode;
                            ReadLookAheadCharacter();

                            // not an entity
                            isNextCharacterEntity = false;
                        }
                    }
                    else
                    {
                        // skip whatever we read after the ampersand
                        // set next character and move on
                        nextCharacter = lookAheadCharacter;
                        ReadLookAheadCharacter();
                        isNextCharacterEntity = false;
                    }
                }
            }
        }

        private void ReadLookAheadCharacter()
        {
            if (lookAheadCharacterCode != -1)
            {
                lookAheadCharacterCode = inputStringReader.Read();
                lookAheadCharacter = (char) lookAheadCharacterCode;
            }
        }

        /// <summary>
        /// skips whitespace in the input string
        /// leaves the first non-whitespace character available in the nextCharacter property
        /// this may be the end-of-file character, it performs no checking.
        /// </summary>
        private void SkipWhiteSpace()
        {
            // TODO: handle character entities while processing comments, CDATA, and directives
            // TODO: SUGGESTION: we could check if lookahead and previous characters are entities also
            while (true)
            {
                if (nextCharacter == '<' && (lookAheadCharacter == '?' || lookAheadCharacter == '!'))
                {
                    GetNextCharacter();

                    if (lookAheadCharacter == '[')
                    {
                        // Skip CDATA block and DTDs(?)
                        while (!IsAtEndOfStream && !(previousCharacter == ']' && nextCharacter == ']' && lookAheadCharacter == '>'))
                        {
                            GetNextCharacter();
                        }

                        if (nextCharacter == '>')
                        {
                            GetNextCharacter();
                        }
                    }
                    else
                    {
                        // Skip processing instruction, comments
                        while (!IsAtEndOfStream && nextCharacter != '>')
                        {
                            GetNextCharacter();
                        }

                        if (nextCharacter == '>')
                        {
                            GetNextCharacter();
                        }
                    }
                }

                if (!char.IsWhiteSpace(nextCharacter))
                {
                    break;
                }

                GetNextCharacter();
            }
        }

        /// <summary>
        /// checks if a character can be used to start a name
        /// if this check is true then the rest of the name can be read.
        /// </summary>
        /// <param name="character">
        /// character value to be checked.
        /// </param>
        /// <returns>
        /// true if the character can be the first character in a name
        /// false otherwise.
        /// </returns>
        private bool IsGoodForNameStart(char character)
        {
            return character == '_' || char.IsLetter(character);
        }

        /// <summary>
        /// checks if a character can be used as a non-starting character in a name
        /// uses the IsExtender and IsCombiningCharacter predicates to see
        /// if a character is an extender or a combining character.
        /// </summary>
        /// <param name="character">
        /// character to be checked for validity in a name.
        /// </param>
        /// <returns>
        /// true if the character can be a valid part of a name.
        /// </returns>
        private bool IsGoodForName(char character)
        {
            // we are not concerned with escaped characters in names
            // we assume that character entities are allowed as part of a name
            return
                IsGoodForNameStart(character) ||
                character == '.' ||
                character == '-' ||
                character == ':' ||
                char.IsDigit(character) ||
                IsCombiningCharacter(character) ||
                IsExtender(character);
        }

        /// <summary>
        /// identifies a character as being a combining character, permitted in a name
        /// TODO: only a placeholder for now but later to be replaced with comparisons against
        /// the list of combining characters in the XML documentation.
        /// </summary>
        /// <param name="character">
        /// character to be checked.
        /// </param>
        /// <returns>
        /// true if the character is a combining character, false otherwise.
        /// </returns>
        private bool IsCombiningCharacter(char character)
        {
            // TODO: put actual code with checks against all combining characters here
            return false;
        }

        /// <summary>
        /// identifies a character as being an extender, permitted in a name
        /// TODO: only a placeholder for now but later to be replaced with comparisons against
        /// the list of extenders in the XML documentation.
        /// </summary>
        /// <param name="character">
        /// character to be checked.
        /// </param>
        /// <returns>
        /// true if the character is an extender, false otherwise.
        /// </returns>
        private bool IsExtender(char character)
        {
            // TODO: put actual code with checks against all extenders here
            return false;
        }

        /// <summary>
        /// skips dynamic content starting with '.<![' and ending with ']>'
        /// </summary>
        private void ReadDynamicContent()
        {
            // verify that we are at dynamic content, which may include CDATA
            Debug.Assert(previousCharacter == '<' && nextCharacter == '!' && lookAheadCharacter == '[');

            // Let's treat this as empty text
            NextTokenType = HtmlTokenType.Text;
            nextToken.Length = 0;

            // advance twice, once to get the lookahead character and then to reach the start of the cdata
            GetNextCharacter();
            GetNextCharacter();

            // NOTE: 10/12/2004: modified this function to check when called if's reading CDATA or something else
            // some directives may start with a <![ and then have some data and they will just end with a ]>
            // this function is modified to stop at the sequence ]> and not ]]>
            // this means that CDATA and anything else expressed in their own set of [] within the <! [...]>
            // directive cannot contain a ]> sequence. However it is doubtful that cdata could contain such
            // sequence anyway, it probably stops at the first ]
            while (!(nextCharacter == ']' && lookAheadCharacter == '>') && !IsAtEndOfStream)
            {
                // advance
                GetNextCharacter();
            }

            if (!IsAtEndOfStream)
            {
                // advance, first to the last >
                GetNextCharacter();

                // then advance past it to the next character after processing directive
                GetNextCharacter();
            }
        }

        /// <summary>
        /// skips comments starting with '.<!-' and ending with '-->'
        /// NOTE: 10/06/2004: processing changed, will now skip anything starting with
        /// the "<!-"  sequence and ending in "!>" or "->", because in practice many HTML pages do not
        /// use the full comment specifying conventions
        /// </summary>
        private void ReadComment()
        {
            // verify that we are at a comment
            Debug.Assert(previousCharacter == '<' && nextCharacter == '!' && lookAheadCharacter == '-');

            // Initialize a token
            NextTokenType = HtmlTokenType.Comment;
            nextToken.Length = 0;

            // advance to the next character, so that to be at the start of comment value
            GetNextCharacter(); // get first '-'
            GetNextCharacter(); // get second '-'
            GetNextCharacter(); // get first character of comment content

            while (true)
            {
                // Read text until end of comment
                // Note that in many actual HTML pages comments end with "!>" (while XML standard is "-->")
                while (!IsAtEndOfStream && !((nextCharacter == '-' && lookAheadCharacter == '-') || (nextCharacter == '!' && lookAheadCharacter == '>')))
                {
                    nextToken.Append(nextCharacter);
                    GetNextCharacter();
                }

                // Finish comment reading
                GetNextCharacter();
                if (previousCharacter == '-' && nextCharacter == '-' && lookAheadCharacter == '>')
                {
                    // Standard comment end. Eat it and exit the loop
                    GetNextCharacter(); // get '>'
                    break;
                }
                else if (previousCharacter == '!' && nextCharacter == '>')
                {
                    // Nonstandard but possible comment end - '!>'. Exit the loop
                    break;
                }
                else
                {
                    // Not an end. Save character and continue reading
                    nextToken.Append(previousCharacter);
                }
            }

            // Read end of comment combination
            if (nextCharacter == '>')
            {
                GetNextCharacter();
            }
        }

        /// <summary>
        /// skips past unknown directives that start with ".<!" but are not comments or Cdata
        /// ignores content of such directives until the next ">" character
        /// applies to directives such as DOCTYPE, etc that we do not presently support
        /// </summary>
        private void ReadUnknownDirective()
        {
            // verify that we are at an unknown directive
            Debug.Assert(previousCharacter == '<' && nextCharacter == '!' && !(lookAheadCharacter == '-' || lookAheadCharacter == '['));

            // Let's treat this as empty text
            NextTokenType = HtmlTokenType.Text;
            nextToken.Length = 0;

            // advance to the next character
            GetNextCharacter();

            // skip to the first tag end we find
            while (!(nextCharacter == '>' && !isNextCharacterEntity) && !IsAtEndOfStream)
            {
                GetNextCharacter();
            }

            if (!IsAtEndOfStream)
            {
                // advance past the tag end
                GetNextCharacter();
            }
        }

        /// <summary>
        /// skips processing directives starting with the characters '<?' and ending with '?>'
        /// NOTE: 10/14/2004: IE also ends processing directives with a />, so this function is
        /// being modified to recognize that condition as well.
        /// </summary>
        private void SkipProcessingDirective()
        {
            // verify that we are at a processing directive
            Debug.Assert(nextCharacter == '<' && lookAheadCharacter == '?');

            // advance twice, once to get the lookahead character and then to reach the start of the directive
            GetNextCharacter();
            GetNextCharacter();

            while (!((nextCharacter == '?' || nextCharacter == '/') && lookAheadCharacter == '>') && !IsAtEndOfStream)
            {
                // advance
                // we don't need to check for entities here because '?' is not an entity
                // and even though > is an entity there is no entity processing when reading lookahead character
                GetNextCharacter();
            }

            if (!IsAtEndOfStream)
            {
                // advance, first to the last >
                GetNextCharacter();

                // then advance past it to the next character after processing directive
                GetNextCharacter();
            }
        }
    }
}
