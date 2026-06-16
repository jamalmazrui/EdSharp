/*
 * SmartyPants  -  Smart punctuation for web sites
 * by John Gruber
 * http://daringfireball.net
 * 
 * Copyright (c) 2004 Michel Fortin - Translation to PHP
 * http://www.michelf.com
 * 
 * Copyright (c) 2004-2005 Milan Negovan - C# translation to .NET
 * http://www.aspnetresources.com
 * 
 */

#region Copyright and license
/*
Copyright (c) 2003 John Gruber
(http://daringfireball.net/)
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

*   Redistributions of source code must retain the above copyright
    notice, this list of conditions and the following disclaimer.

*   Redistributions in binary form must reproduce the above copyright
    notice, this list of conditions and the following disclaimer in the
    documentation and/or other materials provided with the distribution.

*   Neither the name "SmartyPants" nor the names of its contributors may
    be used to endorse or promote products derived from this software
    without specific prior written permission.

This software is provided by the copyright holders and contributors "as is"
and any express or implied warranties, including, but not limited to, the 
implied warranties of merchantability and fitness for a particular purpose 
are disclaimed. In no event shall the copyright owner or contributors be 
liable for any direct, indirect, incidental, special, exemplary, or 
consequential damages (including, but not limited to, procurement of 
substitute goods or services; loss of use, data, or profits; or business 
interruption) however caused and on any theory of liability, whether in 
contract, strict liability, or tort (including negligence or otherwise) 
arising in any way out of the use of this software, even if advised of the
possibility of such damage.
*/
#endregion

using System.Collections;
using System.Text.RegularExpressions;
using System.Web.UI;

namespace anrControls
{
    public enum ConversionMode { 
        /// <summary>
        /// Suppress all transformations. (Do nothing.)
        /// </summary>
        LeaveIntact, 

        /// <summary>
        /// Performs default SmartyPants transformations: quotes (including
        /// backticks-style), em-dashes, and ellipses. `--` (dash dash) is
        /// used to signify an em-dash; there is no support for en-dashes.
        /// </summary>
        EducateDefault, 

        /// <summary>
        /// Same as <seealso cref="EducateDefault"/>, except that it uses the old-school
        /// typewriter shorthand for dashes: `--` (dash dash) for en-dashes,
        /// `---` (dash dash dash) for em-dashes.
        /// </summary>
        EducateOldSchool,

        /// <summary>
        /// Same as smarty_pants="2", but inverts the shorthand for dashes: `--`
        /// (dash dash) for em-dashes, and `---` (dash dash dash) for en-dashes.
        /// </summary>
        EducateOldSchoolInverted,

        /// <summary>
        /// Stupefy mode. Reverses the SmartyPants transformation process,
        /// turning the HTML entities produced by SmartyPants into their ASCII
        /// equivalents. E.g. `&#8220;` is turned into a simple double-quote
        /// (`"`), `&#8212;` is turned into two dashes, etc.
        /// </summary>
        Stupefy
    };

    public class SmartyPants
	{
        // We don't do conversion within these tags
        public const string tagsToSkip = @"<(/?)(?:pre|code|kbd|script|math)[\s>]";

        // Determine at runtime what how to convert dashes
        private delegate string ConvertDashesDelegate (string text);

        // -------------------------------------------------------------------
        public string Transform (string text, ConversionMode mode)
        {
            if (mode == ConversionMode.LeaveIntact)
                return text;


            ArrayList               tokens = null;
            bool                    inPre = false; // Keep track of when we're inside <pre> or <code> tags
            string                  res = null;

            // This is a cheat, used to get some context
            // for one-character tokens that consist of 
            // just a quote char. What we do is remember
            // the last character of the previous text
            // token, to use as context to curl single-
            // character quote tokens correctly.
            string                  prevTokenLastChar = string.Empty;


            // Parse HTML and convert dashes
            tokens = TokenizeHTML (text);

            foreach (Pair token in tokens)
            {
                string value = token.Second.ToString ();

                if (token.First.Equals ("tag")) 
                {
                    // Don't mess with quotes inside tags
                    res += value;
                    Match m = Regex.Match (value, tagsToSkip);

                    if (m.Success)
                        inPre = m.Groups[1].Value.Equals ("/") ? false : true;
                }
                else
                {
                    string lastChar = value.Substring (value.Length-1, 1);

                    if (!inPre)
                    {
                        value = ProcessEscapes (value);

                        if (mode == ConversionMode.Stupefy)
                        {
                            value = StupefyEntities (value);
                        }
                        else
                        {
                            switch (mode)
                            {
                                case ConversionMode.EducateDefault:             value = EducateDashes (value); break;
                                case ConversionMode.EducateOldSchool:           value = EducateDashesOldSchool (value); break;
                                case ConversionMode.EducateOldSchoolInverted:   value = EducateDashesOldSchoolInverted (value); break;
                            }

                            value = EducateEllipses (value);

                            // Note: backticks need to be processed before quotes.
                            value = EducateBackticks (value);

                            if (mode == ConversionMode.EducateOldSchool)
                                value = EducateSingleBackticks (value);
    
                            switch (value)
                            {
                                case "'": 
                                    // Special case: single-character ' token
                                    value = Regex.IsMatch (prevTokenLastChar, @"\S") ? "&#8217;" : "&#8216;"; 
                                    break;

                                case "\"": 
                                    // Special case: single-character " token
                                    value = Regex.IsMatch (prevTokenLastChar, @"\S") ? "&#8221;" : "&#8220;";
                                    break;

                                default:
                                    // Normal case:
                                    value = EducateQuotes (value);
                                    break;
                            }
                        }
                    }

                    prevTokenLastChar = lastChar;
                    res += value;
                }
            }

            return res;
        }

        #region "Smart" punctuation

        #region Process quotes (SmartQuotes)

        public string SmartQuotes (string text, ConversionMode mode)
        {
            // Should we educate ``backticks'' -style quotes?
            bool doBackticks = false;

            switch (mode)
            {
                case ConversionMode.LeaveIntact:
                    return text;

                case ConversionMode.EducateOldSchool:
                    doBackticks = true;
                    break;

                default:
                    doBackticks = false;
                    break;
            }

            /*
                Special case to handle quotes at the very end of $text when preceded by
                an HTML tag. Add a space to give the quote education algorithm a bit of
                context, so that it can guess correctly that it's a closing quote:
            */
            bool addExtraSpace = true;

            if (Regex.IsMatch (text, @">['\x22]\z"))
            {
                // Remember, so we can trim the extra space later.
                addExtraSpace = true;
                text += " ";
            }

            ArrayList               tokens = null;
            bool                    inPre = false; // Keep track of when we're inside <pre> or <code> tags
            string                  res = null;

            // This is a cheat, used to get some context
            // for one-character tokens that consist of 
            // just a quote char. What we do is remember
            // the last character of the previous text
            // token, to use as context to curl single-
            // character quote tokens correctly.
            string                  prevTokenLastChar = string.Empty;

            // Parse HTML and convert dashes
            tokens = TokenizeHTML (text);

            foreach (Pair token in tokens)
            {
                string value = token.Second.ToString ();

                if (token.First.Equals ("tag")) 
                {
                    // Don't mess with quotes inside tags
                    res += value;
                    Match m = Regex.Match (value, tagsToSkip);

                    if (m.Success)
                        inPre = m.Groups[1].Value.Equals ("/") ? false : true;
                }
                else
                {
                    string lastChar = value.Substring (value.Length-1, 1);

                    if (!inPre)
                    {
                        value = ProcessEscapes (value);
                        
                        if (doBackticks)
                            value = EducateBackticks (value);

                        switch (value)
                        {
                            case "'": 
                                // Special case: single-character ' token
                                value = (Regex.IsMatch (prevTokenLastChar, @"\S")) ? "&#8217;" : "&#8216;"; 
                                break;

                            case "\"": 
                                // Special case: single-character " token
                                value =  (Regex.IsMatch (prevTokenLastChar, @"\S")) ? "&#8221;" : "&#8220;";
                                break;

                            default:
                                // Normal case:
                                value = EducateQuotes (value);
                                break;
                        }
                    }

                    prevTokenLastChar = lastChar;
                    res += value;
                }
            }

            // Trim trailing space if we added one earlier.
            if (addExtraSpace)
                res = Regex.Replace (res, @" \z", string.Empty);

            return res;
        }

        #endregion

        #region Process dashes (SmartDashes)

        public string SmartDashes (string text, ConversionMode mode)
        {
            ArrayList               tokens = null;
            bool                    inPre = false; // Keep track of when we're inside <pre> or <code> tags
            string                  res = null;
            ConvertDashesDelegate   cd = null;
            
            switch (mode)
            {
                case ConversionMode.LeaveIntact:
                    return text;

                case ConversionMode.EducateDefault:
                    // Reference to the method to use for dash education, default to EducateDashes:
                    cd = new ConvertDashesDelegate (EducateDashes);
                    break;

                case ConversionMode.EducateOldSchool:
                    // Use old smart dash shortcuts, "--" for en, "---" for em
                    cd = new ConvertDashesDelegate (EducateDashesOldSchool);
                    break;

                case ConversionMode.EducateOldSchoolInverted:
                    // Inverse of previous, "--" for em, "---" for en
                    cd = new ConvertDashesDelegate (EducateDashesOldSchoolInverted);
                    break;
            }

            // Parse HTML and convert dashes
            tokens = TokenizeHTML (text);

            foreach (Pair token in tokens)
            {
                string value = token.Second.ToString ();

                if (token.First.Equals ("tag")) 
                {
                    // Don't mess with quotes inside tags
                    res += value;
                    Match m = Regex.Match (value, tagsToSkip);

                    if (m.Success)
                        inPre = m.Groups[1].Value.Equals ("/") ? false : true;
                }
                else
                {
                    if (!inPre)
                    {
                        value = ProcessEscapes (value);
                        value = cd (value);
                    }
                    res += value;
                }
            }
            return res;
        }

        #endregion

        #region Convert ellipses (SmartEllipses)

        public string SmartEllipses (string text, ConversionMode mode)
        {
            if (mode == ConversionMode.LeaveIntact)
                return text;


            ArrayList   tokens = TokenizeHTML (text);
            bool        inPre = false; // Keep track of when we're inside <pre> or <code> tags
            string      res = null;


            foreach (Pair token in tokens)
            {
                string value = token.Second.ToString ();

                if (token.First.Equals ("tag")) 
                {
                    // Don't mess with quotes inside tags
                    res += value;
                    Match m = Regex.Match (value, tagsToSkip);

                    if (m.Success)
                        inPre = m.Groups[1].Value.Equals ("/") ? false : true;
                }
                else
                {
                    if (!inPre)
                    {
                        value = ProcessEscapes (value);
                        value = EducateEllipses (value);
                    }
                    res += value;
                }
            }
            return res;
        }

        #endregion

        #endregion

        #region Convert quotes (EducateQuotes)

	    private string EducateQuotes (string text)
	    {
	        /*
            Parameter:  String.
            Returns:    The string, with "educated" curly quote HTML entities.

            Example input:  "Isn't this fun?"
            Example output: &#8220;Isn&#8217;t this fun?&#8221;
            */

	        // Make our own "punctuation" character class:
	        string punctClass = @"[!\x22#\$\%'()*+,-.\/:;<=>?\@\[\\\]\^_`{|}~]";

	        // Special case if the very first character is a quote
	        // followed by punctuation at a non-word-break. Close the quotes by brute force:
	        text = Regex.Replace (text, string.Format (@"^'(?={0}\B)", punctClass), "&#8217;");
	        text = Regex.Replace (text, string.Format (@"^\x22(?={0}\B)", punctClass), "&#8221;");

	        // Special case for double sets of quotes, e.g.:
	        // <p>He said, "'Quoted' words in a larger quote."</p>
	        text = Regex.Replace (text, @"\x22'(?=\w)", "&#8220;&#8216;");
	        text = Regex.Replace (text, @"'\x22(?=\w)", "&#8216;&#8220;");

	        // 	Special case for decade abbreviations (the '80s):
	        text = Regex.Replace (text, @"'(?=\d{2}s)", "&#8217;");

	        string closeClass = @"[^\ \t\r\n\[\{\(\-]";
	        string decDashes = @"&\#8211;|&\#8212;";

	        // Get most opening single quotes:
	        text = Regex.Replace (text, string.Format(@"
		            (
			            \s          |   # a whitespace char, or
			            &nbsp;      |   # a non-breaking space entity, or
			            --          |   # dashes, or
			            &[mn]dash;  |   # named dash entities
			            {0}         |   # or decimal entities
			            &\#x201[34];    # or hex
		            )
		            '                   # the quote
		            (?=\w)              # followed by a word character
            ", decDashes), "$1&#8216;", RegexOptions.IgnorePatternWhitespace);

	        // Single closing quotes:
	        text = Regex.Replace (text, string.Format (@"
                ({0})?
		        '
		        (?(1)|          # If $1 captured, then do nothing;
		          (?=\s | s\b)  # otherwise, positive lookahead for a whitespace
		        )               # char or an 's' at a word ending position.
                ", closeClass), "$1&#8217;", RegexOptions.IgnorePatternWhitespace | RegexOptions.IgnoreCase);

	        // Any remaining single quotes should be opening ones:
	        text = text.Replace ("'", "&#8216;");

	        // Get most opening double quotes:
	        text = Regex.Replace (text, string.Format (@"
                    (
			            \s          |   # a whitespace char, or
			            &nbsp;      |   # a non-breaking space entity, or
			            --          |   # dashes, or
			            &[mn]dash;  |   # named dash entities
			            {0}         |   # or decimal entities
			            &\#x201[34];    # or hex
		            )
		            \x22                # the quote
		            (?=\w)              # followed by a word character
                 ", decDashes), "$1&#8220;", RegexOptions.IgnorePatternWhitespace);

	        // Double closing quotes:
	        text = Regex.Replace (text, string.Format(@"
                    ({0})?
		            \x22
		            (?(1)|(?=\s))       # If $1 captured, then do nothing;
                                        # if not, then make sure the next char is whitespace.
                ", closeClass), "$1&#8221;", RegexOptions.IgnorePatternWhitespace);
 
	        // Any remaining quotes should be opening ones.
	        text = text.Replace ("\"", "&#8220;");

	        return text;
	    }

	    #endregion

	    #region Process backticks, dashes, ellipses

	    // --------------------------------------------------------------------------------------
	    private string EducateBackticks (string text)
	    {
	        /*
            Parameter:  String.
            Returns:    The string, with ``backticks'' -style double quotes
                        translated into HTML curly quote entities.

            Example input:  ``Isn't this fun?''
            Example output: &#8220;Isn't this fun?&#8221;
            */

	        return text.Replace ("``", "&#8220;").Replace ("''", "&#8221;");
	    }

	    // --------------------------------------------------------------------------------------
	    private string EducateSingleBackticks (string text)
	    {
	        /*
            Parameter:  String.
            Returns:    The string, with `backticks' -style single quotes
                        translated into HTML curly quote entities.

            Example input:  `Isn't this fun?'
            Example output: &#8216;Isn&#8217;t this fun?&#8217;
            */

	        return text.Replace ("`", "&#8216;").Replace ("'", "&#8217;");
	    }

	    // --------------------------------------------------------------------------------------
	    private string EducateDashes (string text)
	    {
	        /*
            Parameter:  String.

            Returns:    The string, with each instance of "--" translated to
                        an em-dash HTML entity.
            */

	        return text.Replace ("--", "&#8212;");
	    }
        
	    // --------------------------------------------------------------------------------------
	    private string EducateDashesOldSchool (string text)
	    {
	        /*
            Parameter:  String.

            Returns:    The string, with each instance of "--" translated to
                        an en-dash HTML entity, and each "---" translated to
                        an em-dash HTML entity.
            */

	        //                             em                         en
	        return text.Replace ("---", "&#8212;").Replace ("--", "&#8211;");
	    }

	    // --------------------------------------------------------------------------------------
	    private string EducateDashesOldSchoolInverted (string text)
	    {
	        /*
            Parameter:  String.

            Returns:    The string, with each instance of "--" translated to
                        an em-dash HTML entity, and each "---" translated to
                        an en-dash HTML entity. Two reasons why: First, unlike the
                        en- and em-dash syntax supported by
                        EducateDashesOldSchool(), it's compatible with existing
                        entries written before SmartyPants 1.1, back when "--" was
                        only used for em-dashes.  Second, em-dashes are more
                        common than en-dashes, and so it sort of makes sense that
                        the shortcut should be shorter to type. (Thanks to Aaron
                        Swartz for the idea.)
            */

	        //                             en                         em
	        return text.Replace ("---", "&#8211;").Replace ("--", "&#8212;");
	    }

	    // --------------------------------------------------------------------------------------
	    private string EducateEllipses (string text)
	    {
	        /*
            Parameter:  String.
            Returns:    The string, with each instance of "..." translated to
                        an ellipsis HTML entity. Also converts the case where
                        there are spaces between the dots.

            Example input:  Huh...?
            Example output: Huh&#8230;?
            */
	        return text.Replace ("...", "&#8230;").Replace (". . .", "&#8230;");
	    }

	    // --------------------------------------------------------------------------------------
	    private string StupefyEntities (string text)
	    {
	        /*
            Parameter:  String.
            Returns:    The string, with each SmartyPants HTML entity translated to
                        its ASCII counterpart.

            Example input:  &#8220;Hello &#8212; world.&#8221;
            Example output: "Hello -- world."
            */

	        //                     en-dash                  em-dash
	        text = text.Replace ("&#8211;", "-").Replace ("&#8212;", "--");

	        // single quote        open                     close
	        text = text.Replace ("&#8216;", "'").Replace ("&#8217;", "'");

	        // double quote        open                     close
	        text = text.Replace ("&#8220;", "\"").Replace ("&#8221;", "\"");

	        // ellipsis
	        text = text.Replace ("&#8230;", "...");

	        return text;
	    }

	    // --------------------------------------------------------------------------------------
	    private string ProcessEscapes (string text)
	    {
	        /*
               Parameter:  String.
               Returns:    The string, with after processing the following backslash
                           escape sequences. This is useful if you want to force a "dumb"
                           quote or other character to appear.
            
                           Escape  Value
                           ------  -----
                           \\      &#92;
                           \"      &#34;
                           \'      &#39;
                           \.      &#46;
                           \-      &#45;
                           \`      &#96;
            */

	        Hashtable escapes = new Hashtable ();
	        escapes ["\\\\"]    = 0x92;
	        escapes ["\\\""]    = 0x34;
	        escapes ["\\'"]     = 0x39;
	        escapes ["\\."]     = 0x46;
	        escapes ["\\-"]     = 0x45;
	        escapes ["\\`"]     = 0x96;

	        foreach (string key in escapes.Keys)
	            text = text.Replace (key, string.Format("&#{0:x};", escapes [key]));

	        return text;
	    }

	    #endregion

        #region Parse HTML into tokens

        /// <summary>
        /// 
        /// </summary>
        /// <param name="text">String containing HTML markup.</param>
        /// <returns>An array of the tokens comprising the input string. Each token is 
        /// either a tag (possibly with nested, tags contained therein, such 
        /// as &lt;a href="<MTFoo>"&gt;, or a run of text between tags. Each element of the 
        /// array is a two-element array; the first is either 'tag' or 'text'; the second is 
        /// the actual value.
        /// </returns>
        private ArrayList TokenizeHTML (string text)
        {
            // Regular expression derived from the _tokenize() subroutine in 
            // Brad Choate's MTRegex plugin.
            // http://www.bradchoate.com/past/mtregex.php
            int                 pos = 0;
            int                 depth = 6;
            ArrayList           tokens = new ArrayList ();

            string nestedTags = string.Concat (RepeatString (@"(?:<[a-z\/!$](?:[^<>]|", depth), RepeatString (@")*>)", depth));
            string pattern = string.Concat (@"(?s:<!(?:--.*?--\s*)+>)|(?s:<\?.*?\?>)|", nestedTags);

            MatchCollection mc = Regex.Matches (text, pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline);

            foreach (Match m in mc)
            {
                string wholeTag = m.Value;
                int tagStart = m.Index;
                Pair token = null;

                if (pos < tagStart)
                {
                    token = new Pair ();
                    token.First = "text";
                    token.Second = text.Substring (pos, tagStart - pos);
                    tokens.Add (token);
                }

                token = new Pair ();
                token.First = "tag";
                token.Second = wholeTag;
                tokens.Add (token);

                pos = m.Index + m.Length;
            }

            if (pos < text.Length)
            {
                Pair token = new Pair ();
                token.First = "text";
                token.Second = text.Substring (pos, text.Length - pos);
                tokens.Add (token);
            }

            return tokens;
        }

        #endregion

        private string RepeatString (string text, int count)
        {
            string res = null;

            for (int i=0; i<count; i++)
                res += text;

            return res;
        }

    }
}
