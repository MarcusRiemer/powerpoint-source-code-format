using System;
using System.Text;
using System.Windows.Forms;
/// <summary>
/// Helper to encode and set HTML fragment to clipboard.<br/>
/// See http://theartofdev.com/2014/06/12/setting-htmltext-to-clipboard-revisited/.<br/>
/// <seealso cref="CreateDataObject"/>.
/// </summary>
/// <remarks>
/// The MIT License (MIT) Copyright (c) 2014 Arthur Teplitzki.
/// </remarks>
public static class ClipboardHelper
{
    #region Fields and Consts

    /// <summary>
    /// The string contains index references to other spots in the string, so we need placeholders so we can compute the offsets. <br/>
    /// The <![CDATA[<<<<<<<]]>_ strings are just placeholders. We'll back-patch them actual values afterwards. <br/>
    /// The string layout (<![CDATA[<<<]]>) also ensures that it can't appear in the body of the html because the <![CDATA[<]]> <br/>
    /// character must be escaped. <br/>
    /// </summary>
    private const string Header = @"Version:0.9
StartHTML:<<<<<<<<1
EndHTML:<<<<<<<<2
StartFragment:<<<<<<<<3
EndFragment:<<<<<<<<4
StartSelection:<<<<<<<<3
EndSelection:<<<<<<<<4";

    /// <summary>
    /// html comment to point the beginning of html fragment
    /// </summary>
    public const string StartFragment = "<!--StartFragment-->";

    /// <summary>
    /// html comment to point the end of html fragment
    /// </summary>
    public const string EndFragment = @"<!--EndFragment-->";

    /// <summary>
    /// Used to calculate characters byte count in UTF-8
    /// </summary>
    private static readonly char[] _byteCount = new char[1];

    #endregion

    /// <summary>
    /// Generate HTML fragment data string with header that is required for the clipboard.
    /// </summary>
    /// <param name="html">the html to generate for</param>
    /// <returns>the resulted string</returns>
    public static string GetHtmlDataString(string html)
    {
        var sb = new StringBuilder();
        sb.AppendLine(Header);
        sb.AppendLine(@"<!DOCTYPE html>");

        // if given html already provided the fragments we won't add them
        int fragmentStart, fragmentEnd;
        int fragmentStartIdx = html.IndexOf(StartFragment, StringComparison.OrdinalIgnoreCase);
        int fragmentEndIdx = html.LastIndexOf(EndFragment, StringComparison.OrdinalIgnoreCase);

        // if html tag is missing add it surrounding the given html (critical)
        int htmlOpenIdx = html.IndexOf("<html", StringComparison.OrdinalIgnoreCase);
        int htmlOpenEndIdx = htmlOpenIdx > -1 ? html.IndexOf('>', htmlOpenIdx) + 1 : -1;
        int htmlCloseIdx = html.LastIndexOf("</html", StringComparison.OrdinalIgnoreCase);

        if (fragmentStartIdx < 0 && fragmentEndIdx < 0)
        {
            throw new Exception("Start and End fragments are required");
        }
        else
        {
            // handle html with existing start\end fragments just need to calculate the correct bytes offset (surround with html tag if missing)
            if (htmlOpenEndIdx < 0)
                sb.Append("<html>");
            int start = GetByteCount(sb);
            sb.Append(html);
            fragmentStart = start + GetByteCount(sb, start, start + fragmentStartIdx) + StartFragment.Length;
            fragmentEnd = start + GetByteCount(sb, start, start + fragmentEndIdx);
            if (htmlCloseIdx < 0)
                sb.Append("</html>");
        }

        // Back-patch offsets (scan only the header part for performance)
        sb.Replace("<<<<<<<<1", Header.Length.ToString("D9"), 0, Header.Length);
        sb.Replace("<<<<<<<<2", GetByteCount(sb).ToString("D9"), 0, Header.Length);
        sb.Replace("<<<<<<<<3", fragmentStart.ToString("D9"), 0, Header.Length);
        sb.Replace("<<<<<<<<4", fragmentEnd.ToString("D9"), 0, Header.Length);

        return sb.ToString();
    }

    /// <summary>
    /// Calculates the number of bytes produced by encoding the string in the string builder in UTF-8 and not .NET default string encoding.
    /// </summary>
    /// <param name="sb">the string builder to count its string</param>
    /// <param name="start">optional: the start index to calculate from (default - start of string)</param>
    /// <param name="end">optional: the end index to calculate to (default - end of string)</param>
    /// <returns>the number of bytes required to encode the string in UTF-8</returns>
    private static int GetByteCount(StringBuilder sb, int start = 0, int end = -1)
    {
        int count = 0;
        end = end > -1 ? end : sb.Length;
        for (int i = start; i < end; i++)
        {
            _byteCount[0] = sb[i];
            count += Encoding.UTF8.GetByteCount(_byteCount);
        }
        return count;
    }
}