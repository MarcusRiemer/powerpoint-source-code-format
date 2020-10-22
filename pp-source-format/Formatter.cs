using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace pp_source_format
{
    /// <summary>
    /// Encapsulates the heavy lifting of the formatting process. This boils down to the following steps:
    /// 
    /// 1) Prepare the source code by stripping or replacing some Powerpoint-specifics 
    /// 2) Find out which binary to run to do the formatting and run it
    /// 3) Convert the result to be in Win32 HTML Clipboard format which is digestible by Powerpoint
    /// 4) Paste (!!) the snippet in the correct location. Yes, this uses the clipboard.
    /// </summary>
    public static class Formatter
    {
        // I miss Java enums with their own methods :( Extension methods dont quite cut it

        /// <summary>
        /// Output formats that seem to be sensible when pasting in to Powerpoint. The RTFs produced
        /// by pygments worked fine when opening and pasting from Word but had various issues with
        /// Powerpoint so I went for HTML.
        /// </summary>
        enum Format
        {
            HTML,
            RTF,
        };


        /// <summary>
        /// Format -> Powerpoint
        /// </summary>
        static PpPasteDataType PowerpointPasteType(this Format f)
        {
            switch (f)
            {
                case Format.HTML:
                    return PpPasteDataType.ppPasteHTML;
                case Format.RTF:
                    return PpPasteDataType.ppPasteRTF;
                default:
                    throw new Exception("Unknown paste type: " + f.ToString());

            }
        }

        /// <summary>
        /// Format -> Clipboard data format
        /// </summary>
        static string DataFormat(this Format f)
        {
            switch (f)
            {
                case Format.HTML:
                    return DataFormats.Html;
                case Format.RTF:
                    return DataFormats.Rtf;
                default:
                    throw new Exception("Unknown data format: " + f.ToString());

            }
        }

        /// <summary>
        /// Format to pygments -f parameter
        /// </summary>
        static string PygmentsFormat(this Format f)
        {
            switch (f)
            {
                case Format.HTML:
                    return "html";
                case Format.RTF:
                    return "rtf";
                default:
                    throw new Exception("Unknown pygments format: " + f.ToString());

            }
        }

        /// <summary>
        /// Exactly what it says on the tin.
        /// </summary>
        /// <param name="s"></param>
        /// <param name="language"></param>
        /// <param name="style"></param>
        public static void FormatShape(Shape s, string language, string style)
        {
            // Remember the old state of the clipboard to restore it after we used it as
            // a way to interface with Powerpoint.
            var previousClipboard = Clipboard.GetDataObject();

            try
            {

                // Every paste option I tried had different issues, so I use this
                // switch to switch between them.
                Format format = Format.HTML;

                // The source code input string
                var sourceText = s.TextFrame.TextRange.Text;

                // The original font is overridden by the pasting operation, so we remember it here
                var fontName = GuessSensibleFont(s.TextFrame.TextRange);

                // Below this point `formattedText` describes a perfectly valid HTML or
                // RTF document that is displayed just fine by any word processor or
                // browser that I am aware of. But for whatever reason powerpoint
                // badly chokes on the import.

                bool multiPaste = false;
                if (multiPaste)
                {
                    DataObject d = new DataObject();
                    foreach (Format item in Enum.GetValues(typeof(Format)))
                    {
                        var formattedText = RunPygments(item, sourceText, language, style);
                        d.SetData(item.DataFormat(), formattedText);
                    }
                    Clipboard.SetDataObject(d);
                }
                else
                {
                    // Actually run pygments and pipe it to the clipboard
                    var formattedText = RunPygments(format, sourceText, language, style);
                    Clipboard.SetData(format.DataFormat(), formattedText);

                    // Possibly try a hacky HTML -> RTF conversion?
                    // This seems to loose colour informarmation
                    if (false && format == Format.HTML)
                    {
                        formattedText = HtmlToRtf(formattedText);
                        format = Format.RTF;
                        Clipboard.SetData(format.DataFormat(), formattedText);
                    }
                }

                // This seems to be the sensible way to paste into Powerpoint
                // But for whatever reason it leeds to "bleeding" of colours once
                // the color should reset to black
                s.TextFrame.TextRange.PasteSpecial(format.PowerpointPasteType());

                // An alternative way to paste,
                // inspired by https://stackoverflow.com/questions/33493983/vsto-powerpoint-notes-page-different-colored-words-on-same-line/43210187#43210187
                //s.Select();
                //Globals.SourceCodeFormatAddin.Application.CommandBars.ExecuteMso("PasteSourceFormatting");
                //System.Windows.Forms.Application.DoEvents();

                // Pasting has removed the font
                s.TextFrame.TextRange.Font.Name = fontName;
            }
            finally
            {
                //Clipboard.SetDataObject(previousClipboard);
            }
        }

        /// <summary>
        /// Best effort guess to keep the existing font.
        /// </summary>
        /// <param name="t">Sample text range</param>
        /// <returns>A sensible font</returns>
        private static string GuessSensibleFont(TextRange t)
        {
            if (!String.IsNullOrEmpty(t.Font.Name))
            {
                return t.Font.Name;
            }

            foreach (TextRange c in t.Characters())
            {
                if (!String.IsNullOrEmpty(c.Font.Name))
                {
                    return c.Font.Name;
                }
            }

            return "Consolas";
        }

        /// <summary>
        /// Starts an external pygments process and collects the results.
        /// </summary>
        /// <param name="format">The output format to use</param>
        /// <param name="input">The code to format</param>
        /// <param name="language">The programming language to use</param>
        /// <param name="style">The style to use</param>
        /// <returns>A highlighted document, format according to the parameter</returns>
        private static string RunPygments(Format format, string input, string language, string style)
        {
            // For debug purposes it may come in handy to see the actual results in file form
            var filePath = Path.Combine(Path.GetTempPath(), "pp-format." + format.PygmentsFormat());
            bool useFilePath = false;

            var options = new List<string>(new string[] {
                "style=" + style,
            });
            if (format == Format.HTML)
            {
                options.AddRange(new string[] {
                    "noclasses=true",
                    "nowrap=true",
                    //"full=true",
                    "lineseparator=<br>",
                });
            }


            var allArguments = new List<string>(new string[] {
                "-f " + format.PygmentsFormat(),
                "-l " + language,
                "-O " + '"' + string.Join(",", options.ToArray()) + '"'
            });

            if (useFilePath)
            {
                allArguments.Insert(0, "-o " + '"' + filePath + '"');
            }

            var arguments = String.Join(" ", allArguments.ToArray());
            var startInfo = new ProcessStartInfo()
            {
                FileName = Pygments.PygmentizePath,
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                CreateNoWindow = true,
            };

            // The highlighting and powerpoint might do some bad things to the
            // pre-formatted input. So we reverse some of these issues:
            // * Swap "not quite breaks" that powerpoint uses for "proper" breaks (no, \n doesnt work)
            // * Swap all non breaking spaces for normal spaces
            input = input.Replace("\v", "\r").Replace("\u00A0", " ");

            Process p = new Process()
            {
                StartInfo = startInfo
            };

            p.Start();
            p.StandardInput.Write(input);
            p.StandardInput.Close();

            p.WaitForExit();


            var error = p.StandardError.ReadToEnd();
            if (!String.IsNullOrWhiteSpace(error))
            {
                throw new Exception(String.Format("Temp file at {0}, Error running pygmentize with arguments {0}:\n{1}", filePath, arguments, error));
            }

            var result = useFilePath ? File.ReadAllText(filePath) : p.StandardOutput.ReadToEnd();

            if (format == Format.HTML)
            {
                // Wrap the result in a valid, minimal document with the Fragments as required by the HTML Clipboard format
                // https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa767917(v=vs.85)
                result = "<html><body>" + ClipboardHelper.StartFragment + "<pre>" + result + "</pre>" + ClipboardHelper.EndFragment + "</body></html>";

                // Whitespace at the beginning of lines seems to be collapsed, sadly neither of the
                // tricks at https://stackoverflow.com/questions/47475774/how-to-add-spaces-to-html-clipboard-data-so-that-winword-inserts-them-on-pasting
                // worked for me
                //result = result.Replace("<span style=\"", "<span style=\"white-space: pre;");

                // Remove last newline, that line is useless for MS Word
                result = result.RemoveLastOccurrence("<br>");

                // Okay, this is the most nasty part ... We insert &nbsp; after each linebreak to preserve
                // whitespace formatting. Each whitespace must be replaced by a single non breaking space.
                result = Regex.Replace(result, "(<br> *)", delegate (Match m)
                {
                    return m.Value.Replace(" ", "&nbsp;");
                });

                if (useFilePath)
                {
                    File.WriteAllText(filePath, result);
                }


                result = ClipboardHelper.GetHtmlDataString(result);
            }


            return result;
        }

        /// <summary>
        /// Didn't find a method in the standard library so I hacked my own.
        /// </summary>
        private static string RemoveLastOccurrence(this string Source, string Find)
        {
            int place = Source.LastIndexOf(Find);

            if (place == -1)
                return Source;
            else
                return Source.Remove(place, Find.Length);
        }

        /// <summary>
        /// Marvelous hack as proposed at https://stackoverflow.com/questions/33493983/
        /// Sadly it seems to loose all colour information.
        /// </summary>
        /// <param name="toPaste"></param>
        /// <returns></returns>
        private static string HtmlToRtf(string toPaste)
        {
            var web = new WebBrowser();
            web.CreateControl();
            web.DocumentText = toPaste;
            while (web.DocumentText != toPaste)
            {
                System.Windows.Forms.Application.DoEvents();
            }
            web.Document.ExecCommand("SelectAll", false, null);
            web.Document.ExecCommand("Copy", false, null);
            //web.Dispose();
            return Clipboard.GetData(DataFormats.Rtf) as string;
        }
    }
}
