using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace pp_source_format
{
    public static class Formatter
    {
        public static void FormatShape(Shape s, string language, string style)
        {
            var sourceText = s.TextFrame.TextRange.Text;
            var fontName = GuessSensibleFont(s.TextFrame.TextRange);

            var formattedText = RunPygments(sourceText, language, style);
            Clipboard.SetData(DataFormats.Text, formattedText);
            Clipboard.SetData(DataFormats.Rtf, formattedText);
            s.TextFrame.TextRange.PasteSpecial(PpPasteDataType.ppPasteRTF);

            // Pasting has removed the font
            s.TextFrame.TextRange.Font.Name = fontName;
        }

        private static string GuessSensibleFont(TextRange t)
        {
            if (!String.IsNullOrEmpty(t.Font.Name))
            {
                return t.Font.Name;
            }

            foreach (TextRange c in t.Characters())
            {
                if (!String.IsNullOrEmpty(c.Font.Name)) {
                    return c.Font.Name;
                }
            }

            return "Consolas";
        }

        private static string RunPygments(string input, string language, string style)
        {
            var arguments = String.Format("-f rtf -l \"{0}\" -O \"style={1}\"", language, style);
            if (false )
            {
                arguments += "-o \"C:/Users/Marcus Riemer/source/repos/MarcusRiemer/powerpoint-source-code-format/pp-source-format/test.rtf\"";
            }
            var startInfo = new ProcessStartInfo()
            {
                FileName = FindPygmentizePath(),
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                CreateNoWindow = true,
            };

            Process p = new Process()
            {
                StartInfo = startInfo
            };

            p.Start();
            p.StandardInput.Write(input);
            p.StandardInput.Close();

            p.WaitForExit();

            var result = p.StandardOutput.ReadToEnd();
            var error = p.StandardError.ReadToEnd();

            if (!String.IsNullOrWhiteSpace(error))
            {
                throw new Exception(String.Format("Error running pygmentize with arguments {0}:\n{1}", arguments, error));
            } 
            return result;
        }

        public static string FindPygmentizePath()
        {
            return FindExePath("pygmentize.exe");
        }

        private static string FindPygmentizeFromPythonPath()
        {
            var pythonExe = FindExePath("python.exe");
            var pythonDir = Path.GetDirectoryName(pythonExe);
            var pygmentsExe = Path.Combine(pythonDir, "Scripts", "pygmentize.exe");

            if (!File.Exists(pygmentsExe))
            {
                throw new FileNotFoundException(String.Format("Pygments formatter could not be found: {0}", pygmentsExe));
            }

            return Path.GetFullPath(pygmentsExe);
        }

        /// <summary>
        /// Expands environment variables and, if unqualified, locates the exe in the working directory
        /// or the evironment's path.
        /// </summary>
        /// <param name="exe">The name of the executable file</param>
        /// <returns>The fully-qualified path to the file</returns>
        /// <exception cref="System.IO.FileNotFoundException">Raised when the exe was not found</exception>
        private static string FindExePath(string exe)
        {
            exe = Environment.ExpandEnvironmentVariables(exe);
            if (!File.Exists(exe))
            {
                if (Path.GetDirectoryName(exe) == String.Empty)
                {
                    foreach (string test in (Environment.GetEnvironmentVariable("PATH") ?? "").Split(';'))
                    {
                        string path = test.Trim();
                        if (!String.IsNullOrEmpty(path) && File.Exists(path = Path.Combine(path, exe)))
                            return Path.GetFullPath(path);
                    }
                }
                throw new FileNotFoundException(new FileNotFoundException().Message, exe);
            }
            return Path.GetFullPath(exe);
        }
    }
}
