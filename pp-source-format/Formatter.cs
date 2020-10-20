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
        public static void FormatShape(Shape s)
        {
            var sourceText = s.TextFrame.TextRange.Text;
            var formattedText = RunPygments(sourceText);
            Clipboard.SetData(DataFormats.Rtf, formattedText);
            s.TextFrame.TextRange.PasteSpecial(PpPasteDataType.ppPasteRTF);
        }

        private static string RunPygments(string input)
        {
            var startInfo = new ProcessStartInfo()
            {
                FileName = FindPygmentizePath(),
                Arguments = "-f rtf -l java",
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardError = true,
                RedirectStandardOutput = true,
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

            return result;
        }

        public static string FindPygmentizePath()
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
        public static string FindExePath(string exe)
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
