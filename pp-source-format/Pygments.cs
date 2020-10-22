using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pp_source_format
{
    static class Pygments
    {
        public static bool FoundPygmentize
        {
            get
            {
                try
                {
                    return (!string.IsNullOrEmpty(PygmentizePath));
                }
                catch
                {
                    return false;
                }
            }
        }

        public static string PygmentizePath
        {
            get => FindExePath("pygmentize.exe");
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
