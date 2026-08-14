using System;
using System.Globalization;
using System.IO;
using System.IO.Pipes;
using System.Text.RegularExpressions;

namespace RecallTape.ProtocolHandler
{
    /// <summary>
    /// Turns a click on a tape strip into a command the add-in can act on.
    ///
    /// WHY THIS EXISTS AT ALL: OneNote hands an Image's `hyperlink` to the shell, which launches us.
    /// We cannot do the work here -- on Office 16.0.20228 an external process gets a OneNote
    /// Application object whose every method returns E_FAIL (see docs/analysis/). Only code running
    /// inside the add-in surrogate can touch the page. So this executable is a courier: parse the URL,
    /// hand it to the add-in over a named pipe, exit.
    ///
    /// It is also the one part of RecallTape reachable from outside, because a `recalltape://` link
    /// could be planted in any document. It therefore validates hard and trusts nothing.
    /// </summary>
    internal static class Program
    {
        private const string PipeName = "RecallTape.Commands";

        // Deliberately strict: verb is a short lowercase word, id is a bare GUID. Nothing else gets
        // through, so a hostile link cannot smuggle a path, an argument, or a type name.
        private static readonly Regex UrlPattern = new Regex(
            @"^recalltape://(?<verb>[a-z]{1,16})/(?<id>[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})/?$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static int Main(string[] args)
        {
            string raw = args.Length > 0 ? args[0] : "(none)";
            try
            {
                var m = UrlPattern.Match(raw.Trim());
                if (!m.Success)
                {
                    Log("REJECTED (does not match the expected shape): " + raw);
                    return 1;
                }

                string command = m.Groups["verb"].Value.ToLowerInvariant() + " " + m.Groups["id"].Value;
                Log("accepted: " + command);

                using (var pipe = new NamedPipeClientStream(".", PipeName, PipeDirection.Out))
                {
                    // Short timeout: if the add-in is not listening, OneNote is probably not running
                    // and there is nothing useful to do. Fail fast rather than hang a shell process.
                    pipe.Connect(3000);
                    using (var writer = new StreamWriter(pipe) { AutoFlush = true })
                    {
                        writer.WriteLine(command);
                    }
                }

                Log("delivered");
                return 0;
            }
            catch (Exception ex)
            {
                Log("FAILED: " + ex.Message);
                return 2;
            }
        }

        /// <summary>
        /// Its own log file, separate from the add-in's. When a three-process chain breaks -- OneNote,
        /// this courier, the add-in -- the only cheap way to find out how far it got is for each link
        /// to say so independently.
        /// </summary>
        private static void Log(string message)
        {
            try
            {
                string dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "EWC3 Labs", "RecallTape");
                Directory.CreateDirectory(dir);
                File.AppendAllText(Path.Combine(dir, "protocol.log"),
                    DateTime.Now.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture)
                    + "| " + message + Environment.NewLine);
            }
            catch { /* a courier that cannot log still delivers */ }
        }
    }
}
