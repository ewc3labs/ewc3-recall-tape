using System;
using System.IO;
using System.IO.Pipes;
using System.Text.RegularExpressions;
using System.Threading;

namespace RecallTape.OneNote
{
    /// <summary>
    /// Listens for commands from outside the surrogate -- currently just clicks on tape strips,
    /// couriered here by RecallTape.ProtocolHandler.exe.
    ///
    /// This is the "reverse COM proxy" shape: OneNote will not serve an external process, so instead of
    /// the outside world reaching in, the add-in listens. Same pattern OneMore uses for onemore:// links
    /// (see docs/analysis/onemore-onenote-interaction.md).
    ///
    /// It is an inbound surface, so it is deliberately small and suspicious: one line, one verb, one
    /// GUID, validated again here even though the courier already validated it. Two checks, because the
    /// courier is an executable a hostile caller could invoke directly.
    /// </summary>
    public partial class AddIn
    {
        private const string PipeName = "RecallTape.Commands";
        private const int MaxConsecutiveErrors = 5;

        private static readonly Regex CommandPattern = new Regex(
            @"^(?<verb>[a-z]{1,16}) (?<id>[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})$",
            RegexOptions.Compiled);

        private void StartCommandService()
        {
            var thread = new Thread(ListenLoop)
            {
                IsBackground = true,     // must never keep OneNote alive
                Name = "RecallTape.CommandService"
            };
            thread.Start();
            Log("command service listening on pipe " + PipeName);
        }

        private void ListenLoop()
        {
            int errors = 0;

            // Bounded error tolerance rather than an unbounded retry: a pipe that fails forever should
            // give up quietly, not spin a background thread for the life of the OneNote session.
            while (errors < MaxConsecutiveErrors)
            {
                try
                {
                    using (var server = new NamedPipeServerStream(PipeName, PipeDirection.In))
                    {
                        server.WaitForConnection();
                        using (var reader = new StreamReader(server))
                        {
                            string line = reader.ReadLine();
                            if (!string.IsNullOrEmpty(line)) Dispatch(line.Trim());
                        }
                    }
                    errors = 0;
                }
                catch (Exception ex)
                {
                    errors++;
                    Log("command service error " + errors + ": " + ex.Message);
                    Thread.Sleep(500);
                }
            }

            Log("command service stopped after " + errors + " consecutive errors");
        }

        private void Dispatch(string line)
        {
            var m = CommandPattern.Match(line);
            if (!m.Success)
            {
                Log("command REJECTED: " + line);
                return;
            }

            string verb = m.Groups["verb"].Value;
            string id = m.Groups["id"].Value;
            Log("command: " + verb + " " + id);

            switch (verb)
            {
                case "toggle":
                    TogglePeek(id);
                    break;
                default:
                    Log("command: unknown verb " + verb);
                    break;
            }
        }
    }
}
