using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Management.Automation;
using System.Threading;
using System.Threading.Tasks;
using Veritec.Dynamics.CI.Common;

namespace Veritec.Dynamics.CI.PowerShell
{
    [Cmdlet("Import", "DynamicsSolution")]
    [OutputType(typeof(Boolean))]
    public class ImportDynamicsSolution : Cmdlet
    {

        [Parameter(Mandatory = true)]
        public string ConnectionString { get; set; }

        [Parameter(Mandatory = true)]
        public string SolutionName { get; set; }

        [Parameter(Mandatory = false)]
        public int ConnectionTimeOutMinutes { get; set; } = 30;

        [Parameter(Mandatory = false)]
        public bool Managed { get; set; }

        [Parameter(Mandatory = false)]
        public string SolutionDir { get; set; } = Directory.GetCurrentDirectory();

        private ConcurrentQueue<string> MessageQueue { get; set; }

        protected override void ProcessRecord()
        {
            MessageQueue = new ConcurrentQueue<string>();

            try
            {
                var crmParameter = new CrmParameter(ConnectionString)
                {
                    ConnectionTimeOutMinutes = ConnectionTimeOutMinutes
                };

                SolutionDir = SolutionDir + @"\";
                crmParameter.ExecutionDirectory = SolutionDir;

                /* Connect to Dynamics */
                WriteObject($"Connecting ({crmParameter.GetConnectionStringObfuscated()})");

                SolutionTool solutionTool = null;

                var stopwatch = Stopwatch.StartNew();
                var taskInstantiateSolution = Task.Run(() => solutionTool = new SolutionTool(crmParameter));

                while (!taskInstantiateSolution.IsCompleted)
                {
                    WriteVerbose(".");
                    Thread.Sleep(2000);
                }

                solutionTool.MessageLogger += ReceiveMessage;

                var elapsed = $"{stopwatch.Elapsed.Minutes}min {stopwatch.Elapsed.Seconds}s";
                WriteObject($"Done... [{elapsed}]{Environment.NewLine}");

                /* Import Solutions */
                var solutionNames = SolutionName.Split(';');
                foreach (var sol in solutionNames)
                {

                    stopwatch = Stopwatch.StartNew();
                    var fileName = SolutionDir + sol + (Managed ? "_managed" : "") + ".zip";

                    WriteObject($"Uploading '{sol.ToUpper()}' solution ");
                    var solutionBinaryTask = solutionTool.ImportSolutionAsync(fileName);

                    while (!solutionBinaryTask.IsCompleted)
                    {
                        WriteVerbose(".");
                        Thread.Sleep(3000);

                        /* check if any messages from the solution import thread 
                         * have been passed back */
                        while (true)
                        {
                            if (MessageQueue.Count == 0)
                                break; // exit while

                            var MessageFound = MessageQueue.TryDequeue(out string message);

                            if (MessageFound)
                                WriteObject(message);
                        }
                    }

                    if (solutionBinaryTask.IsFaulted)
                    {
                        foreach (var innerException in solutionBinaryTask.Exception.InnerExceptions)
                        {
                            WriteObject($"ERROR - {innerException.Message}");
                            throw new Exception();
                        }
                    }

                    elapsed = $"{stopwatch.Elapsed.Minutes}min {stopwatch.Elapsed.Seconds}s";
                    WriteObject($"Done... [{elapsed}]{Environment.NewLine}");

                    stopwatch = Stopwatch.StartNew();
                    WriteObject($"Publishing ");

                    /* publish solution */
                    var publishTask = solutionTool.PublishAsync();

                    while (!publishTask.IsCompleted)
                    {
                        WriteVerbose(".");
                        Thread.Sleep(3000);
                    }

                    elapsed = $"{stopwatch.Elapsed.Minutes}min {stopwatch.Elapsed.Seconds}s";
                    WriteObject($"Done... [{elapsed}]{Environment.NewLine}");

                }
            }
            catch (Exception ex)
            {
                var errorRecord = new ErrorRecord(new Exception($"Solution Import Failed: {ex.Message}", ex.InnerException), "", ErrorCategory.InvalidResult, null);
                ThrowTerminatingError(errorRecord);
            }
        }

        private void ReceiveMessage(object sender, string e)
        {
            MessageQueue.Enqueue(e);
        }
    }
}
