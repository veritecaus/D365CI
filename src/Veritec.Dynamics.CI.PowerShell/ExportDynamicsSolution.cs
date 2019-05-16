using System;
using System.Diagnostics;
using System.IO;
using System.Management.Automation;
using System.Threading;
using System.Threading.Tasks;
using Veritec.Dynamics.CI.Common;

namespace Veritec.Dynamics.CI.PowerShell
{
    [Cmdlet("Export", "DynamicsSolution")]
    public class ExportDynamicsSolution : Cmdlet
    {
        [Parameter(Mandatory = true)]
        public string ConnectionString { get; set; }

        [Parameter(Mandatory = true)]
        public string SolutionName { get; set; }

        [Parameter(Mandatory = false)]
        public int ConnectionTimeOutMinutes { get; set; } = 10;

        [Parameter(Mandatory = false)]
        public bool Managed { get; set; }

        [Parameter(Mandatory = false)]
        public string SolutionDir { get; set; } = Directory.GetCurrentDirectory();

        protected override void ProcessRecord()
        {
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
                solutionTool = new SolutionTool(crmParameter);

                var stopwatch = Stopwatch.StartNew();
                var taskInstantiateSolution = Task.Run(() => solutionTool = new SolutionTool(crmParameter));

                while (!taskInstantiateSolution.IsCompleted)
                {
                    WriteVerbose(".");
                    Thread.Sleep(2000);
                }

                if (taskInstantiateSolution.IsFaulted)
                {
                    foreach (var innerException in taskInstantiateSolution.Exception.InnerExceptions)
                    {
                        WriteObject($"ERROR - {innerException.Message}");
                    }

                    return;
                }

                var elapsed = $"{stopwatch.Elapsed.Minutes}min {stopwatch.Elapsed.Seconds}s";
                WriteObject($"Done... [{elapsed}]{Environment.NewLine}");              

                /* Export Solutions */
                if (!Directory.Exists(SolutionDir)) Directory.CreateDirectory(SolutionDir);
                var solutionNames = SolutionName.Split(';');
                foreach (var sol in solutionNames)
                {

                    stopwatch = Stopwatch.StartNew();
                    var fileName = SolutionDir + sol + (Managed ? "_managed" : "") + ".zip";

                    WriteObject($"Downloading '{sol.ToUpper()}' solution ");
                    var solutionBinaryTask = solutionTool.ExportSolutionAsync(sol, Managed);

                    while (!solutionBinaryTask.IsCompleted)
                    {
                        WriteVerbose(".");
                        Thread.Sleep(3000);
                    }

                    if (solutionBinaryTask.IsFaulted)
                    {
                        foreach (var innerException in solutionBinaryTask.Exception.InnerExceptions)
                        {
                            WriteObject($"ERROR - {innerException.Message}");
                        }
                    }
                    else
                    {
                        WriteObject($"Saving file {fileName.ToUpper()} ");
                        elapsed = $"{stopwatch.Elapsed.Minutes}min {stopwatch.Elapsed.Seconds}s";
                        File.WriteAllBytes(fileName, solutionBinaryTask.Result);
                        WriteObject($"Done... [{elapsed}]{Environment.NewLine}");

                        if (crmParameter.ExtractSolutionContent)
                        {
                            var folderToUnzip = solutionTool.ExtractSolution(fileName, crmParameter.ExecutionDirectory, null);

                            WriteObject($"Solution is extracted to {folderToUnzip}");
                        }
                    }

                }

            }
            catch (Exception ex)
            {
                WriteObject(ex.Message);
                WriteObject(ex.StackTrace);

                if (ex.InnerException != null)
                {
                    WriteObject(ex.InnerException.Message);
                    WriteObject(ex.InnerException.StackTrace);
                }

            }
        }
    }
}
