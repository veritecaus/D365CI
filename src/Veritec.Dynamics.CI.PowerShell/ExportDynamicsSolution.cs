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
        #region Input Paramaters

        [Parameter(Mandatory = true)]
        public string EncryptedPassword { get; set; }

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

        #endregion

        protected override void ProcessRecord()
        {
            try
            {
                #region Initialise Paramaters

                var cp = new CrmParameter(ConnectionString, true)
                {
                    Password = CredentialTool.MakeSecurityString(EncryptedPassword)
                };

                var crmParameter = new CrmParameter(ConnectionString, true)
                {
                    Password = CredentialTool.MakeSecurityString(EncryptedPassword),
                    ConnectionTimeOutMinutes = ConnectionTimeOutMinutes
                };

                SolutionDir = SolutionDir + @"\";
                crmParameter.ExecutionDirectory = SolutionDir;
                #endregion

                #region Connect to Dynamics
                WriteObject($"Connecting ({ConnectionString}) ");

                SolutionTool solutionTool = null;

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
                WriteObject($" [{elapsed}] {Environment.NewLine}");

                #endregion              

                #region Export Solutions
                if (!Directory.Exists(SolutionDir)) Directory.CreateDirectory(SolutionDir);
                var solutionNames = SolutionName.Split(';');
                foreach (var sol in solutionNames)
                {

                    stopwatch = Stopwatch.StartNew();
                    var fileName = SolutionDir + sol + (Managed ? "_managed" : "") + ".zip";

                    WriteObject(Environment.NewLine);
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
                        elapsed = $"{stopwatch.Elapsed.Minutes}min {stopwatch.Elapsed.Seconds}s";
                        WriteObject($" [{elapsed}] {Environment.NewLine}");

                        WriteObject($"Saving file {fileName.ToUpper()} ");
                        File.WriteAllBytes(fileName, solutionBinaryTask.Result);
                        WriteObject($"... [done]{Environment.NewLine}");

                        if (crmParameter.ExtractSolutionContent)
                        {
                            var folderToUnzip = solutionTool.ExtractSolution(fileName, crmParameter.ExecutionDirectory, null);

                            WriteObject($"Solution is extracted to {folderToUnzip}");
                        }
                    }

                }

                #endregion
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
