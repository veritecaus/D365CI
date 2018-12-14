using System;
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

                var elapsed = $"{stopwatch.Elapsed.Minutes}min {stopwatch.Elapsed.Seconds}s";
                WriteObject($" [{elapsed}] {Environment.NewLine}");

                #endregion              

                #region Import Solutions
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
                    WriteObject($" [{elapsed}] {Environment.NewLine}");

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
                    WriteObject($" [{elapsed}] {Environment.NewLine}");
                    
                }
                #endregion
            }
            catch (Exception ex)
            {
                var errorRecord = new ErrorRecord(new Exception("Solution Import Failed",ex.InnerException), "", ErrorCategory.InvalidResult, null);
                ThrowTerminatingError(errorRecord);
            }
        }
    }
}
