using System;
using System.Diagnostics;
using System.Management.Automation;
using System.Threading;
using Veritec.Dynamics.CI.Common;

namespace Veritec.Dynamics.CI.PowerShell
{
    [Cmdlet("Import", "DynamicsData")]
    public class ImportDynamicsData : Cmdlet
    {
        [Parameter(Mandatory = true)]
        public string ConnectionString { get; set; }

        [Parameter(Mandatory = true)]
        public string TransformFiles { get; set; }

        [Parameter(Mandatory = false)]
        public string InputDataPath { get; set; }

        [Parameter(Mandatory = false)]
        public bool PublishAllCustomizations { get; set; } = false;

        protected override void ProcessRecord()
        {
            WriteObject("Transform File(s): " + TransformFiles);
            WriteObject("Input Data Path: " + InputDataPath);

            var crmParameter = new CrmParameter(ConnectionString)
            {
                ExportDirectoryName = InputDataPath
            };

            ExecuteImportData(crmParameter, crmParameter.GetSourceDataDirectory(), TransformFiles);

            if (PublishAllCustomizations)
            {
                /* some reference data needs the solution to be published for the changes to be applied
                 * eg systemforms enitity */
                PublishSolutions(crmParameter);
            }
        }

        protected virtual void ExecuteImportData(CrmParameter crmParameter, string dataDir, string transformFileNames)
        {
            WriteObject($"Connecting ({crmParameter.GetConnectionStringObfuscated()})");
            var dataImport = new DataImport(crmParameter, transformFileNames);

            dataImport.Logger += Feedback_Received;

            WriteObject("Reading FetchXML data from disk");
            dataImport.ReadFetchXmlQueryResultsFromDiskAndImport(dataDir, null);
        }

        protected void PublishSolutions(CrmParameter crmParameter)
        {
            var solutionTool = new SolutionTool(crmParameter);

            var stopwatch = Stopwatch.StartNew();
            WriteObject($"Publishing ");

            /* publish solution */
            var publishTask = solutionTool.PublishAsync();

            while (!publishTask.IsCompleted)
            {
                WriteVerbose(".");
                Thread.Sleep(3000);
            }

            var elapsed = $"{stopwatch.Elapsed.Minutes}min {stopwatch.Elapsed.Seconds}s";
            WriteObject($"Done... [{elapsed}]{Environment.NewLine}");
        }

        private void Feedback_Received(object sender, string e)
        {
            if (e.IndexOf("ERROR!", StringComparison.OrdinalIgnoreCase) != -1)
            {
                WriteObject(e);
            }
            else
            {
                WriteObject(e);
            }
        }
    }
}
