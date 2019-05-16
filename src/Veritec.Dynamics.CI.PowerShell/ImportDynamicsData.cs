using System;
using System.Management.Automation;
using Veritec.Dynamics.CI.Common;

namespace Veritec.Dynamics.CI.PowerShell
{
    [Cmdlet("Import", "DynamicsData")]
    public class ImportDynamicsData : Cmdlet
    {
        [Parameter(Mandatory = true)]
        public string ConnectionString { get; set; }

        [Parameter(Mandatory = true)]
        public string TransformFile { get; set; }

        [Parameter(Mandatory = false)]
        public string InputDataPath { get; set; }

        protected override void ProcessRecord()
        {
            WriteObject("Transform File: " + TransformFile);
            WriteObject("Input Data Path: " + InputDataPath);

            var crmParameter = new CrmParameter(ConnectionString)
            {
                ExportDirectoryName = InputDataPath
            };

            ExecuteImportData(crmParameter, crmParameter.GetSourceDataDirectory(), TransformFile);
        }

        protected virtual void ExecuteImportData(CrmParameter crmParameter, string dataDir, string transformFileName)
        {
            transformFileName = (string.IsNullOrWhiteSpace(transformFileName) || transformFileName == "NULL" ?
                @"Veritec.Dynamics.DataMigrate.Transform.xml" : transformFileName);

            WriteObject("Transfrom File: " + transformFileName);

            WriteObject($"Connecting ({crmParameter.GetConnectionStringObfuscated()})");
            var dataImport = new DataImport(crmParameter, transformFileName);

            dataImport.Logger += Feedback_Received;

            WriteObject("Reading FetchXML data from disk");
            dataImport.ReadFetchXmlQueryResultsFromDiskAndImport(dataDir, null);
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
