using System;
using System.Management.Automation;
using Veritec.Dynamics.CI.Common;

namespace Veritec.Dynamics.CI.PowerShell
{
    [Cmdlet("Export", "DynamicsData")]
    public class ExportDynamicsData : Cmdlet
    {
        [Parameter(Mandatory = true)]
        public string ConnectionString { get; set; }

        [Parameter(Mandatory = true)]
        public string FetchXmlFile { get; set; }

        [Parameter(Mandatory = false)]
        public string OutputDataPath { get; set; }

        [Parameter(Mandatory = false)]
        public string AttributesToExclude { get; set; } = "languagecode;createdon;createdby;modifiedon;modifiedby;owningbusinessunit;owninguser;owneridtype;importsequencenumber;overriddencreatedon;timezoneruleversionnumber;operatorparam;utcconversiontimezonecode;versionnumber;customertypecode;matchingentitymatchcodetable;baseentitymatchcodetable;slaidunique;slaitemidunique";

        protected override void ProcessRecord()
        {
            WriteObject("");
            WriteObject($"FetchXML File: {FetchXmlFile}");
            WriteObject($"Output Data Path: {OutputDataPath}{Environment.NewLine}");

            var crmParameter =
                new CrmParameter(ConnectionString)
                {
                    ExportDirectoryName = OutputDataPath
                };

            var attributesToExcludeArray = AttributesToExclude?.Split(';');

            string dataDir = crmParameter.GetSourceDataDirectory();

            FetchXmlFile = (string.IsNullOrWhiteSpace(FetchXmlFile) || FetchXmlFile == "NULL" ?
                @"Veritec.Dynamics.DataMigrate.FetchQuery.xml" : FetchXmlFile);

            WriteObject($"Connecting ({crmParameter.GetConnectionStringObfuscated()})");
            var dataExport = new DataExport(crmParameter, FetchXmlFile);

            WriteObject($"Fetch queries loaded {Environment.NewLine}");
            foreach (var fetchXml in dataExport.FetchXmlQueries)
            {
                WriteObject($"Executing Fetch Query: {fetchXml.Key}...");

                var queryResultXml = dataExport.ExecuteFetchXmlQuery(fetchXml.Key, fetchXml.Value, attributesToExcludeArray);

                WriteObject($"Writing Results: {dataDir}{fetchXml.Key}.xml {Environment.NewLine}");
                DataExport.SaveFetchXmlQueryResulttoDisk(dataDir, fetchXml.Key, queryResultXml.ToString());
            }
        }
    }
}
