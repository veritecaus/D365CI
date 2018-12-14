using System;
using System.Management.Automation;
using Veritec.Dynamics.CI.Common;

namespace Veritec.Dynamics.CI.PowerShell
{
    [Cmdlet("Export", "DynamicsData")]
    public class ExportDynamicsData : Cmdlet
    {
        [Parameter(Mandatory = true)]
        public string EncryptedPassword { get; set; }

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
            WriteObject("Starting Export");
            WriteObject("Connection String: " + ConnectionString);
            WriteObject("FetchXML File: " + FetchXmlFile);
            WriteObject("Output Data Path: " + OutputDataPath);
            WriteObject("Encrypted Password: ***********");

            var cp =
                new CrmParameter(ConnectionString, true)
                {
                    Password = CredentialTool.MakeSecurityString(EncryptedPassword),
                    ExportDirectoryName = OutputDataPath
                };

            var attributesToExcludeArray = AttributesToExclude?.Split(';');

            string dataDir = cp.GetSourceDataDirectory();

            FetchXmlFile = (string.IsNullOrWhiteSpace(FetchXmlFile) || FetchXmlFile == "NULL" ?
                @"Veritec.Dynamics.DataMigrate.FetchQuery.xml" : FetchXmlFile);

            WriteObject("Fetch Query source: " + Environment.CurrentDirectory + "\\" + FetchXmlFile);

            WriteObject("Connecting to Dynamics...");
            var dataExport = new DataExport(cp, FetchXmlFile);

            WriteObject("Fetch queries loaded");
            foreach (var fetchXml in dataExport.FetchXmlQueries)
            {
                WriteObject($"Executing Fetch Query: {fetchXml.Key}...");

                var queryResultXml = dataExport.ExecuteFetchXmlQuery(fetchXml.Key, fetchXml.Value, attributesToExcludeArray);

                WriteObject($"Writing Results: {dataDir}{fetchXml.Key}.xml");
                DataExport.SaveFetchXmlQueryResulttoDisk(dataDir, fetchXml.Key, queryResultXml.ToString());
            }
        }
    }
}
