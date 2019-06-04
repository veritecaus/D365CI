using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace Veritec.Dynamics.CI.Common
{
    public class SolutionTool : CiBase
    {
        public event EventHandler<string> MessageLogger;

        public SolutionTool(string dynamicsConnectionString, int timeoutMinutes) : base(dynamicsConnectionString, timeoutMinutes)
        {
        }

        public SolutionTool(CrmParameter crmParameter) : base(crmParameter)
        {

        }

        public async Task ImportSolutionAsync(string solutionName)
        {
            var taskImportSolution = Task.Run(() => Import(solutionName));

            await taskImportSolution;
        }

        public async Task<byte[]> ExportSolutionAsync(string solutionName, bool exportManaged)
        {
            var taskExportSolution = Task.Run(() => ExportSolution(solutionName, exportManaged));
            var taskResult = await taskExportSolution;

            return taskResult;
        }

        public byte[] ExportSolution(string solutionName, bool exportManaged)
        {
            var exportSolutionRequest = new ExportSolutionRequest
            {
                Managed = exportManaged,
                SolutionName = solutionName
            };

            var exportSolutionResponse =
                (ExportSolutionResponse)OrganizationService.Execute(exportSolutionRequest);

            var exportXml = exportSolutionResponse.ExportSolutionFile;

            return exportXml;
        }

        public Dictionary<string, Entity> GetSolutionsInTarget(string[] solutionToBeImported)
        {
            // Check whether it already exists
            var querySolution = new QueryExpression
            {
                EntityName = "solution",
                ColumnSet = new ColumnSet(true),
                Criteria = new FilterExpression()
            };

            // retrieve only those solution info that we're importing
            querySolution.Criteria.FilterOperator = LogicalOperator.Or;
            foreach (var solutionName in solutionToBeImported)
            {
                querySolution.Criteria.AddCondition(new ConditionExpression("uniquename", ConditionOperator.Equal, solutionName));
            }

            // Create the solution if it does not already exist.
            var solutionsInTarget = OrganizationService.RetrieveMultiple(querySolution);
            var solutionsInfo = new Dictionary<string, Entity>();
            if (solutionsInTarget != null && solutionsInTarget.Entities.Count > 0)
            {
                foreach (var curSolution in solutionsInTarget.Entities)
                {
                    solutionsInfo.Add(curSolution.GetAttributeValue<string>("uniquename").ToLower(), curSolution);
                }
            }

            return solutionsInfo;
        }

        public bool DoesSolutionAlreadyExistInTarget(Dictionary<string, Entity> solutionsInTarget, Entity solutionFromSource)
        {
            var solutionName = solutionFromSource.GetAttributeValue<string>("uniquename").ToLower();
            if (!solutionsInTarget.ContainsKey(solutionName))
                return false;

            var targetsolution = solutionsInTarget[solutionName];

            if (targetsolution != null)
            {
                if (targetsolution.GetAttributeValue<string>("uniquename").Equals(solutionName, StringComparison.OrdinalIgnoreCase))
                {
                    //solution exists in target - now check the verion
                    var targetVersion = targetsolution.GetAttributeValue<string>("version");
                    var targetVersionValue = GetVersionValue(targetVersion);

                    var sourceVersion = solutionFromSource.GetAttributeValue<string>("version");
                    var sourceVersionValue = GetVersionValue(sourceVersion);

                    if (targetVersionValue >= sourceVersionValue)
                    {
                        return true; //oops....the solution with same or higher version has already been imported to target
                    }
                }
            }
            return false;
        }

        public int GetVersionValue(string version)
        {
            var versionToken = version.Split('.');
            int versionValue = 0;
            try
            {
                for (int digit = 0; digit < versionToken.Length; digit++)
                {
                    if (versionToken.Length - digit > 3)
                        versionValue = versionValue + int.Parse(versionToken[digit]) * 1000;
                    else if (versionToken.Length - digit > 2)
                        versionValue = versionValue + int.Parse(versionToken[digit]) * 100;
                    else if (versionToken.Length - digit > 1)
                        versionValue = versionValue + int.Parse(versionToken[digit]) * 10;
                    else
                        versionValue = versionValue + int.Parse(versionToken[digit]);
                }
            }
            catch { } //in case of exception eg version has string etc then return 0

            return versionValue;
        }

        public Entity GetSourceSolutionInfo(string solutionFileName)
        {
            using (ZipArchive archive = ZipFile.Open(solutionFileName, ZipArchiveMode.Read))
            {
                ZipArchiveEntry solutionEntry = archive.GetEntry("solution.xml");

                if (solutionEntry != null)
                {
                    try
                    {
                        MemoryStream readerStream = new MemoryStream();
                        var xmlStream = solutionEntry.Open();
                        xmlStream.CopyTo(readerStream);
                        readerStream.Position = 0;

                        XmlDocument solutionXml = new XmlDocument();
                        solutionXml.Load(readerStream);

                        //xmlDoc.FirstChild.FirstChild.ChildNodes

                        string solutionName = null, solutionVersion = null;

                        var solutionManifest = solutionXml.GetElementsByTagName("SolutionManifest");
                        if (solutionManifest != null && solutionManifest.Count > 0)
                        {
                            var solutionManifestNode = solutionManifest.Item(0);
                            if (solutionManifestNode == null) return null;

                            for (int index = 0; index <= solutionManifestNode.ChildNodes.Count; index++)
                            {
                                var curNode = solutionManifestNode.ChildNodes[index];
                                if (curNode != null)
                                {
                                    if (curNode.Name.ToLower().Equals("uniquename"))
                                    {
                                        solutionName = curNode.InnerText;
                                    }
                                    else if (curNode.Name.ToLower().Equals("version"))
                                    {
                                        solutionVersion = curNode.InnerText;
                                    }
                                }
                            }

                            if (solutionName != null)
                            {
                                var solutionInfo = new Entity("solution")
                                {
                                    ["uniquename"] = solutionName,
                                    ["version"] = solutionVersion
                                };
                                return solutionInfo;
                            }
                        }
                    }
                    catch //(Exception ex)
                    {
                        //don't throw - if it return null means the solution file is invalid!!!
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// source: https://community.dynamics.com/crm/b/tsgrdcrmblog/archive/2014/03/06/microsoft-dynamics-crm-2013-application-lifetime-management-part-2 
        /// </summary>
        public void Import(string solutionFilename, bool importAsync = true)
        {

            var solutionFileBytes = File.ReadAllBytes(solutionFilename);
            var importSolutionRequest = new ImportSolutionRequest
            {
                CustomizationFile = solutionFileBytes,
                PublishWorkflows = true,
                ImportJobId = new Guid(),
                OverwriteUnmanagedCustomizations = true
            };

            if (importAsync)
            {
                var asyncRequest = new ExecuteAsyncRequest
                {
                    Request = importSolutionRequest
                };
                var asyncResponse = OrganizationService.Execute(asyncRequest) as ExecuteAsyncResponse;

                Guid? asyncJobId = asyncResponse.AsyncJobId;

                var end = DateTime.Now.AddMinutes(CrmParameter.ConnectionTimeOutMinutes);
                bool finished = false;

                while (!finished)
                {
                    // Wait for 15 Seconds to prevent us overloading the server with too many requests
                    Thread.Sleep(15 * 1000);

                    if (end < DateTime.Now)
                    {
                        throw new Exception(($"Solution import has timed out: {CrmParameter.ConnectionTimeOutMinutes} minutes"));
                    }

                    Entity asyncOperation;

                    try
                    {
                        asyncOperation = OrganizationService.Retrieve("asyncoperation", asyncJobId.Value, new ColumnSet("asyncoperationid", Constant.Entity.StatusCode, "message"));

                    }
                    catch (Exception e)
                    {
                        /* Unfortunately CRM Online seems to lock up the application when importing
                         * Large Solutions, and thus it generates random errors. Mainly they are
                         * SQL Client errors, but we can't guarantee this so we just catch them and report to user
                         * then continue on. */

                        MessageLogger?.Invoke(this, $"{e.Message}");
                        continue;
                    }

                    var statusCode = asyncOperation.GetAttributeValue<OptionSetValue>(Constant.Entity.StatusCode).Value;
                    var message = asyncOperation.GetAttributeValue<string>("message");

                    // Succeeded
                    if (statusCode == 30)
                    {
                        finished = true;
                        break;
                    }
                    // Pausing // Canceling // Failed // Canceled

                    if (statusCode == 21 || statusCode == 22 ||
                        statusCode == 31 || statusCode == 32)
                    {
                        throw new Exception($"Solution Import Failed: {statusCode} {message}");
                    }

                }
            }
            else
            {
                var response = (ImportSolutionResponse)OrganizationService.Execute(importSolutionRequest);
            }
        }

        public async Task PublishAsync()
        {
            var taskPublish = Task.Run(() => Publish());
            await taskPublish;
        }

        public void Publish()
        {
            // Obtain an organization service proxy.
            // The using statement assures that the service proxy will be properly disposed.

            var publishRequest = new PublishAllXmlRequest();
            OrganizationService.Execute(publishRequest);
        }

        public string ExtractSolution(string fileName, string executionFolder, string extractionLogFile)
        {
            var folderToUnzip = Path.GetDirectoryName(fileName) + "\\" + Path.GetFileNameWithoutExtension(fileName);

            if (folderToUnzip != null)
            {
                //sample
                //ProcessStartInfo packagerParameter = new ProcessStartInfo("SolutionPackager.exe")
                //{
                //    CreateNoWindow = true,
                //    UseShellExecute = false,
                //    WorkingDirectory = "C:\\Veritec.Dynamics.CI\\Veritec.Dynamics.CI.Console.VeritecUtilityUnitTests\\bin\\Debug",
                //    Arguments = $"/action:Extract /zipfile:.\\SourcePkg\\{System.IO.Path.GetFileName(fileName)} /folder: {folderToUnzip} /packagetype:Unmanaged /errorlevel:Info /allowDelete:Yes /clobber"
                //};

                ProcessStartInfo packagerParameter = new ProcessStartInfo("SolutionPackager.exe")
                {
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    WorkingDirectory = executionFolder,
                    Arguments = $"/action:Extract /zipfile:\"{fileName}\" /folder:\"{folderToUnzip}\"" +
                        " /packagetype:Unmanaged /errorlevel:Info /allowDelete:Yes /clobber " +
                        (string.IsNullOrWhiteSpace(extractionLogFile) ? "" : "/Log:\"" + extractionLogFile + "\"")
                };

                var extractionProcess = Process.Start(packagerParameter);

                // TODO: Make the wait time a configurable setting.
                extractionProcess.WaitForExit(1800000); //1800sec
            }

            return folderToUnzip;
        }
    }
}

