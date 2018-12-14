using System;
using System.Linq;
using System.Net;
using System.Security;

namespace Veritec.Dynamics.CI.Common
{
    public class CrmParameter
    {
        public string ConnectionString;
        public bool IsOffice365;
        public NetworkCredential Credential = null;
        public string HostName;
        public string HostPort = "80";
        public string Domain;
        public string UserName;
        public SecureString Password = null;
        public string Region;
        public string OrganizationName;
        public bool UseSsl = true;
        public string ExecutionDirectory = null;
        public string ExportDirectoryName = null; //this is where the files will be exported
        public bool WriteLog = true; //by default write log
        public int ConnectionTimeOutMinutes = 10;
        public bool VerboseMode = false;
        public bool ExtractSolutionContent = false; //don't extract solution content by default
        public bool VerifyDataImport = false; //don't verify data import by default

        // If true, then the utility will overwrite the solution from Source into Target regardless of the version of the solution in Target environment.
        public bool ForceOverwriteSolution = true;

        public TimeSpan GetConnectionTimeOut()
        {
            return new TimeSpan(0, ConnectionTimeOutMinutes, 0);
        }

        public CrmParameter(string connectionString, bool populateParameter = false)
        {
            // either connection string or individual parameter
            if (populateParameter)
                PopulateParameter(connectionString);
            else
                ConnectionString = connectionString;
        }

        public void PopulateParameter(string connectionString)
        {
            var parameters = connectionString.Split(';').Select(s => s.Split('=')).ToDictionary(s => s.First().ToLower(), s => s.Last());

            var fields = GetType().GetFields().ToDictionary(s => s.Name.ToLower(), s => s);

            foreach (var curParameter in parameters)
            {
                var propertyName = curParameter.Key.Trim();
                if (fields.ContainsKey(propertyName))
                {
                    var curField = fields[propertyName];
                    curField.SetValue(this, Convert.ChangeType(curParameter.Value, curField.FieldType));
                }
            }
        }

        public string GetSourcePkgDirectory()
        {
            var pkgDirName = (ExportDirectoryName == null ? @"\SourcePkg\" : "\\" + ExportDirectoryName + "\\");
            if (string.IsNullOrWhiteSpace(ExecutionDirectory))
                return Environment.CurrentDirectory + pkgDirName;
            return ExecutionDirectory + pkgDirName;
        }

        public string GetSourceDataDirectory()
        {
            var pkgDirName = (ExportDirectoryName == null ? @"\SourceData\" : "\\" + ExportDirectoryName + "\\");
            if (string.IsNullOrWhiteSpace(ExecutionDirectory))
                return Environment.CurrentDirectory + pkgDirName;
            return ExecutionDirectory + pkgDirName;
        }

        public string GetLogsDirectory()
        {
            if (string.IsNullOrWhiteSpace(ExecutionDirectory))
                return Environment.CurrentDirectory + "\\Logs\\";
            return ExecutionDirectory + "\\Logs\\";
        }

        public override string ToString()
        {
            if (string.IsNullOrWhiteSpace(ConnectionString))
            {
                var userOrg = ", User name: " + UserName + ", Password: *******, Organization: " + OrganizationName;

                if (IsOffice365)
                    return "Region: " + Region + userOrg;
                if (Credential != null)
                    return "Host: " + HostName + ", Port: " + HostPort + ", Domain: " + Credential.Domain + userOrg;
                return "Host: " + HostName + ", Port: " + HostPort + ", Domain: " + Domain + userOrg;
            }

            return CiBase.HidePasswordInConnection(ConnectionString);
        }
    }
}
