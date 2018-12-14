using System;
using System.Diagnostics;
using System.Net;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;

namespace Veritec.Dynamics.CI.Common
{
    public class CiBase
    {
        public IOrganizationService OrganizationService;
        public CrmParameter CrmParameter;

        public CiBase(string dynamicsConnectionString) 
        {
            var crmParameter = new CrmParameter(dynamicsConnectionString);

            ConnectToCrm(crmParameter);
        }

        public CiBase(string dynamicsConnectionString, int timeoutMinutes)
        {
            var crmParameter = new CrmParameter(dynamicsConnectionString)
            {
                ConnectionTimeOutMinutes = timeoutMinutes
            };

            ConnectToCrm(crmParameter);
        }

        public CiBase(CrmParameter crmParameter)
        {
            ConnectToCrm(crmParameter);
        }

        public void ConnectToCrm(CrmParameter crmParameter)
        {
            CrmServiceClient conn;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            if (string.IsNullOrEmpty(crmParameter.ConnectionString) && string.IsNullOrEmpty(crmParameter.UserName)) 
                throw new Exception($"Connection parameter is missing: {crmParameter}");

            CrmParameter = crmParameter;

            if (!string.IsNullOrWhiteSpace(crmParameter.ConnectionString))
                conn = new CrmServiceClient(crmParameter.ConnectionString);
            else if (crmParameter.IsOffice365)
                conn = new CrmServiceClient(crmParameter.UserName, crmParameter.Password, crmParameter.Region, crmParameter.OrganizationName, true, crmParameter.UseSsl, null, true);
            else if (crmParameter.Credential != null)
                conn = new CrmServiceClient(crmParameter.Credential, crmParameter.HostName, crmParameter.HostPort, crmParameter.OrganizationName, true, crmParameter.UseSsl);
            else 
                conn = new CrmServiceClient(crmParameter.UserName, crmParameter.Password, crmParameter.Domain, null, crmParameter.HostName, crmParameter.HostPort, crmParameter.OrganizationName, true, crmParameter.UseSsl, null);

            if (conn?.OrganizationServiceProxy == null) throw new Exception($"Could not connect to Dynamics! Please check connection parameter: " + crmParameter);
            
            conn.OrganizationServiceProxy.Timeout = crmParameter.GetConnectionTimeOut();


            try
            {
                OrganizationService = (IOrganizationService)conn.OrganizationWebProxyClient != null ? conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
            }
            catch (Exception e)
            {
                throw new Exception("Error establishing connection to Dynamics. Inner exception: " + e.InnerException);
            }

            if (OrganizationService == null) throw new Exception("Error establishing connection to Dynamics");
        }
        
        public void LaunchCommandLineApp(string targetExe, string[] args)
        {
            var startInfo = new ProcessStartInfo
            {
                CreateNoWindow = false,
                UseShellExecute = false,
                FileName = targetExe,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            var arguments = "";

            foreach (var arg in args)
            {
                arguments += " " + arg;
            }

            startInfo.Arguments = arguments;

            try
            {
                // Start the process with the info we specified.
                // Call WaitForExit and then the using statement will close.
                var exeProcess = Process.Start(startInfo);
            }
            catch
            {
                // Log error.
            }
        }

        public static string HidePasswordInConnection(string connectionString)
        {
            if (string.IsNullOrWhiteSpace(connectionString)) return connectionString;

            var connectionParams = connectionString.Split(';');
            for (var i = 0; i < connectionParams.Length; i++)
            {
                if (!string.IsNullOrWhiteSpace(connectionParams[i]))
                {
                    var pwdPos = connectionParams[i].Trim().IndexOf("Password", StringComparison.OrdinalIgnoreCase);
                    if (pwdPos == -1)
                    {
                        pwdPos = connectionParams[i].Trim().IndexOf("Pwd", StringComparison.OrdinalIgnoreCase);
                        if (pwdPos != -1)
                            connectionParams[i] = "Pwd=*******";
                    }
                    else
                    {
                        connectionParams[i] = "Password=*******";
                    }
                }
            }

            return string.Join(";", connectionParams);
        }

    }
}
