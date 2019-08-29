using System;
using System.Management.Automation;
using Veritec.Dynamics.CI.Common;
namespace Veritec.Dynamics.CI.PowerShell
{
    /// <summary>
    /// Set enable/disable status of a record
    /// </summary>
    [Cmdlet("Set", "RecordStatus")]
    [OutputType(typeof(Boolean))]
    public class SetRecordStatus : Cmdlet
    {
        [Parameter(Mandatory = true)]
        public string ConnectionString { get; set; }

        [Parameter(Mandatory = false)]
        public int ConnectionTimeOutMinutes { get; set; } = 10;

        [Parameter(Mandatory = true)]
        public string TargetRecordsFetchXML { get; set; }

        [Parameter(Mandatory = true)]
        public bool SetEnabled { get; set; }

        protected override void ProcessRecord()
        {
            var crmParameter = new CrmParameter(ConnectionString)
            {
                ConnectionTimeOutMinutes = ConnectionTimeOutMinutes
            };

            WriteObject($"Connecting ({crmParameter.GetConnectionStringObfuscated()}){Environment.NewLine}");

            var recordManager = new RecordManager(crmParameter);
            recordManager.Logger += ReceiveMessage;

            try
            {
                if (SetEnabled)
                {
                    recordManager.SetStatus(TargetRecordsFetchXML, RecordManager.RecordStatus.Enable);
                }
                else
                {
                    recordManager.SetStatus(TargetRecordsFetchXML, RecordManager.RecordStatus.Disable);
                }
            }
            catch (Exception ex)
            {
                WriteError(new ErrorRecord(ex, "Error setting record status", ErrorCategory.InvalidOperation, TargetRecordsFetchXML));
            }
        }

        private void ReceiveMessage(object sender, string e)
        {
            WriteObject(e);
        }
    }
}