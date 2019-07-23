using System;
using System.Management.Automation;
using Veritec.Dynamics.CI.Common;

namespace Veritec.Dynamics.CI.PowerShell
{
    [Cmdlet("Set", "AutoNumberSeed")]
    [OutputType(typeof(Boolean))]
    public class SetAutoNumberSeed : Cmdlet
    {
        [Parameter(Mandatory = true)]
        public string ConnectionString { get; set; }

        [Parameter(Mandatory = false)]
        public int ConnectionTimeOutMinutes { get; set; } = 10;

        [Parameter(Mandatory = true)]
        public string EntityName { get; set; }

        [Parameter(Mandatory = true)]
        public string AttributeName { get; set; }

        [Parameter(Mandatory = true)]
        public long Value { get; set; }

        [Parameter(Mandatory = false)]
        public bool Force { get; set; } = false;

        protected override void ProcessRecord()
        {
            var crmParameter = new CrmParameter(ConnectionString)
            {
                ConnectionTimeOutMinutes = ConnectionTimeOutMinutes
            };

            WriteObject($"Connecting ({crmParameter.GetConnectionStringObfuscated()}){Environment.NewLine}");

            var autoNumber = new AutoNumber(crmParameter);
            try
            {
                var recordCount = autoNumber.TargetEntityRowCount(EntityName, AttributeName);
                if (Force == true || recordCount == 0)
                {
                    /* only do this if the user forces it or there are no records in the target */
                    autoNumber.SetSeed(EntityName, AttributeName, Value, Force);
                    WriteObject($"Autonumber has been set to {Value}");
                }
                else
                {
                    WriteObject($"No action required to set the autonumber seed for {EntityName} because {recordCount} record(s) already exist.");
                }
               
            }
            catch (Exception ex)
            {
                WriteError(new ErrorRecord(ex, "Error setting autonumber", ErrorCategory.InvalidOperation, AttributeName));
            }
        }
    }
}
