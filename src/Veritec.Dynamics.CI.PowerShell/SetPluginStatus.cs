using System;
using System.Management.Automation;
using Veritec.Dynamics.CI.Common;
namespace Veritec.Dynamics.CI.PowerShell
{
    /// <summary>
    /// Set enable/disable status of a Plugin
    /// </summary>
    [Cmdlet("Set", "PluginStatus")]
    [OutputType(typeof(Boolean))]
    public class SetPluginStatus : Cmdlet
    {
        [Parameter(Mandatory = true)]
        public string ConnectionString { get; set; }

        [Parameter(Mandatory = false)]
        public int ConnectionTimeOutMinutes { get; set; } = 10;
        
        /// <summary>
        /// This is a semicolon delimited list of plugin names from the "SDK Message Processing Steps"
        /// section of the solution manager.
        /// </summary>
        [Parameter(Mandatory = true)]
        public string PluginStepNames { get; set; }

        [Parameter(Mandatory = true)]
        public bool setEnabled { get; set; }

        protected override void ProcessRecord()
        {
            var crmParameter = new CrmParameter(ConnectionString)
            {
                ConnectionTimeOutMinutes = ConnectionTimeOutMinutes
            };

            WriteObject($"Connecting ({crmParameter.GetConnectionStringObfuscated()}){Environment.NewLine}");

            var plugin = new PluginManager(crmParameter);
            plugin.Logger += ReceiveMessage;

            try
            {
                if (setEnabled)
                {
                    plugin.SetStatus(PluginStepNames, PluginManager.PluginStatus.Enable);
                }
                else
                {
                    plugin.SetStatus(PluginStepNames, PluginManager.PluginStatus.Disable);
                }
            }
            catch (Exception ex)
            {
                WriteError(new ErrorRecord(ex, "Error setting pugin status", ErrorCategory.InvalidOperation, PluginStepNames));
            }
        }

        private void ReceiveMessage(object sender, string e)
        {
            WriteObject(e);
        }
    }
}