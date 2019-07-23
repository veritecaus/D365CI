using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
namespace Veritec.Dynamics.CI.Common
{
    /// <summary>
    /// Manipulate Plugin Status
    /// </summary>
    public class PluginManager : CiBase
    {
        public event EventHandler<string> Logger;

        public PluginManager(CrmParameter crmParameter) : base(crmParameter)
        {
        }

        public enum PluginStatus { Enable, Disable };

        public void SetStatus(string pluginStepNames, PluginStatus pluginStatus)
        {
            if (string.IsNullOrWhiteSpace(pluginStepNames))
            {
                return;
            }

            var pluginStepNamesArray = pluginStepNames.Split(';');
            var sdkPluginStepIds = GetPluginSteps(OrganizationService, pluginStepNamesArray);

            if (sdkPluginStepIds.Count == 0)
            {
                Logger?.Invoke(this, $"Warning - No plugins found with the PluginStepNames supplied");
                return;
            }

            foreach (var curSdkSep in sdkPluginStepIds)
            {
                // enable or disable the plugins
                SetPluginStepStatus(base.OrganizationService, sdkPluginStepIds, pluginStatus);
                Logger?.Invoke(this, $"Plugin '{curSdkSep.Value}' set to {pluginStatus}");
            }

        }

        private Dictionary<Guid, string> GetPluginSteps(IOrganizationService orgService, string[] pluginSdkSteps)
        {
            var pluginSdkStepQuery = new QueryExpression("sdkmessageprocessingstep")
            {
                ColumnSet = new ColumnSet(new string[] { "sdkmessageprocessingstepid", "name" }),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("name", ConditionOperator.In, pluginSdkSteps.ToArray())
                    }
                }
            };
            var sdkPluginStepIds = new Dictionary<Guid, string>();
            var sdkPluginStepEntities = orgService.RetrieveMultiple(pluginSdkStepQuery);

            if (sdkPluginStepEntities != null && sdkPluginStepEntities.Entities.Count > 0)
            {
                foreach (var curEntity in sdkPluginStepEntities.Entities)
                {
                    sdkPluginStepIds.Add(curEntity.Id, curEntity.GetAttributeValue<string>("name"));
                }
            }

            return sdkPluginStepIds;
        }

        private void SetPluginStepStatus(IOrganizationService orgService,
            Dictionary<Guid, string> sdkPluginStepIds, PluginStatus pluginStatus)
        {
            foreach (var curSdkSep in sdkPluginStepIds)
            {
                try
                {
                    // Inactive: stateCode = 1 and statusCode = 2
                    // Active: stateCode = 0 and statusCode = 1
                    orgService.Execute(new SetStateRequest
                    {
                        EntityMoniker = new EntityReference("sdkmessageprocessingstep", curSdkSep.Key),
                        State = new OptionSetValue(pluginStatus == PluginStatus.Disable ? 1 : 0),
                        Status = new OptionSetValue(pluginStatus == PluginStatus.Disable ? 2 : 1)
                    });
                }
                catch (Exception ex)
                {
                    throw new Exception($"Failed for {pluginStatus} of plugin step {curSdkSep.Value}", ex);
                }
            }
        }
    }
}
