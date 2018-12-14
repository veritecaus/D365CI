using System;
using System.Threading;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Veritec.Dynamics.CI.Common
{
    public class WFandSlaMigrateUtility
    {
        private readonly IOrganizationService _organizationService;

        public WFandSlaMigrateUtility(IOrganizationService organizationService)
        {
            _organizationService = organizationService;
        }

        public void DeactivateWorkflowBeforeImportingWorkflow(Entity targetEntity)
        {
            if (targetEntity.LogicalName.Equals("workflow", StringComparison.OrdinalIgnoreCase))
            {
                if (!targetEntity.Contains(Constant.Entity.StateCode) || (targetEntity.Contains(Constant.Entity.StateCode) &&
                    targetEntity.GetAttributeValue<OptionSetValue>(Constant.Entity.StateCode).Value == 1))  // Active in Target
                {
                    var deactivateRequest = new SetStateRequest
                    {
                        EntityMoniker = new EntityReference(targetEntity.LogicalName, targetEntity.Id),
                        State = new OptionSetValue(0), //draft
                        Status = new OptionSetValue(1) //draft
                    };
                    _organizationService.Execute(deactivateRequest);

                    //sleep to let the process finish before continuing ;)
                    Thread.Sleep(2000);
                }
            }
        }

        public void DeactivateSlaBeforeImportSla(Entity curEntity)
        {
            if (curEntity.LogicalName.Equals("sla", StringComparison.OrdinalIgnoreCase))
            {
                //if the SLA is active in source, then deactivate 
                var curSla =
                    new Entity(curEntity.LogicalName, curEntity.Id)
                    {
                        [Constant.Entity.StateCode] = new OptionSetValue(0), //inactive
                        [Constant.Entity.StatusCode] = new OptionSetValue(1) //draft
                    };
                
                _organizationService.Update(curSla);
            }
        }

        public void DeactivateSlaRelatedToWorkflow(Entity curWorkflowEntity)
        {
            if (curWorkflowEntity == null || !curWorkflowEntity.LogicalName.Equals("workflow", StringComparison.OrdinalIgnoreCase)) return;

            var hasDeactivatedSla = false;
            //SLA and SLA Item can have one or more Workflow!

            //get the Active SLA of sla-item of the workflow and DEACTIVATE the SLA
            #region Prepare query to retrieve related SLA of the SLA item for the current workflow. Then deactivate the SLA

            var slaofSlaItemOfWorkflow = new QueryExpression("sla")
            {
                ColumnSet = new ColumnSet("slaid"),
                Distinct = true,
                NoLock = true,
                Criteria =
                    {
                        FilterOperator = LogicalOperator.And,
                        Conditions =
                        {
                            new ConditionExpression(Constant.Entity.StateCode, ConditionOperator.Equal, 1) //For SLA active code = 1!
                        }
                    },

                LinkEntities =
                    {
                        new LinkEntity
                        {
                            LinkFromEntityName = "sla",
                            LinkFromAttributeName = "slaid",
                            LinkToEntityName = "slaitem",
                            LinkToAttributeName = "slaid",
                            JoinOperator = JoinOperator.Inner,
                            LinkCriteria =
                            {
                                FilterOperator = LogicalOperator.And,
                                Conditions =
                                {
                                    new ConditionExpression("workflowid", ConditionOperator.Equal, curWorkflowEntity.Id)
                                }
                            }
                        }
                    }
            };

            //Create the solution if it does not already exist.
            var slaEntities = _organizationService.RetrieveMultiple(slaofSlaItemOfWorkflow);

            foreach (var curSla in slaEntities.Entities)
            {
                curSla[Constant.Entity.StateCode] = new OptionSetValue(0); //inactive
                curSla[Constant.Entity.StatusCode] = new OptionSetValue(1); //draft

                _organizationService.Update(curSla);
                hasDeactivatedSla = true;
            }

            #endregion

            //get the Active SLA of the workflow and DEACTIVATE the SLA
            QueryExpression slaOfWorkflow = new QueryExpression("sla")
            {
                ColumnSet = new ColumnSet("slaid"),
                Distinct = true,
                NoLock = true,
                Criteria =
                        {
                           FilterOperator = LogicalOperator.And,
                           Conditions =
                           {
                              new ConditionExpression(Constant.Entity.StateCode, ConditionOperator.Equal, 1), //For SLA active code = 1!
                              new ConditionExpression("workflowid", ConditionOperator.Equal, curWorkflowEntity.Id)
                           }
                        }
            };

            //Create the solution if it does not already exist.
            slaEntities = _organizationService.RetrieveMultiple(slaOfWorkflow);

            foreach (var curSla in slaEntities.Entities)
            {
                curSla[Constant.Entity.StateCode] = new OptionSetValue(0); //inactive
                curSla[Constant.Entity.StatusCode] = new OptionSetValue(1); //draft

                _organizationService.Update(curSla);
                hasDeactivatedSla = true;
            }

            //sleep to let the process finish (deactivate SLA that trigger to deactivated its WF) before continuing ;)
            if (hasDeactivatedSla)
            {
                int sleepCount = 0;
                var curWorkflowToCheckStatus = _organizationService.Retrieve("workflow", curWorkflowEntity.Id, new ColumnSet(Constant.Entity.StateCode));

                //give some extra time when the statecode is still active (1)
                while (curWorkflowToCheckStatus.GetAttributeValue<OptionSetValue>(Constant.Entity.StateCode) != null &&
                    curWorkflowToCheckStatus.GetAttributeValue<OptionSetValue>(Constant.Entity.StateCode).Value == 1)
                {
                    Thread.Sleep(2000);
                    sleepCount++;

                    //can't wait too long
                    if (sleepCount > 3)
                        break;
                    curWorkflowToCheckStatus = _organizationService.Retrieve("workflow", curWorkflowEntity.Id, new ColumnSet(Constant.Entity.StateCode));
                }
            }
        }

        public void ActivateSlaAfterImport(Entity entityBeforeStatusCodeChange, Entity curEntity)
        {
            if (entityBeforeStatusCodeChange.LogicalName.Equals("sla", StringComparison.OrdinalIgnoreCase))
            {
                if (entityBeforeStatusCodeChange.Contains(Constant.Entity.StatusCode) &&
                    entityBeforeStatusCodeChange.GetAttributeValue<OptionSetValue>(Constant.Entity.StatusCode).Value == 2) // "Active"
                {
                    var curSla = new Entity(entityBeforeStatusCodeChange.LogicalName, entityBeforeStatusCodeChange.Id);
                    curSla[Constant.Entity.StateCode] = entityBeforeStatusCodeChange.GetAttributeValue<OptionSetValue>(Constant.Entity.StateCode); // active
                    curSla[Constant.Entity.StatusCode] = entityBeforeStatusCodeChange.GetAttributeValue<OptionSetValue>(Constant.Entity.StatusCode); // active

                    _organizationService.Update(curSla);

                    // update the curEntity to publish state in order to get it matched during VERIFICATION
                    curEntity[Constant.Entity.StateCode] = entityBeforeStatusCodeChange.GetAttributeValue<OptionSetValue>(Constant.Entity.StateCode); // active
                    curEntity[Constant.Entity.StatusCode] = entityBeforeStatusCodeChange.GetAttributeValue<OptionSetValue>(Constant.Entity.StatusCode); // published
                }
            }
        }
    }
}
