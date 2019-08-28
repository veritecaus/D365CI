using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Veritec.Dynamics.CI.Common
{
    /// <summary>
    /// Manipulate Record Status
    /// </summary>
    public class RecordManager : CiBase
    {
        public event EventHandler<string> Logger;

        public RecordManager(CrmParameter crmParameter) : base(crmParameter)
        {
        }

        public enum RecordStatus { Enable, Disable };

        public void SetStatus(string fetchXML, RecordStatus recordStatus)
        {
            if (string.IsNullOrWhiteSpace(fetchXML))
            {
                return;
            }
            var records = GetRecords(OrganizationService, fetchXML);

            if (records.Entities.Count == 0)
            {
                Logger?.Invoke(this, $"Warning - No records that can be disabled found for the FetchXML supplied");
                return;
            }

            foreach (var entity in records.Entities)
            {
                if (stateChanged(entity, recordStatus))
                {
                    SetRecordStatus(OrganizationService, entity, recordStatus);
                    Logger?.Invoke(this, $"Record '{entity.LogicalName} {entity.Id}' set to {recordStatus}");
                }
                else
                {
                    Logger?.Invoke(this, $"Record '{entity.LogicalName} {entity.Id}' already {recordStatus}");
                }
            }

        }

        private bool stateChanged(Entity entity, RecordStatus TargetState)
        {
            RecordStatus currentRecordState = RecordStatus.Disable;

            if (entity.GetAttributeValue<OptionSetValue>("statecode").Value == 0 &&
                entity.GetAttributeValue<OptionSetValue>("statuscode").Value == 1)
            {
                currentRecordState = RecordStatus.Enable;
            }

            return currentRecordState == TargetState ? false : true;
        }

        private EntityCollection GetRecords(IOrganizationService orgService, string fetchXML)
        {
            var entityCollectionResult = orgService.RetrieveMultiple(new FetchExpression());

            foreach (var entity in entityCollectionResult.Entities)
            {
                if ((entity.Attributes.Contains("statecode") && entity.Attributes.Contains("statuscode")) == false)
                {
                    // only include results that contain a statecode and statuscode
                    // to ensure we have a record that can be disabled or enabled
                    entityCollectionResult.Entities.Remove(entity);
                }
            }

            return entityCollectionResult;
        }

        private void SetRecordStatus(IOrganizationService orgService, Entity entity, RecordStatus recordStatus)
        {
            try
            {
                // Inactive: stateCode = 1 and statusCode = 2
                // Active: stateCode = 0 and statusCode = 1
                orgService.Execute(new SetStateRequest
                {
                    EntityMoniker = new EntityReference(entity.LogicalName, entity.Id),
                    State = new OptionSetValue(recordStatus == RecordStatus.Disable ? 1 : 0),
                    Status = new OptionSetValue(recordStatus == RecordStatus.Disable ? 2 : 1)
                });
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to {recordStatus} record '{entity.LogicalName} {entity.Id}'", ex);
            }

        }
    }
}
