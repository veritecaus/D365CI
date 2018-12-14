using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Veritec.Dynamics.CI.Common
{
    public class DataImportVerification 
    {
        private readonly List<string> _excludedColumnsToCompare;
        public DataImportVerification(string[] columnsToExcludeToCompareData)
        {
            _excludedColumnsToCompare = columnsToExcludeToCompareData.ToList();
        }

        public string VerifyDataImport(IOrganizationService organizationService, EntityCollection sourceEntityCollection)
        {
            var verificationMsg = new StringBuilder();
            if (sourceEntityCollection == null || sourceEntityCollection.Entities.Count == 0)
                return string.Empty;

            var entityName = sourceEntityCollection.Entities[0].LogicalName;
            // Get target data to verify
            var targetEntityCollection = LoadTargetDataToVerify(organizationService, sourceEntityCollection);
            if (targetEntityCollection == null || targetEntityCollection.Entities.Count == 0)
            {
                verificationMsg.AppendLine($"\r\nVERIFICATION: [ {entityName} ] has NOT been imported into target environment.");
            }
            else
            {
                // no need to compare fields between FetchXML and Target Dynamics as they are inserted/updated using data from FetchXML
                string compareMsg = CompareSourceAndTargetData(sourceEntityCollection, targetEntityCollection, "FetchXML", "Target Dynamics", true);
                if (string.IsNullOrWhiteSpace(compareMsg))
                {
                    verificationMsg.AppendLine($"\r\nVERIFICATION: All [ {entityName} ] records in FetchXML are imported into Target Dynamics :)");
                }
                else
                {
                    if (sourceEntityCollection.Entities.Count == targetEntityCollection.Entities.Count)
                    {
                        verificationMsg.AppendLine($"\r\nVERIFICATION: All [ {entityName} ] records in FetchXML are imported into Target Dynamics :)");
                    }
                    else
                    {
                        verificationMsg.AppendLine($"\r\nVERIFICATION: There are {sourceEntityCollection.Entities.Count} {entityName} records in FetchXML " +
                            $"where as Target Dynamics has {targetEntityCollection.Entities.Count} records!");
                    }
                    verificationMsg.AppendLine(compareMsg);

                    // check if data from target matches with source!
                    verificationMsg.AppendLine(CompareSourceAndTargetData(targetEntityCollection, sourceEntityCollection, "Target Dynamics", "FetchXML", false));
                }
            }
            return verificationMsg.ToString();
        }

        private EntityCollection LoadTargetDataToVerify(IOrganizationService organizationService, EntityCollection sourceEntities)
        {
            var sourceColumns = new List<string>();
            foreach (var curAttribute in sourceEntities[0].Attributes)
            {
                if (!ShallExcludeFieldsToCompare(curAttribute))
                    sourceColumns.Add(curAttribute.Key);
            }

            var targetQuery = new QueryExpression(sourceEntities[0].LogicalName)
            {
                ColumnSet = new ColumnSet(sourceColumns.ToArray()),
                Criteria = new FilterExpression
                {
                    FilterOperator = LogicalOperator.Or
                }
            };

            //add filter to load matching source PK (guids) records from target
            foreach (var sourceEntity in sourceEntities.Entities)
            {
                targetQuery.Criteria.AddCondition(sourceEntity.LogicalName + "id", ConditionOperator.Equal, sourceEntity.Id);
            }
            return organizationService.RetrieveMultiple(targetQuery);
        }

        private string CompareSourceAndTargetData(EntityCollection sourceEntityCollection, EntityCollection targetEntityCollection,
            string sourceName, string targetName, bool compareAttributes)
        {
            // Compare record count
            var verificationMsg = new StringBuilder();
            
            // Find the source data that are NOT in Target
            foreach (var sourceEntity in sourceEntityCollection.Entities)
            {
                var sourceExistInTarget = false;
                if (!targetEntityCollection.Entities.Contains(sourceEntity))
                {
                    foreach (var targetEntity in targetEntityCollection.Entities)
                    {
                        if (sourceEntity.Id.Equals(targetEntity.Id))
                        {
                            sourceExistInTarget = true;
                            if (compareAttributes)
                            {
                                var dataCompareMsg = CompareSourceAndTargetEntities(sourceEntity, targetEntity, sourceName, targetName);
                                if (!string.IsNullOrWhiteSpace(dataCompareMsg))
                                {
                                    verificationMsg.AppendLine("VERIFICATION: ERROR on [ " + CommonUtility.GetUserFriendlyMessage(sourceEntity) + " ]");
                                    verificationMsg.AppendLine(dataCompareMsg);
                                }
                            }
                        }   
                    }
                }
                else
                {
                    sourceExistInTarget = true;
                }

                if (!sourceExistInTarget)
                {
                    verificationMsg.AppendLine($"VERIFICATION: {CommonUtility.GetUserFriendlyMessage(sourceEntity)} does not exist in {targetName}!");
                }
            }

            return verificationMsg.ToString();
        }

        private string CompareSourceAndTargetEntities(Entity sourceEntity, Entity targetEntity, string sourceName, string targetName)
        {
            var unmatchedData = new StringBuilder();
            foreach (var sourceAttribute in sourceEntity.Attributes)
            {
                if (!ShallExcludeFieldsToCompare(sourceAttribute))
                {
                    // source attribute exist in target entity
                    if (targetEntity.Contains(sourceAttribute.Key))
                    {
                        if (sourceAttribute.Value == null && targetEntity[sourceAttribute.Key] == null)
                        {
                            // same - imported successfully
                        }
                        else if (sourceAttribute.Value != null && targetEntity[sourceAttribute.Key] == null)
                        {
                            unmatchedData.AppendLine($"VERIFICATION: Field: {sourceAttribute.Key}\t\t{sourceName}: {sourceAttribute.Value} \t\t{targetName}: (null)");
                        }
                        else if (sourceAttribute.Value == null && targetEntity[sourceAttribute.Key] != null)
                        {
                            unmatchedData.AppendLine($"VERIFICATION: Field: {sourceAttribute.Key}\t\t{sourceName}: (null) \t\t{targetName}: {targetEntity[sourceAttribute.Key]}");
                        }
                        else if (!sourceAttribute.Value.Equals(targetEntity[sourceAttribute.Key]))
                        {
                            unmatchedData.AppendLine($"VERIFICATION: Field: {sourceAttribute.Key}\t\t{sourceName}: {GetStringValue(sourceAttribute.Value)}" +
                                $"\t\t{targetName}: {GetStringValue(targetEntity[sourceAttribute.Key])}");
                        }
                    }
                    else
                    {
                        // alert only if the source value is null because the targetEntity will not have the attribute if it's value is null
                        if (sourceAttribute.Value != null)
                            unmatchedData.AppendLine($"VERIFICATION: Field: {sourceAttribute.Key}\t\t does not exist in {targetName}.");
                    }
                }
            }

            return unmatchedData.ToString();
        }

        public string GetStringValue(object value)
        {
            switch (value)
            {
                case EntityReference reference:
                    return "Lookup: " + reference.LogicalName + " " + reference.Id;
                case Guid guid:
                    return guid.ToString();
                case OptionSetValue setValue:
                    return setValue.Value.ToString();
            }

            return value.ToString();
        }

        private bool ShallExcludeFieldsToCompare(KeyValuePair<string, object> attribute)
        {
            var fieldName = attribute.Key.ToLower();
            if (fieldName == "iscustomizable" || fieldName == "associatedentitytypecode" ||
                fieldName == "calendarrules" || _excludedColumnsToCompare.Contains(fieldName))
            {
                return true;
            }

            switch (attribute.Value)
            {
                case null:
                    return false;
                // exclude any entityCollection column type as it is reference (calendarrule)
                case EntityCollection _:
                    return true;
            }

            return false;
        }

        public string VerifyTargetDataInSourceFetchXml(IOrganizationService organizationService, Dictionary<string, EntityCollection> sourceData)
        {
            var verificationMsg = new StringBuilder();
            if (sourceData == null || sourceData.Count == 0)
                return string.Empty;

            verificationMsg.AppendLine($"\r\n=============================================================================================\r\n" +
                "VERIFICATION: Checking if Target Dynamics has any extra records that do not exist in Source FetchXML, except for following entities:");
            verificationMsg.AppendLine($"\tAccount, Contact, SystemUser, Workflow\r\n");

            // Get target data to verify
            foreach (var curSourceData in sourceData)
            {
                // we know that these entities can have different data comparing to source
                // workflow - verification done after each workflowset-record import confirms that it has imported. and duplicate workflow is not issue
                if (curSourceData.Key == "account" || curSourceData.Key == "contact" || curSourceData.Key == "systemuser" || 
                    curSourceData.Key == "workflow")
                    continue;

                var targetEntityCollection = LoadTargetDataToVerify(organizationService, curSourceData.Key, curSourceData.Value);
                if (targetEntityCollection != null && targetEntityCollection.Entities.Count > 0)
                {
                    // We just want to check if the target record exist in Source environment.
                    var compareMsg = CompareSourceAndTargetData(targetEntityCollection, curSourceData.Value, "Target Dynamics", "FetchXML", true);
                    if (!string.IsNullOrWhiteSpace(compareMsg))
                    {
                        verificationMsg.AppendLine("===");
                        verificationMsg.AppendLine(compareMsg);
                    }
                }
            }
            verificationMsg.AppendLine($"===");
            return verificationMsg.ToString();
        }

        private static EntityCollection LoadTargetDataToVerify(IOrganizationService organizationService, string entityName, EntityCollection sourceEntityCollection)
        {
            if (sourceEntityCollection == null || sourceEntityCollection.Entities.Count == 0)
                return new EntityCollection();


            var targetQuery = new QueryExpression(entityName)
            {
                ColumnSet = new ColumnSet(true), //get all the columns to display them so that user can quickly find the extra record that is in the target!
                Criteria = new FilterExpression
                {
                    FilterOperator = LogicalOperator.And // must be AND
                }
            };

            var sourceIds = sourceEntityCollection.Entities.ToList().Select(x => x.Id).ToArray();
            targetQuery.Criteria.AddCondition(new ConditionExpression(entityName + "id", ConditionOperator.NotIn, sourceIds)); // must not use "new object[] "

            if (entityName.Equals("calendar", StringComparison.OrdinalIgnoreCase))
            {
                targetQuery.Criteria.AddCondition(new ConditionExpression("type", ConditionOperator.NotIn, 0, -1)); // skip inner calendars
            }
            else if (entityName.Equals("team", StringComparison.OrdinalIgnoreCase))
            {
                targetQuery.Criteria.AddCondition(new ConditionExpression("isdefault", ConditionOperator.Equal, new object[] { false })); // skip default team
                targetQuery.Criteria.AddCondition(new ConditionExpression("teamtype", ConditionOperator.Equal, new object[] { 0 })); // include owner team only
            }

            // load all Target records for <entityName>
            return organizationService.RetrieveMultiple(targetQuery);
        }

        public static void AddInSourceEntityCollection(Dictionary<string, EntityCollection> sourceData, EntityCollection fetchXmlQueryData)
        {
            EntityCollection sourceEntityCollection;
            if (sourceData.ContainsKey(fetchXmlQueryData.Entities[0].LogicalName))
            {
                sourceEntityCollection = sourceData[fetchXmlQueryData.Entities[0].LogicalName];
            }
            else
            {
                sourceEntityCollection = new EntityCollection();
                sourceData.Add(fetchXmlQueryData.Entities[0].LogicalName, sourceEntityCollection);
            }

            foreach (var curEntity in fetchXmlQueryData.Entities)
            {
                var tempEntity = new Entity(curEntity.LogicalName, curEntity.Id);

                // add the entity in the list only if it's not there yet.
                // There can be multiple fetchXML for same entity - so this checking is required to skip duplicate
                if (!sourceEntityCollection.Entities.Contains(tempEntity))
                {
                    sourceEntityCollection.Entities.Add(tempEntity);
                }
            }
        }

    }
}

