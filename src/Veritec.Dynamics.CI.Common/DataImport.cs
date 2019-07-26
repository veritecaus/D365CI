using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Extensions;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using MsCrmTools.DocumentTemplatesMover;

namespace Veritec.Dynamics.CI.Common
{
    public struct ForeignKeyInfo
    {
        public string LogicalName;
        public string FieldName;
        public Guid Id;
    }

    public class DataImport : CiBase
    {
        public event EventHandler<string> Logger;
        private readonly TransformData _transformData;
        public Dictionary<string, EntityCollection> SourceData;
        public Dictionary<string, EntityMetadata> TargetEntitiesMetaData;

        public DataImport(string dynamicsConnectionString, string targetDataReplace) : base(dynamicsConnectionString)
        {
            _transformData = new TransformData(targetDataReplace);
        }

        public DataImport(CrmParameter crmParameter, string targetDataReplaceInputFileName) : base(crmParameter)
        {
            _transformData = new TransformData(targetDataReplaceInputFileName);
        }

        public void ReadFetchXmlQueryResultsFromDiskAndImport(string dataFolder, string[] columnsToExcludeToCompareData, bool verifyDataImport = false)
        {
            SourceData = new Dictionary<string, EntityCollection>();

            if (columnsToExcludeToCompareData == null)
            {
                // match this with "ColumnsToExcludeToCompareData" in app.config - this is included here in case there is no exclusion in the fetch!!!
                var columnsToExclude = "languagecode;createdon;createdby;modifiedon;modifiedby;owningbusinessunit;owninguser;owneridtype;" +
                    "importsequencenumber;overriddencreatedon;timezoneruleversionnumber;operatorparam;utcconversiontimezonecode;versionnumber;" +
                    "customertypecode;matchingentitymatchcodetable;baseentitymatchcodetable;slaidunique;slaitemidunique;ignoreblankvalues";

                columnsToExcludeToCompareData = columnsToExclude.Split(';');
            }

            /* Load all the reference data related to the security framework (bu, team, security role, mailbox, queue) of the target */
            LoadTargetSecurityReferenceData();

            foreach (var file in Directory.EnumerateFiles(dataFolder, "*.xml"))
            {
                // must reset for every fetchxml data 
                var fetchXmlQueriesResultXml = new Dictionary<string, string>();

                Logger?.Invoke(this, "\r\n---\r\nReading : " + file);

                var nodes = XElement.Load(File.OpenRead(file));
                var entityName = XElement.Load(File.OpenRead(file)).FirstAttribute.Value;

                fetchXmlQueriesResultXml.Add(entityName, nodes.ToString());

                // save records into Target CRM
                WriteDatatoTargetCrm(fetchXmlQueriesResultXml, columnsToExcludeToCompareData, verifyDataImport);
            }

            // verify data import if requested :)
            if (verifyDataImport)
            {
                // Check if the target has any extra data (potential duplicate) comparing to source 
                var verificationMsg = new DataImportVerification(columnsToExcludeToCompareData).VerifyTargetDataInSourceFetchXml(
                    OrganizationService, SourceData);

                if (!string.IsNullOrWhiteSpace(verificationMsg)) Logger?.Invoke(this, verificationMsg);
            }
        }

        public void LoadTargetSecurityReferenceData()
        {
            var targetDataLoader = new DataLoader(OrganizationService);

            // load metadata for all
            Logger?.Invoke(this, "\r\nLoading Metadata, Business Units and Teams from target systems...");

            EntityCollection targetSystemUsers = null;
            EntityCollection targetOrgInstanceInfo = null;
            EntityCollection targetBUInfo = null;

            // load all the system reference data from target
            Parallel.Invoke(
                () => TargetEntitiesMetaData = targetDataLoader.GetAllEntitiesMetaData(EntityFilters.Entity),

                () => targetOrgInstanceInfo = targetDataLoader.GetAllEntity(Constant.Organization.EntityLogicalName,
                    new[] { Constant.Organization.Name }),

                () => targetBUInfo = targetDataLoader.GetAllEntity(Constant.BusinessUnit.EntityLogicalName,
                    new[] { Constant.BusinessUnit.Name }, LogicalOperator.And,
                    new[] { new ConditionExpression(Constant.BusinessUnit.ParentBusinessUnitId, ConditionOperator.Null) }),

                () => targetSystemUsers = targetDataLoader.GetAllEntity(Constant.User.EntityLogicalName,
                    new[] { Constant.User.DomainName }, LogicalOperator.And,
                    new[] { new ConditionExpression(Constant.User.DomainName, ConditionOperator.Equal, CrmParameter.UserName) })
            );

            Logger?.Invoke(this, "\r\nPreparing data replacement for Target Organization Info...");
            _transformData.SetOrganizationId(targetOrgInstanceInfo);

            // replace BU constant
            _transformData.ReplaceBUConstant(targetBUInfo);

            // root BU must be mapped before import can proceed
            _transformData.SetSourceAndTargetRootBusinessUnitMapping();

            Logger?.Invoke(this, "\r\nPreparing data replacement for System Administrator...");
            // Important note: if the SysAdmin is already set for Team.AdministratorId using the transform config then 
            // this method will NOT override the value set by the Transforms 
            _transformData.SetSystemAdministrator(targetSystemUsers, CrmParameter.UserName);
        }

        /// <summary>
        /// WriteDatatoTargetCRM
        /// </summary>
        /// <param name="fetchXmlQueriesResultXml">This must contain data for one entity only!</param>
        /// <param name="columnsToExcludeToCompareData"></param>
        /// <param name="verifyDataImport"></param>
        /// <returns></returns>
        private void WriteDatatoTargetCrm(Dictionary<string, string> fetchXmlQueriesResultXml,
            string[] columnsToExcludeToCompareData, bool verifyDataImport)
        {
            // load data
            var fetchXmlQueryData = LoadFetchXmlData(fetchXmlQueriesResultXml);

            // check if there is any data to import
            if (fetchXmlQueryData.Entities.Count == 0) return;

            // Add mapping for OOTB entities (Business Unit, Currency, Team and Security Role)
            var targetDataLoader = new DataLoader(OrganizationService);
            if (!AddTransformsForEntity(fetchXmlQueryData))
                return; //don't continue further - eg Security roles can't be imported using this utility

            // get medata of the entity
            var entityMetaData = TransformData.GetEntityMetaData(TargetEntitiesMetaData, fetchXmlQueryData.Entities[0].LogicalName);
            if (entityMetaData == null)
                throw new Exception($"ERROR: MetaData is missing for {fetchXmlQueryData.Entities[0].LogicalName}");

            EntityMetadata relationshipMetaData = null;
            if (entityMetaData.IsIntersect == true) // is this a many to many entity?
            {
                // load relationship
                relationshipMetaData = targetDataLoader.GetEntityMetaData(fetchXmlQueryData.Entities[0].LogicalName, EntityFilters.Relationships);
            }

            var dataImportVerification = new DataImportVerification(columnsToExcludeToCompareData);
            var allRecordsImported = false;

            var failedEntities = new List<Guid>();
            // this loop is required for self-relationship entity such as Account-ParentAccount, 
            var importCount = 0;
            while (allRecordsImported == false)
            {
                foreach (var curEntity in fetchXmlQueryData.Entities.ToArray())
                {
                    // first time (importCount == 0) - process all the entities. on second try process only the failed ones!
                    if (importCount == 0 || (importCount > 0 && failedEntities.Contains(curEntity.Id)))
                    {
                        var currentEntityImported = false;

                        // replace values using transforms - for the first time only
                        if (importCount == 0)
                        {
                            foreach (var a in curEntity.Attributes.ToArray())
                            {
                                // replace any Target Data (ie value to be replaced) with the value from the transform file
                                //if (curEntity.LogicalName == "contact")
                                    _transformData.TransformValue(curEntity, a.Key, a.Value);
                            }
                        }

                        if (entityMetaData.IsIntersect == true) //is this a many to many entity?
                        {
                            if (relationshipMetaData?.ManyToManyRelationships == null)
                                throw new Exception($"Relationship is missing for {entityMetaData.LogicalName}");

                            currentEntityImported = UpsertManyManyRecord(curEntity, relationshipMetaData.ManyToManyRelationships[0].SchemaName);
                        }
                        else
                        {
                            currentEntityImported = UpsertEntityRecord(curEntity, TargetEntitiesMetaData);
                        }

                        if (!currentEntityImported)
                        {
                            failedEntities.Add(curEntity.Id);
                            allRecordsImported = currentEntityImported;
                        }
                    }
                }
                importCount++;

                // try 3 times max - depth of the self-relation within an entity!!!
                // an account is parent of another account that has more than 1 child accounts.
                if (importCount > 2)
                    allRecordsImported = true;
            }

            // Verify data import if requested :)
            if (verifyDataImport && !entityMetaData.IsIntersect.Value)
            {
                // give extra time to publish the records - this can be extended to use asyncoperation to find the jobs and let the jobs finish before the Verification starts
                if (entityMetaData.LogicalName.Equals(Constant.DuplicateRule.EntityLogicalName, StringComparison.OrdinalIgnoreCase) ||
                    entityMetaData.LogicalName.Equals(Constant.Workflow.EntityLogicalName, StringComparison.OrdinalIgnoreCase) ||
                    entityMetaData.LogicalName.Equals(Constant.Sla.EntityLogicalName, StringComparison.OrdinalIgnoreCase))
                    // sleep to let the publishing finishes before continuing ;)
                    Thread.Sleep(16000);

                // Verify that the data in FetchXML match with saved data in Target
                var verificationMsg = dataImportVerification.VerifyDataImport(OrganizationService, fetchXmlQueryData);
                if (!string.IsNullOrWhiteSpace(verificationMsg)) Logger?.Invoke(this, verificationMsg);

                // Build Entity collection to use with Target data verification
                DataImportVerification.AddInSourceEntityCollection(SourceData, fetchXmlQueryData);
            }
        }

        public EntityCollection LoadFetchXmlData(Dictionary<string, string> fetchXmlQueriesResultXml)
        {
            var fetchXmlQueryData = new EntityCollection();

            foreach (var resultXml in fetchXmlQueriesResultXml)
            {
                var xElements = XElement.Parse(resultXml.Value).Nodes();
                var enumerable = xElements as IList<XNode> ?? xElements.ToList();

                if (enumerable.Any())
                {
                    foreach (var item in enumerable)
                    {
                        var typeList = new List<Type> { typeof(Entity) };
                        var dce = Deserialize<Entity>(item.ToString(), typeList);

                        fetchXmlQueryData.Entities.Add(dce);
                    }
                }
            }

            return fetchXmlQueryData;
        }

        /// <summary>
        /// reference: https://dzone.com/articles/using-datacontractserializer
        /// </summary>
        public static T Deserialize<T>(string xml, List<Type> types)
        {
            using (Stream stream = new MemoryStream())
            {
                var data = Encoding.UTF8.GetBytes(xml);
                stream.Write(data, 0, data.Length);
                stream.Position = 0;
                var dataContractSerializer = new DataContractSerializer(typeof(T), types);
                return (T)dataContractSerializer.ReadObject(stream);
            }
        }

        private bool AddTransformsForEntity(EntityCollection fetchXmlQueryData)
        {
            var targetDataLoader = new DataLoader(OrganizationService);
            var entityName = fetchXmlQueryData.Entities[0].LogicalName;

            switch (entityName)
            {
                case Constant.BusinessUnit.EntityLogicalName:
                    Logger?.Invoke(this, "\r\nPreparing data replacement for Business Units...");
                    var targetBusinessUnits = targetDataLoader.GetAllEntity(Constant.BusinessUnit.EntityLogicalName,
                        new[] { Constant.BusinessUnit.Name, Constant.BusinessUnit.ParentBusinessUnitId });
                    _transformData.PopulateBusinessUnitTransforms(fetchXmlQueryData, targetBusinessUnits);
                    break;

                case Constant.Role.EntityLogicalName:
                    // Note: Alternate key for Security role is "Business Unit Name + Security Role Name"
                    Logger?.Invoke(this, "\r\nPreparing data replacement for Security Roles...");
                    var targetSecurityRoles = targetDataLoader.GetAllEntity(Constant.Role.EntityLogicalName,
                        new[] { Constant.Role.Name, Constant.Role.BusinessUnitId });
                    _transformData.PopulateSecurityRoleTransforms(fetchXmlQueryData, targetSecurityRoles);
                    return false; //security role is imported using D365 solution

                case Constant.Team.EntityLogicalName:
                    // there can more than 1 teams with same name but different business unit - this is similar to security role!!!!!
                    // Note: Alternate key for Security role is "Business Unit Name + Security Role Name"
                    Logger?.Invoke(this, "\r\nPreparing data replacement for Team...");
                    var targetTeams = targetDataLoader.GetAllEntity(Constant.Team.EntityLogicalName,
                        new[] { Constant.Team.Name, Constant.Team.BusinessUnitId, Constant.Team.IsDefault }, LogicalOperator.And,
                        new[] { new ConditionExpression(Constant.Team.TeamType, ConditionOperator.Equal, 0) });
                    _transformData.PopulateTeamTransforms(fetchXmlQueryData, targetTeams);
                    break;

                case Constant.FieldSecurityProfile.EntityLogicalName:
                    // Note: Alternate key for Security role is "Business Unit Name + Security Role Name"
                    Logger?.Invoke(this, "\r\nPreparing data replacement for Field Security Profile...");
                    var targetFieldSecurityProfiles = targetDataLoader.GetAllEntity(Constant.FieldSecurityProfile.EntityLogicalName,
                        new[] { Constant.FieldSecurityProfile.Name });
                    _transformData.PopulateFieldSecurityProfileTransforms(fetchXmlQueryData, targetFieldSecurityProfiles);
                    return false; //security role is imported using D365 solution
 
                case Constant.TransactionCurrency.EntityLogicalName:
                    Logger?.Invoke(this, "\r\nPreparing data replacement for Currencies...");
                    var targetCurrencies = targetDataLoader.GetAllEntity(Constant.TransactionCurrency.EntityLogicalName,
                        new[] { Constant.TransactionCurrency.CurrencyName, Constant.TransactionCurrency.IsoCurrencyCode });
                    _transformData.PopulateCurrencyTransforms(fetchXmlQueryData, targetCurrencies);
                    break;

                case Constant.Queue.EntityLogicalName:
                    Logger?.Invoke(this, "\r\nPreparing data replacement for Queues...");
                    var targetQueues = targetDataLoader.GetAllEntity(Constant.Queue.EntityLogicalName,
                        new[] { Constant.Queue.Name, Constant.Queue.TransactionCurrencyId }, LogicalOperator.And,
                        new[] { new ConditionExpression(Constant.Queue.Name, ConditionOperator.DoesNotBeginWith, "<") });
                    _transformData.PopulateQueueTransforms(fetchXmlQueryData, targetQueues);
                    break;
            }
            return true;
        }

        private bool UpsertEntityRecord(Entity curEntity, Dictionary<string, EntityMetadata> entitiesMetaData)
        {
            var importedSuccessfully = true; //must be true when starting

            var debugInfo = string.Empty;
            try
            {
                debugInfo = "Checking if the source data exist in target";
                var entityMetaData = TransformData.GetEntityMetaData(entitiesMetaData, curEntity.LogicalName);

                Entity targetEntity = null;
                try
                {
                    /* check if record already exists*/
                    targetEntity = OrganizationService.Retrieve(curEntity.LogicalName, curEntity.Id, GetColumnSet(entityMetaData, curEntity.Attributes));
                }
                catch
                {
                    /* do nothing, just try to create below */
                }

                if (curEntity.LogicalName.Equals(Constant.DocumentTemplate.EntityLogicalName, StringComparison.OrdinalIgnoreCase))
                {
                    importedSuccessfully = UpsertDocumentTemplate(curEntity, targetEntity, entitiesMetaData, ref debugInfo);
                }
                else if (curEntity.LogicalName.Equals(Constant.DuplicateRule.EntityLogicalName, StringComparison.OrdinalIgnoreCase) ||
                    curEntity.LogicalName.Equals(Constant.DuplicateRuleCondition.EntityLogicalName, StringComparison.OrdinalIgnoreCase))
                {
                    importedSuccessfully = UpsertDuplicateDetectionRule(curEntity, targetEntity, entitiesMetaData, ref debugInfo);
                }
                // Workflow and SLA
                else if (curEntity.LogicalName.Equals(Constant.Workflow.EntityLogicalName, StringComparison.OrdinalIgnoreCase) ||
                    curEntity.LogicalName.Equals(Constant.Sla.EntityLogicalName, StringComparison.OrdinalIgnoreCase) ||
                    curEntity.LogicalName.Equals(Constant.SlaItem.EntityLogicalName, StringComparison.OrdinalIgnoreCase))
                {
                    importedSuccessfully = UpsertSlaOrWorkflow(curEntity, targetEntity, entitiesMetaData, ref debugInfo);
                }
                // Any other entity :)
                else
                {
                    if (targetEntity != null)
                    {
                        //The exchange rate of the base currency cannot be modified
                        if (curEntity.LogicalName.Equals(Constant.TransactionCurrency.EntityLogicalName, StringComparison.OrdinalIgnoreCase))
                        {
                            curEntity.Attributes.Remove(Constant.TransactionCurrency.ExchangeRate);
                        }

                        if (TargetSameAsSource(curEntity, targetEntity))
                        {
                            debugInfo = "Skipping Update";
                        }
                        else
                        {
                            debugInfo = "Updating";
                            OrganizationService.Update(curEntity);
                        }
                    }
                    else
                    {
                        
                        if (curEntity.Attributes.ContainsKey("statecode") && ((OptionSetValue)curEntity.Attributes["statecode"]).Value  == 1) //inactive
                        {
                            /* user trying to insert inactive record */
                            /* need to create first then update to inactive */
                            debugInfo = "Inserting Inactive";
                            OrganizationService.Create(new Entity(curEntity.LogicalName, curEntity.Id));
                            OrganizationService.Update(curEntity);
                        }
                        else
                        {
                            debugInfo = "Inserting";
                            OrganizationService.Create(curEntity);
                        }
                        
                    }
                }
                Logger?.Invoke(this, (targetEntity == null ? "Inserted" : "Updated") + ": " + CommonUtility.GetUserFriendlyMessage(curEntity));
            }
            catch (Exception exception)
            {
                Logger?.Invoke(this, $"\r\nDEBUG: {curEntity.LogicalName}, {curEntity.Id}");
                Logger?.Invoke(this, $"ERROR! {debugInfo}: {CommonUtility.GetUserFriendlyMessage(curEntity)}, message: {exception.Message}\r\n");

                importedSuccessfully = false;

                //in case of error - try importing next record
            }
            return importedSuccessfully;
        }

        private bool TargetSameAsSource(Entity sourceEntity, Entity targetEntity)
        {
            foreach (var sourceAttribute in sourceEntity.Attributes)
            {
                // make sure all attributes exist in target that are in source
                if (!targetEntity.Attributes.Keys.Contains(sourceAttribute.Key))
                    targetEntity.Attributes.Add(sourceAttribute.Key, null);

                // do the comparison
                if (!targetEntity.Attributes.Contains(sourceAttribute))
                {
                    return false;
                }
            }
            return true;
        }

        private static ColumnSet GetColumnSet(EntityMetadata entityMetaData, AttributeCollection sourceAttributeCollection)
        {
            if (entityMetaData == null)
                return new ColumnSet(false);

            if (entityMetaData.IsCustomEntity.Value || entityMetaData.LogicalName.Equals(Constant.DuplicateRule.EntityLogicalName, StringComparison.OrdinalIgnoreCase) ||
                entityMetaData.LogicalName.Equals(Constant.Sla.EntityLogicalName, StringComparison.OrdinalIgnoreCase) ||
                entityMetaData.LogicalName.Equals(Constant.Workflow.EntityLogicalName, StringComparison.OrdinalIgnoreCase))

                return new ColumnSet(Constant.Entity.StateCode, Constant.Entity.StatusCode);

            if (sourceAttributeCollection.Count > 0)
                return new ColumnSet(sourceAttributeCollection.Keys.ToArray());
            else
                return new ColumnSet(false);
        }

        private bool UpsertDocumentTemplate(Entity curEntity, Entity targetEntity,
            Dictionary<string, EntityMetadata> entitiesMetaData, ref string debugInfo)
        {
            if (!curEntity.LogicalName.Equals(Constant.DocumentTemplate.EntityLogicalName, StringComparison.OrdinalIgnoreCase))
                return true;

            // DocumentTemplate has objectTypeCode (EntityTypeCode) that can be
            // different in target environment. So fix it here.
            debugInfo = "Replacing Document Type code";
            ReplaceDocumentTypeCodes(curEntity, entitiesMetaData);

            if (targetEntity != null)
            {
                debugInfo = "Updating Document Template";
                OrganizationService.Update(curEntity);
            }
            else
            {
                debugInfo = "Inserting Document Template";
                OrganizationService.Create(curEntity);
            }

            return true;
        }

        private static void ReplaceDocumentTypeCodes(Entity curEntity, Dictionary<string, EntityMetadata> entitiesMetaData)
        {
            var entityMetaData = TransformData.GetEntityMetaData(entitiesMetaData, curEntity.LogicalName);

            // Entity type code issue is with Custom entity only!
            if (entityMetaData.IsCustomEntity.Value)
            {
                TemplatesManager templatesManager = new TemplatesManager();

                var associatedEntityName = curEntity.GetAttributeValue<string>(Constant.DocumentTemplate.AssociatedEntityTypeCode);
                // string name = curEntity.GetAttributeValue<string>("name");

                var associatedEntityMetaData = TransformData.GetEntityMetaData(entitiesMetaData, associatedEntityName);

                // ObjectTypeCode is fixed for OOTB entities - so just fix for custom entities
                if (associatedEntityMetaData.IsCustomEntity.Value)
                {
                    int? newEtc = associatedEntityMetaData.ObjectTypeCode; //t emplatesManager.GetEntityTypeCode(OrganizationService, entityName);

                    templatesManager.ReRouteEtcViaOpenXml(curEntity, associatedEntityName, associatedEntityMetaData.ObjectTypeCode);
                }
            }
        }

        private bool UpsertDuplicateDetectionRule(Entity curEntity, Entity targetEntity,
            Dictionary<string, EntityMetadata> targetEntitiesMetaData, ref string debugInfo)
        {
            if (!curEntity.LogicalName.Equals(Constant.DuplicateRule.EntityLogicalName, StringComparison.OrdinalIgnoreCase) &&
                !curEntity.LogicalName.Equals("duplicaterulecondition", StringComparison.OrdinalIgnoreCase))
                return true;

            // get status of the record in Source
            var curEntityBeforeChange = new Entity(curEntity.LogicalName, curEntity.Id);
            if (curEntity.Contains(Constant.Entity.StateCode))
            {
                curEntityBeforeChange[Constant.Entity.StateCode] = curEntity[Constant.Entity.StateCode];
            }
            if (curEntity.Contains(Constant.Entity.StatusCode))
            {
                curEntityBeforeChange[Constant.Entity.StatusCode] = curEntity[Constant.Entity.StatusCode];
            }

            // We can't import Published Duplicate Detection rules - so make them unpublished before we import them!!!
            if (curEntity.LogicalName.Equals(Constant.DuplicateRule.EntityLogicalName, StringComparison.OrdinalIgnoreCase))
            {
                curEntity[Constant.Entity.StateCode] = new OptionSetValue(0); //inactive
                curEntity[Constant.Entity.StatusCode] = new OptionSetValue(0); //Unpublished

                // Fix any baseentitytypecode and matchingentitytypecode if any entity-typecode changed in target environment
                FixObjectTypeCodeForDdRule(targetEntitiesMetaData, curEntity);
            }
            else if (curEntity.LogicalName.Equals(Constant.DuplicateRuleCondition.EntityLogicalName, StringComparison.OrdinalIgnoreCase))
            {
                // https://msdn.microsoft.com/en-us/library/gg334583.aspx
                // Don’t set the OperatorParam to zero during Create or Update operations. 
                if (curEntity.Contains("operatorparam") && curEntity.GetAttributeValue<int>("operatorparam") == 0)
                    curEntity.Attributes.Remove("operatorparam");
            }

            if (targetEntity != null)
            {
                // Must unpulish the rule in target before import!
                debugInfo = "Unpublishing Duplicate Detection rule";
                UnpublishDuplicateDetectionRuleBeforeImport(targetEntity);

                debugInfo = "Updating Duplicate Detection rule";
                OrganizationService.Update(curEntity);
            }
            else
            {
                try
                {
                    debugInfo = "Inserting Duplicate Detection rule";
                    OrganizationService.Create(curEntity);
                }
                catch (Exception ex)
                {
                    // Error: The rule condition cannot be created or updated because it would cause the matchcode length to exceed the maximum limit. 
                    // this error occurred when the target environment has duplicate detection rules that was not imported using the Source Env (DEV) and 
                    // the rule uses one or more text field and combination of all the fields including the one being inserted would result to exceed the max limit
                    if (ex.Message.IndexOf("matchcode length to exceed", StringComparison.OrdinalIgnoreCase) > 0)
                    {
                        throw new Exception("Please delete the duplicate detection rule from the target Dynamics and try again!", ex);
                    }

                    throw ex;
                }
            }

            /////////////////////////////special case
            /////1. Duplicate Detection rule should be published only after importing all the Duplicate Detection rule conditions.
            /////2. SLA should be activated only after importing all the SLA Items.
            /////
            /////So there will be 2 set of fetchxml to export/import these entities.
            /////First one migrate the record with DRAFT status
            /////Second one to update Status of the records
            debugInfo = "Publishing Duplicate Detection rule";
            PublishDuplicateDetectionRule(curEntityBeforeChange, curEntity);

            return true;
        }

        private static void FixObjectTypeCodeForDdRule(Dictionary<string, EntityMetadata> entitiesMetaData, Entity curDdRuleEntity)
        {
            EntityMetadata entityMetaData = null;
            // baseentitytypecode or matchingentitytypecode may not be part of the curDDRuleEntity if we just want to activate the DD rules
            var baseentitytypecode = curDdRuleEntity.GetAttributeValue<OptionSetValue>(Constant.DuplicateRule.BaseEntityTypeCode);

            // it is safe to assume that any entity type code below 10000 are not custom entities - so skip them
            if (baseentitytypecode != null && baseentitytypecode.Value > 9999)
            {
                var baseentitytypecodeName = curDdRuleEntity.FormattedValues[Constant.DuplicateRule.BaseEntityTypeCode];

                // Find metadata using entitytypecodename and update baseentitytypecode if required
                if (!string.IsNullOrWhiteSpace(baseentitytypecodeName))
                {
                    entityMetaData = TransformData.GetEntityMetaData(entitiesMetaData, baseentitytypecodeName);

                    if (entityMetaData.IsCustomEntity.Value && baseentitytypecode.Value != entityMetaData.ObjectTypeCode.Value)
                    {
                        curDdRuleEntity[Constant.DuplicateRule.BaseEntityTypeCode] = new OptionSetValue(entityMetaData.ObjectTypeCode.Value);
                    }
                }
            }

            var matchingentitytypecode = curDdRuleEntity.GetAttributeValue<OptionSetValue>(Constant.DuplicateRule.MatchingEntityTypeCode);
            // it is safe to assume that any entity type code below 10000 are not custom entities - so skip them
            if (matchingentitytypecode != null && matchingentitytypecode.Value > 9999)
            {
                var matchingentitytypecodeName = curDdRuleEntity.FormattedValues[Constant.DuplicateRule.MatchingEntityTypeCode];

                // Find metadata using entitytypecodename and update matchingentitytypecode if required
                if (!string.IsNullOrWhiteSpace(matchingentitytypecodeName))
                {
                    entityMetaData = TransformData.GetEntityMetaData(entitiesMetaData, matchingentitytypecodeName);

                    if (entityMetaData.IsCustomEntity.Value && matchingentitytypecode.Value != entityMetaData.ObjectTypeCode.Value)
                    {
                        curDdRuleEntity[Constant.DuplicateRule.MatchingEntityTypeCode] = new OptionSetValue(entityMetaData.ObjectTypeCode.Value);
                    }
                }
            }
        }

        public void UnpublishDuplicateDetectionRuleBeforeImport(Entity targetEntity)
        {
            if (targetEntity.LogicalName.Equals(Constant.DuplicateRule.EntityLogicalName, StringComparison.OrdinalIgnoreCase))
            {
                if (!targetEntity.Contains(Constant.Entity.StatusCode) || (targetEntity.Contains(Constant.Entity.StatusCode) &&
                    targetEntity.GetAttributeValue<OptionSetValue>(Constant.Entity.StatusCode).Value == 2)) //"Published"
                {
                    // unpublish the DuplicateDetection rule before we import (update)
                    var publishReq = new UnpublishDuplicateRuleRequest { DuplicateRuleId = targetEntity.Id };
                    OrganizationService.Execute(publishReq);

                    // sleep to let the process finish before continuing ;)
                    Thread.Sleep(2000);
                }
            }
        }

        public void PublishDuplicateDetectionRule(Entity entityBeforeStatusCodeChange, Entity curEntity)
        {
            if (entityBeforeStatusCodeChange.LogicalName.Equals(Constant.DuplicateRule.EntityLogicalName, StringComparison.OrdinalIgnoreCase))
            {
                if (entityBeforeStatusCodeChange.Contains(Constant.Entity.StatusCode) &&
                    entityBeforeStatusCodeChange.GetAttributeValue<OptionSetValue>(Constant.Entity.StatusCode).Value == 2) //"Published"
                {
                    var publishReq = new PublishDuplicateRuleRequest { DuplicateRuleId = entityBeforeStatusCodeChange.Id };
                    OrganizationService.Execute(publishReq);

                    // update the curEntity to publish state in order to get it matched during VERIFICATION
                    curEntity[Constant.Entity.StateCode] = entityBeforeStatusCodeChange.GetAttributeValue<OptionSetValue>(Constant.Entity.StateCode); //active
                    curEntity[Constant.Entity.StatusCode] = entityBeforeStatusCodeChange.GetAttributeValue<OptionSetValue>(Constant.Entity.StatusCode); //published

                    // sleep to let the process finish before continuing ;)
                    Thread.Sleep(2000);
                }
            }
        }

        private bool UpsertSlaOrWorkflow(Entity curEntity, Entity targetEntity, Dictionary<string, EntityMetadata> entitiesMetaData, ref string debugInfo)
        {
            if (!curEntity.LogicalName.Equals("workflow", StringComparison.OrdinalIgnoreCase) &&
                !curEntity.LogicalName.Equals("sla", StringComparison.OrdinalIgnoreCase) &&
                !curEntity.LogicalName.Equals("slaitem", StringComparison.OrdinalIgnoreCase))
                return true;

            var curEntityBeforeChange = new Entity(curEntity.LogicalName, curEntity.Id);
            if (curEntity.Contains(Constant.Entity.StateCode)) curEntityBeforeChange[Constant.Entity.StateCode] = curEntity[Constant.Entity.StateCode];
            if (curEntity.Contains(Constant.Entity.StatusCode)) curEntityBeforeChange[Constant.Entity.StatusCode] = curEntity[Constant.Entity.StatusCode];

            // DocumentTemplate and SLA has objectTypeCode (EntityTypeCode) that can be
            // different in target environment. So fix it here.
            if (curEntity.LogicalName.Equals("sla", StringComparison.OrdinalIgnoreCase))
            {
                debugInfo = "Updating ObjectTypeCode of the entity with SLA";
                // ObjecttypecodeName is the Display name which we can't use to get objecttypecode from metadata!!!
                // hence we're using Transforms
                var oldObjectTypeCode = curEntity.GetAttributeValue<OptionSetValue>("objecttypecode");
                if (oldObjectTypeCode != null && oldObjectTypeCode.Value > 9999) // custom entity object type code starts from 10000!!!
                {
                    // get entity metadata for this entity for which sla is created using oldObjectTypeCode
                    var oldObjectTypeCodeName = (curEntity.FormattedValues.Contains("objecttypecode") ? curEntity.FormattedValues["objecttypecode"] : null);
                    var targetSlaEntityMetaData = TransformData.GetEntityMetaData(entitiesMetaData, oldObjectTypeCodeName);

                    // objecttypecode issue is with cutom entities only 
                    if (targetSlaEntityMetaData != null) // metadata found
                    {
                        if (targetSlaEntityMetaData.IsCustomEntity.Value && oldObjectTypeCode.Value != targetSlaEntityMetaData.ObjectTypeCode.Value)
                        {
                            curEntity.Attributes["objecttypecode"] = new OptionSetValue(targetSlaEntityMetaData.ObjectTypeCode.Value);
                            curEntity.Attributes["primaryentityotc"] = targetSlaEntityMetaData.ObjectTypeCode.Value;
                        }
                    }
                    else
                    {
                        // if it can't find in the metadata - then use data transform config
                        var newObjectTypeCode = _transformData.TransformObjectValue(
                            curEntity.LogicalName, "objecttypecode", oldObjectTypeCode.Value.ToString());

                        if (!string.IsNullOrWhiteSpace(newObjectTypeCode) && !oldObjectTypeCode.Value.ToString().Equals(newObjectTypeCode))
                        {
                            curEntity.Attributes["objecttypecode"] = new OptionSetValue(int.Parse(newObjectTypeCode));
                            curEntity.Attributes["primaryentityotc"] = int.Parse(newObjectTypeCode);
                        }
                    }
                }

                // we can't import Active SLA - so make them inactive before we import them.!!!
                curEntity[Constant.Entity.StateCode] = new OptionSetValue(0); //inactive
                curEntity[Constant.Entity.StatusCode] = new OptionSetValue(1); //draft
            }

            var wfAndSlaUtility = new WFandSlaMigrateUtility(OrganizationService);

            if (targetEntity != null)
            {
                debugInfo = "Deactivating SLA related to the workflow";
                wfAndSlaUtility.DeactivateSlaRelatedToWorkflow(curEntity);

                debugInfo = "Deactivatating SLA before importing SLA";
                wfAndSlaUtility.DeactivateSlaBeforeImportSla(curEntity);

                //workflow in Target must be deactivated before it can be imported!
                debugInfo = "Deactivating Workflow before importing workflow";
                wfAndSlaUtility.DeactivateWorkflowBeforeImportingWorkflow(targetEntity);

                debugInfo = "Updating";
                OrganizationService.Update(curEntity);
            }
            else
            {
                debugInfo = "Inserting";
                OrganizationService.Create(curEntity);
            }

            //activate the SLA in target only if it is active in source
            debugInfo = "Activating SLA";
            wfAndSlaUtility.ActivateSlaAfterImport(curEntityBeforeChange, curEntity);

            return true;
        }

        private bool UpsertManyManyRecord(Entity curEntity, string manyToManyRelationshipName)
        {
            var relationships = from a in curEntity.Attributes
                                where a.Value.ToString() != curEntity.Id.ToString()
                                select new ForeignKeyInfo
                                {
                                    Id = new Guid(a.Value.ToString()),
                                    LogicalName = a.Key.Substring(0, a.Key.LastIndexOf("id", StringComparison.Ordinal)),
                                    FieldName = a.Key
                                };

            var relationshipsList = relationships.ToList();

            if (relationshipsList.Count != 2)
            {
                Logger?.Invoke(this, "\r\nERROR: many to many data issue: coudn't find references for " + CommonUtility.GetUserFriendlyMessage(curEntity) + "\r\n");
                return false;
            }

            var foreignKeyInfo1 = relationshipsList[0];
            var foreignKeyInfo2 = relationshipsList[1];

            if (foreignKeyInfo1.LogicalName == null)
            {
                Logger?.Invoke(this, "\r\nERROR: many to many data issue: Target entity reference is missing for " + CommonUtility.GetUserFriendlyMessage(curEntity) + "\r\n");
                return false;
            }

            if (foreignKeyInfo2.LogicalName == null)
            {
                Logger?.Invoke(this, "\r\nERROR: many to many data issue: Related entity reference is missing for " + CommonUtility.GetUserFriendlyMessage(curEntity) + "\r\n");
                return false;
            }

            var qe = new QueryExpression(curEntity.LogicalName);

            //we can't use entityId when creating many-many record.
            //qe.Criteria.AddCondition(new ConditionExpression(e.LogicalName + "id", ConditionOperator.Equal, e.Id));

            //so use two side of the record find matching
            qe.Criteria.AddCondition(new ConditionExpression(foreignKeyInfo1.FieldName, ConditionOperator.Equal, foreignKeyInfo1.Id));
            qe.Criteria.AddCondition(new ConditionExpression(foreignKeyInfo2.FieldName, ConditionOperator.Equal, foreignKeyInfo2.Id));

            var rmr = new RetrieveMultipleRequest { Query = qe };

            var response = (RetrieveMultipleResponse)OrganizationService.Execute(rmr);

            if (response?.EntityCollection?.Entities != null && response.EntityCollection.Entities.Count > 0)
            {
                Logger?.Invoke(this, "many to many exists: " + CommonUtility.GetUserFriendlyMessage(curEntity));
            }
            else
            {
                var associateRequest = new AssociateRequest
                {
                    Target = new EntityReference(foreignKeyInfo1.LogicalName, foreignKeyInfo1.Id),
                    RelatedEntities = new EntityReferenceCollection { new EntityReference(foreignKeyInfo2.LogicalName, foreignKeyInfo2.Id) },
                    Relationship = new Relationship(manyToManyRelationshipName)
                };

                try
                {
                    OrganizationService.Execute(associateRequest);
                    Logger?.Invoke(this, "many to many created: " + CommonUtility.GetUserFriendlyMessage(curEntity));
                }
                catch (Exception ex)
                {
                    #region Generate user-friendly msg

                    string parentEntityInfo = string.Empty;
                    try
                    {
                        Entity targetEntity = OrganizationService.Retrieve(foreignKeyInfo1.LogicalName, foreignKeyInfo1.Id, new ColumnSet(true));
                        parentEntityInfo = "\r\nTarget Entity info: " + CommonUtility.GetUserFriendlyMessage(targetEntity);
                    }
                    catch
                    { }

                    try
                    {
                        Entity relatedEntity = OrganizationService.Retrieve(foreignKeyInfo2.LogicalName, foreignKeyInfo2.Id, new ColumnSet(true));
                        parentEntityInfo = parentEntityInfo + "\r\nRelated Entity info: " + CommonUtility.GetUserFriendlyMessage(relatedEntity);
                    }
                    catch
                    { }

                    Logger?.Invoke(this, $"\r\nDEBUG: {foreignKeyInfo1.LogicalName}, {foreignKeyInfo1.Id}");
                    Logger?.Invoke(this, $"DEBUG: {foreignKeyInfo2.LogicalName}, {foreignKeyInfo2.Id}");
                    Logger?.Invoke(this, $"DEBUG: {curEntity.LogicalName}, {curEntity.Id}");
                    Logger?.Invoke(this, "ERROR! associating many to many: " + CommonUtility.GetUserFriendlyMessage(curEntity) +
                        ", error message:" + ex.Message + parentEntityInfo + "\r\n");

                    return false;

                    #endregion

                    //in case error - try importing next record
                }
            }

            // all is well
            return true;
        }
    }
}
