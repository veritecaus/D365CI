using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Newtonsoft.Json;

namespace Veritec.Dynamics.CI.Common
{
    internal class TransformData
    {
        private readonly List<Guid> _sameGuidsInSourceAndTarget;
        private EntityReference _targetSystemAdministrator;
        private Guid _targetRootBuId = Guid.Empty;
        private Guid _sourceRootBuId = Guid.Empty;
        private readonly TransformConfig _transformConfig;

        public TransformData(string[] targetDataReplaceInputFileNames)
        {
            _sameGuidsInSourceAndTarget = new List<Guid>();

            List<Transform> transforms = new List<Transform>();
            foreach (var ReplaceInputFileName in targetDataReplaceInputFileNames)
            {
                if (File.Exists(ReplaceInputFileName))
                {
                    var jsonConfigTransforms = File.ReadAllText(ReplaceInputFileName);
                    var jsonConfigTransformsDeserialized = JsonConvert.DeserializeObject<IList<Transform>>(jsonConfigTransforms);
                    transforms.AddRange(jsonConfigTransformsDeserialized);
                }
            }

            _transformConfig = new TransformConfig(transforms);
        }

        public void TransformValue(Entity sourceEntity, string sourceAttribute, object sourceValue)
        {
            // don't proceed further if we don't need to replace
            if (sourceEntity == null || string.IsNullOrWhiteSpace(sourceEntity.LogicalName) ||
                string.IsNullOrWhiteSpace(sourceAttribute) || sourceValue == null) return;

            switch (sourceValue)
            {
                case Guid _:
                    // these columns can never be mapped
                    if (sourceAttribute.Equals("address1_addressid") ||
                        sourceAttribute.Equals("address2_addressid") ||
                        sourceAttribute.Equals("address3_addressid"))
                    {
                        return;
                    }

                    TransformTargetGuidValue(sourceEntity, sourceAttribute, sourceValue);
                    break;

                case EntityReference _:
                    TransformTargetEntityReferenceValue(sourceEntity, sourceAttribute, sourceValue);
                    break;

                case object s when (s.GetType().IsPrimitive || s.GetType() == typeof(string)):
                    // the IsPrimitive types are Boolean, Byte, SByte, Int16, UInt16, Int32, UInt32, Int64, UInt64, IntPtr, UIntPtr, Char, Double, and Single.

                    var transform = TransformObject(sourceEntity.LogicalName, sourceAttribute, sourceValue);

                    if (transform != null && !transform.ReplacementValue.ToString().Equals(sourceValue.ToString(), StringComparison.OrdinalIgnoreCase))
                    {
                        sourceEntity[transform.ReplacementAttribute] = Convert.ChangeType(transform.ReplacementValue, sourceValue.GetType());
                    }
                    break;

            }
        }

        private void TransformTargetEntityReferenceValue(Entity sourceEntity, string attributeName, object oldValue)
        {
            var oldEntityReference = (EntityReference)oldValue;

            // if the guid of the record is same in source and target then there is no mapping as no need to replace 
            if (_sameGuidsInSourceAndTarget != null && _sameGuidsInSourceAndTarget.Contains(oldEntityReference.Id))
                return;

            // don't proceed if this field is not in the transform config
            if (!_transformConfig.ContainsKey(sourceEntity.LogicalName, attributeName.ToLower()) &&
                !_transformConfig.ContainsKey(oldEntityReference.LogicalName, (oldEntityReference.LogicalName + "id").ToLower()))
            {
                return;
            }

            var substituteValue = TransformObjectValue(sourceEntity.LogicalName, attributeName, oldEntityReference.Id);

            // Example: Owner of an entity record is a Team. So, if there is no mapping for the "OwnerId" of the entity
            // then try searching mapping of "Team/TeamId"
            if (oldEntityReference.Id.Equals(substituteValue))
            {
                substituteValue = TransformObjectValue(oldEntityReference.LogicalName,
                        oldEntityReference.LogicalName + "id", oldEntityReference.Id);
            }

            if (!substituteValue.Equals(Guid.Empty) && !oldEntityReference.Id.Equals(substituteValue))
            {
                oldEntityReference.Id = substituteValue;
                sourceEntity[attributeName] = oldEntityReference;
            }
        }

        private string GetPrimaryNameAttribute(string entityName)
        {
            if (entityName.Equals(Constant.BusinessUnit.EntityLogicalName, StringComparison.OrdinalIgnoreCase) ||
                entityName.Equals(Constant.Team.EntityLogicalName, StringComparison.OrdinalIgnoreCase) ||
                entityName.Equals(Constant.Queue.EntityLogicalName, StringComparison.OrdinalIgnoreCase))
            {
                return "name";
            }

            return "PRIMARYNAMEATTRIBUTE_" + entityName;
        }

        /// <summary>
        /// Checks if the source value exists in the transform config. The method converts all objects to their string representation
        /// for the purpose of checking the 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="entityName"></param>
        /// <param name="attributeName"></param>
        /// <param name="sourceValue"></param>
        /// <returns></returns>
        public T TransformObjectValue<T>(string entityName, string attributeName, T sourceValue)
        {
            var transformMatch = TransformObject(entityName, attributeName, sourceValue);

            if (transformMatch == null || string.IsNullOrEmpty(transformMatch.ReplacementValue))
                return sourceValue;

            if (sourceValue is Guid)
                return (T)Convert.ChangeType(Guid.Parse(transformMatch.ReplacementValue), sourceValue.GetType());

            return (T)Convert.ChangeType(transformMatch.ReplacementValue, sourceValue.GetType());
        }

        public Transform TransformObject<T>(string entityName, string attributeName, T sourceValue)
        {
            var transformMatch = _transformConfig[entityName, attributeName, sourceValue.ToString().ToLower()] ??
                         _transformConfig[entityName, attributeName, "*"];

            if (transformMatch == null || string.IsNullOrEmpty(transformMatch.ReplacementValue))
                return null;

            return transformMatch;
        }

        private void TransformTargetGuidValue(Entity sourceEntity, string attributeName, object oldValue)
        {
            var oldGuid = (Guid)oldValue;
            var substituteValue = (Guid)oldValue;

            // if the guid of the record is same in source and target then there is no mapping as no need to replace 
            if (_sameGuidsInSourceAndTarget.Contains(oldGuid))
                return;

            // this is for teamrole entity 
            if (sourceEntity.LogicalName.Equals(Constant.TeamRoles.EntityLogicalName, StringComparison.OrdinalIgnoreCase) ||
                sourceEntity.LogicalName.Equals(Constant.TeamProfiles.EntityLogicalName, StringComparison.OrdinalIgnoreCase))
            {
                if (attributeName.Equals(Constant.TeamRoles.TeamId, StringComparison.OrdinalIgnoreCase))
                {
                    substituteValue = TransformObjectValue("team", "teamid", oldGuid);
                }
                else if (attributeName.Equals(Constant.TeamRoles.RoleId, StringComparison.OrdinalIgnoreCase))
                {
                    substituteValue = TransformObjectValue("role", "roleid", oldGuid);
                }
                else if (attributeName.Equals(Constant.TeamProfiles.FieldSecurityProfileId, StringComparison.OrdinalIgnoreCase))
                {
                    substituteValue = TransformObjectValue("fieldsecurityprofile", "fieldsecurityprofileid", oldGuid);
                }
            }

            if (oldGuid.Equals(substituteValue))
            {
                // in case data replacement not found then try using the entity name - this mapping will exist in Transform config
                substituteValue = TransformObjectValue(sourceEntity.LogicalName, attributeName, oldGuid);
            }

            if (substituteValue.Equals(Guid.Empty) || oldGuid.Equals(substituteValue)) return;

            sourceEntity[attributeName] = substituteValue;

            if (attributeName.Equals(sourceEntity.LogicalName + "id", StringComparison.OrdinalIgnoreCase))
            {
                sourceEntity.Id = substituteValue;
            }
        }

        public static EntityMetadata GetEntityMetaData(Dictionary<string, EntityMetadata> entitiesMetaData,
            string entityName, string entityTypeCode = null)
        {
            // must have either entityName or entityTypeCode
            if (entityName == null && entityTypeCode == null)
                return null;

            EntityMetadata entityMetaData = null;
            if (entityName != null && entitiesMetaData.ContainsKey(entityName))
            {
                entityMetaData = entitiesMetaData[entityName];
            }
            else
            {
                if (!string.IsNullOrWhiteSpace(entityName))
                    entityMetaData = FindEntityMetaData(entitiesMetaData, entityName);
                else if (!string.IsNullOrWhiteSpace(entityTypeCode))
                    entityMetaData = FindEntityMetaData(entitiesMetaData, entityTypeCode);
            }

            if (entityMetaData == null)
                throw new Exception($"ERROR! MetaData is missing for {entityName}");
            return entityMetaData;
        }

        public static EntityMetadata FindEntityMetaData(Dictionary<string, EntityMetadata> entitiesMetaData, string valueToMatch)
        {
            EntityMetadata matchedEntityMetadata = null;
            foreach (var curEntityMetaData in entitiesMetaData)
            {
                if (curEntityMetaData.Value.DisplayName.UserLocalizedLabel != null && curEntityMetaData.Value.ObjectTypeCode.HasValue)
                {
                    if (curEntityMetaData.Value.DisplayName.UserLocalizedLabel.Label.Equals(valueToMatch, StringComparison.OrdinalIgnoreCase) ||
                        curEntityMetaData.Value.ObjectTypeCode.Value.ToString().Equals(valueToMatch, StringComparison.OrdinalIgnoreCase))
                    {
                        // this logic will make this method little-bit slow but it's worth it :) 
                        if (matchedEntityMetadata == null)
                            matchedEntityMetadata = curEntityMetaData.Value;
                        else
                            throw new Exception($"ERROR! There are more than one entities that match [ {valueToMatch} ].\r\nPlease fix the issue before trying again!");
                    }
                }
            }

            return matchedEntityMetadata;
        }

        public void SetOrganizationId(EntityCollection targetOrgInstanceInfo)
        {
            if (targetOrgInstanceInfo != null && targetOrgInstanceInfo.Entities.Count > 0)
            {
                _transformConfig.Transforms.Add(new Transform(Constant.Organization.EntityLogicalName, Constant.Organization.Id, "*", targetOrgInstanceInfo[0].Id.ToString()));
            }
            else
                throw new Exception("Organization record missing in Target Dynamics");
        }

        public void ReplaceBUConstant(EntityCollection targetBUInfo)
        {
            if (targetBUInfo == null && targetBUInfo.Entities.Count != 1)
                throw new Exception("Unable to detect root Business Unit");

            foreach (var transform in _transformConfig.Transforms)
            {
                if (transform.ReplacementValue == Constant.TransformConstant.DESTINATIONROOTBU)
                {
                    transform.ReplacementValue = targetBUInfo.Entities[0].ToEntityReference().Id.ToString();
                }
            }

        }

        public void ReplaceFetchXMLEntries(DataLoader dataLoader)
        {
            foreach (var transform in _transformConfig.Transforms)
            {
                if (transform.ReplacementValue.StartsWith("<fetch"))
                {
                    var resultCollection = dataLoader.ExecuteFetch(transform.ReplacementValue);

                    if (resultCollection.Entities.Count != 1 || resultCollection[0].Attributes.Count != 1)
                        throw new Exception($"Only one record with one field can be accepted as a return value from a transform fetchxml. Record Count: {resultCollection.TotalRecordCount}, Attibute Count: {resultCollection[0].Attributes.Count}, Offending FetchXML: {transform.ReplacementValue} ");

                    transform.ReplacementValue = resultCollection.Entities.First().Attributes.First().Value.ToString();
                }
            }
        }

        /// <summary>
        /// Retrieves the root business unit and adds the ID to the target data replacements. It must be included in Transform config
        /// </summary>
        public void SetSourceAndTargetRootBusinessUnitMapping()
        {
            // get Parent Business Id
            var targetRootBu = TransformObjectValue(Constant.BusinessUnit.EntityLogicalName, Constant.BusinessUnit.ParentBusinessUnitId, "*");
            if (string.IsNullOrWhiteSpace(targetRootBu) || !Guid.TryParse(targetRootBu, out _targetRootBuId))
            {
                throw new Exception("Please update the transform config to include parentbusinessunitid for Target Root Business Unit Guid.");
            }

            var transform = _transformConfig[Constant.BusinessUnit.EntityLogicalName, Constant.BusinessUnit.ParentBusinessUnitId, "*"];

            if (transform is null)
                throw new Exception("Please update the transform config to include businessunitid for Root Business Unit guid of Source and Target.");

            _sourceRootBuId = Guid.Parse(transform.ReplacementValue);
        }

        public void SetSystemAdministrator(EntityCollection targetSystemUsers, string systemUserName)
        {
            if (targetSystemUsers != null && targetSystemUsers.Entities.Count > 0)
            {
                foreach (var curSystemUser in targetSystemUsers.Entities)
                {
                    if (curSystemUser.GetAttributeValue<string>(Constant.User.DomainName).ToLower().Equals(systemUserName.ToLower()))
                    {
                        _targetSystemAdministrator = targetSystemUsers[0].ToEntityReference();
                        _targetSystemAdministrator.Name = systemUserName;

                        // make current user (of the OrganizationSvc) as Administrator Id, OwnerId for replacement
                        _transformConfig.Transforms.Add(new Transform(Constant.Team.EntityLogicalName, Constant.Team.AdministratorId, "*", _targetSystemAdministrator.Id.ToString()));
                        _transformConfig.Transforms.Add(new Transform(Constant.Team.EntityLogicalName, Constant.Team.AdministratorId, "*", _targetSystemAdministrator.Id.ToString()));
                        _transformConfig.Transforms.Add(new Transform(Constant.DuplicateRule.EntityLogicalName, Constant.Entity.OwnerId, "*", _targetSystemAdministrator.Id.ToString()));
                        _transformConfig.Transforms.Add(new Transform(Constant.Workflow.EntityLogicalName, Constant.Entity.OwnerId, "*", _targetSystemAdministrator.Id.ToString()));
                        _transformConfig.Transforms.Add(new Transform(Constant.Sla.EntityLogicalName, Constant.Entity.OwnerId, "*", _targetSystemAdministrator.Id.ToString()));
                        _transformConfig.Transforms.Add(new Transform(Constant.SlaItem.EntityLogicalName, Constant.Entity.OwnerId, "*", _targetSystemAdministrator.Id.ToString()));

                        return;
                    }
                }
            }

            throw new Exception($"Current user {systemUserName} must be System Administrator in Target Dynamics!");
        }

        public void PopulateBusinessUnitTransforms(EntityCollection sourceBusinessUnits, EntityCollection targetBusinessUnits)
        {
            foreach (var sourceBusinessUnit in sourceBusinessUnits.Entities)
            {
                var parentBuId = sourceBusinessUnit.GetAttributeValue<EntityReference>(Constant.BusinessUnit.ParentBusinessUnitId);

                foreach (var targetBusinessUnit in targetBusinessUnits.Entities)
                {
                    // source and target guids are same - no benefit of adding to the transform config
                    if (sourceBusinessUnit.Id.Equals(targetBusinessUnit.Id))
                    {
                        _sameGuidsInSourceAndTarget.Add(sourceBusinessUnit.Id);
                        break;
                    }

                    if (parentBuId == null && targetBusinessUnit.GetAttributeValue<EntityReference>(Constant.BusinessUnit.ParentBusinessUnitId) == null)
                    {
                        // the source FetchXML should never have a root BU record - it must be done using the transform config
                        // This is to avoid the scenario when user thinks s/he is running for one environment where the code is running in another environment.
                    }
                    else if (sourceBusinessUnit.GetAttributeValue<string>(Constant.BusinessUnit.Name).Equals(
                        targetBusinessUnit.GetAttributeValue<string>(Constant.BusinessUnit.Name), StringComparison.OrdinalIgnoreCase))
                    {
                        _transformConfig.Transforms.Add(new Transform(Constant.BusinessUnit.EntityLogicalName, Constant.BusinessUnit.Id, sourceBusinessUnit.Id.ToString(), targetBusinessUnit.Id.ToString()));
                    }
                }
            }
        }

        /// <summary>
        /// Find matching security role using BU name and Security Role name. Then map it to transform config (Role, RoleId, SourceRoleId, TargetRoleId).
        /// </summary> 
        public void PopulateSecurityRoleTransforms(EntityCollection sourceSecurityRoles, EntityCollection targetSecurityRoles)
        {
            foreach (var sourceRole in sourceSecurityRoles.Entities)
            {
                var sourceBusinessUnit = sourceRole.GetAttributeValue<EntityReference>(Constant.Role.BusinessUnitId);
                var sourceRoleName = sourceRole.GetAttributeValue<string>(Constant.Role.Name);

                if (!string.IsNullOrWhiteSpace(sourceBusinessUnit.Name))
                {
                    foreach (var targetRole in targetSecurityRoles.Entities)
                    {
                        var targetBusinessUnit = targetRole.GetAttributeValue<EntityReference>(Constant.Role.BusinessUnitId);

                        // if the BU name and Role name match means it's same security role in target
                        if (!string.IsNullOrWhiteSpace(targetBusinessUnit.Name) &&
                            sourceBusinessUnit.Name.Equals(targetBusinessUnit.Name, StringComparison.OrdinalIgnoreCase) &&
                            sourceRoleName.Equals(targetRole.GetAttributeValue<string>(Constant.Role.Name), StringComparison.OrdinalIgnoreCase))
                        {
                            _transformConfig.Transforms.Add(new Transform(Constant.Role.EntityLogicalName, Constant.Role.Id, sourceRole.Id.ToString(), targetRole.Id.ToString()));
                            break;
                        }
                    }
                }
            }
        }

        /// <summary>
        ///  Map Field Security Profile (FSP) between Source and Target Dynamics
        /// </summary>
        /// <param name="sourceFsPs"></param>
        /// <param name="targetFsPs"></param>
        public void PopulateFieldSecurityProfileTransforms(EntityCollection sourceFsPs, EntityCollection targetFsPs)
        {
            foreach (var sourceFsp in sourceFsPs.Entities)
            {
                var matchCount = 0;
                foreach (var targetFsp in targetFsPs.Entities)
                {
                    // source and target guids are same - no benefit of adding in the transform config
                    if (sourceFsp.Id.Equals(targetFsp.Id))
                    {
                        _sameGuidsInSourceAndTarget.Add(sourceFsp.Id);
                        break;
                    }

                    // if the BU name and Role name match means it's same security role in target
                    if (sourceFsp.GetAttributeValue<string>(Constant.FieldSecurityProfile.Name).Equals(
                            targetFsp.GetAttributeValue<string>(Constant.FieldSecurityProfile.Name), StringComparison.OrdinalIgnoreCase))
                    {
                        _transformConfig.Transforms.Add(new Transform(Constant.FieldSecurityProfile.EntityLogicalName, Constant.FieldSecurityProfile.Id, sourceFsp.Id.ToString(), targetFsp.Id.ToString()));
                        //break; don't break - to confirm that there is not multiple record with matching name
                        matchCount++;
                    }
                }

                if (matchCount > 1)
                    throw new Exception($"Duplicate Field Security Profile with same name [ {sourceFsp.GetAttributeValue<string>("name")} ]");

            }
        }

        /// <summary>
        /// Map Currency between Source and Target Dynamics
        /// </summary>
        /// <param name="sourceCurrencies"></param>
        /// <param name="targetCurrencies"></param>
        public void PopulateCurrencyTransforms(EntityCollection sourceCurrencies, EntityCollection targetCurrencies)
        {
            foreach (var sourceCurrency in sourceCurrencies.Entities)
            {
                var sourceCurrencyCode = sourceCurrency.GetAttributeValue<string>(Constant.TransactionCurrency.IsoCurrencyCode);
                foreach (var targetCurrency in targetCurrencies.Entities)
                {
                    // source and target guids are same - no benefit of adding in the transform config
                    if (sourceCurrency.Id.Equals(targetCurrency.Id))
                    {
                        _sameGuidsInSourceAndTarget.Add(sourceCurrency.Id);
                        break;
                    }

                    // if the BU name and Role name match means it's same security role in target
                    if (sourceCurrencyCode.Equals(targetCurrency.GetAttributeValue<string>(Constant.TransactionCurrency.IsoCurrencyCode), StringComparison.OrdinalIgnoreCase))
                    {
                        _transformConfig.Transforms.Add(new Transform(Constant.TransactionCurrency.EntityLogicalName, Constant.TransactionCurrency.Id, sourceCurrency.Id.ToString(), targetCurrency.Id.ToString()));
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Map Team between Source and Target Dynamics
        /// This is specially required for teamrole that has "teamid" which is a "Guid", not an EntityReference - stupid Dynamics data modeller :(
        /// </summary>
        /// <param name="sourceTeams"></param>
        /// <param name="targetTeams"></param>
        public void PopulateTeamTransforms(EntityCollection sourceTeams, EntityCollection targetTeams)
        {
            if (_targetRootBuId.Equals(Guid.Empty) || _sourceRootBuId.Equals(Guid.Empty))
                throw new Exception("Root Business Unit has not been configured yet for Source and Target system in transform config!");

            foreach (var sourceTeam in sourceTeams.Entities)
            {
                var sourceBusinessUnit = sourceTeam.GetAttributeValue<EntityReference>(Constant.Team.BusinessUnitId);

                // find matching BU from replacement config only if the the sourceBusinessUnit.Name is null
                var mappedTargetBuId = (string.IsNullOrWhiteSpace(sourceBusinessUnit.Name) ?
                    TransformObjectValue(Constant.BusinessUnit.EntityLogicalName, Constant.BusinessUnit.Id, sourceBusinessUnit.Id) : Guid.Empty);

                var sourceTeamName = sourceTeam.GetAttributeValue<string>(Constant.Team.Name);
                foreach (var targetTeam in targetTeams.Entities)
                {
                    // source and target guids are same - no benefit of adding in the transform!
                    if (sourceTeam.Id.Equals(targetTeam.Id))
                    {
                        _sameGuidsInSourceAndTarget.Add(sourceTeam.Id);
                        break;
                    }

                    var targetBusinessUnit = targetTeam.GetAttributeValue<EntityReference>(Constant.Team.BusinessUnitId);

                    if (sourceTeamName.Equals(targetTeam.GetAttributeValue<string>(Constant.Team.Name), StringComparison.OrdinalIgnoreCase))
                    {
                        // same team name - now we need to ensure that ie same team of the same business unit

                        if (string.IsNullOrWhiteSpace(sourceBusinessUnit.Name))
                        {
                            // we don't have the BU name from the source team record but we have BU Id of the team record

                            // find matching BU from transform
                            if (!mappedTargetBuId.Equals(Guid.Empty) && mappedTargetBuId.Equals(targetBusinessUnit.Id))
                            {
                                _transformConfig.Transforms.Add(new Transform(Constant.Team.EntityLogicalName, Constant.Team.Id, sourceTeam.Id.ToString(), targetTeam.Id.ToString()));
                                break;
                            }
                        }
                        else
                        {
                            // if the BU name and Team name match with the target record theat means it's same Team in target
                            if (!string.IsNullOrWhiteSpace(targetBusinessUnit.Name) &&
                                sourceBusinessUnit.Name.Equals(targetBusinessUnit.Name, StringComparison.OrdinalIgnoreCase))
                            {
                                _transformConfig.Transforms.Add(new Transform(Constant.Team.EntityLogicalName, Constant.Team.Id, sourceTeam.Id.ToString(), targetTeam.Id.ToString()));
                                break;
                            }
                        }
                    }
                    else
                    {
                        // special condition to map default team of root bu. example:
                        // Source root BU = devclient then default team = "devclient"
                        // Target root BU = testclient then default team = "testclient"
                        // it means if root bu name = team name then it is default team of the target-root-bu
                        if (_sourceRootBuId != null && _targetRootBuId != null &&
                            sourceBusinessUnit.Id.Equals(_sourceRootBuId) && targetBusinessUnit.Id.Equals(_targetRootBuId))
                        {
                            if (sourceTeamName.Equals(sourceBusinessUnit.Name, StringComparison.OrdinalIgnoreCase) &&
                                targetTeam.GetAttributeValue<string>(Constant.Team.Name).Equals(targetBusinessUnit.Name, StringComparison.OrdinalIgnoreCase))
                            {
                                _transformConfig.Transforms.Add(new Transform(Constant.Team.EntityLogicalName, Constant.Team.Id, sourceTeam.Id.ToString(), targetTeam.Id.ToString()));

                                // make the team name in source same as target so that the default team of Root BU is NOT renamed!
                                sourceTeam[Constant.Team.Name] = targetTeam[Constant.Team.Name];
                                break;
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Map Queues between Source and Target Dynamics
        /// </summary>
        /// <param name="sourceQueues"></param>
        /// <param name="targetQueues"></param>
        public void PopulateQueueTransforms(EntityCollection sourceQueues, EntityCollection targetQueues)
        {
            foreach (var sourceQueue in sourceQueues.Entities)
            {
                foreach (var targetQueue in targetQueues.Entities)
                {
                    if (sourceQueue.Id.Equals(targetQueue.Id))
                    {
                        _sameGuidsInSourceAndTarget.Add(sourceQueue.Id);
                        break;
                    }

                    // if the BU name and Role name match means it's same security role in target
                    if (sourceQueue.GetAttributeValue<string>(Constant.Queue.Name).Equals(
                            targetQueue.GetAttributeValue<string>(Constant.Queue.Name), StringComparison.OrdinalIgnoreCase))
                    {
                        _transformConfig.Transforms.Add(new Transform(Constant.Queue.EntityLogicalName, Constant.Queue.Id, sourceQueue.Id.ToString(), targetQueue.Id.ToString()));
                        break;
                    }
                }
            }
        }
    }
}
