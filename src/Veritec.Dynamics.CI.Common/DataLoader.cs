using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace Veritec.Dynamics.CI.Common
{
    public class DataLoader
    {
        public IOrganizationService OrganizationService;

        public DataLoader(IOrganizationService organizationService) 
        {
            OrganizationService = organizationService;
        }

        public Dictionary<string, EntityMetadata> GetAllEntitiesMetaData(EntityFilters entityFilter)
        {
            var retrievesEntitiesRequest = new RetrieveAllEntitiesRequest
            {
                RetrieveAsIfPublished = true,
                EntityFilters = entityFilter
            };

            var retrieveEntityResponse = (RetrieveAllEntitiesResponse)OrganizationService.Execute(retrievesEntitiesRequest);
            return retrieveEntityResponse.EntityMetadata.ToDictionary(curMetadata => curMetadata.LogicalName);
        }

        /// <summary>
        /// source: https://crmpolataydin.wordpress.com/2014/12/02/crm-20112013-get-entity-metadata-from-c/
        /// </summary>
        /// <param name="entityName"></param>
        /// <param name="entityFilter"></param>
        /// <returns></returns>
        public EntityMetadata GetEntityMetaData(string entityName, EntityFilters entityFilter)
        {
            var retrievesEntityRequest = new RetrieveEntityRequest
            {
                LogicalName = entityName,
                RetrieveAsIfPublished = true,
                EntityFilters = entityFilter
            };

            var retrieveEntityResponse = (RetrieveEntityResponse)OrganizationService.Execute(retrievesEntityRequest);
            var entityMetadata = retrieveEntityResponse.EntityMetadata;

            return entityMetadata;
        }

        public EntityCollection GetAllEntity(string entityName, string[] columns, LogicalOperator filterOperator = LogicalOperator.And, 
            ConditionExpression[] conditions = null)
        {
            var query = new QueryExpression(entityName)
            {
                NoLock = true,
                Distinct = false,
                ColumnSet = columns == null ? new ColumnSet(true) : new ColumnSet(columns)
            };

            if (conditions == null)
                return OrganizationService.RetrieveMultiple(query);

            query.Criteria.FilterOperator = filterOperator;
            query.Criteria.Conditions.AddRange(conditions);

            return OrganizationService.RetrieveMultiple(query);
        }
    }
}
