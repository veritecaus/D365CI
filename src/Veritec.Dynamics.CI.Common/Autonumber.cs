using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Veritec.Dynamics.CI.Common
{
    public class Autonumber : CiBase
    {
        public Autonumber(CrmParameter crmParameter) : base(crmParameter)
        {

        }

        public void SetSeed(string entityName, string attributeName, long value, bool force = false)
        {
            SetAutoNumberSeedRequest req = new SetAutoNumberSeedRequest
            {
                EntityName = entityName,
                AttributeName = attributeName,
                Value = value
            };

            OrganizationService.Execute(req);
        }

        /// <summary>
        /// Check if there is data in a target entity
        /// </summary>
        /// <param name="entityName"></param>
        /// <returns>returns count of records that exist</returns>
        public int TargetEntityRowCount(string entityName, string attributeName)
        {
            // using fetch as it's the only query type that supports aggregate functions.
            var fetchXMLCount = $@"<fetch aggregate='true'>
                                     <entity name = '{entityName}'>
                                       <attribute name = '{attributeName}' alias='RecordCount' aggregate='count'/>
                                     </entity >
                                   </fetch>";

            var fetchResult = OrganizationService.RetrieveMultiple(new FetchExpression(fetchXMLCount));
            if (fetchResult.Entities.Count < 1) return 0;
            var recordCount = (int)((Microsoft.Xrm.Sdk.AliasedValue)fetchResult[0]["RecordCount"]).Value;
            return recordCount;

        }
    }
}
