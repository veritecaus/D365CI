namespace Veritec.Dynamics.CI.Common
{
    /// <summary>
    /// Static constants, for example Dynamics entity and attribute names.
    /// </summary>
    public static class Constant
    {
        /// <summary>
        /// Constants that are common to all entities
        /// </summary>
        public static class Entity
        {
            public const string Name = "name";
            public const string OwnerId = "ownerid";
            public const string StateCode = "statecode";
            public const string StatusCode = "statuscode";
        }

        public static class BusinessUnit
        {
            public const string EntityLogicalName = "businessunit";
            public const string Id = "businessunitid";
            public const string Name = "name";
            public const string ParentBusinessUnitId = "parentbusinessunitid";
        }

        public static class Organization
        {
            public const string EntityLogicalName = "organization";
            public const string Id = "organizationid";
            public const string Name = "name";
        }

        public static class Team
        {
            public const string EntityLogicalName = "team";
            public const string Id = "teamid";
            public const string Name = "name";
            public const string AdministratorId = "administratorid";
            public const string BusinessUnitId = "businessunitid";
            public const string TeamType = "teamtype";
            public const string IsDefault = "isdefault"; //0 = No
        }

        public static class Role
        {
            public const string EntityLogicalName = "role";
            public const string Id = "roleid";
            public const string Name = "name";
            public const string BusinessUnitId = "businessunitid";
        }

        public static class FieldSecurityProfile
        {
            public const string EntityLogicalName = "fieldsecurityprofile";
            public const string Id = "fieldsecurityprofileid";
            public const string Name = "name";
        }

        public static class TeamRoles
        {
            public const string EntityLogicalName = "teamroles";
            public const string Id = "teamroleid";
            public const string TeamId = "teamid";
            public const string RoleId = "roleid";
        }

        public static class TeamProfiles
        {
            public const string EntityLogicalName = "teamprofiles";
            public const string Id = "teamprofilesid";
            public const string TeamId = "teamid";
            public const string FieldSecurityProfileId = "fieldsecurityprofileid";
        }
        
        public static class DocumentTemplate
        {
            public const string EntityLogicalName = "documenttemplate";
            public const string Name = "name";
            public const string AssociatedEntityTypeCode = "associatedentitytypecode";
        }

        public static class DuplicateRule
        {
            public const string EntityLogicalName = "duplicaterule";
            public const string Id = "duplicateruleid";
            public const string BaseEntityTypeCode = "baseentitytypecode";
            public const string MatchingEntityTypeCode = "matchingentitytypecode";
        }

        public static class DuplicateRuleCondition
        {
            public const string EntityLogicalName = "duplicaterulecondition";
            public const string Id = "duplicateruleconditionid";
        }

        public static class Workflow
        {
            public const string EntityLogicalName = "workflow";
        }

        public static class Sla
        {
            public const string EntityLogicalName = "sla";
        }

        public static class SlaItem
        {
            public const string EntityLogicalName = "slaitem";
        }

        public static class TransactionCurrency
        {
            public const string EntityLogicalName = "transactioncurrency";
            public const string Id = "transactioncurrencyid";
            public const string ExchangeRate = "exchangerate";
            public const string CurrencyName = "currencyname";
            public const string IsoCurrencyCode = "isocurrencycode";
        }

        public static class Queue
        {
            public const string EntityLogicalName = "queue";
            public const string Id = "queueid";
            public const string Name = "name";
            public const string TransactionCurrencyId = "transactioncurrencyid";
            public const string OwnerId = "ownerid";
        }

        public static class Mailbox
        {
            public const string EntityLogicalName = "mailbox";
            public const string Id = "mailboxid";
            public const string Name = "name";
            public const string OwnerId = "ownerid";       
            public const string RegardingObjectId = "regardingobjectid";
        }

        public static class User
        {
            public const string EntityLogicalName = "systemuser";
            public const string Id = "systemuserid";
            public const string Name = "name";
            public const string DomainName = "domainname";
        }
    }
}
