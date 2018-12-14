using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace Veritec.Dynamics.CI.Common
{
    public class DataExport : CiBase
    {
        public Dictionary<string, string> FetchXmlQueries;
        public Dictionary<string, string> FetchXmlQueriesResultXml;
        
        public DataExport(string dynamicsConnectionString, string fetchXmlFile) : base(dynamicsConnectionString)
        {
            InitDataExport(fetchXmlFile);
        }

        public DataExport(CrmParameter crmParameter, string fetchXmlFile) : base(crmParameter)
        {
            InitDataExport(fetchXmlFile);
        }

        public void InitDataExport(string fetchXmlFile)
        { 
            FetchXmlQueries = new Dictionary<string, string>();
            FetchXmlQueriesResultXml = new Dictionary<string, string>();

            IEnumerable<XElement> xElements;

            try {
                xElements = XElement.Load(fetchXmlFile).Elements("FetchXMLQuery");
            }
            catch (Exception ex)
            {
                throw new Exception(
                    $"Error occurred while loading Fetch Query from: {Environment.CurrentDirectory}\\{fetchXmlFile}\r\n{ex.Message}\r\n{ex.StackTrace}");
            }

            var enumerable = xElements as IList<XElement> ?? xElements.ToList();

            if (enumerable.Any())
            {
                foreach (var item in enumerable)
                {
                    var order = item.Attribute("order");
                    string orderval;
                    if (order == null) orderval = "";
                    else orderval = order.Value + "-";

                    var xAttribute = item.Attribute("name");

                    if (xAttribute != null)
                        FetchXmlQueries.Add(orderval + xAttribute.Value, item.FirstNode.ToString());
                    else
                    {
                        throw new Exception("FetchXML name field is null");
                    }
                }
            }
            else
            {
                throw new Exception("No FetchXML queries found in file: " + Environment.CurrentDirectory + fetchXmlFile);
            }
        }

        public StringBuilder ExecuteFetchXmlQuery(string entityName, string fetchXml, string[] systemFieldsToExclude, bool forComparison = false)
        {
            var queryResultXml = new StringBuilder();

            if (string.IsNullOrWhiteSpace(entityName) || string.IsNullOrWhiteSpace(fetchXml))
                return queryResultXml;

            // convert fetchxml to query expression
            var query =  ConvertFetchXmlToQueryExpression(fetchXml);

            var systemFields = new List<string>();
            if (systemFieldsToExclude == null)
            {
                // match this with "ColumnsToExcludeToCompareData" in app.config - this is included here in case there is no exclusion in the fetch!!!
                const string columnsToExclude = "languagecode;createdon;createdby;modifiedon;modifiedby;owningbusinessunit;owninguser;owneridtype;" +
                                                "importsequencenumber;overriddencreatedon;timezoneruleversionnumber;operatorparam;utcconversiontimezonecode;versionnumber;" +
                                                "customertypecode;matchingentitymatchcodetable;baseentitymatchcodetable;slaidunique;slaitemidunique";

                systemFields.AddRange(columnsToExclude.Split(';'));
            }
            else
            {
                systemFields.AddRange(systemFieldsToExclude);
            }
            systemFields.Sort();

            var sourceDataLoader = new DataLoader(OrganizationService);
            var entityMetadata = sourceDataLoader.GetEntityMetaData(query.EntityName, EntityFilters.Attributes);

            // exclude all fields tha are in the system fields or IsValidForRead=false or isValidForCreate = false
            if (query.ColumnSet.AllColumns)
            {
                query.ColumnSet.AllColumns = false;
                foreach (var attributeMetadata in entityMetadata.Attributes)
                {
                    if (systemFields.Contains(attributeMetadata.LogicalName) || attributeMetadata.IsValidForRead == false ||
                        attributeMetadata.IsValidForCreate == false)
                    {
                        // ignore system fields or those can't be read or used for created
                    }
                    else
                    {
                        query.ColumnSet.AddColumn(attributeMetadata.LogicalName);
                    }
                }
            }

            var queryResults = OrganizationService.RetrieveMultiple(query);
            if (queryResults.Entities.Count == 0) return queryResultXml;

            queryResultXml.AppendLine($"<EntityData Name='{entityName}'>");

            foreach (var queryResult in queryResults.Entities)
            {
                if (!forComparison)
                {
                    foreach (var attributeMetadata in entityMetadata.Attributes)
                    {
                        if (systemFields.Contains(attributeMetadata.LogicalName) ||
                            attributeMetadata.IsValidForCreate == false)
                        {
                            // these columns should not be included (even if they are included in the query)
                        }
                        else if (query.ColumnSet != null)
                        {
                            // if the fetchXml contains the column but the column is not in the queryResult because it has null value
                            // then we need to add the column in the queryResult so that the column-value is replaced in the target
                            if (query.ColumnSet.Columns.Contains(attributeMetadata.LogicalName) &&
                                !queryResult.Attributes.Contains(attributeMetadata.LogicalName))
                            {
                                queryResult.Attributes.Add(attributeMetadata.LogicalName, null);
                            }
                        }
                    }
                }


                if ((queryResult.RowVersion != null) && forComparison) queryResult.RowVersion = null;

                var typeList = new List<Type> { queryResult.GetType() };
                queryResultXml.AppendLine(DataContractSerializeObject(queryResult, typeList));
            }

            queryResultXml.AppendLine("</EntityData>");

            return queryResultXml;
        }

        public QueryExpression ConvertFetchXmlToQueryExpression(string fetchXml)
        {
            var convertRq = new FetchXmlToQueryExpressionRequest
            {
                FetchXml = fetchXml
            };

            var convertedQuery = OrganizationService.Execute(convertRq);
            return ((FetchXmlToQueryExpressionResponse) convertedQuery)?.Query;
        }

        public void SaveFetchXmlQueryResultstoDisk(string dataFolder)
        {
            Directory.CreateDirectory(dataFolder);

            if (FetchXmlQueriesResultXml == null) return;

            foreach (var x in FetchXmlQueriesResultXml)
            {
                File.WriteAllText(dataFolder + x.Key + ".xml", x.Value);
            }
        }

        public static void SaveFetchXmlQueryResulttoDisk(string dataFolder, string queryName, string fetchXmlQueryResultXml)
        {
            Directory.CreateDirectory(dataFolder);

            if (!string.IsNullOrWhiteSpace(fetchXmlQueryResultXml))
            {
                File.WriteAllText(dataFolder + queryName + ".xml", fetchXmlQueryResultXml);
            }
        }

        /// <summary>
        /// persist object to disk as xml
        /// reference: https://dzone.com/articles/using-datacontractserializer
        /// </summary>
        public static string DataContractSerializeObject<T>(T objectToSerialize, List<Type> types)
        {
            using (var output = new StringWriter())
            {
                using (var writer = new XmlTextWriter(output) { Formatting = Formatting.Indented })
                {
                    var dataContractSerializer = new DataContractSerializer(typeof(T), types);
                    dataContractSerializer.WriteObject(writer, objectToSerialize);
                    return output.GetStringBuilder().ToString();
                }
            }
        }
    }
}
