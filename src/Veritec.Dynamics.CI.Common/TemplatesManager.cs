using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace MsCrmTools.DocumentTemplatesMover
{
    /// <summary>
    /// Code sourced from https://github.com/MscrmTools/MscrmTools.DocumentTemplatesMover and then modified to improve performance
    /// Original code credit goes to Magnetism - https://github.com/NZxRMGuy/crm-documentextractor. Reused here with permission from NZxRMGuy under MIT License.
    /// </summary>
    internal class TemplatesManager
    {
        public int? GetEntityTypeCode(IOrganizationService service, string entity)
        {
            var request = new RetrieveEntityRequest
            {
                LogicalName = entity,
                EntityFilters = EntityFilters.Entity
            };

            var response = (RetrieveEntityResponse)service.Execute(request);
            var metadata = response.EntityMetadata;

            return metadata.ObjectTypeCode;
        }

        public void ReRouteEtcViaOpenXml(Entity template, string entityName, int? newEtc)
        {
            if (template == null || string.IsNullOrWhiteSpace(entityName))
                return;

            var content = Convert.FromBase64String(template.GetAttributeValue<string>("content"));
            var contentStream = new MemoryStream();
            contentStream.Write(content, 0, content.Length);
            contentStream.Position = 0;

            var toFind = $@"urn:microsoft-crm\/document-template\/{entityName}\/(.*?)(?:\/|$)";
            var replaceWith = $"urn:microsoft-crm/document-template/{entityName}/{newEtc}/";
                                                
            //this caused big headache - var doc = WordprocessingDocument.Open(contentStream, true, new OpenSettings { AutoSave = true });
            var hasEntityTypeCodeChanged = false;

            using (var doc = WordprocessingDocument.Open(contentStream, true, new OpenSettings { AutoSave = true }))
            {
                var updatedInnerXml = Regex.Replace(doc.MainDocumentPart.Document.InnerXml, toFind, replaceWith);
                hasEntityTypeCodeChanged = !doc.MainDocumentPart.Document.InnerXml.Equals(updatedInnerXml, StringComparison.OrdinalIgnoreCase);

                //proceed only if there is no change of InnerXml means the EntityTypeCode has been changed - to improve performance
                if (hasEntityTypeCodeChanged) 
                {
                    // crm keeps the etc in multiple places; parts here are the actual merge fields
                    doc.MainDocumentPart.Document.InnerXml = updatedInnerXml;

                    // next is the actual namespace declaration
                    foreach (var curDocPart in doc.MainDocumentPart.CustomXmlParts.ToList())
                    {
                        using (var reader = new StreamReader(curDocPart.GetStream()))
                        {
                            var xml = XDocument.Load(reader);

                            // crappy way to replace the xml, couldn't be bothered figuring out xml root attribute replacement...
                            var crappy = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n" + xml;

                            if (Regex.IsMatch(crappy, toFind)) // only replace what is needed
                            {
                                crappy = Regex.Replace(crappy, toFind, replaceWith);

                                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(crappy)))
                                {
                                    curDocPart.FeedData(stream);
                                }
                            }
                        }
                    }
                }
            }

            //proceed only if there is no change of InnerXml means the EntityTypeCode has been changed 
            if (hasEntityTypeCodeChanged)
            {
                template["content"] = Convert.ToBase64String(contentStream.ToArray());
                template["associatedentitytypecode"] = newEtc;
            }
        }

        public Guid TemplateExists(IOrganizationService service, string name)
        {
            var result = Guid.Empty;

            var qe = new QueryExpression("documenttemplate");
            qe.Criteria.AddCondition("status", ConditionOperator.Equal, false); // only get active templates
            qe.Criteria.AddCondition("name", ConditionOperator.Equal, name);

            var results = service.RetrieveMultiple(qe);
            if (results?.Entities != null && results.Entities.Count > 0)
            {
                result = results[0].Id;
            }

            return result;
        }
    }
}
