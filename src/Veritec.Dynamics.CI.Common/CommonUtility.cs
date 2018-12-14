using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Veritec.Dynamics.CI.Common
{
    public class CommonUtility
    {
        public static void StartTraceLogger(string logFileName)
        {
            Trace.Listeners.Clear();

            var twtl = new TextWriterTraceListener(logFileName)
            {
                Name = "DeploymentUtilityLogger",
                TraceOutputOptions = TraceOptions.ThreadId | TraceOptions.DateTime
            };

            var ctl = new ConsoleTraceListener(false) { TraceOutputOptions = TraceOptions.DateTime };
            Trace.Listeners.Add(twtl);
            Trace.Listeners.Add(ctl);
            Trace.AutoFlush = true;
        }

        public static string GetUserFriendlyMessage(Entity entity)
        {
            if (entity?.Attributes == null || entity.Attributes.Count == 0)
                return string.Empty;

            var nameAttribute = entity.Attributes.First();
            var nameAttributeFound = true;
            if (nameAttribute.Key.IndexOf("name", StringComparison.OrdinalIgnoreCase) == -1)
            {
                nameAttributeFound = false;
                foreach (var attribute in entity.Attributes)
                {
                    if (attribute.Key.Equals("name", StringComparison.OrdinalIgnoreCase) ||
                        attribute.Key.Equals("fullname", StringComparison.OrdinalIgnoreCase) ||
                        attribute.Key.Equals("baseattributename", StringComparison.OrdinalIgnoreCase) || //duplicaterulecondition
                        attribute.Key.IndexOf("_name", StringComparison.OrdinalIgnoreCase) > 0)
                    {
                        nameAttribute = attribute;
                        nameAttributeFound = true;
                        break;
                    }
                }
            }

            //include first column value only when it is either string or premitive type
            if (nameAttribute.Key == null || nameAttribute.Value == null || !nameAttributeFound)
            {
                return entity.LogicalName + " " + entity.Id;
            }

            var nameAttributeValue = (!nameAttribute.Value.GetType().IsGenericType &&
                                      (nameAttribute.Value.GetType().IsPrimitive || nameAttribute.Value is string) ?
                nameAttribute.Value.ToString() : "");

            return entity.LogicalName + " " + entity.Id + " " + nameAttributeValue;
        }

        /// <summary>
        /// Waits for async job to complete
        /// </summary>
        /// <param name="organizationService"></param>
        /// <param name="asyncJobIds"></param>
        /// <remarks>Source: https://msdn.microsoft.com/en-us/library/hh547456.aspx </remarks>
        public string WaitForAsyncJobCompletion(IOrganizationService organizationService, IEnumerable<Guid> asyncJobIds)
        {
            var asyncJobList = new List<Guid>(asyncJobIds);
            var cs = new ColumnSet(Constant.Entity.StateCode, "asyncoperationid");
            var retryCount = 100;
            var message = new StringBuilder();

            while (asyncJobList.Count != 0 && retryCount > 0)
            {
                // Retrieve the async operations based on the ids
                var crmAsyncJobs = organizationService.RetrieveMultiple(
                    new QueryExpression("asyncoperation")
                    {
                        ColumnSet = cs,
                        Criteria = new FilterExpression
                        {
                            Conditions =
                            {
                                new ConditionExpression("asyncoperationid", ConditionOperator.In, asyncJobList.ToArray())
                            }
                        }
                    });

                // Check to see if the operations are completed and if so remove them from the Async Guid list
                foreach (var item in crmAsyncJobs.Entities)
                {
                    var crmAsyncJob = item;
                    if (crmAsyncJob.GetAttributeValue<OptionSetValue>(Constant.Entity.StateCode).Value == 3) //AsyncOperationState.Completed
                        asyncJobList.Remove(crmAsyncJob.Id);

                    message.AppendLine(String.Concat("Async operation state is ",
                        crmAsyncJob.GetAttributeValue<OptionSetValue>(Constant.Entity.StateCode).Value.ToString(),
                        ", async operation id: ", crmAsyncJob.Id.ToString()));
                }

                // If there are still jobs remaining, sleep the thread.
                if (asyncJobList.Count > 0)
                    Thread.Sleep(2000);

                retryCount--;
            }

            if (retryCount == 0 && asyncJobList.Count > 0)
            {
                foreach (var asyncJob in asyncJobList)
                {
                    message.AppendLine($"The following async operation has not completed: {asyncJob.ToString()}");
                }
            }
            return message.ToString();
        }
    }
}
