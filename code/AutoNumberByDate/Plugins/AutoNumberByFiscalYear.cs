using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using System.ServiceModel;

namespace AutoNumberByDate.Plugins
{
    public class AutoNumberByFiscalYear : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the tracing service
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.  
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));


            // The InputParameters collection contains all the data passed in the message request.  
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.  
                Entity entity = (Entity)context.InputParameters["Target"];

                // Obtain the IOrganizationService instance which you will need for web service calls.  
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                // Passing null into CreateOrganizationService will get the system organization service which will have access to the fiscal year counter by default.
                IOrganizationService service = serviceFactory.CreateOrganizationService(null);

                try
                {
                    SetNameFieldWithFiscalYearNumber(entity, service);
                }

                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in AutoNumberByFiscalYear.", ex);
                }

                catch (Exception ex)
                {
                    tracingService.Trace("AutoNumberByFiscalYear: {0}", ex.ToString());
                    throw;
                }
            }
        }

        private static void SetNameFieldWithFiscalYearNumber(Entity entity, IOrganizationService service)
        {
            string controlNumber = "0001";
            int currentNumber = 1;
            // Get the current date in UTC.
            DateTime currentDate = DateTime.UtcNow.Date;

            // Look for an existing fiscal year counter entity
            QueryExpression qExpression = new QueryExpression()
            {
                EntityName = "raw_fiscalyearcounter",
                ColumnSet = new ColumnSet("raw_name", "raw_lastnumber"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("raw_start", ConditionOperator.LessEqual, currentDate),
                        new ConditionExpression("raw_end", ConditionOperator.GreaterEqual, currentDate)
                    }
                }
            };
            EntityCollection results = service.RetrieveMultiple(qExpression);
            Entity counter = results.Entities.FirstOrDefault();

            if (counter != null)
            {
                // If the counter is there update it to lock it in the transaction.
                Entity blocker = new Entity() { LogicalName = "raw_fiscalyearcounter", Id = counter.Id };
                blocker["raw_name"] = counter.GetAttributeValue<string>("raw_name");
                service.Update(blocker);

                // Now that the counter is locked retrieve the counter again so that we can get the current number and udpate it.
                QueryExpression qExpression2 = new QueryExpression()
                {
                    EntityName = "raw_fiscalyearcounter",
                    ColumnSet = new ColumnSet("raw_lastnumber", "raw_name"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("raw_start", ConditionOperator.LessEqual, currentDate),
                            new ConditionExpression("raw_end", ConditionOperator.GreaterEqual, currentDate)
                        }
                    }
                };

                // DO NOT refactor RetrieveMultipleRequest to RetrieveMultiple otherwise some caching will occur.
                RetrieveMultipleRequest rmRequest = new RetrieveMultipleRequest() { Query = qExpression2 };
                RetrieveMultipleResponse lockedResults = (RetrieveMultipleResponse)service.Execute(rmRequest);
                Entity lockedCounter = lockedResults.EntityCollection.Entities.First();

                Entity counterUpdater = new Entity() { LogicalName = "raw_fiscalyearcounter", Id = counter.Id };

                // Get the current number and increment it.
                currentNumber = lockedCounter.GetAttributeValue<int>("raw_lastnumber");
                currentNumber = ++currentNumber;
                controlNumber = (currentNumber).ToString().PadLeft(4, '0');

                // Update the counter with the new number.
                counterUpdater["raw_lastnumber"] = currentNumber;
                service.Update(counterUpdater);

                // Set the control number on the invoice {FY}{Sequence #}".
                entity["raw_name"] = $"{lockedCounter.GetAttributeValue<string>("raw_name")}{controlNumber}";
            }
            else
            {
                //if no counter was found thow an error
                throw new InvalidPluginExecutionException("No counter was found for the current fiscal year. Please contact your system administrator.");
            }
        }
    }
}
