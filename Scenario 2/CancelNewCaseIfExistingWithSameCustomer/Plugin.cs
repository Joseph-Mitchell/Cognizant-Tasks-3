using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace CancelNewCaseIfExistingWithSameCustomer
{
    public class Plugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationService service = (IOrganizationService)serviceProvider.GetService(typeof(IOrganizationService));
            ITracingService tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                Entity newCase = GetEntity(context);
                QueryExpression query = BuildQuery((string)newCase["customerid"]);

                if (!CustomerHasExistingCases(service, query))
                    throw new InvalidPluginExecutionException("An existing case belonging to this customer already exists. Check the existing case and either make changes to it or close it before opening a new case.");
            }
            catch (InvalidPluginExecutionException e)
            {
                tracer.Trace("Caught a plugin exception: " + e.Message);
                throw e;
            }
            catch (Exception e)
            {
                tracer.Trace("Caught an unexpected exception: " + e.Message);
                throw new InvalidPluginExecutionException(e.Message);
            }
        }

        private Entity GetEntity(IPluginExecutionContext context)
        {
            if (!context.InputParameters.ContainsKey("Target"))
                throw new InvalidPluginExecutionException("The target cannot be found.");

            Entity newCase = (Entity)context.InputParameters["Target"];
            if (!newCase.Attributes.ContainsKey("customerid"))
                throw new InvalidPluginExecutionException("Could not retrieve cutomer from case");

            return newCase;
        }

        private QueryExpression BuildQuery(string cutsomerId) 
        {
            QueryExpression query = new QueryExpression("account")
            {
                TopCount = 1,
                ColumnSet = new ColumnSet()
            };
            query.Criteria.AddCondition("customerid", ConditionOperator.Equal, cutsomerId);

            return query;
        }

        private bool CustomerHasExistingCases(IOrganizationService service, QueryExpression query)
        {
            try
            {
                int numEntities = service.RetrieveMultiple(query).TotalRecordCount;
                return numEntities > 0;
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException("Error retrieving existing cases");
            }
        }
    }
}
