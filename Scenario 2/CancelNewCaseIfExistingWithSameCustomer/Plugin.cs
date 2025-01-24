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
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                tracer.Trace("CancelNewCaseIfExistingWithSameCustomer: Beginning Execution");

                tracer.Trace("Getting new case");
                Entity newCase = GetEntity(context);

                tracer.Trace("Building query");
                QueryExpression query = BuildQuery(((EntityReference)newCase["customerid"]).Id.ToString());

                tracer.Trace("Checking existing cases with same customer");
                if (CustomerHasExistingCases(service, query, tracer))
                    throw new InvalidPluginExecutionException("An existing case belonging to this customer already exists. Check the existing case and either make changes to it or close it before opening a new case.");
            }
            catch (InvalidPluginExecutionException e)
            {
                tracer.Trace("Caught a plugin exception: " + e.Message);
                throw;
            }
            catch (Exception e)
            {
                tracer.Trace("Caught an unexpected exception: " + e.Message);
                throw new InvalidPluginExecutionException("An unexpected error occurred: " + e.Message);
            }
        }

        private Entity GetEntity(IPluginExecutionContext context)
        {
            if (!context.InputParameters.ContainsKey("Target"))
                throw new InvalidPluginExecutionException("The target cannot be found.");

            Entity newCase = (Entity)context.InputParameters["Target"];
            if (!newCase.Attributes.ContainsKey("customerid"))
                throw new InvalidPluginExecutionException("Could not retrieve customer from case");

            return newCase;
        }

        private QueryExpression BuildQuery(string customerId) 
        {
            QueryExpression query = new QueryExpression("incident")
            {
                TopCount = 2,
                ColumnSet = new ColumnSet("customerid")
            };
            query.Criteria.AddCondition("customerid", ConditionOperator.Equal, customerId);

            return query;
        }

        private bool CustomerHasExistingCases(IOrganizationService service, QueryExpression query, ITracingService tracer)
        {
            try
            {
                EntityCollection collection = service.RetrieveMultiple(query);
                int numEntities = collection.Entities.Count;
                tracer.Trace("Existing cases: " + numEntities);
                return numEntities > 0;
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException("Error retrieving existing cases: " + e.Message);
            }
        }
    }
}
