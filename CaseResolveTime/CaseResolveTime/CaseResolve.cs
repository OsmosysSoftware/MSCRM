using System;
//sdk namespace
using Microsoft.Xrm.Sdk;
using System.IdentityModel;
using Microsoft.Xrm.Sdk.Query;

namespace CaseResolveTime
{
    public class CaseResolve : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {

			//declare service objects.
			IPluginExecutionContext context;
			IOrganizationServiceFactory serviceFactory;
			IOrganizationService service;
            
			//get the context.
            context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = serviceFactory.CreateOrganizationService(context.UserId);
           

			Entity targetCase = new Entity("incident");
            string strResolution = string.Empty;
            int intTotalTime = -1;
            int intTotalBillableTime = -1;
            string strRemarks = string.Empty;


            //get Target.
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                //get the entity.
                Entity target = (Entity)context.InputParameters["Target"];
                
                if (target.LogicalName != "incidentresolution")
                    return;

                try
                {
                    if(target.Contains("incidentid"))
                    {
                       //get related case.
                        targetCase.Id = ((EntityReference)target["incidentid"]).Id;

                        //capture case resolution fields.
                        strResolution = target.Contains("subject") ? target["subject"].ToString() : string.Empty;
                        //intTotalBillableTime = target.Contains("timespent") ? (Int32)target["timespent"] : 0;
                        intTotalBillableTime = GetTotalBillableTime(service, targetCase.Id);
                        strRemarks = target.Contains("description") ? target["description"].ToString() : string.Empty;

                        //get total time for activities.
                        intTotalTime=GetTotalTime(service,targetCase.Id);

                        //update Case with the fields
                        targetCase["tlg_resolution"] = strResolution;
                        targetCase["tlg_billabletime"] = intTotalBillableTime.ToString();
                        targetCase["tlg_totaltime"] = intTotalTime.ToString();
                        targetCase["tlg_remarks"] = strRemarks;
                        service.Update(targetCase);
                    }

                }
                catch(Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message);
                }                
            }

        }

        //method to get total time
        private int GetTotalTime(IOrganizationService service, Guid guidRelatedCaseId)
        {
            //count the Activity Actual Duration Minutes for this Case
            //need to sum time spent of each activity (cannot directly using the actualtdurationminutes) 
 
            int intSumTotalTime = 0;
 
            //Retrieve all related Activities by Case
            QueryExpression query = new QueryExpression("activitypointer");
            query.ColumnSet.AddColumns("actualdurationminutes");
 
            query.Criteria = new FilterExpression();
            query.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, guidRelatedCaseId);
 
            // Execute the Query 
            EntityCollection results = service.RetrieveMultiple(query);
 
            foreach (Entity entity in results.Entities)
            {
                int intActivityTime = 0;
                intActivityTime = entity.Contains("actualdurationminutes") ? (Int32)entity["actualdurationminutes"] : 0;
                intSumTotalTime = intSumTotalTime + intActivityTime;
            }
 

           

            return intSumTotalTime;
        }

		//method to get total billable time for a case for each resolution and reopen.
        private int GetTotalBillableTime(IOrganizationService service, Guid guidRelatedCaseId)
        {
            //retrive all billable hours.
            int intTotalBillable = 0;

            QueryExpression billquery = new QueryExpression("incidentresolution");
            billquery.ColumnSet.AddColumns("timespent");

            billquery.Criteria = new FilterExpression();
            billquery.Criteria.AddCondition("incidentid", ConditionOperator.Equal, guidRelatedCaseId);

            //Execute query.
            EntityCollection entCol = service.RetrieveMultiple(billquery);

            foreach (Entity ent in entCol.Entities)
            {
               
                intTotalBillable += ent.Contains("timespent") ? (Int32)ent["timespent"] : 0;
               

            }

            return intTotalBillable;
        }
    }
    
}
