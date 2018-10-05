using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ELComePlugin.FieldService.ServiceOrder
{
    public class JobSheetTimingPlugin : IPlugin
    {
        public static List<JobSheetFieldMapping> mappingList = null;
        ITracingService tracingService = null;
        IPluginExecutionContext context = null;
        IOrganizationService orgService = null;
        IOrganizationServiceFactory orgFactory = null;
        //If Service ORder has no Legal Entity Value, This Legal Entity will be used to fetch Categories
        public string DefaultLegalEntity = "7A41C8F1-7B94-E711-80E5-3863BB34FA68";
        //Category Id is Static in List of JobSheetTimingFieldMapping Class. 
        //Legal Entity Filter is needed because, Based on Category id, there are multiple Categories.
        public string GetCategoriesFetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='elc_projectcategory'>
                                <attribute name='elc_projectcategoryid' />
                                <attribute name='elc_name' />
                                <filter type='and'>
                                  <condition attribute='elc_categoryid' operator='eq' value='{0}' />
                                  <condition attribute='elc_legalentitycompany' operator='eq' value='{1}' />
                                </filter>
                              </entity>
                            </fetch>";
        public string GetServiceHoursForUpdate = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                      <entity name='msdyn_workorderservice'>
                                                        <attribute name='msdyn_workorderserviceid' />
                                                        <attribute name='msdyn_duration' />
                                                        <order attribute='msdyn_duration' descending='false' />
                                                        <filter type='and'>
                                                          <condition attribute='elc_projectcategory' operator='eq' value='{0}' />
                                                          <condition attribute='msdyn_workorder' operator='eq' value='{1}' />
                                                        </filter>
                                                      </entity>
                                                    </fetch>";
        public const string ProjectCategoryEntitySchemaName = "elc_projectcategory";
        public const string ServiceHourEntitySchemaName = "msdyn_workorderservice";
        public const string LegalEntityEntitySchemaName = "elc_legalentitycompany";
        public void Execute(IServiceProvider serviceProvider)
        {
            #region Fill Mapping List
            mappingList = new List<JobSheetFieldMapping>();
            //Travelling To Vessel
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_traveltovslnormaltime", ServiceCategoryId = "SER-TNH-S", AutomationCategoryId = "ASP-TNH-S", GroupName = "TNH" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_traveltovslovertime", ServiceCategoryId = "SER-TOH-S", AutomationCategoryId = "ASP-TOH-S", GroupName = "TOH" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_traveltovslholidaytime", ServiceCategoryId = "SER-THH-S", AutomationCategoryId = "ASP-THH-S", GroupName = "THH" });
            //Travelling To Vessel 1
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_traveltovslnormaltime1", ServiceCategoryId = "SER-TNH-S", AutomationCategoryId = "ASP-TNH-S", GroupName = "TNH" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_traveltovslovertime1", ServiceCategoryId = "SER-TOH-S", AutomationCategoryId = "ASP-TOH-S", GroupName = "TOH" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_traveltovslholidaytime1", ServiceCategoryId = "SER-THH-S", AutomationCategoryId = "ASP-THH-S", GroupName = "THH" });
            //Waiting           
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_waitingnormaltime", ServiceCategoryId = "SER-WNH-S", AutomationCategoryId = "ASP-WNH-S", GroupName = "WNH" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_waitingovertime", ServiceCategoryId = "SER-WOH-S", AutomationCategoryId = "ASP-WOH-S", GroupName = "WOH" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_waitingholidaytime", ServiceCategoryId = "SER-WHH-S", AutomationCategoryId = "ASP-WHH-S", GroupName = "WHH" });
            //Waiting 1
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_waitingnormaltime1", ServiceCategoryId = "SER-WNH-S", AutomationCategoryId = "ASP-WNH-S", GroupName = "WNH" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_waitingovertime1", ServiceCategoryId = "SER-WOH-S", AutomationCategoryId = "ASP-WOH-S", GroupName = "WOH" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_waitingholidaytime1", ServiceCategoryId = "SER-WHH-S", AutomationCategoryId = "ASP-WHH-S", GroupName = "WHH" });
            //Working           
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_workingnormaltime", ServiceCategoryId = "SER-NH-S", AutomationCategoryId = "ASP-NH-S", GroupName = "NH" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_workingovertime", ServiceCategoryId = "SER-OHL-S", AutomationCategoryId = "ASP-OHL-S", GroupName = "OHL" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_workingholidaytime", ServiceCategoryId = "SER-HH-S", AutomationCategoryId = "ASP-HH-S", GroupName = "HH" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_workingoutstationday", ServiceCategoryId = "SER-OND-S", AutomationCategoryId = "ASP-OND-S", GroupName = "OND" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_workingoutstation", ServiceCategoryId = "SER-OH-S", AutomationCategoryId = "ASP-OH-S", GroupName = "OH" });
            //Working 1         
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_workingnormaltime1", ServiceCategoryId = "SER-NH-S", AutomationCategoryId = "ASP-NH-S", GroupName = "NH" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_workingovertime1", ServiceCategoryId = "SER-OHL-S", AutomationCategoryId = "ASP-OHL-S", GroupName = "OHL" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_workingholidaytime1", ServiceCategoryId = "SER-HH-S", AutomationCategoryId = "ASP-HH-S", GroupName = "HH" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_workingoutstationdaynormal1", ServiceCategoryId = "SER-OND-S", AutomationCategoryId = "ASP-OND-S", GroupName = "OND" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_workingoutstationdayholiday1", ServiceCategoryId = "SER-OH-S", AutomationCategoryId = "ASP-OH-S", GroupName = "OH" });
            //Travelling From Vessel
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_travelfmvslnormaltime", ServiceCategoryId = "SER-TNH-S", AutomationCategoryId = "ASP-TNH-S", GroupName = "TNH" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_travelfmvslovertime", ServiceCategoryId = "SER-TOH-S", AutomationCategoryId = "ASP-TOH-S", GroupName = "TOH" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_travelfmvslholidaytime", ServiceCategoryId = "SER-THH-S", AutomationCategoryId = "ASP-THH-S", GroupName = "THH" });
            //Travelling From Vessel 1 
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_travelfmvslnormaltime1", ServiceCategoryId = "SER-TNH-S", AutomationCategoryId = "ASP-TNH-S", GroupName = "TNH" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_travelfmvslovertime1", ServiceCategoryId = "SER-TOH-S", AutomationCategoryId = "ASP-TOH-S", GroupName = "TOH" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_travelfmvslholidaytime1", ServiceCategoryId = "SER-THH-S", AutomationCategoryId = "ASP-THH-S", GroupName = "THH" });
            //Workshop          
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_workshopnormaltime", ServiceCategoryId = "SER-WNH001-S", AutomationCategoryId = "ASP-WNH001-S", GroupName = "WNH001" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_workshopovertime", ServiceCategoryId = "SER-WOH001-S", AutomationCategoryId = "ASP-WOH001-S", GroupName = "WOH001" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_workshopholidaytime", ServiceCategoryId = "SER-WOH001-S", AutomationCategoryId = "ASP-WOH001-S", GroupName = "WOH001" });
            //Workshop 1        
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_workshopnormaltime1", ServiceCategoryId = "SER-WNH001-S", AutomationCategoryId = "ASP-WNH001-S", GroupName = "WNH001" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_workshopovertime1", ServiceCategoryId = "SER-WOH001-S", AutomationCategoryId = "ASP-WOH001-S", GroupName = "WOH001" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_workshopholidaytime1", ServiceCategoryId = "SER-WOH001-S", AutomationCategoryId = "ASP-WOH001-S", GroupName = "WOH001" });
            //Waiting 2         
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_waitingnormaltime2", ServiceCategoryId = "SER-WNH-S", AutomationCategoryId = "ASP-WNH-S", GroupName = "WNH" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_waitingovertime2", ServiceCategoryId = "SER-WOH-S", AutomationCategoryId = "ASP-WOH-S", GroupName = "WOH" });
            mappingList.Add(new JobSheetFieldMapping { FieldSchemaName = "elc_waitingoutstationdaynormal2", ServiceCategoryId = "SER-WHH-S", AutomationCategoryId = "ASP-WHH-S", GroupName = "WHH" });
            #endregion
            try
            {
                #region Object Instances
                context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                Entity entity = context.InputParameters["Target"] as Entity; ;
                orgFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                orgService = orgFactory.CreateOrganizationService(context.UserId);
                Entity entityPostImage = null;
                List<JobSheetFieldMapping> GroupedFieldList = new List<JobSheetFieldMapping>();
                #endregion

                if (context.Depth > 2) return;

                if (context.MessageName.ToUpper() == "CREATE")
                    entityPostImage = context.PostEntityImages["PostCreateForJobSheetTiming"] as Entity;
                if (context.MessageName.ToUpper() == "UPDATE")
                    entityPostImage = context.PostEntityImages["PostUpdateForJobSheetTiming"] as Entity;

                foreach (var item in entityPostImage.Attributes)
                    if (mappingList.Where(x => x.FieldSchemaName == item.Key).Count() > 0)
                        GroupedFieldList.AddRange(mappingList.Where(x => x.FieldSchemaName == item.Key));

                //Group By All Mapping by Group Name
                var groupedFieldMappingList = GroupedFieldList.GroupBy(u => u.GroupName).Select(grp => grp);

                foreach (var item in groupedFieldMappingList)
                {
                    //Calculate total hours For Specifc Group
                    int EachHours = 0;
                    item.Select(x => x.FieldSchemaName).ToList().ForEach(a => { EachHours = EachHours + Convert.ToInt32(entityPostImage.Attributes.Where(x => x.Key == a).Sum(b => Convert.ToInt32(b.Value))); });

                    //Check for Automation Code or Service Code to pick up
                    string ServiceCategoryId = string.Empty;
                    if (((string)entityPostImage["msdyn_name"]).Contains("AUS"))
                        ServiceCategoryId = mappingList.Where(x => x.FieldSchemaName == item.Select(a => a.FieldSchemaName).FirstOrDefault()).Select(x => x.AutomationCategoryId).FirstOrDefault();
                    else
                        ServiceCategoryId = mappingList.Where(x => x.FieldSchemaName == item.Select(a => a.FieldSchemaName).FirstOrDefault()).Select(x => x.ServiceCategoryId).FirstOrDefault();

                    //Get Category Id from Category Number
                    string GetCategoryFetchExpression = string.Format(GetCategoriesFetchXML, ServiceCategoryId,
                        entityPostImage.Contains(LegalEntityEntitySchemaName) ? ((EntityReference)entityPostImage[LegalEntityEntitySchemaName]).Id.ToString() : DefaultLegalEntity);
                    var Category = orgService.RetrieveMultiple(new FetchExpression(GetCategoryFetchExpression));

                    if (Category.Entities.Count > 0)
                    {
                        if (context.MessageName.ToUpper() == "UPDATE")
                        {
                            //Get Service Hours For Updated Service Order
                            string GetServiceHoursFetchExpression = string.Format(GetServiceHoursForUpdate, Category.Entities[0].Id, entity.Id);
                            var ServicesHours = orgService.RetrieveMultiple(new FetchExpression(GetServiceHoursFetchExpression));
                            //If No Service Hours Created, Create One
                            if (ServicesHours.Entities.Count == 0)
                            {
                                Entity CreateServiceHour = new Entity(ServiceHourEntitySchemaName);
                                CreateServiceHour[ProjectCategoryEntitySchemaName] = new EntityReference(ProjectCategoryEntitySchemaName, Category.Entities[0].Id);
                                CreateServiceHour["msdyn_workorder"] = new EntityReference("msdyn_workorder", entity.Id);
                                CreateServiceHour[LegalEntityEntitySchemaName] = new EntityReference("businessunit", entityPostImage.Contains(LegalEntityEntitySchemaName) ? ((EntityReference)entityPostImage[LegalEntityEntitySchemaName]).Id : new Guid(DefaultLegalEntity));
                                CreateServiceHour["msdyn_duration"] = EachHours;
                                orgService.Create(CreateServiceHour);
                            }
                            //If Service Hours is in system check value of duration, Update only if Not same as new value
                            else if (ServicesHours.Entities[0].Contains("msdyn_duration") && Convert.ToInt32(ServicesHours.Entities[0]["msdyn_duration"]) != EachHours)
                            {
                                Entity UpdateServiceHour = new Entity(ServiceHourEntitySchemaName, ServicesHours.Entities[0].Id);
                                UpdateServiceHour["msdyn_duration"] = EachHours;
                                orgService.Update(UpdateServiceHour);
                            }
                        }
                        if (context.MessageName.ToUpper() == "CREATE")
                        {
                            Entity CreateServiceHour = new Entity(ServiceHourEntitySchemaName);
                            CreateServiceHour[ProjectCategoryEntitySchemaName] = new EntityReference(ProjectCategoryEntitySchemaName, Category.Entities[0].Id);
                            CreateServiceHour["msdyn_workorder"] = new EntityReference("msdyn_workorder", entity.Id);
                            CreateServiceHour[LegalEntityEntitySchemaName] = new EntityReference("businessunit", entityPostImage.Contains(LegalEntityEntitySchemaName) ? ((EntityReference)entityPostImage[LegalEntityEntitySchemaName]).Id : new Guid(DefaultLegalEntity));
                            CreateServiceHour["msdyn_duration"] = EachHours;
                            orgService.Create(CreateServiceHour);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public class JobSheetFieldMapping
        {
            public string FieldSchemaName { get; set; }
            public string ServiceCategoryId { get; set; }
            public string AutomationCategoryId { get; set; }
            public string GroupName { get; set; }
        }
    }
}



