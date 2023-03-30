using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IdentityModel.Metadata;
using System.Linq;
using System.Runtime.Remoting.Services;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace ArthalandPlugins
{
    public class ValidateBuyerInformationSheet : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {

            try
            {
                ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                IPluginExecutionContext context = (IPluginExecutionContext) serviceProvider.GetService(typeof(IPluginExecutionContext));
                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    if (entity.LogicalName == "artha_buyerinformationsheet")
                    {

                        if (context.MessageName.Equals("Update"))
                        {
                            tracingService.Trace("Passed here: Populate CTS Information");
                            PopulateCTSInformation(serviceProvider);
                        }
                        else
                        {
                            tracingService.Trace("Passed here: Validate Buyer Information Sheet");
                            //Get the pipeline guid
                            EntityReference pipeline = (EntityReference)entity.Attributes["artha_pipeline"];

                            //Count all the records in the Buyer Information Sheet based on the Pipeline GUID
                            String fetchXML = @"<fetch>
                                <entity name='artha_buyerinformationsheet'>
                                    <filter>
                                        <condition attribute='artha_pipeline' operator='eq' value='" + pipeline.Id.ToString() + @"' />
                                    </filter>
                                    </entity>
                                </fetch>";

                            String fetchBuyerRolePrincipal = @"<fetch>
                                <entity name='artha_buyerinformationsheet'>
                                    <filter>
                                        <condition attribute='artha_pipeline' operator='eq' value='" + pipeline.Id.ToString() + @"' />
                                        <condition attribute='artha_buyerrole' operator='eq' value='147480000' />
                                    </filter>
                                    </entity>
                                </fetch>";


                            //Retrieve records from Buyer Information Sheet based on FetchXML
                            EntityCollection retrieveBIS = service.RetrieveMultiple(new FetchExpression(fetchXML));
                            EntityCollection validateBISIfPrincipal = service.RetrieveMultiple(new FetchExpression(fetchBuyerRolePrincipal));
                            tracingService.Trace("Passed here: " + retrieveBIS.Entities.Count.ToString());

                            //Validation if record is there are already 3 co-owners and 1 owner
                            if (retrieveBIS.Entities.Count == 4)
                            {
                                throw new InvalidPluginExecutionException("Sorry, you have reached the maximum number of buyer information sheets allowed.");
                            }
                            else
                            {
                                PopulateCTSInformation(serviceProvider);
                                if (validateBISIfPrincipal.Entities.Count == 0)
                                {
                                    //Principal
                                    OptionSetValue BuyerRolePrincipal = new OptionSetValue(147480000);
                                    entity["artha_buyerrole"] = BuyerRolePrincipal;
                                }
                                else
                                {
                                    //Co-Owner
                                    OptionSetValue BuyerRoleCoowner = new OptionSetValue(147480001);
                                    entity["artha_buyerrole"] = BuyerRoleCoowner;
                                }
                            }
                        }
                    }
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        public void PopulateCTSInformation(IServiceProvider serviceProvider)
        {
            try
            {
                ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                tracingService.Trace("Succeeded in try");
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("Context is Entity");

                    Entity preImage = new Entity();

                    if (context.PreEntityImages.Contains("image"))
                    {
                        preImage = context.PreEntityImages["image"];

                        string civilstatus = "";
                        string buyertype = entity.Contains("artha_buyertype") ? ((OptionSetValue)entity["artha_buyertype"]).ToString() : preImage.GetAttributeValue<OptionSetValue>("artha_buyertype").Value.ToString();

                        //Check if buyer type is Corporate
                        if(buyertype != "147480001")
                        {
                            civilstatus = entity.Contains("artha_civilstatus") ? ((OptionSetValue)entity["artha_civilstatus"]).ToString() : preImage.GetAttributeValue<OptionSetValue>("artha_civilstatus").Value.ToString();

                        }

                        string principalfullname = "";
                        string spousefullname = "";
                        string spouseinformation = "";
                        string ownerinformation = "";

                        string principalfirstname = entity.Contains("artha_firstname") ? (string)entity["artha_firstname"] : preImage.GetAttributeValue<string>("artha_firstname");
                        string principalmiddlename = entity.Contains("artha_middlename") ? (string)entity["artha_middlename"] : preImage.GetAttributeValue<string>("artha_middlename");
                        string principallastname = entity.Contains("artha_lastname") ? (string)entity["artha_lastname"] : preImage.GetAttributeValue<string>("artha_lastname");
                        string principalsuffix = entity.Contains("artha_suffix") ? (string)entity["artha_suffix"] : preImage.GetAttributeValue<string>("artha_suffix");

                        tracingService.Trace("Civil Status Value: " + civilstatus);

                        /********************************
                        * Civil Status Optionset Values *
                        * Single              147480000 *
                        * Married             147480001 *
                        * Divorced            147480002 *
                        * Separated           147480003 *
                        * Widowed             147480004 *
                        ********************************/

                        /*********************************************
                        * Registered in the name of Optionset Values *
                        * married to                       147480000 *
                        * as spouses	                   147480001 *
                        *********************************************/

                        //Single
                        if (civilstatus == "147480000")
                        {
                            tracingService.Trace("Passed Here: Single");
                            string spousecitizenship = entity.Contains("artha_citizenship") ? (string)entity["artha_citizenship"] : preImage.GetAttributeValue<string>("artha_citizenship");
                            tracingService.Trace("Passed Here: " + spousecitizenship);
                            string houseunitnumber = entity.Contains("artha_houseunitnumber") ? (string)entity["artha_houseunitnumber"] : preImage.GetAttributeValue<string>("artha_houseunitnumber");
                            string streetname = entity.Contains("artha_streetname") ? (string)entity["artha_streetname"] : preImage.GetAttributeValue<string>("artha_streetname");
                            string buildingtownnumber = entity.Contains("artha_buildingtowernamenumber") ? (string)entity["artha_buildingtowernamenumber"] : preImage.GetAttributeValue<string>("artha_buildingtowernamenumber");
                            string barangaydistrictnumber = entity.Contains("artha_barangaydistrictnumber") ? (string)entity["artha_barangaydistrictnumber"] : preImage.GetAttributeValue<string>("artha_barangaydistrictnumber");
                            string city = entity.Contains("artha_city") ? (string)entity["artha_city"] : preImage.GetAttributeValue<string>("artha_city");
                            string country = entity.Contains("artha_country") ? (string)entity["artha_country"] : preImage.GetAttributeValue<string>("artha_country");
                            string postalcode = entity.Contains("artha_postalcode") ? (string)entity["artha_postalcode"] : preImage.GetAttributeValue<string>("artha_postalcode");

                            principalfullname = principalfirstname + " " + principalmiddlename + " " + principallastname + " " + principalsuffix;
                            ownerinformation = "" +
                                ", " + spousecitizenship +
                                " citizen, of legal age, single, resident of and with postal address at " +
                                houseunitnumber + " " +
                                streetname + " " +
                                buildingtownnumber + ", " +
                                barangaydistrictnumber + ", " +
                                city + ", " +
                                country + ", " +
                                postalcode + ",";

                        }

                        //Married & Married To
                        if (civilstatus == "147480001")
                        {
                            string registeredinthenameof = entity.Contains("artha_registeredinthenameof") ? ((OptionSetValue)entity["artha_registeredinthenameof"]).ToString() : preImage.GetAttributeValue<OptionSetValue>("artha_registeredinthenameof").Value.ToString();
                            string spousefirstname = entity.Contains("artha_coownerspousefirstname") ? (string)entity["artha_coownerspousefirstname"] : preImage.GetAttributeValue<string>("artha_coownerspousefirstname");
                            string spousemiddlename = entity.Contains("artha_coownerspousemiddlename") ? (string)entity["artha_coownerspousemiddlename"] : preImage.GetAttributeValue<string>("artha_coownerspousemiddlename");
                            string spouselastname = entity.Contains("artha_coownerspouselastname") ? (string)entity["artha_coownerspouselastname"] : preImage.GetAttributeValue<string>("artha_coownerspouselastname");
                            string spousesuffix = entity.Contains("artha_coownerspousesuffix") ? (string)entity["artha_coownerspousesuffix"] : preImage.GetAttributeValue<string>("artha_coownerspousesuffix");

                            string spousecitizenship = entity.Contains("artha_coownerspousecitizenship") ? (string)entity["artha_coownerspousecitizenship"] : preImage.GetAttributeValue<string>("artha_coownerspousecitizenship");
                            string houseunitnumber = entity.Contains("artha_coownerspousehouseunitnumber") ? (string)entity["artha_coownerspousehouseunitnumber"] : preImage.GetAttributeValue<string>("artha_coownerspousehouseunitnumber");
                            string streetname = entity.Contains("artha_coownerspousestreetname") ? (string)entity["artha_coownerspousestreetname"] : preImage.GetAttributeValue<string>("artha_coownerspousestreetname");
                            string buildingtownnumber = entity.Contains("artha_coownerspousebuildingtowernamenumber") ? (string)entity["artha_coownerspousebuildingtowernamenumber"] : preImage.GetAttributeValue<string>("artha_coownerspousebuildingtowernamenumber");
                            string barangaydistrictnumber = entity.Contains("artha_coownerspousebarangaydistrictnumber") ? (string)entity["artha_coownerspousebarangaydistrictnumber"] : preImage.GetAttributeValue<string>("artha_coownerspousebarangaydistrictnumber");
                            string city = entity.Contains("artha_coownerspousecity") ? (string)entity["artha_coownerspousecity"] : preImage.GetAttributeValue<string>("artha_coownerspousecity");
                            string country = entity.Contains("artha_coownerspousecountry") ? (string)entity["artha_coownerspousecountry"] : preImage.GetAttributeValue<string>("artha_coownerspousecountry");
                            string postalcode = entity.Contains("artha_coownerspousepostalcode") ? (string)entity["artha_coownerspousepostalcode"] : preImage.GetAttributeValue<string>("artha_coownerspousepostalcode");

                            //Married To
                            if (registeredinthenameof == "147480000")
                            {
                                principalfullname = principalfirstname + " " + principalmiddlename + " " + principallastname + " " + principalsuffix + ",";
                                ownerinformation = " married to ";
                                spousefullname = " " + spousefirstname + " " + spousemiddlename + " " + spouselastname + " " + spousesuffix;
                                spouseinformation = "" +
                                    ", " + spousecitizenship +
                                    " citizens, both of legal age, residents of and with postal address at " +
                                    houseunitnumber + " " +
                                    streetname + " " +
                                    buildingtownnumber + ", " +
                                    barangaydistrictnumber + ", " +
                                    city + ", " +
                                    country + ", " +
                                    postalcode + ",";
                            }

                            //As Spouses
                            else
                            {
                                principalfullname = "SPOUSES " + principalfirstname + " " + principalmiddlename + " " + principallastname + " " + principalsuffix;
                                spousefullname = " and " + spousefirstname + " " + spousemiddlename + " " + spouselastname + " " + spousesuffix;
                                spouseinformation = "" +
                                    ", " + spousecitizenship +
                                    " citizens, both of legal age, residents of and with postal address at " +
                                    houseunitnumber + " " +
                                    streetname + " " +
                                    buildingtownnumber + ", " +
                                    barangaydistrictnumber + ", " +
                                    city + ", " +
                                    country + ", " +
                                    postalcode + ",";
                            }
                            
                        }

                        tracingService.Trace("Spouse Information: " + spouseinformation);
                        tracingService.Trace("Owner Information: " + ownerinformation);
                        tracingService.Trace("Principal Fullname: " + principalfullname);
                        tracingService.Trace("Spouse Fullname: " + spousefullname);

                        //Update the fields on the entity during the pre-operation
                        entity["artha_ctsprincipalspouseinformation"] = spouseinformation;
                        entity["artha_ctsprincipalinformation"] = ownerinformation;
                        entity["artha_ctsprincipalfullname"] = principalfullname;
                        entity["artha_ctsprincipalspousefullname"] = spousefullname;
                    }
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}
