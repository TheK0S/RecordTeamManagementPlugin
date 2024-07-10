using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Services;
using System.ServiceModel;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace TraineeCasesSecurityPlugin
{
    public class TableRowActivationPlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    Guid recordId = entity.Id;

                    if (entity.Contains("statecode"))
                    {
                        OptionSetValue state = (OptionSetValue)entity["statecode"];
                        bool isActive = (state.Value == 0); // Active = 0, Inactive = 1

                        string accessTeamTemplateName = "Security team template";
                        string teamName = "Secret Team";

                        if (isActive)
                            RemoveUsersFromAccessTeam(service, recordId, accessTeamTemplateName, tracingService);
                        else
                            AddUsersToAccessTeam(service, recordId, accessTeamTemplateName, teamName);
                    }
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace($"ManageAccessTeamPlugin: {ex.ToString()}");
                throw new InvalidPluginExecutionException($"An error occurred in ManageAccessTeamPlugin: {ex.Message}");
            }
        }

        private void AddUsersToAccessTeam(IOrganizationService service, Guid recordId, string accessTeamTemplateName, string teamName)
        {
            // Get Security team template ID
            QueryExpression queryTemplate = new QueryExpression("teamtemplate")
            {
                ColumnSet = new ColumnSet("teamtemplateid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("teamtemplatename", ConditionOperator.Equal, accessTeamTemplateName)
                    }
                }
            };

            EntityCollection templates = service.RetrieveMultiple(queryTemplate);
            if (templates.Entities.Count == 0)
                throw new InvalidPluginExecutionException($"Access Team Template with name '{accessTeamTemplateName}' not found.");

            Guid templateId = templates.Entities[0].Id;

            // Get Security Team ID by name
            QueryExpression queryTeam = new QueryExpression("team")
            {
                ColumnSet = new ColumnSet("teamid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("name", ConditionOperator.Equal, teamName)
                    }
                }
            };

            EntityCollection teams = service.RetrieveMultiple(queryTeam);
            if (teams.Entities.Count == 0)
                throw new InvalidPluginExecutionException($"Security Team with name '{teamName}' not found.");

            Guid securityTeamId = teams.Entities[0].Id;

            // Get Security Team users
            QueryExpression queryUsers = new QueryExpression("teammembership")
            {
                ColumnSet = new ColumnSet("systemuserid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("teamid", ConditionOperator.Equal, securityTeamId)
                    }
                }
            };

            EntityCollection users = service.RetrieveMultiple(queryUsers);

            foreach (Entity userEntity in users.Entities)
            {
                Guid userId = userEntity.GetAttributeValue<Guid>("systemuserid");

                // Add users to Access Team
                var addUserRequest = new AddUserToRecordTeamRequest
                {
                    Record = new EntityReference("cre76_sequritytable", recordId),
                    SystemUserId = userId,
                    TeamTemplateId = templateId
                };

                service.Execute(addUserRequest);
            }
        }

        private void RemoveUsersFromAccessTeam(IOrganizationService service, Guid recordId, string accessTeamTemplateName, ITracingService tracingService)
        {
            // Get Security team template ID
            QueryExpression query = new QueryExpression("teamtemplate")
            {
                ColumnSet = new ColumnSet("teamtemplateid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("teamtemplatename", ConditionOperator.Equal, accessTeamTemplateName)
                    }
                }
            };

            EntityCollection templates = service.RetrieveMultiple(query);
            if (templates.Entities.Count == 0)
                throw new InvalidPluginExecutionException($"Security team template with name = '{accessTeamTemplateName}' not found.");

            Guid teamTemplateId = templates.Entities[0].Id;

            //Get record owner
            Entity record = service.Retrieve("cre76_sequritytable", recordId, new ColumnSet("ownerid"));
            Guid ownerId = record.GetAttributeValue<EntityReference>("ownerid").Id;

            //Get users from access team
            string fetchUsersXml = @"
            <fetch>
                <entity name='systemuser'>
                    <attribute name='systemuserid' />
                    <link-entity name='teammembership' from='systemuserid' to='systemuserid' intersect='true' link-type='inner'>
                        <link-entity name='team' from='teamid' to='teamid' alias='te'>
                            <link-entity name='teamtemplate' from='teamtemplateid' to='teamtemplateid'>
                                <filter>
                                    <condition attribute='teamtemplateid' operator='eq' value='" + teamTemplateId + @"' />
                                </filter>
                            </link-entity>
                        </link-entity>
                    </link-entity>
                </entity>
            </fetch>";

            EntityCollection teamUsers = service.RetrieveMultiple(new FetchExpression(fetchUsersXml));
            if (teamUsers.Entities.Count == 0)
                throw new InvalidPluginExecutionException($"No users found in Access Team Template.");
            

            //Remove users from access team without record owner
            foreach (var user in teamUsers.Entities)
            {
                Guid userId = user.GetAttributeValue<Guid>("systemuserid");

                if (userId == ownerId)
                    continue;

                var request = new RemoveUserFromRecordTeamRequest
                {
                    TeamTemplateId = teamTemplateId,
                    SystemUserId = userId,
                    Record = new EntityReference("cre76_sequritytable", recordId)
                };

                service.Execute(request);
            }
        }
    }
}
