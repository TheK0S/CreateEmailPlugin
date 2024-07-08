using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace CreateEmailWhenContactCreUpdDel
{
    public class CreateEmailPlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            Entity contact = null;
            EntityReference contactRef = null;
            string messageName = context.MessageName;

            if (context.InputParameters.Contains("Target"))
            {
                if (context.InputParameters["Target"] is Entity)
                {
                    contact = (Entity)context.InputParameters["Target"];
                    contactRef = contact.ToEntityReference();
                }
                else if (context.InputParameters["Target"] is EntityReference)
                {
                    contactRef = (EntityReference)context.InputParameters["Target"];
                }
            }

            if (contactRef == null || contactRef.LogicalName != "contact")
            {
                return;
            }

            tracingService.Trace($"message Name: {messageName}");

            switch (messageName)
            {
                case "Create":
                    HandleCreate(service, context, contact);
                    break;

                case "Update":
                    HandleUpdate(service, context, contact, context.PreEntityImages["PreImage"]);
                    tracingService.Trace($"case Update is completed");
                    break;

                case "Delete":
                    HandleDelete(service, context, contactRef, context.PreEntityImages["PreImage"]);
                    break;
            }
        }

        private void HandleCreate(IOrganizationService service, IPluginExecutionContext context, Entity contact)
        {
            string email = contact.GetAttributeValue<string>("emailaddress1");
            string fullname = contact.GetAttributeValue<string>("fullname");
            string contactId = contact.Attributes["contactid"].ToString();
            string description = $"New contact created - <a href=\"https://org82a3f762.crm11.dynamics.com/main.aspx?etn=contact&id={contactId}&pagetype=entityrecord\">{fullname}</a>";
            DateTime createdon = contact.GetAttributeValue<DateTime>("createdon");
            string subject = $"New Contact {fullname} created {createdon}";

            CreateEmail(service, subject, description, email, contact.ToEntityReference(), new EntityReference("systemuser", context.UserId));
        }

        private void HandleUpdate(IOrganizationService service, IPluginExecutionContext context, Entity contact, Entity preImage)
        {
            string newEmail = contact.GetAttributeValue<string>("emailaddress1");
            string oldEmail = preImage.GetAttributeValue<string>("emailaddress1");
            string fullname = contact.GetAttributeValue<string>("fullname");
            string description = $"Old email address - {oldEmail} \n New email address - {newEmail}";
            DateTime modifiedon = DateTime.Now;
            string subject = $"Contact {fullname} email address changed {modifiedon}";

            CreateEmail(service, subject, description, newEmail, contact.ToEntityReference(), new EntityReference("systemuser", context.UserId));
        }

        private void HandleDelete(IOrganizationService service, IPluginExecutionContext context, EntityReference contactRef, Entity preImage)
        {
            string email = preImage.GetAttributeValue<string>("emailaddress1");
            string fullname = preImage.GetAttributeValue<string>("fullname");
            DateTime modifiedon = DateTime.Now;

            Entity emailEntity = new Entity("email");
            emailEntity["from"] = new EntityCollection(new[] { new Entity("activityparty") { ["partyid"] = new EntityReference("systemuser", context.UserId) } });
            emailEntity["subject"] = $"Contact {fullname} was deleted {modifiedon}";
            emailEntity["description"] = $"Contact was deleted!";

            service.Create(emailEntity);
        }

        private void CreateEmail(IOrganizationService service, string subject, string description, string email, EntityReference regardingContact, EntityReference sender)
        {
            Entity inputEntity = new Entity("new_SetCustomEmailAction");

            inputEntity["Subject"] = subject;
            inputEntity["Body"] = description;
            inputEntity["RecepientEmail"] = email;
            inputEntity["RegardingContact"] = regardingContact;
            inputEntity["Sender"] = sender;

            OrganizationRequest request = new OrganizationRequest
            {
                RequestName = "new_SetCustomEmailAction",
                Parameters = new ParameterCollection { { "Subject", inputEntity["Subject"] }, { "Body", inputEntity["Body"] }, { "RecepientEmail", inputEntity["RecepientEmail"] }, { "RegardingContact", inputEntity["RegardingContact"] }, { "Sender", inputEntity["Sender"] } }
            };

            service.Execute(request);
        }
    }
}