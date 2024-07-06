using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateEmailWhenContactCreUpdDel
{
    public class CreateEmailPlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

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

            switch (messageName)
            {
                case "Create":
                    HandleCreate(service, context, contact);
                    break;

                case "Update":
                    HandleUpdate(service, context, contact, context.PreEntityImages["PreImage"]);
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
            DateTime createdon = contact.GetAttributeValue<DateTime>("createdon");

            Entity emailEntity = new Entity("email");
            // Set the "to" field
            Entity toParty = new Entity("activityparty");
            toParty["partyid"] = contact.ToEntityReference();
            emailEntity["to"] = new EntityCollection(new List<Entity> { toParty });
            // Set the "from" field
            Entity fromParty = new Entity("activityparty");
            fromParty["partyid"] = new EntityReference("systemuser", context.UserId);
            emailEntity["from"] = new EntityCollection(new List<Entity> { fromParty });

            emailEntity["subject"] = $"New Contact {fullname} created {createdon}";
            emailEntity["description"] = $"New contact created - <a href=\"https://org82a3f762.crm11.dynamics.com/main.aspx?etn=contact&id={contactId}&pagetype=entityrecord\">{fullname}</a>";
            emailEntity["regardingobjectid"] = contact.ToEntityReference();

            service.Create(emailEntity);
        }

        private void HandleUpdate(IOrganizationService service, IPluginExecutionContext context, Entity contact, Entity preImage)
        {
            if (contact.Contains("emailaddress1") && contact.Contains("fullname"))
            {
                string newEmail = contact.GetAttributeValue<string>("emailaddress1");
                string oldEmail = preImage.GetAttributeValue<string>("emailaddress1");
                string fullname = contact.GetAttributeValue<string>("fullname");
                DateTime modifiedon = DateTime.Now;

                Entity emailEntity = new Entity("email");
                // Set the "to" field
                Entity toParty = new Entity("activityparty");
                toParty["partyid"] = contact.ToEntityReference();
                emailEntity["to"] = new EntityCollection(new List<Entity> { toParty });
                // Set the "from" field
                Entity fromParty = new Entity("activityparty");
                fromParty["partyid"] = new EntityReference("systemuser", context.UserId);
                emailEntity["from"] = new EntityCollection(new List<Entity> { fromParty });

                emailEntity["subject"] = $"Contact {fullname} email address changed {modifiedon}";
                emailEntity["description"] = $"Old email address - {oldEmail} \n New email address - {newEmail}";
                emailEntity["regardingobjectid"] = contact.ToEntityReference();

                service.Create(emailEntity);
            }
        }

        private void HandleDelete(IOrganizationService service, IPluginExecutionContext context, EntityReference contactRef, Entity preImage)
        {
            string email = preImage.GetAttributeValue<string>("emailaddress1");
            string fullname = preImage.GetAttributeValue<string>("fullname");
            DateTime modifiedon = DateTime.Now;

            Entity emailEntity = new Entity("email");
            // Set the "from" field
            Entity fromParty = new Entity("activityparty");
            fromParty["partyid"] = new EntityReference("systemuser", context.UserId);
            emailEntity["from"] = new EntityCollection(new List<Entity> { fromParty });

            emailEntity["subject"] = $"Contact {fullname} was deleted {modifiedon}";
            emailEntity["description"] = $"Contact was deleted!";

            service.Create(emailEntity);
        }
    }
}