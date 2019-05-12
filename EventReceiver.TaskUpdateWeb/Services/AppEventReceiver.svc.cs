using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;

namespace EventReceiver.TaskUpdateWeb.Services
{
    public class AppEventReceiver : IRemoteEventService
    {
        private const string RECEIVER_NAME = "EventReceiver.TaskUpdate.ItemUpdatedEvent";
        private const string LIST_TITLE = "Workflow Tasks";

        /// <summary>
        /// Handles app events that occur after the app is installed or upgraded, or when app is being uninstalled.
        /// </summary>
        /// <param name="properties">Holds information about the app event.</param>
        /// <returns>Holds information returned from the app event.</returns>
        public SPRemoteEventResult ProcessEvent(SPRemoteEventProperties properties)
        {
            Helper.ClearLog();
            Helper.AddLog("AppEventReceiver-", "On Started on : "+properties.ItemEventProperties.WebUrl);
            Helper.AddLog("AppEventReceiver-", "List : " + properties.ItemEventProperties.ListTitle+", Item Id : "+ properties.ItemEventProperties.ListItemId.ToString());
            SPRemoteEventResult result = new SPRemoteEventResult();
            Helper.AddLog("AppEventReceiver-Event Type", properties.EventType.ToString());
            switch (properties.EventType)
            {
                case SPRemoteEventType.AppInstalled:
                    HandleAppInstalled(properties);
                    break;
                case SPRemoteEventType.AppUninstalling:
                    HandleAppUninstalling(properties);
                    break;


                case SPRemoteEventType.ItemAdding:
                    HandleItemAdding(properties, result);
                    break;
                case SPRemoteEventType.ItemUpdating:
                    HandleItemUpdating(properties, result);
                    break;

                case SPRemoteEventType.ItemAdded:
                    HandleItemAdded(properties);
                    break;
                case SPRemoteEventType.ItemUpdated:
                    HandleItemUpdated(properties, result);
                    break;
            }

            List<string> emails = new List<string>();
            emails.Add("me@gauravgoyal.onmicrosoft.com");
            Helper.SendEmail(emails, "AppEventReceiver : " + properties.EventType.ToString(), Helper.GetLog());

            return result;
        }

        /// <summary>
        /// This method is a required placeholder, but is not used by app events.
        /// </summary>
        /// <param name="properties">Unused.</param>
        public void ProcessOneWayEvent(SPRemoteEventProperties properties)
        {
            using (ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
            {
                if (clientContext != null)
                {
                    if (properties.EventType.Equals(SPRemoteEventType.ItemUpdated))
                    {
                        var afterProperties = properties.ItemEventProperties.AfterProperties;
                        var beforeProperties = properties.ItemEventProperties.BeforeProperties;
                    }
                }
            }
        }


        /// <summary>
        /// Handles when an app is installed.  Activates a feature in the
        /// host web.  The feature is not required.  
        /// Next, if the Jobs list is
        /// not present, creates it.  Finally it attaches a remote event
        /// receiver to the list.  
        /// </summary>
        /// <param name="properties"></param>
        private void HandleAppInstalled(SPRemoteEventProperties properties)
        {
            using (ClientContext clientContext =
                TokenHelper.CreateAppEventClientContext(properties, false))
            {
                if (clientContext != null)
                {
                    new RemoteEventReceiverManager().AssociateRemoteEventsToHostWeb(clientContext);
                }
            }
        }

        /// <summary>
        /// Removes the remote event receiver from the list and 
        /// adds a new item to the list.
        /// </summary>
        /// <param name="properties"></param>
        private void HandleAppUninstalling(SPRemoteEventProperties properties)
        {
            using (ClientContext clientContext =
                TokenHelper.CreateAppEventClientContext(properties, false))
            {
                if (clientContext != null)
                {
                    new RemoteEventReceiverManager().RemoveEventReceiversFromHostWeb(clientContext, "RemoteEventReceiver1ItemUpdating");
                    new RemoteEventReceiverManager().RemoveEventReceiversFromHostWeb(clientContext, "RemoteEventReceiver1ItemUpdated");
                }
            }
        }

        /// <summary>
        /// Handles the ItemAdding event by check the Description
        /// field of the item.
        /// </summary>
        /// <param name="properties"></param>
        private void HandleItemAdding(SPRemoteEventProperties properties, SPRemoteEventResult result)
        {
            using (ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
            {
                if (clientContext != null)
                {
                    new RemoteEventReceiverManager().ItemAddingToListEventHandler(clientContext, properties, result);
                }
            }
        }

        /// <summary>
        /// Handles the ItemUpdating event by check the Description
        /// field of the item.
        /// </summary>
        /// <param name="properties"></param>
        private void HandleItemUpdating(SPRemoteEventProperties properties, SPRemoteEventResult result)
        {
            using (ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
            {
                if (clientContext != null)
                {
                    new RemoteEventReceiverManager().ItemUpdatingToListEventHandler(clientContext, properties, result);
                }
            }
        }

        private void HandleItemUpdated(SPRemoteEventProperties properties, SPRemoteEventResult result)
        {
            using (ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
            {
                if (clientContext != null)
                {
                    new RemoteEventReceiverManager().ItemUpdatedToListEventHandler(clientContext, properties, result);
                }
            }
        }

        /// <summary>
        /// Handles the ItemAdded event by modifying the Description
        /// field of the item.
        /// </summary>
        /// <param name="properties"></param>
        private void HandleItemAdded(SPRemoteEventProperties properties)
        {
            Helper.AddLog("Under HandleItemAdded");
            using (ClientContext clientContext = TokenHelper.CreateRemoteEventReceiverClientContext(properties))
            {
                if (clientContext != null )
                {
                    Helper.AddLog("Inside of HandleItemAdded");
                    new RemoteEventReceiverManager().ItemAddedToListEventHandler(clientContext, properties.ItemEventProperties.ListId, properties.ItemEventProperties.ListItemId);
                }
            }
            Helper.AddLog("Going out HandleItemAdded");
        }



    }
}
