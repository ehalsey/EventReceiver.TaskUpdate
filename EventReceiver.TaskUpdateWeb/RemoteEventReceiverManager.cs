using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ServiceModel;
using System.ServiceModel.Channels;
using Microsoft.SharePoint.Client.EventReceivers;
using System.Reflection;

using Newtonsoft.Json;

namespace EventReceiver.TaskUpdateWeb
{
    public class RemoteEventReceiverManager
    {
        //private const string RECEIVER_NAME_UPDATED = "ItemAddedEvent";
        private const string RECEIVER_NAME_UPDATED = "EventReceiver.TaskUpdate.ItemUpdatedEvent";
        private const string LIST_TITLE = "Workflow Tasks";

        public void AssociateRemoteEventsToHostWeb(ClientContext clientContext)
        {
            Helper.AddLog("RemoteEventReceiverManager-AssociateRemoteEventsToHostWeb", "On Started");
            try
            {
                List jobsList = clientContext.Web.Lists.GetByTitle(LIST_TITLE);
                clientContext.Load(jobsList, l => l.Title, l => l.EventReceivers);
                clientContext.ExecuteQuery();

                Helper.AddLog("RemoteEventReceiverManager-AssociateRemoteEventsToHostWeb", "Read List");

                bool rerExists = false;

                foreach (var rer in jobsList.EventReceivers)
                {
                    if (rer.ReceiverName.Contains("RemoteEventReceiver1Item"))
                    {
                        Helper.AddLog("RemoteEventReceiverManager-AssociateRemoteEventsToHostWeb", "Removing Event Receiver : "+rer.ReceiverName);
                        RemoveEventReceiversFromHostWeb(clientContext, rer.ReceiverName);
                        Helper.AddLog("RemoteEventReceiverManager-AssociateRemoteEventsToHostWeb", "Removed Event Receiver : " + rer.ReceiverName);
                        //rerExists = true;
                        System.Diagnostics.Trace.WriteLine("Found existing ItemAdded receiver at "
                            + rer.ReceiverUrl);
                    }
                }


                if (!rerExists)
                {
                    Helper.AddLog("RemoteEventReceiverManager-AssociateRemoteEventsToHostWeb", "Adding Event Receiver 1" );
                    //should be in config
                    string url = "https://remoteeventreceiverfortaskbs.azurewebsites.net/Services/AppEventReceiver.svc";
                    Helper.AddLog("RemoteEventReceiverManager-AssociateRemoteEventsToHostWeb", url);
                    EventReceiverDefinitionCreationInformation receiver =
                        new EventReceiverDefinitionCreationInformation();
                    receiver.EventType = EventReceiverType.ItemUpdating;
                    receiver.ReceiverUrl = url;
                    Helper.AddLog("RemoteEventReceiverManager-AssociateRemoteEventsToHostWeb", Assembly.GetExecutingAssembly().FullName);
                    receiver.ReceiverAssembly = Assembly.GetExecutingAssembly().FullName;
                    receiver.ReceiverClass = "AppEventReceiver";
                    receiver.ReceiverName = "RemoteEventReceiver1ItemUpdating";
                    receiver.SequenceNumber = 1009;
                    receiver.Synchronization = EventReceiverSynchronization.Synchronous;

                    //Add the new event receiver to a list in the host web
                    jobsList.EventReceivers.Add(receiver);
                    //clientContext.ExecuteQuery();
                    Helper.AddLog("RemoteEventReceiverManager-AssociateRemoteEventsToHostWeb", "Adding Event Receiver 2");
                    EventReceiverDefinitionCreationInformation receiver2 =
                        new EventReceiverDefinitionCreationInformation();
                    receiver2.EventType = EventReceiverType.ItemUpdated;
                    receiver2.ReceiverUrl = url;
                    Helper.AddLog("RemoteEventReceiverManager-AssociateRemoteEventsToHostWeb", Assembly.GetExecutingAssembly().FullName);
                    receiver2.ReceiverAssembly = Assembly.GetExecutingAssembly().FullName;
                    receiver2.ReceiverClass = "AppEventReceiver";
                    receiver2.ReceiverName = "RemoteEventReceiver1ItemUpdated";
                    receiver2.SequenceNumber = 1008;
                    receiver2.Synchronization = EventReceiverSynchronization.Synchronous;

                    //Add the new event receiver to a list in the host web
                    jobsList.EventReceivers.Add(receiver2);
                    clientContext.ExecuteQuery();

                    EventReceiverDefinitionCreationInformation receiver3 =
                        new EventReceiverDefinitionCreationInformation();
                    receiver3.EventType = EventReceiverType.ItemAdded;
                    receiver3.ReceiverUrl = url;
                    Helper.AddLog("RemoteEventReceiverManager-AssociateRemoteEventsToHostWeb", Assembly.GetExecutingAssembly().FullName);
                    receiver3.ReceiverAssembly = Assembly.GetExecutingAssembly().FullName;
                    receiver3.ReceiverClass = "AppEventReceiver";
                    receiver3.ReceiverName = "RemoteEventReceiver1ItemAdded";
                    receiver3.SequenceNumber = 1007;
                    receiver3.Synchronization = EventReceiverSynchronization.Asynchronous;

                    jobsList.EventReceivers.Add(receiver3);
                    clientContext.ExecuteQuery();

                    Helper.AddLog("RemoteEventReceiverManager-AssociateRemoteEventsToHostWeb", "Added Event Receiver 1 , 2 and 3");
                    System.Diagnostics.Trace.WriteLine("Added ItemAdded receiver at " + receiver.ReceiverUrl);


                }
            }
            catch (Exception ex)
            {
                Helper.AddLog("RemoteEventReceiverManager-AssociateRemoteEventsToHostWeb", "Adding Event Receiver Error "+ ex.ToString());
            }
        }

        public void RemoveEventReceiversFromHostWeb(ClientContext clientContext, string receiverName)
        {
            try
            {
                Helper.AddLog("Removing receiver at ",receiverName);
                List myList = clientContext.Web.Lists.GetByTitle(LIST_TITLE);
                List testList = clientContext.Web.Lists.GetByTitle("TestList");
                clientContext.Load(myList, p => p.EventReceivers);
                clientContext.Load(testList);
                clientContext.ExecuteQuery();
                Helper.AddLog("Loaded Lists");

                var rer = myList.EventReceivers.Where(
                    e => e.ReceiverName == receiverName).FirstOrDefault();

                Helper.AddLog("Got ERs");
                try
                {
                    if(rer == null)
                    {
                        Helper.AddLog("No Event Receiver Find "+receiverName);
                        return;
                    }
                    Helper.AddLog("Removing receiver at ",rer.ReceiverUrl);

                    var rerList = myList.EventReceivers.Where(
                    e => e.ReceiverUrl == rer.ReceiverUrl).ToList<EventReceiverDefinition>();

                    foreach (var rerFromUrl in rerList)
                    {
                        Helper.AddLog("Removing receiver --- ",
                            rerFromUrl.ReceiverName);
                        //This will fail when deploying via F5, but works
                        //when deployed to production
                        rerFromUrl.DeleteObject();
                        Helper.AddLog("Removing receiver --- End ",
                            rerFromUrl.ReceiverName);
                    }
                    clientContext.ExecuteQuery();
                }
                catch (Exception oops)
                {
                    Helper.AddLog("RemoteEventReceiverManager-AssociateRemoteEventsToHostWeb", "Error in Removing Event Receiver : " + oops.ToString());
                    System.Diagnostics.Trace.WriteLine(oops.Message);
                }

                //Now the RER is removed, add a new item to the list
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem newItem = testList.AddItem(itemCreateInfo);
                newItem["Title"] = "App deleted : " + receiverName;
                newItem["TaskDetails"] = "Deleted on " + System.DateTime.Now.ToLongTimeString();
                newItem.Update();

                clientContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
                Helper.AddLog("RemoteEventReceiverManager-AssociateRemoteEventsToHostWeb", "Error in Removing Event Receiver 2 : " + ex.ToString());
            }
        }

        public void ItemAddingToListEventHandler(ClientContext clientContext,
            SPRemoteEventProperties properties, SPRemoteEventResult result)
        {
            try
            {
                // only for demo we check here the Description
                if (properties.ItemEventProperties.AfterProperties["Description"] != null &&
                    !string.IsNullOrEmpty(properties.ItemEventProperties.AfterProperties["Description"].ToString()))
                {
                    throw new Exception("Description should be empty!");
                }
                else
                {
                    result.Status = SPRemoteEventServiceStatus.Continue;
                }
            }
            catch (Exception oops)
            {
                result.Status = SPRemoteEventServiceStatus.CancelWithError;
                result.ErrorMessage = oops.Message;

                System.Diagnostics.Trace.WriteLine(oops.Message);
            }
        }

        public void ItemUpdatingToListEventHandler(ClientContext clientContext,
            SPRemoteEventProperties properties, SPRemoteEventResult result)
        {
            try
            {
                Helper.AddLog("RemoteEventReceiverManager-ItemUpdatingToListEventHandler", "" );
                var relatedItem = new TaskDetail();
                string beforeName = "";
                string assignToFieldName = "AssignedTo";

                relatedItem = GetListItemDetails(clientContext, properties.ItemEventProperties.ListItemId);

                try
                {
                    if (properties.ItemEventProperties.BeforeProperties.ContainsKey(assignToFieldName))
                    {
                        beforeName = Convert.ToString(properties.ItemEventProperties.BeforeProperties[assignToFieldName]);
                        Helper.AddLog("RemoteEventReceiverManager-ItemUpdatingToListEventHandler", "Before :" + beforeName);
                    }
                    else
                    {
                        Helper.AddLog("RemoteEventReceiverManager-ItemUpdatingToListEventHandler", "No Before");
                    }
                }
                catch(Exception ex1)
                {
                    beforeName = "Before Name got Error : " + ex1.ToString();
                    Helper.AddLog("RemoteEventReceiverManager-ItemUpdatingToListEventHandler", beforeName);
                }
                string afterName = "";
                try
                {
                    if (properties.ItemEventProperties.AfterProperties.ContainsKey(assignToFieldName))
                    {
                        afterName = Convert.ToString(properties.ItemEventProperties.AfterProperties[assignToFieldName]);
                        Helper.AddLog("RemoteEventReceiverManager-ItemUpdatingToListEventHandler", "After : " + afterName + ", After Variable Length should be > 5 "+afterName.Length.ToString());
                        if(afterName.Length>5)
                        {
                            Helper.AddLog("RemoteEventReceiverManager-ItemUpdatingToListEventHandler", "Spliting now");
                            string[] splitAfterName = afterName.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                            Helper.AddLog("RemoteEventReceiverManager-ItemUpdatingToListEventHandler", "Splited array length : "+ splitAfterName.Length.ToString());
                            if (splitAfterName.Length > 0)
                            {
                                Helper.AddLog("RemoteEventReceiverManager-ItemUpdatingToListEventHandler", "getting user login name");
                                string userName = splitAfterName[(splitAfterName.Length - 1)];
                                Helper.AddLog("RemoteEventReceiverManager-ItemUpdatingToListEventHandler", "got user login name : "+userName);
                                Helper.AddLog("RemoteEventReceiverManager-ItemUpdatingToListEventHandler", "getting user object");
                                var afterUserObj = clientContext.Web.EnsureUser(userName);
                                clientContext.Load(afterUserObj);
                                clientContext.ExecuteQuery();
                                Helper.AddLog("RemoteEventReceiverManager-ItemUpdatingToListEventHandler", "got user object & checking null");
                                if (afterUserObj != null)
                                {
                                    Helper.AddLog("RemoteEventReceiverManager-ItemUpdatingToListEventHandler", "user object is not null");
                                    relatedItem.NewUserId = afterUserObj.Id.ToString();
                                    relatedItem.NewUserLoginId = afterUserObj.LoginName;
                                    relatedItem.NewUserName = afterUserObj.Title;
                                    Helper.AddLog("RemoteEventReceiverManager-ItemUpdatingToListEventHandler", "user object info : Id="+afterUserObj.Id.ToString()+", LoginName="+afterUserObj.LoginName+", Name ="+afterUserObj.Title);
                                    UpdateDocLib(clientContext, relatedItem);
                                }
                            }
                        }
                        //var userName = (FieldUserValue)properties.ItemEventProperties.AfterProperties[assignToFieldName];
                        //Helper.AddLog("RemoteEventReceiverManager-ItemUpdatingToListEventHandler", "User Id : " + userName.LookupId);
                        //Helper.AddLog("RemoteEventReceiverManager-ItemUpdatingToListEventHandler", "User Name : " + userName.LookupValue);
                        //Helper.AddLog("RemoteEventReceiverManager-ItemUpdatingToListEventHandler", "User Email : " + userName.Email);
                        //relatedItem.NewUserId = userName.LookupId.ToString();
                        //relatedItem.NewUserName = userName.LookupValue;
                        //relatedItem.NewUserLoginId = userName.Email;
                    }
                    else
                        Helper.AddLog("RemoteEventReceiverManager-ItemUpdatingToListEventHandler", "No After");
                }
                catch(Exception ex2)
                {
                    afterName = "After Name got Error : " + ex2.ToString();
                    Helper.AddLog("RemoteEventReceiverManager-ItemUpdatingToListEventHandler", afterName);
                }
                
                // only for demo we check here the Description
                var testList = clientContext.Web.Lists.GetByTitle("TestList");
                clientContext.Load(clientContext.Web);
                clientContext.Load(testList);
                clientContext.ExecuteQuery();
                ListItemCreationInformation newItemInfo = new ListItemCreationInformation();
                var newItem = testList.AddItem(newItemInfo);
                newItem["Title"] = DateTime.Now.ToString();
                newItem["TaskDetails"] = "Before : " + beforeName + ", After : " + afterName;
                newItem["DocDetails"] = relatedItem.ListId;
                newItem.Update();
                clientContext.ExecuteQuery();

                result.Status = SPRemoteEventServiceStatus.Continue;

            }
            catch (Exception oops)
            {
                Helper.AddLog("RemoteEventReceiverManager-ItemUpdatingToListEventHandler", "error : " + oops.ToString());
                result.Status = SPRemoteEventServiceStatus.CancelWithError;
                result.ErrorMessage = oops.Message;

                System.Diagnostics.Trace.WriteLine(oops.Message);
            }
        }

        public void ItemUpdatedToListEventHandler(ClientContext clientContext,
            SPRemoteEventProperties properties, SPRemoteEventResult result)
        {
            try
            {
                Helper.AddLog("RemoteEventReceiverManager-ItemUpdatedToListEventHandler", "");

                //string beforeName = "";
                //try
                //{
                //    beforeName = Convert.ToString(properties.ItemEventProperties.BeforeProperties["AssignedTo"]);
                //    Helper.AddLog("RemoteEventReceiverManager-ItemUpdatedToListEventHandler", beforeName);
                //}
                //catch (Exception ex1)
                //{
                //    beforeName = "Before Name got Error : " + ex1.ToString();
                //    Helper.AddLog("RemoteEventReceiverManager-ItemUpdatedToListEventHandler", beforeName);
                //}
                //string afterName = "";
                //try
                //{
                //    afterName = Convert.ToString(properties.ItemEventProperties.AfterProperties["AssignedTo"]);
                //    Helper.AddLog("RemoteEventReceiverManager-ItemUpdatedToListEventHandler", afterName);
                //}
                //catch (Exception ex2)
                //{
                //    afterName = "After Name got Error : " + ex2.ToString();
                //    Helper.AddLog("RemoteEventReceiverManager-ItemUpdatedToListEventHandler", afterName);
                //}
                //string relatedItemBefore = "";
                //string relatedItemAfter = "";
                //try
                //{
                //    relatedItemBefore = Convert.ToString(properties.ItemEventProperties.BeforeProperties["RelatedItems"]);
                //    Helper.AddLog("RemoteEventReceiverManager-ItemUpdatedToListEventHandler", "RelatedItemBefore: " + relatedItemBefore);
                //}
                //catch (Exception ex)
                //{
                //    relatedItemBefore = "Error : " + ex.ToString();
                //    Helper.AddLog("RemoteEventReceiverManager-ItemUpdatedToListEventHandler", "RelatedItemBefore: " + relatedItemBefore);
                //}
                //try
                //{
                //    relatedItemAfter = Convert.ToString(properties.ItemEventProperties.AfterProperties["RelatedItems"]);
                //    Helper.AddLog("RemoteEventReceiverManager-ItemUpdatedToListEventHandler", "RelatedItemAfter: " + relatedItemAfter);
                //}
                //catch (Exception ex)
                //{
                //    relatedItemAfter = "Error : " + ex.ToString();
                //    Helper.AddLog("RemoteEventReceiverManager-ItemUpdatedToListEventHandler", "RelatedItemAfter: " + relatedItemAfter);
                //}

                //// only for demo we check here the Description
                //var testList = clientContext.Web.Lists.GetByTitle("TestList");
                //clientContext.Load(clientContext.Web);
                //clientContext.Load(testList);
                //clientContext.ExecuteQuery();
                //ListItemCreationInformation newItemInfo = new ListItemCreationInformation();
                //var newItem = testList.AddItem(newItemInfo);
                //newItem["Title"] = DateTime.Now.ToString();
                //newItem["TaskDetails"] = "Before : " + beforeName + ", After : " + afterName;
                //newItem["DocDetails"] = relatedItemAfter + ",,," + relatedItemBefore;
                //newItem.Update();
                //clientContext.ExecuteQuery();

                result.Status = SPRemoteEventServiceStatus.Continue;

            }
            catch (Exception oops)
            {
                Helper.AddLog("RemoteEventReceiverManager-ItemUpdatedToListEventHandler", "error : " + oops.ToString());
                result.Status = SPRemoteEventServiceStatus.CancelWithError;
                result.ErrorMessage = oops.Message;

                System.Diagnostics.Trace.WriteLine(oops.Message);
            }
        }

        public void ItemAddedToListEventHandler(ClientContext clientContext, Guid listId, int listItemId)
        {
            try
            {
                Helper.AddLog("ItemAddedToListEventHandler", "Creating Record");
                var testList = clientContext.Web.Lists.GetByTitle("TestList");
                clientContext.Load(clientContext.Web);
                clientContext.Load(testList);
                clientContext.ExecuteQuery();
                ListItemCreationInformation newItemInfo = new ListItemCreationInformation();
                var newItem = testList.AddItem(newItemInfo);
                newItem["Title"] = DateTime.Now.ToString();
                newItem["TaskDetails"] = "Item Added";
                newItem["DocDetails"] = "";
                newItem.Update();
                clientContext.ExecuteQuery();
                Helper.AddLog("ItemAddedToListEventHandler", "Created Record");
            }
            catch (Exception oops)
            {
                Helper.AddLog("ItemAddedToListEventHandler-Error", oops.ToString());
                System.Diagnostics.Trace.WriteLine(oops.Message);
            }

        }

        
        internal TaskDetail GetListItemDetails(ClientContext context,int itemId)
        {
            Helper.AddLog("GetListItemDetails", "Getting record");
            var taskDetail = new TaskDetail(); 
                List list = context.Web.Lists.GetByTitle(LIST_TITLE);
            ListItem itm = list.GetItemById(itemId);
            context.Load(list);
            context.Load(itm);
            context.ExecuteQuery();

            taskDetail.ListId =Convert.ToString( itm["RelatedItems"]);
            taskDetail.RelatedItem = Convert.ToString(itm["RelatedItems"]);
            if (Convert.ToString(itm["AssignedTo"]).Length > 0)
            {
                var userDetail = (FieldUserValue)itm["AssignedTo"];
                taskDetail.OldUserId = userDetail.LookupId.ToString();
                taskDetail.OldUserName = userDetail.LookupValue;
                taskDetail.OldUserLoginId = userDetail.Email;
            }
            if(taskDetail.RelatedItem.Length>32)
            {
                WFRelatedItemDetails[] json = JsonConvert.DeserializeObject<WFRelatedItemDetails[]>(taskDetail.RelatedItem);
                taskDetail.ItemId = Convert.ToString(json[0].ItemId);
                taskDetail.ListId = Convert.ToString(json[0].ListId);
                taskDetail.WebId = Convert.ToString(json[0].WebId);

                Helper.AddLog("GetListItemDetails", "Related Item ItemId : " + taskDetail.ItemId);
                Helper.AddLog("GetListItemDetails", "Related Item ListId : " + taskDetail.ListId);
                Helper.AddLog("GetListItemDetails", "Related Item WebId : " + taskDetail.WebId);
            }
            Helper.AddLog("GetListItemDetails", "Related Item : "+taskDetail.RelatedItem);
            Helper.AddLog("GetListItemDetails", "AssignedTo Id : " + taskDetail.OldUserId);
            Helper.AddLog("GetListItemDetails", "AssignedTo Name : " + taskDetail.OldUserName);
            Helper.AddLog("GetListItemDetails", "AssignedTo Name : " + taskDetail.OldUserLoginId);
            //Do not execute the call.  We simply create the list in the context, 
            //it's up to the caller to call ExecuteQuery.
            return taskDetail;
        }

        internal bool UpdateDocLib(ClientContext context, TaskDetail taskDetail)
        {
            bool retValue = true;
            try
            {
                Helper.AddLog("UpdateDocLib", "Getting record : Bamert AP Documents, ItemId : " + taskDetail.ItemId + ", AssignedTo : " + taskDetail.NewUserLoginId);
                List list = context.Web.Lists.GetByTitle("Bamert AP Documents");
                ListItem itm = list.GetItemById(Convert.ToInt32(taskDetail.ItemId));

                var newUserId = context.Web.EnsureUser(taskDetail.NewUserLoginId);
                Helper.AddLog("UpdateDocLib", "Loading Item + NewUserLgoinId");

                context.Load(list);
                context.Load(itm);
                context.Load(newUserId);
                context.ExecuteQuery();
                Helper.AddLog("UpdateDocLib", "Loaded Item + NewUserLgoinId");

                if (newUserId == null)
                {
                    Helper.AddLog("User object is null");
                    return (retValue = false);
                }

                Helper.AddLog("UpdateDocLib", "User object is not null");
                
                string approvers = Convert.ToString(itm["Approvers"]);
                Helper.AddLog("UpdateDocLib", "Old Approvers : "+ approvers);
                Helper.AddLog("UpdateDocLib", "Replacing Approvers : " + taskDetail.OldUserName+" With "+ newUserId.Title);
                approvers = approvers.Replace(taskDetail.OldUserName, newUserId.Title);
                Helper.AddLog("UpdateDocLib", "New Approvers : " + approvers);
                itm["Approvers"] = approvers;
                Helper.AddLog("UpdateDocLib", "Updating item ");
                itm.Update();
                context.ExecuteQuery();
                Helper.AddLog("UpdateDocLib", "Updated item ");
                retValue = true;
            }
            catch(Exception ex)
            {
                Helper.AddLog("UpdateDocLib-Error", ex.ToString());
                retValue = false;
            }
            return retValue;
        }
    }
}