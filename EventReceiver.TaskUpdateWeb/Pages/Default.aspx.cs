using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace EventReceiver.TaskUpdateWeb
{
    public partial class Default : System.Web.UI.Page
    {
        protected void Page_PreInit(object sender, EventArgs e)
        {
            Uri redirectUrl;
            switch (SharePointContextProvider.CheckRedirectionStatus(Context, out redirectUrl))
            {
                case RedirectionStatus.Ok:
                    return;
                case RedirectionStatus.ShouldRedirect:
                    Response.Redirect(redirectUrl.AbsoluteUri, endResponse: true);
                    break;
                case RedirectionStatus.CanNotRedirect:
                    Response.Write("An error occurred while processing your request.");
                    Response.End();
                    break;
            }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                if (Page.IsPostBack) return;
                // define initial script, needed to render the chrome control
                string script = @"
            function chromeLoaded() {
                $('body').show();
            }

            //function callback to render chrome after SP.UI.Controls.js loads
            function renderSPChrome() {
                //Set the chrome options for launching Help, Account, and Contact pages
                var options = {
                    'appTitle': document.title,
                    'onCssLoaded': 'chromeLoaded()'
                };

                //Load the Chrome Control in the divSPChrome element of the page
                var chromeNavigation = new SP.UI.Controls.Navigation('divSPChrome', options);
                chromeNavigation.setVisible(true);
            }";


                //register script in page
                Page.ClientScript.RegisterClientScriptBlock(typeof(Default), "BasePageScript", script, true);

                var contextToken = TokenHelper.GetContextTokenFromRequest(Page.Request);

                // The following code gets the client context and Title property by using TokenHelper.
                // To access other properties, the app may need to request permissions on the host web.
                var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);

                using (var clientContext = spContext.CreateAppOnlyClientContextForSPHost())
                {
                    clientContext.Load(clientContext.Web, web => web.Title);
                    clientContext.ExecuteQuery();
                    Response.Write(clientContext.Web.Title);
                }
            }
            catch(Exception ex)
            {
                Response.Write(ex.ToString());
            }
        }

        protected void btnCreateEventReceiver_Click(object sender, EventArgs e)
        {
            try
            {
                Helper.AddLog("btnCreateEventReceiver_Click");
                var contextToken = TokenHelper.GetContextTokenFromRequest(Page.Request);

                // The following code gets the client context and Title property by using TokenHelper.
                // To access other properties, the app may need to request permissions on the host web.
                var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);
                Helper.AddLog("btnCreateEventReceiver_Click","Creating object");
                using (var clientContext = spContext.CreateAppOnlyClientContextForSPHost())
                {
                    clientContext.Load(clientContext.Web, web => web.Title);
                    clientContext.ExecuteQuery();
                    Helper.AddLog("btnCreateEventReceiver_Click", "Created object");
                    new RemoteEventReceiverManager().AssociateRemoteEventsToHostWeb(clientContext);
                    Helper.AddLog("btnCreateEventReceiver_Click", "Creating Receivers");
                }
            }
            catch(Exception ex)
            {
                Helper.AddLog("btnCreateEventReceiver_Click- Error",ex.ToString());
            }
            List<string> emails = new List<string>();
            emails.Add("me@gauravgoyal.onmicrosoft.com");
            Helper.SendEmail(emails, "Default.aspx : btnCreateEventReceiver_Click ", Helper.GetLog());
        }

        protected void btnRemoveEventReceiver_Click(object sender, EventArgs e)
        {
            try
            {
                Helper.AddLog("btnRemoveEventReceiver_Click");
                var contextToken = TokenHelper.GetContextTokenFromRequest(Page.Request);

                // The following code gets the client context and Title property by using TokenHelper.
                // To access other properties, the app may need to request permissions on the host web.
                var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);
                Helper.AddLog("btnRemoveEventReceiver_Click", "Creating object");
                using (var clientContext = spContext.CreateAppOnlyClientContextForSPHost())
                {
                    clientContext.Load(clientContext.Web, web => web.Title);
                    clientContext.ExecuteQuery();
                    Helper.AddLog("btnRemoveEventReceiver_Click", "Created object");
                    new RemoteEventReceiverManager().RemoveEventReceiversFromHostWeb(clientContext, "RemoteEventReceiver1ItemUpdating");
                    new RemoteEventReceiverManager().RemoveEventReceiversFromHostWeb(clientContext, "RemoteEventReceiver1ItemUpdated");
                    new RemoteEventReceiverManager().RemoveEventReceiversFromHostWeb(clientContext, "RemoteEventReceiver1ItemAdded");
                    Helper.AddLog("btnRemoveEventReceiver_Click", "Removed Receivers");
                }
            }
            catch (Exception ex)
            {
                Helper.AddLog("btnRemoveEventReceiver_Click- Error", ex.ToString());
            }
            List<string> emails = new List<string>();
            emails.Add("me@gauravgoyal.onmicrosoft.com");
            Helper.SendEmail(emails, "Default.aspx : btnRemoveEventReceiver_Click ", Helper.GetLog());
        }

        protected void btnRemoveEventReceiver1_Click(object sender, EventArgs e)
        {
            try
            {
                Helper.AddLog("btnRemoveEventReceiver_Click");
                var contextToken = TokenHelper.GetContextTokenFromRequest(Page.Request);

                // The following code gets the client context and Title property by using TokenHelper.
                // To access other properties, the app may need to request permissions on the host web.
                var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);
                Helper.AddLog("btnRemoveEventReceiver_Click", "Creating object");
                using (var clientContext = spContext.CreateAppOnlyClientContextForSPHost())
                {
                    clientContext.Load(clientContext.Web, web => web.Title);
                    clientContext.ExecuteQuery();
                    Helper.AddLog("btnRemoveEventReceiver_Click", "Created object");
                    new RemoteEventReceiverManager().RemoveEventReceiversFromHostWeb(clientContext, "RemoteEventReceiver1ItemUpdating");
                    Helper.AddLog("btnRemoveEventReceiver_Click", "Removed Receivers");
                }
            }
            catch (Exception ex)
            {
                Helper.AddLog("btnRemoveEventReceiver_Click- Error", ex.ToString());
            }
            List<string> emails = new List<string>();
            emails.Add("me@gauravgoyal.onmicrosoft.com");
            Helper.SendEmail(emails, "Default.aspx : btnRemoveEventReceiver_Click ", Helper.GetLog());
        }

        protected void btnRemoveEventReceiver2_Click(object sender, EventArgs e)
        {
            try
            {
                Helper.AddLog("btnRemoveEventReceiver_Click");
                var contextToken = TokenHelper.GetContextTokenFromRequest(Page.Request);

                // The following code gets the client context and Title property by using TokenHelper.
                // To access other properties, the app may need to request permissions on the host web.
                var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);
                Helper.AddLog("btnRemoveEventReceiver_Click", "Creating object");
                using (var clientContext = spContext.CreateAppOnlyClientContextForSPHost())
                {
                    clientContext.Load(clientContext.Web, web => web.Title);
                    clientContext.ExecuteQuery();
                    Helper.AddLog("btnRemoveEventReceiver_Click", "Created object");
                    new RemoteEventReceiverManager().RemoveEventReceiversFromHostWeb(clientContext, "RemoteEventReceiver1ItemUpdated");
                    Helper.AddLog("btnRemoveEventReceiver_Click", "Removed Receivers");
                }
            }
            catch (Exception ex)
            {
                Helper.AddLog("btnRemoveEventReceiver_Click- Error", ex.ToString());
            }
            List<string> emails = new List<string>();
            emails.Add("me@gauravgoyal.onmicrosoft.com");
            Helper.SendEmail(emails, "Default.aspx : btnRemoveEventReceiver_Click ", Helper.GetLog());
        }

        protected void btnRemoveEventReceiver3_Click(object sender, EventArgs e)
        {
            try
            {
                Helper.AddLog("btnRemoveEventReceiver_Click");
                var contextToken = TokenHelper.GetContextTokenFromRequest(Page.Request);

                // The following code gets the client context and Title property by using TokenHelper.
                // To access other properties, the app may need to request permissions on the host web.
                var spContext = SharePointContextProvider.Current.GetSharePointContext(Context);
                Helper.AddLog("btnRemoveEventReceiver_Click", "Creating object");
                using (var clientContext = spContext.CreateAppOnlyClientContextForSPHost())
                {
                    clientContext.Load(clientContext.Web, web => web.Title);
                    clientContext.ExecuteQuery();
                    Helper.AddLog("btnRemoveEventReceiver_Click", "Created object");
                    new RemoteEventReceiverManager().RemoveEventReceiversFromHostWeb(clientContext, "RemoteEventReceiver1ItemAdded");
                    Helper.AddLog("btnRemoveEventReceiver_Click", "Removed Receivers");
                }
            }
            catch (Exception ex)
            {
                Helper.AddLog("btnRemoveEventReceiver_Click- Error", ex.ToString());
            }
            List<string> emails = new List<string>();
            emails.Add("me@gauravgoyal.onmicrosoft.com");
            Helper.SendEmail(emails, "Default.aspx : btnRemoveEventReceiver_Click ", Helper.GetLog());
        }
    }
}