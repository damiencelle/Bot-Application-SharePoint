using AdaptiveCards;
using AuthBot;
using AuthBot.Dialogs;
using AuthBot.Models;
using Bot_Application2.Common;
using Bot_Application2.Luis;
using Bot_Application2.Resources;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Web;

namespace Bot_Application2.Dialogs
{
    [Serializable]
    public class RootDialog : IDialog<object>
    {
        #region ########## ATTRIBUTES / PROPERTIES ############################

        /// <summary>
        /// The authentication result
        /// </summary>
        private AuthResult _authResult;

        #endregion ####### ATTRIBUTES / PROPERTIES ############################

        #region ########## GENERIC ############################################

        /// <summary>
        /// The start of the code that represents the conversational dialog.
        /// </summary>
        /// <param name="context">The dialog context.</param>
        /// <returns>
        /// A task that represents the dialog start.
        /// </returns>
        public Task StartAsync(IDialogContext context)
        {
            context.Wait(MessageReceivedAsync);
            return Task.CompletedTask;
        }

        /// <summary>
        /// Messages the received asynchronous.
        /// </summary>
        /// <param name="context">The context.</param>
        /// <param name="result">The result.</param>
        /// <returns></returns>
        private async Task MessageReceivedAsync(IDialogContext context, IAwaitable<object> result)
        {
            var message = await result as Activity;

            // Check authentication
            if (string.IsNullOrEmpty(await context.GetAccessToken(ConfigurationManager.AppSettings["ActiveDirectory.ResourceId"])))
            {
                // Run authentication dialog
                await context.Forward(new AzureAuthDialog(ConfigurationManager.AppSettings["ActiveDirectory.ResourceId"]), this.ResumeAfterAuth, message, CancellationToken.None);
            }
            else
            {
                string sresourceId = ConfigurationManager.AppSettings["ActiveDirectory.Tenant"].Split('.')[0];
                var tenantUrl = $"https://{sresourceId}-admin.sharepoint.com";

                context.UserData.TryGetValue(ContextConstants.AuthResultKey, out _authResult);

                AuthenticationManager authManager = new AuthenticationManager();

                using (ClientContext ctx = authManager.GetAzureADAccessTokenAuthenticatedContext(tenantUrl, _authResult.AccessToken))
                {
                    await ProcessMessage(context, ctx, message);
                }
                context.Wait(MessageReceivedAsync);
            }
        }

        #endregion ####### GENERIC ############################################

        #region ########## AUTHENTICATION #####################################

        /// <summary>
        /// Resumes the after authentication.
        /// </summary>
        /// <param name="context">The context.</param>
        /// <param name="result">The result.</param>
        /// <returns></returns>
        private async Task ResumeAfterAuth(IDialogContext context, IAwaitable<string> result)
        {
            var message = await result;
            await context.PostAsync(message);
            await context.PostAsync(AppResources.WhatWouldYouLikeMeToDo);
            context.Wait(MessageReceivedAsync);
        }

        #endregion ####### AUTHENTICATION #####################################

        #region ########## PROCESS MESSAGES ###################################

        /// <summary>
        /// Processes the message.
        /// </summary>
        /// <param name="context">The context.</param>
        /// <param name="ctx">The CTX.</param>
        /// <param name="message">The message.</param>
        /// <returns></returns>
        private async Task ProcessMessage(IDialogContext context, ClientContext ctx, Activity message)
        {
            if (!String.IsNullOrEmpty(message.Text))
            {
                String treatedMessage = await CallLuis(message.Text);

                switch (treatedMessage)
                {
                    case AppConstants.HELLO:
                        await SayHelloToUser(context, ctx, message);
                        break;
                    case AppConstants.SUGGESTIONS_SHOW:
                        await ShowSuggestions(context, ctx, message);
                        break;
                    case AppConstants.SITE_COLLECTIONS_SHOW:
                        await ShowAllSiteCollections(context, ctx);
                        break;
                    case AppConstants.LOGO_CHANGE:
                        await ShowLogoChangePage(context, ctx, message);
                        break;
                    case AppConstants.SUBSITE_CREATE:
                        await CreateSubsite(context, ctx, message);
                        break;
                    case AppConstants.TEST:
                        await Test(context, ctx, message);
                        break;
                    default:
                        await ShowSuggestions(context, ctx, message);
                        break;
                }
            }
            else
            {
                if (message.Value.GetType() == typeof(JObject))
                {
                    JObject messageJsonObject = message.Value as JObject;

                    JToken spSiteUrl;
                    JToken newSubsiteName;
                    JToken subsiteWebTemplate;

                    bool existSpSiteUrl = messageJsonObject.TryGetValue(AppConstants.SP_SITE, out spSiteUrl);
                    bool existNewSubsiteName = messageJsonObject.TryGetValue(AppConstants.SUBSITE_NAME, out newSubsiteName);
                    bool existSubsiteWebTemplate = messageJsonObject.TryGetValue(AppConstants.SP_WEBTEMPATE, out subsiteWebTemplate);

                    if (existSpSiteUrl && existNewSubsiteName && existSubsiteWebTemplate)
                    {
                        await CreateSubsiteOnSharePoint(context, ctx, message, spSiteUrl.ToString(), newSubsiteName.ToString(), subsiteWebTemplate.ToString());
                    }
                }
            }
        }

        /// <summary>
        /// Says the hello to user.
        /// </summary>
        /// <param name="context">The context.</param>
        /// <param name="ctx">The CTX.</param>
        /// <param name="message">The message.</param>
        /// <returns></returns>
        private async Task SayHelloToUser(IDialogContext context, ClientContext ctx, Activity message)
        {
            Activity reply = message.CreateReply(AppResources.HiImAnOffice365Bot);
            reply.Type = ActivityTypes.Message;
            reply.TextFormat = TextFormatTypes.Plain;

            await context.PostAsync(reply);
        }

        /// <summary>
        /// Calls the luis.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <returns></returns>
        private async Task<string> CallLuis(string text)
        {
            string treatedMessage = string.Empty;
            LuisJSON data = new LuisJSON();

            using (HttpClient client = new HttpClient())
            {
                string requestUri = ConfigurationManager.AppSettings["LuisServiceURI"];

                string finalQuery = String.Concat(requestUri, HttpUtility.UrlEncode(text));

                HttpResponseMessage msg = await client.GetAsync(finalQuery);

                if (msg.IsSuccessStatusCode)
                {
                    var jsonDataResponse = await msg.Content.ReadAsStringAsync();
                    data = JsonConvert.DeserializeObject<LuisJSON>(jsonDataResponse);

                    if(data.TopScoringIntent.Score > AppConstants.LUIS_TOP_SCORING_INTENT)
                    {
                        // We consider that Luis result is correct
                        treatedMessage = data.TopScoringIntent.intent;
                    }
                }
            }

            return treatedMessage;
        }

        private async Task Test(IDialogContext context, ClientContext ctx, Activity message)
        {
            Tenant tenant = new Tenant(ctx);
            
            ctx.Load(tenant);

            ctx.ExecuteQuery();

            IList<OfficeDevPnP.Core.Entities.SiteEntity> cols = tenant.GetOneDriveSiteCollections();

            foreach (OfficeDevPnP.Core.Entities.SiteEntity col in cols)
            {
                Console.WriteLine(col.Title);
            }
        }

        /// <summary>
        /// Creates the subsite on share point.
        /// </summary>
        /// <param name="context">The context.</param>
        /// <param name="ctx">The CTX.</param>
        /// <param name="message">The message.</param>
        /// <param name="spSiteUrl">The sp site URL.</param>
        /// <param name="newSubsiteName">New name of the subsite.</param>
        /// <param name="webTemplate">The web template.</param>
        /// <returns></returns>
        private async Task CreateSubsiteOnSharePoint(IDialogContext context, ClientContext ctx, Activity message, string spSiteUrl, string newSubsiteName, string webTemplate)
        {
            Tenant tenant = new Tenant(ctx);

            WebCreationInformation wci = new WebCreationInformation()
            {
                WebTemplate = webTemplate,
                Description = newSubsiteName,
                Title = newSubsiteName,
                Url = HttpUtility.UrlEncode(newSubsiteName),
                UseSamePermissionsAsParentSite = true
            };

            Site parentSite = tenant.GetSiteByUrl(spSiteUrl);
            Web newWeb = parentSite.RootWeb.Webs.Add(wci);
            ctx.Load(newWeb);
            ctx.ExecuteQuery();

            // You can access your new subsite by clicking on this link
            Activity reply = message.CreateReply(AppResources.YourSubsiteHasBeenCreated);
            reply.Type = ActivityTypes.Message;
            reply.TextFormat = TextFormatTypes.Plain;

            reply.SuggestedActions = new SuggestedActions();
            reply.SuggestedActions.Actions = new List<CardAction>();
            CardAction newSiteLink = new CardAction()
            {
                Title = AppResources.GoToNewCreatedSubsite,
                Type = ActionTypes.OpenUrl,
                Value = newWeb.Url
            };
            reply.SuggestedActions.Actions.Add(newSiteLink);

            await context.PostAsync(reply);
        }

        /// <summary>
        /// Shows the suggestions.
        /// </summary>
        /// <param name="context">The context.</param>
        /// <param name="ctx">The CTX.</param>
        /// <param name="message">The message.</param>
        /// <returns></returns>
        private async Task ShowSuggestions(IDialogContext context, ClientContext ctx, Activity message)
        {
            var reply = message.CreateReply(AppResources.TheseAreActionsICanDo);
            reply.Type = ActivityTypes.Message;
            reply.TextFormat = TextFormatTypes.Plain;

            reply.SuggestedActions = new SuggestedActions()
            {
                Actions = new List<CardAction>()
                {
                    new CardAction(){ Title = AppResources.ShowAllSiteCollections, Type=ActionTypes.PostBack, Value=AppConstants.SITE_COLLECTIONS_SHOW },
                    new CardAction(){ Title = AppResources.ChangeSiteCollectionLogo, Type=ActionTypes.PostBack, Value=AppConstants.LOGO_CHANGE },
                    new CardAction(){ Title = AppResources.CreateASubsite, Type=ActionTypes.PostBack, Value=AppConstants.SUBSITE_CREATE },
                    new CardAction(){ Title = AppResources.DoSomethingElse, Type=ActionTypes.PostBack, Value=AppConstants.SUGGESTIONS_SHOW }
                }
            };

            await context.PostAsync(reply);
        }

        /// <summary>
        /// Shows the logo change page.
        /// </summary>
        /// <param name="context">The context.</param>
        /// <param name="ctx">The CTX.</param>
        /// <param name="message">The message.</param>
        /// <returns></returns>
        private async Task ShowLogoChangePage(IDialogContext context, ClientContext ctx, Activity message)
        {
            Activity reply = message.CreateReply(AppResources.OnWhichSPSiteDoYouWantToChangeTheLogo);
            reply.Type = ActivityTypes.Message;
            reply.TextFormat = TextFormatTypes.Plain;

            reply.SuggestedActions = new SuggestedActions();
            reply.SuggestedActions.Actions = new List<CardAction>();

            // List all the site collections for the tenant
            SPOSitePropertiesEnumerable prop = null;

            Tenant tenant = new Tenant(ctx);
            prop = tenant.GetSiteProperties(0, true);
            ctx.Load(prop);
            ctx.ExecuteQuery();

            foreach (SiteProperties sp in prop)
            {
                if (!String.IsNullOrEmpty(sp.Title))
                {
                    CardAction spSiteCardAction = new CardAction()
                    {
                        Title = sp.Title,
                        Type = ActionTypes.OpenUrl,
                        Value = String.Concat(sp.Url, AppConstants.PRJ_SETTINGS_URL)
                    };
                    reply.SuggestedActions.Actions.Add(spSiteCardAction);
                }
            }
            await context.PostAsync(reply);
        }

        /// <summary>
        /// Creates the subsite.
        /// </summary>
        /// <param name="context">The context.</param>
        /// <param name="ctx">The CTX.</param>
        /// <param name="message">The message.</param>
        /// <returns></returns>
        private async Task CreateSubsite(IDialogContext context, ClientContext ctx, Activity message)
        {
            var connector = new ConnectorClient(new Uri(message.ServiceUrl));
            Activity replyToConversation = message.CreateReply(AppResources.CreateSubsite);
            replyToConversation.Attachments = new List<Microsoft.Bot.Connector.Attachment>();

            AdaptiveCard card = new AdaptiveCard();

            // Specify speech for the card.
            card.Speak = AppResources.PleaseFillInformationToCreateSubsite;

            // Add text to the card.
            card.Body.Add(new TextBlock()
            {
                Text = AppResources.CreateNewSubsite,
                Size = TextSize.Large,
                Weight = TextWeight.Bolder
            });

            // Add text to the card.
            card.Body.Add(new TextBlock()
            {
                Text = AppResources.SelectSPSite
            });

            ChoiceSet spsites = new ChoiceSet()
            {
                Id = AppConstants.SP_SITE,
                Style = ChoiceInputStyle.Compact,
            };

            // List all the site collections for the tenant
            SPOSitePropertiesEnumerable prop = null;

            Tenant tenant = new Tenant(ctx);
            prop = tenant.GetSiteProperties(0, true);
            ctx.Load(prop);
            ctx.ExecuteQuery();

            foreach (SiteProperties sp in prop)
            {
                if (!String.IsNullOrEmpty(sp.Title))
                {
                    spsites.Choices.Add(
                        new Choice()
                        {
                            Speak = sp.Title,
                            Title = sp.Title,
                            Value = sp.Url
                        });
                }
            }

            card.Body.Add(spsites);

            // Add text to the card.
            card.Body.Add(new TextBlock()
            {
                Text = AppResources.NewSubsiteName
            });

            card.Body.Add(new TextInput()
            {
                Id = AppConstants.SUBSITE_NAME,
                IsMultiline = false
            });


            // Add text to the card.
            card.Body.Add(new TextBlock()
            {
                Text = AppResources.WebTemplateToApply
            });

            ctx.Load(ctx.Web);
            ctx.ExecuteQuery();

            SPOTenantWebTemplateCollection wtc = tenant.GetSPOTenantWebTemplates(ctx.Web.Language, 15);
            ctx.Load(wtc);
            ctx.ExecuteQuery();

            ChoiceSet spWebTemlate = new ChoiceSet()
            {
                Id = AppConstants.SP_WEBTEMPATE,
                Style = ChoiceInputStyle.Compact,
            };

            foreach (SPOTenantWebTemplate wt in wtc)
            {
                if (!String.IsNullOrEmpty(wt.Title))
                {
                    spWebTemlate.Choices.Add(
                        new Choice()
                        {
                            Speak = wt.Title,
                            Title = wt.Title,
                            Value = wt.Name
                        });
                }
            }

            card.Body.Add(spWebTemlate);

            // Add buttons to the card.
            card.Actions.Add(new SubmitAction()
            {
                Title = AppResources.Save
            });

            // Create the attachment.
            Microsoft.Bot.Connector.Attachment attachment = new Microsoft.Bot.Connector.Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card
            };

            replyToConversation.Attachments.Add(attachment);

            var reply = await connector.Conversations.SendToConversationAsync(replyToConversation);
        }

        /// <summary>
        /// Shows all site collections.
        /// </summary>
        /// <param name="context">The context.</param>
        /// <param name="ctx">The CTX.</param>
        /// <returns></returns>
        private async Task ShowAllSiteCollections(IDialogContext context, ClientContext ctx)
        {
            // List all the site collections for the tenant
            SPOSitePropertiesEnumerable prop = null;

            Tenant tenant = new Tenant(ctx);
            prop = tenant.GetSiteProperties(0, true);
            ctx.Load(prop);
            ctx.ExecuteQuery();

            foreach (SiteProperties sp in prop)
            {
                string strSiteColName = AppResources.SiteCollection + sp.Title + " => " + sp.Url;
                await context.PostAsync(strSiteColName);
            }
        }

        #endregion ####### PROCESS MESSAGES ###################################
    }
}