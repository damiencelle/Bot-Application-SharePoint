using AdaptiveCards;
using AuthBot;
using AuthBot.Dialogs;
using AuthBot.Models;
using Bot_Application2.Luis;
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

        private AuthResult _authResult;

        #endregion ####### ATTRIBUTES / PROPERTIES ############################

        #region ########## GENERIC ############################################

        public Task StartAsync(IDialogContext context)
        {
            context.Wait(MessageReceivedAsync);
            return Task.CompletedTask;
        }

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

        private async Task ResumeAfterAuth(IDialogContext context, IAwaitable<string> result)
        {
            var message = await result;
            await context.PostAsync(message);
            await context.PostAsync("What would you like me to do?");
            context.Wait(MessageReceivedAsync);
        }

        #endregion ####### AUTHENTICATION #####################################

        #region ########## PROCESS MESSAGES ###################################

        private async Task ProcessMessage(IDialogContext context, ClientContext ctx, Activity message)
        {
            if (!String.IsNullOrEmpty(message.Text))
            {
                String treatedMessage = await CallLuis(message.Text);

                switch (treatedMessage)
                {
                    case "Hello":
                        await SayHelloToUser(context, ctx, message);
                        break;
                    case "Suggestions.Show":
                        await ShowSuggestions(context, ctx, message);
                        break;
                    case "SiteCollections.Show":
                        await ShowAllSiteCollections(context, ctx);
                        break;
                    case "Logo.Change":
                        await ShowLogoChangePage(context, ctx, message);
                        break;
                    case "Subsite.Create":
                        await CreateSubsite(context, ctx, message);
                        break;
                    case "Test":
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

                    bool existSpSiteUrl = messageJsonObject.TryGetValue("SpSite", out spSiteUrl);
                    bool existNewSubsiteName = messageJsonObject.TryGetValue("SubsiteName", out newSubsiteName);
                    bool existSubsiteWebTemplate = messageJsonObject.TryGetValue("SpWebTemplate", out subsiteWebTemplate);

                    if (existSpSiteUrl && existNewSubsiteName && existSubsiteWebTemplate)
                    {
                        await CreateSubsiteOnSharePoint(context, ctx, message, spSiteUrl.ToString(), newSubsiteName.ToString(), subsiteWebTemplate.ToString());
                    }
                }
            }
        }

        private async Task SayHelloToUser(IDialogContext context, ClientContext ctx, Activity message)
        {
            Activity reply = message.CreateReply("Hi I'm an Office 365 Bot. What can I do for you ?");
            reply.Type = ActivityTypes.Message;
            reply.TextFormat = TextFormatTypes.Plain;

            await context.PostAsync(reply);
        }

        private async Task<string> CallLuis(string text)
        {
            string treatedMessage = string.Empty;
            LuisJSON data = new LuisJSON();

            using (HttpClient client = new HttpClient())
            {
                string requestUri = "https://westus.api.cognitive.microsoft.com/luis/v2.0/apps/b609a2c5-3a83-4cc6-b72a-0052b6453821?subscription-key=8d70230b2ebc4fffb63fc58a3b2a4c4a&staging=true&verbose=true&timezoneOffset=0&q=";

                string finalQuery = String.Concat(requestUri, HttpUtility.UrlEncode(text));

                HttpResponseMessage msg = await client.GetAsync(finalQuery);

                if (msg.IsSuccessStatusCode)
                {
                    var jsonDataResponse = await msg.Content.ReadAsStringAsync();
                    data = JsonConvert.DeserializeObject<LuisJSON>(jsonDataResponse);

                    if(data.TopScoringIntent.Score > 0.95)
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
            Activity reply = message.CreateReply("Your subsite has been created.");
            reply.Type = ActivityTypes.Message;
            reply.TextFormat = TextFormatTypes.Plain;

            reply.SuggestedActions = new SuggestedActions();
            reply.SuggestedActions.Actions = new List<CardAction>();
            CardAction newSiteLink = new CardAction()
            {
                Title = "Go to new created subsite",
                Type = ActionTypes.OpenUrl,
                Value = newWeb.Url
            };
            reply.SuggestedActions.Actions.Add(newSiteLink);

            await context.PostAsync(reply);
        }

        private async Task ShowSuggestions(IDialogContext context, ClientContext ctx, Activity message)
        {
            var reply = message.CreateReply("These are actions I can do. What would you like me to do?");
            reply.Type = ActivityTypes.Message;
            reply.TextFormat = TextFormatTypes.Plain;

            reply.SuggestedActions = new SuggestedActions()
            {
                Actions = new List<CardAction>()
                {
                    new CardAction(){ Title = "Show All Site Collections", Type=ActionTypes.PostBack, Value="ShowAllSiteCollections" },
                    new CardAction(){ Title = "Change Site Collection Logo", Type=ActionTypes.PostBack, Value="ShowLogoChangePage" },
                    new CardAction(){ Title = "Create a subsite", Type=ActionTypes.PostBack, Value="CreateASubsite" },
                    new CardAction(){ Title = "Do something else", Type=ActionTypes.PostBack, Value="ShowSuggestions" }
                }
            };

            await context.PostAsync(reply);
        }

        private async Task ShowLogoChangePage(IDialogContext context, ClientContext ctx, Activity message)
        {
            Activity reply = message.CreateReply("On which SPSite do you want to change the logo ? The click on the following links redirects you to display settings page.");
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
                        Value = String.Concat(sp.Url, "/_layouts/15/prjsetng.aspx")
                    };
                    reply.SuggestedActions.Actions.Add(spSiteCardAction);
                }
            }
            await context.PostAsync(reply);
        }

        private async Task CreateSubsite(IDialogContext context, ClientContext ctx, Activity message)
        {
            var connector = new ConnectorClient(new Uri(message.ServiceUrl));
            Activity replyToConversation = message.CreateReply("Create Subsite");
            replyToConversation.Attachments = new List<Microsoft.Bot.Connector.Attachment>();

            AdaptiveCard card = new AdaptiveCard();

            // Specify speech for the card.
            card.Speak = "<s>Please fill informations to create subsite.</s>";

            // Add text to the card.
            card.Body.Add(new TextBlock()
            {
                Text = "Create new subsite",
                Size = TextSize.Large,
                Weight = TextWeight.Bolder
            });

            // Add text to the card.
            card.Body.Add(new TextBlock()
            {
                Text = "Select SPSite"
            });

            ChoiceSet spsites = new ChoiceSet()
            {
                Id = "SpSite",
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
                Text = "New Subsite Name"
            });

            card.Body.Add(new TextInput()
            {
                Id = "SubsiteName",
                IsMultiline = false
            });


            // Add text to the card.
            card.Body.Add(new TextBlock()
            {
                Text = "Web template to apply"
            });

            ctx.Load(ctx.Web);
            ctx.ExecuteQuery();

            SPOTenantWebTemplateCollection wtc = tenant.GetSPOTenantWebTemplates(ctx.Web.Language, 15);
            ctx.Load(wtc);
            ctx.ExecuteQuery();

            ChoiceSet spWebTemlate = new ChoiceSet()
            {
                Id = "SpWebTemplate",
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
                Title = "Save"
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
                string strSiteColName = "Site Collection: " + sp.Title + " => " + sp.Url;
                await context.PostAsync(strSiteColName);
            }
        }

        #endregion ####### PROCESS MESSAGES ###################################
    }
}