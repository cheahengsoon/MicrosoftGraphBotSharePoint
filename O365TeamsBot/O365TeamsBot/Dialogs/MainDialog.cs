// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

using RestSharp;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;

namespace O365TeamsBot.Dialogs
{
   public class MainDialog : LogoutDialog
    {
        protected readonly ILogger Logger;

        public MainDialog(IConfiguration configuration, ILogger<MainDialog> logger)
            : base(nameof(MainDialog), configuration["ConnectionName"])
        {
            Logger = logger;

            AddDialog(new OAuthPrompt(
                nameof(OAuthPrompt),
                new OAuthPromptSettings
                {
                    ConnectionName = ConnectionName,
                    Text = "Please Sign In",
                    Title = "Sign In",
                    Timeout = 300000, // User has 5 minutes to login (1000 * 60 * 5)
                }));

            AddDialog(new ConfirmPrompt(nameof(ConfirmPrompt)));
            AddDialog(new TextPrompt(nameof(TextPrompt)));
            AddDialog(new WaterfallDialog(nameof(WaterfallDialog), new WaterfallStep[]
            {
                PromptStepAsync,
                LoginStepAsync,
                DisplayTokenPhase1Async,
                DisplayTokenPhase2Async,
            }));

            // The initial child Dialog to run.
            InitialDialogId = nameof(WaterfallDialog);
        }

        private async Task<DialogTurnResult> PromptStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken);
        }

        private async Task<DialogTurnResult> LoginStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            // Get the token from the previous step. Note that we could also have gotten the
            // token directly from the prompt itself. There is an example of this in the next method.
            var tokenResponse = (TokenResponse)stepContext.Result;
            if (tokenResponse != null)
            {
             
                 await stepContext.Context.SendActivityAsync(MessageFactory.Text("You are now logged in."), cancellationToken);
                return await stepContext.PromptAsync(nameof(TextPrompt), new PromptOptions { Prompt = MessageFactory.Text("Would you like to search from SharePoint?") }, cancellationToken);
            }

            await stepContext.Context.SendActivityAsync(MessageFactory.Text("Login was not successful please try again."), cancellationToken);
            return await stepContext.EndDialogAsync();
        }

        private async Task<DialogTurnResult> DisplayTokenPhase1Async(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {

            if (stepContext.Result != null)
            {
                try {
                    var client = new RestClient("https://login.microsoftonline.com/de0122d8-4332-478d-93c7-38594531094a/oauth2/v2.0/token");
                    var request = new RestRequest(Method.POST);
                    request.AddHeader("cache-control", "no-cache");
                    request.AddHeader("Connection", "keep-alive");
                    request.AddHeader("content-length", "186");
                    request.AddHeader("accept-encoding", "gzip, deflate");
                    request.AddHeader("cookie", "fpc=AmGuu7W1bCFBtBkWG1xoSP4zi0p8AQAAALm4j9QOAAAA; x-ms-gateway-slice=prod; stsservicecookie=ests");
                    request.AddHeader("Host", "login.microsoftonline.com");
                    request.AddHeader("Postman-Token", "bcc747d6-f599-461e-abf8-82a68dd0a79c,096392f7-6472-404f-876c-06e5f4cf60a0");
                    request.AddHeader("Cache-Control", "no-cache");
                    request.AddHeader("Accept", "*/*");
                    request.AddHeader("User-Agent", "PostmanRuntime/7.13.0");
                    request.AddHeader("SdkVersion", "postman-graph/v1.0");
                    request.AddHeader("Content-Type", "application/x-www-form-urlencoded");
                    request.AddParameter("undefined", "grant_type=client_credentials&client_id=5f3cbb90-75e3-4143-afc0-1b192a124e79&client_secret=g1mgCHWr%2FVD2%2B5%2FZ5Ly%2B-YdH8%2F%401PkX1&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default", ParameterType.RequestBody);
                    IRestResponse response = client.Execute(request);

                    GraphData token = JsonConvert.DeserializeObject<GraphData>(response.Content);

                    var client2 = new RestClient($"https://graph.microsoft.com/v1.0/sites/root/drive/root/search(q='{stepContext.Result}')");
                    var request2 = new RestRequest(Method.GET);
                    request2.AddHeader("cache-control", "no-cache");
                    request2.AddHeader("Connection", "keep-alive");
                    request2.AddHeader("accept-encoding", "gzip, deflate");
                    request2.AddHeader("Postman-Token", "8849d8c8-ec94-4aff-b524-fcb53e3f7c59,7696f547-3413-40f8-9a4c-ba214f50ddbe");
                    request2.AddHeader("Cache-Control", "no-cache");
                    request2.AddHeader("Accept", "*/*");
                    request2.AddHeader("User-Agent", "PostmanRuntime/7.13.0");
                    request2.AddHeader("Host", "graph.microsoft.com");
                    request2.AddHeader("Content-Length", "0");
                    request2.AddHeader("Content-Type", "application/json");
                    request2.AddHeader("Authorization", "Bearer " + token.access_token);
                    IRestResponse response2 = client2.Execute(request2);

                  //  await stepContext.Context.SendActivityAsync(MessageFactory.Text($"You say : {response2.Content}"), cancellationToken);

                    var cardActivity = Activity.CreateMessageActivity();


                    string sampleJson = response2.Content.ToString();
                    JObject results = JObject.Parse(sampleJson);
                    foreach (var resulto in results["value"])
                    {
                        HeroCard card = new HeroCard
                        {
                            Title = (string)resulto["name"],
                            //Images = new List<CardImage> { new CardImage(System.Web.HttpContext.Current.Server.MapPath(@"~\Images\document.png")) },
                            Buttons = new List<CardAction> { new CardAction(ActionTypes.OpenUrl, "View Online", value: (string)resulto["webUrl"]) }

                        };
                        var cardAttachment = card.ToAttachment();
                        cardActivity.Attachments.Add(cardAttachment);
                    }

                    cardActivity.AttachmentLayout = AttachmentLayoutTypes.Carousel;
                    await stepContext.Context.SendActivityAsync(cardActivity, cancellationToken);
                }
                catch (Exception ex)
                {
                    await stepContext.Context.SendActivityAsync(MessageFactory.Text("Error : "+ex.Message.ToString()), cancellationToken);
                    return await stepContext.EndDialogAsync();
                }
            }
   
            return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
        }

        private async Task<DialogTurnResult> DisplayTokenPhase2Async(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var tokenResponse = (TokenResponse)stepContext.Result;
            if (tokenResponse != null)
            {
                await stepContext.Context.SendActivityAsync(MessageFactory.Text($"Here is your token {tokenResponse.Token}"), cancellationToken);
            }

            return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
        }
    }
}

public class GraphData
{
    public string access_token { get; set; }
    public string name { get; set; }
    public string webUrl { get; set; }

    public IList<string> value { get; set; }
}
