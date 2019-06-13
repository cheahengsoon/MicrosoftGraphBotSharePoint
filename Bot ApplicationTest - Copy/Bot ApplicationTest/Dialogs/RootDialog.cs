using System;
using System.Threading.Tasks;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using AuthBot;
using System.Configuration;
using AuthBot.Dialogs;
using System.Threading;
using RestSharp;
using Newtonsoft.Json.Linq;
using System.IO;
using System.Text;
using System.Runtime.Serialization.Json;
using System.Runtime.Serialization;
using System.Web.Script.Serialization;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace Bot_ApplicationTest.Dialogs
{
    [Serializable]
    public class RootDialog : IDialog<object>
    {
        public Task StartAsync(IDialogContext context)
        {
            context.Wait(MessageReceivedAsync);

            return Task.CompletedTask;
        }

        private async Task MessageReceivedAsync(IDialogContext context, IAwaitable<object> result)
        {
            var message = await result as Activity;
           

            // 認証チェック
            if (string.IsNullOrEmpty(await context.GetAccessToken(ConfigurationManager.AppSettings["ActiveDirectory.ResourceId"])))
            {
                // 認証ダイアログの実行
                await context.Forward(new AzureAuthDialog(ConfigurationManager.AppSettings["ActiveDirectory.ResourceId"]), this.ResumeAfterAuth, message, CancellationToken.None);

            }
            else
            {
               

              //  await context.PostAsync($"Your said {message.Text}");

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

                JavaScriptSerializer js = new JavaScriptSerializer();
                GraphData blogObject = js.Deserialize<GraphData>(response.Content);
                var tokenkey = blogObject.access_token;

                var client2 = new RestClient($"https://graph.microsoft.com/v1.0/sites/root/drive/root/search(q='{message.Text}')");
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
                request2.AddHeader("Authorization", "Bearer " + tokenkey.ToString());
                IRestResponse response2 = client2.Execute(request2);

                //  await context.PostAsync(response2.Content.ToString());
                Activity reply = ((Activity)context.Activity).CreateReply();

                string sampleJSon = response2.Content.ToString();

                JObject results = JObject.Parse(sampleJSon);

                foreach(var resulto in results["value"])
                {
              
                    HeroCard card= new HeroCard
                    {
                        Title=(string)resulto["name"],
                        //Images = new List<CardImage> { new CardImage(System.Web.HttpContext.Current.Server.MapPath(@"~\Images\document.png")) },
                        Buttons = new List<CardAction> { new CardAction(ActionTypes.OpenUrl, "View Online", value: (string)resulto["webUrl"]) }
                    
                };
                    reply.Attachments.Add(card.ToAttachment());
                }
                reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;

                await context.PostAsync(reply);

               // context.Wait(MessageReceivedAsync);
            }

        }

        private async Task ResumeAfterAuth(IDialogContext context, IAwaitable<string> result)
        {
          
          
           await context.PostAsync("What can I help you?");

            context.Wait(MessageReceivedAsync);

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


