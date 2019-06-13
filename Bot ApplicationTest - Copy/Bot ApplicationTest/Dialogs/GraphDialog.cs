using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace Bot_ApplicationTest.Dialogs
{
   [Serializable]
    public class GraphDialog:IDialog<object>
    {
        public async Task StartAsync(IDialogContext context)
        {
            await context.PostAsync("What can I help you?");
            
            Activity reply=((Activity)context.Activity).CreateReply();

            var message = reply.Text;
            await context.PostAsync(message.ToString());
           context.Done(true);


        }
    }
}