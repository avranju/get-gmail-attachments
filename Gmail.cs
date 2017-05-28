using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Google.Apis.Gmail.v1.Data;

namespace get_gmail_attachments
{
    public class Gmail
    {
        public const string UserMe = "me";

        private readonly GmailService _mailService;

        public Gmail(GmailService mailService)
        {
            _mailService = mailService;
        }

        public static async Task<GmailService> Authenticate()
        {
            UserCredential credential;
            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    stream,
                    new[] {GmailService.Scope.GmailReadonly},
                    "user",
                    CancellationToken.None,
                    new FileDataStore("Gmail.Attachments"));
            }

            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "get-gmail-attachments"
            });

            return service;
        }

        public async Task<IList<Message>> Search(string searchString)
        {
            var messages = new List<Message>();
            var request = _mailService.Users.Messages.List(UserMe);
            request.Q = searchString;
            var response = await request.ExecuteAsync();

            while (response.Messages != null)
            {
                messages.AddRange(response.Messages);
                if (!string.IsNullOrEmpty(response.NextPageToken))
                {
                    request = _mailService.Users.Messages.List(UserMe);
                    request.Q = searchString;
                    request.PageToken = response.NextPageToken;
                    response = await request.ExecuteAsync();
                }
                else
                {
                    break;
                }
            }

            return messages;
        }

        public async Task<Message> GetMessage(Message msg)
        {
            return await _mailService.Users.Messages.Get(UserMe, msg.Id).ExecuteAsync();
        }

        public async Task<MessagePartBody> GetAttachment(Message message, string attachmentId)
        {
            return await _mailService.Users.Messages.Attachments.Get(UserMe, message.Id, attachmentId).ExecuteAsync();
        }
    }
}
