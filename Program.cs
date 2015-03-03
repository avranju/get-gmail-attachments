using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace get_gmail_attachments
{
    class Program
    {
        private const string USER_ME = "me";
    
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: get-gmail-attachments <search criteria> <path to write attachments to>");
                return;
            }

            string searchString = args[0];
            string filePath = args[1];

            try
            {
                Run(searchString, filePath).Wait();
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: {0}", ex);
            }
        }

        async static Task Run(string searchString, string filePath)
        {
            List<Message> messages = new List<Message>();
            var mailService = await Authenticate();
            var request = mailService.Users.Messages.List(USER_ME);
            request.Q = searchString;
            var response = await request.ExecuteAsync();

            while (response.Messages != null)
            {
                messages.AddRange(response.Messages);
                if (!String.IsNullOrEmpty(response.NextPageToken))
                {
                    request = mailService.Users.Messages.List(USER_ME);
                    request.Q = searchString;
                    request.PageToken = response.NextPageToken;
                    response = await request.ExecuteAsync();
                }
                else
                {
                    break;
                }
            }

            foreach (var message in messages)
            {
                var msg = await mailService.Users.Messages.Get(USER_ME, message.Id).ExecuteAsync();
                foreach (var part in msg.Payload.Parts)
                {
                    if (!String.IsNullOrEmpty(part.Filename))
                    {
                        Console.WriteLine("Getting " + part.Filename);
                        var attachment = await mailService.Users.Messages.
                                Attachments.Get(USER_ME, message.Id, part.Body.AttachmentId).ExecuteAsync();
                        byte[] attachmentData = Convert.FromBase64String(attachment.Data.Replace('-', '+').Replace('_', '/'));
                        string fileName = Path.Combine(filePath, CleanFileName(part.Filename));
                        File.WriteAllBytes(fileName, attachmentData);
                    }
                }
            }
        }

        /// <summary>
        /// Taken from http://stackoverflow.com/a/7393722/8080.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private static string CleanFileName(string fileName)
        {
            return Path.GetInvalidFileNameChars().Aggregate(fileName, (current, c) => current.Replace(c, '_'));
        }

        async static Task<GmailService> Authenticate()
        {
            UserCredential credential;
            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    stream,
                    new[] { GmailService.Scope.GmailReadonly },
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
    }
}
