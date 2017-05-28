using Google.Apis.Gmail.v1.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace get_gmail_attachments
{
    class Program
    {
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

        static async Task Run(string searchString, string filePath)
        {
            var gmail = new Gmail(await Gmail.Authenticate(appName: "get-gmail-attachments"));
            IList<Message> messages = await gmail.Search(searchString);

            foreach (var message in messages)
            {
                var msg = await gmail.GetMessage(message);
                foreach (var part in msg.Payload.Parts)
                {
                    if (!String.IsNullOrEmpty(part.Filename))
                    {
                        Console.WriteLine("Getting " + part.Filename);
                        var attachment = await gmail.GetAttachment(message, part.Body.AttachmentId);
                        byte[] attachmentData = Convert.FromBase64String(attachment.Data.Replace('-', '+').Replace('_', '/'));
                        string fileName = Path.Combine(filePath, CleanFileName(part.Filename));
                        File.WriteAllBytes(fileName, attachmentData);
                    }
                }
            }
        }

        static async Task MarkAsRead(string searchString)
        {
            var gmail = new Gmail(await Gmail.Authenticate(appName: "mark-as-read"));

            Console.WriteLine($"Searching for messages matching search criteria: {searchString}");
            IList<Message> messages = await gmail.Search(searchString);

            Console.WriteLine($"Marking {messages.Count} messages as read. Hit return to continue.");

            // We don't parallelize requests as much as we can because we don't want to hit
            // Gmail API rate limits (i.e. the number of requests/second limit).
            int count = 0;
            foreach (var message in messages)
            {
                await gmail.MarkAsRead(message);
                Console.Write($"{++count} / {messages.Count} done.\r");
            }

            Console.WriteLine("All done.");
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
    }
}
