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
                // Run(searchString, filePath).Wait();
                MarkAsNotSpam(searchString).Wait();
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

        static async Task RunOp(string opName, string searchString, Func<Gmail, Message, Task<bool>> op)
        {
            var gmail = new Gmail(await Gmail.Authenticate(appName: opName));

            Console.WriteLine($"Searching for messages matching search criteria: {searchString}");
            long count = 0;
            await gmail.Search(searchString, async (message, totalCount) =>
            {
                if (await op(gmail, message) == false)
                {
                    Console.WriteLine($"Exiting after processing {++count} / {totalCount} messages.");
                    return false;
                }
                Console.Write($"{++count} / {totalCount} done.\r");
                return true;
            });

            Console.WriteLine("All done.");
        }

        static Task MarkAsRead(string searchString)
        {
            return RunOp("mark-as-read", searchString, async (gmail, msg) => {
                await gmail.MarkAsRead(msg);
                return true;
            });
        }

        static string GetHeader(IList<MessagePartHeader> headers, string name)
        {
            return headers.Where(h => h.Name == name).Select(h => h.Value).FirstOrDefault();
        }

        static Task MarkAsNotSpam(string searchString)
        {
            return RunOp("mark-as-not-spam", searchString, async (gmail, msg) => {
                await gmail.MarkAsNotSpam(msg);
                return true;
            });
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
