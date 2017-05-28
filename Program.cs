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
            var gmail = new Gmail(await Gmail.Authenticate());
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
