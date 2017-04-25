using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Drawing;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Accord.Vision.Detection;
using Accord.Vision.Detection.Cascades;

namespace GmailQuickstart
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/gmail-dotnet-quickstart.json
        static string[] Scopes = { GmailService.Scope.GmailReadonly };
        static string ApplicationName = "Gmail API .NET Quickstart";

        static void Main(string[] args)
        {
            String outputDir = "H:\\Temp\\";

            UserCredential credential;
            FaceHaarCascade cascade = new FaceHaarCascade();
            HaarObjectDetector detector = new HaarObjectDetector(cascade);

            detector.SearchMode = ObjectDetectorSearchMode.Average;
            detector.ScalingFactor = 1.5F;
            detector.ScalingMode = ObjectDetectorScalingMode.GreaterToSmaller; ;
            detector.UseParallelProcessing = true;
            detector.Suppression = 3;

            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/gmail-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Gmail API service.
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define parameters of request.
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List("me");
            request.Q = "has:attachment";

            // List labels.
            IList<Message> messages = request.Execute().Messages;
            Console.WriteLine("Messages:");
            if (messages != null && messages.Count > 0)
            {
                foreach (var messageItem in messages)
                {
                    //Console.WriteLine("{0}", messageItem.Id);
                    Message message = service.Users.Messages.Get("me", messageItem.Id).Execute();
                    Console.WriteLine(message.Payload.MimeType.ToString());
                    IList<MessagePart> parts = message.Payload.Parts;
                    foreach (MessagePart part in parts)
                    {
                        if (!String.IsNullOrEmpty(part.Filename))
                        {
                            String attId = part.Body.AttachmentId;
                            MessagePartBody attachPart = service.Users.Messages.Attachments.Get("me", messageItem.Id, attId).Execute();
                            Console.WriteLine(part.Filename);

                            // Converting from RFC 4648 base64 to base64url encoding
                            // see http://en.wikipedia.org/wiki/Base64#Implementations_and_history
                            String attachData = attachPart.Data.Replace('-', '+');
                            attachData = attachData.Replace('_', '/');

                            byte[] data = Convert.FromBase64String(attachData);
                            
                            MemoryStream ms = new MemoryStream(data);
                            Bitmap img = new Bitmap(Image.FromStream(ms));
                            Rectangle[] rects = detector.ProcessFrame(img);
                            if(rects.Count() > 0)
                            {
                                Console.WriteLine("Face detected!!!!");
                                File.WriteAllBytes(Path.Combine(outputDir, part.Filename), data);
                            }
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("No messages found.");
            }
            Console.Read();

        }
    }
}