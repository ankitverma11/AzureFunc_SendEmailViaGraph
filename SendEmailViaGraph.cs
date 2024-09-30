using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Collections.Generic;
using Azure.Identity;
using Microsoft.Graph.Users.Item.SendMail;
using Newtonsoft.Json;
using System.IO;

namespace SendEmailViaGraph
{
    public static class SendEmailViaGraph
    {
        [FunctionName("SendEmailViaGraph")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("Processing request to send email. ");
            try
            {
                string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
                log.LogInformation($"Received JSON: {requestBody}");
                dynamic data = JsonConvert.DeserializeObject(requestBody);
                string name = data?.name;
                string email = data?.email;
                string phonenumber = data?.phonenumber;
                string text = data?.text;

                // Fetch secrets and config from environment
                var scopes = new[] { "https://graph.microsoft.com/.default" };
                string clientId = Environment.GetEnvironmentVariable("ClientId");
                string clientSecret = Environment.GetEnvironmentVariable("ClientSecret");
                string tenantId = Environment.GetEnvironmentVariable("TenantId");
                var options = new TokenCredentialOptions
                {
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
                };

                var clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret,options);

                var graphClient = new GraphServiceClient(clientSecretCredential,scopes);

                // Create email message
                log.LogInformation("Create email message");
                // Construct the HTML email body
                string htmlBody = $@"
                    <html>
                    <body>
                        <p>Name : {name},</p>
                        <p>Email: {email}</p>
                        <p>Phone Number : {phonenumber}</p>
                        <p>Text : {text}</p>
                    </body>
                    </html>";

                // Create email message
                var body = new SendMailPostRequestBody
                {
                    Message = new Message
                    {
                        Subject = "Test Email from Azure Function",
                        Body = new ItemBody
                        {
                            ContentType = BodyType.Html,
                            Content = htmlBody
                        },
                        ToRecipients = new List<Recipient>
                        {
                            new Recipient
                            {
                                EmailAddress = new EmailAddress
                                {
                                    Address = "abc@gmail.com"
                                }
                            },
                            new Recipient
                            {
                                EmailAddress = new EmailAddress
                                {
                                    Address = "vca...com"
                                }
                            }
                        }
                    }
                };

                // Send email
                await graphClient.Users["....onmicrosoft.com"].SendMail.PostAsync(body);
                return new OkObjectResult("Email sent successfully.");
            }
            catch (Exception ex)
            {
                log.LogInformation("Email sent function failed");
                log.LogInformation("InnerException : " + ex.InnerException);
                return new OkObjectResult("Email sent failed.");
            }
        }
    }
}
