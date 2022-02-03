using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Identity.Client;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using Microsoft.SharePoint.Client.Utilities;
using Newtonsoft.Json;

namespace FunctionApp_UpdateAuthor
{
    public static class UpdateCreator
    {
        public class ReqData
        {
            public string siteUrl { get; set; }
            public string listName { get; set; }
            public int itemID { get; set; }
            public string emailID { get; set; }
            public string tenantName { get; set; }
        }

        [FunctionName("UpdateCreator")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");

            string siteUrl = "";
            string listName = "";
            string emailID = "";
            Int32 itemID = 0;
            Int32 userID = 0;
            string tenantName = "";

            ReqData data = await req.Content.ReadAsAsync<ReqData>();

            if (data == null)
            {
                return req.CreateResponse(HttpStatusCode.BadRequest,
                    "Please pass all parameter in the request body");
            }

            siteUrl = data.siteUrl;
            listName = data.listName;
            itemID = data.itemID;
            emailID = data.emailID;
            tenantName = data.tenantName;

            if (string.IsNullOrEmpty(siteUrl) ||
               string.IsNullOrEmpty(listName) ||
               string.IsNullOrEmpty(tenantName) ||
               itemID <= 0 || string.IsNullOrEmpty(emailID))
            {
                return req.CreateResponse(HttpStatusCode.BadRequest,
                   "Please pass all parameter in the request body");
            }
            var certName = "AzureSPOAccessPvtKeyCert.pfx";
            var certPassword = "pass@word1";
            var home = Environment.GetEnvironmentVariable("HOME");
            var certPath = home != null ?
                Path.Combine(home, @"site\wwwroot", certName) :
                Path.Combine(@"E:\Projects\DocsNodeFunction\AzureFnDocsNodeCSOMAPI\" + certName);

            log.Info("Cert Path:" + certPath);

            var cert = new X509Certificate2(
                System.IO.File.ReadAllBytes(certPath),
                certPassword,
                X509KeyStorageFlags.Exportable |
                X509KeyStorageFlags.MachineKeySet |
                X509KeyStorageFlags.PersistKeySet);

            var clientId = "cee79eed-97a7-4373-846b-c54892c0eb89";
            var authority =
                "https://login.microsoftonline.com/" +
                $"{tenantName}.onmicrosoft.com/";
            var azureApp =
                ConfidentialClientApplicationBuilder.Create(clientId)
                    .WithAuthority(authority)
                    .WithCertificate(cert)
                    .Build();

            var scopes = new string[] {
                $"https://{tenantName}.sharepoint.com/.default" };
            var authResult = await
                azureApp.AcquireTokenForClient(scopes).ExecuteAsync();
            var token = authResult.AccessToken;

            try
            {
                using (var context = new ClientContext(siteUrl))
                {
                    context.ExecutingWebRequest += (s, e) =>
                    {
                        e.WebRequestExecutor.RequestHeaders["Authorization"] =
                            "Bearer " + token;
                    };

                    userID = GetUserID(emailID, context);

                    List list = context.Web.Lists.GetByTitle(listName);
                    ListItem spItem = list.GetItemById(itemID);
                    spItem["Author"] = userID;
                    spItem["Editor"] = userID;
                    spItem.Update();

                    context.ExecuteQuery();
                }

                var response = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent("Success", Encoding.UTF8,
                        "application/json")
                };

                return response;
            }
            catch (Exception ex)
            {
                var response = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent("Fail", Encoding.UTF8,
                       "application/json")
                };

                return response;
            }
        }
        // Test for Git
        /// <summary>
        /// Returns user ID by processing email ID
        /// </summary>
        /// <param name="emailID"></param>
        /// <returns></returns>
        private static int GetUserID(string emailID, ClientContext context)
        {
            Int32 userID = 0;

            try
            {
                User user = context.Web.EnsureUser(emailID);
                context.Load(user);
                context.ExecuteQuery();

                userID = user.Id;
            }
            catch (Exception ex)
            {
            }
            return userID;
        }
    }
}
