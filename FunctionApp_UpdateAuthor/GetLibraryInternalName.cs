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
    public static class GetLibraryInternalName
    {
        public class ReqData
        {
            public string siteUrl { get; set; }
            public string tenantName { get; set; }
        }

        [FunctionName("GetLibraryInternalName")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");

            string siteUrl = "";
            string tenantName = "docsnode";

            ReqData data = await req.Content.ReadAsAsync<ReqData>();

            if (data == null)
            {
                return req.CreateResponse(HttpStatusCode.BadRequest,
                    "Please pass all parameter in the request body");
            }

            siteUrl = data.siteUrl;
            tenantName = data.tenantName;

            if (string.IsNullOrEmpty(siteUrl))
            {
                return req.CreateResponse(HttpStatusCode.BadRequest,
                   "Please pass siteUrl parameter in the request body");
            }

            log.Info("url:" + data.siteUrl);
            log.Info("tenantName:" + data.tenantName);

            var certName = "AzureSPOAccessPvtKeyCert.pfx";
            var certPassword = "pass@word1";
            //var certPath = @"C:\Users\spdev\Desktop\"
            //    + certName;
            var home = Environment.GetEnvironmentVariable("HOME");
            var certPath = home != null ?
                Path.Combine(home, @"site\wwwroot", certName) :
                Path.Combine(@"C:\Users\spdev\Desktop\" + certName);

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
                string internalName = "";
                using (var context = new ClientContext(siteUrl))
                {
                    context.ExecutingWebRequest += (s, e) =>
                    {
                        e.WebRequestExecutor.RequestHeaders["Authorization"] =
                            "Bearer " + token;
                    };

                    List list = context.Web.Lists.GetByTitle("Dokumenter");
                    Folder rootFolder = list.RootFolder;
                    context.Load(list);
                    context.Load(rootFolder);
                    context.ExecuteQuery();
                    internalName = rootFolder.Name;
                }

                var response = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent(internalName, Encoding.UTF8,
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
    }
}
