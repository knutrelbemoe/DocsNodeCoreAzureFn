using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Identity.Client;
using Microsoft.SharePoint.Client;

namespace FunctionApp_UpdateAuthor
{
    public static class CheckFolderPermission
    {
        public class ReqData
        {
            public string siteUrl { get; set; }
            public string folderPath { get; set; }
            public string emailID { get; set; }
            public string tenantName { get; set; }
        }

        [FunctionName("CheckFolderPermission")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");

            string siteUrl = "";
            string folderPath = "";
            string emailID = "";
            string tenantName = "";
            permissionJSON folderPermission = new permissionJSON();


            ReqData data = await req.Content.ReadAsAsync<ReqData>();

            if (data == null)
            {
                return req.CreateResponse(HttpStatusCode.BadRequest,
                    "Please pass all parameter in the request body");
            }
            siteUrl = data.siteUrl;
            folderPath = data.folderPath;
            emailID = data.emailID;
            tenantName = data.tenantName;

            if (string.IsNullOrEmpty(siteUrl) ||
             string.IsNullOrEmpty(tenantName) ||
             string.IsNullOrEmpty(emailID))
            {
                return req.CreateResponse(HttpStatusCode.BadRequest,
                   "Please pass all parameter in the request body");
            }

            var certName = "AzureSPOAccessPvtKeyCert.pfx";
            var certPassword = "pass@word1";
            var home = Environment.GetEnvironmentVariable("HOME");
            var certPath = home != null ?
                Path.Combine(home, @"site\wwwroot", certName) :
                Path.Combine(@"D:\Projects\DocsNode\Azure\FunctionApp_UpdateAuthor\FunctionApp_UpdateAuthor\" + certName);

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
                using (var context = new ClientContext(siteUrl+"/sites/docsnodeadmin"))
                {
                    context.ExecutingWebRequest += (s, e) =>
                    {
                        e.WebRequestExecutor.RequestHeaders["Authorization"] =
                            "Bearer " + token;
                    };

                     folderPermission = chkUsrPermission(tenantName, siteUrl, emailID, folderPath, context);
                }
            }
            catch
            {

            }

            return object.Equals(folderPermission,new permissionJSON())
                ? req.CreateResponse(HttpStatusCode.BadRequest, "Permission check failed !!!")
                : req.CreateResponse(HttpStatusCode.OK, folderPermission);
        }

        public class permissionJSON
        {
            public string userEmail { get; set; }
            public bool hasPermission { get; set; }
        }
        public static permissionJSON chkUsrPermission(string tenantName, string siteUrl, string userEmail, string folderPath, ClientContext context)
        {

            permissionJSON objJSON = new permissionJSON();
            try
            {
                var userEffectivePermission = new ClientResult<BasePermissions>();

                if (!string.IsNullOrEmpty(folderPath))
                {
                    var Folder = context.Web.GetFolderByServerRelativeUrl("/sites/DocsNodeAdmin/DocsNodeTemplatesLibrary/" + folderPath);
                    context.Load(Folder);
                    context.ExecuteQuery();

                     userEffectivePermission = Folder.ListItemAllFields.GetUserEffectivePermissions("i:0#.f|membership|" + userEmail);
                    context.ExecuteQuery();

                }
                else
                {
                    var Library = context.Web.Lists.GetByTitle("DocsNodeTemplatesLibrary");
                    context.Load(Library);
                    context.ExecuteQuery();

                     userEffectivePermission = Library.GetUserEffectivePermissions("i:0#.f|membership|" + userEmail);
                    context.ExecuteQuery();
                }

               

                //var val = userEffectivePermission.Value;

                objJSON.userEmail = userEmail;

                if (userEffectivePermission.Value.Has(PermissionKind.ViewListItems))
                {
                    objJSON.hasPermission = true;
                }

            }
            catch (Exception ex)
            {

            }
            return objJSON;
        }
    }
}
