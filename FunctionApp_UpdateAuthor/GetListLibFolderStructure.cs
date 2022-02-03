using System;
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

namespace FunctionApp_UpdateAuthor
{
    public static class GetListLibFolderStructure
    {

        public class ReqData
        {
            public string siteUrl { get; set; }
            public string folderPath { get; set; }
            public string emailID { get; set; }
            public string tenantName { get; set; }
            public string resource { get; set; }
            public string type { get; set; }
        }

        [FunctionName("GetListLibFolderStructure")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");

            string siteUrl = "";
            string folderPath = "";
            // string emailID = "";
            string tenantName = "";
            string resource = "";
            string type = "";

            ReqData data = await req.Content.ReadAsAsync<ReqData>();

            if (data == null)
            {
                return req.CreateResponse(HttpStatusCode.BadRequest,
                    "Please pass all parameter in the request body");
            }

            siteUrl = data.siteUrl;
            type = data.type;
            folderPath = data.folderPath;
            resource = data.resource;
            tenantName = data.tenantName;
            bool chkFolder = false;
            if (folderPath != null)
            {
                chkFolder = true;
            }

            if (string.IsNullOrEmpty(siteUrl) ||
              string.IsNullOrEmpty(type) ||
              string.IsNullOrEmpty(tenantName) ||
              string.IsNullOrEmpty(resource))
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
                string lstItems = string.Empty;

                using (var context = new ClientContext(siteUrl + "/sites/docsnodeadmin"))
                {
                    context.ExecutingWebRequest += (s, e) =>
                    {
                        e.WebRequestExecutor.RequestHeaders["Authorization"] =
                            "Bearer " + token;
                    };
                    List olist = context.Web.Lists.GetByTitle("DocsNodeConfiguration");

                    CamlQuery camlQuery = new CamlQuery();
                    camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='ConfigAssestTitle'/>" +
                                         "<Value Type='Text'>" + type + "</Value>" +
                                         "</Eq></Where></Query></View>";

                    ListItemCollection collistItem = olist.GetItems(camlQuery);
                    context.Load(collistItem);
                    context.ExecuteQuery();

                    if (collistItem.Count > 0)
                    {
                        foreach (ListItem oListItem in collistItem)
                        {
                            //  Console.WriteLine("ID: {0} \nTitle: {1} \nBody: {2}", oListItem.Id, oListItem["Title"], oListItem["Body"]);
                            var configlistGuid = oListItem["ConfigSourceListGUID"];
                            var configSiteUrl = oListItem["ConfigSourceListPath"];
                            var configListName = oListItem["ConfigSourceList"];
                            var configListRelUrl = configSiteUrl + "/" + configListName + "/" + folderPath;
                            var configAssestTitle = oListItem["ConfigAssestTitle"];



                            using (var ctx = new ClientContext(siteUrl + configSiteUrl))
                            {
                                context.ExecutingWebRequest += (s, e) =>
                                {
                                    e.WebRequestExecutor.RequestHeaders["Authorization"] =
                                        "Bearer " + token;
                                };

                                if (chkFolder)
                                {
                                    ///sites/DocsNodeAdmin/Lists/DocsNodeText/Microsoft/Word/Client
                                 //   string relFolderPath = configSiteUrl + "/Lists" + "/" + configListName + "/" + folderPath;
                                    olist = context.Web.Lists.GetByTitle(configListName.ToString());
                                    camlQuery = new CamlQuery();
                                    camlQuery.ViewXml = "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FileDirRef'/>" +
                                                        "<Value Type='Text'>" + folderPath + "</Value>" +
                                                        "</Eq></Where></Query></View>";
                                    
                                    collistItem = olist.GetItems(camlQuery);
                                    context.Load(collistItem);
                                    context.ExecuteQuery();
                                    lstItems = string.Empty;

                                    foreach (ListItem itm in collistItem)
                                    {
                                        lstItems += string.Format("{0}{1}{2}{3}{4}", "[", itm["Title"], ",", itm["FSObjType"], "]");
                                    }
                                }
                                else
                                {
                                    olist = context.Web.Lists.GetByTitle(configListName.ToString());
                                    camlQuery = new CamlQuery();
                                    collistItem = olist.GetItems(camlQuery);
                                    context.Load(collistItem);
                                    context.ExecuteQuery();

                                    foreach (ListItem itm in collistItem)
                                    {
                                        lstItems += string.Format("{0}{1}{2}{3}{4}", "[", itm["Title"], ",", itm["FSObjType"], "]");
                                    }

                                }
                            }
                        }

                    }
                }


                var response = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent(lstItems, Encoding.UTF8,
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
            //return name == null
            //? req.CreateResponse(HttpStatusCode.BadRequest, "Please pass a name on the query string or in the request body")
            //: req.CreateResponse(HttpStatusCode.OK, "Hello " + name);
        }
    }
}
