using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.SharePoint.CoreServices;
using OneDriveForBusiness.MVC.Models;
using OneDriveForBusiness.MVC.Utils;
using OneDriveForBusiness.MVC.ADAL;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace OneDriveForBusiness.MVC.Controllers
{
    [Authorize]
    public class OneDriveExplorerController : Controller
    {
        [HttpGet]
        public ActionResult Index()
        {
            if (ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier) == null)
            {
                return RedirectToAction("", "Home");
            }

            return View("Index");
        }

        private static async Task<SharePointClient> GetSharePointClient()
        {
            string signInUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
            string userObjectId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
            string tenantId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;

            AuthenticationContext authContext = new AuthenticationContext(string.Format("{0}/{1}", SettingsHelper.AuthorizationUri, tenantId), new ADALTokenCache(signInUserId));

            DiscoveryClient discovery = new DiscoveryClient(SettingsHelper.DiscoveryServiceEndpointUri,
                    async () =>
                    {
                        var authResult = await authContext.AcquireTokenSilentAsync(SettingsHelper.DiscoveryServiceResourceId, new ClientCredential(SettingsHelper.ClientId, SettingsHelper.ClientSecret), new UserIdentifier(userObjectId, UserIdentifierType.UniqueId));

                        return authResult.AccessToken;
                    });

            CapabilityDiscoveryResult capability = await discovery.DiscoverCapabilityAsync(SettingsHelper.Capability);
            SharePointClient client = new SharePointClient(capability.ServiceEndpointUri,
                async () =>
                {
                    var authResult = await authContext.AcquireTokenSilentAsync(capability.ServiceResourceId, new ClientCredential(SettingsHelper.ClientId, SettingsHelper.ClientSecret), new UserIdentifier(userObjectId, UserIdentifierType.UniqueId));

                    return authResult.AccessToken;
                });
            return client;
        }

        private static string GetContentType(string fileName)
        {
            var reg = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(System.IO.Path.GetExtension(fileName).ToLower());
            string contentType = "application/octet-stream"; //RFC2046 says "The "octet-stream" subtype is used to indicate that a body contains arbitrary binary data

            if (reg != null)
            {
                string registryContentType = reg.GetValue("Content Type") as string;

                if (!String.IsNullOrWhiteSpace(registryContentType))
                {
                    contentType = reg.GetValue("Content Type") as string;
                }
            }
            return contentType;
        }

        [HttpPost]
        public async Task<JsonResult> List(string folderId)
        {
            if (ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier) == null)
            {
                return Json(new { success = false, message = "Authentication Required" });
            }

            List<OneDriveItemViewModel> oneDriveItems = new List<OneDriveItemViewModel>();
            try
            {
                SharePointClient client = await GetSharePointClient();

                if (string.IsNullOrEmpty(folderId))
                {

                    var filesResult = await client.Files.ExecuteAsync();
                    do
                    {
                        var files = filesResult.CurrentPage;
                        foreach (var file in files)
                        {
                            string extension = (file.Type == "File") ? System.IO.Path.GetExtension(file.Name).Split('.')[1] : string.Empty;
                            oneDriveItems.Add(new OneDriveItemViewModel { Id = file.Id, Name = file.Name, Size = file.Size, Type = file.Type, Extension = extension, Creator = file.CreatedBy.User.DisplayName });
                        }
                        filesResult = filesResult.GetNextPageAsync().GetAwaiter().GetResult();
                    } while (filesResult != null);
                }
                else
                {
                    var filesResult = await client.Files.GetById(folderId).ToFolder().Children.ExecuteAsync();
                    do
                    {
                        var files = filesResult.CurrentPage;
                        foreach (var file in files)
                        {
                            string extension = (file.Type == "File") ? System.IO.Path.GetExtension(file.Name).Split('.')[1] : string.Empty;
                            oneDriveItems.Add(new OneDriveItemViewModel { Id = file.Id, Name = file.Name, Size = file.Size, Type = file.Type, Extension = extension, Creator = file.CreatedBy.User.DisplayName });
                        }
                        filesResult = filesResult.GetNextPageAsync().GetAwaiter().GetResult();
                    } while (filesResult != null);

                }
            }
            catch (AdalException exception)
            {
                //handle token acquisition failure
                return Json(new { success = false, message = "Authorization Required" });
            }
            catch
            {
                return Json(new { success = false, message = "Unexpected exception" });
            }
            return Json(new { success = true, items = oneDriveItems.OrderBy(x => x.Type).OrderBy(x => x.Name) });
        }
        
        [HttpPost]
        public async Task Upload()
        {
            if (ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier) == null)
            {
                //not authenticated
                RedirectToAction("", "Home");
            }

            var length = Request.ContentLength;
            if (length > 0)
            {
                SharePointClient client = await GetSharePointClient();
                var fileName = Request.Headers["X-File-Name"];
                try
                {
                    // First check if a file with the same name already exists. If it exists, delete it.
                    var item = await client.Files.GetByPathAsync(fileName);
                    await item.DeleteAsync();
                }
                catch
                {
                    //there is no file with the same name, swallow the error and continue
                }

                var oneDriveFile = new Microsoft.Office365.SharePoint.FileServices.File
                {
                    Name = fileName
                };
                await client.Files.AddItemAsync(oneDriveFile);
                await client.Files.GetById(oneDriveFile.Id).ToFile().UploadAsync(Request.InputStream);
            }
            else
            {
                throw new Exception("Empty piece");
            }
        }

        [HttpGet]
        public async Task Download(string fileId)
        {
            if (ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier) == null)
            {
                RedirectToAction("", "Home");
            }
            SharePointClient client = await GetSharePointClient();

            var file = client.Files.GetById(fileId).ToFile();
            using (Stream stream = await file.DownloadAsync())
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    stream.Seek(0, SeekOrigin.Begin);
                    stream.CopyTo(ms);
                    byte[] buffer = ms.ToArray();

                    var fileData = await file.ExecuteAsync();
                    Response.Clear();
                    Response.ContentType = GetContentType(fileData.Name);
                    Response.AddHeader("Content-Disposition", "attachment; filename=" + HttpUtility.UrlEncode(fileData.Name, System.Text.Encoding.UTF8));
                    Response.AddHeader("Content-Length", buffer.Length.ToString());
                    Response.OutputStream.Write(buffer, 0, buffer.Length);
                    Response.Flush();
                }
            }
        }

        
    }
}