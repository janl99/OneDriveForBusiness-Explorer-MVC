#OneDrive for Business starter project for ASP.NET MVC

**Table of Contents**
- [Overview](#overview)
- [Prerequisites and Configuration](#prerequisites)
- [Quick Look at the SSO Authentication Code](#SSOAuth)
- [Quick Look at the the OneDriveExplorerController.cs Code](#cscode)
- [Quick Look at the the JavaScript Code](#jscode)
- [Project files of interest](#projectfiles)
- [License](https://github.com/robledosm/OneDriveForBusiness-Explorer/blob/master/LICENSE.txt)
- [Questions and Comments](#questions-and-comments)

##Overview
This sample uses the [Office 365 APIs client libraries](http://aka.ms/kbwa5c) to demonstrate basic operations against OneDrive for Business service endpoint in Office 365 from a single-tenant ASP.NET MVC5 application.

Below are the operations that you can perform with this sample:

- Read files and folders
- Browse folders
- Download files
- Upload files

<a name="prerequisites"></a>
##Prerequisites and Configuration
This sample requires the following:

- Visual Studio 2013 Update 3.
- [Microsoft Office 365 API Tools](https://visualstudiogallery.msdn.microsoft.com/a15b85e6-69a7-4fdf-adda-a38066bb5155).
- An [Office 365 developer site](http://aka.ms/ro9c62) or another Office 365 tenant.
- Microsoft IIS enabled on your computer.

<a name="register"></a>
###Register app and configure the sample to consume Office 365 API

You can do this via the Office 365 API Tools for Visual Studio (which automates the registration process). Be sure to download and install the Office 365 API tools from the Visual Studio Gallery before you proceed any further.

  1. Build the project. This will restore the NuGet packages for this solution. 
  2. In the Solution Explorer window, choose **OneDriveForBusiness.MVC** project -> **Add** -> **Connected Service**.
  3. A Services Manager window will appear. Choose **Office 365** -> **Office 365 APIs** and select the **Register your app** link.
  4. If you haven't signed in before, a sign-in dialog box will appear. Enter the user name and password for your Office 365 tenant admin. We recommend that you use your Office 365 Developer Site. Often, this user name will follow the pattern {username}@{tenant}.onmicrosoft.com. If you do not have a developer site, you can get a free Developer Site as part of your MSDN Benefits or sign up for a free trial. Be aware that the user must be a Tenant Admin user—but for tenants created as part of an Office 365 Developer Site, this is likely to be the case already. Also developer accounts are usually limited to one user.
  5. After you're signed in, you will see a list of all the services. Initially, no permissions will be selected, as the app is not registered to consume any services yet. 
  6. To register for the services used in this sample, choose the following permissions, and select the Permissions link to set the following permissions: 
   - (Files) - Read and write to your files (Write)
   - (Users and Groups) – Sign you in and read your profile (Read)
  7. Choose the **App Properties** link in the Services Manager window. Make this app available to a Single Organization. 
  8. After selecting **OK** in the Services Manager window, assemblies for connecting to Office 365 REST APIs will be added to your project and the following entries will be added to your appSettings in the web.config: ClientId, ClientSecret, AADInstance, and TenantId. You can use your tenant name for the value of the TenantId setting instead of using the tenant identifier.
  9. Build the solution. Nuget packages will be added to you project. Now you are ready to run the solution and sign in with your organizational account to Office 365.

<a name="SSOAuth"></a>
##Quick Look at the SSO Authentication Code

The authentication startup class, **App_Start/Startup.Auth.cs** in the project contains the startup logic for Azure AD authentication. These classes demonstrate how a persistent ADAL token cache can be constructed and used to store tokens. You can see the token cache implementation in the following files:
  - [```ADAL/ADALTokenCache.cs```](OneDriveForBusiness.MVC/ADAL/ADALTokenCache.cs)
  - [```ADAL/ADALDbContext.cs```](OneDriveForBusiness.MVC/ADAL/ADALDbContext.cs)
  - [```ADAL/Model/UserTokenCache.cs```](OneDriveForBusiness.MVC/ADAL/Model/UserTokenCache.cs) 

> ###Requiring authentication to access controllers
> Applying Authorize attribute to all controllers in your project will require the user to be authenticated before accessing these controllers. To allow the controller to be accessed anonymously, remove this attribute from the controller. If you want to set the permissions at a more granular level, apply the attribute to each method that requires authorization instead of applying it to the controller class.

<a name="cscode"></a>
##Quick Look at the the OneDriveExplorerController.cs Code

###Getting the SharePoint API Client
Get the *SharePointClient* object. You can call this code from other methods that use the SharePoint client.
```csharp
private static async Task<SharePointClient> GetSharePointClient()
{
    string signInUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
    string userObjectId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier".Value;
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
```

###Browsing folders and files
Returns a list of *OneDriveItemViewModel*, if *folderId* is a null or an empty string, the list will be populated with objects from the root folder.
```csharp
List<OneDriveItemViewModel> oneDriveItems = new List<OneDriveItemViewModel>();
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
```

###Downloading files
Download a file from OneDrive for Business as a Stream.

```csharp
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
```

###Uploading files
Uploads a stream as a file to OneDrive for Business deleting the file if it alrady exists.

```csharp
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
```

<a name="jscode"></a>
##Quick Look at the the JavaScript Code

###Browse folders and files
Gets the list of files and folders from the List method of the OneDriveExplorer Controller.
```javascript
function loadFolder(folderId) {
  var dfd = $.Deferred();
  $('.loading-container').removeClass('loading-inactive');
  tbody.empty();
  $.ajax({
      url: '@Url.Action("List", "OneDriveExplorer")',
      data: { folderId: folderId },
      cache: false,
      method: 'POST',
      dataType: 'json',
      success: function (data) {
          if (data.success === true) {
              var rows = '';
              $.each(data.items, function (idx, item) {
                  rows += '<tr data-element-id="' + item.Id + '" data-element-name="' + item.Name + '" data-element-size="' + item.Size + '">' +
                                     '<td>' + createMenu(item.Type, item.Extension) + '</td>' +
                                     '<td>' + item.Name + '</td>' +
                                     '<td>' + item.Creator + '</td>' +
                                     '<td align="right">' + (item.Type === "Folder" ? '' : getReadableFileSizeString(item.Size)) + '</td>' +
                                 '</tr>';
              });
              tbody.append(rows);
              tbody.find('tr').dblclick(function (e) {
                  var selectedItem = getSelectedItem();
                  switch (selectedItem.type) {
                      case "Folder": //open folder
                          openFolder();
                          break;
                      case "File": //download file
                          downloadFile();
                      default:
                          break;
                  }
              });
              tbody.find('tr').click(function (e) {
                  var id = $(this).attr('data-element-id');
                  tbody.find('tr.active').removeClass('active').removeClass('selectedRow');
                  tbody.find('tr[data-element-id="' + id + '"]').addClass('active').addClass('selectedRow');
              });
              dfd.resolve();
          }
          else {
              if (data.message) {
                  alert(data.message);
              }
              dfd.reject(data);
          }

      },
      complete: function () {
          $('.loading-container').addClass('loading-inactive');
      }
  });
  return dfd.promise();
}
```
###Downloading files
Asks the user for confirmation before downloading the file from OneDrive for Business.
```javascript
function downloadFile() {
  var selectedItem = getSelectedItem();
  if (selectedItem.type === "File") {
      var confirm = bootbox.dialog({
          title: 'Confirm',
          message: 'Are you sure you want to download ' + selectedItem.name + '?',
          backdrop: true,
          show: false,
          buttons: {
              No: {
                  callback: function () {
                      return true;
                  }
              },
              Yes: {
                  className: 'btn-success',
                  callback: function () {
                      var url = '@Url.Action("Download", "OneDriveExplorer")?fileId=' + selectedItem.id;
                      window.open(url, "_blank");
                      return true;
                  }
              }
          }
      });
      confirm.find('.modal-dialog').addClass('modal-md');
      confirm.modal('show');
  }
}
```

###Uploading files
Asks the user for confirmation before uploading a file to OneDrive for Business using XMLHttpRequest().
```javascript
uploadButton.click(function (e) {
  var fileSelect = document.getElementById("file-select");
  var files = fileSelect.files;
  if (files.length > 0) {
      var confirm = bootbox.dialog({
          title: 'Confirm',
          message: 'Are you sure you want to upload the selected file?',
          backdrop: true,
          show: false,
          buttons: {
              No: {
                  callback: function () {
                      return true;
                  }
              },
              Yes: {
                  className: 'btn-success',
                  callback: function () {
                      $('.loading-container').removeClass('loading-inactive');
                      var file = files[0];
                      var xhr = new XMLHttpRequest();

                      // File uploaded
                      xhr.addEventListener("load", function () {
                          //upload complete
                          $('.loading-container').addClass('loading-inactive');
                      }, false);


                      xhr.open("POST", "@Url.Action("Upload", "OneDriveExplorer")", true);

                      // Set headers
                      xhr.setRequestHeader("Content-Type", "multipart/form-data");
                      xhr.setRequestHeader("X-File-Name", file.name);

                      // Send the file
                      xhr.send(file);
                  }
              }
          }
      });
      confirm.find('.modal-dialog').addClass('modal-md');
      confirm.modal('show');
  }
  else
  {
      bootbox.alert({
          title: 'Missing file',
          message: 'Please, select a file to upload'
      });
  }
});
```

<a name="projectfiles"></a>
##Project files of interest
**Controllers**

 - [```OneDriveExplorerController.cs```](OneDriveForBusiness.MVC/Controllers/OneDriveExplorerController.cs)

**Models**
 
 - [```OneDriveItemViewModel.cs```](OneDriveForBusiness.MVC/Models/OneDriveItemViewModel.cs)

**Views**
 
  - [```OneDriveExplorer/Index.cshtml```](OneDriveForBusiness.MVC/Views/OneDriveExplorer/Index.cshtml)

**Utils Folder**
 
  - [```SettingsHelper.cs```](OneDriveForBusiness.MVC/Utils/SettingsHelper.cs)


> This sample uses [Bootbox.js](http://bootboxjs.com/) to render dialog boxes and [XMLHttpRequest](http://www.w3.org/TR/XMLHttpRequest/) to upload files

## Questions and Comments

We'd love to get your feedback on the [OneDriveForBusiness-Explorer-MVC](https://github.com/robledosm/OneDriveForBusiness-Explorer) project. You can send your questions and suggestions to us in the [Issues](https://github.com/robledosm/OneDriveForBusiness-Explorer/issues) section of this repository.


## Copyright

Copyright (c) Lumen21, INC. All rights reserved.

