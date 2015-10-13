# OneDrive for Business Explorer #

##OneDrive for Business starter project for ASP.NET MVC##

**Table of Contents**
- [Overview](#overview)
- [Prerequisites and Configuration](#prerequisites)
- [Quick Look at the SSO Authentication Code](#SSOAuth)
- [Project files of interest](#projectfiles)
- [License](https://github.com/robledosm/OneDriveForBusiness-Explorer/blob/master/LICENSE.txt)
- [Questions and Comments](#questions-and-comments)

##Overview##
This sample uses the [Office 365 APIs client libraries](http://aka.ms/kbwa5c) to demonstrate basic operations against OneDrive for Business service endpoint in Office 365 from a single-tenant ASP.NET MVC5 application.

Below are the operations that you can perform with this sample:

- Read files and folders
- Browse folders
- Download files
- Upload files

<a name="prerequisites"></a>
##Prerequisites and Configuration##
This sample requires the following:
- Visual Studio 2013 Update 3.
- [Microsoft Office 365 API Tools](https://visualstudiogallery.msdn.microsoft.com/a15b85e6-69a7-4fdf-adda-a38066bb5155).
- An [Office 365 developer site](http://aka.ms/ro9c62) or another Office 365 tenant.
- Microsoft IIS enabled on your computer.

<a name="register"></a>
###Register app and configure the sample to consume Office 365 APIs###

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
##Quick Look at the SSO Authentication Code##

The authentication startup class, **App_Start/Startup.Auth.cs** in the project contains the startup logic for Azure AD authentication. These classes demonstrate how a persistent ADAL token cache can be constructed and used to store tokens. You can see the token cache implementation in the following files:
  - [```ADAL/ADALTokenCache.cs```](OneDriveForBusiness.MVC/ADAL/ADALTokenCache.cs)
  - [```ADAL/ADALDbContext.cs```](OneDriveForBusiness.MVC/ADAL/ADALDbContext.cs)
  - [```ADAL/Model/UserTokenCache.cs```](OneDriveForBusiness.MVC/ADAL/Model/UserTokenCache.cs) 

###Requiring authentication to access controllers###

Applying Authorize attribute to all controllers in your project will require the user to be authenticated before accessing these controllers. To allow the controller to be accessed anonymously, remove this attribute from the controller. If you want to set the permissions at a more granular level, apply the attribute to each method that requires authorization instead of applying it to the controller class.

<a name="projectfiles"></a>
##Project files of interest##

**Controllers**

  - [```OneDriveExplorerController.cs```](OneDriveForBusiness.MVC/Controllers/OneDriveExplorerController.cs)

**Models**
  
  - [```OneDriveItemViewModel.cs```](OneDriveForBusiness.MVC/Models/OneDriveItemViewModel.cs)

**Views**
  
  - [```OneDriveExplorer/Index.cshtml```](OneDriveForBusiness.MVC/Views/OneDriveExplorer/Index.cshtml)

**Utils Folder**
  - [```SettingsHelper.cs```](OneDriveForBusiness.MVC/Utils/SettingsHelper.cs)

##3rd party tools##

This sample uses [Bootbox.js](http://bootboxjs.com/) to render dialog boxes and [XMLHttpRequest](http://www.w3.org/TR/XMLHttpRequest/) to upload files

## Questions and Comments

We'd love to get your feedback on the OneDriveForBusiness-Explorer project. You can send your questions and suggestions to us in the [Issues](https://github.com/robledosm/OneDriveForBusiness-Explorer/issues) section of this repository.


## Copyright ##

Copyright (c) Lumen21, INC. All rights reserved.
