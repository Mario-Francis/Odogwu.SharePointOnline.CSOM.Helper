# Odogwu.SharePointOnline.CSOM.Helper

A .Net Standard helper library for the Microsoft.SharePointOnline.CSOM library.

## Description

The helper library is a specialized software tool designed as a wrapper for the Microsoft.SharePointOnline.CSOM library. Its primary objective is to streamline and simplify the process of establishing connections with SharePoint Online, specifically targeting tasks related to uploading and retrieving documents with their associated properties within document libraries. Additionally, the library offers convenient functionality for performing CRUD (Create, Read, Update, Delete) operations on Lists, enabling efficient management of SharePoint data. By abstracting away the complexities of the underlying SharePoint Online API, this helper library empowers developers to interact with SharePoint Online more seamlessly and focus on implementing their desired document and list management features with enhanced productivity.

### Built With

- C#.Net (Targeting .Net standard 2.0)

### Getting Started

#### Prerequisites

 To be able to use this library, the following is needed

- SharePoint Site Url: This is the url to the site you wish to connect to.
- Client ID: This is the SharePoitn App client ID.
- Client Secret: This is the SharePoint App client secret.
- Tenant ID: This is your sharepoint tenant ID.
- Resource: This is a string usually "00000003-0000-0ff1-ce00-000000000000" by default.
- Grant Type: This is a string usually "client_credentials" by default.

To successfully generate the above information, kindly follow the guide [here](https://emmanueladegor.medium.com/sharepoint-online-rest-api-authentication-in-postman-b66d9ea5f0bc) written by Emmanuel A.

#### Installation

To installt this library, execute the below command

    dotnet add package Odogwu.SharePointOnline.CSOM.Helper

#### Usage

##### Create Authentication Manager

    var  siteUrl  =  "your_site_url";
    var  grantType="client_credentials";
    var  resource="00000003-0000-0ff1-ce00-000000000000";
    var  clientId="your_client_id";
    var  clientSecret="your_client_secret";
    var  tenant="your_tenant_id";

    var authManager = AuthenticationManager(siteUrl, grantType, resource, clientId, clientSecret, tenant);

##### Upload a new file

    var  libraryManager  =  new  LibraryManager(authManager);
    var  uploadRequest  =  new  FileUploadRequest()
    {
    Library  =  "your_library_name",
    UploadItem  =  new  FileUploadItem
    {
    File  =  new  byte[0], // file binary
    FileExtension  =  "file_extension(txt, pdf, jpg)",
    FileName  =  "file_name",
    Properties  =  new  List<KeyValuePair<string, object>>(){}
    }
    };
    FileUploadResponse  result  =  await  libraryManager.UploadFile(uploadRequest);
    int  fileId  =  result.Id;
    string  fileObjectId  =  result.Guid;
    string  absoluteUrl  =  result.AbsoluteUrl;
    string  fileName  =  result.FileName;
    string  serverRelativeUrl  =  result.ServerRelativeUrl;

#### Get file by ID

    var  libraryManager  =  new  LibraryManager(authManager);
    int fileId = <your_file_id>;
    SPFile  file  =  await  libraryManager.GetFileById(fileId, "your_library_name");

### Contributing

Contributions are what make the open source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement". Don't forget to give the project a star! Thanks again!

### License

Distributed under the MIT License. See `LICENSE.txt` for more information.
