## SharePoint Framework custom header and footer application customizer extension

### Building the code (for localhost debugging)

```bash
git clone the repo
npm i
gulp serve --nobrowser
```

When debugging locally, you will need to include the following querystring parameter on any modern pages where you wish to test this functionality:

```bash
?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/manifests.js&customActions={"bbe5f3fa-7326-455d-8573-9f0b2b015ff9":{"location":"ClientSideExtension.ApplicationCustomizer"}}
```

### Building the code (for production deployment)

```bash
git clone the repo
npm i
gulp bundle --ship
gulp package-solution --ship
```
### Adding the user custom action for the extension in a tenant-scoped deployment

If you check the box labeled **Make this solution available to all sites in the organization** before pressing **Deploy**, you will need to manually add the user custom action associated with the extension on any site where you would like the custom header and footer to be rendered on modern pages. If you deploy the extension at tenant scope, it is immediately available to all sites and you do not need to explicitly add the app from the Site Contents screen. However, because tenant-scoped extensions cannot leverage the feature framework, you will need to associate the user custom action with the **ClientSideComponentId** of the extension manually. This can be accomplished a number of different ways. Some example code using the .NET Managed Client Object Model in a console application is shown below:

```cs
using (ClientContext ctx = new ClientContext("https://[YOUR TENANT].sharepoint.com"))
{
    SecureString password = new SecureString();
    foreach (char c in "[YOUR PASSWORD]".ToCharArray()) password.AppendChar(c);
    ctx.Credentials = new SharePointOnlineCredentials("[USER]@[YOUR TENANT].onmicrosoft.com", password);

    Web web = ctx.Web;
    UserCustomActionCollection ucaCollection = web.UserCustomActions;
    UserCustomAction uca = ucaCollection.Add();
    uca.Title = "SPFxHeaderFooterApplicationCustomizer";
    // This is the user custom action location for application customizer extensions
    uca.Location = "ClientSideExtension.ApplicationCustomizer";
    // Use the ID from HeaderFooterApplicationCustomizer.manifest.json below
    uca.ClientSideComponentId = new Guid("bbe5f3fa-7326-455d-8573-9f0b2b015ff9");
    uca.Update();

    ctx.Load(web, w => w.UserCustomActions);
    ctx.ExecuteQuery();

    Console.WriteLine("User custom action added to site successfully!");
}
```

You still need to install and configure the [SharePoint-hosted add-in]
([https://github.com/iSharePoint](https://github.com/iSharePoint/PublicRepository/SPFxHeaderFooter) 
on any site where you wish to use this extension and 
[disable NoScript on that site]

### If you do not perform a tenant-scoped installation

If you decline to check the box to allow tenant-scoped installation when you upload the **.sppkg** file to the app catalog, the extension will be made available to manually add to any site via **Site Contents > Add an app**. This will automatically associate the user custom action on any site where you manually add the extension, so no code is necessary to register the user custom action as shown above in this case.
