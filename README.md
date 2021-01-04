[logo]: https://raw.githubusercontent.com/Geeksltd/Zebble.Office365/master/icon.png "Zebble.Office365"


## Zebble.Office365

![logo]

A Zebble plugin that allow you to sign in with Microsoft Office 365 API and access user information, sending an email and so on.


[![NuGet](https://img.shields.io/nuget/v/Zebble.Office365.svg?label=NuGet)](https://www.nuget.org/packages/Zebble.Office365/)

> With this plugin you can easily connect to Microsoft Office 365 API and send some request like reading user information, sending an email in Zebble applications.

<br>


### Setup
* Available on NuGet: [https://www.nuget.org/packages/Zebble.Office365/](https://www.nuget.org/packages/Zebble.Office365/)
* Install in your platform client projects.
* Available for iOS, Android and UWP.
<br>

#### Register your application

You can register the application by visiting https://apps.dev.microsoft.com and clicking on the Add an app button. Once you enter the app details, be sure to note the Application Id generated. Under the Add Platforms section, you can now register an application to multiple platforms. This makes life simpler by having one application ID for various implementations of the same app across mobile and web.
Click **Native Application**. This will allow mobile applications to access the organization’s AD and Graph API.

Once supported platforms are added, we can add permissions on the same screen. In this case, I’ve given the permissions `User.Read`. These are required to get a user’s information.
Save the changes, and your app is now registered.

### Api Usage

To connect to Microsoft Office 365 API you can use `Zebble.Office365.SignIn()` but you need to configure client indentification and scopes and set some setting for each platform before you connect to the API.

```csharp
await Office365.Initialize("Your Client ID", new[] { "User.Read", ... }, "msal[Your Client ID]://auth");
await Office365.SignIn();
```

#### GetRequest

To access user information or reading the contact list or calendar and so on, you can use this method like below:

```csharp
//get user profile picture
var stream = await Office365.GetRequest<Stream>("https://graph.microsoft.com/v1.0/me/photo/$value");
```

#### PostRequest

You can use some other API methods which are need post type request like sending an Email like below:

```csharp
var jsonString = await Office365.PostRequest<string>("url", "body as json string");
```

To see other Microsoft API you can see https://developer.microsoft.com/en-us/graph/docs/concepts/use_the_api

### Platform Specific Notes

Some platforms require certain settings before it will connet to the API.

#### Android

Open `AndroidManifest.xml` and add the `BrowserTabActivity` with intent-filter to register the URL scheme:
```xml
<application ... >
    <activity android:name="microsoft.identity.client.BrowserTabActivity">
      <intent-filter>
        <action android:name="android.intent.action.VIEW" />
        <category android:name="android.intent.category.DEFAULT" />
        <category android:name="android.intent.category.BROWSABLE" />
        <data android:scheme="msal[Your Client ID]" android:host="auth" />
      </intent-filter>
    </activity>
  </application>
```

#### IOS

To configure URL Schemes of IOS platform open the `Info.plist` file and add an URL type under advanced tab and set the role option to `None` and set `msal[Your Client ID]` to the URL Schemes.


### Properties
| Property     | Type         | Android | iOS | Windows |
| :----------- | :----------- | :------ | :-- | :------ |
| ClientId            | string           | x       |  x  |    x    |
| Scopes            | string[]           | x       |  x  |    x    |
| UserName            | string           | x       |  x  |    x    |

### Methods
| Method       | Return Type  | Parameters                          | Android | iOS | Windows |
| :----------- | :----------- | :-----------                        | :------ | :-- | :------ |
| Initialize         | Task| clientId -> string,<br> scopes -> string[],<br> redirectUri -> string| x       | x   | x       |
| SignIn         | Task| -| x       | x   | x       |
| SignOut         | Task| -| x       | x   | x       |
| GetRequest&lt;T&gt;*         | Task&lt;T&gt;| url -> string| x       | x   | x       |
| PostRequest&lt;T&gt;*         | Task&lt;T&gt;| url -> string,<br> body -> string| x       | x   | x       |
    
#### Note
*: T can be string, byte[] or System.IO.Stream
