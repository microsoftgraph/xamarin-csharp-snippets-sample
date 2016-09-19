# Microsoft Graph SDK Snippets Library for Xamarin.Forms

##Table of contents

* [Prerequisites](#prerequisites)
* [Register and configure the app](#register)
* [Build and debug](#build)
* [Run the sample](#run)
* [How the sample affects your account data](#how-the-sample-affects-your-tenant-data)
* [Add a snippet](#add-a-snippet)
* [Questions and comments](#questions)
* [Contributing](#contributing")
* [Additional resources](#additional-resources)

<a name="introduction"></a>
This sample project provides a repository of code snippets that use the Microsoft Graph to perform common tasks, such as sending email, managing groups, and other activities from within a Xamarin.Forms app. It uses the [Microsoft Graph .NET Client SDK](https://github.com/microsoftgraph/msgraph-sdk-dotnet) to work with data returned by Microsoft Graph. 

The sample uses the [Microsoft Authentication Library (MSAL)](https://www.nuget.org/packages/Microsoft.Identity.Client/) for authentication. The MSAL SDK provides features for working with the [Azure AD v2.0 endpoint](https://graph.microsoft.io/en-us/docs/authorization/converged_auth), which enables developers to write a single code flow that handles authentication for both users' work or school (Azure Active Directory) or personal (Microsoft) accounts, including Office 365, Outlook.com, and OneDrive accounts.

> **Note** The MSAL SDK is currently in prerelease, and as such should not be used in production code. It is used here for illustrative purposes only.

The app displays UI representing common user tasks, or 'stories'. Each story is comprised of one or more code snippets. The stories are grouped by the account type and permission level. The user can log into their account and run the selected stories. Each story turns green if it succeeds, and red if it fails. Additional information is sent to the Output window.

<a name="prerequisites"></a>
## Prerequisites ##

This sample requires the following:  

  * [Visual Studio 2015](https://www.visualstudio.com/downloads) 
  * [Xamarin for Visual Studio](https://www.xamarin.com/visual-studio)
  * Windows 10 ([development mode enabled](https://msdn.microsoft.com/library/windows/apps/xaml/dn706236.aspx))
  * Either a [Microsoft](https://www.outlook.com) or [Office 365 for business account](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_Office365Account)

If you want to run the iOS project in this sample, you'll need the following:

  * The latest iOS SDK
  * The latest version of Xcode
  * Mac OS X Yosemite(10.10) & above 
  * [Xamarin.iOS](https://developer.xamarin.com/guides/ios/getting_started/installation/mac/)
  * A [Xamarin Mac agent connected to Visual Studio](https://developer.xamarin.com/guides/ios/getting_started/installation/windows/connecting-to-mac/)

You can use the [Visual Studio Emulator for Android](https://www.visualstudio.com/features/msft-android-emulator-vs.aspx) if you want to run the Android project.

<a name="register"></a>
##Register and configure the app

1. Sign into the [App Registration Portal](https://apps.dev.microsoft.com/) using either your personal or work or school account.
2. Select **Add an app**.
3. Enter a name for the app, and select **Create application**.
	
	The registration page displays, listing the properties of your app.
 
4. Under **Platforms**, select **Add platform**.
5. Select **Mobile application**.
6. Copy the Client Id (App Id) value to the clipboard. You'll need to enter these values into the sample app.

	The app id is a unique identifier for your app.

7. Select **Save**.

<a name="build"></a>
## Build and debug ##

**Note:** If you see any errors while installing packages during step 2, make sure the local path where you placed the solution is not too long/deep. Moving the solution closer to the root of your drive resolves this issue.

1. Open the App.cs file inside the **Graph_Xamarin_CS_Snippets (Portable)** project of the solution.

    ![Screenshot of the Solution Explorer pane in Visual Studio, with App.cs file selected in the Graph_Xamarin_CS_Snippets project](/readme-images/Appdotcs.png "Open App.cs file in Graph_Xamarin_CS_Snippets project")

2. After you've loaded the solution in Visual Studio, configure the sample to use the client id that you registered by making this the value of the **ClientId** variable in the App.cs file.


    ![Screenshot of the ClientId variable in the App.cs file, currently set to an empty string.](/readme-images/appId.png "Client ID value in App.cs file")

2.	If you are planning on signing into the sample with a work or school account that does not have admin permissions, you'll need to comment out  code that requests scopes that require admin permissions. If you don't comment out these lines, you won't be able to sign in with your work or school account (if you sign in with a personal account, these scope requests are ignored.)

	In the `GetTokenForUserAsync()` method of the `AuthenticationHelper.cs` file, comment out the following scope requests:
	
	```
		"https://graph.microsoft.com/Directory.AccessAsUser.All",
	    "https://graph.microsoft.com/User.ReadWrite.All",
	    "https://graph.microsoft.com/Group.ReadWrite.All",
	```

3. Select the project that you want to run. If you select the Universal Windows Platform option, you can run the sample on the local machine. If you want to run the iOS project, you'll need to connect to a [Mac that has the Xamarin tools](https://developer.xamarin.com/guides/ios/getting_started/installation/windows/connecting-to-mac/) installed on it. (You can also open this solution in Xamarin Studio on a Mac and run the sample directly from there.) You can use the [Visual Studio Emulator for Android](https://www.visualstudio.com/features/msft-android-emulator-vs.aspx) if you want to run the Android project. 

    ![Screenshot of the Visual Studio toolbar, with iOS selected as the start-up project.](/readme-images/SelectProject.png "Select project in Visual Studio")

4. Press F5 to build and debug. Run the solution and sign in with either your personal or work or school account.
    > **Note** You might have to open the Build Configuration Manager to make sure that the Build and Deploy steps are selected for the UWP project.

<a name="run"></a>
## Run the sample

When launched, the app displays a list representing common user tasks, or 'stories'. Each story is comprised of one or more code snippets. The stories are grouped by the account type and permission level:

- Tasks that are applicable to both work or school and personal accounts, such as getting and sending email, creating files, etc.
- Tasks that are only applicable to work or school accounts, such as getting a user's manager or account photo.
- Tasks that are only applicable to a work or school account with administrative permissions, such as getting group members or creating new user accounts.

Select the stories you want to execute, and choose the 'run selected' button. You'll be prompted to log in with your work or school or personal account. Be aware that if you log in with an account that doesn't have applicable permissions for the stories you've selected (for example, if you select stories that are applicable only to a work or school account, and then log in with a personal account), those stories will fail.

Each story turns green if it succeeds, and red if it fails. Additional information is sent to the Output window. 

<a name="#how-the-sample-affects-your-tenant-data"></a>
##How the sample affects your account data

This sample runs commands that create, read, update, or delete data. It will not edit or delete your actual account data. However, it may create and leave data artifacts in your account as part of its operation: when running commands that create, update or delete, the sample creates fake entities, such as new users or groups, so as not to affect your actual account data. 

The sample may leave behind such fake entities in your account, if you choose stories that create or update entities. For example, choosing the run the 'update group' story creates a new group, then updates it. In this case, the new group remains in your account after the sample has run.

<a name="add-a-snippet"></a>
##Add a snippet

This project includes two snippets files: 

- Groups\GroupSnippets.cs 
- Users\UserSnippets.cs.

If you have a snippet of your own that you would like to run in this project, just follow these three steps:

1. **Add your snippet to the snippets file.** Be sure to include a try/catch block. 

        public static async Task<string> GetMeAsync()
        {
            string currentUserName = null;

            try
            {
                var graphClient = AuthenticationHelper.GetAuthenticatedClientAsync();

                var currentUserObject = await graphClient.Me.Request().GetAsync();
                currentUserName = currentUserObject.DisplayName;

                if ( currentUserName != null)
                {
                    Debug.WriteLine("Got user: " + currentUserName);
                }

            }

2. **Create a story that uses your snippet and add it to the associated stories file.** For example, the `TryGetMeAsync()` story uses the `GetMeAsync()` snippet inside the Users\UserStories.cs file:

        public static async Task<bool> TryGetMeAsync()
        {
            var currentUser = await UserSnippets.GetMeAsync();

            return currentUser != null;
        }        


Sometimes your story will need to run snippets in addition to the one that you're implementing. For example, if you want to update an event, you first need to use the `CreateEventAsync()` method to create an event. Then you can update it. Always be sure to use snippets that already exist in the snippets file. If the operation you need doesn't exist, you'll have to create it and then include it in your story. It's a best practice to delete any entities that you create in a story, especially if you're working on anything other than a test or developer account.

3. **Add your story to the story list in MainPage.xaml.cs** (inside the `CreateStoryList()` method):
	
	`snippetList.Children.Add(new CheckBox 
		{ StoryName = "Get Me", GroupName = "Users", AccountType = "All", RunStoryAsync = UserStories.TryGetMeAsync });`
	
Now you can test your snippet. When you run the app, your story will appear as a new item. Select the check box for your snippet, and then run it. Use this as an opportunity to debug your snippet.

<a name="questions"></a>
## Questions and comments

We'd love to get your feedback about the Microsoft Graph Snippets Sample for Xamarin.Forms project. You can send your questions and suggestions to us in the [Issues](https://github.com/MicrosoftGraph/xamarin-csharp-snippets-sample/issues) section of this repository.

Your feedback is important to us. Connect with us on [Stack Overflow](http://stackoverflow.com/questions/tagged/office365+or+microsoftgraph). Tag your questions with [MicrosoftGraph].

<a name="contributing"></a>
## Contributing ##

If you'd like to contribute to this sample, see [CONTRIBUTING.MD](/CONTRIBUTING.md).

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.


<a name="additional-resources"></a>
## Additional resources ##

- [Other Microsoft Graph Connect samples](https://github.com/MicrosoftGraph?utf8=%E2%9C%93&query=-Connect)
- [Microsoft Graph overview](http://graph.microsoft.io)
- [Office developer code samples](http://dev.office.com/code-samples)
- [Office dev center](http://dev.office.com/)


## Copyright
Copyright (c) 2016 Microsoft. All rights reserved.


