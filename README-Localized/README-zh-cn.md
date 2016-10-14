# <a name="microsoft-graph-sdk-snippets-library-for-xamarin.forms"></a>Xamarin.Forms 的 Microsoft Graph SDK 代码段库

##<a name="table-of-contents"></a>目录

* [先决条件](#prerequisites)
* [注册和配置应用](#register)
* [构建和调试](#build)
* [运行示例](#run)
* [示例对帐户数据有何影响](#how-the-sample-affects-your-tenant-data)
* [添加代码段](#add-a-snippet)
* [问题和意见](#questions)
* [参与](#contributing")
* [其他资源](#additional-resources)

<a name="introduction"></a> 此示例项目提供使用 Microsoft Graph 执行常见任务的代码段存储库，例如发送电子邮件、管理组和 Xamarin.Forms 应用内的其他活动。它使用 [Microsoft Graph .NET Client SDK](https://github.com/microsoftgraph/msgraph-sdk-dotnet) 以结合使用由 Microsoft Graph 返回的数据。 

此示例使用 [Microsoft 身份验证库 (MSAL)](https://www.nuget.org/packages/Microsoft.Identity.Client/) 进行身份验证。MSAL SDK 提供使用 [Azure AD v2.0 终结点](https://graph.microsoft.io/en-us/docs/authorization/converged_auth)的功能以使开发人员写入单个代码流，用于对用户的工作或学校 (Azure Active Directory) 或个人 (Microsoft) 帐户（包括 Office 365、Outlook.com 和 OneDrive 帐户）进行身份验证。

> **注意** MSAL SDK 目前处于预发布状态，因此，不应该在生产代码中使用。在此处仅用于说明目的。

该应用显示表示常见用户任务或情景的 UI。每个情景由一个或多个代码段组成。情景按帐户类型和权限级别进行分组。用户可以登录到自己的帐户并运行所选情景。每个情景在运行成功时变为绿色，在运行失败时变为红色。其他信息被发送至输出窗口。

<a name="prerequisites"></a>
## <a name="prerequisites"></a>先决条件 ##

此示例需要以下各项：  

  * [Visual Studio 2015](https://www.visualstudio.com/downloads) 
  * [Visual Studio 的 Xamarin](https://www.xamarin.com/visual-studio)
  * Windows 10（[已启用开发模式](https://msdn.microsoft.com/library/windows/apps/xaml/dn706236.aspx)）
  * [Microsoft](https://www.outlook.com) 或 [Office 365 商业版帐户](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_Office365Account)

如果想要在此示例中运行 iOS 项目，则要求如下：

  * 最新的 iOS SDK
  * Xcode 的最新版本
  * Mac OS X Yosemite (10.10) 和更高版本 
  * [Xamarin.iOS](https://developer.xamarin.com/guides/ios/getting_started/installation/mac/)
  * [连接到 Visual Studio 的 Xamarin Mac 代理](https://developer.xamarin.com/guides/ios/getting_started/installation/windows/connecting-to-mac/)

如果想要运行 Android 项目，可以使用 [适用于 Android 的 Visual Studio 模拟器](https://www.visualstudio.com/features/msft-android-emulator-vs.aspx)。

<a name="register"></a>
##<a name="register-and-configure-the-app"></a>注册和配置应用

1. 使用个人或工作或学校帐户登录到 [应用注册门户](https://apps.dev.microsoft.com/)。
2. 选择“**添加应用**”。
3. 为应用输入名称，并选择“**创建应用程序**”。
    
    将显示注册页，其中列出应用的属性。
 
4. 在“**平台**”下，选择“**添加平台**”。
5. 选择“**移动应用程序**”。
6. 将客户端 ID（应用 ID）值复制到剪贴板。将需要在示例应用中输入这些值。

    应用 ID 是应用的唯一标识符。

7. 选择“**保存**”。

<a name="build"></a>
## <a name="build-and-debug"></a>构建和调试 ##

**注意：**如果在步骤 2 安装程序包时出现任何错误，请确保放置该解决方案的本地路径并未太长/太深。将解决方案移动到更接近驱动器根目录的位置可以解决此问题。

1. 打开解决方案的 **Graph_Xamarin_CS_Snippets（可移植）**项目内的 App.cs 文件。

    ![Visual Studio 中的解决方案资源管理器窗格的屏幕截图（在 Graph_Xamarin_CS_Snippets 项目中选择了 App.cs 文件）](/readme-images/Appdotcs.png "在 Graph_Xamarin_CS_Snippets 项目中打开 App.cs 文件")

2. 在 Visual Studio 中加载解决方案后，通过将注册的客户端 ID 生成为 App.cs 文件中的 **ClientId** 变量来配置使用注册的客户端 ID 的示例。


    ![App.cs 文件中 ClientId 变量的屏幕截图，目前被设置为空字符串。](/readme-images/appId.png "App.cs 文件中的客户端 ID 值")

2.  如果计划使用没有管理员权限的工作或学校帐户登录到示例，则需剔除所需的作用域（要求管理员权限）的代码。如果未剔除这些行，则不能使用你的工作或学校帐户登录（如果使用个人帐户登录，则这些作用域请求将被忽略。）

    在 `AuthenticationHelper.cs` 文件的 `GetTokenForUserAsync()` 方法中剔除以下作用域请求：
    
    ```
        "https://graph.microsoft.com/Directory.AccessAsUser.All",
        "https://graph.microsoft.com/User.ReadWrite.All",
        "https://graph.microsoft.com/Group.ReadWrite.All",
    ```

3. 选择想要运行的项目。如果选择“通用 Windows 平台”选项，则可以在本地计算机上运行示例。如果想要运行 iOS 项目，则需连接到安装在其上的 [具有 Xamarin 工具的 Mac](https://developer.xamarin.com/guides/ios/getting_started/installation/windows/connecting-to-mac/)。（还可以在 Mac 上的 Xamarin Studio 中打开此解决方案并直接从此处运行示例。）如果想要运行 Android 项目，可以使用[适用于 Android 的 Visual Studio 模拟器](https://www.visualstudio.com/features/msft-android-emulator-vs.aspx)。 

    ![Visual Studio 工具栏的屏幕截图（iOS 被选为启动项目）。](/readme-images/SelectProject.png "选择 Visual Studio 中的项目")

4. 按 F5 进行构建和调试。运行此解决方案并使用个人或工作或学校帐户登录。
    > **注意** 可能需要打开生成配置管理器，以确保为 UWP 项目选择“生成”和“部署”步骤。

<a name="run"></a>
## <a name="run-the-sample"></a>运行示例

启动时，应用将显示表示常见用户任务或情景的列表。每个情景由一个或多个代码段组成。情景按帐户类型和权限级别进行分组：

- 适用于工作或学校以及个人帐户的任务，如接收和发送电子邮件、创建文件等。
- 仅适用于工作或学校帐户的任务，如获取用户的管理器或帐户照片。
- 仅适用于具有管理权限的工作或学校帐户的任务，如获取组成员或新建用户帐户。

选择想要执行的情景，并选择“运行所选项”按钮。系统将提示你使用你的工作或学校或个人帐户登录。请注意，如果你对所选情景未使用具有相应权限的帐户进行登录（例如，如果你选择只适用于工作或学校帐户的情景，然后使用个人帐户登录），这些情景会失败。

每个情景在运行成功时变为绿色，在运行失败时变为红色。其他信息被发送至输出窗口。 

<a name="#how-the-sample-affects-your-tenant-data"></a>
##<a name="how-the-sample-affects-your-account-data"></a>示例如何影响你的帐户数据

此示例运行创建、读取、更新或删除数据的命令。它不会编辑或删除你的实际帐户数据。但是，它可能会创建数据项目并将其作为操作的一部分保留在你的帐户中：当运行创建、更新或删除命令时，示例将创建虚设实体，如新用户或组，以免影响你的实际帐户数据。 

如果你选择创建或更新实体的情景，则示例可能将该虚设实体遗留在你的帐户中。例如，选择运行“更新组”情景将创建新组，然后对其更新。在这种情况下，示例已运行后，新组仍保留在你的帐户中。

<a name="add-a-snippet"></a>
##<a name="add-a-snippet"></a>添加代码段

该项目包括两个代码段文件： 

- Groups\GroupSnippets.cs 
- Users\UserSnippets.cs.

如果想要在该项目中运行自己的代码段，只需执行以下三个步骤：

1. **将代码段添加到代码段文件。**请务必包括 try/catch 块。 

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

2. **创建使用代码段的情景并将其添加到相关联的情景文件。**例如，`TryGetMeAsync()` 情景使用 Users\UserStories.cs 文件内的 `GetMeAsync()` 代码段：

        public static async Task<bool> TryGetMeAsync()
        {
            var currentUser = await UserSnippets.GetMeAsync();

            return currentUser != null;
        }        


有时，情景需要运行正在实现的代码段以外的代码段。例如，如果想要更新某个事件，首先需要使用 `CreateEventAsync()` 方法来创建一个事件。然后可以对其更新。始终确保使用总是存在于代码段文件中的代码段。如果所需的操作不存在，则需对其创建并将其包括在你的情景中。最好的做法是删除在情境中创建的任何实体，尤其是当你仅进行测试或使用开发人员帐户时。

3. **将你的情景添加到 MainPage.xaml.cs 中的情景列表**（在 `CreateStoryList()` 方法内）：
    
    `snippetList.Children.Add(new CheckBox 
        { StoryName = "Get Me", GroupName = "Users", AccountType = "All", RunStoryAsync = UserStories.TryGetMeAsync });`
    
现在可以测试你的代码段。运行应用时，你的情景将作为新项目出现。为你的代码段选择复选框，然后运行它。将其作为调试你的代码段的机会。

<a name="questions"></a>
## <a name="questions-and-comments"></a>问题和意见

我们乐意倾听你有关 Xamarin.Forms 的 Microsoft Graph 代码段示例项目的反馈。你可以在该存储库中的 [问题](https://github.com/MicrosoftGraph/xamarin-csharp-snippets-sample/issues) 部分将问题和建议发送给我们。

你的反馈对我们意义重大。请在 [Stack Overflow](http://stackoverflow.com/questions/tagged/office365+or+microsoftgraph) 上与我们联系。使用 [MicrosoftGraph] 标记出你的问题。

<a name="contributing"></a>
## <a name="contributing"></a>参与 ##

如果想要参与本示例，请参阅 [CONTRIBUTING.MD](/CONTRIBUTING.md)。

此项目采用 [Microsoft 开源行为准则](https://opensource.microsoft.com/codeofconduct/)。有关详细信息，请参阅 [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/)（行为准则常见问题解答），有任何其他问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。


<a name="additional-resources"></a>
## <a name="additional-resources"></a>其他资源 ##

- [其他 Microsoft Graph Connect 示例](https://github.com/MicrosoftGraph?utf8=%E2%9C%93&query=-Connect)
- [Microsoft Graph 概述](http://graph.microsoft.io)
- [Office 开发人员代码示例](http://dev.office.com/code-samples)
- [Office 开发人员中心](http://dev.office.com/)


## <a name="copyright"></a>版权
版权所有 (c) 2016 Microsoft。保留所有权利。


