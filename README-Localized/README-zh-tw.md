# <a name="microsoft-graph-sdk-snippets-library-for-xamarin.forms"></a>Microsoft Graph SDK 程式碼片段程式庫 (適用於 Xamarin.Forms)

##<a name="table-of-contents"></a>目錄

* [必要條件](#prerequisites)
* [註冊和設定應用程式](#register)
* [建置及偵錯](#build)
* [執行範例](#run)
* [範例如何影響帳戶資料](#how-the-sample-affects-your-tenant-data)
* [新增程式碼片段](#add-a-snippet)
* [問題和建議](#questions)
* [參與](#contributing")
* [其他資源](#additional-resources)

<a name="introduction"></a> 此範例提供程式碼片段的儲存機制，使用 Microsoft Graph 以執行一般工作，例如傳送郵件、管理群組及其他 Xamarin.Forms 應用程式內的活動。它會使用 [Microsoft Graph.NET 用戶端 SDK](https://github.com/microsoftgraph/msgraph-sdk-dotnet)，使用 Microsoft Graph 所傳回的資料。 

範例會使用 [Microsoft 驗證程式庫 (MSAL)](https://www.nuget.org/packages/Microsoft.Identity.Client/) 進行驗證。MSAL SDK 提供功能以使用 [Azure AD v2.0 端點](https://graph.microsoft.io/en-us/docs/authorization/converged_auth)，可讓開發人員撰寫單一程式碼流程，處理使用者的工作或學校 (Azure Active Directory) 和個人 (Microsoft) 帳戶的驗證，包括 Office 365、Outlook.com 和 OneDrive 帳戶。

> **附註** MSAL SDK 目前是發行前版本，因此不應該用於實際執行程式碼。在這裡僅供說明目的使用。

應用程式顯示 UI 代表一般使用者工作，或「故事」。每一個故事是由一或多個程式碼片段所組成。故事是根據帳戶類型和權限層級分組。使用者可以登入其帳戶，並執行選取的故事。如果成功，每個故事會變成綠色，如果失敗則變成紅色。額外的資訊會傳送到 [輸出] 視窗中。

<a name="prerequisites"></a>
## <a name="prerequisites"></a>必要條件 ##

此範例需要下列項目：  

  * [Visual Studio 2015](https://www.visualstudio.com/downloads) 
  * [Xamarin for Visual Studio](https://www.xamarin.com/visual-studio)
  * Windows 10 ([已啟用開發模式](https://msdn.microsoft.com/library/windows/apps/xaml/dn706236.aspx))
  * [Microsoft](https://www.outlook.com) 或[商務用 Office 365 帳戶](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_Office365Account)

如果您想要執行這個範例中的 iOS 專案，您需要下列項目：

  * 最新的 iOS SDK
  * 最新版本的 Xcode
  * Mac OS X Yosemite(10.10) 和更新版本 
  * [Xamarin.iOS](https://developer.xamarin.com/guides/ios/getting_started/installation/mac/)
  * [連接至 Visual Studio 的 Xamarin Mac 代理程式](https://developer.xamarin.com/guides/ios/getting_started/installation/windows/connecting-to-mac/)

如果您想要執行 Android 專案，您可以使用[適用於 Android 的 Visual Studio 模擬器](https://www.visualstudio.com/features/msft-android-emulator-vs.aspx)。

<a name="register"></a>
##<a name="register-and-configure-the-app"></a>註冊和設定應用程式

1. 使用您的個人或工作或學校帳戶登入[應用程式註冊入口網站](https://apps.dev.microsoft.com/)。
2. 選取 [新增應用程式]****。
3. 為應用程式輸入名稱，然後選取 [建立應用程式]****。
    
    [註冊] 頁面隨即顯示，列出您的應用程式的屬性。
 
4. 在 [平台]**** 底下，選取 [新增平台]****。
5. 選取 [行動應用程式]****。
6. 將用戶端識別碼 (應用程式識別碼) 值複製到剪貼簿。您必須將這些值輸入範例應用程式中。

    應用程式識別碼是您的應用程式的唯一識別碼。

7. 選取 [儲存]****。

<a name="build"></a>
## <a name="build-and-debug"></a>建置和偵錯 ##

**附註︰**如果您在步驟 2 安裝封裝時看到任何錯誤，請確定您放置解決方案的本機路徑不會太長/太深。將解決方案移靠近您的磁碟機根目錄可解決這個問題。

1. 開啟解決方案的 **Graph_Xamarin_CS_Snippets (Portable)** 專案內的 App.cs 檔案。

    ![Visual Studio 中方案總管窗格的螢幕擷取畫面，在 Graph_Xamarin_CS_Snippets 專案中選取 App.cs 檔案](/readme-images/Appdotcs.png "在 Graph_Xamarin_CS_Snippets 專案中開啟 App.cs 檔案")

2. 在 Visual Studio 中載入解決方案之後，設定範例以使用用戶端識別碼，該識別碼是您藉由讓其成為 App.cs 檔案中的 **ClientId** 變數的值來註冊的。


    ![App.cs 檔案中 ClientId 變數的螢幕擷取畫面，目前設為空白字串。](/readme-images/appId.png "App.cs 檔案中的用戶端識別碼值")

2.  如果您計劃使用沒有系統管理權限的工作或學校帳戶登入範例，您必須註解化程式碼，該程式碼要求需要系統管理權限的範圍。如果您未註解化這些行，您就無法使用您的工作或學校帳戶登入 (如果您使用個人帳戶登入，會略過這些範圍要求。)

    在 `AuthenticationHelper.cs` 檔案的 `GetTokenForUserAsync()` 方法中，註解化下列範圍要求：
    
    ```
        "https://graph.microsoft.com/Directory.AccessAsUser.All",
        "https://graph.microsoft.com/User.ReadWrite.All",
        "https://graph.microsoft.com/Group.ReadWrite.All",
    ```

3. 選取您要執行的專案。如果您選取 [通用 Windows 平台] 選項，您可以在本機機器上執行範例。如果您想要執行 iOS 專案，您必須連接至[已安裝 Xamarin 工具的 Mac](https://developer.xamarin.com/guides/ios/getting_started/installation/windows/connecting-to-mac/)。(您也可以在 Mac 上，在 Xamarin Studio 中開啟此解決方案，然後直接從那裡執行範例。)如果您想要執行 Android 專案，您可以使用[適用於 Android 的 Visual Studio 模擬器](https://www.visualstudio.com/features/msft-android-emulator-vs.aspx)。 

    ![Visual Studio 工具列的螢幕擷取畫面，選取 iOS 做為啟動專案。](/readme-images/SelectProject.png "在 Visual Studio 中選取專案")

4. 按 F5 進行建置和偵錯。執行解決方案並且登入您的個人或工作或學校帳戶。
    > **附註** 您可能必須開啟組建組態管理員，以確定針對 UWP 專案選取建置和部署步驟。

<a name="run"></a>
## <a name="run-the-sample"></a>執行範例

啟動時，應用程式會顯示清單，代表一般使用者工作，或「故事」。每一個故事是由一或多個程式碼片段所組成。故事是根據帳戶類型和權限層級分組：

- 適用於工作或學校和個人帳戶的工作，例如取得和傳送電子郵件、建立檔案等等。
- 只適用於工作或學校帳戶的工作，例如取得使用者的經理或帳戶相片。
- 只適用於具有管理權限的工作或學校帳戶的工作，例如取得群組成員，或建立新的使用者帳戶。

選取您想要執行的故事，然後選擇 [執行選取] 按鈕。系統會提示您使用您的工作或學校或個人帳戶登入。請注意，如果您使用對於您已選取的故事沒有適用權限的帳戶登入 (例如，如果您選取只適用於工作或學校帳戶的故事，然後使用個人帳戶登入)，這些故事就會失敗。

如果成功，每個故事會變成綠色，如果失敗則變成紅色。額外的資訊會傳送到 [輸出] 視窗中。 

<a name="#how-the-sample-affects-your-tenant-data"></a>
##<a name="how-the-sample-affects-your-account-data"></a>範例如何影響帳戶資料

這個範例會執行命令，該命令會建立、讀取、更新或刪除資料。無法編輯或刪除您的實際帳戶資料。不過，它可能會在您的帳戶中建立並將資料成品留在其中，作為其作業的一部分︰當執行建立、更新或刪除的命令時，這個範例會建立假的實體，例如新的使用者或群組，如此便不會影響您的實際帳戶資料。 

如果您選擇建立或更新實體的故事，範例可能會在您的帳戶中留下這種假實體。例如，選擇執行「更新群組」故事，會建立新的群組，然後更新它。在此情況下，新的群組在執行範例之後會保留在您的帳戶中。

<a name="add-a-snippet"></a>
##<a name="add-a-snippet"></a>新增程式碼片段

此專案包含兩個程式碼片段檔案︰ 

- Groups\GroupSnippets.cs 
- Users\UserSnippets.cs.

如果您有自己的程式碼片段，想要在此專案中執行，請遵循下列三個步驟：

1. **將程式碼片段新增至程式碼片段檔案。**請務必包括 try/catch 區塊。 

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

2. **建立使用程式碼片段的故事，並將它新增至相關聯的故事檔案。**例如，`TryGetMeAsync()` 故事使用 Users\UserStories.cs 檔案內的 `GetMeAsync()` 程式碼片段：

        public static async Task<bool> TryGetMeAsync()
        {
            var currentUser = await UserSnippets.GetMeAsync();

            return currentUser != null;
        }        


有時候您的故事除了您實作的程式碼片段之外，還需要執行程式碼片段。例如，如果您想要更新事件，您首先要使用 `CreateEventAsync()` 方法，以建立事件。然後您可以更新它。務必一律使用已存在於程式碼片段檔案中的程式碼片段。如果您需要的作業不存在，您必須建立它，然後將它包含在您的故事中。最佳作法是刪除您在故事中建立的任何實體，特別是如果您正在測試或開發人員帳戶以外的任何項目上工作。

3. **在 MainPage.xaml.cs 中將您的故事新增至故事清單** (`CreateStoryList()` 方法內)：
    
    `snippetList.Children.Add(new CheckBox 
        { StoryName = "Get Me", GroupName = "Users", AccountType = "All", RunStoryAsync = UserStories.TryGetMeAsync });`
    
現在您可以測試您的程式碼片段。當您執行應用程式時，您的故事會顯示為新的項目。選取您的程式碼片段的核取方塊，然後再執行它。使用這一個機會來偵錯程式碼片段。

<a name="questions"></a>
## <a name="questions-and-comments"></a>問題和建議

我們很樂於收到您對於 Microsoft Graph 程式碼片段範例 (適用於 Xamarin Forms) 專案的意見反應。您可以在此儲存機制的[問題](https://github.com/MicrosoftGraph/xamarin-csharp-snippets-sample/issues)區段中，將您的問題及建議傳送給我們。

我們很重視您的意見。請透過 [Stack Overflow](http://stackoverflow.com/questions/tagged/office365+or+microsoftgraph) 與我們連絡。以 [MicrosoftGraph] 標記您的問題。

<a name="contributing"></a>
## <a name="contributing"></a>參與 ##

如果您想要參與這個範例，請參閱 [CONTRIBUTING.MD](/CONTRIBUTING.md)。

此專案已採用 [Microsoft 開放原始碼執行](https://opensource.microsoft.com/codeofconduct/)。如需詳細資訊，請參閱[程式碼執行常見問題集](https://opensource.microsoft.com/codeofconduct/faq/)，如果有其他問題或意見，請連絡 [opencode@microsoft.com](mailto:opencode@microsoft.com)。


<a name="additional-resources"></a>
## <a name="additional-resources"></a>其他資源 ##

- [其他 Microsoft Graph connect 範例](https://github.com/MicrosoftGraph?utf8=%E2%9C%93&query=-Connect)
- [Microsoft Graph 概觀](http://graph.microsoft.io)
- [Office 開發人員程式碼範例](http://dev.office.com/code-samples)
- [Office 開發人員中心](http://dev.office.com/)


## <a name="copyright"></a>著作權
Copyright (c) 2016 Microsoft.著作權所有，並保留一切權利。


