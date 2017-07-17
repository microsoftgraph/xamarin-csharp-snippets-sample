# Xamarin.Forms 用 Microsoft Graph SDK スニペット ライブラリ
<a id="microsoft-graph-sdk-snippets-library-for-xamarinforms" class="xliff"></a>

## 目次
<a id="table-of-contents" class="xliff"></a>

* [前提条件](#prerequisites)
* [アプリを登録して構成する](#register)
* [ビルドとデバッグ](#build)
* [サンプルの実行](#run)
* [サンプルによるアカウント データへの影響](#how-the-sample-affects-your-tenant-data)
* [スニペットの追加](#add-a-snippet)
* [質問とコメント](#questions)
* [投稿](#contributing")
* [その他のリソース](#additional-resources)

<a name="introduction"></a> このサンプル プロジェクトには、Xamarin.Forms アプリ内からのメール送信、グループ管理、および他のアクティビティなどの一般的なタスクを実行するために Microsoft Graph を使用する、コード スニペットのリポジトリが用意されています。[Microsoft Graph .NET クライアント SDK](https://github.com/microsoftgraph/msgraph-sdk-dotnet) を使用して、Microsoft Graph が返すデータを操作します。 

サンプルでは認証に [Microsoft 認証ライブラリ (MSAL)](https://www.nuget.org/packages/Microsoft.Identity.Client/) を使用します。MSAL SDK には、[Azure AD v2.0 エンドポイント](https://graph.microsoft.io/en-us/docs/authorization/converged_auth)を操作するための機能が用意されており、開発者はユーザーの職場または学校 (Azure Active Directory) アカウント、または Office 365、Outlook.com、および OneDrive の各アカウントなどの個人用 (Microsoft) アカウントの両方に対する認証を処理する 1 つのコード フローを記述することができます。

## MSAL プレビューに関する重要な注意事項
<a id="important-note-about-the-msal-preview" class="xliff"></a>

このライブラリは、運用環境での使用に適しています。 このライブラリに対しては、現在の運用ライブラリと同じ運用レベルのサポートを提供します。 プレビュー中にこのライブラリの API、内部キャッシュの形式、および他のメカニズムを変更する場合があります。これは、バグの修正や機能強化の際に実行する必要があります。 これは、アプリケーションに影響を与える場合があります。 例えば、キャッシュ形式を変更すると、再度サインインが要求されるなどの影響をユーザーに与えます。 API を変更すると、コードの更新が要求される場合があります。 一般提供リリースが実施されると、プレビュー バージョンを使って作成されたアプリケーションは動作しなくなるため、6 か月以内に一般提供バージョンに更新することが求められます。

アプリには、共通のユーザー タスク、'ストーリー' を表す UI が表示されます。各ストーリーは、1 つ以上のコード スニペットで構成されます。ストーリーは、アカウントの種類とアクセス許可のレベルによってグループ化されます。ユーザーは自分のアカウントにログインして、選択したストーリーを実行できます。各ストーリーは、成功した場合は緑、失敗した場合は赤に変わります。追加情報は出力ウィンドウに送信されます。

<a name="prerequisites"></a>
## 前提条件
<a id="prerequisites" class="xliff"></a> ##

このサンプルを実行するには次のものが必要です:  

  * [Visual Studio 2015](https://www.visualstudio.com/downloads) 
  * [Xamarin for Visual Studio](https://www.xamarin.com/visual-studio)
  * Windows 10 ([開発モードが有効](https://msdn.microsoft.com/library/windows/apps/xaml/dn706236.aspx))
  * [Microsoft](https://www.outlook.com) または [Office 365 for Business アカウント](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_Office365Account)のいずれか

このサンプルで iOS プロジェクトを実行する場合は、以下のものが必要です:

  * 最新の iOS SDK
  * 最新バージョンの Xcode
  * Mac OS X Yosemite(10.10) 以上 
  * [Xamarin.iOS](https://developer.xamarin.com/guides/ios/getting_started/installation/mac/)
  * [Visual Studio に接続されている Xamarin Mac エージェント](https://developer.xamarin.com/guides/ios/getting_started/installation/windows/connecting-to-mac/)

Android プロジェクトを実行する場合は、[Visual Studio Emulator for Android](https://www.visualstudio.com/features/msft-android-emulator-vs.aspx) を使用できます。

<a name="register"></a>
## アプリを登録して構成する
<a id="register-and-configure-the-app" class="xliff"></a>

1. 個人用アカウント、あるいは職場または学校アカウントのいずれかを使用して、[アプリ登録ポータル](https://apps.dev.microsoft.com/)にサインインします。
2. **[アプリの追加]** を選択します。
3. アプリの名前を入力して、**[アプリケーションの作成]** を選択します。
    
    登録ページが表示され、アプリのプロパティが一覧表示されます。
 
4. **[プラットフォーム]** で、**[プラットフォームの追加]** を選択します。
5. **[モバイル アプリケーション]** を選択します。
6. クライアント ID (アプリ ID) の値をクリップボードにコピーします。サンプル アプリにこれらの値を入力する必要があります。

    アプリ ID は、アプリの一意識別子です。

7. **[保存]** を選択します。

<a name="build"></a>
## ビルドとデバッグ
<a id="build-and-debug" class="xliff"></a> ##

**注:**手順 2 でパッケージのインストール中にエラーが発生した場合は、ソリューションを保存したローカル パスが長すぎたり深すぎたりしていないかご確認ください。ドライブのルート近くにソリューションを移動すると問題が解決します。

1. ソリューションの **Graph_Xamarin_CS_Snippets (ポータブル)** プロジェクト内にある App.cs ファイルを開きます。

    ![Graph_Xamarin_CS_Snippets プロジェクトで App.cs ファイルが選択されている Visual Studio のソリューション エクスプローラー ウィンドウのスクリーン ショット](/readme-images/Appdotcs.png "Graph_Xamarin_CS_Snippets プロジェクトの App.cs file を開く")

2. Visual Studio にソリューションを読み込んだ後、登録したクライアント ID を App.cs ファイルの **ClientId** 変数の値にして、この値を使用するようにサンプルを構成します。


    ![現在空の文字列に設定されている App.cs ファイルの ClientId 変数のスクリーンショット。](/readme-images/appId.png "App.cs ファイルのクライアント ID の値")

2.  管理者のアクセス許可がない職場または学校のアカウントでサンプルにサインインしようとする場合は、管理者のアクセス許可が必要な適用範囲を要求するコードをコメントにする必要があります。これらの行をコメントにしないと、職場または学校のアカウントでサインインすることはできません (個人用アカウントでサインインする場合は、これらの適用範囲の要求は無視されます。)

    `AuthenticationHelper.cs` ファイルの `GetTokenForUserAsync()` メソッドでは、次の適用範囲の要求をコメントにします:
    
    ```
        "https://graph.microsoft.com/Directory.AccessAsUser.All",
        "https://graph.microsoft.com/User.ReadWrite.All",
        "https://graph.microsoft.com/Group.ReadWrite.All",
    ```

3. 実行するプロジェクトを選びます。ユニバーサル Windows プラットフォームのオプションを選択すると、ローカル コンピューターでサンプルを実行できます。iOS プロジェクトを実行する場合は、[Xamarin ツールがインストールされた Mac](https://developer.xamarin.com/guides/ios/getting_started/installation/windows/connecting-to-mac/) に接続する必要があります。(また、このソリューションを Mac 上の Xamarin Studio で開いて、そこからサンプルを直接実行することもできます。)Android プロジェクトを実行する場合は、[Visual Studio Emulator for Android](https://www.visualstudio.com/features/msft-android-emulator-vs.aspx) を使用できます。 

    ![スタートアップ プロジェクトとして iOS が選択されている Visual Studio ツールバーのスクリーンショット。](/readme-images/SelectProject.png "Visual Studio でプロジェクトを選択する")

4. F5 キーを押して、ビルドとデバッグを実行します。　ソリューションを実行し、個人用アカウント、あるいは職場または学校のアカウントのいずれかでサインインします。
    > **注** ビルド構成マネージャーを開いて、ビルドと展開の手順が UWP プロジェクトに対して選択されていることを確認することが必要な場合があります。

<a name="run"></a>
## サンプルの実行
<a id="run-the-sample" class="xliff"></a>

起動すると、共通のユーザー タスク、'ストーリー' を表す一覧がアプリに表示されます。各ストーリーは、1 つ以上のコード スニペットで構成されます。ストーリーは、アカウントの種類とアクセス許可のレベルによってグループ化されます:

- メールの送受信、ファイルの作成など、職場または学校のアカウントおよび個人用アカウントの両方に適用可能なタスク。
- ユーザーの上司またはアカウントの写真の取得など、職場または学校のアカウントにのみ適用可能なタスク。
- グループ メンバーの取得または新しいユーザーの作成など、管理アクセス許可を持つ職場または学校のアカウントにのみ適用可能なタスク。

実行するストーリーを選択し、[選択項目の実行] ボタンを選択します。職場または学校のアカウント、あるいは個人用アカウントでログインするように求めるメッセージが表示されます。選択したストーリーに適用可能なアクセス許可がないアカウントでログインした場合 (たとえば、学校または職場のアカウントのみに適用可能なストーリーを選択して、個人用アカウントでログインした場合)、それらのストーリーは失敗しますので注意してください。

各ストーリーは、成功した場合は緑、失敗した場合は赤に変わります。追加情報は出力ウィンドウに送信されます。 

<a name="#how-the-sample-affects-your-tenant-data"></a>
##サンプルによるアカウント データへの影響
<a id="how-the-sample-affects-your-account-data" class="xliff"></a>

このサンプルでは、データを作成、読み取り、更新、または削除するコマンドを実行します。実際のアカウント データは編集も削除もされません。ただし、アカウント内のデータの成果物がその操作の一部として作成されて残る可能性があります。作成、更新または削除するコマンドを実行すると、サンプルでは実際のアカウント データに影響を与えないように、新しいユーザーまたはグループなど偽のエンティティが作成されます。 

エンティティを作成または更新するストーリーを選択すると、サンプルではこのような偽のエンティティが削除されずにアカウントに残る場合があります。たとえば、'グループの更新' ストーリーの実行を選択すると、新しいグループが作成されて更新されます。この例では、サンプルの実行後も、新しいグループはアカウントに残ります。

<a name="add-a-snippet"></a>
## スニペットの追加
<a id="add-a-snippet" class="xliff"></a>

このプロジェクトには次の 2 つのスニペット ファイルが含まれています: 

- Groups\GroupSnippets.cs 
- Users\UserSnippets.cs.

このプロジェクトで実行する独自のスニペットがある場合には、次の 3 つの手順のみを実行します:

1. **スニペットをスニペット ファイルに追加します。**try/catch ブロックが含まれていることを確認します。 

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

2. **スニペットを使用するストーリーを作成し、関連付けられているストーリー ファイルにそのストーリーを追加します。**たとえば、`TryGetMeAsync()` ストーリーは、Users\UserStories.cs ファイル内の `GetMeAsync()` スニペットを使用します:

        public static async Task<bool> TryGetMeAsync()
        {
            var currentUser = await UserSnippets.GetMeAsync();

            return currentUser != null;
        }        


ストーリーには、実装しているものに加えて、スニペットを実行することが必要な場合があります。たとえば、イベントを更新する場合は、まず `CreateEventAsync()` メソッドを使用して、イベントを作成します。これにより、そのイベントを更新できます。既にスニペット ファイルに存在するスニペットを常に使用してください。必要な操作が存在しない場合は、その操作を作成し、ストーリーに含める必要があります。ストーリーに作成するすべてのエンティティを削除することが、特に、テスト アカウントまたは開発者アカウント以外のアカウントで作業している場合は、ベスト プラクティスです。

3. **ストーリーを MainPage.xaml.cs のストーリー リストに追加します** (`CreateStoryList()` メソッド内):
    
    `snippetList.Children.Add(new CheckBox 
        { StoryName = "Get Me", GroupName = "Users", AccountType = "All", RunStoryAsync = UserStories.TryGetMeAsync });`
    
これで、スニペットをテストできます。アプリを実行すると、ストーリーは新しいアイテムとして表示されます。スニペットのチェック ボックスを選択して、実行します。スニペットをデバッグする機会として、これを使用します。

<a name="questions"></a>
## 質問とコメント
<a id="questions-and-comments" class="xliff"></a>

Xamarin.Forms プロジェクト用 Microsoft Graph スニペットのサンプルに関するフィードバックをお寄せください。質問や提案につきましては、このリポジトリの「[問題](https://github.com/MicrosoftGraph/xamarin-csharp-snippets-sample/issues)」セクションで送信できます。

お客様からのフィードバックは私たちにとって重要です。[スタック オーバーフロー](http://stackoverflow.com/questions/tagged/office365+or+microsoftgraph)でご連絡いただけます。ご質問には [MicrosoftGraph] のタグを付けてください。

<a name="contributing"></a>
## 投稿
<a id="contributing" class="xliff"></a> ##

このサンプルに投稿する場合は、[CONTRIBUTING.MD](/CONTRIBUTING.md) を参照してください。

このプロジェクトでは、[Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/) が採用されています。詳細については、「[規範に関する FAQ](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。または、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までにお問い合わせください。


<a name="additional-resources"></a>
## 追加リソース
<a id="additional-resources" class="xliff"></a> ##

- [その他の Microsoft Graph Connect サンプル](https://github.com/MicrosoftGraph?utf8=%E2%9C%93&query=-Connect)
- [Microsoft Graph の概要](http://graph.microsoft.io)
- [Office 開発者向けコード サンプル](http://dev.office.com/code-samples)
- [Office デベロッパー センター](http://dev.office.com/)


## 著作権
<a id="copyright" class="xliff"></a>
Copyright (c) 2017 Microsoft.All rights reserved.


