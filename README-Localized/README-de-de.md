# <a name="microsoft-graph-sdk-snippets-library-for-xamarin.forms"></a>Codeausschnittsbibliothek des Microsoft Graph-SDKs für Xamarin.Forms

##<a name="table-of-contents"></a>Inhalt

* [Voraussetzungen](#prerequisites)
* [Registrieren und Konfigurieren der App](#register)
* [Erstellen und Debuggen](#build)
* [Ausführen des Beispiels](#run)
* [Wie sich das Beispiel auf Ihre Kontodaten auswirkt](#how-the-sample-affects-your-tenant-data)
* [Hinzufügen eines Codeausschnitts](#add-a-snippet)
* [Fragen und Kommentare](#questions)
* [Mitwirkung](#contributing")
* [Weitere Ressourcen](#additional-resources)

<a name="introduction"></a> Dieses Beispielprojekt enthält ein Repository von Codeausschnitten, die Microsoft Graph verwenden, um allgemeine Aufgaben, z. B. das Senden von E-Mails, das Verwalten von Gruppen und andere Aktivitäten, aus einer Xamarin.Forms-App heraus auszuführen. Es verwendet das [Microsoft Graph .NET-Client-SDK](https://github.com/microsoftgraph/msgraph-sdk-dotnet), um mit Daten zu arbeiten, die von Microsoft Graph zurückgegeben werden. 

Das Beispiel verwendet die [Microsoft-Authentifizierungsbibliothek (MSAL)](https://www.nuget.org/packages/Microsoft.Identity.Client/) für die Authentifizierung. Das MSAL-SDK bietet Features für die Arbeit mit dem [Azure AD v2.0-Endpunkt](https://graph.microsoft.io/en-us/docs/authorization/converged_auth), der es Entwicklern ermöglicht, einen einzelnen Codefluss zu schreiben, der die Authentifizierung sowohl für Geschäfts- oder Schulkonten von Benutzern (Azure Active Directory) als auch für persönliche Konten (Microsoft) verarbeitet, einschließlich Office 365-, Outlook.com- und OneDrive-Konten.

> **Hinweis** Das MSAL-SDK befindet sich derzeit in der Vorabversion und sollte daher nicht in Produktionscode verwendet werden. Es dient hier nur zur Veranschaulichung

In der App wird die Benutzeroberfläche angezeigt, die allgemeine Aufgaben bzw. Stories darstellen. Jede Story besteht aus einem oder mehreren Codeausschnitten. Die Stories werden nach Kontotyp und Berechtigungsstufe gruppiert. Der Benutzer kann sich bei seinem Konto anmelden und die ausgewählten Stories ausführen. Jede Story wird grün angezeigt, wenn sie erfolgreich ist, und rot, wenn sie fehlschlägt. Zusätzliche Informationen werden an das Ausgabefenster gesendet.

<a name="prerequisites"></a>
## <a name="prerequisites"></a>Anforderungen ##

Für dieses Beispiel ist Folgendes erforderlich:  

  * [Visual Studio 2015](https://www.visualstudio.com/downloads) 
  * [Xamarin für Visual Studio](https://www.xamarin.com/visual-studio)
  * Windows 10 (mit [aktiviertem Entwicklungsmodus](https://msdn.microsoft.com/library/windows/apps/xaml/dn706236.aspx))
  * Entweder ein [Microsoft-Konto](https://www.outlook.com) oder ein [Office 365 for Business-Konto](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_Office365Account).

Wenn Sie das iOS-Projekt in diesem Beispiel ausführen möchten, benötigen Sie Folgendes:

  * Das neueste iOS-SDK
  * Die neueste Version von Xcode
  * Mac OS X Yosemite (10.10) und höher 
  * [Xamarin.iOS](https://developer.xamarin.com/guides/ios/getting_started/installation/mac/)
  * Einen mit [Visual Studio verbundenen Xamarin Mac-Agent](https://developer.xamarin.com/guides/ios/getting_started/installation/windows/connecting-to-mac/)

Sie können den [Visual Studio-Emulator für Android](https://www.visualstudio.com/features/msft-android-emulator-vs.aspx) verwenden, wenn Sie das Android-Projekt ausführen möchten.

<a name="register"></a>
##<a name="register-and-configure-the-app"></a>Registrieren und Konfigurieren der App

1. Melden Sie sich beim [App-Registrierungsportal](https://apps.dev.microsoft.com/) entweder mit Ihrem persönlichen oder geschäftlichen Konto oder mit Ihrem Schulkonto an.
2. Klicken Sie auf **App hinzufügen**.
3. Geben Sie einen Namen für die App ein, und wählen Sie **Anwendung erstellen** aus.
    
    Die Registrierungsseite wird angezeigt, und die Eigenschaften der App werden aufgeführt.
 
4. Wählen Sie unter **Plattformen** die Option **Plattform hinzufügen** aus.
5. Wählen Sie **Mobile Anwendung** aus.
6. Kopieren Sie den Wert für die Client-ID (App-ID) in die Zwischenablage. Sie müssen diese Werte in die Beispiel-App eingeben.

    Die App-ID ist ein eindeutiger Bezeichner für Ihre App.

7. Klicken Sie auf **Speichern**.

<a name="build"></a>
## <a name="build-and-debug"></a>Erstellen und Debuggen ##

**Hinweis:** Wenn beim Installieren der Pakete während des Schritts 2 Fehler angezeigt werden, müssen Sie sicherstellen, dass der lokale Pfad, unter dem Sie die Projektmappe abgelegt haben, weder zu lang noch zu tief ist. Dieses Problem lässt sich beheben, indem Sie den Pfad auf Ihrem Laufwerk verkürzen.

1. Öffnen Sie die Datei „App.cs“ innerhalb des **Graph_Xamarin_CS_Snippets (Portable)**-Projekts der Lösung.

    ![Screenshot des Bereichs „Projektmappen-Explorer“ in Visual Studio, wobei die Datei „App.cs“ im Projekt „Graph_Xamarin_CS_Snippets“ ausgewählt ist](/readme-images/Appdotcs.png "Öffnen Sie die Datei „App.cs“ im Graph_Xamarin_CS_Snippets-Projekt")

2. Nachdem Sie die Lösung in Visual Studio geladen haben, konfigurieren Sie das Beispiel so, dass die registrierte Client-ID verwendet wird, indem Sie diese als Wert der **ClientId**-Variablen in der Datei „App.cs“ zuweisen.


    ![Screenshot der ClientId-Variablen in der Datei“ App.cs“, die derzeit auf eine leere Zeichenfolge festgelegt ist.](/readme-images/appId.png "Client-ID-Wert in App.cs-Datei")

2.  Wenn Sie sich bei dem Beispiel mit einem Geschäfts- oder Schulkonto anmelden möchten, das nicht über Administratorberechtigungen verfügt, müssen Sie Code auskommentieren, der Bereiche anfordert, für die Administratorberechtigungen erforderlich sind. Wenn Sie diese Zeilen nicht auskommentieren, können Sie sich nicht mit Ihrem Geschäfts- oder Schulkonto anmelden (wenn Sie sich mit einem persönlichen Konto anmelden, werden die Bereichsanforderungen ignoriert).

    Kommentieren Sie in der `GetTokenForUserAsync()`-Methode der `AuthenticationHelper.cs`-Datei die folgenden Bereichsanforderungen aus:
    
    ```
        "https://graph.microsoft.com/Directory.AccessAsUser.All",
        "https://graph.microsoft.com/User.ReadWrite.All",
        "https://graph.microsoft.com/Group.ReadWrite.All",
    ```

3. Wählen Sie das auszuführende Projekt aus. Wenn Sie die Option für die universelle Windows-Plattform auswählen, können Sie das Beispiel auf dem lokalen Computer ausführen. Wenn Sie das iOS-Projekt ausführen möchten, müssen Sie eine Verbindung zu einem [Mac herstellen, auf dem die Xamarin-Tools installiert](https://developer.xamarin.com/guides/ios/getting_started/installation/windows/connecting-to-mac/) sind. (Sie können diese Lösung auch in Xamarin Studio auf einem Mac öffnen und das Beispiel direkt dort ausführen.) Sie können den [Visual Studio-Emulator für Android](https://www.visualstudio.com/features/msft-android-emulator-vs.aspx) verwenden, wenn Sie das Android-Projekt ausführen möchten. 

    ![Screenshot der Visual Studio-Symbolleiste, wobei iOS als Startprojekt ausgewählt ist.](/readme-images/SelectProject.png "Projekt auswählen in Visual Studio")

4. Drücken Sie zum Erstellen und Debuggen F5. Führen Sie die Lösung aus, und melden Sie sich entweder mit Ihrem persönlichen Konto oder mit Ihrem Geschäfts- oder Schulkonto an.
    > **Hinweis** Möglicherweise müssen Sie den Buildkonfigurations-Manager öffnen, um sicherzustellen, dass die Build- und Bereitstellungsschritte für das UWP-Projekt ausgewählt sind.

<a name="run"></a>
## <a name="run-the-sample"></a>Ausführen des Beispiels

Nach dem Start wird in der App eine Liste angezeigt, die allgemeine Benutzeraufgaben bzw. so genannte Stories darstellen. Jede Story besteht aus einem oder mehreren Codeausschnitten. Die Stories werden nach Kontotyp und Berechtigungsstufe gruppiert:

- Aufgaben, die sowohl für Geschäfts- oder Schulkonten als auch für persönliche Konten gelten, z. B. das Abrufen und Senden von E-Mails, das Erstellen von Dateien usw.
- Aufgaben, die nur für Geschäfts- oder Schulkonten gelten, z. B. das Abrufen eines Vorgesetzten eines Benutzers oder eines Kontofotos.
- Aufgaben, die nur für Geschäfts- oder Schulkonten mit Administratorberechtigungen gelten, z. B. das Abrufen von Gruppenmitgliedern oder das Erstellen neuer Benutzerkonten.

Wählen Sie die Stories aus, die Sie ausführen möchten, und klicken Sie dann auf die Schaltfläche, um die Auswahl auszuführen. Sie werden aufgefordert, sich mit Ihrem Geschäfts- oder Schulkonto oder mit Ihrem persönlichen Konto anzumelden. Wenn Sie sich mit einem Konto anmelden, das für die ausgewählten Stories nicht über die jeweiligen Berechtigungen verfügt (wenn Sie zum Beispiel Stories auswählen, die nur für ein Geschäfts- oder Schulkonto gelten und sich dann mit einem persönlichen Konto anmelden), tritt bei diesen Stories ein Fehler auf.

Jede Story wird grün angezeigt, wenn sie erfolgreich ist, und rot, wenn sie fehlschlägt. Zusätzliche Informationen werden an das Ausgabefenster gesendet. 

<a name="#how-the-sample-affects-your-tenant-data"></a>
##<a name="how-the-sample-affects-your-account-data"></a>Wie sich das Beispiel auf Ihre Kontodaten auswirkt

In diesem Beispiel werden Befehle ausgeführt, mit denen Daten erstellt, aktualisiert oder gelöscht werden. Ihre tatsächlichen Kontodaten werden nicht bearbeitet oder gelöscht. Im Rahmen der Ausführung des Beispiels werden in Ihrem Konto jedoch möglicherweise Datenartefakte erstellt und hinterlassen: Beim Ausführen von Befehlen, die Daten erstellen, aktualisieren oder löschen, erstellt das Beispiel gefälschte Entitäten, z. B. als neue Benutzer oder Gruppen, damit Ihre tatsächlichen Kontodaten nicht beeinträchtigt werden. 

Das Beispiel kann solche gefälschten Entitäten in Ihrem Konto hinterlassen, wenn Sie Stories auswählen, die Entitäten erstellen oder aktualisieren. Angenommen, durch Auswählen der Story „Gruppe aktualisieren“ wird eine neue Gruppe erstellt, die dann aktualisiert wird. In diesem Fall bleibt die neue Gruppe in Ihrem Konto, nachdem das Beispiel ausgeführt wurde.

<a name="add-a-snippet"></a>
##<a name="add-a-snippet"></a>Hinzufügen eines Codeausschnitts

Dieses Projekt umfasst zwei Codeausschnittdateien: 

- Groups\GroupSnippets.cs 
- Users\UserSnippets.cs.

Wenn Sie über einen eigenen Codeausschnitt verfügen, den Sie in diesem Projekt ausführen möchten, führen Sie die folgenden drei Schritte aus:

1. **Fügen Sie den Codeausschnitt der Codeausschnittdatei hinzu.** Achten Sie darauf, einen Try-Catch-Block einzufügen. 

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

2. **Erstellen Sie eine Story, die Ihren Codeausschnitt verwendet, und fügen Sie diese zu der zugeordneten Storydatei hinzu.** Die Story `TryGetMeAsync()` verwendet beispielsweise den Codeausschnitt `GetMeAsync()` in der Datei „Users\UserStories.cs“:

        public static async Task<bool> TryGetMeAsync()
        {
            var currentUser = await UserSnippets.GetMeAsync();

            return currentUser != null;
        }        


Manchmal muss die Story zusätzlich zu dem von Ihnen implementierten Codeausschnitt weitere Codeausschnitte ausführen. Wenn Sie zum Beispiel ein Ereignis aktualisieren möchten, müssen Sie zuerst die `CreateEventAsync()`-Methode verwenden, um ein Ereignis zu erstellen. Dann können Sie die Aktualisierung vornehmen. Achten Sie unbedingt darauf, Codeausschnitte zu verwenden, die bereits in der Codeausschnittdatei vorhanden sind. Wenn der von Ihnen benötigte Vorgang nicht vorhanden ist, müssen Sie ihn erstellen und dann in die Story einfügen. Es wird empfohlen, dass Sie alle in einer Story erstellten Entitäten löschen, insbesondere, wenn Sie nicht unter einem Test- oder Entwicklerkonto arbeiten.

3. **Fügen Sie Ihre Story der Storyliste in „MainPage.xaml.cs“ hinzu** (innerhalb der `CreateStoryList()`-Methode):
    
    `snippetList.Children.Add(new CheckBox 
        { StoryName = "Get Me", GroupName = "Users", AccountType = "All", RunStoryAsync = UserStories.TryGetMeAsync });`
    
Jetzt können Sie Ihren Codeausschnitt testen. Wenn Sie die App ausführen, wird Ihre Story als neues Element angezeigt. Aktivieren Sie das Kontrollkästchen für Ihren Codeausschnitt, und führen Sie ihn dann aus. Nutzen Sie diese Gelegenheit, um Ihren Codeausschnitt zu debuggen.

<a name="questions"></a>
## <a name="questions-and-comments"></a>Fragen und Kommentare

Wir schätzen Ihr Feedback hinsichtlich des Microsoft Graph-Codeausschnittbeispiels für das Xamarin Forms-Projekt. Sie können uns Ihre Fragen und Vorschläge über den Abschnitt [Probleme](https://github.com/MicrosoftGraph/xamarin-csharp-snippets-sample/issues) dieses Repositorys senden.

Ihr Feedback ist uns wichtig. Nehmen Sie unter [Stack Overflow](http://stackoverflow.com/questions/tagged/office365+or+microsoftgraph) Kontakt mit uns auf. Taggen Sie Ihre Fragen mit [MicrosoftGraph].

<a name="contributing"></a>
## <a name="contributing"></a>Mitwirkung ##

Wenn Sie einen Beitrag zu diesem Beispiel leisten möchten, finden Sie unter [CONTRIBUTING.MD](/CONTRIBUTING.md) weitere Informationen.

In diesem Projekt wurden die [Microsoft Open Source-Verhaltensregeln](https://opensource.microsoft.com/codeofconduct/) übernommen. Weitere Informationen finden Sie unter [Häufig gestellte Fragen zu Verhaltensregeln](https://opensource.microsoft.com/codeofconduct/faq/), oder richten Sie Ihre Fragen oder Kommentare an [opencode@microsoft.com](mailto:opencode@microsoft.com).


<a name="additional-resources"></a>
## <a name="additional-resources"></a>Zusätzliche Ressourcen ##

- [Weitere Microsoft Graph Connect-Beispiele](https://github.com/MicrosoftGraph?utf8=%E2%9C%93&query=-Connect)
- [Microsoft Graph-Übersicht](http://graph.microsoft.io)
- [Office-Entwicklercodebeispiele](http://dev.office.com/code-samples)
- [Office Dev Center](http://dev.office.com/)


## <a name="copyright"></a>Copyright
Copyright (c) 2016 Microsoft. Alle Rechte vorbehalten.


