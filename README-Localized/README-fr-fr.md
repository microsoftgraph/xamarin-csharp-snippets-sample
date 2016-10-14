# <a name="microsoft-graph-sdk-snippets-library-for-xamarin.forms"></a>Bibliothèque d’extraits de code du kit de développement Microsoft Graph pour Xamarin.Forms

##<a name="table-of-contents"></a>Sommaire

* [Conditions préalables](#prerequisites)
* [Enregistrement et configuration de l’application](#register)
* [Création et débogage](#build)
* [Exécution de l’exemple](#run)
* [Impact de l’exemple sur les données de votre compte](#how-the-sample-affects-your-tenant-data)
* [Ajout d’un extrait de code](#add-a-snippet)
* [Questions et commentaires](#questions)
* [Contribution](#contributing")
* [Ressources supplémentaires](#additional-resources)

<a name="introduction"></a> Cet exemple de projet constitue un référentiel des extraits de code qui utilisent Microsoft Graph pour effectuer des tâches courantes, telles que l’envoi des messages électroniques, la gestion des groupes et d’autres activités au sein d’une application Xamarin.Forms. Il utilise le [kit de développement logiciel Microsoft Graph .NET Client](https://github.com/microsoftgraph/msgraph-sdk-dotnet) pour fonctionner avec les données renvoyées par Microsoft Graph. 

L’exemple utilise la [bibliothèque d’authentification Microsoft (MSAL)](https://www.nuget.org/packages/Microsoft.Identity.Client/) pour l’authentification. Le kit de développement logiciel (SDK) MSAL offre des fonctionnalités permettant d’utiliser le [point de terminaison Azure AD v2.0](https://graph.microsoft.io/en-us/docs/authorization/converged_auth), qui permet aux développeurs d’écrire un flux de code unique qui gère l’authentification des comptes professionnels ou scolaires (Azure Active Directory) et personnels (Microsoft), y compris des comptes Office 365, Outlook.com et OneDrive.

> **Remarque** Le kit de développement logiciel MSAL se trouve actuellement dans la version préliminaire et en tant que tel, il ne doit pas être utilisé dans le code de production. Il est utilisé ici à titre indicatif uniquement.

L’application affiche l’interface utilisateur représentant les tâches courantes de l’utilisateur, ou « articles ». Chaque article est constitué d’un ou de plusieurs extraits de code. Les articles sont groupés par niveau d’autorisations et type de compte. L’utilisateur peut se connecter à son compte et exécuter les articles sélectionnés. Chaque article devient vert en cas de succès et rouge en cas d’échec. Des informations supplémentaires sont envoyées vers la fenêtre de sortie.

<a name="prerequisites"></a>
## <a name="prerequisites"></a>Conditions préalables ##

Cet exemple nécessite les éléments suivants :  

  * [Visual Studio 2015](https://www.visualstudio.com/downloads) 
  * [Xamarin pour Visual Studio](https://www.xamarin.com/visual-studio)
  * Windows 10 (avec [mode de développement](https://msdn.microsoft.com/library/windows/apps/xaml/dn706236.aspx))
  * Soit un [compte Microsoft](https://www.outlook.com), soit un [compte Office 365 pour entreprise](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_Office365Account)

Si vous souhaitez exécuter le projet iOS dans cet exemple, vous avez besoin des éléments suivants :

  * Le dernier kit de développement logiciel iOS
  * La dernière version de Xcode
  * Mac OS X Yosemite (10.10) et versions supérieures 
  * [Xamarin.iOS](https://developer.xamarin.com/guides/ios/getting_started/installation/mac/)
  * Un [agent Xamarin Mac connecté à Visual Studio](https://developer.xamarin.com/guides/ios/getting_started/installation/windows/connecting-to-mac/)

Vous pouvez utiliser l’[émulateur Visual Studio pour Android](https://www.visualstudio.com/features/msft-android-emulator-vs.aspx) si vous souhaitez exécuter le projet Android.

<a name="register"></a>
##<a name="register-and-configure-the-app"></a>Enregistrement et configuration de l’application

1. Connectez-vous au [portail d’inscription des applications](https://apps.dev.microsoft.com/) en utilisant votre compte personnel, professionnel ou scolaire.
2. Sélectionnez **Ajouter une application**.
3. Entrez un nom pour l’application, puis sélectionnez **Créer une application**.
    
    La page d’inscription s’affiche, répertoriant les propriétés de votre application.
 
4. Sous **Plateformes**, sélectionnez **Ajouter une plateforme**.
5. Sélectionnez **Application mobile**.
6. Copiez la valeur d’ID client (Id d’application) dans le Presse-papiers. Vous devrez saisir ces valeurs dans l’exemple d’application.

    L’ID d’application est un identificateur unique pour votre application.

7. Cliquez sur **Enregistrer**.

<a name="build"></a>
## <a name="build-and-debug"></a>Création et débogage ##

**Remarque :** si vous constatez des erreurs pendant l’installation des packages à l’étape 2, vérifiez que le chemin d’accès local où vous avez sauvegardé la solution n’est pas trop long/profond. Pour résoudre ce problème, il vous suffit de déplacer la solution dans un dossier plus près du répertoire racine de votre lecteur.

1. Ouvrez le fichier App.cs à l’intérieur du projet **Graph_Xamarin_CS_Snippets (Portable)** de la solution.

    ![Capture d’écran du volet Explorateur de solutions dans Visual Studio, avec fichier App.cs sélectionné dans le projet Graph_Xamarin_CS_Snippets](/readme-images/Appdotcs.png "Fichier App.cs ouvert dans le projet Graph_Xamarin_CS_Snippets")

2. Une fois que vous avez chargé la solution dans Visual Studio, configurez l’exemple pour utiliser l’ID client que vous avez enregistré en l’indiquant comme valeur de la variable **ClientId** dans le fichier App.cs.


    ![Capture d’écran de la variable ClientId dans le fichier App.cs, actuellement définie sur une chaîne vide.](/readme-images/appId.png "Valeur d’ID client dans le fichier App.cs")

2.  Si vous envisagez de vous connecter à l’exemple avec un compte professionnel ou scolaire qui n’a pas les autorisations d’administrateur, vous devrez annuler le code demandant des étendues qui nécessitent des autorisations d’administrateur. Si vous n’annulez pas ces lignes, vous ne serez pas en mesure de vous connecter avec votre compte professionnel ou scolaire (si vous vous connectez avec un compte personnel, ces requêtes d’étendues sont ignorées.)

    Dans la méthode `GetTokenForUserAsync()` du fichier `AuthenticationHelper.cs`, transformez en commentaire les demandes d’étendues suivantes :
    
    ```
        "https://graph.microsoft.com/Directory.AccessAsUser.All",
        "https://graph.microsoft.com/User.ReadWrite.All",
        "https://graph.microsoft.com/Group.ReadWrite.All",
    ```

3. Sélectionnez le projet à exécuter. Si vous sélectionnez l’option Plateforme Windows universelle, vous pouvez exécuter l’exemple sur l’ordinateur local. Si vous souhaitez exécuter le projet iOS, vous devez vous connecter à un [Mac sur lequel les outils de Xamarin](https://developer.xamarin.com/guides/ios/getting_started/installation/windows/connecting-to-mac/) ont été installés. (Vous pouvez également ouvrir cette solution dans Xamarin Studio sur un Mac et exécuter l’exemple directement à partir de là.) Vous pouvez utiliser l’[émulateur Visual Studio pour Android](https://www.visualstudio.com/features/msft-android-emulator-vs.aspx) si vous souhaitez exécuter le projet Android. 

    ![Capture d’écran de la barre d’outils Visual Studio, avec iOS sélectionné comme projet de démarrage.](/readme-images/SelectProject.png "Sélectionnez le projet dans Visual Studio").

4. Appuyez sur F5 pour créer et déboguer l’application. Exécutez la solution et connectez-vous avec votre compte personnel, professionnel ou scolaire.
    > **Remarque** Vous devrez ouvrir le gestionnaire de configurations de build pour vous assurer que les étapes de création et de déploiement sont sélectionnées pour le projet UWP.

<a name="run"></a>
## <a name="run-the-sample"></a>Exécution de l’exemple

Une fois lancée, l’application affiche une liste représentant les tâches courantes de l’utilisateur ou « articles ». Chaque article est constitué d’un ou de plusieurs extraits de code. Les articles sont groupés par niveau d’autorisations et type de compte :

- Tâches qui s’appliquent à la fois aux comptes professionnels, scolaires et personnels, telles que l’obtention et l’envoi de messages électroniques, la création de fichiers, etc.
- Tâches qui s’appliquent uniquement aux comptes professionnels ou scolaires, telles que l’obtention de la photo de compte ou du responsable d’un utilisateur.
- Tâches qui s’appliquent uniquement aux comptes professionnels ou scolaires avec des autorisations administratives appropriées, telles que l’obtention de membres du groupe ou la création de comptes d’utilisateur.

Sélectionnez les articles que vous souhaitez exécuter et cliquez sur le bouton « Exécuter la sélection ». Vous serez invité à vous connecter avec votre compte professionnel, scolaire ou personnel. N’oubliez pas que si vous vous connectez avec un compte qui ne dispose des autorisations applicables pour les articles sélectionnés (par exemple, si vous sélectionnez des articles qui s’appliquent uniquement à un compte professionnel ou scolaire, puis ouvrez une session avec un compte personnel), ces articles échoueront.

Chaque article devient vert en cas de succès et rouge en cas d’échec. Des informations supplémentaires sont envoyées vers la fenêtre de sortie. 

<a name="#how-the-sample-affects-your-tenant-data"></a>
##<a name="how-the-sample-affects-your-account-data"></a>Impact de l’exemple sur les données de votre compte

Cet exemple exécute des commandes qui permettent de créer, lire, mettre à jour ou supprimer des données. Il ne modifie pas et ne supprime pas vos données de compte réelles. Toutefois, il peut créer et laisser des artefacts de données dans votre compte dans le cadre de son fonctionnement : lors de l’exécution des commandes de création, mise à jour ou suppression, l’exemple crée des entités fictives, telles que des utilisateurs ou des groupes, de façon à ne pas affecter vos données de compte réelles. 

L’exemple peut épargner ces entités fictives dans votre compte, si vous choisissez des articles qui créent ou mettent à jour des entités. Par exemple, si vous choisissez d’exécuter l’article « mettre à jour un groupe », un groupe est créé, puis mis à jour. Dans ce cas, le nouveau groupe reste dans votre compte après l’exécution de l’exemple.

<a name="add-a-snippet"></a>
##<a name="add-a-snippet"></a>Ajout d’un extrait de code

Ce projet inclut deux fichiers d’extraits : 

- Groups\GroupSnippets.cs 
- Users\UserSnippets.cs.

Si vous souhaitez exécuter votre propre extrait de code dans ce projet, il suffit de suivre ces trois étapes :

1. **Ajoutez votre extrait au fichier d’extraits.** Veillez à inclure un bloc try/catch. 

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

2. **Créez un article qui utilise votre extrait et ajoutez-le au fichier d’articles associé.** Par exemple, l’article `TryGetMeAsync()` utilise l’extrait `GetMeAsync()` dans le fichier Users\UserStories.cs :

        public static async Task<bool> TryGetMeAsync()
        {
            var currentUser = await UserSnippets.GetMeAsync();

            return currentUser != null;
        }        


Il peut arriver que votre article nécessite l’exécution d’extraits en plus de celui que vous mettez en œuvre. Par exemple, si vous souhaitez mettre à jour un événement, vous devez abord utiliser la méthode `CreateEventAsync()` pour créer un événement. Vous pouvez ensuite le mettre à jour. Veillez à toujours utiliser des extraits de code qui existent déjà dans le fichier d’extraits. Si l’opération dont vous avez besoin n’existe pas, vous devez la créer et l’inclure dans votre article. Il est conseillé de supprimer toutes les entités que vous créez dans un article, en particulier si vous travaillez sur autre chose qu’un compte de test ou de développement.

3. **Ajoutez votre article à la liste d’articles dans MainPage.xaml.cs** (à l’intérieur de la méthode `CreateStoryList()`) :
    
    `snippetList.Children.Add(new CheckBox 
        { StoryName = "Get Me", GroupName = "Users", AccountType = "All", RunStoryAsync = UserStories.TryGetMeAsync });`
    
Vous pouvez à présent tester votre extrait de code. Lorsque vous exécutez l’application, votre article s’affiche sous la forme d’un nouvel élément. Activez la case à cocher de votre extrait de code, puis exécutez-le. Utilisez cette option comme une opportunité pour déboguer votre extrait de code.

<a name="questions"></a>
## <a name="questions-and-comments"></a>Questions et commentaires

Nous serions ravis de connaître votre opinion sur l’exemple d’extraits de code Microsoft Graph pour Xamarin.Forms. Vous pouvez nous faire part de vos questions et suggestions dans la rubrique [Problèmes](https://github.com/MicrosoftGraph/xamarin-csharp-snippets-sample/issues) de ce référentiel.

Votre avis compte beaucoup pour nous. Communiquez avec nous sur [Stack Overflow](http://stackoverflow.com/questions/tagged/office365+or+microsoftgraph). Posez vos questions avec la balise [MicrosoftGraph].

<a name="contributing"></a>
## <a name="contributing"></a>Contribution ##

Si vous souhaitez contribuer à cet exemple, voir [CONTRIBUTING.MD](/CONTRIBUTING.md).

Ce projet a adopté le [code de conduite Microsoft Open Source](https://opensource.microsoft.com/codeofconduct/). Pour plus d’informations, reportez-vous à la [FAQ relative au code de conduite](https://opensource.microsoft.com/codeofconduct/faq/) ou contactez [opencode@microsoft.com](mailto:opencode@microsoft.com) pour toute question ou tout commentaire.


<a name="additional-resources"></a>
## <a name="additional-resources"></a>Ressources supplémentaires ##

- [Autres exemples de connexion avec Microsoft Graph](https://github.com/MicrosoftGraph?utf8=%E2%9C%93&query=-Connect)
- [Présentation de Microsoft Graph](http://graph.microsoft.io)
- [Exemples de code du développeur Office](http://dev.office.com/code-samples)
- [Centre de développement Office](http://dev.office.com/)


## <a name="copyright"></a>Copyright
Copyright (c) 2016 Microsoft. Tous droits réservés.


