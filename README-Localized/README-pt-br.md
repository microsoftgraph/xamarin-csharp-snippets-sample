# <a name="microsoft-graph-sdk-snippets-library-for-xamarin.forms"></a>Biblioteca de Trechos de Código do SDK do Microsoft Graph para Xamarin.Forms

##<a name="table-of-contents"></a>Sumário

* [Pré-requisitos](#prerequisites)
* [Registrar e configurar o aplicativo](#register)
* [Compilar e depurar](#build)
* [Executar o exemplo](#run)
* [De que maneira o exemplo afeta os dados da conta](#how-the-sample-affects-your-tenant-data)
* [Adicionar um trecho de código](#add-a-snippet)
* [Perguntas e comentários](#questions)
* [Colaboração](#contributing")
* [Recursos adicionais](#additional-resources)

<a name="introduction"></a> Este exemplo de projeto fornece um repositório de trechos de código que usa o Microsoft Graph para realizar tarefas comuns, como envio de emails, gerenciamento de grupos e outras atividades diretamente de um aplicativo Xamarin.Forms. O exemplo usa o [SDK do Cliente .NET para Microsoft Graph](https://github.com/microsoftgraph/msgraph-sdk-dotnet) para trabalhar com dados retornados pelo Microsoft Graph. 

O exemplo usa a [Biblioteca de Autenticação da Microsoft (MSAL)](https://www.nuget.org/packages/Microsoft.Identity.Client/) para autenticação. O SDK da MSAL fornece recursos para trabalhar com o [ponto de extremidade do Microsoft Azure AD versão 2.0](https://graph.microsoft.io/en-us/docs/authorization/converged_auth), que permite aos desenvolvedores gravar um único fluxo de código para tratar da autenticação de contas pessoais (Microsoft), corporativas ou de estudantes (Azure Active Directory), inclusive contas das plataformas Office 365, Outlook.com e OneDrive.

> **Observação** No momento, o SDK da MSAL encontra-se em pré-lançamento e como tal não deve ser usado no código de produção. Isso é usado apenas para fins ilustrativos

O aplicativo exibe uma interface de usuário que representa tarefas comuns do usuário ou “histórias”. Cada história é composta de um ou mais trechos de código. As histórias são agrupadas por tipo de conta e nível de permissão. O usuário pode fazer logon em sua conta e executar as histórias selecionadas. As histórias ficam verdes se tiverem êxito e vermelhas se falharem. Informações adicionais são enviadas para a janela de Saída.

<a name="prerequisites"></a>
## <a name="prerequisites"></a>Pré-requisitos ##

Este exemplo requer o seguinte:  

  * [Visual Studio 2015](https://www.visualstudio.com/downloads) 
  * [Xamarin para Visual Studio](https://www.xamarin.com/visual-studio)
  * Windows 10 ([modo de desenvolvedor habilitado](https://msdn.microsoft.com/library/windows/apps/xaml/dn706236.aspx))
  * Uma [conta da Microsoft](https://www.outlook.com) ou a [conta do Office 365 para empresas](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_Office365Account)

Se quiser executar o projeto do iOS neste exemplo, você precisará do seguinte:

  * O SDK mais recente do iOS
  * A versão mais recente do Xcode
  * Mac OS X Yosemite (10.10) e versões superiores 
  * [Xamarin.iOS](https://developer.xamarin.com/guides/ios/getting_started/installation/mac/)
  * Um [agente do Xamarin Mac conectado ao Visual Studio](https://developer.xamarin.com/guides/ios/getting_started/installation/windows/connecting-to-mac/)

Você pode usar o [Emulador do Visual Studio para Android](https://www.visualstudio.com/features/msft-android-emulator-vs.aspx) se quiser executar o projeto do Android.

<a name="register"></a>
##<a name="register-and-configure-the-app"></a>Registrar e configurar o aplicativo

1. Entre no [Portal de Registro do Aplicativo](https://apps.dev.microsoft.com/) usando sua conta pessoal ou sua conta comercial ou escolar.
2. Selecione **Adicionar um aplicativo**.
3. Insira um nome para o aplicativo e selecione **Criar aplicativo**.
    
    A página de registro é exibida, listando as propriedades do seu aplicativo.
 
4. Em **Plataformas**, selecione **Adicionar plataforma**.
5. Escolha **Aplicativo móvel**.
6. Copie o valor da ID de Cliente (ID de Aplicativo) para a área de transferência. Você precisará inserir esses valores no exemplo de aplicativo.

    Essa ID de aplicativo é o identificador exclusivo do aplicativo.

7. Selecione **Salvar**.

<a name="build"></a>
## <a name="build-and-debug"></a>Compilar e depurar ##

**Observação:** Caso receba mensagens de erro durante a instalação de pacotes na etapa 2, verifique se o caminho para o local onde você colocou a solução não é muito longo ou extenso. Para resolver esse problema, coloque a solução junto à raiz da unidade.

1. Abra o arquivo App.cs no projeto **Graph_Xamarin_CS_Snippets (Portátil)** da solução.

    ![Captura de tela do painel do Gerenciador de Soluções no Visual Studio, com o arquivo App.cs selecionado no projeto Graph_Xamarin_CS_Snippets](/readme-images/Appdotcs.png "Abrir o arquivo App.cs no projeto Graph_Xamarin_CS_Snippets")

2. Após carregar a solução no Visual Studio, configure o exemplo para usar a ID do cliente registrada transformando-a no valor da variável **ClientId** variável no arquivo App.cs.


    ![Captura de tela da variável ClientId no arquivo App.cs, atualmente definida como uma cadeia de caracteres vazia.](/readme-images/appId.png "Valor da ID do Cliente no arquivo App.cs")

2.  Se você estiver planejando entrar no exemplo com uma conta comercial ou escolar que não tenha permissões de administrador, será preciso comentar o código que solicita escopos que exigem permissões de administrador. Se não comentar essas linhas, você não será capaz de entrar com sua conta comercial ou escolar (se você entrar com uma conta pessoal, essas solicitações de escopo serão ignoradas.)

    No método `GetTokenForUserAsync()` do arquivo `AuthenticationHelper.cs`, fale sobre as seguintes solicitações de escopo:
    
    ```
        "https://graph.microsoft.com/Directory.AccessAsUser.All",
        "https://graph.microsoft.com/User.ReadWrite.All",
        "https://graph.microsoft.com/Group.ReadWrite.All",
    ```

3. Escolha o projeto que você deseja excluir. Se escolher a opção Plataforma Universal do Windows, você poderá executar o exemplo no computador local. Se quiser executar o projeto do iOS, você precisará se conectar a um [Mac que tenha as ferramentas Xamarin](https://developer.xamarin.com/guides/ios/getting_started/installation/windows/connecting-to-mac/) instaladas nele. (Você também pode abrir esta solução no Xamarin Studio em um Mac e executar o exemplo diretamente de lá). Você pode usar o [Emulador do Microsoft Visual Studio para Android](https://www.visualstudio.com/features/msft-android-emulator-vs.aspx), caso pretenda executar o projeto do Android. 

    ![Captura de tela da barra de ferramentas do Visual Studio, com o iOS selecionado como o projeto de inicialização.](/readme-images/SelectProject.png "Selecionar o projeto no Visual Studio")

4. Pressione F5 para criar e depurar. Execute a solução e entre com sua conta pessoal ou sua conta comercial ou escolar.
    > **Observação** Talvez seja necessário abrir o Gerenciador de Configuração de Compilação para certificar-se de que as etapas de Compilar e Implantar estejam selecionadas para o projeto do UWP.

<a name="run"></a>
## <a name="run-the-sample"></a>Executar o exemplo

Quando iniciado, o aplicativo exibe uma lista que representam tarefas comuns do usuário ou “histórias”. Cada história é composta de um ou mais trechos de código. As histórias são agrupadas por tipo de conta e nível de permissão:

- Tarefas que são aplicáveis a contas comerciais ou escolares e contas pessoais, como receber e enviar emails, criar arquivos, etc.
- Tarefas que só são aplicáveis a contas comerciais ou escolares, como receber a foto da conta e o gerenciador do usuário.
- Tarefas que só são aplicáveis a contas comerciais ou escolares com permissões administrativas, como receber membros do grupo ou criar novas contas de usuário.

Escolha as histórias que você deseja executar e escolha o botão “executar selecionadas”. Você será solicitado a fazer logon em sua conta comercial ou escolar ou conta pessoal. Lembre-se de que se fizer logon com uma conta que não tenha permissões aplicáveis para as histórias selecionadas (por exemplo, se você escolher histórias que se aplicam apenas a uma conta comercial ou escolar e, em seguida, fizer logon com uma conta pessoal), essas histórias falharão.

As histórias ficam verdes se tiverem êxito e vermelhas se falharem. Informações adicionais são enviadas para a janela de Saída. 

<a name="#how-the-sample-affects-your-tenant-data"></a>
##<a name="how-the-sample-affects-your-account-data"></a>Como o exemplo afeta os dados da conta

Este exemplo executa comandos que criam, leem, atualizam ou excluem dados. Ele não editará ou excluirá os dados reais da conta. No entanto, o exemplo pode criar e deixar artefatos de dados em sua conta como parte da operação: durante a execução de comandos que criam, atualizam ou excluem, o exemplo cria entidades falsas, como novos usuários ou grupos, para não afetar os dados reais da conta. 

O exemplo pode deixar tais entidades falsas em sua conta, se você escolher histórias que criam ou atualizam entidades. Por exemplo, ao escolher a história de execução ‘atualizar grupo’ cria um novo grupo e, em seguida, atualiza-o. Nesse caso, o novo grupo permanece na sua conta após o exemplo ser executado.

<a name="add-a-snippet"></a>
##<a name="add-a-snippet"></a>Adicionar um trecho de código

Este projeto inclui dois arquivos de trechos de código: 

- Groups\GroupSnippets.cs 
- Users\UserSnippets.cs.

Se você tem seu próprio trecho de código e quer executá-lo neste projeto, basta seguir estas três etapas:

1. **Adicione seu trecho de código ao arquivo de trechos de código.** Certifique-se de incluir um bloco try/catch. 

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

2. **Crie uma história que usa o trecho de código e adicione-a ao arquivo de histórias associado.** Por exemplo, a história `TryGetMeAsync()` usa o trecho de código `GetMeAsync()` no arquivo Users\UserStories.cs:

        public static async Task<bool> TryGetMeAsync()
        {
            var currentUser = await UserSnippets.GetMeAsync();

            return currentUser != null;
        }        


Às vezes, sua história precisará executar trechos de código além do que você está implementando. Por exemplo, se você quiser atualizar um evento, primeiro é necessário usar o método `CreateEventAsync()` para criar um evento. Em seguida, você já poderá atualizá-lo. Sempre use trechos de código que já existam no arquivo de trechos de código. Se a operação que necessária não existir, você terá que criá-la e, em seguida, inclui-la em sua história. Excluir entidades criadas em uma história é uma prática recomendada, especialmente se você estiver trabalhando em uma conta que não seja uma conta de teste ou de desenvolvedor.

3. **Adicione sua história à lista de histórias em MainPage.xaml.cs** (no método `CreateStoryList()`):
    
    `snippetList.Children.Add(new CheckBox 
        { StoryName = "Get Me", GroupName = "Users", AccountType = "All", RunStoryAsync = UserStories.TryGetMeAsync });`
    
Você já pode testar seu trecho de código. Ao executar o aplicativo, sua história aparecerá como um novo item. Escolha a caixa de seleção do trecho de código e execute-o. Use isso como uma oportunidade para depurar seu trecho de código.

<a name="questions"></a>
## <a name="questions-and-comments"></a>Perguntas e comentários

Adoraríamos receber seus comentários sobre o Exemplo de Trechos de Código do Microsoft Graph para o projeto Xamarin.Forms. Você pode nos enviar suas perguntas e sugestões por meio da seção [Issues](https://github.com/MicrosoftGraph/xamarin-csharp-snippets-sample/issues) deste repositório.

Seus comentários são importantes para nós. Junte-se a nós na página [Stack Overflow](http://stackoverflow.com/questions/tagged/office365+or+microsoftgraph). Marque suas perguntas com [MicrosoftGraph].

<a name="contributing"></a>
## <a name="contributing"></a>Colaboração ##

Se quiser contribuir para esse exemplo, confira [CONTRIBUTING.MD](/CONTRIBUTING.md).

Este projeto adotou o [Código de Conduta do Código Aberto da Microsoft](https://opensource.microsoft.com/codeofconduct/). Para saber mais, confira as [Perguntas frequentes do Código de Conduta](https://opensource.microsoft.com/codeofconduct/faq/) ou contate [opencode@microsoft.com](mailto:opencode@microsoft.com) se tiver outras dúvidas ou comentários.


<a name="additional-resources"></a>
## <a name="additional-resources"></a>Recursos adicionais ##

- [Outros exemplos de conexão usando o Microsoft Graph](https://github.com/MicrosoftGraph?utf8=%E2%9C%93&query=-Connect)
- [Visão geral do Microsoft Graph](http://graph.microsoft.io)
- [Exemplos de código para desenvolvedores do Office](http://dev.office.com/code-samples)
- [Centro de Desenvolvimento do Office](http://dev.office.com/)


## <a name="copyright"></a>Direitos autorais
Copyright © 2016 Microsoft. Todos os direitos reservados.


