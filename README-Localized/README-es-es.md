# Biblioteca de fragmentos de código de SDK de Microsoft Graph para Xamarin.Forms
<a id="microsoft-graph-sdk-snippets-library-for-xamarinforms" class="xliff"></a>

## Tabla de contenido
<a id="table-of-contents" class="xliff"></a>

* [Requisitos previos](#prerequisites)
* [Registrar y configurar la aplicación](#register)
* [Compilar y depurar](#build)
* [Ejecutar el ejemplo](#run)
* [Repercusión de la muestra en los datos en su cuenta](#how-the-sample-affects-your-tenant-data)
* [Agregar un fragmento de código](#add-a-snippet)
* [Preguntas y comentarios](#questions)
* [Colaboradores](#contributing")
* [Recursos adicionales](#additional-resources)

<a name="introduction"></a> Este proyecto de ejemplo proporciona un repositorio de fragmentos de código que usa Microsoft Graph para realizar tareas comunes, como enviar correos electrónicos, administrar grupos y otras actividades desde una aplicación de Xamarin.Forms. Usa el [SDK del cliente de Microsoft Graph .NET](https://github.com/microsoftgraph/msgraph-sdk-dotnet) para trabajar con los datos devueltos por Microsoft Graph. 

El ejemplo usa la [biblioteca de autenticación de Microsoft (MSAL)](https://www.nuget.org/packages/Microsoft.Identity.Client/) para la autenticación. El SDK de MSAL ofrece características para trabajar con el [punto de conexión v2.0 de Azure AD](https://graph.microsoft.io/en-us/docs/authorization/converged_auth), lo que permite a los desarrolladores escribir un flujo de código único que controla la autenticación para las cuentas profesionales o educativas (Azure Active Directory) o las cuentas personales de (Microsoft), incluidas las cuentas de Office 365, Outlook.com y OneDrive de los usuarios.

## Nota importante acerca de MSAL Preview
<a id="important-note-about-the-msal-preview" class="xliff"></a>

Esta biblioteca es apta para utilizarla en un entorno de producción. Ofrecemos la misma compatibilidad de nivel de producción de esta biblioteca que la de las bibliotecas de producción actual. Durante la vista previa podemos realizar cambios en la API, el formato de caché interna y otros mecanismos de esta biblioteca, que deberá tomar junto con correcciones o mejoras. Esto puede afectar a la aplicación. Por ejemplo, un cambio en el formato de caché puede afectar a los usuarios, como que se les pida que vuelvan a iniciar sesión. Un cambio de API puede requerir que actualice su código. Cuando ofrecemos la versión de disponibilidad General, deberá actualizar a la versión de disponibilidad General dentro de seis meses, ya que las aplicaciones escritas mediante una versión de vista previa de biblioteca puede que ya no funcionen.

La aplicación muestra la interfaz de usuario que representa tareas o "casos" comunes del usuario. Cada caso está formado por uno o más fragmentos de código. Los casos se agrupan por el nivel de permiso y el tipo de cuenta. El usuario puede iniciar sesión en su cuenta y ejecutar las historias seleccionadas. Cada caso se vuelve verde si se acepta y rojo si se produce un error. Se envía información adicional a la ventana de salida.

<a name="prerequisites"></a>
## Requisitos previos
<a id="prerequisites" class="xliff"></a> ##

Este ejemplo necesita lo siguiente:  

  * [Visual Studio 2015](https://www.visualstudio.com/downloads) 
  * [Xamarin para Visual Studio](https://www.xamarin.com/visual-studio)
  * Windows 10 ([modo de desarrollo habilitado](https://msdn.microsoft.com/library/windows/apps/xaml/dn706236.aspx))
  * Una [cuenta de Microsoft](https://www.outlook.com) o bien de [Office 365 para empresas](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_Office365Account)

Si desea ejecutar el proyecto de iOS en este ejemplo, necesita lo siguiente:

  * El SDK de iOS más reciente
  * La versión de Xcode más reciente
  * Mac OS X Yosemite (10.10) o superior 
  * [Xamarin.iOS](https://developer.xamarin.com/guides/ios/getting_started/installation/mac/)
  * Un [agente Xamarin Mac conectado a Visual Studio](https://developer.xamarin.com/guides/ios/getting_started/installation/windows/connecting-to-mac/)

Puede usar el [emulador de Visual Studio para Android](https://www.visualstudio.com/features/msft-android-emulator-vs.aspx) si desea ejecutar el proyecto Android.

<a name="register"></a>
## Registrar y configurar la aplicación
<a id="register-and-configure-the-app" class="xliff"></a>

1. Inicie sesión en el [Portal de registro de la aplicación](https://apps.dev.microsoft.com/) mediante su cuenta personal, profesional o educativa.
2. Seleccione **Agregar una aplicación**.
3. Escriba un nombre para la aplicación y seleccione **Crear aplicación**.
    
    Se muestra la página de registro, indicando las propiedades de la aplicación.
 
4. En **Plataformas**, seleccione **Agregar plataforma**.
5. Seleccione **Aplicación móvil**.
6. Copie el valor del Id. de cliente (Id. de la aplicación) en el portapapeles. Deberá especificar estos valores en la aplicación de ejemplo.

    El id. de la aplicación es un identificador único para su aplicación.

7. Seleccione **Guardar**.

<a name="build"></a>
## Compilar y depurar
<a id="build-and-debug" class="xliff"></a> ##

**Nota:** Si observa algún error durante la instalación de los paquetes en el paso 2, asegúrese de que la ruta de acceso local donde colocó la solución no es demasiado larga o profunda. Para resolver este problema, mueva la solución más cerca de la raíz de la unidad.

1. Abra el archivo App.cs en el proyecto **Graph_Xamarin_CS_Snippets (portátil)** de la solución.

    ![Captura de pantalla del panel Explorador de soluciones en Visual Studio, con el archivo App.cs seleccionado en el proyecto Graph_Xamarin_CS_Snippets](/readme-images/Appdotcs.png "Abra el archivo App.cs en el proyecto Graph_Xamarin_CS_Snippets")

2. Después de cargar la solución en Visual Studio, configure la muestra para usar el Id. de cliente que registró convirtiendo este valor en el de la variable **ClientId** en el archivo App.cs.


    ![Captura de pantalla de la variable ClientId en el archivo App.cs, actualmente establecida en una cadena vacía.](/readme-images/appId.png "Valor de id. de cliente en el archivo App.cs.")

2.  Si va a iniciar la sesión en el ejemplo con una cuenta profesional o educativa que no tiene permisos de administrador, deberá comentar el código que solicita los ámbitos que requieren permisos de administrador. Si no realiza comentarios en estas líneas, no podrá iniciar sesión con su cuenta profesional o educativa (si se inicia sesión con una cuenta personal, estas solicitudes de ámbito se ignoran).

    En el método `GetTokenForUserAsync()` del archivo `AuthenticationHelper.cs`, comente las siguientes solicitudes de ámbito:
    
    ```
        "https://graph.microsoft.com/Directory.AccessAsUser.All",
        "https://graph.microsoft.com/User.ReadWrite.All",
        "https://graph.microsoft.com/Group.ReadWrite.All",
    ```

3. Seleccione el proyecto que desee ejecutar. Si selecciona la opción de plataforma universal de Windows, puede ejecutar el ejemplo en el equipo local. Si desea ejecutar el proyecto iOS, necesitará conectarse a un [Mac que tenga las herramientas Xamarin](https://developer.xamarin.com/guides/ios/getting_started/installation/windows/connecting-to-mac/) instaladas. (También puede abrir esta solución en Xamarin Studio en un Mac y ejecutar el ejemplo directamente desde allí). Puede usar el [Emulador de Visual Studio para Android](https://www.visualstudio.com/features/msft-android-emulator-vs.aspx) si desea ejecutar el proyecto de Android. 

    ![Captura de pantalla de la barra de herramientas de Visual Studio, con iOS seleccionado como proyecto de inicio.](/readme-images/SelectProject.png "Seleccione el proyecto en Visual Studio.")

4. Presione F5 para compilar y depurar. Ejecute la solución e inicie sesión con su cuenta personal, profesional o educativa.
    > **Nota** Es posible que tenga que abrir el administrador de configuración de compilación para asegurarse de que los pasos de compilación e implementación están seleccionados para el proyecto UWP.

<a name="run"></a>
## Ejecutar el ejemplo
<a id="run-the-sample" class="xliff"></a>

Al iniciarse, la aplicación muestra una lista que representan las tareas o "casos" de usuario comunes. Cada caso está formado por uno o más fragmentos de código. Los casos se agrupan por el nivel de permiso y de tipo de cuenta:

- Tareas que son aplicables a cuentas profesionales, educativas y personales, como recibir y enviar correo electrónico, crear archivos, etc.
- Tareas que solamente son aplicables a cuentas profesionales o educativas, como obtener fotos de administrador o de la cuenta de un usuario.
- Tareas que solo son aplicables a una cuenta profesional o educativa con permisos administrativos, como obtener miembros del grupo o crear nuevas cuentas de usuario.

Seleccione las historias que desea ejecutar y elija el botón 'ejecutar selección'. Se le pedirá que inicie sesión con su cuenta profesional, educativa o personal. Tenga en cuenta que si inicia una sesión con una cuenta que no tiene permisos aplicables para las historias que ha seleccionado (por ejemplo, si selecciona historias que son aplicables solo para una cuenta profesional o educativa y, después, inicia sesión con una cuenta personal), se producirá un error en esas historias.

Cada caso se vuelve verde si se acepta y rojo si se produce un error. Se envía información adicional a la ventana de salida. 

<a name="#how-the-sample-affects-your-tenant-data"></a>
##Repercusión de la muestra en los datos en su cuenta
<a id="how-the-sample-affects-your-account-data" class="xliff"></a>

Este ejemplo ejecuta comandos que crean, leen, actualizan o eliminan datos. No editarán ni eliminarán datos reales de la cuenta. Pero puede crear y dejar artefactos de datos en su cuenta como parte de su funcionamiento: al ejecutar comandos que crean, actualizan o eliminan, el ejemplo crea entidades falsas, como nuevos usuarios o grupos, para no afectar a los datos reales de la cuenta. 

El ejemplo puede dejar estas entidades falsas en su cuenta, si elige historias que crean o actualizan entidades. Por ejemplo, al seleccionar la ejecución de la historia 'actualizar grupo' se crea un nuevo grupo y, después, se actualiza. En este caso, el nuevo grupo permanece en su cuenta después de ejecutar el ejemplo.

<a name="add-a-snippet"></a>
## Agregar un fragmento de código
<a id="add-a-snippet" class="xliff"></a>

Este proyecto incluye dos archivos de fragmentos de código: 

- Groups\GroupSnippets.cs 
- Users\UserSnippets.cs.

Si tiene un fragmento de código propio que le gustaría ejecutar en este proyecto, solo tiene que seguir estos tres pasos:

1. **Agregue su fragmento de código al archivo de fragmentos de código.** No olvide incluir un bloque try/catch. 

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

2. **Cree un caso que use su fragmento de código y agréguelo al archivo de casos asociado.** Por ejemplo, el caso `TryGetMeAsync()` usa el `GetMeAsync()` fragmento de código dentro del archivo Users\UserStories.cs:

        public static async Task<bool> TryGetMeAsync()
        {
            var currentUser = await UserSnippets.GetMeAsync();

            return currentUser != null;
        }        


En ocasiones el caso deberá ejecutar fragmentos de código además del que se está implementando. Por ejemplo, si desea actualizar un evento, primero debe usar el método `CreateEventAsync()` para crear un evento. Después puede actualizarlo. Asegúrese siempre de usar fragmentos de código que ya existan en el archivo de fragmentos de código. Si la operación que necesita no existe, deberá crearla y, después, incluirla en su caso. Es recomendable eliminar todas las entidades que crea en un caso, especialmente si está trabajando en algo diferente a una cuenta de prueba o desarrollador.

3. **Agregue su caso a la lista de casos en MainPage.xaml.cs** (en el método `CreateStoryList()`):
    
    `snippetList.Children.Add(new CheckBox 
        { StoryName = "Get Me", GroupName = "Users", AccountType = "All", RunStoryAsync = UserStories.TryGetMeAsync });`
    
Ahora puede probar su fragmento de código. Cuando ejecute la aplicación, su caso se mostrarán como un nuevo elemento. Active la casilla para el fragmento de código y, después, ejecútelo. Esto le servirá para depurar el fragmento de código.

<a name="questions"></a>
## Preguntas y comentarios
<a id="questions-and-comments" class="xliff"></a>

Nos encantaría recibir sus comentarios acerca del ejemplo de fragmentos de código de Microsoft Graph para el proyecto Xamarin.Forms. Puede enviarnos sus preguntas y sugerencias a través de la sección [Problemas](https://github.com/MicrosoftGraph/xamarin-csharp-snippets-sample/issues) de este repositorio.

Su opinión es importante para nosotros. Conecte con nosotros en [Desbordamiento de pila](http://stackoverflow.com/questions/tagged/office365+or+microsoftgraph). Etiquete sus preguntas con [MicrosoftGraph].

<a name="contributing"></a>
## Colaboradores
<a id="contributing" class="xliff"></a> ##

Si le gustaría contribuir a este ejemplo, consulte [CONTRIBUTING.MD](/CONTRIBUTING.md).

Este proyecto ha adoptado el [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/) (Código de conducta de código abierto de Microsoft). Para obtener más información, consulte las [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) (Preguntas más frecuentes del código de conducta) o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) con otras preguntas o comentarios.


<a name="additional-resources"></a>
## Recursos adicionales
<a id="additional-resources" class="xliff"></a> ##

- [Otros ejemplos de Connect de Microsoft Graph](https://github.com/MicrosoftGraph?utf8=%E2%9C%93&query=-Connect)
- [Información general de Microsoft Graph](http://graph.microsoft.io)
- [Ejemplos de código de Office Developer](http://dev.office.com/code-samples)
- [Centro de desarrollo de Office](http://dev.office.com/)


## Copyright
<a id="copyright" class="xliff"></a>
Copyright (c) 2017 Microsoft. Todos los derechos reservados.


