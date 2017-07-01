//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Xamarin.Forms;
using Microsoft.Identity.Client;


namespace Graph_Xamarin_CS_Snippets
{
    public class App : Application
    {
        public static PublicClientApplication IdentityClientApp = null;
        public static string ClientID = "ENTER_YOUR_CLIENT_ID";
        public static string RedirectUri = "msal" + ClientID + "://auth";
        public static UIParent UiParent = null;
        public static string[] Scopes = {
                        "https://graph.microsoft.com/User.Read",
                        "https://graph.microsoft.com/User.ReadWrite",
                        "https://graph.microsoft.com/User.ReadBasic.All",
                        "https://graph.microsoft.com/Mail.Send",
                        "https://graph.microsoft.com/Calendars.ReadWrite",
                        "https://graph.microsoft.com/Mail.ReadWrite",
                        "https://graph.microsoft.com/Files.ReadWrite",

                        // Admin-only scopes. Comment these out if you're running the sample with a non-admin work account.
                        // You won't be able to sign in with a non-admin work account if you request these scopes.
                        // These scopes will be ignored if you leave them uncommented and run the sample with a consumer account.
                        // See the MainPage.xaml.cs file for all of the operations that won't work if you're not running the 
                        // sample with an admin work account.
                        "https://graph.microsoft.com/Directory.AccessAsUser.All",
                        "https://graph.microsoft.com/User.ReadWrite.All",
                        "https://graph.microsoft.com/Group.ReadWrite.All",

                    };



        public App()
        {
            IdentityClientApp = new PublicClientApplication(ClientID);
            MainPage = new MainPage
            {

            };
        }


        protected override void OnStart()
        {
            // Handle when your app starts
        }

        protected override void OnSleep()
        {
            // Handle when your app sleeps
        }

        protected override void OnResume()
        {
            // Handle when your app resumes
        }
    }
}
