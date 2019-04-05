//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using System.Resources;
using System.Globalization;
using Xamarin.Forms;

namespace Graph_Xamarin_CS_Snippets
{
    public partial class MainPage : ContentPage
    {

        private static GraphServiceClient graphClient = null;
        public MainPage()
        {
            InitializeComponent();
            AppTitle.Text = AppResources.AppTitle;
            runSelected.Text = AppResources.RunSelected;
            clearSelected.Text = AppResources.ClearSelected;
            disconnect.Text = AppResources.Disconnect;
            CreateStoryList();

        }

        protected override void OnAppearing()
        {
            
            // Developer code - if you haven't registered the app yet, we warn you. 
            if (App.ClientID == "")
            {
                infoText.Text = AppResources.PleaseRegister;
                infoText.IsVisible = true;               
                runSelected.IsEnabled = false;
            }

        }

        private void CreateStoryList()
        {
            // Stories applicable to both work or school and personal accounts
            // NOTE: All of these snippets require permissions available whether you sign into the sample 
            // with a or work or school (commerical) or personal (consumer) account

            snippetList.Children.Add(new Label { Text = AppResources.PersonalWorkAccess, FontSize = Xamarin.Forms.Device.GetNamedSize(NamedSize.Medium, typeof(Label)) });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.GetMe, GroupName = "Users", AccountType = "All", RunStoryAsync = UserStories.TryGetMeAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.ReadUsers, GroupName = "Users", AccountType = "All", RunStoryAsync = UserStories.TryGetUsersAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.GetDrive, GroupName = "Users", AccountType = "All", RunStoryAsync = UserStories.TryGetCurrentUserDriveAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.GetEvents, GroupName = "Users", AccountType = "All", RunStoryAsync = UserStories.TryGetEventsAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.CreateEvent, GroupName = "Users", AccountType = "All", RunStoryAsync = UserStories.TryCreateEventAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.UpdateEvent, GroupName = "Users", AccountType = "All", RunStoryAsync = UserStories.TryUpdateEventAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.DeleteEvent, GroupName = "Users", AccountType = "All", RunStoryAsync = UserStories.TryDeleteEventAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.GetMessages, GroupName = "Users", AccountType = "All", RunStoryAsync = UserStories.TryGetMessages });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.SendMessage, GroupName = "Users", AccountType = "All", RunStoryAsync = UserStories.TrySendMailAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.GetUserFiles, GroupName = "Users", AccountType = "All", RunStoryAsync = UserStories.TryGetCurrentUserFilesAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.CreateTextFile, GroupName = "Users", AccountType = "All", RunStoryAsync = UserStories.TryCreateFileAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.DownloadFile, GroupName = "Users", AccountType = "All", RunStoryAsync = UserStories.TryDownloadFileAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.UpdateFile, GroupName = "Users", AccountType = "All", RunStoryAsync = UserStories.TryUpdateEventAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.RenameFile, GroupName = "Users", AccountType = "All", RunStoryAsync = UserStories.TryRenameFileAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.DeleteFile, GroupName = "Users", AccountType = "All", RunStoryAsync = UserStories.TryDeleteFileAsync });

            // Stories applicable only to work or school accounts
            // NOTE: All of these snippets will fail for lack of permissions if you sign into the sample with a personal (consumer) account

            snippetList.Children.Add(new Label { Text = AppResources.WorkAccess, FontSize = Xamarin.Forms.Device.GetNamedSize(NamedSize.Medium, typeof(Label)) });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.GetManager, GroupName = "Users", AccountType = "Work", RunStoryAsync = UserStories.TryGetCurrentUserManagerAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.GetDirects, GroupName = "Users", AccountType = "Work", RunStoryAsync = UserStories.TryGetDirectReportsAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.GetPhoto, GroupName = "Users", AccountType = "Work", RunStoryAsync = UserStories.TryGetCurrentUserPhotoAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.CreateFolder, GroupName = "Users", AccountType = "Work", RunStoryAsync = UserStories.TryCreateFolderAsync });


            // Stories applicable only to work or school accounts with admin access
            // NOTE: All of these snippets will fail for lack of permissions if you sign into the sample with a non-admin work account
            // or any consumer account. 

            snippetList.Children.Add(new Label { Text = AppResources.WorkAdminAccess, FontSize = Xamarin.Forms.Device.GetNamedSize(NamedSize.Medium, typeof(Label)) });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.CreateUser, GroupName = "Users", AccountType = "WorkAdmin", RunStoryAsync = UserStories.TryCreateUserAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.GetUserGroups, GroupName = "Users", AccountType = "WorkAdmin", RunStoryAsync = UserStories.TryGetCurrentUserGroupsAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.GetAllGroups, GroupName = "Groups", AccountType = "WorkAdmin", RunStoryAsync = GroupStories.TryGetGroupsAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.GetAGroup, GroupName = "Groups", AccountType = "WorkAdmin", RunStoryAsync = GroupStories.TryGetGroupAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.GetMembers, GroupName = "Groups", AccountType = "WorkAdmin", RunStoryAsync = GroupStories.TryGetGroupMembersAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.GetOwners, GroupName = "Groups", AccountType = "WorkAdmin", RunStoryAsync = GroupStories.TryGetGroupOwnersAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.CreateGroup, GroupName = "Groups", AccountType = "WorkAdmin", RunStoryAsync = GroupStories.TryCreateGroupAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.UpdateGroup, GroupName = "Groups", AccountType = "WorkAdmin", RunStoryAsync = GroupStories.TryUpdateGroupAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.DeleteGroup, GroupName = "Groups", AccountType = "WorkAdmin", RunStoryAsync = GroupStories.TryDeleteGroupAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.AddMember, GroupName = "Groups", AccountType = "WorkAdmin", RunStoryAsync = GroupStories.TryAddUserToGroup });


            // Excel snippets. These stories are applicable only to work or school accounts.
            // NOTE: All of these snippets will fail for lack of permissions if you sign into the sample with a personal (consumer) account

            snippetList.Children.Add(new CheckBox { StoryName = AppResources.UploadXLFile, GroupName = "Users", AccountType = "WorkAdmin", RunStoryAsync = UserStories.TryUploadExcelFileAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.CreateXLChart, GroupName = "Users", AccountType = "WorkAdmin", RunStoryAsync = UserStories.TryCreateExcelChartFromTableAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.GetXLRange, GroupName = "Users", AccountType = "WorkAdmin", RunStoryAsync = UserStories.TryGetExcelRangeAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.UpdateXLRange, GroupName = "Users", AccountType = "WorkAdmin", RunStoryAsync = UserStories.TryUpdateExcelRangeAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.ChangeXLNumFormat, GroupName = "Users", AccountType = "WorkAdmin", RunStoryAsync = UserStories.TryChangeExcelNumberFormatAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.XLABSFunction, GroupName = "Users", AccountType = "WorkAdmin", RunStoryAsync = UserStories.TryAbsExcelFunctionAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.SetXLFormula, GroupName = "Users", AccountType = "WorkAdmin", RunStoryAsync = UserStories.TrySetExcelFormulaAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.AddXLTable, GroupName = "Users", AccountType = "WorkAdmin", RunStoryAsync = UserStories.TryAddExcelTableToUsedRangeAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.AddXLTableRow, GroupName = "Users", AccountType = "WorkAdmin", RunStoryAsync = UserStories.TryAddExcelRowToTableAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.SortXLTable, GroupName = "Users", AccountType = "WorkAdmin", RunStoryAsync = UserStories.TrySortExcelTableOnFirstColumnValueAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.FilterXLTable, GroupName = "Users", AccountType = "WorkAdmin", RunStoryAsync = UserStories.TryFilterExcelTableValuesAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.ProtectWorksheet, GroupName = "Users", AccountType = "WorkAdmin", RunStoryAsync = UserStories.TryProtectExcelWorksheetAsync });
            snippetList.Children.Add(new CheckBox { StoryName = AppResources.UnprotectWorksheet, GroupName = "Users", AccountType = "WorkAdmin", RunStoryAsync = UserStories.TryUnprotectExcelWorksheetAsync });

            snippetList.Children.Add(new CheckBox { StoryName = AppResources.DeleteXLFile, GroupName = "Users", AccountType = "WorkAdmin", RunStoryAsync = UserStories.TryDeleteExcelFileAsync });
            





        }

        async void RunSelectedStories_Click(object sender, EventArgs args)
        {
            bool all = false;
            bool work = false;
            bool admin = false;

            //Iterate through the list of stories displayed in the UI
            //hide the ones that aren't selected
            //then hide any category label that doesn't have an associated story displayed

            for (int i = 0; i < snippetList.Children.Count; i++)
            {
                var item = snippetList.Children[i];

                if (item is CheckBox)
                {
                    CheckBox checkbox = (CheckBox)item;

                    //If item isn't selected, then hide it 

                    if (checkbox.IsChecked)
                    {
                        switch (checkbox.AccountType)
                        {
                            case "All":
                                all = true;
                                break;
                            case "Work":
                                work = true;
                                break;
                            case "WorkAdmin":
                                admin = true;
                                break;
                        }

                    }
                    else
                    {
                        checkbox.IsVisible = false;
                    }
                }

            }
            this.snippetList.Children[0].IsVisible = all;
            this.snippetList.Children[16].IsVisible = work;
            this.snippetList.Children[21].IsVisible = admin;

            //Now that we're going to connect to the service
            //enable the disconnect button
            disconnect.IsEnabled = true;

            //Iterate through the list of stories again
            //and this time, run the associated delegate method
            //then update checkbox color to show success or failure

            for (int i = 0; i < snippetList.Children.Count; i++)
            {
                bool result = false;
                var item = snippetList.Children[i];


                if (item is CheckBox)
                {
                    CheckBox checkbox = (CheckBox)item;

                    //If item is selected, run the associated story 
                    //and format the item according to story result

                    if (checkbox.IsChecked)
                    {

                        try
                        {
                            result = await checkbox.RunStoryAsync();
                            
                            Debug.WriteLine(String.Format("{0}.{1} {2}", checkbox.GroupName, checkbox.StoryName, (result) ? "passed" : "failed"));
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine("{0}.{1} failed. Exception: {2}", checkbox.GroupName, checkbox.StoryName, ex.Message);
                            result = false;
                        }
                        checkbox.showResultTest(result);

                    }
                }

            }

        }



        void ClearSelection_Click(object sender, EventArgs args)
        {
            for (int i = 0; i < snippetList.Children.Count; i++)
            {
                var item = snippetList.Children[i];

                VisualElement element = (VisualElement)item;
                element.IsVisible = true;

                if (item is CheckBox)
                {
                    CheckBox checkbox = (CheckBox)item;
                    checkbox.deselect();
                    //checkbox.IsVisible = true;
                }

            }
        }

        void Disconnect_Click(object sender, EventArgs args)
        {
            AuthenticationHelper.SignOut();
            ClearSelection_Click(sender, args);
            disconnect.IsEnabled = false;
        }
    }
}
