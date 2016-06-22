//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xamarin.Forms;
using Graph_Xamarin_CS_Snippets;

namespace Graph_Xamarin_CS_Snippets
{
    //Implementation of custom checkbox view based on example in
    //Creating Mobile Apps with Xamarin.Forms by Charles Petzold
    // https://developer.xamarin.com/guides/xamarin-forms/creating-mobile-apps-xamarin-forms/

    public partial class CheckBox : ContentView
    {
        string storyName;
        string groupName;
        string accountType;

        public string StoryName
        {
            set
            {
                storyName = value;
                textLabel.Text = value;
            }
            get
            {
                return storyName;
            }
        }

        public string GroupName
        {
            set
            {
                groupName = value;
            }
            get
            {
                return groupName;
            }
        }

        public string AccountType
        {
            set
            {
                accountType = value;
            }
            get
            {
                return accountType;
            }
        }

        // Delegate method to call
        public Func<Task<bool>> RunStoryAsync { get; set; }

        public static readonly BindableProperty IsCheckedProperty =
            BindableProperty.Create(
                "IsChecked", typeof(bool), typeof(CheckBox),
                false,
                propertyChanged: (bindable, oldValue, newValue) =>
                {
                    //Updates checkbox to be checked or not
                    //and updates checkbox formatting accordingly

                    CheckBox checkbox = (CheckBox)bindable;

                    if ((bool)newValue)
                    {
                        checkbox.boxLabel.Text = "\u2611";
                        checkbox.frame.BackgroundColor = Color.Accent;
                        checkbox.textLabel.TextColor = Color.White;
                    }
                    else
                    {
                        checkbox.boxLabel.Text = "\u2610";
                        checkbox.frame.BackgroundColor = Color.Default;
                        checkbox.textLabel.TextColor = Color.Default;
                    }

                    //Raises the event
                    EventHandler<bool> eventHandler = checkbox.CheckChanged;
                    if (eventHandler != null)
                    {
                        eventHandler(checkbox, (bool)newValue);
                    }
                });

        //Sets background color based on delegate method result
        public void showResultTest(bool result)
        {

            if (result)
            {
                this.frame.BackgroundColor = Color.Green;
            }
            else
            {
                this.frame.BackgroundColor = Color.Red;
            }
        }



        //Changes checkbox to be unchecked
        //Sets background color back to default
        public void deselect()
        {
            this.IsChecked = false;
            this.frame.BackgroundColor = Color.Default;
        }

        public event EventHandler<bool> CheckChanged;
        public CheckBox()
        {
            InitializeComponent();
        }
        public bool IsChecked
        {
            set { SetValue(IsCheckedProperty, value); }
            get { return (bool)GetValue(IsCheckedProperty); }
        }

        //TapGestureRecognizer handler
        void OnCheckBoxTapped(object sender, EventArgs args)
        {
            IsChecked = !IsChecked;
        }
    }

}
