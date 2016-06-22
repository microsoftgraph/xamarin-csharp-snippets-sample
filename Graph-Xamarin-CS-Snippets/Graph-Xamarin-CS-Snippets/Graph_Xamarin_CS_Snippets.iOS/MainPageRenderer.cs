using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Text;
using Graph_Xamarin_CS_Snippets;
using Graph_Xamarin_CS_Snippets.iOS;
using Xamarin.Forms;
using Xamarin.Forms.Platform.iOS;

[assembly: ExportRenderer(typeof(MainPage), typeof(MainPageRenderer))]

namespace Graph_Xamarin_CS_Snippets.iOS
{
    class MainPageRenderer : PageRenderer
    {
        MainPage page;
        protected override void OnElementChanged(VisualElementChangedEventArgs e)
        {
            base.OnElementChanged(e);
            page = e.NewElement as MainPage;
        }
        public override void ViewDidLoad()
        {
            base.ViewDidLoad();
            page.platformParameters = new PlatformParameters(this);
        }
    }
}
