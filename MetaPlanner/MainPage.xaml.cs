using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using System.Threading.Tasks;
using System.Net.Http.Headers;
using Windows.UI.Popups;
using MetaPlanner.Model;
using MetaPlanner.Output;
using System.IO;
using Windows.Storage;
using System.Text;
using Serilog;
using System.Xml.Serialization;
using Windows.UI.Xaml.Data;
using System.Collections.Immutable;
using MetaPlanner.Control;
using System.Runtime.CompilerServices;



// La plantilla de elemento Página en blanco está documentada en https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0xc0a

namespace MetaPlanner
{
    /// <summary>
    /// Página vacía que se puede usar de forma independiente o a la que se puede navegar dentro de un objeto Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {

        //string redirectURI = Windows.Security.Authentication.Web.WebAuthenticationBroker.GetCurrentApplicationCallbackUri().ToString();
        // ms-app://s-1-15-2-148375016-475961868-2312470711-1599034693-979352800-1769312473-2847594358/


        private Command Control;

        public MainPage()
        {
            Control = new Command(this);
            this.InitializeComponent();
            lblMessage.Text = Command.config.Tenant;
        }
        


        /// <summary>
        /// Displays a message in the ResultText. Can be called from any thread.
        /// </summary>
        private async Task DisplayMessageAsync(string message)
        {
            await Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal,
                   () =>
                   {
                       lblMessage.Text = message;      
                    });

            // Create the message dialog and set its content
            var messageDialog = new MessageDialog(message,"Error");

            // Set the command that will be invoked by default
           // messageDialog.DefaultCommandIndex = 0;

            // Set the command to be invoked when escape is pressed
           // messageDialog.CancelCommandIndex = 1;

            // Show the message dialog
            await messageDialog.ShowAsync();
        }

        public void DisplayMessage(string message)
        {
            lblMessage.Text = message; //DisplayMessageAsync(message);
        }
        private async void btnLoadUsers_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                App.logger.Information("Start ProcessUsers");
                Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Wait, 1);

                await Control.ProcessUsers();
                RadDataGrid.DataContext = Control.PlannerUsers.Values;

                Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Arrow, 1);
                App.logger.Information("End ProcessUsers");
             }
            catch (Exception exception)
            {
                await DisplayMessageAsync($"ProcessUsers:{System.Environment.NewLine}{exception}");
                App.logger.Error(exception.Message);
                return;
            }
        }

        private async void btnLoadTasks_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                App.logger.Information("Start ProcessTasks");
                Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Wait, 1);

                await Control.ProcessTasks();
                RadDataGrid.DataContext = Control.PlannerTasks.Values;

                Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Arrow, 1);
                App.logger.Information("End ProcessTasks");
            }
            catch (Exception exception)
            {
                await DisplayMessageAsync($"ProcessTasks:{System.Environment.NewLine}{exception}");
                App.logger.Error(exception.Message);
                return;
            }
        }

        private async void btnLoadBuckets_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                App.logger.Information("Start ProcessBuckets");
                Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Wait, 1);

                await Control.ProcessBuckets();
                RadDataGrid.DataContext = Control.PlannerBuckets.Values;

                Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Arrow, 1);
                App.logger.Information("End ProcessBuckets");
            }
            catch (Exception exception)
            {
                await DisplayMessageAsync($"ProcessBuckets:{System.Environment.NewLine}{exception}");
                App.logger.Error(exception.Message);
                return;
            }
        }

        private async void btnLoadPlans_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                App.logger.Information("Start ProcessPlans");
                Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Wait, 1);

                await Control.ProcessPlans();
                RadDataGrid.DataContext = Control.PlannerPlans.Values;

                Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Arrow, 1);
                App.logger.Information("End ProcessPlans");
            }
            catch (Exception exception)
            {
                await DisplayMessageAsync($"ProcessPlans:{System.Environment.NewLine}{exception}");
                App.logger.Error(exception.Message);
                return;
            }
        }

        private async void btnClean_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                App.logger.Information("Start CleanAllSharePointLists");
                Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Wait, 1);

                await Control.CleanAllSharePointLists();

                Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Arrow, 1);
                App.logger.Information("End CleanAllSharePointLists");
            }
            catch (Exception exception)
            {
                await DisplayMessageAsync($"CleanAllSharePointLists:{System.Environment.NewLine}{exception}");
                App.logger.Error(exception.Message);
                return;
            }
        }

        private async void btnLoadAll_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                App.logger.Information("Start ProcessAll");
                Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Wait, 1);

                await Control.ProcessAll();
                RadDataGrid.DataContext = Control.PlannerPlans.Values;

                Windows.UI.Xaml.Window.Current.CoreWindow.PointerCursor = new Windows.UI.Core.CoreCursor(Windows.UI.Core.CoreCursorType.Arrow, 1);
                App.logger.Information("End ProcessAll");
            }
            catch (Exception exception)
            {
                await DisplayMessageAsync($"ProcessAll:{System.Environment.NewLine}{exception}");
                App.logger.Error(exception.Message);
                return;
            }

        }
    }
}
