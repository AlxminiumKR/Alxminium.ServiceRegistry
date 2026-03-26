using System.Configuration;
using System.Data;
using System.Windows;

namespace Alxminium.ServiceRegistry
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            LoginWindow loginWindow = new LoginWindow();
            loginWindow.Show();
        }
    }
}
