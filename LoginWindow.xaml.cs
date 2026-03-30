using Alxminium.ServiceRegistry.Models;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace Alxminium.ServiceRegistry
{
    public partial class LoginWindow : Window
    {
        public LoginWindow()
        {
            InitializeComponent();
            InputLanguageManager.Current.InputLanguageChanged += (s, e) =>
            {
                UpdateLanguageIndicator();
            };
            /*this.PreviewKeyDown += (s, e) => {
                if (e.Key == Key.CapsLock) UpdateCapsIndicator();
            };*/
        }
        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.CapsLock || TxtPassword.IsFocused)
            {
                UpdateCapsIndicator();
            }
        }
        private void UpdateCapsIndicator()
        {
            bool isCapsLock = Keyboard.IsKeyToggled(Key.CapsLock);
            IconCaps.Visibility = (isCapsLock && TxtPassword.IsFocused)
                                  ? Visibility.Visible
                                  : Visibility.Collapsed;
        }

        private void UpdateLanguageIndicator()
        {
            var currentLang = InputLanguageManager.Current.CurrentInputLanguage.Name;

            if (currentLang.StartsWith("ru"))
            {
                TxtLang.Text = "RU";
                TxtLang.Foreground = Brushes.OrangeRed;
            }
            else
            {
                TxtLang.Text = "EN";
                TxtLang.Foreground = Brushes.Black;
            }
        }

        private void TxtPassword_GotFocus(object sender, RoutedEventArgs e)
        {
            TxtLang.Visibility = Visibility.Visible;
            UpdateLanguageIndicator();
            UpdateCapsIndicator();
        }

        private void TxtPassword_LostFocus(object sender, RoutedEventArgs e)
        {
            TxtLang.Visibility = Visibility.Collapsed;
            UpdateCapsIndicator();
        }

        /*private void TxtPassword_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() => {
                UpdateLanguageIndicator();
            }), System.Windows.Threading.DispatcherPriority.ContextIdle);
        }*/

        private void BtnLogin_Click(object sender, RoutedEventArgs e)
        {
            string login = TxtLogin.Text.Trim();
            string password = TxtPassword.Password;

            if (string.IsNullOrEmpty(login) || string.IsNullOrEmpty(password))
            {
                MessageBox.Show("Введите логин и пароль! - alxminium", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                DatabaseManager db = new DatabaseManager();
                User user = db.Authenticate(login, password);

                if (user != null)
                {
                    MainWindow main = new MainWindow(user);
                    main.Show();
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Неверный логин или пароль! - alxminium", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка базы данных: {ex.Message} - alxminium");
            }
        }
    }
}