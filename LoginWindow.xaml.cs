using Alxminium.ServiceRegistry.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Alxminium.ServiceRegistry
{
    public partial class LoginWindow : Window
    {
        public LoginWindow()
        {
            InitializeComponent();
        }

        private void BtnLogin_Click(object sender, RoutedEventArgs e)
        {
            string login = TxtLogin.Text;
            string password = TxtPassword.Password;

            if (string.IsNullOrEmpty(login) || string.IsNullOrEmpty(password))
            {
                MessageBox.Show("Введите логин и пароль!");
                return;
            }

            DatabaseManager db = new DatabaseManager();
            User user = db.Authenticate(login, password);

            if (user != null)
            {
                // Если залогинились — открываем главное окно и передаем туда юзера
                MainWindow main = new MainWindow(user);
                main.Show();
                this.Close(); // Закрываем окно логина
            }
            else
            {
                MessageBox.Show("Неверный логин или пароль!");
            }
        }
    }
}