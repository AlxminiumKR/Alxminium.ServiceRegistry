using MySql.Data.MySqlClient;
using System;
using System.Windows;

namespace Alxminium.ServiceRegistry
{
    public class DatabaseManager
    {
        private string _connectionString = "Server=localhost;Port=3306;Database=alxminium_db;Uid=root;Pwd=1234;";

        public MySqlConnection GetConnection()
        {
            return new MySqlConnection(_connectionString);
        }

        public bool CheckConnection()
        {
            try
            {
                using (var conn = GetConnection())
                {
                    conn.Open();
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка подключения к MySQL: " + ex.Message);
                return false;
            }
        }
    }
}