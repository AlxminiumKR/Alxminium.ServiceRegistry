using MySql.Data.MySqlClient;
using System;
using System.Windows;
using Alxminium.ServiceRegistry.Models;

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

        public User Authenticate(string login, string password)
        {
            try
            {
                using (var conn = GetConnection())
                {
                    conn.Open();
                    string sql = "SELECT id, login, role, section FROM users WHERE login = @login AND password = @pass";

                    using (var cmd = new MySqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@login", login);
                        cmd.Parameters.AddWithValue("@pass", password);

                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                return new User
                                {
                                    Id = reader.GetInt32("id"),
                                    Login = reader.GetString("login"),
                                    Role = reader.GetString("role"),
                                    Section = reader.IsDBNull(reader.GetOrdinal("section")) ? "" : reader.GetString("section")
                                };
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при авторизации: " + ex.Message);
            }

            return null;
        }
    }
}