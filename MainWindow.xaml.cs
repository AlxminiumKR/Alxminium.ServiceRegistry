using Alxminium.ServiceRegistry.Models;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Win32;
using MySql.Data.MySqlClient;
using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Animation;

namespace Alxminium.ServiceRegistry
{
    public partial class MainWindow : Window
    {
        public User CurrentUser { get; set; }
        public MainWindow(User user)
        {
            InitializeComponent();
            CurrentUser = user;
            ApplyPermissions();
            ShowWelcomeAnimation(CurrentUser.Login);

            LoadReferenceData();
            LoadRequestsFromDb();
            //DataStorage.InitializeMockData();

            ComboObjects.ItemsSource = DataStorage.Objects;
            ComboServices.ItemsSource = DataStorage.ServiceTasks;
            this.Title += $" - Вы вошли как: {CurrentUser.Login} ({CurrentUser.Role})";
            //GridRequests.ItemsSource = DataStorage.Requests;
            UpdateStatistics();
        }
        private void ShowWelcomeAnimation(string userName)
        {
            TxtWelcomeUser.Text = $"Рады видеть вас, {userName}!";

            DoubleAnimation fadeAnimation = new DoubleAnimation
            {
                From = 1.0,
                To = 0.0,
                Duration = TimeSpan.FromSeconds(2),
                BeginTime = TimeSpan.FromSeconds(2)
            };

            fadeAnimation.Completed += (s, e) => {
                WelcomeOverlay.Visibility = Visibility.Collapsed;
            };

            WelcomeOverlay.BeginAnimation(Grid.OpacityProperty, fadeAnimation);
        }
        private void ApplyPermissions()
        {
            if (CurrentUser.Role != "Admin")
            {
                if (TabAdmin != null)
                {
                    TabAdmin.Visibility = Visibility.Hidden;
                    TabAdmin.IsEnabled = false;

                    if (GridAllObjects != null) GridAllObjects.ContextMenu = null;
                    if (GridAllServices != null) GridAllServices.ContextMenu = null;
                }
            }
            else
            {
                if (TabAdmin != null)
                {
                    TabAdmin.Visibility = Visibility.Visible;
                    TabAdmin.IsEnabled = true;
                }
            }
        }
        private void LoadRequestsFromDb()
        {
            var db = new DatabaseManager();
            try
            {
                using (var conn = db.GetConnection())
                {
                    conn.Open();

                    string sql = @"
                    SELECT id, author, section, object_id, object_name, 
                    work_name, work_type, deadline_days, unit, price, 
                    volume, total_cost, status, description, created_at 
                    FROM requests";

                    if (CurrentUser.Role != "Admin")
                    {
                        sql += " WHERE author = @author";
                    }

                    using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(sql, conn))
                    {
                        if (CurrentUser.Role != "Admin")
                        {
                            cmd.Parameters.AddWithValue("@author", CurrentUser.Login);
                        }

                        using (var reader = cmd.ExecuteReader())
                        {
                            DataStorage.Requests.Clear();

                            while (reader.Read())
                            {
                                var req = new WorkRequest
                                {
                                    Id = reader.GetInt32("id"),
                                    Author = reader.IsDBNull(reader.GetOrdinal("author")) ? "" : reader.GetString("author"),
                                    Section = reader.IsDBNull(reader.GetOrdinal("section")) ? "" : reader.GetString("section"),
                                    ObjectId = reader.IsDBNull(reader.GetOrdinal("object_id")) ? 0 : reader.GetInt32("object_id"),

                                    ObjectName = reader.IsDBNull(reader.GetOrdinal("object_name")) ? "Не указан" : reader.GetString("object_name"),
                                    WorkName = reader.IsDBNull(reader.GetOrdinal("work_name")) ? "Без названия" : reader.GetString("work_name"),
                                    WorkType = reader.IsDBNull(reader.GetOrdinal("work_type")) ? "" : reader.GetString("work_type"),

                                    DeadlineDays = reader.IsDBNull(reader.GetOrdinal("deadline_days")) ? 0 : reader.GetInt32("deadline_days"),
                                    Unit = reader.IsDBNull(reader.GetOrdinal("unit")) ? "" : reader.GetString("unit"),
                                    Price = reader.IsDBNull(reader.GetOrdinal("price")) ? 0m : reader.GetDecimal("price"),

                                    Volume = reader.IsDBNull(reader.GetOrdinal("volume")) ? 0 : reader.GetDouble("volume"),
                                    TotalCost = reader.IsDBNull(reader.GetOrdinal("total_cost")) ? 0m : reader.GetDecimal("total_cost"),
                                    Status = reader.IsDBNull(reader.GetOrdinal("status")) ? "В очереди" : reader.GetString("status"),
                                    Description = reader.IsDBNull(reader.GetOrdinal("description")) ? "" : reader.GetString("description"),

                                    CreatedAt = reader.IsDBNull(reader.GetOrdinal("created_at")) ? DateTime.Now : reader.GetDateTime("created_at")
                                };

                                DataStorage.Requests.Add(req);
                            }
                        }
                    }
                }
                RefreshRequestsGrid();
                UpdateStatistics();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке заявок: " + ex.Message);
            }
        }

        private async void BtnSendFeedback_Click(object sender, RoutedEventArgs e)
        {
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
            string message = TxtFeedbackMessage.Text;
            if (string.IsNullOrWhiteSpace(message)) return;

            FeedbackProgress.Visibility = Visibility.Visible;
            BtnSendFeedback.Content = "";
            BtnSendFeedback.IsEnabled = false;
            TxtFeedbackMessage.IsEnabled = false;

            string senderName = Environment.UserName;
            string pcName = Environment.MachineName;

            string text = $"🚀 *Новый фидбек!*\n👤 *От:* {senderName} (ПК: {pcName})\n📝 *Сообщение:* {message}";
            string url = $"{Secrets.ProxyUrl}/bot{Secrets.Token}/sendMessage?chat_id={Secrets.ChatId}&text={Uri.EscapeDataString(text)}&parse_mode=Markdown";

            try
            {
                using (var client = new System.Net.Http.HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(7);

                    var response = await client.GetAsync(url);

                    if (response.IsSuccessStatusCode)
                    {
                        FeedbackDialog.IsOpen = false;
                        TxtFeedbackMessage.Clear();
                        MessageBox.Show("Сообщение успешно доставлено!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        string errorContent = await response.Content.ReadAsStringAsync();
                        MessageBox.Show($"Ошибка: {response.StatusCode}\n{errorContent} Возможно, требуется VPN.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
            catch (TaskCanceledException)
            {
                MessageBox.Show("Время ожидания истекло. Возможно, требуется VPN.",
                                "Ошибка соединения", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось отправить сообщение: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                FeedbackProgress.Visibility = Visibility.Collapsed;
                BtnSendFeedback.Content = "ОТПРАВИТЬ";
                BtnSendFeedback.IsEnabled = true;
                TxtFeedbackMessage.IsEnabled = true;
            }
        }

        private void BtnOpenFeedback_Click(object sender, RoutedEventArgs e) => FeedbackDialog.IsOpen = true;
        private void BtnCloseFeedback_Click(object sender, RoutedEventArgs e) => FeedbackDialog.IsOpen = false;

        private void UpdateRequestDetailsInDb(WorkRequest req)
        {
            var db = new DatabaseManager();
            using (var conn = db.GetConnection())
            {
                conn.Open();
                string sql = @"UPDATE requests 
                       SET object_id = @objId, 
                           object_name = @objName, 
                           work_name = @wName, 
                           work_type = @wType, 
                           unit = @unit, 
                           price = @price, 
                           volume = @vol, 
                           total_cost = @total, 
                           description = @desc 
                       WHERE id = @id";

                using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@objId", req.ObjectId);
                    cmd.Parameters.AddWithValue("@objName", req.ObjectName);
                    cmd.Parameters.AddWithValue("@wName", req.WorkName);
                    cmd.Parameters.AddWithValue("@wType", req.WorkType);
                    cmd.Parameters.AddWithValue("@unit", req.Unit);
                    cmd.Parameters.AddWithValue("@price", req.Price);
                    cmd.Parameters.AddWithValue("@vol", req.Volume);
                    cmd.Parameters.AddWithValue("@total", req.TotalCost);
                    cmd.Parameters.AddWithValue("@desc", req.Description);
                    cmd.Parameters.AddWithValue("@id", req.Id);

                    cmd.ExecuteNonQuery();
                }
            }
        }

        private void MenuEditRequest_Click(object sender, RoutedEventArgs e)
        {
            if (GridRequests.SelectedItem is WorkRequest selectedRequest)
            {
                if (selectedRequest.Status == "Выполнена" && CurrentUser.Role != "Admin")
                {
                    MessageBox.Show("Вы не можете редактировать уже выполненную заявку.",
                                    "Доступ ограничен", MessageBoxButton.OK, MessageBoxImage.Stop);
                    return;
                }

                ComboObjects.SelectedItem = DataStorage.Objects.FirstOrDefault(o => o.Id == selectedRequest.ObjectId);
                ComboServices.SelectedItem = DataStorage.ServiceTasks.FirstOrDefault(s => s.WorkName == selectedRequest.WorkName);

                TxtVolume.Text = selectedRequest.Volume.ToString();
                TxtDescription.Text = selectedRequest.Description;

                BtnCreateRequest.Content = "Сохранить изменения";
                BtnCreateRequest.Tag = selectedRequest;
                MainTabControl.SelectedIndex = 0;
            }
        }
        private void DeleteRequestFromDb(int requestId)
        {
            var db = new DatabaseManager();
            try
            {
                using (var conn = db.GetConnection())
                {
                    conn.Open();
                    string sql = "DELETE FROM requests WHERE id = @id";

                    using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@id", requestId);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при удалении из базы: " + ex.Message);
            }
        }
        private void LoadReferenceData()
        {
            var db = new DatabaseManager();
            try
            {
                using (var conn = db.GetConnection())
                {
                    conn.Open();

                    string sqlObj = "SELECT * FROM objects";
                    using (var cmdObj = new MySql.Data.MySqlClient.MySqlCommand(sqlObj, conn))
                    using (var reader = cmdObj.ExecuteReader())
                    {
                        DataStorage.Objects.Clear();
                        while (reader.Read())
                        {
                            DataStorage.Objects.Add(new ServiceObject
                            {
                                Id = reader.GetInt32("id"),
                                Name = reader.IsDBNull(reader.GetOrdinal("name")) ? "Без названия" : reader.GetString("name"),
                                Address = reader.IsDBNull(reader.GetOrdinal("address")) ? "" : reader.GetString("address"),
                                ResponsiblePerson = reader.IsDBNull(reader.GetOrdinal("responsible")) ? "" : reader.GetString("responsible")
                            });
                        }
                    }

                    string sqlServ = "SELECT * FROM services";
                    using (var cmdServ = new MySql.Data.MySqlClient.MySqlCommand(sqlServ, conn))
                    using (var reader = cmdServ.ExecuteReader())
                    {
                        DataStorage.ServiceTasks.Clear();
                        while (reader.Read())
                        {
                            DataStorage.ServiceTasks.Add(new ServiceTask
                            {
                                Id = reader.GetInt32("id"),
                                WorkName = reader.IsDBNull(reader.GetOrdinal("work_name")) ? "Без названия" : reader.GetString("work_name"),
                                WorkType = reader.IsDBNull(reader.GetOrdinal("work_type")) ? "" : reader.GetString("work_type"),
                                Description = reader.IsDBNull(reader.GetOrdinal("description")) ? "" : reader.GetString("description"),
                                Unit = reader.IsDBNull(reader.GetOrdinal("unit")) ? "шт." : reader.GetString("unit"),
                                Price = reader.IsDBNull(reader.GetOrdinal("price")) ? 0m : reader.GetDecimal("price"),
                                DeadlineDays = reader.IsDBNull(reader.GetOrdinal("deadline_days")) ? 0 : reader.GetInt32("deadline_days")
                            });
                        }
                    }
                }

                GridAllObjects.ItemsSource = null;
                GridAllObjects.ItemsSource = DataStorage.Objects;

                GridAllServices.ItemsSource = null;
                GridAllServices.ItemsSource = DataStorage.ServiceTasks;

                ComboObjects.ItemsSource = null;
                ComboObjects.ItemsSource = DataStorage.Objects;

                ComboServices.ItemsSource = null;
                ComboServices.ItemsSource = DataStorage.ServiceTasks;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки справочников: " + ex.Message);
            }
        }

        private void SaveRequestToDb(WorkRequest req)
        {
            var db = new DatabaseManager();
            try
            {
                using (var conn = db.GetConnection())
                {
                    conn.Open();
                    string sql = @"INSERT INTO requests 
                                (author, section, object_id, object_name, work_name, work_type, deadline_days, unit, price, volume, total_cost, description, status) 
                                VALUES 
                                (@author, @section, @objId, @objName, @workName, @workType, @deadline, @unit, @price, @vol, @total, @desc, @status);
                                SELECT LAST_INSERT_ID();";

                    using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@author", req.Author);
                        cmd.Parameters.AddWithValue("@section", req.Section);
                        cmd.Parameters.AddWithValue("@objId", req.ObjectId);
                        cmd.Parameters.AddWithValue("@objName", req.ObjectName);
                        cmd.Parameters.AddWithValue("@workName", req.WorkName);
                        cmd.Parameters.AddWithValue("@workType", req.WorkType);
                        cmd.Parameters.AddWithValue("@deadline", req.DeadlineDays);
                        cmd.Parameters.AddWithValue("@unit", req.Unit);
                        cmd.Parameters.AddWithValue("@price", req.Price);
                        cmd.Parameters.AddWithValue("@vol", req.Volume);
                        cmd.Parameters.AddWithValue("@total", req.TotalCost);
                        cmd.Parameters.AddWithValue("@desc", req.Description ?? (object)DBNull.Value);
                        cmd.Parameters.AddWithValue("@status", req.Status);

                        var newId = cmd.ExecuteScalar();
                        if (newId != null) req.Id = Convert.ToInt32(newId);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка сохранения в базу: " + ex.Message);
            }
        }
        private void UpdateRequestStatusInDb(WorkRequest req)
        {
            var db = new DatabaseManager();
            try
            {
                using (var conn = db.GetConnection())
                {
                    conn.Open();
                    string sql = "UPDATE requests SET status = @status WHERE id = @id";

                    using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@status", req.Status);
                        cmd.Parameters.AddWithValue("@id", req.Id);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка обновления статуса в БД: " + ex.Message);
            }
        }
        /*private void BtnImportServices_Click(object sender, RoutedEventArgs e) #старый метод импорта услуг
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xlsm;*.xlsb;*.xls";

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    using (var workbook = new XLWorkbook(openFileDialog.FileName))
                    {
                        var worksheet = workbook.Worksheet(1);
                        var range = worksheet.RangeUsed();
                        if (range == null) return;
                        var rows = range.RowsUsed().Skip(1);

                        foreach (var row in rows)
                        {
                            var newTask = new ServiceTask
                            {
                                Id = DataStorage.ServiceTasks.Count + 1,
                                WorkName = row.Cell(1).GetValue<string>(),  
                                WorkType = row.Cell(2).GetValue<string>(),   
                                DeadlineDays = row.Cell(3).GetValue<int>()   
                            };
                            DataStorage.ServiceTasks.Add(newTask);
                        }
                    }
                    MessageBox.Show("Данные успешно импортированы! - alxminium");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при импорте: {ex.Message}");
                }
            }
        }*/

        private void BtnImportServices_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xlsm;*.xlsb;*.xls";

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    int addedCount = 0;
                    int duplicateCount = 0;

                    var db = new DatabaseManager();

                    using (var workbook = new ClosedXML.Excel.XLWorkbook(openFileDialog.FileName))
                    using (var conn = db.GetConnection())
                    {
                        var worksheet = workbook.Worksheet(1);
                        var range = worksheet.RangeUsed();
                        if (range == null) return;

                        var rows = range.RowsUsed().Skip(1);
                        conn.Open();

                        foreach (var row in rows)
                        {
                            string name = row.Cell(1).GetValue<string>()?.Trim();
                            string category = row.Cell(2).GetValue<string>()?.Trim();
                            string description = row.Cell(3).GetValue<string>()?.Trim();
                            string unit = row.Cell(4).GetValue<string>()?.Trim() ?? "шт.";

                            decimal price = row.Cell(5).TryGetValue<decimal>(out var p) ? p : 0m;
                            int deadline = row.Cell(6).TryGetValue<int>(out var d) ? d : 7;

                            if (string.IsNullOrWhiteSpace(name)) continue;

                            bool exists = DataStorage.ServiceTasks.Any(t =>
                                t.WorkName != null && t.WorkName.Equals(name, StringComparison.OrdinalIgnoreCase));

                            if (!exists)
                            {

                                string sql = @"INSERT INTO services (work_name, work_type, description, unit, price, deadline_days) 
                                       VALUES (@name, @type, @desc, @unit, @price, @deadline);
                                       SELECT LAST_INSERT_ID();";

                                using (var cmd = new MySqlCommand(sql, conn))
                                {
                                    cmd.Parameters.AddWithValue("@name", name);
                                    cmd.Parameters.AddWithValue("@type", string.IsNullOrEmpty(category) ? (object)DBNull.Value : category);
                                    cmd.Parameters.AddWithValue("@desc", string.IsNullOrEmpty(description) ? (object)DBNull.Value : description);
                                    cmd.Parameters.AddWithValue("@unit", unit);
                                    cmd.Parameters.AddWithValue("@price", price);
                                    cmd.Parameters.AddWithValue("@deadline", deadline);

                                    int newId = Convert.ToInt32(cmd.ExecuteScalar());

                                    var newTask = new ServiceTask
                                    {
                                        Id = newId,
                                        WorkName = name,
                                        WorkType = category,
                                        Description = description,
                                        Unit = unit,
                                        Price = price,
                                        DeadlineDays = deadline
                                    };

                                    DataStorage.ServiceTasks.Add(newTask);
                                }
                                addedCount++;
                            }
                            else
                            {
                                duplicateCount++;
                            }
                        }
                    }

                    MessageBox.Show($"Данные успешно импортированы!\nДобавлено: {addedCount}\nДубликатов пропущено: {duplicateCount} - alxminium",
                                    "Импорт завершен", MessageBoxButton.OK, MessageBoxImage.Information);

                    if (addedCount > 0) GridAllServices.Items.Refresh();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при импорте: {ex.Message}", "Критический сбой", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void RefreshRequestsGrid()
        {
            string searchText = TxtSearch.Text.Trim().ToLower();
            var filtered = DataStorage.Requests.AsEnumerable();

            if (ChkShowCompleted.IsChecked != true)
            {
                filtered = filtered.Where(r => r.Status == "В работе");
            }

            if (!string.IsNullOrWhiteSpace(searchText))
            {
                filtered = filtered.Where(r =>
                    (r.ObjectName != null && r.ObjectName.ToLower().Contains(searchText)) ||
                    (r.WorkName != null && r.WorkName.ToLower().Contains(searchText)) ||
                    (r.Description != null && r.Description.ToLower().Contains(searchText))
                );
            }

            GridRequests.ItemsSource = null;
            GridRequests.ItemsSource = filtered.ToList();

            UpdateStatistics();
        }

        private void TxtSearch_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            RefreshRequestsGrid();
        }
        private void UpdateStatistics()
        {
            var allRequests = DataStorage.Requests;

            if (allRequests != null)
            {
                int total = allRequests.Count;
                int completed = allRequests.Count(r => r.Status == "Выполнена");
                int pending = total - completed;

                TxtTotalCount.Text = total.ToString();
                TxtCompletedCount.Text = completed.ToString();
                TxtPendingCount.Text = pending.ToString();
            }
        }
        private void Filter_Changed(object sender, RoutedEventArgs e)
        {
            RefreshRequestsGrid();
        }

        private void BtnDeleteRequest_Click(object sender, RoutedEventArgs e)
        {
            if (CurrentUser?.Role != "Admin")
            {
                MessageBox.Show("Ошибка доступа: Удалять заявки может только Администратор.",
                                "alxminium Security", MessageBoxButton.OK, MessageBoxImage.Stop);
                return;
            }

            if (GridRequests.SelectedItem is WorkRequest selectedRequest)
            {
                var result = MessageBox.Show($"Вы уверены, что хотите безвозвратно удалить заявку №{selectedRequest.Id}?",
                                           "Подтверждение удаления",
                                           MessageBoxButton.YesNo,
                                           MessageBoxImage.Warning);

                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        DeleteRequestFromDb(selectedRequest.Id);

                        DataStorage.Requests.Remove(selectedRequest);

                        UpdateStatistics();
                        RefreshRequestsGrid();

                        MessageBox.Show("Заявка успешно удалена из базы! - alxminium");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Не удалось удалить: " + ex.Message);
                    }
                }
            }
            else
            {
                MessageBox.Show("Сначала выберите заявку в таблице, которую хотите удалить!");
            }
        }
        /*private void BtnImportObjects_Click(object sender, RoutedEventArgs e) #старый метод импорта объектов
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xlsm;*.xlsb;*.xls";

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    using (var workbook = new XLWorkbook(openFileDialog.FileName))
                    {
                        var worksheet = workbook.Worksheet(1);
                        var range = worksheet.RangeUsed();
                        if (range == null) return;
                        var rows = range.RowsUsed().Skip(1);

                        foreach (var row in rows)
                        {
                            DataStorage.Objects.Add(new ServiceObject
                            {
                                Id = DataStorage.Objects.Count + 1,
                                Name = row.Cell(1).GetValue<string>(),          
                                Address = row.Cell(2).GetValue<string>(),       
                                ResponsiblePerson = row.Cell(3).GetValue<string>() 
                            });
                        }
                    }
                    MessageBox.Show("Объекты успешно загружены! - alxminium");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка импорта объектов: {ex.Message}");
                }
            }
        }*/
        private void BtnImportObjects_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xlsm;*.xlsb;*.xls";

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    int addedCount = 0;
                    int duplicateCount = 0;

                    var db = new DatabaseManager();

                    using (var workbook = new ClosedXML.Excel.XLWorkbook(openFileDialog.FileName))
                    using (var conn = db.GetConnection())
                    {
                        var worksheet = workbook.Worksheet(1);
                        var range = worksheet.RangeUsed();
                        if (range == null) return;

                        var rows = range.RowsUsed().Skip(1);
                        conn.Open();

                        foreach (var row in rows)
                        {
                            string name = row.Cell(1).GetValue<string>()?.Trim();
                            string address = row.Cell(2).GetValue<string>()?.Trim();
                            string person = row.Cell(3).GetValue<string>()?.Trim();

                            if (string.IsNullOrWhiteSpace(name)) continue;

                            bool exists = DataStorage.Objects.Any(o =>
                                o.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

                            if (!exists)
                            {
                                string sql = @"INSERT INTO objects (name, address, responsible) 
                                       VALUES (@name, @address, @person);
                                       SELECT LAST_INSERT_ID();";

                                using (var cmd = new MySqlCommand(sql, conn))
                                {
                                    cmd.Parameters.AddWithValue("@name", name);
                                    cmd.Parameters.AddWithValue("@address", string.IsNullOrEmpty(address) ? (object)DBNull.Value : address);
                                    cmd.Parameters.AddWithValue("@person", string.IsNullOrEmpty(person) ? (object)DBNull.Value : person);

                                    int newId = Convert.ToInt32(cmd.ExecuteScalar());

                                    DataStorage.Objects.Add(new ServiceObject
                                    {
                                        Id = newId,
                                        Name = name,
                                        Address = address,
                                        ResponsiblePerson = person
                                    });
                                }
                                addedCount++;
                            }
                            else
                            {
                                duplicateCount++;
                            }
                        }
                    }

                    MessageBox.Show($"Загрузка завершена!\nДобавлено: {addedCount}\nПропущено дублей: {duplicateCount} - alxminium");
                    if (addedCount > 0) GridAllObjects.Items.Refresh();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка импорта объектов: {ex.Message}");
                }
            }
        }
        private void GridRequests_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (GridRequests.SelectedItem is WorkRequest selectedRequest)
            {
                if (CurrentUser.Role != "Admin")
                {
                    MessageBox.Show("У вас недостаточно прав для изменения статуса заявки. Обратитесь к администратору.",
                                    "Доступ ограничен", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (selectedRequest.Status == "Выполнена")
                {
                    selectedRequest.Status = "В работе";
                    MessageBox.Show($"Заявка №{selectedRequest.Id} возвращена в работу!");
                }
                else
                {
                    selectedRequest.Status = "Выполнена";
                    MessageBox.Show($"Заявка №{selectedRequest.Id} завершена!");
                }

                UpdateRequestStatusInDb(selectedRequest);
                UpdateStatistics();
                RefreshRequestsGrid();
            }
        }
        private void BtnDownloadObjectsTemplate_Click(object sender, RoutedEventArgs e)
        {
            string[] headers = { "name", "address", "responsible" };
            GenerateExcelTemplate("Шаблон_Объекты.xlsx", headers);
        }

        private void BtnDownloadServicesTemplate_Click(object sender, RoutedEventArgs e)
        {
            string[] headers = { "work_name", "work_type", "description", "unit", "price", "deadline_days" };
            GenerateExcelTemplate("Шаблон_Услуги.xlsx", headers);
        }

        private void GenerateExcelTemplate(string fileName, string[] headers)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    FileName = fileName,
                    DefaultExt = ".xlsx",
                    Filter = "Excel файлы (*.xlsx)|*.xlsx"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    using (SpreadsheetDocument document = SpreadsheetDocument.Create(saveFileDialog.FileName, SpreadsheetDocumentType.Workbook))
                    {
                        WorkbookPart workbookPart = document.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();

                        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                        worksheetPart.Worksheet = new Worksheet(new SheetData());

                        Sheets sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
                        Sheet sheet = new Sheet() { Id = document.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Шаблон" };
                        sheets.Append(sheet);

                        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                        Row headerRow = new Row() { RowIndex = 1 };
                        foreach (string header in headers)
                        {
                            Cell cell = new Cell() { DataType = CellValues.String, CellValue = new CellValue(header) };
                            headerRow.AppendChild(cell);
                        }
                        sheetData.AppendChild(headerRow);

                        workbookPart.Workbook.Save();
                    }
                    string folderPath = System.IO.Path.GetDirectoryName(saveFileDialog.FileName);

                    MainSnackbar.MessageQueue?.Enqueue(
                        "Шаблон сохранен!",
                        "ОТКРЫТЬ ПАПКУ",
                        () =>
                        {
                            try { System.Diagnostics.Process.Start("explorer.exe", folderPath); }
                            catch { /* тут я это вставил на случай ошибки доступа, чет не придумал лучшее решение */ }
                        }
                    );
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при генерации Excel: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            bool hasCompleted = DataStorage.Requests.Any(r => r.Status == "Выполнена");

            if (!hasCompleted)
            {
                MessageBox.Show("В архиве нет заявок со статусом 'Выполнена'. Сначала завершите работу над заявкой (смените статус), чтобы сформировать акт.",
                                "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var saveFileDialog = new SaveFileDialog { Filter = "Excel|*.xlsx", FileName = "Отчет_alxminium" };
            if (saveFileDialog.ShowDialog() == true)
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Акты выполненных работ");
                    int currentRow = 1;

                    var titleRange = worksheet.Range(currentRow, 1, currentRow, 6);
                    titleRange.Merge();
                    titleRange.Value = $"Отчет по выполненным заявкам на {DateTime.Now:dd.MM.yyyy}";
                    titleRange.Style.Font.Bold = true;
                    titleRange.Style.Font.FontSize = 16;
                    titleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    titleRange.Style.Font.FontColor = XLColor.DarkSlateGray;
                    currentRow += 2;

                    var groups = DataStorage.Requests
                        .Where(r => r.Status == "Выполнена")
                        .GroupBy(r => r.WorkType);

                    foreach (var group in groups)
                    {
                        var headerRange = worksheet.Range(currentRow, 1, currentRow, 6);
                        headerRange.Merge();
                        headerRange.Value = group.Key;
                        headerRange.Style.Font.Bold = true;
                        headerRange.Style.Font.FontSize = 14;
                        headerRange.Style.Fill.BackgroundColor = XLColor.LightSteelBlue;
                        headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                        currentRow++;

                        string[] columns = { "Наименование услуги", "Объем", "Ед. изм.", "Цена", "Сумма", "Объект" };
                        for (int i = 0; i < columns.Length; i++)
                        {
                            var cell = worksheet.Cell(currentRow, i + 1);
                            cell.Value = columns[i];
                            cell.Style.Font.Bold = true;
                            cell.Style.Fill.BackgroundColor = XLColor.WhiteSmoke;
                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        }
                        currentRow++;

                        foreach (var item in group)
                        {
                            worksheet.Cell(currentRow, 1).Value = item.WorkName;
                            worksheet.Cell(currentRow, 2).Value = item.Volume;
                            worksheet.Cell(currentRow, 3).Value = item.Unit;
                            worksheet.Cell(currentRow, 4).Value = item.Price;
                            worksheet.Cell(currentRow, 4).Style.NumberFormat.Format = "#,##0.00 ₽";

                            worksheet.Cell(currentRow, 5).Value = item.TotalCost;
                            worksheet.Cell(currentRow, 5).Style.NumberFormat.Format = "#,##0.00 ₽";

                            worksheet.Cell(currentRow, 6).Value = item.ObjectName;

                            worksheet.Range(currentRow, 1, currentRow, 6).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            worksheet.Range(currentRow, 1, currentRow, 6).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                            currentRow++;
                        }

                        worksheet.Range(currentRow, 1, currentRow, 4).Merge();
                        var footerText = worksheet.Cell(currentRow, 1);
                        footerText.Value = $"Итого по категории «{group.Key}»:";
                        footerText.Style.Font.Italic = true;
                        footerText.Style.Font.Bold = true;
                        footerText.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                        var footerSum = worksheet.Cell(currentRow, 5);
                        footerSum.Value = group.Sum(x => x.TotalCost);
                        footerSum.Style.Font.Bold = true;
                        footerSum.Style.NumberFormat.Format = "#,##0.00 ₽";

                        worksheet.Range(currentRow, 1, currentRow, 6).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                        worksheet.Range(currentRow, 1, currentRow, 6).Style.Border.BottomBorder = XLBorderStyleValues.Medium;

                        currentRow += 2;
                    }

                    worksheet.Columns().AdjustToContents();

                    workbook.SaveAs(saveFileDialog.FileName);
                }

                MessageBox.Show("Отчет успешно выгружен! - alxminium", "Экспорт завершен", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void BtnCreateRequest_Click(object sender, RoutedEventArgs e)
        {
            var selectedObject = ComboObjects.SelectedItem as ServiceObject;
            var selectedTask = ComboServices.SelectedItem as ServiceTask;

            bool isVolumeValid = double.TryParse(TxtVolume.Text, out var volVal);
            string description = TxtDescription.Text.Trim();

            if (selectedObject == null || selectedTask == null || !isVolumeValid || string.IsNullOrEmpty(description))
            {
                MessageBox.Show("Пожалуйста, заполните все поля: Объект, Услуга, Объем и Описание!",
                                "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                if (BtnCreateRequest.Tag is WorkRequest editingReq)
                {
                    editingReq.ObjectId = selectedObject.Id;
                    editingReq.ObjectName = selectedObject.Name;
                    editingReq.WorkName = selectedTask.WorkName;
                    editingReq.WorkType = selectedTask.WorkType;
                    editingReq.Unit = selectedTask.Unit;
                    editingReq.Price = selectedTask.Price;
                    editingReq.Volume = volVal;
                    editingReq.Description = description;
                    editingReq.TotalCost = editingReq.Price * (decimal)volVal;

                    UpdateRequestDetailsInDb(editingReq);

                    MessageBox.Show($"Заявка №{editingReq.Id} успешно обновлена!");

                    BtnCreateRequest.Content = "Создать заявку";
                    BtnCreateRequest.Tag = null;
                }
                else
                {
                    var newRequest = new WorkRequest
                    {
                        Id = 0,
                        Author = CurrentUser.Login,
                        Section = "Участок №1",
                        ObjectId = selectedObject.Id,
                        ObjectName = selectedObject.Name,
                        WorkName = selectedTask.WorkName,
                        WorkType = selectedTask.WorkType,
                        DeadlineDays = selectedTask.DeadlineDays,
                        Unit = selectedTask.Unit,
                        Price = selectedTask.Price,
                        Volume = volVal,
                        TotalCost = selectedTask.Price * (decimal)volVal,
                        Status = "В работе",
                        Description = description,
                        CreatedAt = DateTime.Now
                    };

                    DataStorage.Requests.Add(newRequest);
                    SaveRequestToDb(newRequest);

                    MessageBox.Show("Заявка успешно создана!");
                }

                ClearRequestInputs();
                UpdateStatistics();
                LoadRequestsFromDb();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ClearRequestInputs()
        {
            TxtVolume.Clear();
            TxtDescription.Clear();
            ComboObjects.SelectedIndex = -1;
            ComboServices.SelectedIndex = -1;
        }

        private void DeleteObjectFromDb(int objectId)
        {
            var db = new DatabaseManager();
            using (var conn = db.GetConnection())
            {
                conn.Open();
                string sql = "DELETE FROM objects WHERE id = @id";
                using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@id", objectId);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        private void BtnDeleteObject_Click(object sender, RoutedEventArgs e)
        {
            if (CurrentUser.Role != "Admin") return;

            if (GridAllObjects.SelectedItem is ServiceObject selected)
            {
                var result = MessageBox.Show($"Удалить объект \"{selected.Name}\"?\nВНИМАНИЕ: Если на этот объект завязаны заявки, база может выдать ошибку.",
                    "Удаление объекта", MessageBoxButton.YesNo, MessageBoxImage.Warning);

                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        DeleteObjectFromDb(selected.Id);
                        DataStorage.Objects.Remove(selected);
                        ComboObjects.ItemsSource = null;
                        ComboObjects.ItemsSource = DataStorage.Objects;

                        MessageBox.Show("Объект успешно удален.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка! Возможно, объект используется в заявках.\nДетали: " + ex.Message);
                    }
                }
            }
        }

        private void DeleteServiceFromDb(int serviceId)
        {
            var db = new DatabaseManager();
            using (var conn = db.GetConnection())
            {
                conn.Open();
                string sql = "DELETE FROM services WHERE id = @id";
                using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@id", serviceId);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        private void BtnDeleteService_Click(object sender, RoutedEventArgs e)
        {
            if (CurrentUser.Role != "Admin") return;

            if (GridAllServices.SelectedItem is ServiceTask selected)
            {
                var result = MessageBox.Show($"Удалить услугу \"{selected.WorkName}\"?",
                    "Удаление услуги", MessageBoxButton.YesNo, MessageBoxImage.Warning);

                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        DeleteServiceFromDb(selected.Id);
                        DataStorage.ServiceTasks.Remove(selected);
                        ComboServices.ItemsSource = null;
                        ComboServices.ItemsSource = DataStorage.ServiceTasks;

                        MessageBox.Show("Услуга удалена.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Не удалось удалить услугу: " + ex.Message);
                    }
                }
            }
        }
    }
}