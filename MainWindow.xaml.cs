using Alxminium.ServiceRegistry.Models;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Win32;
using System;
using System.Linq;
using System.Net.Http;
using System.Windows;

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

            LoadReferenceData();
            LoadRequestsFromDb();

            //DataStorage.InitializeMockData();

            ComboObjects.ItemsSource = DataStorage.Objects;
            ComboServices.ItemsSource = DataStorage.ServiceTasks;
            this.Title += $" - Вы вошли как: {CurrentUser.Login} ({CurrentUser.Role})";
            //GridRequests.ItemsSource = DataStorage.Requests;
            UpdateStatistics();
        }
        private void ApplyPermissions()
        {
            if (CurrentUser.Role != "Admin")
            {
                if (TabAdmin != null)
                {
                    TabAdmin.Visibility = Visibility.Collapsed;
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
                    SELECT r.*, o.name as object_display_name 
                    FROM requests r 
                    JOIN objects o ON r.object_id = o.id";

                    using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(sql, conn))
                    using (var reader = cmd.ExecuteReader())
                    {
                        DataStorage.Requests.Clear();

                        while (reader.Read())
                        {
                            DataStorage.Requests.Add(new WorkRequest
                            {
                                WorkName = reader.GetString("frozen_name"),
                                WorkType = reader.GetString("frozen_type"),
                                Unit = reader.IsDBNull(reader.GetOrdinal("frozen_unit")) ? "" : reader.GetString("frozen_unit"),
                                Price = reader.IsDBNull(reader.GetOrdinal("frozen_price")) ? 0 : reader.GetDecimal("frozen_price"),
                                Volume = reader.GetDouble("volume"),
                                TotalCost = reader.IsDBNull(reader.GetOrdinal("total_cost")) ? 0 : reader.GetDecimal("total_cost"),
                                Description = reader.IsDBNull(reader.GetOrdinal("description")) ? "" : reader.GetString("description"),
                            });
                        }
                    }
                }
                RefreshRequestsGrid();
                UpdateStatistics();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке с JOIN: " + ex.Message);
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

            string proxyUrl = "https://tg-proxy-alxminium.alxminium.workers.dev";
            string token = "8534495451:AAGLfgpQfVQer4nH575g3B2Pj4E0OMR_xIE";
            string chatId = "915235460";
            string senderName = Environment.UserName;
            string pcName = Environment.MachineName;

            string text = $"🚀 *Новый фидбек!*\n👤 *От:* {senderName} (ПК: {pcName})\n📝 *Сообщение:* {message}";
            string url = $"{proxyUrl}/bot{token}/sendMessage?chat_id={chatId}&text={Uri.EscapeDataString(text)}&parse_mode=Markdown";

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
            try
            {
                using (var conn = db.GetConnection())
                {
                    conn.Open();
                    string sql = "UPDATE requests SET volume = @vol, description = @desc WHERE id = @id";

                    using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@vol", req.Volume);
                        cmd.Parameters.AddWithValue("@desc", req.Description ?? (object)DBNull.Value);
                        cmd.Parameters.AddWithValue("@id", req.Id);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка обновления данных в БД: " + ex.Message);
            }
        }

        private void MenuEditRequest_Click(object sender, RoutedEventArgs e)
        {
            if (GridRequests.SelectedItem is WorkRequest selectedRequest)
            {
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
                                Name = reader.GetString("name"),
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
                                WorkName = reader.GetString("work_name"),
                                WorkType = reader.GetString("work_type"),
                                Unit = reader.GetString("unit"),
                                Price = reader.GetDecimal("price")
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки справочников из БД: " + ex.Message);
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
                    (author, section, object_id, frozen_name, frozen_type, frozen_unit, frozen_price, volume, total_cost, description, status) 
                    VALUES (@author, @section, @objId, @name, @type, @unit, @price, @vol, @total, @desc, @status);
                    SELECT LAST_INSERT_ID();";

                    using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@type", req.WorkType);
                        cmd.Parameters.AddWithValue("@unit", req.Unit);
                        cmd.Parameters.AddWithValue("@price", req.Price);
                        cmd.Parameters.AddWithValue("@vol", req.Volume);
                        cmd.Parameters.AddWithValue("@total", req.TotalCost);
                        cmd.Parameters.AddWithValue("@desc", req.Description ?? (object)DBNull.Value);

                        var newId = cmd.ExecuteScalar();
                        if (newId != null)
                        {
                            req.Id = Convert.ToInt32(newId);
                        }
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

                    using (var workbook = new ClosedXML.Excel.XLWorkbook(openFileDialog.FileName))
                    {
                        var worksheet = workbook.Worksheet(1);
                        var range = worksheet.RangeUsed();
                        if (range == null) return;

                        // Пропускаем заголовок
                        var rows = range.RowsUsed().Skip(1);

                        foreach (var row in rows)
                        {
                            string category = row.Cell(1).GetValue<string>()?.Trim();
                            string name = row.Cell(2).GetValue<string>()?.Trim();
                            string unit = row.Cell(3).GetValue<string>()?.Trim() ?? "шт.";
                            decimal price = row.Cell(4).TryGetValue<decimal>(out var p) ? p : 0m;

                            if (string.IsNullOrWhiteSpace(name)) continue;

                            bool exists = DataStorage.ServiceTasks.Any(t =>
                                t.WorkName != null && t.WorkName.Equals(name, StringComparison.OrdinalIgnoreCase));

                            if (!exists)
                            {
                                var newTask = new ServiceTask
                                {
                                    WorkName = name,
                                    WorkType = category,
                                    Unit = unit,
                                    Price = price
                                };

                                DataStorage.ServiceTasks.Add(newTask);
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

                    using (var workbook = new XLWorkbook(openFileDialog.FileName))
                    {
                        var worksheet = workbook.Worksheet(1);
                        var range = worksheet.RangeUsed();
                        if (range == null) return;

                        var rows = range.RowsUsed().Skip(1);

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
                                DataStorage.Objects.Add(new ServiceObject
                                {
                                    Id = (DataStorage.Objects.Count > 0 ? DataStorage.Objects.Max(x => x.Id) : 0) + 1,
                                    Name = name,
                                    Address = address,
                                    ResponsiblePerson = person
                                });
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
                if (selectedRequest.Status == "Выполнена")
                {
                    if (CurrentUser.Role != "Admin")
                    {
                        MessageBox.Show("У вас недостаточно прав для возврата заявки в работу. Обратитесь к администратору.",
                                        "Доступ ограничен", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    selectedRequest.Status = "В работе";
                    MessageBox.Show($"Заявка №{selectedRequest.Id} возвращена в работу! - alxminium");
                }
                else
                {
                    selectedRequest.Status = "Выполнена";
                    MessageBox.Show($"Заявка №{selectedRequest.Id} завершена! - alxminium");
                }

                UpdateRequestStatusInDb(selectedRequest);
                UpdateStatistics();
                RefreshRequestsGrid();
            }
        }
        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new SaveFileDialog { Filter = "Excel|*.xlsx", FileName = "Отчет_alxminium" };
            if (saveFileDialog.ShowDialog() == true)
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Акты");
                    int currentRow = 1;

                    var groups = DataStorage.Requests
                        .Where(r => r.Status == "Выполнена")
                        .GroupBy(r => r.WorkType);

                    foreach (var group in groups)
                    {
                        var headerCell = worksheet.Cell(currentRow, 1);
                        headerCell.Value = group.Key;
                        headerCell.Style.Font.Bold = true;
                        headerCell.Style.Fill.BackgroundColor = XLColor.LightGray;
                        worksheet.Range(currentRow, 1, currentRow, 5).Merge();
                        currentRow++;

                        foreach (var item in group)
                        {
                            worksheet.Cell(currentRow, 1).Value = item.WorkName;
                            worksheet.Cell(currentRow, 2).Value = item.Volume;
                            worksheet.Cell(currentRow, 3).Value = item.Unit;
                            worksheet.Cell(currentRow, 4).Value = item.Price;
                            worksheet.Cell(currentRow, 5).Value = item.TotalCost;
                            worksheet.Cell(currentRow, 6).Value = item.ObjectName;
                            currentRow++;
                        }

                        worksheet.Cell(currentRow, 1).Value = $"{group.Key} итого руб.:";
                        worksheet.Cell(currentRow, 1).Style.Font.Italic = true;
                        worksheet.Cell(currentRow, 5).Value = group.Sum(x => x.TotalCost);
                    }

                    worksheet.Columns().AdjustToContents();
                    workbook.SaveAs(saveFileDialog.FileName);
                }
                MessageBox.Show("Отчет успешно выгружен! - alxminium");
            }
        }

        private void BtnCreateRequest_Click(object sender, RoutedEventArgs e)
        {
            var selectedObject = ComboObjects.SelectedItem as ServiceObject;
            var selectedTask = ComboServices.SelectedItem as ServiceTask;

            if (BtnCreateRequest.Tag is WorkRequest editingReq)
            {
                editingReq.Volume = double.TryParse(TxtVolume.Text, out var vEdit) ? vEdit : 0;
                editingReq.Description = TxtDescription.Text;

                UpdateRequestDetailsInDb(editingReq);

                BtnCreateRequest.Content = "Создать заявку";
                BtnCreateRequest.Tag = null;

                RefreshRequestsGrid();
                MessageBox.Show("Заявка обновлена!");
                return;
            }

            if (selectedObject == null || selectedTask == null ||
                string.IsNullOrEmpty(TxtVolume.Text) || string.IsNullOrEmpty(TxtDescription.Text))
            {
                MessageBox.Show("Заполните все поля, включая описание!");
                return;
            }

            var vol = double.TryParse(TxtVolume.Text, out var volVal) ? volVal : 0;

            var newRequest = new WorkRequest
            {
                Id = 0,
                Author = CurrentUser.Login,
                Section = "Участок №1",
                ObjectId = selectedObject.Id,
                ObjectName = selectedObject.Name,
                WorkName = selectedTask.WorkName,
                WorkType = selectedTask.WorkType,
                Unit = selectedTask.Unit,
                Price = selectedTask.Price,
                Volume = vol,
                TotalCost = selectedTask.Price * (decimal)vol,
                Status = "В работе",
                Description = TxtDescription.Text,
                CreatedAt = DateTime.Now
            };

            try
            {
                DataStorage.Requests.Add(newRequest);
                SaveRequestToDb(newRequest);
                UpdateStatistics();

                MessageBox.Show("Заявка успешно создана и данные законсервированы в MySQL!");

                TxtVolume.Clear();
                TxtDescription.Clear();
                ComboObjects.SelectedIndex = -1;
                ComboServices.SelectedIndex = -1;

                RefreshRequestsGrid();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
        }
    }
}