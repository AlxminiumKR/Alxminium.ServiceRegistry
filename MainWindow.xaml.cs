using Alxminium.ServiceRegistry.Models;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Win32;
using System;
using System.Linq;
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
                                Id = reader.GetInt32("id"),
                                Author = reader.GetString("author"),
                                Section = reader.GetString("section"),
                                ObjectId = reader.GetInt32("object_id"),

                                ObjectName = reader.GetString("object_display_name"),

                                WorkName = reader.GetString("frozen_name"),
                                WorkType = reader.GetString("frozen_type"),
                                DeadlineDays = reader.GetInt32("frozen_deadline"),
                                Volume = reader.GetDouble("volume"),
                                Status = reader.GetString("status"),
                                CreatedAt = reader.GetDateTime("created_at")
                            });
                        }
                    }
                }
                RefreshRequestsGrid();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке с JOIN: " + ex.Message);
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
                                DeadlineDays = reader.GetInt32("deadline_days")
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
                (author, section, object_id, frozen_name, frozen_type, frozen_deadline, volume, status) 
                VALUES (@author, @section, @objId, @name, @type, @deadline, @vol, @status);
                SELECT LAST_INSERT_ID();";

                    using (var cmd = new MySql.Data.MySqlClient.MySqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@author", req.Author);
                        cmd.Parameters.AddWithValue("@section", req.Section);
                        cmd.Parameters.AddWithValue("@objId", req.ObjectId);
                        cmd.Parameters.AddWithValue("@name", req.WorkName);
                        cmd.Parameters.AddWithValue("@type", req.WorkType);
                        cmd.Parameters.AddWithValue("@deadline", req.DeadlineDays);
                        cmd.Parameters.AddWithValue("@vol", req.Volume);
                        cmd.Parameters.AddWithValue("@status", req.Status);

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
        private void BtnImportServices_Click(object sender, RoutedEventArgs e)
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
        }
        private void RefreshRequestsGrid()
        {
            if (ChkShowCompleted.IsChecked == true)
            {
                GridRequests.ItemsSource = DataStorage.Requests;
            }
            else
            {
                GridRequests.ItemsSource = DataStorage.Requests
                    .Where(r => r.Status == "В очереди")
                    .ToList();
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
        private void BtnImportObjects_Click(object sender, RoutedEventArgs e)
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
        }
        private void GridRequests_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (GridRequests.SelectedItem is WorkRequest selectedRequest)
            {
                selectedRequest.Status = "Выполнена";
                UpdateRequestStatusInDb(selectedRequest);
                RefreshRequestsGrid();
                MessageBox.Show($"Заявка №{selectedRequest.Id} завершена! - alxminium");
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
                            worksheet.Cell(currentRow, 3).Value = item.DeadlineDays + " дн.";
                            worksheet.Cell(currentRow, 4).Value = item.ObjectName;
                            worksheet.Cell(currentRow, 5).Value = item.CreatedAt.ToString("dd/MM/yy");
                            currentRow++;
                        }

                        worksheet.Cell(currentRow, 1).Value = $"{group.Key} итого";
                        worksheet.Cell(currentRow, 1).Style.Font.Italic = true;
                        worksheet.Cell(currentRow, 5).Value = group.Sum(x => x.Volume);
                        currentRow += 2;
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

            if (selectedObject == null || selectedTask == null || string.IsNullOrEmpty(TxtVolume.Text))
            {
                MessageBox.Show("Заполните все поля!");
                return;
            }

            var newRequest = new WorkRequest
            {

                Id = 0,
                Author = "alxminium_user",
                Section = "Участок №1",
                ObjectId = selectedObject.Id,
                ObjectName = selectedObject.Name,

                WorkName = selectedTask.WorkName,
                WorkType = selectedTask.WorkType,
                DeadlineDays = selectedTask.DeadlineDays,

                Volume = double.TryParse(TxtVolume.Text, out var vol) ? vol : 0,
                Status = "В очереди",
                CreatedAt = DateTime.Now
            };

            try
            {
                DataStorage.Requests.Add(newRequest);

                SaveRequestToDb(newRequest);

                MessageBox.Show("Заявка успешно создана и данные законсервированы в MySQL!");

                RefreshRequestsGrid();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
        }
    }
}