using System;
using System.Collections.ObjectModel;
using Alxminium.ServiceRegistry.Models;

namespace Alxminium.ServiceRegistry
{
    public static class DataStorage
    {
        public static ObservableCollection<ServiceObject> Objects { get; set; } = new();
        public static ObservableCollection<ServiceTask> ServiceTasks { get; set; } = new();
        public static ObservableCollection<WorkRequest> Requests { get; set; } = new();

        public static void ClearAll()
        {
            Objects.Clear();
            ServiceTasks.Clear();
            Requests.Clear();
        }
        /*public static void InitializeMockData()
        {
            Objects.Add(new ServiceObject { Id = 1, Name = "Офис Центр", Address = "ул. Ленина, 10", ResponsiblePerson = "Иванов И.И." });
            Objects.Add(new ServiceObject { Id = 2, Name = "Склад №3", Address = "промзона Южная", ResponsiblePerson = "Петров В.В." });

            ServiceTasks.Add(new ServiceTask { Id = 1, WorkName = "Ремонт принтера", WorkType = "ТО оргтехники", DeadlineDays = 7 });
            ServiceTasks.Add(new ServiceTask { Id = 2, WorkName = "Замена ламп", WorkType = "Текущий ремонт", DeadlineDays = 2 });
        }*/
    }
}