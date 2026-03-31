using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Alxminium.ServiceRegistry.Models
{
    public class ServiceTask
    {
        public int Id { get; set; }
        public string WorkName { get; set; }
        public string WorkType { get; set; }
        public string Description { get; set; }
        public string Unit { get; set; }
        public decimal Price { get; set; }
        public int DeadlineDays { get; set; }
    }
}