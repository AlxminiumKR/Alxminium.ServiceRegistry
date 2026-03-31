using System;

namespace Alxminium.ServiceRegistry.Models
{
    public class WorkRequest
    {
        public int Id { get; set; }
        public string Author { get; set; }
        public string Section { get; set; }
        public int ObjectId { get; set; }
        public string ObjectName { get; set; }
        public string WorkName { get; set; }
        public string WorkType { get; set; }

        public string Unit { get; set; } 
        public decimal Price { get; set; } 
        public double Volume { get; set; }
        public decimal TotalCost { get; set; }

        public string Status { get; set; }
        public string Description { get; set; }
        public DateTime CreatedAt { get; set; } = DateTime.Now;
    }
}