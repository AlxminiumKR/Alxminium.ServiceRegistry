using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Alxminium.ServiceRegistry.Models
{
    public class ServiceObject
    {
        public int Id { get; set; }
        public string Name { get; set; }        
        public string Address { get; set; }     
        public string ResponsiblePerson { get; set; } 
    }
}