using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGeneration_Klimov.Models
{
    public class Evaluation
    {
        [Key]
        public int Id { get; set; }
        public int IdWork { get; set; }
        public int IdStudent { get; set; }
        public string Value { get; set; }
        public string Lateness { get; set; }
    }
}
