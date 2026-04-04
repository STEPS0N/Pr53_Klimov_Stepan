using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGeneration_Klimov.Models
{
    public class Work
    {
        [Key]
        public int Id { get; set; }
        public int IdDiscipline { get; set; }
        public int IdType { get; set; }
        public DateTime Date { get; set; }
        public string Name { get; set; }
        public int Semester { get; set; }
    }
}
