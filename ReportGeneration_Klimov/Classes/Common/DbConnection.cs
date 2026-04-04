using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.EntityFrameworkCore;
using Pomelo.EntityFrameworkCore.MySql.Storage;
using ReportGeneration_Klimov.Models;

namespace ReportGeneration_Klimov.Classes.Common
{
    public class DbConnection : DbContext
    {
        public DbSet<Discipline> Disciplines { get; set; }
        public DbSet<Evaluation> Evaluations { get; set; }
        public DbSet<Group> Groups { get; set; }
        public DbSet<Student> Students { get; set; }
        public DbSet<Work> Works { get; set; }

        public DbConnection()
        {
            try
            {
                Database.EnsureCreated();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseMySql("server=localhost;port=3307;database=Stud;uid=root;pwd=;", new ServerVersion(new Version(8, 0, 11)));
        }
    }
}
