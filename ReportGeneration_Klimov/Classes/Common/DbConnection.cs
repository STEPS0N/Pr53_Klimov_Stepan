using Microsoft.EntityFrameworkCore;
using Pomelo.EntityFrameworkCore.MySql;
using Pomelo.EntityFrameworkCore.MySql.Infrastructure;
using Pomelo.EntityFrameworkCore.MySql.Storage;
using ReportGeneration_Klimov.Models;
using System;
using System.Windows;


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
            optionsBuilder.UseMySql("server=localhost;port=3307;database=Stud;uid=root;pwd=;",
                mySqlOptions => mySqlOptions.ServerVersion("8.0.11-mysql"));
        }
    }
}
