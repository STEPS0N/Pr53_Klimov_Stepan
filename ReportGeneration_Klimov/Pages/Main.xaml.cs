using ReportGeneration_Klimov.Classes.Common;
using ReportGeneration_Klimov.Models;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ReportGeneration_Klimov.Pages
{
    /// <summary>
    /// Логика взаимодействия для Main.xaml
    /// </summary>
    public partial class Main : Page
    {
        public static DbConnection connection = new DbConnection();

        public Main()
        {
            InitializeComponent();
            CreateGroupUI();
            CreateStudents();
        }

        public void CreateGroupUI()
        {
            cbGroups.Items.Clear();

            var groups = connection.Groups.ToList();

            foreach (var items in groups)
            {
                cbGroups.Items.Add(items.Name);
            }

            cbGroups.Items.Add("Выберите");
            cbGroups.SelectedIndex = cbGroups.Items.Count - 1;
        }

        public void CreateStudents(List<Student> studentsList)
        {
            Parent.Children.Clear();

            foreach (var items in studentsList)
            {
                Parent.Children.Add(new Items.Student(items, this));
            }
        }

        private void SelectGroup(object sender, SelectionChangedEventArgs e)
        {
            if (cbGroups.SelectedIndex != cbGroups.Items.Count - 1)
            {
                int IdGroup = connection.Groups.ToList().Find(x => x.Name == cbGroups.SelectedItem).Id;
                CreateStudents(connection.Students.ToList().FindAll(x => x.IdGroup == IdGroup));
            }
        }

        private void SelectStudents(object sender, KeyEventArgs e)
        {
            var students = connection.Students.ToList();

            if (cbGroups.SelectedIndex != cbGroups.Items.Count - 1)
            {
                int IdGroup = connection.Groups.ToList().Find(x => x.Name == cbGroups.SelectedItem).Id;
                students = students.FindAll(x => x.IdGroup == IdGroup);
            }
            CreateStudents(students.FindAll(x => $"{x.LastName} {x.FirstName}".Contains(tbFIO.Text)));
        }

        private void ReportGeneration(object sender, RoutedEventArgs e)
        {

        }
    }
}
