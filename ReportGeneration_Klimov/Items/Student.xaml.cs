using ReportGeneration_Klimov.Pages;
using System.Linq;
using System.Windows.Controls;
using ReportGeneration_Klimov.Models;
using System.Collections.Generic;
using System;

namespace ReportGeneration_Klimov.Items
{
    /// <summary>
    /// Логика взаимодействия для Student.xaml
    /// </summary>
    public partial class Student : UserControl
    {
        public Models.Student student;
        public Main main;

        public Student(Models.Student student, Main main)
        {
            InitializeComponent();
            this.student = student;
            this.main = main;

            tbFIO.Text = $"{student.LastName} {student.FirstName}";
            cbExpelled.IsChecked = student.Expelled;

            List<Discipline> studentDisciplines = Main.connection.Disciplines.ToList().FindAll(
                x => x.IdGroup == student.IdGroup);

            int NecessarilyCount = 0;
            int WorksCount = 0;
            int DoneCount = 0;
            int MissedCount = 0;

            foreach (Discipline students in studentDisciplines)
            {
                List<Work> StudentWorks = Main.connection.Works.ToList().FindAll(x =>
                    (x.IdType == 1 || x.IdType == 2 || x.IdType == 3) &&
                    x.IdDiscipline == students.Id);
                NecessarilyCount += StudentWorks.Count;

                foreach (Work StudentWork in StudentWorks)
                {
                    Evaluation Evaluation = Main.connection.Evaluations.ToList().Find(x =>
                        x.IdWork == StudentWork.Id &&
                        x.IdStudent == student.Id);
                    if (Evaluation != null && Evaluation.Value.Trim() != "" && Evaluation.Value.Trim() != "2")
                        DoneCount++;
                }

                StudentWorks = Main.connection.Works.ToList().FindAll(x =>
                    x.IdType != 4 && x.IdType != 3);
                WorksCount += StudentWorks.Count;

                foreach (Work StudentWork in StudentWorks)
                {
                    Evaluation Evaluation = Main.connection.Evaluations.ToList().Find(x =>
                        x.IdWork == StudentWork.Id &&
                        x.IdStudent == student.Id);
                    if (Evaluation != null && Evaluation.Lateness.Trim() != "")
                        MissedCount += Convert.ToInt32(Evaluation.Lateness);
                }
            }
            doneWorks.Value = (100f / (float)NecessarilyCount) * ((float)DoneCount);
            missedCount.Value = (100f / ((float)WorksCount * 90f)) * ((float)MissedCount);
            tbGroup.Text = Main.connection.Groups.ToList().Find(x => x.Id == student.IdGroup).Name;
        }
    }
}
