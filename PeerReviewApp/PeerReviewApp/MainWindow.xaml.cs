using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace PeerReviewApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Dictionary<string, Student> students;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void uxButton_LoadStudents_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            if ((bool)open.ShowDialog())
            {
                Excel.Application xlApp = null;
                Excel.Workbook xlWorkbook = null;
                Excel._Worksheet sheet = null;
                Excel.Range xlRange = null;
                try
                {
                    students = new Dictionary<string, Student>();
                    xlApp = new Excel.Application();
                    xlWorkbook = xlApp.Workbooks.Open(open.FileName);
                    sheet = xlWorkbook.Sheets[1];
                    xlRange = sheet.UsedRange;
                    object[,] values = xlRange.Value2;
                    for (int i = 2; i <= xlRange.Rows.Count; i++)
                    {
                        students[values[i, (int)Roster.Email].ToString()] = new Student(
                            values[i, (int)Roster.FirstName].ToString() + " " + values[i, (int)Roster.LastName].ToString());

                    }
                    uxButton_LoadSurvey.IsEnabled = true;
                    uxButton_SaveResults.IsEnabled = false;
                    MessageBox.Show("Loaded " + students.Count + " students from the roster.");
                }

                catch (Exception err)
                {
                    MessageBox.Show(err.Message);
                }
                finally
                {
                    CleanUp(xlApp, xlWorkbook, xlRange, sheet);
                }
            }
        }

        private void uxButton_LoadSurvey_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            if ((bool)open.ShowDialog())
            {
                Excel.Application xlApp = null;
                Excel.Workbook xlWorkbook = null;
                Excel._Worksheet sheet = null;
                Excel.Range xlRange = null;
                try
                {
                    xlApp = new Excel.Application();
                    xlWorkbook = xlApp.Workbooks.Open(open.FileName);
                    sheet = xlWorkbook.Sheets[1];
                    xlRange = sheet.UsedRange;
                    object[,] values = xlRange.Value2;
                    int colStart = Convert.ToInt32(uxSurveyColSkip.Text) + 1;
                    int maxGroupSize = Convert.ToInt32(uxMaxGroupSize.Text);
                    int questionCount = Convert.ToInt32(uxQuestionCount.Text);

                    int TeamScoreStartIndex = colStart + maxGroupSize + questionCount;
                    int selfCommentIndex = colStart + maxGroupSize + questionCount * maxGroupSize;
                    int teamEmailEndIndex = colStart + maxGroupSize - 1;

                    for (int row = Convert.ToInt32(uxSurveyRowSkip.Text) + 1; row <= xlRange.Rows.Count; row++)
                    {
                        if (values[row, colStart] != null)
                        {
                            string email = values[row, colStart].ToString().Trim().ToLower();
                            if (email == "" || email == "blank")
                            {
                                //Console.WriteLine("Blank email: " + row);
                            }
                            else if (!students.ContainsKey(email))
                            {
                                Console.WriteLine("No matching student email: " + email);
                            }
                            else
                            {
                                //self scores
                                students[email].CompletedSurvey = true;
                                for (int k = colStart + maxGroupSize; k < TeamScoreStartIndex; k++)
                                {
                                    students[email].SelfScores.Add(Convert.ToInt32(values[row, k]));
                                }
                                students[email].CommentsByStudent.Append(values[row, selfCommentIndex]);

                                //teammate scores
                                for (int k = colStart + 1; k <= teamEmailEndIndex; k++)
                                {
                                    string teamEmail = values[row, k].ToString().Trim().ToLower();
                                    if (teamEmail == "" || teamEmail == "blank")
                                    {
                                        //Console.WriteLine("Blank email: " + row + "," + k);
                                    }
                                    else if (!students.ContainsKey(teamEmail))
                                    {
                                        Console.WriteLine("No matching student email: " + teamEmail);
                                    }
                                    else
                                    {
                                        students[teamEmail].TeamScores[email] = new List<int>();
                                        int qStart = (TeamScoreStartIndex) + ((k - colStart - 1) * questionCount);
                                        for (int q = qStart; q < qStart + questionCount; q++)
                                        {
                                            string score = values[row, q].ToString();
                                            if (score != "" || score != "blank")
                                                students[teamEmail].TeamScores[email].Add(Convert.ToInt32(score));
                                        }
                                        students[teamEmail].CommentsByTeam.Append(students[email].Name);
                                        students[teamEmail].CommentsByTeam.Append(":");
                                        students[teamEmail].CommentsByTeam.Append(values[row, selfCommentIndex]);
                                        students[teamEmail].CommentsByTeam.Append(",");
                                    }
                                }
                            }
                        }

                    }
                    MessageBox.Show("Loaded survey.");
                    uxButton_SaveResults.IsEnabled = true;
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message);
                }

                finally
                {
                    CleanUp(xlApp, xlWorkbook, xlRange, sheet);
                }
            }
        }

        private void uxButton_SaveResults_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            if ((bool)open.ShowDialog())
            {
                Excel.Application xlApp = null;
                Excel.Workbook xlWorkbook = null;
                Excel._Worksheet sheet = null;
                Excel.Range xlRange = null;
                try
                {
                    xlApp = new Excel.Application();
                    xlWorkbook = xlApp.Workbooks.Open(Filename: open.FileName, ReadOnly: false);
                    sheet = xlWorkbook.Sheets[1];
                    xlRange = sheet.UsedRange;
                    object[,] values = xlRange.Value2;

                    int questionCount = Convert.ToInt32(uxQuestionCount.Text);

                    for (int row = 3; row <= xlRange.Rows.Count; row++)
                    {
                        string email = values[row, 4].ToString().Trim().ToLower();
                        if (students.ContainsKey(email))
                        {
                            sheet.Cells[row, 8] = students[email].CompletedSurvey;
                            int col = 9;
                            if (students[email].SelfScores.Count != questionCount)
                            {
                                Console.WriteLine("Student did not fully evaluate theirself: " + email);
                            }
                            else
                            {
                                for (int q = 0; q < questionCount; q++, col++)
                                {
                                    sheet.Cells[row, col] = students[email].SelfScores[q];
                                }
                            }
                            col = 9+ questionCount;
                            for (int q = 0; q < questionCount; q++, col++)
                            {
                                List<double> totals = new List<double>();
                                
                                foreach (List<int> scores in students[email].TeamScores.Values)
                                {
                                    if (scores.Count == questionCount)
                                    {
                                        totals.Add(scores[q]);
                                    }
                                    else
                                    {
                                        Console.WriteLine("Student did not fully evaluatev a team member: " + email);
                                    }
                                }
                                if (totals.Count > 0)
                                    sheet.Cells[row, col] = Math.Round(totals.Average(), 2);                                
                            }
                            col += 5;
                            sheet.Cells[row, col] = students[email].CommentsByStudent.ToString();
                            col++;
                            sheet.Cells[row, col] = students[email].CommentsByTeam.ToString();
                        }
                        else
                        {
                            Console.WriteLine("Student does not match roster: " + email);
                        }
                    }
                    MessageBox.Show("Finshed outputting results.");
                    uxButton_SaveResults.IsEnabled = true;
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.StackTrace.ToString());
                }

                finally
                {
                    CleanUp(xlApp, xlWorkbook, xlRange, sheet);
                }
            }
        }


        private void CleanUp(Excel.Application xlApp, Excel.Workbook xlWorkbook, Excel.Range xlRange, Excel._Worksheet sheet)
        {
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);

            Marshal.ReleaseComObject(sheet);

            //close and release
            if (xlWorkbook != null)
                xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            if (xlApp != null)
                xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

    }
}
