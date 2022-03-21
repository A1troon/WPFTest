using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;
namespace WPFTest
{

    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        static int countQ;
        StackPanel[] Panels;
        Test[] tests;
        Label lblTime;
        Button btn;
        DispatcherTimer timer;
        private void Window_Initialized(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            try
            {
                Excel.Workbook excelBook = excelApp.get_Workbooks().Open(@"D:\c#projects\WPFTest\WPFTest\myTest.xlsx", 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel.Worksheet excelSheet = (Excel.Worksheet)excelBook.get_ActiveSheet();
                Excel.Range excelRange = excelSheet.UsedRange;

                countQ = excelRange.Rows.Count;

                Panels = new StackPanel[countQ];
                tests = new Test[countQ];

                timer = new DispatcherTimer();
                lblTime = new Label();
                lblTime.FontSize = 48;
                lblTime.HorizontalAlignment = HorizontalAlignment.Center;
                lblTime.VerticalAlignment = VerticalAlignment.Center;
                timer.Interval = TimeSpan.FromSeconds(1);
                timer.Tick += timer_Tick;
                timer.Start();
                
                lblTime.Content = "300";
                MainPanel.Children.Add(lblTime);

                for (int i = 0; i < countQ; i++)
                {
                    Panels[i] = new StackPanel();
                    Panels[i].Orientation = Orientation.Vertical;
                    Panels[i].Background = new SolidColorBrush(Colors.CadetBlue);
                    tests[i] = new Test(
                        Convert.ToString(excelSheet.Cells[i + 1, 1].Value),
                        Convert.ToInt32(excelSheet.Cells[i + 1, 7].Value),
                        new string[] {
                        Convert.ToString(excelSheet.Cells[i + 1, 2].Value),
                        Convert.ToString(excelSheet.Cells[i + 1, 3].Value),
                        Convert.ToString(excelSheet.Cells[i + 1, 4].Value),
                        Convert.ToString(excelSheet.Cells[i + 1, 5].Value),
                        Convert.ToString(excelSheet.Cells[i + 1, 6].Value)
                        }  
                        );
                    Label label = new Label();
                    label.Content = tests[i].Question;
                    Panels[i].Children.Add(label);
                    for(int j=0; j<tests[i].V.Length; j++)
                    {
                        RadioButton radioButton = new RadioButton();
                        radioButton.Content = tests[i].V[j];
                        radioButton.FontSize = 20;
                        radioButton.Name = "Q" + Convert.ToString(j+1);
                        radioButton.GroupName = Convert.ToString(i);
                        Panels[i].Children.Add(radioButton);

                    }
                    Panels[i].Name += "Q"+Convert.ToString(i);
                    Panels[i].AddHandler(RadioButton.CheckedEvent, new RoutedEventHandler(RadioButton_Click));

                    Panels[i].Margin = new Thickness(3, 5, 0, 0);
                    MainPanel.Children.Add(Panels[i]);
                }


                


                btn = new Button();
                btn.Content = "Ok";
                btn.Width = 85;
                btn.Margin = new Thickness(10, 10, 10, 10);
                btn.HorizontalAlignment = HorizontalAlignment.Right;
                btn.Click += Btn_Click;
                MainPanel.Children.Add(btn);
                excelBook.Close(true, null, null);
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                excelApp.Quit();
            }
        }
        void timer_Tick(object sender, EventArgs e)
        {
            int t = Int32.Parse(lblTime.Content.ToString());
            t = t - 1;
            lblTime.Content = t.ToString();
            if (t <= 0)
            {
                Btn_Click(btn, null);
            }
        }

        private void RadioButton_Click(object sender, RoutedEventArgs e)
        { 
            tests[Convert.ToInt32(Convert.ToString(((StackPanel)sender).Name[1]))].AnswerChecked=Convert.ToInt32(Convert.ToString(((RadioButton)e.Source).Name[1]));
        }


        private void Btn_Click(object sender, RoutedEventArgs e)
        {
            timer.Stop();
            ((Button)sender).Visibility = Visibility.Hidden;
            int counter = 0;
            Answered.Content += "Отвеченные:";
            Correct.Content += "Правильные:";
            foreach (var test in tests)
            {
                if (test.AnswerChecked != 0)
                {
                    Answered.Content += "\n" + test.V[test.AnswerChecked - 1];
                    Correct.Content += "\n" + test.V[test.Answer - 1];
                }
                else
                {
                    Answered.Content += "\n-";
                    Correct.Content += "\n" + test.V[test.Answer - 1];
                }

                if (test.AnswerChecked == test.Answer)
                    counter++;
            }
            Rating.Content += "\n" + counter + "/" + tests.Length + "Верно\n";
            if (Convert.ToDouble(counter) / tests.Length < 0.56)
                Rating.Content += "Неудовлетворительно";
            else if (Convert.ToDouble(counter) / tests.Length < 0.71)
                Rating.Content += "Удовлетворительно";
            else if (Convert.ToDouble(counter) / tests.Length < 0.85)
                Rating.Content += "Хорошо";
            else
                Rating.Content += "Великолепно";
        }

       
    }
}
