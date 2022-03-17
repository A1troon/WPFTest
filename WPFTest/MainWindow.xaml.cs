using System;
using System.Windows;

namespace WPFTest
{

    public partial class MainWindow : Window
    {
        public static int ColVoprt; Test[] myTest;
        int i, k;
        public static int pv; public MainWindow()
        {
            InitializeComponent();
        }

        private void btnNext_Click(object sender, RoutedEventArgs e)
        {
            if (myTest[i - 1].PV == pv) k++;
            MethodNext(i);

            if (i > ColVoprt - 2)
            {
                btnNext.IsEnabled = false;
                MessageBox.Show(k.ToString());
            }
            i++;
        }

        private void Window_Initialized(object sender, EventArgs e)
        {
            ReadExcelMethod();

            i = 1;
            k = 0;

            MethodNext(0);

        }

        private void ReadExcelMethod()
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook excelBook =
            excelApp.Workbooks.Open(@"E:\KFU\WPFTest\WPFTest\myTest.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet =(Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(1); ;
            Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;

            ColVoprt = excelRange.Rows.Count; myTest = new Test[ColVoprt + 1]; pv = 0;
            for (i = 0; i < ColVoprt; i++)
            {
                myTest[i] = new Test();
            }

            for (int i = 1; i < ColVoprt; i++)
            {
                myTest[i - 1].Q = Convert.ToString((excelRange.Cells[i + 1, 1] as Microsoft.Office.Interop.Excel.Range).Value2);
                myTest[i - 1].V1 = Convert.ToString((excelRange.Cells[i + 1, 2] as Microsoft.Office.Interop.Excel.Range).Value2);
                myTest[i - 1].V2 = Convert.ToString((excelRange.Cells[i + 1, 3] as Microsoft.Office.Interop.Excel.Range).Value2);
                myTest[i - 1].V3 = Convert.ToString((excelRange.Cells[i + 1, 4] as Microsoft.Office.Interop.Excel.Range).Value2);
                myTest[i - 1].PV = Convert.ToInt32((excelRange.Cells[i + 1, 5] as Microsoft.Office.Interop.Excel.Range).Value2);

            }

            excelBook.Close(true, null, null); excelApp.Quit();
        }

        private void Rb1_Checked(object sender, RoutedEventArgs e)
        {
            pv = 1;
        }

        private void Rb2_Checked(object sender, RoutedEventArgs e)
        {
            pv = 2;
        }

        private void Rb3_Checked(object sender, RoutedEventArgs e)
        {
            pv = 3;
        }

        private void MethodNext(int i)
        {
            lbl1.Content = myTest[i].Q; Rb1.Content = myTest[i].V1; Rb2.Content = myTest[i].V2; Rb3.Content = myTest[i].V3;
        }
    }
}
