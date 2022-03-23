using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
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
using Newtonsoft.Json;
using Srez.Models;
using Excel = Microsoft.Office.Interop.Excel;
using word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.IO;

namespace Srez
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        public async void LoadSales()
        {
            using (HttpClient client = new HttpClient())
            {
                //var values = new Dictionary<string, string>
                //{
                //    {"dateStart",DpOt.SelectedDate.ToString()},
                //    {"dateEnd",DpDo.SelectedDate.ToString()}
                //};
                //var content = new FormUrlEncodedContent(values);

                //var jsonText = JsonConvert.SerializeObject(DpOt.SelectedDate.ToString());
                //var jsonData = Encoding.UTF8.GetBytes(jsonText);
                ////var content = new StringContent($"dateStart={DpOt.SelectedDate.ToString()}", Encoding.UTF8, "application/json");
                //var response = await client.PostAsync($"https://localhost:7100/api/Sale",jsonData);
                //response.EnsureSuccessStatusCode();
            }
        }

        private void BtnData_Click(object sender, RoutedEventArgs e)
        {
            if (DpOt.SelectedDate == null || DpDo.SelectedDate == null)
            {
                MessageBox.Show("Выберите начальную и конечную дату!");
                return;
            }
            if(DpOt.SelectedDate > DpDo.SelectedDate)
            {
                MessageBox.Show("Дата начала не может быть позже чем дата окончания!");
            }
            //LoadSales();
            using (WebClient client = new WebClient())
            {
                Sale sale = new Sale();
                client.Encoding = Encoding.UTF8;
                client.Headers.Add("Content-Type", "application/json");
                var jsonText = JsonConvert.SerializeObject(DpOt.SelectedDate.ToString());
                var jsonData = Encoding.UTF8.GetBytes(jsonText);
                string ad = "https://localhost:7100/api/Sale";
                client.UploadData(ad, WebRequestMethods.Http.Post, jsonData);
                //List<Sale> sales = JsonConvert.DeserializeObject<List<Sale>>(data);
                //LvSales.ItemsSource = sales;
            }
        }

        private void btnWordotch_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Text files(*.docx)|*.docx|All files(*.*)|*.*";
            string source = $@"{Directory.GetCurrentDirectory()}\Шаблон отчета по продажам.docx";
            word.Application app = new word.Application();
            word.Document doc = app.Documents.Open(source);
            doc.Activate();

            try
            {
                if (sfd.ShowDialog() == false)
                {
                    doc.Close();
                    doc = null;
                    app.Quit();
                    return;
                }

                doc.SaveAs2(sfd.FileName);
                doc.Close();
                doc = null;
                app.Quit();
                MessageBox.Show("Файл успешно создан");
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                doc.Close();
                doc = null;
                app.Quit();
            }
        }

        private void BtnExelOtch_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Text files(*.xlsx)|*.xlsx|All files(*.*)|*.*";
            try

            {
                if (sfd.ShowDialog() == false)
                {
                    return;
                }
                
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workBook;
                Excel.Worksheet workSheet;
                workBook = excelApp.Workbooks.Add();
                workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);
                workSheet.Columns.EntireColumn.AutoFit();
                //Excel.Range range = workSheet.get_Range("A1", "C12");
                //range.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
               
                //Excel.ChartObjects xlCharts = (Excel.ChartObjects)workSheet.ChartObjects(Type.Missing);
                //Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(450, 0, 500, 250);
                //Excel.Chart chart = myChart.Chart;
                //Excel.SeriesCollection seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection(Type.Missing);
                //Excel.Series series = seriesCollection.NewSeries();
                //int colx = 1 + stat.Count();
                //int coly = 1 + stat.Count();
                //series.XValues = workSheet.get_Range("A2", "A" + colx);
                //series.Values = workSheet.get_Range("B2", "B" + coly);
                //chart.ChartType = Excel.XlChartType.xl3DColumnStacked;
                //chart.HasTitle = true;
                excelApp.Application.ActiveWorkbook.SaveAs(sfd.FileName);
                excelApp.Quit();
                MessageBox.Show("Файл успешно сформирован");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void BtnWordChek_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Text files(*.doc)|*.doc|All files(*.*)|*.*";
            string source = $@"{Directory.GetCurrentDirectory()}\товарный чек.doc";
            word.Application app = new word.Application();
            word.Document doc = app.Documents.Open(source);
            doc.Activate();

            try
            {
                if (sfd.ShowDialog() == false)
                {
                    doc.Close();
                    doc = null;
                    app.Quit();
                    return;
                }

                doc.SaveAs2(sfd.FileName);
                doc.Close();
                doc = null;
                app.Quit();
                MessageBox.Show("Файл успешно создан");
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                doc.Close();
                doc = null;
                app.Quit();
            }
        }

        private void BtnExelChek_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Text files(*.xlsx)|*.xlsx|All files(*.*)|*.*";
            try

            {
                if (sfd.ShowDialog() == false)
                {
                    return;
                }

                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workBook;
                Excel.Worksheet workSheet;
                workBook = excelApp.Workbooks.Add();
                workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);
                workSheet.Columns.EntireColumn.AutoFit();
                //Excel.Range range = workSheet.get_Range("A1", "C12");
                //range.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;

                //Excel.ChartObjects xlCharts = (Excel.ChartObjects)workSheet.ChartObjects(Type.Missing);
                //Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(450, 0, 500, 250);
                //Excel.Chart chart = myChart.Chart;
                //Excel.SeriesCollection seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection(Type.Missing);
                //Excel.Series series = seriesCollection.NewSeries();
                //int colx = 1 + stat.Count();
                //int coly = 1 + stat.Count();
                //series.XValues = workSheet.get_Range("A2", "A" + colx);
                //series.Values = workSheet.get_Range("B2", "B" + coly);
                //chart.ChartType = Excel.XlChartType.xl3DColumnStacked;
                //chart.HasTitle = true;
                excelApp.Application.ActiveWorkbook.SaveAs(sfd.FileName);
                excelApp.Quit();
                MessageBox.Show("Файл успешно сформирован");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
