using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
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
using Excel = Microsoft.Office.Interop.Excel;

namespace DelfaRabota
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void buildButton_Click(object sender, RoutedEventArgs e)
        {

            string name = "";
            string adress = "";
            string route = "";
            string phone = "";

            string[] mass = { "", "", "" }; //linkText.Text.Split(',');

            //try
            {

                var application = new Excel.Application();
                application.SheetsInNewWorkbook = mass.Count();

                Excel.Workbook wb = application.Workbooks.Add(Type.Missing);

                int startRowIndex = 1;

                for (int i = 0; i < mass.Count(); i++)
                {
                    WebClient web = new WebClient();
                    //web.Headers.Add("user-agent", "Mozilla / 5.0(Windows NT 10.0; Win64; x64) AppleWebKit / 537.36(KHTML, like Gecko) Chrome / 96.0.4664.174 YaBrowser / 22.1.5.810 Yowser / 2.5 Safari / 537.36");
                    //web.Headers.Add("sec-ch-ua", @""" Not A; Brand"";v=""99"", ""Chromium"";v=""96"", ""Yandex"";v=""22""");
                    //web.Headers.Add("sec-ch-ua-mobile", "?0");
                    //web.Headers.Add("sec-ch-ua-platform", @"""Windows""");
                    //web.Headers.Add("sec-fetch-dest", "document");
                    //web.Headers.Add("sec-fetch-mode", "navigate");
                    //web.Headers.Add("sec-fetch-site", "same-origin");
                    //web.Headers.Add("sec-fetch-user", "?1");
                    //web.Headers.Add("upgrade-insecure-requests", "1");
                    //web.Headers.Add("accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9");
                    //web.Headers.Add("accept-encoding", "");
                    //web.Headers.Add("cookie", @"_2gis_webapi_user=2e3af2ce-336c-4771-b9e4-0ec03e91d770; dg5_pos=65.533853;57.199953;11; _ga=GA1.2.570225354.1647451826; _gid=GA1.2.1831526861.1647451826; _ym_d=1647451826; _ym_uid=1647451826246052650; dg5_jur={""ru_sng"":{""status"":""agree"",""ts"":1647452349553,""v"":2}}; ipp_uid=1647452622316/jaIG9ZlPmVfjfcX6/wrps19rYwhD8lU63+w8M6w==; ipp_uid1=; ipp_uid2=; ipp_uid_tst=; ipp_static_key=; _ym_isad=1; _2gis_webapi_session=048c45d5-83cc-4963-927f-d99470185c83; ipp_key=v1647681673926/v3394bd400b5e53a13cfc65163aeca6afa04ab3/FwQTtQy1/xhaed99voaBvg==; _gat_online5=1");

                    string path = @"C:\Users\Danila\Desktop\f.txt";
                    StreamReader reader = new StreamReader(path);

                    string text = reader.ReadToEnd();

                    reader.Close();

                    String allData = text;      //web.DownloadString(mass[i]);

                    Match match = Regex.Match(allData.ToString(), @"_oqoid"">(.*?)</span>");
                    name = match.Groups[1].Value.ToString();

                    match = Regex.Match(allData.ToString(), @"""address_name"":""(.*?)"",""");
                    adress = match.Groups[1].Value.ToString();

                    match = Regex.Match(allData.ToString(), @"phone"",""text"":""(.*?)"",""print_text"":");
                    phone = match.Groups[1].Value.ToString();

                    string target = "";
                    Regex regex = new Regex(@"\D");
                    string result = regex.Replace(phone, target);

                    for (int j = 0; j < mass.Count(); j++)
                    {

                        Excel.Worksheet worksheet = application.Worksheets.Item[j+1];

                        worksheet.Cells[1][startRowIndex] = "Название";
                        worksheet.Cells[2][startRowIndex] = "Ссылка";
                        worksheet.Cells[3][startRowIndex] = "Телефон";
                        worksheet.Cells[4][startRowIndex] = "Адрес";

                        worksheet.Cells[1][startRowIndex + 1] = name;
                        worksheet.Cells[2][startRowIndex + 1] = route;
                        worksheet.Cells[3][startRowIndex + 1] = phone;
                        worksheet.Cells[4][startRowIndex + 1] = adress;

                        Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[4][startRowIndex + 1]];

                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;

                        worksheet.Columns.AutoFit();

                        startRowIndex++;
                    }
                }
                application.Visible = true;
            }
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message + "\nНеверная ссылка или нет подключения к интернету");
            //}
        }
    }
}
