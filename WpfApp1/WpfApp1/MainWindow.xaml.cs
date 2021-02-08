using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Data.SQLite;
using System.Configuration;
using System.Data;
using OfficeOpenXml;
using System.IO;
using System.Threading;
//using System.Net.Mail;
//using System.Net;
using EASendMail;

namespace WpfApp1
{
    public partial class MainWindow : Window
    {
        SQLiteConnection sqlConnection;
        public MainWindow()
        {
            InitializeComponent();
            Directory.CreateDirectory(Directory.GetDirectoryRoot(Directory.GetCurrentDirectory()) + "TechnoProbe");
            string connectionString = "Data Source=" + Directory.GetCurrentDirectory().Substring(0, Directory.GetCurrentDirectory().Length - 9) + "TechnoProbe.db;Version=3;New=False;Compress=True;";
            sqlConnection = new SQLiteConnection(connectionString);
            initializeUI();
        }

        private void Type_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            String value = (sender as ComboBox).SelectedItem.ToString();
            if (value == "Assembled")
            {
                Layer.IsEnabled = false;
                TopBottom.IsEnabled = false;
            } else
            {
                Layer.IsEnabled = true;
                TopBottom.IsEnabled = true;
            }
        }

        private void Defect_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            String value = (sender as ComboBox).SelectedItem.ToString();
            switch (value)
            {
                case "Parylene": makeParylene(); break;
                case "Graffio": makeGraffio(); break;
                case "Macchia": makeMacchia(); break;
                case "Sbeccatura": makeSbeccature(); break;
                case "Gap": makeGap(); break;
                case "Altri": makeAltri(); break;
            }
        }

        private void Input1_PreviewLostKeyboardFocus(object sender, System.Windows.Input.KeyboardFocusChangedEventArgs e)
        {
            String value = Defect.SelectedItem.ToString();
            switch (value)
            {
                case "Graffio": Decision.Text = Int32.Parse(Input1.Text) < 5 ? "La piastra è adatta per passare alla fase successiva" : "Contatta ingegnere di processo"; break;
                case "Macchia": calcMacchia(); break;
                case "Gap": Decision.Text = Int32.Parse(Input1.Text) < 20 ? "La piastra è adatta per passare alla fase successiva" : "Rifare l'incollaggio"; break;
            }
        }

        private void Input2_PreviewLostKeyboardFocus(object sender, System.Windows.Input.KeyboardFocusChangedEventArgs e)
        {
            calcMacchia();
        }

        private void calcMacchia()
        {
            if (Input1.Text != "" && Input2.Text != "")
            {
                if (Int32.Parse(Input1.Text) < 5 && Int32.Parse(Input2.Text) > 5)
                {
                    Decision.Text = "La piastra è adatta per passare alla fase successiva";
                }
                else
                {
                    Decision.Text = "Contatta ingegnere di processo";
                }
            }
        }

        private void makeParylene()
        {
            Input1.IsEnabled = false;
            Input1_Lbl.Content = "";
            Input2.Visibility = Visibility.Hidden;
            Input2_Lbl.Visibility = Visibility.Hidden;
            Input2.Text = "";
            Decision.Text = "Rifare parylene";
        }

        private void makeGraffio()
        {
            Input1.IsEnabled = true;
            Input1_Lbl.Content = "Inserire le misure";
            Input2.Visibility = Visibility.Hidden;
            Input2_Lbl.Visibility = Visibility.Hidden;
            Input2.Text = "";
        }

        private void makeMacchia()
        {
            Input1.IsEnabled = true;
            Input1_Lbl.Content = "Inserire lungezza massima";
            Input2.Visibility = Visibility.Visible;
            Input2_Lbl.Visibility = Visibility.Visible;
            Input2.Text = "";
            Input2_Lbl.Content = "Distanza area attiva";
        }

        private void makeSbeccature()
        {
            Input1.IsEnabled = true;
            Input1_Lbl.Content = "Inserire posizione";
            Input2.Visibility = Visibility.Hidden;
            Input2_Lbl.Visibility = Visibility.Hidden;
            Input2.Text = "";
            Decision.Text = "Contatta ingegnere di processo(inserire la foto) (e - mail: alessandro.bonetto @technoprobe.com)";
        }

        private void makeGap()
        {
            Input1.IsEnabled = true;
            Input1_Lbl.Content = "Inserire la misura";
            Input2.Visibility = Visibility.Hidden;
            Input2_Lbl.Visibility = Visibility.Hidden;
            Input2.Text = "";
        }

        private void makeAltri()
        {
            Input1.IsEnabled = false;
            Input1_Lbl.Content = "";
            Input2.Visibility = Visibility.Hidden;
            Input2_Lbl.Visibility = Visibility.Hidden;
            Input2.Text = "";
            Decision.Text = "Contatta ingegnere di processo (inserire la foto) (e-mail: alessandro.bonetto@technoprobe.com)";
        }

        private void initializeUI()
        {
            Type.ItemsSource = new List<String>() { "Assembled", "Kit" };
            Layer.ItemsSource = new List<String>() { "UD1", "UD2", "LD" };
            TopBottom.ItemsSource = new List<String>() { "Top", "Bottom" };
            Defect.ItemsSource = new List<String>() { "Parylene", "Graffio", "Macchia", "Sbeccatura", "Gap", "Altri" };
            Layer.IsEnabled = false;
            TopBottom.IsEnabled = false;
            Input2.Visibility = Visibility.Hidden;
            Input2_Lbl.Visibility = Visibility.Hidden;
        }

        private void emptyFields()
        {
            Name.Text = "";
            ID_Code.Text = "";
            Input1.Text = "";
            Input2.Text = "";

        }

        private String generateEmailText()
        {
            String email = "";
            switch (Defect.SelectedItem.ToString())
            {
                case "Parylene":
                case "Sbeccatura": email = "Ciao Alessandro, (picture) (table) Il seguente difetto non è derogabile, procediamo con il rework. Grazie!"; break;
                case "Altri": email = "Ciao Alessandro,  (picture) Il seguente difetto non è presente nel CP, che recovery plan possiamo attuare? Grazie! "; break;
                case "Graffio":
                    if (Double.Parse(Input1.Text) > 4.8 && Double.Parse(Input1.Text) < 5.2)
                        email = "Ciao Alessandro,  (picture) Il seguente difetto è derogabile? Grazie! (borderline)";
                    else
                        email = "Ciao Alessandro, (picture) (table) Il seguente difetto non è derogabile, procediamo con il rework. Grazie!";
                    break;
                case "Macchia":
                    if (Double.Parse(Input1.Text) < 4.8 && Double.Parse(Input2.Text) < 5.2)
                        email = "Ciao Alessandro,  (picture) Il seguente difetto è derogabile? Grazie! (borderline)";
                    else
                        email = "Ciao Alessandro, (picture) (table) Il seguente difetto non è derogabile, procediamo con il rework. Grazie!";
                    break;
                case "Gap":
                    if (Double.Parse(Input1.Text) > 9.8 && Double.Parse(Input1.Text) < 20.2)
                        email = "Ciao Alessandro,  (picture) Il seguente difetto è derogabile? Grazie! (borderline)";
                    else
                        email = "Ciao Alessandro, (picture) (table) Il seguente difetto non è derogabile, procediamo con il rework. Grazie!";
                    break;
            }
            return email;
        }

        private void saveToDatabase()
        {
            try
            {
                string query = "INSERT INTO FinalResults (Type, Name, IdCode, Layer, TopBottom, Defect, Input1, Input2, Decision, Date) values (@Type, @Name, @IdCode, @Layer, @TopBottom, @Defect, @Input1, @Input2, @Decision, @Date)";
                SQLiteCommand sqLiteCommand = new SQLiteCommand(query, sqlConnection);
                sqlConnection.Open();
                sqLiteCommand.Parameters.AddWithValue("@Type", Type.SelectedValue);
                sqLiteCommand.Parameters.AddWithValue("@Name", Name.Text);
                sqLiteCommand.Parameters.AddWithValue("@IdCode", ID_Code.Text);
                sqLiteCommand.Parameters.AddWithValue("@Layer", Layer.SelectedValue);
                sqLiteCommand.Parameters.AddWithValue("@TopBottom", TopBottom.SelectedValue);
                sqLiteCommand.Parameters.AddWithValue("@Defect", Defect.SelectedValue);
                sqLiteCommand.Parameters.AddWithValue("@Input1", Input1.Text);
                sqLiteCommand.Parameters.AddWithValue("@Input2", Input2.Text);
                sqLiteCommand.Parameters.AddWithValue("@Decision", Decision.Text);
                sqLiteCommand.Parameters.AddWithValue("@Date", (DateTime.Now.Date + DateTime.Now.TimeOfDay).ToString());
                sqLiteCommand.ExecuteScalar();
            }
            catch (Exception ex)
            {
                MessageBox.Show("saveclick" + ex.ToString());
            }
            finally
            {
                sqlConnection.Close();
            }
        }

        private void generateExcel()
        {
            generateListFromDatabase();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Worksheet1");
                excel.Workbook.Worksheets.Add("Worksheet2");
                excel.Workbook.Worksheets.Add("Worksheet3");
                var data = generateListFromDatabase();
                var headerRow = new List<string[]>() { new string[] { "Type", "Name", "IdCode", "Layer", "TopBottom", "IdCode", "Defect", "Input1", "Input2", "Decision", "Date" } };
                string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";
                string borderRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + (data.Count + 1).ToString();
                var worksheet = excel.Workbook.Worksheets["Worksheet1"];
                setExcelStyle(worksheet, headerRange, borderRange);
                worksheet.Cells[headerRange].LoadFromArrays(headerRow);
                worksheet.Cells[2, 1].LoadFromArrays(data);
                var dir = Directory.GetDirectoryRoot(Directory.GetCurrentDirectory()) + "TechnoProbe";
                FileInfo excelFile = new FileInfo(dir + "\\TechnoProb.xlsx");
                excel.SaveAs(excelFile);
            }
        }

        public List<string[]> generateListFromDatabase()
        {
            var cellData = new List<string[]>();
            try
            {
                string query = "SELECT Type, Name, IdCode, Layer, TopBottom, Defect, Input1, Input2, Decision, Date FROM FinalResults";
                sqlConnection.Open();
                String a;
                String b;
                String c;
                String d;
                String e;
                String f;
                String g;
                String h;
                String i;
                String j;
                String k;
                var cmd = new SQLiteCommand(query, sqlConnection);
                using (SQLiteDataReader rdr = cmd.ExecuteReader())
                {
                    while (rdr.Read())
                    {
                        try
                        {
                            a = rdr.GetString(0);
                        }
                        catch (Exception ex)
                        {
                            a = "";
                        }
                        try
                        {
                            b = rdr.GetString(1);
                        }
                        catch (Exception ex)
                        {
                            b = "";
                        }
                        try
                        {
                            c = rdr.GetString(2);
                        }
                        catch (Exception ex)
                        {
                            c = "";
                        }
                        try
                        {
                            d = rdr.GetString(3);
                        }
                        catch (Exception ex)
                        {
                            d = "";
                        }
                        try
                        {
                            e = rdr.GetString(4);
                        }
                        catch (Exception ex)
                        {
                            e = "";
                        }
                        try
                        {
                            f = rdr.GetString(5);
                        }
                        catch (Exception ex)
                        {
                            f = "";
                        }
                        try
                        {
                            g = rdr.GetString(6);
                        }
                        catch (Exception ex)
                        {
                            g = "";
                        }
                        try
                        {
                            h = rdr.GetString(7);
                        }
                        catch (Exception ex)
                        {
                            h = "";
                        }
                        try
                        {
                            i = rdr.GetString(8);
                        }
                        catch (Exception ex)
                        {
                            i = "";
                        }
                        try
                        {
                            j = rdr.GetString(9);
                        }
                        catch (Exception ex)
                        {
                            j = "";
                        }
                        var cell = new string[] { a, b, c, d, e, f, g, h, i, j };
                        cellData.Add(cell);
                    }
                }
                return cellData;
            }
            catch (Exception ex)
            {
                MessageBox.Show("generatelistfromdatabase" + ex.Message.ToString());
            }
            finally
            {
                sqlConnection.Close();
            }
            return cellData;
        }

        public void setExcelStyle(OfficeOpenXml.ExcelWorksheet worksheet, String headerRange, String borderRange)
        {
            worksheet.Cells[headerRange].Style.Font.Bold = true;
            worksheet.Cells[headerRange].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[headerRange].Style.Fill.BackgroundColor.SetColor(0, 150, 150, 150);
            worksheet.Cells.AutoFitColumns();
            worksheet.Column(2).Width = 20;
            worksheet.Column(3).Width = 20;
            worksheet.Column(4).Width = 20;
            worksheet.Column(5).Width = 20;
            worksheet.Column(6).Width = 20;
            worksheet.Column(7).Width = 20;
            worksheet.Column(8).Width = 20;
            worksheet.Column(9).Width = 20;
            worksheet.Column(10).Width = 20;
            worksheet.Cells[borderRange].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[borderRange].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[borderRange].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[borderRange].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            saveToDatabase();
            string text = generateEmailText();
            sendEmail(text);
            Thread thread1 = new Thread(generateExcel);
            thread1.Start();
        }

        public void sendEmail(String text)
        {
            try
            {
                SmtpMail oMail = new SmtpMail("TryIt");

                oMail.From = "raslsam082@gmail.com";
                oMail.To = "raslsam082@gmail.com";

                oMail.Subject = "test email from gmail account";
                oMail.TextBody = text;

                SmtpServer oServer = new SmtpServer("smtp.gmail.com");
                oServer.User = "raslsam082@gmail.com";
                oServer.Password = "LucaHouse-2020";

                oServer.Port = 587;
                oServer.ConnectType = SmtpConnectType.ConnectSSLAuto;

                Console.WriteLine("start to send email over SSL ...");

                SmtpClient oSmtp = new SmtpClient();
                oSmtp.SendMail(oServer, oMail);

                Console.WriteLine("email was sent successfully!");
            }
            catch (Exception ep) {
                Console.WriteLine("failed to send email with the following error:");
                Console.WriteLine(ep.Message);
            }
        }
    }
}