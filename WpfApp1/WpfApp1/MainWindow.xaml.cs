﻿using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Data.SQLite;
using OfficeOpenXml;
using System.IO;
using System.Threading;
using EASendMail;
using System.Windows.Media.Imaging;

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
            Logo.Source = new BitmapImage(new Uri(Directory.GetCurrentDirectory().Substring(0, Directory.GetCurrentDirectory().Length - 9) + "Image\\Logo.jpeg", UriKind.Absolute));
            initializeUI();
        }

        private void Type_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            String value = (sender as ComboBox).SelectedItem.ToString();
            if (value == "Assembled")
            {
                Layer.SelectedIndex = -1;
                TopBottom.SelectedIndex = -1;
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

        private void calcMacchia()
        {
            if (Double.TryParse(Input1.Text, out double input1) && Double.TryParse(Input2.Text, out double input2))
            {
                if (input1 < 5 && input2 > 5)
                {
                    Decision.Text = "La piastra è adatta per passare alla fase successiva";
                }
                else
                {
                    Decision.Text = "Contatta ingegnere di processo";
                }
            }
            else
            {
                Decision.Text = "";
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
                case "Sbeccatura": email = "Ciao Alessandro, Il seguente difetto non è derogabile, procediamo con il rework. Grazie!\n\n"; break;
                case "Altri": email = "Ciao Alessandro, Il seguente difetto non è presente nel CP, che recovery plan possiamo attuare? Grazie!\n\n"; break;
                case "Graffio":
                    if (Double.Parse(Input1.Text) > 4.8 && Double.Parse(Input1.Text) < 5.2)
                        email = "Ciao Alessandro,  Il seguente difetto è derogabile? Grazie!\n\n";
                    else if (Double.Parse(Input1.Text) > 5.2)
                        email = "Ciao Alessandro,  Il seguente difetto non è derogabile, procediamo con il rework. Grazie!\n\n";
                    break;
                case "Macchia":
                    if (Double.Parse(Input1.Text) < 4.8 && Double.Parse(Input2.Text) < 5.2)
                        email = "Ciao Alessandro,  Il seguente difetto è derogabile? Grazie!\n\n";
                    else
                        email = "Ciao Alessandro,  Il seguente difetto non è derogabile, procediamo con il rework. Grazie!\n\n";
                    break;
                case "Gap":
                    if (Double.Parse(Input1.Text) > 9.8 && Double.Parse(Input1.Text) < 20.2)
                        email = "Ciao Alessandro,  Il seguente difetto è derogabile? Grazie!\n\n";
                    else
                        email = "Ciao Alessandro, Il seguente difetto non è derogabile, procediamo con il rework. Grazie!\n\n";
                    break;
            }
            String[] fields = new String[] { "Nome", "IDCode", "Tipo di Layer", "Top/Bottom", "Difetto", "Misura", "Misura2", "Decisione", "Data" };
            String[] values = new String[] { checkInput(Name.Text), checkInput(ID_Code.Text), checkInput(Layer.Text),
                checkInput(TopBottom.Text), checkInput(Defect.Text), checkInput(Input1.Text),
                checkInput(Input2.Text), checkInput(Decision.Text), (DateTime.Now.Date + DateTime.Now.TimeOfDay).ToString()};
            for (int i = 0; i < fields.Length; i++)
            {
                if (values[i] != "")
                {
                    email = $"{email} {fields[i]} - {values[i]}\n";
                }
            }
            return email;
        }

        private String checkInput(object input)
        {
            if (input == null)
            {
                input = "";
            }
            return input.ToString();
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

        public void setExcelStyle(ExcelWorksheet worksheet, String headerRange, String borderRange)
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
            Thread thread1 = new Thread(generateExcel);
            Thread thread2 = new Thread(()=>sendEmail(text));
            thread1.Start();
            thread2.Start();
            emptyFields();
        }

        public void sendEmail(String text)
        {
            try
            {
                SmtpMail oMail = new SmtpMail("TryIt");

                oMail.From = "technoprobe.finalwash@gmail.com";
                oMail.To = "technoprobe.finalwash@gmail.com";

                oMail.Subject = "Difetto";
                oMail.TextBody = text;
                oMail.AddAttachment("probe.png", File.ReadAllBytes(Directory.GetCurrentDirectory().Substring(0, Directory.GetCurrentDirectory().Length - 9) + "\\Image\\probe.png"));

                SmtpServer oServer = new SmtpServer("smtp.gmail.com");
                oServer.User = "technoprobe.finalwash@gmail.com";
                oServer.Password = "Technoprobe2021";

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

        private void Input1_TextChanged(object sender, TextChangedEventArgs e)
        {
            String value = Defect.SelectedItem.ToString();
            bool calculate = Double.TryParse(Input1.Text, out double realValue);
            if (calculate)
            {
                switch (value)
                {
                    case "Graffio": Decision.Text = realValue < 5.0 ? "La piastra è adatta per passare alla fase successiva" : "Contatta ingegnere di processo"; break;
                    case "Macchia": calcMacchia(); break;
                    case "Gap": Decision.Text = realValue < 20 ? "La piastra è adatta per passare alla fase successiva" : "Rifare l'incollaggio"; break;
                }
            }
            else
            {
                Decision.Text = "";
            }
        }

        private void Input2_TextChanged(object sender, TextChangedEventArgs e)
        {
            calcMacchia();
        }
    }
}