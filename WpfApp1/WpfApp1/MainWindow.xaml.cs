using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.IO;
using System.Threading;
using System.Windows.Media.Imaging;
using Microsoft.Win32;

namespace WpfApp1
{
    public partial class MainWindow : Window
    {
        DatabaseHandling databaseHandling;
        ExcelHandling excelHandling;
        UtilityHelpers utilityHelpers;
        EmailHandling emailHandling;
        Controller controller;

        public MainWindow()
        {
            InitializeComponent();

            databaseHandling = new DatabaseHandling();
            excelHandling = new ExcelHandling();
            utilityHelpers = new UtilityHelpers();
            emailHandling = new EmailHandling();
            controller = new Controller();

            Directory.CreateDirectory(Directory.GetDirectoryRoot(Directory.GetCurrentDirectory()) + "TechnoProbe");
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
            utilityHelpers.adjustUI(Input1, Input2, Input1_Lbl, Input2_Lbl, Decision, value);
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

        private void Input1_TextChanged(object sender, TextChangedEventArgs e)
        {
            String value = Defect.SelectedItem.ToString();
            bool calculate = Double.TryParse(Input1.Text, out double realValue);
            if (calculate)
            {
                switch (value)
                {
                    case "Graffio": Decision.Text = controller.calcGraffio(realValue); break;
                    case "Macchia": Decision.Text = controller.calcMacchia(Input1.Text, Input2.Text); break;
                    case "Gap": Decision.Text = controller.calcGap(realValue); break;
                }
            }
            else
            {
                Decision.Text = "";
            }
        }

        private void Input2_TextChanged(object sender, TextChangedEventArgs e)
        {
            Decision.Text = controller.calcMacchia(Input1.Text, Input2.Text);
        }

        private Data getFieldValues()
        {
            string type = controller.checkInput(Type.SelectedValue);
            string name = controller.checkInput(Name.Text);
            string idCode = controller.checkInput(ID_Code.Text);
            string layer = controller.checkInput(Layer.SelectedValue);
            string topBottom = controller.checkInput(TopBottom.SelectedValue);
            string defect = controller.checkInput(Defect.SelectedValue);
            string input1 = controller.checkInput(Input1.Text);
            string input2 = controller.checkInput(Input2.Text);
            string decision = controller.checkInput(Decision.Text);

            return new Data
            {
                Type = type,
                Name = name,
                ID_Code = idCode,
                Layer = layer,
                TopBottom = topBottom,
                Defect = defect,
                Input1 = input1,
                Input2 = input2,
                Decision = decision
            };
        }

        private void Select_Image(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
            
            if (openFileDialog.ShowDialog() == true)
            {
                emailHandling.attachImagePath(openFileDialog.FileName);
            }
        }

        private void Final_Click(object sender, RoutedEventArgs e)
        {
            Data fields = getFieldValues();
            databaseHandling.saveToDatabase(fields);
            emailHandling.generateEmailText(fields);

            Thread thread1 = new Thread(excelHandling.generateExcel);
            Thread thread2 = new Thread(() => emailHandling.sendEmail());

            thread1.Start();
            thread2.Start();

            emptyFields();
        }
    }
}