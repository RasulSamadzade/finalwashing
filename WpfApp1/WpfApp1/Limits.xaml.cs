using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace WpfApp1
{
    public partial class Limits : Window
    {
        public struct Numbers
        {
            public double Graffio;
            public double Macchia1;
            public double Macchia2;
            public double Gap;
            public double Tolerance;
        }

        public Limits()
        {
            InitializeComponent();
            Graffio.Text = "5";
            Macchia1.Text = "5";
            Macchia2.Text = "5";
            Gap.Text = "20";
            Tolerance.Text = "0.2";
        }

        public Numbers getValues()
        {
            Numbers numbers = new Numbers();
            numbers.Graffio = Double.TryParse(Graffio.Text, out numbers.Graffio) ? numbers.Graffio : 5.0;
            numbers.Macchia1 = Double.TryParse(Macchia1.Text, out numbers.Macchia1) ? numbers.Macchia1 : 5.0;
            numbers.Macchia2 = Double.TryParse(Macchia2.Text, out numbers.Macchia2) ? numbers.Macchia2 : 5.0;
            numbers.Gap = Double.TryParse(Gap.Text, out numbers.Gap) ? numbers.Gap : 20.0;
            numbers.Tolerance = Double.TryParse(Tolerance.Text, out numbers.Tolerance) ? numbers.Tolerance : 0.2;
            return numbers;
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
            this.Visibility = Visibility.Hidden;
        }
    }
}
