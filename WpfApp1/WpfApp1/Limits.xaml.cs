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
        public Limits()
        {
            InitializeComponent();
            Graffio.Text = GlobalLimits.Graffio.ToString();
            Macchia1.Text = GlobalLimits.Macchia1.ToString();
            Macchia2.Text = GlobalLimits.Macchia2.ToString();
            Gap.Text = GlobalLimits.Gap.ToString();
            Tolerance.Text = GlobalLimits.Tolerance.ToString();
        }

        protected override void OnClosed(EventArgs e)
        {
            GlobalLimits.Graffio = Double.TryParse(Graffio.Text, out GlobalLimits.Graffio) ? GlobalLimits.Graffio : 5;
            GlobalLimits.Macchia1 = Double.TryParse(Macchia1.Text, out GlobalLimits.Macchia1) ? GlobalLimits.Macchia1 : 5;
            GlobalLimits.Macchia2 = Double.TryParse(Macchia2.Text, out GlobalLimits.Macchia2) ? GlobalLimits.Macchia2 : 5;
            GlobalLimits.Gap = Double.TryParse(Gap.Text, out GlobalLimits.Gap) ? GlobalLimits.Gap : 20;
            GlobalLimits.Tolerance = Double.TryParse(Tolerance.Text, out GlobalLimits.Tolerance) ? GlobalLimits.Tolerance : 0.2;
        }
    }
}
