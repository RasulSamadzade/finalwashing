using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace WpfApp1
{
    class UtilityHelpers
    {
        private TextBox input1;
        private TextBox input2;
        private Label input1_Lbl;
        private Label input2_Lbl;
        private TextBlock decision;

        public void adjustUI(TextBox input1, TextBox input2, Label input1_Lbl, Label input2_Lbl, TextBlock decision, string type)
        {
            this.input1 = input1;
            this.input2 = input2;
            this.input1_Lbl = input1_Lbl;
            this.input2_Lbl = input2_Lbl;
            this.decision = decision;

            switch (type)
            {
                case "Parylene": makeParylene(); break;
                case "Graffio": makeGraffio(); break;
                case "Macchia": makeMacchia(); break;
                case "Sbeccatura": makeSbeccature(); break;
                case "Gap": makeGap(); break;
                case "Altri": makeAltri(); break;
            }
        }

        private void makeParylene()
        {
            input1.IsEnabled = false;
            input1_Lbl.Content = "";
            input2.Visibility = Visibility.Hidden;
            input2_Lbl.Visibility = Visibility.Hidden;
            input2.Text = "";
            decision.Text = "Rifare parylene";
        }

        private void makeGraffio()
        {
            input1.IsEnabled = true;
            input1_Lbl.Content = "Inserire le misure";
            input2.Visibility = Visibility.Hidden;
            input2_Lbl.Visibility = Visibility.Hidden;
            input2.Text = "";
        }

        private void makeMacchia()
        {
            input1.IsEnabled = true;
            input1_Lbl.Content = "Inserire lungezza massima";
            input2.Visibility = Visibility.Visible;
            input2_Lbl.Visibility = Visibility.Visible;
            input2.Text = "";
            input2_Lbl.Content = "Distanza area attiva";
        }

        private void makeSbeccature()
        {
            input1.IsEnabled = true;
            input1_Lbl.Content = "Inserire posizione";
            input2.Visibility = Visibility.Hidden;
            input2_Lbl.Visibility = Visibility.Hidden;
            input2.Text = "";
            decision.Text = "Contatta ingegnere di processo";
        }

        private void makeGap()
        {
            input1.IsEnabled = true;
            input1_Lbl.Content = "Inserire la misura";
            input2.Visibility = Visibility.Hidden;
            input2_Lbl.Visibility = Visibility.Hidden;
            input2.Text = "";
        }

        private void makeAltri()
        {
            input1.IsEnabled = false;
            input1_Lbl.Content = "";
            input2.Visibility = Visibility.Hidden;
            input2_Lbl.Visibility = Visibility.Hidden;
            input2.Text = "";
            decision.Text = "Contatta ingegnere di processo";
        }
    }
}
