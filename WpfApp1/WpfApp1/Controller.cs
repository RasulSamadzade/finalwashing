using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp1
{
    class Controller
    {
        public string calcMacchia(string input1, string input2)
        {
            string decision;
            if (Double.TryParse(input1, out double Input1) && Double.TryParse(input2, out double Input2))
            {
                decision = Input1 < 5 && Input2 > 5 ? "La piastra è adatta per passare alla fase successiva" : "Contatta ingegnere di processo";
            }
            else
            {
                decision = "";
            }
            return decision;
        }

        public string calcGraffio(double input)
        {
            return input < 5.0 ? "La piastra è adatta per passare alla fase successiva" : "Contatta ingegnere di processo";
        }

        public string calcGap(double input)
        {
            return input < 20 ? "La piastra è adatta per passare alla fase successiva" : "Rifare l'incollaggio";
        }

        public String checkInput(object input)
        {
            return input == null ? "" : input.ToString();
        }
    }
}
