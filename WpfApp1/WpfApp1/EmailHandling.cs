using System;
using System.IO;
using EASendMail;

namespace WpfApp1
{
    class EmailHandling
    {
        private string email;
        private bool emailStatus = false;
        private string imagePath;

        public void sendEmail()
        {
            if (emailStatus)
            {
                try
                {
                    SmtpMail oMail = new SmtpMail("TryIt");

                    oMail.From = "technoprobe.finalwash@gmail.com";
                    oMail.To = "technoprobe.finalwash@gmail.com";

                    oMail.Subject = "Difetto";
                    oMail.TextBody = email;
                    if (imagePath != null)
                        oMail.AddAttachment("probe.png", File.ReadAllBytes(imagePath));

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
                catch (Exception ep)
                {
                    Console.WriteLine("failed to send email with the following error:");
                    Console.WriteLine(ep.Message);
                }
            }
        }

        public void generateEmailText(Data parameters, Limits limits)
        {
            emailStatus = true;
            email = "";
            var limit = limits.getValues();
            switch (parameters.Defect)
            {
                case "Sbeccatura": email = "Ciao Alessandro, Il seguente difetto non è derogabile, procediamo con il rework. Grazie!\n\n"; break;
                case "Altri": email = "Ciao Alessandro, Il seguente difetto non è presente nel CP, che recovery plan possiamo attuare? Grazie!\n\n"; break;
                case "Graffio":
                    if (Double.Parse(parameters.Input1) < limit.Graffio - limit.Tolerance && Double.Parse(parameters.Input1) > limit.Graffio)
                        email = "Ciao Alessandro, Il seguente difetto è derogabile? Grazie!";
                    else if (Double.Parse(parameters.Input1) >= limit.Graffio + limit.Tolerance)
                        email = "Ciao Alessandro,  Il seguente difetto non è derogabile, procediamo con il rework. Grazie!\n\n";
                    else
                        emailStatus = false;
                    break;
                case "Macchia":
                    if ((Double.Parse(parameters.Input1) < limit.Macchia1 + limit.Tolerance && Double.Parse(parameters.Input1) > limit.Macchia1) || (Double.Parse(parameters.Input2) > limit.Macchia2 - limit.Tolerance && Double.Parse(parameters.Input2) < limit.Macchia2))
                        email = "Ciao Alessandro, Il seguente difetto è derogabile? Grazie!";
                    else if (Double.Parse(parameters.Input1) >= limit.Macchia1 + limit.Tolerance && Double.Parse(parameters.Input2) <= limit.Macchia2 - limit.Tolerance)
                        email = "Ciao Alessandro,  Il seguente difetto non è derogabile, procediamo con il rework. Grazie!\n\n";
                    else
                        emailStatus = false;
                    break;
                case "Gap":
                    if (Double.Parse(parameters.Input1) > limit.Gap && Double.Parse(parameters.Input1) < limit.Gap + limit.Tolerance)
                        email = "Ciao Alessandro, Il seguente difetto è derogabile? Grazie!";
                    else if (Double.Parse(parameters.Input1) >= limit.Gap + limit.Tolerance)
                        email = "Ciao Alessandro,  Il seguente difetto è derogabile? Grazie!\n\n";
                    else
                        emailStatus = false;
                    break;
            }
            if (emailStatus)
            {
                String[] fields = new String[] { "Nome", "IDCode", "Tipo di Layer", "Top/Bottom", "Difetto", "Misura", "Misura2", "Decisione", "Data" };
                String[] values = new String[] { parameters.Name, parameters.ID_Code, parameters.Layer,parameters.TopBottom, parameters.Defect,
                parameters.Input1, parameters.Input2, parameters.Decision, (DateTime.Now.Date + DateTime.Now.TimeOfDay).ToString()};
                for (int i = 0; i < fields.Length; i++)
                {
                    if (values[i] != "")
                    {
                        email = $"{email} {fields[i]} - {values[i]}\n";
                    }
                }
            }
        }

        public void attachImagePath(string imagePath) {
            this.imagePath = imagePath;
        }
    }
}
