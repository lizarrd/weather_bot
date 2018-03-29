using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Net;
using System.Xml.Linq;
using System.Net.Mail;
using ImapX;
using ImapX.Enums;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.Timers;
using Excel;

namespace Weather
{
    public partial class Service1 : ServiceBase
    {
        
        public Service1()
        {
            InitializeComponent();
        }

                  

       protected override void OnStart(string[] args)
        {
            Debugger.Launch();
            Timer timer = new Timer();
            timer.Interval = 10000;
            timer.Elapsed += timer_Elapsed;
            timer.Enabled = true;
        }

        bool auth(ImapX.Message x, IExcelDataReader excelReader)
        {
            while (excelReader.Read())
            {

                if (excelReader.GetString(0) == x.From.Address)
                {
                    return true;
                }
            }
          return false;
        }
        private void timer_Elapsed(object sender, EventArgs e)
        {
            kod();
        }
        string GetAPI()
        {
            RegistryKey currentUserKey = Registry.Users;
            RegistryKey API = currentUserKey.OpenSubKey(".DEFAULT", true);
            API = API.OpenSubKey("API", true);
            string APIKEY = API.GetValue("API-key").ToString();
            API.Close();
            currentUserKey.Close();

            return APIKEY;
        }

        string GetEmail()
        {
            RegistryKey currentUserKey = Registry.Users;
            RegistryKey API = currentUserKey.OpenSubKey(".DEFAULT", true);
            API = API.OpenSubKey("API", true);
            string MYEMAIL = API.GetValue("E-mail").ToString();
            API.Close();
            currentUserKey.Close();

            return MYEMAIL;
        }

        string GetPass()
        {
            RegistryKey currentUserKey = Registry.Users;
            RegistryKey API = currentUserKey.OpenSubKey(".DEFAULT", true);
            API = API.OpenSubKey("API", true);
            string MYPASS = API.GetValue("Password").ToString();
            API.Close();
            currentUserKey.Close();

            return MYPASS;      
        }

       void kod()
        {
      
      string Call;

            string APIKEY = GetAPI();
            string MYEMAIL = GetEmail();
            string MYPASS = GetPass();


            FileStream stream = File.Open(@"C:\Users\430\Documents\Visual Studio 2015\Projects\Weather\Weather\bin\Debug\3.xlsx", FileMode.Open, FileAccess.Read);

            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
  
            DataSet Eresult = excelReader.AsDataSet();
            var Iclient = new ImapClient("imap.gmail.com", true);
            Iclient.Connect();

           string mailto = "null";  string message = "0";

            Iclient.Login(MYEMAIL, MYPASS);
            Iclient.Behavior.MessageFetchMode = MessageFetchMode.Tiny;
            var folder = Iclient.Folders.Inbox;
            folder.Messages.Download("Unseen");

            MailMessage mail = new MailMessage();
           // StreamWriter file = new StreamWriter(@"C:\Users\430\Documents\Visual Studio 2015\Projects\Weather\Weather\bin\Debug\Base.txt");
            foreach (ImapX.Message x in folder.Messages)
            {
                string datetime = DateTime.Now.ToLongTimeString() + DateTime.Now.ToLongDateString();
                datetime = datetime.Replace(@":", "-");
                mail.From = new System.Net.Mail.MailAddress(x.From.Address);
                mail.To.Add(new System.Net.Mail.MailAddress(x.From.Address));
                mail.Subject = "Hi from weather bot";
                SmtpClient client = new SmtpClient();
                client.Host = "smtp.gmail.com";
                client.Port = 587;
                client.EnableSsl = true;
                client.Credentials = new NetworkCredential(MYEMAIL, MYPASS);
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
               x.Subject= x.Subject.ToLower();
                x.Subject = Regex.Replace(x.Subject, "[^0-9a-zA-Z]+", "");
                
                x.Body.Text.ToLower();






                if (auth(x,excelReader))
                {
                 
                    File.AppendAllText(@"C:\Users\430\Documents\Visual Studio 2015\Projects\Weather\Weather\bin\Debug\Base.txt", "Sender: " + x.From + "resuested:  " + x.Subject + "weather in " + x.Body.Text + " at " + datetime + "\n", Encoding.UTF8);
                    if (x.Subject == "current")
                    {
                        Call = "http://api.openweathermap.org/data/2.5/weather?q=" + x.Body.Text + "&units=metric&mode=xml&appid=" + APIKEY;
                        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Call);
                        request.Method = "GET";
                        HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                        Stream responseStream = response.GetResponseStream();
                        XDocument doc = XDocument.Load(responseStream);
                        message = doc.Root.Element("city").Attribute("name").Value + " " + doc.Root.Element("temperature").Attribute("value").Value;
                        doc.Save(datetime + ".xml");

                    }

                    if (x.Subject == "forecast")
                    {
                        Call = "http://api.openweathermap.org/data/2.5/forecast?q=" + x.Body.Text + "&units=metric&mode=xml&appid=" + APIKEY;
                        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Call);
                        request.Method = "GET";
                        HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                        Stream responseStream = response.GetResponseStream();
                        XDocument doc = XDocument.Load(responseStream);
                        doc.Save(datetime + ".xml");

                        message = doc.Root.Element("location").Element("name").Value + "\n";
                        foreach (XElement element in doc.Root.Descendants("time"))  //создает массив из всех time'ов; descendants - потомки
                        {
                            message +=
                                element.Attribute("from").Value + " " +
                                element.Attribute("to") + " " + "\n" +
                                element.Element("temperature").Attribute("value") + "\n" + "\n";

                        }
                    }

                    if (x.Subject == "forecastdaily")
                    {
                        Call = "http://api.openweathermap.org/data/2.5/forecast/daily?q=" + x.Body.Text + "&units=metric&mode=xml&cnt=16&appid=" + APIKEY;
                        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Call);
                        request.Method = "GET";
                        HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                        Stream responseStream = response.GetResponseStream();
                        XDocument doc = XDocument.Load(responseStream);
                        doc.Save(datetime + ".xml");
                        message = doc.Root.Element("location").Element("name").Value + "\n";
                        foreach (XElement element in doc.Root.Descendants("time"))
                        {
                            message +=
                                element.Attribute("day").Value + " " +
                                element.Element("temperature").Attribute("day") + "\n";
                        }
                    }
                    System.Net.Mail.Attachment add = new System.Net.Mail.Attachment(datetime + ".xml");
                    mail.Attachments.Add(add);
                    mail.Body = message;
                    client.Send(mail);
                    mail.Dispose();
                    x.Seen = true;
                    add.Dispose();
                    File.Delete(datetime + ".xml");
                }
                else
                {
                   File.AppendAllText(@"C:\Users\430\Documents\Visual Studio 2015\Projects\Weather\Weather\bin\Debug\Base.txt", "Sender: " + x.From + "acces denied " + " at " + datetime + "\n", Encoding.UTF8);
                    message = "Нет доступа к сервису";
                    mail.Body = message;
                    client.Send(mail);
                    mail.Dispose();
                    x.Seen = true;

                }
            


        }
            //file.Close();
        } 
        protected override void OnStop()
        {
            
        }
    }
}
