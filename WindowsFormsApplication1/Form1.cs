using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        
        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text=="" || textBox2.Text=="" || textBox3.Text=="")
                    {
                MessageBox.Show("Заполните все поля", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
            else 
            {
                RegistryKey currentUserKey = Registry.Users;
                RegistryKey API = currentUserKey.OpenSubKey(".DEFAULT", true);
                RegistryKey values = API.CreateSubKey("API");
                values.SetValue("E-mail", textBox2.Text);
                values.SetValue("Password", textBox1.Text);
                values.SetValue("API-Key", textBox3.Text);
                values.Close();
                API.Close();
                currentUserKey.Close();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            RegistryKey currentUserKey = Registry.Users;
            RegistryKey API = currentUserKey.OpenSubKey(".DEFAULT");
            RegistryKey values = API.OpenSubKey("API", true);
            values.SetValue("E-mail", textBox2.Text);
            values.SetValue("Password", textBox1.Text);
            values.SetValue("API-Key", textBox3.Text);
            values.Close();
            API.Close();
            currentUserKey.Close();

        }
    }
}
