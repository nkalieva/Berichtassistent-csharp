using Microsoft.Office.Tools;
using Microsoft.Win32;
using RibbonTaskPane;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar;


namespace RibbonTaskPane
{
    public partial class MyUserControl : UserControl
    {
        private RegistryKey key;
        private const string RegistrySubKey = "BA2";
        private const string RegistryValueServer = "BA2_Server";
        private const string RegistryValuePort = "BA2_Port";
        //private const string RegistryValueName2 = "BA2";
        public MyUserControl()
        {
            InitializeComponent();
            textBox1.TextChanged += textBox1_TextChanged;
            textBox2.TextChanged += textBox2_TextChanged;
        }
        private void LoadRegistryValue()
        {
            key = Registry.CurrentUser.OpenSubKey(RegistrySubKey, true);
            if (key != null)
            {
                string value = (string)key.GetValue(RegistryValueServer);

                if (!string.IsNullOrEmpty(value))
                {
                    textBox1.Text = value;

                }
            }
            key = Registry.CurrentUser.OpenSubKey(RegistrySubKey, true);
            if (key != null)
            {
                string value = (string)key.GetValue(RegistryValuePort);

                if (!string.IsNullOrEmpty(value))
                {
                    textBox2.Text = value;

                }
            }
        }
        private void SaveRegistryValue()
        {
            key = Registry.CurrentUser.CreateSubKey(RegistrySubKey);
            key.SetValue(RegistryValueServer, textBox1.Text);
            key.SetValue(RegistryValuePort, textBox2.Text);

            key.Close();
        }
        // key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("BA2");
        // key.SetValue("BA2", "localhost:6666");
        // key.Close();

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void UpdateTaskPane(string information)
        {
            throw new NotImplementedException();
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            //key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("BA2", true);
        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void ButtonRegistry(object sender, EventArgs e)
        {
            SaveRegistryValue();
        }

        private void MyUserControl_SizeChanged(object sender, EventArgs e)
        {

        }

        private void MyUserControl_Load(object sender, EventArgs e)
        {
            LoadRegistryValue();
        }
    }

}



