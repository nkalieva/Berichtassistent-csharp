using RibbonTaskPane;
using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Registy
{
    public partial class Form1 : Form
    {
        Microsoft.Win32.RegistryKey key;
        private MyUserControl userControl;

        public Form1()
        {
            
        }

        /* private void InitializeComponent()
         {
             throw new NotImplementedException();
         }
        */
        private void Form1_Load(object sender, EventArgs e)
        {
            key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("BA2", true);
            userControl = new MyUserControl(); 
            userControl.Dock = DockStyle.Fill;
            this.Controls.Add(userControl);
        }

        private void RegistryLocal(object sender, EventArgs e)
        {
            key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("BA2");
            key.SetValue("BA2", "localhost:6666");
            key.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                Microsoft.Win32.Registry.CurrentUser.DeleteSubKey("BA2");
            }
            catch (Exception) { }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                if (key.GetValue("BA2").ToString() == "LocalHost:6666")
                {
                    MessageBox.Show("Willkommen und viel Spaß");
                    key.SetValue("BA2", "Überpruft");
                }
                else
                {
                    MessageBox.Show("Danke");
                }
            }
            catch (Exception)
            {
            }
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(348, 293);
            this.Name = "Form1";
            this.ResumeLayout(false);

        }
    }
}