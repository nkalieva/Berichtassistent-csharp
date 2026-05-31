using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;



namespace RibbonTaskPane
{
    public partial class TaskPaneView : UserControl
    {

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // TaskPaneView
            // 
            this.Name = "TaskPaneView";
            this.Size = new System.Drawing.Size(397, 336);
            this.ResumeLayout(false);

        }
        /* private Form1 _registryForm; // Instanz der Registry-Formularklasse

public TaskPaneView()
{
    InitializeComponent();
}

private void TaskPaneView_Load(object sender, EventArgs e)
{
    _registryForm = new Form1(); // Erstellen Sie eine Instanz des Registry-Formulars
    _registryForm.Show(); // Zeigen Sie das Registry-Formular an
} */
    }
}