using Microsoft.Office.Tools.Ribbon;
using RibbonTaskPane;
using System;

namespace Berichtsassistent2
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
            //this.dropDown1.SelectionChanged += new RibbonControlEventHandler(dropDown1_SelectionChanged);
        }

        /// <summary> 
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">"true", wenn verwaltete Ressourcen gelöscht werden sollen, andernfalls "false".</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Komponenten-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.BA2Btn = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "Berichtsassistent 2";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.BA2Btn);
            this.group1.Label = "Einstellung";
            this.group1.Name = "group1";
            // 
            // BA2Btn
            // 
            this.BA2Btn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BA2Btn.Image = global::RibbonTaskPane.Properties.Resources.MicrosoftTeams_image__19_;
            this.BA2Btn.Label = "BA2";
            this.BA2Btn.Name = "BA2";
            this.BA2Btn.ShowImage = true;
            this.BA2Btn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BA2Btn_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.button1);
            this.group2.Label = "Benutzerunterstützung";
            this.group2.Name = "group2";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::RibbonTaskPane.Properties.Resources.paper_pencil_icon_icons_com_52928_300x300;
            this.button1.Label = "Mandanten Auswahl";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MandantenAuswahl);
            // 
            // group3
            // 
            this.group3.Items.Add(this.button2);
            this.group3.Items.Add(this.button3);
            this.group3.Label = "Editieren";
            this.group3.Name = "group3";
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = global::RibbonTaskPane.Properties.Resources.Download__2_;
            this.button2.Label = "layout";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Layout);
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Image = global::RibbonTaskPane.Properties.Resources.Download__3_;
            this.button3.Label = "Bereich";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load_1);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }


        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            // throw new NotImplementedException();
        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BA2Btn;
        internal RibbonGroup group2;
        internal RibbonButton button1;
        internal RibbonGroup group3;
        internal RibbonButton button2;
        internal RibbonButton button3;
    }

    /*partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
    */
}
