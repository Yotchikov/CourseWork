namespace VisioAddIn
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group = this.Factory.CreateRibbonGroup();
            this.openFileButton = this.Factory.CreateRibbonButton();
            this.exportGraphButton = this.Factory.CreateRibbonButton();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.tab1.SuspendLayout();
            this.group.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group);
            this.tab1.Label = "Импорт графа";
            this.tab1.Name = "tab1";
            // 
            // group
            // 
            this.group.Items.Add(this.openFileButton);
            this.group.Items.Add(this.exportGraphButton);
            this.group.Name = "group";
            // 
            // openFileButton
            // 
            this.openFileButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.openFileButton.Description = "Выбрать файл";
            this.openFileButton.Image = global::VisioAddIn.Properties.Resources.Browse;
            this.openFileButton.ImageName = "Выбрать файл";
            this.openFileButton.Label = "Выбрать файл";
            this.openFileButton.Name = "openFileButton";
            this.openFileButton.ShowImage = true;
            this.openFileButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.openFileButton_Click);
            // 
            // exportGraphButton
            // 
            this.exportGraphButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.exportGraphButton.Image = global::VisioAddIn.Properties.Resources.Save;
            this.exportGraphButton.Label = "Сохранить изменения";
            this.exportGraphButton.Name = "exportGraphButton";
            this.exportGraphButton.ShowImage = true;
            this.exportGraphButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.exportGraphButton_Click);
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "Graph File";
            this.openFileDialog.Filter = "DOT files (*.gv;*.dot)|*.gv;*.dot";
            // 
            // saveFileDialog
            // 
            this.saveFileDialog.DefaultExt = "gv";
            this.saveFileDialog.Filter = "DOT files (*.gv;*.dot)|*.gv;*.dot";
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Visio.Drawing";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group.ResumeLayout(false);
            this.group.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton openFileButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton exportGraphButton;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
