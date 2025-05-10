namespace Kurs
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpTopicsToolStripMenuItem;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton toolStripButtonOpen;
        private System.Windows.Forms.ToolStripButton toolStripButtonSave;
        private System.Windows.Forms.ToolStripButton toolStripButtonCompute;
        private System.Windows.Forms.ToolStripButton toolStripButtonHelp;
        private ToolStripButton toolStripButtonExcel;
        private ToolStripButton toolStripButtonWord;
        private ToolStripButton toolStripButtonPpt;
        private System.Windows.Forms.TextBox txtIniPath;
        private System.Windows.Forms.CheckBox chkAnimate;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.DataGridView dataGridViewPoints;
        private System.Windows.Forms.DataVisualization.Charting.Chart chartRegression;
        private System.Windows.Forms.Label labelCoefA;
        private System.Windows.Forms.TextBox txtCoefA;
        private System.Windows.Forms.Label labelCoefB;
        private System.Windows.Forms.TextBox txtCoefB;
        private ContextMenuStrip contextMenu;
        private ToolStripMenuItem ctxHelp;
        private ToolStripMenuItem ctxAbout;

        #region Windows Form Designer generated code

        /// <summary>
        /// Method for Designer support - do not modify
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            // Меню
            menuStrip1 = new MenuStrip();
            fileToolStripMenuItem = new ToolStripMenuItem();
            openToolStripMenuItem = new ToolStripMenuItem();
            saveToolStripMenuItem = new ToolStripMenuItem();
            wordToolStripMenuItem = new ToolStripMenuItem();
            excelToolStripMenuItem = new ToolStripMenuItem();
            powerPointToolStripMenuItem = new ToolStripMenuItem();
            helpToolStripMenuItem = new ToolStripMenuItem();
            aboutToolStripMenuItem = new ToolStripMenuItem();
            helpTopicsToolStripMenuItem = new ToolStripMenuItem();

            // ToolStrip
            toolStrip1 = new ToolStrip();
            toolStripButtonOpen = new ToolStripButton();
            toolStripButtonSave = new ToolStripButton();
            toolStripButtonCompute = new ToolStripButton();
            toolStripButtonHelp = new ToolStripButton();
            toolStripButtonWord = new ToolStripButton();
            toolStripButtonExcel = new ToolStripButton();
            toolStripButtonPpt = new ToolStripButton();

            // Остальные контролы
            txtIniPath = new TextBox();
            chkAnimate = new CheckBox();
            splitContainer1 = new SplitContainer();
            dataGridViewPoints = new DataGridView();
            chartRegression = new System.Windows.Forms.DataVisualization.Charting.Chart();
            labelCoefA = new Label();
            txtCoefA = new TextBox();
            labelCoefB = new Label();
            txtCoefB = new TextBox();

            contextMenu = new ContextMenuStrip(components);
            ctxHelp = new ToolStripMenuItem();
            ctxAbout = new ToolStripMenuItem();
            ctxPowerPoint = new ToolStripMenuItem();
            ctxExcel = new ToolStripMenuItem();
            ctxWord = new ToolStripMenuItem();

            // 
            // menuStrip1
            // 
            menuStrip1.Items.AddRange(new ToolStripItem[] {
        fileToolStripMenuItem,
        helpToolStripMenuItem
    });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Size = new Size(1003, 28);

            // 
            // fileToolStripMenuItem
            // 
            fileToolStripMenuItem.Text = "Файл";
            fileToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] {
        openToolStripMenuItem,
        saveToolStripMenuItem,
        new ToolStripSeparator(),
        wordToolStripMenuItem,
        excelToolStripMenuItem,
        powerPointToolStripMenuItem
    });

            openToolStripMenuItem.Text = "Открыть";
            openToolStripMenuItem.Click += OpenToolStripMenuItem_Click;

            saveToolStripMenuItem.Text = "Сохранить";
            saveToolStripMenuItem.Click += SaveToolStripMenuItem_Click;

            wordToolStripMenuItem.Text = "Экспорт в Word";
            wordToolStripMenuItem.Click += ToolStripButtonWord_Click;

            excelToolStripMenuItem.Text = "Экспорт в Excel";
            excelToolStripMenuItem.Click += ToolStripButtonExcel_Click;

            powerPointToolStripMenuItem.Text = "Экспорт в PowerPoint";
            powerPointToolStripMenuItem.Click += ToolStripButtonPpt_Click;

            // 
            // helpToolStripMenuItem
            // 
            helpToolStripMenuItem.Text = "Help";
            helpToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] {
        aboutToolStripMenuItem,
        helpTopicsToolStripMenuItem
    });

            aboutToolStripMenuItem.Text = "О программе";
            aboutToolStripMenuItem.Click += AboutToolStripMenuItem_Click;

            helpTopicsToolStripMenuItem.Text = "Справка";
            helpTopicsToolStripMenuItem.Click += HelpTopicsToolStripMenuItem_Click;

            // 
            // toolStrip1
            // 
            toolStrip1.Items.AddRange(new ToolStripItem[] {
        toolStripButtonOpen,
        toolStripButtonSave,
        toolStripButtonCompute,
        toolStripButtonHelp,
        toolStripButtonWord,
        toolStripButtonExcel,
        toolStripButtonPpt
    });
            toolStrip1.Location = new Point(0, 28);
            toolStrip1.Size = new Size(1003, 27);

            toolStripButtonOpen.Text = "Открыть";
            toolStripButtonOpen.Click += ToolStripButtonOpen_Click;

            toolStripButtonSave.Text = "Сохранить";
            toolStripButtonSave.Click += ToolStripButtonSave_Click;

            toolStripButtonCompute.Text = "Вычислить";
            toolStripButtonCompute.Click += ToolStripButtonCompute_Click;

            toolStripButtonHelp.Text = "Справка";
            toolStripButtonHelp.Click += ToolStripButtonHelp_Click;

            toolStripButtonWord.DisplayStyle = ToolStripItemDisplayStyle.Image;
            toolStripButtonWord.Image = (Image)resources.GetObject("toolStripButtonWord.Image");
            toolStripButtonWord.Click += ToolStripButtonWord_Click;

            toolStripButtonExcel.DisplayStyle = ToolStripItemDisplayStyle.Image;
            toolStripButtonExcel.Image = (Image)resources.GetObject("toolStripButtonExcel.Image");
            toolStripButtonExcel.Click += ToolStripButtonExcel_Click;

            toolStripButtonPpt.DisplayStyle = ToolStripItemDisplayStyle.Image;
            toolStripButtonPpt.Image = (Image)resources.GetObject("toolStripButtonPpt.Image");
            toolStripButtonPpt.Click += ToolStripButtonPpt_Click;

            // 
            // txtIniPath
            // 
            txtIniPath.Location = new Point(12, 62);
            txtIniPath.Size = new Size(600, 27);

            // 
            // chkAnimate
            // 
            chkAnimate.Location = new Point(630, 62);
            chkAnimate.Text = "Анимация";

            // 
            // splitContainer1
            // 
            splitContainer1.Location = new Point(12, 92);
            splitContainer1.Size = new Size(776, 348);
            splitContainer1.SplitterDistance = 300;
            splitContainer1.Panel1.Controls.Add(dataGridViewPoints);
            splitContainer1.Panel2.Controls.Add(chartRegression);

            // 
            // dataGridViewPoints
            // 
            dataGridViewPoints.Dock = DockStyle.Fill;

            // 
            // chartRegression
            // 
            chartRegression.Dock = DockStyle.Fill;
            var chartArea = new System.Windows.Forms.DataVisualization.Charting.ChartArea("ChartArea1");
            chartRegression.ChartAreas.Add(chartArea);

            // 
            // labelCoefA, txtCoefA, labelCoefB, txtCoefB
            // 
            labelCoefA.Location = new Point(794, 92);
            labelCoefA.Text = "Коэфф. a:";
            txtCoefA.Location = new Point(794, 115);
            txtCoefA.ReadOnly = true;

            labelCoefB.Location = new Point(794, 166);
            labelCoefB.Text = "Коэфф. b:";
            txtCoefB.Location = new Point(794, 189);
            txtCoefB.ReadOnly = true;

            // 
            // contextMenu
            // 
            contextMenu.Items.AddRange(new ToolStripItem[] {
        ctxHelp, ctxAbout,
        new ToolStripSeparator(),
        ctxWord, ctxExcel, ctxPowerPoint
    });
            ctxHelp.Text = "Справка"; ctxHelp.Click += HelpTopicsToolStripMenuItem_Click;
            ctxAbout.Text = "О программе"; ctxAbout.Click += AboutToolStripMenuItem_Click;
            ctxWord.Text = "Экспорт в Word"; ctxWord.Click += ToolStripButtonWord_Click;
            ctxExcel.Text = "Экспорт в Excel"; ctxExcel.Click += ToolStripButtonExcel_Click;
            ctxPowerPoint.Text = "Экспорт в PowerPoint"; ctxPowerPoint.Click += ToolStripButtonPpt_Click;

            // 
            // Form1
            // 
            this.ContextMenuStrip = contextMenu;
            this.Controls.Add(txtCoefB);
            this.Controls.Add(labelCoefB);
            this.Controls.Add(txtCoefA);
            this.Controls.Add(labelCoefA);
            this.Controls.Add(splitContainer1);
            this.Controls.Add(chkAnimate);
            this.Controls.Add(txtIniPath);
            this.Controls.Add(toolStrip1);
            this.Controls.Add(menuStrip1);
            this.MainMenuStrip = menuStrip1;
            this.ClientSize = new Size(1003, 593);
            this.Name = "Form1";
            this.Text = "Метод наименьших квадратов";
        }


        #endregion

        private ToolStripMenuItem ctxPowerPoint;
        private ToolStripMenuItem ctxExcel;
        private ToolStripMenuItem ctxWord;
        private ToolStripMenuItem wordToolStripMenuItem;
        private ToolStripMenuItem excelToolStripMenuItem;
        private ToolStripMenuItem powerPointToolStripMenuItem;
    }
}