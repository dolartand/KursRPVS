// Kurs/MainForm.Designer.cs
using System.Windows.Forms;
using System.Drawing;

namespace Kurs
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            // Серии (Series) для графика будут добавлены программно в MainForm.cs, поэтому здесь их не инициализируем.
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.wordToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.excelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.powerPointToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpTopicsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripButtonOpen = new System.Windows.Forms.ToolStripButton();
            this.toolStripButtonSave = new System.Windows.Forms.ToolStripButton();
            this.toolStripButtonCompute = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripButtonWord = new System.Windows.Forms.ToolStripButton();
            this.toolStripButtonExcel = new System.Windows.Forms.ToolStripButton();
            this.toolStripButtonPpt = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparatorMatlab = new System.Windows.Forms.ToolStripSeparator(); // Разделитель перед MatLab
            this.toolStripButtonMatlab = new System.Windows.Forms.ToolStripButton(); // Кнопка MatLab
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripButtonHelp = new System.Windows.Forms.ToolStripButton();
            this.txtIniPath = new System.Windows.Forms.TextBox();
            this.chkAnimate = new System.Windows.Forms.CheckBox();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.dataGridViewPoints = new System.Windows.Forms.DataGridView();
            this.chartRegression = new System.Windows.Forms.DataVisualization.Charting.Chart(); // Инициализация графика
            this.contextMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.ctxOpen = new System.Windows.Forms.ToolStripMenuItem();
            this.ctxSaveAs = new System.Windows.Forms.ToolStripMenuItem();
            this.ctxSeparatorFileActions = new System.Windows.Forms.ToolStripSeparator();
            this.ctxWord = new System.Windows.Forms.ToolStripMenuItem();
            this.ctxExcel = new System.Windows.Forms.ToolStripMenuItem();
            this.ctxPowerPoint = new System.Windows.Forms.ToolStripMenuItem();
            this.ctxSeparatorHelp = new System.Windows.Forms.ToolStripSeparator();
            this.ctxHelp = new System.Windows.Forms.ToolStripMenuItem();
            this.ctxAbout = new System.Windows.Forms.ToolStripMenuItem();
            this.labelCoefA = new System.Windows.Forms.Label();
            this.txtCoefA = new System.Windows.Forms.TextBox();
            this.labelCoefB = new System.Windows.Forms.Label();
            this.txtCoefB = new System.Windows.Forms.TextBox();
            this.groupBoxRegressionType = new System.Windows.Forms.GroupBox();
            this.rbPolynomial = new System.Windows.Forms.RadioButton();
            this.rbLinear = new System.Windows.Forms.RadioButton();
            this.lblPolynomialDegree = new System.Windows.Forms.Label();
            this.nudPolynomialDegree = new System.Windows.Forms.NumericUpDown();
            this.lblCoefficientsOutput = new System.Windows.Forms.Label();
            this.txtCoefficientsOutput = new System.Windows.Forms.TextBox();
            this.menuStrip1.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout(); // Панель для графика
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewPoints)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chartRegression)).BeginInit(); // BeginInit для графика
            this.contextMenu.SuspendLayout();
            this.groupBoxRegressionType.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudPolynomialDegree)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1082, 28); // Размер может отличаться
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem,
            this.saveToolStripMenuItem,
            this.toolStripSeparator1,
            this.wordToolStripMenuItem,
            this.excelToolStripMenuItem,
            this.powerPointToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(59, 24);
            this.fileToolStripMenuItem.Text = "&Файл";
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("openToolStripMenuItem.Image"))); // Предполагается наличие ресурса
            this.openToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O)));
            this.openToolStripMenuItem.Size = new System.Drawing.Size(258, 26);
            this.openToolStripMenuItem.Text = "&Открыть...";
            this.openToolStripMenuItem.Click += new System.EventHandler(this.OpenToolStripMenuItem_Click);
            // 
            // saveToolStripMenuItem
            // 
            this.saveToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("saveToolStripMenuItem.Image"))); // Предполагается наличие ресурса
            this.saveToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.saveToolStripMenuItem.Name = "saveToolStripMenuItem";
            this.saveToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.S)));
            this.saveToolStripMenuItem.Size = new System.Drawing.Size(258, 26);
            this.saveToolStripMenuItem.Text = "&Сохранить как...";
            this.saveToolStripMenuItem.Click += new System.EventHandler(this.SaveToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(255, 6);
            // 
            // wordToolStripMenuItem
            // 
            this.wordToolStripMenuItem.Name = "wordToolStripMenuItem";
            this.wordToolStripMenuItem.Size = new System.Drawing.Size(258, 26);
            this.wordToolStripMenuItem.Text = "Экспорт в &Word...";
            this.wordToolStripMenuItem.Click += new System.EventHandler(this.wordToolStripMenuItem_Click);
            // 
            // excelToolStripMenuItem
            // 
            this.excelToolStripMenuItem.Name = "excelToolStripMenuItem";
            this.excelToolStripMenuItem.Size = new System.Drawing.Size(258, 26);
            this.excelToolStripMenuItem.Text = "Экспорт в &Excel...";
            this.excelToolStripMenuItem.Click += new System.EventHandler(this.excelToolStripMenuItem_Click);
            // 
            // powerPointToolStripMenuItem
            // 
            this.powerPointToolStripMenuItem.Name = "powerPointToolStripMenuItem";
            this.powerPointToolStripMenuItem.Size = new System.Drawing.Size(258, 26);
            this.powerPointToolStripMenuItem.Text = "Экспорт в &PowerPoint...";
            this.powerPointToolStripMenuItem.Click += new System.EventHandler(this.powerPointToolStripMenuItem_Click);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aboutToolStripMenuItem,
            this.helpTopicsToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(83, 24); // "Справка" немного длиннее
            this.helpToolStripMenuItem.Text = "С&правка";
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(221, 26); // Размер может отличаться
            this.aboutToolStripMenuItem.Text = "&О программе...";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.AboutToolStripMenuItem_Click);
            // 
            // helpTopicsToolStripMenuItem
            // 
            this.helpTopicsToolStripMenuItem.Name = "helpTopicsToolStripMenuItem";
            this.helpTopicsToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F1;
            this.helpTopicsToolStripMenuItem.Size = new System.Drawing.Size(221, 26); // Размер может отличаться
            this.helpTopicsToolStripMenuItem.Text = "Вызов &справки";
            this.helpTopicsToolStripMenuItem.Click += new System.EventHandler(this.HelpTopicsToolStripMenuItem_Click);
            // 
            // toolStrip1
            // 
            this.toolStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripButtonOpen,
            this.toolStripButtonSave,
            this.toolStripButtonCompute,
            this.toolStripSeparator2,
            this.toolStripButtonWord,
            this.toolStripButtonExcel,
            this.toolStripButtonPpt,
            this.toolStripSeparatorMatlab, // Добавлен разделитель
            this.toolStripButtonMatlab,   // Добавлена кнопка
            this.toolStripSeparator3,
            this.toolStripButtonHelp});
            this.toolStrip1.Location = new System.Drawing.Point(0, 28);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(1082, 27); // Размер может отличаться
            this.toolStrip1.TabIndex = 1;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripButtonOpen
            // 
            this.toolStripButtonOpen.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButtonOpen.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonOpen.Image")));
            this.toolStripButtonOpen.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonOpen.Name = "toolStripButtonOpen";
            this.toolStripButtonOpen.Size = new System.Drawing.Size(75, 24); // "Открыть"
            this.toolStripButtonOpen.Text = "Открыть";
            this.toolStripButtonOpen.Click += new System.EventHandler(this.ToolStripButtonOpen_Click);
            // 
            // toolStripButtonSave
            // 
            this.toolStripButtonSave.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButtonSave.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonSave.Image")));
            this.toolStripButtonSave.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonSave.Name = "toolStripButtonSave";
            this.toolStripButtonSave.Size = new System.Drawing.Size(87, 24); // "Сохранить"
            this.toolStripButtonSave.Text = "Сохранить";
            this.toolStripButtonSave.Click += new System.EventHandler(this.ToolStripButtonSave_Click);
            // 
            // toolStripButtonCompute
            // 
            this.toolStripButtonCompute.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButtonCompute.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonCompute.Image")));
            this.toolStripButtonCompute.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonCompute.Name = "toolStripButtonCompute";
            this.toolStripButtonCompute.Size = new System.Drawing.Size(88, 24); // "Вычислить"
            this.toolStripButtonCompute.Text = "Вычислить";
            this.toolStripButtonCompute.Click += new System.EventHandler(this.ToolStripButtonCompute_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 27);
            // 
            // toolStripButtonWord
            // 
            this.toolStripButtonWord.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButtonWord.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonWord.Image")));
            this.toolStripButtonWord.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonWord.Name = "toolStripButtonWord";
            this.toolStripButtonWord.Size = new System.Drawing.Size(29, 24);
            this.toolStripButtonWord.Text = "Экспорт в Word";
            this.toolStripButtonWord.Click += new System.EventHandler(this.ToolStripButtonWord_Click);
            // 
            // toolStripButtonExcel
            // 
            this.toolStripButtonExcel.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButtonExcel.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonExcel.Image")));
            this.toolStripButtonExcel.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonExcel.Name = "toolStripButtonExcel";
            this.toolStripButtonExcel.Size = new System.Drawing.Size(29, 24);
            this.toolStripButtonExcel.Text = "Экспорт в Excel";
            this.toolStripButtonExcel.Click += new System.EventHandler(this.ToolStripButtonExcel_Click);
            // 
            // toolStripButtonPpt
            // 
            this.toolStripButtonPpt.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButtonPpt.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonPpt.Image")));
            this.toolStripButtonPpt.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonPpt.Name = "toolStripButtonPpt";
            this.toolStripButtonPpt.Size = new System.Drawing.Size(29, 24);
            this.toolStripButtonPpt.Text = "Экспорт в PowerPoint";
            this.toolStripButtonPpt.Click += new System.EventHandler(this.ToolStripButtonPpt_Click);
            // 
            // toolStripSeparatorMatlab
            // 
            this.toolStripSeparatorMatlab.Name = "toolStripSeparatorMatlab";
            this.toolStripSeparatorMatlab.Size = new System.Drawing.Size(6, 27);
            // 
            // toolStripButtonMatlab
            // 
            this.toolStripButtonMatlab.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            // this.toolStripButtonMatlab.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonMatlab.Image"))); // Если есть иконка
            this.toolStripButtonMatlab.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonMatlab.Name = "toolStripButtonMatlab";
            this.toolStripButtonMatlab.Size = new System.Drawing.Size(63, 24); // "MatLab"
            this.toolStripButtonMatlab.Text = "MatLab";
            this.toolStripButtonMatlab.ToolTipText = "Вычислить и построить график в MatLab";
            this.toolStripButtonMatlab.Click += new System.EventHandler(this.ToolStripButtonMatlab_Click); // ИСПРАВЛЕНО
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(6, 27);
            // 
            // toolStripButtonHelp
            // 
            this.toolStripButtonHelp.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButtonHelp.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonHelp.Image")));
            this.toolStripButtonHelp.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonHelp.Name = "toolStripButtonHelp";
            this.toolStripButtonHelp.Size = new System.Drawing.Size(71, 24); // "Справка"
            this.toolStripButtonHelp.Text = "Справка";
            this.toolStripButtonHelp.Click += new System.EventHandler(this.ToolStripButtonHelp_Click);
            // 
            // txtIniPath
            // 
            this.txtIniPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtIniPath.Location = new System.Drawing.Point(12, 62); // Позиция может отличаться
            this.txtIniPath.Name = "txtIniPath";
            this.txtIniPath.Size = new System.Drawing.Size(678, 27); // Используем шрифт по умолчанию для .NET 8, высота 27
            this.txtIniPath.TabIndex = 2;
            // 
            // chkAnimate
            // 
            this.chkAnimate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.chkAnimate.AutoSize = true;
            this.chkAnimate.Location = new System.Drawing.Point(700, 63); // Позиция может отличаться
            this.chkAnimate.Name = "chkAnimate";
            this.chkAnimate.Size = new System.Drawing.Size(104, 24); // Используем шрифт по умолчанию, высота 24
            this.chkAnimate.TabIndex = 3;
            this.chkAnimate.Text = "Анимация";
            this.chkAnimate.UseVisualStyleBackColor = true;
            // 
            // splitContainer1
            // 
            this.splitContainer1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.splitContainer1.Location = new System.Drawing.Point(12, 95); // Сдвинуто из-за groupBoxRegressionType
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.dataGridViewPoints);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.chartRegression); // Добавление графика на панель
            this.splitContainer1.Size = new System.Drawing.Size(798, 571); // Высота может измениться
            this.splitContainer1.SplitterDistance = 266; // Может измениться
            this.splitContainer1.TabIndex = 4; // Индекс может измениться
            // 
            // dataGridViewPoints
            // 
            this.dataGridViewPoints.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewPoints.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewPoints.Location = new System.Drawing.Point(0, 0);
            this.dataGridViewPoints.Name = "dataGridViewPoints";
            this.dataGridViewPoints.RowHeadersWidth = 51;
            this.dataGridViewPoints.RowTemplate.Height = 24; // Стандартная высота строки
            this.dataGridViewPoints.Size = new System.Drawing.Size(266, 571); // Размер панели
            this.dataGridViewPoints.TabIndex = 0;
            // 
            // chartRegression
            // 
            chartArea1.Name = "ChartArea1";
            this.chartRegression.ChartAreas.Add(chartArea1);
            this.chartRegression.ContextMenuStrip = this.contextMenu;
            this.chartRegression.Dock = System.Windows.Forms.DockStyle.Fill;
            legend1.Name = "Legend1";
            this.chartRegression.Legends.Add(legend1);
            this.chartRegression.Location = new System.Drawing.Point(0, 0);
            this.chartRegression.Name = "chartRegression";
            // Серии добавляются программно
            this.chartRegression.Size = new System.Drawing.Size(528, 571); // Размер панели
            this.chartRegression.TabIndex = 0;
            this.chartRegression.Text = "chartRegression";
            // 
            // contextMenu
            // 
            this.contextMenu.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.contextMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ctxOpen,
            this.ctxSaveAs,
            this.ctxSeparatorFileActions,
            this.ctxWord,
            this.ctxExcel,
            this.ctxPowerPoint,
            this.ctxSeparatorHelp,
            this.ctxHelp,
            this.ctxAbout});
            this.contextMenu.Name = "contextMenu";
            this.contextMenu.Size = new System.Drawing.Size(233, 184); // Размер может измениться
            // 
            // ctxOpen
            // 
            this.ctxOpen.Name = "ctxOpen";
            this.ctxOpen.Size = new System.Drawing.Size(232, 24);
            this.ctxOpen.Text = "Открыть...";
            this.ctxOpen.Click += new System.EventHandler(this.OpenToolStripMenuItem_Click);
            // 
            // ctxSaveAs
            // 
            this.ctxSaveAs.Name = "ctxSaveAs";
            this.ctxSaveAs.Size = new System.Drawing.Size(232, 24);
            this.ctxSaveAs.Text = "Сохранить как...";
            this.ctxSaveAs.Click += new System.EventHandler(this.SaveToolStripMenuItem_Click);
            // 
            // ctxSeparatorFileActions
            // 
            this.ctxSeparatorFileActions.Name = "ctxSeparatorFileActions";
            this.ctxSeparatorFileActions.Size = new System.Drawing.Size(229, 6);
            // 
            // ctxWord
            // 
            this.ctxWord.Name = "ctxWord";
            this.ctxWord.Size = new System.Drawing.Size(232, 24);
            this.ctxWord.Text = "Экспорт в Word...";
            this.ctxWord.Click += new System.EventHandler(this.ctxWord_click);
            // 
            // ctxExcel
            // 
            this.ctxExcel.Name = "ctxExcel";
            this.ctxExcel.Size = new System.Drawing.Size(232, 24);
            this.ctxExcel.Text = "Экспорт в Excel...";
            this.ctxExcel.Click += new System.EventHandler(this.ctxExcel_click);
            // 
            // ctxPowerPoint
            // 
            this.ctxPowerPoint.Name = "ctxPowerPoint";
            this.ctxPowerPoint.Size = new System.Drawing.Size(232, 24);
            this.ctxPowerPoint.Text = "Экспорт в PowerPoint...";
            this.ctxPowerPoint.Click += new System.EventHandler(this.ctxPpt_click);
            // 
            // ctxSeparatorHelp
            // 
            this.ctxSeparatorHelp.Name = "ctxSeparatorHelp";
            this.ctxSeparatorHelp.Size = new System.Drawing.Size(229, 6);
            // 
            // ctxHelp
            // 
            this.ctxHelp.Name = "ctxHelp";
            this.ctxHelp.Size = new System.Drawing.Size(232, 24);
            this.ctxHelp.Text = "Справка";
            this.ctxHelp.Click += new System.EventHandler(this.HelpTopicsToolStripMenuItem_Click);
            // 
            // ctxAbout
            // 
            this.ctxAbout.Name = "ctxAbout";
            this.ctxAbout.Size = new System.Drawing.Size(232, 24);
            this.ctxAbout.Text = "О программе";
            this.ctxAbout.Click += new System.EventHandler(this.AboutToolStripMenuItem_Click);
            // 
            // labelCoefA
            // 
            this.labelCoefA.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.labelCoefA.AutoSize = true;
            this.labelCoefA.Location = new System.Drawing.Point(818, 219); // Позиция может отличаться
            this.labelCoefA.Name = "labelCoefA";
            this.labelCoefA.Size = new System.Drawing.Size(134, 20); // Для шрифта Segoe UI 9pt (по умолчанию в .NET 8)
            this.labelCoefA.TabIndex = 7; // Индекс может отличаться
            this.labelCoefA.Text = "Коэфф. наклона а:";
            // 
            // txtCoefA
            // 
            this.txtCoefA.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.txtCoefA.Location = new System.Drawing.Point(821, 242); // Позиция может отличаться
            this.txtCoefA.Name = "txtCoefA";
            this.txtCoefA.ReadOnly = true;
            this.txtCoefA.Size = new System.Drawing.Size(249, 27);
            this.txtCoefA.TabIndex = 8; // Индекс может отличаться
            // 
            // labelCoefB
            // 
            this.labelCoefB.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.labelCoefB.AutoSize = true;
            this.labelCoefB.Location = new System.Drawing.Point(818, 278); // Позиция может отличаться
            this.labelCoefB.Name = "labelCoefB";
            this.labelCoefB.Size = new System.Drawing.Size(143, 20);
            this.labelCoefB.TabIndex = 9; // Индекс может отличаться
            this.labelCoefB.Text = "Свободный член b:";
            // 
            // txtCoefB
            // 
            this.txtCoefB.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.txtCoefB.Location = new System.Drawing.Point(821, 301); // Позиция может отличаться
            this.txtCoefB.Name = "txtCoefB";
            this.txtCoefB.ReadOnly = true;
            this.txtCoefB.Size = new System.Drawing.Size(249, 27);
            this.txtCoefB.TabIndex = 10; // Индекс может отличаться
            // 
            // groupBoxRegressionType
            // 
            this.groupBoxRegressionType.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBoxRegressionType.Controls.Add(this.rbPolynomial);
            this.groupBoxRegressionType.Controls.Add(this.rbLinear);
            this.groupBoxRegressionType.Location = new System.Drawing.Point(818, 95); // Позиция может отличаться
            this.groupBoxRegressionType.Name = "groupBoxRegressionType";
            this.groupBoxRegressionType.Size = new System.Drawing.Size(252, 55); // Размер может отличаться
            this.groupBoxRegressionType.TabIndex = 5; // Индекс может измениться
            this.groupBoxRegressionType.TabStop = false;
            this.groupBoxRegressionType.Text = "Тип регрессии";
            // 
            // rbPolynomial
            // 
            this.rbPolynomial.AutoSize = true;
            this.rbPolynomial.Location = new System.Drawing.Point(110, 22);
            this.rbPolynomial.Name = "rbPolynomial";
            this.rbPolynomial.Size = new System.Drawing.Size(110, 24);
            this.rbPolynomial.TabIndex = 1;
            this.rbPolynomial.Text = "Полином-я";
            this.rbPolynomial.UseVisualStyleBackColor = true;
            this.rbPolynomial.CheckedChanged += new System.EventHandler(this.rbRegressionType_CheckedChanged);
            // 
            // rbLinear
            // 
            this.rbLinear.AutoSize = true;
            this.rbLinear.Checked = true;
            this.rbLinear.Location = new System.Drawing.Point(6, 22);
            this.rbLinear.Name = "rbLinear";
            this.rbLinear.Size = new System.Drawing.Size(100, 24);
            this.rbLinear.TabIndex = 0;
            this.rbLinear.TabStop = true;
            this.rbLinear.Text = "Линейная";
            this.rbLinear.UseVisualStyleBackColor = true;
            this.rbLinear.CheckedChanged += new System.EventHandler(this.rbRegressionType_CheckedChanged);
            // 
            // lblPolynomialDegree
            // 
            this.lblPolynomialDegree.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblPolynomialDegree.AutoSize = true;
            this.lblPolynomialDegree.Location = new System.Drawing.Point(818, 160); // Позиция может отличаться
            this.lblPolynomialDegree.Name = "lblPolynomialDegree";
            this.lblPolynomialDegree.Size = new System.Drawing.Size(106, 20);
            this.lblPolynomialDegree.TabIndex = 6; // Индекс может измениться
            this.lblPolynomialDegree.Text = "Ст. полинома:";
            // 
            // nudPolynomialDegree
            // 
            this.nudPolynomialDegree.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.nudPolynomialDegree.Location = new System.Drawing.Point(950, 158); // Позиция может отличаться
            this.nudPolynomialDegree.Maximum = new decimal(new int[] {
            7,
            0,
            0,
            0});
            this.nudPolynomialDegree.Minimum = new decimal(new int[] {
            0, // Минимальная степень 0 (константа)
            0,
            0,
            0});
            this.nudPolynomialDegree.Name = "nudPolynomialDegree";
            this.nudPolynomialDegree.Size = new System.Drawing.Size(120, 27);
            this.nudPolynomialDegree.TabIndex = 7; // Индекс может измениться
            this.nudPolynomialDegree.Value = new decimal(new int[] {
            2,
            0,
            0,
            0});
            // 
            // lblCoefficientsOutput
            // 
            this.lblCoefficientsOutput.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblCoefficientsOutput.AutoSize = true;
            this.lblCoefficientsOutput.Location = new System.Drawing.Point(818, 337); // Позиция может отличаться
            this.lblCoefficientsOutput.Name = "lblCoefficientsOutput";
            this.lblCoefficientsOutput.Size = new System.Drawing.Size(194, 20);
            this.lblCoefficientsOutput.TabIndex = 11; // Индекс может отличаться
            this.lblCoefficientsOutput.Text = "Коэффициенты полинома:";
            // 
            // txtCoefficientsOutput
            // 
            this.txtCoefficientsOutput.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtCoefficientsOutput.Location = new System.Drawing.Point(821, 360); // Позиция может отличаться
            this.txtCoefficientsOutput.Multiline = true;
            this.txtCoefficientsOutput.Name = "txtCoefficientsOutput";
            this.txtCoefficientsOutput.ReadOnly = true;
            this.txtCoefficientsOutput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtCoefficientsOutput.Size = new System.Drawing.Size(249, 306); // Высота может отличаться
            this.txtCoefficientsOutput.TabIndex = 12; // Индекс может отличаться
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F); // Для шрифта Segoe UI 9pt
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1082, 678); // Размер может отличаться
            this.ContextMenuStrip = this.contextMenu;
            this.Controls.Add(this.groupBoxRegressionType);
            this.Controls.Add(this.txtCoefficientsOutput);
            this.Controls.Add(this.lblCoefficientsOutput);
            this.Controls.Add(this.nudPolynomialDegree);
            this.Controls.Add(this.lblPolynomialDegree);
            this.Controls.Add(this.txtCoefB);
            this.Controls.Add(this.labelCoefB);
            this.Controls.Add(this.txtCoefA);
            this.Controls.Add(this.labelCoefA);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.chkAnimate);
            this.Controls.Add(this.txtIniPath);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.MinimumSize = new System.Drawing.Size(950, 600); // Минимальный размер
            this.Name = "MainForm";
            this.Text = "Метод наименьших квадратов";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false); // ДляEndInit
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewPoints)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chartRegression)).EndInit(); // EndInit для графика
            this.contextMenu.ResumeLayout(false);
            this.groupBoxRegressionType.ResumeLayout(false);
            this.groupBoxRegressionType.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudPolynomialDegree)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem wordToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem excelToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem powerPointToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpTopicsToolStripMenuItem;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton toolStripButtonOpen;
        private System.Windows.Forms.ToolStripButton toolStripButtonSave;
        private System.Windows.Forms.ToolStripButton toolStripButtonCompute;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripButton toolStripButtonWord;
        private System.Windows.Forms.ToolStripButton toolStripButtonExcel;
        private System.Windows.Forms.ToolStripButton toolStripButtonPpt;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparatorMatlab; // Объявление разделителя
        private System.Windows.Forms.ToolStripButton toolStripButtonMatlab;   // Объявление кнопки
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private System.Windows.Forms.ToolStripButton toolStripButtonHelp;
        private System.Windows.Forms.TextBox txtIniPath;
        private System.Windows.Forms.CheckBox chkAnimate;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.DataGridView dataGridViewPoints;
        private System.Windows.Forms.DataVisualization.Charting.Chart chartRegression; // Объявление графика
        private System.Windows.Forms.ContextMenuStrip contextMenu;
        private System.Windows.Forms.ToolStripMenuItem ctxOpen;
        private System.Windows.Forms.ToolStripMenuItem ctxSaveAs;
        private System.Windows.Forms.ToolStripSeparator ctxSeparatorFileActions;
        private System.Windows.Forms.ToolStripMenuItem ctxWord;
        private System.Windows.Forms.ToolStripMenuItem ctxExcel;
        private System.Windows.Forms.ToolStripMenuItem ctxPowerPoint;
        private System.Windows.Forms.ToolStripSeparator ctxSeparatorHelp;
        private System.Windows.Forms.ToolStripMenuItem ctxHelp;
        private System.Windows.Forms.ToolStripMenuItem ctxAbout;
        private System.Windows.Forms.Label labelCoefA;
        private System.Windows.Forms.TextBox txtCoefA;
        private System.Windows.Forms.Label labelCoefB;
        private System.Windows.Forms.TextBox txtCoefB;
        private System.Windows.Forms.GroupBox groupBoxRegressionType;
        private System.Windows.Forms.RadioButton rbPolynomial;
        private System.Windows.Forms.RadioButton rbLinear;
        private System.Windows.Forms.Label lblPolynomialDegree;
        private System.Windows.Forms.NumericUpDown nudPolynomialDegree;
        private System.Windows.Forms.Label lblCoefficientsOutput;
        private System.Windows.Forms.TextBox txtCoefficientsOutput;
    }
}