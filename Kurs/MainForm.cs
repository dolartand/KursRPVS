using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using KursCore;
using UIHelpers;

using FormsTimer = System.Windows.Forms.Timer;

namespace Kurs
{
    /// <summary>
    /// �������� ����� ���������� ��� ������� ������� ���������� ���������
    /// </summary>
    public partial class MainForm : Form
    {
        private dynamic _comCalc;
        private bool _useCom;
        private readonly LeastSquaresCalculator _calculator;
        private FormsTimer _animationTimer;
        private List<double> _animX;
        private List<double> _animY;
        private int _animIndex;

        /// <summary>
        /// ����������� �� ���������. �������������� ���������� �����,
        /// COM-����������, ������� ������, ������ � ������������ ������
        /// </summary>
        public MainForm()
        {
            InitializeComponent();
            this.KeyPreview = true;
            this.KeyDown += Form1_KeyDown;
            _calculator = new LeastSquaresCalculator();
            TryInitializeCom();
            InitializeDataGrid();
            InitializeChart();
            InitializeTimer();
        }

        /// <summary>
        /// �������� ���������������� COM-������ ��� ��������
        /// </summary>
        /// <remarks>
        /// � ������ ������� ������������� ���� _useCom � false � ���������� ��������� ����������
        /// </remarks>
        private void TryInitializeCom()
        {
            const string progId = "KursComServer.LeastSquaresCom";
            try
            {
                Type comType = Type.GetTypeFromProgID(progId);
                if (comType == null) throw new InvalidOperationException($"ProgID '{progId}' �� ������ � �������.");

                _comCalc = Activator.CreateInstance(comType);
                _useCom = true;
            }
            catch (Exception ex)
            {
                _useCom = false;
                MessageBox.Show(
                    $"�� ������� ���������������� COM-������:\n{ex.Message}\n����� �������������� ��������� �����������.",
                    "COM Init Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// �������������� ������� ������ � ��������� X � Y
        /// </summary>
        private void InitializeDataGrid()
        {
            dataGridViewPoints.Columns.Clear();
            dataGridViewPoints.Columns.Add("X", "X");
            dataGridViewPoints.Columns.Add("Y", "Y");
            dataGridViewPoints.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        /// <summary>
        /// ����������� ������ ��� ����������� ����� ���������
        /// </summary>
        private void InitializeChart()
        {
            chartRegression.Series.Clear();
            var series = new Series("RegressionLine")
            {
                ChartType = SeriesChartType.Line,
                BorderWidth = 2
            };
            chartRegression.Series.Add(series);
        }

        /// <summary>
        /// �������������� ������ ��� �������� ���������� �������
        /// </summary>
        private void InitializeTimer()
        {
            _animationTimer = new FormsTimer { Interval = 50 };
            _animationTimer.Tick += AnimationTimer_Tick;
        }

        /// <summary>
        /// ���������� ����� �� ������ "���������"
        /// </summary>
        /// <param name="sender">�������� �������</param>
        /// <param name="e">��������� �������</param>
        private void ToolStripButtonCompute_Click(object sender, EventArgs e)
        {
            PerformComputation();
        }

        /// <summary>
        /// ��������� ������ ������ INI-����� � ������� ������
        /// </summary>
        private void OpenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using var dlg = new OpenFileDialog { Filter = "INI files (*.ini)|*.ini|All files (*.*)|*.*" };
            if (dlg.ShowDialog() == DialogResult.OK)
                txtIniPath.Text = dlg.FileName;
        }

        /// <summary>
        /// ��������� ������� ����� ������ � INI-����
        /// </summary>
        /// <exception cref="InvalidOperationException">
        /// ������������� ��� ������� ������������ ������ � �������
        /// </exception>
        private void SaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using var dlg = new SaveFileDialog
            {
                Title = "��������� ����� � INI",
                Filter = "INI files (*.ini)|*.ini|All files (*.*)|*.*",
                DefaultExt = "ini",
                FileName = txtIniPath.Text
            };

            if (dlg.ShowDialog() != DialogResult.OK)
                return;

            try
            {
                LoadDataIntoLocalCalculator();

                _calculator.SaveToIni(dlg.FileName);

                // ��������� ���� ���� � ���������� ������������
                txtIniPath.Text = dlg.FileName;
                MessageBox.Show("������ ������� ��������� � INI.", "�����",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"�� ������� ��������� INI:\n{ex.Message}", "������ ����������",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void ToolStripButtonOpen_Click(object sender, EventArgs e) => OpenToolStripMenuItem_Click(sender, e);
        private void ToolStripButtonSave_Click(object sender, EventArgs e) => SaveToolStripMenuItem_Click(sender, e);

        private void AboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // ���������� ����� "� ���������" �� UIHelpers.dll
            using var about = new AboutForm();
            about.ShowDialog(this);
        }

        private void HelpTopicsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using var help = new HelpForm();
            help.ShowDialog(this);
        }

        private void ToolStripButtonHelp_Click(object sender, EventArgs e) => HelpTopicsToolStripMenuItem_Click(sender, e);

        /// <summary>
        /// ��������� ������ � ��������� ����������� �� INI-����� ��� �������
        /// </summary>
        /// <exception cref="InvalidOperationException">
        /// ������������� ��� ������������� ���������� ����� ��� ������� �������
        /// </exception>
        private void LoadDataIntoLocalCalculator()
        {
            _calculator.DataX.Clear();
            _calculator.DataY.Clear();

            // ���� ������ �������� INI-���� � ������ �� ����
            if (!string.IsNullOrWhiteSpace(txtIniPath.Text) && File.Exists(txtIniPath.Text))
            {
                _calculator.LoadFromIni(txtIniPath.Text);
                return;
            }

            // ����� � ������ �� �������
            foreach (DataGridViewRow row in dataGridViewPoints.Rows)
            {
                if (row.IsNewRow) continue;
                if (double.TryParse(row.Cells["X"].Value?.ToString(), out double x)
                    && double.TryParse(row.Cells["Y"].Value?.ToString(), out double y))
                {
                    _calculator.DataX.Add(x);
                    _calculator.DataY.Add(y);
                }
                else
                {
                    throw new InvalidOperationException("� ������� ���� ������ ��� ������������ ������ X/Y.");
                }
            }

            if (_calculator.DataX.Count < 2)
                throw new InvalidOperationException("����� ������ ���� �� ��� ����� ��� ����������.");
        }

        /// <summary>
        /// ���������� ���������� �������� �� ������� � � �������
        /// </summary>
        /// <param name="a">����������� ������� ����� ���������</param>
        /// <param name="b">��������� ���� ��������� ���������</param>
        private void RenderResults(double a, double b)
        {
            dataGridViewPoints.Rows.Clear();
            for (int i = 0; i < _calculator.DataX.Count; i++)
                dataGridViewPoints.Rows.Add(_calculator.DataX[i], _calculator.DataY[i]);

            double minX = _calculator.DataX.Min();
            double maxX = _calculator.DataX.Max();
            int pointsCount = 100;

            _animX = new List<double>(pointsCount + 1);
            _animY = new List<double>(pointsCount + 1);

            for (int i = 0; i <= pointsCount; i++)
            {
                double raw = minX + i * (maxX - minX) / pointsCount;
                double x = Math.Round(raw, 2);      // ���������� �� 2 ������
                double y = a * x + b;
                _animX.Add(x);
                _animY.Add(y);
            }

            var area = chartRegression.ChartAreas[0];
            area.AxisX.LabelStyle.Format = "0.##";    // ������ �����
            area.AxisX.Interval = (maxX - minX) / Math.Min(pointsCount, 10); // ������� �� ~10 �����
            area.AxisX.Minimum = _animX.First();
            area.AxisX.Maximum = _animX.Last();

            var series = chartRegression.Series["RegressionLine"];
            series.Points.Clear();

            if (chkAnimate.Checked)
            {
                _animIndex = 0;
                _animationTimer.Start();
            }
            else
            {
                foreach (var (x, y) in _animX.Zip(_animY, (xx, yy) => (xx, yy)))
                    series.Points.AddXY(x, y);
            }
        }

        /// <summary>
        /// ��������� �������� ������ ������� ���������� ���������
        /// </summary>
        /// <remarks>
        /// ���������� COM-������ ��� �����������, ����� ��������� ������
        /// </remarks>
        private void PerformComputation()
        {
            try
            {
                LoadDataIntoLocalCalculator();

                string iniPath = txtIniPath.Text;
                if (string.IsNullOrWhiteSpace(iniPath) || !File.Exists(iniPath))
                {
                    iniPath = Path.Combine(Path.GetTempPath(), "KursData.ini");   
                    _calculator.SaveToIni(iniPath);
                }

                double a, b;

                if (_useCom)
                {
                    try
                    {
                        if (string.IsNullOrWhiteSpace(txtIniPath.Text) || !File.Exists(txtIniPath.Text))
                        {
                            var xs = _calculator.DataX.ToArray();
                            var ys = _calculator.DataY.ToArray();
                            _comCalc.SetData(xs, ys);
                            _comCalc.Compute(out a, out b);
                        }
                        else
                        {
                            _comCalc.LoadFromIni(iniPath);
                            _comCalc.Compute(out a, out b);
                        }
                    }
                    catch (Exception comEx)
                    {
                        MessageBox.Show(
                            $"������ ��� ������ COM-�������:\n{comEx.Message}\n����� ����������� ��������� ��������.",
                            "COM Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        _calculator.Compute(out a, out b);
                    }
                }
                else
                {
                    _calculator.Compute(out a, out b);
                }

                txtCoefA.Text = a.ToString("0.##");
                txtCoefB.Text = b.ToString("0.##");

                RenderResults(a, b);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "������ ����������", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// ���������� ����� ������������� ������� ��� ������������ ���������� �������
        /// </summary>
        private void AnimationTimer_Tick(object sender, EventArgs e)
        {
            if (_animIndex < _animX.Count)
            {
                chartRegression.Series["RegressionLine"].Points.AddXY(_animX[_animIndex], _animY[_animIndex]);
                _animIndex++;
            }
            else
            {
                _animationTimer.Stop();
            }
        }

        /// <summary>
        /// ���������� ������� ������ (F1 - ����� �������)
        /// </summary>
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
            {
                using var help = new HelpForm();
                help.ShowDialog();
                e.Handled = true;
            }
        }

        /// <summary>
        /// ������������ ���������� � Excel � ��������
        /// </summary>
        /// <exception cref="Exception">������ ��������</exception>
        private void ToolStripButtonExcel_Click(object sender, EventArgs e)
        {
            try
            {
                LoadDataIntoLocalCalculator();
                _calculator.Compute(out double a, out double b);
                var xs = _calculator.DataX.ToArray();
                var ys = _calculator.DataY.ToArray();
                OfficeExporter.ExportToExcel(chartRegression, xs, ys, a, b);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"�� ������� �������������� � Excel:\n{ex.Message}",
                                "������", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// ������������ ���������� � Word � ��������
        /// </summary>
        /// <exception cref="Exception">������ ��������</exception>
        private void ToolStripButtonWord_Click(object sender, EventArgs e)
        {
            try
            {
                LoadDataIntoLocalCalculator();
                _calculator.Compute(out double a, out double b);
                var xs = _calculator.DataX.ToArray();
                var ys = _calculator.DataY.ToArray();
                OfficeExporter.ExportToWord(chartRegression, xs, ys, a, b);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"�� ������� �������������� � Word:\n{ex.Message}",
                                "������", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // <summary>
        /// ������������ ���������� � PowerPoint � �������� � �����������
        /// </summary>
        /// <exception cref="Exception">������ ��������</exception>
        private void ToolStripButtonPpt_Click(object sender, EventArgs e)
        {
            try
            {
                LoadDataIntoLocalCalculator();
                _calculator.Compute(out double a, out double b);
                var xs = _calculator.DataX.ToArray();
                var ys = _calculator.DataY.ToArray();
                OfficeExporter.ExportToPowerPoint(
                    chartRegression, xs, ys, a, b,
                    "����� ���������� ���������",
                    "������� ������� �.�.",
                    "���������� ���������� � ���������� ������"
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show($"�� ������� �������������� � PowerPoint:\n{ex.Message}",
                                "������", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void wordToolStripMenuItem_Click(object sender, EventArgs e) => ToolStripButtonWord_Click(sender, e);
        private void excelToolStripMenuItem_Click(object sender, EventArgs e) => ToolStripButtonExcel_Click(sender, e);
        private void powerPointToolStripMenuItem_Click(object sender, EventArgs e) => ToolStripButtonPpt_Click(sender, e);
        private void ctxPpt_click(object sender, EventArgs e) => ToolStripButtonPpt_Click(sender, e);
        private void ctxExcel_click(object sender, EventArgs e) => ToolStripButtonExcel_Click(sender, e);
        private void ctxWord_click(object sender, EventArgs e) => ToolStripButtonWord_Click(sender, e);
    }
}
