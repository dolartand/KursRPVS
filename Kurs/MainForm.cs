using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using KursCore;
using UIHelpers;
using KursComServer;

using FormsTimer = System.Windows.Forms.Timer;

namespace Kurs
{
    /// <summary>
    /// Главная форма приложения для выполнения расчетов методом наименьших квадратов
    /// и визуализации результатов.
    /// </summary>
    public partial class MainForm : Form
    {
        private ILeastSquaresCom _comCalc;
        private bool _useCom;
        private readonly LeastSquaresCalculator _calculator;
        private FormsTimer _animationTimer;
        private List<double> _animX;
        private List<double> _animY;
        private int _animIndex;

        private KursCore.RegressionModelType _currentModelType = KursCore.RegressionModelType.Linear;
        private double _coeffA;
        private double _coeffB;
        private double[] _polynomialCoefficients;

        /// <summary>
        /// Инициализирует новый экземпляр класса <see cref="MainForm"/>.
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
            UpdateRegressionControlsVisibility();
        }

        /// <summary>
        /// Пытается инициализировать COM-компонент для вычислений.
        /// В случае неудачи переключается на использование локального калькулятора.
        /// </summary>
        private void TryInitializeCom()
        {
            try
            {
                _comCalc = new KursComServer.LeastSquaresCom();
                _useCom = true;

                if (_comCalc == null)
                {
                    throw new InvalidOperationException("Не удалось создать экземпляр KursComServer.LeastSquaresCom.");
                }
            }
            catch (FileNotFoundException fnfEx)
            {
                _useCom = false;
                MessageBox.Show(
                   $"Не удалось найти компонент KursComServer (возможно, отсутствует KursComServer.dll):\n{fnfEx.Message}\nБудет использоваться локальный калькулятор.",
                   "Component Load Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (TypeLoadException tlEx)
            {
                _useCom = false;
                MessageBox.Show(
                   $"Ошибка при загрузке типа из компонента KursComServer:\n{tlEx.Message}\nБудет использоваться локальный калькулятор.",
                   "Component Type Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                _useCom = false;
                MessageBox.Show(
                    $"Не удалось инициализировать компонент KursComServer:\n{ex.Message}\nБудет использоваться локальный калькулятор.",
                    "Component Init Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// Инициализирует таблицу для ввода данных (DataGridView).
        /// </summary>
        private void InitializeDataGrid()
        {
            dataGridViewPoints.Columns.Clear();
            dataGridViewPoints.Columns.Add("X", "X");
            dataGridViewPoints.Columns.Add("Y", "Y");
            dataGridViewPoints.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        /// <summary>
        /// Инициализирует элемент управления Chart для отображения графика регрессии.
        /// </summary>
        private void InitializeChart()
        {
            chartRegression.Series.Clear();
            var series = new Series("RegressionLine")
            {
                ChartType = SeriesChartType.Line,
                BorderWidth = 3,
                Color = Color.SteelBlue
            };
            var pointsSeries = new Series("DataPoints")
            {
                ChartType = SeriesChartType.Point,
                MarkerStyle = MarkerStyle.Circle,
                MarkerSize = 8,
                Color = Color.Firebrick
            };

            chartRegression.Series.Add(series);
            chartRegression.Series.Add(pointsSeries);
            if (chartRegression.Legends.Count > 0) // Добавлена проверка
            {
                chartRegression.Legends[0].Enabled = false;
            }
        }

        /// <summary>
        /// Инициализирует таймер для анимации построения графика.
        /// </summary>
        private void InitializeTimer()
        {
            _animationTimer = new FormsTimer { Interval = 30 };
            _animationTimer.Tick += AnimationTimer_Tick;
        }

        /// <summary>
        /// Обработчик нажатия кнопки "Вычислить" на панели инструментов.
        /// </summary>
        private void ToolStripButtonCompute_Click(object sender, EventArgs e)
        {
            PerformComputation();
        }

        /// <summary>
        /// Обработчик выбора пункта меню "Открыть" для загрузки данных из INI-файла.
        /// </summary>
        private void OpenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using var dlg = new OpenFileDialog { Filter = "INI files (*.ini)|*.ini|All files (*.*)|*.*" };
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txtIniPath.Text = dlg.FileName;
                try
                {
                    LoadDataIntoLocalCalculator(true);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при загрузке данных из файла:\n{ex.Message}", "Ошибка загрузки", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// Загружает данные из DataGridView или INI-файла в локальный калькулятор.
        /// </summary>
        /// <param name="fromIniPath">Указывает, следует ли пытаться загрузить данные из INI-файла, указанного в <see cref="txtIniPath"/>.</param>
        /// <exception cref="FormatException">Выбрасывается, если в таблице присутствуют некорректные числовые значения.</exception>
        /// <exception cref="InvalidOperationException">Выбрасывается, если количество точек данных недостаточно для выбранного типа регрессии.</exception>
        private void LoadDataIntoLocalCalculator(bool fromIniPath = false)
        {
            _calculator.DataX.Clear();
            _calculator.DataY.Clear();

            if (fromIniPath && !string.IsNullOrWhiteSpace(txtIniPath.Text) && File.Exists(txtIniPath.Text))
            {
                _calculator.LoadFromIni(txtIniPath.Text);
                PopulateDataGridViewFromCalculator();
            }
            else
            {
                foreach (DataGridViewRow row in dataGridViewPoints.Rows)
                {
                    if (row.IsNewRow) continue;

                    bool xParsed = double.TryParse(row.Cells["X"].Value?.ToString(),
                        System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out double x);
                    if (!xParsed)
                        xParsed = double.TryParse(row.Cells["X"].Value?.ToString(),
                            System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out x);

                    bool yParsed = double.TryParse(row.Cells["Y"].Value?.ToString(),
                        System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out double y);
                    if (!yParsed)
                        yParsed = double.TryParse(row.Cells["Y"].Value?.ToString(),
                            System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out y);

                    if (xParsed && yParsed)
                    {
                        _calculator.DataX.Add(x);
                        _calculator.DataY.Add(y);
                    }
                    else if (row.Cells["X"].Value != null && !string.IsNullOrWhiteSpace(row.Cells["X"].Value.ToString()) ||
                             row.Cells["Y"].Value != null && !string.IsNullOrWhiteSpace(row.Cells["Y"].Value.ToString()))
                    {
                        throw new FormatException($"В таблице есть некорректные значения X/Y в строке {row.Index + 1}. Используйте точку или запятую как десятичный разделитель.");
                    }
                }
            }

            int requiredPoints = _currentModelType == KursCore.RegressionModelType.Linear ? 2 : (int)nudPolynomialDegree.Value + 1;
            if (_calculator.DataX.Count < requiredPoints)
                throw new InvalidOperationException($"Нужно ввести хотя бы {requiredPoints} точек для вычисления регрессии типа '{_currentModelType}' (степень полинома {(int)nudPolynomialDegree.Value}).");
        }

        /// <summary>
        /// Заполняет DataGridView данными из локального калькулятора.
        /// </summary>
        private void PopulateDataGridViewFromCalculator()
        {
            dataGridViewPoints.Rows.Clear();
            if (_calculator.DataX != null && _calculator.DataY != null)
            {
                for (int i = 0; i < _calculator.DataX.Count; i++)
                {
                    dataGridViewPoints.Rows.Add(_calculator.DataX[i].ToString(System.Globalization.CultureInfo.CurrentCulture),
                                                _calculator.DataY[i].ToString(System.Globalization.CultureInfo.CurrentCulture));
                }
            }
        }

        /// <summary>
        /// Обработчик выбора пункта меню "Сохранить как" для сохранения данных в INI-файл.
        /// </summary>
        private void SaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using var dlg = new SaveFileDialog
            {
                Title = "Сохранить точки в INI",
                Filter = "INI files (*.ini)|*.ini|All files (*.*)|*.*",
                DefaultExt = "ini",
                FileName = string.IsNullOrWhiteSpace(txtIniPath.Text) ? "data.ini" : Path.GetFileName(txtIniPath.Text)
            };

            if (dlg.ShowDialog() != DialogResult.OK)
                return;

            try
            {
                LoadDataIntoLocalCalculator();
                _calculator.SaveToIni(dlg.FileName);
                txtIniPath.Text = dlg.FileName;
                MessageBox.Show("Данные успешно сохранены в INI.", "Успех",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось сохранить INI:\n{ex.Message}", "Ошибка сохранения",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Обработчик нажатия кнопки "Открыть" на панели инструментов.
        /// </summary>
        private void ToolStripButtonOpen_Click(object sender, EventArgs e) => OpenToolStripMenuItem_Click(sender, e);

        /// <summary>
        /// Обработчик нажатия кнопки "Сохранить" на панели инструментов.
        /// </summary>
        private void ToolStripButtonSave_Click(object sender, EventArgs e) => SaveToolStripMenuItem_Click(sender, e);

        /// <summary>
        /// Обработчик выбора пункта меню "О программе".
        /// </summary>
        private void AboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using var about = new AboutForm();
            about.ShowDialog(this);
        }

        /// <summary>
        /// Обработчик выбора пункта меню "Вызов справки" (F1).
        /// </summary>
        private void HelpTopicsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                using var help = new HelpForm();
                help.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось отобразить справку: {ex.Message}", "Ошибка справки", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Обработчик нажатия кнопки "Справка" на панели инструментов.
        /// </summary>
        private void ToolStripButtonHelp_Click(object sender, EventArgs e) => HelpTopicsToolStripMenuItem_Click(sender, e);

        /// <summary>
        /// Отображает результаты вычислений на графике.
        /// </summary>
        private void RenderResults()
        {
            PopulateDataGridViewFromCalculator();
            chartRegression.Series["DataPoints"].Points.Clear();

            if (_calculator.DataX == null || !_calculator.DataX.Any())
            {
                chartRegression.Series["RegressionLine"].Points.Clear();
                return;
            }

            for (int i = 0; i < _calculator.DataX.Count; i++)
            {
                chartRegression.Series["DataPoints"].Points.AddXY(_calculator.DataX[i], _calculator.DataY[i]);
            }

            double minX = _calculator.DataX.Min();
            double maxX = _calculator.DataX.Max();
            if (Math.Abs(minX - maxX) < 1e-9)
            {
                minX -= 1;
                maxX += 1;
            }

            int pointsCount = 200;

            _animX = new List<double>(pointsCount + 1);
            _animY = new List<double>(pointsCount + 1);

            for (int i = 0; i <= pointsCount; i++)
            {
                double x = minX + i * (maxX - minX) / pointsCount;
                double yValue;

                if (_currentModelType == KursCore.RegressionModelType.Linear)
                {
                    yValue = _coeffA * x + _coeffB;
                }
                else
                {
                    if (_polynomialCoefficients == null || _polynomialCoefficients.Length == 0)
                    {
                        chartRegression.Series["RegressionLine"].Points.Clear();
                        return;
                    }
                    yValue = 0;
                    for (int j = 0; j < _polynomialCoefficients.Length; j++)
                    {
                        yValue += _polynomialCoefficients[j] * Math.Pow(x, j);
                    }
                }
                _animX.Add(x);
                _animY.Add(yValue);
            }

            var area = chartRegression.ChartAreas[0];
            area.AxisX.LabelStyle.Format = "0.##";
            area.AxisY.LabelStyle.Format = "0.##";

            double dataMinX = _calculator.DataX.Min();
            double dataMaxX = _calculator.DataX.Max();
            double dataMinY = _calculator.DataY.Min();
            double dataMaxY = _calculator.DataY.Max();

            double xRange = dataMaxX - dataMinX;
            double yRange = dataMaxY - dataMinY;

            area.AxisX.Minimum = dataMinX - (xRange == 0 ? 1 : xRange * 0.1);
            area.AxisX.Maximum = dataMaxX + (xRange == 0 ? 1 : xRange * 0.1);
            area.AxisY.Minimum = dataMinY - (yRange == 0 ? 1 : yRange * 0.1);
            area.AxisY.Maximum = dataMaxY + (yRange == 0 ? 1 : yRange * 0.1);

            var series = chartRegression.Series["RegressionLine"];
            series.Points.Clear();

            if (chkAnimate.Checked && _animX.Any())
            {
                _animIndex = 0;
                if (_animationTimer.Enabled) _animationTimer.Stop();
                _animationTimer.Start();
            }
            else if (_animX.Any())
            {
                for (int k = 0; k < _animX.Count; k++)
                    series.Points.AddXY(_animX[k], _animY[k]);
            }
        }

        /// <summary>
        /// Выполняет вычисление регрессии на основе введенных данных и выбранного типа модели.
        /// Использует COM-компонент, если доступен, иначе локальный калькулятор.
        /// </summary>
        private void PerformComputation()
        {
            try
            {
                LoadDataIntoLocalCalculator(File.Exists(txtIniPath.Text) && !string.IsNullOrWhiteSpace(txtIniPath.Text));

                string iniPathForCom = txtIniPath.Text;
                bool tempIniCreated = false;

                if (_useCom && (string.IsNullOrWhiteSpace(iniPathForCom) || !File.Exists(iniPathForCom)))
                {
                    if (_calculator.DataX.Any()) // Только если есть данные для сохранения
                    {
                        iniPathForCom = Path.Combine(Path.GetTempPath(), $"KursData_{Guid.NewGuid()}.ini");
                        _calculator.SaveToIni(iniPathForCom);
                        tempIniCreated = true;
                    }
                }

                if (_currentModelType == KursCore.RegressionModelType.Linear)
                {
                    if (_useCom && _comCalc != null)
                    {
                        try
                        {
                            if (!string.IsNullOrWhiteSpace(iniPathForCom) && File.Exists(iniPathForCom))
                            {
                                _comCalc.LoadFromIni(iniPathForCom);
                                _comCalc.Compute(out _coeffA, out _coeffB);
                            }
                            else
                            {
                                if (_calculator.DataX.Count >= 2) // Проверка на достаточность данных для COM
                                {
                                    _comCalc.SetData(_calculator.DataX.ToArray(), _calculator.DataY.ToArray());
                                    _comCalc.Compute(out _coeffA, out _coeffB);
                                }
                                else
                                {
                                    _calculator.Compute(out _coeffA, out _coeffB);
                                }
                            }
                        }
                        catch (Exception comEx)
                        {
                            MessageBox.Show(
                                $"Ошибка при вызове компонента для линейной регрессии:\n{comEx.Message}\nБудет использован локальный алгоритм.",
                                "Component Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            _calculator.Compute(out _coeffA, out _coeffB);
                        }
                    }
                    else
                    {
                        _calculator.Compute(out _coeffA, out _coeffB);
                    }
                    txtCoefA.Text = _coeffA.ToString("F4");
                    txtCoefB.Text = _coeffB.ToString("F4");
                    txtCoefficientsOutput.Text = "";
                }
                else
                {
                    int degree = (int)nudPolynomialDegree.Value;
                    if (_useCom && _comCalc != null)
                    {
                        try
                        {
                            if (!string.IsNullOrWhiteSpace(iniPathForCom) && File.Exists(iniPathForCom))
                            {
                                _comCalc.LoadFromIni(iniPathForCom);
                                _comCalc.SetPolynomialDegree(degree);
                                _polynomialCoefficients = _comCalc.ComputePolynomialCoefficients(degree);
                            }
                            else
                            {
                                if (_calculator.DataX.Count >= degree + 1)
                                {
                                    _comCalc.SetData(_calculator.DataX.ToArray(), _calculator.DataY.ToArray());
                                    _comCalc.SetPolynomialDegree(degree);
                                    _polynomialCoefficients = _comCalc.ComputePolynomialCoefficients(degree);
                                }
                                else
                                {
                                    _polynomialCoefficients = _calculator.ComputePolynomial(degree);
                                }
                            }
                        }
                        catch (NotImplementedException nie)
                        {
                            MessageBox.Show(
                                $"Метод для полиномиальной регрессии не реализован в компоненте KursComServer:\n{nie.Message}\nБудет использован локальный алгоритм.",
                                "Component Method Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            _polynomialCoefficients = _calculator.ComputePolynomial(degree);
                        }
                        catch (Exception comEx)
                        {
                            MessageBox.Show(
                                $"Ошибка при вызове компонента для полиномиальной регрессии:\n{comEx.Message}\nБудет использован локальный алгоритм.",
                                "Component Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            _polynomialCoefficients = _calculator.ComputePolynomial(degree);
                        }
                    }
                    else
                    {
                        _polynomialCoefficients = _calculator.ComputePolynomial(degree);
                    }

                    StringBuilder sb = new StringBuilder();
                    if (_polynomialCoefficients != null)
                    {
                        for (int i = 0; i < _polynomialCoefficients.Length; i++)
                        {
                            sb.AppendLine($"c[{i}] (x^{i}): {_polynomialCoefficients[i]:F4}");
                        }
                    }
                    txtCoefficientsOutput.Text = sb.ToString();
                    txtCoefA.Text = "";
                    txtCoefB.Text = "";
                }

                if (tempIniCreated)
                {
                    try { File.Delete(iniPathForCom); } catch { /* Игнорируем */ }
                }

                RenderResults();
            }
            catch (FormatException ex)
            {
                MessageBox.Show(ex.Message, "Ошибка формата данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show(ex.Message, "Ошибка вычисления", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (ArgumentOutOfRangeException ex)
            {
                MessageBox.Show(ex.Message, "Ошибка параметров", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла непредвиденная ошибка:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Обработчик тика таймера для анимации графика.
        /// </summary>
        private void AnimationTimer_Tick(object sender, EventArgs e)
        {
            if (_animX == null || _animY == null || _animIndex >= _animX.Count)
            {
                _animationTimer.Stop();
                return;
            }

            chartRegression.Series["RegressionLine"].Points.AddXY(_animX[_animIndex], _animY[_animIndex]);
            _animIndex++;

            if (_animIndex >= _animX.Count)
            {
                _animationTimer.Stop();
            }
        }

        /// <summary>
        /// Обработчик нажатия клавиш на форме (для F1 - вызов справки).
        /// </summary>
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
            {
                HelpTopicsToolStripMenuItem_Click(sender, e);
                e.Handled = true;
            }
        }

        /// <summary>
        /// Обработчик изменения выбора типа регрессии (линейная/полиномиальная).
        /// </summary>
        private void rbRegressionType_CheckedChanged(object sender, EventArgs e)
        {
            if (rbLinear.Checked)
                _currentModelType = KursCore.RegressionModelType.Linear;
            else if (rbPolynomial.Checked)
                _currentModelType = KursCore.RegressionModelType.Polynomial;

            UpdateRegressionControlsVisibility();
            chartRegression.Series["RegressionLine"].Points.Clear();
            txtCoefA.Text = ""; txtCoefB.Text = ""; txtCoefficientsOutput.Text = "";
            _coeffA = double.NaN; _coeffB = double.NaN; _polynomialCoefficients = null;
        }

        /// <summary>
        /// Обновляет видимость элементов управления в зависимости от выбранного типа регрессии.
        /// </summary>
        private void UpdateRegressionControlsVisibility()
        {
            bool isLinear = (_currentModelType == KursCore.RegressionModelType.Linear);

            labelCoefA.Visible = isLinear;
            txtCoefA.Visible = isLinear;
            labelCoefB.Visible = isLinear;
            txtCoefB.Visible = isLinear;

            lblPolynomialDegree.Visible = !isLinear;
            nudPolynomialDegree.Visible = !isLinear;
            lblCoefficientsOutput.Visible = !isLinear;
            txtCoefficientsOutput.Visible = !isLinear;
        }

        /// <summary>
        /// Обработчик нажатия кнопки "Экспорт в Excel" на панели инструментов.
        /// </summary>
        private void ToolStripButtonExcel_Click(object sender, EventArgs e)
        {
            try
            {
                PerformComputation();

                if ((_currentModelType == RegressionModelType.Linear && (double.IsNaN(_coeffA) || double.IsNaN(_coeffB))) ||
                    (_currentModelType == RegressionModelType.Polynomial && (_polynomialCoefficients == null || _polynomialCoefficients.Length == 0)))
                {
                    if (_calculator.DataX.Count == 0)
                    {
                        MessageBox.Show("Нет данных для экспорта.", "Экспорт", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }

                if (_currentModelType == KursCore.RegressionModelType.Linear)
                {
                    OfficeExporter.ExportToExcel(chartRegression,
                                                 _calculator.DataX.ToArray(), _calculator.DataY.ToArray(),
                                                 _coeffA, _coeffB);
                }
                else
                {
                    int degree = (int)nudPolynomialDegree.Value;
                    OfficeExporter.ExportToExcel(chartRegression,
                                                 _calculator.DataX.ToArray(), _calculator.DataY.ToArray(),
                                                 _polynomialCoefficients, degree);
                }
                MessageBox.Show("Данные успешно экспортированы в Excel.", "Экспорт завершен", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось экспортировать в Excel:\n{ex.Message}",
                                "Ошибка экспорта", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Обработчик нажатия кнопки "Экспорт в Word" на панели инструментов.
        /// </summary>
        private void ToolStripButtonWord_Click(object sender, EventArgs e)
        {
            try
            {
                PerformComputation();

                if ((_currentModelType == RegressionModelType.Linear && (double.IsNaN(_coeffA) || double.IsNaN(_coeffB))) ||
                    (_currentModelType == RegressionModelType.Polynomial && (_polynomialCoefficients == null || _polynomialCoefficients.Length == 0)))
                {
                    if (_calculator.DataX.Count == 0)
                    {
                        MessageBox.Show("Нет данных для экспорта.", "Экспорт", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }

                if (_currentModelType == KursCore.RegressionModelType.Linear)
                {
                    OfficeExporter.ExportToWord(chartRegression,
                                                _calculator.DataX.ToArray(), _calculator.DataY.ToArray(),
                                                _coeffA, _coeffB);
                }
                else
                {
                    int degree = (int)nudPolynomialDegree.Value;
                    OfficeExporter.ExportToWord(chartRegression,
                                                _calculator.DataX.ToArray(), _calculator.DataY.ToArray(),
                                                _polynomialCoefficients, degree);
                }
                MessageBox.Show("Данные успешно экспортированы в Word.", "Экспорт завершен", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось экспортировать в Word:\n{ex.Message}",
                                "Ошибка экспорта", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Обработчик нажатия кнопки "Экспорт в PowerPoint" на панели инструментов.
        /// </summary>
        private void ToolStripButtonPpt_Click(object sender, EventArgs e)
        {
            try
            {
                PerformComputation();

                if ((_currentModelType == RegressionModelType.Linear && (double.IsNaN(_coeffA) || double.IsNaN(_coeffB))) ||
                    (_currentModelType == RegressionModelType.Polynomial && (_polynomialCoefficients == null || _polynomialCoefficients.Length == 0)))
                {
                    if (_calculator.DataX.Count == 0)
                    {
                        MessageBox.Show("Нет данных для экспорта.", "Экспорт", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }

                string topic = _currentModelType == KursCore.RegressionModelType.Linear ?
                               "Метод наименьших квадратов" :
                               $"Метод наименьших квадратов, полином степени {(int)nudPolynomialDegree.Value}";

                if (_currentModelType == KursCore.RegressionModelType.Linear)
                {
                    OfficeExporter.ExportToPowerPoint(chartRegression,
                                                      _calculator.DataX.ToArray(), _calculator.DataY.ToArray(),
                                                      _coeffA, _coeffB,
                                                      topic, "Студент Доленко А.А.", "Разработка приложений в визуальных средах");
                }
                else
                {
                    int degree = (int)nudPolynomialDegree.Value;
                    OfficeExporter.ExportToPowerPoint(chartRegression,
                                                      _calculator.DataX.ToArray(), _calculator.DataY.ToArray(),
                                                      _polynomialCoefficients, degree,
                                                      topic, "Студент Доленко А.А.", "Разработка приложений в визуальных средах");
                }
                MessageBox.Show("Данные успешно экспортированы в PowerPoint.", "Экспорт завершен", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось экспортировать в PowerPoint:\n{ex.Message}",
                                "Ошибка экспорта", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Обработчик выбора пункта меню "Экспорт в Word".
        /// </summary>
        private void wordToolStripMenuItem_Click(object sender, EventArgs e) => ToolStripButtonWord_Click(sender, e);

        /// <summary>
        /// Обработчик выбора пункта меню "Экспорт в Excel".
        /// </summary>
        private void excelToolStripMenuItem_Click(object sender, EventArgs e) => ToolStripButtonExcel_Click(sender, e);

        /// <summary>
        /// Обработчик выбора пункта меню "Экспорт в PowerPoint".
        /// </summary>
        private void powerPointToolStripMenuItem_Click(object sender, EventArgs e) => ToolStripButtonPpt_Click(sender, e);

        /// <summary>
        /// Обработчик выбора пункта "Экспорт в PowerPoint" из контекстного меню.
        /// </summary>
        private void ctxPpt_click(object sender, EventArgs e) => ToolStripButtonPpt_Click(sender, e);

        /// <summary>
        /// Обработчик выбора пункта "Экспорт в Excel" из контекстного меню.
        /// </summary>
        private void ctxExcel_click(object sender, EventArgs e) => ToolStripButtonExcel_Click(sender, e);

        /// <summary>
        /// Обработчик выбора пункта "Экспорт в Word" из контекстного меню.
        /// </summary>
        private void ctxWord_click(object sender, EventArgs e) => ToolStripButtonWord_Click(sender, e);

        /// <summary>
        /// Обработчик нажатия кнопки "MatLab" для выполнения вычислений и построения графика в MatLab.
        /// </summary>
        private async void ToolStripButtonMatlab_Click(object sender, EventArgs e)
        {
            List<double> xData = new List<double>();
            List<double> yData = new List<double>();

            try
            {
                foreach (DataGridViewRow row in dataGridViewPoints.Rows)
                {
                    if (row.IsNewRow) continue;

                    bool xParsedCurrent = double.TryParse(row.Cells["X"].Value?.ToString(),
                        System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out double xVal);
                    if (!xParsedCurrent)
                        xParsedCurrent = double.TryParse(row.Cells["X"].Value?.ToString(),
                            System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out xVal);

                    bool yParsedCurrent = double.TryParse(row.Cells["Y"].Value?.ToString(),
                        System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out double yVal);
                    if (!yParsedCurrent)
                        yParsedCurrent = double.TryParse(row.Cells["Y"].Value?.ToString(),
                            System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out yVal);

                    if (xParsedCurrent && yParsedCurrent)
                    {
                        xData.Add(xVal);
                        yData.Add(yVal);
                    }
                    else if (row.Cells["X"].Value != null && !string.IsNullOrWhiteSpace(row.Cells["X"].Value.ToString()) ||
                             row.Cells["Y"].Value != null && !string.IsNullOrWhiteSpace(row.Cells["Y"].Value.ToString()))
                    {
                        MessageBox.Show($"В таблице есть некорректные значения X/Y в строке {row.Index + 1} для передачи в MatLab.",
                                        "Ошибка данных MatLab", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сборе данных для MatLab: {ex.Message}", "Ошибка данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int requiredPoints = _currentModelType == RegressionModelType.Linear ? 2 : (int)nudPolynomialDegree.Value + 1;
            if (xData.Count < requiredPoints)
            {
                MessageBox.Show($"Нужно ввести хотя бы {requiredPoints} точек для вычислений в MatLab.",
                                "Недостаточно данных для MatLab", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            await MatlabHelper.GenerateAndRunMatlabScript(
                xData,
                yData,
                _currentModelType,
                (int)nudPolynomialDegree.Value,
                this,
                (message, isBusy) =>
                {
                    if (this.InvokeRequired)
                    {
                        this.Invoke((MethodInvoker)delegate {
                            toolStripButtonMatlab.Text = message;
                        });
                    }
                    else
                    {
                        toolStripButtonMatlab.Text = message;
                    }
                }
            );
        }
    }
}