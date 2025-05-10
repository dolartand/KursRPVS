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
    /// Основная форма приложения для расчета методом наименьших квадратов
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
        /// Конструктор по умолчанию. Инициализирует компоненты формы,
        /// COM-соединение, таблицу данных, график и анимационный таймер
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
        /// Пытается инициализировать COM-сервер для расчетов
        /// </summary>
        /// <remarks>
        /// В случае неудачи устанавливает флаг _useCom в false и использует локальные вычисления
        /// </remarks>
        private void TryInitializeCom()
        {
            const string progId = "KursComServer.LeastSquaresCom";
            try
            {
                Type comType = Type.GetTypeFromProgID(progId);
                if (comType == null) throw new InvalidOperationException($"ProgID '{progId}' не найден в реестре.");

                _comCalc = Activator.CreateInstance(comType);
                _useCom = true;
            }
            catch (Exception ex)
            {
                _useCom = false;
                MessageBox.Show(
                    $"Не удалось инициализировать COM-сервер:\n{ex.Message}\nБудет использоваться локальный калькулятор.",
                    "COM Init Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// Инициализирует таблицу данных с колонками X и Y
        /// </summary>
        private void InitializeDataGrid()
        {
            dataGridViewPoints.Columns.Clear();
            dataGridViewPoints.Columns.Add("X", "X");
            dataGridViewPoints.Columns.Add("Y", "Y");
            dataGridViewPoints.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        /// <summary>
        /// Настраивает график для отображения линии регрессии
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
        /// Инициализирует таймер для анимации построения графика
        /// </summary>
        private void InitializeTimer()
        {
            _animationTimer = new FormsTimer { Interval = 50 };
            _animationTimer.Tick += AnimationTimer_Tick;
        }

        /// <summary>
        /// Обработчик клика по кнопке "Вычислить"
        /// </summary>
        /// <param name="sender">Источник события</param>
        /// <param name="e">Параметры события</param>
        private void ToolStripButtonCompute_Click(object sender, EventArgs e)
        {
            PerformComputation();
        }

        /// <summary>
        /// Открывает диалог выбора INI-файла с точками данных
        /// </summary>
        private void OpenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using var dlg = new OpenFileDialog { Filter = "INI files (*.ini)|*.ini|All files (*.*)|*.*" };
            if (dlg.ShowDialog() == DialogResult.OK)
                txtIniPath.Text = dlg.FileName;
        }

        /// <summary>
        /// Сохраняет текущие точки данных в INI-файл
        /// </summary>
        /// <exception cref="InvalidOperationException">
        /// Выбрасывается при наличии некорректных данных в таблице
        /// </exception>
        private void SaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using var dlg = new SaveFileDialog
            {
                Title = "Сохранить точки в INI",
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

                // Обновляем поле пути и уведомляем пользователя
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


        private void ToolStripButtonOpen_Click(object sender, EventArgs e) => OpenToolStripMenuItem_Click(sender, e);
        private void ToolStripButtonSave_Click(object sender, EventArgs e) => SaveToolStripMenuItem_Click(sender, e);

        private void AboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Показываем форму "О программе" из UIHelpers.dll
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
        /// Загружает данные в локальный калькулятор из INI-файла или таблицы
        /// </summary>
        /// <exception cref="InvalidOperationException">
        /// Выбрасывается при недостаточном количестве точек или ошибках формата
        /// </exception>
        private void LoadDataIntoLocalCalculator()
        {
            _calculator.DataX.Clear();
            _calculator.DataY.Clear();

            // Если указан валидный INI-файл — читаем из него
            if (!string.IsNullOrWhiteSpace(txtIniPath.Text) && File.Exists(txtIniPath.Text))
            {
                _calculator.LoadFromIni(txtIniPath.Text);
                return;
            }

            // Иначе — читаем из таблицы
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
                    throw new InvalidOperationException("В таблице есть пустые или некорректные ячейки X/Y.");
                }
            }

            if (_calculator.DataX.Count < 2)
                throw new InvalidOperationException("Нужно ввести хотя бы две точки для вычисления.");
        }

        /// <summary>
        /// Отображает результаты расчетов на графике и в таблице
        /// </summary>
        /// <param name="a">Коэффициент наклона линии регрессии</param>
        /// <param name="b">Свободный член уравнения регрессии</param>
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
                double x = Math.Round(raw, 2);      // округление до 2 знаков
                double y = a * x + b;
                _animX.Add(x);
                _animY.Add(y);
            }

            var area = chartRegression.ChartAreas[0];
            area.AxisX.LabelStyle.Format = "0.##";    // формат меток
            area.AxisX.Interval = (maxX - minX) / Math.Min(pointsCount, 10); // разбить на ~10 шагов
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
        /// Выполняет основной расчет методом наименьших квадратов
        /// </summary>
        /// <remarks>
        /// Использует COM-сервер при доступности, иначе локальный расчет
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
                            $"Ошибка при вызове COM-сервера:\n{comEx.Message}\nБудет использован локальный алгоритм.",
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
                MessageBox.Show(ex.Message, "Ошибка вычисления", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Обработчик тиков анимационного таймера для постепенного построения графика
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
        /// Обработчик нажатия клавиш (F1 - вызов справки)
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
        /// Экспортирует результаты в Excel с графиком
        /// </summary>
        /// <exception cref="Exception">Ошибки экспорта</exception>
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
                MessageBox.Show($"Не удалось экспортировать в Excel:\n{ex.Message}",
                                "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Экспортирует результаты в Word с графиком
        /// </summary>
        /// <exception cref="Exception">Ошибки экспорта</exception>
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
                MessageBox.Show($"Не удалось экспортировать в Word:\n{ex.Message}",
                                "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // <summary>
        /// Экспортирует результаты в PowerPoint с графиком и метаданными
        /// </summary>
        /// <exception cref="Exception">Ошибки экспорта</exception>
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
                    "Метод наименьших квадратов",
                    "Студент Доленко А.А.",
                    "Разработка приложений в визуальных средах"
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось экспортировать в PowerPoint:\n{ex.Message}",
                                "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
