using System;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using PPT = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Charting = System.Windows.Forms.DataVisualization.Charting.Chart;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;

namespace Kurs
{
    /// <summary>
    /// Статический класс, предоставляющий методы для экспорта данных и графиков регрессии
    /// в приложения Microsoft Office: Excel, Word и PowerPoint.
    /// </summary>
    public static class OfficeExporter
    {
        /// <summary>
        /// Сохраняет указанный элемент управления Chart во временный PNG файл.
        /// </summary>
        /// <param name="chart">Элемент управления Chart для сохранения.</param>
        /// <returns>Путь к созданному временному PNG файлу.</returns>
        private static string SaveChartToPng(Charting chart)
        {
            string tmpDir = Path.Combine(Path.GetTempPath(), "KursExporter");
            Directory.CreateDirectory(tmpDir);
            string tmpFile = Path.Combine(tmpDir, Guid.NewGuid() + ".png");

            int chartWidth = chart.Width > 0 ? chart.Width : 600;
            int chartHeight = chart.Height > 0 ? chart.Height : 400;

            using (var bmp = new Bitmap(chartWidth, chartHeight))
            {
                if (chart.Series.Count > 0 && chart.Series.Any(s => s.Points.Count > 0))
                {
                    chart.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height));
                }
                else
                {
                    using (Graphics g = Graphics.FromImage(bmp))
                    {
                        g.Clear(Color.White);
                        TextRenderer.DrawText(g, "Нет данных для графика", chart.Font ?? SystemFonts.DefaultFont, new Point(10, 10), Color.Black);
                    }
                }
                bmp.Save(tmpFile, ImageFormat.Png);
            }
            return tmpFile;
        }

        /// <summary>
        /// Форматирует уравнение линейной регрессии в строку.
        /// </summary>
        /// <param name="a">Коэффициент наклона.</param>
        /// <param name="b">Свободный член.</param>
        /// <returns>Строковое представление уравнения линейной регрессии.</returns>
        private static string FormatEquationLinear(double a, double b)
        {
            return $"y = {a.ToString("F4", CultureInfo.InvariantCulture)} \u00B7 x {(b < 0 ? "-" : "+")} {Math.Abs(b).ToString("F4", CultureInfo.InvariantCulture)}";
        }

        /// <summary>
        /// Форматирует уравнение полиномиальной регрессии в строку.
        /// </summary>
        /// <param name="coeffs">Массив коэффициентов полинома (от c0 до cn).</param>
        /// <param name="degree">Степень полинома (используется для информации, фактическое форматирование зависит от длины массива coeffs).</param>
        /// <returns>Строковое представление уравнения полиномиальной регрессии.</returns>
        private static string FormatEquationPolynomial(double[] coeffs, int degree)
        {
            var sb = new StringBuilder("y = ");
            bool firstTerm = true;
            if (coeffs == null || coeffs.Length == 0) return "y = (коэффициенты не вычислены)";

            for (int i = 0; i < coeffs.Length; i++)
            {
                double coeff = coeffs[i];
                int power = i;

                if (Math.Abs(coeff) < 1e-9)
                {
                    if (coeffs.Length == 1 && i == 0) { }
                    else if (i == 0 && coeffs.Length > 1 && coeffs.Skip(1).All(c => Math.Abs(c) < 1e-9)) { }
                    else if (i == 0 && coeffs.Length > 1 && coeffs.Skip(1).Any(c => Math.Abs(c) >= 1e-9)) { }
                    else if (i > 0) continue;
                }

                if (firstTerm)
                {
                    if (coeff < 0) sb.Append("-");
                    firstTerm = false;
                }
                else
                {
                    sb.Append(coeff < 0 ? " - " : " + ");
                }
                sb.Append(Math.Abs(coeff).ToString("F4", CultureInfo.InvariantCulture));

                if (power == 1) sb.Append(" \u00B7 x");
                else if (power > 1) sb.Append($" \u00B7 x^{power}");
            }
            if (sb.ToString() == "y = " || sb.ToString() == "y = -" || sb.ToString() == "y =  - ") sb.Append("0");
            return sb.ToString();
        }

        /// <summary>
        /// Экспортирует данные линейной регрессии и график в Microsoft Excel.
        /// </summary>
        /// <param name="chart">График для экспорта.</param>
        /// <param name="xs">Массив X-координат исходных данных.</param>
        /// <param name="ys">Массив Y-координат исходных данных.</param>
        /// <param name="a">Коэффициент наклона линейной регрессии.</param>
        /// <param name="b">Свободный член линейной регрессии.</param>
        public static void ExportToExcel(Charting chart, double[] xs, double[] ys, double a, double b)
        {
            ExportToExcelInternal(chart, xs, ys, FormatEquationLinear(a, b), new[] { a, b }, -1, "Метод наименьших квадратов");
        }

        /// <summary>
        /// Экспортирует данные полиномиальной регрессии и график в Microsoft Excel.
        /// </summary>
        /// <param name="chart">График для экспорта.</param>
        /// <param name="xs">Массив X-координат исходных данных.</param>
        /// <param name="ys">Массив Y-координат исходных данных.</param>
        /// <param name="polyCoefficients">Массив коэффициентов полиномиальной регрессии.</param>
        /// <param name="degree">Степень полинома.</param>
        public static void ExportToExcel(Charting chart, double[] xs, double[] ys, double[] polyCoefficients, int degree)
        {
            ExportToExcelInternal(chart, xs, ys, FormatEquationPolynomial(polyCoefficients, degree), polyCoefficients, degree, $"Метод наименьших квадратов (степень полинома {degree})");
        }

        /// <summary>
        /// Внутренний метод для экспорта данных и графика регрессии в Microsoft Excel.
        /// </summary>
        private static void ExportToExcelInternal(Charting chart, double[] xs, double[] ys, string equation, double[] coeffs, int degree, string title)
        {
            var dlg = new SaveFileDialog
            {
                Title = "Сохранить в Excel",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx",
                FileName = "RegressionResults.xlsx"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            string imgPath = "";
            Excel.Application excelApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                imgPath = SaveChartToPng(chart);
                excelApp = new Excel.Application { Visible = false, DisplayAlerts = false };
                wb = excelApp.Workbooks.Add();
                ws = (Excel.Worksheet)wb.Worksheets[1];
                ws.Name = "Результаты Регрессии";

                Excel.Range cell;

                cell = (Excel.Range)ws.Cells[1, 1];
                cell.Value2 = title;
                Excel.Range mergeRange = ws.Range[ws.Cells[1, 1], ws.Cells[1, 4]];
                mergeRange.Merge();
                cell.Font.Bold = true;
                cell.Font.Size = 14;

                int currentRow = 3;
                cell = (Excel.Range)ws.Cells[currentRow, 1];
                cell.Value2 = "Исходные данные:";
                cell.Font.Bold = true;
                currentRow++;

                ((Excel.Range)ws.Cells[currentRow, 1]).Value2 = "X";
                ((Excel.Range)ws.Cells[currentRow, 2]).Value2 = "Y";
                Excel.Range xyHeaderRange = ws.Range[ws.Cells[currentRow, 1], ws.Cells[currentRow, 2]];
                xyHeaderRange.Font.Bold = true;
                currentRow++;

                if (xs != null && ys != null && xs.Length == ys.Length)
                {
                    for (int i = 0; i < xs.Length; i++)
                    {
                        ((Excel.Range)ws.Cells[currentRow + i, 1]).Value2 = xs[i].ToString(CultureInfo.CurrentCulture);
                        ((Excel.Range)ws.Cells[currentRow + i, 2]).Value2 = ys[i].ToString(CultureInfo.CurrentCulture);
                    }
                    currentRow += xs.Length;
                }
                currentRow++;

                cell = (Excel.Range)ws.Cells[currentRow, 1];
                cell.Value2 = "Уравнение регрессии:";
                cell.Font.Bold = true;
                currentRow++;
                ((Excel.Range)ws.Cells[currentRow, 1]).Value2 = equation;
                currentRow += 2;

                cell = (Excel.Range)ws.Cells[currentRow, 1];
                cell.Value2 = "Коэффициенты:";
                cell.Font.Bold = true;
                currentRow++;

                if (degree == -1)
                {
                    if (coeffs != null && coeffs.Length >= 2)
                    {
                        ((Excel.Range)ws.Cells[currentRow, 1]).Value2 = "a (наклон)";
                        ((Excel.Range)ws.Cells[currentRow, 2]).Value2 = coeffs[0].ToString("F4", CultureInfo.CurrentCulture);
                        currentRow++;
                        ((Excel.Range)ws.Cells[currentRow, 1]).Value2 = "b (свободный член)";
                        ((Excel.Range)ws.Cells[currentRow, 2]).Value2 = coeffs[1].ToString("F4", CultureInfo.CurrentCulture);
                    }
                    else { ((Excel.Range)ws.Cells[currentRow, 1]).Value2 = "(нет данных)"; }
                }
                else
                {
                    if (coeffs != null)
                    {
                        for (int i = 0; i < coeffs.Length; i++)
                        {
                            ((Excel.Range)ws.Cells[currentRow + i, 1]).Value2 = $"c[{i}] (коэф. при x^{i})";
                            ((Excel.Range)ws.Cells[currentRow + i, 2]).Value2 = coeffs[i].ToString("F4", CultureInfo.CurrentCulture);
                        }
                    }
                    else
                    {
                        ((Excel.Range)ws.Cells[currentRow, 1]).Value2 = "(нет данных)";
                    }
                }

                float chartRenderWidth = chart.Width > 0 ? chart.Width : 600;
                float chartRenderHeight = chart.Height > 0 ? chart.Height : 400;
                float chartLeft = 250;
                float chartTopPosition = Convert.ToSingle(((Excel.Range)ws.Cells[3, 1]).Top) + 10;
                float imageWidth = (float)(chartRenderWidth * 0.65);
                float imageHeight = (float)(chartRenderHeight * 0.65);

                ws.Shapes.AddPicture(imgPath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, chartLeft, chartTopPosition, imageWidth, imageHeight);

                ((Excel.Range)ws.Columns["A:B"]).AutoFit();

                wb.SaveAs(dlg.FileName, Excel.XlFileFormat.xlOpenXMLWorkbook, AccessMode: Excel.XlSaveAsAccessMode.xlNoChange);
            }
            finally
            {
                if (ws != null) { Marshal.ReleaseComObject(ws); ws = null; }
                if (wb != null) { wb.Close(false); Marshal.ReleaseComObject(wb); wb = null; }
                if (excelApp != null) { excelApp.Quit(); Marshal.ReleaseComObject(excelApp); excelApp = null; }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                try { if (!string.IsNullOrEmpty(imgPath) && File.Exists(imgPath)) File.Delete(imgPath); } catch { }
            }
        }

        /// <summary>
        /// Экспортирует данные линейной регрессии и график в Microsoft Word.
        /// </summary>
        /// <param name="chart">График для экспорта.</param>
        /// <param name="xs">Массив X-координат исходных данных.</param>
        /// <param name="ys">Массив Y-координат исходных данных.</param>
        /// <param name="a">Коэффициент наклона линейной регрессии.</param>
        /// <param name="b">Свободный член линейной регрессии.</param>
        public static void ExportToWord(Charting chart, double[] xs, double[] ys, double a, double b)
        {
            ExportToWordInternal(chart, xs, ys, FormatEquationLinear(a, b), new[] { a, b }, -1, "Метод наименьших квадратов");
        }

        /// <summary>
        /// Экспортирует данные полиномиальной регрессии и график в Microsoft Word.
        /// </summary>
        /// <param name="chart">График для экспорта.</param>
        /// <param name="xs">Массив X-координат исходных данных.</param>
        /// <param name="ys">Массив Y-координат исходных данных.</param>
        /// <param name="polyCoefficients">Массив коэффициентов полиномиальной регрессии.</param>
        /// <param name="degree">Степень полинома.</param>
        public static void ExportToWord(Charting chart, double[] xs, double[] ys, double[] polyCoefficients, int degree)
        {
            ExportToWordInternal(chart, xs, ys, FormatEquationPolynomial(polyCoefficients, degree), polyCoefficients, degree, $"Метод наименьших квадратов (степень полинома {degree})");
        }

        /// <summary>
        /// Внутренний метод для экспорта данных и графика регрессии в Microsoft Word.
        /// </summary>
        private static void ExportToWordInternal(Charting chart, double[] xs, double[] ys, string equation, double[] coeffs, int degree, string title)
        {
            var dlg = new SaveFileDialog
            {
                Title = "Сохранить в Word",
                Filter = "Word Document (*.docx)|*.docx",
                DefaultExt = "docx",
                FileName = "RegressionResults.docx"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            string imgPath = "";
            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                imgPath = SaveChartToPng(chart);
                wordApp = new Word.Application { Visible = false };
                doc = wordApp.Documents.Add();

                Word.Paragraph p = doc.Content.Paragraphs.Add();
                p.Range.Text = title;
                p.Range.set_Style(Word.WdBuiltinStyle.wdStyleHeading1);
                p.Range.InsertParagraphAfter();

                p = doc.Content.Paragraphs.Add();
                p.Range.Text = "Исходные данные:";
                p.Range.Font.Bold = 1;
                p.Range.InsertParagraphAfter();

                if (xs != null && ys != null && xs.Length == ys.Length && xs.Length > 0)
                {
                    Word.Range tableRange = doc.Bookmarks.get_Item("\\endofdoc").Range;
                    Word.Table table = doc.Tables.Add(tableRange, xs.Length + 1, 2);
                    table.Borders.Enable = 1;
                    table.Cell(1, 1).Range.Text = "X";
                    table.Cell(1, 2).Range.Text = "Y";
                    table.Rows[1].Range.Font.Bold = 1;
                    for (int i = 0; i < xs.Length; i++)
                    {
                        table.Cell(i + 2, 1).Range.Text = xs[i].ToString(CultureInfo.CurrentCulture);
                        table.Cell(i + 2, 2).Range.Text = ys[i].ToString(CultureInfo.CurrentCulture);
                    }
                    p = doc.Content.Paragraphs.Add();
                    p.Range.InsertParagraphAfter();
                }
                else
                {
                    p = doc.Content.Paragraphs.Add();
                    p.Range.Text = "(Нет исходных данных для таблицы)";
                    p.Range.InsertParagraphAfter();
                }

                p = doc.Content.Paragraphs.Add();
                p.Range.Text = "Уравнение регрессии:";
                p.Range.Font.Bold = 1;
                p.Range.InsertParagraphAfter();

                p = doc.Content.Paragraphs.Add();
                p.Range.Text = equation;
                p.Range.InsertParagraphAfter();

                p = doc.Content.Paragraphs.Add();
                p.Range.Text = "Коэффициенты:";
                p.Range.Font.Bold = 1;
                p.Range.InsertParagraphAfter();

                Word.Range coeffRange = doc.Bookmarks.get_Item("\\endofdoc").Range;
                if (degree == -1)
                {
                    if (coeffs != null && coeffs.Length >= 2)
                    {
                        coeffRange.InsertAfter($"a (наклон): {coeffs[0].ToString("F4", CultureInfo.CurrentCulture)}\n");
                        coeffRange.InsertAfter($"b (свободный член): {coeffs[1].ToString("F4", CultureInfo.CurrentCulture)}\n");
                    }
                    else { coeffRange.InsertAfter("(нет данных)\n"); }
                }
                else
                {
                    if (coeffs != null && coeffs.Length > 0)
                    {
                        for (int i = 0; i < coeffs.Length; i++)
                        {
                            coeffRange.InsertAfter($"c[{i}] (коэф. при x^{i}): {coeffs[i].ToString("F4", CultureInfo.CurrentCulture)}\n");
                        }
                    }
                    else
                    {
                        coeffRange.InsertAfter("(нет данных)\n");
                    }
                }

                p = doc.Content.Paragraphs.Add();
                p.Range.Text = "\nГрафик регрессии:";
                p.Range.Font.Bold = 1;
                p.Range.InsertParagraphAfter();

                Word.Range chartInsertRange = doc.Bookmarks.get_Item("\\endofdoc").Range;
                chartInsertRange.InlineShapes.AddPicture(imgPath, LinkToFile: false, SaveWithDocument: true);

                doc.SaveAs2(dlg.FileName);
            }
            finally
            {
                if (doc != null) { doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges); Marshal.ReleaseComObject(doc); doc = null; }
                if (wordApp != null) { wordApp.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); Marshal.ReleaseComObject(wordApp); wordApp = null; }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                try { if (!string.IsNullOrEmpty(imgPath) && File.Exists(imgPath)) File.Delete(imgPath); } catch { }
            }
        }

        /// <summary>
        /// Экспортирует данные линейной регрессии и график в Microsoft PowerPoint.
        /// </summary>
        /// <param name="chart">График для экспорта.</param>
        /// <param name="xs">Массив X-координат исходных данных.</param>
        /// <param name="ys">Массив Y-координат исходных данных.</param>
        /// <param name="a">Коэффициент наклона линейной регрессии.</param>
        /// <param name="b">Свободный член линейной регрессии.</param>
        /// <param name="topic">Тема презентации (для титульного слайда).</param>
        /// <param name="author">Автор презентации.</param>
        /// <param name="discipline">Дисциплина (для титульного слайда).</param>
        public static void ExportToPowerPoint(Charting chart, double[] xs, double[] ys, double a, double b, string topic, string author, string discipline)
        {
            ExportToPowerPointInternal(chart, xs, ys, FormatEquationLinear(a, b), new[] { a, b }, -1, topic, author, discipline);
        }

        /// <summary>
        /// Экспортирует данные полиномиальной регрессии и график в Microsoft PowerPoint.
        /// </summary>
        /// <param name="chart">График для экспорта.</param>
        /// <param name="xs">Массив X-координат исходных данных.</param>
        /// <param name="ys">Массив Y-координат исходных данных.</param>
        /// <param name="polyCoefficients">Массив коэффициентов полиномиальной регрессии.</param>
        /// <param name="degree">Степень полинома.</param>
        /// <param name="topic">Тема презентации (для титульного слайда).</param>
        /// <param name="author">Автор презентации.</param>
        /// <param name="discipline">Дисциплина (для титульного слайда).</param>
        public static void ExportToPowerPoint(Charting chart, double[] xs, double[] ys, double[] polyCoefficients, int degree, string topic, string author, string discipline)
        {
            ExportToPowerPointInternal(chart, xs, ys, FormatEquationPolynomial(polyCoefficients, degree), polyCoefficients, degree, topic, author, discipline);
        }

        /// <summary>
        /// Внутренний метод для экспорта данных и графика регрессии в Microsoft PowerPoint.
        /// </summary>
        private static void ExportToPowerPointInternal(Charting chart, double[] xs, double[] ys, string equation, double[] coeffs, int degree, string topic, string author, string discipline)
        {
            var dlg = new SaveFileDialog
            {
                Title = "Сохранить в PowerPoint",
                Filter = "PowerPoint Presentation (*.pptx)|*.pptx",
                DefaultExt = "pptx",
                FileName = "RegressionResults.pptx"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            string imgPath = "";
            PPT.Application pptApp = null;
            PPT.Presentation pres = null;
            PPT.Slide s1 = null, s2 = null, s3 = null;

            try
            {
                imgPath = SaveChartToPng(chart);
                pptApp = new PPT.Application();
                pres = pptApp.Presentations.Add(Office.MsoTriState.msoTrue);

                s1 = pres.Slides.Add(1, PPT.PpSlideLayout.ppLayoutTitle);
                if (s1.Shapes.HasTitle == Office.MsoTriState.msoTrue && s1.Shapes.Title != null)
                {
                    s1.Shapes.Title.TextFrame.TextRange.Text = topic;
                }

                PPT.Shape subtitlePlaceholder = null;
                if (s1.Shapes.Placeholders.Count >= 2)
                {
                    subtitlePlaceholder = s1.Shapes.Placeholders[2];
                }
                else if (s1.Shapes.Placeholders.Count == 1 && s1.Shapes.HasTitle == Office.MsoTriState.msoFalse)
                {
                    subtitlePlaceholder = s1.Shapes.Placeholders[1];
                }

                if (subtitlePlaceholder != null && subtitlePlaceholder.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    subtitlePlaceholder.TextFrame.TextRange.Text = $"Автор: {author}\nДисциплина: {discipline}\nДата: {DateTime.Now:dd.MM.yyyy}";
                }

                s2 = pres.Slides.Add(2, PPT.PpSlideLayout.ppLayoutText);
                if (s2.Shapes.HasTitle == Office.MsoTriState.msoTrue && s2.Shapes.Title != null)
                {
                    s2.Shapes.Title.TextFrame.TextRange.Text = "Исходные данные и результаты";
                }

                var sb = new StringBuilder();
                sb.AppendLine("Исходные точки (X | Y):");
                int maxPointsToShow = 10;
                if (xs != null && ys != null && xs.Length == ys.Length)
                {
                    for (int i = 0; i < Math.Min(xs.Length, maxPointsToShow); i++)
                        sb.AppendLine($"{xs[i].ToString("0.##", CultureInfo.InvariantCulture),8} | {ys[i].ToString("0.##", CultureInfo.InvariantCulture),8}");
                    if (xs.Length > maxPointsToShow) sb.AppendLine("...");
                }
                else { sb.AppendLine("(Нет данных)"); }
                sb.AppendLine();
                sb.AppendLine("Уравнение регрессии:");
                sb.AppendLine(equation);
                sb.AppendLine();
                sb.AppendLine("Коэффициенты:");
                if (degree == -1)
                {
                    if (coeffs != null && coeffs.Length >= 2)
                    {
                        sb.AppendLine($"  a (наклон): {coeffs[0].ToString("F4", CultureInfo.InvariantCulture)}");
                        sb.AppendLine($"  b (свободный член): {coeffs[1].ToString("F4", CultureInfo.InvariantCulture)}");
                    }
                    else { sb.AppendLine("  (нет данных)"); }
                }
                else
                {
                    if (coeffs != null && coeffs.Length > 0)
                    {
                        for (int i = 0; i < coeffs.Length; i++)
                        {
                            sb.AppendLine($"  c[{i}] (коэф. при x^{i}): {coeffs[i].ToString("F4", CultureInfo.InvariantCulture)}");
                        }
                    }
                    else
                    {
                        sb.AppendLine("  (нет данных)");
                    }
                }

                PPT.Shape bodyPlaceholderS2 = null;
                if (s2.Shapes.Placeholders.Count >= 2)
                {
                    bodyPlaceholderS2 = s2.Shapes.Placeholders[2];
                }
                else if (s2.Shapes.Placeholders.Count == 1 && s2.Shapes.HasTitle == Office.MsoTriState.msoFalse)
                {
                    bodyPlaceholderS2 = s2.Shapes.Placeholders[1];
                }

                if (bodyPlaceholderS2 != null && bodyPlaceholderS2.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    PPT.TextRange contentTextRange = bodyPlaceholderS2.TextFrame.TextRange;
                    contentTextRange.Text = sb.ToString();
                    contentTextRange.Font.Name = "Consolas";
                    contentTextRange.Font.Size = 12;
                }

                s3 = pres.Slides.Add(3, PPT.PpSlideLayout.ppLayoutTitleOnly);
                if (s3.Shapes.HasTitle == Office.MsoTriState.msoTrue && s3.Shapes.Title != null)
                {
                    s3.Shapes.Title.TextFrame.TextRange.Text = "График регрессии";
                }

                float slideWidth = pres.PageSetup.SlideWidth;
                float slideHeight = pres.PageSetup.SlideHeight;

                float chartRenderWidth = chart.Width > 0 ? chart.Width : 600;
                float chartRenderHeight = chart.Height > 0 ? chart.Height : 400;

                float imageScaleFactor = 0.70f;
                float imageWidth = chartRenderWidth * imageScaleFactor;
                float imageHeight = chartRenderHeight * imageScaleFactor;

                float titleShapeHeightValue = 0;
                if (s3.Shapes.HasTitle == Office.MsoTriState.msoTrue && s3.Shapes.Title != null)
                {
                    titleShapeHeightValue = s3.Shapes.Title.Height;
                }

                float left = (slideWidth - imageWidth) / 2;
                float topPosition = titleShapeHeightValue + 20;
                if (left < 10) left = 10;
                if (topPosition < titleShapeHeightValue + 5) topPosition = titleShapeHeightValue + 5;

                if (topPosition + imageHeight > slideHeight - 20)
                {
                    imageHeight = slideHeight - topPosition - 20;
                    if (imageHeight < 50) imageHeight = 50;
                    imageWidth = (chartRenderWidth / (chartRenderHeight > 0 ? chartRenderHeight : 1)) * imageHeight;
                }
                if (left + imageWidth > slideWidth - 10)
                {
                    imageWidth = slideWidth - left - 10;
                    if (imageWidth < 50) imageWidth = 50;
                    imageHeight = ((chartRenderHeight > 0 ? chartRenderHeight : 1) / (chartRenderWidth > 0 ? chartRenderWidth : 1)) * imageWidth;
                    if (topPosition + imageHeight > slideHeight - 20)
                    {
                        imageHeight = slideHeight - topPosition - 20;
                        if (imageHeight < 50) imageHeight = 50;
                    }
                }

                s3.Shapes.AddPicture(imgPath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, left, topPosition, imageWidth, imageHeight);

                pres.SaveAs(dlg.FileName, PPT.PpSaveAsFileType.ppSaveAsDefault);
            }
            finally
            {
                if (s1 != null) { Marshal.ReleaseComObject(s1); s1 = null; }
                if (s2 != null) { Marshal.ReleaseComObject(s2); s2 = null; }
                if (s3 != null) { Marshal.ReleaseComObject(s3); s3 = null; }

                if (pres != null) { pres.Close(); Marshal.ReleaseComObject(pres); pres = null; }
                if (pptApp != null) { pptApp.Quit(); Marshal.ReleaseComObject(pptApp); pptApp = null; }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                try { if (!string.IsNullOrEmpty(imgPath) && File.Exists(imgPath)) File.Delete(imgPath); } catch { }
            }
        }
    }
}