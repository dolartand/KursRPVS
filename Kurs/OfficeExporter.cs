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
using Charting = System.Windows.Forms.DataVisualization.Charting;
using System.Diagnostics;

namespace Kurs
{
    /// <summary>
    /// Класс для экспорта результатов расчета в форматы Microsoft Office
    /// </summary>
    public static class OfficeExporter
    {
        /// <summary>
        /// Сохраняет график в PNG-файл во временной директории
        /// </summary>
        /// <param name="chart">Объект графика для сохранения</param>
        /// <returns>Путь к временному PNG-файлу</returns>
        private static string SaveChartToPng(Charting.Chart chart)
        {
            string tmp = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".png");
            using (var bmp = new Bitmap(chart.Width, chart.Height))
            {
                chart.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height));
                bmp.Save(tmp, ImageFormat.Png);
            }
            return tmp;
        }

        /// <summary>
        /// Экспортирует данные и график в документ Excel
        /// </summary>
        /// <param name="chart">График для экспорта</param>
        /// <param name="xs">Массив значений X</param>
        /// <param name="ys">Массив значений Y</param>
        /// <param name="a">Коэффициент наклона</param>
        /// <param name="b">Свободный член</param>
        /// <exception cref="System.Runtime.InteropServices.COMException">
        /// Ошибка взаимодействия с приложением Excel
        /// </exception>
        public static void ExportToExcel(Charting.Chart chart, double[] xs, double[] ys, double a, double b)
        {
            var dlg = new SaveFileDialog
            {
                Title = "Сохранить в Excel",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx",
                FileName = "LeastSquares.xlsx"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            string imgPath = SaveChartToPng(chart);

            var excel = new Excel.Application();
            var wb = excel.Workbooks.Add();
            var ws = (Excel.Worksheet)wb.Worksheets[1];
            ws.Name = "Data";

            int n = xs.Length;
            ws.Cells[1, 1] = "X"; ws.Cells[1, 2] = "Y";
            for (int i = 0; i < n; i++)
            {
                ws.Cells[i + 2, 1] = xs[i];
                ws.Cells[i + 2, 2] = ys[i];
            }

            int solRow = n + 4;
            ws.Cells[solRow, 1] = "Коэффициенты:";
            ws.Cells[solRow + 1, 1] = "a"; ws.Cells[solRow + 1, 2] = a;
            ws.Cells[solRow + 2, 1] = "b"; ws.Cells[solRow + 2, 2] = b;

            ws.Shapes.AddPicture(
                imgPath,
                Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoTrue,
                300, 10,
                (float)(chart.Width * 0.75),
                (float)(chart.Height * 0.75)
            );

            wb.SaveAs(
                dlg.FileName,
                Excel.XlFileFormat.xlOpenXMLWorkbook,
                Type.Missing, Type.Missing,
                false, false,
                Excel.XlSaveAsAccessMode.xlNoChange);
            wb.Close(false);
            excel.Quit();

            try { File.Delete(imgPath); } catch { }
        }

        /// <summary>
        /// Экспортирует данные и график в документ Word
        /// </summary>
        /// <param name="chart">График для экспорта</param>
        /// <param name="xs">Массив значений X</param>
        /// <param name="ys">Массив значений Y</param>
        /// <param name="a">Коэффициент наклона</param>
        /// <param name="b">Свободный член</param>
        /// <exception cref="System.Runtime.InteropServices.COMException">
        /// Ошибка взаимодействия с приложением Word
        /// </exception>
        public static void ExportToWord(Charting.Chart chart, double[] xs, double[] ys, double a, double b)
        {
            var dlg = new SaveFileDialog
            {
                Title = "Сохранить в Word",
                Filter = "Word Document (*.docx)|*.docx",
                DefaultExt = "docx",
                FileName = "LeastSquares.docx"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            string imgPath = SaveChartToPng(chart);

            var word = new Word.Application();
            var doc = word.Documents.Add();

            var p1 = doc.Content.Paragraphs.Add();
            p1.Range.Text = "Результаты метода наименьших квадратов";
            p1.Range.set_Style(Word.WdBuiltinStyle.wdStyleHeading1);
            p1.Range.InsertParagraphAfter();

            int n = xs.Length;
            var table = doc.Tables.Add(doc.Bookmarks.get_Item("\\endofdoc").Range, n + 1, 2);
            table.Cell(1, 1).Range.Text = "X";
            table.Cell(1, 2).Range.Text = "Y";
            for (int i = 0; i < n; i++)
            {
                table.Cell(i + 2, 1).Range.Text = xs[i].ToString("0.##");
                table.Cell(i + 2, 2).Range.Text = ys[i].ToString("0.##");
            }
            doc.Bookmarks.get_Item("\\endofdoc").Range.InsertParagraphAfter();

            var p2 = doc.Bookmarks.get_Item("\\endofdoc").Range.Paragraphs.Add();
            p2.Range.Text = $"Прямая: y = {a:0.##}·x + {b:0.##}";
            p2.Range.InsertParagraphAfter();

            var rng = doc.Bookmarks.get_Item("\\endofdoc").Range;
            doc.InlineShapes.AddPicture(
                imgPath,
                LinkToFile: false,
                SaveWithDocument: true,
                Range: rng
            );

            doc.SaveAs2(dlg.FileName);
            doc.Close();
            word.Quit();

            try { File.Delete(imgPath); } catch { }
        }

        /// <summary>
        /// Экспортирует данные и график в презентацию PowerPoint
        /// </summary>
        /// <param name="chart">График для экспорта</param>
        /// <param name="xs">Массив значений X</param>
        /// <param name="ys">Массив значений Y</param>
        /// <param name="a">Коэффициент наклона</param>
        /// <param name="b">Свободный член</param>
        /// <param name="topic">Тема исследования</param>
        /// <param name="author">Автор работы</param>
        /// <param name="discipline">Название дисциплины</param>
        /// <exception cref="System.Runtime.InteropServices.COMException">
        /// Ошибка взаимодействия с приложением PowerPoint
        /// </exception>
        public static void ExportToPowerPoint(
            Charting.Chart chart,
            double[] xs, double[] ys,
            double a, double b,
            string topic,
            string author,
            string discipline)
        {
            var dlg = new SaveFileDialog
            {
                Title = "Сохранить в PowerPoint",
                Filter = "PowerPoint Presentation (*.pptx)|*.pptx",
                DefaultExt = "pptx",
                FileName = "LeastSquares.pptx"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            string imgPath = SaveChartToPng(chart);

            var ppt = new PPT.Application();
            var pres = ppt.Presentations.Add(Office.MsoTriState.msoTrue);

            var s1 = pres.Slides.Add(1, PPT.PpSlideLayout.ppLayoutTitle);
            s1.Shapes.Title.TextFrame.TextRange.Text = topic;
            s1.Shapes[2].TextFrame.TextRange.Text = $"Автор: {author}\nДисциплина: {discipline}";

            var s2 = pres.Slides.Add(2, PPT.PpSlideLayout.ppLayoutText);
            s2.Shapes[1].TextFrame.TextRange.Text = "Таблица точек и коэффициенты";
            var sb = new StringBuilder();
            for (int i = 0; i < xs.Length; i++)
                sb.AppendLine($"{xs[i],6:0.##} | {ys[i],6:0.##}");
            sb.AppendLine();
            sb.AppendLine($"Прямая: y = {a:0.##}·x + {b:0.##}");
            s2.Shapes[2].TextFrame.TextRange.Text = sb.ToString();

            var s3 = pres.Slides.Add(3, PPT.PpSlideLayout.ppLayoutBlank);
            s3.Shapes.AddPicture(
                imgPath,
                Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoTrue,
                50, 50,
                chart.Width * 0.75f,
                chart.Height * 0.75f
            );

            pres.SaveAs(dlg.FileName);
            pres.Close();
            ppt.Quit();

            try { File.Delete(imgPath); } catch { }

            Process.Start(new ProcessStartInfo(dlg.FileName) { UseShellExecute = true });
        }
    }
}
