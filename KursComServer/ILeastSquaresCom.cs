using System.Runtime.InteropServices;

namespace KursComServer
{
    /// <summary>
    /// COM-интерфейс для работы с методом наименьших квадратов
    /// </summary>
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("39856D54-B3DE-4EC9-A547-056EFF777C35")]
    public interface ILeastSquaresCom
    {
        /// <summary>
        /// Загружает данные из INI-файла
        /// </summary>
        /// <param name="path">Путь к INI-файлу с данными</param>
        void LoadFromIni(string path);

        /// <summary>
        /// Сохраняет текущие данные в INI-файл
        /// </summary>
        /// <param name="path">Путь для сохранения файла</param>
        void SaveToIni(string path);

        /// <summary>
        /// Вычисляет коэффициенты линейной регрессии
        /// </summary>
        /// <param name="a">Коэффициент наклона (выходной параметр)</param>
        /// <param name="b">Свободный член (выходной параметр)</param>
        void Compute(out double a, out double b);

        /// <summary>
        /// Устанавливает данные напрямую без использования файла
        /// </summary>
        /// <param name="xs">Массив значений X</param>
        /// <param name="ys">Массив значений Y</param>
        void SetData(double[] xs, double[] ys);
    }
}
