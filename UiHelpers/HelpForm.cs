using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UIHelpers
{
    /// <summary>
    /// Форма для отображения HTML-справки приложения
    /// </summary>
    /// <remarks>
    /// Автоматически ищет файл справки index.html в поддиректории "help"
    /// рядом с исполняемым файлом приложения. Использует компонент WebBrowser
    /// для отображения HTML-контента.
    /// </remarks>
    public partial class HelpForm : Form
    {
        /// <summary>
        /// Инициализирует форму справки и загружает содержание
        /// </summary>
        public HelpForm()
        {
            InitializeComponent();
            LoadIndex();
        }

        /// <summary>
        /// Загружает главную страницу справки из HTML-файла
        /// </summary>
        /// <remarks>
        /// Логика работы:
        /// 1. Формирует путь к help/Index.html относительно EXE-файла
        /// 2. Проверяет существование файла
        /// 3. При ошибке показывает сообщение и закрывает форму
        /// 4. Загружает контент в WebBrowser
        /// </remarks>
        private void LoadIndex()
        {
            string helpDir = Path.Combine(Application.StartupPath, "help");
            string indexPath = Path.Combine(helpDir, "index.html");
            if (!File.Exists(indexPath))
            {
                MessageBox.Show($"Не найден файл справки: {indexPath}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close(); return;
            }
            webBrowser1.Url = new Uri(indexPath);
        }
    }
}
