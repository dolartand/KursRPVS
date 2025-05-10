namespace Kurs
{
    /// <summary>
    /// Стартовая форма приложения с навигацией в главное окно
    /// </summary>
    public partial class StartForm : Form
    {
        /// <summary>
        /// Инициализирует стартовую форму и устанавливает заголовок окна
        /// </summary>
        public StartForm()
        {
            InitializeComponent();
            this.Text = "Программная реализация метода наименьших квадратов";
        }

        /// <summary>
        /// Обработчик клика для кнопки выхода из приложения
        /// </summary>
        private void btnCloseProgram_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        /// <summary>
        /// Обработчик открытия главного окна приложения
        /// </summary>
        /// <remarks>
        /// Скрывает стартовую форму и создает главное окно.
        /// При закрытии главного окна автоматически закрывает стартовую форму.
        /// </remarks>
        private void openMainWindow(object sender, EventArgs e)
        {
            this.Hide();
            var main = new MainForm();
            main.FormClosed += (s, args) => this.Close();
            main.Show();
        }
    }
}