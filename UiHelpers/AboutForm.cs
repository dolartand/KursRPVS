using System;
using System.Drawing;
using System.Windows.Forms;

namespace UIHelpers
{
    /// <summary>
    /// Форма "О программе" для отображения информации о приложении.
    /// </summary>
    public class AboutForm : Form
    {
        private Label lblTitle;
        private Label lblVersion;
        private Label lblAuthor;
        private Button btnClose;

        /// <summary>
        /// Конструктор формы, инициализирующий компоненты
        /// </summary>
        public AboutForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Инициализация компонентов формы
        /// </summary>
        /// <remarks>
        /// Настройки формы:
        /// - Фиксированный размер окна 400x180
        /// - Центрирование относительно родителя
        /// - Отключение кнопок максимизации/минимизации
        /// </remarks>
        private void InitializeComponent()
        {
            this.Text = "О программе";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.ClientSize = new System.Drawing.Size(400, 180);

            lblTitle = new Label
            {
                Text = "Метод наименьших квадратов",
                AutoSize = true,
                Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold),
                Location = new System.Drawing.Point(20, 20)
            };

            lblVersion = new Label
            {
                Text = $"Версия: {Application.ProductVersion}",
                AutoSize = true,
                Location = new System.Drawing.Point(20, 60)
            };

            lblAuthor = new Label
            {
                Text = "Автор: Доленко А.А.",
                AutoSize = true,
                Location = new System.Drawing.Point(20, 90)
            };

            btnClose = new Button
            {
                Text = "Закрыть",
                DialogResult = DialogResult.OK,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Size = new System.Drawing.Size(80, 30),
                Location = new System.Drawing.Point(200, 130)
            };

            this.Controls.Add(lblTitle);
            this.Controls.Add(lblVersion);
            this.Controls.Add(lblAuthor);
            this.Controls.Add(btnClose);
        }
    }
}
