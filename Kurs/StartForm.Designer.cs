using static System.Drawing.Font;
using System.Windows.Forms;

namespace Kurs
{
    partial class StartForm
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            label1 = new Label();
            label2 = new Label();
            label3 = new Label();
            label4 = new Label();
            label5 = new Label();
            label6 = new Label();
            label7 = new Label();
            label8 = new Label();
            label9 = new Label();
            btnCloseProgram = new Button();
            button2 = new Button();
            SuspendLayout();
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Arial", 10.2F);
            label1.Location = new Point(186, 31);
            label1.Name = "label1";
            label1.Size = new Size(519, 19);
            label1.TabIndex = 0;
            label1.Text = "БЕЛАРУСКИЙ НАЦИОНАЛЬНЫЙ ТЕХНИЧЕСКИЙ УНИВЕРСИТЕТ ";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Arial", 10.2F);
            label2.Location = new Point(206, 80);
            label2.Name = "label2";
            label2.Size = new Size(448, 19);
            label2.TabIndex = 1;
            label2.Text = "Факультет информационных технологий и робототехники";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Font = new Font("Arial", 10.2F);
            label3.Location = new Point(153, 107);
            label3.Name = "label3";
            label3.Size = new Size(600, 19);
            label3.TabIndex = 2;
            label3.Text = "Кафедра программного обеспечения информационных систем и  технологий";
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Font = new Font("Arial", 10.8F, FontStyle.Bold, GraphicsUnit.Point, 204);
            label4.Location = new Point(330, 170);
            label4.Name = "label4";
            label4.Size = new Size(163, 21);
            label4.TabIndex = 3;
            label4.Text = "Курсовая работа";
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Font = new Font("Arial", 10.8F, FontStyle.Bold, GraphicsUnit.Point, 204);
            label5.Location = new Point(138, 204);
            label5.Name = "label5";
            label5.Size = new Size(589, 21);
            label5.TabIndex = 4;
            label5.Text = "по дисциплине \"Разработка приложений в визуальных средах\"";
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Font = new Font("Arial Narrow", 12F, FontStyle.Bold, GraphicsUnit.Point, 204);
            label6.Location = new Point(150, 261);
            label6.Name = "label6";
            label6.Size = new Size(577, 24);
            label6.TabIndex = 5;
            label6.Text = "Разработка программной реализации метода наименьших квадратов";
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Font = new Font("Arial", 10.8F);
            label7.Location = new Point(366, 391);
            label7.Name = "label7";
            label7.Size = new Size(433, 21);
            label7.TabIndex = 6;
            label7.Text = "Выполнил: студент группы 10701323 Доленко А.А.";
            // 
            // label8
            // 
            label8.AutoSize = true;
            label8.Font = new Font("Arial", 10.8F);
            label8.Location = new Point(366, 425);
            label8.Name = "label8";
            label8.Size = new Size(259, 21);
            label8.TabIndex = 7;
            label8.Text = "Принял: доцент Гурский Н. Н.";
            // 
            // label9
            // 
            label9.AutoSize = true;
            label9.Font = new Font("Arial", 10.8F);
            label9.Location = new Point(386, 631);
            label9.Name = "label9";
            label9.Size = new Size(107, 21);
            label9.TabIndex = 8;
            label9.Text = "Минск 2025";
            // 
            // btnCloseProgram
            // 
            btnCloseProgram.BackColor = SystemColors.MenuHighlight;
            btnCloseProgram.Font = new Font("Arial", 10.8F);
            btnCloseProgram.ForeColor = Color.White;
            btnCloseProgram.Location = new Point(502, 667);
            btnCloseProgram.Name = "btnCloseProgram";
            btnCloseProgram.Size = new Size(302, 61);
            btnCloseProgram.TabIndex = 10;
            btnCloseProgram.Text = "Выход";
            btnCloseProgram.UseVisualStyleBackColor = false;
            btnCloseProgram.Click += btnCloseProgram_Click;
            // 
            // button2
            // 
            button2.BackColor = SystemColors.MenuHighlight;
            button2.Font = new Font("Arial", 10.8F);
            button2.ForeColor = Color.White;
            button2.Location = new Point(28, 667);
            button2.Name = "button2";
            button2.Size = new Size(326, 61);
            button2.TabIndex = 11;
            button2.Text = "Далее";
            button2.UseVisualStyleBackColor = false;
            button2.Click += openMainWindow;
            // 
            // StartForm
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(833, 742);
            Controls.Add(button2);
            Controls.Add(btnCloseProgram);
            Controls.Add(label9);
            Controls.Add(label8);
            Controls.Add(label7);
            Controls.Add(label6);
            Controls.Add(label5);
            Controls.Add(label4);
            Controls.Add(label3);
            Controls.Add(label2);
            Controls.Add(label1);
            Name = "StartForm";
            Text = "Form1";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
        private Label label5;
        private Label label6;
        private Label label7;
        private Label label8;
        private Label label9;
        private Button btnCloseProgram;
        private Button button2;
    }
}