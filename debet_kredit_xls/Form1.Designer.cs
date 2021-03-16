namespace debet_kredit_xls
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.добавитьДанныеИзExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.очитьДанныеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.сохранитьОтсортированныеДанныеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.ContextMenuStrip = this.contextMenuStrip1;
            this.dataGridView1.Location = new System.Drawing.Point(2, 47);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(482, 508);
            this.dataGridView1.TabIndex = 0;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.добавитьДанныеИзExcelToolStripMenuItem,
            this.очитьДанныеToolStripMenuItem,
            this.сохранитьОтсортированныеДанныеToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(290, 70);
            // 
            // добавитьДанныеИзExcelToolStripMenuItem
            // 
            this.добавитьДанныеИзExcelToolStripMenuItem.Name = "добавитьДанныеИзExcelToolStripMenuItem";
            this.добавитьДанныеИзExcelToolStripMenuItem.Size = new System.Drawing.Size(289, 22);
            this.добавитьДанныеИзExcelToolStripMenuItem.Text = "Добавить данные из excel...";
            this.добавитьДанныеИзExcelToolStripMenuItem.Click += new System.EventHandler(this.добавитьДанныеИзExcelToolStripMenuItem_Click);
            // 
            // очитьДанныеToolStripMenuItem
            // 
            this.очитьДанныеToolStripMenuItem.Name = "очитьДанныеToolStripMenuItem";
            this.очитьДанныеToolStripMenuItem.Size = new System.Drawing.Size(289, 22);
            this.очитьДанныеToolStripMenuItem.Text = "Очить данные";
            this.очитьДанныеToolStripMenuItem.Click += new System.EventHandler(this.очитьДанныеToolStripMenuItem_Click);
            // 
            // сохранитьОтсортированныеДанныеToolStripMenuItem
            // 
            this.сохранитьОтсортированныеДанныеToolStripMenuItem.Name = "сохранитьОтсортированныеДанныеToolStripMenuItem";
            this.сохранитьОтсортированныеДанныеToolStripMenuItem.Size = new System.Drawing.Size(289, 22);
            this.сохранитьОтсортированныеДанныеToolStripMenuItem.Text = "Сохранить отфильтрованные данные...";
            this.сохранитьОтсортированныеДанныеToolStripMenuItem.Click += new System.EventHandler(this.сохранитьОтсортированныеДанныеToolStripMenuItem_Click);
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "Все",
            "Должники",
            "Без долга"});
            this.comboBox1.Location = new System.Drawing.Point(68, 20);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(121, 21);
            this.comboBox1.TabIndex = 1;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(50, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Фильтр;";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(486, 557);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.dataGridView1);
            this.Name = "Form1";
            this.Text = "Дебет_Кредит v.10";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem добавитьДанныеИзExcelToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem очитьДанныеToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem сохранитьОтсортированныеДанныеToolStripMenuItem;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label1;
    }
}

