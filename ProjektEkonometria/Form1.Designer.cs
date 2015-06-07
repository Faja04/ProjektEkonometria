namespace ProjektEkonometria
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.getExelY = new System.Windows.Forms.Button();
            this.getExelX = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.dataGridView3 = new System.Windows.Forms.DataGridView();
            this.X1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.X2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.X3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.X4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridView4 = new System.Windows.Forms.DataGridView();
            this.Y1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.rStar = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.rowMod = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.r2text = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.odchylenie = new System.Windows.Forms.TextBox();
            this.dataGridView5 = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.daDGV = new System.Windows.Forms.DataGridView();
            this.label8 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.wspLos = new System.Windows.Forms.TextBox();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.daDGV)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 23);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.Size = new System.Drawing.Size(338, 225);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // dataGridView2
            // 
            this.dataGridView2.AllowUserToAddRows = false;
            this.dataGridView2.AllowUserToDeleteRows = false;
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Location = new System.Drawing.Point(356, 23);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.ReadOnly = true;
            this.dataGridView2.Size = new System.Drawing.Size(137, 225);
            this.dataGridView2.TabIndex = 1;
            this.dataGridView2.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView2_CellContentClick);
            // 
            // getExelY
            // 
            this.getExelY.Location = new System.Drawing.Point(356, 254);
            this.getExelY.Name = "getExelY";
            this.getExelY.Size = new System.Drawing.Size(137, 23);
            this.getExelY.TabIndex = 3;
            this.getExelY.Text = "Pobierz dane Y";
            this.getExelY.UseVisualStyleBackColor = true;
            this.getExelY.Click += new System.EventHandler(this.getExelY_Click);
            // 
            // getExelX
            // 
            this.getExelX.Location = new System.Drawing.Point(12, 254);
            this.getExelX.Name = "getExelX";
            this.getExelX.Size = new System.Drawing.Size(338, 23);
            this.getExelX.TabIndex = 4;
            this.getExelX.Text = "Pobierz dane X";
            this.getExelX.UseVisualStyleBackColor = true;
            this.getExelX.Click += new System.EventHandler(this.getExelX_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label1.Location = new System.Drawing.Point(9, 7);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(224, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Pobieranie danych (plik Excel musi być otwrty)";
            // 
            // dataGridView3
            // 
            this.dataGridView3.AllowUserToAddRows = false;
            this.dataGridView3.AllowUserToDeleteRows = false;
            this.dataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView3.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.X1,
            this.X2,
            this.X3,
            this.X4});
            this.dataGridView3.Location = new System.Drawing.Point(12, 319);
            this.dataGridView3.Name = "dataGridView3";
            this.dataGridView3.ReadOnly = true;
            this.dataGridView3.Size = new System.Drawing.Size(338, 126);
            this.dataGridView3.TabIndex = 6;
            this.dataGridView3.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView3_CellContentClick);
            // 
            // X1
            // 
            this.X1.HeaderText = "X1";
            this.X1.Name = "X1";
            this.X1.ReadOnly = true;
            this.X1.Width = 70;
            // 
            // X2
            // 
            this.X2.HeaderText = "X2";
            this.X2.Name = "X2";
            this.X2.ReadOnly = true;
            this.X2.Width = 70;
            // 
            // X3
            // 
            this.X3.HeaderText = "X3";
            this.X3.Name = "X3";
            this.X3.ReadOnly = true;
            this.X3.Width = 70;
            // 
            // X4
            // 
            this.X4.HeaderText = "X4";
            this.X4.Name = "X4";
            this.X4.ReadOnly = true;
            this.X4.Width = 70;
            // 
            // dataGridView4
            // 
            this.dataGridView4.AllowUserToAddRows = false;
            this.dataGridView4.AllowUserToDeleteRows = false;
            this.dataGridView4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView4.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Y1});
            this.dataGridView4.Location = new System.Drawing.Point(356, 319);
            this.dataGridView4.Name = "dataGridView4";
            this.dataGridView4.ReadOnly = true;
            this.dataGridView4.Size = new System.Drawing.Size(137, 126);
            this.dataGridView4.TabIndex = 7;
            // 
            // Y1
            // 
            this.Y1.HeaderText = "R0";
            this.Y1.Name = "Y1";
            this.Y1.ReadOnly = true;
            this.Y1.Width = 70;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label2.Location = new System.Drawing.Point(9, 295);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(117, 13);
            this.label2.TabIndex = 8;
            this.label2.Text = "Współczynnik korelacji";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(135, 290);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(358, 23);
            this.button1.TabIndex = 9;
            this.button1.Text = "Oblicz";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // rStar
            // 
            this.rStar.Location = new System.Drawing.Point(545, 46);
            this.rStar.Name = "rStar";
            this.rStar.ReadOnly = true;
            this.rStar.Size = new System.Drawing.Size(182, 20);
            this.rStar.TabIndex = 10;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(582, 27);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(105, 13);
            this.label3.TabIndex = 11;
            this.label3.Text = "Wartość krytyczna r*";
            // 
            // rowMod
            // 
            this.rowMod.Location = new System.Drawing.Point(545, 85);
            this.rowMod.Name = "rowMod";
            this.rowMod.ReadOnly = true;
            this.rowMod.Size = new System.Drawing.Size(182, 20);
            this.rowMod.TabIndex = 12;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(582, 69);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(92, 13);
            this.label4.TabIndex = 13;
            this.label4.Text = "Równanie modelu";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(559, 108);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(157, 13);
            this.label5.TabIndex = 14;
            this.label5.Text = "Współczynnik determinacji R^2";
            // 
            // r2text
            // 
            this.r2text.Location = new System.Drawing.Point(545, 124);
            this.r2text.Name = "r2text";
            this.r2text.ReadOnly = true;
            this.r2text.Size = new System.Drawing.Size(182, 20);
            this.r2text.TabIndex = 15;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(519, 147);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(247, 13);
            this.label7.TabIndex = 18;
            this.label7.Text = "Odchylenie standardowe składnika resztowego Su";
            // 
            // odchylenie
            // 
            this.odchylenie.Location = new System.Drawing.Point(545, 163);
            this.odchylenie.Name = "odchylenie";
            this.odchylenie.ReadOnly = true;
            this.odchylenie.Size = new System.Drawing.Size(182, 20);
            this.odchylenie.TabIndex = 19;
            // 
            // dataGridView5
            // 
            this.dataGridView5.AllowUserToAddRows = false;
            this.dataGridView5.AllowUserToDeleteRows = false;
            this.dataGridView5.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView5.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1});
            this.dataGridView5.Location = new System.Drawing.Point(775, 23);
            this.dataGridView5.Name = "dataGridView5";
            this.dataGridView5.ReadOnly = true;
            this.dataGridView5.Size = new System.Drawing.Size(143, 211);
            this.dataGridView5.TabIndex = 20;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Y\'";
            this.Column1.Name = "Column1";
            this.Column1.ReadOnly = true;
            // 
            // daDGV
            // 
            this.daDGV.AllowUserToAddRows = false;
            this.daDGV.AllowUserToDeleteRows = false;
            this.daDGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.daDGV.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn2,
            this.dataGridViewTextBoxColumn3,
            this.dataGridViewTextBoxColumn4});
            this.daDGV.Location = new System.Drawing.Point(506, 254);
            this.daDGV.Name = "daDGV";
            this.daDGV.ReadOnly = true;
            this.daDGV.Size = new System.Drawing.Size(347, 191);
            this.daDGV.TabIndex = 21;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(510, 238);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(194, 13);
            this.label8.TabIndex = 22;
            this.label8.Text = "Błędy średnie szacunku parametru D(a)";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(548, 184);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(168, 13);
            this.label6.TabIndex = 23;
            this.label6.Text = "Współczynnik zmienności losowej";
            // 
            // wspLos
            // 
            this.wspLos.Location = new System.Drawing.Point(545, 200);
            this.wspLos.Name = "wspLos";
            this.wspLos.ReadOnly = true;
            this.wspLos.Size = new System.Drawing.Size(182, 20);
            this.wspLos.TabIndex = 24;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.HeaderText = "D(a1)";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.HeaderText = "D(a2)";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.HeaderText = "D(a3)";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            this.dataGridViewTextBoxColumn4.ReadOnly = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(923, 451);
            this.Controls.Add(this.wspLos);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.daDGV);
            this.Controls.Add(this.dataGridView5);
            this.Controls.Add(this.odchylenie);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.r2text);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.rowMod);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.rStar);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dataGridView4);
            this.Controls.Add(this.dataGridView3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.getExelX);
            this.Controls.Add(this.getExelY);
            this.Controls.Add(this.dataGridView2);
            this.Controls.Add(this.dataGridView1);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.daDGV)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.Button getExelY;
        private System.Windows.Forms.Button getExelX;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView dataGridView3;
        private System.Windows.Forms.DataGridView dataGridView4;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.DataGridViewTextBoxColumn X1;
        private System.Windows.Forms.DataGridViewTextBoxColumn X2;
        private System.Windows.Forms.DataGridViewTextBoxColumn X3;
        private System.Windows.Forms.DataGridViewTextBoxColumn X4;
        private System.Windows.Forms.TextBox rStar;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Y1;
        private System.Windows.Forms.TextBox rowMod;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox r2text;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox odchylenie;
        private System.Windows.Forms.DataGridView dataGridView5;
        private System.Windows.Forms.DataGridView daDGV;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox wspLos;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;


    }
}

