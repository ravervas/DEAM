namespace WindowsFormsApplication8
{
    partial class Form11
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
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series2 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series3 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series4 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series5 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series6 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form11));
            this.chart2 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton2 = new System.Windows.Forms.ToolStripButton();
            this.toolStripComboBox4 = new System.Windows.Forms.ToolStripComboBox();
            this.toolStripDropDownButton1 = new System.Windows.Forms.ToolStripDropDownButton();
            this.toolStripComboBox1 = new System.Windows.Forms.ToolStripComboBox();
            this.toolStripSplitButton1 = new System.Windows.Forms.ToolStripSplitButton();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem3 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem4 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSplitButton2 = new System.Windows.Forms.ToolStripSplitButton();
            this.аппроксимацииToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.распределение1ToolStripMenuItem3 = new System.Windows.Forms.ToolStripMenuItem();
            this.распределение2ToolStripMenuItem3 = new System.Windows.Forms.ToolStripMenuItem();
            this.общаябимодалToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.оКРToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.распределение1ToolStripMenuItem4 = new System.Windows.Forms.ToolStripMenuItem();
            this.распределение2ToolStripMenuItem4 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripButton3 = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton4 = new System.Windows.Forms.ToolStripButton();
            this.panel1 = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.chart2)).BeginInit();
            this.toolStrip1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // chart2
            // 
            this.chart2.BorderlineColor = System.Drawing.Color.Black;
            this.chart2.BorderlineDashStyle = System.Windows.Forms.DataVisualization.Charting.ChartDashStyle.Solid;
            this.chart2.BorderlineWidth = 0;
            chartArea1.AxisX.TitleFont = new System.Drawing.Font("Times New Roman", 48F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            chartArea1.AxisX2.Enabled = System.Windows.Forms.DataVisualization.Charting.AxisEnabled.True;
            chartArea1.AxisX2.LabelStyle.Enabled = false;
            chartArea1.AxisX2.MajorGrid.Enabled = false;
            chartArea1.AxisX2.MajorTickMark.Enabled = false;
            chartArea1.AxisY2.Enabled = System.Windows.Forms.DataVisualization.Charting.AxisEnabled.True;
            chartArea1.AxisY2.LabelStyle.Enabled = false;
            chartArea1.AxisY2.MajorGrid.Enabled = false;
            chartArea1.AxisY2.MajorTickMark.Enabled = false;
            chartArea1.Name = "ChartArea1";
            this.chart2.ChartAreas.Add(chartArea1);
            this.chart2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.chart2.IsSoftShadows = false;
            legend1.Alignment = System.Drawing.StringAlignment.Center;
            legend1.Docking = System.Windows.Forms.DataVisualization.Charting.Docking.Bottom;
            legend1.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            legend1.IsTextAutoFit = false;
            legend1.Name = "Legend1";
            this.chart2.Legends.Add(legend1);
            this.chart2.Location = new System.Drawing.Point(0, 0);
            this.chart2.Name = "chart2";
            this.chart2.Palette = System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.None;
            series1.ChartArea = "ChartArea1";
            series1.Color = System.Drawing.Color.CornflowerBlue;
            series1.Font = new System.Drawing.Font("Times New Roman", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            series1.IsVisibleInLegend = false;
            series1.Legend = "Legend1";
            series1.Name = "Гис";
            series1.SmartLabelStyle.AllowOutsidePlotArea = System.Windows.Forms.DataVisualization.Charting.LabelOutsidePlotAreaStyle.Yes;
            series2.ChartArea = "ChartArea1";
            series2.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            series2.Color = System.Drawing.Color.Red;
            series2.Font = new System.Drawing.Font("Times New Roman", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            series2.IsVisibleInLegend = false;
            series2.Legend = "Legend1";
            series2.MarkerBorderColor = System.Drawing.Color.White;
            series2.Name = "Series2";
            series3.ChartArea = "ChartArea1";
            series3.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            series3.Color = System.Drawing.Color.Orange;
            series3.Font = new System.Drawing.Font("Times New Roman", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            series3.IsVisibleInLegend = false;
            series3.Legend = "Legend1";
            series3.Name = "Series3";
            series4.BorderDashStyle = System.Windows.Forms.DataVisualization.Charting.ChartDashStyle.Dash;
            series4.ChartArea = "ChartArea1";
            series4.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            series4.Color = System.Drawing.Color.LawnGreen;
            series4.IsVisibleInLegend = false;
            series4.Legend = "Legend1";
            series4.Name = "ОКР 1";
            series5.BorderDashStyle = System.Windows.Forms.DataVisualization.Charting.ChartDashStyle.Dash;
            series5.ChartArea = "ChartArea1";
            series5.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            series5.Color = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            series5.IsVisibleInLegend = false;
            series5.Legend = "Legend1";
            series5.Name = "ОКР 2";
            series6.ChartArea = "ChartArea1";
            series6.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
            series6.Color = System.Drawing.Color.Blue;
            series6.Font = new System.Drawing.Font("Times New Roman", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            series6.IsVisibleInLegend = false;
            series6.Legend = "Legend1";
            series6.Name = "Общая (бимодал)";
            this.chart2.Series.Add(series1);
            this.chart2.Series.Add(series2);
            this.chart2.Series.Add(series3);
            this.chart2.Series.Add(series4);
            this.chart2.Series.Add(series5);
            this.chart2.Series.Add(series6);
            this.chart2.Size = new System.Drawing.Size(780, 528);
            this.chart2.TabIndex = 2;
            this.chart2.Text = "chart2";
            this.chart2.TextAntiAliasingQuality = System.Windows.Forms.DataVisualization.Charting.TextAntiAliasingQuality.Normal;
            // 
            // toolStrip1
            // 
            this.toolStrip1.BackColor = System.Drawing.SystemColors.Control;
            this.toolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripButton1,
            this.toolStripButton2,
            this.toolStripComboBox4,
            this.toolStripDropDownButton1,
            this.toolStripSplitButton1,
            this.toolStripSplitButton2,
            this.toolStripButton3,
            this.toolStripButton4});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(784, 25);
            this.toolStrip1.Stretch = true;
            this.toolStrip1.TabIndex = 60;
            this.toolStrip1.Text = "toolStrip1";
            this.toolStrip1.MouseEnter += new System.EventHandler(this.toolStrip1_MouseEnter);
            this.toolStrip1.MouseLeave += new System.EventHandler(this.toolStrip1_MouseLeave);
            // 
            // toolStripButton1
            // 
            this.toolStripButton1.BackColor = System.Drawing.SystemColors.Control;
            this.toolStripButton1.Checked = true;
            this.toolStripButton1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.toolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton1.Image = global::WindowsFormsApplication8.Properties.Resources.сетка;
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(23, 22);
            this.toolStripButton1.Text = "toolStripButton1";
            this.toolStripButton1.Click += new System.EventHandler(this.toolStripButton1_Click);
            this.toolStripButton1.MouseMove += new System.Windows.Forms.MouseEventHandler(this.toolStripButton1_MouseMove);
            // 
            // toolStripButton2
            // 
            this.toolStripButton2.BackColor = System.Drawing.SystemColors.Control;
            this.toolStripButton2.Checked = true;
            this.toolStripButton2.CheckState = System.Windows.Forms.CheckState.Checked;
            this.toolStripButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton2.Image = global::WindowsFormsApplication8.Properties.Resources.Легенда;
            this.toolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton2.Name = "toolStripButton2";
            this.toolStripButton2.Size = new System.Drawing.Size(23, 22);
            this.toolStripButton2.Text = "toolStripButton2";
            this.toolStripButton2.Click += new System.EventHandler(this.toolStripButton2_Click);
            this.toolStripButton2.MouseMove += new System.Windows.Forms.MouseEventHandler(this.toolStripButton2_MouseMove);
            // 
            // toolStripComboBox4
            // 
            this.toolStripComboBox4.AutoSize = false;
            this.toolStripComboBox4.Font = new System.Drawing.Font("Times New Roman", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.toolStripComboBox4.Items.AddRange(new object[] {
            "8",
            "9",
            "10",
            "11",
            "12",
            "14",
            "16",
            "18",
            "20",
            "22",
            "24",
            "26",
            "28",
            "36",
            "48",
            "72"});
            this.toolStripComboBox4.Name = "toolStripComboBox4";
            this.toolStripComboBox4.Size = new System.Drawing.Size(40, 25);
            this.toolStripComboBox4.Text = "12";
            this.toolStripComboBox4.SelectedIndexChanged += new System.EventHandler(this.toolStripComboBox4_SelectedIndexChanged);
            this.toolStripComboBox4.MouseMove += new System.Windows.Forms.MouseEventHandler(this.toolStripComboBox4_MouseMove);
            // 
            // toolStripDropDownButton1
            // 
            this.toolStripDropDownButton1.BackColor = System.Drawing.SystemColors.Control;
            this.toolStripDropDownButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripDropDownButton1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripComboBox1});
            this.toolStripDropDownButton1.Image = global::WindowsFormsApplication8.Properties.Resources.Толщина_линии;
            this.toolStripDropDownButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripDropDownButton1.Name = "toolStripDropDownButton1";
            this.toolStripDropDownButton1.Size = new System.Drawing.Size(29, 22);
            this.toolStripDropDownButton1.Text = "toolStripDropDownButton1";
            this.toolStripDropDownButton1.MouseMove += new System.Windows.Forms.MouseEventHandler(this.toolStripDropDownButton1_MouseMove);
            // 
            // toolStripComboBox1
            // 
            this.toolStripComboBox1.AutoSize = false;
            this.toolStripComboBox1.Font = new System.Drawing.Font("Times New Roman", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.toolStripComboBox1.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5"});
            this.toolStripComboBox1.Name = "toolStripComboBox1";
            this.toolStripComboBox1.Size = new System.Drawing.Size(40, 25);
            this.toolStripComboBox1.Text = "1";
            this.toolStripComboBox1.SelectedIndexChanged += new System.EventHandler(this.toolStripComboBox1_SelectedIndexChanged);
            // 
            // toolStripSplitButton1
            // 
            this.toolStripSplitButton1.BackColor = System.Drawing.SystemColors.Menu;
            this.toolStripSplitButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripSplitButton1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem2,
            this.toolStripMenuItem3,
            this.toolStripMenuItem4});
            this.toolStripSplitButton1.Font = new System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.toolStripSplitButton1.Image = global::WindowsFormsApplication8.Properties.Resources.поворот_оси_Х;
            this.toolStripSplitButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripSplitButton1.Name = "toolStripSplitButton1";
            this.toolStripSplitButton1.Size = new System.Drawing.Size(32, 22);
            this.toolStripSplitButton1.Text = "toolStripSplitButton1";
            this.toolStripSplitButton1.MouseMove += new System.Windows.Forms.MouseEventHandler(this.toolStripSplitButton1_MouseMove);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.AutoSize = false;
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(60, 22);
            this.toolStripMenuItem2.Text = "0°";
            this.toolStripMenuItem2.Click += new System.EventHandler(this.toolStripMenuItem2_Click);
            // 
            // toolStripMenuItem3
            // 
            this.toolStripMenuItem3.AutoSize = false;
            this.toolStripMenuItem3.Name = "toolStripMenuItem3";
            this.toolStripMenuItem3.Size = new System.Drawing.Size(60, 22);
            this.toolStripMenuItem3.Text = "45°";
            this.toolStripMenuItem3.Click += new System.EventHandler(this.toolStripMenuItem3_Click);
            // 
            // toolStripMenuItem4
            // 
            this.toolStripMenuItem4.AutoSize = false;
            this.toolStripMenuItem4.Name = "toolStripMenuItem4";
            this.toolStripMenuItem4.Size = new System.Drawing.Size(60, 22);
            this.toolStripMenuItem4.Text = "90°";
            this.toolStripMenuItem4.Click += new System.EventHandler(this.toolStripMenuItem4_Click);
            // 
            // toolStripSplitButton2
            // 
            this.toolStripSplitButton2.BackColor = System.Drawing.SystemColors.Menu;
            this.toolStripSplitButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripSplitButton2.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.аппроксимацииToolStripMenuItem,
            this.оКРToolStripMenuItem});
            this.toolStripSplitButton2.Font = new System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.toolStripSplitButton2.Image = ((System.Drawing.Image)(resources.GetObject("toolStripSplitButton2.Image")));
            this.toolStripSplitButton2.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripSplitButton2.Name = "toolStripSplitButton2";
            this.toolStripSplitButton2.Size = new System.Drawing.Size(32, 22);
            this.toolStripSplitButton2.Text = "toolStripSplitButton2";
            this.toolStripSplitButton2.MouseMove += new System.Windows.Forms.MouseEventHandler(this.toolStripSplitButton2_MouseMove);
            // 
            // аппроксимацииToolStripMenuItem
            // 
            this.аппроксимацииToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.распределение1ToolStripMenuItem3,
            this.распределение2ToolStripMenuItem3,
            this.общаябимодалToolStripMenuItem1});
            this.аппроксимацииToolStripMenuItem.Name = "аппроксимацииToolStripMenuItem";
            this.аппроксимацииToolStripMenuItem.Size = new System.Drawing.Size(167, 22);
            this.аппроксимацииToolStripMenuItem.Text = "Аппроксимации";
            // 
            // распределение1ToolStripMenuItem3
            // 
            this.распределение1ToolStripMenuItem3.Name = "распределение1ToolStripMenuItem3";
            this.распределение1ToolStripMenuItem3.Size = new System.Drawing.Size(171, 22);
            this.распределение1ToolStripMenuItem3.Text = "Распределение 1";
            this.распределение1ToolStripMenuItem3.Click += new System.EventHandler(this.распределение1ToolStripMenuItem3_Click);
            // 
            // распределение2ToolStripMenuItem3
            // 
            this.распределение2ToolStripMenuItem3.Name = "распределение2ToolStripMenuItem3";
            this.распределение2ToolStripMenuItem3.Size = new System.Drawing.Size(171, 22);
            this.распределение2ToolStripMenuItem3.Text = "Распределение 2";
            this.распределение2ToolStripMenuItem3.Click += new System.EventHandler(this.распределение2ToolStripMenuItem3_Click);
            // 
            // общаябимодалToolStripMenuItem1
            // 
            this.общаябимодалToolStripMenuItem1.Name = "общаябимодалToolStripMenuItem1";
            this.общаябимодалToolStripMenuItem1.Size = new System.Drawing.Size(171, 22);
            this.общаябимодалToolStripMenuItem1.Text = "Общая (бимодал)";
            this.общаябимодалToolStripMenuItem1.Click += new System.EventHandler(this.общаябимодалToolStripMenuItem1_Click);
            // 
            // оКРToolStripMenuItem
            // 
            this.оКРToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.распределение1ToolStripMenuItem4,
            this.распределение2ToolStripMenuItem4});
            this.оКРToolStripMenuItem.Name = "оКРToolStripMenuItem";
            this.оКРToolStripMenuItem.Size = new System.Drawing.Size(167, 22);
            this.оКРToolStripMenuItem.Text = "ОКР";
            // 
            // распределение1ToolStripMenuItem4
            // 
            this.распределение1ToolStripMenuItem4.Name = "распределение1ToolStripMenuItem4";
            this.распределение1ToolStripMenuItem4.Size = new System.Drawing.Size(166, 22);
            this.распределение1ToolStripMenuItem4.Text = "Распределение 1";
            this.распределение1ToolStripMenuItem4.Click += new System.EventHandler(this.распределение1ToolStripMenuItem4_Click);
            // 
            // распределение2ToolStripMenuItem4
            // 
            this.распределение2ToolStripMenuItem4.Name = "распределение2ToolStripMenuItem4";
            this.распределение2ToolStripMenuItem4.Size = new System.Drawing.Size(166, 22);
            this.распределение2ToolStripMenuItem4.Text = "Распределение 2";
            this.распределение2ToolStripMenuItem4.Click += new System.EventHandler(this.распределение2ToolStripMenuItem4_Click);
            // 
            // toolStripButton3
            // 
            this.toolStripButton3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton3.Image = global::WindowsFormsApplication8.Properties.Resources.Среднее_значение;
            this.toolStripButton3.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton3.Name = "toolStripButton3";
            this.toolStripButton3.Size = new System.Drawing.Size(23, 22);
            this.toolStripButton3.Text = "toolStripButton3";
            this.toolStripButton3.Click += new System.EventHandler(this.toolStripButton3_Click);
            this.toolStripButton3.MouseMove += new System.Windows.Forms.MouseEventHandler(this.toolStripButton3_MouseMove);
            // 
            // toolStripButton4
            // 
            this.toolStripButton4.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButton4.Font = new System.Drawing.Font("Times New Roman", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.toolStripButton4.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton4.Name = "toolStripButton4";
            this.toolStripButton4.Size = new System.Drawing.Size(28, 22);
            this.toolStripButton4.Text = "Ln";
            this.toolStripButton4.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.toolStripButton4.TextImageRelation = System.Windows.Forms.TextImageRelation.TextAboveImage;
            this.toolStripButton4.Click += new System.EventHandler(this.toolStripButton4_Click);
            this.toolStripButton4.MouseMove += new System.Windows.Forms.MouseEventHandler(this.toolStripButton4_MouseMove);
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.chart2);
            this.panel1.Location = new System.Drawing.Point(0, 28);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(784, 532);
            this.panel1.TabIndex = 61;
            // 
            // Form11
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 561);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.toolStrip1);
            this.MaximizeBox = false;
            this.Name = "Form11";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Лин. распр.";
            this.Activated += new System.EventHandler(this.Form11_Activated);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form11_FormClosing);
            this.Load += new System.EventHandler(this.Form11_Load);
            this.Click += new System.EventHandler(this.Form11_Click);
            ((System.ComponentModel.ISupportInitialize)(this.chart2)).EndInit();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.DataVisualization.Charting.Chart chart2;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton toolStripButton1;
        private System.Windows.Forms.ToolStripButton toolStripButton2;
        private System.Windows.Forms.ToolStripComboBox toolStripComboBox4;
        private System.Windows.Forms.ToolStripDropDownButton toolStripDropDownButton1;
        private System.Windows.Forms.ToolStripComboBox toolStripComboBox1;
        private System.Windows.Forms.ToolStripSplitButton toolStripSplitButton1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem3;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem4;
        public System.Windows.Forms.ToolStripSplitButton toolStripSplitButton2;
        public System.Windows.Forms.ToolStripMenuItem аппроксимацииToolStripMenuItem;
        public System.Windows.Forms.ToolStripMenuItem распределение1ToolStripMenuItem3;
        public System.Windows.Forms.ToolStripMenuItem распределение2ToolStripMenuItem3;
        public System.Windows.Forms.ToolStripMenuItem общаябимодалToolStripMenuItem1;
        public System.Windows.Forms.ToolStripMenuItem оКРToolStripMenuItem;
        public System.Windows.Forms.ToolStripMenuItem распределение1ToolStripMenuItem4;
        public System.Windows.Forms.ToolStripMenuItem распределение2ToolStripMenuItem4;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ToolStripButton toolStripButton3;
        public System.Windows.Forms.ToolStripButton toolStripButton4;
    }
}