using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication8
{
    public partial class Form10 : Form
    {

        int y, clickML = 0, Loc, x_position1, x_position2; // click_toolStr_button1=0;
        double x;
        Point lastpoint = Point.Empty;
        private bool MD = false, Line = false, sizeWE0 = false, sizeWE1 = false, Discr=true; 
        private double x_mouseMove, y_mouseMove, Di;
        public string LineChart = "Линейная диаграмма", ColumnChart = "Столбчатая диаграмма", Interval = "Измерение интервала", Smooth="Сглаживание";


        private void toolStripButton1_MouseMove(object sender, MouseEventArgs e)
        {
            toolStripButton1.ToolTipText = Program.f1.Grid;
        }

        private void toolStripButton2_MouseMove(object sender, MouseEventArgs e)
        {
            toolStripButton2.ToolTipText = LineChart;
        }

        private void toolStripComboBox1_MouseMove(object sender, MouseEventArgs e)
        {
            toolStripComboBox1.ToolTipText = Program.f1.font;
        }

        private void toolStripButton3_MouseMove(object sender, MouseEventArgs e)
        {
            toolStripButton3.ToolTipText = ColumnChart;
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            toolStripButton3.Checked = false;
            toolStripButton2.Checked = true;
            chart1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            chart1.Series[0].Color = Color.Navy;
            Program.f1.tSBut2 = true;
            Program.f1.tSBut3 = false;
        }

        

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            toolStripMenuItem2.Image = Properties.Resources.radio_button_on;
            toolStripMenuItem3.Image = Properties.Resources.radio_button_off;
            Discr = true;
        }

        private void toolStripButton4_MouseMove(object sender, MouseEventArgs e)
        {
            toolStripButton4.ToolTipText = Smooth;
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            // сглаживание по  трем точкам
            Program.f1.dataGridView4.Rows[0].Cells[1].Value = (5 * Convert.ToDouble(Program.f1.dataGridView4.Rows[0].Cells[2].Value) + 2 * Convert.ToDouble(Program.f1.dataGridView4.Rows[1].Cells[2].Value) - Convert.ToDouble(Program.f1.dataGridView4.Rows[2].Cells[2].Value)) / 6;
            for (int i=1;i<=Program.f1.col_pix-1;i++)
            {
                Program.f1.dataGridView4.Rows[i].Cells[1].Value = (Convert.ToDouble(Program.f1.dataGridView4.Rows[i - 1].Cells[2].Value) + Convert.ToDouble(Program.f1.dataGridView4.Rows[i].Cells[2].Value) + Convert.ToDouble(Program.f1.dataGridView4.Rows[i + 1].Cells[2].Value)) / 3;
                
            }
            Program.f1.dataGridView4.Rows[Program.f1.col_pix].Cells[1].Value=(-1* Convert.ToDouble(Program.f1.dataGridView4.Rows[Program.f1.col_pix - 2].Cells[2].Value)+ 2*Convert.ToDouble(Program.f1.dataGridView4.Rows[Program.f1.col_pix - 1].Cells[2].Value)+5* Convert.ToDouble(Program.f1.dataGridView4.Rows[Program.f1.col_pix].Cells[2].Value)) / 6;

            for (int i = 0; i <= Program.f1.col_pix; i++)
                Program.f1.dataGridView4.Rows[i].Cells[2].Value = Program.f1.dataGridView4.Rows[i].Cells[1].Value;

            Program.f1.Y_min = Convert.ToInt32(Program.f1.dataGridView4.Rows[0].Cells[2].Value);
            Program.f1.Y_max = Convert.ToInt32(Program.f1.dataGridView4.Rows[0].Cells[2].Value);
            for (int i = 1; i <= Program.f1.col_pix; i++)
            {
                if (Convert.ToInt32(Program.f1.dataGridView4.Rows[i].Cells[2].Value) < Program.f1.Y_min)
                    Program.f1.Y_min = Convert.ToInt32(Program.f1.dataGridView4.Rows[i].Cells[2].Value);
                if (Convert.ToInt32(Program.f1.dataGridView4.Rows[i].Cells[2].Value) > Program.f1.Y_max)
                    Program.f1.Y_max = Convert.ToInt32(Program.f1.dataGridView4.Rows[i].Cells[2].Value);
            }
            chart1.Series[0].Points.Clear();           
            for (int i = 0; i <= Program.f1.col_pix; i++)
            {
                x = Convert.ToDouble(Program.f1.dataGridView4.Rows[i].Cells[0].Value);
                y = Convert.ToInt32(Program.f1.dataGridView4.Rows[i].Cells[2].Value);
                chart1.Series[0].Points.AddXY(x, y);
            }
            if (Program.f1.Y_min - 10 <= 0)
                Program.f1.Y_min = 0;
            else
                Program.f1.Y_min = Program.f1.Y_min - 10;
            Program.f1.Y_max = Program.f1.Y_max + 5;
            chart1.ChartAreas[0].AxisY.Maximum = Program.f1.Y_max;
            chart1.ChartAreas[0].AxisY.Minimum = Program.f1.Y_min;
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            toolStripMenuItem2.Image = Properties.Resources.radio_button_off;
            toolStripMenuItem3.Image = Properties.Resources.radio_button_on;
            Discr = false;
        }

        private void toolStripSplitButton1_MouseMove(object sender, MouseEventArgs e)
        {
            toolStripSplitButton1.ToolTipText = Interval;
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            toolStripButton3.Checked = true;
            toolStripButton2.Checked = false;
            chart1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
            chart1.Series[0]["PointWidth"] = "1";
            chart1.Series[0].Color = Color.MediumTurquoise;
            Program.f1.tSBut3 = true;
            Program.f1.tSBut2 = false;
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Program.f1.I_size = Convert.ToInt32(toolStripComboBox1.Text);            
            chart1.ChartAreas[0].AxisX.TitleFont = new Font("Times New Roman", Program.f1.I_size, FontStyle.Regular);
            chart1.ChartAreas[0].AxisX.LabelStyle.Font = new Font("Times New Roman", Program.f1.I_size, FontStyle.Regular);
            chart1.ChartAreas[0].AxisY.LabelStyle.Font = new Font("Times New Roman", Program.f1.I_size, FontStyle.Regular);
            label1.Font = new Font("Times New Roman", Program.f1.I_size, FontStyle.Regular);
            if (Program.f1.I_size == 12)
                label1.Location = new Point(label1.Location.X, 338);
            if (Program.f1.I_size == 14)
                label1.Location = new Point(label1.Location.X, 328);
            if (Program.f1.I_size == 16)
                label1.Location = new Point(label1.Location.X, 318);
            if (Program.f1.I_size == 18)
                label1.Location = new Point(label1.Location.X, 308);
            if (Program.f1.I_size == 20)
                label1.Location = new Point(label1.Location.X, 298);
            if (Program.f1.I_size == 22)
                label1.Location = new Point(label1.Location.X, 288);
            if (Program.f1.I_size == 24)
                label1.Location = new Point(label1.Location.X, 278);
            if (Program.f1.I_size == 26)
                label1.Location = new Point(label1.Location.X, 268);

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            ++Program.f1.click_toolStr_button1;
            switch (Program.f1.click_toolStr_button1)
            {
                case 1:                   
                    chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
                    chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
                    break;
                case 2:                  
                    chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = true;
                    chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true;
                    Program.f1.click_toolStr_button1 = 0;
                    break;
            }
        }

        private void chart1_MouseUp(object sender, MouseEventArgs e)
        {
            sizeWE0 = false;
            sizeWE1 = false;
        }

      
        public Form10()
        {
            Program.f10 = this;
            InitializeComponent();
        }

        private void Form10_Load(object sender, EventArgs e)
        {
            Program.f1.FormLineProphile = true;
            label1.Parent = chart1;
            label1.BackColor = Color.Transparent;
            if (Program.f1.Eng)
            {
                this.Text = "Line profile";
                chart1.ChartAreas[0].AxisX.Title = Program.f1.textBox5.Text;
                LineChart = "Line chart";
                ColumnChart = "Column chart";
                Interval = "Interval measurement";
                toolStripMenuItem2.Text = "Discrete";
                toolStripMenuItem3.Text = "Continuous";
                Smooth = "Smoothing";
            }
            LineProphile();

            if (Program.f1.tSBut2)            
                toolStripButton2_Click(sender, e);
            else
                toolStripButton3_Click(sender, e);

            if (Program.f1.click_toolStr_button1==1)
            {
                chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
                chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
            }

            chart1.ChartAreas[0].AxisX.TitleFont = new Font("Times New Roman", Program.f1.I_size, FontStyle.Regular);
            chart1.ChartAreas[0].AxisX.LabelStyle.Font = new Font("Times New Roman", Program.f1.I_size, FontStyle.Regular);
            chart1.ChartAreas[0].AxisY.LabelStyle.Font = new Font("Times New Roman", Program.f1.I_size, FontStyle.Regular);
            label1.Font = new Font("Times New Roman", Program.f1.I_size, FontStyle.Regular);
            if (Program.f1.I_size == 12)
                label1.Location = new Point(label1.Location.X, 338);
            if (Program.f1.I_size == 14)
                label1.Location = new Point(label1.Location.X, 328);
            if (Program.f1.I_size == 16)
                label1.Location = new Point(label1.Location.X, 318);
            if (Program.f1.I_size == 18)
                label1.Location = new Point(label1.Location.X, 308);
            if (Program.f1.I_size == 20)
                label1.Location = new Point(label1.Location.X, 298);
            if (Program.f1.I_size == 22)
                label1.Location = new Point(label1.Location.X, 288);
            if (Program.f1.I_size == 24)
                label1.Location = new Point(label1.Location.X, 278);
            if (Program.f1.I_size == 26)
                label1.Location = new Point(label1.Location.X, 268);
            toolStripComboBox1.Text = Program.f1.I_size.ToString();

        }

        public void LineProphile()
        {
            for (int i=0;i<6;i++)
            chart1.Series[i].Points.Clear();
            label1.Visible = false;
            for (int i = 0; i <= Program.f1.col_pix; i++)
            {
                x = Convert.ToDouble(Program.f1.dataGridView4.Rows[i].Cells[0].Value);
                y = Convert.ToInt32(Program.f1.dataGridView4.Rows[i].Cells[2].Value);
                chart1.Series[0].Points.AddXY(x, y);
            }
            chart1.ChartAreas[0].AxisX.Minimum = 0;
            chart1.ChartAreas[0].AxisX.Title = Program.f1.textBox5.Text;
            chart1.ChartAreas[0].AxisX.TitleFont = new Font("Times New Roman", 12, FontStyle.Regular);       
            if (Program.f1.Y_min - 10 <= 0)
                Program.f1.Y_min = 0;
            else
                Program.f1.Y_min = Program.f1.Y_min - 10;
            Program.f1.Y_max = Program.f1.Y_max + 5;
            chart1.ChartAreas[0].AxisY.Maximum = Program.f1.Y_max;
            chart1.ChartAreas[0].AxisY.Minimum = Program.f1.Y_min;
            chart1.ChartAreas[0].AxisX.Maximum = Convert.ToDouble(Program.f1.dataGridView4.Rows[Program.f1.col_pix].Cells[0].Value);
          
            
        }

        private void chart1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button==MouseButtons.Left)
            {
                lastpoint = e.Location;
                x_mouseMove = Math.Round(chart1.ChartAreas[0].AxisX.PixelPositionToValue(lastpoint.X), 2);
                y_mouseMove = Math.Round(chart1.ChartAreas[0].AxisY.PixelPositionToValue(lastpoint.Y), 2);

                if (Discr)
                {
                    for (int i = 0; i < Program.f1.col_pix; i++)
                        if (Math.Abs(x_mouseMove - Convert.ToDouble(Program.f1.dataGridView4.Rows[i].Cells[0].Value)) < Math.Abs(x_mouseMove - Convert.ToDouble(Program.f1.dataGridView4.Rows[i + 1].Cells[0].Value)))
                        {
                            x_mouseMove = Convert.ToDouble(Program.f1.dataGridView4.Rows[i].Cells[0].Value);
                            break;
                        }
                }
               

                if (chart1.Cursor==Cursors.VSplit)
                {
                    ++clickML;
                    switch (clickML)
                    {
                        case 1:

                            {
                                Line = false;
                                label1.Visible = true;
                                chart1.Series[1].Points.Clear();
                                chart1.Series[2].Points.Clear();
                                chart1.Series[3].Points.Clear();
                                chart1.Series[4].Points.Clear();
                                chart1.Series[5].Points.Clear();
                                MD = true;
                                dataGridView1.Rows.Add();
                                dataGridView1.Rows[0].Cells[0].Value = x_mouseMove;
                                chart1.Series[1].Points.AddXY(x_mouseMove, 0);
                                chart1.Series[1].Points.AddXY(x_mouseMove, Program.f1.Y_max);                               
                                x_position1 = lastpoint.X;
                            }
                            break;

                        case 2:
                            {
                                
                                clickML = 0;
                                MD = false;
                                Line = true;
                            }
                            break;
                    }
                   
                }

                if (chart1.Cursor == Cursors.SizeWE)
                {                    
                   if(x_mouseMove > Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value) - 1 * Program.f1.n / Program.f1.n_const && x_mouseMove < Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value) + 1 * Program.f1.n / Program.f1.n_const)
                    sizeWE0 = true;
                else
                    sizeWE1 = true;
                }


                
            }
        }

        private void chart1_MouseMove(object sender, MouseEventArgs e)
        {
            try
            {
                lastpoint = e.Location;
                x_mouseMove = Math.Round(chart1.ChartAreas[0].AxisX.PixelPositionToValue(lastpoint.X), 2);
                y_mouseMove = Math.Round(chart1.ChartAreas[0].AxisY.PixelPositionToValue(lastpoint.Y), 2);

                if (Discr)
                {
                    for (int i = 0; i < Program.f1.col_pix; i++)
                        if (Math.Abs(x_mouseMove - Convert.ToDouble(Program.f1.dataGridView4.Rows[i].Cells[0].Value)) < Math.Abs(x_mouseMove - Convert.ToDouble(Program.f1.dataGridView4.Rows[i + 1].Cells[0].Value)))
                        {
                            x_mouseMove = Convert.ToDouble(Program.f1.dataGridView4.Rows[i].Cells[0].Value);
                            break;
                        }
                }
               

                if (sizeWE0 == false && sizeWE1 == false)
                {
                    if (x_mouseMove > 0 && y_mouseMove > Program.f1.Y_min && x_mouseMove< chart1.ChartAreas[0].AxisX.Maximum && y_mouseMove< Program.f1.Y_max)
                    {
                        if (Line)
                        {
                            if (x_mouseMove > Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value) - 1 * Program.f1.n / Program.f1.n_const && x_mouseMove < Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value) + 1 * Program.f1.n / Program.f1.n_const || x_mouseMove > Convert.ToDouble(dataGridView1.Rows[1].Cells[0].Value) - 1 * Program.f1.n / Program.f1.n_const && x_mouseMove < Convert.ToDouble(dataGridView1.Rows[1].Cells[0].Value) + 1 * Program.f1.n / Program.f1.n_const)
                                chart1.Cursor = Cursors.SizeWE;
                            else
                                chart1.Cursor = Cursors.VSplit;
                        }
                        else
                            chart1.Cursor = Cursors.VSplit;
                    }
                    else
                        chart1.Cursor = Cursors.Default;

                    
                }
                else
                    chart1.Cursor = Cursors.SizeWE;


            }
            catch
            {
                return;
            }
            

            if (MD)
            {               

                if (x_mouseMove > 0 && y_mouseMove > Program.f1.Y_min)
                {
                    chart1.Series[2].Points.Clear();
                    chart1.Series[2].Points.AddXY(x_mouseMove, 0);
                    chart1.Series[2].Points.AddXY(x_mouseMove, Program.f1.Y_max);
                    chart1.Series[3].Points.Clear();
                    chart1.Series[3].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value), Program.f1.Y_min + 10);
                    chart1.Series[3].Points.AddXY(x_mouseMove, Program.f1.Y_min + 10);
                    x_position2 = lastpoint.X;
                    if (Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value) < x_mouseMove)
                    {
                        chart1.Series[4].Points.Clear();
                        chart1.Series[4].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value) + 3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 5);
                        chart1.Series[4].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value), Program.f1.Y_min+10);
                        chart1.Series[4].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value) + 3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 15);
                        chart1.Series[5].Points.Clear();
                        chart1.Series[5].Points.AddXY(x_mouseMove-3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 5);
                        chart1.Series[5].Points.AddXY(x_mouseMove, Program.f1.Y_min + 10);
                        chart1.Series[5].Points.AddXY(x_mouseMove - 3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 15);

                        if (Math.Abs(x_position1 - x_position2) < label1.Size.Width)
                            label1.Location = new Point(x_position2 + 3, label1.Location.Y);
                        else
                        {
                            Loc = x_position1 + (x_position2 - x_position1) / 2 - label1.Size.Width/2;
                            label1.Location = new Point(Loc+5, label1.Location.Y);
                        }
                    }
                    else
                    {
                        chart1.Series[4].Points.Clear();
                        chart1.Series[4].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value) - 3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 5);
                        chart1.Series[4].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value), Program.f1.Y_min + 10);
                        chart1.Series[4].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value) - 3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 15);
                        chart1.Series[5].Points.Clear();
                        chart1.Series[5].Points.AddXY(x_mouseMove + 3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 5);
                        chart1.Series[5].Points.AddXY(x_mouseMove, Program.f1.Y_min + 10);
                        chart1.Series[5].Points.AddXY(x_mouseMove + 3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 15);    
                        
                        if (Math.Abs(x_position1 - x_position2) < label1.Size.Width)
                            label1.Location = new Point(x_position1 - label1.Size.Width - 3, label1.Location.Y);
                        else
                        {
                            Loc = x_position1 - (x_position1 - x_position2) / 2 - label1.Size.Width/2;
                            label1.Location = new Point(Loc+5, label1.Location.Y);
                        }
                    }
                    dataGridView1.Rows[1].Cells[0].Value = x_mouseMove;
                    dataGridView1.Rows[1].Cells[1].Value = y_mouseMove;                  
                    Di = Math.Round(Math.Abs(Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[1].Cells[0].Value)), 3);
                                        
                    label1.Text = Di.ToString();
                    
                }
            }

            if (sizeWE0)
            {
               
                chart1.Series[1].Points.Clear();
                    chart1.Series[1].Points.AddXY(x_mouseMove, 0);
                    chart1.Series[1].Points.AddXY(x_mouseMove, Program.f1.Y_max);
                    dataGridView1.Rows[0].Cells[0].Value = x_mouseMove;
                    chart1.Series[3].Points.Clear();                    
                    chart1.Series[3].Points.AddXY(x_mouseMove, Program.f1.Y_min + 10);
                    chart1.Series[3].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[1].Cells[0].Value), Program.f1.Y_min + 10);
                    Di = Math.Round(Math.Abs(Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[1].Cells[0].Value)), 3);
                    
                    label1.Text = Di.ToString();
                    x_position1 = lastpoint.X;
                    if (Convert.ToDouble(dataGridView1.Rows[1].Cells[0].Value) > x_mouseMove)
                    {
                        chart1.Series[4].Points.Clear();
                        chart1.Series[4].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value) + 3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 5);
                        chart1.Series[4].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value), Program.f1.Y_min + 10);
                        chart1.Series[4].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value) + 3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 15);
                        chart1.Series[5].Points.Clear();
                        chart1.Series[5].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[1].Cells[0].Value) - 3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 5);
                        chart1.Series[5].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[1].Cells[0].Value), Program.f1.Y_min + 10);
                        chart1.Series[5].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[1].Cells[0].Value) - 3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 15);
                        if (Math.Abs(x_position1 - x_position2) < label1.Size.Width)
                            label1.Location = new Point(x_position2 + 3, label1.Location.Y);
                        else
                        {
                            Loc = x_position1 + (x_position2 - x_position1) / 2 - label1.Size.Width/2;
                            label1.Location = new Point(Loc+5, label1.Location.Y);
                        }
                    }
                    else
                    {
                        chart1.Series[4].Points.Clear();
                        chart1.Series[4].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value) - 3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 5);
                        chart1.Series[4].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value), Program.f1.Y_min + 10);
                        chart1.Series[4].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value) - 3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 15);
                        chart1.Series[5].Points.Clear();
                        chart1.Series[5].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[1].Cells[0].Value) + 3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 5);
                        chart1.Series[5].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[1].Cells[0].Value), Program.f1.Y_min + 10);
                        chart1.Series[5].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[1].Cells[0].Value) + 3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 15);
                        if (Math.Abs(x_position1 - x_position2) < label1.Size.Width)
                            label1.Location = new Point(x_position1 - label1.Size.Width - 3, label1.Location.Y);
                        else
                        {
                            Loc = x_position1 - (x_position1 - x_position2) / 2 - label1.Size.Width/2;
                            label1.Location = new Point(Loc+5, label1.Location.Y);
                        }
                    }

                
            }

            if (sizeWE1)
            {

                chart1.Series[2].Points.Clear();
                chart1.Series[2].Points.AddXY(x_mouseMove, 0);
                chart1.Series[2].Points.AddXY(x_mouseMove, Program.f1.Y_max);
                dataGridView1.Rows[1].Cells[0].Value = x_mouseMove;
                chart1.Series[3].Points.Clear();
                chart1.Series[3].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value), Program.f1.Y_min + 10);
                chart1.Series[3].Points.AddXY(x_mouseMove, Program.f1.Y_min + 10);
                Di = Math.Round(Math.Abs(Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[1].Cells[0].Value)), 3);
                
                label1.Text = Di.ToString();
                x_position2 = lastpoint.X;
                if (Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value) < x_mouseMove)
                {
                    chart1.Series[4].Points.Clear();
                    chart1.Series[4].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value) + 3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 5);
                    chart1.Series[4].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value), Program.f1.Y_min + 10);
                    chart1.Series[4].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value) + 3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 15);
                    chart1.Series[5].Points.Clear();
                    chart1.Series[5].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[1].Cells[0].Value) - 3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 5);
                    chart1.Series[5].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[1].Cells[0].Value), Program.f1.Y_min + 10);
                    chart1.Series[5].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[1].Cells[0].Value) - 3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 15);
                    if (Math.Abs(x_position1 - x_position2) < label1.Size.Width)
                        label1.Location = new Point(x_position2 + 3, label1.Location.Y);
                    else
                    {
                        Loc = x_position1 + (x_position2 - x_position1) / 2 - label1.Size.Width/2;
                        label1.Location = new Point(Loc+5, label1.Location.Y);
                    }
                }
                else
                {
                    chart1.Series[4].Points.Clear();
                    chart1.Series[4].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value) - 3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 5);
                    chart1.Series[4].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value), Program.f1.Y_min + 10);
                    chart1.Series[4].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value) - 3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 15);
                    chart1.Series[5].Points.Clear();
                    chart1.Series[5].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[1].Cells[0].Value) + 3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 5);
                    chart1.Series[5].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[1].Cells[0].Value), Program.f1.Y_min + 10);
                    chart1.Series[5].Points.AddXY(Convert.ToDouble(dataGridView1.Rows[1].Cells[0].Value) + 3 * Program.f1.n / Program.f1.n_const, Program.f1.Y_min + 15);
                    if (Math.Abs(x_position1 - x_position2) < label1.Size.Width)
                        label1.Location = new Point(x_position1 - label1.Size.Width - 3, label1.Location.Y);
                    else
                    {
                        Loc = x_position1 - (x_position1 - x_position2) / 2 - label1.Size.Width/2;
                        label1.Location = new Point(Loc+5, label1.Location.Y);
                    }
                }


            }

        }

        private void Form10_FormClosing(object sender, FormClosingEventArgs e)
        {
            Program.f1.FormLineProphile = false;
            Line = false;
            Program.f1.chart1.Series[Program.f1.I].Points.Clear();
            Program.f1.chart1.Series.Remove(Program.f1.chart1.Series[Program.f1.I]);
            Program.f1.dataGridView1.Rows.Remove(Program.f1.dataGridView1.Rows[Program.f1.num_line]);
            Program.f1.dataGridView1.Rows.Remove(Program.f1.dataGridView1.Rows[Program.f1.num_line - 1]);
            Program.f1.num_line = Program.f1.num_line - 2;
            Program.f1.I = Program.f1.I - 1;
            Program.f1.kol_count_di = Program.f1.kol_count_di - 1;
            Program.f1.button10.BackColor = System.Drawing.SystemColors.Control;
            Program.f1.clickLineProfile = 0;
            if (Program.f1.loupe)
            {
                Program.f2.chart1.Series[Program.f2.I].Points.Clear();
                Program.f2.chart1.Series.Remove(Program.f2.chart1.Series[Program.f2.I]);
                Program.f2.dataGridView1.Rows.Remove(Program.f2.dataGridView1.Rows[Program.f2.num_line]);
                Program.f2.dataGridView1.Rows.Remove(Program.f2.dataGridView1.Rows[Program.f2.num_line - 1]);
                Program.f2.num_line = Program.f2.num_line - 2;
                Program.f2.I = Program.f2.I - 1;
            }
        }
    }
}
