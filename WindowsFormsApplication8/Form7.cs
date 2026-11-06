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
    public partial class Form7 : Form
    {
        private double x, y, x_mouseMove, y_mouseMove;
        Point lastpoint = Point.Empty;
        private bool But1 = false;
        private int n_row;
        public Form7()
        {
            Program.f7 = this;            
            InitializeComponent();
        }

        private void Form7_Load(object sender, EventArgs e)
        {
            Program.f3.F7 = true;
            ToolTip t = new ToolTip();
            t.SetToolTip(button1, "Изменить положение максимумов");
            for (int i = 0; i < Program.f3.N; i++)
            {
                x = Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[0].Value);
                y = Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[2].Value);
                chart1.Series[0].Points.AddXY(x, y);
                //Program.f3.X_max = Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[0].Value);
            }
            chart1.ChartAreas[0].AxisX.Title = "Размер частиц, нм";
            chart1.ChartAreas[0].AxisX.TitleFont = new Font("Times New Roman", 12, FontStyle.Bold);
            chart1.ChartAreas[0].AxisY.Title = "Доля частиц, %";
            chart1.ChartAreas[0].AxisY.TitleFont = new Font("Times New Roman", 12, FontStyle.Bold);
            chart1.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Times New Roman", 12, System.Drawing.FontStyle.Regular);
            chart1.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Times New Roman", 12, System.Drawing.FontStyle.Regular);
            Program.f3.r= Convert.ToDouble(Program.f3.dataGridView1.Rows[1].Cells[0].Value) - Convert.ToDouble(Program.f3.dataGridView1.Rows[0].Cells[0].Value);
            Program.f3.X_max= Convert.ToDouble(Program.f3.dataGridView1.Rows[Program.f3.N-1].Cells[0].Value)+ Program.f3.r;
            chart1.ChartAreas[0].AxisY.Minimum = 0;
            chart1.ChartAreas[0].AxisX.Maximum = Program.f3.X_max;
            if (Program.f3.toolStripComboBox2.SelectedIndex == 0)
            {
                if (Program.f1.Rus == true)                
                        chart1.ChartAreas[0].AxisX.Title = "Размер частиц, мм";                                    
                if (Program.f1.Eng == true)                                  
                        chart1.ChartAreas[0].AxisX.Title = "Particle size, mm";  
            }
            if (Program.f3.toolStripComboBox2.SelectedIndex == 1)
            {
                if (Program.f1.Rus == true)                                 
                        chart1.ChartAreas[0].AxisX.Title = "Размер частиц, мкм";                      
                if (Program.f1.Eng == true)               
                        chart1.ChartAreas[0].AxisX.Title = "Particle size, mkm";                 
            }
            if (Program.f3.toolStripComboBox2.SelectedIndex == 2)
            {
                if (Program.f1.Rus == true)                                    
                        chart1.ChartAreas[0].AxisX.Title = "Размер частиц, нм";                   
                if (Program.f1.Eng == true)
                        chart1.ChartAreas[0].AxisX.Title = "Particle size, nm";                    
            }
            if (Program.f3.dataGridView1.Rows.Count > 2)
            {
                Program.f3.X_interval = Convert.ToDouble(Program.f3.dataGridView1.Rows[1].Cells[0].Value) - Convert.ToDouble(Program.f3.dataGridView1.Rows[0].Cells[0].Value);
                Program.f3.X_min = Convert.ToDouble(Program.f3.dataGridView1.Rows[0].Cells[0].Value) - Program.f3.X_interval;
                chart1.ChartAreas[0].AxisX.Minimum = Program.f3.X_min;
                chart1.ChartAreas[0].AxisX.Interval = Program.f3.X_interval;
                if (Program.f3.X_min <= 0)
                {
                    chart1.ChartAreas[0].AxisX.IsStartedFromZero = true;
                    for (int i = 0; i < Program.f3.N; i++)
                        chart1.Series[0].Points[i].AxisLabel = (Program.f3.dataGridView1.Rows[i].Cells[0].Value).ToString();
                }
            }
            Program.f3.R2 = 0;
            Program.f3.SSres = 0;
            Program.f3.SStot = 0;
            Program.f3.error_sumi = 0;
            Program.f3.kol_y_max = 0;
            Program.f3.sumni = 0;
            Program.f3.sum_Y = 0;
            Program.f3.S_ni_lndi = 0;
            Program.f3.S_ni = 0;
            Program.f3.sum1 = 0;
            Program.f3.N1 = 0;
            Program.f3.S = 0;
            Program.f3.S1 = 0;
            Program.f3.S2 = 0;
            for (int i = 0; i < Program.f3.N; i++)
            {
                Program.f3.sumni = Program.f3.sumni + Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[1].Value);
                Program.f3.S_ni = Program.f3.S_ni + Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[1].Value) * Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[0].Value);
                Program.f3.sum_Y = Program.f3.sum_Y + Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[2].Value);
            }
            //Dcp = S_ni / sumni;
            Program.f3.sumni = 0;
            try
            {
                Program.f3.y_min = 101;
                    W_max: while (Program.f3.N1 < Program.f3.N - 1 && Convert.ToDouble(Program.f3.dataGridView1.Rows[Program.f3.N1].Cells[2].Value) < Convert.ToDouble(Program.f3.dataGridView1.Rows[Program.f3.N1 + 1].Cells[2].Value))
                    Program.f3.N1 = Program.f3.N1 + 1;
                Program.f3.y_max[Program.f3.kol_y_max] = Convert.ToDouble(Program.f3.dataGridView1.Rows[Program.f3.N1].Cells[2].Value);
                Program.f3.N_max[Program.f3.kol_y_max] = Program.f3.N1;
                Program.f3.kol_y_max = Program.f3.kol_y_max + 1;
                    if (Program.f3.N1 < Program.f3.N - 2)
                        goto W_min;
                    W_min: while (Program.f3.N1 < Program.f3.N - 1 && Convert.ToDouble(Program.f3.dataGridView1.Rows[Program.f3.N1].Cells[2].Value) >= Convert.ToDouble(Program.f3.dataGridView1.Rows[Program.f3.N1 + 1].Cells[2].Value))
                    Program.f3.N1 = Program.f3.N1 + 1;
                    if (Program.f3.N1 < Program.f3.N - 2)
                        goto W_max;
                    //for (int i = 0; i <= Program.f3.kol_y_max; i++)
                   // {
                     //   for (int j = 0; j <= Program.f3.kol_y_max; j++)
                     //   {
                        //    if (Program.f3.y_max[j] <= Program.f3.y_max[j + 1])
                        //    {
                         //       double max = Program.f3.y_max[j + 1];
                         //   Program.f3.y_max[j] = max;
                         //   }
                       // }
                   // }
                Program.f3.y_max1 = Program.f3.y_max[0];
                Program.f3.y_max2 = Program.f3.y_max[1];
                    for (int i = 0; i <= Program.f3.kol_y_max; i++)
                    {
                        if (Convert.ToDouble(Program.f3.dataGridView1.Rows[Program.f3.N_max[i]].Cells[2].Value) == Program.f3.y_max1)
                        Program.f3.N_max1 = Program.f3.N_max[i];
                    }
                    for (int i = 0; i <= Program.f3.kol_y_max; i++)
                    {
                        if (Convert.ToDouble(Program.f3.dataGridView1.Rows[Program.f3.N_max[i]].Cells[2].Value) == Program.f3.y_max2)
                        Program.f3.N_max2 = Program.f3.N_max[i];
                    }
                    if (Program.f3.N_max1 < Program.f3.N_max2)
                    {
                    Program.f3.N1 = Program.f3.N_max1;
                    Program.f3.N2 = Program.f3.N_max2;
                    }
                    else
                    {
                    Program.f3.N1 = Program.f3.N_max2;
                    Program.f3.N2 = Program.f3.N_max1;
                    }
                    for (int i = Program.f3.N1; i <= Program.f3.N2; i++)
                    {
                        if (Convert.ToDouble(Program.f3.dataGridView1.Rows[i + 1].Cells[0].Value) - Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[0].Value) > Program.f3.r)
                        {
                        Program.f3.r1 = Convert.ToDouble(Program.f3.dataGridView1.Rows[i + 1].Cells[0].Value) - Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[0].Value);
                        Program.f3.N1_ = i;
                        Program.f3.N2_ = i + 1;
                        }

                    }
                                    
                    x = Convert.ToDouble(Program.f3.dataGridView1.Rows[Program.f3.N1].Cells[0].Value);
                    y= Convert.ToDouble(Program.f3.dataGridView1.Rows[Program.f3.N1].Cells[2].Value);
                chart1.Series[1].Points.AddXY(x, y);
                x = Convert.ToDouble(Program.f3.dataGridView1.Rows[Program.f3.N2].Cells[0].Value);
                y = Convert.ToDouble(Program.f3.dataGridView1.Rows[Program.f3.N2].Cells[2].Value);
                chart1.Series[1].Points.AddXY(x, y);

            }
            catch
            {

            }
        }
        

        private void chart1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (But1)
            {
                if (e.Button==MouseButtons.Left)
                {
                    if (n_row < 2)
                    {
                        for (int i = 0; i < Program.f3.N; i++)
                        {
                            if (x_mouseMove > Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[0].Value) - Program.f3.r / 2 && x_mouseMove < Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[0].Value) + Program.f3.r / 2)
                            {
                                x = Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[0].Value);
                                y = Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[2].Value);
                                chart1.Series[1].Points.AddXY(x, y);
                                if (n_row == 0)                                
                                    Program.f3.N1 = i;                                                                   
                                else
                                    Program.f3.N2 = i;
                                n_row = n_row + 1;
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Достигнуто максимальное количество выделяемых максимумов! Если хотите поменять положения, нажмите кнопку \"Изменить\"" , "Сообщение!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (Program.f3.N1>Program.f3.N2)
            {
                int NN = Program.f3.N1;
                Program.f3.N1 = Program.f3.N2;
                Program.f3.N2 = NN;
            }

            for (int i = Program.f3.N1; i <= Program.f3.N2; i++)
            {
                if (Convert.ToDouble(Program.f3.dataGridView1.Rows[i + 1].Cells[0].Value) - Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[0].Value) > Program.f3.r)
                {
                    Program.f3.r1 = Convert.ToDouble(Program.f3.dataGridView1.Rows[i + 1].Cells[0].Value) - Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[0].Value);
                    Program.f3.N1_ = i;
                    Program.f3.N2_ = i + 1;
                }

            }

            Program.f3.button4_Click = false;
            Program.f3.pictureBox2.Visible = false;
            Program.f3.textBox11.Visible = false;
            Program.f3.pictureBox3.Visible = false;
            Program.f3.textBox12.Visible = false;
            Program.f3.pictureBox4.Visible = false;
            Program.f3.label5.Visible = false;
            Program.f3.label9.Visible = false;
            Program.f3.textBox22.Visible = true;
            Program.f3.textBox22.Location = new Point(234, 480);

            Program.f3.chart2.Series[1].Points.Clear();
            Program.f3.chart2.Series[2].Points.Clear();
            Program.f3.chart2.Series[5].Points.Clear();
            Program.f3.richTextBox1.Clear();

            if (Program.f3.r < Program.f3.r1) // случай когда есть разрыв между столбцами с двумя распределениями
                    {
                        for (int i = 0; i <= Program.f3.N1_; i++)
                        {
                    Program.f3.sumni = Program.f3.sumni + Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[1].Value);
                    Program.f3.S_ni_lndi = Program.f3.S_ni_lndi + Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[1].Value) * Math.Log(Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[0].Value)); 
                    Program.f3.S1 = Program.f3.S1 + Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[2].Value);
                        }
                Program.f3.par_a1 = Program.f3.S_ni_lndi / Program.f3.sumni;

                        for (int i = 0; i <= Program.f3.N1_; i++)
                    Program.f3.sum1 = Program.f3.sum1 + Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[1].Value) * (Math.Log(Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[0].Value)) - Program.f3.par_a1) * (Math.Log(Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[0].Value)) - Program.f3.par_a1);
                Program.f3.par_b1 = Math.Sqrt(Program.f3.sum1 / (Program.f3.sumni - 1));

                        for (double i = 0.1; i < Program.f3.X_max + Program.f3.r; i = i + 0.1)
                        {
                    Program.f3.fDi = Program.f3.r / (Program.f3.par_b1 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - Program.f3.par_a1) * (Math.Log(i) - Program.f3.par_a1) / (2 * Program.f3.par_b1 * Program.f3.par_b1)) * Program.f3.S1;
                    Program.f3.chart2.Series[1].Points.AddXY(i, Program.f3.fDi);
                            if (Program.f3.MaxShiftFDi1 < Program.f3.fDi)
                            {
                        Program.f3.MaxShiftFDi1 = Program.f3.fDi;
                        Program.f3.position_MaxShiftFDi1 = i;
                            }
                        }

                Program.f3.richTextBox1.SelectedRtf = Properties.Resources.Bi_log_norm;
                Program.f3.textBox4.Text = Math.Round(Program.f3.par_b1, 3).ToString("0.000");
                Program.f3.textBox3.Text = Math.Round(Program.f3.par_a1, 3).ToString("0.000");
                Program.f3.textBox5.Text = Math.Round(Program.f3.par_b1, 3).ToString("0.000");
                Program.f3.richTextBox1.Width = 309;
                Program.f3.richTextBox1.Height = 130;
                Program.f3.sumni = 0;
                Program.f3.S_ni_lndi = 0;
                Program.f3.sum1 = 0;
                        for (int i = Program.f3.N2_; i < Program.f3.N - 1; i++)
                        {
                    Program.f3.sumni = Program.f3.sumni + Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[1].Value);
                    Program.f3.S_ni_lndi = Program.f3.S_ni_lndi + Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[1].Value) * Math.Log(Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[0].Value));
                    Program.f3.S2 = Program.f3.S2 + Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[2].Value);
                        }
                Program.f3.par_a2 = Program.f3.S_ni_lndi / Program.f3.sumni;

                        for (int i = Program.f3.N2_; i < Program.f3.N - 1; i++)
                    Program.f3.sum1 = Program.f3.sum1 + Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[1].Value) * (Math.Log(Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[0].Value)) - Program.f3.par_a2) * (Math.Log(Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[0].Value)) - Program.f3.par_a2);
                Program.f3.par_b2 = Math.Sqrt(Program.f3.sum1 / (Program.f3.sumni - 1));

                Program.f3.textBox3.Location = new Point(247, 378);
                Program.f3.textBox5.Location = new Point(234, 401);


                        for (double i = 0.1; i < Program.f3.X_max + Program.f3.r; i = i + 0.1)
                        {
                    Program.f3.fDi = Program.f3.r / (Program.f3.par_b2 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - Program.f3.par_a2) * (Math.Log(i) - Program.f3.par_a2) / (2 * Program.f3.par_b2 * Program.f3.par_b2)) * Program.f3.S2;
                    Program.f3.chart2.Series[2].Points.AddXY(i, Program.f3.fDi);
                            if (Program.f3.MaxShiftFDi2 < Program.f3.fDi)
                            {
                        Program.f3.MaxShiftFDi2 = Program.f3.fDi;
                        Program.f3.position_MaxShiftFDi2 = i;
                            }
                    Program.f3.fDiCom = (Program.f3.r / (Program.f3.par_b1 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - Program.f3.par_a1) * (Math.Log(i) - Program.f3.par_a1) / (2 * Program.f3.par_b1 * Program.f3.par_b1)) * Program.f3.S1) + (Program.f3.r / (Program.f3.par_b2 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - Program.f3.par_a2) * (Math.Log(i) - Program.f3.par_a2) / (2 * Program.f3.par_b2 * Program.f3.par_b2)) * Program.f3.S2);
                    Program.f3.chart2.Series[5].Points.AddXY(i, Program.f3.fDiCom);
                        }
                        for (int i = 0; i < Program.f3.N; i++)
                        {
                            if (Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[2].Value) != 0)
                             {
                        Program.f3.SSres = Program.f3.SSres + Math.Pow((Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[2].Value) - (Program.f3.r / (Program.f3.par_b1 * (Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[0].Value)) - Program.f3.par_a1) * (Math.Log(Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[0].Value)) - Program.f3.par_a1) / (2 * Program.f3.par_b1 * Program.f3.par_b1)) * Program.f3.S1) - (Program.f3.r / (Program.f3.par_b2 * (Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[0].Value)) - Program.f3.par_a2) * (Math.Log(Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[0].Value)) - Program.f3.par_a2) / (2 * Program.f3.par_b2 * Program.f3.par_b2)) * Program.f3.S2)), 2);
                        Program.f3.SStot = Program.f3.SStot + Math.Pow((Convert.ToDouble(Program.f3.dataGridView1.Rows[i].Cells[2].Value) - Program.f3.sum_Y / Program.f3.N), 2);
                             }
                        }
                Program.f3.R2 = 1 - Program.f3.SSres / Program.f3.SStot;
                Program.f3.textBox22.Text = Math.Round(Program.f3.R2, 4).ToString("0.0000");
                    }
            Program.f3.Norm = false;
            Program.f3.LogNorm = true;
            Program.f3.button3_Click = true;           
            Program.f3.chart2.Series[1].IsVisibleInLegend = true;
            if (Program.f1.Rus == true)
                Program.f3.chart2.Series[1].Name = "Лог-норм 1";
            if (Program.f1.Eng == true)
                Program.f3.chart2.Series[1].Name = "Log-Norm 1";
            Program.f3.chart2.Series[2].IsVisibleInLegend = true;
            if (Program.f1.Rus == true)
                Program.f3.chart2.Series[2].Name = "Лог-норм 2";
            if (Program.f1.Eng == true)
                Program.f3.chart2.Series[2].Name = "Log-Norm 2";
            if (Program.f1.Rus == true)
                Program.f3.chart2.Series[5].Name = "Общая Лог-норм";
            if (Program.f1.Eng == true)
                Program.f3.chart2.Series[5].Name = "Log-Norm common";
            Program.f3.chart2.Series[5].IsVisibleInLegend = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            chart1.Series[1].Points.Clear();
            But1 = true;
            n_row = 0;
        }

        private void chart1_MouseMove(object sender, MouseEventArgs e)
        {
            try
            {
                lastpoint = e.Location;
                x_mouseMove = Math.Round(chart1.ChartAreas[0].AxisX.PixelPositionToValue(lastpoint.X),1);
                
                
            }
            catch
            {
                return;
            }
        }

        private void Form7_FormClosing(object sender, FormClosingEventArgs e)
        {
            Program.f3.F7 = false;
            
        }
    }
}
