using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using ExcelObj = Microsoft.Office.Interop.Excel;
using System.IO;
using WindowsFormsApplication8;
using System.Runtime.InteropServices;

namespace WindowsFormsApplication8
{

    public partial class Form3 : Form
    {
        private int X = Screen.PrimaryScreen.WorkingArea.Width;
        private int Y = Screen.PrimaryScreen.WorkingArea.Height;
        private string Dcp10 = "Средний линейный размер", median = "Медиана распределения", mode = "Мода распределения", Dcp30="Средний кубический размер",
            Dcp32="Средний объемно-поверхностный размер";
        public string line_raspr = "Линейное распределение", cub_raspr = "Кубическое распределение", surf_raspr = "Поверхностное распределение", vol_surf_raspr = "Объемно-поверхностное распределение", vol_weight_raspr = "Взвешенное по объему распределение"; 
        public Form f7, f11, f12, f13, f14, f15;
        public int size_H;
        public int size_W;
        private double W1;
        private double H1;
        private double C1H;
        private double C1W;
        public int size_H1 = 150;
        public int size_W1 = 150;
        public int L = 100;
        public int N;
        public int N1, N1_;
        public int N2, N2_;
        public int N_max1;
        public int N_max2;
        public int N30, N32, N43, N20;
        public double Sum = 0;
        public double V;
        public double x;
        public double y;
        public double X_min, X_min30, X_min32, X_min43, X_min20;
        public double X_max = 0, X_max30, X_max32, X_max43, X_max20;
        public double X_interval, X_interval30, X_interval32, X_interval43, X_interval20;
        public int I = 12, I3 = 1;
        public double sumni;
        public double S_ni_lndi;
        public double par_a, par_a_30, par_a_32, par_a_43, par_a_20, par_a1, par_a2, par_a0, par_a0_30, par_a0_32, par_a0_43, par_a0_20;
        public double sum1, Area, Area0, Area_30, Area0_30, Area_32, Area0_32, Area_43, Area0_43, Area_20, Area0_20;
        public double par_b, par_b_30, par_b_32, par_b_43, par_b_20, par_b1, par_b1_2, par_b2, par_b2_2, Mo, m, par_b0, par_b0_30, par_b0_32, par_b0_43, par_b0_20;
        public double r, r1 = 0;
        public double fDi, fvDi, fDiCom;
        public double S_ni;
        public double[] y_max;
        public int kol_y_max;
        public int[] N_max;
        private int N_min;
        public double y_min;
        public double y_max1;
        public double y_max2;
        public double S, S1, S2, sum_Y, sum_Y_30, sum_Y_32, sum_Y_43, sum_Y_20;
        //private double[] Vol;
        public List<double> Vol = new List<double>();
        private double sum_Vol;
        //private double[] percent_Vol;
        private double[] bi_modal;
        private double[] error;
        private double error_sum;
        public int error_sumi;
        private double error_sumj1, dataGV, dataGV_zap;
        private double[] error_sum1;
        private double[] error_sum2;
        private bool flag_cub_raspr = false, flag_surf_raspr=false, flag_vol_surf_raspr=false, flag_volweight_raspr=false;
        public bool flag_lin_raspr = true;
        public bool open_f11 = false, open_f12=false, open_f13=false, open_f14=false, open_f15=false;
        public int tSCB2 = 2;
        public bool exit_Log_Norm = false;
        public double MaxShiftFDi1 = 0, MaxShiftFDi2 = 0;
        public double position_MaxShiftFDi1, position_MaxShiftFDi2, R02;
        private int pos_trackBar1 = 0, pos_trackBar2 = 0;
        public bool button3_Click = false;
        public bool button4_Click = false;
        public bool Norm = false, LogNorm = false, Norm_30=false, LogNorm_30=false, Norm_32=false, LogNorm_32=false, Norm_43=false, LogNorm_43=false, LogNorm_20=false, Norm_20=false;
        public bool F7 = false, Ln=false, Ln30=false, Ln32=false, Ln43=false, Ln20=false;
        public double D10, D30, D32, D43, D20;
        private double[] crystalSize1;
        private double[] crystalSize2;
        private int NCSR1, NCSR2;
        private double[] crystalFreq1;
        private double[] crystalFreq2;
        public bool CSR = false, CSR30=false, CSR32=false, CSR43=false, CSR20=false;
        public bool CSR1 = false, CSR1_30=false, CSR1_32=false, CSR1_43=false, CSR1_20=false;
        public bool CSR2 = false, CSR2_30=false, CSR2_32=false, CSR2_43=false, CSR2_20=false;
        public bool Part = false, Part30=false, Part32=false, Part43=false, Part20=false;
        private bool MD = false, KD = false, TB = false;
        private double traB1, traB5;
        private double DCSR1;
        private double DvCSR1;
        private double DCSR2;
        private double DvCSR2;
        private int clickRaspred1 = 0;
        private int clickRaspred2 = 0;
        private int clickRaspredCommon = 0;
        private int clickCSR1 = 0;
        private int clickCSR2 = 0;
        private int click_toolStr_button1 = 0;
        private int click_toolStr_button2 = 0;
        public double R2, SSres, SStot, Dcp;
        private string[] X_X;
        private ToolTip t = new ToolTip();
        private int MouseUp_X, MouseUp_Y;
        public bool tS1 = false, analyt_func=false, analyt_func_20=false, analyt_func_30=false, analyt_func_32=false, analyt_func_43=false;
        private bool tB = false;
        public Form3()
        {
            Program.f3 = this; // теперь f3 будет ссылкой на форму Form3
            InitializeComponent();
            ClientSize = new System.Drawing.Size(X, Y);
            y_max = new double[1000];
            N_max = new int[1000];
            //Vol = new double[1000];            
            bi_modal = new double[100];
            //percent_Vol = new double[1000];
            error = new double[1000];
            error_sum1 = new double[1000];
            error_sum2 = new double[1000];
            crystalSize1 = new double[2000];
            crystalSize2 = new double[2000];
            crystalFreq1 = new double[2000];
            crystalFreq2 = new double[2000];
            X_X = new string[100];
        }
        // private string controlsInfoStr;

        [DllImport("user32.dll")]
        static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        [Flags]
        public enum MouseEventFlags
        {
            LEFTDOWN = 0x00000002,
            LEFTUP = 0x00000004,
            MIDDLEDOWN = 0x00000020,
            MIDDLEUP = 0x00000040,
            MOVE = 0x00000001,
            ABSOLUTE = 0x00008000,
            RIGHTDOWN = 0x00000008,
            RIGHTUP = 0x00000010
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            Program.f1.f3_open = true;
            ControlMoverOrResizerForForm3.Init(chart2);
            ControlMoverOrResizerForForm3.WorkType = ControlMoverOrResizerForForm3.MoveOrResize.MoveAndResize;
            //controlsInfoStr = ControlMoverOrResizer.GetSizeAndPositionOfControlsToString(this);
            textBox4.BorderStyle = BorderStyle.None;
            textBox3.BorderStyle = BorderStyle.None;
            textBox5.BorderStyle = BorderStyle.None;
            textBox22.BorderStyle = BorderStyle.None;
            textBox34.BorderStyle = BorderStyle.None;
            textBox30.BorderStyle = BorderStyle.None;
            textBox32.BorderStyle = BorderStyle.None;
            textBox31.BorderStyle = BorderStyle.None;
            textBox47.BorderStyle = BorderStyle.None;
            textBox43.BorderStyle = BorderStyle.None;
            textBox45.BorderStyle = BorderStyle.None;
            textBox44.BorderStyle = BorderStyle.None;
            textBox56.BorderStyle = BorderStyle.None;
            textBox58.BorderStyle = BorderStyle.None;
            textBox60.BorderStyle = BorderStyle.None;
            textBox57.BorderStyle = BorderStyle.None;
            textBox67.BorderStyle = BorderStyle.None;
            textBox66.BorderStyle = BorderStyle.None;
            textBox68.BorderStyle = BorderStyle.None;
            textBox70.BorderStyle = BorderStyle.None;
            tabControl1.TabPages.Remove(tabPage2);
            tabControl1.TabPages.Remove(tabPage3);
            tabControl1.TabPages.Remove(tabPage4);
            tabControl1.TabPages.Remove(tabPage13);
            tabControl2.TabPages[1].Parent = null;
            tabControl3.TabPages[1].Parent = null;
            tabControl4.TabPages[1].Parent = null;
            tabControl5.TabPages[1].Parent = null;
            tabControl6.TabPages[1].Parent = null;

            pictureBox2.Visible = false;
            pictureBox3.Visible = false;
            //textBox11.Visible = false;
            textBox12.Visible = false;
            label5.Visible = false;
            label9.Visible = false;
            size_H = this.Height;
            size_W = this.Width;
            C1H = size_H - size_H1;
            C1W = size_W - size_W1;
            W1 = chart2.Width;
            H1 = chart2.Height;
            if (size_H == size_W)
                chart2.Width = size_W - size_H1;
            else
            {
                chart2.Width = size_H - size_H1;
                chart2.Height = size_H - size_H1;
                // ratio_s = W1 / H1;
                // ratio_p = C1W / C1H;
                // while (Math.Round(ratio_p, 2) != (Math.Round(ratio_s, 2)))
                // {
                // size_W1 = size_W1 + 1;
                // chart2.Width = size_W - size_W1;
                // C1W = size_W - size_W1;
                // ratio_p = C1W / C1H;
            }

            // }
            chart2.Location = new System.Drawing.Point(this.Width - chart2.Width - L, size_H1 - L + 15);
            dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;  // выравнивание заголовок в dataGridView
            if (Program.f1.Eng == true)
            {
                английсикйToolStripMenuItem_Click(sender, e);
            }
            if (Program.f1.mm == true)
            {
                toolStripComboBox2.SelectedIndex = 0;
                tSCB2 = toolStripComboBox2.SelectedIndex;
                if (Program.f1.Rus == true)
                {
                    dataGridView1.Columns[0].HeaderCell.Value = "Размер частиц, мм";
                    label5.Text = "мм";
                    label9.Text = "мм";
                }
                if (Program.f1.Eng == true)
                {
                    dataGridView1.Columns[0].HeaderCell.Value = "Particle size, mm";
                    label5.Text = "mm";
                    label9.Text = "mm";
                }
            }
            if (Program.f1.mkm == true)
            {
                toolStripComboBox2.SelectedIndex = 1;
                tSCB2 = toolStripComboBox2.SelectedIndex;
                if (Program.f1.Rus == true)
                {
                    dataGridView1.Columns[0].HeaderCell.Value = "Размер частиц, мкм";
                    label5.Text = "мкм";
                    label9.Text = "мкм";
                }
                if (Program.f1.Eng == true)
                {
                    dataGridView1.Columns[0].HeaderCell.Value = "Particle size, mkm";
                    label5.Text = "mkm";
                    label9.Text = "mkm";
                }
            }
            if (Program.f1.nm == true)
            {
                toolStripComboBox2.SelectedIndex = 2;
                tSCB2 = toolStripComboBox2.SelectedIndex;
                if (Program.f1.Rus == true)
                {
                    dataGridView1.Columns[0].HeaderCell.Value = "Размер частиц, нм";
                    label5.Text = "нм";
                    label9.Text = "нм";
                }

                if (Program.f1.Eng == true)
                {
                    dataGridView1.Columns[0].HeaderCell.Value = "Particle size, nm";
                    label5.Text = "nm";
                    label9.Text = "nm";
                }
            }






        }


        public void button2_Click(object sender, EventArgs e)
        {


            if (Convert.ToString(dataGridView1.Rows[0].Cells[0].Value) == string.Empty)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Размеры частиц не заданы! Попробуйте снова!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles size are not determined! Try again!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (Convert.ToString(dataGridView1.Rows[0].Cells[2].Value) == string.Empty)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Доля частиц не задана! Попробуйте снова!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles fraction are not determined! Try again!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try
            {

                tabControl2.Visible = false;
                if (!open_f11)
                {
                    f11 = new Form11();
                    f11.Owner = this;
                    open_f11 = true;
                    f11.Show();
                }

                //if (open_f11)
                //{
                    //Program.f11.toolStripButton4.Checked = false;
                    //Program.f11.click_Ln = 0;
               // }
                Program.f11.chart2.Series[0].Points.Clear();
                Program.f11.chart2.Series[1].Points.Clear();
                Program.f11.chart2.Series[2].Points.Clear();
                Program.f11.chart2.Series[5].Points.Clear();
                richTextBox1.Clear();
                textBox3.Clear();
                textBox4.Clear();
                textBox5.Clear();
                pictureBox2.Visible = false;
                pictureBox3.Visible = false;
                //pictureBox4.Visible = false;
                //textBox11.Visible = false;
                textBox12.Visible = false;
                label5.Visible = false;
                label9.Visible = false;
                //Ln = false;
                analyt_func = false;


                //textBox22.Visible = false;
                N = Convert.ToInt32(dataGridView1.RowCount);

                while (Convert.ToString(dataGridView1.Rows[N - 1].Cells[0].Value) == string.Empty || Convert.ToString(dataGridView1.Rows[N - 1].Cells[2].Value) == string.Empty)
                    N = N - 1;
                if (Sum == 0)
                {
                    for (int i = 0; i < N; i++)
                    {
                        Sum = Sum + Convert.ToDouble(dataGridView1[1, i].Value); // сначал номер стобца, затем номер строки
                    }
                }
                if (Ln)
                {
                    for (int i = 0; i < N; i++)
                    {
                        x = Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value);
                        y = Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                        Program.f11.chart2.Series[0].Points.AddXY(x, y);
                        X_max = Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value);

                    }
                    Program.f11.r = Convert.ToDouble(dataGridView1.Rows[N - 1].Cells[3].Value) - Convert.ToDouble(dataGridView1.Rows[N - 2].Cells[3].Value);
                    X_max = X_max + Program.f11.r;
                    Program.f11.chart2.ChartAreas[0].AxisY.Minimum = 0;
                    Program.f11.chart2.ChartAreas[0].AxisX.Maximum = X_max;
                    Program.f11.chart2.ChartAreas[0].AxisX.Minimum = Convert.ToDouble(dataGridView1.Rows[0].Cells[3].Value) - Program.f11.r;
                    Program.f11.chart2.ChartAreas[0].AxisX.Title = "Ln(D)";
                    Program.f11.chart2.ChartAreas[0].AxisX.Interval = 0;

                   
                    //Ln = true;
                }
                else
                {
                    for (int i = 0; i < N; i++)
                    {
                        x = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                        y = Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                        Program.f11.chart2.Series[0].Points.AddXY(x, y);
                        X_max = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                    }
                    Program.f11.chart2.Series[0].IsVisibleInLegend = true;
                    Program.f11.chart2.Series[1].IsVisibleInLegend = false;
                    Program.f11.chart2.Series[2].IsVisibleInLegend = false;
                    Program.f11.chart2.Series[3].IsVisibleInLegend = false;
                    Program.f11.chart2.Series[4].IsVisibleInLegend = false;
                    Program.f11.chart2.Series[5].IsVisibleInLegend = false;
                    if (Program.f1.Rus == true)
                        Program.f11.chart2.Series[0].Name = "Гист.";
                    if (Program.f1.Eng == true)
                        Program.f11.chart2.Series[0].Name = "Hist.";
                    Part = true;
                    r = Convert.ToDouble(dataGridView1.Rows[1].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value);
                    X_max = X_max + r;
                    Program.f11.chart2.ChartAreas[0].AxisY.Minimum = 0;
                    Program.f11.chart2.ChartAreas[0].AxisX.Maximum = X_max;
                    if (Part == true && CSR == true)
                        Program.f11.chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), нм";
                    else
                        Program.f11.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, нм";
                    Program.f11.chart2.ChartAreas[0].AxisX.TitleFont = new Font("Times New Roman", Program.f11.I, FontStyle.Bold);
                    if (Part == true && CSR == true)
                        Program.f11.chart2.ChartAreas[0].AxisY.Title = "Доля частиц (ОКР), %";
                    else
                        Program.f11.chart2.ChartAreas[0].AxisY.Title = "Доля частиц, %";
                    if (Program.f1.Eng == true)
                        Program.f11.chart2.ChartAreas[0].AxisY.Title = "Frequency, %";
                    Program.f11.chart2.ChartAreas[0].AxisY.TitleFont = new Font("Times New Roman", Program.f11.I, FontStyle.Bold);
                    Program.f11.chart2.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Times New Roman", Program.f11.I, System.Drawing.FontStyle.Regular);
                    Program.f11.chart2.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Times New Roman", Program.f11.I, System.Drawing.FontStyle.Regular);
                    if (toolStripComboBox2.SelectedIndex == 0)
                    {
                        if (Program.f1.Rus == true)
                        {
                            if (Part == true && CSR == true)
                                Program.f11.chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), мм (нм)";
                            else
                                Program.f11.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, мм";
                            label5.Text = "мм";
                            label21.Text = "мм";
                            label20.Text = "мм";
                        }
                        if (Program.f1.Eng == true)
                        {
                            if (Part == true && CSR == true)
                                Program.f11.chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), mm (nm)";
                            else
                                Program.f11.chart2.ChartAreas[0].AxisX.Title = "Particle size, mm";
                            label5.Text = "mm";
                            label21.Text = "mm";
                            label20.Text = "mm";
                        }
                    }
                    if (toolStripComboBox2.SelectedIndex == 1)
                    {
                        if (Program.f1.Rus == true)
                        {
                            if (Part == true && CSR == true)
                                Program.f11.chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), мкм (нм)";
                            else
                                Program.f11.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, мкм";
                            label5.Text = "мкм";
                            label21.Text = "мкм";
                            label20.Text = "мкм";
                        }
                        if (Program.f1.Eng == true)
                        {
                            if (Part == true && CSR == true)
                                Program.f11.chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), mkm (nm)";
                            else
                                Program.f11.chart2.ChartAreas[0].AxisX.Title = "Particle size, mkm";
                            label5.Text = "mkm";
                            label21.Text = "mkm";
                            label20.Text = "mkm";
                        }
                    }
                    if (toolStripComboBox2.SelectedIndex == 2)
                    {
                        if (Program.f1.Rus == true)
                        {
                            if (Part == true && CSR == true)
                                Program.f11.chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), нм";
                            else
                                Program.f11.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, нм";
                            label5.Text = "нм";
                            label21.Text = "нм";
                            label20.Text = "нм";
                        }
                        if (Program.f1.Eng == true)
                        {
                            if (Part == true && CSR == true)
                                Program.f11.chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), nm";
                            else
                                Program.f11.chart2.ChartAreas[0].AxisX.Title = "Particle size, nm";
                            label5.Text = "nm";
                            label21.Text = "nm";
                            label20.Text = "nm";
                        }


                    }


                    if (dataGridView1.Rows.Count > 2)
                    {
                        X_interval = Convert.ToDouble(dataGridView1.Rows[1].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value);
                        X_min = Convert.ToDouble(dataGridView1.Rows[0].Cells[0].Value) - X_interval;
                        Program.f11.chart2.ChartAreas[0].AxisX.Minimum = X_min;
                        Program.f11.chart2.ChartAreas[0].AxisX.Interval = X_interval;
                        if (X_min <= 0)
                        {
                            Program.f11.chart2.ChartAreas[0].AxisX.IsStartedFromZero = true;
                            for (int i = 0; i < N; i++)
                                Program.f11.chart2.Series[0].Points[i].AxisLabel = (dataGridView1.Rows[i].Cells[0].Value).ToString();
                        }
                    }
                    if (CSR1 == true)
                    {
                        Program.f11.chart2.Series[3].Points.Clear();
                        for (int i = 0; i < NCSR1; i++)
                        {
                            x = crystalSize1[i];
                            y = crystalFreq1[i] * 50 * r;
                            Program.f11.chart2.Series[3].Points.AddXY(x, y);
                        }
                    }
                    if (CSR2 == true)
                    {
                        Program.f11.chart2.Series[4].Points.Clear();
                        for (int i = 0; i < NCSR2; i++)
                        {
                            x = crystalSize2[i];
                            y = crystalFreq2[i] * 50 * r;
                            Program.f11.chart2.Series[4].Points.AddXY(x, y);
                        }
                    }
                    if (CSR1 == true && CSR2 == true)
                    {
                        groupBox2.Visible = true;
                        trackBar5.Value = 500;
                        traB5 = trackBar5.Value / 10;
                        textBox20.Text = Math.Round(traB5, 1).ToString("0.0");
                    }
                }
            }
            catch
            {
                return;
            }
        }




        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private void английсикйToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Program.f1.Eng = true;
            Program.f1.Rus = false;
            Program.f1.Grid = "Show / hide the grid";
            Program.f1.Legend = "Show/hide legend";
            Program.f1.font = "Font size";
            Program.f1.Rotate = "Rotate X values on";
            Program.f1.Line_thickness = "Line thickness";
            Program.f1.Approx = "Show/hide curves of";
            Program.f1.Calibrate = "Calibrate Image";
            Program.f1.Histogram = "Plot histogram";
            Program.f1.Return = "Return";
            Program.f1.Zoom = "Zoom Image";
            Program.f1.LogNorm = "Go to Log & Norm DF 2.0";
            Program.f1.Electronogram = "Electron Diffraction Pattern Analysis";
            Program.f1.Average = "Average Size";
            Program.f1.Calculate_d = "Interatomic distances calculation";
            Program.f1.Calculate_C = "Microscope constant calculation";
            Program.f1.LineProfile = "Draw a line profile";
            Program.f1.groupBox3.Text = Program.f1.Electronogram;
            Program.f1.tabControl1.TabPages[0].Text = "Ring diameters";
            Program.f1.tabControl1.TabPages[1].Text = "Calibration sample";
            Program.f1.button3.Text = "Calculation d";
            Program.f1.button4.Text = "Calculation C";
            Program.f1.button16.Text = "Clear lines";
            Program.f1.groupBox4.Text = "Image control";
            Program.f1.groupBox1.Text = "Instruments";
            Program.f1.groupBox2.Text = "Measurement results";
            файлToolStripMenuItem.Text = "File";
            открытьФайлtxtToolStripMenuItem.Text = "Load file at the table";
            новоеРаспределениеToolStripMenuItem.Text = "New data distribution";
            сохранитьИзображениеToolStripMenuItem.Text = "Save image";
            выходToolStripMenuItem.Text = "Exit";
            свойстваТаблицыToolStripMenuItem.Text = "Diagram";
            описатьToolStripMenuItem.Text = "Fit";
            логнормФияToolStripMenuItem.Text = "Log-normal function";
            нормФцияГауссToolStripMenuItem.Text = "Normal function (Gauss)";
            бимодальноеToolStripMenuItem.Text = "Bimodal";
            логнормФцияToolStripMenuItem2.Text = "Log-normal function";
            нормФцияГауссToolStripMenuItem2.Text = "Normal function (Gauss)";
            единицаИзмеренияToolStripMenuItem.Text = "Unit measure";
            toolStripComboBox2.Items.Clear();
            toolStripComboBox2.Text = "nm";
            toolStripComboBox2.Items.AddRange(new object[] { "mm", "mkm", "nm" });
            toolStripComboBox2.SelectedIndex = tSCB2;
            кубическоеРаспределениеToolStripMenuItem.Text = "Cubic distribution";
            линейноеРаспределениеToolStripMenuItem.Text = "Linear distribution";
            максЗначениеОсиXToolStripMenuItem.Text = "Max. value of the X-axis";
            groupBox3.Text = "Max. value of the X-axis";
            распределениеОКРToolStripMenuItem.Text = "CSR distribution";
            распределение1ToolStripMenuItem.Text = "Load the 1-st distribution";
            распределение2ToolStripMenuItem.Text = "Load the 2-nd distribution";
            отобразитьToolStripMenuItem.Text = "Plot the graph";
            label12.Text = "nm";
            label13.Text = "nm";
            label14.Text = "nm";
            label15.Text = "nm";
            pictureBox7.Image = Properties.Resources.DCSR1;
            pictureBox8.Image = Properties.Resources.D_v_CSR1;
            pictureBox9.Image = Properties.Resources.DCSR2;
            pictureBox10.Image = Properties.Resources.D_v_CSR2;
            groupBox2.Text = "1-st CSR distribution area";
            chart2.Series[0].Name = "Hist.";
            if (chart2.Series[1].Name == "Лог-норм 1")
                chart2.Series[1].Name = "Log-Norm 1";
            if (chart2.Series[1].Name == "Гаусс 1")
                chart2.Series[1].Name = "Gauss 1";
            if (chart2.Series[2].Name == "Лог-норм 2")
                chart2.Series[2].Name = "Log-Norm 2";
            if (chart2.Series[2].Name == "Гаусс 2")
                chart2.Series[2].Name = "Gauss 2";
            chart2.Series[3].Name = "CSR 1";
            chart2.Series[4].Name = "CSR 2";
            chart2.Series[5].Name = "Log-Norm common";
            языкToolStripMenuItem.Text = "Language";
            русскийToolStripMenuItem.Text = "Russian";
            английсикйToolStripMenuItem.Text = "English";
            ифнформацияToolStripMenuItem.Text = "Information";
            chart2.ChartAreas[0].AxisY.Title = "Frequency, %";
            Program.f11.chart2.ChartAreas[0].AxisY.Title = "Frequency, %";
            dataGridView1.Columns[0].HeaderCell.Value = "Particle size, nm";
            dataGridView1.Columns[1].HeaderCell.Value = "Particles number";
            dataGridView1.Columns[2].HeaderCell.Value = "Particles fraction, %";
            button2.Text = "Plot a histogram";
            line_raspr = "Linear distribution";
            surf_raspr = "Surface distribution";
            cub_raspr = "Volume distribution";
            vol_surf_raspr = "Surface-weighted distribution";
            vol_weight_raspr = "Volume-weighted distribution";

            if (toolStripComboBox2.SelectedIndex == 0)
            {               
                if (CSR == true)
                    chart2.ChartAreas[0].AxisX.Title = "Crystallite size, nm";
                if (Part == true)
                    chart2.ChartAreas[0].AxisX.Title = "Particle size, mm";
                if (Part == true && CSR == true)
                    chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), mm (nm)";
                dataGridView1.Columns[0].HeaderCell.Value = "Particle size, mm";
                label5.Text = "mm";
                label9.Text = "mm";
            }
            if (toolStripComboBox2.SelectedIndex == 1)
            {
                if (CSR == true)
                    chart2.ChartAreas[0].AxisX.Title = "Crystallite size, nm";
                if (Part == true)
                    chart2.ChartAreas[0].AxisX.Title = "Particle size, mkm";
                if (Part == true && CSR == true)
                    chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), mkm (nm)";
                dataGridView1.Columns[0].HeaderCell.Value = "Particle size, mkm";
                label5.Text = "mkm";
                label9.Text = "mkm";
            }
            if (toolStripComboBox2.SelectedIndex == 2)
            {
                Program.f11.chart2.ChartAreas[0].AxisX.Title = "Particle size, nm";
                if (CSR == true)
                    chart2.ChartAreas[0].AxisX.Title = "Crystallite size, nm";
                if (Part == true)
                    chart2.ChartAreas[0].AxisX.Title = "Particle size, nm";
                if (Part == true && CSR == true)
                    chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), nm";
                dataGridView1.Columns[0].HeaderCell.Value = "Particle size, nm";
                label5.Text = "nm";
                label9.Text = "nm";
            }

            оПрограммеИИнструкцииПоРаботеToolStripMenuItem.Text = "About DEAM and work instruction";
            форматОсиХToolStripMenuItem.Text = "X-axis format";
            диапазонРазмеровToolStripMenuItem.Text = "Sizes range";
            среднийРазмерДиапазонаToolStripMenuItem.Text = "Average sizes of the range";
            аппроксимацииToolStripMenuItem.Text = "Approximation";
            оКРToolStripMenuItem.Text = "CSR";
            распределение1ToolStripMenuItem3.Text = "Distribution 1";
            распределение2ToolStripMenuItem3.Text = "Distribution 2";
            общаябимодалToolStripMenuItem1.Text = "Common (bimodal)";
            распределение1ToolStripMenuItem4.Text = "Distribution 1";
            распределение2ToolStripMenuItem4.Text = "Distribution 2";
            Program.f1.файлToolStripMenuItem.Text = "File";
            Program.f1.открытьToolStripMenuItem.Text = "Open";
            Program.f1.сохранитьКакToolStripMenuItem.Text = "Save data";
            Program.f1.новоеРаспределениеToolStripMenuItem.Text = "New data distribution";
            Program.f1.загрузитьЭкспериментальныеДанныеToolStripMenuItem.Text = "Import experimental data";
            Program.f1.выходToolStripMenuItem.Text = "Exit";
            Program.f1.инструментыToolStripMenuItem.Text = "Gadgets";
            Program.f1.системаЕдиницToolStripMenuItem.Text = "Unit measure";
            Program.f1.ммToolStripMenuItem.Text = "mm";
            Program.f1.мкмToolStripMenuItem.Text = "mkm";
            Program.f1.нмToolStripMenuItem.Text = "nm";
            Program.f1.коэффициентЛупыToolStripMenuItem.Text = "Loupe magnification";
            Program.f1.информацияToolStripMenuItem.Text = "Information";
            Program.f1.button5.Text = "Change image position";
            Program.f1.языкToolStripMenuItem.Text = "Language";
            Program.f1.русскийToolStripMenuItem.Text = "Russian";
            Program.f1.englishToolStripMenuItem.Text = "English";
            Program.f1.оПрограммеИИнструкцииПоРаботеToolStripMenuItem.Text = "About DEAM and work instruction";

            Program.f1.groupBox1.Text = "Instruments";
            Program.f1.groupBox2.Text = "Measurement results";
            Program.f1.label1.Text = "d, nm";
            Program.f1.label4.Text = "d, nm";
            if (Program.f1.textBox5.Text == "пик")
            {
                Program.f1.textBox5.Text = "pix";
                Program.f1.textBox6.Text = "pix";
                Program.f1.label2.Text = "D, pix";
                Program.f1.label5.Text = "D, pix";
            }
            if (Program.f1.textBox5.Text == "нм")
            {
                Program.f1.textBox5.Text = "nm";
                Program.f1.textBox6.Text = "nm";
                Program.f1.label2.Text = "D, nm";
                Program.f1.label5.Text = "D, nm";
            }
            if (Program.f1.textBox5.Text == "мкм")
            {
                Program.f1.textBox5.Text = "mkm";
                Program.f1.textBox6.Text = "mkm";
                Program.f1.label2.Text = "D, mkm";
                Program.f1.label5.Text = "D, mkm";
            }
            if (Program.f1.textBox5.Text == "мм")
            {
                Program.f1.textBox5.Text = "mm";
                Program.f1.textBox6.Text = "mm";
                Program.f1.label2.Text = "D, mm";
                Program.f1.label5.Text = "D, mm";
            }
            if (Program.f1.textBox9.Text != String.Empty)
            {
                if (Program.f1.textBox9.Text == "пик")
                {
                    Program.f1.textBox9.Text = "pix";
                    Program.f1.textBox10.Text = "pix";
                }
                if (Program.f1.textBox9.Text == "нм")
                {
                    Program.f1.textBox9.Text = "nm";
                    Program.f1.textBox10.Text = "nm";
                }
                if (Program.f1.textBox9.Text == "мкм")
                {
                    Program.f1.textBox9.Text = "mkm";
                    Program.f1.textBox10.Text = "mkm";
                }
                if (Program.f1.textBox9.Text == "мм")
                {
                    Program.f1.textBox9.Text = "mm";
                    Program.f1.textBox10.Text = "mm";
                }
            }

            if (Program.f1.FormLineProphile)
            {
                Program.f10.Text = "Line profile";
                Program.f10.chart1.ChartAreas[0].AxisX.Title = Program.f1.textBox5.Text;
                Program.f10.LineChart = "Line chart";
                Program.f10.ColumnChart = "Column chart";
                Program.f10.Interval = "Interval measurement";
                Program.f10.toolStripMenuItem2.Text = "Discrete";
                Program.f10.toolStripMenuItem3.Text = "Continuous";
            }

            if (Program.f1.loupe)
                Program.f8.Text = "Zoom factor";

            if (Program.f1.clickCountCalibrationImage == 1)
            {
                Program.f9.Text = Program.f1.Calibrate;
                Program.f9.label1.Text = "Measure the feature of known length and enter the appropriate value";
                Program.f9.label2.Text = "Known length:";
                Program.f9.label2.Location = new Point(Program.f9.label2.Location.X + 10, Program.f9.label2.Location.Y);
                Program.f9.label3.Text = "Units:";
                Program.f9.label3.Location = new Point(Program.f9.label3.Location.X + 60, Program.f9.label3.Location.Y);
                Program.f9.comboBox1.Items.Clear();
                string[] mas = { "nm", "mkm", "mm", "pix" };
                Program.f9.comboBox1.Items.AddRange(mas);
                if (Program.f9.comboBox1.Text == "пик")
                    Program.f9.comboBox1.Text = "pix";
                if (Program.f9.comboBox1.Text == "нм")
                    Program.f9.comboBox1.Text = "nm";
                if (Program.f9.comboBox1.Text == "мкм")
                    Program.f9.comboBox1.Text = "mkm";
                if (Program.f9.comboBox1.Text == "мм")
                    Program.f9.comboBox1.Text = "mm";
                Program.f9.button2.Text = "Cancel";
            }

            if (Program.f1.form4)
            {
                Program.f4.Text = Program.f1.Histogram;
                Program.f4.label1.Text = "Enter the number of columns to plot the histogram";
                Program.f4.label2.Text = "Minimum size:";
                Program.f4.label3.Text = "Maximum size:";
                Program.f4.label4.Text = "Columns number:";
                Program.f4.button2.Text = "Cancel";
            }
        }

        private void toolStripComboBox2_Click(object sender, EventArgs e)
        {
            // if (Program.f1.Rus==true)
            //{
            //  if (toolStripComboBox2.SelectedIndex == 0)
            //dataGridView1.Columns[0].HeaderCell.Value = "Размер частиц, мм";
            //if (toolStripComboBox2.SelectedIndex == 1)
            //dataGridView1.Columns[0].HeaderCell.Value = "Размер частиц, мкм";
            //if (toolStripComboBox2.SelectedIndex == 2)
            //dataGridView1.Columns[0].HeaderCell.Value = "Размер частиц, нм";
            //}
        }

        private void русскийToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Program.f1.Eng = false;
            Program.f1.Rus = true;
            Program.f1.Grid = "Наложить/Убрать сетку";
            Program.f1.Legend = "Показать/скрыть легенду";
            Program.f1.font = "Размер шрифта";
            Program.f1.Rotate = "Поворот значений оси Х на";
            Program.f1.Line_thickness = "Толщина линий";
            Program.f1.Approx = "Показать/скрыть кривые";
            Program.f1.Calibrate = "Калибровка изображения";
            Program.f1.Histogram = "Построить гистограмму";
            Program.f1.Return = "Назад";
            Program.f1.Zoom = "Увеличить изображение";
            Program.f1.LogNorm = "Перейти в Log & Norm DF 2.0";
            Program.f1.Electronogram = "Анализ электронограмм";
            Program.f1.Average = "Средний размер";
            Program.f1.Calculate_d = "Расчет межатомных расстояний";
            Program.f1.Calculate_C = "Расчет постоянной микроскопа";
            Program.f1.button16.Text = "Очистить линии";
            Program.f1.LineProfile = "Линейный профиль";
            Program.f1.groupBox3.Text = Program.f1.Electronogram;
            Program.f1.tabControl1.TabPages[0].Text = "Диаметры колец";
            Program.f1.tabControl1.TabPages[1].Text = "Эталон";
            Program.f1.button3.Text = "Рассчитать d";
            Program.f1.button4.Text = "Рассчитать C";
            Program.f1.button16.Text = "Очистить линии";
            Program.f1.groupBox4.Text = "Параметры изображения";
            Program.f1.groupBox1.Text = "Инструменты";
            Program.f1.groupBox2.Text = "Результаты измерений";
            файлToolStripMenuItem.Text = "Файл";
            открытьФайлtxtToolStripMenuItem.Text = "Загрузить файл в таблицу";
            новоеРаспределениеToolStripMenuItem.Text = "Новое распределение";
            сохранитьИзображениеToolStripMenuItem.Text = "Сохранить изображение";
            выходToolStripMenuItem.Text = "Выход";
            свойстваТаблицыToolStripMenuItem.Text = "График";
            описатьToolStripMenuItem.Text = "Описать";
            логнормФияToolStripMenuItem.Text = "Лог-норм ф-ция";
            нормФцияГауссToolStripMenuItem.Text = "Норм ф-ция (Гаусс)";
            бимодальноеToolStripMenuItem.Text = "Бимодальное";
            логнормФцияToolStripMenuItem2.Text = "Лог-норм ф-ция";
            нормФцияГауссToolStripMenuItem2.Text = "Норм ф-ция (Гаусс)";
            единицаИзмеренияToolStripMenuItem.Text = "Единица измерения";
            toolStripComboBox2.Items.Clear();
            toolStripComboBox2.Text = "нм";
            toolStripComboBox2.Items.AddRange(new object[] { "мм", "мкм", "нм" });
            toolStripComboBox2.SelectedIndex = tSCB2;
            кубическоеРаспределениеToolStripMenuItem.Text = "Кубическое распределение";
            линейноеРаспределениеToolStripMenuItem.Text = "Линейное распределение";
            максЗначениеОсиXToolStripMenuItem.Text = "Макс. значение оси Х";
            groupBox3.Text = "Макс. значение оси Х";
            распределениеОКРToolStripMenuItem.Text = "Распределение ОКР";
            распределение1ToolStripMenuItem.Text = "Загрузить распределение 1";
            распределение2ToolStripMenuItem.Text = "Загрузить распределение 2";
            отобразитьToolStripMenuItem.Text = "Отобразить";
            label12.Text = "нм";
            label13.Text = "нм";
            label14.Text = "нм";
            label15.Text = "нм";
            pictureBox7.Image = Properties.Resources.Dокр1;
            pictureBox8.Image = Properties.Resources.D_v_окр1;
            pictureBox9.Image = Properties.Resources.Dокр2;
            pictureBox10.Image = Properties.Resources.D_v_окр2;
            groupBox2.Text = "Площадь распределения ОКР 1";
            chart2.Series[0].Name = "Гист.";
            if (chart2.Series[1].Name == "Log-Norm 1")
                chart2.Series[1].Name = "Лог-норм 1";
            if (chart2.Series[1].Name == "Gauss 1")
                chart2.Series[1].Name = "Гаусс 1";
            if (chart2.Series[2].Name == "Log-Norm 2")
                chart2.Series[2].Name = "Лог-норм 2";
            if (chart2.Series[2].Name == "Gauss 2")
                chart2.Series[2].Name = "Гаусс 2";
            chart2.Series[3].Name = "ОКР 1";
            chart2.Series[4].Name = "ОКР 2";
            chart2.Series[5].Name = "Общая Лог-норм";
            языкToolStripMenuItem.Text = "Язык";
            русскийToolStripMenuItem.Text = "Русский";
            английсикйToolStripMenuItem.Text = "Английский";
            ифнформацияToolStripMenuItem.Text = "Информация";
            chart2.ChartAreas[0].AxisY.Title = "Доля частиц, %";
            Program.f11.chart2.ChartAreas[0].AxisY.Title = "Доля частиц, %";
            dataGridView1.Columns[0].HeaderCell.Value = "Размер частиц, нм";
            dataGridView1.Columns[1].HeaderCell.Value = "Число частиц";
            dataGridView1.Columns[2].HeaderCell.Value = "Доля частиц, %";
            button2.Text = "Построить гистограмму";
            line_raspr = "Линейное распределение";
            surf_raspr = "Поверхностное распределение";
            cub_raspr = "Объемное распределение";
            vol_surf_raspr = "Объемно-поверхностное распределение";
            vol_weight_raspr = "Взвешенное по объему распределение";

            if (toolStripComboBox2.SelectedIndex == 0)
            {
                if (CSR == true)
                    chart2.ChartAreas[0].AxisX.Title = "Размер кристаллита, нм";
                if (Part == true)
                    chart2.ChartAreas[0].AxisX.Title = "Размер частиц, мм";
                if (Part == true && CSR == true)
                    chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), мм (нм)";
                dataGridView1.Columns[0].HeaderCell.Value = "Размер частиц, мм";
                label5.Text = "мм";
                label9.Text = "мм";
            }
            if (toolStripComboBox2.SelectedIndex == 1)
            {
                if (CSR == true)
                    chart2.ChartAreas[0].AxisX.Title = "Размер кристаллита, нм";
                if (Part == true)
                    chart2.ChartAreas[0].AxisX.Title = "Размер частиц, мкм";
                if (Part == true && CSR == true)
                    chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), мкм (нм)";
                dataGridView1.Columns[0].HeaderCell.Value = "Размер частиц, мкм";
                label5.Text = "мкм";
                label9.Text = "мкм";
            }
            if (toolStripComboBox2.SelectedIndex == 2)
            {
                Program.f11.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, нм";
                if (CSR == true)
                    chart2.ChartAreas[0].AxisX.Title = "Размер кристаллита, нм";
                if (Part == true)
                    chart2.ChartAreas[0].AxisX.Title = "Размер частиц, нм";
                if (Part == true && CSR == true)
                    chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), нм";
                dataGridView1.Columns[0].HeaderCell.Value = "Размер частиц, нм";
                label5.Text = "нм";
                label9.Text = "нм";
            }

            оПрограммеИИнструкцииПоРаботеToolStripMenuItem.Text = "О программе и инструкции по работе";
            форматОсиХToolStripMenuItem.Text = "Формат оси Х";
            диапазонРазмеровToolStripMenuItem.Text = "Диапазон размеров";
            среднийРазмерДиапазонаToolStripMenuItem.Text = "Средние размеры диапазона";
            аппроксимацииToolStripMenuItem.Text = "Аппроксимации";
            оКРToolStripMenuItem.Text = "ОКР";
            распределение1ToolStripMenuItem3.Text = "Распределение 1";
            распределение2ToolStripMenuItem3.Text = "Распределение 2";
            общаябимодалToolStripMenuItem1.Text = "Общая (бимодал)";
            распределение1ToolStripMenuItem4.Text = "Распределение 1";
            распределение2ToolStripMenuItem4.Text = "Распределение 2";
            Program.f1.файлToolStripMenuItem.Text = "Файл";
            Program.f1.открытьToolStripMenuItem.Text = "Открыть";
            Program.f1.сохранитьКакToolStripMenuItem.Text = "Сохранить распределение";
            Program.f1.новоеРаспределениеToolStripMenuItem.Text = "Новое распределение";
            Program.f1.загрузитьЭкспериментальныеДанныеToolStripMenuItem.Text = "Загрузить экспериментальные данные";
            Program.f1.выходToolStripMenuItem.Text = "Выход";
            Program.f1.инструментыToolStripMenuItem.Text = "Инструменты";
            Program.f1.системаЕдиницToolStripMenuItem.Text = "Единица измерения";
            Program.f1.ммToolStripMenuItem.Text = "мм";
            Program.f1.мкмToolStripMenuItem.Text = "мкм";
            Program.f1.нмToolStripMenuItem.Text = "нм";
            Program.f1.коэффициентЛупыToolStripMenuItem.Text = "Коэффициент лупы";
            Program.f1.информацияToolStripMenuItem.Text = "Информация";
            Program.f1.button5.Text = "Изменить положение изображения";
            Program.f1.языкToolStripMenuItem.Text = "Язык";
            Program.f1.оПрограммеИИнструкцииПоРаботеToolStripMenuItem.Text = "О программе и инструкции по работе";

            Program.f1.label1.Text = "d, нм";
            Program.f1.label4.Text = "d, нм";
            if (Program.f1.textBox5.Text == "pix")
            {
                Program.f1.textBox5.Text = "пик";
                Program.f1.textBox6.Text = "пик";
                Program.f1.label2.Text = "D, пик";
                Program.f1.label5.Text = "D, пик";
            }
            if (Program.f1.textBox5.Text == "nm")
            {
                Program.f1.textBox5.Text = "нм";
                Program.f1.textBox6.Text = "нм";
                Program.f1.label2.Text = "D, нм";
                Program.f1.label5.Text = "D, нм";
            }
            if (Program.f1.textBox5.Text == "mkm")
            {
                Program.f1.textBox5.Text = "мкм";
                Program.f1.textBox6.Text = "мкм";
                Program.f1.label2.Text = "D, мкм";
                Program.f1.label5.Text = "D, мкм";
            }
            if (Program.f1.textBox5.Text == "mm")
            {
                Program.f1.textBox5.Text = "мм";
                Program.f1.textBox6.Text = "мм";
                Program.f1.label2.Text = "D, мм";
                Program.f1.label5.Text = "D, мм";
            }

            if (Program.f1.textBox9.Text != String.Empty)
            {
                if (Program.f1.textBox9.Text == "pix")
                {
                    Program.f1.textBox9.Text = "пик";
                    Program.f1.textBox10.Text = "пик";
                }
                if (Program.f1.textBox9.Text == "nm")
                {
                    Program.f1.textBox9.Text = "нм";
                    Program.f1.textBox10.Text = "нм";
                }
                if (Program.f1.textBox9.Text == "mkm")
                {
                    Program.f1.textBox9.Text = "мкм";
                    Program.f1.textBox10.Text = "мкм";
                }
                if (Program.f1.textBox9.Text == "mm")
                {
                    Program.f1.textBox9.Text = "мм";
                    Program.f1.textBox10.Text = "мм";
                }
            }

            if (Program.f1.FormLineProphile)
            {
                Program.f10.Text = "Линейный профиль";
                Program.f10.chart1.ChartAreas[0].AxisX.Title = Program.f1.textBox5.Text;
                Program.f10.LineChart = "Линейная диаграмма";
                Program.f10.ColumnChart = "Столбчатая диаграмма";
                Program.f10.Interval = "Измерение интервала";
                Program.f10.toolStripMenuItem2.Text = "Дискретное";
                Program.f10.toolStripMenuItem3.Text = "Непрерывное";
            }

            if (Program.f1.loupe)
                Program.f8.Text = "Коэффициент лупы";

            if (Program.f1.clickCountCalibrationImage == 1)
            {
                Program.f9.Text = Program.f1.Calibrate;
                Program.f9.label1.Text = "Измерьте масштабную линейку и введите соответствующее значение";
                Program.f9.label2.Text = "Длина шкалы:";
                Program.f9.label2.Location = new Point(Program.f9.label2.Location.X - 10, Program.f9.label2.Location.Y);
                Program.f9.label3.Text = "Ед. измерения:";
                Program.f9.label3.Location = new Point(Program.f9.label3.Location.X - 60, Program.f9.label3.Location.Y);
                Program.f9.comboBox1.Items.Clear();
                string[] mas = { "нм", "мкм", "мм", "пик" };
                Program.f9.comboBox1.Items.AddRange(mas);
                if (Program.f9.comboBox1.Text == "pix")
                    Program.f9.comboBox1.Text = "пик";
                if (Program.f9.comboBox1.Text == "nm")
                    Program.f9.comboBox1.Text = "нм";
                if (Program.f9.comboBox1.Text == "mkm")
                    Program.f9.comboBox1.Text = "мкм";
                if (Program.f9.comboBox1.Text == "mm")
                    Program.f9.comboBox1.Text = "мм";
                Program.f9.button2.Text = "Отмена";
            }

            if (Program.f1.form4)
            {
                Program.f4.Text = "Построение гистограммы";
                Program.f4.label1.Text = "Укажите количество столбцов для построения гистограммы";
                Program.f4.label2.Text = "Мин. размер:";
                Program.f4.label3.Text = "Макс. размер:";
                Program.f4.label4.Text = "Кол-во столбцов:";
                Program.f4.button2.Text = "Отмена";
            }
        }

        private void trackBar2_Scroll(object sender, EventArgs e)
        {
            if (button3_Click == true)
            {
                chart2.Series[1].Points.Clear();
                chart2.Series[5].Points.Clear();

                for (double i = 0.1; i < X_max + r; i = i + 0.1)
                {
                    if (LogNorm == true)
                    {
                        fDi = r / (par_b1 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a1) * (Math.Log(i) - par_a1) / (2 * par_b1 * par_b1)) * S1;
                        fDiCom = (r / (par_b1 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a1) * (Math.Log(i) - par_a1) / (2 * par_b1 * par_b1)) * S1) + (r / (par_b2 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a2) * (Math.Log(i) - par_a2) / (2 * par_b2 * par_b2)) * S2);
                    }
                    if (Norm == true)
                    {
                        fDi = r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a1) * (i - par_a1) / (2 * par_b1 * par_b1)) * S1;
                        fDiCom = (r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a1) * (i - par_a1) / (2 * par_b1 * par_b1)) * S1) + (r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a2) * (i - par_a2) / (2 * par_b2 * par_b2)) * S2);
                    }
                    chart2.Series[1].Points.AddXY((i + Convert.ToDouble(pos_trackBar1) / 10), fDi);
                    chart2.Series[5].Points.AddXY((i + Convert.ToDouble(pos_trackBar1) / 10), fDiCom);
                    textBox3.Text = Math.Round(par_a1, 3).ToString("0.000");
                }
                R2 = 0;
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N; i++)
                {
                    //if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) != 0)
                    //{
                    if (LogNorm == true)
                    {
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1)) * S1) - (r / (par_b2 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) / (2 * par_b2 * par_b2)) * S2)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - Dcp), 2);
                    }
                    if (Norm == true)
                    {

                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1)) * S1) - (r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) / (2 * par_b2 * par_b2)) * S2)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - Dcp), 2);
                    }
                    // }
                }
                R2 = 1 - SSres / SStot;
                textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                if (LogNorm == true)
                {
                    D10 = Math.Exp(par_a1) * Math.Exp(0.5 * par_b1 * par_b1);

                }
                if (Norm == true)
                {
                    D10 = par_a1;

                }
            }
            if (button4_Click == true)
            {
                chart2.Series[2].Points.Clear();
                chart2.Series[5].Points.Clear();

                for (double i = 0.1; i < X_max + r; i = i + 0.1)
                {
                    if (LogNorm == true)
                    {
                        fDi = r / (par_b2 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a2) * (Math.Log(i) - par_a2) / (2 * par_b2 * par_b2)) * S2;
                        fDiCom = (r / (par_b1 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a1) * (Math.Log(i) - par_a1) / (2 * par_b1 * par_b1)) * S1) + (r / (par_b2 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a2) * (Math.Log(i) - par_a2) / (2 * par_b2 * par_b2)) * S2);
                    }
                    if (Norm == true)
                    {
                        fDi = r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a2) * (i - par_a2) / (2 * par_b2 * par_b2)) * S2;
                        fDiCom = (r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a1) * (i - par_a1) / (2 * par_b1 * par_b1)) * S1) + (r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a2) * (i - par_a2) / (2 * par_b2 * par_b2)) * S2);

                    }
                    chart2.Series[2].Points.AddXY((i + Convert.ToDouble(pos_trackBar2) / 10), fDi);
                    chart2.Series[5].Points.AddXY((i + Convert.ToDouble(pos_trackBar2) / 10), fDiCom);
                    //textBox6.Text = Math.Round(par_a2, 3).ToString("0.000");
                }
                R2 = 0;
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N; i++)
                {
                    //if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) != 0)
                    //{
                    if (LogNorm == true)
                    {
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1)) * S1) - (r / (par_b2 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) / (2 * par_b2 * par_b2)) * S2)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - Dcp), 2);
                    }
                    if (Norm == true)
                    {
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1)) * S1) - (r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) / (2 * par_b2 * par_b2)) * S2)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - Dcp), 2);
                    }
                    // }
                }
                R2 = 1 - SSres / SStot;
                textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                if (LogNorm == true)
                {
                    D10 = Math.Exp(par_a2) * Math.Exp(0.5 * par_b2 * par_b2);

                }
                if (Norm == true)
                {
                    D10 = par_a2;

                }
            }
        }

        private void распределение1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "Файлы excel (*.xls; *.xlsx)|*.xls; *.xlsx";
            if (Program.f1.Rus == true)
                ofd.Title = "Выберите документ для загрузки данных";
            if (Program.f1.Eng == true)
                ofd.Title = "Chose the document to load data";
            ExcelObj.Application app = new ExcelObj.Application();
            ExcelObj.Workbook workbook;
            ExcelObj.Worksheet NwSheet;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    workbook = app.Workbooks.Open(ofd.FileName, Missing.Value,
                      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                      Missing.Value);
                    NwSheet = (ExcelObj.Worksheet)workbook.Sheets.get_Item(1);
                    ////ShtRange = NwSheet.UsedRange;
                    var lastCell = NwSheet.Cells.SpecialCells(ExcelObj.XlCellType.xlCellTypeLastCell);
                    for (int Cnum = 1; Cnum <= lastCell.Column; Cnum++)
                    {
                        if (NwSheet.Cells[7, Cnum] != null)
                            crystalSize1[Cnum - 1] = Convert.ToDouble(NwSheet.Cells[7, Cnum].Text);
                        if (NwSheet.Cells[8, Cnum] != null)
                            crystalFreq1[Cnum - 1] = Convert.ToDouble(NwSheet.Cells[8, Cnum].Text);
                        NCSR1 = Cnum;
                        if (NwSheet.Cells[7, Cnum] == null)
                            goto M;
                    }
                    M: if (Program.f1.Rus == true)
                        MessageBox.Show("Данные успешно загружены!", "Сообщение!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (Program.f1.Eng == true)
                        MessageBox.Show("Data are imported successfully!!", "Message!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    DCSR1 = Convert.ToDouble(NwSheet.Cells[4, 16].Text);
                    DvCSR1 = Convert.ToDouble(NwSheet.Cells[4, 11].Text);
                    CSR1 = true;
                    app.Quit();
                }
                catch
                {
                    if (Program.f1.Rus == true)
                        MessageBox.Show("Невозможно открыть выбранный файл", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (Program.f1.Eng == true)
                        MessageBox.Show("It is impossible to open the selected file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    app.Quit();
                }

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            groupBox3.Visible = false;
        }

        private void максЗначениеОсиXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Part == true)
            {
                groupBox3.BringToFront();
                groupBox3.Visible = true;
                trackBar1.Maximum = 200 * 10;
                trackBar1.Value = Convert.ToInt32(X_max * 10);
                traB1 = trackBar1.Value;
                traB1 = traB1 / 10;
                textBox19.Text = Math.Round(traB1, 1).ToString("0.0");

            }
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            chart2.ChartAreas[0].AxisX.Maximum = (trackBar1.Value) / 10;
            traB1 = trackBar1.Value;
            traB1 = traB1 / 10;
            textBox19.Text = Math.Round(traB1, 1).ToString("0.0");
            if ((traB1 - 1) == X_max)
            {
                dataGridView1.Rows.Add();
                N = N + 1;
                dataGridView1.Rows[N - 1].Cells[0].Value = X_max;
                dataGridView1.Rows[N - 1].Cells[2].Value = 0;
                x = Convert.ToDouble(dataGridView1.Rows[N - 1].Cells[0].Value);
                y = Convert.ToDouble(dataGridView1.Rows[N - 1].Cells[2].Value);
                chart2.Series[0].Points.AddXY(x, y);
                chart2.Series[0].Points[N - 1].AxisLabel = (dataGridView1.Rows[N - 1].Cells[0].Value).ToString();
                X_max = X_max + X_interval;
            }

        }










        private void toolStripComboBox4_MouseMove(object sender, MouseEventArgs e)
        {

            toolStripComboBox4.ToolTipText = Program.f1.font;
        }

        private void toolStripButton1_MouseMove(object sender, MouseEventArgs e)
        {
            toolStripButton1.ToolTipText = Program.f1.Grid;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            ++click_toolStr_button1;
            switch (click_toolStr_button1)
            {
                case 1:
                    toolStripButton1.BackColor = SystemColors.Window;
                    toolStripButton1.Checked = true;
                    chart2.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
                    chart2.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
                    break;
                case 2:
                    toolStripButton1.BackColor = SystemColors.Control;
                    toolStripButton1.Checked = false;
                    chart2.ChartAreas[0].AxisX.MajorGrid.Enabled = true;
                    chart2.ChartAreas[0].AxisY.MajorGrid.Enabled = true;
                    click_toolStr_button1 = 0;
                    break;
            }
        }

        private void toolStripButton2_MouseMove(object sender, MouseEventArgs e)
        {
            toolStripButton2.ToolTipText = Program.f1.Legend;
        }

        private void toolStripSplitButton1_MouseMove(object sender, MouseEventArgs e)
        {
            toolStripSplitButton1.ToolTipText = Program.f1.Rotate;
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            chart2.ChartAreas[0].AxisX.LabelStyle.Angle = 0;
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            chart2.ChartAreas[0].AxisX.LabelStyle.Angle = -45;
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            chart2.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
        }

        private void toolStripDropDownButton1_MouseMove(object sender, MouseEventArgs e)
        {
            toolStripDropDownButton1.ToolTipText = Program.f1.Line_thickness;
        }

        private void toolStripSplitButton2_MouseMove(object sender, MouseEventArgs e)
        {
            toolStripSplitButton2.ToolTipText = Program.f1.Approx;
        }

        private void распределение1ToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            ++clickRaspred1;
            switch (clickRaspred1)
            {
                case 1:
                    chart2.Series[1].Enabled = false;
                    break;
                case 2:
                    chart2.Series[1].Enabled = true;
                    clickRaspred1 = 0;
                    break;
            }
        }

        private void распределение2ToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            ++clickRaspred2;
            switch (clickRaspred2)
            {
                case 1:
                    chart2.Series[2].Enabled = false;
                    break;
                case 2:
                    chart2.Series[2].Enabled = true;
                    clickRaspred2 = 0;
                    break;
            }
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex==0)
            {
                dataGridView1.Rows[e.RowIndex].Cells[3].Value = Math.Round(Math.Log(Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[0].Value)),2);

                dataGridRefresh();

            }

            if (e.ColumnIndex == 1)
            {
                Sum = 0;
                N = Convert.ToInt32(dataGridView1.RowCount);
                for (int i = 0; i < N; i++)
                    dataGridView1.Rows[i].Cells[2].Value = string.Empty;

                while (Convert.ToString(dataGridView1.Rows[N - 1].Cells[1].Value) == string.Empty)
                    N = N - 1;
                for (int i = 0; i < N; i++)
                {
                    Sum = Sum + Convert.ToDouble(dataGridView1[1, i].Value); // сначал номер стобца, затем номер строки
                }


                for (int i = 0; i < N; i++)
                {
                    V = Convert.ToDouble(dataGridView1[1, i].Value) / Sum * 100;
                    dataGridView1.Rows[i].Cells[2].Value = Convert.ToString(Math.Round(V, 2));
                }
                dataGridView1.Refresh();

                dataGridRefresh();
            }
        }

        private void dataGridRefresh()
        {
            if (open_f11)
            {
                if (analyt_func)
                {
                    button2_Click(this, new EventArgs());
                    if (LogNorm)
                        Log_Norm_10();
                    if (Norm)
                        Norm_10();
                }
                else
                    button2_Click(this, new EventArgs());
            }
            if (tabControl1.TabPages.ContainsKey("tabPage2"))
            {
                dataGridView2.Rows.Clear();
                sum_Vol = 0;
                N = Convert.ToInt32(dataGridView1.RowCount);
                N30 = Convert.ToInt32(dataGridView2.RowCount);
                if (N30 < N)
                    dataGridView2.Rows.Add(N - N30);
                while (dataGridView1.Rows[N - 1].Cells[0].Value == null)// || dataGridView1.Rows[N - 1].Cells[1].Value == null)
                    N = N - 1;

                if (Sum == 0)
                {
                    for (int i = 0; i < N; i++)
                    {
                        Sum = Sum + Convert.ToDouble(dataGridView1[1, i].Value); // сначал номер стобца, затем номер строки
                    }
                }

                for (int i = 0; i < N; i++)
                {
                    dataGridView2.Rows[i].Cells[0].Value = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                    Vol.Insert(i, Convert.ToDouble(dataGridView1[0, i].Value) * Convert.ToDouble(dataGridView1[1, i].Value));
                    dataGridView2.Rows[i].Cells[1].Value = Vol[i];
                    dataGridView2.Rows[i].Cells[3].Value = Math.Round(Math.Log(Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)), 2);
                    sum_Vol = sum_Vol + Vol[i];
                }

                for (int i = 0; i < N; i++)
                    dataGridView2.Rows[i].Cells[2].Value = Math.Round(Vol[i] / sum_Vol * 100, 2);

                if (open_f12)
                {
                    if (analyt_func_30)
                    {
                        button6_Click(this, new EventArgs());
                        if (LogNorm_30)
                            Log_Norm_30();
                        if (Norm_30)
                            Norm30();
                    }
                    else
                        button6_Click(this, new EventArgs());
                }
            }

            if (tabControl1.TabPages.ContainsKey("tabPage3"))
            {
                dataGridView3.Rows.Clear();
                sum_Vol = 0;
                N = Convert.ToInt32(dataGridView1.RowCount);
                N32 = Convert.ToInt32(dataGridView3.RowCount);
                if (N32 < N)
                    dataGridView3.Rows.Add(N - N32);
                while (dataGridView1.Rows[N - 1].Cells[0].Value == null)// || dataGridView1.Rows[N - 1].Cells[1].Value == null)
                    N = N - 1;

                if (Sum == 0)
                {
                    for (int i = 0; i < N; i++)
                    {
                        Sum = Sum + Convert.ToDouble(dataGridView1[1, i].Value); // сначал номер стобца, затем номер строки
                    }
                }

                for (int i = 0; i < N; i++)
                {
                    dataGridView3.Rows[i].Cells[0].Value = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                    Vol.Insert(i, Math.Pow(Convert.ToDouble(dataGridView1[0, i].Value), 2) * Convert.ToDouble(dataGridView1[1, i].Value));
                    dataGridView3.Rows[i].Cells[1].Value = Vol[i];
                    dataGridView3.Rows[i].Cells[3].Value = Math.Round(Math.Log(Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value)), 2);
                    sum_Vol = sum_Vol + Vol[i];
                }

                for (int i = 0; i < N; i++)
                    dataGridView3.Rows[i].Cells[2].Value = Math.Round(Vol[i] / sum_Vol * 100, 2);
                if (open_f13)
                {
                    if (analyt_func_32)
                    {
                        button7_Click(this, new EventArgs());
                        if (LogNorm_32)
                            Log_Norm_32();
                        if (Norm_32)
                            Norm32();
                    }
                    else
                        button7_Click(this, new EventArgs());
                }
            }

            if (tabControl1.TabPages.ContainsKey("tabPage4"))
            {
                dataGridView4.Rows.Clear();
                sum_Vol = 0;
                N = Convert.ToInt32(dataGridView1.RowCount);
                N43 = Convert.ToInt32(dataGridView4.RowCount);
                if (N43 < N)
                    dataGridView4.Rows.Add(N - N43);
                while (dataGridView1.Rows[N - 1].Cells[0].Value == null)// || dataGridView1.Rows[N - 1].Cells[1].Value == null)
                    N = N - 1;

                if (Sum == 0)
                {
                    for (int i = 0; i < N; i++)
                    {
                        Sum = Sum + Convert.ToDouble(dataGridView1[1, i].Value); // сначал номер стобца, затем номер строки
                    }
                }

                for (int i = 0; i < N; i++)
                {
                    dataGridView4.Rows[i].Cells[0].Value = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                    Vol.Insert(i, Math.Pow(Convert.ToDouble(dataGridView1[0, i].Value), 3) * Convert.ToDouble(dataGridView1[1, i].Value));
                    dataGridView4.Rows[i].Cells[1].Value = Vol[i];
                    dataGridView4.Rows[i].Cells[3].Value = Math.Round(Math.Log(Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value)), 2);
                    sum_Vol = sum_Vol + Vol[i];
                }

                for (int i = 0; i < N; i++)
                    dataGridView4.Rows[i].Cells[2].Value = Math.Round(Vol[i] / sum_Vol * 100, 2);

                if (open_f14)
                {
                    if (analyt_func_43)
                    {
                        button8_Click(this, new EventArgs());
                        if (LogNorm_43)
                            Log_Norm_43();
                        if (Norm_43)
                            Norm43();
                    }
                    else
                        button8_Click(this, new EventArgs());
                }
            }

            if (tabControl1.TabPages.ContainsKey("tabPage13"))
            {
                dataGridView5.Rows.Clear();
                sum_Vol = 0;
                N = Convert.ToInt32(dataGridView1.RowCount);
                N20 = Convert.ToInt32(dataGridView5.RowCount);
                if (N20 < N)
                    dataGridView5.Rows.Add(N - N20);
                while (dataGridView1.Rows[N - 1].Cells[0].Value == null)// || dataGridView1.Rows[N - 1].Cells[1].Value == null)
                    N = N - 1;

                if (Sum == 0)
                {
                    for (int i = 0; i < N; i++)
                    {
                        Sum = Sum + Convert.ToDouble(dataGridView1[1, i].Value); // сначал номер стобца, затем номер строки
                    }
                }

                for (int i = 0; i < N; i++)
                {
                    dataGridView5.Rows[i].Cells[0].Value = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                    Vol.Insert(i, Math.Sqrt(Convert.ToDouble(dataGridView1[0, i].Value)) * Convert.ToDouble(dataGridView1[1, i].Value));
                    dataGridView5.Rows[i].Cells[1].Value = Vol[i];
                    dataGridView5.Rows[i].Cells[3].Value = Math.Round(Math.Log(Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value)), 2);
                    sum_Vol = sum_Vol + Vol[i];
                }

                for (int i = 0; i < N; i++)
                    dataGridView5.Rows[i].Cells[2].Value = Math.Round(Vol[i] / sum_Vol * 100, 2);

                if (open_f15)
                {
                    if (analyt_func_20)
                    {
                        button1_Click(this, new EventArgs());
                        if (LogNorm_20)
                            Log_Norm_20();
                        if (Norm_20)
                            Norm20();
                    }
                    else
                        button1_Click(this, new EventArgs());
                }
            }
        }

        private void textBox4_Enter(object sender, EventArgs e)
        {
            label5.Select();
        }

        private void textBox3_Enter(object sender, EventArgs e)
        {
            label5.Select();
        }

        private void textBox5_Enter(object sender, EventArgs e)
        {
            label5.Select();
        }

        private void pictureBox2_MouseHover(object sender, EventArgs e)
        {
            t.SetToolTip(pictureBox2, Dcp10);
        }

        private void label19_MouseHover(object sender, EventArgs e)
        {
            t.SetToolTip(label19, median);
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && (e.KeyChar != (char)Keys.Back) && !(e.KeyChar == ','))
            {
                e.Handled = true;
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && (e.KeyChar != (char)Keys.Back) && !(e.KeyChar == ','))
            {
                e.Handled = true;
            }
        }

        private void textBox9_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && (e.KeyChar != (char)Keys.Back) && !(e.KeyChar == ','))
            {
                e.Handled = true;
            }
        }

        private void trackBar8_MouseUp(object sender, MouseEventArgs e)
        {
            MD = false;                       
            trackBar8.Value = 0;
        }

        private void textBox34_Enter(object sender, EventArgs e)
        {
            label33.Select();
        }

        private void textBox30_Enter(object sender, EventArgs e)
        {
            label33.Select();
        }

        private void textBox32_Enter(object sender, EventArgs e)
        {
            label33.Select();
        }

        private void textBox31_Enter(object sender, EventArgs e)
        {
            label33.Select();
        }

        private void trackBar10_ValueChanged(object sender, EventArgs e)
        {
            if (!MD && !KD)
            {
                trackBar10.Value = 0;
                par_a_30 = par_a0_30;
            }
            par_a0_30 = par_a_30 + Convert.ToDouble(trackBar10.Value) * 0.001;
            textBox26.Text = Math.Round(par_a0_30, 3).ToString("0.000");
            Program.f12.chart2.Series[1].Points.Clear();
            if (LogNorm_30)
            {
                for (double i = 0.1; i < X_max30 + r; i = i + 0.1)
                {
                    fDi = r / (par_b_30 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a0_30) * (Math.Log(i) - par_a0_30) / (2 * par_b_30 * par_b_30)) * Area_30;
                    Program.f12.chart2.Series[1].Points.AddXY(i, fDi);
                }
                D30 = Math.Exp(par_a0_30) * Math.Exp(0.5 * par_b_30 * par_b_30);
                Mo = Math.Exp(par_a0_30 - par_b_30 * par_b_30);
                m = Math.Exp(par_a0_30);
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N30; i++)
                {
                    if (!Ln30)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a0_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a0_30) / (2 * par_b_30 * par_b_30)) * Area_30)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a0_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a0_30) / (2 * par_b_30 * par_b_30)) * Area_30)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox31.Text = Math.Round(R2, 4).ToString("0.0000");
                textBox30.Text = Math.Round(par_a0_30, 3).ToString("0.000");
                textBox27.Text = Math.Round(D30, 1).ToString("0.0");
                textBox28.Text = Math.Round(m, 1).ToString("0.0");
                textBox29.Text = Math.Round(Mo, 1).ToString("0.0");
            }
            if (Norm_30)
            {
                for (double i = 0.1; i < X_max30 + r; i = i + 0.1)
                {
                    fDi = r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a0_30) * (i - par_a0_30) / (2 * par_b_30 * par_b_30)) * Area_30;
                    Program.f12.chart2.Series[1].Points.AddXY(i, fDi);
                }
                D30 = par_a0_30;
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N30; i++)
                {
                    if (!Ln30)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) - par_a0_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) - par_a0_30) / (2 * par_b_30 * par_b_30)) * Area_30)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a0_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a0_30) / (2 * par_b_30 * par_b_30)) * Area_30)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox31.Text = Math.Round(R2, 4).ToString("0.0000");
                textBox30.Text = Math.Round(par_a0_30, 3).ToString("0.000");
                textBox28.Text = Math.Round(Math.Exp(D30), 1).ToString("0.0");
            }
        }

        private void trackBar10_MouseUp(object sender, MouseEventArgs e)
        {
            MD = false;
            trackBar10.Value = 0;
        }

        private void trackBar10_MouseDown(object sender, MouseEventArgs e)
        {
            MD = true;
        }

        private void trackBar11_ValueChanged(object sender, EventArgs e)
        {
            if (!MD && !KD)
            {
                trackBar11.Value = 0;
                par_b_30 = par_b0_30;
            }
            par_b0_30 = par_b_30 + Convert.ToDouble(trackBar11.Value) * 0.001;
            textBox25.Text = Math.Round(par_b0_30, 3).ToString("0.000");
            Program.f12.chart2.Series[1].Points.Clear();
            if (LogNorm_30)
            {
                for (double i = 0.1; i < X_max30 + r; i = i + 0.1)
                {
                    fDi = r / (par_b0_30 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_30) * (Math.Log(i) - par_a_30) / (2 * par_b0_30 * par_b0_30)) * Area_30;
                    Program.f12.chart2.Series[1].Points.AddXY(i, fDi);
                }
                D30 = Math.Exp(par_a_30) * Math.Exp(0.5 * par_b0_30 * par_b0_30);
                Mo = Math.Exp(par_a_30 - par_b0_30 * par_b0_30);

                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N30; i++)
                {
                    if (!Ln30)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b0_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) / (2 * par_b0_30 * par_b0_30)) * Area_30)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b0_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) / (2 * par_b0_30 * par_b0_30)) * Area_30)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox31.Text = Math.Round(R2, 4).ToString("0.0000");
                textBox34.Text = Math.Round(par_b0_30, 3).ToString("0.000");
                textBox32.Text = Math.Round(par_b0_30, 3).ToString("0.000");
                textBox27.Text = Math.Round(D30, 1).ToString("0.0");

                textBox29.Text = Math.Round(Mo, 1).ToString("0.0");
            }
            if (Norm_30)
            {
                for (double i = 0.1; i < X_max30 + r; i = i + 0.1)
                {
                    fDi = r / (par_b0_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_30) * (i - par_a_30) / (2 * par_b0_30 * par_b0_30)) * Area_30;
                    Program.f12.chart2.Series[1].Points.AddXY(i, fDi);
                }

                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N30; i++)
                {
                    if (!Ln30)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b0_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) - par_a_30) / (2 * par_b0_30 * par_b0_30)) * Area_30)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b0_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) / (2 * par_b0_30 * par_b0_30)) * Area_30)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox31.Text = Math.Round(R2, 4).ToString("0.0000");
                textBox34.Text = Math.Round(par_b0_30, 3).ToString("0.000");
                textBox32.Text = Math.Round(par_b0_30, 3).ToString("0.000");
                
            }
        }

        private void trackBar11_MouseUp(object sender, MouseEventArgs e)
        {
            MD = false;
            trackBar11.Value = 0;
        }

        private void trackBar11_MouseDown(object sender, MouseEventArgs e)
        {
            MD = true;
        }

        private void trackBar11_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = false;
                trackBar11.Value = 0;
            }
        }

        private void trackBar9_ValueChanged(object sender, EventArgs e)
        {
            if (!MD && !KD)
            {
                Area_30 = Area0_30;
                trackBar9.Value = 0;
            }
            Area0_30 = Area_30 + Convert.ToDouble(trackBar9.Value) / 10;
            Program.f12.chart2.Series[1].Points.Clear();
            if (LogNorm_30)
            {
                for (double i = 0.1; i < X_max30 + r; i = i + 0.1)
                {
                    fDi = r / (par_b_30 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_30) * (Math.Log(i) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30;
                    Program.f12.chart2.Series[1].Points.AddXY(i, fDi);
                }

                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N30; i++)
                {
                    if (!Ln30)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox24.Text = Math.Round(Area0_30, 1).ToString();
                textBox31.Text = Math.Round(R2, 4).ToString("0.0000");

            }
            if (Norm_30)
            {
                for (double i = 0.1; i < X_max30 + r; i = i + 0.1)
                {
                    fDi = r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_30) * (i - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30;
                    Program.f12.chart2.Series[1].Points.AddXY(i, fDi);
                }

                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N30; i++)
                {
                    if (!Ln30)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox24.Text = Math.Round(Area0_30, 1).ToString();
                textBox31.Text = Math.Round(R2, 4).ToString("0.0000");

            }
        }

        private void trackBar9_MouseUp(object sender, MouseEventArgs e)
        {
            MD = false;
            trackBar9.Value = 0;
        }

        private void trackBar13_ValueChanged(object sender, EventArgs e)
        {
            if (!MD && !KD)
            {
                trackBar13.Value = 0;
                par_a_32 = par_a0_32;
            }
            par_a0_32 = par_a_32 + Convert.ToDouble(trackBar13.Value) * 0.001;
            textBox39.Text = Math.Round(par_a0_32, 3).ToString("0.000");
            Program.f13.chart2.Series[1].Points.Clear();
            if (LogNorm_32)
            {
                for (double i = 0.1; i < X_max32 + r; i = i + 0.1)
                {
                    fDi = r / (par_b_32 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a0_32) * (Math.Log(i) - par_a0_32) / (2 * par_b_32 * par_b_32)) * Area_32;
                    Program.f13.chart2.Series[1].Points.AddXY(i, fDi);
                }
                D32 = Math.Exp(par_a0_32) * Math.Exp(0.5 * par_b_32 * par_b_32);
                Mo = Math.Exp(par_a0_32 - par_b_32 * par_b_32);
                m = Math.Exp(par_a0_32);
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N32; i++)
                {
                    if (!Ln32)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a0_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a0_32) / (2 * par_b_32 * par_b_32)) * Area_32)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a0_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a0_32) / (2 * par_b_32 * par_b_32)) * Area_32)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox44.Text = Math.Round(R2, 4).ToString("0.0000");
                textBox43.Text = Math.Round(par_a0_32, 3).ToString("0.000");
                textBox40.Text = Math.Round(D32, 1).ToString("0.0");
                textBox41.Text = Math.Round(m, 1).ToString("0.0");
                textBox42.Text = Math.Round(Mo, 1).ToString("0.0");
            }
            if (Norm_32)
            {
                for (double i = 0.1; i < X_max32 + r; i = i + 0.1)
                {
                    fDi = r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a0_32) * (i - par_a0_32) / (2 * par_b_32 * par_b_32)) * Area_32;
                    Program.f13.chart2.Series[1].Points.AddXY(i, fDi);
                }
                D32 = par_a0_32;
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N32; i++)
                {
                    //if (!Ln30)
                      //  SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) - par_a0_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) - par_a0_30) / (2 * par_b_30 * par_b_30)) * Area_30)), 2);
                    //else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a0_32) * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a0_32) / (2 * par_b_32 * par_b_32)) * Area_32)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox44.Text = Math.Round(R2, 4).ToString("0.0000");
                textBox43.Text = Math.Round(par_a0_32, 3).ToString("0.000");
                textBox41.Text = Math.Round(Math.Exp(D32), 1).ToString("0.0");
            }
        }

        private void Form3_Activated(object sender, EventArgs e)
        {
            MouseUp_X = Cursor.Position.X;
            MouseUp_Y = Cursor.Position.Y;
            Form3_Click(sender, e);
        }

        private void Form3_Click(object sender, EventArgs e)
        {
            if (!tS1 && !tB)
            {
                // нажимам левую кнопку
                mouse_event((int)(MouseEventFlags.LEFTDOWN), MouseUp_X, MouseUp_Y, 0, 0);
                // отпускаем левую кнопку
                mouse_event((int)(MouseEventFlags.LEFTUP), MouseUp_X, MouseUp_Y, 0, 0);
            }
            if (tS1 || tB)
                // нажимам левую кнопку
                mouse_event((int)(MouseEventFlags.LEFTDOWN), MouseUp_X, MouseUp_Y, 0, 0);
        }

        private void toolStrip1_MouseEnter(object sender, EventArgs e)
        {
            tS1 = true;      
        }

        private void toolStrip1_MouseLeave(object sender, EventArgs e)
        {
            tS1 = false;
        }

        private void взвешенноеПоОбъемуРаспределениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Convert.ToString(dataGridView1.Rows[0].Cells[0].Value) == string.Empty)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Размеры частиц не заданы! Попробуйте снова!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles size are not determined! Try again!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (Convert.ToString(dataGridView1.Rows[0].Cells[1].Value) == string.Empty)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Число частиц не задано! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles number are not determined! Try again!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            flag_volweight_raspr = true;
            //CSR = false;
            //chart2.Series[3].Points.Clear();
            //chart2.Series[4].Points.Clear();
            if (!tabControl1.TabPages.ContainsKey("tabPage4"))
            {
                tabControl1.TabPages.Add(tabPage4);

            }
            tabControl1.SelectedTab = tabControl1.TabPages["tabPage4"];

            dataGridView4.Rows.Clear();
            sum_Vol = 0;
            N = Convert.ToInt32(dataGridView1.RowCount);
            N43 = Convert.ToInt32(dataGridView4.RowCount);
            if (N43 < N)
                dataGridView4.Rows.Add(N - N43);
            while (dataGridView1.Rows[N - 1].Cells[0].Value == null)// || dataGridView1.Rows[N - 1].Cells[1].Value == null)
                N = N - 1;

            if (Sum == 0)
            {
                for (int i = 0; i < N; i++)
                {
                    Sum = Sum + Convert.ToDouble(dataGridView1[1, i].Value); // сначал номер стобца, затем номер строки
                }
            }

            for (int i = 0; i < N; i++)
            {
                dataGridView4.Rows[i].Cells[0].Value = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                Vol.Insert(i, Math.Pow(Convert.ToDouble(dataGridView1[0, i].Value), 3) * Convert.ToDouble(dataGridView1[1, i].Value));
                dataGridView4.Rows[i].Cells[1].Value = Vol[i];
                dataGridView4.Rows[i].Cells[3].Value = Math.Round(Math.Log(Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value)), 2);
                sum_Vol = sum_Vol + Vol[i];
            }

            for (int i = 0; i < N; i++)
                dataGridView4.Rows[i].Cells[2].Value = Math.Round(Vol[i] / sum_Vol * 100, 2);

            button8_Click(sender, e);


            //if (CSR1 == true)
           // {
             //   textBox15.Visible = false;
               // textBox16.Visible = false;
              //  pictureBox7.Visible = false;
              //  pictureBox8.Visible = false;
              //  label12.Visible = false;
               // label13.Visible = false;
           // }
           // if (CSR2 == true)
           // {
             //   textBox17.Visible = false;
              //  textBox18.Visible = false;
              //  pictureBox9.Visible = false;
              //  pictureBox10.Visible = false;
              //  label14.Visible = false;
               // label15.Visible = false;
           // }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (Convert.ToString(dataGridView4.Rows[0].Cells[0].Value) == string.Empty)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Размеры частиц не заданы! Попробуйте снова!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles size are not determined! Try again!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (Convert.ToString(dataGridView4.Rows[0].Cells[2].Value) == string.Empty)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Доля частиц не задана! Попробуйте снова!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles fraction are not determined! Try again!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            tabControl5.Visible = false;
            if (!open_f14)
            {
                f14 = new Form14();
                f14.Owner = this;
                open_f14 = true;
                f14.Show();
            }
            Program.f14.chart2.Series[0].Points.Clear();
            Program.f14.chart2.Series[1].Points.Clear();
            Program.f14.chart2.Series[2].Points.Clear();
            Program.f14.chart2.Series[5].Points.Clear();
            analyt_func_43 = false;

            N43 = Convert.ToInt32(dataGridView4.RowCount);
            while (Convert.ToString(dataGridView4.Rows[N43 - 1].Cells[0].Value) == string.Empty || Convert.ToString(dataGridView4.Rows[N43 - 1].Cells[2].Value) == string.Empty)
                N43 = N43 - 1;

            if (Sum == 0)
            {
                for (int i = 0; i < N43; i++)
                {
                    Sum = Sum + Convert.ToDouble(dataGridView4[1, i].Value); // сначал номер стобца, затем номер строки
                }
            }

            if (Ln43)
            {
                for (int i = 0; i < N43; i++)
                {
                    x = Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value);
                    y = Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value);
                    Program.f14.chart2.Series[0].Points.AddXY(x, y);
                    X_max43 = Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value);

                }
                Program.f14.r = Convert.ToDouble(dataGridView4.Rows[N43 - 1].Cells[3].Value) - Convert.ToDouble(dataGridView4.Rows[N43 - 2].Cells[3].Value);
                X_max43 = X_max43 + Program.f14.r;
                Program.f14.chart2.ChartAreas[0].AxisY.Minimum = 0;
                Program.f14.chart2.ChartAreas[0].AxisX.Maximum = X_max43;
                Program.f14.chart2.ChartAreas[0].AxisX.Minimum = Convert.ToDouble(dataGridView4.Rows[0].Cells[3].Value) - Program.f14.r;
                Program.f14.chart2.ChartAreas[0].AxisX.Title = "Ln(D)";
                Program.f14.chart2.ChartAreas[0].AxisX.Interval = 0;
            }
            else
            {
                for (int i = 0; i < N43; i++)
                {
                    x = Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value);
                    y = Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value);
                    Program.f14.chart2.Series[0].Points.AddXY(x, y);
                    X_max43 = Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value);
                }
                Program.f14.chart2.Series[0].IsVisibleInLegend = true;
                Program.f14.chart2.Series[1].IsVisibleInLegend = false;
                Program.f14.chart2.Series[2].IsVisibleInLegend = false;
                Program.f14.chart2.Series[3].IsVisibleInLegend = false;
                Program.f14.chart2.Series[4].IsVisibleInLegend = false;
                Program.f14.chart2.Series[5].IsVisibleInLegend = false;
                if (Program.f1.Rus == true)
                    Program.f14.chart2.Series[0].Name = "Гист.";
                if (Program.f1.Eng == true)
                    Program.f14.chart2.Series[0].Name = "Hist.";
                Part43 = true;
                r = Convert.ToDouble(dataGridView4.Rows[1].Cells[0].Value) - Convert.ToDouble(dataGridView4.Rows[0].Cells[0].Value);
                X_max43 = X_max43 + r;
                Program.f14.chart2.ChartAreas[0].AxisY.Minimum = 0;
                Program.f14.chart2.ChartAreas[0].AxisX.Maximum = X_max43;
                if (Part43 == true && CSR43 == true)
                    Program.f14.chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), нм";
                else
                    Program.f14.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, нм";
                Program.f14.chart2.ChartAreas[0].AxisX.TitleFont = new Font("Times New Roman", Program.f14.I, FontStyle.Bold);
                if (Part43 == true && CSR43 == true)
                    Program.f14.chart2.ChartAreas[0].AxisY.Title = "Доля частиц (ОКР), %";
                else
                    Program.f14.chart2.ChartAreas[0].AxisY.Title = "Доля частиц, %";
                if (Program.f1.Eng == true)
                    Program.f14.chart2.ChartAreas[0].AxisY.Title = "Frequency, %";
                Program.f14.chart2.ChartAreas[0].AxisY.TitleFont = new Font("Times New Roman", Program.f14.I, FontStyle.Bold);
                Program.f14.chart2.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Times New Roman", Program.f14.I, System.Drawing.FontStyle.Regular);
                Program.f14.chart2.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Times New Roman", Program.f14.I, System.Drawing.FontStyle.Regular);
                if (toolStripComboBox2.SelectedIndex == 0)
                {
                    if (Program.f1.Rus == true)
                    {
                        if (Part43 == true && CSR43 == true)
                            Program.f14.chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), мм (нм)";
                        else
                            Program.f14.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, мм";
                        label61.Text = "мм";
                        label62.Text = "мм";
                        label64.Text = "мм";
                    }
                    if (Program.f1.Eng == true)
                    {
                        if (Part43 == true && CSR43 == true)
                            Program.f14.chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), mm (nm)";
                        else
                            Program.f14.chart2.ChartAreas[0].AxisX.Title = "Particle size, mm";
                        label61.Text = "mm";
                        label62.Text = "mm";
                        label64.Text = "mm";
                    }
                }
                if (toolStripComboBox2.SelectedIndex == 1)
                {
                    if (Program.f1.Rus == true)
                    {
                        if (Part43 == true && CSR43 == true)
                            Program.f14.chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), мкм (нм)";
                        else
                            Program.f14.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, мкм";
                        label61.Text = "мкм";
                        label62.Text = "мкм";
                        label64.Text = "мкм";
                    }
                    if (Program.f1.Eng == true)
                    {
                        if (Part43 == true && CSR43 == true)
                            Program.f14.chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), mkm (nm)";
                        else
                            Program.f14.chart2.ChartAreas[0].AxisX.Title = "Particle size, mkm";
                        label61.Text = "mkm";
                        label62.Text = "mkm";
                        label64.Text = "mkm";
                    }
                }
                if (toolStripComboBox2.SelectedIndex == 2)
                {
                    if (Program.f1.Rus == true)
                    {
                        if (Part43 == true && CSR43 == true)
                            Program.f14.chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), нм";
                        else
                            Program.f14.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, нм";
                        label61.Text = "нм";
                        label62.Text = "нм";
                        label64.Text = "нм";
                    }
                    if (Program.f1.Eng == true)
                    {
                        if (Part43 == true && CSR43 == true)
                            Program.f14.chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), nm";
                        else
                            Program.f14.chart2.ChartAreas[0].AxisX.Title = "Particle size, nm";
                        label61.Text = "nm";
                        label62.Text = "nm";
                        label64.Text = "nm";
                    }


                }


                if (dataGridView4.Rows.Count > 2)
                {
                    X_interval43 = Convert.ToDouble(dataGridView4.Rows[1].Cells[0].Value) - Convert.ToDouble(dataGridView4.Rows[0].Cells[0].Value);
                    X_min43 = Convert.ToDouble(dataGridView4.Rows[0].Cells[0].Value) - X_interval43;
                    Program.f14.chart2.ChartAreas[0].AxisX.Minimum = X_min43;
                    Program.f14.chart2.ChartAreas[0].AxisX.Interval = X_interval43;
                    if (X_min43 <= 0)
                    {
                        Program.f14.chart2.ChartAreas[0].AxisX.IsStartedFromZero = true;
                        for (int i = 0; i < N43; i++)
                            Program.f14.chart2.Series[0].Points[i].AxisLabel = (dataGridView4.Rows[i].Cells[0].Value).ToString();
                    }
                }
            }
        }

        private void dataGridView4_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                dataGridView4.Rows[e.RowIndex].Cells[3].Value = Math.Round(Math.Log(Convert.ToDouble(dataGridView4.Rows[e.RowIndex].Cells[0].Value)), 2);
            }
        }

        private void trackBar16_ValueChanged(object sender, EventArgs e)
        {
            if (!MD && !KD)
            {
                trackBar16.Value = 0;
                par_a_43 = par_a0_43;
            }
            par_a0_43 = par_a_43 + Convert.ToDouble(trackBar16.Value) * 0.001;
            textBox52.Text = Math.Round(par_a0_43, 3).ToString("0.000");
            Program.f14.chart2.Series[1].Points.Clear();
            if (LogNorm_43)
            {
                for (double i = 0.1; i < X_max43 + r; i = i + 0.1)
                {
                    fDi = r / (par_b_43 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a0_43) * (Math.Log(i) - par_a0_43) / (2 * par_b_43 * par_b_43)) * Area_43;
                    Program.f14.chart2.Series[1].Points.AddXY(i, fDi);
                }
                D43 = Math.Exp(par_a0_43) * Math.Exp(0.5 * par_b_43 * par_b_43);
                Mo = Math.Exp(par_a0_43 - par_b_43 * par_b_43);
                m = Math.Exp(par_a0_43);
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N43; i++)
                {
                    if (!Ln43)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a0_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a0_43) / (2 * par_b_43 * par_b_43)) * Area_43)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a0_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a0_43) / (2 * par_b_43 * par_b_43)) * Area_43)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox57.Text = Math.Round(R2, 4).ToString("0.0000");
                textBox56.Text = Math.Round(par_a0_43, 3).ToString("0.000");
                textBox53.Text = Math.Round(D43, 1).ToString("0.0");
                textBox54.Text = Math.Round(m, 1).ToString("0.0");
                textBox55.Text = Math.Round(Mo, 1).ToString("0.0");
            }
            if (Norm_43)
            {
                for (double i = 0.1; i < X_max43 + r; i = i + 0.1)
                {
                    fDi = r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a0_43) * (i - par_a0_43) / (2 * par_b_43 * par_b_43)) * Area_43;
                    Program.f14.chart2.Series[1].Points.AddXY(i, fDi);
                }
                D43 = par_a0_43;
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N43; i++)
                {
                    //if (!Ln30)
                    //  SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) - par_a0_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) - par_a0_30) / (2 * par_b_30 * par_b_30)) * Area_30)), 2);
                    //else
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a0_43) * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a0_43) / (2 * par_b_43 * par_b_43)) * Area_43)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox57.Text = Math.Round(R2, 4).ToString("0.0000");
                textBox56.Text = Math.Round(par_a0_43, 3).ToString("0.000");
                textBox54.Text = Math.Round(Math.Exp(D43), 1).ToString("0.0");
            }
        }

        private void trackBar16_MouseUp(object sender, MouseEventArgs e)
        {
            MD = false;
            trackBar16.Value = 0;
        }

        private void trackBar16_MouseDown(object sender, MouseEventArgs e)
        {
            MD = true;
        }

        private void trackBar16_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = true;
            }
        }

        private void trackBar16_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = false;
                trackBar16.Value = 0;
            }
        }

        private void trackBar17_ValueChanged(object sender, EventArgs e)
        {
            if (!MD && !KD)
            {
                trackBar17.Value = 0;
                par_b_43 = par_b0_43;
            }
            par_b0_43 = par_b_43 + Convert.ToDouble(trackBar17.Value) * 0.001;
            textBox51.Text = Math.Round(par_b0_43, 3).ToString("0.000");
            Program.f14.chart2.Series[1].Points.Clear();
            if (LogNorm_43)
            {
                for (double i = 0.1; i < X_max43 + r; i = i + 0.1)
                {
                    fDi = r / (par_b0_43 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_43) * (Math.Log(i) - par_a_43) / (2 * par_b0_43 * par_b0_43)) * Area_43;
                    Program.f14.chart2.Series[1].Points.AddXY(i, fDi);
                }
                D43 = Math.Exp(par_a_43) * Math.Exp(0.5 * par_b0_43 * par_b0_43);
                Mo = Math.Exp(par_a_43 - par_b0_43 * par_b0_43);

                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N43; i++)
                {
                    if (!Ln43)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b0_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) / (2 * par_b0_43 * par_b0_43)) * Area_43)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b0_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) / (2 * par_b0_43 * par_b0_43)) * Area_43)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox57.Text = Math.Round(R2, 4).ToString("0.0000");
                textBox60.Text = Math.Round(par_b0_43, 3).ToString("0.000");
                textBox58.Text = Math.Round(par_b0_43, 3).ToString("0.000");
                textBox53.Text = Math.Round(D43, 1).ToString("0.0");
                textBox55.Text = Math.Round(Mo, 1).ToString("0.0");
            }
            if (Norm_43)
            {
                for (double i = 0.1; i < X_max43 + r; i = i + 0.1)
                {
                    fDi = r / (par_b0_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_43) * (i - par_a_43) / (2 * par_b0_43 * par_b0_43)) * Area_43;
                    Program.f14.chart2.Series[1].Points.AddXY(i, fDi);
                }

                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N43; i++)
                {
                    // if (!Ln32)
                    //SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b0_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) - par_a_30) / (2 * par_b0_30 * par_b0_30)) * Area_30)), 2);
                    //else
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b0_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) / (2 * par_b0_43 * par_b0_43)) * Area_43)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox57.Text = Math.Round(R2, 4).ToString("0.0000");
                textBox60.Text = Math.Round(par_b0_43, 3).ToString("0.000");
                textBox58.Text = Math.Round(par_b0_43, 3).ToString("0.000");

            }
        }

        private void trackBar17_MouseUp(object sender, MouseEventArgs e)
        {
            MD = false;
            trackBar17.Value = 0;
        }

        private void trackBar17_MouseDown(object sender, MouseEventArgs e)
        {
            MD = true;
        }

        private void trackBar17_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = true;
            }
        }

        private void trackBar17_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = false;
                trackBar17.Value = 0;
            }
        }

        private void trackBar15_ValueChanged(object sender, EventArgs e)
        {
            if (!MD && !KD)
            {
                Area_43 = Area0_43;
                trackBar15.Value = 0;
            }
            Area0_43 = Area_43 + Convert.ToDouble(trackBar15.Value) / 10;
            Program.f14.chart2.Series[1].Points.Clear();
            if (LogNorm_43)
            {
                for (double i = 0.1; i < X_max43 + r; i = i + 0.1)
                {
                    fDi = r / (par_b_43 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_43) * (Math.Log(i) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43;
                    Program.f14.chart2.Series[1].Points.AddXY(i, fDi);
                }

                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N43; i++)
                {
                    if (!Ln43)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox50.Text = Math.Round(Area0_43, 1).ToString();
                textBox57.Text = Math.Round(R2, 4).ToString("0.0000");

            }
            if (Norm_43)
            {
                for (double i = 0.1; i < X_max43 + r; i = i + 0.1)
                {
                    fDi = r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_43) * (i - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43;
                    Program.f14.chart2.Series[1].Points.AddXY(i, fDi);
                }

                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N43; i++)
                {
                    if (!Ln43)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value) - par_a_43) * (Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox50.Text = Math.Round(Area0_43, 1).ToString();
                textBox57.Text = Math.Round(R2, 4).ToString("0.0000");

            }
        }

        private void trackBar3_MouseEnter(object sender, EventArgs e)
        {
            tB = true;
        }

        private void trackBar3_MouseLeave(object sender, EventArgs e)
        {
            tB = false;
        }

        private void trackBar8_MouseEnter(object sender, EventArgs e)
        {
            tB = true;
        }

        private void trackBar8_MouseLeave(object sender, EventArgs e)
        {
            tB = false;
        }

        private void trackBar7_MouseEnter(object sender, EventArgs e)
        {
            tB = true;
        }

        private void trackBar7_MouseLeave(object sender, EventArgs e)
        {
            tB = false;
        }

        private void trackBar6_MouseEnter(object sender, EventArgs e)
        {
            tB = true;
        }

        private void trackBar6_MouseLeave(object sender, EventArgs e)
        {
            tB = false;
        }

        private void trackBar10_MouseEnter(object sender, EventArgs e)
        {
            tB = true;
        }

        private void trackBar10_MouseLeave(object sender, EventArgs e)
        {
            tB = false;
        }

        private void trackBar11_MouseEnter(object sender, EventArgs e)
        {
            tB = true;
        }

        private void trackBar11_MouseLeave(object sender, EventArgs e)
        {
            tB = false;
        }

        private void trackBar9_MouseEnter(object sender, EventArgs e)
        {
            tB = true;
        }

        private void trackBar9_MouseLeave(object sender, EventArgs e)
        {
            tB = false;
        }

        private void trackBar13_MouseEnter(object sender, EventArgs e)
        {
            tB = true;
        }

        private void trackBar13_MouseLeave(object sender, EventArgs e)
        {
            tB = false;
        }

        private void trackBar14_MouseEnter(object sender, EventArgs e)
        {
            tB = true;
        }

        private void trackBar14_MouseLeave(object sender, EventArgs e)
        {
            tB = false;
        }

        private void trackBar12_MouseEnter(object sender, EventArgs e)
        {
            tB = true;
        }

        private void trackBar12_MouseLeave(object sender, EventArgs e)
        {
            tB = false;
        }

        private void trackBar16_MouseEnter(object sender, EventArgs e)
        {
            tB = true;
        }

        private void trackBar16_MouseLeave(object sender, EventArgs e)
        {
            tB = false;
        }

        private void trackBar17_MouseEnter(object sender, EventArgs e)
        {
            tB = true;
        }

        private void trackBar17_MouseLeave(object sender, EventArgs e)
        {
            tB = false;
        }

        private void trackBar15_MouseEnter(object sender, EventArgs e)
        {
            tB = true;
        }

        private void trackBar15_MouseLeave(object sender, EventArgs e)
        {
            tB = false;
        }

        private void поверхностноеРаспределениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Convert.ToString(dataGridView1.Rows[0].Cells[0].Value) == string.Empty)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Размеры частиц не заданы! Попробуйте снова!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles size are not determined! Try again!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (Convert.ToString(dataGridView1.Rows[0].Cells[1].Value) == string.Empty)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Число частиц не задано! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles number are not determined! Try again!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            flag_surf_raspr = true;
            //CSR = false;
            //chart2.Series[3].Points.Clear();
            //chart2.Series[4].Points.Clear();
            if (!tabControl1.TabPages.ContainsKey("tabPage13"))
            {
                tabControl1.TabPages.Add(tabPage13);

            }
            tabControl1.SelectedTab = tabControl1.TabPages["tabPage13"];

            dataGridView5.Rows.Clear();
            sum_Vol = 0;
            N = Convert.ToInt32(dataGridView1.RowCount);
            N20 = Convert.ToInt32(dataGridView5.RowCount);
            if (N20 < N)
                dataGridView5.Rows.Add(N - N20);
            while (dataGridView1.Rows[N - 1].Cells[0].Value == null)// || dataGridView1.Rows[N - 1].Cells[1].Value == null)
                N = N - 1;

            if (Sum == 0)
            {
                for (int i = 0; i < N; i++)
                {
                    Sum = Sum + Convert.ToDouble(dataGridView1[1, i].Value); // сначал номер стобца, затем номер строки
                }
            }

            for (int i = 0; i < N; i++)
            {
                dataGridView5.Rows[i].Cells[0].Value = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                Vol.Insert(i, Math.Sqrt(Convert.ToDouble(dataGridView1[0, i].Value)) * Convert.ToDouble(dataGridView1[1, i].Value));
                dataGridView5.Rows[i].Cells[1].Value = Vol[i];
                dataGridView5.Rows[i].Cells[3].Value = Math.Round(Math.Log(Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value)), 2);
                sum_Vol = sum_Vol + Vol[i];
            }

            for (int i = 0; i < N; i++)
                dataGridView5.Rows[i].Cells[2].Value = Math.Round(Vol[i] / sum_Vol * 100, 2);

            button1_Click(sender, e);


            //if (CSR1 == true)
            // {
            //   textBox15.Visible = false;
            // textBox16.Visible = false;
            //  pictureBox7.Visible = false;
            //  pictureBox8.Visible = false;
            //  label12.Visible = false;
            // label13.Visible = false;
            // }
            // if (CSR2 == true)
            // {
            //   textBox17.Visible = false;
            //  textBox18.Visible = false;
            //  pictureBox9.Visible = false;
            //  pictureBox10.Visible = false;
            //  label14.Visible = false;
            // label15.Visible = false;
            // }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Convert.ToString(dataGridView5.Rows[0].Cells[0].Value) == string.Empty)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Размеры частиц не заданы! Попробуйте снова!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles size are not determined! Try again!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (Convert.ToString(dataGridView5.Rows[0].Cells[2].Value) == string.Empty)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Доля частиц не задана! Попробуйте снова!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles fraction are not determined! Try again!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            tabControl6.Visible = false;
            if (!open_f15)
            {
                f15 = new Form15();
                f15.Owner = this;
                open_f15 = true;
                f15.Show();
            }
            Program.f15.chart2.Series[0].Points.Clear();
            Program.f15.chart2.Series[1].Points.Clear();
            Program.f15.chart2.Series[2].Points.Clear();
            Program.f15.chart2.Series[5].Points.Clear();
            analyt_func_20 = false;

            N20 = Convert.ToInt32(dataGridView5.RowCount);
            while (Convert.ToString(dataGridView5.Rows[N20 - 1].Cells[0].Value) == string.Empty || Convert.ToString(dataGridView5.Rows[N20 - 1].Cells[2].Value) == string.Empty)
                N20 = N20 - 1;

            if (Sum == 0)
            {
                for (int i = 0; i < N20; i++)
                {
                    Sum = Sum + Convert.ToDouble(dataGridView4[1, i].Value); // сначал номер стобца, затем номер строки
                }
            }

            if (Ln20)
            {
                for (int i = 0; i < N20; i++)
                {
                    x = Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value);
                    y = Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value);
                    Program.f15.chart2.Series[0].Points.AddXY(x, y);
                    X_max20 = Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value);

                }
                Program.f15.r = Convert.ToDouble(dataGridView5.Rows[N20 - 1].Cells[3].Value) - Convert.ToDouble(dataGridView5.Rows[N20 - 2].Cells[3].Value);
                X_max20 = X_max20 + Program.f15.r;
                Program.f15.chart2.ChartAreas[0].AxisY.Minimum = 0;
                Program.f15.chart2.ChartAreas[0].AxisX.Maximum = X_max20;
                Program.f15.chart2.ChartAreas[0].AxisX.Minimum = Convert.ToDouble(dataGridView5.Rows[0].Cells[3].Value) - Program.f15.r;
                Program.f15.chart2.ChartAreas[0].AxisX.Title = "Ln(D)";
                Program.f15.chart2.ChartAreas[0].AxisX.Interval = 0;
            }
            else
            {
                for (int i = 0; i < N20; i++)
                {
                    x = Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value);
                    y = Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value);
                    Program.f15.chart2.Series[0].Points.AddXY(x, y);
                    X_max20 = Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value);
                }
                Program.f15.chart2.Series[0].IsVisibleInLegend = true;
                Program.f15.chart2.Series[1].IsVisibleInLegend = false;
                Program.f15.chart2.Series[2].IsVisibleInLegend = false;
                Program.f15.chart2.Series[3].IsVisibleInLegend = false;
                Program.f15.chart2.Series[4].IsVisibleInLegend = false;
                Program.f15.chart2.Series[5].IsVisibleInLegend = false;
                if (Program.f1.Rus == true)
                    Program.f15.chart2.Series[0].Name = "Гист.";
                if (Program.f1.Eng == true)
                    Program.f15.chart2.Series[0].Name = "Hist.";
                Part20 = true;
                r = Convert.ToDouble(dataGridView5.Rows[1].Cells[0].Value) - Convert.ToDouble(dataGridView5.Rows[0].Cells[0].Value);
                X_max20 = X_max20 + r;
                Program.f15.chart2.ChartAreas[0].AxisY.Minimum = 0;
                Program.f15.chart2.ChartAreas[0].AxisX.Maximum = X_max20;
                if (Part20 == true && CSR20 == true)
                    Program.f15.chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), нм";
                else
                    Program.f15.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, нм";
                Program.f15.chart2.ChartAreas[0].AxisX.TitleFont = new Font("Times New Roman", Program.f15.I, FontStyle.Bold);
                if (Part20 == true && CSR20 == true)
                    Program.f15.chart2.ChartAreas[0].AxisY.Title = "Доля частиц (ОКР), %";
                else
                    Program.f15.chart2.ChartAreas[0].AxisY.Title = "Доля частиц, %";
                if (Program.f1.Eng == true)
                    Program.f15.chart2.ChartAreas[0].AxisY.Title = "Frequency, %";
                Program.f15.chart2.ChartAreas[0].AxisY.TitleFont = new Font("Times New Roman", Program.f15.I, FontStyle.Bold);
                Program.f15.chart2.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Times New Roman", Program.f15.I, System.Drawing.FontStyle.Regular);
                Program.f15.chart2.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Times New Roman", Program.f15.I, System.Drawing.FontStyle.Regular);
                if (toolStripComboBox2.SelectedIndex == 0)
                {
                    if (Program.f1.Rus == true)
                    {
                        if (Part20 == true && CSR20 == true)
                            Program.f15.chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), мм (нм)";
                        else
                            Program.f15.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, мм";
                        label70.Text = "мм";
                        label71.Text = "мм";
                        label73.Text = "мм";
                    }
                    if (Program.f1.Eng == true)
                    {
                        if (Part20 == true && CSR20 == true)
                            Program.f15.chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), mm (nm)";
                        else
                            Program.f15.chart2.ChartAreas[0].AxisX.Title = "Particle size, mm";
                        label70.Text = "mm";
                        label71.Text = "mm";
                        label73.Text = "mm";
                    }
                }
                if (toolStripComboBox2.SelectedIndex == 1)
                {
                    if (Program.f1.Rus == true)
                    {
                        if (Part20 == true && CSR20 == true)
                            Program.f15.chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), мкм (нм)";
                        else
                            Program.f15.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, мкм";
                        label70.Text = "мкм";
                        label71.Text = "мкм";
                        label73.Text = "мкм";
                    }
                    if (Program.f1.Eng == true)
                    {
                        if (Part20 == true && CSR20 == true)
                            Program.f15.chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), mkm (nm)";
                        else
                            Program.f15.chart2.ChartAreas[0].AxisX.Title = "Particle size, mkm";
                        label70.Text = "mkm";
                        label71.Text = "mkm";
                        label73.Text = "mkm";
                    }
                }
                if (toolStripComboBox2.SelectedIndex == 2)
                {
                    if (Program.f1.Rus == true)
                    {
                        if (Part20 == true && CSR20 == true)
                            Program.f15.chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), нм";
                        else
                            Program.f15.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, нм";
                        label70.Text = "нм";
                        label71.Text = "нм";
                        label73.Text = "нм";
                    }
                    if (Program.f1.Eng == true)
                    {
                        if (Part20 == true && CSR20 == true)
                            Program.f15.chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), nm";
                        else
                            Program.f15.chart2.ChartAreas[0].AxisX.Title = "Particle size, nm";
                        label70.Text = "nm";
                        label71.Text = "nm";
                        label73.Text = "nm";
                    }


                }


                if (dataGridView5.Rows.Count > 2)
                {
                    X_interval20 = Convert.ToDouble(dataGridView5.Rows[1].Cells[0].Value) - Convert.ToDouble(dataGridView5.Rows[0].Cells[0].Value);
                    X_min20 = Convert.ToDouble(dataGridView5.Rows[0].Cells[0].Value) - X_interval20;
                    Program.f15.chart2.ChartAreas[0].AxisX.Minimum = X_min20;
                    Program.f15.chart2.ChartAreas[0].AxisX.Interval = X_interval20;
                    if (X_min20 <= 0)
                    {
                        Program.f15.chart2.ChartAreas[0].AxisX.IsStartedFromZero = true;
                        for (int i = 0; i < N20; i++)
                            Program.f15.chart2.Series[0].Points[i].AxisLabel = (dataGridView5.Rows[i].Cells[0].Value).ToString();
                    }
                }
            }
        }

        private void dataGridView5_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                dataGridView5.Rows[e.RowIndex].Cells[3].Value = Math.Round(Math.Log(Convert.ToDouble(dataGridView5.Rows[e.RowIndex].Cells[0].Value)), 2);
            }
        }

        private void trackBar15_MouseUp(object sender, MouseEventArgs e)
        {
            MD = false;
            trackBar15.Value = 0;
        }

        private void trackBar15_MouseDown(object sender, MouseEventArgs e)
        {
            MD = true;
        }

        private void trackBar3_ValueChanged(object sender, EventArgs e)
        {
            if (!MD && !KD)
            {
                trackBar3.Value = 0;
                par_a_20 = par_a0_20;
            }
            par_a0_20 = par_a_20 + Convert.ToDouble(trackBar3.Value) * 0.001;
            textBox21.Text = Math.Round(par_a0_20, 3).ToString("0.000");
            Program.f15.chart2.Series[1].Points.Clear();
            if (LogNorm_20)
            {
                for (double i = 0.1; i < X_max20 + r; i = i + 0.1)
                {
                    fDi = r / (par_b_20 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a0_20) * (Math.Log(i) - par_a0_20) / (2 * par_b_20 * par_b_20)) * Area_20;
                    Program.f15.chart2.Series[1].Points.AddXY(i, fDi);
                }
                D20 = Math.Exp(par_a0_20) * Math.Exp(0.5 * par_b_20 * par_b_20);
                Mo = Math.Exp(par_a0_20 - par_b_20 * par_b_20);
                m = Math.Exp(par_a0_20);
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N20; i++)
                {
                    if (!Ln20)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a0_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a0_20) / (2 * par_b_20 * par_b_20)) * Area_20)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a0_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a0_20) / (2 * par_b_20 * par_b_20)) * Area_20)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox67.Text = Math.Round(R2, 4).ToString("0.0000");
                textBox66.Text = Math.Round(par_a0_20, 3).ToString("0.000");
                textBox63.Text = Math.Round(D20, 1).ToString("0.0");
                textBox64.Text = Math.Round(m, 1).ToString("0.0");
                textBox65.Text = Math.Round(Mo, 1).ToString("0.0");
            }
            if (Norm_20)
            {
                for (double i = 0.1; i < X_max20 + r; i = i + 0.1)
                {
                    fDi = r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a0_20) * (i - par_a0_20) / (2 * par_b_20 * par_b_20)) * Area_20;
                    Program.f15.chart2.Series[1].Points.AddXY(i, fDi);
                }
                D20 = par_a0_20;
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N20; i++)
                {
                    //if (!Ln30)
                    //  SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) - par_a0_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) - par_a0_30) / (2 * par_b_30 * par_b_30)) * Area_30)), 2);
                    //else
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a0_20) * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a0_20) / (2 * par_b_20 * par_b_20)) * Area_20)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox67.Text = Math.Round(R2, 4).ToString("0.0000");
                textBox66.Text = Math.Round(par_a0_20, 3).ToString("0.000");
                textBox64.Text = Math.Round(Math.Exp(D20), 1).ToString("0.0");
            }
        }

        private void trackBar3_MouseUp(object sender, MouseEventArgs e)
        {
            MD = false;
            trackBar3.Value = 0;
        }

        private void trackBar3_MouseDown(object sender, MouseEventArgs e)
        {
            MD = true;
        }

        private void trackBar3_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = false;
                trackBar3.Value = 0;
            }
        }

        private void trackBar3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = true;
            }
        }

        private void trackBar4_ValueChanged(object sender, EventArgs e)
        {
            if (!MD && !KD)
            {
                trackBar4.Value = 0;
                par_b_20 = par_b0_20;
            }
            par_b0_20 = par_b_20 + Convert.ToDouble(trackBar4.Value) * 0.001;
            textBox14.Text = Math.Round(par_b0_20, 3).ToString("0.000");
            Program.f15.chart2.Series[1].Points.Clear();
            if (LogNorm_20)
            {
                for (double i = 0.1; i < X_max20 + r; i = i + 0.1)
                {
                    fDi = r / (par_b0_20 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_20) * (Math.Log(i) - par_a_20) / (2 * par_b0_20 * par_b0_20)) * Area_20;
                    Program.f15.chart2.Series[1].Points.AddXY(i, fDi);
                }
                D20 = Math.Exp(par_a_20) * Math.Exp(0.5 * par_b0_20 * par_b0_20);
                Mo = Math.Exp(par_a_20 - par_b0_20 * par_b0_20);

                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N20; i++)
                {
                    if (!Ln20)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b0_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) / (2 * par_b0_20 * par_b0_20)) * Area_20)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b0_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) / (2 * par_b0_20 * par_b0_20)) * Area_20)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox67.Text = Math.Round(R2, 4).ToString("0.0000");
                textBox70.Text = Math.Round(par_b0_20, 3).ToString("0.000");
                textBox68.Text = Math.Round(par_b0_20, 3).ToString("0.000");
                textBox63.Text = Math.Round(D20, 1).ToString("0.0");
                textBox65.Text = Math.Round(Mo, 1).ToString("0.0");
            }
            if (Norm_20)
            {
                for (double i = 0.1; i < X_max20 + r; i = i + 0.1)
                {
                    fDi = r / (par_b0_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_20) * (i - par_a_20) / (2 * par_b0_20 * par_b0_20)) * Area_20;
                    Program.f15.chart2.Series[1].Points.AddXY(i, fDi);
                }

                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N20; i++)
                {
                    // if (!Ln32)
                    //SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b0_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) - par_a_30) / (2 * par_b0_30 * par_b0_30)) * Area_30)), 2);
                    //else
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b0_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) / (2 * par_b0_20 * par_b0_20)) * Area_20)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox67.Text = Math.Round(R2, 4).ToString("0.0000");
                textBox70.Text = Math.Round(par_b0_20, 3).ToString("0.000");
                textBox68.Text = Math.Round(par_b0_20, 3).ToString("0.000");

            }
        }

        private void trackBar4_MouseUp(object sender, MouseEventArgs e)
        {
            MD = false;
            trackBar4.Value = 0;
        }

        private void trackBar4_MouseLeave(object sender, EventArgs e)
        {
            tB = false;
        }

        private void trackBar4_MouseEnter(object sender, EventArgs e)
        {
            tB = true;
        }

        private void trackBar4_MouseDown(object sender, MouseEventArgs e)
        {
            MD = true;
        }

        private void trackBar4_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = false;
                trackBar4.Value = 0;
            }
        }

        private void trackBar4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = true;
            }
        }

        private void trackBar2_ValueChanged(object sender, EventArgs e)
        {
            if (!MD && !KD)
            {
                Area_20 = Area0_20;
                trackBar2.Value = 0;
            }
            Area0_20 = Area_20 + Convert.ToDouble(trackBar2.Value) / 10;
            Program.f15.chart2.Series[1].Points.Clear();
            if (LogNorm_20)
            {
                for (double i = 0.1; i < X_max20 + r; i = i + 0.1)
                {
                    fDi = r / (par_b_20 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_20) * (Math.Log(i) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20;
                    Program.f15.chart2.Series[1].Points.AddXY(i, fDi);
                }

                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N20; i++)
                {
                    if (!Ln20)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox13.Text = Math.Round(Area0_20, 1).ToString();
                textBox67.Text = Math.Round(R2, 4).ToString("0.0000");

            }
            if (Norm_20)
            {
                for (double i = 0.1; i < X_max20 + r; i = i + 0.1)
                {
                    fDi = r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_20) * (i - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20;
                    Program.f15.chart2.Series[1].Points.AddXY(i, fDi);
                }

                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N20; i++)
                {
                    if (!Ln20)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value) - par_a_20) * (Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox13.Text = Math.Round(Area0_20, 1).ToString();
                textBox67.Text = Math.Round(R2, 4).ToString("0.0000");

            }
        }

        private void trackBar2_MouseUp(object sender, MouseEventArgs e)
        {
            MD = false;
            trackBar2.Value = 0;
        }

        private void trackBar2_MouseLeave(object sender, EventArgs e)
        {
            tB = false;
        }

        private void trackBar2_MouseEnter(object sender, EventArgs e)
        {
            tB = true;
        }

        private void trackBar2_MouseDown(object sender, MouseEventArgs e)
        {
            MD = true;
        }

        private void trackBar2_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = false;
                trackBar2.Value = 0;
            }
        }

        private void trackBar2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = true;
            }
        }

        private void textBox21_TextChanged(object sender, EventArgs e)
        {
            if (TB)
            {
                try
                {
                    par_a_20 = double.Parse(textBox21.Text);
                    Program.f15.chart2.Series[1].Points.Clear();
                    if (LogNorm_20)
                    {
                        for (double i = 0.1; i < X_max20 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_20 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_20) * (Math.Log(i) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area_20;
                            Program.f15.chart2.Series[1].Points.AddXY(i, fDi);
                        }
                        D20 = Math.Exp(par_a_20) * Math.Exp(0.5 * par_b_20 * par_b_20);
                        Mo = Math.Exp(par_a_20 - par_b_20 * par_b_20);
                        m = Math.Exp(par_a_20);
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N20; i++)
                        {
                            if (!Ln20)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area_20)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area_20)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        par_a0_20 = par_a_20;
                        textBox67.Text = Math.Round(R2, 4).ToString("0.0000");
                        textBox66.Text = Math.Round(par_a_20, 3).ToString("0.000");
                        textBox63.Text = Math.Round(D20, 1).ToString("0.0");
                        textBox64.Text = Math.Round(m, 1).ToString("0.0");
                        textBox65.Text = Math.Round(Mo, 1).ToString("0.0");
                    }
                    if (Norm_20)
                    {
                        for (double i = 0.1; i < X_max20 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_20) * (i - par_a_20) / (2 * par_b_20 * par_b_20)) * Area_20;
                            Program.f15.chart2.Series[1].Points.AddXY(i, fDi);
                        }
                        D20 = par_a_20;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N20; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area_20)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        par_a0_20 = par_a_20;
                        textBox67.Text = Math.Round(R2, 4).ToString("0.0000");
                        textBox66.Text = Math.Round(par_a_20, 3).ToString("0.000");
                        textBox64.Text = Math.Round(Math.Exp(D20), 1).ToString("0.0");
                    }
                }
                catch
                {
                    return;
                }
            }
        }

        private void textBox21_Leave(object sender, EventArgs e)
        {
            TB = false;
        }

        private void textBox21_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && (e.KeyChar != (char)Keys.Back) && !(e.KeyChar == ','))
            {
                e.Handled = true;
            }
        }

        private void textBox21_Enter(object sender, EventArgs e)
        {
            TB = true;
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            if (TB)
            {
                try
                {
                    par_b_20 = double.Parse(textBox14.Text);
                    Program.f15.chart2.Series[1].Points.Clear();
                    if (LogNorm_20)
                    {
                        for (double i = 0.1; i < X_max20 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_20 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_20) * (Math.Log(i) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area_20;
                            Program.f15.chart2.Series[1].Points.AddXY(i, fDi);
                        }
                        D20 = Math.Exp(par_a_20) * Math.Exp(0.5 * par_b_20 * par_b_20);
                        Mo = Math.Exp(par_a_20 - par_b_20 * par_b_20);

                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N20; i++)
                        {
                            if (!Ln20)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area_20)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area_20)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        par_b0_20 = par_b_20;
                        textBox67.Text = Math.Round(R2, 4).ToString("0.0000");
                        textBox70.Text = Math.Round(par_b_20, 3).ToString("0.000");
                        textBox68.Text = Math.Round(par_b_20, 3).ToString("0.000");
                        textBox63.Text = Math.Round(D20, 1).ToString("0.0");
                        textBox65.Text = Math.Round(Mo, 1).ToString("0.0");
                    }
                    if (Norm_20)
                    {
                        for (double i = 0.1; i < X_max20 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_20) * (i - par_a_20) / (2 * par_b_20 * par_b_20)) * Area_20;
                            Program.f15.chart2.Series[1].Points.AddXY(i, fDi);
                        }

                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N20; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area_20)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        par_b0_20 = par_b_20;
                        textBox67.Text = Math.Round(R2, 4).ToString("0.0000");
                        textBox68.Text = Math.Round(par_b_20, 3).ToString("0.000");
                        textBox70.Text = Math.Round(par_b_20, 3).ToString("0.000");
                        //textBox41.Text = Math.Round(Math.Exp(par_a_32), 1).ToString("0.0");
                        //textBox3.Text = Math.Round(par_a, 3).ToString("0.000");
                        //textBox11.Text = Math.Round(D10, 1).ToString("0.0");
                    }
                }
                catch
                {
                    return;
                }
            }
        }

        private void textBox14_Leave(object sender, EventArgs e)
        {
            TB = false;
        }

        private void textBox14_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && (e.KeyChar != (char)Keys.Back) && !(e.KeyChar == ','))
            {
                e.Handled = true;
            }
        }

        private void textBox14_Enter(object sender, EventArgs e)
        {
            TB = true;
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            if (TB)
            {
                try
                {
                    Area_20 = double.Parse(textBox13.Text);
                    Program.f15.chart2.Series[1].Points.Clear();
                    if (LogNorm_20)
                    {
                        for (double i = 0.1; i < X_max20 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_20 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_20) * (Math.Log(i) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area_20;
                            Program.f15.chart2.Series[1].Points.AddXY(i, fDi);
                        }

                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N20; i++)
                        {
                            if (!Ln20)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area_20)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area_20)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        Area0_20 = Area_20;
                        textBox67.Text = Math.Round(R2, 4).ToString("0.0000");

                    }
                    if (Norm_20)
                    {
                        for (double i = 0.1; i < X_max20 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_20) * (i - par_a_20) / (2 * par_b_20 * par_b_20)) * Area_20;
                            Program.f15.chart2.Series[1].Points.AddXY(i, fDi);
                        }

                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N20; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area_20)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        Area0_20 = Area_20;
                        textBox67.Text = Math.Round(R2, 4).ToString("0.0000");

                    }
                }
                catch
                {
                    return;
                }

            }
        }

        private void textBox13_Leave(object sender, EventArgs e)
        {
            TB = false;
        }

        private void textBox13_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && (e.KeyChar != (char)Keys.Back) && !(e.KeyChar == ','))
            {
                e.Handled = true;
            }
        }

        private void textBox13_Enter(object sender, EventArgs e)
        {
            TB = true;
        }

        private void trackBar15_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = true;
            }
        }

        private void trackBar15_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = false;
                trackBar15.Value = 0;
            }
        }

        private void textBox52_TextChanged(object sender, EventArgs e)
        {
            if (TB)
            {
                try
                {
                    par_a_43 = double.Parse(textBox52.Text);
                    Program.f14.chart2.Series[1].Points.Clear();
                    if (LogNorm_43)
                    {
                        for (double i = 0.1; i < X_max43 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_43 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_43) * (Math.Log(i) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area_43;
                            Program.f14.chart2.Series[1].Points.AddXY(i, fDi);
                        }
                        D43 = Math.Exp(par_a_43) * Math.Exp(0.5 * par_b_43 * par_b_43);
                        Mo = Math.Exp(par_a_43 - par_b_43 * par_b_43);
                        m = Math.Exp(par_a_43);
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N43; i++)
                        {
                            if (!Ln43)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area_43)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area_43)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        par_a0_43 = par_a_43;
                        textBox57.Text = Math.Round(R2, 4).ToString("0.0000");
                        textBox56.Text = Math.Round(par_a_43, 3).ToString("0.000");
                        textBox53.Text = Math.Round(D43, 1).ToString("0.0");
                        textBox54.Text = Math.Round(m, 1).ToString("0.0");
                        textBox55.Text = Math.Round(Mo, 1).ToString("0.0");
                    }
                    if (Norm_43)
                    {
                        for (double i = 0.1; i < X_max43 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_43) * (i - par_a_43) / (2 * par_b_43 * par_b_43)) * Area_43;
                            Program.f14.chart2.Series[1].Points.AddXY(i, fDi);
                        }
                        D43 = par_a_43;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N43; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area_43)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        par_a0_43 = par_a_43;
                        textBox57.Text = Math.Round(R2, 4).ToString("0.0000");
                        textBox56.Text = Math.Round(par_a_43, 3).ToString("0.000");
                        textBox54.Text = Math.Round(Math.Exp(D43), 1).ToString("0.0");
                    }
                }
                catch
                {
                    return;
                }
            }
        }

        private void textBox52_Leave(object sender, EventArgs e)
        {
            TB = false;
        }

        private void textBox52_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && (e.KeyChar != (char)Keys.Back) && !(e.KeyChar == ','))
            {
                e.Handled = true;
            }
        }

        private void textBox52_Enter(object sender, EventArgs e)
        {
            TB = true;
        }

        private void textBox51_TextChanged(object sender, EventArgs e)
        {
            if (TB)
            {
                try
                {
                    par_b_43 = double.Parse(textBox51.Text);
                    Program.f14.chart2.Series[1].Points.Clear();
                    if (LogNorm_43)
                    {
                        for (double i = 0.1; i < X_max43 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_43 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_43) * (Math.Log(i) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area_43;
                            Program.f14.chart2.Series[1].Points.AddXY(i, fDi);
                        }
                        D43 = Math.Exp(par_a_43) * Math.Exp(0.5 * par_b_43 * par_b_43);
                        Mo = Math.Exp(par_a_43 - par_b_43 * par_b_43);

                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N43; i++)
                        {
                            if (!Ln43)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area_43)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area_43)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        par_b0_43 = par_b_43;
                        textBox57.Text = Math.Round(R2, 4).ToString("0.0000");
                        textBox60.Text = Math.Round(par_b_43, 3).ToString("0.000");
                        textBox58.Text = Math.Round(par_b_43, 3).ToString("0.000");
                        textBox53.Text = Math.Round(D43, 1).ToString("0.0");
                        textBox55.Text = Math.Round(Mo, 1).ToString("0.0");
                    }
                    if (Norm_43)
                    {
                        for (double i = 0.1; i < X_max43 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_43) * (i - par_a_43) / (2 * par_b_43 * par_b_43)) * Area_43;
                            Program.f14.chart2.Series[1].Points.AddXY(i, fDi);
                        }

                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N43; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area_43)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        par_b0_43 = par_b_43;
                        textBox57.Text = Math.Round(R2, 4).ToString("0.0000");
                        textBox58.Text = Math.Round(par_b_43, 3).ToString("0.000");
                        textBox60.Text = Math.Round(par_b_43, 3).ToString("0.000");
                        //textBox41.Text = Math.Round(Math.Exp(par_a_32), 1).ToString("0.0");
                        //textBox3.Text = Math.Round(par_a, 3).ToString("0.000");
                        //textBox11.Text = Math.Round(D10, 1).ToString("0.0");
                    }
                }
                catch
                {
                    return;
                }
            }
        }

        private void textBox51_Leave(object sender, EventArgs e)
        {
            TB = false;
        }

        private void textBox51_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && (e.KeyChar != (char)Keys.Back) && !(e.KeyChar == ','))
            {
                e.Handled = true;
            }
        }

        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void textBox51_Enter(object sender, EventArgs e)
        {
            TB = true;
        }

        private void textBox50_TextChanged(object sender, EventArgs e)
        {
            if (TB)
            {
                try
                {
                    Area_43 = double.Parse(textBox50.Text);
                    Program.f14.chart2.Series[1].Points.Clear();
                    if (LogNorm_43)
                    {
                        for (double i = 0.1; i < X_max43 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_43 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_43) * (Math.Log(i) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area_43;
                            Program.f14.chart2.Series[1].Points.AddXY(i, fDi);
                        }

                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N43; i++)
                        {
                            if (!Ln43)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area_43)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area_43)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        Area0_43 = Area_43;
                        textBox57.Text = Math.Round(R2, 4).ToString("0.0000");

                    }
                    if (Norm_43)
                    {
                        for (double i = 0.1; i < X_max43 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_43) * (i - par_a_43) / (2 * par_b_43 * par_b_43)) * Area_43;
                            Program.f14.chart2.Series[1].Points.AddXY(i, fDi);
                        }

                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N43; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area_43)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        Area0_43 = Area_43;
                        textBox57.Text = Math.Round(R2, 4).ToString("0.0000");

                    }
                }
                catch
                {
                    return;
                }

            }
        }

        private void textBox50_Leave(object sender, EventArgs e)
        {
            TB = false;
        }

        private void textBox50_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && (e.KeyChar != (char)Keys.Back) && !(e.KeyChar == ','))
            {
                e.Handled = true;
            }
        }

        private void textBox50_Enter(object sender, EventArgs e)
        {
            TB = true;
        }

        private void trackBar13_MouseUp(object sender, MouseEventArgs e)
        {
            MD = false;
            trackBar13.Value = 0;
        }

        private void trackBar13_MouseDown(object sender, MouseEventArgs e)
        {
            MD = true;
        }

        private void trackBar13_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = true;
            }
        }

        private void trackBar13_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = false;
                trackBar13.Value = 0;
            }
        }

        private void trackBar14_ValueChanged(object sender, EventArgs e)
        {
            if (!MD && !KD)
            {
                trackBar14.Value = 0;
                par_b_32 = par_b0_32;
            }
            par_b0_32 = par_b_32 + Convert.ToDouble(trackBar14.Value) * 0.001;
            textBox38.Text = Math.Round(par_b0_32, 3).ToString("0.000");
            Program.f13.chart2.Series[1].Points.Clear();
            if (LogNorm_32)
            {
                for (double i = 0.1; i < X_max32 + r; i = i + 0.1)
                {
                    fDi = r / (par_b0_32 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_32) * (Math.Log(i) - par_a_32) / (2 * par_b0_32 * par_b0_32)) * Area_32;
                    Program.f13.chart2.Series[1].Points.AddXY(i, fDi);
                }
                D32 = Math.Exp(par_a_32) * Math.Exp(0.5 * par_b0_32 * par_b0_32);
                Mo = Math.Exp(par_a_32 - par_b0_32 * par_b0_32);

                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N32; i++)
                {
                    if (!Ln32)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b0_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) / (2 * par_b0_32 * par_b0_32)) * Area_32)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b0_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) / (2 * par_b0_32 * par_b0_32)) * Area_32)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox44.Text = Math.Round(R2, 4).ToString("0.0000");
                textBox47.Text = Math.Round(par_b0_32, 3).ToString("0.000");
                textBox45.Text = Math.Round(par_b0_32, 3).ToString("0.000");
                textBox40.Text = Math.Round(D32, 1).ToString("0.0");

                textBox42.Text = Math.Round(Mo, 1).ToString("0.0");
            }
            if (Norm_32)
            {
                for (double i = 0.1; i < X_max32 + r; i = i + 0.1)
                {
                    fDi = r / (par_b0_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_32) * (i - par_a_32) / (2 * par_b0_32 * par_b0_32)) * Area_32;
                    Program.f13.chart2.Series[1].Points.AddXY(i, fDi);
                }

                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N32; i++)
                {
                   // if (!Ln32)
                        //SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b0_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) - par_a_30) / (2 * par_b0_30 * par_b0_30)) * Area_30)), 2);
                    //else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b0_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) / (2 * par_b0_32 * par_b0_32)) * Area_32)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox44.Text = Math.Round(R2, 4).ToString("0.0000");
                textBox45.Text = Math.Round(par_b0_32, 3).ToString("0.000");
                textBox47.Text = Math.Round(par_b0_32, 3).ToString("0.000");

            }
        }

        private void trackBar14_MouseUp(object sender, MouseEventArgs e)
        {
            MD = false;
            trackBar14.Value = 0;
        }

        private void trackBar14_MouseDown(object sender, MouseEventArgs e)
        {
            MD = true;
        }

        private void trackBar14_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = false;
                trackBar14.Value = 0;
            }
        }

        private void trackBar14_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = true;
            }
        }

        private void trackBar12_ValueChanged(object sender, EventArgs e)
        {
            if (!MD && !KD)
            {
                Area_32 = Area0_32;
                trackBar12.Value = 0;
            }
            Area0_32 = Area_32 + Convert.ToDouble(trackBar12.Value) / 10;
            Program.f13.chart2.Series[1].Points.Clear();
            if (LogNorm_32)
            {
                for (double i = 0.1; i < X_max32 + r; i = i + 0.1)
                {
                    fDi = r / (par_b_32 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_32) * (Math.Log(i) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32;
                    Program.f13.chart2.Series[1].Points.AddXY(i, fDi);
                }

                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N32; i++)
                {
                    if (!Ln32)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox37.Text = Math.Round(Area0_32, 1).ToString();
                textBox44.Text = Math.Round(R2, 4).ToString("0.0000");

            }
            if (Norm_32)
            {
                for (double i = 0.1; i < X_max32 + r; i = i + 0.1)
                {
                    fDi = r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_32) * (i - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32;
                    Program.f13.chart2.Series[1].Points.AddXY(i, fDi);
                }

                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N32; i++)
                {
                    if (!Ln32)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value) - par_a_32) * (Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox37.Text = Math.Round(Area0_32, 1).ToString();
                textBox44.Text = Math.Round(R2, 4).ToString("0.0000");

            }
        }

        private void trackBar12_MouseUp(object sender, MouseEventArgs e)
        {
            MD = false;
            trackBar12.Value = 0;
        }

        private void trackBar12_MouseDown(object sender, MouseEventArgs e)
        {
            MD = true;
        }

        private void trackBar12_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = true;
            }
        }

        private void trackBar12_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = false;
                trackBar12.Value = 0;
            }
        }

        private void textBox39_TextChanged(object sender, EventArgs e)
        {
            if (TB)
            {
                try
                {
                    par_a_32 = double.Parse(textBox39.Text);
                    Program.f13.chart2.Series[1].Points.Clear();
                    if (LogNorm_32)
                    {
                        for (double i = 0.1; i < X_max32 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_32 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_32) * (Math.Log(i) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area_32;
                            Program.f13.chart2.Series[1].Points.AddXY(i, fDi);
                        }
                        D32 = Math.Exp(par_a_32) * Math.Exp(0.5 * par_b_32 * par_b_32);
                        Mo = Math.Exp(par_a_32 - par_b_32 * par_b_32);
                        m = Math.Exp(par_a_32);
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N32; i++)
                        {
                            if (!Ln32)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area_32)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area_32)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        par_a0_32 = par_a_32;
                        textBox44.Text = Math.Round(R2, 4).ToString("0.0000");
                        textBox43.Text = Math.Round(par_a_32, 3).ToString("0.000");
                        textBox40.Text = Math.Round(D32, 1).ToString("0.0");
                        textBox41.Text = Math.Round(m, 1).ToString("0.0");
                        textBox42.Text = Math.Round(Mo, 1).ToString("0.0");
                    }
                    if (Norm_32 && Ln32)
                    {
                        for (double i = 0.1; i < X_max32 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_32) * (i - par_a_32) / (2 * par_b_32 * par_b_32)) * Area_32;
                            Program.f13.chart2.Series[1].Points.AddXY(i, fDi);
                        }
                        D32 = par_a_32;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N32; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area_32)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        par_a0_32 = par_a_32;
                        textBox44.Text = Math.Round(R2, 4).ToString("0.0000");
                        textBox43.Text = Math.Round(par_a_32, 3).ToString("0.000");
                        textBox41.Text = Math.Round(Math.Exp(D32), 1).ToString("0.0");
                    }
                }
                catch
                {
                    return;
                }
            }
        }

        private void textBox39_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && (e.KeyChar != (char)Keys.Back) && !(e.KeyChar == ','))
            {
                e.Handled = true;
            }
        }

        private void textBox39_Leave(object sender, EventArgs e)
        {
            TB = false;
        }

        private void textBox39_Enter(object sender, EventArgs e)
        {
            TB = true;
        }

        private void textBox38_TextChanged(object sender, EventArgs e)
        {
            if (TB)
            {
                try
                {
                    par_b_32 = double.Parse(textBox38.Text);
                    Program.f13.chart2.Series[1].Points.Clear();
                    if (LogNorm_32)
                    {
                        for (double i = 0.1; i < X_max32 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_32 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_32) * (Math.Log(i) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area_32;
                            Program.f13.chart2.Series[1].Points.AddXY(i, fDi);
                        }
                        D32 = Math.Exp(par_a_32) * Math.Exp(0.5 * par_b_32 * par_b_32);
                        Mo = Math.Exp(par_a_32 - par_b_32 * par_b_32);

                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N32; i++)
                        {
                            if (!Ln32)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area_32)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area_32)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        par_b0_32 = par_b_32;
                        textBox44.Text = Math.Round(R2, 4).ToString("0.0000");
                        textBox45.Text = Math.Round(par_b_32, 3).ToString("0.000");
                        textBox47.Text = Math.Round(par_b_32, 3).ToString("0.000");
                        textBox40.Text = Math.Round(D32, 1).ToString("0.0");

                        textBox42.Text = Math.Round(Mo, 1).ToString("0.0");
                    }
                    if (Norm_32 && Ln32)
                    {
                        for (double i = 0.1; i < X_max32 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_32) * (i - par_a_32) / (2 * par_b_32 * par_b_32)) * Area_32;
                            Program.f13.chart2.Series[1].Points.AddXY(i, fDi);
                        }

                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N32; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area_32)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        par_b0_32 = par_b_32;
                        textBox44.Text = Math.Round(R2, 4).ToString("0.0000");
                        textBox47.Text = Math.Round(par_b_32, 3).ToString("0.000");
                        textBox45.Text = Math.Round(par_b_32, 3).ToString("0.000");
                        //textBox41.Text = Math.Round(Math.Exp(par_a_32), 1).ToString("0.0");
                        //textBox3.Text = Math.Round(par_a, 3).ToString("0.000");
                        //textBox11.Text = Math.Round(D10, 1).ToString("0.0");
                    }
                }
                catch
                {
                    return;
                }
            }
        }

        private void textBox38_Leave(object sender, EventArgs e)
        {
            TB = false;
        }

        private void textBox38_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && (e.KeyChar != (char)Keys.Back) && !(e.KeyChar == ','))
            {
                e.Handled = true;
            }
        }

        private void textBox38_Enter(object sender, EventArgs e)
        {
            TB = true;
        }

        private void tabControl1_MouseHover(object sender, EventArgs e)
        {
            tabControl1.TabPages["tabPage1"].ToolTipText = line_raspr;
            if (tabControl1.TabPages.ContainsKey("tabPage2"))
                tabControl1.TabPages["tabPage2"].ToolTipText = cub_raspr;
            if (tabControl1.TabPages.ContainsKey("tabPage3"))
                tabControl1.TabPages["tabPage3"].ToolTipText = vol_surf_raspr;
            if (tabControl1.TabPages.ContainsKey("tabPage4"))
                tabControl1.TabPages["tabPage4"].ToolTipText = vol_weight_raspr;
            if (tabControl1.TabPages.ContainsKey("tabPage13"))
                tabControl1.TabPages["tabPage13"].ToolTipText = surf_raspr;
        }

        private void textBox37_TextChanged(object sender, EventArgs e)
        {
            if (TB)
            {
                try
                {
                    Area_32 = double.Parse(textBox37.Text);
                    Program.f13.chart2.Series[1].Points.Clear();
                    if (LogNorm_32)
                    {
                        for (double i = 0.1; i < X_max32 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_32 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_32) * (Math.Log(i) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area_32;
                            Program.f13.chart2.Series[1].Points.AddXY(i, fDi);
                        }

                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N32; i++)
                        {
                            if (!Ln32)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area_32)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area_32)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        Area0_32 = Area_32;
                        textBox44.Text = Math.Round(R2, 4).ToString("0.0000");

                    }
                    if (Norm_32 && Ln32)
                    {
                        for (double i = 0.1; i < X_max32 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_32) * (i - par_a_32) / (2 * par_b_32 * par_b_32)) * Area_32;
                            Program.f13.chart2.Series[1].Points.AddXY(i, fDi);
                        }

                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N32; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area_32)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        Area0_32 = Area_32;
                        textBox44.Text = Math.Round(R2, 4).ToString("0.0000");

                    }
                }
                catch
                {
                    return;
                }

            }
        }

        private void textBox37_Leave(object sender, EventArgs e)
        {
            TB = false;
        }

        private void textBox37_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && (e.KeyChar != (char)Keys.Back) && !(e.KeyChar == ','))
            {
                e.Handled = true;
            }
        }

        private void textBox37_Enter(object sender, EventArgs e)
        {
            TB = true;
        }

        private void trackBar9_MouseDown(object sender, MouseEventArgs e)
        {
            MD = true;
        }

        private void trackBar9_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = false;
                trackBar9.Value = 0;
            }
        }

        private void trackBar9_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = true;
            }
        }

        private void dataGridView2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                dataGridView2.Rows[e.RowIndex].Cells[3].Value = Math.Round(Math.Log(Convert.ToDouble(dataGridView2.Rows[e.RowIndex].Cells[0].Value)), 2);
            }
        }

        private void textBox26_TextChanged(object sender, EventArgs e)
        {
            if (TB)
            {
                try
                {
                    par_a_30 = double.Parse(textBox26.Text);
                    Program.f12.chart2.Series[1].Points.Clear();
                    if (LogNorm_30)
                    {
                        for (double i = 0.1; i < X_max30 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_30 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_30) * (Math.Log(i) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area_30;
                            Program.f12.chart2.Series[1].Points.AddXY(i, fDi);
                        }
                        D30 = Math.Exp(par_a_30) * Math.Exp(0.5 * par_b_30 * par_b_30);
                        Mo = Math.Exp(par_a_30 - par_b_30 * par_b_30);
                        m = Math.Exp(par_a_30);
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N30; i++)
                        {
                            if (!Ln30)
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area_30)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area_30)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        par_a0_30 = par_a_30;
                        textBox31.Text = Math.Round(R2, 4).ToString("0.0000");
                        textBox30.Text = Math.Round(par_a_30, 3).ToString("0.000");
                        textBox27.Text = Math.Round(D30, 1).ToString("0.0");
                        textBox28.Text = Math.Round(m, 1).ToString("0.0");
                        textBox29.Text = Math.Round(Mo, 1).ToString("0.0");
                    }
                    if (Norm_30 && Ln30)
                    {
                        for (double i = 0.1; i < X_max30 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_30) * (i - par_a_30) / (2 * par_b_30 * par_b_30)) * Area_30;
                            Program.f12.chart2.Series[1].Points.AddXY(i, fDi);
                        }
                        D30 = par_a_30;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N30; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area_30)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        par_a0_30 = par_a_30;
                        textBox31.Text = Math.Round(R2, 4).ToString("0.0000");
                        textBox30.Text = Math.Round(par_a_30, 3).ToString("0.000");
                        textBox28.Text = Math.Round(Math.Exp(D30), 1).ToString("0.0");
                    }
                }
                catch
                {
                    return;
                }
            }
        }

        private void textBox26_Leave(object sender, EventArgs e)
        {
            TB = false;
        }

        private void textBox26_Enter(object sender, EventArgs e)
        {
            TB = true;
        }

        private void label51_MouseHover(object sender, EventArgs e)
        {
            t.SetToolTip(label51, median);
        }

        private void label49_MouseHover(object sender, EventArgs e)
        {
            t.SetToolTip(label49, mode);
        }

        private void pictureBox11_MouseHover(object sender, EventArgs e)
        {
            t.SetToolTip(pictureBox11, Dcp32);
        }

        private void dataGridView3_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                dataGridView3.Rows[e.RowIndex].Cells[3].Value = Math.Round(Math.Log(Convert.ToDouble(dataGridView3.Rows[e.RowIndex].Cells[0].Value)), 2);
            }
        }

        private void textBox26_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && (e.KeyChar != (char)Keys.Back) && !(e.KeyChar == ','))
            {
                e.Handled = true;
            }
        }

        private void textBox25_TextChanged(object sender, EventArgs e)
        {
            if (TB)
            {
                try
                {
                    par_b_30 = double.Parse(textBox25.Text);
                    Program.f12.chart2.Series[1].Points.Clear();
                    if (LogNorm_30)
                    {
                        for (double i = 0.1; i < X_max30 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_30 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_30) * (Math.Log(i) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area_30;
                            Program.f12.chart2.Series[1].Points.AddXY(i, fDi);
                        }
                        D30 = Math.Exp(par_a_30) * Math.Exp(0.5 * par_b_30 * par_b_30);
                        Mo = Math.Exp(par_a_30 - par_b_30 * par_b_30);

                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N30; i++)
                        {
                            if (!Ln30)
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area_30)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area_30)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        par_b0_30 = par_b_30;
                        textBox31.Text = Math.Round(R2, 4).ToString("0.0000");
                        textBox34.Text = Math.Round(par_b_30, 3).ToString("0.000");
                        textBox32.Text = Math.Round(par_b_30, 3).ToString("0.000");
                        textBox27.Text = Math.Round(D30, 1).ToString("0.0");

                        textBox29.Text = Math.Round(Mo, 1).ToString("0.0");
                    }
                    if (Norm_30 && Ln30)
                    {
                        for (double i = 0.1; i < X_max30 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_30) * (i - par_a_30) / (2 * par_b_30 * par_b_30)) * Area_30;
                            Program.f12.chart2.Series[1].Points.AddXY(i, fDi);
                        }

                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N30; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area_30)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        par_b0_30 = par_b_30;
                        textBox31.Text = Math.Round(R2, 4).ToString("0.0000");
                        textBox34.Text = Math.Round(par_b_30, 3).ToString("0.000");
                        textBox32.Text = Math.Round(par_b_30, 3).ToString("0.000");
                        //textBox28.Text=Math.Round(Math.Exp(par_a), 1).ToString("0.0");
                        //textBox3.Text = Math.Round(par_a, 3).ToString("0.000");
                        //textBox11.Text = Math.Round(D10, 1).ToString("0.0");
                    }
                }
                catch
                {
                    return;
                }
            }
        }

        private void textBox25_Leave(object sender, EventArgs e)
        {
            TB = false;
        }

        private void textBox25_Enter(object sender, EventArgs e)
        {
            TB = true;
        }

        private void textBox25_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && (e.KeyChar != (char)Keys.Back) && !(e.KeyChar == ','))
            {
                e.Handled = true;
            }
        }

        private void объемноповерхностноеРаспределениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Convert.ToString(dataGridView1.Rows[0].Cells[0].Value) == string.Empty)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Размеры частиц не заданы! Попробуйте снова!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles size are not determined! Try again!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (Convert.ToString(dataGridView1.Rows[0].Cells[1].Value) == string.Empty)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Число частиц не задано! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles number are not determined! Try again!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            flag_vol_surf_raspr = true;
            CSR = false;
            chart2.Series[3].Points.Clear();
            chart2.Series[4].Points.Clear();
            if (!tabControl1.TabPages.ContainsKey("tabPage3"))
            {
                tabControl1.TabPages.Add(tabPage3);
                
            }
            tabControl1.SelectedTab = tabControl1.TabPages["tabPage3"];

            dataGridView3.Rows.Clear();
            sum_Vol = 0;
            N = Convert.ToInt32(dataGridView1.RowCount);
            N32 = Convert.ToInt32(dataGridView3.RowCount);
            if (N32 < N)
                dataGridView3.Rows.Add(N - N32);
            while (dataGridView1.Rows[N - 1].Cells[0].Value == null)// || dataGridView1.Rows[N - 1].Cells[1].Value == null)
                N = N - 1;

            if (Sum == 0)
            {
                for (int i = 0; i < N; i++)
                {
                    Sum = Sum + Convert.ToDouble(dataGridView1[1, i].Value); // сначал номер стобца, затем номер строки
                }
            }

            for (int i = 0; i < N; i++)
            {
                dataGridView3.Rows[i].Cells[0].Value = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                Vol.Insert(i, Math.Pow(Convert.ToDouble(dataGridView1[0, i].Value),2) * Convert.ToDouble(dataGridView1[1, i].Value));
                dataGridView3.Rows[i].Cells[1].Value = Vol[i];
                dataGridView3.Rows[i].Cells[3].Value = Math.Round(Math.Log(Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value)), 2);
                sum_Vol = sum_Vol + Vol[i];
            }

            for (int i = 0; i < N; i++)
                dataGridView3.Rows[i].Cells[2].Value = Math.Round(Vol[i] / sum_Vol * 100, 2);

            button7_Click(sender, e);


            if (CSR1 == true)
            {
                textBox15.Visible = false;
                textBox16.Visible = false;
                pictureBox7.Visible = false;
                pictureBox8.Visible = false;
                label12.Visible = false;
                label13.Visible = false;
            }
            if (CSR2 == true)
            {
                textBox17.Visible = false;
                textBox18.Visible = false;
                pictureBox9.Visible = false;
                pictureBox10.Visible = false;
                label14.Visible = false;
                label15.Visible = false;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (Convert.ToString(dataGridView3.Rows[0].Cells[0].Value) == string.Empty)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Размеры частиц не заданы! Попробуйте снова!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles size are not determined! Try again!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (Convert.ToString(dataGridView3.Rows[0].Cells[2].Value) == string.Empty)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Доля частиц не задана! Попробуйте снова!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles fraction are not determined! Try again!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            tabControl4.Visible = false;
            if (!open_f13)
            {
                f13 = new Form13();
                f13.Owner = this;
                open_f13 = true;
                f13.Show();
            }
            Program.f13.chart2.Series[0].Points.Clear();
            Program.f13.chart2.Series[1].Points.Clear();
            Program.f13.chart2.Series[2].Points.Clear();
            Program.f13.chart2.Series[5].Points.Clear();
            analyt_func_32 = false;

            N32 = Convert.ToInt32(dataGridView3.RowCount);
            while (Convert.ToString(dataGridView3.Rows[N32 - 1].Cells[0].Value) == string.Empty || Convert.ToString(dataGridView3.Rows[N32 - 1].Cells[2].Value) == string.Empty)
                N32 = N32 - 1;

            if (Sum == 0)
            {
                for (int i = 0; i < N32; i++)
                {
                    Sum = Sum + Convert.ToDouble(dataGridView3[1, i].Value); // сначал номер стобца, затем номер строки
                }
            }

            if (Ln32)
            {
                for (int i = 0; i < N32; i++)
                {
                    x = Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value);
                    y = Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value);
                    Program.f13.chart2.Series[0].Points.AddXY(x, y);
                    X_max32 = Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value);

                }
                Program.f13.r = Convert.ToDouble(dataGridView3.Rows[N32 - 1].Cells[3].Value) - Convert.ToDouble(dataGridView3.Rows[N32 - 2].Cells[3].Value);
                X_max32 = X_max32 + Program.f13.r;
                Program.f13.chart2.ChartAreas[0].AxisY.Minimum = 0;
                Program.f13.chart2.ChartAreas[0].AxisX.Maximum = X_max32;
                Program.f13.chart2.ChartAreas[0].AxisX.Minimum = Convert.ToDouble(dataGridView3.Rows[0].Cells[3].Value) - Program.f13.r;
                Program.f13.chart2.ChartAreas[0].AxisX.Title = "Ln(D)";
                Program.f13.chart2.ChartAreas[0].AxisX.Interval = 0;
            }
            else
            {
                for (int i = 0; i < N32; i++)
                {
                    x = Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value);
                    y = Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value);
                    Program.f13.chart2.Series[0].Points.AddXY(x, y);
                    X_max32 = Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value);
                }
                Program.f13.chart2.Series[0].IsVisibleInLegend = true;
                Program.f13.chart2.Series[1].IsVisibleInLegend = false;
                Program.f13.chart2.Series[2].IsVisibleInLegend = false;
                Program.f13.chart2.Series[3].IsVisibleInLegend = false;
                Program.f13.chart2.Series[4].IsVisibleInLegend = false;
                Program.f13.chart2.Series[5].IsVisibleInLegend = false;
                if (Program.f1.Rus == true)
                    Program.f13.chart2.Series[0].Name = "Гист.";
                if (Program.f1.Eng == true)
                    Program.f13.chart2.Series[0].Name = "Hist.";
                Part32 = true;
                r = Convert.ToDouble(dataGridView3.Rows[1].Cells[0].Value) - Convert.ToDouble(dataGridView3.Rows[0].Cells[0].Value);
                X_max32 = X_max32 + r;
                Program.f13.chart2.ChartAreas[0].AxisY.Minimum = 0;
                Program.f13.chart2.ChartAreas[0].AxisX.Maximum = X_max32;
                if (Part32 == true && CSR32 == true)
                    Program.f13.chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), нм";
                else
                    Program.f13.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, нм";
                Program.f13.chart2.ChartAreas[0].AxisX.TitleFont = new Font("Times New Roman", Program.f13.I, FontStyle.Bold);
                if (Part32 == true && CSR32 == true)
                    Program.f13.chart2.ChartAreas[0].AxisY.Title = "Доля частиц (ОКР), %";
                else
                    Program.f13.chart2.ChartAreas[0].AxisY.Title = "Доля частиц, %";
                if (Program.f1.Eng == true)
                    Program.f13.chart2.ChartAreas[0].AxisY.Title = "Frequency, %";
                Program.f13.chart2.ChartAreas[0].AxisY.TitleFont = new Font("Times New Roman", Program.f13.I, FontStyle.Bold);
                Program.f13.chart2.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Times New Roman", Program.f13.I, System.Drawing.FontStyle.Regular);
                Program.f13.chart2.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Times New Roman", Program.f13.I, System.Drawing.FontStyle.Regular);
                if (toolStripComboBox2.SelectedIndex == 0)
                {
                    if (Program.f1.Rus == true)
                    {
                        if (Part32 == true && CSR32 == true)
                            Program.f13.chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), мм (нм)";
                        else
                            Program.f13.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, мм";
                        label47.Text = "мм";
                        label48.Text = "мм";
                        label50.Text = "мм";
                    }
                    if (Program.f1.Eng == true)
                    {
                        if (Part32 == true && CSR32 == true)
                            Program.f13.chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), mm (nm)";
                        else
                            Program.f13.chart2.ChartAreas[0].AxisX.Title = "Particle size, mm";
                        label47.Text = "mm";
                        label48.Text = "mm";
                        label50.Text = "mm";
                    }
                }
                if (toolStripComboBox2.SelectedIndex == 1)
                {
                    if (Program.f1.Rus == true)
                    {
                        if (Part32 == true && CSR32 == true)
                            Program.f13.chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), мкм (нм)";
                        else
                            Program.f13.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, мкм";
                        label47.Text = "мкм";
                        label48.Text = "мкм";
                        label50.Text = "мкм";
                    }
                    if (Program.f1.Eng == true)
                    {
                        if (Part32 == true && CSR32 == true)
                            Program.f13.chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), mkm (nm)";
                        else
                            Program.f13.chart2.ChartAreas[0].AxisX.Title = "Particle size, mkm";
                        label47.Text = "mkm";
                        label48.Text = "mkm";
                        label50.Text = "mkm";
                    }
                }
                if (toolStripComboBox2.SelectedIndex == 2)
                {
                    if (Program.f1.Rus == true)
                    {
                        if (Part32 == true && CSR32 == true)
                            Program.f13.chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), нм";
                        else
                            Program.f13.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, нм";
                        label47.Text = "нм";
                        label48.Text = "нм";
                        label50.Text = "нм";
                    }
                    if (Program.f1.Eng == true)
                    {
                        if (Part32 == true && CSR32 == true)
                            Program.f13.chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), nm";
                        else
                            Program.f13.chart2.ChartAreas[0].AxisX.Title = "Particle size, nm";
                        label47.Text = "nm";
                        label48.Text = "nm";
                        label50.Text = "nm";
                    }


                }


                if (dataGridView3.Rows.Count > 2)
                {
                    X_interval32 = Convert.ToDouble(dataGridView3.Rows[1].Cells[0].Value) - Convert.ToDouble(dataGridView3.Rows[0].Cells[0].Value);
                    X_min32 = Convert.ToDouble(dataGridView3.Rows[0].Cells[0].Value) - X_interval32;
                    Program.f13.chart2.ChartAreas[0].AxisX.Minimum = X_min32;
                    Program.f13.chart2.ChartAreas[0].AxisX.Interval = X_interval32;
                    if (X_min32 <= 0)
                    {
                        Program.f13.chart2.ChartAreas[0].AxisX.IsStartedFromZero = true;
                        for (int i = 0; i < N32; i++)
                            Program.f13.chart2.Series[0].Points[i].AxisLabel = (dataGridView3.Rows[i].Cells[0].Value).ToString();
                    }
                }
            }
        }

        private void textBox24_TextChanged(object sender, EventArgs e)
        {
            if (TB)
            {
                try
                {
                    Area_30 = double.Parse(textBox24.Text);
                    Program.f12.chart2.Series[1].Points.Clear();
                    if (LogNorm_30)
                    {
                        for (double i = 0.1; i < X_max30 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_30 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_30) * (Math.Log(i) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area_30;
                            Program.f12.chart2.Series[1].Points.AddXY(i, fDi);
                        }

                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N30; i++)
                        {
                            if (!Ln30)
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area_30)), 2);
                            if (Ln30)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area_30)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        Area0_30 = Area_30;
                        textBox31.Text = Math.Round(R2, 4).ToString("0.0000");

                    }
                    if (Norm_30 && Ln30)
                    {
                        for (double i = 0.1; i < X_max30 + r; i = i + 0.1)
                        {
                            fDi = r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_30) * (i - par_a_30) / (2 * par_b_30 * par_b_30)) * Area_30;
                            Program.f12.chart2.Series[1].Points.AddXY(i, fDi);
                        }

                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N30; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area_30)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        Area0_30 = Area_30;
                        textBox31.Text = Math.Round(R2, 4).ToString("0.0000");

                    }
                }
                catch
                {
                    return;
                }

            }
        }

        private void textBox24_Leave(object sender, EventArgs e)
        {
            TB = false;
        }

        private void textBox24_Enter(object sender, EventArgs e)
        {
            TB = true;
        }

        private void textBox24_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && (e.KeyChar != (char)Keys.Back) && !(e.KeyChar == ','))
            {
                e.Handled = true;
            }
        }

        private void trackBar11_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = true;
            }
        }

        private void trackBar10_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = false;
                trackBar10.Value = 0;
            }
        }

        private void trackBar10_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = true;                
            }
        }

        private void trackBar8_ValueChanged(object sender, EventArgs e)
        {
            if (!MD && !KD)
            {
                trackBar8.Value = 0;
                par_a = par_a0;
            }            
            par_a0 = par_a + Convert.ToDouble(trackBar8.Value) * 0.001;
            textBox1.Text = Math.Round(par_a0, 3).ToString("0.000");
            Program.f11.chart2.Series[1].Points.Clear();
            if (LogNorm)
            {                
                for (double i = 0.1; i < X_max + r; i = i + 0.1)
                {
                    fDi = r / (par_b * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a0) * (Math.Log(i) - par_a0) / (2 * par_b * par_b)) * Area;
                    Program.f11.chart2.Series[1].Points.AddXY(i, fDi);
                }
                D10 = Math.Exp(par_a0) * Math.Exp(0.5 * par_b * par_b);
                Mo = Math.Exp(par_a0 - par_b * par_b);
                m = Math.Exp(par_a0);
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N; i++)
                {
                    if (!Ln)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a0) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a0) / (2 * par_b * par_b)) * Area)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a0) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a0) / (2 * par_b * par_b)) * Area)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                }
                R2 = 1 - SSres / SStot;
                if (Ln)
                {
                    D10 = Math.Exp(D10);
                    m = Math.Exp(m);
                    Mo = Math.Exp(Mo);
                }
                textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                textBox3.Text = Math.Round(par_a0, 3).ToString("0.000");
                textBox11.Text = Math.Round(D10, 1).ToString("0.0");
                textBox23.Text = Math.Round(m, 1).ToString("0.0");
                textBox10.Text = Math.Round(Mo, 1).ToString("0.0");
            }
            if (Norm)
            {
                for (double i = 0.1; i < X_max + r; i = i + 0.1)
                {
                    fDi = r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a0) * (i - par_a0) / (2 * par_b * par_b)) * Area;
                    Program.f11.chart2.Series[1].Points.AddXY(i, fDi);
                }
                D10 = par_a0;             
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N; i++)
                {
                    if (!Ln)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a0) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a0) / (2 * par_b * par_b)) * Area)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a0) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a0) / (2 * par_b * par_b)) * Area)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                textBox3.Text = Math.Round(par_a0, 3).ToString("0.000");
                if (!Ln)
                    textBox11.Text = Math.Round(D10, 1).ToString("0.0");
                else
                    textBox23.Text = Math.Round(Math.Exp(D10), 1).ToString("0.0");
            }
        }

        private void trackBar8_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = false;
                trackBar8.Value = 0;
            }
        }

        private void trackBar7_ValueChanged(object sender, EventArgs e)
        {
            if (!MD && !KD)
            {
                trackBar7.Value = 0;
                par_b = par_b0;
            }
            par_b0 = par_b + Convert.ToDouble(trackBar7.Value) * 0.001;
            textBox2.Text = Math.Round(par_b0, 3).ToString("0.000");
            Program.f11.chart2.Series[1].Points.Clear();
            if (LogNorm)
            {
                for (double i = 0.1; i < X_max + r; i = i + 0.1)
                {
                    fDi = r / (par_b0 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a) * (Math.Log(i) - par_a) / (2 * par_b0 * par_b0)) * Area;
                    Program.f11.chart2.Series[1].Points.AddXY(i, fDi);
                }
                D10 = Math.Exp(par_a) * Math.Exp(0.5 * par_b0 * par_b0);
                Mo = Math.Exp(par_a - par_b0 * par_b0);
                
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N; i++)
                {
                    if (!Ln)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b0 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) / (2 * par_b0 * par_b0)) * Area)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b0 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) / (2 * par_b0 * par_b0)) * Area)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                }
                R2 = 1 - SSres / SStot;
                if (Ln)
                {
                    D10 = Math.Exp(D10);
                    //m = Math.Exp(m);
                    Mo = Math.Exp(Mo);
                }
                textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                textBox4.Text = Math.Round(par_b0, 3).ToString("0.000");
                textBox5.Text = Math.Round(par_b0, 3).ToString("0.000");
                textBox11.Text = Math.Round(D10, 1).ToString("0.0");
                
                textBox10.Text = Math.Round(Mo, 1).ToString("0.0");
            }
            if (Norm)
            {
                for (double i = 0.1; i < X_max + r; i = i + 0.1)
                {
                    fDi = r / (par_b0 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a) * (i - par_a) / (2 * par_b0 * par_b0)) * Area;
                    Program.f11.chart2.Series[1].Points.AddXY(i, fDi);
                }
                
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N; i++)
                {
                    if (!Ln)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b0 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) / (2 * par_b0 * par_b0)) * Area)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b0 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) / (2 * par_b0 * par_b0)) * Area)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                textBox4.Text = Math.Round(par_b0, 3).ToString("0.000");
                textBox5.Text = Math.Round(par_b0, 3).ToString("0.000");
                // textBox3.Text = Math.Round(par_a, 3).ToString("0.000");
                //textBox11.Text = Math.Round(D10, 1).ToString("0.0");
            }
        }

        private void trackBar7_MouseDown(object sender, MouseEventArgs e)
        {
            MD = true;
        }

        private void trackBar7_MouseUp(object sender, MouseEventArgs e)
        {
            MD = false;
            trackBar7.Value = 0;
        }

        private void trackBar7_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = true;
            }
        }

        private void pictureBox4_MouseHover(object sender, EventArgs e)
        {
            t.SetToolTip(pictureBox4, Dcp30);
        }

        private void label37_MouseHover(object sender, EventArgs e)
        {
            t.SetToolTip(label37, median);
        }

        private void label35_MouseHover(object sender, EventArgs e)
        {
            t.SetToolTip(label35, mode);
        }

        private void trackBar6_ValueChanged(object sender, EventArgs e)
        {
            if (!MD && !KD)
            {
                Area = Area0;
                trackBar6.Value = 0;                
            }
            Area0 = Area + Convert.ToDouble(trackBar6.Value)/10;
            Program.f11.chart2.Series[1].Points.Clear();
            if (LogNorm)
            {
                for (double i = 0.1; i < X_max + r; i = i + 0.1)
                {
                    fDi = r / (par_b * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a) * (Math.Log(i) - par_a) / (2 * par_b * par_b)) * Area0;
                    Program.f11.chart2.Series[1].Points.AddXY(i, fDi);
                }
                
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N; i++)
                {
                    if (!Ln)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox9.Text = Math.Round(Area0,1).ToString();
                textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                
            }
            if (Norm)
            {
                for (double i = 0.1; i < X_max + r; i = i + 0.1)
                {
                    fDi = r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a) * (i - par_a) / (2 * par_b * par_b)) * Area0;
                    Program.f11.chart2.Series[1].Points.AddXY(i, fDi);
                }
                
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N; i++)
                {
                    if (!Ln)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                }
                R2 = 1 - SSres / SStot;
                textBox9.Text = Math.Round(Area0,1).ToString();
                textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                
            }
        }

        private void trackBar6_MouseDown(object sender, MouseEventArgs e)
        {
            MD = true;
        }

        private void trackBar6_MouseUp(object sender, MouseEventArgs e)
        {
            MD = false;
            trackBar6.Value = 0;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (TB)
            {
                try
                {
                    par_a = double.Parse(textBox1.Text);
                    Program.f11.chart2.Series[1].Points.Clear();
                    if (LogNorm)
                    {
                        for (double i = 0.1; i < X_max + r; i = i + 0.1)
                        {
                            fDi = r / (par_b * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a) * (Math.Log(i) - par_a) / (2 * par_b * par_b)) * Area;
                            Program.f11.chart2.Series[1].Points.AddXY(i, fDi);
                        }
                        D10 = Math.Exp(par_a) * Math.Exp(0.5 * par_b * par_b);
                        Mo = Math.Exp(par_a - par_b * par_b);
                        m = Math.Exp(par_a);
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N; i++)
                        {
                            if (!Ln)
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) / (2 * par_b * par_b)) * Area)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) / (2 * par_b * par_b)) * Area)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        if (Ln)
                        {
                            D10 = Math.Exp(D10);
                            m = Math.Exp(m);
                            Mo = Math.Exp(Mo);
                        }
                        par_a0 = par_a;                     
                        textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                        textBox3.Text = Math.Round(par_a, 3).ToString("0.000");
                        textBox11.Text = Math.Round(D10, 1).ToString("0.0");
                        textBox23.Text = Math.Round(m, 1).ToString("0.0");
                        textBox10.Text = Math.Round(Mo, 1).ToString("0.0");
                    }
                    if (Norm)
                    {
                        for (double i = 0.1; i < X_max + r; i = i + 0.1)
                        {
                            fDi = r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a) * (i - par_a) / (2 * par_b * par_b)) * Area;
                            Program.f11.chart2.Series[1].Points.AddXY(i, fDi);
                        }
                        D10 = par_a;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N; i++)
                        {
                            if (!Ln)
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) / (2 * par_b * par_b)) * Area)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) / (2 * par_b * par_b)) * Area)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        par_a0 = par_a;
                        textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                        textBox3.Text = Math.Round(par_a, 3).ToString("0.000");
                        if (!Ln)
                            textBox11.Text = Math.Round(D10, 1).ToString("0.0");
                        else
                            textBox23.Text = Math.Round(Math.Exp(D10), 1).ToString("0.0");
                    }
                }
                catch
                {
                    return;
                }
            }
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            TB = true;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (TB)
            {
                try
                {
                    par_b = double.Parse(textBox2.Text);
                    Program.f11.chart2.Series[1].Points.Clear();
                    if (LogNorm)
                    {
                        for (double i = 0.1; i < X_max + r; i = i + 0.1)
                        {
                            fDi = r / (par_b * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a) * (Math.Log(i) - par_a) / (2 * par_b * par_b)) * Area;
                            Program.f11.chart2.Series[1].Points.AddXY(i, fDi);
                        }
                        D10 = Math.Exp(par_a) * Math.Exp(0.5 * par_b * par_b);
                        Mo = Math.Exp(par_a - par_b * par_b);

                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N; i++)
                        {
                            if (!Ln)
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) / (2 * par_b * par_b)) * Area)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) / (2 * par_b * par_b)) * Area)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        if (Ln)
                        {
                            D10 = Math.Exp(D10);
                            //m = Math.Exp(m);
                            Mo = Math.Exp(Mo);
                        }
                        par_b0 = par_b;
                        textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                        textBox4.Text = Math.Round(par_b, 3).ToString("0.000");
                        textBox5.Text = Math.Round(par_b, 3).ToString("0.000");
                        textBox11.Text = Math.Round(D10, 1).ToString("0.0");

                        textBox10.Text = Math.Round(Mo, 1).ToString("0.0");
                    }
                    if (Norm)
                    {
                        for (double i = 0.1; i < X_max + r; i = i + 0.1)
                        {
                            fDi = r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a) * (i - par_a) / (2 * par_b * par_b)) * Area;
                            Program.f11.chart2.Series[1].Points.AddXY(i, fDi);
                        }

                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N; i++)
                        {
                            if (!Ln)
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) / (2 * par_b * par_b)) * Area)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) / (2 * par_b * par_b)) * Area)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                        }
                        par_b0 = par_b;
                        R2 = 1 - SSres / SStot;
                        textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                        textBox4.Text = Math.Round(par_b, 3).ToString("0.000");
                        textBox5.Text = Math.Round(par_b, 3).ToString("0.000");
                    }
                }
                catch
                {
                    return;
                }
            }
        }

        private void textBox2_Enter(object sender, EventArgs e)
        {
            TB = true;
        }

        private void textBox2_Leave(object sender, EventArgs e)
        {
            TB = false;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab==tabControl1.TabPages["TabPage1"])
            {
                flag_lin_raspr = true;
                flag_cub_raspr = false;
                flag_surf_raspr = false;
                flag_vol_surf_raspr = false;
                flag_volweight_raspr = false;
                if (open_f11)
                    f11.Select();
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["TabPage2"])
            {
                flag_lin_raspr = false;
                flag_cub_raspr = true;
                flag_surf_raspr = false;
                flag_vol_surf_raspr = false;
                flag_volweight_raspr = false;
                if (open_f12)
                    f12.Select();
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["TabPage3"])
            {
                flag_lin_raspr = false;
                flag_cub_raspr = false;
                flag_surf_raspr = false;
                flag_vol_surf_raspr = true;
                flag_volweight_raspr = false;
                if (open_f13)
                    f13.Select();
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["TabPage4"])
            {
                flag_lin_raspr = false;
                flag_cub_raspr = false;
                flag_surf_raspr = false;
                flag_vol_surf_raspr = false;
                flag_volweight_raspr = true;
                if (open_f14)
                    f14.Select();
            }
            if (tabControl1.SelectedTab == tabControl1.TabPages["TabPage13"])
            {
                flag_lin_raspr = false;
                flag_cub_raspr = false;
                flag_surf_raspr = true;
                flag_vol_surf_raspr = false;
                flag_volweight_raspr = false;
                if (open_f15)
                    f15.Select();
            }
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            if (TB)
            {
                try
                {
                    Area = double.Parse(textBox9.Text);
                    Program.f11.chart2.Series[1].Points.Clear();
                    if (LogNorm)
                    {
                        for (double i = 0.1; i < X_max + r; i = i + 0.1)
                        {
                            fDi = r / (par_b * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a) * (Math.Log(i) - par_a) / (2 * par_b * par_b)) * Area;
                            Program.f11.chart2.Series[1].Points.AddXY(i, fDi);
                        }

                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N; i++)
                        {
                            if (!Ln)
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) / (2 * par_b * par_b)) * Area)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) / (2 * par_b * par_b)) * Area)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        Area0 = Area;
                        textBox9.Text = Area.ToString();
                        textBox22.Text = Math.Round(R2, 4).ToString("0.0000");

                    }
                    if (Norm)
                    {
                        for (double i = 0.1; i < X_max + r; i = i + 0.1)
                        {
                            fDi = r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a) * (i - par_a) / (2 * par_b * par_b)) * Area;
                            Program.f11.chart2.Series[1].Points.AddXY(i, fDi);
                        }

                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N; i++)
                        {
                            if (!Ln)
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) / (2 * par_b * par_b)) * Area)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) / (2 * par_b * par_b)) * Area)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                        }
                        R2 = 1 - SSres / SStot;
                        Area0 = Area;
                        textBox9.Text = Area.ToString();
                        textBox22.Text = Math.Round(R2, 4).ToString("0.0000");

                    }
                }
                catch
                {
                    return;
                }

            }
        }

        private void textBox9_Enter(object sender, EventArgs e)
        {
            TB = true;
        }

        private void textBox9_Leave(object sender, EventArgs e)
        {
            TB = false;
        }

        private void button6_Click(object sender, EventArgs e) // куб. распр.
        {
            if (Convert.ToString(dataGridView2.Rows[0].Cells[0].Value) == string.Empty)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Размеры частиц не заданы! Попробуйте снова!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles size are not determined! Try again!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (Convert.ToString(dataGridView2.Rows[0].Cells[2].Value) == string.Empty)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Доля частиц не задана! Попробуйте снова!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles fraction are not determined! Try again!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            
            tabControl3.Visible = false;
            if (!open_f12)
            {
                f12 = new Form12();
                f12.Owner = this;
                open_f12 = true;
                f12.Show();
            }
            Program.f12.chart2.Series[0].Points.Clear();
            Program.f12.chart2.Series[1].Points.Clear();
            Program.f12.chart2.Series[2].Points.Clear();
            Program.f12.chart2.Series[5].Points.Clear();
            analyt_func_30 = false;

            N30 = Convert.ToInt32(dataGridView2.RowCount);
            while (Convert.ToString(dataGridView2.Rows[N30 - 1].Cells[0].Value) == string.Empty || Convert.ToString(dataGridView2.Rows[N30 - 1].Cells[2].Value) == string.Empty)
                N30 = N30 - 1;
            
            if (Sum == 0)
            {
                for (int i = 0; i < N30; i++)
                {
                    Sum = Sum + Convert.ToDouble(dataGridView2[1, i].Value); // сначал номер стобца, затем номер строки
                }
            }

            if (Ln30)
            {
                for (int i = 0; i < N30; i++)
                {
                    x = Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value);
                    y = Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value);
                    Program.f12.chart2.Series[0].Points.AddXY(x, y);
                    X_max30 = Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value);

                }
                Program.f12.r = Convert.ToDouble(dataGridView2.Rows[N30 - 1].Cells[3].Value) - Convert.ToDouble(dataGridView2.Rows[N30 - 2].Cells[3].Value);
                X_max30 = X_max30 + Program.f12.r;
                Program.f12.chart2.ChartAreas[0].AxisY.Minimum = 0;
                Program.f12.chart2.ChartAreas[0].AxisX.Maximum = X_max30;
                Program.f12.chart2.ChartAreas[0].AxisX.Minimum = Convert.ToDouble(dataGridView2.Rows[0].Cells[3].Value) - Program.f12.r;
                Program.f12.chart2.ChartAreas[0].AxisX.Title = "Ln(D)";
                Program.f12.chart2.ChartAreas[0].AxisX.Interval = 0;
            }

            else
            {
                for (int i = 0; i < N30; i++)
                {
                    x = Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value);
                    y = Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value);
                    Program.f12.chart2.Series[0].Points.AddXY(x, y);
                    X_max30 = Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value);
                }
                Program.f12.chart2.Series[0].IsVisibleInLegend = true;
                Program.f12.chart2.Series[1].IsVisibleInLegend = false;
                Program.f12.chart2.Series[2].IsVisibleInLegend = false;
                Program.f12.chart2.Series[3].IsVisibleInLegend = false;
                Program.f12.chart2.Series[4].IsVisibleInLegend = false;
                Program.f12.chart2.Series[5].IsVisibleInLegend = false;
                if (Program.f1.Rus == true)
                    Program.f12.chart2.Series[0].Name = "Гист.";
                if (Program.f1.Eng == true)
                    Program.f12.chart2.Series[0].Name = "Hist.";
                Part30 = true;
                r = Convert.ToDouble(dataGridView2.Rows[1].Cells[0].Value) - Convert.ToDouble(dataGridView2.Rows[0].Cells[0].Value);
                X_max30 = X_max30 + r;
                Program.f12.chart2.ChartAreas[0].AxisY.Minimum = 0;
                Program.f12.chart2.ChartAreas[0].AxisX.Maximum = X_max30;
                if (Part30 == true && CSR30 == true)
                    Program.f12.chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), нм";
                else
                    Program.f12.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, нм";
                Program.f12.chart2.ChartAreas[0].AxisX.TitleFont = new Font("Times New Roman", Program.f12.I, FontStyle.Bold);
                if (Part30 == true && CSR30 == true)
                    Program.f12.chart2.ChartAreas[0].AxisY.Title = "Доля частиц (ОКР), %";
                else
                    Program.f12.chart2.ChartAreas[0].AxisY.Title = "Доля частиц, %";
                if (Program.f1.Eng == true)
                    Program.f12.chart2.ChartAreas[0].AxisY.Title = "Frequency, %";
                Program.f12.chart2.ChartAreas[0].AxisY.TitleFont = new Font("Times New Roman", Program.f12.I, FontStyle.Bold);
                Program.f12.chart2.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Times New Roman", Program.f12.I, System.Drawing.FontStyle.Regular);
                Program.f12.chart2.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Times New Roman", Program.f12.I, System.Drawing.FontStyle.Regular);
                if (toolStripComboBox2.SelectedIndex == 0)
                {
                    if (Program.f1.Rus == true)
                    {
                        if (Part30 == true && CSR30 == true)
                            Program.f12.chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), мм (нм)";
                        else
                            Program.f12.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, мм";
                        label33.Text = "мм";
                        label34.Text = "мм";
                        label36.Text = "мм";
                    }
                    if (Program.f1.Eng == true)
                    {
                        if (Part30 == true && CSR30 == true)
                            Program.f12.chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), mm (nm)";
                        else
                            Program.f12.chart2.ChartAreas[0].AxisX.Title = "Particle size, mm";
                        label33.Text = "mm";
                        label34.Text = "mm";
                        label36.Text = "mm";
                    }
                }
                if (toolStripComboBox2.SelectedIndex == 1)
                {
                    if (Program.f1.Rus == true)
                    {
                        if (Part30 == true && CSR30 == true)
                            Program.f12.chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), мкм (нм)";
                        else
                            Program.f12.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, мкм";
                        label33.Text = "мкм";
                        label34.Text = "мкм";
                        label36.Text = "мкм";
                    }
                    if (Program.f1.Eng == true)
                    {
                        if (Part30 == true && CSR30 == true)
                            Program.f12.chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), mkm (nm)";
                        else
                            Program.f12.chart2.ChartAreas[0].AxisX.Title = "Particle size, mkm";
                        label33.Text = "mkm";
                        label34.Text = "mkm";
                        label36.Text = "mkm";
                    }
                }
                if (toolStripComboBox2.SelectedIndex == 2)
                {
                    if (Program.f1.Rus == true)
                    {
                        if (Part30 == true && CSR30 == true)
                            Program.f12.chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), нм";
                        else
                            Program.f12.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, нм";
                        label33.Text = "нм";
                        label34.Text = "нм";
                        label36.Text = "нм";
                    }
                    if (Program.f1.Eng == true)
                    {
                        if (Part30 == true && CSR30 == true)
                            Program.f12.chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), nm";
                        else
                            Program.f12.chart2.ChartAreas[0].AxisX.Title = "Particle size, nm";
                        label33.Text = "nm";
                        label34.Text = "nm";
                        label36.Text = "nm";
                    }


                }


                if (dataGridView2.Rows.Count > 2)
                {
                    X_interval30 = Convert.ToDouble(dataGridView2.Rows[1].Cells[0].Value) - Convert.ToDouble(dataGridView2.Rows[0].Cells[0].Value);
                    X_min30 = Convert.ToDouble(dataGridView2.Rows[0].Cells[0].Value) - X_interval30;
                    Program.f12.chart2.ChartAreas[0].AxisX.Minimum = X_min30;
                    Program.f12.chart2.ChartAreas[0].AxisX.Interval = X_interval30;
                    if (X_min30 <= 0)
                    {
                        Program.f12.chart2.ChartAreas[0].AxisX.IsStartedFromZero = true;
                        for (int i = 0; i < N30; i++)
                            Program.f12.chart2.Series[0].Points[i].AxisLabel = (dataGridView2.Rows[i].Cells[0].Value).ToString();
                    }
                }
            }
            //if (CSR1_30 == true)
           // {
             //   Program.f12.chart2.Series[3].Points.Clear();
              //  for (int i = 0; i < NCSR1; i++)
               // {
                 //   x = crystalSize1[i];
                  //  y = crystalFreq1[i] * 50 * r;
                  //  Program.f11.chart2.Series[3].Points.AddXY(x, y);
              //  }
           // }
           // if (CSR2 == true)
           // {
             //   Program.f11.chart2.Series[4].Points.Clear();
              //  for (int i = 0; i < NCSR2; i++)
              //  {
                 //   x = crystalSize2[i];
                 //   y = crystalFreq2[i] * 50 * r;
                  //  Program.f11.chart2.Series[4].Points.AddXY(x, y);
               // }
           // }
           // if (CSR1 == true && CSR2 == true)
           // {
             //   groupBox2.Visible = true;
             //   trackBar5.Value = 500;
             //   traB5 = trackBar5.Value / 10;
              //  textBox20.Text = Math.Round(traB5, 1).ToString("0.0");
           // }
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            TB = false;
        }

        private void trackBar6_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = true;
            }
        }

        private void trackBar6_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = false;
                trackBar6.Value = 0;
            }
        }

        private void trackBar7_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode == Keys.Left)
            {
                KD = false;
                trackBar7.Value = 0;
            }
        }

        private void trackBar8_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Right || e.KeyCode==Keys.Left)
            {
                KD = true;
            }
        }

        private void trackBar8_MouseDown(object sender, MouseEventArgs e)
        {
            MD = true;
        }

        private void label22_MouseHover(object sender, EventArgs e)
        {
            t.SetToolTip(label22, mode);
        }

        private void textBox22_Enter(object sender, EventArgs e)
        {
            label5.Select();
        }

        private void textBox11_Enter(object sender, EventArgs e)
        {
            label5.Select();
        }

        private void textBox23_Enter(object sender, EventArgs e)
        {
            label5.Select();
        }

        private void textBox10_Enter(object sender, EventArgs e)
        {
            label5.Select();
        }

        private void общаябимодалToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ++clickRaspredCommon;
            switch (clickRaspredCommon)
            {
                case 1:
                    chart2.Series[5].Enabled = false;
                    break;
                case 2:
                    chart2.Series[5].Enabled = true;
                    clickRaspredCommon = 0;
                    break;
            }
        }

        private void распределение1ToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            ++clickCSR1;
            switch (clickCSR1)
            {
                case 1:
                    chart2.Series[3].Enabled = false;
                    break;
                case 2:
                    chart2.Series[3].Enabled = true;
                    clickCSR1 = 0;
                    break;
            }
        }

        private void распределение2ToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            ++clickCSR2;
            switch (clickCSR2)
            {
                case 1:
                    chart2.Series[4].Enabled = false;
                    break;
                case 2:
                    chart2.Series[4].Enabled = true;
                    clickCSR2 = 0;
                    break;
            }
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            I3 = (int)(toolStripComboBox1.SelectedIndex) + 1;
            chart2.Series[1].BorderWidth = I3;
            chart2.Series[2].BorderWidth = I3;
            chart2.Series[3].BorderWidth = I3;
            chart2.Series[4].BorderWidth = I3;
            chart2.Series[5].BorderWidth = I3;
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            ++click_toolStr_button2;
            switch (click_toolStr_button2)
            {
                case 1:
                    toolStripButton2.BackColor = SystemColors.Window;
                    toolStripButton2.Checked = true;
                    chart2.Legends[0].Enabled = false;
                    break;
                case 2:
                    toolStripButton2.BackColor = SystemColors.Control;
                    toolStripButton2.Checked = false;
                    chart2.Legends[0].Enabled = true;
                    click_toolStr_button2 = 0;
                    break;
            }
        }

        private void toolStripComboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            I = (int)(toolStripComboBox4.SelectedIndex) + 8;
            chart2.ChartAreas[0].AxisX.TitleFont = new Font("Times New Roman", I, FontStyle.Bold);
            chart2.ChartAreas[0].AxisY.TitleFont = new Font("Times New Roman", I, FontStyle.Bold);
            chart2.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Times New Roman", I, System.Drawing.FontStyle.Regular);
            chart2.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Times New Roman", I, System.Drawing.FontStyle.Regular);
            chart2.Legends[0].Font = new System.Drawing.Font("Times New Roman", I, System.Drawing.FontStyle.Regular);
        }
      

        private void диапазонРазмеровToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < N; i++)
            {
                double X_X1= Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)-r/2;
                double X_X2= Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)+r/2;
                X_X[i] = Convert.ToString(Math.Round(X_X1,1)) + "-" + Convert.ToString(Math.Round(X_X2, 1));
                Program.f11.chart2.Series[0].Points[i].AxisLabel = X_X[i];
            }
        }

        private void среднийРазмерДиапазонаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < N; i++)
                Program.f11.chart2.Series[0].Points[i].AxisLabel = (dataGridView1.Rows[i].Cells[0].Value).ToString();
        }

        

        private void распределение2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "Excel (*.xls; *.xlsx)|*.xls; *.xlsx";
            if (Program.f1.Rus == true)
                ofd.Title = "Выберите документ для загрузки данных";
            if (Program.f1.Eng == true)
                ofd.Title = "Chose the document to load data";
            ExcelObj.Application app = new ExcelObj.Application();
            ExcelObj.Workbook workbook;
            ExcelObj.Worksheet NwSheet;
            if (ofd.ShowDialog() == DialogResult.OK)
            {

                try
                {
                    workbook = app.Workbooks.Open(ofd.FileName, Missing.Value,
                      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                      Missing.Value);
                    NwSheet = (ExcelObj.Worksheet)workbook.Sheets.get_Item(1);
                    ////ShtRange = NwSheet.UsedRange;
                    var lastCell = NwSheet.Cells.SpecialCells(ExcelObj.XlCellType.xlCellTypeLastCell);
                    for (int Cnum = 1; Cnum <= lastCell.Column; Cnum++)
                    {
                        if (NwSheet.Cells[7, Cnum] != null)
                            crystalSize2[Cnum - 1] = Convert.ToDouble(NwSheet.Cells[7, Cnum].Text);
                        if (NwSheet.Cells[8, Cnum] != null)
                            crystalFreq2[Cnum - 1] = Convert.ToDouble(NwSheet.Cells[8, Cnum].Text);
                        NCSR2 = Cnum;                     
                         if (NwSheet.Cells[7, Cnum] == null)
                            goto M;
                    }
                    M: if (Program.f1.Rus == true)
                        MessageBox.Show("Данные успешно загружены!", "Сообщение!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (Program.f1.Eng == true)
                        MessageBox.Show("Data are imported successfully!!", "Message!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    DCSR2 = Convert.ToDouble(NwSheet.Cells[4, 16].Text);
                    DvCSR2 = Convert.ToDouble(NwSheet.Cells[4, 11].Text);
                    CSR2 = true;
                    app.Quit();
                }
                catch
                {
                    if (Program.f1.Rus == true)
                        MessageBox.Show("Невозможно открыть выбранный файл", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (Program.f1.Eng == true)
                        MessageBox.Show("It is impossible to open the selected file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    app.Quit();
                }

            }
        }

        private void открытьФайлtxtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.txt;*.dat;*.xls;*.xlsx";
            if (Program.f1.Rus == true)
            {
                ofd.Title = "Выберите документ для загрузки данных";
                ofd.Filter = "Текстовый файл (*.txt;*.dat)|*.txt;*.dat|Файлы excel (*.xls; *.xlsx)|*.xls; *.xlsx";
            }
            if (Program.f1.Eng == true)
            {
                ofd.Title = "Chose the document to load data";
                ofd.Filter = "Textfile (*.txt;*.dat)|*.txt;*.dat|Excel file (*.xls; *.xlsx)|*.xls; *.xlsx";
            }
            ExcelObj.Application app = new ExcelObj.Application();
            ExcelObj.Workbook workbook;
            ExcelObj.Worksheet NwSheet;
            ExcelObj.Range ShtRange;
            DataTable dt = new DataTable();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string ext = System.IO.Path.GetExtension(ofd.FileName);
                if (ext == ".txt" || ext == ".dat")
                    try
                    {
                        dataGridView1.Rows.Clear();
                        dataGridView2.Rows.Clear();
                        dataGridView3.Rows.Clear();
                        dataGridView4.Rows.Clear();
                        dataGridView5.Rows.Clear();
                        StreamReader reader = new StreamReader(ofd.FileName);
                        string s = reader.ReadLine();
                        while (s != null)
                        {
                            string[] rvalue = System.Text.RegularExpressions.Regex.Split(s, " ");
                            dataGridView1.Rows.Add(rvalue);
                            s = reader.ReadLine();
                        }
                        dataGridView1.Refresh();
                        X_max = 0;
                        Sum = 0;
                        chart2.Series[0].Points.Clear();
                        chart2.Series[1].Points.Clear();
                        chart2.Series[2].Points.Clear();
                        chart2.Series[5].Points.Clear();
                        richTextBox1.Clear();
                        textBox3.Clear();
                        textBox4.Clear();
                        textBox5.Clear();                     
                        pictureBox2.Visible = false;
                        pictureBox3.Visible = false;
                        //pictureBox4.Visible = false;
                        //textBox11.Visible = false;
                        textBox12.Visible = false;
                        label5.Visible = false;
                        label9.Visible = false;
                        chart2.Series[0].IsVisibleInLegend = false;
                        chart2.Series[1].IsVisibleInLegend = false;
                        chart2.Series[2].IsVisibleInLegend = false;
                        chart2.Series[3].IsVisibleInLegend = false;
                        chart2.Series[4].IsVisibleInLegend = false;
                        textBox22.Visible = false;
                    }
                    catch
                    {
                        if (Program.f1.Rus == true)
                            MessageBox.Show("Невозможно открыть выбранный файл", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        if (Program.f1.Eng == true)
                            MessageBox.Show("It is impossible to open the selected file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                if (ext == ".xls" || ext == ".xlsx")
                    try
                {
                    workbook = app.Workbooks.Open(ofd.FileName, Missing.Value,
                      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                      Missing.Value);
                    NwSheet = (ExcelObj.Worksheet)workbook.Sheets.get_Item(1);
                    ShtRange = NwSheet.UsedRange;
                    dataGridView1.Rows.Clear();
                    dataGridView2.Rows.Clear();
                    dataGridView3.Rows.Clear();
                    dataGridView4.Rows.Clear();
                    dataGridView5.Rows.Clear();
                        for (int Rnum = 1; Rnum <= ShtRange.Rows.Count; Rnum++)
                    {
                        dataGridView1.Rows.Add();
                        for (int Cnum = 1; Cnum <= 3; Cnum++)
                        {
                            if ((ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2 != null)
                                dataGridView1.Rows[Rnum - 1].Cells[Cnum - 1].Value = (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2.ToString();
                        }

                    }
                    dataGridView1.Refresh();
                    X_max = 0;
                    Sum = 0;
                    chart2.Series[0].Points.Clear();
                    chart2.Series[1].Points.Clear();
                    chart2.Series[2].Points.Clear();
                    chart2.Series[5].Points.Clear();
                    richTextBox1.Clear();
                    textBox3.Clear();
                    textBox4.Clear();
                    textBox5.Clear();
                    pictureBox2.Visible = false;
                    pictureBox3.Visible = false;
                    //pictureBox4.Visible = false;
                    //textBox11.Visible = false;
                    textBox12.Visible = false;
                    label5.Visible = false;
                    label9.Visible = false;
                        chart2.Series[0].IsVisibleInLegend = false;
                        chart2.Series[1].IsVisibleInLegend = false;
                        chart2.Series[2].IsVisibleInLegend = false;
                        chart2.Series[3].IsVisibleInLegend = false;
                        chart2.Series[4].IsVisibleInLegend = false;
                        textBox22.Visible = false;
                        app.Quit();
                }
                catch
                {
                    if (Program.f1.Rus == true)
                        MessageBox.Show("Невозможно открыть выбранный файл", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (Program.f1.Eng == true)
                        MessageBox.Show("It is impossible to open the selected file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                
                N = Convert.ToInt32(dataGridView1.RowCount);
               

                while (Convert.ToString(dataGridView1.Rows[N - 1].Cells[1].Value) == string.Empty)
                    N = N - 1;
                for (int i = 0; i < N; i++)
                {
                    Sum = Sum + Convert.ToDouble(dataGridView1[1, i].Value); // сначал номер стобца, затем номер строки
                }


                for (int i = 0; i < N; i++)
                {
                    V = Convert.ToDouble(dataGridView1[1, i].Value) / Sum * 100;
                    dataGridView1.Rows[i].Cells[2].Value = Convert.ToString(Math.Round(V, 2));
                    dataGridView1.Rows[i].Cells[3].Value = Math.Round(Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)), 2);
                }
                dataGridView1.Refresh();
                tabControl2.Visible = false;
                if (open_f11)
                    f11.Close();
                if (open_f12)
                    f12.Close();

            }
        }
     

        private void trackBar5_Scroll(object sender, EventArgs e)
        {
            chart2.Series[3].Points.Clear();
            chart2.Series[4].Points.Clear();
            traB5 = Convert.ToDouble(trackBar5.Value) / 10;
            textBox20.Text = Math.Round(traB5, 1).ToString("0.0");
            for (int i = 0; i < NCSR1; i++)
            {
                x = crystalSize1[i];
                if (Part == true)
                    y = crystalFreq1[i] * traB5*r;
                else
                    y = crystalFreq1[i]*100;
                chart2.Series[3].Points.AddXY(x, y);
            }
            for (int i = 0; i < NCSR2; i++)
            {
                x = crystalSize2[i];
                if (Part == true)
                    y = crystalFreq2[i] * (100-traB5)*r;
                else
                    y = crystalFreq2[i]*100;
                chart2.Series[4].Points.AddXY(x, y);
            }
        }

        private void отобразитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (flag_cub_raspr == true)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Распределение ОКР по размерам недоступно!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Crystal size distribution is not available!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (CSR1 == true || CSR2 == true)
            {
                CSR = true;
                chart2.Series[3].Points.Clear();
                chart2.Series[4].Points.Clear();
                //chart2.ChartAreas[0].AxisY.Minimum = 0;
                //chart2.ChartAreas[0].AxisX.Minimum = 0;
                textBox15.Text = Math.Round(DCSR1, 1).ToString("0.0");
                textBox16.Text = Math.Round(DvCSR1, 1).ToString("0.0");
                if (CSR1 == true)
                {
                    textBox15.Visible = true;
                    textBox16.Visible = true;
                    pictureBox7.Visible = true;
                    pictureBox8.Visible = true;
                    label12.Visible = true;
                    label13.Visible = true;
                    chart2.Series[3].IsVisibleInLegend = true;
                    if (Program.f1.Rus == true)
                        chart2.Series[3].Name = "ОКР 1";
                    if (Program.f1.Eng == true)
                        chart2.Series[3].Name = "CSR 1";
                }
                textBox17.Text = Math.Round(DCSR2, 1).ToString("0.0");
                textBox18.Text = Math.Round(DvCSR2, 1).ToString("0.0");
                if (CSR2 == true)
                {
                    textBox17.Visible = true;
                    textBox18.Visible = true;
                    pictureBox9.Visible = true;
                    pictureBox10.Visible = true;
                    label14.Visible = true;
                    label15.Visible = true;
                    chart2.Series[4].IsVisibleInLegend = true;
                    if (Program.f1.Rus == true)
                        chart2.Series[4].Name = "ОКР 2";
                    if (Program.f1.Eng == true)
                        chart2.Series[4].Name = "CSR 2";
                }
                
                if (CSR == true)
                {
                    chart2.ChartAreas[0].AxisX.Title = "Размер кристаллита, нм";
                    if (Program.f1.Eng == true)
                        chart2.ChartAreas[0].AxisX.Title = "Crystallite size, nm";
                    chart2.ChartAreas[0].AxisX.TitleFont = new Font("Times New Roman", I, FontStyle.Bold);
                    chart2.ChartAreas[0].AxisY.Title = "Доля кристаллитов, %";
                    if (Program.f1.Eng == true)
                        chart2.ChartAreas[0].AxisY.Title = "Frequency, %";
                }
                if (Part == true && CSR == true)
                {
                    chart2.ChartAreas[0].AxisY.Title = "Доля частиц (ОКР), %";
                    if (Program.f1.Eng == true)
                        chart2.ChartAreas[0].AxisY.Title = "Frequency, %";
                    if (toolStripComboBox2.SelectedIndex == 0)
                    {
                        if (Program.f1.Rus == true)
                        {
                            if (Part == true && CSR == true)
                                chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), мм (нм)";
                            else
                                chart2.ChartAreas[0].AxisX.Title = "Размер частиц, мм";
                            label5.Text = "мм";
                            label9.Text = "мм";
                        }
                        if (Program.f1.Eng == true)
                        {
                            if (Part == true && CSR == true)
                                chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), mm (nm)";
                            else
                                chart2.ChartAreas[0].AxisX.Title = "Particle size, mm";
                            label5.Text = "mm";
                            label9.Text = "mm";
                        }
                    }
                    if (toolStripComboBox2.SelectedIndex == 1)
                    {
                        if (Program.f1.Rus == true)
                        {
                            if (Part == true && CSR == true)
                                chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), мкм (нм)";
                            else
                                chart2.ChartAreas[0].AxisX.Title = "Размер частиц, мкм";
                            label5.Text = "мкм";
                            label9.Text = "мкм";
                        }
                        if (Program.f1.Eng == true)
                        {
                            if (Part == true && CSR == true)
                                chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), mkm (nm)";
                            else
                                chart2.ChartAreas[0].AxisX.Title = "Particle size, mkm";
                            label5.Text = "mkm";
                            label9.Text = "mkm";
                        }
                    }
                    if (toolStripComboBox2.SelectedIndex == 2)
                    {
                        if (Program.f1.Rus == true)
                        {
                            if (Part == true && CSR == true)
                                chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), нм";
                            else
                                chart2.ChartAreas[0].AxisX.Title = "Размер частиц, нм";
                            label5.Text = "нм";
                            label9.Text = "нм";
                        }
                        if (Program.f1.Eng == true)
                        {
                            if (Part == true && CSR == true)
                                chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), nm";
                            else
                                chart2.ChartAreas[0].AxisX.Title = "Particle size, nm";
                            label5.Text = "nm";
                            label9.Text = "nm";
                        }


                    }

                }
                chart2.ChartAreas[0].AxisX.TitleFont = new Font("Times New Roman", I, FontStyle.Bold);
                chart2.ChartAreas[0].AxisY.TitleFont = new Font("Times New Roman", I, FontStyle.Bold);
                chart2.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Times New Roman", I, System.Drawing.FontStyle.Regular);
                chart2.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Times New Roman", I, System.Drawing.FontStyle.Regular);
                if (CSR1 == true && CSR2 == true)
                {
                    if (Part == true)
                    {
                        groupBox2.Visible = true;
                        trackBar5.Value = 500;
                        traB5 = trackBar5.Value / 10;
                        textBox20.Text = Math.Round(traB5, 1).ToString("0.0");
                    }
                    for (int i = 0; i < NCSR1; i++)
                    {
                        x = crystalSize1[i];
                        if (Part == true)
                            y = crystalFreq1[i] * 50*r;
                        else
                            y = crystalFreq1[i]*100;
                        chart2.Series[3].Points.AddXY(x, y);
                    }
                    for (int i = 0; i < NCSR2; i++)
                    {
                        x = crystalSize2[i];
                        if (Part == true)
                            y = crystalFreq2[i] * 50*r;
                        else
                            y = crystalFreq2[i]*100;
                        chart2.Series[4].Points.AddXY(x, y);
                    }
                }
                else
                {
                    for (int i = 0; i < NCSR1; i++)
                    {
                        x = crystalSize1[i];
                        if (Part == true)
                            y = crystalFreq1[i]*100 * r;
                        else
                            y = crystalFreq1[i]*100;
                        chart2.Series[3].Points.AddXY(x, y);
                    }
                    for (int i = 0; i < NCSR2; i++)
                    {
                        x = crystalSize2[i];
                        if (Part == true)
                            y = crystalFreq2[i]*100 * r;
                        else
                            y = crystalFreq2[i]*100;
                        chart2.Series[4].Points.AddXY(x, y);
                    }
                }
            }
            else
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Данные по распределению ОКР отсутствуют!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);                
                if (Program.f1.Eng == true)
                    MessageBox.Show("Crystallite size distribution data are missing!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
       

        private void оПрограммеИИнструкцииПоРаботеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.OverwritePrompt = true; // Отображать ли предупреждение, если пользователь указывает имя уже существующего файла
            dialog.CheckPathExists = true; // Отображать ли предупреждение, если пользователь указывает несуществующий путь
            dialog.Filter = "Текст в формате RTF|*.rtf";
            if (Program.f1.Rus == true)
            {
                string resString = Properties.Resources.О_программе_и_инструкции_по_работе_с_ней;
                dialog.FileName = "О программе и инструкции по работе";
                if (dialog.ShowDialog() == DialogResult.OK)
                    File.WriteAllText(dialog.FileName, resString);
            }
            if (Program.f1.Eng == true)
            {
                string resString = Properties.Resources.About_DEAM_and_work_instruction;
                dialog.FileName = "About DEAM and work instruction";
                if (dialog.ShowDialog() == DialogResult.OK)
                    File.WriteAllText(dialog.FileName, resString);
            }
        }

        private void trackBar3_Scroll(object sender, EventArgs e)
        {
            if (button3_Click == true)
            {
                chart2.Series[1].Points.Clear();
                chart2.Series[5].Points.Clear();
                
                for (double i = 0.1; i < X_max + r; i = i + 0.1)
                {
                    if (LogNorm == true)
                    {
                        fDi = r / (par_b1 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a1) * (Math.Log(i) - par_a1) / (2 * par_b1 * par_b1)) * S1;
                        fDiCom = (r / (par_b1 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a1) * (Math.Log(i) - par_a1) / (2 * par_b1 * par_b1)) * S1) + (r / (par_b2 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a2) * (Math.Log(i) - par_a2) / (2 * par_b2 * par_b2)) * S2);
                    }
                    if (Norm == true)
                    {
                        fDi = r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a1) * (i - par_a1) / (2 * par_b1 * par_b1)) * S1;
                        fDiCom = (r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a1) * (i - par_a1) / (2 * par_b1 * par_b1)) * S1) + (r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a2) * (i - par_a2) / (2 * par_b2 * par_b2)) * S2);
                    }
                    chart2.Series[1].Points.AddXY((i + Convert.ToDouble(pos_trackBar1) / 10), fDi);
                    chart2.Series[5].Points.AddXY((i + Convert.ToDouble(pos_trackBar1) / 10), fDiCom);
                    textBox4.Text = Math.Round(par_b1, 3).ToString("0.000");
                    textBox5.Text = Math.Round(par_b1, 3).ToString("0.000");
                }
                R2 = 0;
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N; i++)
                {
                    //if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) != 0)
                    //{
                    if (LogNorm == true)
                    {
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1)) * S1) - (r / (par_b2 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) / (2 * par_b2 * par_b2)) * S2)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - Dcp), 2);
                    }
                    if (Norm == true)
                    {
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1)) * S1) - (r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) / (2 * par_b2 * par_b2)) * S2)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - Dcp), 2);
                    }
                    // }
                }
                R2 = 1 - SSres / SStot;
                textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                if (LogNorm == true)
                {
                    D10 = Math.Exp(par_a1) * Math.Exp(0.5 * par_b1 * par_b1);
                    
                }
            }
            if (button4_Click == true)
            {
                chart2.Series[2].Points.Clear();
                chart2.Series[5].Points.Clear();
                
                for (double i = 0.1; i < X_max + r; i = i + 0.1)
                {
                    if (LogNorm == true)
                    {
                        fDi = r / (par_b2 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a2) * (Math.Log(i) - par_a2) / (2 * par_b2 * par_b2)) * S2;
                        fDiCom = (r / (par_b1 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a1) * (Math.Log(i) - par_a1) / (2 * par_b1 * par_b1)) * S1) + (r / (par_b2 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a2) * (Math.Log(i) - par_a2) / (2 * par_b2 * par_b2)) * S2);
                    }
                    if (Norm == true)
                    {
                        fDi = r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a2) * (i - par_a2) / (2 * par_b2 * par_b2)) * S2;
                        fDiCom = (r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a1) * (i - par_a1) / (2 * par_b1 * par_b1)) * S1) + (r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a2) * (i - par_a2) / (2 * par_b2 * par_b2)) * S2);
                    }
                    chart2.Series[2].Points.AddXY((i + Convert.ToDouble(pos_trackBar2) / 10), fDi);
                    chart2.Series[5].Points.AddXY((i + Convert.ToDouble(pos_trackBar2) / 10), fDiCom);
                    //textBox7.Text = Math.Round(par_b2, 3).ToString("0.000");
                    //textBox8.Text = Math.Round(par_b2, 3).ToString("0.000");
                }
                R2 = 0;
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N; i++)
                {
                    //if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) != 0)
                    //{
                    if (LogNorm == true)
                    {
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1)) * S1) - (r / (par_b2 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) / (2 * par_b2 * par_b2)) * S2)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - Dcp), 2);
                    }
                    if (Norm == true)
                    {
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1)) * S1) - (r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) / (2 * par_b2 * par_b2)) * S2)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - Dcp), 2);
                    }
                    // }
                }
                R2 = 1 - SSres / SStot;
                textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                if (LogNorm == true)
                {
                    D10 = Math.Exp(par_a2) * Math.Exp(0.5 * par_b2 * par_b2);
                    
                }
                
            }
        }

        private void trackBar4_Scroll(object sender, EventArgs e)
        {
            if (button3_Click == true)
            {
                chart2.Series[1].Points.Clear();
                chart2.Series[2].Points.Clear();
                chart2.Series[5].Points.Clear();
                
                S2 = 100 - S1;
                for (double i = 0.1; i < X_max + r; i = i + 0.1)
                {
                    if (LogNorm == true)
                    {
                        fDi = r / (par_b1 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a1) * (Math.Log(i) - par_a1) / (2 * par_b1 * par_b1)) * S1;
                        chart2.Series[1].Points.AddXY((i + Convert.ToDouble(pos_trackBar1) / 10), fDi);
                        fDi = r / (par_b2 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a2) * (Math.Log(i) - par_a2) / (2 * par_b2 * par_b2)) * S2;
                        chart2.Series[2].Points.AddXY((i + Convert.ToDouble(pos_trackBar2) / 10), fDi);
                        fDiCom = (r / (par_b1 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a1) * (Math.Log(i) - par_a1) / (2 * par_b1 * par_b1)) * S1) + (r / (par_b2 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a2) * (Math.Log(i) - par_a2) / (2 * par_b2 * par_b2)) * S2);
                        chart2.Series[5].Points.AddXY((i + Convert.ToDouble(pos_trackBar2) / 10), fDiCom);
                    }
                    if (Norm==true)
                    {
                        fDi = r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a1) * (i - par_a1) / (2 * par_b1 * par_b1)) * S1;
                        chart2.Series[1].Points.AddXY((i + Convert.ToDouble(pos_trackBar1) / 10), fDi);
                        fDi = r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a2) * (i - par_a2) / (2 * par_b2 * par_b2)) * S2;
                        chart2.Series[2].Points.AddXY((i + Convert.ToDouble(pos_trackBar2) / 10), fDi);
                        fDiCom = (r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a1) * (i - par_a1) / (2 * par_b1 * par_b1)) * S1) + (r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a2) * (i - par_a2) / (2 * par_b2 * par_b2)) * S2);
                        chart2.Series[5].Points.AddXY((i + Convert.ToDouble(pos_trackBar2) / 10), fDiCom);
                    }


                }
                R2 = 0;
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N; i++)
                {
                    //if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) != 0)
                    //{
                    if (LogNorm == true)
                    {
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1)) * S1) - (r / (par_b2 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) / (2 * par_b2 * par_b2)) * S2)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - Dcp), 2);
                    }
                    if (Norm == true)
                    {
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1)) * S1) - (r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) / (2 * par_b2 * par_b2)) * S2)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - Dcp), 2);
                    }
                    // }
                }
                R2 = 1 - SSres / SStot;
                textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                
            }
            if (button4_Click==true)
            {
                chart2.Series[1].Points.Clear();
                chart2.Series[2].Points.Clear();
                chart2.Series[5].Points.Clear();
                
                S1 = 100 - S2;
                for (double i = 0.1; i < X_max + r; i = i + 0.1)
                {
                    if (LogNorm == true)
                    {
                        fDi = r / (par_b1 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a1) * (Math.Log(i) - par_a1) / (2 * par_b1 * par_b1)) * S1;
                        chart2.Series[1].Points.AddXY((i + Convert.ToDouble(pos_trackBar1) / 10), fDi);
                        fDi = r / (par_b2 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a2) * (Math.Log(i) - par_a2) / (2 * par_b2 * par_b2)) * S2;
                        chart2.Series[2].Points.AddXY((i + Convert.ToDouble(pos_trackBar2) / 10), fDi);
                        fDiCom = (r / (par_b1 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a1) * (Math.Log(i) - par_a1) / (2 * par_b1 * par_b1)) * S1) + (r / (par_b2 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a2) * (Math.Log(i) - par_a2) / (2 * par_b2 * par_b2)) * S2);
                        chart2.Series[5].Points.AddXY((i + Convert.ToDouble(pos_trackBar2) / 10), fDiCom);
                    }
                    if (Norm == true)
                    {
                        fDi = r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a1) * (i - par_a1) / (2 * par_b1 * par_b1)) * S1;
                        chart2.Series[1].Points.AddXY((i + Convert.ToDouble(pos_trackBar1) / 10), fDi);
                        fDi = r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a2) * (i - par_a2) / (2 * par_b2 * par_b2)) * S2;
                        chart2.Series[2].Points.AddXY((i + Convert.ToDouble(pos_trackBar2) / 10), fDi);
                        fDiCom = (r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a1) * (i - par_a1) / (2 * par_b1 * par_b1)) * S1) + (r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a2) * (i - par_a2) / (2 * par_b2 * par_b2)) * S2);
                        chart2.Series[5].Points.AddXY((i + Convert.ToDouble(pos_trackBar2) / 10), fDiCom);
                    }
                }
                R2 = 0;
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N; i++)
                {
                    //if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) != 0)
                    //{
                    if (LogNorm == true)
                    {
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1)) * S1) - (r / (par_b2 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) / (2 * par_b2 * par_b2)) * S2)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - Dcp), 2);
                    }
                    if (Norm == true)
                    {
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1)) * S1) - (r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) / (2 * par_b2 * par_b2)) * S2)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - Dcp), 2);
                    }
                    // }
                }
                R2 = 1 - SSres / SStot;
                textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                
            }
        }
      

        


            

        private void button3_Click_1(object sender, EventArgs e)
        {
            button3_Click = true;
            button4_Click = false;
            
            if (LogNorm == true)
            {
                D10 = Math.Exp(par_a1) * Math.Exp(0.5 * par_b1 * par_b1);
                
            }
            if (Norm==true)
            {
                D10 = par_a1;
                
            }
            
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            button3_Click = false;
            button4_Click = true;
            
            if (LogNorm == true)
            {
                D10 = Math.Exp(par_a2) * Math.Exp(0.5 * par_b2 * par_b2);
                
            }
            if (Norm == true)
            {
                D10 = par_a2;
                
            }
            
        }
            

        private void toolStripComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (toolStripComboBox2.SelectedIndex == 0)
            {
                if (Program.f1.Rus == true)
                {
                    if (CSR == true)
                        chart2.ChartAreas[0].AxisX.Title = "Размер кристаллита, нм";                    
                    if (Part == true)
                        chart2.ChartAreas[0].AxisX.Title = "Размер частиц, мм";
                    if (Part == true && CSR == true)
                        chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), мм (нм)";
                    dataGridView1.Columns[0].HeaderCell.Value = "Размер частиц, мм";
                    tSCB2 = toolStripComboBox2.SelectedIndex;
                    label5.Text = "мм";
                    label9.Text = "мм";
                }
                if (Program.f1.Eng==true)
                {
                    if (CSR == true)
                        chart2.ChartAreas[0].AxisX.Title = "Crystallite size, nm";                   
                    if (Part == true)
                        chart2.ChartAreas[0].AxisX.Title = "Particle size, mm";
                    if (Part == true && CSR == true)
                        chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), mm (nm)";
                    dataGridView1.Columns[0].HeaderCell.Value = "Particle size, mm";
                    tSCB2 = toolStripComboBox2.SelectedIndex;
                    label5.Text = "mm";
                    label9.Text = "mm";
                }
            }
            if (toolStripComboBox2.SelectedIndex == 1)
            {
                if (Program.f1.Rus == true)
                {
                    if (CSR == true)
                        chart2.ChartAreas[0].AxisX.Title = "Размер кристаллита, нм";                    
                    if (Part == true)
                        chart2.ChartAreas[0].AxisX.Title = "Размер частиц, мкм";
                    if (Part == true && CSR == true)
                        chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), мкм (нм)";
                    dataGridView1.Columns[0].HeaderCell.Value = "Размер частиц, мкм";
                    tSCB2 = toolStripComboBox2.SelectedIndex;
                    label5.Text = "мкм";
                    label9.Text = "мкм";
                }
                if (Program.f1.Eng==true)
                {
                    if (CSR == true)
                        chart2.ChartAreas[0].AxisX.Title = "Crystallite size, nm";                   
                    if (Part == true)
                        chart2.ChartAreas[0].AxisX.Title = "Particle size, mkm";
                    if (Part == true && CSR == true)
                        chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), mkm (nm)";
                    dataGridView1.Columns[0].HeaderCell.Value = "Particle size, mkm";
                    tSCB2 = toolStripComboBox2.SelectedIndex;
                    label5.Text = "mkm";
                    label9.Text = "mkm";
                }
            }
            if (toolStripComboBox2.SelectedIndex == 2)
            {
                if (Program.f1.Rus == true)
                {
                    if (CSR == true)
                        chart2.ChartAreas[0].AxisX.Title = "Размер кристаллита, нм";                    
                    if (Part == true)
                        chart2.ChartAreas[0].AxisX.Title = "Размер частиц, нм";
                    if (Part == true && CSR == true)
                        chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), нм";
                    dataGridView1.Columns[0].HeaderCell.Value = "Размер частиц, нм";
                    tSCB2 = toolStripComboBox2.SelectedIndex;
                    label5.Text = "нм";
                    label9.Text = "нм";
                }
                if (Program.f1.Eng==true)
                {
                    if (CSR == true)
                        chart2.ChartAreas[0].AxisX.Title = "Crystallite size, nm";                   
                    if (Part == true)
                        chart2.ChartAreas[0].AxisX.Title = "Particle size, nm";
                    if (Part == true && CSR == true)
                        chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), nm";
                    dataGridView1.Columns[0].HeaderCell.Value = "Particle size, nm";
                    tSCB2 = toolStripComboBox2.SelectedIndex;
                    label5.Text = "nm";
                    label9.Text = "nm";
                }
            }
            

        }

        private void сохранитьИзображениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            if (Program.f1.Eng == true)
                sfd.Title = "Save";
            sfd.OverwritePrompt = true; // Отображать ли предупреждение, если пользователь указывает имя уже существующего файла
            sfd.CheckPathExists = true; // Отображать ли предупреждение, если пользователь указывает несуществующий путь
            sfd.Filter = "*.jpg|*.jpg|;*.bmp|*.bmp;|*.png|*.png";
            sfd.AddExtension = true;
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                switch (sfd.FilterIndex)
                {
                    case 1: chart2.SaveImage(sfd.FileName, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Bmp); break;
                    case 2: chart2.SaveImage(sfd.FileName, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Png); break;
                    case 3: chart2.SaveImage(sfd.FileName, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Jpeg); break;
                }
            }

        }

        public void Log_Norm_10()
        {
            if (dataGridView1.Rows[0].Cells[2].Value == null)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Значения таблицы не заданы! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("The table values are missing! Try again", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (dataGridView1.Rows[0].Cells[1].Value == null)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Число частиц не задано! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles number are not determined! Try again!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (!open_f11)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Сначала отобразите гистограмму на графике!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Plot the histogram first!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            par_a0 = 0;
            par_b0 = 0;
            Area = 100;
            Area0 = 100;
            LogNorm = true;
            Norm = false;
            tabControl2.Visible = true;
            tabControl2.TabPages[0].Text = "Лог-норм";
            label23.Text = "ln(μ)";
            if (open_f11)
            {
                Program.f11.chart2.Series[1].Points.Clear();
                Program.f11.chart2.Series[2].Points.Clear();
                Program.f11.chart2.Series[5].Points.Clear();
            }
            richTextBox1.Clear();
            pictureBox2.Visible = true;
            textBox11.Visible = true;
            label5.Visible = true;
            label9.Visible = true;
            textBox23.Visible = true;
            label21.Visible = true;
            textBox10.Visible = true;
            label20.Visible = true;
            label19.Visible = true;
            label22.Visible = true;
            textBox22.Visible = true;
            textBox22.Location = new Point(226, 58);
            R2 = 0;
            SSres = 0;
            SStot = 0;
            sumni = 0;
            S_ni_lndi = 0;
            sum1 = 0;
            sum_Y = 0;

            for (int i = 0; i < N; i++)
            {
                //sumni = sumni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                if (!Ln)
                    S_ni_lndi = S_ni_lndi + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) * Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value));
                else
                    S_ni_lndi = S_ni_lndi + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) * Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value));
                sum_Y = sum_Y + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
            }
            //par_a = S_ni_lndi / sumni;
            par_a = S_ni_lndi / 100;

            for (int i = 0; i < N; i++)
            {
                if (!Ln)
                    sum1 = sum1 + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a);
                else
                    sum1 = sum1 + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value)) - par_a) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value)) - par_a);
            }
            //par_b = Math.Sqrt(sum1 / (sumni - 1));
            par_b = Math.Sqrt(sum1 / 100);

            for (int i = 0; i < N; i++)
            {
                if (!Ln)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) / (2 * par_b * par_b)) * 100)), 2);
                else
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) / (2 * par_b * par_b)))), 2);
                SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
            }
            R2 = 1 - SSres / SStot;
            if (Ln)
            {
                Area0 = 1;
                R02 = 0;
                while (R02 < R2 || R2 < 0)
                {
                    R02 = R2;
                    Area0 = Area0 + 0.1;
                    SSres = 0;
                    SStot = 0;
                    for (int i = 0; i < N; i++)
                    {
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                    }
                    R2 = 1 - SSres / SStot;
                }
            }
            R02 = R2;
            for (int j = 0; j < 3; j++)
            {
                par_b = par_b - 0.001; // сужение
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N; i++)
                {
                    if (!Ln)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_b = par_b + 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_b = par_b - 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N; i++)
                        {
                            if (!Ln)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_b = par_b + 0.001; // расширение
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N; i++)
                {
                    if (!Ln)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_b = par_b - 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_b = par_b + 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N; i++)
                        {
                            if (!Ln)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_a = par_a - 0.001; // шаг влево
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N; i++)
                {
                    if (!Ln)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_a = par_a + 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_a = par_a - 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N; i++)
                        {
                            if (!Ln)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_a = par_a + 0.001; // шаг вправо
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N; i++)
                {
                    if (!Ln)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_a = par_a - 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_a = par_a + 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N; i++)
                        {
                            if (!Ln)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }
            }

            Area0 = Area0 - 0.1; // уменьшение площади
            SSres = 0;
            SStot = 0;
            for (int i = 0; i < N; i++)
            {
                if (!Ln)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                else
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
            }
            R2 = 1 - SSres / SStot;
            if (R02 > R2)
                Area0 = Area0 + 0.1;
            else
                while (R02 < R2)
                {
                    R02 = R2;
                    Area0 = Area0 - 0.1;
                    SSres = 0;
                    SStot = 0;
                    for (int i = 0; i < N; i++)
                    {
                        if (!Ln)
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                        else
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                    }
                    R2 = 1 - SSres / SStot;
                }

            Area0 = Area0 + 0.1; // увеличение площади
            SSres = 0;
            SStot = 0;
            for (int i = 0; i < N; i++)
            {
                if (!Ln)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                else
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
            }
            R2 = 1 - SSres / SStot;
            if (R02 > R2)
                Area0 = Area0 - 0.1;
            else
                while (R02 < R2)
                {
                    R02 = R2;
                    Area0 = Area0 + 0.1;
                    SSres = 0;
                    SStot = 0;
                    for (int i = 0; i < N; i++)
                    {
                        if (!Ln)
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                        else
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) * (Math.Log((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value))) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                    }
                    R2 = 1 - SSres / SStot;
                }

            R2 = R02;
            Area = Area0;
            par_a0 = par_a;
            par_b0 = par_b;
            for (double i = 0.1; i < X_max + r; i = i + 0.1)
            {
                fDi = r / (par_b * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a) * (Math.Log(i) - par_a) / (2 * par_b * par_b)) * Area0;
                Program.f11.chart2.Series[1].Points.AddXY(i, fDi);
            }


            D10 = Math.Exp(par_a) * Math.Exp(0.5 * par_b * par_b);
            if (Ln)
                D10 = Math.Exp(D10);
            Mo = Math.Exp(par_a - par_b * par_b);
            if (Ln)
                Mo = Math.Exp(Mo);
            m = Math.Exp(par_a);
            if (Ln)
                m = Math.Exp(m);


            textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
            richTextBox1.SelectedRtf = Properties.Resources.log_norm;
            richTextBox1.Height = 81;
            textBox4.Text = Math.Round(par_b, 3).ToString("0.000");
            textBox3.Text = Math.Round(par_a, 3).ToString("0.000");
            textBox5.Text = Math.Round(par_b, 3).ToString("0.000");
            textBox5.Location = new Point(225, 31);
            textBox3.Location = new Point(239, 8);
            textBox11.Text = Math.Round(D10, 1).ToString("0.0");
            textBox23.Text = Math.Round(m, 1).ToString("0.0");
            textBox10.Text = Math.Round(Mo, 1).ToString("0.0");
            textBox1.Text = Math.Round(par_a, 3).ToString("0.000");
            textBox2.Text = Math.Round(par_b, 3).ToString("0.000");
            textBox9.Text = Math.Round(Area0, 1).ToString();
            analyt_func = true;

            Program.f11.chart2.Series[2].IsVisibleInLegend = false;
            Program.f11.chart2.Series[5].IsVisibleInLegend = false;
            Program.f11.chart2.Series[1].IsVisibleInLegend = true;
            if (Program.f1.Rus == true)
                Program.f11.chart2.Series[1].Name = "Лог-норм 1";
            if (Program.f1.Eng == true)
                Program.f11.chart2.Series[1].Name = "Log-Norm 1";
            
        }

        private void Log_Norm_30()
        {
            if (dataGridView2.Rows[0].Cells[2].Value == null)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Значения таблицы не заданы! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("The table values are missing! Try again", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (dataGridView2.Rows[0].Cells[1].Value == null)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Число частиц не задано! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles number are not determined! Try again!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (!open_f12)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Сначала отобразите гистограмму на графике!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Plot the histogram first!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            par_a0_30 = 0;
            par_b0_30 = 0;
            Area_30 = 100;
            Area0_30 = 100;
            LogNorm_30 = true;
            Norm_30 = false;
            tabControl3.Visible = true;
            tabControl3.TabPages[0].Text = "Лог-норм";
            label28.Text = "ln(μ)";
            if (open_f12)
            {
                Program.f12.chart2.Series[1].Points.Clear();
                Program.f12.chart2.Series[2].Points.Clear();
                Program.f12.chart2.Series[5].Points.Clear();
            }
            richTextBox2.Clear();
            pictureBox4.Visible = true;
            //textBox11.Visible = true;
            //textBox12.Visible = true;
            label37.Visible = true;
            label35.Visible = true;
            textBox29.Visible = true;
            label34.Visible = true;
            textBox28.Visible = true;
            label36.Visible = true;
            label33.Visible = true;
            //label19.Visible = true;
            //label22.Visible = true;
            textBox27.Visible = true;
            textBox31.Location = new Point(226, 58);
            R2 = 0;
            SSres = 0;
            SStot = 0;
            sumni = 0;
            S_ni_lndi = 0;
            sum1 = 0;
            sum_Y_30 = 0;
            for (int i = 0; i < N30; i++)
            {
                //sumni = sumni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                if (!Ln30)
                    S_ni_lndi = S_ni_lndi + Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) * Math.Log(Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value));
                else
                    S_ni_lndi = S_ni_lndi + Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) * Math.Log(Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value));
                sum_Y_30 = sum_Y_30 + Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value);
            }
            //par_a = S_ni_lndi / sumni;
            par_a_30 = S_ni_lndi / 100;

            for (int i = 0; i < N30; i++)
            {
                if (!Ln30)
                    sum1 = sum1 + Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) * (Math.Log(Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)) - par_a_30) * (Math.Log(Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)) - par_a_30);
                else
                    sum1 = sum1 + Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) * (Math.Log(Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)) - par_a_30) * (Math.Log(Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)) - par_a_30);
            }
            //par_b = Math.Sqrt(sum1 / (sumni - 1));
            par_b_30 = Math.Sqrt(sum1 / 100);

            for (int i = 0; i < N30; i++)
            {
                if (!Ln30)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * 100)), 2);
                else
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)))), 2);
                SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
            }
            R2 = 1 - SSres / SStot;

            if (Ln30)
            {
                Area0_30 = 1;
                R02 = 0;
                while (R02 < R2 || R2 < 0)
                {
                    R02 = R2;
                    Area0_30 = Area0_30 + 0.1;
                    SSres = 0;
                    SStot = 0;
                    for (int i = 0; i < N30; i++)
                    {
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                    }
                    R2 = 1 - SSres / SStot;
                }
            }
            R02 = R2;
            for (int j = 0; j < 3; j++)
            {
                par_b_30 = par_b_30 - 0.001; // сужение
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N30; i++)
                {
                    if (!Ln30)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * 100)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_b_30 = par_b_30 + 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_b_30 = par_b_30 - 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N30; i++)
                        {
                            if (!Ln30)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * 100)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_b_30 = par_b_30 + 0.001; // расширение
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N30; i++)
                {
                    if (!Ln30)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * 100)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_b_30 = par_b_30 - 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_b_30 = par_b_30 + 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N30; i++)
                        {
                            if (!Ln30)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * 100)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_a_30 = par_a_30 - 0.001; // шаг влево
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N30; i++)
                {
                    if (!Ln30)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * 100)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_a_30 = par_a_30 + 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_a_30 = par_a_30 - 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N30; i++)
                        {
                            if (!Ln30)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * 100)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_a_30 = par_a_30 + 0.001; // шаг вправо
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N30; i++)
                {
                    if (!Ln30)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * 100)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_a_30 = par_a_30 - 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_a_30 = par_a_30 + 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N30; i++)
                        {
                            if (!Ln30)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * 100)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

            }
            Area0_30 = Area0_30 - 0.1; // уменьшение площади
            SSres = 0;
            SStot = 0;
            for (int i = 0; i < N30; i++)
            {
                if (!Ln30)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                else
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
            }
            R2 = 1 - SSres / SStot;
            if (R02 > R2)
                Area0_30 = Area0_30 + 0.1;
            else
                while (R02 < R2)
                {
                    R02 = R2;
                    Area0_30 = Area0_30 - 0.1;
                    SSres = 0;
                    SStot = 0;
                    for (int i = 0; i < N30; i++)
                    {
                        if (!Ln30)
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                        else
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                    }
                    R2 = 1 - SSres / SStot;
                }

            Area0_30 = Area0_30 + 0.1; // увеличение площади
            SSres = 0;
            SStot = 0;
            for (int i = 0; i < N30; i++)
            {
                if (!Ln30)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                else
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
            }
            R2 = 1 - SSres / SStot;
            if (R02 > R2)
                Area0_30 = Area0_30 - 0.1;
            else
                while (R02 < R2)
                {
                    R02 = R2;
                    Area0_30 = Area0_30 + 0.1;
                    SSres = 0;
                    SStot = 0;
                    for (int i = 0; i < N30; i++)
                    {
                        if (!Ln30)
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                        else
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) * (Math.Log((Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                    }
                    R2 = 1 - SSres / SStot;
                }

            R2 = R02;
            Area_30 = Area0_30;
            par_a0_30 = par_a_30;
            par_b0_30 = par_b_30;

            for (double i = 0.1; i < X_max30 + r; i = i + 0.1)
            {
                fDi = r / (par_b_30 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_30) * (Math.Log(i) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area_30;
                Program.f12.chart2.Series[1].Points.AddXY(i, fDi);
            }
            D30 = Math.Exp(par_a_30) * Math.Exp(0.5 * par_b_30 * par_b_30);
            Mo = Math.Exp(par_a_30 - par_b_30 * par_b_30);
            m = Math.Exp(par_a_30);
            if (Ln30)
            {
                D30 = Math.Exp(D30);
                Mo = Math.Exp(Mo);
                m = Math.Exp(m);
            }

            textBox31.Text = Math.Round(R2, 4).ToString("0.0000");
            richTextBox2.SelectedRtf = Properties.Resources.log_norm;
            richTextBox2.Height = 81;
            textBox34.Text = Math.Round(par_b_30, 3).ToString("0.000");
            textBox30.Text = Math.Round(par_a_30, 3).ToString("0.000");
            textBox32.Text = Math.Round(par_b_30, 3).ToString("0.000");
            textBox32.Location = new Point(225, 31);
            textBox30.Location = new Point(239, 8);
            textBox27.Text = Math.Round(D30, 1).ToString("0.0");
            textBox28.Text = Math.Round(m, 1).ToString("0.0");
            textBox29.Text = Math.Round(Mo, 1).ToString("0.0");
            textBox26.Text = Math.Round(par_a_30, 3).ToString("0.000");
            textBox25.Text = Math.Round(par_b_30, 3).ToString("0.000");
            textBox24.Text = Math.Round(Area_30, 1).ToString();
            analyt_func_30 = true;

            Program.f12.chart2.Series[2].IsVisibleInLegend = false;
            Program.f12.chart2.Series[5].IsVisibleInLegend = false;
            Program.f12.chart2.Series[1].IsVisibleInLegend = true;
            if (Program.f1.Rus == true)
                Program.f12.chart2.Series[1].Name = "Лог-норм 1";
            if (Program.f1.Eng == true)
                Program.f12.chart2.Series[1].Name = "Log-Norm 1";
        }

        private void Log_Norm_32()
        {
            if (dataGridView3.Rows[0].Cells[2].Value == null)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Значения таблицы не заданы! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("The table values are missing! Try again", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (dataGridView3.Rows[0].Cells[1].Value == null)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Число частиц не задано! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles number are not determined! Try again!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (!open_f13)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Сначала отобразите гистограмму на графике!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Plot the histogram first!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            par_a0_32 = 0;
            par_b0_32 = 0;
            Area_32 = 100;
            Area0_32 = 100;
            LogNorm_32 = true;
            Norm_32 = false;
            tabControl4.Visible = true;
            tabControl4.TabPages[0].Text = "Лог-норм";
            label42.Text = "ln(μ)";
            if (open_f13)
            {
                Program.f13.chart2.Series[1].Points.Clear();
                Program.f13.chart2.Series[2].Points.Clear();
                Program.f13.chart2.Series[5].Points.Clear();
            }
            richTextBox3.Clear();
            pictureBox11.Visible = true;
            //textBox11.Visible = true;
            //textBox12.Visible = true;
            label51.Visible = true;
            label47.Visible = true;
            textBox40.Visible = true;
            label48.Visible = true;
            textBox41.Visible = true;
            label49.Visible = true;
            label50.Visible = true;
            //label19.Visible = true;
            //label22.Visible = true;
            textBox42.Visible = true;
            textBox44.Location = new Point(226, 58);
            R2 = 0;
            SSres = 0;
            SStot = 0;
            sumni = 0;
            S_ni_lndi = 0;
            sum1 = 0;
            sum_Y_32 = 0;
            for (int i = 0; i < N32; i++)
            {
                //sumni = sumni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                if (!Ln32)
                    S_ni_lndi = S_ni_lndi + Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) * Math.Log(Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value));
                else
                    S_ni_lndi = S_ni_lndi + Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) * Math.Log(Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value));
                sum_Y_32 = sum_Y_32 + Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value);
            }
            //par_a = S_ni_lndi / sumni;
            par_a_32 = S_ni_lndi / 100;

            for (int i = 0; i < N32; i++)
            {
                if (!Ln32)
                    sum1 = sum1 + Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) * (Math.Log(Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value)) - par_a_32) * (Math.Log(Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value)) - par_a_32);
                else
                    sum1 = sum1 + Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) * (Math.Log(Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value)) - par_a_32) * (Math.Log(Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value)) - par_a_32);
            }
            //par_b = Math.Sqrt(sum1 / (sumni - 1));
            par_b_32 = Math.Sqrt(sum1 / 100);

            for (int i = 0; i < N32; i++)
            {
                if (!Ln32)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * 100)), 2);
                else
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)))), 2);
                SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
            }
            R2 = 1 - SSres / SStot;

            if (Ln32)
            {
                Area0_32 = 1;
                R02 = 0;
                while (R02 < R2 || R2 < 0)
                {
                    R02 = R2;
                    Area0_32 = Area0_32 + 0.1;
                    SSres = 0;
                    SStot = 0;
                    for (int i = 0; i < N32; i++)
                    {
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                    }
                    R2 = 1 - SSres / SStot;
                }
            }
            R02 = R2;
            for (int j = 0; j < 3; j++)
            {
                par_b_32 = par_b_32 - 0.001; // сужение
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N32; i++)
                {
                    if (!Ln32)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * 100)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_b_32 = par_b_32 + 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_b_32 = par_b_32 - 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N32; i++)
                        {
                            if (!Ln32)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * 100)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_b_32 = par_b_32 + 0.001; // расширение
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N32; i++)
                {
                    if (!Ln32)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * 100)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_b_32 = par_b_32 - 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_b_32 = par_b_32 + 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N32; i++)
                        {
                            if (!Ln32)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * 100)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_a_32 = par_a_32 - 0.001; // шаг влево
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N32; i++)
                {
                    if (!Ln32)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * 100)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_a_32 = par_a_32 + 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_a_32 = par_a_32 - 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N32; i++)
                        {
                            if (!Ln32)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * 100)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_a_32 = par_a_32 + 0.001; // шаг вправо
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N32; i++)
                {
                    if (!Ln32)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * 100)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_a_32 = par_a_32 - 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_a_32 = par_a_32 + 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N32; i++)
                        {
                            if (!Ln32)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * 100)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }
            }

            Area0_32 = Area0_32 - 0.1; // уменьшение площади
            SSres = 0;
            SStot = 0;
            for (int i = 0; i < N32; i++)
            {
                if (!Ln32)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                else
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
            }
            R2 = 1 - SSres / SStot;
            if (R02 > R2)
                Area0_32 = Area0_32 + 0.1;
            else
                while (R02 < R2)
                {
                    R02 = R2;
                    Area0_32 = Area0_32 - 0.1;
                    SSres = 0;
                    SStot = 0;
                    for (int i = 0; i < N32; i++)
                    {
                        if (!Ln32)
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                        else
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                    }
                    R2 = 1 - SSres / SStot;
                }

            Area0_32 = Area0_32 + 0.1; // увеличение площади
            SSres = 0;
            SStot = 0;
            for (int i = 0; i < N32; i++)
            {
                if (!Ln32)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                else
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
            }
            R2 = 1 - SSres / SStot;
            if (R02 > R2)
                Area0_32 = Area0_32 - 0.1;
            else
                while (R02 < R2)
                {
                    R02 = R2;
                    Area0_32 = Area0_32 + 0.1;
                    SSres = 0;
                    SStot = 0;
                    for (int i = 0; i < N32; i++)
                    {
                        if (!Ln32)
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                        else
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) * (Math.Log((Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value))) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                    }
                    R2 = 1 - SSres / SStot;
                }

            R2 = R02;
            Area_32 = Area0_32;
            par_a0_32 = par_a_32;
            par_b0_32 = par_b_32;

            for (double i = 0.1; i < X_max32 + r; i = i + 0.1)
            {
                fDi = r / (par_b_32 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_32) * (Math.Log(i) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area_32;
                Program.f13.chart2.Series[1].Points.AddXY(i, fDi);
            }
            D32 = Math.Exp(par_a_32) * Math.Exp(0.5 * par_b_32 * par_b_32);
            Mo = Math.Exp(par_a_32 - par_b_32 * par_b_32);
            m = Math.Exp(par_a_32);
            if (Ln32)
            {
                D32 = Math.Exp(D32);
                Mo = Math.Exp(Mo);
                m = Math.Exp(m);
            }

            textBox44.Text = Math.Round(R2, 4).ToString("0.0000");
            richTextBox3.SelectedRtf = Properties.Resources.log_norm;
            richTextBox3.Height = 81;
            textBox47.Text = Math.Round(par_b_32, 3).ToString("0.000");
            textBox43.Text = Math.Round(par_a_32, 3).ToString("0.000");
            textBox45.Text = Math.Round(par_b_32, 3).ToString("0.000");
            textBox45.Location = new Point(225, 31);
            textBox43.Location = new Point(239, 8);
            textBox40.Text = Math.Round(D32, 1).ToString("0.0");
            textBox41.Text = Math.Round(m, 1).ToString("0.0");
            textBox42.Text = Math.Round(Mo, 1).ToString("0.0");
            textBox39.Text = Math.Round(par_a_32, 3).ToString("0.000");
            textBox38.Text = Math.Round(par_b_32, 3).ToString("0.000");
            textBox37.Text = Math.Round(Area_32, 1).ToString();
            analyt_func_32 = true;

            Program.f13.chart2.Series[2].IsVisibleInLegend = false;
            Program.f13.chart2.Series[5].IsVisibleInLegend = false;
            Program.f13.chart2.Series[1].IsVisibleInLegend = true;
            if (Program.f1.Rus == true)
                Program.f13.chart2.Series[1].Name = "Лог-норм 1";
            if (Program.f1.Eng == true)
                Program.f13.chart2.Series[1].Name = "Log-Norm 1";
        }

        private void Log_Norm_43()
        {
            if (dataGridView4.Rows[0].Cells[2].Value == null)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Значения таблицы не заданы! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("The table values are missing! Try again", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (dataGridView4.Rows[0].Cells[1].Value == null)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Число частиц не задано! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles number are not determined! Try again!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (!open_f14)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Сначала отобразите гистограмму на графике!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Plot the histogram first!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            par_a0_43 = 0;
            par_b0_43 = 0;
            Area_43 = 100;
            Area0_43 = 100;
            LogNorm_43 = true;
            Norm_43 = false;
            tabControl5.Visible = true;
            tabControl5.TabPages[0].Text = "Лог-норм";
            label56.Text = "ln(μ)";
            if (open_f14)
            {
                Program.f14.chart2.Series[1].Points.Clear();
                Program.f14.chart2.Series[2].Points.Clear();
                Program.f14.chart2.Series[5].Points.Clear();
            }
            richTextBox4.Clear();
            pictureBox12.Visible = true;
            //textBox11.Visible = true;
            //textBox12.Visible = true;
            label65.Visible = true;
            label63.Visible = true;
            textBox53.Visible = true;
            label61.Visible = true;
            textBox54.Visible = true;
            label62.Visible = true;
            label64.Visible = true;
            //label19.Visible = true;
            //label22.Visible = true;
            textBox55.Visible = true;
            textBox57.Location = new Point(226, 58);
            R2 = 0;
            SSres = 0;
            SStot = 0;
            sumni = 0;
            S_ni_lndi = 0;
            sum1 = 0;
            sum_Y_43 = 0;
            for (int i = 0; i < N43; i++)
            {
                //sumni = sumni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                if (!Ln43)
                    S_ni_lndi = S_ni_lndi + Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) * Math.Log(Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value));
                else
                    S_ni_lndi = S_ni_lndi + Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) * Math.Log(Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value));
                sum_Y_43 = sum_Y_43 + Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value);
            }
            //par_a = S_ni_lndi / sumni;
            par_a_43 = S_ni_lndi / 100;

            for (int i = 0; i < N43; i++)
            {
                if (!Ln43)
                    sum1 = sum1 + Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) * (Math.Log(Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value)) - par_a_43) * (Math.Log(Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value)) - par_a_43);
                else
                    sum1 = sum1 + Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) * (Math.Log(Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value)) - par_a_43) * (Math.Log(Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value)) - par_a_43);
            }
            //par_b = Math.Sqrt(sum1 / (sumni - 1));
            par_b_43 = Math.Sqrt(sum1 / 100);

            for (int i = 0; i < N43; i++)
            {
                if (!Ln43)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * 100)), 2);
                else
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)))), 2);
                SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
            }
            R2 = 1 - SSres / SStot;

            if (Ln43)
            {
                Area0_43 = 1;
                R02 = 0;
                while (R02 < R2 || R2 < 0)
                {
                    R02 = R2;
                    Area0_43 = Area0_43 + 0.1;
                    SSres = 0;
                    SStot = 0;
                    for (int i = 0; i < N43; i++)
                    {
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                    }
                    R2 = 1 - SSres / SStot;
                }
            }
            R02 = R2;
            for (int j = 0; j < 3; j++)
            {
                par_b_43 = par_b_43 - 0.001; // сужение
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N43; i++)
                {
                    if (!Ln43)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * 100)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_b_43 = par_b_43 + 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_b_43 = par_b_43 - 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N43; i++)
                        {
                            if (!Ln43)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * 100)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_b_43 = par_b_43 + 0.001; // расширение
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N43; i++)
                {
                    if (!Ln43)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * 100)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_b_43 = par_b_43 - 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_b_43 = par_b_43 + 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N43; i++)
                        {
                            if (!Ln43)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * 100)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_a_43 = par_a_43 - 0.001; // шаг влево
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N43; i++)
                {
                    if (!Ln43)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * 100)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_a_43 = par_a_43 + 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_a_43 = par_a_43 - 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N43; i++)
                        {
                            if (!Ln43)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * 100)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_a_43 = par_a_43 + 0.001; // шаг вправо
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N43; i++)
                {
                    if (!Ln43)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * 100)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_a_43 = par_a_43 - 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_a_43 = par_a_43 + 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N43; i++)
                        {
                            if (!Ln43)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * 100)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }
            }

            Area0_43 = Area0_43 - 0.1; // уменьшение площади
            SSres = 0;
            SStot = 0;
            for (int i = 0; i < N43; i++)
            {
                if (!Ln43)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                else
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
            }
            R2 = 1 - SSres / SStot;
            if (R02 > R2)
                Area0_43 = Area0_43 + 0.1;
            else
                while (R02 < R2)
                {
                    R02 = R2;
                    Area0_43 = Area0_43 - 0.1;
                    SSres = 0;
                    SStot = 0;
                    for (int i = 0; i < N43; i++)
                    {
                        if (!Ln43)
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                        else
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                    }
                    R2 = 1 - SSres / SStot;
                }

            Area0_43 = Area0_43 + 0.1; // увеличение площади
            SSres = 0;
            SStot = 0;
            for (int i = 0; i < N43; i++)
            {
                if (!Ln43)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                else
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
            }
            R2 = 1 - SSres / SStot;
            if (R02 > R2)
                Area0_43 = Area0_43 - 0.1;
            else
                while (R02 < R2)
                {
                    R02 = R2;
                    Area0_43 = Area0_43 + 0.1;
                    SSres = 0;
                    SStot = 0;
                    for (int i = 0; i < N43; i++)
                    {
                        if (!Ln43)
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[0].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                        else
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) * (Math.Log((Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value))) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                    }
                    R2 = 1 - SSres / SStot;
                }

            R2 = R02;
            Area_43 = Area0_43;
            par_a0_43 = par_a_43;
            par_b0_43 = par_b_43;

            for (double i = 0.1; i < X_max43 + r; i = i + 0.1)
            {
                fDi = r / (par_b_43 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_43) * (Math.Log(i) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area_43;
                Program.f14.chart2.Series[1].Points.AddXY(i, fDi);
            }
            D43 = Math.Exp(par_a_43) * Math.Exp(0.5 * par_b_43 * par_b_43);
            Mo = Math.Exp(par_a_43 - par_b_43 * par_b_43);
            m = Math.Exp(par_a_43);
            if (Ln43)
            {
                D43 = Math.Exp(D43);
                Mo = Math.Exp(Mo);
                m = Math.Exp(m);
            }

            textBox57.Text = Math.Round(R2, 4).ToString("0.0000");
            richTextBox4.SelectedRtf = Properties.Resources.log_norm;
            richTextBox4.Height = 81;
            textBox60.Text = Math.Round(par_b_43, 3).ToString("0.000");
            textBox56.Text = Math.Round(par_a_43, 3).ToString("0.000");
            textBox58.Text = Math.Round(par_b_43, 3).ToString("0.000");
            textBox58.Location = new Point(225, 31);
            textBox56.Location = new Point(239, 8);
            textBox53.Text = Math.Round(D43, 1).ToString("0.0");
            textBox54.Text = Math.Round(m, 1).ToString("0.0");
            textBox55.Text = Math.Round(Mo, 1).ToString("0.0");
            textBox52.Text = Math.Round(par_a_43, 3).ToString("0.000");
            textBox51.Text = Math.Round(par_b_43, 3).ToString("0.000");
            textBox50.Text = Math.Round(Area_43, 1).ToString();
            analyt_func_43 = true;

            Program.f14.chart2.Series[2].IsVisibleInLegend = false;
            Program.f14.chart2.Series[5].IsVisibleInLegend = false;
            Program.f14.chart2.Series[1].IsVisibleInLegend = true;
            if (Program.f1.Rus == true)
                Program.f14.chart2.Series[1].Name = "Лог-норм 1";
            if (Program.f1.Eng == true)
                Program.f14.chart2.Series[1].Name = "Log-Norm 1";
        }

        private void Log_Norm_20()
        {
            if (dataGridView5.Rows[0].Cells[2].Value == null)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Значения таблицы не заданы! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("The table values are missing! Try again", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (dataGridView5.Rows[0].Cells[1].Value == null)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Число частиц не задано! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles number are not determined! Try again!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (!open_f15)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Сначала отобразите гистограмму на графике!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Plot the histogram first!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            par_a0_20 = 0;
            par_b0_20 = 0;
            Area_20 = 100;
            Area0_20 = 100;
            LogNorm_20 = true;
            Norm_20 = false;
            tabControl6.Visible = true;
            tabControl6.TabPages[0].Text = "Лог-норм";
            label10.Text = "ln(μ)";
            if (open_f15)
            {
                Program.f15.chart2.Series[1].Points.Clear();
                Program.f15.chart2.Series[2].Points.Clear();
                Program.f15.chart2.Series[5].Points.Clear();
            }
            richTextBox5.Clear();
            pictureBox5.Visible = true;
            //textBox11.Visible = true;
            //textBox12.Visible = true;
            label70.Visible = true;
            label71.Visible = true;
            textBox63.Visible = true;
            label72.Visible = true;
            textBox64.Visible = true;
            label73.Visible = true;
            label74.Visible = true;
            //label19.Visible = true;
            //label22.Visible = true;
            textBox65.Visible = true;
            textBox67.Location = new Point(226, 58);
            R2 = 0;
            SSres = 0;
            SStot = 0;
            sumni = 0;
            S_ni_lndi = 0;
            sum1 = 0;
            sum_Y_20 = 0;
            for (int i = 0; i < N20; i++)
            {
                //sumni = sumni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                if (!Ln20)
                    S_ni_lndi = S_ni_lndi + Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) * Math.Log(Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value));
                else
                    S_ni_lndi = S_ni_lndi + Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) * Math.Log(Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value));
                sum_Y_20 = sum_Y_20 + Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value);
            }
            //par_a = S_ni_lndi / sumni;
            par_a_20 = S_ni_lndi / 100;

            for (int i = 0; i < N20; i++)
            {
                if (!Ln20)
                    sum1 = sum1 + Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) * (Math.Log(Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value)) - par_a_20) * (Math.Log(Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value)) - par_a_20);
                else
                    sum1 = sum1 + Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) * (Math.Log(Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value)) - par_a_20) * (Math.Log(Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value)) - par_a_20);
            }
            //par_b = Math.Sqrt(sum1 / (sumni - 1));
            par_b_20 = Math.Sqrt(sum1 / 100);

            for (int i = 0; i < N20; i++)
            {
                if (!Ln20)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * 100)), 2);
                else
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)))), 2);
                SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
            }
            R2 = 1 - SSres / SStot;

            if (Ln20)
            {
                Area0_20 = 1;
                R02 = 0;
                while (R02 < R2 || R2 < 0)
                {
                    R02 = R2;
                    Area0_20 = Area0_20 + 0.1;
                    SSres = 0;
                    SStot = 0;
                    for (int i = 0; i < N20; i++)
                    {
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) / (2 * par_b_20 * par_b_20))*Area0_20)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                    }
                    R2 = 1 - SSres / SStot;
                }
            }
            R02 = R2;
            for (int j = 0; j < 3; j++)
            {
                par_b_20 = par_b_20 - 0.001; // сужение
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N20; i++)
                {
                    if (!Ln20)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * 100)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) / (2 * par_b_20 * par_b_20))*Area0_20)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_b_20 = par_b_20 + 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_b_20 = par_b_20 - 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N20; i++)
                        {
                            if (!Ln20)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * 100)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_b_20 = par_b_20 + 0.001; // расширение
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N20; i++)
                {
                    if (!Ln20)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * 100)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_b_20 = par_b_20 - 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_b_20 = par_b_20 + 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N20; i++)
                        {
                            if (!Ln20)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * 100)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_a_20 = par_a_20 - 0.001; // шаг влево
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N20; i++)
                {
                    if (!Ln20)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * 100)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_a_20 = par_a_20 + 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_a_20 = par_a_20 - 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N20; i++)
                        {
                            if (!Ln20)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * 100)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_a_20 = par_a_20 + 0.001; // шаг вправо
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N20; i++)
                {
                    if (!Ln20)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * 100)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_a_20 = par_a_20 - 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_a_20 = par_a_20 + 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N20; i++)
                        {
                            if (!Ln20)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * 100)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }
            }

            Area0_20 = Area0_20 - 0.1; // уменьшение площади
            SSres = 0;
            SStot = 0;
            for (int i = 0; i < N20; i++)
            {
                if (!Ln20)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * 100)), 2);
                else
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
            }
            R2 = 1 - SSres / SStot;
            if (R02 > R2)
                Area0_20 = Area0_20 + 0.1;
            else
                while (R02 < R2)
                {
                    R02 = R2;
                    Area0_20 = Area0_20 - 0.1;
                    SSres = 0;
                    SStot = 0;
                    for (int i = 0; i < N20; i++)
                    {
                        if (!Ln20)
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * 100)), 2);
                        else
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                    }
                    R2 = 1 - SSres / SStot;
                }

            Area0_20 = Area0_20 + 0.1; // увеличение площади
            SSres = 0;
            SStot = 0;
            for (int i = 0; i < N20; i++)
            {
                if (!Ln20)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * 100)), 2);
                else
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
            }
            R2 = 1 - SSres / SStot;
            if (R02 > R2)
                Area0_20 = Area0_20 - 0.1;
            else
                while (R02 < R2)
                {
                    R02 = R2;
                    Area0_20 = Area0_20 + 0.1;
                    SSres = 0;
                    SStot = 0;
                    for (int i = 0; i < N20; i++)
                    {
                        if (!Ln20)
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[0].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * 100)), 2);
                        else
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) * (Math.Log((Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value))) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                    }
                    R2 = 1 - SSres / SStot;
                }

            R2 = R02;
            Area_20 = Area0_20;
            par_a0_20 = par_a_20;
            par_b0_20 = par_b_20;

            for (double i = 0.1; i < X_max20 + r; i = i + 0.1)
            {
                fDi = r / (par_b_20 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a_20) * (Math.Log(i) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area_20;
                Program.f15.chart2.Series[1].Points.AddXY(i, fDi);
            }
            D20 = Math.Exp(par_a_20) * Math.Exp(0.5 * par_b_20 * par_b_20);
            Mo = Math.Exp(par_a_20 - par_b_20 * par_b_20);
            m = Math.Exp(par_a_20);
            if (Ln20)
            {
                D20 = Math.Exp(D20);
                Mo = Math.Exp(Mo);
                m = Math.Exp(m);
            }

            textBox67.Text = Math.Round(R2, 4).ToString("0.0000");
            richTextBox5.SelectedRtf = Properties.Resources.log_norm;
            richTextBox5.Height = 81;
            textBox70.Text = Math.Round(par_b_20, 3).ToString("0.000");
            textBox66.Text = Math.Round(par_a_20, 3).ToString("0.000");
            textBox68.Text = Math.Round(par_b_20, 3).ToString("0.000");
            textBox68.Location = new Point(225, 31);
            textBox66.Location = new Point(239, 8);
            textBox63.Text = Math.Round(D20, 1).ToString("0.0");
            textBox64.Text = Math.Round(m, 1).ToString("0.0");
            textBox65.Text = Math.Round(Mo, 1).ToString("0.0");
            textBox21.Text = Math.Round(par_a_20, 3).ToString("0.000");
            textBox14.Text = Math.Round(par_b_20, 3).ToString("0.000");
            textBox13.Text = Math.Round(Area_20, 1).ToString();
            analyt_func_20 = true;

            Program.f15.chart2.Series[2].IsVisibleInLegend = false;
            Program.f15.chart2.Series[5].IsVisibleInLegend = false;
            Program.f15.chart2.Series[1].IsVisibleInLegend = true;
            if (Program.f1.Rus == true)
                Program.f15.chart2.Series[1].Name = "Лог-норм 1";
            if (Program.f1.Eng == true)
                Program.f15.chart2.Series[1].Name = "Log-Norm 1";
        }

        private void логнормФияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (flag_lin_raspr)
                {
                    Log_Norm_10();
                    return;
                }
                if (flag_cub_raspr)
                {
                    Log_Norm_30();
                    return;
                }

                if (flag_vol_surf_raspr)
                {
                    Log_Norm_32();
                    return;
                }

                if (flag_volweight_raspr)
                {
                    Log_Norm_43();
                    return;
                }

                if (flag_surf_raspr)
                    Log_Norm_20();
            }
            catch
            {
                return;
            }
        }

        public void Norm_10()
        {
            if (dataGridView1.Rows[0].Cells[2].Value == null)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Значения таблицы не заданы! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("The table values are missing! Try again", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (dataGridView1.Rows[0].Cells[1].Value == null)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Число частиц не задано! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles number are not determined! Try again!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (!open_f11)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Сначала отобразите гистограмму на графике!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Plot the histogram first!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            par_a0 = 0;
            par_b0 = 0;
            Area = 100;
            Area0 = 100;
            LogNorm = false;
            Norm = true;
            tabControl2.Visible = true;
            tabControl2.TabPages[0].Text = "Гаусс";
            if (!Ln)
            {
                label19.Visible = false;
                textBox23.Visible = false;
                label21.Visible = false;
                pictureBox2.Visible = true;
                textBox11.Visible = true;
                label5.Visible = true;
            }
            else
            {
                label19.Visible = true;
                textBox23.Visible = true;
                label21.Visible = true;
                pictureBox2.Visible = false;
                textBox11.Visible = false;
                label5.Visible = false;
            }
            label23.Text = "μ";
            label22.Visible = false;
            textBox10.Visible = false;
            label20.Visible = false;
            pictureBox3.Visible = false;
            textBox12.Visible = false;
            //pictureBox4.Visible = false;

            if (open_f11)
            {
                Program.f11.chart2.Series[1].Points.Clear();
                Program.f11.chart2.Series[2].Points.Clear();
                Program.f11.chart2.Series[5].Points.Clear();
            }
            richTextBox1.Clear();
            label9.Visible = false;
            textBox22.Visible = true;
            textBox22.Location = new Point(226, 58);
            sumni = 0;
            S_ni = 0;
            R2 = 0;
            SSres = 0;
            SStot = 0;
            sum1 = 0;
            sum_Y = 0;

            for (int i = 0; i < N; i++)
            {
                //sumni = sumni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                if (!Ln)
                    S_ni = S_ni + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) * Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                else
                    S_ni = S_ni + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) * Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value);
                sum_Y = sum_Y + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
            }
            par_a = S_ni / 100;
            D10 = par_a;

            for (int i = 0; i < N; i++)
            {
                if (!Ln)
                    sum1 = sum1 + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a);
                else
                    sum1 = sum1 + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a);
            }
            par_b = Math.Sqrt(sum1 / 100);

            for (int i = 0; i < N; i++)
            {
                if (!Ln)
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) / (2 * par_b * par_b)) * 100)), 2);
                else
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) / (2 * par_b * par_b)))), 2);
                SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
            }
            R2 = 1 - SSres / SStot;
            if (Ln)
            {
                R02 = 0;
                Area0 = 1;
                while (R02 < R2 || R2 < 0)
                {
                    R02 = R2;
                    Area0 = Area0 + 0.1;
                    SSres = 0;
                    SStot = 0;
                    for (int i = 0; i < N; i++)
                    {

                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);

                    }
                    R2 = 1 - SSres / SStot;
                }
            }

            R02 = R2;
            for (int j = 0; j < 3; j++)
            {
                par_a = par_a - 0.001; // шаг влево
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N; i++)
                {
                    if (!Ln)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_a = par_a + 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_a = par_a - 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N; i++)
                        {
                            if (!Ln)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_a = par_a + 0.001; // шаг вправо
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N; i++)
                {
                    if (!Ln)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_a = par_a - 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_a = par_a + 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N; i++)
                        {
                            if (!Ln)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_b = par_b - 0.001; // сужение
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N; i++)
                {
                    if (!Ln)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_b = par_b + 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_b = par_b - 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N; i++)
                        {
                            if (!Ln)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_b = par_b + 0.001; // расширение
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N; i++)
                {
                    if (!Ln)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_b = par_b - 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_b = par_b + 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N; i++)
                        {
                            if (!Ln)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                Area0 = Area0 - 0.1; // уменьшение площади
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N; i++)
                {
                    if (!Ln)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    Area0 = Area0 + 0.1;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        Area0 = Area0 - 0.1;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N; i++)
                        {
                            if (!Ln)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                Area0 = Area0 + 0.1; // увеличение площади
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N; i++)
                {
                    if (!Ln)
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    else
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    Area0 = Area0 - 0.1;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        Area0 = Area0 + 0.1;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N; i++)
                        {
                            if (!Ln)
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                            else
                                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) - par_a) / (2 * par_b * par_b)) * Area0)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }
            }
            for (double i = 0.1; i < X_max + r; i = i + 0.1)
            {
                fDi = r / (par_b * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a) * (i - par_a) / (2 * par_b * par_b)) * Area0;
                Program.f11.chart2.Series[1].Points.AddXY(i, fDi);
            }
            R2 = R02;
            Area = Area0;
            par_a0 = par_a;
            par_b0 = par_b;

            textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
            richTextBox1.SelectedRtf = Properties.Resources.Gauss;
            richTextBox1.Height = 81;
            textBox4.Text = Math.Round(par_b, 3).ToString("0.000");
            textBox3.Text = Math.Round(par_a, 3).ToString("0.000");
            textBox5.Text = Math.Round(par_b, 3).ToString("0.000");
            textBox3.Location = new Point(215, 8);
            textBox5.Location = new Point(212, 31);
            if (!Ln)
                textBox11.Text = Math.Round(par_a, 1).ToString("0.0");
            else
                textBox23.Text = Math.Round(Math.Exp(par_a), 1).ToString("0.0");
            textBox1.Text = Math.Round(par_a, 3).ToString("0.000");
            textBox2.Text = Math.Round(par_b, 3).ToString("0.000");
            textBox9.Text = Math.Round(Area0, 2).ToString();
            analyt_func = true;

            Program.f11.chart2.Series[2].IsVisibleInLegend = false;
            Program.f11.chart2.Series[5].IsVisibleInLegend = false;
            Program.f11.chart2.Series[1].IsVisibleInLegend = true;
            if (Program.f1.Rus == true)
                Program.f11.chart2.Series[1].Name = "Гаусс 1";
            if (Program.f1.Eng == true)
                Program.f11.chart2.Series[1].Name = "Gauss 1";
            
        }

        private void Norm30()
        {
            if (!Ln30)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Функция Гаусса недоступна в кубическом распределении", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("The Gaussian function is not available in the cubic distribution", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (dataGridView2.Rows[0].Cells[2].Value == null)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Значения таблицы не заданы! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("The table values are missing! Try again", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (dataGridView2.Rows[0].Cells[1].Value == null)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Число частиц не задано! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles number are not determined! Try again!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (!open_f12)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Сначала отобразите гистограмму на графике!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Plot the histogram first!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            par_a0_30 = 0;
            par_b0_30 = 0;
            Area_30 = 100;
            Area0_30 = 1;
            R02 = 0;
            LogNorm_30 = false;
            Norm_30 = true;
            tabControl3.Visible = true;
            tabControl3.TabPages[0].Text = "Гаусс";


            pictureBox4.Visible = false;
            textBox27.Visible = false;
            label33.Visible = false;
            label37.Visible = true;
            label34.Visible = true;
            textBox28.Visible = true;


            label28.Text = "μ";

            label35.Visible = false;

            textBox29.Visible = false;
            label36.Visible = false;
            //textBox11.Visible = true;
            //pictureBox3.Visible = false;
            //textBox12.Visible = false;
            //pictureBox4.Visible = false;

            if (open_f12)
            {
                Program.f12.chart2.Series[1].Points.Clear();
                Program.f12.chart2.Series[2].Points.Clear();
                Program.f12.chart2.Series[5].Points.Clear();
            }
            richTextBox2.Clear();
            //label5.Visible = true;
            //label9.Visible = false;
            //textBox22.Visible = true;
            textBox31.Location = new Point(226, 58);
            //sumni = 0;
            S_ni = 0;
            R2 = 0;
            SSres = 0;
            SStot = 0;
            sum1 = 0;
            sum_Y_30 = 0;
            for (int i = 0; i < N30; i++)
            {
                //sumni = sumni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                S_ni = S_ni + Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) * Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value);
                sum_Y_30 = sum_Y_30 + Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value);
            }
            par_a_30 = S_ni / 100;
            D30 = par_a_30;

            for (int i = 0; i < N30; i++)
                sum1 = sum1 + Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30);
            par_b_30 = Math.Sqrt(sum1 / 100);


            for (int i = 0; i < N30; i++)
            {
                //if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) != 0)
                //{
                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                // }
            }
            R2 = 1 - SSres / SStot;

            while (R02 < R2 || R2 < 0)
            {
                R02 = R2;
                Area0_30 = Area0_30 + 0.1;
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N30; i++)
                {
                    //if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) != 0)
                    //{
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                    // }
                }
                R2 = 1 - SSres / SStot;
            }
            for (int j = 0; j < 3; j++)
            {
                par_a_30 = par_a_30 - 0.001; // шаг влево
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N30; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_a_30 = par_a_30 + 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_a_30 = par_a_30 - 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N30; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_a_30 = par_a_30 + 0.001; // шаг вправо
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N30; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_a_30 = par_a_30 - 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_a_30 = par_a_30 + 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N30; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_b_30 = par_b_30 - 0.001; // сужение
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N30; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_b_30 = par_b_30 + 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_b_30 = par_b_30 - 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N30; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_b_30 = par_b_30 + 0.001; // расширение
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N30; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_b_30 = par_b_30 - 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_b_30 = par_b_30 + 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N30; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                Area0_30 = Area0_30 - 0.1; // уменьшение площади
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N30; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    Area0_30 = Area0_30 + 0.1;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        Area0_30 = Area0_30 - 0.1;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N30; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                Area0_30 = Area0_30 + 0.1; // увеличение площади
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N30; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    Area0_30 = Area0_30 - 0.1;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        Area0_30 = Area0_30 + 0.1;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N30; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - (r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) * (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value) - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value) - sum_Y_30 / N30), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }
            }
            for (double i = 0.1; i < X_max30 + r; i = i + 0.1)
            {
                fDi = r / (par_b_30 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_30) * (i - par_a_30) / (2 * par_b_30 * par_b_30)) * Area0_30;
                Program.f12.chart2.Series[1].Points.AddXY(i, fDi);
            }
            R2 = R02;
            Area_30 = Area0_30;
            par_a0_30 = par_a_30;
            par_b0_30 = par_b_30;

            textBox31.Text = Math.Round(R2, 4).ToString("0.0000");
            richTextBox2.SelectedRtf = Properties.Resources.Gauss;
            richTextBox2.Height = 81;
            textBox34.Text = Math.Round(par_b_30, 3).ToString("0.000");
            textBox30.Text = Math.Round(par_a_30, 3).ToString("0.000");
            textBox32.Text = Math.Round(par_b_30, 3).ToString("0.000");
            textBox30.Location = new Point(215, 8);
            textBox32.Location = new Point(212, 31);
            if (!Ln30)
                textBox27.Text = Math.Round(par_a_30, 1).ToString("0.0");
            else
                textBox28.Text = Math.Round(Math.Exp(par_a_30), 1).ToString("0.0");
            textBox26.Text = Math.Round(par_a_30, 3).ToString("0.000");
            textBox25.Text = Math.Round(par_b_30, 3).ToString("0.000");
            textBox24.Text = Math.Round(Area_30, 1).ToString();
            analyt_func_30 = true;

            Program.f12.chart2.Series[2].IsVisibleInLegend = false;
            Program.f12.chart2.Series[5].IsVisibleInLegend = false;
            Program.f12.chart2.Series[1].IsVisibleInLegend = true;
            if (Program.f1.Rus == true)
                Program.f12.chart2.Series[1].Name = "Гаусс 1";
            if (Program.f1.Eng == true)
                Program.f12.chart2.Series[1].Name = "Gauss 1";
        }

        private void Norm32()
        {
            if (!Ln32)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Функция Гаусса недоступна в объемно-поверхностном распределении", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("The Gaussian function is not available in the volume-surface distribution", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (dataGridView3.Rows[0].Cells[2].Value == null)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Значения таблицы не заданы! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("The table values are missing! Try again", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (dataGridView3.Rows[0].Cells[1].Value == null)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Число частиц не задано! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles number are not determined! Try again!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (!open_f13)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Сначала отобразите гистограмму на графике!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Plot the histogram first!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            par_a0_32 = 0;
            par_b0_32 = 0;
            Area_32 = 100;
            Area0_32 = 1;
            R02 = 0;
            LogNorm_32 = false;
            Norm_32 = true;
            tabControl4.Visible = true;
            tabControl4.TabPages[0].Text = "Гаусс";


            pictureBox11.Visible = false;
            textBox40.Visible = false;
            label47.Visible = false;
            label51.Visible = true;
            label48.Visible = true;
            textBox41.Visible = true;


            label42.Text = "μ";

            label49.Visible = false;

            textBox42.Visible = false;
            label50.Visible = false;
            //textBox11.Visible = true;
            //pictureBox3.Visible = false;
            //textBox12.Visible = false;
            //pictureBox4.Visible = false;

            if (open_f13)
            {
                Program.f13.chart2.Series[1].Points.Clear();
                Program.f13.chart2.Series[2].Points.Clear();
                Program.f13.chart2.Series[5].Points.Clear();
            }
            richTextBox3.Clear();
            //label5.Visible = true;
            //label9.Visible = false;
            //textBox22.Visible = true;
            textBox44.Location = new Point(226, 58);
            //sumni = 0;
            S_ni = 0;
            R2 = 0;
            SSres = 0;
            SStot = 0;
            sum1 = 0;
            sum_Y_32 = 0;
            for (int i = 0; i < N32; i++)
            {
                //sumni = sumni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                S_ni = S_ni + Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) * Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value);
                sum_Y_32 = sum_Y_32 + Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value);
            }
            par_a_32 = S_ni / 100;
            D32 = par_a_32;

            for (int i = 0; i < N32; i++)
                sum1 = sum1 + Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32);
            par_b_32 = Math.Sqrt(sum1 / 100);


            for (int i = 0; i < N32; i++)
            {
                //if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) != 0)
                //{
                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                // }
            }
            R2 = 1 - SSres / SStot;

            while (R02 < R2 || R2 < 0)
            {
                R02 = R2;
                Area0_32 = Area0_32 + 0.1;
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N32; i++)
                {
                    //if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) != 0)
                    //{
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                    // }
                }
                R2 = 1 - SSres / SStot;
            }
            for (int j = 0; j < 3; j++)
            {
                par_a_32 = par_a_32 - 0.001; // шаг влево
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N32; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_a_32 = par_a_32 + 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_a_32 = par_a_32 - 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N32; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_a_32 = par_a_32 + 0.001; // шаг вправо
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N32; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_a_32 = par_a_32 - 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_a_32 = par_a_32 + 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N32; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_b_32 = par_b_32 - 0.001; // сужение
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N32; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_b_32 = par_b_32 + 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_b_32 = par_b_32 - 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N32; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_b_32 = par_b_32 + 0.001; // расширение
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N32; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_b_32 = par_b_32 - 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_b_32 = par_b_32 + 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N32; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                Area0_32 = Area0_32 - 0.1; // уменьшение площади
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N32; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    Area0_32 = Area0_32 + 0.1;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        Area0_32 = Area0_32 - 0.1;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N32; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                Area0_32 = Area0_32 + 0.1; // увеличение площади
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N32; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    Area0_32 = Area0_32 - 0.1;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        Area0_32 = Area0_32 + 0.1;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N32; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - (r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) * (Convert.ToDouble(dataGridView3.Rows[i].Cells[3].Value) - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value) - sum_Y_32 / N32), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }
            }
            for (double i = 0.1; i < X_max32 + r; i = i + 0.1)
            {
                fDi = r / (par_b_32 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_32) * (i - par_a_32) / (2 * par_b_32 * par_b_32)) * Area0_32;
                Program.f13.chart2.Series[1].Points.AddXY(i, fDi);
            }
            R2 = R02;
            Area_32 = Area0_32;
            par_a0_32 = par_a_32;
            par_b0_32 = par_b_32;

            textBox44.Text = Math.Round(R2, 4).ToString("0.0000");
            richTextBox3.SelectedRtf = Properties.Resources.Gauss;
            richTextBox3.Height = 81;
            textBox47.Text = Math.Round(par_b_32, 3).ToString("0.000");
            textBox43.Text = Math.Round(par_a_32, 3).ToString("0.000");
            textBox45.Text = Math.Round(par_b_32, 3).ToString("0.000");
            textBox43.Location = new Point(215, 8);
            textBox45.Location = new Point(212, 31);
            textBox41.Text = Math.Round(Math.Exp(par_a_32), 1).ToString("0.0");
            textBox39.Text = Math.Round(par_a_32, 3).ToString("0.000");
            textBox38.Text = Math.Round(par_b_32, 3).ToString("0.000");
            textBox37.Text = Math.Round(Area_32, 1).ToString();
            analyt_func_32 = true;

            Program.f13.chart2.Series[2].IsVisibleInLegend = false;
            Program.f13.chart2.Series[5].IsVisibleInLegend = false;
            Program.f13.chart2.Series[1].IsVisibleInLegend = true;
            if (Program.f1.Rus == true)
                Program.f13.chart2.Series[1].Name = "Гаусс 1";
            if (Program.f1.Eng == true)
                Program.f13.chart2.Series[1].Name = "Gauss 1";
        }

        private void Norm43()
        {
            if (!Ln43)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Функция Гаусса недоступна во взвешенном по объему распределении", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("The Gaussian function is not available in the Volume-weighted distribution", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (dataGridView4.Rows[0].Cells[2].Value == null)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Значения таблицы не заданы! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("The table values are missing! Try again", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (dataGridView4.Rows[0].Cells[1].Value == null)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Число частиц не задано! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles number are not determined! Try again!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (!open_f14)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Сначала отобразите гистограмму на графике!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Plot the histogram first!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            par_a0_43 = 0;
            par_b0_43 = 0;
            Area_43 = 100;
            Area0_43 = 1;
            R02 = 0;
            LogNorm_43 = false;
            Norm_43 = true;
            tabControl5.Visible = true;
            tabControl5.TabPages[0].Text = "Гаусс";


            pictureBox12.Visible = false;
            textBox53.Visible = false;
            label61.Visible = false;
            label65.Visible = true;
            label62.Visible = true;
            textBox54.Visible = true;


            label56.Text = "μ";

            label63.Visible = false;

            textBox55.Visible = false;
            label64.Visible = false;
            //textBox11.Visible = true;
            //pictureBox3.Visible = false;
            //textBox12.Visible = false;
            //pictureBox4.Visible = false;

            if (open_f14)
            {
                Program.f14.chart2.Series[1].Points.Clear();
                Program.f14.chart2.Series[2].Points.Clear();
                Program.f14.chart2.Series[5].Points.Clear();
            }
            richTextBox4.Clear();
            //label5.Visible = true;
            //label9.Visible = false;
            //textBox22.Visible = true;
            textBox57.Location = new Point(226, 58);
            //sumni = 0;
            S_ni = 0;
            R2 = 0;
            SSres = 0;
            SStot = 0;
            sum1 = 0;
            sum_Y_43 = 0;
            for (int i = 0; i < N43; i++)
            {
                //sumni = sumni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                S_ni = S_ni + Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) * Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value);
                sum_Y_43 = sum_Y_43 + Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value);
            }
            par_a_43 = S_ni / 100;
            D43 = par_a_43;

            for (int i = 0; i < N43; i++)
                sum1 = sum1 + Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43);
            par_b_43 = Math.Sqrt(sum1 / 100);


            for (int i = 0; i < N43; i++)
            {
                //if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) != 0)
                //{
                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                // }
            }
            R2 = 1 - SSres / SStot;

            while (R02 < R2 || R2 < 0)
            {
                R02 = R2;
                Area0_43 = Area0_43 + 0.1;
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N43; i++)
                {
                    //if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) != 0)
                    //{
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                    // }
                }
                R2 = 1 - SSres / SStot;
            }
            for (int j = 0; j < 3; j++)
            {
                par_a_43 = par_a_43 - 0.001; // шаг влево
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N43; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_a_43 = par_a_43 + 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_a_43 = par_a_43 - 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N43; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_a_43 = par_a_43 + 0.001; // шаг вправо
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N43; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_a_43 = par_a_43 - 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_a_43 = par_a_43 + 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N43; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_b_43 = par_b_43 - 0.001; // сужение
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N43; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_b_43 = par_b_43 + 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_b_43 = par_b_43 - 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N43; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_b_43 = par_b_43 + 0.001; // расширение
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N43; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_b_43 = par_b_43 - 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_b_43 = par_b_43 + 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N43; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                Area0_43 = Area0_43 - 0.1; // уменьшение площади
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N43; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    Area0_43 = Area0_43 + 0.1;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        Area0_43 = Area0_43 - 0.1;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N43; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                Area0_43 = Area0_43 + 0.1; // увеличение площади
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N43; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    Area0_43 = Area0_43 - 0.1;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        Area0_43 = Area0_43 + 0.1;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N43; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - (r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) * (Convert.ToDouble(dataGridView4.Rows[i].Cells[3].Value) - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView4.Rows[i].Cells[2].Value) - sum_Y_43 / N43), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }
            }
            for (double i = 0.1; i < X_max43 + r; i = i + 0.1)
            {
                fDi = r / (par_b_43 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_43) * (i - par_a_43) / (2 * par_b_43 * par_b_43)) * Area0_43;
                Program.f14.chart2.Series[1].Points.AddXY(i, fDi);
            }
            R2 = R02;
            Area_43 = Area0_43;
            par_a0_43 = par_a_43;
            par_b0_43 = par_b_43;

            textBox57.Text = Math.Round(R2, 4).ToString("0.0000");
            richTextBox4.SelectedRtf = Properties.Resources.Gauss;
            richTextBox4.Height = 81;
            textBox60.Text = Math.Round(par_b_43, 3).ToString("0.000");
            textBox56.Text = Math.Round(par_a_43, 3).ToString("0.000");
            textBox58.Text = Math.Round(par_b_43, 3).ToString("0.000");
            textBox56.Location = new Point(215, 8);
            textBox58.Location = new Point(212, 31);
            textBox54.Text = Math.Round(Math.Exp(par_a_43), 1).ToString("0.0");
            textBox52.Text = Math.Round(par_a_43, 3).ToString("0.000");
            textBox51.Text = Math.Round(par_b_43, 3).ToString("0.000");
            textBox50.Text = Math.Round(Area_43, 1).ToString();
            analyt_func_43 = true;

            Program.f14.chart2.Series[2].IsVisibleInLegend = false;
            Program.f14.chart2.Series[5].IsVisibleInLegend = false;
            Program.f14.chart2.Series[1].IsVisibleInLegend = true;
            if (Program.f1.Rus == true)
                Program.f14.chart2.Series[1].Name = "Гаусс 1";
            if (Program.f1.Eng == true)
                Program.f14.chart2.Series[1].Name = "Gauss 1";
        }

        private void Norm20()
        {
            if (!Ln20)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Функция Гаусса недоступна в поверхностном распределении", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("The Gaussian function is not available in the surface distribution", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (dataGridView5.Rows[0].Cells[2].Value == null)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Значения таблицы не заданы! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("The table values are missing! Try again", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (dataGridView5.Rows[0].Cells[1].Value == null)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Число частиц не задано! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles number are not determined! Try again!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (!open_f15)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Сначала отобразите гистограмму на графике!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Plot the histogram first!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            par_a0_20 = 0;
            par_b0_20 = 0;
            Area_20 = 100;
            Area0_20 = 1;
            R02 = 0;
            LogNorm_20 = false;
            Norm_20 = true;
            tabControl6.Visible = true;
            tabControl6.TabPages[0].Text = "Гаусс";


            pictureBox5.Visible = false;
            textBox63.Visible = false;
            label70.Visible = false;
            label74.Visible = true;
            label71.Visible = true;
            textBox64.Visible = true;


            label10.Text = "μ";

            label72.Visible = false;

            textBox65.Visible = false;
            label73.Visible = false;
            //textBox11.Visible = true;
            //pictureBox3.Visible = false;
            //textBox12.Visible = false;
            //pictureBox4.Visible = false;

            if (open_f15)
            {
                Program.f15.chart2.Series[1].Points.Clear();
                Program.f15.chart2.Series[2].Points.Clear();
                Program.f15.chart2.Series[5].Points.Clear();
            }
            richTextBox5.Clear();
            //label5.Visible = true;
            //label9.Visible = false;
            //textBox22.Visible = true;
            textBox67.Location = new Point(226, 58);
            //sumni = 0;
            S_ni = 0;
            R2 = 0;
            SSres = 0;
            SStot = 0;
            sum1 = 0;
            sum_Y_20 = 0;
            for (int i = 0; i < N20; i++)
            {
                //sumni = sumni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                S_ni = S_ni + Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) * Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value);
                sum_Y_20 = sum_Y_20 + Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value);
            }
            par_a_20 = S_ni / 100;
            D20 = par_a_20;

            for (int i = 0; i < N20; i++)
                sum1 = sum1 + Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20);
            par_b_20 = Math.Sqrt(sum1 / 100);


            for (int i = 0; i < N20; i++)
            {
                //if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) != 0)
                //{
                SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                // }
            }
            R2 = 1 - SSres / SStot;

            while (R02 < R2 || R2 < 0)
            {
                R02 = R2;
                Area0_20 = Area0_20 + 0.1;
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N20; i++)
                {
                    //if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) != 0)
                    //{
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                    // }
                }
                R2 = 1 - SSres / SStot;
            }
            for (int j = 0; j < 3; j++)
            {
                par_a_20 = par_a_20 - 0.001; // шаг влево
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N20; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_a_20 = par_a_20 + 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_a_20 = par_a_20 - 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N20; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_a_20 = par_a_20 + 0.001; // шаг вправо
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N20; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_a_20 = par_a_20 - 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_a_20 = par_a_20 + 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N20; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_b_20 = par_b_20 - 0.001; // сужение
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N20; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_b_20 = par_b_20 + 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_b_20 = par_b_20 - 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N20; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                par_b_20 = par_b_20 + 0.001; // расширение
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N20; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    par_b_20 = par_b_20 - 0.001;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        par_b_20 = par_b_20 + 0.001;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N20; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                Area0_20 = Area0_20 - 0.1; // уменьшение площади
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N20; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    Area0_20 = Area0_20 + 0.1;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        Area0_20 = Area0_20 - 0.1;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N20; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }

                Area0_20 = Area0_20 + 0.1; // увеличение площади
                SSres = 0;
                SStot = 0;
                for (int i = 0; i < N20; i++)
                {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                }
                R2 = 1 - SSres / SStot;
                if (R02 > R2)
                    Area0_20 = Area0_20 - 0.1;
                else
                    while (R02 < R2)
                    {
                        R02 = R2;
                        Area0_20 = Area0_20 + 0.1;
                        SSres = 0;
                        SStot = 0;
                        for (int i = 0; i < N20; i++)
                        {
                            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - (r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) * (Convert.ToDouble(dataGridView5.Rows[i].Cells[3].Value) - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20)), 2);
                            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView5.Rows[i].Cells[2].Value) - sum_Y_20 / N20), 2);
                        }
                        R2 = 1 - SSres / SStot;
                    }
            }
            for (double i = 0.1; i < X_max20 + r; i = i + 0.1)
            {
                fDi = r / (par_b_20 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a_20) * (i - par_a_20) / (2 * par_b_20 * par_b_20)) * Area0_20;
                Program.f15.chart2.Series[1].Points.AddXY(i, fDi);
            }
            R2 = R02;
            Area_20 = Area0_20;
            par_a0_20 = par_a_20;
            par_b0_20 = par_b_20;

            textBox67.Text = Math.Round(R2, 4).ToString("0.0000");
            richTextBox5.SelectedRtf = Properties.Resources.Gauss;
            richTextBox5.Height = 81;
            textBox70.Text = Math.Round(par_b_20, 3).ToString("0.000");
            textBox66.Text = Math.Round(par_a_20, 3).ToString("0.000");
            textBox68.Text = Math.Round(par_b_20, 3).ToString("0.000");
            textBox66.Location = new Point(215, 8);
            textBox68.Location = new Point(212, 31);
            textBox64.Text = Math.Round(Math.Exp(par_a_20), 1).ToString("0.0");
            textBox21.Text = Math.Round(par_a_20, 3).ToString("0.000");
            textBox14.Text = Math.Round(par_b_20, 3).ToString("0.000");
            textBox13.Text = Math.Round(Area_20, 1).ToString();
            analyt_func_20 = true;

            Program.f15.chart2.Series[2].IsVisibleInLegend = false;
            Program.f15.chart2.Series[5].IsVisibleInLegend = false;
            Program.f15.chart2.Series[1].IsVisibleInLegend = true;
            if (Program.f1.Rus == true)
                Program.f15.chart2.Series[1].Name = "Гаусс 1";
            if (Program.f1.Eng == true)
                Program.f15.chart2.Series[1].Name = "Gauss 1";
        }

        private void нормФцияГауссToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (flag_lin_raspr)
                {
                    Norm_10();
                    return;
                }

                if (flag_cub_raspr)
                {
                    Norm30();
                    return;
                }

                if (flag_vol_surf_raspr)
                {
                    Norm32();
                    return;
                }

                if (flag_volweight_raspr)
                {
                    Norm43();
                    return;
                }

                if (flag_surf_raspr)
                {
                    Norm20();
                }
            }
            catch
            {
                return;
            }
        }

        //private void логнормФцияToolStripMenuItem_Click(object sender, EventArgs e)
        //   {
        //    if (dataGridView1.Rows[0].Cells[2].Value == null)
        //  {
        //     MessageBox.Show("Значения таблицы не заданы!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //   return;
        // }
        // f4 = new Form4(this);
        // f4.Show();
        // }

        //  private void нормФцияToolStripMenuItem_Click(object sender, EventArgs e)
        // {
        //   if (dataGridView1.Rows[0].Cells[2].Value == null)
        //  {
        //     MessageBox.Show("Значения таблицы не заданы!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //    return;
        //  }
        // f5 = new Form5(this);
        // f5.Show();
        // }

        private void новоеРаспределениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Program.f1.Rus == true)
            {
                DialogResult result = MessageBox.Show("Действительно хотите создать новое рапределение?",
            "Уведомление",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question,
            MessageBoxDefaultButton.Button1);
                if (result == DialogResult.Yes)
                {
                    //if (size_H == size_W)
                      //  chart2.Width = size_W - size_H1;
                   // else
                    //{
                      //  chart2.Width = size_H - size_H1;
                      //  chart2.Height = size_H - size_H1;
                    //}
                   // chart2.Location = new System.Drawing.Point(this.Width - chart2.Width - L, size_H1 - L+15);
                    if (open_f11)
                        f11.Close();
                    if (open_f12)
                        f12.Close();
                    richTextBox1.Height = 81;
                    dataGridView1.Rows.Clear();
                    dataGridView2.Rows.Clear();
                    dataGridView3.Rows.Clear();
                    dataGridView4.Rows.Clear();
                    dataGridView5.Rows.Clear();

                    X_max = 0;
                    Sum = 0;
                    tabControl2.Visible = false;
                   // chart2.Series[0].Points.Clear();
                   // chart2.Series[1].Points.Clear();
                   // chart2.Series[2].Points.Clear();
                   // chart2.Series[3].Points.Clear();
                   // chart2.Series[4].Points.Clear();
                   // chart2.Series[5].Points.Clear();
                   // chart2.Series[0].IsVisibleInLegend = false;
                   // chart2.Series[1].IsVisibleInLegend = false;
                   // chart2.Series[2].IsVisibleInLegend = false;
                   // chart2.Series[3].IsVisibleInLegend = false;
                   // chart2.Series[4].IsVisibleInLegend = false;
                   // chart2.Series[5].IsVisibleInLegend = false;
                    groupBox2.Visible =false;
                    groupBox3.Visible = false;                                      
                    richTextBox1.Clear();
                    textBox3.Clear();
                    textBox4.Clear();
                    textBox5.Clear();                                    
                    pictureBox2.Visible = false;
                    pictureBox3.Visible = false;
                    //pictureBox4.Visible = false;
                    //textBox11.Visible = false;
                    textBox12.Visible = false;
                    label5.Visible = false;
                    label9.Visible = false;
                    Part = false;
                    CSR = false;
                    CSR1 = false;
                    CSR2 = false;
                    textBox15.Visible = false;
                    textBox16.Visible = false;
                    pictureBox7.Visible = false;
                    pictureBox8.Visible = false;
                    label12.Visible = false;
                    label13.Visible = false;
                    textBox17.Visible = false;
                    textBox18.Visible = false;
                    pictureBox9.Visible = false;
                    pictureBox10.Visible = false;
                    label14.Visible = false;
                    label15.Visible = false;
                    textBox22.Visible = false;
                }
            }
            if (Program.f1.Eng == true)
            {
                Program.f1.new_distr_f5 = true;
                Program.f5 = new Form5();
                Program.f5.ShowDialog();
                richTextBox1.Height = 81;

            }
        }

        private void логнормФцияToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (F7==false)
            {

                if (flag_cub_raspr == true)
                {
                    if (Program.f1.Rus == true)
                        MessageBox.Show("Функция описания в кубическом распределении недоступна!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (Program.f1.Eng == true)
                        MessageBox.Show("Fitting function in cubic distribution is not available!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (dataGridView1.Rows[0].Cells[2].Value == null)
                {
                    if (Program.f1.Rus == true)
                        MessageBox.Show("Значения таблицы не заданы! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (Program.f1.Eng == true)
                        MessageBox.Show("The table values are missing! Try again", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (dataGridView1.Rows[0].Cells[1].Value == null)
                {
                    if (Program.f1.Rus == true)
                        MessageBox.Show("Число частиц не задано! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (Program.f1.Eng == true)
                        MessageBox.Show("Particles number are not determined! Try again!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (flag_lin_raspr == false)
                {
                    if (Program.f1.Rus == true)
                        MessageBox.Show("Сначала отобразите гистограмму на графике!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (Program.f1.Eng == true)
                        MessageBox.Show("Plot the histogram first!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                f7 = new Form7();
                f7.Owner = this;
                f7.Show();

                
                
                
                //try
                //{
                //    
                //    if (r < r1) // случай когда есть разрыв между столбцами с двумя распределениями
                //    {
                //        for (int i = 0; i <= N1_; i++)
                //        {
                //            sumni = sumni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                //            S_ni_lndi = S_ni_lndi + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value));
                //            S1 = S1 + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                //        }
                //        par_a1 = S_ni_lndi / sumni;
                //        textBox1.Text = Math.Round(par_a1, 3).ToString("0.000");
                //        for (int i = 0; i <= N1_; i++)
                //            sum1 = sum1 + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1);
                //        par_b1 = Math.Sqrt(sum1 / (sumni - 1));
                //        textBox2.Text = Math.Round(par_b1, 3).ToString("0.000");
                //        for (double i = 0.1; i < X_max + r; i = i + 0.1)
                //        {
                //            fDi = r / (par_b1 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a1) * (Math.Log(i) - par_a1) / (2 * par_b1 * par_b1)) * S1;
                //            chart2.Series[1].Points.AddXY(i, fDi);
                //            if (MaxShiftFDi1 < fDi)
                //            {
                //                MaxShiftFDi1 = fDi;
                //                position_MaxShiftFDi1 = i;
                //            }
                //        }

                //        richTextBox1.SelectedRtf = Properties.Resources.Bi_log_norm;
                //        textBox4.Text = Math.Round(par_b1, 3).ToString("0.000");
                //        textBox3.Text = Math.Round(par_a1, 3).ToString("0.000");
                //        textBox5.Text = Math.Round(par_b1, 3).ToString("0.000");
                //        richTextBox1.Width = 309;
                //        richTextBox1.Height = 130;
                //        sumni = 0;
                //        S_ni_lndi = 0;
                //        sum1 = 0;
                //        for (int i = N2_; i < N - 1; i++)
                //        {
                //            sumni = sumni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                //            S_ni_lndi = S_ni_lndi + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value));
                //            S2 = S2 + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                //        }
                //        par_a2 = S_ni_lndi / sumni;
                //        textBox9.Visible = true;
                //        textBox9.BackColor = System.Drawing.Color.White;
                //        textBox9.Text = Math.Round(par_a2, 3).ToString("0.000");
                //        for (int i = N2_; i < N - 1; i++)
                //            sum1 = sum1 + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2);
                //        par_b2 = Math.Sqrt(sum1 / (sumni - 1));
                //        textBox6.Text = Math.Round(par_a2, 3).ToString("0.000");
                //        textBox8.Text = Math.Round(par_b2, 3).ToString("0.000");
                //        textBox7.Text = Math.Round(par_b2, 3).ToString("0.000");
                //        textBox3.Location = new Point(247, 378);
                //        textBox5.Location = new Point(234, 401);
                //        textBox7.Location = new Point(61, 455);
                //        textBox6.Location = new Point(247, 431);
                //        textBox8.Location = new Point(234, 454);
                //        textBox10.Visible = true;
                //        textBox10.BackColor = System.Drawing.Color.White;
                //        textBox10.Text = Math.Round(par_b2, 3).ToString("0.000");
                //        for (double i = 0.1; i < X_max + r; i = i + 0.1)
                //        {
                //            fDi = r / (par_b2 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a2) * (Math.Log(i) - par_a2) / (2 * par_b2 * par_b2)) * S2;
                //            chart2.Series[2].Points.AddXY(i, fDi);
                //            if (MaxShiftFDi2 < fDi)
                //            {
                //                MaxShiftFDi2 = fDi;
                //                position_MaxShiftFDi2 = i;
                //            }
                //            fDiCom = (r / (par_b1 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a1) * (Math.Log(i) - par_a1) / (2 * par_b1 * par_b1)) * S1) + (r / (par_b2 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a2) * (Math.Log(i) - par_a2) / (2 * par_b2 * par_b2)) * S2);
                //            chart2.Series[5].Points.AddXY(i, fDiCom);
                //        }
                //        for (int i = 0; i < N; i++)
                //        {
                //            //if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) != 0)
                //            // {
                //            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1)) * S1) - (r / (par_b2 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) / (2 * par_b2 * par_b2)) * S2)), 2);
                //            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                //            // }
                //        }
                //        R2 = 1 - SSres / SStot;
                //        textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                //        textBox21.Text = Math.Round(S1, 1).ToString("0.0");
                //    }
                //    else
                //    {
                //        for (int i = N1; i <= N2; i++)
                //        {
                //            if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) < Convert.ToDouble(dataGridView1.Rows[i + 1].Cells[2].Value) & Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) <= y_min)
                //            {
                //                y_min = Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                //                N_min = i;
                //            }
                //        }
                //        for (int i = 0; i < N; i++)
                //            S = S + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                //        Sum = 0;
                //        for (int i = 0; i < N; i++)
                //            Sum = Sum + Convert.ToDouble(dataGridView1[1, i].Value); // сначал номер стобца, затем номер строки               
                //        N1 = N_min;
                //        dataGV = Convert.ToDouble(dataGridView1.Rows[N1].Cells[1].Value);
                //        dataGV_zap = Convert.ToDouble(dataGridView1.Rows[0].Cells[1].Value);
                //        for (double j = 0; j <= dataGV; j = j + 0.1)
                //        {
                //            dataGridView1.Rows[N1].Cells[1].Value = j;
                //            error_sum = 0;
                //            sum1 = 0;
                //            sumni = 0;
                //            S_ni_lndi = 0;
                //            S1 = 0;
                //            for (int i = 0; i <= N1; i++)
                //            {
                //                sumni = sumni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                //                S_ni_lndi = S_ni_lndi + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value));

                //            }
                //            par_a1 = S_ni_lndi / sumni;

                //            for (int i = 0; i <= N1; i++)
                //                sum1 = sum1 + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1);
                //            par_b1 = Math.Sqrt(sum1 / (sumni - 1));

                //            for (double i = r; i < X_max + r; i = i + r)
                //                fDi = r / (par_b1 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a1) * (Math.Log(i) - par_a1) / (2 * par_b1 * par_b1));
                //            for (int i = 0; i < N1; i++)
                //                S1 = S1 + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                //            for (int i = N1; i < N; i++)
                //                S1 = S1 + r / (par_b1 * Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1)) * S;
                //            N2 = N1;

                //            bi_modal[N2] = dataGV - Convert.ToDouble(dataGridView1.Rows[N1].Cells[1].Value);
                //            for (int i = N2 + 1; i < N; i++)
                //                bi_modal[i] = (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) / 100 - r / (par_b1 * Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1))) * Sum;
                //            sumni = 0;
                //            S_ni_lndi = 0;
                //            sum1 = 0;
                //            S2 = 0;
                //            for (int i = N2; i < N; i++)
                //            {
                //                sumni = sumni + bi_modal[i];
                //                S_ni_lndi = S_ni_lndi + bi_modal[i] * Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value));
                //            }
                //            par_a2 = S_ni_lndi / sumni;
                //            for (int i = N2; i < N; i++)
                //                sum1 = sum1 + bi_modal[i] * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2);
                //            par_b2 = Math.Sqrt(sum1 / (sumni - 1));
                //            for (double i = r; i < X_max + r; i = i + r)
                //                fDi = r / (par_b2 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a2) * (Math.Log(i) - par_a2) / (2 * par_b2 * par_b2));

                //            for (int i = 0; i < N2; i++)
                //                S2 = S2 + r / (par_b2 * Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) / (2 * par_b2 * par_b2)) * S;
                //            S1 = S1 - S2;
                //            S2 = S - S1;
                //            for (int i = 0; i < N; i++)
                //                error_sum = error_sum + Math.Abs(Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b1 * Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1)) * S1) + (r / (par_b2 * Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) / (2 * par_b2 * par_b2)) * S2));
                //            error_sum1[error_sumi] = error_sum;
                //            error_sumi = error_sumi + 1;
                //        }
                //        error_sum = error_sum1[0];
                //        error_sumj1 = 0;
                //        for (int i = 0; i < error_sumi - 1; i++)
                //            if (error_sum > error_sum1[i + 1])
                //            {
                //                error_sum = error_sum1[i + 1];
                //                error_sumj1 = i + 1;
                //                error_sumj1 = error_sumj1 / 10;
                //            }
                //        dataGridView1.Rows[N1].Cells[1].Value = error_sumj1;
                //        sumni = 0;
                //        S_ni_lndi = 0;
                //        S1 = 0;
                //        sum1 = 0;
                //        for (int i = 0; i <= N1; i++)
                //        {
                //            sumni = sumni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                //            S_ni_lndi = S_ni_lndi + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value));
                //        }
                //        par_a1 = S_ni_lndi / sumni;
                //        textBox1.Text = Math.Round(par_a1, 3).ToString("0.000");
                //        for (int i = 0; i <= N1; i++)
                //            sum1 = sum1 + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1);
                //        par_b1 = Math.Sqrt(sum1 / (sumni - 1));
                //        textBox2.Text = Math.Round(par_b1, 3).ToString("0.000");
                //        for (double i = r; i < X_max + r; i = i + r)
                //            fDi = r / (par_b1 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a1) * (Math.Log(i) - par_a1) / (2 * par_b1 * par_b1));
                //        for (int i = 0; i < N1; i++)
                //            S1 = S1 + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                //        for (int i = N1; i < N; i++)
                //            S1 = S1 + r / (par_b1 * Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1)) * S;
                //        dataGridView1.Rows[N2].Cells[1].Value = dataGV;
                //        bi_modal[N2] = (dataGV - error_sumj1);
                //        for (int i = N2 + 1; i < N; i++)
                //            bi_modal[i] = (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) / 100 - r / (par_b1 * Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1))) * Sum;
                //        richTextBox1.SelectedRtf = Properties.Resources.Bi_log_norm;
                //        textBox4.Text = Math.Round(par_b1, 3).ToString("0.000");
                //        textBox3.Text = Math.Round(par_a1, 3).ToString("0.000");
                //        textBox5.Text = Math.Round(par_b1, 3).ToString("0.000");
                //        richTextBox1.Width = 309;
                //        richTextBox1.Height = 130;
                //        sumni = 0;
                //        S_ni_lndi = 0;
                //        sum1 = 0;
                //        S2 = 0;
                //        for (int i = N2; i < N; i++)
                //        {
                //            sumni = sumni + bi_modal[i];
                //            S_ni_lndi = S_ni_lndi + bi_modal[i] * Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value));
                //        }
                //        par_a2 = S_ni_lndi / sumni;
                //        textBox9.Visible = true;
                //        textBox9.BackColor = System.Drawing.Color.White;
                //        textBox9.Text = Math.Round(par_a2, 3).ToString("0.000");
                //        for (int i = N2; i < N; i++)
                //            sum1 = sum1 + bi_modal[i] * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2);
                //        par_b2 = Math.Sqrt(sum1 / (sumni - 1));
                //        textBox10.Visible = true;
                //        textBox10.BackColor = System.Drawing.Color.White;
                //        textBox10.Text = Math.Round(par_b2, 3).ToString("0.000");
                //        for (double i = r; i < X_max + r; i = i + r)
                //            fDi = r / (par_b2 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a2) * (Math.Log(i) - par_a2) / (2 * par_b2 * par_b2));
                //        for (int i = 0; i < N2; i++)
                //            S2 = S2 + r / (par_b2 * Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) / (2 * par_b2 * par_b2)) * S;
                //        S1 = S1 - S2;
                //        S2 = S - S1;
                //        textBox7.Text = Math.Round(par_b2, 3).ToString("0.000");
                //        textBox6.Text = Math.Round(par_a2, 3).ToString("0.000");
                //        textBox8.Text = Math.Round(par_b2, 3).ToString("0.000");
                //        textBox3.Location = new Point(247, 378);
                //        textBox5.Location = new Point(234, 401);
                //        textBox7.Location = new Point(61, 455);
                //        textBox6.Location = new Point(247, 431);
                //        textBox8.Location = new Point(234, 454);
                //        for (double i = 0.1; i < X_max + r; i = i + 0.1)
                //        {
                //            fDi = r / (par_b1 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a1) * (Math.Log(i) - par_a1) / (2 * par_b1 * par_b1)) * S1;
                //            chart2.Series[1].Points.AddXY(i, fDi);
                //            if (MaxShiftFDi1 < fDi)
                //            {
                //                MaxShiftFDi1 = fDi;
                //                position_MaxShiftFDi1 = i;
                //            }
                //            fDi = r / (par_b2 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a2) * (Math.Log(i) - par_a2) / (2 * par_b2 * par_b2)) * S2;
                //            chart2.Series[2].Points.AddXY(i, fDi);
                //            if (MaxShiftFDi2 < fDi)
                //            {
                //                MaxShiftFDi2 = fDi;
                //                position_MaxShiftFDi2 = i;
                //            }
                //            // добавить общую кривую
                //            fDiCom = (r / (par_b1 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a1) * (Math.Log(i) - par_a1) / (2 * par_b1 * par_b1)) * S1) + (r / (par_b2 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a2) * (Math.Log(i) - par_a2) / (2 * par_b2 * par_b2)) * S2);
                //            chart2.Series[5].Points.AddXY(i, fDiCom);
                //        }
                //        for (int i = 0; i < N; i++)
                //        {
                //            //if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) != 0)
                //            //{
                //            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1)) * S1) - (r / (par_b2 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) / (2 * par_b2 * par_b2)) * S2)), 2);
                //            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                //            //}
                //        }
                //        R2 = 1 - SSres / SStot;
                //        textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                //    }
                //    if (double.IsNaN(par_a1) || double.IsNaN(par_a2) || double.IsNaN(par_b1) || double.IsNaN(par_b2) || double.IsNaN(S1) || double.IsNaN(S2))
                //    {
                //        dataGridView1.Rows[0].Cells[1].Value = dataGV_zap;
                //        if (Program.f1.Rus == true)
                //            MessageBox.Show("Невозможно построить бимодальное распределение в автоматическом режиме! Вам предлагается подобрать параметры в ручную", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //        if (Program.f1.Eng == true)
                //            MessageBox.Show("It is impossible to plot a bimodal distribution in automatic! You are offered to fit the parameters manually", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //        richTextBox1.Clear();
                //        sumni = 0;
                //        S_ni_lndi = 0;
                //        sum1 = 0;
                //        for (int i = 0; i < N; i++)
                //        {
                //            sumni = sumni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                //            S_ni_lndi = S_ni_lndi + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value));
                //        }
                //        par_a1 = S_ni_lndi / sumni;
                //        textBox1.Text = Math.Round(par_a1, 3).ToString("0.000");
                //        for (int i = 0; i < N; i++)
                //            sum1 = sum1 + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1);
                //        par_b1 = Math.Sqrt(sum1 / (sumni - 1));
                //        textBox2.Text = Math.Round(par_b1, 3).ToString("0.000");
                //        for (double i = 0.1; i < X_max + r; i = i + 0.1)
                //        {
                //            S1 = 90;
                //            fDi = r / (par_b1 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a1) * (Math.Log(i) - par_a1) / (2 * par_b1 * par_b1)) * S1;
                //            chart2.Series[1].Points.AddXY(i, fDi);
                //        }
                //        richTextBox1.SelectedRtf = Properties.Resources.Bi_log_norm;
                //        textBox4.Text = Math.Round(par_b1, 3).ToString("0.000");
                //        textBox3.Text = Math.Round(par_a1, 3).ToString("0.000");
                //        textBox5.Text = Math.Round(par_b1, 3).ToString("0.000");
                //        richTextBox1.Width = 309;
                //        richTextBox1.Height = 130;
                //        par_a2 = 2.5;
                //        par_b2 = 0.5;
                //        S2 = 10;
                //        for (double i = 0.1; i < X_max + r; i = i + 0.1)
                //        {
                //            fDi = r / (par_b2 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a2) * (Math.Log(i) - par_a2) / (2 * par_b2 * par_b2)) * S2;
                //            chart2.Series[2].Points.AddXY(i, fDi);
                //            fDiCom = (r / (par_b1 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a1) * (Math.Log(i) - par_a1) / (2 * par_b1 * par_b1)) * S1) + (r / (par_b2 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a2) * (Math.Log(i) - par_a2) / (2 * par_b2 * par_b2)) * S2);
                //            chart2.Series[5].Points.AddXY(i, fDiCom);
                //        }
                //        R2 = 0;
                //        SSres = 0;
                //        SStot = 0;
                //        for (int i = 0; i < N; i++)
                //        {
                //            //if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) != 0)
                //            //{
                //            SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1)) * S1) - (r / (par_b2 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) / (2 * par_b2 * par_b2)) * S2)), 2);
                //            SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                //            // }
                //        }
                //        R2 = 1 - SSres / SStot;
                //        textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                //        textBox9.Visible = true;
                //        textBox9.BackColor = System.Drawing.Color.White;
                //        textBox9.Text = Math.Round(par_a2, 3).ToString("0.000");
                //        textBox10.Visible = true;
                //        textBox10.BackColor = System.Drawing.Color.White;
                //        textBox10.Text = Math.Round(par_b2, 3).ToString("0.000");
                //        textBox7.Text = Math.Round(par_b2, 3).ToString("0.000");
                //        textBox6.Text = Math.Round(par_a2, 3).ToString("0.000");
                //        textBox8.Text = Math.Round(par_b2, 3).ToString("0.000");
                //        textBox3.Location = new Point(247, 378);
                //        textBox7.Location = new Point(61, 455);
                //        textBox6.Location = new Point(247, 431);
                //        textBox8.Location = new Point(234, 454);
                //    }
                //}
                //catch
                //{
                //    dataGridView1.Rows[0].Cells[1].Value = dataGV_zap;
                //    if (Program.f1.Rus == true)
                //        MessageBox.Show("Невозможно построить бимодальное распределение в автоматическом режиме! Вам предлагается подобрать параметры в ручную", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //    if (Program.f1.Eng == true)
                //        MessageBox.Show("It is impossible to plot a bimodal distribution in automatic! You are offered to fit the parameters manually", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //    richTextBox1.Clear();
                //    sumni = 0;
                //    S_ni_lndi = 0;
                //    sum1 = 0;
                //    sum_Y = 0;
                //    for (int i = 0; i < N; i++)
                //    {
                //        sumni = sumni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                //        S_ni_lndi = S_ni_lndi + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value));
                //        sum_Y = sum_Y + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                //    }
                //    par_a1 = S_ni_lndi / sumni;
                //    textBox1.Text = Math.Round(par_a1, 3).ToString("0.000");
                //    for (int i = 0; i < N; i++)
                //        sum1 = sum1 + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1);
                //    par_b1 = Math.Sqrt(sum1 / (sumni - 1));
                //    textBox2.Text = Math.Round(par_b1, 3).ToString("0.000");
                //    for (double i = 0.1; i < X_max + r; i = i + 0.1)
                //    {
                //        S1 = 90;
                //        fDi = r / (par_b1 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a1) * (Math.Log(i) - par_a1) / (2 * par_b1 * par_b1)) * S1;
                //        chart2.Series[1].Points.AddXY(i, fDi);
                //    }
                //    richTextBox1.SelectedRtf = Properties.Resources.Bi_log_norm;
                //    textBox4.Text = Math.Round(par_b1, 3).ToString("0.000");
                //    textBox3.Text = Math.Round(par_a1, 3).ToString("0.000");
                //    textBox5.Text = Math.Round(par_b1, 3).ToString("0.000");
                //    richTextBox1.Width = 309;
                //    richTextBox1.Height = 130;
                //    par_a2 = 2.5;
                //    par_b2 = 0.5;
                //    S2 = 10;
                //    for (double i = 0.1; i < X_max + r; i = i + 0.1)
                //    {
                //        fDi = r / (par_b2 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a2) * (Math.Log(i) - par_a2) / (2 * par_b2 * par_b2)) * S2;
                //        chart2.Series[2].Points.AddXY(i, fDi);
                //        fDiCom = (r / (par_b1 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a1) * (Math.Log(i) - par_a1) / (2 * par_b1 * par_b1)) * S1) + (r / (par_b2 * i * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(i) - par_a2) * (Math.Log(i) - par_a2) / (2 * par_b2 * par_b2)) * S2);
                //        chart2.Series[5].Points.AddXY(i, fDiCom);
                //    }
                //    R2 = 0;
                //    SSres = 0;
                //    SStot = 0;
                //    for (int i = 0; i < N; i++)
                //    {
                //        //if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) != 0)
                //        //{
                //        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1)) * S1) - (r / (par_b2 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * (Math.Log(Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) / (2 * par_b2 * par_b2)) * S2)), 2);
                //        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y / N), 2);
                //        // }
                //    }
                //    R2 = 1 - SSres / SStot;
                //    textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                //    textBox9.Visible = true;
                //    textBox9.BackColor = System.Drawing.Color.White;
                //    textBox9.Text = Math.Round(par_a2, 3).ToString("0.000");
                //    textBox10.Visible = true;
                //    textBox10.BackColor = System.Drawing.Color.White;
                //    textBox10.Text = Math.Round(par_b2, 3).ToString("0.000");
                //    textBox7.Text = Math.Round(par_b2, 3).ToString("0.000");
                //    textBox6.Text = Math.Round(par_a2, 3).ToString("0.000");
                //    textBox8.Text = Math.Round(par_b2, 3).ToString("0.000");
                //    textBox3.Location = new Point(247, 378);
                //    textBox5.Location = new Point(234, 401);
                //    textBox7.Location = new Point(61, 455);
                //    textBox6.Location = new Point(247, 431);
                //    textBox8.Location = new Point(234, 454);
                //}

                
            }
        }

    private void кубическоеРаспределениеToolStripMenuItem_Click(object sender, EventArgs e)
        {           
            if (Convert.ToString(dataGridView1.Rows[0].Cells[0].Value) == string.Empty)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Размеры частиц не заданы! Попробуйте снова!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles size are not determined! Try again!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (Convert.ToString(dataGridView1.Rows[0].Cells[1].Value) == string.Empty)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Число частиц не задано! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles number are not determined! Try again!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            dataGridView2.Rows.Clear();
            flag_cub_raspr = true;
            CSR = false;
            chart2.Series[3].Points.Clear();
            chart2.Series[4].Points.Clear();
            if (!tabControl1.TabPages.ContainsKey("tabPage2"))
            {
                tabControl1.TabPages.Add(tabPage2);
                
            }
            tabControl1.SelectedTab = tabControl1.TabPages["tabPage2"];


            sum_Vol = 0;
                N = Convert.ToInt32(dataGridView1.RowCount);
                N30= Convert.ToInt32(dataGridView2.RowCount);
                if (N30 < N)
                    dataGridView2.Rows.Add(N - N30);
                while (dataGridView1.Rows[N - 1].Cells[0].Value == null)// || dataGridView1.Rows[N - 1].Cells[1].Value == null)
                    N = N - 1;
                
                if (Sum == 0)
                {
                    for (int i = 0; i < N; i++)
                    {
                        Sum = Sum + Convert.ToDouble(dataGridView1[1, i].Value); // сначал номер стобца, затем номер строки
                    }
                }

                for (int i = 0; i < N; i++)
                {
                    dataGridView2.Rows[i].Cells[0].Value = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                    Vol.Insert(i, Convert.ToDouble(dataGridView1[0, i].Value) * Convert.ToDouble(dataGridView1[1, i].Value));
                    dataGridView2.Rows[i].Cells[1].Value = Vol[i];
                    dataGridView2.Rows[i].Cells[3].Value = Math.Round(Math.Log(Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value)), 2);
                    sum_Vol = sum_Vol + Vol[i];
                }

                for (int i = 0; i < N; i++)
                    dataGridView2.Rows[i].Cells[2].Value = Math.Round(Vol[i] / sum_Vol * 100, 2);

                button6_Click(sender, e);
            

            if (CSR1 == true)
            {
                textBox15.Visible = false;
                textBox16.Visible = false;
                pictureBox7.Visible = false;
                pictureBox8.Visible = false;
                label12.Visible = false;
                label13.Visible = false;
            }           
            if (CSR2 == true)
            {
                textBox17.Visible = false;
                textBox18.Visible = false;
                pictureBox9.Visible = false;
                pictureBox10.Visible = false;
                label14.Visible = false;
                label15.Visible = false;
            }
            //dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[0]; // меняем фокус на другую ячейку с целью завершения редактирования последней набранной нами
            
        }

        private void линейноеРаспределениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabControl1.TabPages["tabPage1"];
            button2_Click(sender, e);            
        }

        private void Form3_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (Program.f1.form4)
            {
                Program.f1.f3_open = false;
                //this.Close();
            }
            else
            {
                if (Program.f1.Rus == true)
                {
                    DialogResult result = MessageBox.Show("Уже уходите?", "Уведомление", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        Program.f1.f3_open = false;
                        return;
                    }
                    if (result == DialogResult.No)
                    {
                        e.Cancel = true;
                    }
                }
                if (Program.f1.Eng == true)
                {
                    exit_Log_Norm = true;
                    Program.f5 = new Form5();
                    Program.f5.ShowDialog();
                    if (Program.f5.exite_f5 == true)
                    {
                        Program.f5.exite_f5 = false;
                        e.Cancel = true;
                    }
                    else
                        Program.f1.f3_open = false;
                }
            }
        }

        private void нормФцияГауссToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (flag_cub_raspr == true)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Функция описания в кубическом распределении недоступна!", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Fitting function in cubic distribution is not available!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (dataGridView1.Rows[0].Cells[2].Value == null)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Значения таблицы не заданы! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("The table values are missing! Try again", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (dataGridView1.Rows[0].Cells[1].Value == null)
            {
                if (Program.f1.Rus == true)
                    MessageBox.Show("Число частиц не задано! Попробуйте снова", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("Particles number are not determined! Try again!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            button4_Click = false;
            pictureBox2.Visible = false;
            //textBox11.Visible = false;
            pictureBox3.Visible = false;
            textBox12.Visible = false;
            //pictureBox4.Visible = false;
            label5.Visible = false;
            label9.Visible = false;
            textBox22.Visible = true;
            textBox22.Location = new Point(234, 480);
            R2 = 0;
            SSres = 0;
            SStot = 0;
            error_sumi = 0;
            kol_y_max = 0;
            chart2.Series[1].Points.Clear();
            chart2.Series[2].Points.Clear();
            chart2.Series[5].Points.Clear();
            richTextBox1.Clear();
            sumni = 0;
            sum_Y = 0;
            S_ni = 0;
            sum1 = 0;
            N1 = 0;
            S = 0;
            S1 = 0;
            S2 = 0;
            for (int i = 0; i < N; i++)
            {
                sumni = sumni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                S_ni = S_ni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                sum_Y = sum_Y + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
            }
            Dcp = S_ni / sumni;
            sumni = 0;
            try
            {
                y_min = 101;
                W_max: while (N1 < N - 1 && Convert.ToDouble(dataGridView1.Rows[N1].Cells[2].Value) < Convert.ToDouble(dataGridView1.Rows[N1 + 1].Cells[2].Value))
                    N1 = N1 + 1;
                y_max[kol_y_max] = Convert.ToDouble(dataGridView1.Rows[N1].Cells[2].Value);
                N_max[kol_y_max] = N1;
                kol_y_max = kol_y_max + 1;
                if (N1 < N - 2)
                    goto W_min;
                W_min: while (N1 < N - 1 && Convert.ToDouble(dataGridView1.Rows[N1].Cells[2].Value) >= Convert.ToDouble(dataGridView1.Rows[N1 + 1].Cells[2].Value))
                    N1 = N1 + 1;
                if (N1 < N - 2)
                    goto W_max;
                for (int i = 0; i <= kol_y_max; i++)
                {
                    for (int j = 0; j <= kol_y_max; j++)
                    {
                        if (y_max[j] <= y_max[j + 1])
                        {
                            double max = y_max[j + 1];
                            y_max[j + 1] = y_max[j];
                            y_max[j] = max;
                        }
                    }
                }
                y_max1 = y_max[0];
                y_max2 = y_max[1];
                for (int i = 0; i <= kol_y_max; i++)
                {
                    if (Convert.ToDouble(dataGridView1.Rows[N_max[i]].Cells[2].Value) == y_max1)
                        N_max1 = N_max[i];
                }
                for (int i = 0; i <= kol_y_max; i++)
                {
                    if (Convert.ToDouble(dataGridView1.Rows[N_max[i]].Cells[2].Value) == y_max2)
                        N_max2 = N_max[i];
                }
                if (N_max1 < N_max2)
                {
                    N1 = N_max1;
                    N2 = N_max2;
                }
                else
                {
                    N1 = N_max2;
                    N2 = N_max1;
                }
                for (int i = N1; i <= N2; i++)
                {
                    if (Convert.ToDouble(dataGridView1.Rows[i + 1].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) > r)
                    {
                        r1 = Convert.ToDouble(dataGridView1.Rows[i + 1].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                        N1_ = i;
                        N2_ = i + 1;
                    }

                }
                if (r < r1) // случай когда есть разрыв между столбцами с двумя распределениями
                {
                    for (int i = 0; i <= N1_; i++)
                    {
                        sumni = sumni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                        S_ni = S_ni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                        S1 = S1 + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                    }
                    par_a1 = S_ni / sumni;

                    for (int i = 0; i <= N1_; i++)
                        sum1 = sum1 + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a1) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a1);
                    par_b1 = Math.Sqrt(sum1 / (sumni - 1));

                    for (double i = 0.1; i < X_max + r; i = i + 0.1)
                    {
                        fDi = r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a1) * (i - par_a1) / (2 * par_b1 * par_b1)) * S1;
                        chart2.Series[1].Points.AddXY(i, fDi);
                        if (MaxShiftFDi1 < fDi)
                        {
                            MaxShiftFDi1 = fDi;
                            position_MaxShiftFDi1 = i;
                        }
                    }
                    richTextBox1.SelectedRtf = Properties.Resources.Bi_Gauss;
                    textBox4.Text = Math.Round(par_b1, 3).ToString("0.000");
                    textBox3.Text = Math.Round(par_a1, 3).ToString("0.000");
                    textBox5.Text = Math.Round(par_b1, 3).ToString("0.000");
                    textBox3.Location = new Point(225, 378);
                    textBox5.Location = new Point(220, 401);
                    richTextBox1.Width = 309;
                    richTextBox1.Height = 130;
                    sumni = 0;
                    S_ni = 0;
                    sum1 = 0;
                    for (int i = N2_; i < N - 1; i++)
                    {
                        sumni = sumni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                        S_ni = S_ni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                        S2 = S2 + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                    }
                    par_a2 = S_ni / sumni;

                    for (int i = N2_; i < N - 1; i++)
                        sum1 = sum1 + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a2) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a2);
                    par_b2 = Math.Sqrt(sum1 / (sumni - 1));


                    for (double i = 0.1; i < X_max + r; i = i + 0.1)
                    {
                        fDi = r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a2) * (i - par_a2) / (2 * par_b2 * par_b2)) * S2;
                        chart2.Series[2].Points.AddXY(i, fDi);
                        if (MaxShiftFDi2 < fDi)
                        {
                            MaxShiftFDi2 = fDi;
                            position_MaxShiftFDi2 = i;
                        }
                        fDiCom = (r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a1) * (i - par_a1) / (2 * par_b1 * par_b1)) * S1) + (r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a2) * (i - par_a2) / (2 * par_b2 * par_b2)) * S2);
                        chart2.Series[5].Points.AddXY(i, fDiCom);
                    }
                    for (int i = 0; i < N; i++)
                    {
                        //if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) != 0)
                        // {
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1)) * S1) - (r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) / (2 * par_b2 * par_b2)) * S2)),2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y/N), 2);
                        // }
                    }
                    R2 = 1 - SSres / SStot;
                    textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                }
                else // описывается обычный случай
                {
                    for (int i = N1; i <= N2; i++)
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) < Convert.ToDouble(dataGridView1.Rows[i + 1].Cells[2].Value) & Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) <= y_min)
                        {
                            y_min = Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                            N_min = i;
                        }
                    }
                    for (int i = 0; i < N; i++)
                        S = S + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                    Sum = 0;
                    for (int i = 0; i < N; i++)
                        Sum = Sum + Convert.ToDouble(dataGridView1[1, i].Value); // сначал номер стобца, затем номер строки               
                    N1 = N_min;
                    dataGV = Convert.ToDouble(dataGridView1.Rows[N1].Cells[1].Value);
                    dataGV_zap = Convert.ToDouble(dataGridView1.Rows[0].Cells[1].Value);
                    for (double j = 0; j <= dataGV; j = j + 0.1)
                    {
                        dataGridView1.Rows[N1].Cells[1].Value = j;
                        error_sum = 0;
                        sum1 = 0;
                        sumni = 0;
                        S_ni = 0;
                        S1 = 0;
                        for (int i = 0; i <= N1; i++)
                        {
                            sumni = sumni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                            S_ni = S_ni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                        }

                        par_a1 = S_ni / sumni;
                        for (int i = 0; i <= N1; i++)
                            sum1 = sum1 + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a1) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a1);
                        par_b1 = Math.Sqrt(sum1 / (sumni - 1));

                        for (double i = r; i < X_max + r; i = i + r)
                            fDi = r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a1) * (i - par_a1) / (2 * par_b1 * par_b1));
                        for (int i = 0; i < N1; i++)
                            S1 = S1 + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                        for (int i = N1; i < N; i++)
                            S1 = S1 + r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a1) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a1) / (2 * par_b1 * par_b1)) * S;
                        N2 = N1;

                        bi_modal[N2] = dataGV - Convert.ToDouble(dataGridView1.Rows[N1].Cells[1].Value);
                        for (int i = N2 + 1; i < N; i++)
                            bi_modal[i] = (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) / 100 - r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a1) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a1) / (2 * par_b1 * par_b1))) * Sum;
                        sumni = 0;
                        S_ni = 0;
                        sum1 = 0;
                        S2 = 0;
                        for (int i = N2; i < N; i++)
                        {
                            sumni = sumni + bi_modal[i];
                            S_ni = S_ni + bi_modal[i] * Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                        }
                        par_a2 = S_ni / sumni;
                        for (int i = N2; i < N; i++)
                            sum1 = sum1 + bi_modal[i] * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a2) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a2);
                        par_b2 = Math.Sqrt(sum1 / (sumni - 1));
                        for (double i = r; i < X_max + r; i = i + r)
                            fDi = r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a2) * (i - par_a2) / (2 * par_b2 * par_b2));
                        for (int i = 0; i < N2; i++)
                            S2 = S2 + r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a2) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a2) / (2 * par_b2 * par_b2)) * S;
                        S1 = S1 - S2;
                        S2 = S - S1;
                        for (int i = 0; i < N; i++)
                            error_sum = error_sum + Math.Abs(Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a1) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a1) / (2 * par_b1 * par_b1)) * S1) + (r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a1) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a1) / (2 * par_b1 * par_b1)) * S2));
                        error_sum1[error_sumi] = error_sum;
                        error_sumi = error_sumi + 1;
                    }
                    error_sum = error_sum1[0];
                    error_sumj1 = 0;
                    for (int i = 0; i < error_sumi - 1; i++)
                        if (error_sum > error_sum1[i + 1])
                        {
                            error_sum = error_sum1[i + 1];
                            error_sumj1 = i + 1;
                            error_sumj1 = error_sumj1 / 10;
                        }
                    dataGridView1.Rows[N1].Cells[1].Value = error_sumj1;
                    sumni = 0;
                    S_ni = 0;
                    S1 = 0;
                    sum1 = 0;
                    for (int i = 0; i <= N1; i++)
                    {
                        sumni = sumni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                        S_ni = S_ni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                    }
                    par_a1 = S_ni / sumni;

                    for (int i = 0; i <= N1; i++)
                        sum1 = sum1 + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a1) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a1);
                    par_b1 = Math.Sqrt(sum1 / (sumni - 1));

                    for (double i = r; i < X_max + r; i = i + r)
                        fDi = r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a1) * (i - par_a1) / (2 * par_b1 * par_b1));
                    for (int i = 0; i < N1; i++)
                        S1 = S1 + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                    for (int i = N1; i < N; i++)
                        S1 = S1 + r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a1) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a1) / (2 * par_b1 * par_b1)) * S;
                    dataGridView1.Rows[N2].Cells[1].Value = dataGV;
                    bi_modal[N2] = (dataGV - error_sumj1);
                    for (int i = N2 + 1; i < N; i++)
                        bi_modal[i] = (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) / 100 - r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a1) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a1) / (2 * par_b1 * par_b1))) * Sum;
                    richTextBox1.SelectedRtf = Properties.Resources.Bi_Gauss;
                    textBox4.Text = Math.Round(par_b1, 3).ToString("0.000");
                    textBox3.Text = Math.Round(par_a1, 3).ToString("0.000");
                    textBox5.Text = Math.Round(par_b1, 3).ToString("0.000");
                    richTextBox1.Width = 309;
                    richTextBox1.Height = 130;
                    sumni = 0;
                    S_ni = 0;
                    sum1 = 0;
                    S2 = 0;
                    for (int i = N2; i < N; i++)
                    {
                        sumni = sumni + bi_modal[i];
                        S_ni = S_ni + bi_modal[i] * Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                    }
                    par_a2 = S_ni / sumni;

                    for (int i = N2; i < N; i++)
                        sum1 = sum1 + bi_modal[i] * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a2) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a2);
                    par_b2 = Math.Sqrt(sum1 / (sumni - 1));

                    for (double i = r; i < X_max + r; i = i + r)
                        fDi = r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a2) * (i - par_a2) / (2 * par_b2 * par_b2));
                    for (int i = 0; i < N2; i++)
                        S2 = S2 + r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a2) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a2) / (2 * par_b2 * par_b2)) * S;
                    S1 = S1 - S2;
                    S2 = S - S1;
                    textBox3.Location = new Point(225, 378);
                    textBox5.Location = new Point(220, 401);

                    for (double i = 0.1; i < X_max + r; i = i + 0.1)
                    {
                        fDi = r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a1) * (i - par_a1) / (2 * par_b1 * par_b1)) * S1;
                        chart2.Series[1].Points.AddXY(i, fDi);
                        if (MaxShiftFDi1 < fDi)
                        {
                            MaxShiftFDi1 = fDi;
                            position_MaxShiftFDi1 = i;
                        }
                        fDi = r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a2) * (i - par_a2) / (2 * par_b2 * par_b2)) * S2;
                        chart2.Series[2].Points.AddXY(i, fDi);
                        if (MaxShiftFDi2 < fDi)
                        {
                            MaxShiftFDi2 = fDi;
                            position_MaxShiftFDi2 = i;
                        }
                        fDiCom = (r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a1) * (i - par_a1) / (2 * par_b1 * par_b1)) * S1) + (r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a2) * (i - par_a2) / (2 * par_b2 * par_b2)) * S2);
                        chart2.Series[5].Points.AddXY(i, fDiCom);
                    }
                    for (int i = 0; i < N; i++)
                    {
                        //if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) != 0)
                        // {
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1)) * S1) - (r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) / (2 * par_b2 * par_b2)) * S2)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y/N), 2);
                        // }
                    }
                    R2 = 1 - SSres / SStot;
                    textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                }
                if (double.IsNaN(par_a1) || double.IsNaN(par_a2) || double.IsNaN(par_b1) || double.IsNaN(par_b2) || double.IsNaN(S1) || double.IsNaN(S2))
                {
                    dataGridView1.Rows[N1].Cells[1].Value = dataGV_zap;
                    if (Program.f1.Rus == true)
                        MessageBox.Show("Невозможно построить бимодальное распределение в автоматическом режиме! Вам предлагается подобрать параметры в ручную", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (Program.f1.Eng == true)
                        MessageBox.Show("It is impossible to plot a bimodal distribution in automatic! You are offered to fit the parameters manually", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    richTextBox1.Clear();
                    sumni = 0;
                    S_ni = 0;
                    sum1 = 0;
                    sum_Y = 0;
                    for (int i = 0; i < N; i++)
                    {
                        sumni = sumni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                        S_ni = S_ni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                        sum_Y = sum_Y + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                    }
                    par_a1 = S_ni / sumni;

                    for (int i = 0; i < N; i++)
                        sum1 = sum1 + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a1) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a1);
                    par_b1 = Math.Sqrt(sum1 / (sumni - 1));

                    for (double i = 0.1; i < X_max + r; i = i + 0.1)
                    {
                        S1 = 90;
                        fDi = r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a1) * (i - par_a1) / (2 * par_b1 * par_b1)) * S1;
                        chart2.Series[1].Points.AddXY(i, fDi);
                    }
                    richTextBox1.SelectedRtf = Properties.Resources.Bi_Gauss;
                    textBox4.Text = Math.Round(par_b1, 3).ToString("0.000");
                    textBox3.Text = Math.Round(par_a1, 3).ToString("0.000");
                    textBox5.Text = Math.Round(par_b1, 3).ToString("0.000");
                    richTextBox1.Width = 309;
                    richTextBox1.Height = 130;
                    par_a2 = par_a1 * 1.5;
                    par_b2 = par_b1;
                    S2 = 10;
                    for (double i = 0.1; i < X_max + r; i = i + 0.1)
                    {
                        fDi = r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a2) * (i - par_a2) / (2 * par_b2 * par_b2)) * S2;
                        chart2.Series[2].Points.AddXY(i, fDi);
                        fDiCom = (r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a1) * (i - par_a1) / (2 * par_b1 * par_b1)) * S1) + (r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a2) * (i - par_a2) / (2 * par_b2 * par_b2)) * S2);
                        chart2.Series[5].Points.AddXY(i, fDiCom);
                    }
                    for (int i = 0; i < N; i++)
                    {
                        //if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) != 0)
                        // {
                        SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1)) * S1) - (r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) / (2 * par_b2 * par_b2)) * S2)), 2);
                        SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y/N), 2);
                        // }
                    }
                    R2 = 1 - SSres / SStot;
                    textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                    textBox5.Location = new Point(220, 401);
                    textBox3.Location = new Point(225, 378);
                }
            }
            catch
            {
                dataGridView1.Rows[N1].Cells[1].Value = dataGV_zap;
                if (Program.f1.Rus == true)
                    MessageBox.Show("Невозможно построить бимодальное распределение в автоматическом режиме! Вам предлагается подобрать параметры в ручную", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (Program.f1.Eng == true)
                    MessageBox.Show("It is impossible to plot a bimodal distribution in automatic! You are offered to fit the parameters manually", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                richTextBox1.Clear();
                sumni = 0;
                S_ni = 0;
                sum1 = 0;
                sum_Y = 0;
                for (int i = 0; i < N; i++)
                {
                    sumni = sumni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                    S_ni = S_ni + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                    sum_Y = sum_Y + Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                }
                par_a1 = S_ni / sumni;

                for (int i = 0; i < N; i++)
                    sum1 = sum1 + Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a1) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value) - par_a1);
                par_b1 = Math.Sqrt(sum1 / (sumni - 1));

                for (double i = 0.1; i < X_max + r; i = i + 0.1)
                {
                    S1 = 90;
                    fDi = r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a1) * (i - par_a1) / (2 * par_b1 * par_b1)) * S1;
                    chart2.Series[1].Points.AddXY(i, fDi);
                }
                richTextBox1.SelectedRtf = Properties.Resources.Bi_Gauss;
                textBox4.Text = Math.Round(par_b1, 3).ToString("0.000");
                textBox3.Text = Math.Round(par_a1, 3).ToString("0.000");
                textBox5.Text = Math.Round(par_b1, 3).ToString("0.000");
                richTextBox1.Width = 309;
                richTextBox1.Height = 130;
                par_a2 = par_a1 * 1.5;
                par_b2 = par_b1;
                S2 = 10;
                for (double i = 0.1; i < X_max + r; i = i + 0.1)
                {
                    fDi = r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a2) * (i - par_a2) / (2 * par_b2 * par_b2)) * S2;
                    chart2.Series[2].Points.AddXY(i, fDi);
                    fDiCom = (r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a1) * (i - par_a1) / (2 * par_b1 * par_b1)) * S1) + (r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * (i - par_a2) * (i - par_a2) / (2 * par_b2 * par_b2)) * S2);
                    chart2.Series[5].Points.AddXY(i, fDiCom);
                }
                for (int i = 0; i < N; i++)
                {
                    //if (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) != 0)
                    // {
                    SSres = SSres + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - (r / (par_b1 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a1) / (2 * par_b1 * par_b1)) * S1) - (r / (par_b2 * Math.Sqrt(2 * Math.PI)) * Math.Exp(-1 * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) * ((Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value)) - par_a2) / (2 * par_b2 * par_b2)) * S2)), 2);
                    SStot = SStot + Math.Pow((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) - sum_Y/N), 2);
                    // }
                }
                R2 = 1 - SSres / SStot;
                textBox22.Text = Math.Round(R2, 4).ToString("0.0000");
                textBox5.Location = new Point(220, 401);
                textBox3.Location = new Point(225, 378);
            }
            Norm = true;
            LogNorm = false;            
            button3_Click = true;
            D10 = par_a1;
            chart2.Series[1].IsVisibleInLegend = true;
            if (Program.f1.Rus == true)
                chart2.Series[1].Name = "Гаусс 1";
            if (Program.f1.Eng == true)
                chart2.Series[1].Name = "Gauss 1";
            chart2.Series[2].IsVisibleInLegend = true;
            if (Program.f1.Rus == true)
                chart2.Series[2].Name = "Гаусс 2";
            if (Program.f1.Eng == true)
                chart2.Series[2].Name = "Gauss 2";
        }      
       

        

        

        

        

        

        

        
    }
}
