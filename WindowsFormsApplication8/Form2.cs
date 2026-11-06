using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace WindowsFormsApplication8
{
    
    public partial class Form2 : Form
    {
        Form1 parentForm;
        Pen myPen = new Pen(Color.Red, 2);
        Point lastpoint = Point.Empty;
        //Graphics gr;
        private int clickCountForm2 = 0;
        private double x_MouseDown, y_MouseDown;
        private int W;
        private int H;
        private int Px, Py, P, Rc;      
        public Image img;
        private bool MD=false;
        public int I = 1;
        public int num_line=0;
        private double L, xc, yc, coeff;

        public Form2(Form1 parent)
        {
            Program.f2 = this;
            InitializeComponent();
            parentForm = parent; 
            

        }

        private void Form2_Deactivate(object sender, EventArgs e)
        {
            if (Program.f1.FormLineProphile)
                Program.f10.TopMost = false;
        }

        



        //public class ColoredLine
        // {
        //public int Previous_MouseMoveXF, Previous_MouseMoveYF, Previous_MouseMoveXT, Previous_MouseMoveYT;
        //public Color colorOfLine;

        // public ColoredLine(int fx, int fy, int tx, int ty, Color _color)
        // {
        // Previous_MouseMoveXF = fx;
        // Previous_MouseMoveYF = fy;
        // Previous_MouseMoveXT = tx;
        // Previous_MouseMoveYT = ty;
        // colorOfLine = _color;
        //}
        // }

        [DllImport("kernel32.dll")]// // отнимаем у приложения лишнюю непотребляемую в данный момент память
        static extern bool SetProcessWorkingSetSize(IntPtr hProcess,
int dwMinimumWorkingSetSize, int dwMaximumWorkingSetSize);

        private void Form2_Load(object sender, EventArgs e)
        {
            dataGridView1.Rows.Add();
            Program.f1.dataGridView1.Rows.Add();
            openImage();    
            
        }

        public void openImage()
        {
            using (MemoryStream ms = new MemoryStream(CommunicateClass.imageData))
            {
                img = Image.FromStream(ms);
                panel1.BackgroundImage = img;
                //pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
                W = panel1.BackgroundImage.Width;
                H = panel1.BackgroundImage.Height;
                double W_H = Convert.ToDouble(W) / Convert.ToDouble(H);
                W_H = Math.Round(W_H, 5);
                P = Convert.ToInt32(Program.f1.P * 100000);
                panel1.Width = panel1.Height * P / 100000;
                Px = panel1.Width;
                Py = panel1.Height;
                this.Size = new System.Drawing.Size(Px + 15, Py + 40);
                chart1.ChartAreas[0].Position.Width = 100;
                chart1.ChartAreas[0].Position.Height = 100;

                for (int i = I; i > 1; i--)
                {
                    chart1.Series[I].Points.Clear();
                    chart1.Series.Remove(chart1.Series[I]);
                    dataGridView1.Rows.Remove(dataGridView1.Rows[num_line]);
                    dataGridView1.Rows.Remove(dataGridView1.Rows[num_line - 1]);
                    num_line = num_line - 2;
                    I = I - 1;
                }

                for (int i = 0; i <= parentForm.num_line; i++)
                {
                    dataGridView1.Rows.Add();
                }

                dataGridView1.Rows[0].Cells[0].Value = 0;
                dataGridView1.Rows[0].Cells[1].Value = 0;
                dataGridView1.Rows[1].Cells[0].Value = Program.f1.magnificX;
                dataGridView1.Rows[1].Cells[1].Value = Program.f1.magnificY;


                chart1.ChartAreas[0].AxisX.Minimum = 0;
                chart1.ChartAreas[0].AxisY.Minimum = 0;
                chart1.ChartAreas[0].AxisX.Maximum = Program.f1.magnificX;
                chart1.ChartAreas[0].AxisY.Maximum = Program.f1.magnificY;

                for (int i = 0; i <= 1; i++)
                {
                    chart1.Series[0].Points.AddXY(dataGridView1.Rows[i].Cells[0].Value, dataGridView1.Rows[i].Cells[1].Value);

                }
                chart1.Series[0].Points.Clear();
                chart1.Series[0].Color = Color.Black;

                if (Program.f1.FormLineProphile)
                {
                    while (num_line != parentForm.num_line - 2)
                    {
                        dataGridView1.Rows[num_line].Cells[0].Value = Convert.ToDouble(parentForm.dataGridView1.Rows[num_line].Cells[0].Value) - parentForm.X_chart;
                        dataGridView1.Rows[num_line].Cells[1].Value = Convert.ToDouble(parentForm.dataGridView1.Rows[num_line].Cells[1].Value) - parentForm.H + parentForm.magnificY + parentForm.Y_chart;
                        num_line = num_line + 1;
                        dataGridView1.Rows[num_line].Cells[0].Value = Convert.ToDouble(parentForm.dataGridView1.Rows[num_line].Cells[0].Value) - parentForm.X_chart;
                        dataGridView1.Rows[num_line].Cells[1].Value = Convert.ToDouble(parentForm.dataGridView1.Rows[num_line].Cells[1].Value) - parentForm.H + parentForm.magnificY + parentForm.Y_chart;
                        I = I + 1;
                        chart1.Series.Add(Convert.ToString(I));
                        chart1.Series[I].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                        chart1.Series[I].Color = Color.Red;
                        for (int i = num_line - 1; i <= num_line; i++)
                            chart1.Series[I].Points.AddXY(dataGridView1.Rows[i].Cells[0].Value, dataGridView1.Rows[i].Cells[1].Value);
                        num_line = num_line + 1;
                    }
                    dataGridView1.Rows[num_line].Cells[0].Value = Convert.ToDouble(parentForm.dataGridView1.Rows[num_line].Cells[0].Value) - parentForm.X_chart;
                    dataGridView1.Rows[num_line].Cells[1].Value = Convert.ToDouble(parentForm.dataGridView1.Rows[num_line].Cells[1].Value) - parentForm.H + parentForm.magnificY + parentForm.Y_chart;
                    num_line = num_line + 1;
                    dataGridView1.Rows[num_line].Cells[0].Value = Convert.ToDouble(parentForm.dataGridView1.Rows[num_line].Cells[0].Value) - parentForm.X_chart;
                    dataGridView1.Rows[num_line].Cells[1].Value = Convert.ToDouble(parentForm.dataGridView1.Rows[num_line].Cells[1].Value) - parentForm.H + parentForm.magnificY + parentForm.Y_chart;                  
                    I = I + 1;
                    chart1.Series.Add(Convert.ToString(I));
                    chart1.Series[I].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                    chart1.Series[I].MarkerSize = 3;
                    chart1.Series[I].MarkerStyle = System.Windows.Forms.DataVisualization.Charting.MarkerStyle.Square;
                    chart1.Series[I].MarkerColor = Color.White;
                    chart1.Series[I].MarkerBorderColor = Color.Aqua;
                    chart1.Series[I].Color = Color.Aqua;
                    for (int i = num_line - 1; i <= num_line; i++)
                        chart1.Series[I].Points.AddXY(dataGridView1.Rows[i].Cells[0].Value, dataGridView1.Rows[i].Cells[1].Value);
                    num_line = num_line + 1;

                }
                else
                {
                    while (num_line != parentForm.num_line)
                    {
                        dataGridView1.Rows[num_line].Cells[0].Value = Convert.ToDouble(parentForm.dataGridView1.Rows[num_line].Cells[0].Value) - parentForm.X_chart;
                        dataGridView1.Rows[num_line].Cells[1].Value = Convert.ToDouble(parentForm.dataGridView1.Rows[num_line].Cells[1].Value) - parentForm.H + parentForm.magnificY + parentForm.Y_chart;
                        num_line = num_line + 1;
                        dataGridView1.Rows[num_line].Cells[0].Value = Convert.ToDouble(parentForm.dataGridView1.Rows[num_line].Cells[0].Value) - parentForm.X_chart;
                        dataGridView1.Rows[num_line].Cells[1].Value = Convert.ToDouble(parentForm.dataGridView1.Rows[num_line].Cells[1].Value) - parentForm.H + parentForm.magnificY + parentForm.Y_chart;
                        I = I + 1;
                        chart1.Series.Add(Convert.ToString(I));
                        chart1.Series[I].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;                       
                            chart1.Series[I].Color = Color.Red;
                        for (int i = num_line - 1; i <= num_line; i++)
                            chart1.Series[I].Points.AddXY(dataGridView1.Rows[i].Cells[0].Value, dataGridView1.Rows[i].Cells[1].Value);
                        num_line = num_line + 1;
                    }
                }

            }

            SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle, -1, -1);
        }

        private void chart1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ++clickCountForm2;
                switch (clickCountForm2)
                {
                    case 1:
                        if (Program.f1.calibrateImage)
                            Program.f9.Activate();                       
                        if (Program.f1.clickLineProfile != 1)
                        {
                            try
                            {
                                if (parentForm.calibrateImage)
                                    if (parentForm.N_calImage == 1)
                                    {
                                        if (parentForm.Rus == true)
                                            MessageBox.Show("Сначала удалите предыдущее измерение", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        if (parentForm.Eng == true)
                                            MessageBox.Show("Delete the previous measurement first", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        clickCountForm2 = 0;
                                        return;
                                    }
                                dataGridView1.Rows.Add();
                                dataGridView1.Rows.Add();
                                parentForm.dataGridView1.Rows.Add();
                                parentForm.dataGridView1.Rows.Add();
                                MD = true;
                                lastpoint = e.Location;
                                x_MouseDown = Math.Round(chart1.ChartAreas[0].AxisX.PixelPositionToValue(lastpoint.X), 0);
                                y_MouseDown = Math.Round(chart1.ChartAreas[0].AxisY.PixelPositionToValue(lastpoint.Y), 0);
                                dataGridView1.Rows[num_line].Cells[0].Value = x_MouseDown;
                                dataGridView1.Rows[num_line].Cells[1].Value = y_MouseDown;
                                parentForm.dataGridView1.Rows[parentForm.num_line].Cells[0].Value = x_MouseDown + parentForm.X_chart;
                                parentForm.dataGridView1.Rows[parentForm.num_line].Cells[1].Value = y_MouseDown + parentForm.H-parentForm.magnificY- parentForm.Y_chart;
                                num_line = num_line + 1;
                                parentForm.num_line = parentForm.num_line + 1;
                                I = I + 1;
                                parentForm.I = parentForm.I + 1;
                                chart1.Series.Add(Convert.ToString(I));
                                parentForm.chart1.Series.Add(Convert.ToString(parentForm.I));
                                chart1.Series[I].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                                chart1.Series[I].Color = Color.Red;
                                parentForm.chart1.Series[parentForm.I].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                                parentForm.chart1.Series[parentForm.I].Color = Color.Red;
                            }
                            catch
                            {
                                return;
                            }
                        }

                        if (Program.f1.clickLineProfile == 1)
                        {
                            try
                            {
                                if (Program.f1.FormLineProphile)
                                {
                                    chart1.Series[I].Points.Clear();
                                    Program.f1.chart1.Series[Program.f1.I].Points.Clear();
                                    chart1.Series.Remove(chart1.Series[I]);
                                    Program.f1.chart1.Series.Remove(Program.f1.chart1.Series[Program.f1.I]);
                                    dataGridView1.Rows.Remove(dataGridView1.Rows[num_line]);
                                    dataGridView1.Rows.Remove(dataGridView1.Rows[num_line - 1]);
                                    num_line = num_line - 2;
                                    
                                    //Program.f1.dataGridView1.Rows.Remove(Program.f1.dataGridView1.Rows[Program.f1.num_line]);
                                    //Program.f1.dataGridView1.Rows.Remove(Program.f1.dataGridView1.Rows[Program.f1.num_line - 1]);
                                    Program.f1.num_line = Program.f1.num_line - 2;
                                    Program.f1.kol_count_di = Program.f1.kol_count_di - 1;
                                    I = I - 1;
                                    Program.f1.I = Program.f1.I - 1;
                                }
                                dataGridView1.Rows.Add();
                                //Program.f1.dataGridView1.Rows.Add();
                                MD = true;
                                lastpoint = e.Location;
                                x_MouseDown = Math.Round(chart1.ChartAreas[0].AxisX.PixelPositionToValue(lastpoint.X), 0);
                                y_MouseDown = Math.Round(chart1.ChartAreas[0].AxisY.PixelPositionToValue(lastpoint.Y), 0);
                                dataGridView1.Rows[num_line].Cells[0].Value = x_MouseDown;
                                dataGridView1.Rows[num_line].Cells[1].Value = y_MouseDown;
                                parentForm.dataGridView1.Rows[parentForm.num_line].Cells[0].Value = x_MouseDown + parentForm.X_chart;
                                parentForm.dataGridView1.Rows[parentForm.num_line].Cells[1].Value = y_MouseDown + parentForm.H - parentForm.magnificY - parentForm.Y_chart; ;
                                num_line = num_line + 1;
                                parentForm.num_line = parentForm.num_line + 1;
                                I = I + 1;
                                parentForm.I = parentForm.I + 1;
                                chart1.Series.Add(Convert.ToString(I));
                                parentForm.chart1.Series.Add(Convert.ToString(parentForm.I));

                                chart1.Series[I].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                                chart1.Series[I].MarkerSize = 3;
                                chart1.Series[I].MarkerStyle = System.Windows.Forms.DataVisualization.Charting.MarkerStyle.Square;
                                chart1.Series[I].MarkerColor = Color.White;
                                chart1.Series[I].MarkerBorderColor = Color.Aqua;
                                chart1.Series[I].Color = Color.Aqua;

                                Program.f1.chart1.Series[Program.f1.I].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                                Program.f1.chart1.Series[Program.f1.I].MarkerSize = 3;
                                Program.f1.chart1.Series[Program.f1.I].MarkerStyle = System.Windows.Forms.DataVisualization.Charting.MarkerStyle.Square;
                                Program.f1.chart1.Series[Program.f1.I].MarkerColor = Color.White;
                                Program.f1.chart1.Series[Program.f1.I].MarkerBorderColor = Color.Aqua;
                                Program.f1.chart1.Series[Program.f1.I].Color = Color.Aqua;
                            }

                            catch
                            {
                                return;
                            }
                        }

                            break;
                        
                    case 2:
                        if (Program.f1.calibrateImage)
                            Program.f9.Activate();
                        if (Program.f1.FormLineProphile)
                            Program.f10.Activate();
                        try
                            {

                                MD = false;
                                clickCountForm2 = 0;
                                dataGridView1.Rows.Add();
                                Program.f1.Di = Math.Round(Math.Sqrt((Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value)) * (Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value)) + (Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value)) * (Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value))), 2);
                                Program.f1.Di = Program.f1.Di * Program.f1.n / Program.f1.n_const;
                                parentForm.dii.Insert(Program.f1.kol_count_di, Program.f1.Di);
                                Program.f1.kol_count_di = Program.f1.kol_count_di + 1;
                                if (parentForm.calibrateImage)
                                {
                                    Program.f9.textBox1.Text = Math.Round(Program.f1.Di, 2).ToString();
                                    parentForm.N_calImage = parentForm.N_calImage + 1;
                                    parentForm.number_N_meas_fault = parentForm.number_N_meas_fault + 1;
                                    parentForm.N_meas_faultt.Insert(parentForm.kol_meas_fault, parentForm.number_N_meas_fault);
                                }
                                else if (Program.f1.clickLineProfile != 1)
                                {
                                    parentForm.textBox1.Text = Math.Round(Program.f1.Di, 2).ToString();
                                    parentForm.textBox7.Text = Program.f1.kol_count_di.ToString();
                                    parentForm.number_N_meas_fault = parentForm.number_N_meas_fault + 1;
                                    parentForm.N_meas_faultt.Insert(parentForm.kol_meas_fault, parentForm.number_N_meas_fault);
                                }
                                chart1.Series[0].Points.Clear();
                                chart1.Series[1].Points.Clear();
                                for (int i = parentForm.num_line - 1; i <= parentForm.num_line; i++)
                                    parentForm.chart1.Series[parentForm.I].Points.AddXY(parentForm.dataGridView1.Rows[i].Cells[0].Value, parentForm.dataGridView1.Rows[i].Cells[1].Value);
                                parentForm.num_line = parentForm.num_line + 1;
                                num_line = num_line + 1;
                                //parentForm.number_N_meas_fault = parentForm.number_N_meas_fault + 1;
                                //parentForm.N_meas_faultt.Insert(parentForm.kol_meas_fault, parentForm.number_N_meas_fault);

                            if (Program.f1.clickLineProfile == 1)
                            {
                                if (Program.f1.FormLineProphile)
                                {
                                    Program.f1.dataGridView4.Rows.Clear();
                                }
                                //panel1.BackgroundImage = img;
                                Program.f1.k = (Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 2].Cells[1].Value)) / (Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 2].Cells[0].Value));
                                Program.f1.b = Convert.ToDouble(dataGridView1.Rows[num_line - 2].Cells[1].Value) - Program.f1.k * Convert.ToDouble(dataGridView1.Rows[num_line - 2].Cells[0].Value);
                                if (Math.Abs(Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[1].Value) - Convert.ToInt32(dataGridView1.Rows[num_line - 1].Cells[1].Value)) > Math.Abs(Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[0].Value) - Convert.ToInt32(dataGridView1.Rows[num_line - 1].Cells[0].Value)))

                                {
                                    Program.f1.col_pix = Math.Abs(Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[1].Value) - Convert.ToInt32(dataGridView1.Rows[num_line - 1].Cells[1].Value));
                                    coeff = Program.f1.Di / Program.f1.col_pix;                                    Color pixel;
                                    pixel = (panel1.BackgroundImage as Bitmap).GetPixel(Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[0].Value), H - Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[1].Value));
                                    Program.f1.dataGridView4.Rows.Add();
                                    Program.f1.dataGridView4.Rows[0].Cells[2].Value = (pixel.R + pixel.G + pixel.B) / 3;
                                    Program.f1.dataGridView4.Rows[0].Cells[0].Value = 0;
                                    for (int i = 1; i < Program.f1.col_pix; i++)
                                    {
                                        if (Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[1].Value) > Convert.ToInt32(dataGridView1.Rows[num_line - 1].Cells[1].Value))
                                            Program.f1.YY = Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[1].Value) - i;
                                        else
                                            Program.f1.YY = Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[1].Value) + i;

                                        Program.f1.XX = Convert.ToInt32((Program.f1.YY - Program.f1.b) / Program.f1.k);
                                        pixel = (panel1.BackgroundImage as Bitmap).GetPixel(Program.f1.XX, H - Program.f1.YY);
                                        Program.f1.dataGridView4.Rows.Add();
                                        Program.f1.dataGridView4.Rows[i].Cells[2].Value = (pixel.R+pixel.G+pixel.B)/3;
                                        Program.f1.dataGridView4.Rows[i].Cells[0].Value = i * coeff;
                                    }
                                    pixel = (panel1.BackgroundImage as Bitmap).GetPixel(Convert.ToInt32(dataGridView1.Rows[num_line - 1].Cells[0].Value), H - Convert.ToInt32(dataGridView1.Rows[num_line - 1].Cells[1].Value));
                                    Program.f1.dataGridView4.Rows[Program.f1.col_pix].Cells[2].Value = (pixel.R + pixel.G + pixel.B) / 3;
                                    Program.f1.dataGridView4.Rows[Program.f1.col_pix].Cells[0].Value = Program.f1.col_pix * coeff;
                                    Program.f1.Y_min = Convert.ToInt32(Program.f1.dataGridView4.Rows[0].Cells[2].Value);
                                    Program.f1.Y_max = Convert.ToInt32(Program.f1.dataGridView4.Rows[0].Cells[2].Value);

                                    for (int i = 1; i <= Program.f1.col_pix; i++)
                                    {
                                        if (Convert.ToInt32(Program.f1.dataGridView4.Rows[i].Cells[2].Value) < Program.f1.Y_min)
                                            Program.f1.Y_min = Convert.ToInt32(Program.f1.dataGridView4.Rows[i].Cells[2].Value);
                                        if (Convert.ToInt32(Program.f1.dataGridView4.Rows[i].Cells[2].Value) > Program.f1.Y_max)
                                            Program.f1.Y_max = Convert.ToInt32(Program.f1.dataGridView4.Rows[i].Cells[2].Value);
                                    }

                                }
                                else
                                {
                                    Program.f1.col_pix = Math.Abs(Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[0].Value) - Convert.ToInt32(dataGridView1.Rows[num_line - 1].Cells[0].Value));
                                    coeff = Program.f1.Di / Program.f1.col_pix;
                                    Color pixel;
                                    pixel = (panel1.BackgroundImage as Bitmap).GetPixel(Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[0].Value), H - Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[1].Value));
                                    Program.f1.dataGridView4.Rows.Add();
                                    Program.f1.dataGridView4.Rows[0].Cells[2].Value = (pixel.R + pixel.G + pixel.B) / 3;
                                    Program.f1.dataGridView4.Rows[0].Cells[0].Value = 0;
                                    for (int i = 1; i < Program.f1.col_pix; i++)
                                    {
                                        if (Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[0].Value) > Convert.ToInt32(dataGridView1.Rows[num_line - 1].Cells[0].Value))
                                            Program.f1.XX = Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[0].Value) - i;
                                        else
                                            Program.f1.XX = Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[0].Value) + i;

                                        Program.f1.YY = Convert.ToInt32(Program.f1.k * Program.f1.XX + Program.f1.b);
                                        pixel = (panel1.BackgroundImage as Bitmap).GetPixel(Program.f1.XX, H - Program.f1.YY);
                                        Program.f1.dataGridView4.Rows.Add();
                                        Program.f1.dataGridView4.Rows[i].Cells[2].Value = (pixel.R + pixel.G + pixel.B) / 3;
                                        Program.f1.dataGridView4.Rows[i].Cells[0].Value = i * coeff;
                                    }
                                    pixel = (panel1.BackgroundImage as Bitmap).GetPixel(Convert.ToInt32(dataGridView1.Rows[num_line - 1].Cells[0].Value), H - Convert.ToInt32(dataGridView1.Rows[num_line - 1].Cells[1].Value));
                                    Program.f1.dataGridView4.Rows[Program.f1.col_pix].Cells[2].Value = (pixel.R + pixel.G + pixel.B) / 3;
                                    Program.f1.dataGridView4.Rows[Program.f1.col_pix].Cells[0].Value = Program.f1.col_pix * coeff;
                                    Program.f1.Y_min = Convert.ToInt32(Program.f1.dataGridView4.Rows[0].Cells[2].Value);
                                    Program.f1.Y_max = Convert.ToInt32(Program.f1.dataGridView4.Rows[0].Cells[2].Value);
                                    for (int i = 1; i <= Program.f1.col_pix; i++)
                                    {
                                        if (Convert.ToInt32(Program.f1.dataGridView4.Rows[i].Cells[2].Value) < Program.f1.Y_min)
                                            Program.f1.Y_min = Convert.ToInt32(Program.f1.dataGridView4.Rows[i].Cells[2].Value);
                                        if (Convert.ToInt32(Program.f1.dataGridView4.Rows[i].Cells[2].Value) > Program.f1.Y_max)
                                            Program.f1.Y_max = Convert.ToInt32(Program.f1.dataGridView4.Rows[i].Cells[2].Value);
                                    }

                                }

                                if (Program.f1.FormLineProphile)
                                    Program.f10.LineProphile();
                                else
                                {
                                    Program.f1.f10 = new Form10();
                                    Program.f1.f10.Owner = this;
                                    Program.f1.f10.Show();
                                }
                            }
                        }
                            catch
                            {
                                return;
                            }
                        

                        
                        break;


                }
            }
            if (e.Button == MouseButtons.Right)
            {
                if (clickCountForm2 == 1)
                {
                    if (Program.f1.calibrateImage)
                        Program.f9.Activate();                    
                    MD = false;
                    clickCountForm2 = 0;
                    chart1.Series[0].Points.Clear();
                    chart1.Series[1].Points.Clear();
                    chart1.Series[I].Points.Clear();
                    chart1.Series.Remove(chart1.Series[I]);
                    dataGridView1.Rows.Remove(dataGridView1.Rows[num_line]);
                    I = I - 1;
                    num_line = num_line - 1;

                    parentForm.chart1.Series[parentForm.I].Points.Clear();
                    parentForm.chart1.Series.Remove(parentForm.chart1.Series[parentForm.I]);
                    parentForm.dataGridView1.Rows.Remove(parentForm.dataGridView1.Rows[parentForm.num_line]);
                    parentForm.I = parentForm.I - 1;
                    parentForm.num_line = parentForm.num_line - 1;
                }
            }
        }

        private void chart1_MouseMove(object sender, MouseEventArgs e)
        {
            if (MD)
            {
                    try
                    {
                        lastpoint = e.Location;
                        x_MouseDown = Math.Round(chart1.ChartAreas[0].AxisX.PixelPositionToValue(lastpoint.X), 0);
                        y_MouseDown = Math.Round(chart1.ChartAreas[0].AxisY.PixelPositionToValue(lastpoint.Y), 0);
                        dataGridView1.Rows[num_line].Cells[0].Value = x_MouseDown;
                        dataGridView1.Rows[num_line].Cells[1].Value = y_MouseDown;
                        parentForm.dataGridView1.Rows[parentForm.num_line].Cells[0].Value = x_MouseDown + parentForm.X_chart;
                        parentForm.dataGridView1.Rows[parentForm.num_line].Cells[1].Value = y_MouseDown + parentForm.H - parentForm.magnificY - parentForm.Y_chart;
                        chart1.Series[I].Points.Clear();
                        for (int i = num_line - 1; i <= num_line; i++)
                            chart1.Series[I].Points.AddXY(dataGridView1.Rows[i].Cells[0].Value, dataGridView1.Rows[i].Cells[1].Value);

                        dataGridView1.Rows[0].Cells[3].Value = 0;
                        dataGridView1.Rows[0].Cells[4].Value = Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value) - ((Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value)) * (0 - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value))) / (Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value));
                        dataGridView1.Rows[0].Cells[5].Value = Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value) - ((Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value)) * (0 - Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value))) / (Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value));
                        dataGridView1.Rows[1].Cells[3].Value = Program.f1.magnificX;
                        dataGridView1.Rows[1].Cells[4].Value = Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value) - ((Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value)) * (Program.f1.magnificX - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value))) / (Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value));
                        dataGridView1.Rows[1].Cells[5].Value = Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value) - ((Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value)) * (Program.f1.magnificX - Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value))) / (Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value));

                        chart1.Series[0].Points.Clear();
                        chart1.Series[1].Points.Clear();
                        if (double.IsInfinity(Convert.ToDouble(dataGridView1.Rows[0].Cells[4].Value)) || double.IsInfinity(Convert.ToDouble(dataGridView1.Rows[0].Cells[5].Value)) || double.IsInfinity(Convert.ToDouble(dataGridView1.Rows[1].Cells[4].Value)) || double.IsInfinity(Convert.ToDouble(dataGridView1.Rows[1].Cells[5].Value)))
                            return;
                        else
                        {
                            for (int i = 0; i <= 1; i++)
                            {
                                chart1.Series[0].Points.AddXY(dataGridView1.Rows[i].Cells[3].Value, dataGridView1.Rows[i].Cells[4].Value);
                                chart1.Series[1].Points.AddXY(dataGridView1.Rows[i].Cells[3].Value, dataGridView1.Rows[i].Cells[5].Value);
                            }
                        }

                    }
                    catch
                    {
                        return;
                    }
                
            }
        }

        private void Form2_MouseClick(object sender, MouseEventArgs e)
        {
            this.TopMost = true;
           
        }

      

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            
            Program.f1.loupe1 = false;
            Program.f1.chart1.Series[2].Points.Clear();
            Program.f1.chart1.Series[3].Points.Clear();
            Program.f1.clickCountLoupe2x = 0;
            Program.f1.clickCountLoupe4x = 0;
            Program.f1.clickCountLoupe8x = 0;
            Program.f1.clickCountLoupe16x = 0;
            Program.f1.openf2 = false;
            Program.f1.loupe = false;
            SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle, -1, -1);

            Program.f1.button12.Visible = false;
            Program.f1.button13.Visible = false;
            Program.f1.button14.Visible = false;
            Program.f1.button15.Visible = false;

        }
        

        

        

        
    }
}
