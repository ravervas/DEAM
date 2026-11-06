using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using WindowsFormsApplication8;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Globalization;
using ExcelObj = Microsoft.Office.Interop.Excel;
using System.Drawing.Imaging;

namespace WindowsFormsApplication8
{
    public partial class Form1 : Form
    {

        private int X = Screen.PrimaryScreen.WorkingArea.Width;
        private int Y = Screen.PrimaryScreen.WorkingArea.Height;
        public int YY, XX;
        public Bitmap bmp;
        private Bitmap bmp1;
        //Graphics gr;
        Point lastpoint = Point.Empty;
        private bool MD = false;
        public bool flag1 = false;
        public bool loupe = false;
        public bool move = false;
        public bool loupe1 = false;
        public bool openf2 = false;
        public bool new_distr = false;
        public bool exit_DEAM = false;
        public bool mm = false;
        public bool mkm = false;
        public bool nm = false;
        public bool Eng = false;
        public bool Rus = true;
        public bool f3_open = false;
        public bool new_distr_f5 = false;
        public bool flag = false;
        public bool form4 = false;
        private bool formclose = false;
        public bool But6MM = false;
        public bool l4x, l8x, l16x = false;
        public bool l2x = true;
        public bool calibrateImage = false;
        public bool FormLineProphile = false;
        public bool tSBut2 = true, tSBut3 = false;
        private int W;
        private double W1;
        public int H;
        public int I_dGV2 = 0;
        public int I_dGV3 = 0;
        public int I_size = 12;
        public double ConstMicro;
        public int N_calImage = 0;
        public int x_Rect, y_Rect, x1_Rect, y1_Rect;
        public double n_s = 1;
        public double n_const_s = 1;
        public double Di = 0;
        public double coeff;
        public int kol_count_di, Y_min, Y_max;
        public int magnificY, magnificX;
        public double P, k, b;
        //private double koeff_mag_y = 1;
        public int I = 3;
        public int numDGV3 = 1;
        public int numDGV2 = 1;
        public double L, xc, yc;
        public int Rc, col_pix;
        public int num_line = 0;
        private double H1;
        public int Ly, Lx;
        private double ratio_p; // Picture
        private double ratio_s; // Screen
        public double n = 1;
        public double n_const = 1;
        private int size_H;
        private int size_W;
        //private int size_H1 = 100;
        private int size_W1 = 100;
        //private double pB1H;
        private double pB1W;
        private int clickMM = 0;
        //private int clickCount = 0;
        public int clickCountLoupe2x = 0, clickCountLoupe4x=0, clickCountLoupe8x=0, clickCountLoupe16x=0;
        private int clickCountButton5 = 0;
        public int clickCountCalibrationImage = 0;
        public int clickCountElectronogram = 0;
        public int clickLineProfile = 0;
        public int click_toolStr_button1 = 0;
        private double x_MouseDown, y_MouseDown, x_M, y_M;
        public int magnificationX;
        public int magnificationY;
        public double X_chart, Y_chart;
        private double sum_Di, Dcp;
        public double dpi;
        public double measuring_fault, delta_measuring_fault;
        //private double[] meas_fault;
        public List<double> meas_faultt = new List<double>();
        public int kol_meas_fault = -1;
        public int number_N_meas_fault = 0;
        private string FN;
        public string Units="пик";
        public string UnitsEn = "pix";
        public List<int> N_meas_faultt = new List<int>();
        public Form f2, f3, f4, f5, f6, f7, f8, f9, f10;       
        Pen myPen = new Pen(Color.Yellow, 2);
        //Pen myPenBlack = new Pen(Color.Silver, 2);
        //List<ColoredLine> ListOfLines = new List<ColoredLine>();
        public List<Bitmap> ListBmp, ListBmp2, ListBmp3;
        public List<double> dii = new List<double>();
        public int pass = 0;
        public string password;
        public string Grid = "Наложить/Убрать сетку";
        public string Legend = "Показать/скрыть легенду";
        public string font = "Размер шрифта";
        public string Rotate = "Поворот значений оси Х на";
        public string Line_thickness = "Толщина линий";
        public string Approx = "Показать/cкрыть кривые";
        public string Calibrate = "Калибровка изображения";
        public string Histogram = "Построить гистограмму";
        public string Return = "Назад";
        public string Zoom = "Увеличить изображение";
        public string LogNorm = "Перейти в Log & Norm DF 2.0";
        public string Electronogram = "Анализ электронограмм";
        public string Average = "Средний размер";
        public string Calculate_d = "Расчет межатомных расстояний";
        public string Calculate_C = "Расчет постоянной микроскопа";
        public string LineProfile = "Линейный профиль";
        ToolTip t = new ToolTip();
        public static float brightness=0, divRed=0, contrast=1, RedContrast=1, GreenContrast=1, BlueContrast=1, Red=1, Green=1, Blue=1;       


        public Form1()
        {
            Program.f1 = this;
            InitializeComponent();
            // Подгоняем размер приложения под размер экрана монитора пользователя

            ClientSize = new System.Drawing.Size(X, Y);
            size_H = this.Height;
            size_W = this.Width;
            bmp = new Bitmap(pictureBox1.Height, pictureBox1.Width);
            //ListBmp = new List<Bitmap>();
            //ListBmp2 = new List<Bitmap>();
            //meas_fault = new double[1000];
            //ListBmp.Insert(0, new Bitmap(bmp));
            //ListBmp.Add(new Bitmap(bmp));
            bmp1 = new Bitmap(pictureBox1.Height, pictureBox1.Width);



            //raspred = new double[1000];
        }

        [DllImport("kernel32.dll")]// // отнимаем у приложения лишнюю непотребляемую в данный момент память
        static extern bool SetProcessWorkingSetSize(IntPtr hProcess,
    int dwMinimumWorkingSetSize, int dwMaximumWorkingSetSize);
        // [DllImport("myDll.dll")]
        // private static extern bool SetProcessWorkingSetSize(IntPtr handle, int minimumWorkingSetSize, int maximumWorkingSetSize);
        // public static void Collect()
        // {
        // GC.Collect();
        // GC.WaitForPendingFinalizers();
        // SetProcessWorkingSetSize(System.Diagnostics.Process.GetCurrentProcess().Handle, -1, -1);
        //}

        private void файлToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }


        private void сохранитьРаспределениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            if (Eng == true)
                sfd.Title = "Save";
            sfd.OverwritePrompt = true;
            sfd.CheckPathExists = true;
            sfd.DefaultExt = "*.dat;*.txt";
            sfd.Filter = "dat files (*.dat)|*.dat| txt files (*.txt)|*.txt";
            sfd.FileName = FN;
            if (sfd.ShowDialog() == DialogResult.OK)
                try
                {
                    using (StreamWriter str = new StreamWriter(sfd.FileName))
                    {
                        for (int i = 0; i < kol_count_di; i++)
                            str.Write(Math.Round(dii[i], 2) + "\r\n");
                        str.Close();
                    }
                    if (Rus)
                        MessageBox.Show("Успешно сохранено!", "Сообщение!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (Eng)
                        MessageBox.Show("The file was successfully saved!", "Message!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch
                {
                    if (Rus)
                        MessageBox.Show("Невозможно сохранить", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (Eng)
                        MessageBox.Show("The file can't be saved", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {

            {
                OpenFileDialog open_dialog = new OpenFileDialog(); // создание диалогового окна для выбора файла                  
                open_dialog.Filter = "Image Files (.JPEG; BMP; PNG; TIF)|*.jpg; *.bmp; *.png; *tif";
                if (Eng == true)
                    open_dialog.Title = "Open";
                if (open_dialog.ShowDialog() == DialogResult.OK) // если в окне была нажата кнопка ОК
                {
                    try
                    {
                        n_const = 1;
                        n = 1;                       
                        number_N_meas_fault = 0;
                        for (int i = 0; i < I_dGV2; i++)
                            dataGridView2.Rows[i].Cells[1].Value = string.Empty;
                        I_dGV2 = 0;                                               
                        bmp = new Bitmap(open_dialog.FileName);
                        panel1.BackgroundImage = bmp;
                        //ListBmp.Insert(0, new Bitmap(bmp));
                       // ListBmp2.Insert(0, new Bitmap(bmp));
                        //bmp2 = ListBmp[0];

                        //pixel = new UInt32[bmp.Height, bmp.Width];
                       // for (int y = 0; y < bmp.Height; y++)
                         //   for (int x = 0; x < bmp.Width; x++)
                            //    pixel[y, x] = (UInt32)(bmp.GetPixel(x, y).ToArgb());

                        W = panel1.BackgroundImage.Width;
                        H = panel1.BackgroundImage.Height;
                        W1 = panel1.BackgroundImage.Width;
                        H1 = panel1.BackgroundImage.Height;
                        if (W == H)
                        {

                            panel1.Height = this.Height - 100;
                            panel1.Width = this.Height - 100;
                        }
                        if (W > H)
                        {
                            ratio_s = W1 / H1;
                            ratio_p = (this.Width - 100) / (this.Height - 100);
                            while (Math.Round(ratio_p, 2) != (Math.Round(ratio_s, 2)))
                            {
                                size_W1 = size_W1 + 1;
                                //panel1.Width = size_W - size_W1;
                                pB1W = size_W - size_W1;
                                ratio_p = pB1W / (this.Height - 100);
                            }
                            panel1.Height = this.Height - 100;
                            panel1.Width = size_W - size_W1;
                        }
                        //panel1.BackgroundImage = bmp;
                        chart1.ChartAreas[0].Position.Width = 100;
                        chart1.ChartAreas[0].Position.Height = 100;
                        dataGridView1.Rows.Add();


                        dataGridView1.Rows[0].Cells[0].Value = 0;
                        dataGridView1.Rows[0].Cells[1].Value = 0;
                        dataGridView1.Rows[1].Cells[0].Value = W;
                        dataGridView1.Rows[1].Cells[1].Value = H;


                        chart1.ChartAreas[0].AxisX.Minimum = 0;
                        chart1.ChartAreas[0].AxisY.Minimum = 0;
                        chart1.ChartAreas[0].AxisX.Maximum = W;
                        chart1.ChartAreas[0].AxisY.Maximum = H;

                        for (int i = 0; i <= 1; i++)
                        {
                            chart1.Series[0].Points.AddXY(dataGridView1.Rows[i].Cells[0].Value, dataGridView1.Rows[i].Cells[1].Value);

                        }
                       // gr = Graphics.FromImage(bmp);

                        panel1.Visible = true;
                        dpi = panel1.BackgroundImage.VerticalResolution * 0.2 / 25.4;
                        chart1.Series[0].Points.Clear();
                        chart1.Series[0].Color = Color.Black;
                        
                        //magnificationX = W / 4;
                        //magnificationY = H / 4;
                        P = Convert.ToDouble(W) / Convert.ToDouble(H);

                        Ly = H;
                        Lx = W;
                        P = Convert.ToDouble(Lx) / Convert.ToDouble(Ly);
                        P = Math.Round(P, 2);
                        //magnificY = Ly / 2;
                        //magnificX = Lx / 2;
                        
                        flag1 = true;
                        flag = true;
                        string filename = System.IO.Path.GetFileNameWithoutExtension(open_dialog.FileName);
                        FN = filename;
                        this.Text = "DEAM | " + FN;

                        for (int i = I; i > 3; i--)
                        {
                            chart1.Series[I].Points.Clear();
                            chart1.Series.Remove(chart1.Series[I]);
                            dataGridView1.Rows.Remove(dataGridView1.Rows[num_line]);
                            dataGridView1.Rows.Remove(dataGridView1.Rows[num_line - 1]);
                            num_line = num_line - 2;
                            I = I - 1;
                        }
                        kol_meas_fault = kol_meas_fault + 1;
                    }

                    catch
                    {
                        panel1.Visible = false;
                        if (Rus == true)
                            MessageBox.Show("Невозможно открыть выбранный файл", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        if (Eng == true)
                            MessageBox.Show("It is impossible to open the selected file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }

                }
                SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle, -1, -1);

            }


        }




        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();

        }









        private void Form1_Load(object sender, EventArgs e)
        {

            dataGridView1.Rows.Add();
            this.WindowState = FormWindowState.Maximized;
            SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle, -1, -1);

            N_meas_faultt.Insert(0, 0);
            meas_faultt.Insert(0, 0);

            dataGridView2.Rows[0].Cells[0].Value = 5;
            dataGridView3.Rows[0].Cells[0].Value = 5;

            for (int i = 0; i < 4; i++)
            {
                dataGridView2.Rows.Add();
                dataGridView3.Rows.Add();
            }

            
        }


        private void button1_MouseClick(object sender, MouseEventArgs e)
        {
            if (flag1 == true && clickCountElectronogram!=1) //&& flag == true)
            {
                //++clickCountLoupe;

                //switch (clickCountLoupe)
               // {
                 //   case 1:

                        if (e.Button == MouseButtons.Left)
                        {
                            //button1.BackColor = System.Drawing.SystemColors.Info;                                                       
                            loupe = true;
                            SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle, -1, -1);
                            //f8 = new Form8();
                            //f8.Owner = this;
                            //f8.Show();        
                            button12.Visible = true;                           
                            button13.Visible = true;
                            button14.Visible = true;
                            button15.Visible = true;
                            if (l2x)
                                button12_Click(sender, e);
                            if (l4x)
                                button13_Click(sender, e);
                            if (l8x)
                                button14_Click(sender, e);
                            if (l16x)
                                button15_Click(sender, e);

                        }

                    //    break;
                   // case 2:
                     //   if (e.Button == MouseButtons.Left)
                       // {
                         //   button1.BackColor = System.Drawing.SystemColors.Control;
                         //   if (loupe1 == true)
                           // {
                             //   f2.Close();                                
                               // loupe1 = false;
                          //  }
                            // bmp = ListBmp[0];
                            // panel1.BackgroundImage = bmp;
                            // ListBmp.Insert(0, new Bitmap(bmp));
                         //   chart1.Series[2].Points.Clear();
                         //   chart1.Series[3].Points.Clear();
                          //  clickCountLoupe = 0;
                          //  openf2 = false;
                          //  loupe = false;
                          //  SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle, -1, -1);
                            //f8.Close();
                          //  button12.Visible = false;
                          //  button13.Visible = false;
                          //  button14.Visible = false;
                          //  button15.Visible = false;
                       // }
                       // break;

               // }
            }
        }


        private void ммToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mm = true;
            mkm = false;
            nm = false;
        }

        private void мкмToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mm = false;
            mkm = true;
            nm = false;
        }

        private void нмToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mm = false;
            mkm = false;
            nm = true;
        }


        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < kol_count_di; i++)
            {
                sum_Di = sum_Di + dii[i];
            }
            Dcp = sum_Di / kol_count_di;            
            sum_Di = 0;
            if (kol_count_di == 0)
            {
                textBox2.Text = "–";
                textBox4.Text = "–";
            }
            else
            {
                textBox2.Text = Math.Round(Dcp, 2).ToString();
                delta_measuring_fault = 0;
                for (int i = 0; i <= kol_meas_fault; i++)
                {
                    delta_measuring_fault = delta_measuring_fault + N_meas_faultt[i] * meas_faultt[i];
                }                            
                delta_measuring_fault = delta_measuring_fault / kol_count_di;

                textBox4.Text = Math.Round(delta_measuring_fault, 2).ToString();
                textBox9.Text = Units;
                textBox10.Text = Units;
            }
        }



        private void х2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            magnificationX = Lx / 4;
            magnificationY = Ly / 4;
            magnificX = Lx / 2;
            magnificY = Ly / 2;
            //koeff_mag_y = 1;
        }

        private void х4ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            magnificationX = Lx / 8;
            magnificationY = Ly / 8;
            magnificX = Lx / 4;
            magnificY = Ly / 4;
            //koeff_mag_y = 3;
        }

        private void х8ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            magnificationX = Lx / 16;
            magnificationY = Ly / 16;
            magnificX = Lx / 8;
            magnificY = Ly / 8;
            //koeff_mag_y = 7.03;
        }

        private void х16ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            magnificationX = Lx / 32;
            magnificationY = Ly / 32;
            magnificX = Lx / 16;
            magnificY = Ly / 16;
            //koeff_mag_y = 15.06;
        }

        
        

        private void новоеРаспределениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Rus == true)
            {
                DialogResult result = MessageBox.Show("Действительно хотите создать новое рапределение?",
            "Уведомление",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question,
            MessageBoxDefaultButton.Button1);
                if (result == DialogResult.Yes)
                {
                    for (int i = I; i > 3; i--)
                    {
                        chart1.Series[I].Points.Clear();
                        chart1.Series.Remove(chart1.Series[I]);
                        dataGridView1.Rows.Remove(dataGridView1.Rows[num_line]);
                        dataGridView1.Rows.Remove(dataGridView1.Rows[num_line - 1]);
                        num_line = num_line - 2;
                        I = I - 1;
                    }
                    for (int i = 0; i < I_dGV2; i++)
                        dataGridView2.Rows[i].Cells[1].Value = string.Empty;
                    I_dGV2 = 0;                    
                    kol_count_di = 0;
                    textBox7.Text = null;
                    Di = 0;
                    kol_meas_fault = -1;
                    number_N_meas_fault = 0;
                    N_meas_faultt.Insert(0, 0);
                    meas_faultt.Insert(0, 0);
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox3.Text = null;
                    textBox4.Text = null;
                    textBox5.Text = "пик";
                    textBox6.Text = "пик";
                    textBox9.Text = String.Empty;
                    textBox10.Text = String.Empty;
                    panel1.BackgroundImage = null;
                    flag1 = false;
                    flag = false;
                    button1.BackColor = System.Drawing.SystemColors.Control;
                    if (loupe1 == true)
                    {
                        f2.Close();
                        loupe1 = false;
                    }
                    else
                        chart1.Series[3].Points.Clear();
                    clickCountLoupe2x = 0;
                    clickCountLoupe4x = 0;
                    clickCountLoupe8x = 0;
                    clickCountLoupe16x = 0;
                    if (loupe)
                    {
                        button12.Visible = false;
                        button13.Visible = false;
                        button14.Visible = false;
                        button15.Visible = false;
                    }
                    loupe = false;                   
                }
            }
            if (Eng == true)
            {
                new_distr = true;
                f5 = new Form5();
                f5.ShowDialog();
                
            }

        }

        private void englishToolStripMenuItem_Click(object sender, EventArgs e)
        {            
            Rus = false;
            Eng = true;
            Grid = "Show / hide the grid";
            Legend = "Show/hide legend";
            font = "Font size";
            Rotate = "Rotate X values on";
            Line_thickness = "Line thickness";
            Approx = "Show/hide curves of";
            Calibrate = "Calibrate Image";
            Histogram = "Plot histogram";
            Return = "Return";
            Zoom = "Zoom Image";
            LogNorm = "Go to Log & Norm DF 2.0";
            Electronogram = "Electron Diffraction Pattern Analysis";
            Average = "Average size";
            Calculate_d = "Interatomic distances calculation";
            Calculate_C = "Microscope constant calculation";
            LineProfile = "Draw a line profile";
            groupBox3.Text = Electronogram;
            tabControl1.TabPages[0].Text = "Ring diameters";
            tabControl1.TabPages[1].Text = "Calibration sample";
            button3.Text = "Calculation d";
            button4.Text = "Calculation C";
            button16.Text = "Clear lines";
            groupBox4.Text = "Image control";
            файлToolStripMenuItem.Text = "File";
            открытьToolStripMenuItem.Text = "Open";
            сохранитьКакToolStripMenuItem.Text = "Save data";
            новоеРаспределениеToolStripMenuItem.Text = "New data distribution";
            загрузитьЭкспериментальныеДанныеToolStripMenuItem.Text = "Import experimental data";
            выходToolStripMenuItem.Text = "Exit";
            инструментыToolStripMenuItem.Text = "Gadgets";
            системаЕдиницToolStripMenuItem.Text = "Unit measure";
            ммToolStripMenuItem.Text = "mm";
            мкмToolStripMenuItem.Text = "mkm";
            нмToolStripMenuItem.Text = "nm";
            коэффициентЛупыToolStripMenuItem.Text = "Loupe magnification";
            информацияToolStripMenuItem.Text = "Information";
            button5.Text = "Change image position";
            языкToolStripMenuItem.Text = "Language";
            русскийToolStripMenuItem.Text = "Russian";
            englishToolStripMenuItem.Text = "English";         
            оПрограммеИИнструкцииПоРаботеToolStripMenuItem.Text = "About DEAM and work instruction";
            groupBox1.Text = "Instruments";
            groupBox2.Text = "Measurement results";
            label1.Text = "d, nm";
            label4.Text = "d, nm";
            if (textBox5.Text == "пик")
            {
                textBox5.Text = "pix";
                textBox6.Text = "pix";
                label2.Text = "D, pix";
                label5.Text = "D, pix";
            }
            if (textBox5.Text == "нм")
            {
                textBox5.Text = "nm";
                textBox6.Text = "nm";
                label2.Text = "D, nm";
                label5.Text = "D, nm";
            }
            if (textBox5.Text == "мкм")
            {
                textBox5.Text = "mkm";
                textBox6.Text = "mkm";
                label2.Text = "D, mkm";
                label5.Text = "D, mkm";
            }
            if (textBox5.Text == "мм")
            {
                textBox5.Text = "mm";
                textBox6.Text = "mm";
                label2.Text = "D, mm";
                label5.Text = "D, mm";
            }
            if (textBox9.Text!=String.Empty)
            {
                if (textBox9.Text == "пик")
                {
                    textBox9.Text = "pix";
                    textBox10.Text = "pix";
                }
                if (textBox9.Text == "нм")
                {
                    textBox9.Text = "nm";
                    textBox10.Text = "nm";
                }
                if (textBox9.Text == "мкм")
                {
                    textBox9.Text = "mkm";
                    textBox10.Text = "mkm";
                }
                if (textBox9.Text == "мм")
                {
                    textBox9.Text = "mm";
                    textBox10.Text = "mm";
                }
            }

            if (FormLineProphile)
            {
                Program.f10.Text = "Line profile";
                Program.f10.chart1.ChartAreas[0].AxisX.Title = Program.f1.textBox5.Text;
                Program.f10.LineChart = "Line chart";
                Program.f10.ColumnChart = "Column chart";
                Program.f10.Interval = "Interval measurement";
                Program.f10.toolStripMenuItem2.Text = "Discrete";
                Program.f10.toolStripMenuItem3.Text = "Continuous";
                Program.f10.Smooth = "Smoothing";
            }

            if (loupe)
                Program.f8.Text = "Zoom factor";

            if (f3_open == true)
            {
                Program.f3.файлToolStripMenuItem.Text = "File";
                Program.f3.открытьФайлtxtToolStripMenuItem.Text = "Load file at the table";               
                Program.f3.новоеРаспределениеToolStripMenuItem.Text = "New data distribution";
                Program.f3.сохранитьИзображениеToolStripMenuItem.Text = "Save image";
                Program.f3.выходToolStripMenuItem.Text = "Exit";
                Program.f3.свойстваТаблицыToolStripMenuItem.Text = "Diagram";
                Program.f3.описатьToolStripMenuItem.Text = "Fit";
                Program.f3.логнормФияToolStripMenuItem.Text = "Log-normal function";
                Program.f3.нормФцияГауссToolStripMenuItem.Text = "Normal function (Gauss)";
                Program.f3.бимодальноеToolStripMenuItem.Text = "Bimodal";
                Program.f3.логнормФцияToolStripMenuItem2.Text = "Log-normal function";
                Program.f3.нормФцияГауссToolStripMenuItem2.Text = "Normal function (Gauss)";               
                Program.f3.единицаИзмеренияToolStripMenuItem.Text = "Unit measure";
                Program.f3.toolStripComboBox2.Items.Clear();
                Program.f3.toolStripComboBox2.Text = "nm";
                Program.f3.toolStripComboBox2.Items.AddRange(new object[] { "mm", "mkm", "nm" });
                Program.f3.toolStripComboBox2.SelectedIndex = Program.f3.tSCB2;                                                            
                Program.f3.кубическоеРаспределениеToolStripMenuItem.Text = "Cubic distribution";
                Program.f3.линейноеРаспределениеToolStripMenuItem.Text = "Linear distribution";
                Program.f3.максЗначениеОсиXToolStripMenuItem.Text = "Max. value of the X-axis";
                Program.f3.groupBox3.Text = "Max. value of the X-axis";               
                Program.f3.распределениеОКРToolStripMenuItem.Text = "CSR distribution";
                Program.f3.распределение1ToolStripMenuItem.Text = "Load the 1-st distribution";
                Program.f3.распределение2ToolStripMenuItem.Text = "Load the 2-nd distribution";
                Program.f3.отобразитьToolStripMenuItem.Text = "Plot the graph";
                Program.f3.label12.Text = "nm";
                Program.f3.label13.Text = "nm";
                Program.f3.label14.Text = "nm";
                Program.f3.label15.Text = "nm";
                Program.f3.pictureBox7.Image = Properties.Resources.DCSR1;
                Program.f3.pictureBox8.Image = Properties.Resources.D_v_CSR1;
                Program.f3.pictureBox9.Image = Properties.Resources.DCSR2;
                Program.f3.pictureBox10.Image = Properties.Resources.D_v_CSR2;
                Program.f3.groupBox2.Text = "1-st CSR distribution area";
                Program.f3.chart2.Series[0].Name = "Hist.";
                if (Program.f3.chart2.Series[1].Name == "Лог-норм 1")
                    Program.f3.chart2.Series[1].Name = "Log-Norm 1";
                if (Program.f3.chart2.Series[1].Name == "Гаусс 1")
                    Program.f3.chart2.Series[1].Name = "Gauss 1";
                if (Program.f3.chart2.Series[2].Name == "Лог-норм 2")
                    Program.f3.chart2.Series[2].Name = "Log-Norm 2";
                if (Program.f3.chart2.Series[2].Name == "Гаусс 2")
                    Program.f3.chart2.Series[2].Name = "Gauss 2";
                Program.f3.chart2.Series[3].Name = "CSR 1";
                Program.f3.chart2.Series[4].Name = "CSR 2";
                Program.f3.chart2.Series[5].Name = "Log-Norm common";
                Program.f3.языкToolStripMenuItem.Text = "Language";
                Program.f3.русскийToolStripMenuItem.Text = "Russian";
                Program.f3.английсикйToolStripMenuItem.Text = "English";
                Program.f3.ифнформацияToolStripMenuItem.Text = "Information";
                Program.f3.chart2.ChartAreas[0].AxisY.Title = "Frequency, %";
                Program.f3.dataGridView1.Columns[0].HeaderCell.Value = "Particle size, nm";
                Program.f3.dataGridView1.Columns[1].HeaderCell.Value = "Particles number";
                Program.f3.dataGridView1.Columns[2].HeaderCell.Value = "Particles fraction, %";
                Program.f3.button2.Text = "Plot a histogram";


                if (Program.f3.toolStripComboBox2.SelectedIndex == 0)
                {
                    Program.f3.chart2.ChartAreas[0].AxisX.Title = "Particle size, mm";
                    Program.f3.dataGridView1.Columns[0].HeaderCell.Value = "Particle size, mm";
                }
                if (Program.f3.toolStripComboBox2.SelectedIndex == 1)
                {
                    Program.f3.chart2.ChartAreas[0].AxisX.Title = "Particle size, mkm";
                    Program.f3.dataGridView1.Columns[0].HeaderCell.Value = "Particle size, mkm";
                }
                if (Program.f3.toolStripComboBox2.SelectedIndex == 2)
                {
                    Program.f3.chart2.ChartAreas[0].AxisX.Title = "Particle size, nm";
                    Program.f3.dataGridView1.Columns[0].HeaderCell.Value = "Particle size, nm";
                }
                Program.f3.оПрограммеИИнструкцииПоРаботеToolStripMenuItem.Text = "About DEAM and work instruction";
                Program.f3.форматОсиХToolStripMenuItem.Text = "X-axis format";
                Program.f3.диапазонРазмеровToolStripMenuItem.Text = "Sizes range";
                Program.f3.среднийРазмерДиапазонаToolStripMenuItem.Text = "Average sizes of the range";
                Program.f3.аппроксимацииToolStripMenuItem.Text = "Approximation";
                Program.f3.оКРToolStripMenuItem.Text = "CSR";
                Program.f3.распределение1ToolStripMenuItem3.Text = "Distribution 1";
                Program.f3.распределение2ToolStripMenuItem3.Text = "Distribution 2";
                Program.f3.общаябимодалToolStripMenuItem1.Text = "Common (bimodal)";
                Program.f3.распределение1ToolStripMenuItem4.Text = "Distribution 1";
                Program.f3.распределение2ToolStripMenuItem4.Text = "Distribution 2";
            }

            if (clickCountCalibrationImage==1)
            {
                f9.Text = Calibrate;
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

            if (form4)
            {
                f4.Text = Histogram;
                Program.f4.label1.Text = "Enter the number of columns to plot the histogram";
                Program.f4.label2.Text = "Minimum size:";
                Program.f4.label3.Text = "Maximum size:";
                Program.f4.label4.Text = "Columns number:";
                Program.f4.button2.Text = "Cancel";
            }

        }

        private void русскийToolStripMenuItem_Click(object sender, EventArgs e)
        {          
             Rus = true;
            Eng = false;
            Grid = "Наложить/Убрать сетку";
            Legend = "Показать/скрыть легенду";
            font = "Размер шрифта";
            Rotate = "Поворот значений оси Х на";
            Line_thickness = "Толщина линий";
            Approx = "Показать/скрыть кривые";
            Calibrate = "Калибровка изображения";
            Histogram = "Построить гистограмму";
            Return = "Назад";
            Zoom = "Увеличить изображение";
            LogNorm = "Перейти в Log & Norm DF 2.0";
            Electronogram = "Анализ электронограмм";
            Average = "Средний размер";
            Calculate_d = "Расчет межатомных расстояний";
            Calculate_C = "Расчет постоянной микроскопа";
            button16.Text = "Очистить линии";
            LineProfile = "Линейный профиль";
            groupBox3.Text = Electronogram;
            tabControl1.TabPages[0].Text = "Диаметры колец";
            tabControl1.TabPages[1].Text = "Эталон";           
            button3.Text = "Рассчитать d";
            button4.Text = "Рассчитать C";
            button16.Text = "Очистить линии";
            groupBox4.Text = "Параметры изображения";
            файлToolStripMenuItem.Text = "Файл";
            открытьToolStripMenuItem.Text = "Открыть";
            сохранитьКакToolStripMenuItem.Text = "Сохранить распределение";
            новоеРаспределениеToolStripMenuItem.Text = "Новое распределение";
            загрузитьЭкспериментальныеДанныеToolStripMenuItem.Text = "Загрузить экспериментальные данные";
            выходToolStripMenuItem.Text = "Выход";
             инструментыToolStripMenuItem.Text = "Инструменты";
            системаЕдиницToolStripMenuItem.Text = "Единица измерения";
            ммToolStripMenuItem.Text = "мм";
            мкмToolStripMenuItem.Text = "мкм";
            нмToolStripMenuItem.Text = "нм";
            коэффициентЛупыToolStripMenuItem.Text = "Коэффициент лупы";
             информацияToolStripMenuItem.Text = "Информация";
            button5.Text = "Изменить положение изображения";
            языкToolStripMenuItem.Text = "Язык";        
            оПрограммеИИнструкцииПоРаботеToolStripMenuItem.Text = "О программе и инструкции по работе";
            groupBox1.Text = "Инструменты";
            groupBox2.Text = "Результаты измерений";
            label1.Text = "d, нм";
            label4.Text = "d, нм";
            if (textBox5.Text == "pix")
            {
                textBox5.Text = "пик";
                textBox6.Text = "пик";
                label2.Text = "D, пик";
                label5.Text = "D, пик";
            }
            if (textBox5.Text == "nm")
            {
                textBox5.Text = "нм";
                textBox6.Text = "нм";
                label2.Text = "D, нм";
                label5.Text = "D, нм";
            }
            if (textBox5.Text == "mkm")
            {
                textBox5.Text = "мкм";
                textBox6.Text = "мкм";
                label2.Text = "D, мкм";
                label5.Text = "D, мкм";
            }
            if (textBox5.Text == "mm")
            {
                textBox5.Text = "мм";
                textBox6.Text = "мм";
                label2.Text = "D, мм";
                label5.Text = "D, мм";
            }

            if (textBox9.Text != String.Empty)
            {
                if (textBox9.Text == "pix")
                {
                    textBox9.Text = "пик";
                    textBox10.Text = "пик";
                }
                if (textBox9.Text == "nm")
                {
                    textBox9.Text = "нм";
                    textBox10.Text = "нм";
                }
                if (textBox9.Text == "mkm")
                {
                    textBox9.Text = "мкм";
                    textBox10.Text = "мкм";
                }
                if (textBox9.Text == "mm")
                {
                    textBox9.Text = "мм";
                    textBox10.Text = "мм";
                }
            }

            if (FormLineProphile)
            {
                Program.f10.Text = "Линейный профиль";
                Program.f10.chart1.ChartAreas[0].AxisX.Title = Program.f1.textBox5.Text;
                Program.f10.LineChart = "Линейная диаграмма";
                Program.f10.ColumnChart = "Столбчатая диаграмма";
                Program.f10.Interval = "Измерение интервала";
                Program.f10.toolStripMenuItem2.Text = "Дискретное";
                Program.f10.toolStripMenuItem3.Text = "Непрерывное";
                Program.f10.Smooth = "Сглаживание";
            }

            if (loupe)            
              Program.f8.Text = "Коэффициент лупы";

            if (f3_open==true)
            {
                Program.f3.файлToolStripMenuItem.Text = "Файл";
                Program.f3.открытьФайлtxtToolStripMenuItem.Text = "Загрузить файл в таблицу";              
                Program.f3.новоеРаспределениеToolStripMenuItem.Text = "Новое распределение";
                Program.f3.сохранитьИзображениеToolStripMenuItem.Text = "Сохранить изображение";
                Program.f3.выходToolStripMenuItem.Text = "Выход";
                Program.f3.свойстваТаблицыToolStripMenuItem.Text = "График";
                Program.f3.описатьToolStripMenuItem.Text = "Описать";
                Program.f3.логнормФияToolStripMenuItem.Text = "Лог-норм ф-ция";
                Program.f3.нормФцияГауссToolStripMenuItem.Text = "Норм ф-ция (Гаусс)";
                Program.f3.бимодальноеToolStripMenuItem.Text = "Бимодальное";
                Program.f3.логнормФцияToolStripMenuItem2.Text = "Лог-норм ф-ция";
                Program.f3.нормФцияГауссToolStripMenuItem2.Text = "Норм ф-ция (Гаусс)";               
                Program.f3.единицаИзмеренияToolStripMenuItem.Text = "Единица измерения";
                Program.f3.toolStripComboBox2.Items.Clear();
                Program.f3.toolStripComboBox2.Text = "нм";
                Program.f3.toolStripComboBox2.Items.AddRange(new object[] { "мм", "мкм", "нм" });
                Program.f3.toolStripComboBox2.SelectedIndex = Program.f3.tSCB2;                                                             
                Program.f3.кубическоеРаспределениеToolStripMenuItem.Text = "Кубическое распределение";
                Program.f3.линейноеРаспределениеToolStripMenuItem.Text = "Линейное распределение";
                Program.f3.максЗначениеОсиXToolStripMenuItem.Text = "Макс. значение оси Х";
                Program.f3.groupBox3.Text = "Макс. значение оси Х";               
                Program.f3.распределениеОКРToolStripMenuItem.Text = "Распределение ОКР";
                Program.f3.распределение1ToolStripMenuItem.Text = "Загрузить распределение 1";
                Program.f3.распределение2ToolStripMenuItem.Text = "Загрузить распределение 2";
                Program.f3.отобразитьToolStripMenuItem.Text = "Отобразить";
                Program.f3.label12.Text = "нм";
                Program.f3.label13.Text = "нм";
                Program.f3.label14.Text = "нм";
                Program.f3.label15.Text = "нм";
                Program.f3.pictureBox7.Image = Properties.Resources.Dокр1;
                Program.f3.pictureBox8.Image = Properties.Resources.D_v_окр1;
                Program.f3.pictureBox9.Image = Properties.Resources.Dокр2;
                Program.f3.pictureBox10.Image = Properties.Resources.D_v_окр2;
                Program.f3.groupBox2.Text = "Площадь распределения ОКР 1";
                Program.f3.chart2.Series[0].Name = "Гист.";
                if (Program.f3.chart2.Series[1].Name == "Log-Norm 1")
                    Program.f3.chart2.Series[1].Name = "Лог-норм 1";
                if (Program.f3.chart2.Series[1].Name == "Gauss 1")
                    Program.f3.chart2.Series[1].Name = "Гаусс 1";
                if (Program.f3.chart2.Series[2].Name == "Log-Norm 2")
                    Program.f3.chart2.Series[2].Name = "Лог-норм 2";
                if (Program.f3.chart2.Series[2].Name == "Gauss 2")
                    Program.f3.chart2.Series[2].Name = "Гаусс 2";
                Program.f3.chart2.Series[3].Name = "ОКР 1";
                Program.f3.chart2.Series[4].Name = "ОКР 2";
                Program.f3.chart2.Series[5].Name = "Общая Лог-норм";
                Program.f3.языкToolStripMenuItem.Text = "Язык";
                Program.f3.русскийToolStripMenuItem.Text = "Русский";
                Program.f3.английсикйToolStripMenuItem.Text = "Английский";
                Program.f3.ифнформацияToolStripMenuItem.Text = "Информация";
                Program.f3.chart2.ChartAreas[0].AxisY.Title = "Доля частиц, %";
                Program.f3.dataGridView1.Columns[0].HeaderCell.Value = "Размер частиц, нм";
                Program.f3.dataGridView1.Columns[1].HeaderCell.Value = "Число частиц";
                Program.f3.dataGridView1.Columns[2].HeaderCell.Value = "Доля частиц, %";
                Program.f3.button2.Text = "Построить гистограмму";
                if (Program.f3.toolStripComboBox2.SelectedIndex == 0)
                {
                    Program.f3.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, мм";
                    Program.f3.dataGridView1.Columns[0].HeaderCell.Value = "Размер частиц, мм";
                }
                if (Program.f3.toolStripComboBox2.SelectedIndex == 1)
                {
                    Program.f3.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, мкм";
                    Program.f3.dataGridView1.Columns[0].HeaderCell.Value = "Размер частиц, мкм";
                }
                if (Program.f3.toolStripComboBox2.SelectedIndex == 2)
                {
                    Program.f3.chart2.ChartAreas[0].AxisX.Title = "Размер частиц, нм";
                    Program.f3.dataGridView1.Columns[0].HeaderCell.Value = "Размер частиц, нм";
                }
                Program.f3.оПрограммеИИнструкцииПоРаботеToolStripMenuItem.Text = "О программе и инструкции по работе";
                Program.f3.форматОсиХToolStripMenuItem.Text = "Формат оси Х";
                Program.f3.диапазонРазмеровToolStripMenuItem.Text = "Диапазон размеров";
                Program.f3.среднийРазмерДиапазонаToolStripMenuItem.Text = "Средние размеры диапазона";
                Program.f3.аппроксимацииToolStripMenuItem.Text = "Аппроксимации";
                Program.f3.оКРToolStripMenuItem.Text = "ОКР";
                Program.f3.распределение1ToolStripMenuItem3.Text = "Распределение 1";
                Program.f3.распределение2ToolStripMenuItem3.Text = "Распределение 2";
                Program.f3.общаябимодалToolStripMenuItem1.Text = "Общая (бимодал)";
                Program.f3.распределение1ToolStripMenuItem4.Text = "Распределение 1";
                Program.f3.распределение2ToolStripMenuItem4.Text = "Распределение 2";
            }

            if (clickCountCalibrationImage == 1)
            {
                f9.Text = Calibrate;
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

            if (form4)
            {
                f4.Text = "Построение гистограммы";
                Program.f4.label1.Text = "Укажите количество столбцов для построения гистограммы";
                Program.f4.label2.Text = "Мин. размер:";
                Program.f4.label3.Text = "Макс. размер:";
                Program.f4.label4.Text = "Кол-во столбцов:";
                Program.f4.button2.Text = "Отмена";
            }

        }
       

        private void button5_MouseClick(object sender, MouseEventArgs e)
        {
            ++clickCountButton5;
            switch (clickCountButton5)
            {
                case 1:
                    if (e.Button == MouseButtons.Left)
                    {
                        chart1.Visible = false;
                        button5.BackColor = Color.Red;
                        ControlMoverOrResizer.Init(panel1);
                        ControlMoverOrResizer.WorkType = ControlMoverOrResizer.MoveOrResize.Move;
                        move = true;
                    }
                    break;
                case 2:
                    if (e.Button == MouseButtons.Left)
                    {
                        button5.BackColor = System.Drawing.SystemColors.Control;
                        ControlMoverOrResizer.WorkType = ControlMoverOrResizer.MoveOrResize.Not;
                        move = false;
                        clickCountButton5 = 0;
                        chart1.Visible = true;
                    }
                    break;
                    

            }
            
        }

        
            

        private void chart1_MouseEnter(object sender, EventArgs e)
        {
            if (flag1 == true)//flag == true &&
                this.Cursor = Cursors.Cross;
            if (flag1 == true && loupe == true)//flag == true &&
            {
                Cursor cur = new Cursor(new System.IO.MemoryStream(global::WindowsFormsApplication8.Properties.Resources.Cur));
                this.Cursor = cur;
            }
        }

        private void chart1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ++clickMM;
                switch (clickMM)
                {
                    case 1:
                        if (flag1 == true && loupe == true)
                        {
                            try
                            {



                                // bmp = ListBmp[0];
                                  //panel1.BackgroundImage = bmp;
                                // ListBmp.Insert(0, new Bitmap(bmp));
                                chart1.Series[2].Points.Clear();
                                
                                loupe1 = true;
                                //gr = Graphics.FromImage(bmp);
                                lastpoint = e.Location;

                                //x_MouseDown = e.X * W / panel1.Width;
                                x_MouseDown= Math.Round(chart1.ChartAreas[0].AxisX.PixelPositionToValue(lastpoint.X), 0);
                                //y_MouseDown = e.Y * H / panel1.Height;
                                y_MouseDown= Math.Round(chart1.ChartAreas[0].AxisY.PixelPositionToValue(lastpoint.Y), 0);
                                //x_M = x_MouseDown;
                                //y_M = y_MouseDown;
                                x_Rect = Convert.ToInt32(x_MouseDown) - magnificationX;
                                y_Rect = Convert.ToInt32(y_MouseDown) - magnificationY;
                                
                                X_chart = (Math.Round(chart1.ChartAreas[0].AxisX.PixelPositionToValue(lastpoint.X), 0) - magnificationX) - (Math.Round(chart1.ChartAreas[0].AxisX.PixelPositionToValue(lastpoint.X), 0) - x_MouseDown);
                                Y_chart = (Math.Round(chart1.ChartAreas[0].AxisY.PixelPositionToValue(lastpoint.Y), 0) - magnificationY) - (Math.Round(chart1.ChartAreas[0].AxisY.PixelPositionToValue(lastpoint.Y), 0) - (H - y_MouseDown));

                                if (x_Rect < 0)
                                {
                                    x_Rect = 0;
                                    X_chart = 0;
                                }
                                if (y_Rect < 0)
                                {
                                    y_Rect = 0;
                                    Y_chart = H - magnificY;
                                }
                                if (x_Rect > W - magnificX)
                                {
                                    x_Rect = W - magnificX;
                                    X_chart = x_Rect;
                                }
                                if (y_Rect > H - magnificY)
                                {
                                    y_Rect = H - magnificY;
                                    Y_chart = 0;

                                }
                                
                                chart1.Series[2].Points.AddXY(x_Rect, y_Rect);
                                chart1.Series[2].Points.AddXY(x_Rect, y_Rect+ magnificY);
                                chart1.Series[2].Points.AddXY(x_Rect + magnificX, y_Rect + magnificY);
                                chart1.Series[2].Points.AddXY(x_Rect + magnificX, y_Rect);
                                chart1.Series[2].Points.AddXY(x_Rect, y_Rect);
                                //y_Rect = y_Rect;
                                //gr.DrawRectangle(Pens.Blue, x_Rect, y_Rect, magnificX, magnificY);
                                Rectangle rect = new Rectangle(x_Rect, H-y_Rect-magnificY, magnificX, magnificY);
                                var pic = (Bitmap)panel1.BackgroundImage;
                                bmp1 = pic.Clone(rect, System.Drawing.Imaging.PixelFormat.Format48bppRgb);

                                //ListBmp.Insert(1, new Bitmap(bmp));


                                using (MemoryStream ms = new MemoryStream())
                                {
                                    bmp1.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                    CommunicateClass.imageData = ms.ToArray();
                                }
                                if (openf2 == false)
                                {
                                    f2 = new Form2(this);
                                    f2.Owner = this;
                                    f2.Show();
                                }
                                else
                                    Program.f2.openImage();

                                clickMM = 0;
                                openf2 = true;
                               // bmp = ListBmp[1];
                               // panel1.BackgroundImage = bmp;
                               // ListBmp.Insert(2, new Bitmap(bmp));
                            }
                            catch
                            {
                                return;
                            }

                        }
                        if (flag1 == true && loupe == false && clickCountElectronogram != 1 && clickLineProfile!=1)
                        {
                            try
                            {
                                if (calibrateImage)
                                    if (N_calImage == 1)
                                    {
                                        if (Rus == true)
                                            MessageBox.Show("Сначала удалите предыдущее измерение", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        if (Eng == true)
                                            MessageBox.Show("Delete the previous measurement first", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        clickMM = 0;
                                        return;
                                    }
                                dataGridView1.Rows.Add();
                                MD = true;
                                lastpoint = e.Location;
                                x_MouseDown = Math.Round(chart1.ChartAreas[0].AxisX.PixelPositionToValue(lastpoint.X), 0);
                                y_MouseDown = Math.Round(chart1.ChartAreas[0].AxisY.PixelPositionToValue(lastpoint.Y), 0);
                                dataGridView1.Rows[num_line].Cells[0].Value = x_MouseDown;
                                dataGridView1.Rows[num_line].Cells[1].Value = y_MouseDown;

                                num_line = num_line + 1;
                                I = I + 1;
                                chart1.Series.Add(Convert.ToString(I));
                                chart1.Series[I].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                                chart1.Series[I].Color = Color.Red;
                            }
                            catch
                            {
                                return;
                            }
                        }


                        if (flag1 == true && loupe == false && clickCountElectronogram == 1 && clickLineProfile!=1)
                        {
                            try
                            {
                                dataGridView1.Rows.Add();
                                MD = true;
                                lastpoint = e.Location;
                                x_MouseDown = Math.Round(chart1.ChartAreas[0].AxisX.PixelPositionToValue(lastpoint.X), 0);
                                y_MouseDown = Math.Round(chart1.ChartAreas[0].AxisY.PixelPositionToValue(lastpoint.Y), 0);
                                dataGridView1.Rows[num_line].Cells[0].Value = x_MouseDown;
                                dataGridView1.Rows[num_line].Cells[1].Value = y_MouseDown;
                                num_line = num_line + 1;
                                I = I + 1;
                                chart1.Series.Add(Convert.ToString(I));
                                chart1.Series[I].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Point;
                                chart1.Series[I].MarkerSize = 2;
                                chart1.Series[I].MarkerStyle = System.Windows.Forms.DataVisualization.Charting.MarkerStyle.Circle;
                                chart1.Series[I].Color = Color.Lime;
                            }
                            catch
                            {
                                return;
                            }
                        }

                        if (flag1 == true && loupe == false && clickCountElectronogram != 1 && clickLineProfile == 1)
                            try
                            {
                                if (FormLineProphile)
                                {
                                    chart1.Series[I].Points.Clear();
                                    chart1.Series.Remove(chart1.Series[I]);
                                    dataGridView1.Rows.Remove(dataGridView1.Rows[num_line]);
                                    dataGridView1.Rows.Remove(dataGridView1.Rows[num_line - 1]);
                                    num_line = num_line - 2;
                                    kol_count_di = kol_count_di - 1;
                                    I = I - 1;
                                }
                                dataGridView1.Rows.Add();
                                MD = true;
                                lastpoint = e.Location;
                                x_MouseDown = Math.Round(chart1.ChartAreas[0].AxisX.PixelPositionToValue(lastpoint.X), 0);
                                y_MouseDown = Math.Round(chart1.ChartAreas[0].AxisY.PixelPositionToValue(lastpoint.Y), 0);
                                dataGridView1.Rows[num_line].Cells[0].Value = x_MouseDown;
                                dataGridView1.Rows[num_line].Cells[1].Value = y_MouseDown;
                                num_line = num_line + 1;
                                I = I + 1;
                                chart1.Series.Add(Convert.ToString(I));
                                
                                chart1.Series[I].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                                chart1.Series[I].MarkerSize = 3;
                                chart1.Series[I].MarkerStyle = System.Windows.Forms.DataVisualization.Charting.MarkerStyle.Square;
                                chart1.Series[I].MarkerColor = Color.White;
                                chart1.Series[I].MarkerBorderColor = Color.Aqua;
                                chart1.Series[I].Color = Color.Aqua;
                            }

                            catch
                            {
                                return;
                            }

                            //I = I + 1;

                            break;
            
                    case 2:
                        if (flag1 == true && loupe == false && clickCountElectronogram != 1)
                        {
                            try
                            {
                                MD = false;
                                clickMM = 0;
                                dataGridView1.Rows.Add();                               
                                chart1.Series[0].Points.Clear();
                                chart1.Series[1].Points.Clear();                            
                                Di = Math.Round(Math.Sqrt((Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value)) * (Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value)) + (Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value)) * (Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value))), 2);
                                Di = Di * n / n_const;                             
                                dii.Insert(kol_count_di, Di);
                                kol_count_di = kol_count_di + 1;
                                if (calibrateImage)
                                {
                                    Program.f9.textBox1.Text = Math.Round(Di, 2).ToString();
                                    N_calImage = N_calImage + 1;
                                    number_N_meas_fault = number_N_meas_fault + 1;
                                    N_meas_faultt.Insert(kol_meas_fault, number_N_meas_fault);
                                }
                                else if (clickLineProfile != 1)
                                {
                                    textBox1.Text = Math.Round(Di, 2).ToString();
                                    textBox7.Text = kol_count_di.ToString();
                                    number_N_meas_fault = number_N_meas_fault + 1;
                                    N_meas_faultt.Insert(kol_meas_fault, number_N_meas_fault);
                                }
                                num_line = num_line + 1;
                                
                                if (clickLineProfile==1)
                                {
                                    if (FormLineProphile)
                                    {
                                        dataGridView4.Rows.Clear();
                                    }
                                    //panel1.BackgroundImage = bmp;
                                    k = (Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 2].Cells[1].Value)) / (Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 2].Cells[0].Value));
                                    b = Convert.ToDouble(dataGridView1.Rows[num_line - 2].Cells[1].Value) - k * Convert.ToDouble(dataGridView1.Rows[num_line - 2].Cells[0].Value);                                    
                                    if (Math.Abs(Convert.ToInt32(dataGridView1.Rows[num_line-2].Cells[1].Value)- Convert.ToInt32(dataGridView1.Rows[num_line - 1].Cells[1].Value))> Math.Abs(Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[0].Value) - Convert.ToInt32(dataGridView1.Rows[num_line - 1].Cells[0].Value)))
                                        
                                        {
                                            col_pix = Math.Abs(Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[1].Value) - Convert.ToInt32(dataGridView1.Rows[num_line - 1].Cells[1].Value));
                                            coeff = Di / col_pix;
                                            Color pixel;                                            
                                            pixel = (panel1.BackgroundImage as Bitmap).GetPixel(Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[0].Value), H-Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[1].Value));
                                            dataGridView4.Rows.Add();
                                            dataGridView4.Rows[0].Cells[2].Value = (pixel.R+pixel.G+pixel.B)/3;
                                            dataGridView4.Rows[0].Cells[0].Value = 0;
                                            for (int i=1;i<col_pix;i++)
                                            {
                                               if (Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[1].Value) > Convert.ToInt32(dataGridView1.Rows[num_line - 1].Cells[1].Value))                                                
                                                   YY = Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[1].Value) - i;  
                                                else
                                                    YY = Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[1].Value) + i;
                                                
                                                XX = Convert.ToInt32((YY - b) / k);
                                                pixel = (panel1.BackgroundImage as Bitmap).GetPixel(XX, H-YY);
                                                dataGridView4.Rows.Add();
                                                dataGridView4.Rows[i].Cells[2].Value= (pixel.R + pixel.G + pixel.B) / 3;
                                                dataGridView4.Rows[i].Cells[0].Value = i*coeff;
                                            
                                            }
                                            pixel = (panel1.BackgroundImage as Bitmap).GetPixel(Convert.ToInt32(dataGridView1.Rows[num_line - 1].Cells[0].Value), H - Convert.ToInt32(dataGridView1.Rows[num_line - 1].Cells[1].Value));
                                            dataGridView4.Rows[col_pix].Cells[2].Value = (pixel.R + pixel.G + pixel.B) / 3;
                                            dataGridView4.Rows[col_pix].Cells[0].Value = col_pix * coeff;
                                        Y_min = Convert.ToInt32(dataGridView4.Rows[0].Cells[2].Value);
                                        Y_max= Convert.ToInt32(dataGridView4.Rows[0].Cells[2].Value);

                                        for (int i = 1; i <= col_pix; i++)
                                        {
                                            if (Convert.ToInt32(dataGridView4.Rows[i].Cells[2].Value) < Y_min)
                                                Y_min = Convert.ToInt32(dataGridView4.Rows[i].Cells[2].Value);
                                            if (Convert.ToInt32(dataGridView4.Rows[i].Cells[2].Value) > Y_max)
                                                Y_max = Convert.ToInt32(dataGridView4.Rows[i].Cells[2].Value);
                                        }

                                        }
                                    else
                                    {
                                        col_pix = Math.Abs(Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[0].Value) - Convert.ToInt32(dataGridView1.Rows[num_line - 1].Cells[0].Value));
                                        coeff = Di / col_pix;
                                        Color pixel;
                                        pixel = (panel1.BackgroundImage as Bitmap).GetPixel(Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[0].Value), H - Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[1].Value));
                                        dataGridView4.Rows.Add();
                                        dataGridView4.Rows[0].Cells[2].Value = (pixel.R + pixel.G + pixel.B) / 3;
                                        dataGridView4.Rows[0].Cells[0].Value = 0;
                                        for (int i = 1; i < col_pix; i++)
                                        {
                                            if (Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[0].Value) > Convert.ToInt32(dataGridView1.Rows[num_line - 1].Cells[0].Value))
                                                XX = Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[0].Value) - i;
                                            else
                                                XX = Convert.ToInt32(dataGridView1.Rows[num_line - 2].Cells[0].Value) + i;

                                            YY = Convert.ToInt32(k*XX+b);
                                            pixel = (panel1.BackgroundImage as Bitmap).GetPixel(XX, H - YY);
                                            dataGridView4.Rows.Add();
                                            dataGridView4.Rows[i].Cells[2].Value = (pixel.R + pixel.G + pixel.B) / 3;
                                            dataGridView4.Rows[i].Cells[0].Value = i * coeff;
                                        }
                                        pixel = (panel1.BackgroundImage as Bitmap).GetPixel(Convert.ToInt32(dataGridView1.Rows[num_line - 1].Cells[0].Value), H - Convert.ToInt32(dataGridView1.Rows[num_line - 1].Cells[1].Value));
                                        dataGridView4.Rows[col_pix].Cells[2].Value = (pixel.R + pixel.G + pixel.B) / 3;
                                        dataGridView4.Rows[col_pix].Cells[0].Value = col_pix * coeff;
                                        Y_min = Convert.ToInt32(dataGridView4.Rows[0].Cells[2].Value);
                                        Y_max = Convert.ToInt32(dataGridView4.Rows[0].Cells[2].Value);
                                        for (int i = 1; i <= col_pix; i++)
                                        {
                                            if (Convert.ToInt32(dataGridView4.Rows[i].Cells[2].Value) < Y_min)
                                                Y_min = Convert.ToInt32(dataGridView4.Rows[i].Cells[2].Value);
                                            if (Convert.ToInt32(dataGridView4.Rows[i].Cells[2].Value) > Y_max)
                                                Y_max = Convert.ToInt32(dataGridView4.Rows[i].Cells[2].Value);
                                        }

                                    }
                                    
                                    if (FormLineProphile)
                                        Program.f10.LineProphile();
                                    else
                                    {
                                        f10 = new Form10();
                                        f10.Owner = this;
                                        f10.Show();
                                    }
                                }
                            }
                            catch
                            {
                                return;
                            }
                        }

                        if (flag1 == true && loupe == false && clickCountElectronogram == 1)
                        {
                            try
                            {
                                MD = false;
                                clickMM = 0;
                                chart1.Series[1].Points.Clear();
                                dataGridView1.Rows.Add();
                                Di = Math.Round(Math.Sqrt((Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value)) * (Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value)) + (Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value)) * (Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value))), 2);
                                Di = Di * n / n_const;
                                dii.Insert(kol_count_di, Di);
                                kol_count_di = kol_count_di + 1;
                                num_line = num_line + 1;
                                if (tabControl1.SelectedIndex == 0)
                                {
                                    if (numDGV2-I_dGV2 < 2)
                                    {
                                        dataGridView2.Rows.Add();
                                        dataGridView2.Rows[I_dGV2+1].Cells[0].Value = I_dGV2+2;
                                    }
                                    dataGridView2.Rows[I_dGV2].Cells[1].Value = Math.Round(Di,2);
                                    I_dGV2 = I_dGV2 + 1;
                                }
                                if (tabControl1.SelectedIndex == 1)
                                {
                                    if (numDGV3 - I_dGV3 < 2)
                                    {
                                        dataGridView3.Rows.Add();
                                        dataGridView3.Rows[I_dGV3+1].Cells[0].Value = I_dGV3 + 2;
                                    }
                                    dataGridView3.Rows[I_dGV3].Cells[1].Value = Math.Round(Di, 2);
                                    I_dGV3 = I_dGV3 + 1;
                                }
                            }
                            catch
                            {
                                return;
                            }
                        }

                        break;

                       
                }

                
        }
            if (e.Button == MouseButtons.Right)
            {
                if (flag1 == true && loupe == false && clickMM == 1)
                {

                    MD = false;
                    clickMM = 0;
                    chart1.Series[0].Points.Clear();
                    chart1.Series[1].Points.Clear();
                    chart1.Series[I].Points.Clear();
                    chart1.Series.Remove(chart1.Series[I]);
                    
                    dataGridView1.Rows.Remove(dataGridView1.Rows[num_line]);
                    num_line = num_line - 1;
                    I = I - 1;
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (kol_count_di < 1)
            {
                if (Rus == true)
                {
                    MessageBox.Show("Экспериментальные данные отсуствуют!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (Eng == true)
                {
                    MessageBox.Show("No experimental data available!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
            

            dii.Sort();
            if (form4 == false)
            {
                f4 = new Form4(this);
                f4.Owner = this;
                f4.Show();
                form4 = true;
            }
        }

        public void button8_Click(object sender, EventArgs e)
        {
            if (clickLineProfile == 1)
                return;
            if (I - 4 >= 0)
            {
                chart1.Series[I].Points.Clear();
                chart1.Series.Remove(chart1.Series[I]);
                dataGridView1.Rows.Remove(dataGridView1.Rows[num_line]);
                dataGridView1.Rows.Remove(dataGridView1.Rows[num_line - 1]);
                num_line = num_line - 2;
                if (calibrateImage)
                    if (N_calImage == 1)
                        N_calImage = 0;
              
                    number_N_meas_fault = number_N_meas_fault - 1;
                    N_meas_faultt.Insert(kol_meas_fault, number_N_meas_fault);
              
                kol_count_di = kol_count_di - 1;                
                I = I - 1;
                if (N_calImage==0)
                {
                    textBox7.Text = kol_count_di.ToString();
                    if (kol_count_di == 0)
                    {
                        textBox1.Text = 0.ToString();
                        // number_N_meas_fault = 0;
                        //if (kol_meas_fault >= 0)
                        // N_meas_fault[kol_meas_fault] = number_N_meas_fault;
                    }
                    else
                        textBox1.Text = Math.Round(dii[kol_count_di - 1], 2).ToString();
                 
                }
                if (loupe1)
                {
                    Program.f2.chart1.Series[Program.f2.I].Points.Clear();
                    Program.f2.chart1.Series.Remove(Program.f2.chart1.Series[Program.f2.I]);
                    Program.f2.dataGridView1.Rows.Remove(Program.f2.dataGridView1.Rows[Program.f2.num_line]);
                    Program.f2.dataGridView1.Rows.Remove(Program.f2.dataGridView1.Rows[Program.f2.num_line - 1]);
                    Program.f2.num_line = Program.f2.num_line - 2;
                    Program.f2.I = Program.f2.I - 1;
                }
            }

            if (groupBox3.Visible == true)
            {
                if (tabControl1.SelectedIndex == 0)
                {
                    if (I_dGV2 > 0)
                    {
                        dataGridView2.Rows[I_dGV2 - 1].Cells[1].Value = string.Empty;
                        I_dGV2 = I_dGV2 - 1;
                    }
                }
                if (tabControl1.SelectedIndex == 1)
                {
                    if (I_dGV3 > 0)
                    {
                        dataGridView3.Rows[I_dGV3 - 1].Cells[1].Value = string.Empty;
                        I_dGV3 = I_dGV3 - 1;
                    }
                }
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (Program.f1.f3_open)
            {
                form4 = true;
                Program.f3.Close();
                form4 = false;
            }
            f3 = new Form3();
            f3.Show();
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            trackBar1.Value = (int)(numericUpDown1.Value - 50) * 2;
            trackBar1_Scroll(sender, e);
        }

        private void numericUpDown1_KeyUp(object sender, KeyEventArgs e)
        {
            trackBar1.Value = (int)(numericUpDown1.Value - 50) * 2;
            trackBar1_Scroll(sender, e);
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            trackBar2.Value = (int)(numericUpDown2.Value - 50) * 2;
            trackBar2_Scroll(sender, e);
        }

        private void numericUpDown2_KeyUp(object sender, KeyEventArgs e)
        {
            trackBar2.Value = (int)(numericUpDown2.Value - 50) * 2;
            trackBar2_Scroll(sender, e);
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            trackBar3.Value = (int)(numericUpDown3.Value - 50) * 2;
            trackBar3_Scroll(sender, e);
        }

        private void numericUpDown3_KeyUp(object sender, KeyEventArgs e)
        {
            trackBar3.Value = (int)(numericUpDown3.Value - 50) * 2;
            trackBar3_Scroll(sender, e);
        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            trackBar4.Value = (int)(numericUpDown4.Value - 50) * 2;
            trackBar4_Scroll(sender, e);
        }

        private void numericUpDown4_KeyUp(object sender, KeyEventArgs e)
        {
            trackBar4.Value = (int)(numericUpDown4.Value - 50) * 2;
            trackBar4_Scroll(sender, e);
        }

        private void numericUpDown5_ValueChanged(object sender, EventArgs e)
        {
            trackBar5.Value = (int)(numericUpDown5.Value - 50) * 2;
            trackBar5_Scroll(sender, e);
        }

        private void numericUpDown5_KeyUp(object sender, KeyEventArgs e)
        {
            trackBar5.Value = (int)(numericUpDown5.Value - 50) * 2;
            trackBar5_Scroll(sender, e);
        }

        private void panel2_MouseEnter(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Arrow;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            ++clickCountLoupe2x;

            switch (clickCountLoupe2x)
            {
                case 1:
                    {
                        button12.BackColor = System.Drawing.SystemColors.Info;
                        button13.BackColor = System.Drawing.SystemColors.Control;
                        button14.BackColor = System.Drawing.SystemColors.Control;
                        button15.BackColor = System.Drawing.SystemColors.Control;
                        magnificationX = Lx / 4;
                        magnificationY = Ly / 4;
                        magnificX = Lx / 2;
                        magnificY = Ly / 2;
                        l2x = true;
                        l4x = false;
                        l8x = false;
                        l16x = false;
                        clickCountLoupe4x = 0;
                        clickCountLoupe8x = 0;
                        clickCountLoupe16x = 0;
                    }
                    break;

                case 2:
                    {
                        button12.Visible = false;
                        button13.Visible = false;
                        button14.Visible = false;
                        button15.Visible = false;
                        clickCountLoupe2x = 0;
                        loupe = false;
                        if (loupe1 == true)
                        {
                            f2.Close();
                            loupe1 = false;
                        }
                        else
                            chart1.Series[3].Points.Clear();
                    }
                    break;
            }
            
        }

        private void button16_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < I_dGV3; i++)
                dataGridView3.Rows[i].Cells[1].Value = string.Empty;           
            I_dGV3 = 0;
            for (int i = I; i > 3; i--)
            {
                chart1.Series[I].Points.Clear();
                chart1.Series.Remove(chart1.Series[I]);
                dataGridView1.Rows.Remove(dataGridView1.Rows[num_line]);
                dataGridView1.Rows.Remove(dataGridView1.Rows[num_line - 1]);
                num_line = num_line - 2;
                I = I - 1;
            }
        }

        

        private void chart1_MouseLeave(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Arrow;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            ++clickCountLoupe4x;
            switch (clickCountLoupe4x)
            {
                case 1:
                    {
                        button12.BackColor = System.Drawing.SystemColors.Control;
                        button13.BackColor = System.Drawing.SystemColors.Info;
                        button14.BackColor = System.Drawing.SystemColors.Control;
                        button15.BackColor = System.Drawing.SystemColors.Control;
                        magnificationX = Lx / 8;
                        magnificationY = Ly / 8;
                        magnificX = Lx / 4;
                        magnificY = Ly / 4;
                        l2x = false;
                        l4x = true;
                        l8x = false;
                        l16x = false;
                        clickCountLoupe2x = 0;
                        clickCountLoupe8x = 0;
                        clickCountLoupe16x = 0;
                    }
                    break;
                case 2:
                    {
                        button12.Visible = false;
                        button13.Visible = false;
                        button14.Visible = false;
                        button15.Visible = false;
                        clickCountLoupe4x = 0;
                        loupe = false;
                        if (loupe1 == true)
                        {
                            f2.Close();
                            loupe1 = false;
                        }
                        else
                            chart1.Series[3].Points.Clear();
                    }
                    break;
            }

            
        }

        private void button14_Click(object sender, EventArgs e)
        {
            ++clickCountLoupe8x;
            switch(clickCountLoupe8x)
            {
                case 1:
                    {
                        button12.BackColor = System.Drawing.SystemColors.Control;
                        button13.BackColor = System.Drawing.SystemColors.Control;
                        button14.BackColor = System.Drawing.SystemColors.Info;
                        button15.BackColor = System.Drawing.SystemColors.Control;
                        magnificationX = Lx / 16;
                        magnificationY = Ly / 16;
                        magnificX = Lx / 8;
                        magnificY = Ly / 8;
                        l2x = false;
                        l4x = false;
                        l8x = true;
                        l16x = false;
                        clickCountLoupe2x = 0;
                        clickCountLoupe4x = 0;
                        clickCountLoupe16x = 0;
                    }
                    break;
                case 2:
                    {
                        button12.Visible = false;
                        button13.Visible = false;
                        button14.Visible = false;
                        button15.Visible = false;
                        clickCountLoupe8x = 0;
                        loupe = false;
                        if(loupe1 == true)
                            {
                              f2.Close();                                
                              loupe1 = false;
                             }
                        else                        
                        chart1.Series[3].Points.Clear();
                    }
                    break;
            }
            
        }

        private void button15_Click(object sender, EventArgs e)
        {
            ++clickCountLoupe16x;
            switch(clickCountLoupe16x)
                {
                case 1:
                    {
                        button12.BackColor = System.Drawing.SystemColors.Control;
                        button13.BackColor = System.Drawing.SystemColors.Control;
                        button14.BackColor = System.Drawing.SystemColors.Control;
                        button15.BackColor = System.Drawing.SystemColors.Info;
                        magnificationX = Lx / 32;
                        magnificationY = Ly / 32;
                        magnificX = Lx / 16;
                        magnificY = Ly / 16;
                        l2x = false;
                        l4x = false;
                        l8x = false;
                        l16x = true;
                        clickCountLoupe2x = 0;
                        clickCountLoupe4x = 0;
                        clickCountLoupe8x = 0;
                    }
                    break;
                case 2:
                    {
                        button12.Visible = false;
                        button13.Visible = false;
                        button14.Visible = false;
                        button15.Visible = false;
                        clickCountLoupe16x = 0;
                        loupe = false;
                        if (loupe1 == true)
                        {
                            f2.Close();
                            loupe1 = false;
                        }
                        else
                            chart1.Series[3].Points.Clear();
                    }
                    break;
                }
           
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (this.WindowState==FormWindowState.Minimized)
            {
                if (calibrateImage)
                Program.f9.TopMost = false;
                if (FormLineProphile)
                Program.f10.TopMost = false;
            }
            if (this.WindowState == FormWindowState.Maximized)
            {
                if (calibrateImage)
                    Program.f9.TopMost = true;
                if (FormLineProphile)
                    Program.f10.TopMost = true;
            }
        }

        private void Form1_Deactivate(object sender, EventArgs e)
        {
            
                if (calibrateImage)
                    Program.f9.TopMost = false;
                if (FormLineProphile)
                    Program.f10.TopMost = false;
            
        }

        


        private void trackBar1_Scroll(object sender, EventArgs e) // яркость
        {
            numericUpDown1.Value = 50+trackBar1.Value / 2;
            if (flag1)
            {
                    panel1.BackgroundImage = AdjustBrightnessContrast(bmp, trackBar1.Value, trackBar2.Value, trackBar3.Value, trackBar4.Value, trackBar5.Value);   
                if (openf2)                               
                    Program.f2.panel1.BackgroundImage = AdjustBrightnessContrast((Bitmap)Program.f2.img, trackBar1.Value, trackBar2.Value, trackBar3.Value, trackBar4.Value, trackBar5.Value);
                
            }
        }

        private void trackBar2_Scroll(object sender, EventArgs e) // контраст
        {
            numericUpDown2.Value = 50 + trackBar2.Value / 2;
            if (flag1)
            {
                panel1.BackgroundImage = AdjustBrightnessContrast(bmp, trackBar1.Value, trackBar2.Value, trackBar3.Value, trackBar4.Value, trackBar5.Value);
                if (openf2)
                    Program.f2.panel1.BackgroundImage = AdjustBrightnessContrast((Bitmap)Program.f2.img, trackBar1.Value, trackBar2.Value, trackBar3.Value, trackBar4.Value, trackBar5.Value);
            }
        }

       

        private void trackBar3_Scroll(object sender, EventArgs e) // красный
        {
            numericUpDown3.Value = 50 + trackBar3.Value / 2;
            if (flag1)
            {
                
                panel1.BackgroundImage = AdjustBrightnessContrast(bmp, trackBar1.Value, trackBar2.Value, trackBar3.Value, trackBar4.Value, trackBar5.Value);
                if (openf2)
                    Program.f2.panel1.BackgroundImage = AdjustBrightnessContrast((Bitmap)Program.f2.img, trackBar1.Value, trackBar2.Value, trackBar3.Value, trackBar4.Value, trackBar5.Value);
            }
        }

        private void trackBar4_Scroll(object sender, EventArgs e) // зеленый
        {
            numericUpDown4.Value = 50 + trackBar4.Value / 2;
            if (flag1)
            {
                
                panel1.BackgroundImage = AdjustBrightnessContrast(bmp, trackBar1.Value, trackBar2.Value, trackBar3.Value, trackBar4.Value, trackBar5.Value);
                if (openf2)
                    Program.f2.panel1.BackgroundImage = AdjustBrightnessContrast((Bitmap)Program.f2.img, trackBar1.Value, trackBar2.Value, trackBar3.Value, trackBar4.Value, trackBar5.Value);
            }
        }

        private void trackBar5_Scroll(object sender, EventArgs e) // синий
        {
            numericUpDown5.Value = 50 + trackBar5.Value / 2;
            if (flag1)
            {                
                panel1.BackgroundImage = AdjustBrightnessContrast(bmp, trackBar1.Value, trackBar2.Value, trackBar3.Value, trackBar4.Value, trackBar5.Value);
                if (openf2)
                    Program.f2.panel1.BackgroundImage = AdjustBrightnessContrast((Bitmap)Program.f2.img, trackBar1.Value, trackBar2.Value, trackBar3.Value, trackBar4.Value, trackBar5.Value);
            }
        }

        public static Bitmap AdjustBrightnessContrast(Bitmap Image, int brightnessValue, int contrastValue, int RedValue, int GreenValue, int BlueValue) // Яркость/Контраст 
        {

            Bitmap TempBitmap = Image;

            Bitmap NewBitmap = new Bitmap(TempBitmap.Width, TempBitmap.Height);

            Graphics NewGraphics = Graphics.FromImage(NewBitmap);

            
            brightness = brightnessValue / 100.0f;
            RedContrast = RedValue / 100.0f + 0.50f;
            GreenContrast=GreenValue/100.0f + 0.50f;
            BlueContrast = BlueValue / 100.0f + 0.50f;
            
            contrast = contrastValue / 100.0f+0.50f;

            float[][] FloatColorMatrix ={

                    new float[] {contrast+RedContrast, 0, 0, 0, 0},

                    new float[] {0, contrast+GreenContrast, 0, 0, 0},

                    new float[] {0, 0, contrast+BlueContrast, 0, 0},

                    new float[] {0, 0, 0, 1, 0},

                    new float[] {brightness, brightness, brightness, 1, 1}
                };

            ColorMatrix NewColorMatrix = new ColorMatrix(FloatColorMatrix);

            ImageAttributes Attributes = new ImageAttributes();

            Attributes.SetColorMatrix(NewColorMatrix);

            NewGraphics.DrawImage(TempBitmap, new Rectangle(0, 0, TempBitmap.Width, TempBitmap.Height), 0, 0, TempBitmap.Width, TempBitmap.Height, GraphicsUnit.Pixel, Attributes);

            Attributes.Dispose();

            NewGraphics.Dispose();

            return NewBitmap;
        }

        private void button3_MouseHover(object sender, EventArgs e)
        {
            t.SetToolTip(button3, Calculate_d);
        }

        private void button10_MouseHover(object sender, EventArgs e)
        {
            t.SetToolTip(button10, LineProfile);
        }

        private void button4_MouseHover(object sender, EventArgs e)
        {
            t.SetToolTip(button4, Calculate_C);
        }

        private void dataGridView3_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            
                dataGridView3.Rows[numDGV3 - 1].Cells[0].Value = numDGV3;
                numDGV3 = numDGV3 + 1;
          
        
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
                ConstMicro = 0;
                for (int i = 0; i < I_dGV3; i++)
                {
                    ConstMicro = ConstMicro + Convert.ToDouble(dataGridView3.Rows[i].Cells[1].Value) * Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value);
                }
                ConstMicro = Math.Round(ConstMicro / I_dGV3, 2);
            
            if (double.IsNaN(ConstMicro) || ConstMicro==0)
                {
                   if (Rus)
                    MessageBox.Show("Введите данные", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                   if (Eng)
                    MessageBox.Show("No data available", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button4.BackColor = SystemColors.Control;
                return;
                }

            if (Rus)
            {
                MessageBox.Show("Постоянная микроскопа равна " + ConstMicro, "Сообщение!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                button4.BackColor = Color.Lime;
            }
            if (Eng)
            {
                MessageBox.Show("Microscope constant is equal " + ConstMicro, "Message!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                button4.BackColor = Color.Lime;
            }
            
            
                
            
        }

        private void dataGridView2_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            
                dataGridView2.Rows[numDGV2 - 1].Cells[0].Value = numDGV2;
                numDGV2 = numDGV2 + 1;
           
        }

        private void button3_Click(object sender, EventArgs e)
        {
                      
                for (int i = 0; i < I_dGV2; i++)
                    dataGridView2.Rows[i].Cells[2].Value = Math.Round(ConstMicro / Convert.ToDouble(dataGridView2.Rows[i].Cells[1].Value),2);
            
            
        }


        private void dataGridView3_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (dataGridView3.CurrentCell.ColumnIndex == 2)
            {
                TextBox tb = (TextBox)e.Control;
                tb.KeyPress += new KeyPressEventHandler(tb_KeyPress);
            }
            else
            {
                TextBox tb = (TextBox)e.Control;
                tb.KeyPress -= tb_KeyPress;
            }
        }

        void tb_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar)) && (e.KeyChar != (char)Keys.Back) && !(e.KeyChar == ','))
            {
                e.Handled = true;
            }
        }

        private void button10_MouseClick(object sender, MouseEventArgs e)
        {
            if (flag1 == true)
            {
                if (loupe)
                {
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows.Add();
                }

                if (clickCountCalibrationImage==1)
                {
                    f9.Close();
                }

                if (clickCountElectronogram==1)
                {
                    button11.BackColor = System.Drawing.SystemColors.Control;
                    groupBox3.Visible = false;
                    clickCountElectronogram = 0;
                }

                ++clickLineProfile;
                switch (clickLineProfile)
                {
                    case 1:
                        if (e.Button == MouseButtons.Left)
                        {
                            button10.BackColor = System.Drawing.SystemColors.Info;
                        }
                        break;
                    case 2:
                        if (e.Button == MouseButtons.Left)
                        {
                            button10.BackColor = System.Drawing.SystemColors.Control;
                            clickLineProfile = 0;
                            if (FormLineProphile)                                                          
                                f10.Close();
                        }
                        break;
                }
                
                
            }
                
            
        }

        private void button6_MouseEnter(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Arrow;
        }

        private void button11_MouseClick(object sender, MouseEventArgs e)
        {
            if (flag1 == true && loupe==false)
            {
                if (clickCountCalibrationImage==1)
                {
                    f9.Close();
                }

                if (clickLineProfile==1)
                {
                    button10.BackColor = System.Drawing.SystemColors.Control;
                    clickLineProfile = 0;
                    if (FormLineProphile)
                        f10.Close();
                }

                ++clickCountElectronogram;

                switch (clickCountElectronogram)
                {
                    case 1:

                        if (e.Button==MouseButtons.Left)
                        {
                            button11.BackColor = System.Drawing.SystemColors.Info;
                            groupBox3.Location = new Point(groupBox2.Location.X, groupBox2.Location.Y);
                            groupBox3.Visible = true;
                            

                        }
                        break;

                    case 2:
                        if (e.Button == MouseButtons.Left)
                        {
                            button11.BackColor = System.Drawing.SystemColors.Control;
                            groupBox3.Visible = false;
                            clickCountElectronogram = 0;
                            for (int i = I; i > 3; i--)
                            {
                                chart1.Series[I].Points.Clear();
                                chart1.Series.Remove(chart1.Series[I]);
                                dataGridView1.Rows.Remove(dataGridView1.Rows[num_line]);
                                dataGridView1.Rows.Remove(dataGridView1.Rows[num_line - 1]);
                                num_line = num_line - 2;
                                I = I - 1;
                            }
                            for (int i = 0; i < I_dGV2; i++)
                                dataGridView2.Rows[i].Cells[1].Value = string.Empty;
                            I_dGV2 = 0;                    
                            kol_count_di = 0;
                        }
                        break;
                }
            }
        }

        private void button7_MouseEnter(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Arrow;
        }

        private void button8_MouseEnter(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Arrow;
        }

        private void button1_MouseEnter(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Arrow;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (flag1)
            {
                if (clickLineProfile==1)
                {
                    button10.BackColor = System.Drawing.SystemColors.Control;
                    clickLineProfile = 0;
                    if (FormLineProphile)
                        f10.Close();
                }

                if (clickCountElectronogram==1)
                {
                    button11.BackColor = System.Drawing.SystemColors.Control;
                    groupBox3.Visible = false;
                    clickCountElectronogram = 0;
                }
                ++clickCountCalibrationImage;

                    switch (clickCountCalibrationImage)
                    {
                    case 1:
                        {
                            button6.BackColor = System.Drawing.SystemColors.Info;
                            calibrateImage = true;
                            f9 = new Form9();
                            f9.Owner = this;
                            f9.Show();
                        }
                        break;
                    case 2:
                        {
                            f9.Close();
                        }
                        break;
                }
                
            }
        }

        private void button9_MouseEnter(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Arrow;
        }

        private void button11_MouseEnter(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Arrow;
        }

        private void button2_MouseEnter(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Arrow;
        }


        private void button6_MouseHover(object sender, EventArgs e)
        {
            t.SetToolTip(button6, Calibrate);
        }

        private void chart1_MouseMove_1(object sender, MouseEventArgs e)
        {

            if (MD)
            {
                if (clickCountElectronogram!=1 && clickLineProfile!=1)                
                try
                {
                    lastpoint = e.Location;
                    x_MouseDown = Math.Round(chart1.ChartAreas[0].AxisX.PixelPositionToValue(lastpoint.X), 0);
                    y_MouseDown = Math.Round(chart1.ChartAreas[0].AxisY.PixelPositionToValue(lastpoint.Y), 0);
                    dataGridView1.Rows[num_line].Cells[0].Value = x_MouseDown;
                    dataGridView1.Rows[num_line].Cells[1].Value = y_MouseDown;
                    chart1.Series[I].Points.Clear();
                    for (int i = num_line - 1; i <= num_line; i++)
                        chart1.Series[I].Points.AddXY(dataGridView1.Rows[i].Cells[0].Value, dataGridView1.Rows[i].Cells[1].Value);
                    dataGridView1.Rows[0].Cells[2].Value = 0;
                    dataGridView1.Rows[0].Cells[3].Value = Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value) - ((Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value)) * (0 - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value))) / (Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value));
                    dataGridView1.Rows[0].Cells[4].Value = Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value) - ((Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value)) * (0 - Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value))) / (Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value));
                    dataGridView1.Rows[1].Cells[2].Value = W;
                    dataGridView1.Rows[1].Cells[3].Value = Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value) - ((Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value)) * (W - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value))) / (Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value));
                    dataGridView1.Rows[1].Cells[4].Value = Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value) - ((Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value)) * (W - Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value))) / (Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value));

                    chart1.Series[0].Points.Clear();
                    chart1.Series[1].Points.Clear();
                    if (double.IsInfinity(Convert.ToDouble(dataGridView1.Rows[0].Cells[3].Value)) || double.IsInfinity(Convert.ToDouble(dataGridView1.Rows[0].Cells[4].Value)) || double.IsInfinity(Convert.ToDouble(dataGridView1.Rows[1].Cells[3].Value)) || double.IsInfinity(Convert.ToDouble(dataGridView1.Rows[1].Cells[4].Value)))
                        return;
                    else
                    {
                        for (int i = 0; i <= 1; i++)
                        {
                            chart1.Series[0].Points.AddXY(dataGridView1.Rows[i].Cells[2].Value, dataGridView1.Rows[i].Cells[3].Value);
                            chart1.Series[1].Points.AddXY(dataGridView1.Rows[i].Cells[2].Value, dataGridView1.Rows[i].Cells[4].Value);
                        }
                    }
                }
                catch
                {
                    return;
                }
                if (clickCountElectronogram==1 && clickLineProfile!=1)
                    try
                    {
                        lastpoint = e.Location;
                        x_MouseDown = Math.Round(chart1.ChartAreas[0].AxisX.PixelPositionToValue(lastpoint.X), 0);
                        y_MouseDown = Math.Round(chart1.ChartAreas[0].AxisY.PixelPositionToValue(lastpoint.Y), 0);
                        dataGridView1.Rows[num_line].Cells[0].Value = x_MouseDown;
                        dataGridView1.Rows[num_line].Cells[1].Value = y_MouseDown;
                        chart1.Series[I].Points.Clear();
                        L = Math.Round(Math.Sqrt((Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value)) * (Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value) - (Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value))) + (Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value)) * (Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value))),2);
                        if (L == 0)
                            return;
                        Rc = Convert.ToInt32(L / 2);
                        if (Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value) > Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value))
                            xc = Math.Round((Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value)) / 2 + Convert.ToDouble(dataGridView1.Rows[num_line-1].Cells[0].Value), 2);
                        else
                            xc= Math.Round((Convert.ToDouble(dataGridView1.Rows[num_line-1].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value)) / 2 + Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value), 2);
                        if (Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value) > Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value))
                            yc = Math.Round((Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value)) / 2 + Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value), 2);
                        else
                            yc = Math.Round((Convert.ToDouble(dataGridView1.Rows[num_line-1].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value)) / 2 + Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value), 2);
                        int xcoord = Convert.ToInt32(xc);
                        for (int i=(xcoord-Rc);i<=(xcoord+Rc);i=i+3)
                        {
                            double C = yc * yc + i * i - 2 * i * xc + xc * xc - Rc * Rc;
                            double Discriminant = 4 * yc * yc - 4 * C;
                            double y = (2 * yc - Math.Sqrt(Discriminant)) / 2;
                            chart1.Series[I].Points.AddXY(i, y);                           
                        }
                        for (int i=(xcoord+Rc);i>=(xcoord-Rc);i=i-3)
                        {
                            double C = yc * yc + i * i - 2 * i * xc + xc * xc - Rc * Rc;
                            double Discriminant = 4 * yc * yc - 4 * C;
                            double y = (2 * yc + Math.Sqrt(Discriminant)) / 2;
                            chart1.Series[I].Points.AddXY(i, y);
                        }
                        chart1.Series[1].Points.Clear();
                        if (double.IsInfinity(Convert.ToDouble(dataGridView1.Rows[0].Cells[3].Value)) || double.IsInfinity(Convert.ToDouble(dataGridView1.Rows[0].Cells[4].Value)) || double.IsInfinity(Convert.ToDouble(dataGridView1.Rows[1].Cells[3].Value)) || double.IsInfinity(Convert.ToDouble(dataGridView1.Rows[1].Cells[4].Value)))
                            return;
                        else
                        {
                            for (int i = num_line-1; i <= num_line; i++)
                            {
                                chart1.Series[1].Points.AddXY(dataGridView1.Rows[i].Cells[0].Value, dataGridView1.Rows[i].Cells[1].Value);                              
                            }
                        }
                    }
                    catch
                    {
                        return;
                    }
                if (clickLineProfile==1)
                {
                    try
                    {
                        lastpoint = e.Location;
                        x_MouseDown = Math.Round(chart1.ChartAreas[0].AxisX.PixelPositionToValue(lastpoint.X), 0);
                        y_MouseDown = Math.Round(chart1.ChartAreas[0].AxisY.PixelPositionToValue(lastpoint.Y), 0);
                        dataGridView1.Rows[num_line].Cells[0].Value = x_MouseDown;
                        dataGridView1.Rows[num_line].Cells[1].Value = y_MouseDown;
                        chart1.Series[I].Points.Clear();
                        for (int i = num_line - 1; i <= num_line; i++)
                            chart1.Series[I].Points.AddXY(dataGridView1.Rows[i].Cells[0].Value, dataGridView1.Rows[i].Cells[1].Value);
                        dataGridView1.Rows[0].Cells[2].Value = 0;
                        dataGridView1.Rows[0].Cells[3].Value = Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value) - ((Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value)) * (0 - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value))) / (Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value));
                        dataGridView1.Rows[0].Cells[4].Value = Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value) - ((Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value)) * (0 - Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value))) / (Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value));
                        dataGridView1.Rows[1].Cells[2].Value = W;
                        dataGridView1.Rows[1].Cells[3].Value = Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value) - ((Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value)) * (W - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value))) / (Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value));
                        dataGridView1.Rows[1].Cells[4].Value = Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value) - ((Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[0].Value)) * (W - Convert.ToDouble(dataGridView1.Rows[num_line].Cells[0].Value))) / (Convert.ToDouble(dataGridView1.Rows[num_line].Cells[1].Value) - Convert.ToDouble(dataGridView1.Rows[num_line - 1].Cells[1].Value));

                        chart1.Series[0].Points.Clear();
                        chart1.Series[1].Points.Clear();
                        if (double.IsInfinity(Convert.ToDouble(dataGridView1.Rows[0].Cells[3].Value)) || double.IsInfinity(Convert.ToDouble(dataGridView1.Rows[0].Cells[4].Value)) || double.IsInfinity(Convert.ToDouble(dataGridView1.Rows[1].Cells[3].Value)) || double.IsInfinity(Convert.ToDouble(dataGridView1.Rows[1].Cells[4].Value)))
                            return;
                        else
                        {
                            for (int i = 0; i <= 1; i++)
                            {
                                chart1.Series[0].Points.AddXY(dataGridView1.Rows[i].Cells[2].Value, dataGridView1.Rows[i].Cells[3].Value);
                                chart1.Series[1].Points.AddXY(dataGridView1.Rows[i].Cells[2].Value, dataGridView1.Rows[i].Cells[4].Value);
                            }
                        }
                    }
                    catch
                    {
                        return;
                    }
                }
            }
            if (flag1 == true && loupe == true)
            {
                //  bmp = ListBmp2[0];
                // panel1.BackgroundImage = bmp;

                // ListBmp2.Insert(0, new Bitmap(bmp));

                //gr = Graphics.FromImage(bmp);
                chart1.Series[3].Points.Clear();
                lastpoint = e.Location;

                x_MouseDown = Math.Round(chart1.ChartAreas[0].AxisX.PixelPositionToValue(lastpoint.X),0);
                y_MouseDown = Math.Round(chart1.ChartAreas[0].AxisY.PixelPositionToValue(lastpoint.Y),0);
                x1_Rect = Convert.ToInt32(x_MouseDown) - magnificationX;
                y1_Rect = Convert.ToInt32(y_MouseDown) - magnificationY;

                //X1_chart = (Math.Round(chart1.ChartAreas[0].AxisX.PixelPositionToValue(lastpoint.X), 0) - magnificationX) - (Math.Round(chart1.ChartAreas[0].AxisX.PixelPositionToValue(lastpoint.X), 0) - x_MouseDown);
               // Y1_chart = (Math.Round(chart1.ChartAreas[0].AxisY.PixelPositionToValue(lastpoint.Y), 0) - magnificationY) - (Math.Round(chart1.ChartAreas[0].AxisY.PixelPositionToValue(lastpoint.Y), 0) - (H - y_MouseDown));

                if (x1_Rect < 0)
                {
                    x1_Rect = 0;
                   // X1_chart = 0;
                }
                if (y1_Rect < 0)
                {
                    y1_Rect = 0;
                   // Y1_chart = H - magnificY;
                }
                if (x1_Rect > W - magnificX)
                {
                    x1_Rect = W - magnificX;
                    //X1_chart = x_Rect;
                }
                if (y1_Rect > H - magnificY)
                {
                    y1_Rect = H - magnificY;
                    //Y1_chart = 0;

                }

                chart1.Series[3].Points.AddXY(x1_Rect, y1_Rect);
                chart1.Series[3].Points.AddXY(x1_Rect, y1_Rect + magnificY);
                chart1.Series[3].Points.AddXY(x1_Rect + magnificX, y1_Rect + magnificY);
                chart1.Series[3].Points.AddXY(x1_Rect + magnificX, y1_Rect);
                chart1.Series[3].Points.AddXY(x1_Rect, y1_Rect);

                if (openf2)
                if (x1_Rect == x_Rect && y1_Rect == y_Rect)
                    chart1.Series[3].Points.Clear();

                //gr.DrawRectangle(myPen, x_Rect, y_Rect, magnificX, magnificY);
                //Rectangle rect = new Rectangle(x_Rect, H-y_Rect-magnificY, magnificX, magnificY);
              

               // ListBmp2.Insert(1, new Bitmap(bmp));

               // bmp = ListBmp2[1];
               // panel1.BackgroundImage = bmp;
                
            }

            
        }

        private void button7_MouseHover(object sender, EventArgs e)
        {
            t.SetToolTip(button7, Histogram);
        }

        private void button8_MouseHover(object sender, EventArgs e)
        {
            t.SetToolTip(button8, Return);
        }

        private void button1_MouseHover(object sender, EventArgs e)
        {
            t.SetToolTip(button1, Zoom);
        }

        private void button9_MouseHover(object sender, EventArgs e)
        {
            t.SetToolTip(button9, LogNorm);
        }

        private void button11_MouseHover(object sender, EventArgs e)
        {
            t.SetToolTip(button11, Electronogram);
        }

        private void button2_MouseHover(object sender, EventArgs e)
        {
            t.SetToolTip(button2, Average);
        }


        private void textBox1_Enter(object sender, EventArgs e)
        {
            button5.Select();
        }

        private void textBox3_Enter(object sender, EventArgs e)
        {
            button5.Select();
        }


        private void оПрограммеИИнструкцииПоРаботеToolStripMenuItem_Click(object sender, EventArgs e)
        {                       
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.OverwritePrompt = true; // Отображать ли предупреждение, если пользователь указывает имя уже существующего файла
            dialog.CheckPathExists = true; // Отображать ли предупреждение, если пользователь указывает несуществующий путь
            dialog.Filter = "Текст в формате RTF|*.rtf";
            if (Rus == true)
            {
                string resString = Properties.Resources.О_программе_и_инструкции_по_работе_с_ней;
                dialog.FileName = "О программе и инструкции по работе";
                if (dialog.ShowDialog() == DialogResult.OK)
                    File.WriteAllText(dialog.FileName, resString);
                if (Rus)
                    MessageBox.Show("Успешно сохранено!", "Сообщение!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
            }
            if (Eng==true)
            {
                string resString = Properties.Resources.About_DEAM_and_work_instruction;
                dialog.FileName = "About DEAM and work instruction";
                if (dialog.ShowDialog() == DialogResult.OK)
                    File.WriteAllText(dialog.FileName, resString);
                if (Eng)
                    MessageBox.Show("The file was successfully saved!", "Message!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            

        }

        private void загрузитьЭкспериментальныеДанныеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.txt;*.dat;*.xls;*.xlsx";
            if (Rus == true)
            {
                ofd.Title = "Выберите документ для загрузки данных";
                ofd.Filter = "Текстовый файл (*.txt;*.dat)|*.txt;*.dat|Файлы excel (*.xls; *.xlsx)|*.xls; *.xlsx";
            }
            if (Eng == true)
            {
                ofd.Title = "Chose the document to load data";
                ofd.Filter = "Textfile (*.txt;*.dat)|*.txt;*.dat|Excel file (*.xls; *.xlsx)|*.xls; *.xlsx";
            }
            ExcelObj.Application app = new ExcelObj.Application();
            ExcelObj.Workbook workbook;
            ExcelObj.Worksheet NwSheet;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string ext = System.IO.Path.GetExtension(ofd.FileName);   
                dii.Clear();
                if (ext== ".txt" || ext==".dat")
                try
                {
                    
                    kol_count_di = 0;
                    StreamReader f = new StreamReader(ofd.FileName);                   
                    string s = f.ReadLine();
                    while (s != null && s != string.Empty)
                    {
                        dii.Insert(kol_count_di,Convert.ToDouble(s));
                        kol_count_di = kol_count_di + 1;
                        s = f.ReadLine();
                    }
                    if (Rus == true)
                        MessageBox.Show("Экспериментальные данные успешно загружены!", "Сообщение!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (Eng == true)
                        MessageBox.Show("Experimental data are imported successfully!", "Message!", MessageBoxButtons.OK, MessageBoxIcon.Information);                       
                }
                catch
                {
                    if (Rus == true)
                        MessageBox.Show("Невозможно загрузить выбранный файл", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (Eng == true)
                        MessageBox.Show("It is impossible to import the selected file", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                if(ext == ".xls" || ext == ".xlsx")                   
                try
                    {
                        kol_count_di = 0;
                        workbook = app.Workbooks.Open(ofd.FileName, Missing.Value,
                      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                      Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                      Missing.Value);
                        NwSheet = (ExcelObj.Worksheet)workbook.Sheets.get_Item(1);
                        var lastCell = NwSheet.Cells.SpecialCells(ExcelObj.XlCellType.xlCellTypeLastCell);
                        for (int Rnum = 1; Rnum <= lastCell.Row; Rnum++)
                        {
                            if (NwSheet.Cells[Rnum, 1] != null)// && NwSheet.Cells[Rnum, 1] != string.Empty)
                            {
                                dii.Insert(kol_count_di, Convert.ToDouble(NwSheet.Cells[Rnum, 1].Text));
                                kol_count_di = kol_count_di + 1;
                            }
                        }
                        if (Rus == true)
                            MessageBox.Show("Экспериментальные данные успешно загружены!", "Сообщение!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        if (Eng == true)
                            MessageBox.Show("Experimental data are imported successfully!", "Message!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        app.Quit();                        
                    }
                    catch
                    {
                        if (Rus == true)
                            MessageBox.Show("Невозможно загрузить выбранный файл", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        if (Eng == true)
                            MessageBox.Show("It is impossible to import the selected file", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
            }
            SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle, -1, -1);

        }

            private void Form1_FormClosing_1(object sender, FormClosingEventArgs e)
            {
            if (formclose == false)
            {
                if (loupe1 == true)
                {
                    if (Rus == true)
                        MessageBox.Show("Сначала закройте лупу!", "Сообщение!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (Eng == true)
                        MessageBox.Show("Close the magnifier first!", "Message!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (Rus == true)
                {
                    DialogResult result = MessageBox.Show("Уже уходите?", "Уведомление", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        
                        return;
                    }
                    if (result == DialogResult.No)
                    {
                        e.Cancel = true;
                    }
                }
                if (Eng == true)
                {
                    exit_DEAM = true;
                    f5 = new Form5();
                    f5.ShowDialog();
                    if (Program.f5.exite == true)
                    {
                        Program.f5.exite = false;
                        e.Cancel = true;
                    }
                }
            }
            }          

            







        }
    } 
