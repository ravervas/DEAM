using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication8
{
    public partial class Form12 : Form
    {

        private int click_toolStr_button1=0,
                    click_toolStr_button2=0, I3=1,
                    clickRaspred1=0,
                    clickRaspred2=0,
                    clickRaspredCommon=0,
                    clickCSR1=0, clickCSR2=0;
        public int I = 12, click_Ln = 0;
        private double sumNiDi, sumNi, Xcp, x, y;
        public string Ln = "Логарифмические координаты";
        private int MouseUp_X, MouseUp_Y;
        private bool tS1 = false, tSCB=false, tSB1=false;
        public double r;

        public Form12()
        {
            Program.f12 = this;
            InitializeComponent();
        }

        private void Form12_Load(object sender, EventArgs e)
        {

        }

        private void Form12_FormClosing(object sender, FormClosingEventArgs e)
        {
            Program.f3.open_f12 = false;
            Program.f3.tabControl3.Visible = false;
            Program.f3.Ln30 = false;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            ++click_toolStr_button1;
            switch (click_toolStr_button1)
            {
                case 1:
                    
                    toolStripButton1.Checked = false;
                    chart2.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
                    chart2.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
                    break;
                case 2:
                    
                    toolStripButton1.Checked = true;
                    chart2.ChartAreas[0].AxisX.MajorGrid.Enabled = true;
                    chart2.ChartAreas[0].AxisY.MajorGrid.Enabled = true;
                    click_toolStr_button1 = 0;
                    break;
            }
        }

        private void toolStripButton1_MouseMove(object sender, MouseEventArgs e)
        {
            toolStripButton1.ToolTipText = Program.f1.Grid;
        }

        private void toolStripButton2_MouseMove(object sender, MouseEventArgs e)
        {
            toolStripButton2.ToolTipText = Program.f1.Legend;
        }

        private void toolStripComboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            I = Convert.ToInt32(toolStripComboBox4.Text);
            chart2.ChartAreas[0].AxisX.TitleFont = new Font("Times New Roman", I, FontStyle.Bold);
            chart2.ChartAreas[0].AxisY.TitleFont = new Font("Times New Roman", I, FontStyle.Bold);
            chart2.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Times New Roman", I, System.Drawing.FontStyle.Regular);
            chart2.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Times New Roman", I, System.Drawing.FontStyle.Regular);
            chart2.Legends[0].Font = new System.Drawing.Font("Times New Roman", I, System.Drawing.FontStyle.Regular);
        }

        private void toolStripComboBox4_MouseMove(object sender, MouseEventArgs e)
        {
            toolStripComboBox4.ToolTipText = Program.f1.font;
        }

        private void toolStripDropDownButton1_MouseMove(object sender, MouseEventArgs e)
        {
            toolStripDropDownButton1.ToolTipText = Program.f1.Line_thickness;
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

        private void toolStripSplitButton2_MouseMove(object sender, MouseEventArgs e)
        {
            toolStripSplitButton2.ToolTipText = Program.f1.Approx;
        }

        private void Form12_Click(object sender, EventArgs e)
        {
            if (!tS1 && !tSCB)
            {
                // нажимам левую кнопку
                mouse_event((int)(MouseEventFlags.LEFTDOWN), MouseUp_X, MouseUp_Y, 0, 0);
                // отпускаем левую кнопку
                mouse_event((int)(MouseEventFlags.LEFTUP), MouseUp_X, MouseUp_Y, 0, 0);
            }
            if (tS1)
                // нажимам левую кнопку
                mouse_event((int)(MouseEventFlags.LEFTDOWN), MouseUp_X, MouseUp_Y, 0, 0);
            

        }

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

        

        private void toolStrip1_MouseEnter(object sender, EventArgs e)
        {
            tS1 = true;
        }

        private void toolStrip1_MouseLeave(object sender, EventArgs e)
        {
            tS1 = false;
            
        }

       

        private void toolStrip1_Click(object sender, EventArgs e)
        {
           // mouse_event((int)(MouseEventFlags.LEFTDOWN), MouseUp_X, MouseUp_Y, 0, 0);
        }

        private void toolStripButton1_MouseLeave(object sender, EventArgs e)
        {
            tSB1 = false;
        }

        private void toolStripButton1_MouseEnter(object sender, EventArgs e)
        {
            tSB1 = true;
        }

        private void toolStripSplitButton2_MouseLeave(object sender, EventArgs e)
        {
            tS1 = false;
        }

        [DllImport("user32.dll")]
        static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);
    

        private void Form12_Activated(object sender, EventArgs e)
        {
            Program.f3.tabControl1.SelectedTab = Program.f3.tabControl1.TabPages["tabPage2"];
            Program.f3.tabControl1.Refresh();
            MouseUp_X = Cursor.Position.X;
            MouseUp_Y = Cursor.Position.Y;

            Form12_Click(sender, e);
            
            
        }

        private void toolStripButton4_MouseMove(object sender, MouseEventArgs e)
        {
            toolStripButton4.ToolTipText = Ln;
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            ++click_Ln;
            switch (click_Ln)
            {
                case 1:
                    try
                    {
                        toolStripButton4.Checked = true;
                        chart2.Series[0].Points.Clear();
                        chart2.Series[1].Points.Clear();
                        chart2.Series[2].Points.Clear();
                        chart2.Series[5].Points.Clear();
                        Program.f3.tabControl3.Visible = false;
                        //chart2.ChartAreas[0].AxisX.LogarithmBase = Math.Exp(1);
                        // chart2.ChartAreas[0].AxisX.IsLogarithmic = true;

                        for (int i = 0; i < Program.f3.N30; i++)
                        {
                            x = Convert.ToDouble(Program.f3.dataGridView2.Rows[i].Cells[3].Value);
                            y = Convert.ToDouble(Program.f3.dataGridView2.Rows[i].Cells[2].Value);
                            chart2.Series[0].Points.AddXY(x, y);
                            Program.f3.X_max30 = Convert.ToDouble(Program.f3.dataGridView2.Rows[i].Cells[3].Value);

                        }
                        r = Convert.ToDouble(Program.f3.dataGridView2.Rows[Program.f3.N30-1].Cells[3].Value) - Convert.ToDouble(Program.f3.dataGridView2.Rows[Program.f3.N30-2].Cells[3].Value);
                        Program.f3.X_max30 = Program.f3.X_max30 + r;
                        chart2.ChartAreas[0].AxisY.Minimum = 0;
                        chart2.ChartAreas[0].AxisX.Maximum = Program.f3.X_max30;
                        chart2.ChartAreas[0].AxisX.Minimum = Convert.ToDouble(Program.f3.dataGridView2.Rows[0].Cells[3].Value) - r;
                        chart2.ChartAreas[0].AxisX.Title = "Ln(D)";
                        chart2.ChartAreas[0].AxisX.Interval = 0;
                        Program.f3.Ln30 = true;
                        Program.f3.analyt_func_30 = false;

                        break;
                    }
                    catch
                    {
                        return;
                    }

                case 2:
                    {
                        toolStripButton4.Checked = false;
                        chart2.Series[0].Points.Clear();
                        chart2.Series[1].Points.Clear();
                        Program.f3.tabControl3.Visible = false;
                        for (int i = 0; i < Program.f3.N30; i++)
                        {
                            x = Convert.ToDouble(Program.f3.dataGridView2.Rows[i].Cells[0].Value);
                            y = Convert.ToDouble(Program.f3.dataGridView2.Rows[i].Cells[2].Value);
                            chart2.Series[0].Points.AddXY(x, y);
                            Program.f3.X_max30 = Convert.ToDouble(Program.f3.dataGridView2.Rows[i].Cells[0].Value);
                        }
                        Program.f3.X_max30 = Program.f3.X_max30 + Program.f3.r;
                        chart2.ChartAreas[0].AxisY.Minimum = 0;
                        chart2.ChartAreas[0].AxisX.Maximum = Program.f3.X_max30;
                        if (Program.f3.Part30 == true && Program.f3.CSR30 == true)
                            chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), нм";
                        else
                            chart2.ChartAreas[0].AxisX.Title = "Размер частиц, нм";

                        if (Program.f3.Part30 == true && Program.f3.CSR30 == true)
                            chart2.ChartAreas[0].AxisY.Title = "Доля частиц (ОКР), %";
                        else
                            chart2.ChartAreas[0].AxisY.Title = "Доля частиц, %";
                        if (Program.f1.Eng == true)
                            chart2.ChartAreas[0].AxisY.Title = "Frequency, %";
                        if (Program.f3.toolStripComboBox2.SelectedIndex == 0)
                        {
                            if (Program.f1.Rus == true)
                            {
                                if (Program.f3.Part30 == true && Program.f3.CSR30 == true)
                                    chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), мм (нм)";
                                else
                                    chart2.ChartAreas[0].AxisX.Title = "Размер частиц, мм";
                                Program.f3.label33.Text = "мм";
                                Program.f3.label34.Text = "мм";
                                Program.f3.label36.Text = "мм";
                            }
                            if (Program.f1.Eng == true)
                            {
                                if (Program.f3.Part30 == true && Program.f3.CSR30 == true)
                                    chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), mm (nm)";
                                else
                                    chart2.ChartAreas[0].AxisX.Title = "Particle size, mm";
                                Program.f3.label33.Text = "mm";
                                Program.f3.label34.Text = "mm";
                                Program.f3.label36.Text = "mm";
                            }
                        }
                        if (Program.f3.toolStripComboBox2.SelectedIndex == 1)
                        {
                            if (Program.f1.Rus == true)
                            {
                                if (Program.f3.Part30 == true && Program.f3.CSR30 == true)
                                    chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), мкм (нм)";
                                else
                                    chart2.ChartAreas[0].AxisX.Title = "Размер частиц, мкм";
                                Program.f3.label33.Text = "мкм";
                                Program.f3.label34.Text = "мкм";
                                Program.f3.label36.Text = "мкм";
                            }
                            if (Program.f1.Eng == true)
                            {
                                if (Program.f3.Part30 == true && Program.f3.CSR30 == true)
                                    chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), mkm (nm)";
                                else
                                    chart2.ChartAreas[0].AxisX.Title = "Particle size, mkm";
                                Program.f3.label33.Text = "mkm";
                                Program.f3.label34.Text = "mkm";
                                Program.f3.label36.Text = "mkm";
                            }
                        }
                        if (Program.f3.toolStripComboBox2.SelectedIndex == 2)
                        {
                            if (Program.f1.Rus == true)
                            {
                                if (Program.f3.Part30 == true && Program.f3.CSR30 == true)
                                    chart2.ChartAreas[0].AxisX.Title = "Размер частиц (ОКР), нм";
                                else
                                    chart2.ChartAreas[0].AxisX.Title = "Размер частиц, нм";
                                Program.f3.label33.Text = "нм";
                                Program.f3.label34.Text = "нм";
                                Program.f3.label36.Text = "нм";
                            }
                            if (Program.f1.Eng == true)
                            {
                                if (Program.f3.Part30 == true && Program.f3.CSR30 == true)
                                    chart2.ChartAreas[0].AxisX.Title = "Particle size (CSR), nm";
                                else
                                    chart2.ChartAreas[0].AxisX.Title = "Particle size, nm";
                                Program.f3.label33.Text = "nm";
                                Program.f3.label34.Text = "nm";
                                Program.f3.label36.Text = "nm";
                            }


                        }


                        if (Program.f3.dataGridView1.Rows.Count > 2)
                        {
                            Program.f3.X_interval30 = Convert.ToDouble(Program.f3.dataGridView2.Rows[1].Cells[0].Value) - Convert.ToDouble(Program.f3.dataGridView2.Rows[0].Cells[0].Value);
                            Program.f3.X_min30 = Convert.ToDouble(Program.f3.dataGridView2.Rows[0].Cells[0].Value) - Program.f3.X_interval30;
                            chart2.ChartAreas[0].AxisX.Minimum = Program.f3.X_min30;
                            chart2.ChartAreas[0].AxisX.Interval = Program.f3.X_interval30;
                            if (Program.f3.X_min30 <= 0)
                            {
                                chart2.ChartAreas[0].AxisX.IsStartedFromZero = true;
                                for (int i = 0; i < Program.f3.N30; i++)
                                    chart2.Series[0].Points[i].AxisLabel = (Program.f3.dataGridView2.Rows[i].Cells[0].Value).ToString();
                            }
                        }
                        click_Ln = 0;
                        Program.f3.Ln30 = false;
                        Program.f3.analyt_func_30 = false;
                    }
                    break;
            }
        }

        private void toolStripButton3_MouseMove(object sender, MouseEventArgs e)
        {
            toolStripButton3.ToolTipText = Program.f1.Average;
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            sumNiDi = 0;
            sumNi = 0;
            for (int i = 0; i < Program.f3.N30; i++)
            {
                sumNiDi = sumNiDi + Convert.ToDouble(Program.f3.dataGridView2.Rows[i].Cells[0].Value) * Convert.ToDouble(Program.f3.dataGridView2.Rows[i].Cells[1].Value);
                sumNi = sumNi + Convert.ToDouble(Program.f3.dataGridView2.Rows[i].Cells[1].Value);
            }
            Xcp = Math.Round(sumNiDi / sumNi, 2);
            if (Program.f1.Rus == true)
                MessageBox.Show("Математическое ожидание равно " + Xcp + " " + Program.f3.toolStripComboBox2.Text, "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            if (Program.f1.Eng == true)
                MessageBox.Show("Mathematical expectation is " + Xcp + " " + Program.f3.toolStripComboBox2.Text, "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            I3 = Convert.ToInt32(toolStripComboBox1.Text);
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
                    
                    toolStripButton2.Checked = false;
                    chart2.Legends[0].Enabled = false;
                    break;
                case 2:
                    
                    toolStripButton2.Checked = true;
                    chart2.Legends[0].Enabled = true;
                    click_toolStr_button2 = 0;
                    break;
            }
        }
    }
}
