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
    public partial class Form5 : Form
    {
        public bool exite = false;
        public bool exite_f5 = false;
        public Form5()
            
        {
            Program.f5 = this;
            InitializeComponent();
        }

        private void Form5_Load(object sender, EventArgs e)
        {
            if (Program.f1.new_distr == true || Program.f1.new_distr_f5==true)
            {
                label1.Text = "Do you really want to create a new distribution?";
                this.Text = "Notification";
                this.Width = 392;
                this.Height = 162;
                button1.Location = new System.Drawing.Point(186, 88);
                button2.Location= new System.Drawing.Point(277, 88);
                pictureBox1.Image = Properties.Resources.help_support_question_mark_512;
            }
            if (Program.f1.exit_DEAM==true)
            {
                label1.Text = "Do you really want to exit?";
                this.Text = "Notification";
                this.Width = 292;
                this.Height = 162;
                button1.Location = new System.Drawing.Point(86, 88);
                button2.Location = new System.Drawing.Point(177, 88);
                pictureBox1.Image = Properties.Resources.help_support_question_mark_512;
            }
            if (Program.f1.f3_open==true && Program.f3.exit_Log_Norm == true)
            {
                label1.Text = "Do you really want to exit?";
                this.Text = "Notification";
                this.Width = 292;
                this.Height = 162;
                button1.Location = new System.Drawing.Point(86, 88);
                button2.Location = new System.Drawing.Point(177, 88);
                pictureBox1.Image = Properties.Resources.help_support_question_mark_512;
            }
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Program.f1.new_distr==true)
            {
                for (int i = Program.f1.I; i > 3; i--)
                {
                    Program.f1.chart1.Series[Program.f1.I].Points.Clear();
                    Program.f1.chart1.Series.Remove(Program.f1.chart1.Series[Program.f1.I]);
                    Program.f1.dataGridView1.Rows.Remove(Program.f1.dataGridView1.Rows[Program.f1.num_line]);
                    Program.f1.dataGridView1.Rows.Remove(Program.f1.dataGridView1.Rows[Program.f1.num_line - 1]);
                    Program.f1.num_line = Program.f1.num_line - 2;
                    Program.f1.I = Program.f1.I - 1;
                }
                for (int i = 0; i < Program.f1.I_dGV2; i++)
                    Program.f1.dataGridView2.Rows[i].Cells[1].Value = string.Empty;
                Program.f1.I_dGV2 = 0;
                Program.f1.I_dGV3 = 0;
                Program.f1.kol_count_di = 0;
                //GO.Constants.kol_Di = 0;
                Program.f1.Di = 0;
                Program.f1.kol_meas_fault = -1;
                Program.f1.number_N_meas_fault = 0;
                Program.f1.N_meas_faultt.Insert(0, 0);
                Program.f1.meas_faultt.Insert(0, 0);
                Program.f1.textBox1.Text = null;
                Program.f1.textBox2.Text = null;
                Program.f1.textBox3.Text = null;
                Program.f1.textBox4.Text = null;
                Program.f1.textBox7.Text = null;
                Program.f1.textBox5.Text = "pix";
                Program.f1.textBox6.Text = "pix";
                Program.f1.textBox9.Text = String.Empty;
                Program.f1.textBox10.Text = String.Empty;
                Program.f1.panel1.BackgroundImage = null;
                Program.f1.new_distr = false;
                Program.f1.flag1 = false;
                Program.f1.flag = false;
                
                if (Program.f1.loupe1 == true)
                {
                    Program.f1.f2.Close();
                    Program.f1.loupe1 = false;
                    //Program.f1.button1.BackColor = System.Drawing.SystemColors.Control;
                }
                else
                    Program.f1.chart1.Series[3].Points.Clear();
                Program.f1.clickCountLoupe2x = 0;
                Program.f1.clickCountLoupe4x = 0;
                Program.f1.clickCountLoupe8x = 0;
                Program.f1.clickCountLoupe16x = 0;
                if (Program.f1.loupe)
                {
                    Program.f1.button12.Visible = false;
                    Program.f1.button13.Visible = false;
                    Program.f1.button14.Visible = false;
                    Program.f1.button15.Visible = false;
                }
                Program.f1.loupe = false;                
                this.Close();
            }
            if (Program.f1.exit_DEAM==true)
            {               
                this.Close();                
            }

            if (Program.f1.f3_open==true && Program.f1.new_distr_f5==true)
            {
                if (Program.f3.size_H == Program.f3.size_W)
                    Program.f3.chart2.Width = Program.f3.size_W - Program.f3.size_H1;
                else
                {
                    Program.f3.chart2.Width = Program.f3.size_H - Program.f3.size_H1;
                    Program.f3.chart2.Height = Program.f3.size_H - Program.f3.size_H1;
                }
                Program.f3.chart2.Location = new System.Drawing.Point(Program.f3.Width - Program.f3.chart2.Width - Program.f3.L, Program.f3.size_H1 - Program.f3.L);
                Program.f3.dataGridView1.Rows.Clear();
                Program.f3.dataGridView2.Rows.Clear();
                Program.f3.dataGridView3.Rows.Clear();
                Program.f3.dataGridView4.Rows.Clear();
                Program.f3.dataGridView5.Rows.Clear();
                Program.f3.X_max = 0;
                Program.f3.Sum = 0;
                Program.f3.chart2.Series[0].Points.Clear();
                Program.f3.chart2.Series[1].Points.Clear();
                Program.f3.chart2.Series[2].Points.Clear();
                Program.f3.chart2.Series[3].Points.Clear();
                Program.f3.chart2.Series[4].Points.Clear();
                Program.f3.chart2.Series[0].IsVisibleInLegend = false;
                Program.f3.chart2.Series[1].IsVisibleInLegend = false;
                Program.f3.chart2.Series[2].IsVisibleInLegend = false;
                Program.f3.chart2.Series[3].IsVisibleInLegend = false;
                Program.f3.chart2.Series[4].IsVisibleInLegend = false;
                Program.f3.groupBox2.Visible = false;
                Program.f3.groupBox3.Visible = false;
                Program.f3.richTextBox1.Clear();
                Program.f3.textBox3.Clear();
                Program.f3.textBox4.Clear();
                Program.f3.textBox5.Clear();
                Program.f3.pictureBox2.Visible = false;
                Program.f3.pictureBox3.Visible = false;
                Program.f3.pictureBox4.Visible = false;
                Program.f3.textBox11.Visible = false;
                Program.f3.textBox12.Visible = false;
                Program.f3.label5.Visible = false;
                Program.f3.label9.Visible = false;
                Program.f3.Part = false;
                Program.f3.CSR = false;
                Program.f3.CSR1 = false;
                Program.f3.CSR2 = false;
                Program.f3.textBox15.Visible = false;
                Program.f3.textBox16.Visible = false;
                Program.f3.pictureBox7.Visible = false;
                Program.f3.pictureBox8.Visible = false;
                Program.f3.label12.Visible = false;
                Program.f3.label13.Visible = false;
                Program.f3.textBox17.Visible = false;
                Program.f3.textBox18.Visible = false;
                Program.f3.pictureBox9.Visible = false;
                Program.f3.pictureBox10.Visible = false;
                Program.f3.label14.Visible = false;
                Program.f3.label15.Visible = false;
                this.Close();
            }

            if (Program.f1.f3_open == true && Program.f3.exit_Log_Norm == true)
                this.Close();
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (Program.f1.new_distr == true)
            {
                Program.f1.new_distr = false;
                this.Close();
            }

            if (Program.f1.exit_DEAM == true)
            {
                Program.f1.exit_DEAM = false;               
                exite = true;
                this.Close();
            }

            if (Program.f1.f3_open == true && Program.f1.new_distr_f5 == true)
            {
                Program.f1.new_distr_f5 = false;
                this.Close();
            }

            if (Program.f1.f3_open == true && Program.f3.exit_Log_Norm == true)
            {
                Program.f3.exit_Log_Norm = false;
                exite_f5 = true;
                this.Close();
            }
        }
    }
}
