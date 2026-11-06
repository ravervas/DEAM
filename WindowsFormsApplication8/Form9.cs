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
    public partial class Form9 : Form
    {        

        public Form9()
        {
            Program.f9 = this;
            InitializeComponent();
        }

        private void Form9_Load(object sender, EventArgs e)
        {
            Program.f1.n = 1;
            Program.f1.n_const = 1;
            if (Program.f1.Eng)
            {
                this.Text = Program.f1.Calibrate;
                label1.Text = "Measure the feature of known length and enter the appropriate value";
                label2.Text = "Known length:";
                label2.Location = new Point(Program.f9.label2.Location.X + 10, Program.f9.label2.Location.Y);
                label3.Text = "Units:";
                label3.Location = new Point(Program.f9.label3.Location.X + 60, Program.f9.label3.Location.Y);
                comboBox1.Items.Clear();
                string[] mas = { "nm", "mkm", "mm", "pic" };
                comboBox1.Items.AddRange(mas);
                comboBox1.Text = "pic";               
                button2.Text = "Cancel";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Program.f1.N_calImage==1)
            {
                if (textBox1.Text==string.Empty || textBox1.Text==0.ToString())
                {
                    if (Program.f1.Rus == true)
                    {
                        MessageBox.Show("Не задана длина шкалы", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    if (Program.f1.Eng == true)
                    {
                        MessageBox.Show("The scale length is not entered", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }
                Program.f1.n = double.Parse(textBox1.Text);
                Program.f1.n_const = Program.f1.dii[Program.f1.kol_count_di - 1];
                Program.f1.n_s = double.Parse(textBox1.Text);
                Program.f1.n_const_s = Program.f1.dii[Program.f1.kol_count_di - 1];
                Program.f1.dii.Insert(Program.f1.kol_count_di - 1, Program.f1.n);
                //textBox1.Text = Math.Round(Program.f1.n, 2).ToString();
                Program.f1.measuring_fault = Math.Round(Program.f1.n * Program.f1.dpi / Program.f1.n_const, 2);
                Program.f1.textBox3.Text = Program.f1.measuring_fault.ToString();
                //Program.f1.button8_Click(sender, e);
                //Program.f1.kol_meas_fault = Program.f1.kol_meas_fault + 1;
                if (Program.f1.kol_meas_fault >= 0)
                {
                    
                    Program.f1.meas_faultt.Insert(Program.f1.kol_meas_fault, Program.f1.measuring_fault);
                }

                    Program.f1.Units = comboBox1.Text;               
                    Program.f1.textBox5.Text = Program.f1.Units;
                    Program.f1.textBox6.Text = Program.f1.Units;
                if (Program.f1.textBox9.Text!=String.Empty)
                {
                    Program.f1.textBox9.Text= Program.f1.Units;
                    Program.f1.textBox10.Text = Program.f1.Units;
                }

                if (comboBox1.Text=="пик")
                {
                    Program.f1.label2.Text = "D, пик";
                    Program.f1.label5.Text = "D, пик";
                }
                if (comboBox1.Text == "pic")
                {
                    Program.f1.label2.Text = "D, pic";
                    Program.f1.label5.Text = "D, pic";
                }
                if (comboBox1.Text == "нм")
                {
                    Program.f1.label2.Text = "D, нм";
                    Program.f1.label5.Text = "D, нм";
                }
                if (comboBox1.Text == "nm")
                {
                    Program.f1.label2.Text = "D, nm";
                    Program.f1.label5.Text = "D, nm";
                }
                if (comboBox1.Text == "мкм")
                {
                    Program.f1.label2.Text = "D, мкм";
                    Program.f1.label5.Text = "D, мкм";
                }
                if (comboBox1.Text == "mkm")
                {
                    Program.f1.label2.Text = "D, mkm";
                    Program.f1.label5.Text = "D, mkm";
                }
                if (comboBox1.Text == "мм")
                {
                    Program.f1.label2.Text = "D, мм";
                    Program.f1.label5.Text = "D, мм";
                }
                if (comboBox1.Text == "mm")
                {
                    Program.f1.label2.Text = "D, mm";
                    Program.f1.label5.Text = "D, mm";
                }


                this.Close();
            }
            else
            {
                if (Program.f1.Rus == true)
                {
                    MessageBox.Show("Измерьте длину шкалы!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (Program.f1.Eng == true)
                {
                    MessageBox.Show("You have to measure the scale length!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }

        }

        public void Form9_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (Program.f1.N_calImage == 1)
                Program.f1.button8_Click(sender, e);
            Program.f1.N_calImage = 0;
            Program.f1.calibrateImage = false;
            Program.f1.clickCountCalibrationImage = 0;
            Program.f1.button6.BackColor = System.Drawing.SystemColors.Control;
            Program.f1.n = Program.f1.n_s;
            Program.f1.n_const = Program.f1.n_const_s;
        }

        private void button2_Click(object sender, EventArgs e)
        {                    
            this.Close();
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
        
            if (!(Char.IsDigit(e.KeyChar)) && (e.KeyChar != (char)Keys.Back) && !(e.KeyChar == ','))
            {
                e.Handled = true;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            button3.Select();
        }
    }
}
