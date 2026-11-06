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
    public partial class Form8 : Form
    {
        public Form8()
        {
            Program.f8 = this;
            InitializeComponent();
        }

        private void Form8_Load(object sender, EventArgs e)
        {
            if (Program.f1.Eng)
                this.Text = "Zoom factor";
            if (Program.f1.l2x)
                button1_Click(sender, e);
            if (Program.f1.l4x)
                button2_Click(sender, e);
            if (Program.f1.l8x)
                button3_Click(sender, e);
            if (Program.f1.l16x)
                button4_Click(sender, e);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.BackColor = System.Drawing.SystemColors.Info;
            button2.BackColor = System.Drawing.SystemColors.Control;
            button3.BackColor = System.Drawing.SystemColors.Control;
            button4.BackColor = System.Drawing.SystemColors.Control;
            Program.f1.magnificationX = Program.f1.Lx / 4;
            Program.f1.magnificationY = Program.f1.Ly / 4;
            Program.f1.magnificX = Program.f1.Lx / 2;
            Program.f1.magnificY = Program.f1.Ly / 2;
            Program.f1.l2x = true;
            Program.f1.l4x = false;
            Program.f1.l8x = false;
            Program.f1.l16x = false;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            button1.BackColor = System.Drawing.SystemColors.Control;
            button2.BackColor = System.Drawing.SystemColors.Info;
            button3.BackColor = System.Drawing.SystemColors.Control;
            button4.BackColor = System.Drawing.SystemColors.Control;
            Program.f1.magnificationX = Program.f1.Lx / 8;
            Program.f1.magnificationY = Program.f1.Ly / 8;
            Program.f1.magnificX = Program.f1.Lx / 4;
            Program.f1.magnificY = Program.f1.Ly / 4;
            Program.f1.l2x = false;
            Program.f1.l4x = true; ;
            Program.f1.l8x = false;
            Program.f1.l16x = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            button1.BackColor = System.Drawing.SystemColors.Control;
            button2.BackColor = System.Drawing.SystemColors.Control;
            button3.BackColor = System.Drawing.SystemColors.Info;
            button4.BackColor = System.Drawing.SystemColors.Control;
            Program.f1.magnificationX = Program.f1.Lx / 16;
            Program.f1.magnificationY = Program.f1.Ly / 16;
            Program.f1.magnificX = Program.f1.Lx / 8;
            Program.f1.magnificY = Program.f1.Ly / 8;
            Program.f1.l2x = false;
            Program.f1.l4x = false;
            Program.f1.l8x = true;
            Program.f1.l16x = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            button1.BackColor = System.Drawing.SystemColors.Control;
            button2.BackColor = System.Drawing.SystemColors.Control;
            button3.BackColor = System.Drawing.SystemColors.Control;
            button4.BackColor = System.Drawing.SystemColors.Info;
            Program.f1.magnificationX = Program.f1.Lx / 32;
            Program.f1.magnificationY = Program.f1.Ly / 32;
            Program.f1.magnificX = Program.f1.Lx / 16;
            Program.f1.magnificY = Program.f1.Ly / 16;
            Program.f1.l2x = false;
            Program.f1.l4x = false;
            Program.f1.l8x = false;
            Program.f1.l16x = true;
        }
    }
}
