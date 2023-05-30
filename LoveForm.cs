using System;
using System.Windows.Forms;

namespace Students
{
    public partial class LoveForm : Form
    {
        public LoveForm()
        {
            InitializeComponent();
        }
        # region var
        # endregion

        private void LoveForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            pictureBox1.Visible = true;
            pictureBox2.Visible = true;
            pictureBox3.Visible = false;
            pictureBox4.Visible = false;
            pictureBox5.Visible = false;
            timer.Start();
            MessageBox.Show("Приятного дня моё Любимое Солнышко! Целую, я с тобой... ", "Люблю тебя...",
                             MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            pictureBox3.Dock = DockStyle.Fill; 
            pictureBox4.Visible = false;
            pictureBox5.Visible = false; 
            timer.Start();
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            pictureBox4.Dock = DockStyle.Fill; 
            pictureBox3.Visible = false;
            pictureBox5.Visible = false; 
            timer.Start();
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            pictureBox5.Dock = DockStyle.Fill; 
            pictureBox3.Visible = false;
            pictureBox4.Visible = false; 
            timer.Start();
        }

        private void timer_Tick(object sender, EventArgs e)
        {
            pictureBox3.Visible = true;
            pictureBox4.Visible = true;
            pictureBox5.Visible = true; 
            pictureBox3.Dock = DockStyle.None;
            pictureBox4.Dock = DockStyle.None;
            pictureBox5.Dock = DockStyle.None;
            timer.Stop();
        }

        

        

       

      

            
     

     

    

    }
}
