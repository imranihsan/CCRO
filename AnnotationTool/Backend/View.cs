using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GraphRepresentation
{
    public partial class View : MetroFramework.Forms.MetroForm
    {
        public View()
        {
            InitializeComponent();
        }

        private void View_Load(object sender, EventArgs e)
        {
            textBox1.Text = Form1.cite;
            textBox2.Text = Form1.citing;
            textBox3.Text = Form1.citation;
            textBox4.Text = Form1.lematize;
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
