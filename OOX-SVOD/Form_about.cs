﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OOX_SVOD
{
    public partial class Form_about : Form
    {
        public Form_about()
        {
            InitializeComponent();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var info = new ProcessStartInfo
            {
                FileName = linkLabel1.Text,
                UseShellExecute = true
            };
            Process.Start(info);
        }
    }
}
