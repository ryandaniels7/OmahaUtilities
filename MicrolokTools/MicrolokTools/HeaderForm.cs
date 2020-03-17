using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MicrolokTools
{
    public partial class HeaderForm : Form
    {
        public HeaderForm()
        {
            InitializeComponent();
        }
        private void CancelButton_Click(object sender, EventArgs e)
        {
            Dispose();
            
        }
        public void OKButton_Click(object sender, EventArgs e)
        {
            Dispose();
        }
    }
}
