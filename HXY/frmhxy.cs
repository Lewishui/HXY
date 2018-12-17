using QM.Buiness;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;

namespace HXY
{
    public partial class frmhxy : Form
    {
        public frmhxy()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            List<string> newlisttime = new List<string>();

            if (this.checkedListBox1.CheckedItems.Count > 0)
            {

                foreach (string status in this.checkedListBox1.CheckedItems)
                    newlisttime.Add(status);
            }

            clsAllnew BusinessHelp = new clsAllnew();
            BusinessHelp.run_type = 1;

           BusinessHelp.ReadGeckoWEN(newlisttime);

            BusinessHelp.run_type = 2;
            BusinessHelp.ReadGeckoWEN(newlisttime);
            // BusinessHelp.ReadWEN( );



        }
    }
}
