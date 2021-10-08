using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace week4
{
    public partial class Form1 : Form
    {
        RealEstateEntities1 context = new RealEstateEntities1();
        List<Flat> Flats;
        public Form1()
        {
            InitializeComponent();
            LoadData();
        }

        private void LoadData()
        {
            Flats = context.Flats.ToList();
        }

        /*private void Form1_Load(object sender, EventArgs e)
        {
        }*/
    }
}
