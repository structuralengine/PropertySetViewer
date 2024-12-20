using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PropertySetViewer
{
    public partial class PropertySetForm : Form
    {
        public PropertySetForm(List<string> dataList)
        {
            InitializeComponent();

            // データをリストボックスに表示
            foreach (var data in dataList)
            {
                listBoxData.Items.Add(data);
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close(); // フォームを閉じる
        }
    }
}
