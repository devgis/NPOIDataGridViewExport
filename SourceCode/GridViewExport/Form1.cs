using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GridViewExport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //需要设置的参数
        const string colID = "CID";
        const string colName = "CName";
        const string colSpec = "CSpec";
        const string colItems = "CItems";

        const int maxCount = 100;//模拟最多的检测项目数量

        Random rd = new Random();
        //添加测试数据
        private void AddTestData()
        {
            //模拟100行数据
            for (int i = 0; i < 100; i++)
            {
                int index = dataGridView1.Rows.Add();
                dataGridView1.Rows[index].Cells[0].Value = "编号" + i;
                dataGridView1.Rows[index].Cells[1].Value = "名称" + i;
                dataGridView1.Rows[index].Cells[2].Value = "规格" + i;
                string sitems = "";
                int itemcount = rd.Next(1, maxCount);
                for (int j = 0; j < itemcount; j++)
                {
                    if (j == itemcount - 1)
                    {
                        sitems += ("C" + j);
                    }
                    else
                    {
                        sitems += ("C" + j+",");
                    }
                }
                    dataGridView1.Rows[index].Cells[3].Value = sitems;

            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            AddTestData();
        }

        public class ExportItem
        {
            public string ID
            {
                get;
                set;
            }

            public string Name
            {
                get;
                set;
            }

            public string Spec
            {
                get;
                set;
            }

            public List<string> Items
            {
                get;
                set;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            List<ExportItem> list = new List<ExportItem>();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                ExportItem intem = new ExportItem();
                intem.ID = row.Cells[colID].Value == null ? "" : row.Cells[colID].Value.ToString();
                if (string.IsNullOrEmpty(intem.ID))
                {
                    continue;
                }
                intem.Name = row.Cells[colName].Value == null ? "" : row.Cells[colName].Value.ToString();
                intem.Spec = row.Cells[colSpec].Value == null ? "" : row.Cells[colSpec].Value.ToString();
                string[] sarr = (row.Cells[colItems].Value == null ? "" : row.Cells[colItems].Value.ToString()).Split(',');
                intem.Items = new List<string>();
                foreach (string s in sarr)
                {
                    intem.Items.Add(s);
                }
                list.Add(intem);
            }

            if (list.Count > 0)
            {
                ExportExcel.GridToExcel("导出结果", list);
            }
            else
            {
                MessageBox.Show("无可用数据！");
            }
        }
    }
}
