using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenXmlLab
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

		string docLocation;
		CustomProp document;

		private void button1_Click(object sender, EventArgs e)
		{
			openFileDialog1.ShowDialog();
			docLocation = openFileDialog1.FileName;
			if (docLocation != null)
			{
				document = new CustomProp(docLocation);
				fillListView(document.GetCustomProperities());
			}
			button4.Enabled = true;
		}

		private void fillListView(List<CustomPropertyEntity> docProps)
		{
			if (docProps == null)
			{
				MessageBox.Show("This Document Does not have custom Properties!");
				return;
			}
			listView1.Items.Clear();
			foreach (CustomPropertyEntity cProp in docProps)
			{
				addListViewItem(cProp);
			}
		}

		private void addListViewItem(CustomPropertyEntity cProp)
		{
			ListViewItem item = new ListViewItem(cProp.Name);
			item.SubItems.Add(cProp.Value);
			int index = ListViewContains(cProp.Name);
			if (index>=0)
			{
				listView1.Items.RemoveAt(index);
				listView1.Items.Insert(index, item);
			}
			else
			{
				listView1.Items.Add(item);
			}
			
		}

		private int ListViewContains(string itemText)
		{
			foreach (ListViewItem item in listView1.Items)
			{
				if (String.Compare(item.Text ,itemText,true)==0)
				{
					return item.Index;
				}
			}
			return -1;
		}

		private void button2_Click(object sender, EventArgs e)
		{
			listView1.Items.Clear();
		}

		private void button3_Click(object sender, EventArgs e)
		{
			if ((textBox1.Text==String.Empty) || (textBox2.Text == String.Empty))
			{
				MessageBox.Show("Property Name of Value is missing!");
				return;
			}
			addListViewItem(new CustomPropertyEntity(textBox1.Text, textBox2.Text));
			textBox1.Text = textBox2.Text = "";
		}

		private void button4_Click(object sender, EventArgs e)
		{
			foreach (ListViewItem item in listView1.Items)
			{
				CustomPropertyEntity cProp = new CustomPropertyEntity(item.Text, item.SubItems[1].Text);
				if (document != null)
				{
					document.SetCustomProperty(cProp);
				}
			}
		}

		private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
		{
			textBox1.Text = listView1.SelectedItems[0].Text;
			textBox2.Text = listView1.SelectedItems[0].SubItems[1].Text;
		}
	}
}
