using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Office.Interop;
using System.Runtime.InteropServices;
using System.IO;

namespace WindowsFormsApp1
{

    public partial class Form1 : Form
    {
        PeopleCollect workers = new PeopleCollect();
        ListViewItem dragItem = null;
        TreeNode dragNode = null;
        bool NodeMode = false;
        bool ItemMode = false;
        ListViewItem lvlItem;

        string redact;
        public Form1()
        {

            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //var doc=new XmlDocument();
            //doc.Load ("XMLFile1.xml");
            //XmlNode root = doc.DocumentElement;
            ////Console.WriteLine(root.Attributes.Item(0).Value);
            //foreach(XmlNode n in root.ChildNodes)
            //{
            //    treeView1.Text = n.Attributes.Item(0).Value;
            //    foreach (XmlNode c in n.ChildNodes)
            //    {
            //        treeView1.Text = c.Attributes.Item(1).Value;
            //    }
            //}

            var doc = new XmlDocument();
            doc.Load("XMLFile1.xml");
            XmlNode root = doc.DocumentElement;
            TreeNode rootNode = new TreeNode();
            rootNode.Text = root.Attributes.Item(0).Value;
            rootNode.ImageIndex = Convert.ToInt32(root.Attributes.Item(1).Value);
            rootNode.SelectedImageIndex = Convert.ToInt32(root.Attributes.Item(1).Value);
            rootNode.Tag = root.Attributes.Item(2).Value;
            treeView1.Nodes.Add(rootNode);

            if ((root.ChildNodes != null) && (root.ChildNodes.Count > 0))
            {
                RecursiveTreeBuilder(rootNode, root);
            }
            treeView1.ExpandAll();
            treeView1.SelectedNode = rootNode;

            var xmlWriter = new XmlTextWriter("XMLFile1.xml", null);
            xmlWriter.Formatting = Formatting.Indented;
            xmlWriter.IndentChar = '\t';
            xmlWriter.Indentation = 1;

            TreeNode parent =treeView1.Nodes[0];
            xmlWriter.WriteStartDocument();
            xmlWriter.WriteStartElement("rootNode");
            xmlWriter.WriteStartAttribute("text");
            xmlWriter.WriteString(parent.Text);
            xmlWriter.WriteEndAttribute();
            xmlWriter.WriteStartAttribute("vid");
            xmlWriter.WriteString(parent.ImageIndex.ToString());

            xmlWriter.WriteEndAttribute();
            xmlWriter.WriteStartAttribute("tag");
            xmlWriter.WriteString(parent.Tag.ToString());
            xmlWriter.WriteEndAttribute();
           
           
            

           


            if ((parent.Nodes!= null) && (parent.Nodes.Count > 0))
            {
                RecursiveXMLWriter(xmlWriter, parent);



            }
            xmlWriter.WriteEndElement();
            xmlWriter.Close();


        }
        private void forXMl()
        {

        }
        private void RecursiveXMLWriter(XmlTextWriter tn, TreeNode xn)
        {
            //throw new NotImplementedException();
            foreach (TreeNode c in xn.Nodes)
            {
                tn.WriteStartElement("node");
                tn.WriteStartAttribute("text");
                tn.WriteString(c.Text);
                tn.WriteEndAttribute();
                tn.WriteStartAttribute("vid");
                tn.WriteString(c.ImageIndex.ToString());
                tn.WriteEndAttribute();
                tn.WriteStartAttribute("tag");
                tn.WriteString(c.Tag.ToString());
                tn.WriteEndAttribute();
               
               
          

                if ((c.Nodes != null) && (c.Nodes.Count > 0))
                {
                    RecursiveXMLWriter(tn, c);

                }
                tn.WriteEndElement();


            }
        }

        private void RecursiveTreeBuilder(TreeNode tn, XmlNode xn)
        {
            //throw new NotImplementedException();
            foreach (XmlNode c in xn.ChildNodes)
            {
                TreeNode childNode = new TreeNode();
                childNode.Text = c.Attributes.Item(0).Value;
                childNode.ImageIndex = Convert.ToInt32(c.Attributes.Item(1).Value);
                childNode.SelectedImageIndex = Convert.ToInt32(c.Attributes.Item(1).Value);
                childNode.Tag = c.Attributes.Item(2).Value;
                tn.Nodes.Add(childNode);
                if ((c.ChildNodes != null) && (c.ChildNodes.Count > 0))
                {
                    RecursiveTreeBuilder(childNode, c);
                }
            }
        }

        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }

        private void splitContainer2_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void hScrollBar1_Scroll(object sender, ScrollEventArgs e)
        {

        }

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
          
            int Kol_workers = 0, max_salary = 0, min_salary = 0, midle_salary, kol_stager = 0, max_age = 0, min_age = 0, midle_age, kol_otdelB = 0, kol_otdelK = 0, kol_otdelP = 0;
            int tag = Int32.Parse(e.Node.Tag.ToString());
            string fillname;
            fillname = e.Node.Text;
            if (tag < 10)//filial
            {
                foreach (Pers p in workers)
                {
                    if (p.filial == tag)
                    {
                        max_salary = p.money;
                        min_salary = p.money;


                        break;

                    }
                }
            }
            else//otdel
            {
                foreach (Pers p in workers)
                {
                    if (p.code == tag)
                    {
                        max_salary = p.money;
                        min_salary = p.money;
                        max_age = p.age;
                        min_age = p.age;

                        break;
                    }
                }
            }

            listView1.Items.Clear();
            if (tag < 10)//filial
            {
                foreach (Pers p in workers)
                {
                    if (p.filial == tag)
                    {
                        ListViewItem item1 = new ListViewItem(p.fio.ToString());
                        // item1.SubItems.Add(p.fio.ToString());

                        item1.SubItems.Add(p.nomer.ToString());

                        item1.SubItems.Add(p.age.ToString());
                        item1.SubItems.Add(p.money.ToString());
                        if (p.stag == 1)
                        {
                            item1.SubItems.Add("Да");
                        }
                        else
                        {
                            item1.SubItems.Add("Нет");
                        }
                        item1.Tag = p.nomer.ToString();
                        item1.ImageIndex = 5;
                        listView1.Items.Add(item1);

                        Kol_workers++;
                        if (p.code == tag * 10 + 1)
                        {
                            kol_otdelB++;
                        }
                        else if (p.code == tag * 10 + 2)
                        {
                            kol_otdelK++;
                        }
                        else if (p.code == tag * 10 + 3)
                        {
                            kol_otdelP++;
                        }

                        if (p.stag == 1)
                        {
                            kol_stager++;
                        }
                        if (p.money > max_salary)
                        {
                            max_salary = p.money;
                        }
                        if (p.money < min_salary)
                        {
                            min_salary = p.money;
                        }
                    }
                }
                if (tag >= 1)
                {
                    midle_salary = (min_salary + max_salary) / 2;
                    label9.Text = "Информация по филиалу " + fillname;
                    label1.Text = "Количество сотрудников: " + Kol_workers.ToString();
                    label2.Text = "Количество сотрудников в Бухгалтерии: " + kol_otdelB.ToString();
                    label3.Text = "Количество сотрудников в Кадрах: " + kol_otdelK.ToString();
                    label4.Text = "Количество сотрудников в Производстве: " + kol_otdelP.ToString();
                    label5.Text = "Количество стажеров: " + kol_stager.ToString();
                    label6.Text = "Макс. зарплата: " + max_salary.ToString();
                    label7.Text = "Мин. зарплата: " + min_salary.ToString();
                    label8.Text = "Средняя зарплата: " + midle_salary.ToString();
                }
                else
                {
                    label9.Text = "";
                    label1.Text = "";
                    label2.Text = "";
                    label3.Text = "";
                    label4.Text = "";
                    label5.Text = "";
                    label6.Text = "";
                    label7.Text = "";
                    label8.Text = "";
                }
            }
            else
            {
                foreach (Pers p in workers)
                {
                    if (p.code == tag)
                    {
                        ListViewItem item1 = new ListViewItem(p.fio);
                        item1.SubItems.Add(p.nomer.ToString());
                        item1.SubItems.Add(p.age.ToString());
                        item1.SubItems.Add(p.money.ToString());
                        if (p.stag == 1)
                        {
                            item1.SubItems.Add("Да");
                        }
                        else
                        {
                            item1.SubItems.Add("Нет");
                        }
                        item1.Tag = p.nomer.ToString();
                        item1.ImageIndex = 5;
                        listView1.Items.Add(item1);


                        Kol_workers++;


                        if (p.stag == 1)
                        {
                            kol_stager++;
                        }
                        if (p.money > max_salary)
                        {
                            max_salary = p.money;
                        }
                        if (p.money < min_salary)
                        {
                            min_salary = p.money;
                        }

                        if (p.age > max_age)
                        {
                            max_age = p.age;
                        }
                        if (p.age < min_age)
                        {
                            min_age = p.age;
                        }
                    }
                }

                midle_salary = (min_salary + max_salary) / 2;
                midle_age = (max_age + min_age) / 2;
                label9.Text = "Информация по отделу " + fillname;
                label1.Text = "Количество сотрудников: " + Kol_workers.ToString();
                label2.Text = "Макс. возраст: " + max_age.ToString();
                label3.Text = "Мин. возраст: " + min_age.ToString();
                label4.Text = "Средний возраст: " + midle_age.ToString();
                label5.Text = "Количество стажеров: " + kol_stager.ToString();
                label6.Text = "Макс. зарплата: " + max_salary.ToString();
                label7.Text = "Мин. зарплата: " + min_salary.ToString();
                label8.Text = "Средняя зарплата: " + midle_salary.ToString();

            }
           


        }



        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void editToolStripMenuItem_Click(object sender, EventArgs e)
        {
            splitContainer2.Enabled = true;
            splitContainer2.Panel1.Show();
            editToolStripMenuItem.Checked = true;
            splitContainer2.Panel2.Show();
            splitContainer2.SplitterDistance = splitContainer2.Height / 5;
        }

        private void treeView1_MouseUp(object sender, MouseEventArgs e)
        {

            TreeNode targetNode = treeView1.GetNodeAt(e.X, e.Y);

            if (targetNode != null)
            {
                if (ItemMode)
                {
                    int tag = Int32.Parse(targetNode.Tag.ToString());
                    if (tag < 10)
                    {
                        ListViewItem newpos_Item = dragItem;
                        foreach (Pers p in workers)
                        {
                            if (newpos_Item.Tag.ToString() == p.nomer.ToString())
                            {
                                p.filial = tag;
                                int otdel = p.code % 10;
                                p.code = tag * 10 + otdel;
                            }
                        }
                    }
                    else
                    {
                        ListViewItem newpos_Item = dragItem;
                        foreach (Pers p in workers)
                        {
                            if (newpos_Item.Tag.ToString() == p.nomer.ToString())
                            {
                                p.filial = Int32.Parse(targetNode.Parent.Tag.ToString());
                                p.code = tag;
                            }
                        }
                    }

                    dragItem.Remove();
                    listView1.Refresh();
                }
                if (NodeMode)
                {
                    targetNode.Nodes.Add(dragNode.Clone() as TreeNode);
                    dragNode.Remove();
                    treeView1.Refresh();
                    treeView1.ExpandAll();
                }
            }
            CancelDrag();
        }

        private void treeView1_MouseDown(object sender, MouseEventArgs e)
        {
            dragNode = treeView1.GetNodeAt(e.X, e.Y);

        }

        private void treeView1_MouseMove(object sender, MouseEventArgs e)
        {
            if ((!NodeMode) && (dragNode != null))
            {
                treeView1.Cursor = AdvancedCursor.Create("stuff.cur");
                NodeMode = true;
            }
        }

        private void listView1_MouseDown(object sender, MouseEventArgs e)
        {
            dragItem = listView1.GetItemAt(e.X, e.Y);
        }

        private void listView1_MouseUp(object sender, MouseEventArgs e)
        {
            CancelDrag();

        }

        private void listView1_MouseMove(object sender, MouseEventArgs e)
        {
            if ((!ItemMode) && (dragItem != null))
            {
                listView1.Cursor = AdvancedCursor.Create("man.cur");
                ItemMode = true;
            }
        }
        private void CancelDrag()
        {
            treeView1.Cursor = Cursors.Hand;
            listView1.Cursor = Cursors.Default;
            dragItem = null;
            dragNode = null;
            ItemMode = false;
            NodeMode = false;
        }
        private void FormateExelToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {

        }

        private void saveFileDialog2_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void развернутьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode selectsNode = treeView1.SelectedNode;

            if (selectsNode != null)
            {
                selectsNode.Toggle();
            }
        }




        private void toolStripButton1_Click_1(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app1 = new Microsoft.Office.Interop.Excel.Application();
            app1.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook book1 = app1.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet otchet = book1.Worksheets[1];
            otchet.Cells[1, 1] = "Номер";
            otchet.Cells[1, 2] = "Coтрудники";
            for (int i = 1; i <= listView1.Items.Count; i++)
            {
                otchet.Cells[i + 1, 1] = i.ToString();
                otchet.Cells[i + 1, 2] = listView1.Items[i - 1].Text;

            }
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            splitContainer2.Panel1.Hide();
            splitContainer2.Panel2.Hide();
            splitContainer2.Enabled = false;
            editToolStripMenuItem.Checked = false;
            splitContainer2.SplitterDistance = splitContainer2.Height;
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
            TreeNode selectsNode = treeView1.SelectedNode;
            int tag = Int32.Parse(selectsNode.Tag.ToString());
            if (selectsNode != null && selectsNode.IsExpanded == true)
            {
                contextMenuStrip1.Items[0].Text = "Cвернуть";
                contextMenuStrip1.Items[0].Enabled = true;
            }
            else
            {
                contextMenuStrip1.Items[0].Text = "Развернуть";
                if (selectsNode.FirstNode != null && selectsNode != null)
                {
                    contextMenuStrip1.Items[0].Enabled = true;
                }
                else
                {
                    contextMenuStrip1.Items[0].Enabled = false;
                }
            }

        }

        private void переименоватьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode selectsNode = treeView1.SelectedNode;
            if (selectsNode != null)
            {
                selectsNode.BeginEdit();
            }
        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode selectsNode = treeView1.SelectedNode;
            if (selectsNode != null)
            {
                selectsNode.Remove();
            }


        }

        private void стажерыToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            saveFileDialog1.Filter = "Text Files(*.txt)|*.txt|All files(*.*)|*.*";
            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            string filename = saveFileDialog1.FileName;
         
            int i = 0;
            Pers pp = workers.Item(1);
            File.WriteAllText(filename, pp.nomer.ToString() + "+" + pp.filial.ToString() + "+" + pp.code.ToString() + "+" + pp.fio + "+" + pp.age.ToString()
                      + "+" + pp.money.ToString() + "+" + pp.stag.ToString() + Environment.NewLine);
            foreach (Pers p in workers)
            {
                if (1 <= i && i < workers.Count() - 1)
                {
                    File.AppendAllText(filename, p.nomer.ToString() + "+" + p.filial.ToString() + "+" + p.code.ToString() + "+" + p.fio + "+" + p.age.ToString()
                        + "+" + p.money.ToString() + "+" + p.stag.ToString() + Environment.NewLine);
                }
                i++;
            }

            pp = workers.Item(workers.Count());
            File.AppendAllText(filename, pp.nomer.ToString() + "+" + pp.filial.ToString() + "+" + pp.code.ToString() + "+" + pp.fio + "+" + pp.age.ToString()
                      + "+" + pp.money.ToString() + "+" + pp.stag.ToString());
            MessageBox.Show("Файл сохранен","Сохранение в "+filename,MessageBoxButtons.OK,MessageBoxIcon.Information);
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "C:\\";
                openFileDialog1.Filter = "Text Files(*.txt)|*.txt|All files(*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = openFileDialog.FileName;
                    var fileStream = openFileDialog.OpenFile();
                    using (StreamReader reader = new StreamReader(fileStream))
                    {
                        fileContent = reader.ReadToEnd();
                    }

                    DialogResult dr = DialogResult.None;

                    dr = MessageBox.Show("Вы точно хотите открыть этот файл?", "File Content at path:" + filePath, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dr == DialogResult.Yes)
                    {
                        NewOpenFile(filePath);
                    }
                }
            }


        }

        private void NewOpenFile(string path)
        {


            workers = null;
            listView1.Items.Clear();
            PeopleCollect.filepath = path;
            workers = new PeopleCollect();
            MessageBox.Show("Данные успешно загружены\n Перейдите в раздел еще раз,чтобы увидеть данные", "Обновление данных", MessageBoxButtons.OK, MessageBoxIcon.Information);


        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Text Files(*.txt)|*.txt|All files(*.*)|*.*";
            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            string filename = saveFileDialog1.FileName;
            int i = 1;

            ListViewItem p = listView1.Items[0];
            foreach (Pers pp in workers)
            {

                if (p.Tag.ToString() == pp.nomer.ToString())
                {
                    File.WriteAllText(filename, pp.nomer.ToString() + "+" + pp.filial.ToString() + "+" + pp.code.ToString() + "+" + pp.fio + "+" + pp.age.ToString()
                      + "+" + pp.money.ToString() + "+" + pp.stag.ToString() + Environment.NewLine);
                    break;
                }
            }

            foreach (Pers pp in workers)
            {

                p = listView1.Items[i];
                if (p.Tag.ToString() == pp.nomer.ToString() && i < listView1.Items.Count - 1)
                {
                    File.AppendAllText(filename, pp.nomer.ToString() + "+" + pp.filial.ToString() + "+" + pp.code.ToString() + "+" + pp.fio + "+" + pp.age.ToString()
                      + "+" + pp.money.ToString() + "+" + pp.stag.ToString() + Environment.NewLine);
                    i++;
                }
                else if (p.Tag.ToString() == pp.nomer.ToString() && i >= listView1.Items.Count - 1)
                {
                    File.AppendAllText(filename, pp.nomer.ToString() + "+" + pp.filial.ToString() + "+" + pp.code.ToString() + "+" + pp.fio + "+" + pp.age.ToString()
                  + "+" + pp.money.ToString() + "+" + pp.stag.ToString());
                    break;
                }

            }
            MessageBox.Show("Файл сохранен", "Сохранение в " + filename, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void hToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void изменитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
 
        }
        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {

            foreach (Pers pp in workers)
            {
                if (lvlItem.Tag.ToString() == pp.nomer.ToString())
                {
                    if (redact == "ФИО")
                    {
                        pp.fio = comboBox1.Text;
                        lvlItem.Text = this.comboBox1.Text;
                        break;
                    }
                    else if (redact == "Age")
                    {
                        char[] prov = comboBox1.Text.ToCharArray();
                        for (int i = 0; i < prov.Length; i++)
                        {
                            if (!Char.IsDigit(prov[i]))
                            {
                                MessageBox.Show("Данные некорректны", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                break;

                            }
                            else
                            {
                                pp.age = Int32.Parse(comboBox1.Text);
                                lvlItem.SubItems[2].Text = this.comboBox1.Text;
                                break;
                            }
                        }
                    }
                    else if (redact == "Money")
                    {
                        char[] prov = comboBox1.Text.ToCharArray();
                        for (int i = 0; i < prov.Length; i++)
                        {
                            if (!Char.IsDigit(prov[i]))
                            {
                                MessageBox.Show("Данные некорректны", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                break;

                            }
                            else
                            {
                                pp.money = Int32.Parse(comboBox1.Text);
                                lvlItem.SubItems[3].Text = this.comboBox1.Text;
                                break;
                            }
                        }
                    }
                    else if (redact == "Stag")
                    {
                        if (this.comboBox1.Text == "Да")
                        {
                            pp.stag = 1;
                            lvlItem.SubItems[4].Text = "Да";
                            break;
                        }
                        else if (this.comboBox1.Text == "Нет")
                        {
                            pp.stag = 0;
                            lvlItem.SubItems[4].Text = "Нет";
                            break;
                        }
                        else
                        {
                            //MessageBox.Show("Данные некорректны", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            //break;
                        }


                    }
                }
            }
            listView1.HoverSelection = true;
            this.comboBox1.Visible = false;
            this.comboBox1.Items.Clear();
            listView1.Refresh();
        }

        private void comboBox1_Leave(object sender, EventArgs e)
        {
            
              foreach (Pers pp in workers)
               {
                    if (lvlItem.Tag.ToString() == pp.nomer.ToString())
                    {
                        if (redact == "ФИО")
                        {
                            pp.fio = comboBox1.Text;
                        lvlItem.Text = this.comboBox1.Text;
                        break;
                        }
                        else if (redact == "Age")
                        {
                            char[] prov = comboBox1.Text.ToCharArray();
                            for (int i = 0; i < prov.Length; i++)
                            {
                                if (!Char.IsDigit(prov[i]))
                                {
                                    MessageBox.Show("Данные некорректны", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;

                                }
                                else
                                {
                                    pp.age = Int32.Parse(comboBox1.Text);
                                    lvlItem.SubItems[2].Text = this.comboBox1.Text;
                                    break;
                                }
                            }
                        }
                        else if (redact == "Money")
                        {
                            char[] prov = comboBox1.Text.ToCharArray();
                            for (int i = 0; i < prov.Length; i++)
                            {
                                if (!Char.IsDigit(prov[i]))
                                {
                                    MessageBox.Show("Данные некорректны", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;

                                }
                                else
                                {
                                    pp.money = Int32.Parse(comboBox1.Text);
                                    lvlItem.SubItems[3].Text = this.comboBox1.Text;
                                    break;
                                }
                            }
                        }
                        else if (redact == "Stag")
                        {
                            if (this.comboBox1.Text=="Да")
                            {
                                pp.stag = 1;
                                lvlItem.SubItems[4].Text = "Да";
                                break;
                            }
                            else if (this.comboBox1.Text == "Нет")
                            {
                                pp.stag = 0;
                                lvlItem.SubItems[4].Text = "Нет";
                                break;
                            }
                            else
                            {
                                //MessageBox.Show("Данные некорректны", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                //break;
                            }


                        }
                    }
                }
            listView1.HoverSelection = true;
            this.comboBox1.Visible = false;
            this.comboBox1.Items.Clear();
            listView1.Refresh();
        }

        private void comboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            switch (e.KeyChar)
            {
                case (char)(int)Keys.Escape:
                    {
                        //this.comboBox1.Text = lvlItem.Text;
                        foreach (Pers pp in workers)
                        {
                            if (lvlItem.Tag.ToString() == pp.nomer.ToString())
                            {
                                if (redact == "ФИО")
                                {
                                    pp.fio = comboBox1.Text;
                                    lvlItem.Text = this.comboBox1.Text;
                                    break;
                                }
                                else if (redact == "Age")
                                {
                                    char[] prov= comboBox1.Text.ToCharArray();
                                    for(int i=0;i<prov.Length;i++)
                                    {
                                        if(!Char.IsDigit(prov[i]))
                                        {
                                            MessageBox.Show("Данные некорректны", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            break;

                                        }
                                        else
                                        {
                                            pp.age = Int32.Parse(comboBox1.Text);
                                            lvlItem.SubItems[2].Text = this.comboBox1.Text;
                                            break;
                                        }
                                    }
                                }
                                else if (redact == "Money")
                                {
                                    char[] prov = comboBox1.Text.ToCharArray();
                                    for (int i = 0; i < prov.Length; i++)
                                    {
                                        if (!Char.IsDigit(prov[i]))
                                        {
                                            MessageBox.Show("Данные некорректны", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            break;

                                        }
                                        else
                                        {
                                            pp.money = Int32.Parse(comboBox1.Text);
                                            lvlItem.SubItems[3].Text = this.comboBox1.Text;
                                            break;
                                        }
                                    }
                                }
                                else if (redact=="Stag")
                                {
                                    if (this.comboBox1.Text == "Да")
                                    {
                                        pp.stag= 1;
                                        lvlItem.SubItems[4].Text = this.comboBox1.Text;
                                        break;
                                    }
                                    else if(this.comboBox1.Text=="Нет")
                                    {
                                        pp.stag = 0;
                                        lvlItem.SubItems[4].Text = this.comboBox1.Text;
                                        break;
                                    }
                                    else
                                    {
                                        //MessageBox.Show("Данные некорректны", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        //break;
                                    }
                                    
                                        
                                }
                            }
                           
                           
                        }
                        listView1.HoverSelection = true;
                        this.comboBox1.Visible = false;
                        listView1.Refresh();
                        break;
                    }
                case (char)(int)Keys.Enter:
                    {
                        foreach (Pers pp in workers)
                        {

                            if (lvlItem.Tag.ToString() == pp.nomer.ToString())
                            {
                                if (redact == "ФИО")
                                {
                                    pp.fio = lvlItem.Text;
                                    lvlItem.Text = this.comboBox1.Text;
                                    break;
                                }
                                else if (redact == "Age")
                                {
                                    char[] prov = comboBox1.Text.ToCharArray();
                                    for (int i = 0; i < prov.Length; i++)
                                    {
                                        if (!Char.IsDigit(prov[i]))
                                        {
                                            MessageBox.Show("Данные некорректны", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            break;

                                        }
                                        else
                                        {
                                            pp.age = Int32.Parse(comboBox1.Text);
                                            lvlItem.SubItems[2].Text = this.comboBox1.Text;
                                            break;
                                        }
                                    }
                                }
                                else if (redact == "Money")
                                {
                                    char[] prov = comboBox1.Text.ToCharArray();
                                    for (int i = 0; i < prov.Length; i++)
                                    {
                                        if (!Char.IsDigit(prov[i]))
                                        {
                                            MessageBox.Show("Данные некорректны", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            break;

                                        }
                                        else 
                                        {
                                            pp.money = Int32.Parse(comboBox1.Text);
                                            lvlItem.SubItems[3].Text = this.comboBox1.Text;
                                            break;
                                        }
                                    }
                                }
                                else if (redact == "Stag")
                                {
                                    if (this.comboBox1.Text == "Да")
                                    {
                                        pp.stag = 1;
                                        lvlItem.SubItems[4].Text = this.comboBox1.Text;
                                        break;
                                    }
                                    else if (this.comboBox1.Text == "Нет")
                                    {
                                        pp.stag = 0;
                                        lvlItem.SubItems[4].Text = this.comboBox1.Text;
                                        break;
                                    }
                                    else
                                    {
                                       // MessageBox.Show("Данные некорректны", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        //break;
                                    }


                                }
                            }
                        }
                        listView1.HoverSelection = true;
                        this.comboBox1.Visible = false;
                        this.comboBox1.Items.Clear();
                        listView1.Refresh();
                        break;
                    }
            }

        }

        private void фиоToolStripMenuItem1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void фиоToolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {

        }

        private void удалитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ListViewItem deleteItem;
            if (listView1.Items.Count != 0)
            {
                deleteItem = listView1.SelectedItems[0];
            }
            else
            {
                deleteItem = null;
            }

            if (deleteItem != null)
            {
                foreach (Pers pp in workers)
                {


                    if (deleteItem.Tag.ToString() == pp.nomer.ToString())
                    {
                        int d = pp.nomer;



                        //workers.Remove(d-1);
                        //deleteItem.Remove();
                        //PeopleCollect.lst.RemoveAt(d - 1);

                        pp.code = 9999;
                        pp.filial = 9999;
                        pp.nomer = 45667654;
                        pp.fio = "ddrrrrrrrt";
                        deleteItem.Remove();
                        listView1.Refresh();
                        MessageBox.Show("Сотрудник успешно удален\n Перейдите в раздел еще раз,чтобы обновить Справку", "Обновление данных", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        break;
                    }
                }

            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (Char.IsLetter(ch)) return;
            if (Char.IsControl(ch)) return;
            if (Char.IsWhiteSpace(ch)) return;
            e.Handled = true;
            //{
            //}
            //else if(ch == (char)8)
            //if (Char.IsDigit(ch)||!Char.IsLetter(ch)|| Char.IsNumber(ch)||!Char.IsWhiteSpace(ch)||ch!=(char)8)
            //{
            //    e.Handled = true;
            //}
            //else 
            //{

            //    e.Handled = false;
            //}
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox1.KeyPress += textBox1_KeyPress;
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (Char.IsDigit(ch)) return;
            if (Char.IsControl(ch)) return;
            e.Handled = true;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            textBox2.KeyPress += textBox2_KeyPress;
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (Char.IsDigit(ch)) return;
            if (Char.IsControl(ch)) return;
            e.Handled = true;
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            textBox3.KeyPress += textBox3_KeyPress;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                string fios = textBox1.Text;
                string ages = textBox2.Text;
                string moneys = textBox3.Text;
                int f = 1, c = 11, s = 0;

                if (fios != null && fios != "" && fios != " " && ages != "" && moneys != "" && comboBox2.Text != "" && comboBox3.Text != "")
                {
                    if (comboBox2.Text == "Восточный")
                    {
                        f = 1;
                        if (comboBox3.Text == "Бухгалтерия")
                        {
                            c = 11;
                        }
                        else if (comboBox3.Text == "Кадры")
                        {
                            c = 12;
                        }
                        else if (comboBox3.Text == "Производство")
                        {
                            c = 13;
                        }
                    }
                    else if (comboBox2.Text == "Южный")
                    {
                        f = 2;
                        if (comboBox3.Text == "Бухгалтерия")
                        {
                            c = 21;
                        }
                        else if (comboBox3.Text == "Кадры")
                        {
                            c = 22;
                        }
                        else if (comboBox3.Text == "Производство")
                        {
                            c = 23;
                        }
                    }
                    else if (comboBox2.Text == "Западный")
                    {
                        f = 3;
                        if (comboBox3.Text == "Бухгалтерия")
                        {
                            c = 31;
                        }
                        else if (comboBox3.Text == "Кадры")
                        {
                            c = 32;
                        }
                        else if (comboBox3.Text == "Производство")
                        {
                            c = 33;
                        }
                    }
                    else if (comboBox2.Text == "Северный")
                    {
                        f = 4;
                        if (comboBox3.Text == "Бухгалтерия")
                        {
                            c = 41;
                        }
                        else if (comboBox3.Text == "Кадры")
                        {
                            c = 42;
                        }
                        else if (comboBox3.Text == "Производство")
                        {
                            c = 43;
                        }
                    }
                    else if (comboBox2.Text == "Центральный")
                    {
                        f = 5;
                        if (comboBox3.Text == "Бухгалтерия")
                        {
                            c = 51;
                        }
                        else if (comboBox3.Text == "Кадры")
                        {
                            c = 52;
                        }
                        else if (comboBox3.Text == "Производство")
                        {
                            c = 53;
                        }
                    }
                    if (checkBox1.Checked == true)
                    {
                        s = 1;
                    }
                    Pers p = new Pers(workers.Count() + 1, f, c, fios, Int32.Parse(ages), Int32.Parse(moneys), s);
                    workers.Add(p);
                    groupBox1.Visible = false;
                    textBox1.Text = "";
                    textBox2.Text = "";
                    textBox3.Text = "";
                    checkBox2.Checked = false;
                    MessageBox.Show("Новый сотрудник уcпешно добавлен", "Добавление данных", MessageBoxButtons.OK, MessageBoxIcon.Information);



                }
                else
                {
                    checkBox2.Checked = false;
                    MessageBox.Show("Данные некорректны", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }





            }


        }

        private void toolStripButton4_Click_1(object sender, EventArgs e)
        {
            groupBox1.Visible = true;
        }

        private void toolStripComboBox1_Click(object sender, EventArgs e)
        {

        }

        private void toolStripTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (toolStripComboBox1.Text == "По ФИО")
            {  
                char ch = e.KeyChar;
               
                if (Char.IsLetter(ch)) return;
                if (Char.IsControl(ch)) return;
                if (Char.IsWhiteSpace(ch)) return;
                e.Handled = true;
            }
            else if (toolStripComboBox1.Text == "По номеру")
            {
                char ch = e.KeyChar;
                if (Char.IsDigit(ch)) return;
                if (Char.IsControl(ch)) return;
                e.Handled = true;
            }
            else if (toolStripComboBox1.Text == "")
            {
                MessageBox.Show("Пожалуйста, выберите режим поиска в окне Поиск", "Не задан режим поиска", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
            }
        }


        private void toolStripTextBox1_Click(object sender, EventArgs e)
        {

        }

        private void toolStripTextBox1_TextChanged(object sender, EventArgs e)
        {
            toolStripTextBox1.KeyPress += toolStripTextBox1_KeyPress;
        }

        private void toolStripTextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            //if (toolStripComboBox1.Text == "")
            //{
            //    MessageBox.Show("Пожалуйста, выберите режим поиска в окне Поиск", "Не задан режим поиска", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    e.Handled = true;
            //}
            //else
            if(e.KeyCode==Keys.Enter && toolStripComboBox1.Text != "")
            {
                Poisk();
            }
        }

        private void Poisk()
        {
            label9.Text = "";
            label1.Text = "";
            label2.Text = "";
            label3.Text = "";
            label4.Text = "";
            label5.Text = "";
            label6.Text = "";
            label7.Text = "";
            label8.Text = "";
            if (toolStripComboBox1.Text == "По фамилии")
            {
                string fio_p = toolStripTextBox1.Text;
                char[] c = fio_p.ToCharArray();
                int j = -1;
                Pers[] sovpad=new Pers[workers.Count()];
                sovpad[0] = null;
                bool check = true;
                foreach (Pers p in workers)
                {
                    check = true;
                    char[] cp = p.fio.ToCharArray();
                    for(int i=0;i<c.Length;i++)
                    {
                        if(cp[i]!=c[i])
                        {
                            check = false;
                            break;
                        }
                    }
                    if(check)
                    {
                        j++;
                        sovpad[j] = p;
                    }
                }
                if (sovpad[0]!=null)
                {
                    listView1.Items.Clear();
                    //listView1.Refresh();
                    for (int i = 0; i <= j; i++)
                    {


                        ListViewItem item1 = new ListViewItem(sovpad[i].fio);
                        item1.SubItems.Add(sovpad[i].nomer.ToString());
                        item1.SubItems.Add(sovpad[i].age.ToString());
                        item1.SubItems.Add(sovpad[i].money.ToString());
                        if (sovpad[i].stag == 1)
                        {
                            item1.SubItems.Add("Да");
                        }
                        else
                        {
                            item1.SubItems.Add("Нет");
                        }
                        item1.Tag = sovpad[i].nomer.ToString();
                        item1.ImageIndex = 5;
                        listView1.Items.Add(item1);




                    }

                    
                }
                else
                {
                    listView1.Items.Clear();
                    MessageBox.Show("По данному запросу ничего не найдено", "Результат поиска", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                
            }
            else if (toolStripComboBox1.Text == "По номеру")
            {
                string num_s = toolStripTextBox1.Text;
                if(num_s!="")
                {
                    int num_p = Int32.Parse(num_s);

                    Pers sov = null;
                    foreach (Pers p in workers)
                    {
                        if (num_p == p.nomer)
                        {
                            sov = p;
                            break;
                        }

                    }
                    if (sov != null)
                    {

                        listView1.Items.Clear();
                        //listView1.Refresh();


                        ListViewItem item1 = new ListViewItem(sov.fio);
                        item1.SubItems.Add(sov.nomer.ToString());
                        item1.SubItems.Add(sov.age.ToString());
                        item1.SubItems.Add(sov.money.ToString());
                        if (sov.stag == 1)
                        {
                            item1.SubItems.Add("Да");
                        }
                        else
                        {
                            item1.SubItems.Add("Нет");
                        }
                        item1.Tag = sov.nomer.ToString();
                        item1.ImageIndex = 5;
                        listView1.Items.Add(item1);

                    }
                    else
                    {
                        listView1.Items.Clear();
                        MessageBox.Show("По данному запросу ничего не найдено", "Результат поиска", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    listView1.Items.Clear();
                    MessageBox.Show("По данному запросу ничего не найдено", "Результат поиска", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }




            }


            toolStripTextBox1.Text = "";

        }

        private void фИОToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            if (listView1.Items.Count != 0)
            {
                lvlItem = listView1.SelectedItems[0];
            }
            else
            {
                lvlItem = null;
            }

          
            if (lvlItem != null)
            {
                redact = "ФИО";
                
                Rectangle ClickedItem = lvlItem.Bounds;

              
                if ((ClickedItem.Left + this.listView1.Columns[0].Width) < 0)
                {
                   
                    return;
                }

                
                else if (ClickedItem.Left < 0)
                {
                   
                    if ((ClickedItem.Left + this.listView1.Columns[0].Width) > this.listView1.Width)
                    {
                        
                        ClickedItem.Width = this.listView1.Width;
                        ClickedItem.X = 0;
                    }
                    else
                    {
                       
                        ClickedItem.Width = this.listView1.Columns[0].Width + ClickedItem.Left;
                        ClickedItem.X = 2;
                    }
                }
                else if (this.listView1.Columns[0].Width > this.listView1.Width)
                {
                    ClickedItem.Width = this.listView1.Width;
                }
                else
                {
                    ClickedItem.Width = this.listView1.Columns[0].Width;
                    ClickedItem.X = 2;
                }

                
                ClickedItem.Y += this.listView1.Top;
                ClickedItem.X += this.listView1.Left;

               
                this.comboBox1.Bounds = ClickedItem;

               
                this.comboBox1.Text = lvlItem.Text;
                this.comboBox1.DropDownStyle = ComboBoxStyle.DropDown;
               
                this.comboBox1.Visible = true;
                this.comboBox1.BringToFront();
                this.comboBox1.Focus();
            }
        }

        private void возрастToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listView1.Items.Count != 0)
            {
                lvlItem = listView1.SelectedItems[0];
            }
            else
            {
                lvlItem = null;
            }


            if (lvlItem != null)
            {
                redact = "Age";
              
                Rectangle ClickedItem = lvlItem.Bounds;
              
                
                

          
                 if (ClickedItem.Left < 0)
                {
         
                    if ((ClickedItem.Left + this.listView1.Columns[2].Width) > this.listView1.Width)
                    {
                        
                        ClickedItem.Width = listView1.Columns[2].Width;
                       
                        ClickedItem.X = listView1.Columns[0].Width + listView1.Columns[1].Width;
                    }
                    else
                    {
                        
                        ClickedItem.Width = this.listView1.Columns[2].Width + ClickedItem.Left;
                        ClickedItem.X = listView1.Columns[0].Width + listView1.Columns[1].Width;
                    }
                }
                else if (this.listView1.Columns[2].Width > this.listView1.Width)
                {
                    ClickedItem.Width = this.listView1.Width;
                }
                else
                {
                    ClickedItem.Width = this.listView1.Columns[2].Width;
                    ClickedItem.X = listView1.Columns[0].Width + listView1.Columns[1].Width;
                }

          
                ClickedItem.Y += this.listView1.Top;
                ClickedItem.X += this.listView1.Left;

              
                this.comboBox1.Bounds = ClickedItem;

                
                this.comboBox1.Text = lvlItem.SubItems[2].Text;
                this.comboBox1.DropDownStyle = ComboBoxStyle.DropDown;
              
                this.comboBox1.Visible = true;
                this.comboBox1.BringToFront();
                this.comboBox1.Focus();
            }
        }

        private void зарплатуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listView1.Items.Count != 0)
            {
                lvlItem = listView1.SelectedItems[0];
            }
            else
            {
                lvlItem = null;
            }


            if (lvlItem != null)
            {
                redact = "Money";
             
                Rectangle ClickedItem = lvlItem.Bounds;




              
                if (ClickedItem.Left < 0)
                {
                    if ((ClickedItem.Left + this.listView1.Columns[3].Width) > this.listView1.Width)
                    {
                        
                        ClickedItem.Width = listView1.Columns[3].Width;

                        ClickedItem.X = listView1.Columns[0].Width + listView1.Columns[1].Width+ listView1.Columns[2].Width;
                    }
                    else
                    {
                     
                        ClickedItem.Width = this.listView1.Columns[3].Width + ClickedItem.Left;
                        ClickedItem.X = listView1.Columns[0].Width + listView1.Columns[1].Width+ listView1.Columns[2].Width;
                    }
                }
                else if (this.listView1.Columns[3].Width > this.listView1.Width)
                {
                    ClickedItem.Width = this.listView1.Width;
                }
                else
                {
                    ClickedItem.Width = this.listView1.Columns[3].Width;
                    ClickedItem.X = listView1.Columns[0].Width + listView1.Columns[1].Width+ listView1.Columns[2].Width;
                }

                ClickedItem.Y += this.listView1.Top;
                ClickedItem.X += this.listView1.Left;

                this.comboBox1.Bounds = ClickedItem;

                
                this.comboBox1.Text = lvlItem.SubItems[3].Text;
                this.comboBox1.DropDownStyle = ComboBoxStyle.DropDown;
              
                this.comboBox1.Visible = true;
                this.comboBox1.BringToFront();
                this.comboBox1.Focus();
            }
        }

        private void стажировкаToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (listView1.Items.Count != 0)
            {
                lvlItem = listView1.SelectedItems[0];
            }
            else
            {
                lvlItem = null;
            }
            listView1.HoverSelection = false;

            if (lvlItem != null)
            {
                redact = "Stag";
            
                Rectangle ClickedItem = lvlItem.Bounds;




                ClickedItem.Width = listView1.Width - listView1.Columns[0].Width - listView1.Columns[1].Width - listView1.Columns[2].Width - listView1.Columns[3].Width-15;
                ClickedItem.X = listView1.Columns[0].Width + listView1.Columns[1].Width + listView1.Columns[2].Width+ listView1.Columns[3].Width;
               
                ClickedItem.Y += this.listView1.Top;
                ClickedItem.X += this.listView1.Left;

               
                this.comboBox1.Bounds = ClickedItem;

               
                this.comboBox1.Text = lvlItem.SubItems[4].Text;

                this.comboBox1.Items.Add("Да");
                this.comboBox1.Items.Add("Нет");
                this.comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
              
              
                this.comboBox1.Visible = true;
                this.comboBox1.BringToFront();
                this.comboBox1.Focus();
            }
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }
    }
}