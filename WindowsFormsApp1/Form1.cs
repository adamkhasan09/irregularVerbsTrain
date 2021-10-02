using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;


namespace WindowsFormsApp1
{
    
    public partial class Form1 : Form
    {
        static string fileName = @"\source.xlsx";
        ExcelHelper excelObj = new ExcelHelper(path + fileName, 1);
        BaseHelper baseH = new BaseHelper();
        Random random = new Random();
        public static string path = Directory.GetCurrentDirectory();
        DataTable tb = new DataTable();
        DataColumn column;
        DataRow row;
        int[] indexs;
        List<string> dKnow = new List<string>();
        List<string> know = new List<string>();
        int startPoint, endPoint;
        bool cell_mode = false;
        public Form1()
        {
            InitializeComponent();
            Application.ApplicationExit += new EventHandler(this.OnApplicationExit);
            tb = CreateStruct();
            dataGridView1.Visible = false;
            label7.Visible = false;
        }
        
        // select aria
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                cell_mode = false;
                textBox7.Visible = true;
                tb.Clear();
                startPoint = int.Parse(textBox1.Text);
                endPoint = int.Parse(textBox2.Text);
                tb = fillStruct(excelObj, tb, startPoint, endPoint);
                dataGridView1.DataSource = tb;
                indexs = shakeArray(tableGetIndexs());
                //MessageBox.Show(strIntArr(indexs));
                shakeTable();
                Stydy();

            }
            catch
            {
                MessageBox.Show("Введены не корректные данные");
            }
        }
        //next 
        private void button7_Click(object sender, EventArgs e)
        {
            if (tb.Rows.Count > 0)
            {
                
                if (cell_mode)
                {
                    know.Add(tb.Rows[0]["id"].ToString());
                    string[] strArr = know.ToArray().Distinct().ToArray();
                    know = strArr.ToList();
                    label22.Text = baseH.inplode(know.ToArray(), ",");
                    //MessageBox.Show(strIntArr(indexs));
                }
                tb.Rows[0].Delete();
                tb.AcceptChanges();
                Stydy();
            }
            else
            {
                if (cell_mode)
                {
                    MessageBox.Show("Выбранный диапазон из сохранной строки закончился");
                }
            }

        }
        // dont know
        private void button8_Click(object sender, EventArgs e)
        {
            dataGridView1.Visible = !dataGridView1.Visible;
            label7.Visible = !label7.Visible;
            label22.Visible = !label22.Visible;
        }
        // save
        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                if (!cell_mode)
                {
                    excelObj.Close();
                    ExcelHelper excelObj2 = new ExcelHelper(path + fileName, 2);
                    int save_cell = int.Parse(textBox7.Text);
                    string cell = excelObj2.ReadCell(save_cell,1);
                    //MessageBox.Show(cell);
                    if (cell == "")
                    {
                        excelObj2.WriteCell(save_cell, 1, baseH.inplode(dKnow.ToArray(), ","));
                    }
                    else
                    {
                        List<string> currentIds = baseH.explode(",", cell).ToList();
                        List<string> res = currentIds.Concat(dKnow).ToList().Distinct().ToList();
                        MessageBox.Show(baseH.inplode(res.ToArray(),","));
                        excelObj2.WriteCell(save_cell, 1, baseH.inplode(res.ToArray(), ","));
                    }
                   
                    excelObj2.Save();
                    excelObj2.Close();
                    excelObj = new ExcelHelper(path +fileName, 1);
                    MessageBox.Show("Ваши слова успешно схраненены строка для извлечения: " + textBox7.Text);
                }
                if (cell_mode)
                {
                    excelObj.Close();
                    int saveIdx = int.Parse(textBox6.Text);
                    ExcelHelper excelObj2 = new ExcelHelper(path +fileName, 2);
                    string ids = excelObj2.ReadCell(saveIdx, 1);
                    
                    if (ids == "")
                    {
                       
                    }
                    else
                    {
                        
                        DialogResult dr = MessageBox.Show("Удалить слова которые вы знаете?", "Выберете действие", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information);
                        if (dr == DialogResult.Yes)
                        {
                            string[] idsArr = baseH.explode(",", ids);
                            List<string> idsList = idsArr.ToList();
                            //MessageBox.Show(baseH.inplode(idsList.ToArray(), ","));
                            for (int i = 0; i < idsArr.Count(); i++)
                            {
                                int id = know.IndexOf(idsArr[i]);
                                //MessageBox.Show(id.ToString());
                                if(id >= 0)
                                {
                                   idsList.Remove(know[id]);
                                }
                            }
                            string res = baseH.inplode(idsList.ToArray(), ",");
                            excelObj2.WriteCell(saveIdx, 1, res);
                            excelObj2.Save();
                            MessageBox.Show("Изменения успешно сохранены");
                        }
                       
                    }
                    excelObj2.Close();
                    excelObj = new ExcelHelper(path +fileName, 1);


                }
            }
            catch
            {
                MessageBox.Show("not corrent data");
            }

           
            
        }
        //learning from cells
        private void button2_Click(object sender, EventArgs e)
        {
            cell_mode = true;
            textBox7.Visible = false;
            try{
                excelObj.Close();
                int cellNo = int.Parse(textBox6.Text);
                ExcelHelper excelObj2 = new ExcelHelper(path +fileName, 2);
                string idxs = excelObj2.ReadCell(cellNo, 1);
                excelObj2.Close();
                excelObj = new ExcelHelper(path +fileName, 1);
                //MessageBox.Show(idxs);
                
                string[] idxsArr = baseH.explode(",", idxs);
                int[] idxsIntArr = new int[idxsArr.Count()];
                for (int i = 0; i < idxsArr.Count(); i++)
                {
                    idxsIntArr[i] = int.Parse(idxsArr[i]);
                }
                indexs = idxsIntArr;
                tb.Clear();
                string range = "";
                foreach (int id in indexs)
                {
                    string[] rangeArray;
                    range = excelObj.ReadRange(id, 4, 0, "@");
                    rangeArray = baseH.explode("@", range);
                    row = tb.NewRow();
                    row["id"] = id;
                    row["translate"] = rangeArray[0];
                    row["first_form"] = rangeArray[1];
                    row["second_form"] = rangeArray[2];
                    row["third_form"] = rangeArray[3];
                    tb.Rows.Add(row);

                }
                dataGridView1.DataSource = tb;
                indexs = shakeArray(tableGetIndexs());
                //MessageBox.Show(strIntArr(indexs));
                shakeTable();
                Stydy();
                
            }
            catch
            {

            }
        }
        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                if (!cell_mode)
                {
                    dKnow.Add(tb.Rows[0]["id"].ToString());
                }
                string[] strArr = dKnow.ToArray().Distinct().ToArray();
                dKnow = strArr.ToList();
                label7.Text = baseH.inplode(dKnow.ToArray(), ",");
                List<int> indxList = tableGetIndexs().ToList();
                Shuffle(indxList);
                indexs = indxList.ToArray();
                //MessageBox.Show(strIntArr(indexs));
                shakeTable();
                Stydy();

            }
            catch
            {
                MessageBox.Show("select new aria");
            }
        }
        DataTable fillStruct(ExcelHelper obj , DataTable table, int startRange, int endRange)
        {
            int count = endRange - startRange;
            string range = "";
            
            for (int i = 0; i < count; i++)
            {
                string[] rangeArray;
                range = obj.ReadRange(startRange + i, 4, 0, "@");
                rangeArray = baseH.explode("@", range);
                row = table.NewRow();
                row["id"] = startRange + i;
                row["translate"] = rangeArray[0];
                row["first_form"] = rangeArray[1];
                row["second_form"] = rangeArray[2];
                row["third_form"] = rangeArray[3];
                table.Rows.Add(row);

            }  
            return table;
        }

        public static void Shuffle<T>(IList<T> list)
        {
            Random random = new Random();
            int n = list.Count;
            while (n > 1)
            {
                n--;
                int k = random.Next(n + 1);
                T value = list[k];
                list[k] = list[n];
                list[n] = value;
            }
        }
        void Stydy()
        {
            setUnVisible();
            int index = comboBox1.SelectedIndex;
            if (tb.Rows.Count > 0)
            {
                if (index == 0)
                {
                    word.Text = tb.Rows[0]["first_form"].ToString();
                    trans.Text = tb.Rows[0]["translate"].ToString();
                    second.Text = tb.Rows[0]["second_form"].ToString();
                    third.Text = tb.Rows[0]["third_form"].ToString();
                    button3.Visible = false;
                    textBox3.Visible = false;

                }
                if (index == 1)
                {
                    word.Text = tb.Rows[0]["translate"].ToString();
                    trans.Text = tb.Rows[0]["first_form"].ToString();
                    second.Text = tb.Rows[0]["second_form"].ToString();
                    third.Text = tb.Rows[0]["third_form"].ToString();
                    button3.Visible = true;
                    textBox3.Visible = true;
                }
            }
            else
            {
                MessageBox.Show("Выбранный дипазон закончился");
            }
            
        }
        private void OnApplicationExit(object sender, EventArgs e)
        {
            excelObj.Close();
        }

       
        int[] tableGetIndexs()
        {
            int[] indexs = new int[tb.Rows.Count];
            for (int i = 0; i < tb.Rows.Count; i++)
            {
                indexs[i] = int.Parse(tb.Rows[i]["id"].ToString());
            }
            return indexs;
        }
        int[] shakeArray(int[] data)
        {
            Random random = new Random();
            for (int i = data.Length - 1; i >= 1; i--)
            {
                int j = random.Next(i + 1);
                // обменять значения data[j] и data[i]
                var temp = data[j];
                data[j] = data[i];
                data[i] = temp;
            }
            return data;
        }
        void shakeTable()
        {
            DataRow[] tableRows = tb.Select();
            for (int i = 0; i < indexs.Count(); i++)
            {
                string expression = "id =" + indexs[i];
                DataRow[] rows = tb.Select(expression);
                DataRow row = rows[0];
                int id = tb.Rows.IndexOf(row);
                DataRow tmpRow = tableRows[i];
                tableRows[i] = tableRows[id];
                tableRows[id] = tmpRow;
                tb = tableRows.CopyToDataTable();
                dataGridView1.DataSource = tb;
            }
        }

       

        string strIntArr(int[] array)
        {
            string[] strArr = new string[array.Count()];
            for (int i = 0; i < array.Count(); i++)
            {
                strArr[i] = array[i].ToString();

            }
            string res = baseH.inplode(strArr, ",");
            return res;

        }

        private void label8_Click(object sender, EventArgs e)
        {
            trans.Visible = !trans.Visible;
        }

        private void label10_Click(object sender, EventArgs e)
        {
            second.Visible = !second.Visible;
        }

        private void label11_Click(object sender, EventArgs e)
        {
            third.Visible = !third.Visible;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(trans.Text != textBox3.Text)
            {
                label14.Visible = true;
                label9.Visible = false;
            }
            else
            {
                label14.Visible = false;
                label9.Visible = true;
            }
        }
        void setUnVisible()
        {
            
            trans.Visible = false;
            second.Visible = false;
            third.Visible = false;
            label14.Visible = false;
            label9.Visible = false;
            label12.Visible = false;
            label15.Visible = false;
            label16.Visible = false;
            label13.Visible = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (second.Text != textBox4.Text)
            {
                label15.Visible = true;
                label12.Visible = false;
            }
            else
            {
                label15.Visible = false;
                label12.Visible = true;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (third.Text != textBox5.Text)
            {
                label16.Visible = true;
                label13.Visible = false;
            }
            else
            {
                label16.Visible = false;
                label13.Visible = true;
            }
        }

        

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        DataTable CreateStruct()
        {
            DataTable table = new DataTable();
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Int32");
            column.ColumnName = "id";
            table.Columns.Add(column);

            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "translate";
            table.Columns.Add(column);

            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "first_form";
            table.Columns.Add(column);

            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "second_form";
            table.Columns.Add(column);

            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "third_form";
            table.Columns.Add(column);

            return table;
        }
    }
}
