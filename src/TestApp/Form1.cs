using System.Data;
using ExcelDataReader;
using System.Diagnostics;
using System.Text;
using System.Xml.Linq;

namespace TestApp
{
    public partial class Form1 : Form
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.dataSet1 = new System.Data.DataSet();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.sheetCombo = new System.Windows.Forms.ComboBox();
            this.Sheet = new System.Windows.Forms.Label();
            this.firstRowNamesCheckBox = new System.Windows.Forms.CheckBox();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.label1 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.Filter = "Supported files|*.xls;*.xlsx;*.xlsb;*.csv|xls|*.xls|xlsx|*.xlsx|xlsb|*.xlsb|csv|*" +
    ".csv|All|*.*";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(518, 7);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(105, 29);
            this.button1.TabIndex = 0;
            this.button1.Text = "Select file";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Button1Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(69, 8);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(440, 28);
            this.textBox1.TabIndex = 1;
            this.textBox1.Text = "C:\\Users\\3\\Desktop\\excel\\MaterialLibrary.xlsx";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(16, 71);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(110, 40);
            this.button2.TabIndex = 2;
            this.button2.Text = "Process";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.Button2Click);
            // 
            // dataSet1
            // 
            this.dataSet1.DataSetName = "NewDataSet";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(16, 158);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowHeadersWidth = 62;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(1505, 1029);
            this.dataGridView1.TabIndex = 3;
            // 
            // sheetCombo
            // 
            this.sheetCombo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.sheetCombo.FormattingEnabled = true;
            this.sheetCombo.Location = new System.Drawing.Point(132, 123);
            this.sheetCombo.Name = "sheetCombo";
            this.sheetCombo.Size = new System.Drawing.Size(378, 26);
            this.sheetCombo.TabIndex = 4;
            this.sheetCombo.SelectedIndexChanged += new System.EventHandler(this.SheetComboSelectedIndexChanged);
            // 
            // Sheet
            // 
            this.Sheet.AutoSize = true;
            this.Sheet.Location = new System.Drawing.Point(18, 127);
            this.Sheet.Name = "Sheet";
            this.Sheet.Size = new System.Drawing.Size(116, 18);
            this.Sheet.TabIndex = 5;
            this.Sheet.Text = "Choose sheet";
            // 
            // firstRowNamesCheckBox
            // 
            this.firstRowNamesCheckBox.AutoSize = true;
            this.firstRowNamesCheckBox.Checked = true;
            this.firstRowNamesCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.firstRowNamesCheckBox.Location = new System.Drawing.Point(22, 42);
            this.firstRowNamesCheckBox.Name = "firstRowNamesCheckBox";
            this.firstRowNamesCheckBox.Size = new System.Drawing.Size(313, 22);
            this.firstRowNamesCheckBox.TabIndex = 6;
            this.firstRowNamesCheckBox.Text = "first row contains column names";
            this.firstRowNamesCheckBox.UseVisualStyleBackColor = true;
            // 
            // statusStrip1
            // 
            this.statusStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 1254);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Padding = new System.Windows.Forms.Padding(2, 0, 15, 0);
            this.statusStrip1.Size = new System.Drawing.Size(1537, 22);
            this.statusStrip1.TabIndex = 7;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(0, 15);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(18, 12);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(44, 18);
            this.label1.TabIndex = 8;
            this.label1.Text = "Path";
            // 
            // button3
            // 
            this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button3.Location = new System.Drawing.Point(1281, 1193);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(241, 40);
            this.button3.TabIndex = 9;
            this.button3.Text = "generate python script";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1537, 1276);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.firstRowNamesCheckBox);
            this.Controls.Add(this.Sheet);
            this.Controls.Add(this.sheetCombo);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button button2;
        private System.Data.DataSet dataSet1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.ComboBox sheetCombo;
        private System.Windows.Forms.Label Sheet;
        private System.Windows.Forms.CheckBox firstRowNamesCheckBox;
        private StatusStrip statusStrip1;
        private ToolStripStatusLabel toolStripStatusLabel1;
        private Label label1;
        private Button button3;
        private DataSet ds;

        public Form1()
        {
            InitializeComponent();
            if(textBox1.Text.Length > 0 )
            {
                loadXlsx();
            }
        }

        /*
        public static void GetValues(DataSet dataset, string sheetName)
        {
            foreach (DataRow row in dataset.Tables[sheetName].Rows)
            {
                foreach (var value in row.ItemArray)
                {
                    Console.WriteLine("{0}, {1}", value, value.GetType());
                }
            }
        }
        */

        private static IList<string> GetTablenames(DataTableCollection tables)
        {
            var tableList = new List<string>();
            foreach (var table in tables)
            {
                tableList.Add(table.ToString());
            }

            return tableList;
        }

        private void Button1Click(object sender, EventArgs e)
        {
            var result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        private void Button2Click(object sender, EventArgs e)
        {
            loadXlsx();
        }

        private void loadXlsx()
        {
            try
            {
                using var stream = new FileStream(textBox1.Text, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                var sw = new Stopwatch();
                sw.Start();

                using IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);

                var openTiming = sw.ElapsedMilliseconds;
                // reader.IsFirstRowAsColumnNames = firstRowNamesCheckBox.Checked;
                ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    UseColumnDataType = false,
                    ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = firstRowNamesCheckBox.Checked
                    }
                });

                toolStripStatusLabel1.Text = "Elapsed: " + sw.ElapsedMilliseconds.ToString() + " ms (" + openTiming.ToString() + " ms to open)";

                var tablenames = GetTablenames(ds.Tables);
                sheetCombo.DataSource = tablenames;

                if (tablenames.Count > 0)
                    sheetCombo.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SelectTable()
        {
            var tablename = sheetCombo.SelectedItem.ToString();

            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.DataSource = ds; // dataset
            dataGridView1.DataMember = tablename;

            // GetValues(ds, tablename);
        }

        private void SheetComboSelectedIndexChanged(object sender, EventArgs e)
        {
            SelectTable();
        }

        private void writePython()
        {
            // 指定python文件的路径  
            string pyFilePath = "C:\\Users\\3\\Desktop\\excel\\MaterialLibrary.py";
            // 使用StreamWriter来写入文件  
            using (StreamWriter sw = new StreamWriter(pyFilePath, false, Encoding.UTF8))
            {
                foreach (DataTable table in ds.Tables)
                {
                    if (table.Columns.Count != 8)
                    {
                        return;
                    }
                    // 输出当前DataTable的名称  
                    //Console.WriteLine("Table Name: " + table.TableName);

                    // 遍历DataTable中的所有行  
                    foreach (DataRow row in table.Rows)
                    {
                        string enName = row[1] != DBNull.Value ? row[0].ToString() : "NULL";
                        string Permittivity = row[2] != DBNull.Value ? row[2].ToString() : "NULL";
                        string Permeability = row[3] != DBNull.Value ? row[3].ToString() : "NULL";
                        string Conductivity = row[4] != DBNull.Value ? row[4].ToString() : "NULL";
                        string DielectricLossTangent = row[5] != DBNull.Value ? row[5].ToString() : "NULL";
                        string MagneticLossTangent = row[6] != DBNull.Value ? row[6].ToString() : "NULL";
                        string cnName = row[7] != DBNull.Value ? row[7].ToString() : "NULL";

                        string rowData1 = "MainWindow.createSoftwareMaterial(";
                        rowData1 = rowData1 + "\"" + enName + "\",";
                        rowData1 = rowData1 + "\"" + enName + "\",";
                        rowData1 = rowData1 + "\"Permittivity,0," + Permittivity + "|0|0|0|0|0|0|0|0,,0\", ";
                        rowData1 = rowData1 + "\"Permeability,0," + Permeability + "|0|0|0|0|0|0|0|0,,0\", ";
                        rowData1 = rowData1 + "\"Conductivity,0," + Conductivity + "|0|0|0|0|0|0|0|0,siemens/m,0\",";
                        rowData1 = rowData1 + "\"Dielectric Loss Tangent,0," + DielectricLossTangent + "|0|0|0|0|0|0|0|0,,0\",";
                        rowData1 = rowData1 + "\"Magnetic Loss Tangent,0," + MagneticLossTangent + "|0|0|0|0|0|0|0|0,,0\")";

                        sw.WriteLine(rowData1);
                    }
                }
            }

            Console.WriteLine("python文件已成功创建！");
        }

        private void writeJson()
        {
            // 指定python文件的路径  
            string FilePath = "C:\\Users\\3\\Desktop\\excel\\MaterialLibrary.json";
            // 使用StreamWriter来写入文件  
            using (StreamWriter sw = new StreamWriter(FilePath, false, Encoding.UTF8))
            {
                foreach (DataTable table in ds.Tables)
                {
                    if (table.Columns.Count != 8)
                    {
                        return;
                    }
                    // 输出当前DataTable的名称  
                    //Console.WriteLine("Table Name: " + table.TableName);

                    string rowData1 = "{\n";
                    rowData1 += "    \"m_materialDataArray\": [\n";
                    int rawIndex = 0;
                    // 遍历DataTable中的所有行  
                    foreach (DataRow row in table.Rows)
                    {
                        rawIndex++;
                        string enName = row[1] != DBNull.Value ? row[0].ToString() : "NULL";
                        string Permittivity = row[2] != DBNull.Value ? row[2].ToString() : "NULL";
                        string Permeability = row[3] != DBNull.Value ? row[3].ToString() : "NULL";
                        string Conductivity = row[4] != DBNull.Value ? row[4].ToString() : "NULL";
                        string DielectricLossTangent = row[5] != DBNull.Value ? row[5].ToString() : "NULL";
                        string MagneticLossTangent = row[6] != DBNull.Value ? row[6].ToString() : "NULL";
                        string cnName = row[7] != DBNull.Value ? row[7].ToString() : "NULL";

                        rowData1 += "        {\n";
                        rowData1 += "            \"m_conductivity\": {\n";
                        rowData1 += "                \"m_isSymmetric\": false,\n";
                        rowData1 += "                \"m_matrix\": [\n";
                        rowData1 += "                    " + Conductivity + ",\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0\n";
                        rowData1 += "                ],\n";
                        rowData1 += "                \"m_unitName\": \"siemens/m\",\n";
                        rowData1 += "                \"m_valueType\": 0\n";
                        rowData1 += "            },\n";
                        rowData1 += "            \"m_dielectricLossTangent\": {\n";
                        rowData1 += "                \"m_isSymmetric\": false,\n";
                        rowData1 += "                \"m_matrix\": [\n";
                        rowData1 += "                    " + DielectricLossTangent + ",\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0\n";
                        rowData1 += "                ],\n";
                        rowData1 += "                \"m_unitName\": \"\",\n";
                        rowData1 += "                \"m_valueType\": 0\n";
                        rowData1 += "            },\n";
                        rowData1 += "            \"m_editable\": false,\n";
                        rowData1 += "            \"m_fakeDelete\": false,\n";
                        rowData1 += "            \"m_id\": " + rawIndex + ",\n";
                        rowData1 += "            \"m_magneticLossTangent\": {\n";
                        rowData1 += "                \"m_isSymmetric\": false,\n";
                        rowData1 += "                \"m_matrix\": [\n";
                        rowData1 += "                    " + MagneticLossTangent + ",\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0\n";
                        rowData1 += "                ],\n";
                        rowData1 += "                \"m_unitName\": \"\",\n";
                        rowData1 += "                \"m_valueType\": 0\n";
                        rowData1 += "            },\n";
                        rowData1 += "            \"m_name\": \"" + enName + "\",\n";
                        rowData1 += "            \"m_nameEN\": \"" + enName + "\",\n";
                        rowData1 += "            \"m_permeability\": {\n";
                        rowData1 += "                \"m_isSymmetric\": false,\n";
                        rowData1 += "                \"m_matrix\": [\n";
                        rowData1 += "                    " + Permeability + ",\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0\n";
                        rowData1 += "                ],\n";
                        rowData1 += "                \"m_unitName\": \"\",\n";
                        rowData1 += "                \"m_valueType\": 0\n";
                        rowData1 += "            },\n";
                        rowData1 += "            \"m_permittivity\": {\n";
                        rowData1 += "                \"m_isSymmetric\": false,\n";
                        rowData1 += "                \"m_matrix\": [\n";
                        rowData1 += "                    " + Permittivity + ",\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0,\n";
                        rowData1 += "                    0\n";
                        rowData1 += "                ],\n";
                        rowData1 += "                \"m_unitName\": \"\",\n";
                        rowData1 += "                \"m_valueType\": 0\n";
                        rowData1 += "            }\n";
                        rowData1 += "        },\n";
                    }
                    rowData1 = rowData1.TrimEnd('\n');
                    rowData1 = rowData1.TrimEnd(',');
                    rowData1 += "\n";
                    rowData1 += "    ]\n";
                    rowData1 += "}\n";
                    sw.WriteLine(rowData1);
                }
            }
            Console.WriteLine("json文件已成功创建！");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            writePython();
            //writeJson();


            // 遍历DataSet中的所有DataTable  
            //foreach (DataTable table in ds.Tables)
            //{
            //    // 输出当前DataTable的名称  
            //    Console.WriteLine("Table Name: " + table.TableName);

            //    // 遍历DataTable中的所有行  
            //    foreach (DataRow row in table.Rows)
            //    {
            //        // 初始化一个字符串来存储行的内容  
            //        string rowData = "";

            //        // 遍历DataTable中的所有列  
            //        foreach (DataColumn column in table.Columns)
            //        {
            //            // 获取当前单元格的值（注意可能需要转换为字符串或检查是否为DBNull）  
            //            object item = row[column];
            //            rowData += item != DBNull.Value ? item.ToString() + "\t" : "NULL\t";
            //        }

            //        // 输出整行的数据  
            //        Console.WriteLine(rowData.TrimEnd('\t')); // 移除末尾的制表符  
            //    }

            //    // 在每个表后输出一个空行以便于阅读  
            //    Console.WriteLine();
            //}
        }
    }
}
