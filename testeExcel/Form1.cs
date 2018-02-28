using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data.Common;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Data.Sql;
using System.Runtime.InteropServices;
using System.Configuration;
using Microsoft.Office.Interop.Excel;
using ExcelIt = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Drawing;
using testeCampos;
using System.ComponentModel;
using System.Text.RegularExpressions;

namespace testeExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public static string path;
        public static string excelConnectionString;
        public string[] files;
        public string conexao;
        public string baseDeDados;
        public string tabela;
        public string caminho;
        public string directoryPath;
        private static Excel.Application MyApp = null;
        public List<string> filesAdionado = new List<string>();
        public List<string> colunas = new List<string>();
        public List<string> colunasCreate = new List<string>();
        public string tipoArquivo;
        Stream myStream = null;
        string nomeSheet;
        StringBuilder camposDataGrid = new StringBuilder();
        public List<String> itemsDataGrid = new List<String>();
        public Clientes clientes = new Clientes();
        public Fornecedores fornecedores = new Fornecedores();
        public Produtos produtos = new Produtos();

        private void button1_Click(object sender, EventArgs e)
        {
            if(comboBox1.SelectedItem.ToString() == "D_Clientes")
            {
                clientes.geraCliente(filesAdionado, MyApp, caminho, directoryPath, nomeSheet, excelConnectionString, colunas, colunasCreate, itemsDataGrid, dataGridView1);
            }
            if(comboBox1.SelectedItem.ToString() == "D_Fornecedores")
            {
                fornecedores.geraFornecedores(filesAdionado, MyApp, caminho, directoryPath, nomeSheet, excelConnectionString, colunas, colunasCreate, itemsDataGrid, dataGridView1);
            }
            if(comboBox1.SelectedItem.ToString() == "D_Produtos")
            {
                produtos.geraProdutos(filesAdionado, MyApp, caminho, directoryPath, nomeSheet, excelConnectionString, colunas, colunasCreate, itemsDataGrid, dataGridView1);
            }
        }

        private void buttonAbrir_Click(object sender, EventArgs e)
        {

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "C:\\a\\";
            openFileDialog1.Filter = "Csv files (*.csv*)|*.csv*|Excel files (*.xls*)|*.xls*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.Multiselect = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        using (myStream)
                        {
                            caminho = openFileDialog1.FileName;
                            directoryPath = Path.GetDirectoryName(openFileDialog1.FileName);
                            files = (openFileDialog1.SafeFileNames);
                            foreach (string file in files)
                            {
                                filesAdionado.Add(file);
                                listBox1.Items.Add(file);
                            }
                            carregaLinhas();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        public void carregaLinhas()
        {
            SqlConnection conn = new SqlConnection(@"Data Source=BRCAENRODRIGUES\SQLEXPRESS; Initial Catalog=my_database; Integrated Security=True");

            SqlCommand cmdCampos = conn.CreateCommand();

            cmdCampos.CommandText = @"IF OBJECT_ID('dbo.[campos]', 'U') IS NOT NULL 
                  DROP TABLE dbo.[campos]
                 CREATE TABLE [dbo].[campos](
	                                [campo_excel] [varchar](70) NULL,
	                                [campo_sql] [varchar](70) NULL,
                                    [tipo] [varchar](70) NULL,
                                    [tabela] [varchar](70) NULL)";

            SqlTransaction trA = null;

            conn.Open();
            trA = conn.BeginTransaction();
            cmdCampos.Transaction = trA;
            cmdCampos.ExecuteNonQuery();
            trA.Commit();
            conn.Close();

            label2.Text = caminho;

            MyApp = new Excel.Application();
            excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + caminho + ";Extended Properties=Excel 12.0;";
            MyApp.Workbooks.Add("");
            MyApp.Workbooks.Add(caminho);
            SqlTransaction trAx = null;


                for (int i = 1; i <= MyApp.Workbooks[2].Worksheets.Count; i++)
                {
                    comboBox2.Items.Add(MyApp.Workbooks[2].Worksheets[i].Name);
                }
             
            for (int k = 1; k <= MyApp.Workbooks[2].Worksheets[1].UsedRange.Columns.Count; k++)
            {
                if (Convert.ToString((MyApp.Workbooks[2].Worksheets[1].Cells[1, k].Value2)) != null && Convert.ToString(MyApp.Workbooks[2].Worksheets[1].Cells[1, k].Value2.ToString()) != "")
                {
                    string coluna = Convert.ToString(MyApp.Workbooks[2].Worksheets[1].Cells[1, k].Value2);
                    colunas.Add(coluna);
                    colunasCreate.Add(coluna.Trim());

                    cmdCampos.CommandText = "INSERT INTO CAMPOS (CAMPO_EXCEL) VALUES ('" + coluna + "');";
                    conn.Open();
                    trAx = conn.BeginTransaction();
                    cmdCampos.Transaction = trAx;
                    cmdCampos.ExecuteNonQuery();
                    trAx.Commit();
                    conn.Close();
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(@"Data Source=BRCAENRODRIGUES\SQLEXPRESS; Initial Catalog=my_database; Integrated Security=True");

            SqlCommand cmdCampos = conn.CreateCommand();

            cmdCampos.CommandText = @"IF OBJECT_ID('dbo.[campos]', 'U') IS NOT NULL 
                  DROP TABLE dbo.[campos]
                 CREATE TABLE [dbo].[campos](
	                                [campo_excel] [varchar](70) NULL,
	                                [campo_sql] [varchar](70) NULL,
                                    [tipo] [varchar](70) NULL,
                                    [tabela] [varchar](70) NULL)";

            SqlTransaction trA = null;

            conn.Open();
            trA = conn.BeginTransaction();
            cmdCampos.Transaction = trA;
            cmdCampos.ExecuteNonQuery();
            trA.Commit();
            conn.Close();
          

        }

        private void dataGridView2_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(@"Data Source=BRCAENRODRIGUES\SQLEXPRESS; Initial Catalog=my_database; Integrated Security=True");

            string connectionString = "Data Source=BRCAENRODRIGUES\\SQLEXPRESS;Initial Catalog=my_database;Integrated Security=True";

            string sql = "select campo_descr as Campos_SQL, min(campo_excel) as Campos_Excel " +
                         "from I_MAP a left join campos on campo_descr like '%' + campo_excel + '%' " +
                         "where a.tabela = '" + comboBox1.SelectedItem.ToString() + "' "+
                          "group by CAMPO_DESCR, ordem " +
                          "order by ordem ";

            SqlConnection connection = new SqlConnection(connectionString);
            SqlDataAdapter dataadapter = new SqlDataAdapter(sql, connection);
            DataSet ds = new DataSet();
            connection.Open();
            dataadapter.Fill(ds, "campos");
            connection.Close();
            dataGridView1.DataSource = ds;
            dataGridView1.DataMember = "campos";
        }

        static System.Data.DataTable ConvertListToDataTable(List<string> list)
        {
            // New table.
            System.Data.DataTable table = new System.Data.DataTable();

            // Add columns.
            for (int i = 0; i < 1; i++)
            {
                table.Columns.Add();
                table.Columns[0].ColumnName = "Campos Excel";
            }

            // Add rows.
            foreach (var array in list)
            {
                table.Rows.Add(array);
            }

            return table;
        }

        private void dataGridView1_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            dataGridView2.DoDragDrop(dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString(), DragDropEffects.Copy);
        }

        private void dataGridView1_DragDrop_1(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(System.String)))
            {
                System.Drawing.Point clientPoint = dataGridView1.PointToClient(new System.Drawing.Point(e.X, e.Y));
                dataGridView1.Rows[dataGridView1.HitTest(clientPoint.X, clientPoint.Y).RowIndex].Cells[dataGridView1.HitTest(clientPoint.X, clientPoint.Y).ColumnIndex].Value = (System.String)e.Data.GetData(typeof(System.String));
            }
        }

        private void dataGridView1_DragEnter_1(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(System.String)))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
          
            colunas.Clear();
 
            for (int k = 1; k <= MyApp.Workbooks[2].Worksheets[comboBox2.SelectedIndex + 1].UsedRange.Columns.Count; k++)
            {
                if (Convert.ToString((MyApp.Workbooks[2].Worksheets[comboBox2.SelectedIndex + 1].Cells[1, k].Value2)) != null && Convert.ToString(MyApp.Workbooks[2].Worksheets[comboBox2.SelectedIndex + 1].Cells[1, k].Value2.ToString()) != "")
                {
                    string coluna = Convert.ToString(MyApp.Workbooks[2].Worksheets[comboBox2.SelectedIndex + 1].Cells[1, k].Value2);
                    colunas.Add(coluna);
                    colunasCreate.Add(coluna.Trim());
                }
            }

            System.Data.DataTable table = ConvertListToDataTable(colunas);
            dataGridView2.DataSource = table;
            label5.Text = comboBox2.SelectedItem.ToString();

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
    }

}
