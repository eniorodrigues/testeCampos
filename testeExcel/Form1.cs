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
        private void button1_Click(object sender, EventArgs e)
        {

            foreach (string element in filesAdionado)
            {

                MyApp = new Excel.Application();
                MyApp.Workbooks.Add(caminho);
                Workbook wb = MyApp.Workbooks.Add(caminho);
                Worksheet ws = wb.Sheets[1];
                MyApp.DisplayAlerts = false;
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(caminho);
                Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1]; // assume it is the first sheet
                int columnCount = xlWorksheet.UsedRange.Columns.Count;
                List<string> columnNames = new List<string>();
                 
                for (int c = 1; c < columnCount; c++)
                {
                    if (xlWorksheet.Cells[1, c].Value2 != null)
                    {
                        string columnName = xlWorksheet.Columns[c].Address;
                        Regex reg = new Regex(@"(\$)(\w*):");
                        if (reg.IsMatch(columnName))
                        {
                            Match match = reg.Match(columnName);
                            columnNames.Add(match.Groups[2].Value);

                            if (xlWorksheet.Cells[1, c].Value2.Contains("digo"))
                            {
                                ws.Range[match.Groups[2].Value + ":" + match.Groups[2].Value].NumberFormat = "@";
                            }
                            if (xlWorksheet.Cells[1, c].Value2.Contains("CNPJ"))
                            {
                                ws.Range[match.Groups[2].Value + ":" + match.Groups[2].Value].EntireColumn.NumberFormat = "General";
                            }
                            if (xlWorksheet.Cells[1, c].Value2.Contains("data"))
                            {
                                ws.Range[match.Groups[2].Value + ":" + match.Groups[2].Value].Replace(".", "/");
                            }

                        }
                    }
                }
                xlApp.Quit();

                //MyApp.Range["A:A"].NumberFormat = "@";
                //MyApp.Range["B:B"].NumberFormat = "@";

                wb.SaveAs(directoryPath + "\\" + System.IO.Path.GetFileNameWithoutExtension(element) + "-(formatado).xlsx");
                wb.Close();

                SqlConnection conn = new SqlConnection(@"Data Source=BRCAENRODRIGUES\SQLEXPRESS; Initial Catalog=my_database; Integrated Security=True");
                string sqlConnectionString = "Data Source=BRCAENRODRIGUES\\SQLEXPRESS;Initial Catalog=my_database;Integrated Security=True";

                for (int i = 1; i <= MyApp.Workbooks.Count; i++)
                {
                    for (int j = 1; j <= MyApp.Workbooks[i].Worksheets.Count; j++)
                    {
                        nomeSheet = MyApp.Workbooks[1].Sheets[1].Name.ToString();
                    }
                }

                SqlCommand cmdColuna = conn.CreateCommand();

                cmdColuna.CommandText =
                  "IF OBJECT_ID('dbo.clientes', 'U') IS NOT NULL " +
                      "DROP TABLE dbo.clientes; " +
                        "CREATE TABLE [dbo].[Clientes](" +
                            "[Cli_ID] [varchar](70) NULL," +
                            "[Cli_Nome] [varchar](255) NULL," +
                            "[Cli_Pss_ID] [int] NULL," +
                            "[Cli_Vinc] [varchar](1) NULL," +
                            "[Cli_Vinc_DT_Ini] [datetime] NULL," +
                            "[Cli_Vinc_DT_Fim] [datetime] NULL," +
                            "[Cli_CNPJ] [varchar](40) NULL," +
                            "[Cli_Vinc_Justific] [varchar](2) NULL," +
                            "[Cli_Paraiso_Fiscal] [varchar](1) NULL CONSTRAINT [DF_Clientes_Cli_Paraiso_Fiscal]  DEFAULT ('N')," +
                            "[Arq_Origem_ID] [int] NULL," +
                            "[ID] [int] IDENTITY(1,1) NOT NULL" +
                        ") ON [PRIMARY]";

                SqlTransaction trA = null;

                conn.Open();
                trA = conn.BeginTransaction();
                cmdColuna.Transaction = trA;
                cmdColuna.ExecuteNonQuery();
                trA.Commit();
                conn.Close();

                excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + directoryPath + "\\" + System.IO.Path.GetFileNameWithoutExtension(element) + "-(formatado).xlsx; Extended Properties=Excel 12.0;";

                //for (int i = 1; i <= MyApp.Workbooks.Count; i++)
                //{
                //    for (int j = 1; j <= MyApp.Workbooks[i].Worksheets.Count; j++)
                //    {

                        using (OleDbConnection connection = new OleDbConnection(excelConnectionString))
                        {

                            StringBuilder comandoExcel = new StringBuilder();

                            for (int h = 0; h < colunas.Count; h++)
                            {
                                if (h == colunas.Count - 1)
                                {
                                    comandoExcel.Append("[" + Convert.ToString(colunas[h]).Replace(".", "#") + "] ");
                                }
                                else
                                {
                                    comandoExcel.Append("[" + Convert.ToString(colunas[h]).Replace(".", "#") + "], ");
                                }
                            }

                            //string sheet = MyApp.Workbooks[i].Worksheets[j].name;
                            string arquivo = element;
                            string campos = Convert.ToString(comandoExcel);

                            //List<string> items = new List<string>();
                            //foreach (DataGridViewRow dr in dataGridView1.Rows)
                            //{
                            //    string item = dr.Cells[1].ToString();
                            //    foreach (DataGridViewCell dc in dr.Cells)
                            //    {
                            //        items.Add(item);
                            //        MessageBox.Show(item.ToString());
                            //    }
                            //}

                            StringBuilder camposExcel = new StringBuilder();
                            for (int f = 0; f < dataGridView1.Rows.Count -1; f++)
                            {
        
                                    if (f == colunas.Count - 1)
                                    {
                                    camposExcel.Append("[" + Convert.ToString(colunas[f]).Replace(".", "#") + "] ");
                                    }
                                    else
                                    {
                                    camposExcel.Append("[" + Convert.ToString(colunas[f]).Replace(".", "#") + "], ");
                                    }
                            }

                           // MessageBox.Show(camposExcel.ToString());


                            OleDbCommand command = new OleDbCommand
                            ("Select "+ camposExcel + "  FROM [clientes$]", connection);

                            connection.Open();

                            OleDbDataReader dReader = command.ExecuteReader();

                            using (SqlBulkCopy sqlBulk = new SqlBulkCopy(sqlConnectionString))
                            {
                                sqlBulk.DestinationTableName = "Clientes";
                                sqlBulk.WriteToServer(dReader);
                            }
                             
                    SqlCommand cmdCopPedido = conn.CreateCommand();
                    cmdCopPedido.CommandText =
                        @"INSERT INTO D_CLIENTES (CLI_ID, CLI_NOME, CLI_VINC, CLI_PSS_ID, [Cli_Vinc_DT_Ini], [Cli_Vinc_DT_Fim], [Cli_CNPJ], Lin_Origem_id)
                        SELECT CLI_ID, max(CLI_NOME), CLI_VINC, CLI_PSS_ID, [Cli_Vinc_DT_Ini], [Cli_Vinc_DT_Fim], max([Cli_CNPJ]), max(id)
                        FROM clientes
                        where cli_id is not null and cli_vinc is not null
                        GROUP BY CLI_ID, CLI_VINC, CLI_PSS_ID, [Cli_Vinc_DT_Ini], [Cli_Vinc_DT_Fim]";
                    SqlTransaction tr = null;
                    try
                    {
                        conn.Open();
                        tr = conn.BeginTransaction();
                        cmdCopPedido.Transaction = tr;
                        cmdCopPedido.ExecuteNonQuery();
                        tr.Commit();
                        label1.Text = "Tabela clientes copiada ";
                    }
                    catch (Exception ex)
                    {
                        tr.Rollback();
                        MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                        conn.Close();
                    }

                }
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
             
           // MessageBox.Show(MyApp.Workbooks[2].Worksheets.Count.ToString());

            if(MyApp.Workbooks[2].Worksheets.Count == 1)
            {
                for (int i = 0; i < MyApp.Workbooks[2].Worksheets.Count; i++)
                {
                    i++;
                    comboBox2.Items.Add(MyApp.Workbooks[2].Worksheets[i].Name);
                }
            }
            else
            {
                for (int i = 0; i < MyApp.Workbooks[2].Worksheets.Count; i++)
                {
                    comboBox2.Items.Add(MyApp.Workbooks[2].Worksheets[i].Name);
                }
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

        }

        private void dataGridView2_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(@"Data Source=BRCAENRODRIGUES\SQLEXPRESS; Initial Catalog=my_database; Integrated Security=True");

            string connectionString = "Data Source=BRCAENRODRIGUES\\SQLEXPRESS;Initial Catalog=my_database;Integrated Security=True";

            string sql = "select ordem, campo_descr as Campos_SQL, min(campo_excel) as Campos_Excel " +
                         "from I_MAP a left join campos on campo_descr like '%' + campo_excel + '%' " +
                         "where a.tabela = 'd_clientes' "+
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
            dataGridView1.Columns.Add("Campos", "Campos");
        }

        //private void button2_Click(object sender, EventArgs e)
        //{
        //    //dataGridView1.DataSource = null;
        //    //System.Data.DataTable table = ConvertListToDataTable(colunas);
        //    //dataGridView2.DataSource = table;
        //    //label5.Text = comboBox2.SelectedItem.ToString();
        //}

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
            // dataGridView1.DataSource = null;
            System.Data.DataTable table = ConvertListToDataTable(colunas);
            dataGridView2.DataSource = table;
            label5.Text = comboBox2.SelectedItem.ToString();

            colunas.Clear();
            for (int k = 1; k <= MyApp.Workbooks[2].Worksheets[comboBox2.SelectedIndex + 1].UsedRange.Columns.Count; k++)
            {
                if (Convert.ToString((MyApp.Workbooks[2].Worksheets[comboBox2.SelectedIndex + 1].Cells[1, k].Value2)) != null && Convert.ToString(MyApp.Workbooks[2].Worksheets[comboBox2.SelectedIndex + 1].Cells[1, k].Value2.ToString()) != "")
                {
                    string coluna = Convert.ToString(MyApp.Workbooks[2].Worksheets[comboBox2.SelectedIndex + 1].Cells[1, k].Value2);
                    colunas.Add(coluna);
                    colunasCreate.Add(coluna.Trim());

                    //cmdCampos.CommandText = "INSERT INTO CAMPOS (CAMPO_EXCEL) VALUES ('" + coluna + "');";
                    //conn.Open();
                    //trAx = conn.BeginTransaction();
                    //cmdCampos.Transaction = trAx;
                    //cmdCampos.ExecuteNonQuery();
                    //trAx.Commit();
                    //conn.Close();
                }
            }

           
        }
    }

}
