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

namespace testeCampos
{
   public class Clientes
    {
        public void geraCliente(List<string> filesAdionado, Excel.Application MyApp, string caminho, string directoryPath, string nomeSheet, string excelConnectionString, List<string> colunas, List<string> colunasCreate, List<String> itemsDataGrid, DataGridView dataGridView1)
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

                    string arquivo = element;
                    string campos = Convert.ToString(comandoExcel);

                    for (int a = 0; a < dataGridView1.Rows.Count; a++)
                    {
                        itemsDataGrid.Add(dataGridView1.Rows[a].Cells[1].Value.ToString());
                    }

                    StringBuilder camposExcel = new StringBuilder();
                    for (int f = 0; f < dataGridView1.Rows.Count; f++)
                    {
                        if (f == itemsDataGrid.Count - 1)
                        {
                            camposExcel.Append("[" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ");
                        }
                        else
                        {
                            camposExcel.Append("[" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ");
                        }
                    }

                    MessageBox.Show(camposExcel.ToString());

                    OleDbCommand command = new OleDbCommand
                    ("Select " + camposExcel + "  FROM [clientes$]", connection);

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
                        MessageBox.Show("Tabela clientes copiada ");
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
    }
}
