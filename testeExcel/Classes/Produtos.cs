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
   public class Produtos
    {
        public void geraProdutos(List<string> filesAdionado, Excel.Application MyApp, string caminho, string directoryPath, string nomeSheet, string excelConnectionString, List<string> colunas, List<string> colunasCreate, List<String> itemsDataGrid, DataGridView dataGridView1, SqlConnection conn)
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

                wb.SaveAs(directoryPath + "\\" + System.IO.Path.GetFileNameWithoutExtension(element) + "-(formatado).xlsx");
                wb.Close();
                
               // SqlConnection conn = new SqlConnection(@"Data Source=BRCAENRODRIGUES\SQLSERVER; Initial Catalog=MANN_2017; Integrated Security=True");
               // string sqlConnectionString = "Data Source=BRCAENRODRIGUES\\SQLSERVER;Initial Catalog=MANN_2017;Integrated Security=True";
               
                for (int i = 1; i <= MyApp.Workbooks.Count; i++)
                {
                    for (int j = 1; j <= MyApp.Workbooks[i].Worksheets.Count; j++)
                    {
                        nomeSheet = MyApp.Workbooks[1].Sheets[1].Name.ToString();
                    }
                }

                SqlCommand cmdColuna = conn.CreateCommand();

                cmdColuna.CommandText =
                  "IF OBJECT_ID('dbo.produtos', 'U') IS NOT NULL " +
                      "DROP TABLE dbo.produtos; " +
                        "CREATE TABLE[dbo].[produtos](" +
                        "[Pro_ID][varchar](70) NULL," +
                        "[Pro_Descricao] [varchar] (255) NULL," +
	                    "[Pro_Und_ID] [int] NULL," +
	                    "[Pro_NCM] [varchar] (8) NULL," +
	                    "[Pro_Margem] [int] NULL," +
	                    "[Pro_Perc_Quebras_Perdas] [numeric] (24, 12) NULL, " +
                        "[id] [int] IDENTITY(1,1) NOT NULL, " +
	                    "[Arq_Origem_ID] [int] NULL)"; 

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

                    MessageBox.Show(comandoExcel.ToString());

                    string arquivo = element;
                    string campos = Convert.ToString(comandoExcel);

                    for (int a = 0; a < dataGridView1.Rows.Count; a++)
                    {
                        if(dataGridView1.Rows[a].Cells[1].Value.ToString() != "")
                        {
                            itemsDataGrid.Add(dataGridView1.Rows[a].Cells[1].Value.ToString());
                        }
                    }

                    StringBuilder camposExcel = new StringBuilder();
                    for (int f = 0; f < itemsDataGrid.Count; f++)
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

                    OleDbCommand command = new OleDbCommand
                    ("Select " + camposExcel + "  FROM [Produtos$]", connection);

                    connection.Open();

                    OleDbDataReader dReader = command.ExecuteReader();
                    conn.Open();
                    using (SqlBulkCopy sqlBulk = new SqlBulkCopy(conn))
                    {
                        sqlBulk.DestinationTableName = "Produtos";
                        sqlBulk.WriteToServer(dReader);
                    }

                    SqlCommand cmdCopPedido = conn.CreateCommand();
                    cmdCopPedido.CommandText =
                        @"INSERT INTO D_PRODUTOS (PRO_ID, Pro_Descricao, Pro_Und_ID, Pro_NCM, Pro_Margem,  Lin_Origem_ID)
                        SELECT PRO_ID, MAX(Pro_Descricao), max(Pro_Und_ID), MAX(Pro_NCM), MAX(Pro_Margem), max(id)
                        FROM produtos
                        WHERE PRO_ID is not null and Pro_Descricao is not null
                        GROUP BY  PRO_ID ";
                    SqlTransaction tr = null;

                    try
                    {
                        //conn.Open();
                        tr = conn.BeginTransaction();
                        cmdCopPedido.Transaction = tr;
                        cmdCopPedido.ExecuteNonQuery();
                        tr.Commit();
                        MessageBox.Show("Tabela produtos copiada ");
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
