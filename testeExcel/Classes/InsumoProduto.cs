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
   public class InsumoProduto
    {
        public void geraInsumoProduto(List<string> filesAdionado, Excel.Application MyApp, string caminho, string directoryPath, string nomeSheet, string excelConnectionString, List<string> colunas, List<string> colunasCreate, List<String> itemsDataGrid, DataGridView dataGridView1, SqlConnection conn, int sheet)
        {
            foreach (string element in filesAdionado)
            {
                MyApp = new Excel.Application();
                MyApp.Workbooks.Add(caminho);
                Workbook wb = MyApp.Workbooks.Add(caminho);
                MessageBox.Show(sheet.ToString());
                Worksheet ws = wb.Sheets[sheet + 1];
                MyApp.DisplayAlerts = false;
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(caminho);
                Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[sheet + 1]; 
                int columnCount = xlWorksheet.UsedRange.Columns.Count;
                List<string> columnNames = new List<string>();
                int group = 0;

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

                //formata pra não ficar nulo e replace pra data com .
                //devera condicionar formatação pelo nome da coluna e tambem criar janela para
                //personalização da formatação

                MyApp.Range["A:A"].NumberFormat = "@";
                MyApp.Range["B:B"].NumberFormat = "@";
                MyApp.Range["C:C"].NumberFormat = "@";
                MyApp.Range["D:D"].NumberFormat = "@";
                MyApp.Range["E:E"].NumberFormat = "@";
                MyApp.Range["F:F"].NumberFormat = "@";
                MyApp.Range["G:G"].NumberFormat = "@";
                MyApp.Range["H:H"].NumberFormat = "@";

                wb.SaveAs(directoryPath + "\\" + System.IO.Path.GetFileNameWithoutExtension(element) + "-(formatado).xlsx");
                wb.Close();

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
                        if (dataGridView1.Rows[a].Cells[1].Value.ToString() != "")
                        {
                            itemsDataGrid.Add(dataGridView1.Rows[a].Cells[1].Value.ToString());
                        }
                    }

                    List<String> itemsDataGridInsert = new List<String>();

                    for (int a = 0; a < dataGridView1.Rows.Count; a++)
                    {
                        if (dataGridView1.Rows[a].Cells[1].Value.ToString() != "" && dataGridView1.Rows[a].Cells[1].Value.ToString() != "Ins_Id")
                        {
                            itemsDataGridInsert.Add(dataGridView1.Rows[a].Cells[0].Value.ToString());
                        }
                    }

                    StringBuilder camposTabela = new StringBuilder();
                    for (int f = 0; f < itemsDataGrid.Count; f++)
                    {
                        camposTabela.Append("[" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] [varchar](max) NULL, ");
                    }

                    StringBuilder camposTabelaSQL = new StringBuilder();
                    for (int f = 0; f < itemsDataGridInsert.Count; f++)
                    {
                        if (f == itemsDataGrid.Count - 1)
                        {
                            camposTabelaSQL.Append("[" + Convert.ToString(itemsDataGridInsert[f]).Replace(".", "#") + "], [Lin_Origem_ID] ");
                        }
                        else
                        {
                            camposTabelaSQL.Append("[" + Convert.ToString(itemsDataGridInsert[f]).Replace(".", "#") + "],  ");
                        }
                    }

                    SqlCommand cmdColuna = conn.CreateCommand();
                    cmdColuna.CommandText =
                      "IF OBJECT_ID('dbo.Insumo_Produto', 'U') IS NOT NULL " +
                          "DROP TABLE dbo.Insumo_Produto; " +
                            "CREATE TABLE [dbo].[Insumo_Produto](" +
                               camposTabela +
                                "[ID] [int] IDENTITY(1,1) NOT NULL)" +
                            " ON [PRIMARY]";

                    SqlTransaction trA = null;
                    conn.Open();
                    trA = conn.BeginTransaction();
                    cmdColuna.Transaction = trA;
                    cmdColuna.ExecuteNonQuery();
                    trA.Commit();
                    conn.Close();

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

                    StringBuilder camposTabelaInsert = new StringBuilder();
                    StringBuilder camposGroupBy = new StringBuilder();
                    StringBuilder camposWhere = new StringBuilder();

                    for (int f = 0; f < itemsDataGridInsert.Count; f++)
                    {
                        if (f < itemsDataGrid.Count - 1)
                        {
                            if (itemsDataGridInsert[f] == "Ins_PA_Pro_Id")
                            {
                                camposTabelaInsert.Append(" [" + Convert.ToString(itemsDataGrid[f])  + "], ");
                                camposWhere.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] IS NOT NULL AND ");
                                camposGroupBy.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ");
                            }
                            else if (itemsDataGridInsert[f] == "Ins_MP_Pro_Id")
                            {
                                camposTabelaInsert.Append(" [" + Convert.ToString(itemsDataGrid[f]) + "],  ");
                                camposWhere.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "]  IS NOT NULL ");
                                camposGroupBy.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ");
                            }
                            else if (itemsDataGridInsert[f] == "Ins_DT_Ini")
                            {
                                camposTabelaInsert.Append(" (convert( datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)), ");
                                camposGroupBy.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ");
                            }
                            else if (itemsDataGridInsert[f] == "Ins_DT_Fim")
                            {
                                camposTabelaInsert.Append(" (convert( datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)), ");
                                camposGroupBy.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ");
                            }
                            else if (itemsDataGridInsert[f] == "Ins_Qtd_Produzida")
                            {
                                camposTabelaInsert.Append(" (convert(decimal(12,4), replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ',' , '.'))), ");
                                camposGroupBy.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ");
                            }
                            else if (itemsDataGridInsert[f] == "Ins_Qtd_Requisitada")
                            {
                                camposTabelaInsert.Append(" (convert(decimal(12,4), replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ',' , '.'))), ");
                                camposGroupBy.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ");
                            }
                            else if (itemsDataGridInsert[f] == "Lin_Origem_ID")
                            {
                                camposTabelaInsert.Append(" [ " + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ");
                                camposGroupBy.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ");
                            }
                            else
                            {
                                camposTabelaInsert.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ");
                                camposGroupBy.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ");
                            }
                        }
                        else
                        {
                            //if (itemsDataGridInsert[f] == "Ins_PA_Pro_Id")
                            //{
                            //    camposTabelaInsert.Append(" [" + Convert.ToString(itemsDataGrid[f]) + "] , max(ID) ");
                            //    camposWhere.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "]  IS NOT NULL  ");
                            //    camposGroupBy.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ");
                            //}
                            //if (itemsDataGridInsert[f] == "Ins_MP_Pro_Id")
                            //{
                            //    camposTabelaInsert.Append(" [" + Convert.ToString(itemsDataGrid[f]) + "] , max(ID)");
                            //    camposWhere.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "]  IS NOT NULL ");
                            //    camposGroupBy.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ");
                            //}
                            if (itemsDataGridInsert[f] == "Ins_DT_Ini")
                            {
                                camposTabelaInsert.Append(" (convert( datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)), max(ID)");
                                camposGroupBy.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ");
                            }
                            else if (itemsDataGridInsert[f] == "Ins_DT_Fim")
                            {
                                camposTabelaInsert.Append(" (convert( datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)), max(ID)");
                                camposGroupBy.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ");
                            }
                            else if (itemsDataGridInsert[f] == "Ins_Qtd_Produzida")
                            {
                                camposTabelaInsert.Append("  (convert(decimal(12,4), replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ',' , '.'))), max(ID)");
                                camposGroupBy.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ");
                            }
                            else if (itemsDataGridInsert[f] == "Ins_Qtd_Requisitada")
                            {
                                camposTabelaInsert.Append("  (convert(decimal(12,4), replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ',' , '.'))), max(ID)");
                                camposGroupBy.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ");
                            }
                            else if (itemsDataGridInsert[f] == "Lin_Origem_ID")
                            {
                                camposTabelaInsert.Append(" [ " + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], max(ID) ");
                                camposGroupBy.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ");
                            }
                            else
                            {
                                camposTabelaInsert.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], max(ID) ");
                                camposGroupBy.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "]  ");
                            }
                        }
                    }
                    conn.Open();
                    OleDbCommand command = new OleDbCommand
                        
                    ("Select " + camposExcel + "  FROM [" + nomeSheet + "$]", connection);
                    connection.Open();
                    OleDbDataReader dReader = command.ExecuteReader();
                    using (SqlBulkCopy sqlBulk = new SqlBulkCopy(conn))
                    {
                        sqlBulk.DestinationTableName = "Insumo_Produto";
                        sqlBulk.WriteToServer(dReader);
                    }
                    SqlCommand cmdCopPedido = conn.CreateCommand();
 
                        cmdCopPedido.CommandText =
                       " INSERT INTO [dbo].[D_Insumo_Produto] ( " + camposTabelaSQL + " ) " +
                       " SELECT " + camposTabelaInsert +
                       " FROM [dbo].[Insumo_Produto] " +
                       " WHERE " + camposWhere  +
                       " GROUP BY " + camposGroupBy;

                //    MessageBox.Show(cmdCopPedido.CommandText);
                    SqlTransaction tr = null;
                        conn.Close();

                    try
                    {
                        conn.Open();
                        tr = conn.BeginTransaction();
                        cmdCopPedido.Transaction = tr;
                        cmdCopPedido.ExecuteNonQuery();
                        tr.Commit();
                        MessageBox.Show("Tabela Insumo Produto Copiada ");
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
