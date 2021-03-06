﻿using System;
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
   public class ClientesTeste
    {
        public void geraClienteTeste(List<string> filesAdionado, Excel.Application MyApp, string caminho, string directoryPath, string nomeSheet, string excelConnectionString, List<string> colunas, List<string> colunasCreate, List<String> itemsDataGrid, DataGridView dataGridView1, SqlConnection conn)
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
                        if (dataGridView1.Rows[a].Cells[1].Value.ToString() != "")
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

                    //temporario
                    conn = new SqlConnection(@"Data Source=BRCAENRODRIGUES\SQLSERVER; Initial Catalog=MANN_2017; Integrated Security=True");

                    SqlCommand cmdColuna = conn.CreateCommand();
                    cmdColuna.CommandText =
                      "IF OBJECT_ID('dbo.clientes', 'U') IS NOT NULL " +
                      "DROP TABLE dbo.clientes; " +
                        "CREATE TABLE[dbo].[clientes](" +
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
                            if (itemsDataGridInsert[f] == "Cli_Vinc_DT_Ini")
                            {
                                camposTabelaInsert.Append(" max (convert( datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cli_Vinc_DT_Fim")
                            {
                                camposTabelaInsert.Append(" max (convert( datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cli_Pss_ID")
                            {
                                camposTabelaInsert.Append(" max (convert(int, replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ',' , '.'))), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cli_Vinc")
                            {
                                camposTabelaInsert.Append(" max (replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ',' , '.')), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cli_Vinc_Justific")
                            {
                                camposTabelaInsert.Append(" max (replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ',' , '.')), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cli_Paraiso_Fiscal")
                            {
                                camposTabelaInsert.Append(" max (replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ',' , '.')), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cli_CNPJ")
                            {
                                camposTabelaInsert.Append(" max (isnull([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "],0)) , ");
                            }
                            else if (itemsDataGridInsert[f] == "Cli_Nome")
                            {
                                camposTabelaInsert.Append(" max([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "]), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cli_ID")
                            {
                                camposTabelaInsert.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] , ");
                                camposWhere.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ");
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
                            if (itemsDataGridInsert[f] == "Cli_Vinc_DT_Ini")
                            {
                                camposTabelaInsert.Append(" max (convert( datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)), max(ID) ");
                            }
                            else if (itemsDataGridInsert[f] == "Cli_Vinc_DT_Fim")
                            {
                                camposTabelaInsert.Append(" max(convert( datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)), max(ID) ");
                            }
                            else if (itemsDataGridInsert[f] == "Cli_Pss_ID")
                            {
                                camposTabelaInsert.Append(" max (convert(int, replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ',' , '.'))), max(ID) ");
                            }
                            else if (itemsDataGridInsert[f] == "Cli_Vinc")
                            {
                                camposTabelaInsert.Append(" max(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ',' , '.')), max(ID) ");
                            }
                            else if (itemsDataGridInsert[f] == "Cli_Vinc_Justific")
                            {
                                camposTabelaInsert.Append(" max(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ',' , '.')), max(ID) ");
                            }
                            else if (itemsDataGridInsert[f] == "Cli_Paraiso_Fiscal")
                            {
                                camposTabelaInsert.Append(" max(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ',' , '.')), max(ID) ");
                            }
                            else if (itemsDataGridInsert[f] == "Cli_CNPJ")
                            {
                                camposTabelaInsert.Append(" max (isnull([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "],0)) , max(ID) ");
                            }
                            else if (itemsDataGridInsert[f] == "Cli_Nome")
                            {
                                camposTabelaInsert.Append(" max([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "]), max(ID) ");
                            }
                            else if (itemsDataGridInsert[f] == "Cli_ID")
                            {
                                camposTabelaInsert.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] , max(ID)");
                                camposWhere.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "]  ");
                                camposGroupBy.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ");
                            }
                            else
                            {
                                camposTabelaInsert.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] , max(ID) ");
                                camposGroupBy.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ID ");
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
                        sqlBulk.DestinationTableName = "clientes";
                        sqlBulk.WriteToServer(dReader);
                    }
                    SqlCommand cmdCopPedido = conn.CreateCommand();

                    cmdCopPedido.CommandText =
                   " INSERT INTO [dbo].[D_Clientes] ( " + camposTabelaSQL + " ) " +
                   " SELECT " + camposTabelaInsert +
                   " FROM [dbo].[Clientes] " +
                   " WHERE " + camposWhere + " IS NOT NULL " +
                   " GROUP BY " + camposGroupBy;

                    MessageBox.Show(" === " + cmdCopPedido.CommandText);
                    SqlTransaction tr = null;
                    conn.Close();

                    try
                    {
                        conn.Open();
                        tr = conn.BeginTransaction();
                        cmdCopPedido.Transaction = tr;
                        cmdCopPedido.ExecuteNonQuery();
                        tr.Commit();
                        MessageBox.Show("Tabela Clientes Copiada ");
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
