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
   public class Vendas
    {
        public void geraVendas(List<string> filesAdionado, Excel.Application MyApp, string caminho, string directoryPath, string nomeSheet, string excelConnectionString, List<string> colunas, List<string> colunasCreate, List<String> itemsDataGrid, DataGridView dataGridView1, SqlConnection conn, int sheet)
        {
            foreach (string element in filesAdionado)
            {
                MyApp = new Excel.Application();
                MyApp.Workbooks.Add(caminho);
                Workbook wb = MyApp.Workbooks.Add(caminho);
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

                MyApp.Range["A:AZ"].NumberFormat = "@";
                //MyApp.Range["B:B"].NumberFormat = "@";
                //MyApp.Range["C:C"].NumberFormat = "@";
                //MyApp.Range["D:D"].NumberFormat = "@";
                //MyApp.Range["E:E"].NumberFormat = "@";
                //MyApp.Range["F:F"].NumberFormat = "@";
                //MyApp.Range["G:G"].NumberFormat = "@";
                //MyApp.Range["H:H"].NumberFormat = "@";

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
                        if (dataGridView1.Rows[a].Cells[1].Value.ToString() != "" && dataGridView1.Rows[a].Cells[1].Value.ToString() != "Vnd_ID")
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
                      "IF OBJECT_ID('dbo.Vendas_Itens', 'U') IS NOT NULL " +
                          "DROP TABLE dbo.Vendas_Itens; " +
                            "CREATE TABLE [dbo].[Vendas_Itens](" +
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
                            if (itemsDataGridInsert[f] == "Vnd_Cli_ID")
                            {
                                camposTabelaInsert.Append(" isnull([" +Convert.ToString(itemsDataGrid[f]).Replace('.', '#') +" ], 0), " );
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_NF_ID")
                            {
                                 camposTabelaInsert.Append(" isnull([" +Convert.ToString(itemsDataGrid[f]).Replace('.', '#') +" ], 0), " );
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Dias")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] <> '' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] not like  '%N/A%' then [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] else NULL end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_CFOP")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_NF_Serie")
                            {
                                 camposTabelaInsert.Append(" isnull([" +Convert.ToString(itemsDataGrid[f]).Replace('.', '#') +" ], 0), " );
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Item")
                            {
                                 camposTabelaInsert.Append(" isnull([" +Convert.ToString(itemsDataGrid[f]).Replace('.', '#') +" ], 0), " );
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Pro_id")
                            {
                                 camposTabelaInsert.Append(" isnull([" +Convert.ToString(itemsDataGrid[f]).Replace('.', '#') +" ], 0), " );
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Vinculo")
                            {
                                 camposTabelaInsert.Append(" isnull([" +Convert.ToString(itemsDataGrid[f]).Replace('.', '#') +" ], 0), " );
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Paraiso_Fiscal")
                            {
                                 camposTabelaInsert.Append(" isnull([" +Convert.ToString(itemsDataGrid[f]).Replace('.', '#') +" ], 0), " );
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Dt_Emissao")
                            {
                                camposTabelaInsert.Append(" isnull((convert(datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)),''), ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_DT_Vencimento")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] <> '' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] not like  '%N/A%' then(convert(datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)) else NULL end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Item")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "],''), ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Qtde")
                            {
                                  camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Vl_Nota")
                            {
                                  camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Desconto")
                            {
                                  camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_ICMS")
                            {
                                  camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_PIS")
                            {
                                  camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_COFINS")
                            {
                                  camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_ISS")
                            {
                                  camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Comissao")
                            {
                                  camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Frete")
                            {
                                  camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Seguro")
                            {
                                  camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Vl_Moeda")
                            {
                                  camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Custo")
                            {
                                  camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Despesas")
                            {
                                  camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Deducoes")
                            {
                                  camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Ajuste")
                            {
                                  camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Vl_Presente")
                            {
                                  camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Tx_Libor_Selic")
                            {
                                  camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Cambio_Dt_Emissao")
                            {
                                  camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Cambio_Dt_Embarque")
                            {
                                  camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Dt_Embarque")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] <> '' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] not like  '%N/A%' then(convert(datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)) else NULL end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Vl_Reais")
                            {
                                  camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Cotacao")
                            {
                                  camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else
                            {
                                camposTabelaInsert.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ");
                            }
                        }
                        else
                        {
                            if (itemsDataGridInsert[f] == "Vnd_Cli_ID")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_NF_ID")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Dias")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] <> '' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + " not like '%N/A%' then [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] else NULL end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_CFOP")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_NF_Serie")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Item")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Pro_id")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Vinculo")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Paraiso_Fiscal")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Dt_Emissao")
                            {
                                camposTabelaInsert.Append(" isnull((convert(datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)),''), ID ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_DT_Vencimento")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] <> '' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] not like  '%N/A%' then(convert(datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)) else NULL end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Item")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "],''), ID ");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Qtde")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Vl_Nota")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Desconto")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_ICMS")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_PIS")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_COFINS")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_ISS")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Comissao")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Frete")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Seguro")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Vl_Moeda")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Custo")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Despesas")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Deducoes")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Ajuste")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Vl_Presente")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Tx_Libor_Selic")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Cambio_Dt_Emissao")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Cambio_Dt_Embarque")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Dt_Embarque")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] <> '' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] not like  '%N/A%' then(convert(datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)) else NULL end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Vl_Reais")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Vnd_Cotacao")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else
                            {
                                camposTabelaInsert.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ");
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
                        sqlBulk.DestinationTableName = "Vendas_Itens";
                        sqlBulk.WriteToServer(dReader);
                    }
                    SqlCommand cmdCopPedido = conn.CreateCommand();

                    cmdCopPedido.CommandText =
                   " INSERT INTO [dbo].[D_Vendas_Itens] ( " + camposTabelaSQL + " ) " +
                   " SELECT " + camposTabelaInsert +
                   " FROM [dbo].[Vendas_Itens] ";
                     //  " WHERE " + camposWhere;

                   MessageBox.Show(cmdCopPedido.CommandText);
                    Clipboard.SetText(cmdCopPedido.CommandText);
                    SqlTransaction tr = null;
                    conn.Close();

                    try
                    {
                        conn.Open();
                        tr = conn.BeginTransaction();
                        cmdCopPedido.Transaction = tr;
                        cmdCopPedido.ExecuteNonQuery();
                        tr.Commit();
                        MessageBox.Show("Tabela Vendas Copiada ");
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
