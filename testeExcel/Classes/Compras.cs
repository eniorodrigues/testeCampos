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
   public class Compras
    {
        public void geraCompras(List<string> filesAdionado, Excel.Application MyApp, string caminho, string directoryPath, string nomeSheet, string excelConnectionString, List<string> colunas, List<string> colunasCreate, List<String> itemsDataGrid, DataGridView dataGridView1, SqlConnection conn, int sheet)
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

               // MyApp.Range["A:AZ"].NumberFormat = "@";
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
                        if (dataGridView1.Rows[a].Cells[1].Value.ToString() != "" && dataGridView1.Rows[a].Cells[1].Value.ToString() != "Cmp_ID")
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
                      "IF OBJECT_ID('dbo.Compras', 'U') IS NOT NULL " +
                          "DROP TABLE dbo.Compras; " +
                            "CREATE TABLE [dbo].[Compras](" +
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
                            if (itemsDataGridInsert[f] == "Cmp_Pro_ID")
                            {
                                camposTabelaInsert.Append(" isnull([" +Convert.ToString(itemsDataGrid[f]).Replace('.', '#') +" ], 0), " );
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_For_ID")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Vinculo")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Paraiso_Fiscal")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_CNPJ")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_NF_Entrada")
                            {
                                 camposTabelaInsert.Append(" isnull([" +Convert.ToString(itemsDataGrid[f]).Replace('.', '#') +" ], 0), " );
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_NF_Serie")
                            {
                                 camposTabelaInsert.Append(" isnull([" +Convert.ToString(itemsDataGrid[f]).Replace('.', '#') +" ], 0), " );
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Cod_Divisao")
                            {
                                 camposTabelaInsert.Append(" isnull([" +Convert.ToString(itemsDataGrid[f]).Replace('.', '#') +" ], 0), " );
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_DI_ID")
                            {
                                 camposTabelaInsert.Append(" isnull([" +Convert.ToString(itemsDataGrid[f]).Replace('.', '#') +" ], 0), " );
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Incoterm")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Lanc_Cont")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Fat_Coml")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Tipo_Moeda")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_For_id_Frete")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Vinculo_Frete")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Paraiso_Fiscal_Frete")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_For_id_Seguro")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Vinculo_Seguro")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Paraiso_Fiscal_Seguro")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Valida")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Qtde_NF")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Valor_Fob_Original")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Siscomex")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_AFRMM")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Despesas")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_NF_DT")
                            {
                               camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] != '0' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] <> '' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] not like  '%N/A%' then(convert(datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)) else NULL end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_DI_DT_Emissao")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] != '0' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] <> '' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] not like  '%N/A%' then(convert(datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)) else NULL end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_BL_DT")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] != '0' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] <> '' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] not like  '%N/A%' then(convert(datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)) else NULL end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_DI_DT_Emissao_Corrigida")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] != '0' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] <> '' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] not like  '%N/A%' then(convert(datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)) else NULL end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_DI_DT_Vencimento")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] != '0' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] <> '' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] not like  '%N/A%' then(convert(datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)) else NULL end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Qtde")
                            {
                                  camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Valor_Fob")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Cambio_DT_DI")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_VL_Cred_Note")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Vl_Reais_FOB")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Par_US")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Valor_US")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Unit_US")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Vl_Frete_Moeda")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Cambio_Frete")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_VL_Frete_Reais")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_VL_Seguro_Moeda")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Cambio_Seguro")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_VL_Seguro_Reais")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Imposto_Import")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Vl_Presente")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Desconto_Incondicional")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_ICMS")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_PIS")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_COFINS")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_IPI")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Vl_Reais_12715")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Tx_Libor_Selic")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Ajuste")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Vl_Presente_FOB")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Vl_Presente_12715")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Vl_PCI_Reais")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_VL_CPL_Moeda")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Cambio_CPL")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Vl_Reais_CPL")
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
                            if (itemsDataGridInsert[f] == "Cmp_Pro_ID")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_For_ID")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Vinculo")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Paraiso_Fiscal")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_CNPJ")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_NF_Entrada")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_NF_Serie")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Cod_Divisao")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_DI_ID")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Incoterm")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Lanc_Cont")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Fat_Coml")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Tipo_Moeda")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_For_id_Frete")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Vinculo_Frete")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Paraiso_Fiscal_Frete")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_For_id_Seguro")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Vinculo_Seguro")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Paraiso_Fiscal_Seguro")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Valida")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Qtde_NF")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Valor_Fob_Original")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Siscomex")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_AFRMM")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Despesas")
                            {
                                camposTabelaInsert.Append(" isnull([" + Convert.ToString(itemsDataGrid[f]).Replace('.', '#') + " ], 0), ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_NF_DT")
                            {
                               camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] != '0' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] <> '' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] not like  '%N/A%' then(convert(datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)) else NULL end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_DI_DT_Emissao")
                            {
                               camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] != '0' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] <> '' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] not like  '%N/A%' then(convert(datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)) else NULL end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_BL_DT")
                            {
                               camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] != '0' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] <> '' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] not like  '%N/A%' then(convert(datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)) else NULL end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_DI_DT_Emissao_Corrigida")
                            {
                               camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] != '0' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] <> '' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] not like  '%N/A%' then(convert(datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)) else NULL end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_DI_DT_Vencimento")
                            {
                               camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] != '0' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] <> '' and [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] not like  '%N/A%' then(convert(datetime, [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], 103)) else NULL end, ");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Qtde")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Valor_Fob")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Cambio_DT_DI")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_VL_Cred_Note")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Vl_Reais_FOB")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Par_US")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Valor_US")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Unit_US")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Vl_Frete_Moeda")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Cambio_Frete")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_VL_Frete_Reais")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_VL_Seguro_Moeda")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Cambio_Seguro")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_VL_Seguro_Reais")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Imposto_Import")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Vl_Presente")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Desconto_Incondicional")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_ICMS")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_PIS")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_COFINS")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_IPI")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Vl_Reais_12715")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Tx_Libor_Selic")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Ajuste")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Vl_Presente_FOB")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Vl_Presente_12715")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Vl_PCI_Reais")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_VL_CPL_Moeda")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Cambio_CPL")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else if (itemsDataGridInsert[f] == "Cmp_Vl_Reais_CPL")
                            {
                                camposTabelaInsert.Append(" case when [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "] ='-' then 0 else isnull(convert(decimal(12,4),replace(replace(replace(replace([" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], '.',''), ',' , '.'),')',''),'(','-')),0) end, ID");
                            }
                            else
                            {
                                camposTabelaInsert.Append(" [" + Convert.ToString(itemsDataGrid[f]).Replace(".", "#") + "], ID ");
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
                        sqlBulk.DestinationTableName = "Compras";
                        sqlBulk.WriteToServer(dReader);
                    }
                    SqlCommand cmdCopPedido = conn.CreateCommand();

                    cmdCopPedido.CommandText =
                   " INSERT INTO [dbo].[D_Compras] ( " + camposTabelaSQL + " ) " +
                   " SELECT " + camposTabelaInsert +
                   " FROM [dbo].[Compras] ";
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
                        MessageBox.Show("Tabela Comparas Copiada ");
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
