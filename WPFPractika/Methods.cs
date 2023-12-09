using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;
using LiveCharts.Wpf.Charts.Base;
using LiveCharts.Wpf;
using LiveCharts;
using System.Data.SqlClient;

namespace WPFPractika
{
    internal class Methods
    {
        public static void ExportExcel(DataGrid grid,string TableName)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Application.Workbooks.Add(Type.Missing);
            excelApp.Cells.Range[excelApp.Cells[1, 1], excelApp.Cells[1, grid.Columns.Count]].Merge();
            excelApp.Cells[1, 1] = TableName;           
            for (int i = 0; i < grid.Columns.Count; i++)
            {
                excelApp.Cells[2, i + 1] = grid.Columns[i].Header;
                excelApp.Cells[2, i + 1].Borders.Value = Excel.XlLineStyle.xlContinuous;
            }
            for (int i = 0; i < grid.Items.Count; i++)
            {
                DataRowView row = (DataRowView)grid.Items[i];
                for (int j = 0; j < grid.Columns.Count; j++)
                {
                    excelApp.Cells[i + 3, j +1] = row[j+1].ToString();
                    excelApp.Cells[i + 3, j +1].Borders.Value = Excel.XlLineStyle.xlContinuous;
                }
            }
            excelApp.Cells[grid.Items.Count + 4, 1] = "Cоставил:";
            excelApp.Columns.AutoFit();
            excelApp.Columns.Style.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
            excelApp.Visible = true;
        }
        public static void Search(DataGrid data, TextBox text)
        {
            if (text.Text.Length > 0)
            {
                for (int i = 0; i < data.Items.Count; i++)
                {
                    DataRowView row = (DataRowView)data.Items[i];
                    for (int j = 1; j < data.Columns.Count; j++)
                    {
                        if (row[j].ToString().IndexOf(text.Text, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            ((DataGridRow)data.ItemContainerGenerator.ContainerFromIndex(i)).IsSelected = true;
                            break;
                        }
                        else
                        {
                            ((DataGridRow)data.ItemContainerGenerator.ContainerFromIndex(i)).IsSelected = false;
                        }
                    }
                }
            }
            else
                data.UnselectAll();
        }
        public static decimal LoadFullPrice(DataGrid data)
        {
            decimal fullPrice = 0;
            for (int i = 0; i < data.Items.Count; i++)
            {
                PriceList row = (PriceList)data.Items[i];
                fullPrice += row.Count * row.Price;
            } 
            return fullPrice;
        }
        public static void LoadChart(LiveCharts.Wpf.CartesianChart chart)
        {
            chart.Series.Clear();            
            List<string> key = new List<string>();
            List<int> value = new List<int>();
            SqlCommand command = new SqlCommand("Select Concat(Trim(Name),' ',Trim(Surname),' ',Trim(Patronymic)) as 'FIO' from Patient", DBManager.DentistryDBConnetion);
            DBManager.ConnectOpen();
            SqlDataReader reader = command.ExecuteReader();
            while (reader.Read())
                key.Add(reader["FIO"].ToString());
            reader.Close();
            for (int i = 0; i < key.Count; i++)
            {
                SqlCommand countsql = new SqlCommand($"Select count(Record.Id) from Record" +
                    $" inner join Patient on Patient.Id = Record.IdPatient where Concat(Trim(Name),' ',Trim(Surname),' ',Trim(Patronymic)) = N'{key[i]}'", DBManager.DentistryDBConnetion);
                value.Add(int.Parse(countsql.ExecuteScalar().ToString()));
            }
            DBManager.ConnectClose();

            ChartValues<int> values = new ChartValues<int>();
            values.AddRange(value);
            chart.Series.Add(new ColumnSeries()
            {
                Values = values
            });
            chart.AxisX.Add(new LiveCharts.Wpf.Axis
            {
                Labels = key
            });
        }
    }
}

