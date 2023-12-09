using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Xceed.Wpf.Toolkit;

namespace WPFPractika
{
    public partial class AddRecord : Window
    {
        string query;
        public AddRecord()
        {
            InitializeComponent();
        }

        //private void Window_Activated(object sender, EventArgs e)
        //{
            
        //}

        private void choosePriceList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            query = $"Select Price From PriceList where id = {choosePriceList.SelectedValue}";
            DBManager.ConnectOpen();
            SqlCommand command = new SqlCommand(query,DBManager.DentistryDBConnetion);
            infoPrice.Content = command.ExecuteScalar();
            DBManager.ConnectClose();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            bool isHave = false;
            int selectedRow = -1;
            if (choosePriceList.SelectedIndex != -1 && uint.TryParse(entryCount.Text,out uint value)) 
            {
                for (int i = 0; i < gridPriceListInfo.Items.Count; i++)
                {
                    PriceList row = (PriceList)gridPriceListInfo.Items[i];
                    for (int j = 0; j < gridPriceListInfo.Columns.Count; j++)
                    {
                        if (row.Id == Convert.ToInt32(choosePriceList.SelectedValue))
                        {
                            isHave= true;
                            selectedRow = i;
                            break;                            
                        }                            
                    }
                }
            if (isHave)
            {
                    PriceList row = (PriceList)gridPriceListInfo.Items[selectedRow];
                    row.Count += value;
                    gridPriceListInfo.Items[selectedRow] = new PriceList {Id=row.Id,Title=row.Title,Price=row.Price,Count=row.Count };
            }
                else
                gridPriceListInfo.Items.Add(new PriceList { Id = Convert.ToInt32(choosePriceList.SelectedValue), Title = choosePriceList.Text.Trim(), Price = Convert.ToDecimal(infoPrice.Content),Count=value });
                checkFullPrice.Content = Methods.LoadFullPrice(gridPriceListInfo).ToString();
                //choosePriceList.SelectedIndex = -1;
                entryCount.Text = "";
            }            
            else Xceed.Wpf.Toolkit.MessageBox.Show("Требуется выбрать услугу и ввести кол-во");
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (gridPriceListInfo.SelectedIndex != -1) 
            {
                if(uint.TryParse(entrydeleteCount.Text,out uint value)) {
                    int index = gridPriceListInfo.SelectedIndex;
                    PriceList row = (PriceList)gridPriceListInfo.Items[index];
                    if (row.Count > value)
                    {
                        row.Count -= value;
                        gridPriceListInfo.Items[index] = new PriceList { Id = row.Id, Title = row.Title, Price = row.Price, Count = row.Count };
                    }
                    else
                        Xceed.Wpf.Toolkit.MessageBox.Show("Невозможно удалить больше возможного");
                }
                else
                   gridPriceListInfo.Items.RemoveAt(gridPriceListInfo.SelectedIndex);
                checkFullPrice.Content = Methods.LoadFullPrice(gridPriceListInfo).ToString();
                entrydeleteCount.Text = "";
            }
            else Xceed.Wpf.Toolkit.MessageBox.Show("Требуется выбрать строку для удаления");
        }

        private void addRecord_Click(object sender, RoutedEventArgs e)
        {   
            if (chooseDoctor.SelectedIndex != -1 && choosePatient.SelectedIndex != -1 && chooseDate.SelectedDate.HasValue && chooseTime.Value.HasValue && gridPriceListInfo.Items.Count>0)
            {
                if (chooseDate.SelectedDate.Value > DateTime.Now)
                {
                    if (chooseDate.SelectedDate.Value.Month <= DateTime.Now.Month + 1)
                    {
                        query = $"Select Count(Record.Id) From Record Inner Join Doctor on Doctor.Id = Record.IdDoctor" +
                            $" where Time = '{chooseTime.Text}' And Date = '{chooseDate.SelectedDate.Value.ToString("yyyy.MM.dd")}' And Doctor.Id = {chooseDoctor.SelectedValue}";
                        DBManager.ConnectOpen();
                        SqlCommand con = new SqlCommand(query,DBManager.DentistryDBConnetion);
                        int count = int.Parse(con.ExecuteScalar().ToString());
                        DBManager.ConnectClose();
                        if (count == 0)
                        {
                            decimal fullPrice = Methods.LoadFullPrice(gridPriceListInfo);
                            query = $"Insert into Record values({choosePatient.SelectedValue},{chooseDoctor.SelectedValue},'{chooseTime.Value}','{chooseDate.SelectedDate.Value}',{fullPrice})";
                            DBManager.ExecuteQuery(query);
                            for (int i = 0; i < gridPriceListInfo.Items.Count; i++)
                            {
                                PriceList row = (PriceList)gridPriceListInfo.Items[i];
                                query = $"Insert into ListOfServices Values({row.Id},(Select Max(Id) From Record),{row.Count})";
                                DBManager.ExecuteQuery(query);
                            }
                            Hide();
                        }
                        else
                            Xceed.Wpf.Toolkit.MessageBox.Show("У данного врача имеется запись на это время");
                    }
                    else
                        Xceed.Wpf.Toolkit.MessageBox.Show("Можно записаться только на месяц вперёд");
                }
                else Xceed.Wpf.Toolkit.MessageBox.Show("Нельзя записать на предыдущее число");
            }
            else
                Xceed.Wpf.Toolkit.MessageBox.Show("Требуется заполнить все поля");
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            MainWindow mainWindow = this.Owner as MainWindow;
            query = "Select Id as'Id',Title as 'Имя' from PriceList";
            DBManager.LoadDateInComboBox(choosePriceList, query, "Id", "Имя");
            query = "Select id as 'Id', Concat(Trim(Name),' ',Trim(Surname),' ',Trim(Patronymic)) as 'ФИО' from Patient";
            DBManager.LoadDateInComboBox(choosePatient, query, "Id", "ФИО");
            query = "Select id as 'Id',  Concat(Trim(Name),' ',Trim(Surname),' ',Trim(Patronymic)) as 'ФИО' from Doctor";
            DBManager.LoadDateInComboBox(chooseDoctor, query, "Id", "ФИО");
        }
    }
}
