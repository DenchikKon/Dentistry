using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Windows.Controls;
using System.Windows;

namespace WPFPractika
{    
    internal class DBManager
    {
        public static SqlConnection DentistryDBConnetion = new SqlConnection(ConfigurationManager.ConnectionStrings["Dentistry"].ToString());

        public static string mainRecordQuery = @"Select Record.Id,Concat(Trim(Patient.Name),' ',Trim(Patient.Surname),' ',Trim(Patient.Patronymic)) as 'Пациент',
        Concat(Trim(Doctor.Name),' ',Trim(Doctor.Surname),' ',Trim(Doctor.Patronymic)) as 'Доктор', Format(Record.Date,'dd.MM.yyyy') as 'Дата',
        Format(Time,'hh\:mm') as 'Время', Record.FullPrice as 'Стоимость' 
        from Record inner join Patient on Record.IdPatient = Patient.Id
        inner join Doctor on Doctor.Id = Record.IdDoctor";
        public static string mainPatientQuery = "Select Id, CONCAT(Trim(Name),' ',Trim(Surname),' ',Trim(Patronymic)) as 'ФИО', Address as 'Адрес', Phone as 'Телефон', Format(DateOfBirthDay,'dd.MM.yyyy') as 'Дата рождения' From Patient";
        public static string mainPriceListQuery = "Select Id, Title as 'Название', Price as 'Цена' From PriceList";
        public static string mainDoctorQuery = "Select Id, CONCAT(Trim(Name),' ',Trim(Surname),' ',Trim(Patronymic)) as 'ФИО', Spetialization as 'Специализация'," +
            "Phone as 'Телефон', Format(DateOfBirthDay,'dd.MM.yyyy') as 'Дата рождения' From Doctor";
        public static void ConnectOpen()
        {
            DentistryDBConnetion.Open();
        }
        public static void ConnectClose()
        {
            DentistryDBConnetion.Close();
        }
        public static void LoadData(DataGrid data, string query)
        {
            DentistryDBConnetion.Open();
            SqlDataAdapter adapter = new SqlDataAdapter(query, DentistryDBConnetion);
            DentistryDBConnetion.Close();
            System.Data.DataTable table = new System.Data.DataTable();
            adapter.Fill(table);
            data.ItemsSource = table.DefaultView;
        }

        public static void ExecuteQuery(string query)
        {
            DentistryDBConnetion.Open();
            SqlCommand command = new SqlCommand(query,DentistryDBConnetion);
            command.ExecuteNonQuery();
            DentistryDBConnetion.Close();
        }
        public static void LoadDateInComboBox(ComboBox comboBox,string query, string valueMember,string displayMember)
        {
            DentistryDBConnetion.Open();
            SqlDataAdapter adapter = new SqlDataAdapter(query, DentistryDBConnetion);
            DataTable table= new DataTable();
            adapter.Fill(table);
            DentistryDBConnetion.Close();
            comboBox.ItemsSource= table.DefaultView;
            comboBox.DisplayMemberPath= $"{displayMember}";
            comboBox.SelectedValuePath= $"{valueMember}";
        }

    }
}
