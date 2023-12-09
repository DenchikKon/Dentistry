using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace WPFPractika
{
    /// <summary>
    /// Interaction logic for RecordInfo.xaml
    /// </summary>
    public partial class RecordInfo : Window
    {
        public RecordInfo()
        {
            InitializeComponent();
        }

        private void Window_Activated(object sender, EventArgs e)
        {
            MainWindow main = (MainWindow)this.Owner;
            
        }
    }
}
