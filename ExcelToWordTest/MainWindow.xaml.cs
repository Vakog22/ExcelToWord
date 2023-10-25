using ExcelToWordTest.Class;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExcelToWordTest
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btn_Go_Click(object sender, RoutedEventArgs e)
        {
            TestData testData = new TestData();

            var helper = new WordInserter("C:\\Users\\Alexander\\source\\repos\\ExcelToWordTest\\ExcelToWordTest\\Docs\\ZVedomost.doc");

            var items = new Dictionary<string, string>()
            {
                {"<Discipline_name>",  testData.Discipline},
                {"<Group_name>", testData.Group },
                {"<Course_name>", testData.Course },
                {"<Date>", DateTime.Now.ToString("yyyy.MM.dd") },
                {"<Speciality_name>", testData.Specialtity },
                {"<Teacher_name>", testData.Muchitel },
                {"<Signer>", testData.Signer }
            };
            

            helper.Process(items,testData.Pidorasi);
        }
    }
}
