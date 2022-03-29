using System;
using System.Collections.Generic;
using System.Data.SqlClient;
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

namespace ChangeQHeading
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string connectionstring = "Data Source=.;Initial Catalog=WordProcess;Integrated Security=True";
        List<type> types = new List<type>();
        public MainWindow()
        {
            InitializeComponent();
        }

        public void insert(List<type> ls)
        {
            using (SqlConnection con = new SqlConnection(connectionstring))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand(con.ConnectionString);
                foreach(type item in ls)
                {
                    cmd.CommandText = "Insert into Questions(Qheading) values('" + item.heading + "') where Qid = '" + item.id + "'";
                    cmd.ExecuteNonQuery();
                }
                con.Close();
            }
        }

        public List<type> getdate()
        {
            List<type> ls = new List<type>();
            using (SqlConnection con = new SqlConnection(connectionstring))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = "select QDesc,Qid from Questions where QDesc like '%Choose the correct answer%'";
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    ls.Add(new type { id = reader.GetInt32(1), heading = reader.GetString(0) });
                }
                con.Close();
            }

            return ls;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            types = getdate();
               

        }
    }

    public class type
    {

        public int id { get; set; }
        public string heading { get; set; }


    }
}
