using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Media;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Xps.Packaging;
using MessageBox = System.Windows.MessageBox;

namespace Script
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        WordProcessHelper word = new WordProcessHelper();
        string connectionstring = "Data Source=.;Initial Catalog=WordProcess;Integrated Security=True";
        List<FileItem> WordFilePathList = new List<FileItem>();

        
        List<MyQuestions> repos = new List<MyQuestions>();
        List<FileItem> XPSfilePathList = new List<FileItem>();

        MyQuestions questions = new MyQuestions();
        string RootFolder = string.Empty;


        string Grade = string.Empty;

        int Start = 0;
        int End = 0;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

            int count = 0;
            //Generate word files for all quetsion item in Database
            RootWordService service = new RootWordService(connectionstring,RootFolder);
            if (repos.Count() > 0)
            {
                #region Old
                //foreach (MyQuestions item in repos)
                //{
                //    FileItem obj;
                //    count++;
                //    if (item.HasImage)
                //    {
                //        obj = service.Generate_WordDoc_Image(item);
                //        GenerateXpsFile(obj);
                //    }
                //    else
                //    {
                //        obj = service.Generate_WordDoc(item);
                //        GenerateXpsFile(obj);
                //    }
                //    if (count == CountOfQue)
                //    {
                //        break;
                //    }   
                //}
                #endregion

                for (int i = Start; i <= End; i++)
                {
                    FileItem obj;
                    count++;
                    if (!repos[i].HasImage)
                    {
                        obj = service.Generate_WordDoc(repos[i]);
                        GenerateXpsFile(obj);
                    }
                    //if (repos[i].HasImage)
                    //{
                    //    obj = service.Generate_WordDoc_Image(repos[i]);
                    //    GenerateXpsFile(obj);
                    //}
                    //else
                    //{
                    //    obj = service.Generate_WordDoc(repos[i]);
                    //    GenerateXpsFile(obj);
                    //}
                }

                xMainPanel.Background = new SolidColorBrush(Colors.LightGreen);
                StartBeep();
                xInsertBtn.Visibility = Visibility.Visible;
            }
            else
            {
                string msg = "No Question for " + Grade + " grade";
                MessageBox.Show(msg);
            }
        }

       

        public async void GenerateXpsFile(FileItem item)
        {
            await Task.Run(() =>
            {
                XPSfilePathList.Add(new FileItem(item.Qno, word.Convert_WordToXPS(item.FilePath)));
            });
        }


        public void InsertIntoXPS(List<FileItem> itemsList)
        {
            using (SqlConnection con = new SqlConnection(connectionstring))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                foreach(FileItem item in itemsList)
                {
                    cmd.Parameters.Clear();
                    cmd.CommandText = "Insert into Xpstable(Qid,XpsFile) values(" + item.Qno + ",@XpsFile)";
                    cmd.Parameters.AddWithValue("@XpsFile",(byte[])item.XpsByteData);
                    cmd.ExecuteNonQuery();
                }
                con.Close();
            }
        }


        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            List<FileItem> FINALtoINSERT = new List<FileItem>();
            foreach (FileItem item in XPSfilePathList)
            {
                FileItem toIn = item;
                toIn.XpsByteData = word.Convert_XpsTOByteArray(new XpsDocument(item.FilePath, System.IO.FileAccess.Read));
                FINALtoINSERT.Add(toIn);
            }
            xMainPanel.Background = new SolidColorBrush(Colors.White);
            InsertIntoXPS(FINALtoINSERT);
            repos = new List<MyQuestions>();
            XPSfilePathList = new List<FileItem>();
            MessageBox.Show("Inserted sucessfully...!");
            xInsertBtn.Visibility = Visibility.Collapsed;
        }

      

        void StartBeep()
        {
            Console.Beep(3000, 4000);
        }

        private void xChoosePath_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog openDialog = new FolderBrowserDialog();
            var result = openDialog.ShowDialog();
            if (result.ToString() == "OK")
            {
                RootFolder = openDialog.SelectedPath;
                RootFolder += "\\";
                xPathNameTB.Text = RootFolder;
            }
        }

        

        private void xGetQuestion_Click(object sender, RoutedEventArgs e)
        {
            xGetQuestion.Visibility = Visibility.Collapsed;
            Grade = xGradeTB.Text;
            repos = MyQuestions.Get_QuestionItemNotInXpsTable(Grade);
            xCountTotal.Text = repos.Count.ToString();
        }

        private void xStartIndex_TextChanged(object sender, TextChangedEventArgs e)
        {
            this.Start = Convert.ToInt32(xStartIndex.Text);
            this.Start--;
        }

        private void xEndIndex_TextChanged(object sender, TextChangedEventArgs e)
        {
            this.End = Convert.ToInt32(xEndIndex.Text);
            this.End--;
        }
    }
}
