using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Script
{
    public class MyQuestions
    {
        // Methods

        // First gather Question ID for class, subject and sheet type.
        // 1. List<int> Get_QuestionID_List(string sClass, string subject, string sheetType)

        // Get Single Question item from Database as MyQuestion type for giver question ID.
        // 2. MyQuestions Get_QuestionItem(string questionID)

        // Load all the question item for given question id List.
        // 3. List<MyQuestions> Get_QuestionList(IEnumerable<string> QuestionID_List)
        static string connectionstring = "Data Source=.;Initial Catalog=WordProcess;Integrated Security=True";
        public int Qno { get; set; }
        public string Question { get; set; }
        public string InnterText { get; set; }
        public int TextLength { get; set; }
        public bool HasImage { get; set; }
        public string Subject { get; set; }
        public string Sclass { get; set; }
        public string QType { get; set; }
        public int Marks { get; set; }
        public string QDesc { get; set; }
        public string Topic { get; set; }
        public string SubTopic { get; set; }
        public string SheetType { get; set; }
        public string QName { get; set; }
        public List<ImageTable> ImagesList { get; set; }

        public MyQuestions()
        {
            this.QType = "";
        }

        public List<string> Qids = new List<string>();

        public List<string> Get_QuestionID_List(string sClass, string subject, string topic)
        {
            List<string> toReturn = new List<string>();
            using (SqlConnection con = new SqlConnection(connectionstring))
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                con.Open();
                cmd.CommandText = "select Qid from Questions where QType = 'R' and SClass = '" + sClass + "' and Subject = '" + subject + "' and Topic = '" + topic + "'";
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    toReturn.Add(dr.GetInt32(0).ToString());
                }
                dr.Close();
                con.Close();
            }
            return toReturn;
        }
        public static MyQuestions Get_QuestionItem(string questionID)
        {
            MyQuestions Que_Object = new MyQuestions();
            using (SqlConnection con = new SqlConnection(connectionstring))
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                con.Open();
                cmd.CommandText = "select Qid, Question, Hasimage, Subject, SClass, QType, Marks, QDesc,Topic, SubTopic, SheetType, QName from Questions where Qid = '" + questionID + "'";
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    Que_Object.Qno = dr.GetInt32(0);
                    Que_Object.Question = dr.GetString(1);
                    Que_Object.HasImage = dr.GetBoolean(2);
                    Que_Object.Subject = dr.GetString(3);
                    Que_Object.Sclass = dr.GetString(4);
                    Que_Object.QType = dr.GetString(5);
                    Que_Object.Marks = dr.GetInt32(6);
                    Que_Object.QDesc = dr.GetString(7);
                    Que_Object.Topic = dr.GetString(8);
                    Que_Object.SubTopic = dr.GetString(9);
                    Que_Object.SheetType = dr.GetString(10);
                    Que_Object.QName = dr.GetString(11);
                }

                if (Que_Object.HasImage)
                    Que_Object.ImagesList = GetImages(Que_Object.Qno.ToString());

                dr.Close();
                con.Close();

            }
            return Que_Object;
        }

        public static List<MyQuestions> Get_QuestionItemNotInXpsTable(string grade)
        {
            List<MyQuestions> toReturn = new List<MyQuestions>();
            using (SqlConnection con = new SqlConnection(connectionstring))
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                con.Open();
                cmd.CommandText = "select (Questions.Qid), Question, Hasimage, Subject, SClass, QType, Marks, QDesc,Topic, SubTopic, SheetType, QName from Questions where Qid not in (select qid from Xpstable) and SClass = "+ grade + " and Hasimage = 0";
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    MyQuestions Que_Object = new MyQuestions();
                    Que_Object.Qno = dr.GetInt32(0);
                    Que_Object.Question = dr.GetString(1);
                    Que_Object.HasImage = dr.GetBoolean(2);
                    Que_Object.Subject = dr.GetString(3);
                    Que_Object.Sclass = dr.GetString(4);
                    Que_Object.QType = dr.GetString(5);
                    Que_Object.Marks = dr.GetInt32(6);
                    Que_Object.QDesc = dr.GetString(7);
                    Que_Object.Topic = dr.GetString(8);
                    Que_Object.SubTopic = dr.GetString(9);
                    Que_Object.SheetType = dr.GetString(10);
                    Que_Object.QName = dr.GetString(11);

                    if (Que_Object.HasImage)
                        Que_Object.ImagesList = GetImages(Que_Object.Qno.ToString());
                    toReturn.Add(Que_Object);
                }

                dr.Close();
                con.Close();

            }
            return toReturn;
        }

        static List<ImageTable> GetImages(string Qid)
        {
            List<ImageTable> imgList = new List<ImageTable>();
            using (SqlConnection con = new SqlConnection(connectionstring))
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                con.Open();
                cmd.CommandText = "select Qid, imagenumber, ImageByte from imagetable where Qid = '" + Qid + "'";
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    ImageTable obj = new ImageTable();
                    obj.Qid = int.Parse(Qid);
                    obj.Imagenumber = dr.GetString(1);
                    obj.Imagebyte = (byte[])dr.GetValue(2);
                    imgList.Add(obj);
                }
            }
            return imgList;
        }

        public List<MyQuestions> Get_QuestionList(IEnumerable<string> QuestionID_List)
        {
            List<MyQuestions> toReturn = new List<MyQuestions>();
            foreach (string id in QuestionID_List)
            {
                toReturn.Add(Get_QuestionItem(id));
            }
            return toReturn;
        }

        public static List<string> GetTopics(string sClass, string subject)
        {
            List<string> toReturn = new List<string>();
            using (SqlConnection con = new SqlConnection(connectionstring))
            {
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                con.Open();
                cmd.CommandText = "select distinct(Topic) from Questions where SClass = '" + sClass + "' and Subject = '" + subject + "'";
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                    toReturn.Add(dr.GetString(0));
            }
            return toReturn;
        }
    }
}
