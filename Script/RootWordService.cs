using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Paket;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace Script
{
    public class RootWordService
    {
        static string ConnectionString = string.Empty;
        SqlConnection con;
        SqlCommand cmd = new SqlCommand();

        // For proccess WORD file
        string Prefix_String = string.Empty;
        string InterStart_String = string.Empty;
        string InterEnd_String = string.Empty;
        string Suffix_String = string.Empty;
        string Shape_String = string.Empty;
        string NonShaped_String = string.Empty;

        int rid = 10;

        string RootFolder = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "WordTesting\\OutputFolder\\");

        // Word file paths
        static string w_UnzippedPath = "Questions.docx.unzipped\\";

        static string w_MediaPath = "word\\media\\";
        static string w_DocRelsPath = "word\\_rels\\document.xml.rels";
        static string w_docXmlPath = "word\\document.xml";

        static string UnZipInto = string.Empty;
        static string ImageDB_Path = string.Empty;
        static string MediaFile_Dest = string.Empty;
        static string RelsFile_Dest = string.Empty;
        static string XmlDocFile_Dest = string.Empty;

        static string FINALOUTPUTFOLDER = string.Empty;

        //int sClass = 0;
        //string sSubject = string.Empty;
        string sSheetType = string.Empty;


        List<MyQuestions> All_Question = new List<MyQuestions>();

        //For Db
        string FileName = string.Empty;
        public RootWordService(string connectionString, string RootPath)
        {
            con = new SqlConnection(connectionString);
            cmd.Connection = con;
            con.Open();

            //this.sClass = int.Parse(Grade);
            //this.sSubject = SubjectName;

            RootFolder = RootPath;

            UnZipInto = RootFolder + w_UnzippedPath;
            MediaFile_Dest = UnZipInto + w_MediaPath;
            RelsFile_Dest = UnZipInto + w_DocRelsPath;
            XmlDocFile_Dest = UnZipInto + w_docXmlPath;

            ImageDB_Path = RootFolder + "Images\\";
            FINALOUTPUTFOLDER = RootFolder + "FINAL\\";
            CreateFolders();
        }

        public List<MyQuestions> ReadWordFile(string filename)
        {
            string[] temp = filename.Split(new[] { "\\" }, StringSplitOptions.None);
            FileName = temp[temp.Length - 1];
            ConvertToUnZip(filename);
            List<MyQuestions> MyQuestionList = new List<MyQuestions>();
            bool HasImage = false;
            string newimg = string.Empty;
            string selectedsclass = string.Empty;
            string selectsub = string.Empty;
            string SelectSheettype = string.Empty;
            string pathbox = string.Empty;


            var xpic = "";
            var xr = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filename, true))
            {
                string sclass = string.Empty, section = string.Empty, TestType = string.Empty, subject = string.Empty, Duration = string.Empty, Topic = string.Empty, SubTopic = string.Empty;
                IEnumerable<TableRow> rows;
                var tr = wordDoc.MainDocumentPart.Document.Body.Descendants<Table>().ToList().FirstOrDefault();
                rows = tr.Elements<TableRow>();
                var numberofrows = rows.Count();


                //iterating table rows
                int totalquestion = 0;
                int fCount = 0;
                foreach (TableRow row in rows)
                {

                    string questiondescription = string.Empty;
                    // Find the first cell in the row.
                    TableCell LeftCell = row.Elements<TableCell>().ElementAt(0);
                    // Find the second cell in the row.
                    TableCell RightCell = row.Elements<TableCell>().ElementAt(1);
                    if (LeftCell.InnerText.Trim().ToUpper() == "Q")
                    {
                        questiondescription = RightCell.InnerText.ToString();
                    }
                    if (LeftCell.InnerText.Trim().ToUpper() == "CLASS")
                    {
                        sclass = RightCell.InnerText.ToString();
                    }
                    else if (LeftCell.InnerText.Trim().ToUpper() == "SUBJECT")
                    {
                        subject = RightCell.InnerText.ToString();
                    }
                    else if (LeftCell.InnerText.Trim().ToUpper() == "TOPIC")
                    {
                        Topic = RightCell.InnerText.ToString();
                    }
                    else if (LeftCell.InnerText.Trim().ToUpper() == "SECTION")
                    {
                        section = RightCell.InnerText.ToString();
                    }
                    else if (LeftCell.InnerText.Trim().ToUpper() == "DURATION")
                    {
                        Duration = RightCell.InnerText.ToString();
                    }
                    else if (LeftCell.InnerText.Trim().ToUpper() == "TESTTYPE")
                    {
                        TestType = RightCell.InnerText.ToString();
                        SelectSheettype = TestType;
                    }
                    else if (LeftCell.InnerText.Trim().ToUpper() == "SUBTOPIC")
                    {
                        SubTopic = RightCell.InnerText.ToString();
                    }
                    else
                    {
                        totalquestion = totalquestion + 1;
                        HasImage = false;
                        string Questiontype = string.Empty;
                        int marks = 0;

                        Questiontype = LeftCell.InnerText.Trim().Substring(0, 1).ToString();
                        if (Questiontype.ToUpper() != "Q")
                        {
                            var sst = LeftCell.InnerText.Trim().Substring(1);
                            marks = Convert.ToInt32(LeftCell.InnerText.Trim().Substring(1));
                        }
                        var imagepart1 = from graphicdata in RightCell.Descendants<DocumentFormat.OpenXml.Drawing.GraphicData>() select graphicdata.ToList();
                        if (imagepart1 != null && imagepart1.Count() > 0)
                        {
                            List<ImageTable> ImagesInQuestionItem = new List<ImageTable>();

                            IEnumerable<DocumentFormat.OpenXml.Wordprocessing.Drawing> drawings = RightCell.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().ToList();
                            foreach (DocumentFormat.OpenXml.Wordprocessing.Drawing drawing in drawings)
                            {
                                DocProperties dpr = drawing.Descendants<DocProperties>().FirstOrDefault();
                                if (dpr != null)
                                {
                                    foreach (DocumentFormat.OpenXml.Drawing.Blip b in drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().ToList())
                                    {
                                        fCount++;
                                        newimg = "img" + fCount;
                                        string newimgpath = "img" + fCount + ".jpg";
                                        string opfilename = ImageDB_Path + "{0}";

                                        var outputFilename = string.Format(opfilename, newimgpath);
                                        var imageData = wordDoc.MainDocumentPart.GetPartById(b.Embed);

                                        var data2 = imageData.Uri.ToString();
                                        var imagepath = data2.Replace('/', '\\');
                                        var currentfilename = filename + ".unzipped";
                                        byte[] imagedata = converterDemo(currentfilename + imagepath);
                                        HasImage = true;

                                        ImagesInQuestionItem.Add(new ImageTable
                                        {
                                            Qid = totalquestion,
                                            Imagenumber = newimg,
                                            Imagebyte = imagedata
                                        });
                                    }
                                }
                            }
                            MyQuestionList.Add(new MyQuestions
                            {
                                Qno = totalquestion,
                                Question = Addtablecell(RightCell.InnerXml.ToString()),
                                InnterText = RightCell.InnerText,
                                TextLength = RightCell.InnerText.Length,
                                HasImage = HasImage,
                                Subject = subject,
                                Sclass = sclass,
                                QType = Questiontype,
                                QDesc = questiondescription,
                                Marks = marks,
                                Topic = Topic,
                                SubTopic = SubTopic,
                                SheetType = SelectSheettype,
                                QName = pathbox,
                                ImagesList = ImagesInQuestionItem
                            });
                        }
                        else
                        {
                            MyQuestionList.Add(new MyQuestions
                            {
                                Qno = totalquestion,
                                Question = Addtablecell(RightCell.InnerXml.ToString()),
                                InnterText = RightCell.InnerText,
                                TextLength = RightCell.InnerText.Length,
                                HasImage = HasImage,
                                Subject = subject,
                                Sclass = sclass,
                                QType = Questiontype,
                                QDesc = questiondescription,
                                Marks = marks,
                                Topic = Topic,
                                SubTopic = SubTopic,
                                SheetType = SelectSheettype,
                                QName = pathbox
                            });
                        }
                    }
                }
            }
            Deleteunzippedfolder(filename);
            All_Question.Clear();
            All_Question = MyQuestionList;
            return MyQuestionList;
        }
        public string GenerateWordDocumentFile(List<MyQuestions> MyQuestionList)
        {
            GenerateDirectory(UnZipInto);
            foreach (var image in Directory.EnumerateFiles(ImageDB_Path, "*.png"))
            {
                var imgname = System.IO.Path.GetFileName(image);
                if (imgname.ToString() == "image1" + ".png")
                {
                    string dest = UnZipInto + w_MediaPath + imgname;
                    File.Copy(image, dest);
                }
            }

            Prefix_String = Get_PreFix();
            Suffix_String = Get_Suffix();
            InterStart_String = Get_InterStart();
            InterEnd_String = Get_InterEnd();
            NonShaped_String = Get_NonShape();
            Shape_String = Get_Shaped("single");

            StringBuilder sb = new StringBuilder();
            sb.Append(Prefix_String);
            foreach (var question in MyQuestionList)
            {
                var qtype = question.QType;
                if (question.HasImage)
                {
                    var qimage = question.ImagesList.Where(temp => temp.Qid == question.Qno).ToList();
                    var questionstring = question.Question;
                    sb.Append(InterStart_String);
                    for (int i = 0; i < qimage.Count(); i++)
                    {
                        rid++;
                        string newstring = questionstring.ToString().Replace("pic:cNvPr", "piccNvPr").Replace("a:blip", "ablip").Replace("r:embed", "rembed").Replace("w:tc", "wzzztc");
                        XDocument xdocument = XDocument.Parse(newstring);
                        var testelement = xdocument.Descendants("piccNvPr").Attributes("name").ToList();
                        var element = xdocument.Descendants("piccNvPr").Attributes("name").ToList()[i];
                        element.Value = qimage[i].Imagenumber + ".jpg";
                        var testelement1 = xdocument.Descendants("ablip").Attributes("rembed").ToList();
                        var element1 = xdocument.Descendants("ablip").Attributes("rembed").ToList()[i];
                        element1.Value = "rId" + rid;
                        var doc1 = xdocument.ToString().Replace("piccNvPr", "pic:cNvPr").Replace("ablip", "a:blip").Replace("rembed", "r:embed").Replace("wzzztc", "w:tc");
                        questionstring = string.Empty;
                        questionstring = doc1.ToString();
                        ByteArrayToFile(MediaFile_Dest, qimage[i].Imagenumber, qimage[i].Imagebyte);
                        AddImageResourceIdAndNameInDocXmlRels(RelsFile_Dest, qimage[i].Imagenumber, "rId" + rid);
                    }
                    if (qtype.ToString().ToUpper() == "Q")
                    {
                        sb.Append(NonShaped_String);
                    }
                    else
                    {
                        sb.Append(Shape_String);
                    }
                    sb.Append(questionstring);
                    sb.Append(InterEnd_String);
                }
                else
                {
                    sb.Append(InterStart_String);
                    if (qtype.ToString().ToUpper() == "Q")
                    {
                        sb.Append(NonShaped_String);
                    }
                    else
                    {
                        sb.Append(Shape_String);
                    }
                    sb.Append(question.Question);
                    sb.Append(InterEnd_String);
                }
            }
            sb.Append(Suffix_String);

            string allxmlfile = sb.ToString();
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(allxmlfile.ToString());
            doc.Save(XmlDocFile_Dest);

            string fullPath = ResolveErrorandSaveDocument();
            FindPageType(fullPath);
            return fullPath;
        }
        public FileItem Generate_WordDoc(MyQuestions questionItem)
        {
            GenerateDirectory(UnZipInto);

            Prefix_String = Get_PreFix();
            Suffix_String = Get_Suffix();
            InterStart_String = Get_InterStart();
            InterEnd_String = Get_InterEnd();
            NonShaped_String = Get_NonShape();
            Shape_String = Get_Shaped(questionItem.QType);

            StringBuilder sb = new StringBuilder();
            sb.Append(Prefix_String);
            sb.Append(InterStart_String);

            if (questionItem.QType.ToUpper() == "Q")
                sb.Append(NonShaped_String);
            else
                sb.Append(Shape_String);

            sb.Append(questionItem.Question);
            sb.Append(InterEnd_String);
            sb.Append(Suffix_String);

            TOSAVE(XmlDocFile_Dest, sb.ToString());

            string fullPath = ResolveErrorandSaveDocument(questionItem.Qno.ToString());
            FindPageType(fullPath);
            return new FileItem(questionItem.Qno, fullPath);
        }
        public FileItem Generate_WordDoc_Image(MyQuestions questionItem)
        {
            GenerateDirectory(UnZipInto);

            Prefix_String = Get_PreFix();
            Suffix_String = Get_Suffix();
            InterStart_String = Get_InterStart();
            InterEnd_String = Get_InterEnd();
            NonShaped_String = Get_NonShape();
            Shape_String = Get_Shaped(questionItem.QType);


            StringBuilder sb = new StringBuilder();
            sb.Append(Prefix_String);
            sb.Append(InterStart_String);

            string Question_ItemString = questionItem.Question;
            for (int i = 0; i < questionItem.ImagesList.Count(); i++)
            {
                rid++;
                string New_String_Replaced_Stuff = Question_ItemString.ToString().Replace("pic:cNvPr", "piccNvPr").Replace("a:blip", "ablip").Replace("r:embed", "rembed").Replace("w:tc", "wzzztc");

                XDocument xdocument = XDocument.Parse(New_String_Replaced_Stuff);
                var testelement = xdocument.Descendants("piccNvPr").Attributes("name").ToList();
                var element = xdocument.Descendants("piccNvPr").Attributes("name").ToList()[i];
                element.Value = questionItem.ImagesList[i].Imagenumber + ".jpg";

                var testelement1 = xdocument.Descendants("ablip").Attributes("rembed").ToList();
                var element1 = xdocument.Descendants("ablip").Attributes("rembed").ToList()[i];
                element1.Value = "rId" + rid;

                var doc1 = xdocument.ToString().Replace("piccNvPr", "pic:cNvPr").Replace("ablip", "a:blip").Replace("rembed", "r:embed").Replace("wzzztc", "w:tc");
                Question_ItemString = string.Empty;
                Question_ItemString = doc1.ToString();

                ByteArrayToFile(MediaFile_Dest, questionItem.ImagesList[i].Imagenumber, questionItem.ImagesList[i].Imagebyte);
                AddImageResourceIdAndNameInDocXmlRels(RelsFile_Dest, questionItem.ImagesList[i].Imagenumber, "rId" + rid);
            }

            if (questionItem.QType.ToUpper() == "Q")
                sb.Append(NonShaped_String);
            else
                sb.Append(Shape_String);

            sb.Append(Question_ItemString);
            sb.Append(InterEnd_String);
            sb.Append(Suffix_String);

            TOSAVE(XmlDocFile_Dest, sb.ToString());


            string fullPath = ResolveErrorandSaveDocument(questionItem.Qno.ToString());
            FindPageType(fullPath);
            return new FileItem(questionItem.Qno, fullPath);
        }


        public void InsertIntoDatabase(List<MyQuestions> questionList)
        {
            string TestID = GetTestID();
            foreach (MyQuestions que_item in questionList)
            {
                if (que_item.HasImage)
                {
                    foreach (ImageTable img_item in que_item.ImagesList)
                    {
                        InsertIntoImageTable(img_item);
                    }
                }
                int dbQuestionID = GetQuestionId();
                InsertInto_Question(dbQuestionID, que_item);
                InsertInto_FileUpload(dbQuestionID, que_item.Sclass, que_item.Subject, que_item.SheetType);
                InsertInto_TestInfo(TestID, dbQuestionID);
            }
        }
        public void InsertIntoDatabase()
        {
            string TestID = GetTestID();
            foreach (MyQuestions que_item in All_Question)
            {
                if (que_item.HasImage)
                {
                    foreach (ImageTable img_item in que_item.ImagesList)
                    {
                        InsertIntoImageTable(img_item);
                    }
                }
                int dbQuestionID = GetQuestionId();
                InsertInto_Question(dbQuestionID, que_item);
                InsertInto_FileUpload(dbQuestionID, que_item.Sclass, que_item.Subject, que_item.SheetType);
                InsertInto_TestInfo(TestID, dbQuestionID);
            }
        }

        void InsertInto_Question(int questionid, MyQuestions question)
        {
            cmd.Parameters.Clear();
            cmd.CommandText = "insert into Questions values(@Qid,@Question,@Hasimage,@Subject,@SClass,@QType,@Marks,@QDesc,@Topic,@SubTopic,@SheetType,@QName)";
            cmd.Parameters.AddWithValue("@Qid", questionid);
            //  cmd.Parameters.AddWithValue("@Question", cell.InnerXml);
            cmd.Parameters.AddWithValue("@Question", question.Question);
            cmd.Parameters.AddWithValue("@Hasimage", question.HasImage);
            cmd.Parameters.AddWithValue("@Subject", question.Subject);
            cmd.Parameters.AddWithValue("@SClass", question.Sclass);
            // cmd.Parameters.AddWithValue("@QType", cell1.InnerText.Trim().ToString());
            cmd.Parameters.AddWithValue("@QType", question.QType);
            cmd.Parameters.AddWithValue("@QDesc", question.QDesc);
            cmd.Parameters.AddWithValue("@Marks", question.Marks);
            cmd.Parameters.AddWithValue("@Topic", question.Topic);
            cmd.Parameters.AddWithValue("@SubTopic", question.SubTopic);
            cmd.Parameters.AddWithValue("@SheetType", question.SheetType);
            cmd.Parameters.AddWithValue("@QName", question.QName);
            cmd.ExecuteNonQuery();
        }
        void InsertIntoImageTable(ImageTable img)
        {
            cmd.Parameters.Clear();
            cmd.CommandText = "insert into imagetable values(@Qid,@imagenumber,@ImageByte)";
            cmd.Parameters.AddWithValue("@Qid", img.Qid);
            cmd.Parameters.AddWithValue("@imagenumber", img.Imagenumber);
            cmd.Parameters.AddWithValue("@ImageByte", img.Imagebyte);
            cmd.ExecuteNonQuery();
        }
        void InsertInto_TestInfo(string currentTestId, int Qid)
        {
            cmd.Parameters.Clear();
            cmd.CommandText = "insert into TestInfo values(@TestId,@QID)";
            cmd.Parameters.AddWithValue("@TestId", currentTestId);
            cmd.Parameters.AddWithValue("@QID", Qid);
            cmd.ExecuteNonQuery();
        }
        void InsertInto_FileUpload(int qid, string sClass, string subject, string sheetType)
        {
            cmd.Parameters.Clear();
            cmd.CommandText = "insert into UploadFileName values(@Qid,@FileName,@SheetType,@Sclass,@Subjects)";
            cmd.Parameters.AddWithValue("@Qid", qid);
            cmd.Parameters.AddWithValue("@FileName", FileName);
            cmd.Parameters.AddWithValue("@SheetType", sheetType);
            cmd.Parameters.AddWithValue("@Sclass", sClass);
            cmd.Parameters.AddWithValue("@Subjects", subject);
            cmd.ExecuteNonQuery();
            cmd.Parameters.Clear();
        }


        // Essentials
        int GetQuestionId()
        {
            cmd.Parameters.Clear();
            cmd.CommandText = "select count(*) from questions";
            int questionid = Convert.ToInt32(cmd.ExecuteScalar()) + 1;
            return questionid;
        }
        string GetTestID()
        {
            cmd.Parameters.Clear();
            cmd.CommandText = "select Count(distinct(Testid)) from TestInfo ";
            int nooftestid = Convert.ToInt32(cmd.ExecuteScalar()) + 1;
            string currentTestid = GetCurrentTestId(nooftestid);
            return currentTestid;
        }
        public string GetCurrentTestId(int existcount)
        {
            int existtestcountlength = Convert.ToInt32(Math.Floor(Math.Log10(existcount) + 1));
            string testid = "00000";
            if (existcount == 0)
            {
                testid = "00001";
            }
            else if (existtestcountlength == 1)
            {
                testid = "0000" + existcount.ToString();
            }
            else if (existtestcountlength == 2)
            {
                testid = "000" + existcount.ToString();
            }
            else if (existtestcountlength == 3)
            {
                testid = "00" + existcount.ToString();
            }
            else if (existtestcountlength == 4)
            {
                testid = "0" + existcount.ToString();
            }
            else if (existtestcountlength == 5)
            {
                testid = existcount.ToString();
            }
            else
            {
                return testid;
            }
            return testid;
        }
        void GenerateDirectory(string destinationDirectorypath)
        {
            if (Directory.Exists(destinationDirectorypath))
            {
                Directory.Delete(destinationDirectorypath, true);
                Directory.CreateDirectory(destinationDirectorypath);
                CreateDefaultFoldersAndFileswithHeader(destinationDirectorypath);
            }
            else
            {
                Directory.CreateDirectory(destinationDirectorypath);
                CreateDefaultFoldersAndFileswithHeader(destinationDirectorypath);
            }
        }
        void CreateDefaultFoldersAndFileswithHeader(string Folderpath)
        {
            string relsFolder = Folderpath + "\\_rels";
            string docProsFolder = Folderpath + "\\docProps";
            string wordProsFolder = Folderpath + "\\word";
            string customxmlFolder = Folderpath + "\\customXml";
            cmd.CommandText = "select ContentTypeFile from DefaultFilesTable1";
            var contenttypefile = cmd.ExecuteScalar().ToString();
            XmlDocument docctype = new XmlDocument();
            docctype.LoadXml(contenttypefile);
            docctype.Save(Folderpath + "\\[Content_Types].xml");
            //create _rels folder
            if (!Directory.Exists(relsFolder))
            {
                Directory.CreateDirectory(relsFolder);
            }
            else
            {
                Directory.Delete(relsFolder, true);
                Directory.CreateDirectory(relsFolder);
            }
            //create docProsFolder
            if (!Directory.Exists(docProsFolder))
            {
                Directory.CreateDirectory(docProsFolder);
            }
            else
            {
                Directory.Delete(docProsFolder, true);
                Directory.CreateDirectory(docProsFolder);
            }
            //create wordProsFolder
            if (!Directory.Exists(wordProsFolder))
            {
                Directory.CreateDirectory(wordProsFolder);
            }
            else
            {
                Directory.Delete(wordProsFolder, true);
                Directory.CreateDirectory(wordProsFolder);
            }
            //create CustomxmlFolder
            if (!Directory.Exists(customxmlFolder))
            {
                Directory.CreateDirectory(customxmlFolder);
            }
            else
            {
                Directory.Delete(customxmlFolder, true);
                Directory.CreateDirectory(customxmlFolder);
            }
            // create _rels file
            cmd.CommandText = "select RelsFile from DefaultFilesTable1";
            var RelsFile = cmd.ExecuteScalar().ToString();
            XmlDocument docrelsfile = new XmlDocument();
            docrelsfile.LoadXml(RelsFile);
            docrelsfile.Save(relsFolder + "\\.rels");
            //create rels folder in customxml folder
            string custom_rels_folder = customxmlFolder + "\\" + "_rels";
            if (!Directory.Exists(custom_rels_folder))
            {
                Directory.CreateDirectory(custom_rels_folder);
            }
            else
            {
                Directory.Delete(custom_rels_folder, true);
                Directory.CreateDirectory(custom_rels_folder);
            }

            //create item1xms.rels file in custom xml rels folder
            cmd.CommandText = "select CustomRelsItem from DefaultFilesTable1";
            var customrelsitem = cmd.ExecuteScalar().ToString();
            XmlDocument customitem = new XmlDocument();
            customitem.LoadXml(customrelsitem);
            customitem.Save(custom_rels_folder + "\\item1.xml.rels");
            //create item1.xml file in customxmlfolder
            cmd.CommandText = "select ItemFile from DefaultFilesTable1";
            var ItemFile = cmd.ExecuteScalar().ToString();
            XmlDocument itemfiledoc = new XmlDocument();
            itemfiledoc.LoadXml(ItemFile);
            itemfiledoc.Save(customxmlFolder + "\\item1.xml");
            //create itemProps1.xml file in customxmlfolder
            cmd.CommandText = "select ItemProsFile from DefaultFilesTable1";
            var ItemProsFile = cmd.ExecuteScalar().ToString();
            XmlDocument ItemProsFiledoc = new XmlDocument();
            ItemProsFiledoc.LoadXml(ItemProsFile);
            ItemProsFiledoc.Save(customxmlFolder + "\\itemProps1.xml");
            //create files in docpros folder
            cmd.CommandText = "select AppFile from DefaultFilesTable1";
            var appFile = cmd.ExecuteScalar().ToString();
            XmlDocument docappfile = new XmlDocument();
            docappfile.LoadXml(appFile);
            docappfile.Save(docProsFolder + "\\app.xml");
            //create second file in docpros folder
            cmd.CommandText = "select CoreFile from DefaultFilesTable1";
            var coreFile = cmd.ExecuteScalar().ToString();
            XmlDocument doccorefile = new XmlDocument();
            docappfile.LoadXml(coreFile);
            docappfile.Save(docProsFolder + "\\core.xml");
            //create _rels folder in Word folder
            string word_relsFolder = wordProsFolder + "\\" + "_rels";
            if (!Directory.Exists(word_relsFolder))
            {
                Directory.CreateDirectory(word_relsFolder);
            }
            else
            {
                Directory.Delete(word_relsFolder, true);
                Directory.CreateDirectory(word_relsFolder);
            }
            //create theme folder in word folder
            string word_themeFolder = wordProsFolder + "\\" + "theme";
            if (!Directory.Exists(word_themeFolder))
            {
                Directory.CreateDirectory(word_themeFolder);
            }
            else
            {
                Directory.Delete(word_themeFolder, true);
                Directory.CreateDirectory(word_themeFolder);
            }
            //create media folder in word folder
            string word_mediaFolder = wordProsFolder + "\\" + "media";
            if (!Directory.Exists(word_mediaFolder))
            {
                Directory.CreateDirectory(word_mediaFolder);
            }
            else
            {
                Directory.Delete(word_mediaFolder, true);
                Directory.CreateDirectory(word_mediaFolder);
            }
            //create fonttablefile in word folder
            cmd.CommandText = "select FontTableFile from DefaultFilesTable1";
            var fonttableFile = cmd.ExecuteScalar().ToString();
            XmlDocument docfonttablefile = new XmlDocument();
            docfonttablefile.LoadXml(fonttableFile);
            docfonttablefile.Save(wordProsFolder + "\\fontTable.xml");
            //create settingsfile in word folder
            cmd.CommandText = "select settingsFile from DefaultFilesTable1";
            var settingsFile = cmd.ExecuteScalar().ToString();
            XmlDocument docsettings = new XmlDocument();
            docsettings.LoadXml(settingsFile);
            docsettings.Save(wordProsFolder + "\\settings.xml");
            //create stylefile in word folder
            cmd.CommandText = "select stylesfile from DefaultFilesTable1";
            var styleFile = cmd.ExecuteScalar().ToString();
            XmlDocument docstyle = new XmlDocument();
            docstyle.LoadXml(styleFile);
            docstyle.Save(wordProsFolder + "\\styles.xml");
            //create websettings in word folder
            cmd.CommandText = "select WebsettingsFile from DefaultFilesTable1";
            var websettingFile = cmd.ExecuteScalar().ToString();
            XmlDocument websettingstyle = new XmlDocument();
            websettingstyle.LoadXml(websettingFile);
            websettingstyle.Save(wordProsFolder + "\\webSettings.xml");
            //create endnotes.xml in word folder
            cmd.CommandText = "select Endnotes from DefaultFilesTable1";
            var Endnotes = cmd.ExecuteScalar().ToString();
            XmlDocument Endnotesdoc = new XmlDocument();
            Endnotesdoc.LoadXml(Endnotes);
            Endnotesdoc.Save(wordProsFolder + "\\endnotes.xml");
            //create Footnotes.xml in word folder
            cmd.CommandText = "select Footnotes from DefaultFilesTable1";
            var Footnotes = cmd.ExecuteScalar().ToString();
            XmlDocument Footnotesdoc = new XmlDocument();
            Footnotesdoc.LoadXml(Footnotes);
            Footnotesdoc.Save(wordProsFolder + "\\footnotes.xml");
            //create header1.xml in word folder

            //cmd.CommandText = "select Header1File from DefaultFilesTable1";
            //var HeaderFile = cmd.ExecuteScalar().ToString();            
            //var Header1File = ChangeClassSecionDetails(HeaderFile);\
            //XmlDocument Header1Filedoc = new XmlDocument();
            //Header1Filedoc.LoadXml(Header1File);
            //Header1Filedoc.Save(wordProsFolder + "\\header1.xml");



            //create numbering.xml in word folder
            cmd.CommandText = "select Header1File from DefaultFilesTable1";
            var Numbering1File = cmd.ExecuteScalar().ToString();
            XmlDocument Numbering1Filedoc = new XmlDocument();
            Numbering1Filedoc.LoadXml(Numbering1File);
            Numbering1Filedoc.Save(wordProsFolder + "\\numbering.xml");
            //create file in word _rels folder
            cmd.CommandText = "select DocRelsFile from DefaultFilesTable1";
            var docrelsFile = cmd.ExecuteScalar().ToString();
            XmlDocument xdocrelsfile = new XmlDocument();
            xdocrelsfile.LoadXml(docrelsFile);
            xdocrelsfile.Save(word_relsFolder + "\\document.xml.rels");
            //create header1.xml.rels file in _rels Folder
            cmd.CommandText = "select HeaderRels from DefaultFilesTable1";
            var HeaderRels = cmd.ExecuteScalar().ToString();
            XmlDocument HeaderRelsdoc = new XmlDocument();
            HeaderRelsdoc.LoadXml(HeaderRels);
            HeaderRelsdoc.Save(word_relsFolder + "\\header1.xml.rels");


            //create file in word theme folder
            cmd.CommandText = "select ThemeFile from DefaultFilesTable1";
            var themeFile = cmd.ExecuteScalar().ToString();
            XmlDocument xdocthemefile = new XmlDocument();
            xdocthemefile.LoadXml(themeFile);
            xdocthemefile.Save(word_themeFolder + "\\theme1.xml");
        }
        //string ChangeClassSecionDetails(string headerxmlfile)
        //{
        //    string sclass = "";
        //    string subject = "";
        //    string activity = "";
        //    if (sClass != 0)
        //    {
        //        sclass = sClass.ToString();
        //    }
        //    if (sSubject != "")
        //    {
        //        subject = sSubject;
        //    }
        //    if (sSheetType != "")
        //    {
        //        activity = sSheetType;
        //    }
        //    //return headerxmlfile.Replace("Class:", "Class:   " + sclass + "").Replace("Section:", " ").Replace("Subject:", "Subject: " + subject + "").Replace("Activity:", "" + activity + "   ").ToString();
        //    return headerxmlfile.Replace("Class:", string.Empty).Replace("Section:", string.Empty).Replace("Subject:", string.Empty).Replace("Activity:", string.Empty).ToString();
        //}
        void AddImageResourceIdAndNameInDocXmlRels(string documentxmlrels, string imagename, string rid)
        {
            try
            {
                string newstring = documentxmlrels.ToString();
                XDocument xdocument = XDocument.Load(newstring);
                XElement newTag1 = new XElement("Relationship",
                new XAttribute("Id", rid),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"),
                new XAttribute("Target", "media/" + imagename + ".jpg"));
                xdocument.Root.LastNode.AddAfterSelf(newTag1);
                xdocument.Save(documentxmlrels);
            }
            catch (Exception ex)
            {
            }
        }
        void TOSAVE(string docFilePath, string xmlData)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlData.ToString());
            doc.Save(docFilePath);
        }
        static string ResolveErrorandSaveDocument(string filename = "Question")
        {
            //string outputfilepath = filename + ques_id + " .docx";
            if (!Directory.Exists(FINALOUTPUTFOLDER))
            {
                Directory.CreateDirectory(FINALOUTPUTFOLDER);
            }
            if (!filename.Contains(".docx")) filename += ".docx";
            if (!filename.Contains("Question")) filename = "Question" + filename;

            string outputfilepath = FINALOUTPUTFOLDER + filename;
            ConvertToZipFile(UnZipInto, outputfilepath);

            object oMissing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = app.Documents.OpenNoRepairDialog(outputfilepath, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, true, ref oMissing, ref oMissing, ref oMissing);
            doc.SaveAs2(outputfilepath, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            doc.Close();
            app.Quit();

            if (Directory.Exists(UnZipInto))
                Directory.Delete(UnZipInto, true);
            return outputfilepath;
        }
        static void FindPageType(string filename)
        {
            //string filename = @"D:\WordTesting\OutputFolder\Question13.docx";
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filename, true))
            {
                string docText = null;
                int pageCount = Convert.ToInt32(wordDoc.ExtendedFilePropertiesPart.Properties.Pages.Text);
                if (pageCount == 1)
                {
                    using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                    {
                        docText = sr.ReadToEnd().Replace("#S~M~C", "");
                    }
                    using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                    {
                        sw.Write(docText);
                    }
                }
                else
                {
                    using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                    {
                        docText = sr.ReadToEnd().Replace("#S~M~C", "##");
                    }
                    using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                    {
                        sw.Write(docText);
                    }
                }
            }
        }
        void CreateFolders()
        {
            if (!Directory.Exists(RootFolder))
            {
                Directory.CreateDirectory(RootFolder);
            }
            if (!Directory.Exists(ImageDB_Path))
            {
                try
                {
                    var currentfile = AppDomain.CurrentDomain.BaseDirectory + "\\Img\\image1.png";
                    var desfile = ImageDB_Path + "\\image1.png";
                    Directory.CreateDirectory(ImageDB_Path);
                    File.Copy(currentfile, desfile);
                }
                catch (Exception ex) { }
            }
        }

        // Conversions
        void ConvertToUnZip(string filename)
        {
            using (ZipArchive archive = ZipFile.OpenRead(filename))
            {
                archive.ExtractToDirectory(filename + ".unzipped");
            }
        }
        void Deleteunzippedfolder(string filename)
        {
            string filefullpath = filename + ".unzipped";
            if (Directory.Exists(filefullpath))
            {
                Directory.Delete(filefullpath, true);
            }
        }
        public static byte[] converterDemo(string x)
        {
            FileStream fs = new FileStream(x, FileMode.Open, FileAccess.Read);
            byte[] imgByteArr = new byte[fs.Length];
            fs.Read(imgByteArr, 0, Convert.ToInt32(fs.Length));
            fs.Close();
            return imgByteArr;
        }
        private void ByteArrayToFile(string bPath, string fName, byte[] content)
        {
            //Save the Byte Array as File.
            string filePath = bPath + fName + ".jpg";
            File.WriteAllBytes(filePath, content);
        }
        static void ConvertToZipFile(string filename, string destinationFileName)
        {
            if (File.Exists(destinationFileName))
            {
                File.Delete(destinationFileName);
            }
            ZipFile.CreateFromDirectory(filename, destinationFileName);
        }
        static bool Kill_WordFileOfFullQuestion()
        {
            bool toReturn = false;
            var MyProcess = Process.GetProcesses();
            foreach (Process p in MyProcess)
            {
                if ("Question.docx - Microsoft Word" == p.MainWindowTitle)
                {
                    p.Kill();
                    toReturn = true;
                    break;
                }
            }
            return toReturn;
        }
        string Addtablecell(string inputstring)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<w:tc>");
            sb.Append(inputstring);
            sb.Append("</w:tc>");
            return sb.ToString();
        }

        // Needs
        static string Get_PreFix()
        {
            //using (SqlConnection con = new SqlConnection(Properties.Settings.Default.Database))
            //{
            //    SqlCommand cmd = new SqlCommand();
            //    con.Open();
            //    cmd.Connection = con;
            //    cmd.CommandText = "select max(testid) from TestInfo";
            //    string testname = cmd.ExecuteScalar().ToString();

            //    var prefixtext = File.ReadAllText(MyDirectory() + @"\prefix12NewFormat.txt");
            //    if (testname != null && testname != "")
            //    {
            //        return prefixtext = prefixtext.Replace("#S~#R~#N~#1", testname[4].ToString()).Replace("#S~#R~#N~#2", testname[1].ToString()).Replace("#S~#R~#N~#3", testname[0].ToString()).Replace("#S~#R~#N~#4", testname[3].ToString()).Replace("#S~#R~#N~#5", testname[2].ToString());
            //    }
            //    else
            //    {
            //        return prefixtext.ToString();
            //    }
            //}
            var prefixtext = File.ReadAllText(MyDirectory() + @"\prefix.txt");
            return prefixtext.ToString();
        }
        static string Get_Suffix()
        {
            var suf = File.ReadAllLines(MyDirectory() + @"\suffix1.txt").ToList();
            StringBuilder ssb = new StringBuilder();
            foreach (var n in suf)
            {
                ssb.Append(n.ToString());
            }
            return ssb.ToString();
        }
        static string Get_InterStart()
        {
            string interstart = string.Empty;
            var s1 = File.ReadAllLines(MyDirectory() + @"\inter1.txt").ToList();
            foreach (var n in s1)
            {
                interstart = n.ToString();
            }
            return interstart;
        }
        static string Get_InterEnd()
        {
            string interend = string.Empty;
            var s2 = File.ReadAllLines(MyDirectory() + @"\inter2.txt").ToList();
            foreach (var n in s2)
            {
                interend = n.ToString();
            }
            return interend;
        }
        static string Get_Shaped(string Qtype)
        {
            List<string> s3 = new List<string>();
            if (Qtype.ToUpper() == "single".ToUpper())
            {
                s3 = File.ReadAllLines(MyDirectory() + @"\shape2Rectangle.txt").ToList();
            }
            else if (Qtype.ToUpper() == "multiple".ToUpper())
            {
                s3 = File.ReadAllLines(MyDirectory() + @"\shape6Rectangle.txt").ToList();
            }
            else 
            {
                s3 = File.ReadAllLines(MyDirectory() + @"\shape2Rectangle.txt").ToList();
            }

            //Set Shaped file
            string shapefile = string.Empty;
            foreach (var n in s3)
            {
                shapefile = n.ToString();
            }
            return shapefile;
        }
        static string Get_NonShape()
        {
            var s4 = File.ReadAllLines(MyDirectory() + @"\nonshape.txt").ToList();
            string nonshape = string.Empty;
            foreach (var n in s4)
            {
                nonshape = n.ToString();
            }
            return nonshape;
        }
        static string MyDirectory()
        {
            return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        }
    }
}
