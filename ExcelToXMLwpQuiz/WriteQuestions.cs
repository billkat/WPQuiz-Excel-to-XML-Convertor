using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Net;
using System.Windows.Documents;
using System.Xml;

namespace ExcelToXMLwpQuiz
{
   public class WriteQuestions
    {
        protected static Dictionary<string, int> answerDictionary;
        protected static Dictionary<Questions, List<Answers>> questionDictionary;

      public WriteQuestions()
       {
           answerDictionary = new Dictionary<string, int>();
           answerDictionary.Add("A", 0);
           answerDictionary.Add("B", 1);
           answerDictionary.Add("C", 2);
           answerDictionary.Add("D", 3);
           answerDictionary.Add("E", 4);    
       }

       public void ProcessExcel(string ExcelPath,XmlWriter writer)
       {
           string connectionString = "";
           if (ExcelPath.Length>0)
           {
               string fileName = ExcelPath;
               string fileExtension = Path.GetExtension(ExcelPath);
               
               //Check whether file extension is xls or xslx

               if (fileExtension == ".xls")
               {
                   connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ExcelPath +
                                      ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
               }
               else if (fileExtension == ".xlsx")
               {
                   connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ExcelPath +
                                      ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
               }

               //Create OleDB Connection and OleDb Command

               OleDbConnection con = new OleDbConnection(connectionString);
               OleDbCommand cmd = new OleDbCommand();
               cmd.CommandType = System.Data.CommandType.Text;
               cmd.Connection = con;
               OleDbDataAdapter dAdapter = new OleDbDataAdapter(cmd);
               DataTable dtExcelRecords = new DataTable();
               con.Open();
               DataTable dtExcelSheetName = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
               string getExcelSheetName = dtExcelSheetName.Rows[0]["Table_Name"].ToString();
               cmd.CommandText = "SELECT * FROM [" + getExcelSheetName + "]";
               dAdapter.SelectCommand = cmd;
               dAdapter.Fill(dtExcelRecords);

               questionDictionary = new Dictionary<Questions, List<Answers>>();
               int qno = 1;
               foreach (DataRow row in dtExcelRecords.Rows)
               {
                   int count = dtExcelRecords.Columns.Count;
                   Questions question = new Questions();
                   List<Answers> answers = new List<Answers>();
                   Answers ans = null;
                   for (int c = 0; c < count; c++)
                   {
                       //string str = row.Cells[c].Text.Trim().Replace("&#160;", " ").Replace("&quot;", "\"").Replace("&#34;", "\"").Replace("&#39;", "\'").Replace("&nbsp;", "").Replace("&lt;", "<").Replace("&gt;", ">");
                       string str = WebUtility.HtmlDecode(row[c].ToString());
                       
                       if (c == 0)
                       {
                           question.title = "Question " + qno;
                           question.questionText = str;
                           question.points = 1;
                           question.correctMsg = "Correct!";
                           question.incorrectMsg = "Sorry, that was incorrect.";
                           question.category = "Multichoice";
                           qno++;
                       }
                       else if (c < count - 1)
                       {
                           if (str.Trim().Length >= 1)
                           {
                               ans = new Answers();
                               ans.answerText = str;
                               ans.points = 1;
                               ans.stortText = "false";
                               ans.correct = "false";
                               answers.Add(ans);
                           }
                       }
                       if (c == count - 1)
                       {
                           string text = str;
                           if (answerDictionary.ContainsKey(text))
                           {
                               int answerIndex = answerDictionary[text];
                               text = row[answerIndex].ToString().Trim();
                               answers[answerIndex].correct = "true";
                           }
                       }


                   }
                   questionDictionary.Add(question, answers);
               }

               con.Close();

               GenerateXML(questionDictionary, writer);

           }
       }

       private void GenerateXML(Dictionary<Questions, List<Answers>> questionDictionary, XmlWriter writer)
       {
           writer.WriteStartElement("questions");
           foreach (KeyValuePair<Questions, List<Answers>> pair in questionDictionary)
           {
               Questions q = pair.Key;
               List<Answers> ans = pair.Value;
               

               writer.WriteStartElement("question");
               writer.WriteAttributeString("answerType","single");
               writer.WriteStartElement("title");
               writer.WriteValue(q.title);
               writer.WriteEndElement();

               writer.WriteStartElement("points");
               writer.WriteValue(q.points);
               writer.WriteEndElement();

               writer.WriteStartElement("questionText");
               writer.WriteValue(q.questionText);
               writer.WriteEndElement();

               writer.WriteStartElement("correctMsg");
               writer.WriteValue(q.correctMsg);
               writer.WriteEndElement();

               writer.WriteStartElement("incorrectMsg");
               writer.WriteValue(q.incorrectMsg);
               writer.WriteEndElement();

               writer.WriteStartElement("category");
               writer.WriteValue(q.category);
               writer.WriteEndElement();

               writer.WriteStartElement("answers");
               foreach (var an in ans)
               {
                   writer.WriteStartElement("answer");
                   writer.WriteAttributeString("points", an.points.ToString());
                   writer.WriteAttributeString("correct", an.correct);

                   writer.WriteStartElement("answerText");
                   writer.WriteAttributeString("html","false");
                   writer.WriteValue(an.answerText);
                   writer.WriteEndElement();

                   writer.WriteStartElement("stortText");
                   writer.WriteAttributeString("html", "false");
                   writer.WriteEndElement();
                   
                   writer.WriteEndElement(); //end of single answer
               }
               writer.WriteEndElement(); //end of answers


               writer.WriteEndElement(); //end of single question
               
           }
           writer.WriteEndElement(); //end of questions
       }
        
    }
}
