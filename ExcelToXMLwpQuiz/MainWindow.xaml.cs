using Microsoft.Win32;
using System.Windows;
using System.Xml;

namespace ExcelToXMLwpQuiz
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {


        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
                lblFileName.Content = openFileDialog.FileName;
        }

        private void cmdGenerate_Click(object sender, RoutedEventArgs e)
        {
            using (XmlWriter writer = XmlWriter.Create("Quiz.xml"))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("wpProQuiz");
                writer.WriteStartElement("header");
                writer.WriteAttributeString("version", "0.28");
                writer.WriteAttributeString("exportVersion", "1");
                writer.WriteEndElement();

                writer.WriteStartElement("data");// start of data element
                writer.WriteStartElement("quiz");

                #region Header Element

                writer.WriteStartElement("title");
                writer.WriteAttributeString("titleHidden", "false");
                writer.WriteValue(txtQuizName.Text);
                writer.WriteEndElement();

                writer.WriteElementString("text", txtQuizName.Text);

                writer.WriteStartElement("resultText");
                writer.WriteAttributeString("gradeEnabled", "false");
                writer.WriteEndElement();

                writer.WriteStartElement("btnRestartQuizHidden");
                writer.WriteValue("false");
                writer.WriteEndElement();

                writer.WriteStartElement("btnViewQuestionHidden");
                writer.WriteValue("false");
                writer.WriteEndElement();

                writer.WriteStartElement("questionRandom");
                writer.WriteValue("true");
                writer.WriteEndElement();


                writer.WriteStartElement("answerRandom");
                writer.WriteValue("true");
                writer.WriteEndElement();

                writer.WriteElementString("timeLimit", "");


                writer.WriteStartElement("showPoints");
                writer.WriteValue("true");
                writer.WriteEndElement();

                writer.WriteStartElement("quizRunOnce");
                writer.WriteAttributeString("type", "1");
                writer.WriteAttributeString("cookie", "false");
                writer.WriteAttributeString("time", "0");
                writer.WriteValue("false");
                writer.WriteEndElement();

                writer.WriteStartElement("numberedAnswer");
                writer.WriteValue("true");
                writer.WriteEndElement();


                writer.WriteStartElement("hideAnswerMessageBox");
                writer.WriteValue("false");
                writer.WriteEndElement();


                writer.WriteStartElement("disabledAnswerMark");
                writer.WriteValue("true");
                writer.WriteEndElement();


                writer.WriteStartElement("toplist");
                writer.WriteAttributeString("activated", "true");

                writer.WriteStartElement("toplistDataAddPermissions");
                writer.WriteValue("1");
                writer.WriteEndElement();

                writer.WriteStartElement("toplistDataSort");
                writer.WriteValue("1");
                writer.WriteEndElement();

                writer.WriteStartElement("toplistDataAddMultiple");
                writer.WriteValue("true");
                writer.WriteEndElement();

                writer.WriteStartElement("toplistDataAddBlock");
                writer.WriteValue("1");
                writer.WriteEndElement();

                writer.WriteStartElement("toplistDataShowLimit");
                writer.WriteValue("10");
                writer.WriteEndElement();


                writer.WriteStartElement("toplistDataShowIn");
                writer.WriteValue("1");
                writer.WriteEndElement();

                writer.WriteStartElement("toplistDataCaptcha");
                writer.WriteValue("false");
                writer.WriteEndElement();

                writer.WriteStartElement("toplistDataAddAutomatic");
                writer.WriteValue("true");
                writer.WriteEndElement();

                writer.WriteEndElement();

                writer.WriteStartElement("skipQuestionDisabled");
                writer.WriteValue("true");
                writer.WriteEndElement();


                writer.WriteStartElement("showCategoryScore");
                writer.WriteValue("true");
                writer.WriteEndElement();

                writer.WriteStartElement("hideResultPoints");
                writer.WriteValue("false");
                writer.WriteEndElement();

                writer.WriteStartElement("autostart");
                writer.WriteValue("false");
                writer.WriteEndElement();

                writer.WriteStartElement("forcingQuestionSolve");
                writer.WriteValue("true");
                writer.WriteEndElement();

                writer.WriteStartElement("hideQuestionPositionOverview");
                writer.WriteValue("true");
                writer.WriteEndElement();

                writer.WriteStartElement("hideQuestionNumbering");
                writer.WriteValue("false");
                writer.WriteEndElement();

                writer.WriteStartElement("sortCategories");
                writer.WriteValue("true");
                writer.WriteEndElement();

                writer.WriteStartElement("showCategory");
                writer.WriteValue("true");
                writer.WriteEndElement();

                writer.WriteStartElement("quizModus");
                writer.WriteAttributeString("questionsPerPage", "");
                writer.WriteValue("3");
                writer.WriteEndElement();

                writer.WriteStartElement("startOnlyRegisteredUser");
                writer.WriteValue("false");
                writer.WriteEndElement();

                #endregion
                
                WriteQuestions q = new WriteQuestions();
                q.ProcessExcel(lblFileName.Content.ToString(),writer);


                writer.WriteEndElement(); //end of quiz
                writer.WriteEndElement(); //end of data element


                writer.WriteEndElement();
                writer.WriteEndDocument();
            }

        }
    }
}
