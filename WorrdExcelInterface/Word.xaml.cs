using Microsoft.Win32;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.IO;
using Spire.Doc;
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



namespace WorrdExcelInterface
{



    public partial class Word : Window
    {

        private void SaveRtfFile(string _fileName)
        {
            Spire.Doc.Document doc = new Spire.Doc.Document();

            TextRange range = new TextRange(Myrtb.Document.ContentStart, Myrtb.Document.ContentEnd);
            FileStream fstream = new FileStream("converted.rtf", FileMode.Create);
            range.Save(fstream, DataFormats.Rtf);
            fstream.Close();
            doc.LoadFromFile("converted.rtf");
            doc.SaveToFile(_fileName);

        }

        private  void LoadRtfFile(string _fileName)
        {

            if (File.Exists(_fileName))
            {

                TextRange range = new TextRange(Myrtb.Document.ContentStart, Myrtb.Document.ContentEnd);
                FileStream fileStream = new FileStream(_fileName, FileMode.OpenOrCreate);
                range.Load(fileStream, DataFormats.Rtf);
                fileStream.Close();


            }
        }




        public Word()
        {
            InitializeComponent();
        }



        private void SendEmailWord_Click(object sender, RoutedEventArgs e)
        {
            Spire.Doc.Document doc = new Spire.Doc.Document();
            SaveRtfFile("WordMessage.docx");

            MailInterface mailInterface = new MailInterface();
            mailInterface.Show();



        }

        private void SaveFile_Click(object sender, RoutedEventArgs e)
        {

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Text files (*.docx)|*.docx|All files (*.*)|*.*";

            if (saveFileDialog1.ShowDialog() == true)
            {
                string filePath = saveFileDialog1.FileName;
                SaveRtfFile(filePath);
            }
        }

        private void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            Spire.Doc.Document doc = new Spire.Doc.Document();

            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Filter = "Text files (*.docx)|*.docx|All files (*.*)|*.*";

            if (openFileDialog2.ShowDialog() == true)
            {
                string filePath = openFileDialog2.FileName;
                doc.LoadFromFile(filePath);
                doc.SaveToFile("Converted.rtf", Spire.Doc.FileFormat.Rtf);
                LoadRtfFile("Converted.rtf");
            }
        }
    }
}
