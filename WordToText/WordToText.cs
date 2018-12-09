//Added references:
//DocumentFormat.OpenXml.dll for "WordprocessingDocument"
//WindowsBase.dll for "using"

using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace WordToText
{
    public partial class WordToText : Form
    {
        public WordToText()
        {
            InitializeComponent();
        }

        List<String> WordFilesPathes = new List<String>();
        private void btnSelectWordFiles_Click(object sender, EventArgs e)
        {
            try
            {
                var dlg = new OpenFileDialog();
                dlg.Filter = "Word Documents|*.doc;*.docx";
                dlg.Multiselect = true;

                if (dlg.ShowDialog() != DialogResult.OK)
                    return;

                WordFilesPathes = dlg.FileNames.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Occured: " + ex.Message);
            }
        }

        private void btnExtract_Click(object sender, EventArgs e)
        {
            try
            {
                if (WordFilesPathes.Count <= 0)
                {
                    MessageBox.Show("No files are selected!");
                    return;
                }

                string selectedDestination;
                using (var fbd = new FolderBrowserDialog())
                {
                    fbd.Description = "Save Extracted Text File At:";

                    DialogResult result = fbd.ShowDialog();
                    if (result != DialogResult.OK || string.IsNullOrWhiteSpace(fbd.SelectedPath))
                    {
                        return;
                    }
                    selectedDestination = fbd.SelectedPath;
                }

                List<FileInfo> files = new List<FileInfo>();
                //add files into list
                foreach (var path in WordFilesPathes)
                {
                    FileInfo file = new FileInfo(path);

                    string fileName = file.Name; //file name with extension
                    string targetFileName = Path.GetFileNameWithoutExtension(fileName) + ".txt"; //target file name with new extension

                    string txt = getTextFromWord(path);//extract text from Word File
                    if (String.IsNullOrEmpty(txt))
                    {
                        //MessageBox.Show("Empty File!");
                        continue;
                    }

                    //save extracted text into .txt file
                    File.WriteAllText(selectedDestination + "\\" + targetFileName, txt);
                }

                WordFilesPathes.Clear();
                MessageBox.Show("Done!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Occured: " + ex.Message);
            }
        }

        public string getTextFromWord(string filePath)
        {
            try
            {
                //shout out to KyleM - Stack Overflow
                const string wordmlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

                StringBuilder textBuilder = new StringBuilder();
                using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(filePath, false))
                {
                    // Manage namespaces to perform XPath queries.  
                    NameTable nt = new NameTable();
                    XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                    nsManager.AddNamespace("w", wordmlNamespace);

                    // Get the document part from the package.  
                    // Load the XML in the document part into an XmlDocument instance.  
                    XmlDocument xdoc = new XmlDocument(nt);
                    xdoc.Load(wdDoc.MainDocumentPart.GetStream());

                    XmlNodeList paragraphNodes = xdoc.SelectNodes("//w:p", nsManager);
                    foreach (XmlNode paragraphNode in paragraphNodes)
                    {
                        XmlNodeList textNodes = paragraphNode.SelectNodes(".//w:t", nsManager);
                        foreach (System.Xml.XmlNode textNode in textNodes)
                        {
                            textBuilder.Append(textNode.InnerText);
                        }
                        textBuilder.Append(Environment.NewLine);
                    }

                    wdDoc.MainDocumentPart.GetStream().Flush();
                    wdDoc.MainDocumentPart.GetStream().Close();
                }

                return textBuilder.ToString();
            }
            catch (Exception e)
            {
                MessageBox.Show("Error Occured on file: " + filePath + ". " + e.Message);
                return "";
            }
        }

        private void WordToText_Load(object sender, EventArgs e)
        {

        }
 
    }
}
