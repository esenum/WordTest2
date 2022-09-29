using System;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Word;//<- this is what I am talking about
using System.Reflection;
using Microsoft.Office.Core;

namespace WordTest2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            //  create offer letter
            try
            {
                //  Just to kill WINWORD.EXE if it is running
                //  killprocess("winword");
                //  copy letter format to temp.doc
                File.Copy("C:\\Users\\HP\\Desktop\\OfferLetter.docx", "c:\\temp.docx", true);
                //  create missing object
                object missing = Missing.Value;
                //  create Word application object
                Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application(); // for other solution ..ApplicationClass();

                //Type wordType = Type.GetTypeFromProgID("Word.Application");
                //dynamic wordApp = Activator.CreateInstance(wordType);

                //  create Word document object
                Document aDoc = null;
                //  create & define filename object with temp.docx
                object filename = "c:\\temp.docx";
                //  if temp.doc available
                if (File.Exists((string)filename))
                {
                    object readOnly = false;
                    object isVisible = false;
                    //  make visible Word application
                    wordApp.Visible = false;
                    //  open Word document named temp.doc
                    aDoc = wordApp.Documents.Open(ref filename, ref missing,
                    ref readOnly, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref isVisible, ref missing, ref missing,
                    ref missing, ref missing);
                    aDoc.Activate();
                    //  Call FindAndReplace()function for each change
                    this.FindAndReplace(wordApp, "<Date>", "");
                    this.FindAndReplace(wordApp, "<Name>", "");
                    this.FindAndReplace(wordApp, "<Subject>", "");
                    //  save temp.doc after modified
                    aDoc.Save();
                }
                else
                    MessageBox.Show("File does not exist.",
                    "No File", MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                    //killprocess("winword");
            }
            catch (Exception ex)
            {
                     MessageBox.Show("Error in process.", ex.ToString(),
                     MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object findText, object replaceText)
        {
                object matchCase = true;
                object matchWholeWord = true;
                object matchWildCards = false;
                object matchSoundsLike = false;
                object matchAllWordForms = false;
                object forward = true;
                object format = false;
                object matchKashida = false;
                object matchDiacritics = false;
                object matchUmut = false;
                object matchControl = false;
                object read_only = false;
                object visible = true;
                object replace = 2;
                object wrap = 1;
                wordApp.Selection.Find.Execute(ref findText, ref matchCase,
                ref matchWholeWord, ref matchWildCards, ref matchSoundsLike,
                ref matchAllWordForms, ref forward, ref wrap, ref format,
                ref replaceText, ref replace, ref matchKashida,
                        ref matchDiacritics,
                ref matchUmut, ref matchControl);
        }
    }
    
}

