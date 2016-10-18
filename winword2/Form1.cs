using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Diagnostics;
using System.Drawing.Drawing2D;

namespace winword2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        //To find the bookmarks
        private void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object findText,
            object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref findText,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundLike,
                ref nmatchAllForms, ref forward,
                ref wrap, ref format, ref replaceWithText,
                ref replace, ref matchKashida,
                ref matchDiactitics, ref matchAlefHamza,
                ref matchControl);
        }


        //Methode Create the document :
        private void CreateWordDocument(object filename, object savaAs)
        {
            List<int> processesbeforegen = getRunningProcesses();
            object missing = Missing.Value;
            string tempPath = null;

            Word.Application wordApp = new Word.Application();

            Word.Document aDoc = null;

            if (File.Exists((string) filename))
            {
                

                object readOnly = false; 
                object isVisible = false;

                wordApp.Visible = false;

                aDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing);

                aDoc.Activate();

                //Find and replace:
                this.FindAndReplace(wordApp, "<eid>", textBoxID.Text);
                this.FindAndReplace(wordApp, "<firstname>", textBoxFname.Text);
                this.FindAndReplace(wordApp, "<middlename>", textBoxMName.Text);
                this.FindAndReplace(wordApp, "<lastname>", textBoxLName.Text);
                this.FindAndReplace(wordApp, "<gender>", textBoxGender.Text);
                this.FindAndReplace(wordApp, "<nickname>", textBoxNickName.Text);
                this.FindAndReplace(wordApp, "<cityzenship>", textBoxCityzenship.Text);
                this.FindAndReplace(wordApp, "<birthplace>", textBoxBirthPlace.Text);
                this.FindAndReplace(wordApp, "<homeadd1>", textBoxHomeAdd1.Text);
                this.FindAndReplace(wordApp, "< homeadd2>", textBoxHomeAdd2.Text);
                this.FindAndReplace(wordApp, "<country>", textBoxCountry.Text);
                this.FindAndReplace(wordApp, "<homephone>", textBoxHomePhone.Text);
                this.FindAndReplace(wordApp, "<cellphone>", textBoxCellPhone.Text);
                this.FindAndReplace(wordApp, "<homefax>", textBoxHomeFax.Text);
                this.FindAndReplace(wordApp, "<email>", textBoxEmail.Text);
                this.FindAndReplace(wordApp, "<birthdate>", textBoxBirthDate.Text);
                this.FindAndReplace(wordApp, "<nid>", textBoxNID.Text);
                this.FindAndReplace(wordApp, "<pno>", textBoxPNO.Text);

                //Save Document
                
                aDoc.SaveAs2(ref savaAs, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing);

                //Close Document:
				MessageBox.Show("File created.");
                List<int> processesaftergen = getRunningProcesses();
                killProcesses(processesbeforegen, processesaftergen);

            }

        }

        public List<int> getRunningProcesses()
        {
            List<int> ProcessIDs = new List<int>();
            foreach (Process clsProcess in Process.GetProcesses())
            {
                if (Process.GetCurrentProcess().Id == clsProcess.Id)
                    continue;
                if (clsProcess.ProcessName.Contains("WINWORD"))
                {
                    ProcessIDs.Add(clsProcess.Id);
                }
            }
            return ProcessIDs;
        }

        private void killProcesses(List<int> processesbeforegen, List<int> processesaftergen)
        {
            foreach (int pidafter in processesaftergen)
            {
                bool processfound = false;
                foreach (int pidbefore in processesbeforegen)
                {
                    if (pidafter == pidbefore)
                    {
                        processfound = true;
                    }
                }

                if (processfound == false)
                {
                    Process clsProcess = Process.GetProcessById(pidafter);
                    clsProcess.Kill();
                }
            }
        }

        //Enable controls
        private void tEnabled(bool state)
        {
            textBoxID.Enabled = state;
            textBoxFname.Enabled = state;
            textBoxMName.Enabled = state;
            textBoxLName.Enabled = state;
            textBoxGender.Enabled = state;
            textBoxNickName.Enabled = state;
            textBoxCityzenship.Enabled = state;
            textBoxBirthPlace.Enabled = state;
            textBoxHomeAdd1.Enabled = state;
            textBoxHomeAdd2.Enabled = state;
            textBoxCountry.Enabled = state;
            textBoxHomePhone.Enabled = state;
            textBoxCellPhone.Enabled = state;
            textBoxHomeFax.Enabled = state;
            textBoxEmail.Enabled = state;
            textBoxBirthDate.Enabled = state;
            textBoxNID.Enabled = state;
            textBoxPNO.Enabled = state;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (saveDoc.ShowDialog() == DialogResult.OK)
            {
                CreateWordDocument(textBoxFilePath.Text, saveDoc.FileName);
                tEnabled(false);
            }
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            if (loadDoc.ShowDialog() == DialogResult.OK)
            {
                textBoxFilePath.Text = loadDoc.FileName;
                tEnabled(true);
            }
        }
    }
}
