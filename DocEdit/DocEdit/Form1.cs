using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Office;

namespace DocEdit
{
    public partial class frmMain : Form
    {
        #region Fields
        // The word application to handle MSWord operations
        Microsoft.Office.Interop.Word.Application wordApp = 
            new Microsoft.Office.Interop.Word.Application();

        // Create the documents to be manipulated
        Microsoft.Office.Interop.Word.Document loadedDoc;
        Microsoft.Office.Interop.Word.Document tmpDoc;

        private int   pgNums   = 0;
        private bool hasLoaded = false;
        private string tmpFilePath = System.IO.Directory.GetCurrentDirectory();
        #endregion

        #region Controls
        public frmMain()
        {
            InitializeComponent();
            this.Size = new Size(381, 481);
            loadedDoc = new Microsoft.Office.Interop.Word.Document();
            tmpDoc = new Microsoft.Office.Interop.Word.Document();
        }

        private void btnLoadFile_Click(object sender, EventArgs e)
        {
            openFileDialog.ShowDialog();
        }
        private void openFileDialog_FileOk(object sender, CancelEventArgs e)
        {
            try
            {
                var fileInfo = new System.IO.FileInfo(openFileDialog.FileName);

                // If the file is bigger than 30 mb
                if (fileInfo.Length * (9.35 * Math.Pow(10, -7)) > 30)
                    MessageBox.Show("This is a relatively large file.  Please allow load times in between 1 - 2 minutes.",
                        "BIG FILE", MessageBoxButtons.OK, MessageBoxIcon.Information);

                loadedDoc = wordApp.Documents.Open(openFileDialog.FileName);
                tmpDoc = wordApp.Documents.Add();

                loadedDoc.Content.Copy();
                tmpDoc.Content.Paste();

                hasLoaded = true;

                this.Size = new Size(Screen.GetWorkingArea(this.Bounds).Width,
                    Screen.GetWorkingArea(this.Bounds).Height);
                this.Location = new Point(0, 0);

                lblLoadedFile.Text = "Loaded Doc: " + loadedDoc.Path + "\\" + loadedDoc.Name;
                stslblFileLoaded.Text = "File Loaded: True";
                pgNums = loadedDoc.ComputeStatistics(Microsoft.Office.Interop.Word.WdStatistic.wdStatisticPages);
                lblPagesInDoc.Text = "Pages in Document: " + pgNums.ToString();

                pdfReader.Visible = true;

                SaveTempPDF();
                pdfReader.Refresh();
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                PrintErrorMessage("ERROR! The document failed to load.  Either" +
                    " the document is in use or you don't have permission to open it.", ex);
            }
            catch (Exception ex)
            {
                PrintErrorMessage("ERROR! Could not open the file.", ex);
            }
        }

        private void btnSaveFile_Click(object sender, EventArgs e)
        {
            saveFileDialog.ShowDialog();
        }
        private void saveFileDialog_FileOk(object sender, CancelEventArgs e)
        {
            try
            {
                if (saveFileDialog.FilterIndex == 0)
                {
                    // Save the document as a pdf
                    tmpDoc.SaveAs2(saveFileDialog.FileName, FileFormat:
                        Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF);
                }
                else
                {
                    // Save the document as a .doc
                    tmpDoc.SaveAs2(saveFileDialog.FileName, FileFormat:
                        Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault);
                }
            }
            catch (Exception ex)
            {
                PrintErrorMessage("ERROR! Could not save the file.", ex);
            }
        }

        private void btnExecute_Click(object sender, EventArgs e)
        {
            if (chkDelSlideNums.Checked)
                DeleteSlideNums();

            if (chkScaleSlides.Checked)
                ScaleImage();

            SaveTempPDF();
            MessageBox.Show(this, "Finished executing tasks.", "Finished", MessageBoxButtons.OK, MessageBoxIcon.Information);

            pbExecutionStatus.Value = 100;
            lblPercentage.Text = "100%";
        }

        private void btnPreview_Click(object sender, EventArgs e)
        {
            PreviewPDF();
        }
        
        private void btnUnload_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(this, "WARNING! This feature is very unstable.  Will you still continue?",
                "WARNING", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                Reset();
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (hasLoaded)
                if (AttemptCloseProgram())
                    Debug.Print("Closed with no errors!");

            try
            {
                foreach (Process p in Process.GetProcessesByName("winword"))
                {
                    p.Kill();
                    p.WaitForExit();
                }
            }
            catch (InvalidOperationException ex)
            {
                PrintErrorMessage("ERROR! A winword task doesn't exist!", ex);
            }
            catch (Exception ex)
            {
                PrintErrorMessage("ERROR! Could not exit a winword task.", ex);
            }
        }
        
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new AboutBox1().ShowDialog();
        }
        #endregion

        #region Tasks
        private void DeleteSlideNums()
        {
            // Find and delete the numbers + 
            for (int i = 1; i <= pgNums; i++)
            {
                tmpDoc.Content.Find.Execute(FindText:"Slide " + i.ToString() + "\r\r", ReplaceWith: "",
                    Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll,
                    Wrap: Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue);

                pbExecutionStatus.Value = (i / pgNums) * 50;
                lblPercentage.Text = pbExecutionStatus.Value.ToString() + "%";
            }
        }

        private void ScaleImage()
        {
            int tmpStatusVal = 0;
            int    counter   = 0;

            foreach (Microsoft.Office.Interop.Word.InlineShape pict in tmpDoc.InlineShapes)
            {
                pict.ScaleWidth = 100;
                pict.ScaleHeight = 100;

                tmpStatusVal = (counter / pgNums) * 50;
                pbExecutionStatus.Value = tmpStatusVal + 50;
                lblPercentage.Text = pbExecutionStatus.Value.ToString() + "%";

                counter += 1;
            }
        }
        #endregion

        #region Helper Functions

        private void PrintErrorMessage(string message, Exception ex)
        {
            MessageBox.Show(this, message + "\n\n" + ex.Message, "ERROR", 
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        private void PrintErrorMessage(string message)
        {
            MessageBox.Show(this, message, "ERROR", MessageBoxButtons.OK, 
                MessageBoxIcon.Error);
        }

        private void SaveTempPDF()
        {
            string tmpFilePathLocal = tmpFilePath + "\\preview.pdf";
            tmpDoc.SaveAs2(FileName:tmpFilePathLocal, FileFormat:Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF);
            UpdatePDFReader(tmpFilePathLocal);
        }

        private void UpdatePDFReader(string tempFilePath)
        {
            pdfReader.LoadFile(tempFilePath);
            pdfReader.Refresh();
        }

        private void PreviewPDF()
        {
            // Creates a temporary file called preview.pdf
            string tmpFilePath = this.tmpFilePath + "\\tmpPreview.pdf";
            tmpDoc.SaveAs2(FileName: tmpFilePath, FileFormat: Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF);

            Process adobeReader = Process.Start(tmpFilePath);
            adobeReader.WaitForExit();

            try
            {
                System.IO.File.Delete(tmpFilePath);
            }
            catch (IOException e)
            {
                Debug.Print("File not found! \n\n" + e.ToString());
            }
        }

        private void Reset()
        {
            if (AttemptCloseProgram())
                Debug.Print("Closed with no errors!");
            wordApp = new Microsoft.Office.Interop.Word.Application();
            
            pgNums = 0;
            lblLoadedFile.Text    = "Loaded Doc: None";
            stslblFileLoaded.Text = "File Loaded: False";
            lblPagesInDoc.Text    = "Pages in Document: " + pgNums.ToString();
        }

        private bool AttemptCloseProgram()
        {
            if (hasLoaded)
            {
                try
                {
                    loadedDoc.Close(false);
                    tmpDoc.Close(false);
                    System.IO.File.Delete(tmpFilePath + "\\preview.pdf");
                    System.IO.File.Delete(tmpFilePath + "\\temp.pdf");
                    System.IO.File.Delete(tmpFilePath + "\\temp.docx");
                }
                catch (Exception ex)
                {
                    Debug.Print(ex.ToString());
                    return false;
                }

                hasLoaded = false;
                pbExecutionStatus.Value = 0;
                lblPercentage.Text = "0%";
            }
            return true;
        }
        #endregion
    }
}
