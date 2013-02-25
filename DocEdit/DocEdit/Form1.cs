﻿using System;
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
    public enum Theme
    {
        LIGHT, DARK
    }

    public partial class frmMain : Form
    {
        #region Theme Variables
        Theme theme = Theme.LIGHT;

        Color formBackColor = Color.FromArgb(242, 242, 249);
        Color menuStripsColor = Color.FromArgb(246, 251, 243);
        Color toolStripsForeColor = Color.Black;
        Color buttonsNoHoverColor = Color.FromArgb(189, 178, 183);
        Color buttonsHoverColor = ColorTranslator.FromHtml("#757072");
        Color buttonsHoverUnloadColor = ColorTranslator.FromHtml("#F01D1D");
        Color pBarBackColor = Color.Gainsboro;
        Color pBarForeColor = Color.FromArgb(255, 128, 0);
        Color textColor = Color.Black;

        Image imgXButton;
        Image imgXButtonHover;
        Image imgMinButton;
        Image imgMinButtonHover;
        #endregion

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

        // Pont of reference when dragging the form
        Point dragOffset;
        #endregion

        public frmMain()
        {
            InitializeComponent();
            this.Size = new Size(371, 481);
            loadedDoc = new Microsoft.Office.Interop.Word.Document();
            tmpDoc = new Microsoft.Office.Interop.Word.Document();
            RefreshTheme();
        }

        #region Controls
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
            if (hasLoaded)
                saveFileDialog.ShowDialog();
            else
                PrintErrorMessage("No file has been loaded!");
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
            if (hasLoaded)
            {
                if (chkDelSlideNums.Checked)
                    DeleteSlideNums();

                if (chkScaleSlides.Checked)
                    ScaleImage();

                SaveTempPDF();
                MessageBox.Show(this, "Finished executing tasks.", "Finished", MessageBoxButtons.OK, MessageBoxIcon.Information);
                pBar1.Value = 100;
            }
            else
                PrintErrorMessage("No file is loaded!");
        }

        private void btnPreview_Click(object sender, EventArgs e)
        {
            if (hasLoaded)
                PreviewPDF();
            else
                PrintErrorMessage("No file has been loaded!");
        }

        private void btnUnload_Click(object sender, EventArgs e)
        {
            if (hasLoaded)
            {
                if (MessageBox.Show(this, "WARNING! This feature is very unstable.  Will you still continue?",
                  "WARNING", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                    Reset();
            }
            else
                PrintErrorMessage("No file has been loaded!");
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new AboutBox1().ShowDialog();
        }
        private void lightThemeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!(theme == Theme.LIGHT))
            {
                theme = Theme.LIGHT;
                RefreshTheme();
            }
            else
            {
                PrintErrorMessage("This theme is already selected.");
            }
        }
        private void darkThemeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!(theme == Theme.DARK))
            {
                theme = Theme.DARK;
                RefreshTheme();
            }
            else
            {
                PrintErrorMessage("This theme is already selected.");
            }
        }

        // The x button control box and minimize control box
        private void closeFormButtton_MouseEnter(object sender, EventArgs e)
        {
            closeFormButtton.Image = imgXButtonHover;
        }
        private void closeFormButtton_MouseLeave(object sender, EventArgs e)
        {
            closeFormButtton.Image = imgXButton;
        }
        private void closeFormButtton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void minimizeBox_MouseEnter(object sender, EventArgs e)
        {
            minimizeBox.Image = imgMinButtonHover;
        }
        private void minimizeBox_MouseLeave(object sender, EventArgs e)
        {
            minimizeBox.Image = imgMinButton;
        }
        private void minimizeBox_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        // Dragging the form during run time
        private void menuStrip1_MouseDown(object sender, MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button == MouseButtons.Left)
            {
                dragOffset = this.PointToScreen(e.Location);
                var formLocation = FindForm().Location;
                dragOffset.X -= formLocation.X;
                dragOffset.Y -= formLocation.Y;
            }
        }
        private void menuStrip1_MouseMove(object sender, MouseEventArgs e)
        {
            base.OnMouseMove(e);

            if (e.Button == MouseButtons.Left)
            {
                Point newLocation = this.PointToScreen(e.Location);

                newLocation.X -= dragOffset.X;
                newLocation.Y -= dragOffset.Y;

                FindForm().Location = newLocation;
            }
        }
        // Changing button colors when the mouse enters
        private void btn_MouseEnter(object sender, EventArgs e)
        {
            Control control = sender as Control;
            if (!(sender.Equals(btnUnload)))
                control.BackColor = buttonsHoverColor;
            else
                control.BackColor = buttonsHoverUnloadColor;
        }
        private void btn_MouseLeave(object sender, EventArgs e)
        {
            Control control = sender as Control;
            control.BackColor = buttonsNoHoverColor;
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

                pBar1.Value = (i / pgNums) * 50;
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
                pBar1.Value = tmpStatusVal + 50;

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
                pBar1.Value = 0;
            }
            return true;
        }
        #endregion

        #region Customized GUI
        private void RefreshTheme()
        {
            switch (theme)
            {
                case (Theme.DARK):
                    formBackColor = Color.FromArgb(63, 63, 63);
                    buttonsHoverColor = Color.FromArgb(140, 209, 255);
                    buttonsNoHoverColor = Color.FromArgb(77, 184, 255);
                    buttonsHoverUnloadColor = Color.FromArgb(240, 29, 29);
                    menuStripsColor = Color.FromArgb(4, 4, 4);
                    toolStripsForeColor = Color.White;
                    pBarBackColor = Color.DarkGray;
                    pBarForeColor = Color.FromArgb(177, 255, 125);
                    textColor = Color.White;

                    imgXButton = DocEdit.Properties.Resources.X;
                    imgXButtonHover = DocEdit.Properties.Resources.X_Hover;
                    imgMinButton = DocEdit.Properties.Resources.Minimize;
                    imgMinButtonHover = DocEdit.Properties.Resources.Minimize_Hover;

                    break;

                case (Theme.LIGHT):
                    formBackColor = Color.FromArgb(242, 242, 249);
                    menuStripsColor = Color.FromArgb(246, 251, 243);
                    toolStripsForeColor = Color.Black;
                    buttonsNoHoverColor = Color.FromArgb(189, 178, 183);
                    buttonsHoverColor = ColorTranslator.FromHtml("#757072");
                    buttonsHoverUnloadColor = ColorTranslator.FromHtml("#F01D1D");
                    pBarBackColor = Color.Gainsboro;
                    pBarForeColor = Color.FromArgb(255, 128, 0);
                    textColor = Color.Black;

                    imgXButton = DocEdit.Properties.Resources.XLight;
                    imgXButtonHover = DocEdit.Properties.Resources.X_HoverLight;
                    imgMinButton = DocEdit.Properties.Resources.MinimizeLight;
                    imgMinButtonHover = DocEdit.Properties.Resources.Minimize_HoverLight;

                    break;
            }

            this.BackColor = formBackColor;

            foreach (Control c in this.Controls)
            {
                if (c is Button)
                {
                    c.BackColor = buttonsNoHoverColor;
                    c.ForeColor = Color.Black;
                }
                else if (c is MenuStrip)
                {
                    foreach (ToolStripMenuItem t in menuStrip1.Items)
                    {
                        foreach (ToolStripMenuItem i in t.DropDownItems)
                        {
                            t.BackColor = menuStripsColor;
                            t.ForeColor = toolStripsForeColor;
                            i.BackColor = menuStripsColor;
                            i.ForeColor = toolStripsForeColor;
                        }
                    }
                    c.BackColor = menuStripsColor;
                }
                c.ForeColor = textColor;
            }

            foreach (CheckBox c in groupBox1.Controls)
            {
                c.ForeColor = textColor;
            }

            pBar1.BackColor = pBarBackColor;
            pBar1.ForeColor = pBarForeColor;
            statusStrip1.BackColor = menuStripsColor;
            statusStrip1.ForeColor = textColor;
            closeFormButtton.Image = imgXButton;
            minimizeBox.Image = imgMinButton;
        }
        #endregion
    }
}
