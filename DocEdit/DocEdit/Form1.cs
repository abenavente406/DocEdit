using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DocEdit
{
    public partial class Form1 : Form
    {
        #region Fields

        
            //        Dim wordApp As New Word.Application

            //Dim loadedDoc As Word.Document
            //Dim tmpDoc As New Word.Document

            //Dim pgNums As Integer
            //Dim tmpFilePathDir As String = My.Application.Info.DirectoryPath

            //Dim hasLoaded As Boolean = False
        #endregion

        public Form1()
        {
            InitializeComponent();
        }

        private void btnExecute_Click(object sender, EventArgs e)
        {

        }

        #region Tasks
        private void DeleteSlideNums()
        {

            //    For i As Integer = 1 To pgNums
            //        tmpDoc.Content.Find.Execute(FindText:=i.ToString & vbCr & vbCr, ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)

            //        pbStatus.Value = (i / pgNums) * 50
            //        lblPercent.Text = pbStatus.Value & "%"
            //    Next

            //    tmpDoc.Content.Find.Execute(FindText:="Slide", ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)

            //End Sub
        }

        private void ScaleImage()
        {
            //        Private Sub scaleImage()
            //    Static i As Integer = 0

            //    Dim pict As Word.InlineShape
            //    Dim tmpStatusVal As Integer = 0

            //    For Each pict In tmpDoc.InlineShapes
            //        pict.ScaleHeight = 100
            //        pict.ScaleWidth = 100

            //        tmpStatusVal = (i / pgNums) * 50
            //        pbStatus.Value = tmpStatusVal + 50
            //        lblPercent.Text = pbStatus.Value & "%"

            //        i += 1
            //    Next
            //End Sub
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
            //    Dim tmpFilePath As String = My.Application.Info.DirectoryPath & "\preview.pdf"
            //    tmpDoc.SaveAs2(FileName:=tmpFilePath, FileFormat:=Word.WdSaveFormat.wdFormatPDF)
            //    updatePdfReader(tmpFilePath)
        }

        private void UpdatePDFReader()
        {
            //    pdfReader.LoadFile(tmpFilePath)
            //    pdfReader.Refresh()
        }

        private void PreviewPDF()
        {
            //        Private Sub previewPDF()

            //    ' Create the temp file called preview.pdf
            //    Dim tmpFilePath As String = My.Application.Info.DirectoryPath & "\tmpPreview.pdf"
            //    tmpDoc.SaveAs2(FileName:=tmpFilePath, FileFormat:=Word.WdSaveFormat.wdFormatPDF)
            //    ' Start and wait for the process to finish
            //    Dim tmpProcess As Process = Process.Start(tmpFilePath)
            //    tmpProcess.WaitForExit()

            //    ' Check if the tmp file can be deleted
            //    Try
            //        Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile(tmpFilePath)
            //    Catch ex As Exception
            //        printErrorMessage(ex.Message)
            //    End Try

            //End Sub
        }

        private void Reset()
        {
            //            Private Sub resetProgram()
            //    attemptCloseProgram()
            //    wordApp = New Word.Application

            //    lblLoad.Text = "Loaded Document: None"
            //    pgNums = 0
            //    lblPages.Text = "Pages in Document: " & pgNums.ToString
            //    stslblFileLoaded.Text = "File Loaded: False"

            //End Sub
        }

        private bool AttemptCloseProgram()
        {
            //        Private Sub attemptCloseProgram()
            //    If hasLoaded Then
            //        Try
            //            loadedDoc.Close()
            //        Catch ex As Exception
            //            printErrorMessage("Couldn't close loadedDoc.", ex)
            //        End Try
            //        Try
            //            tmpDoc.SaveAs2(My.Application.Info.DirectoryPath & "\temp", Word.WdSaveFormat.wdFormatDocument97)
            //        Catch ex As Exception
            //            printErrorMessage("Could not save tmpDoc.")
            //        End Try
            //        Try
            //            Debug.WriteLine(tmpDoc.Path)
            //            tmpDoc.Close()
            //        Catch ex As Exception
            //            printErrorMessage("Could not close tmpDoc.")
            //        End Try
            //        Try
            //            wordApp.Quit()
            //        Catch ex As Exception
            //            printErrorMessage("FATAL ERROR. Could not quit wordApp_1. Go to task manager and close the task.")
            //        End Try
            //        Try
            //            Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile(tmpFilePathDir & "\preview.pdf")
            //        Catch ex As Exception
            //            printErrorMessage("Could not find the temporary file")
            //        End Try
            //        Try
            //            Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile(tmpFilePathDir & "\temp.doc")
            //        Catch ex As Exception
            //            printErrorMessage("Could not find " & tmpDoc.Path)
            //        End Try

            //        hasLoaded = False

            //        pbStatus.Value = 0
            //        lblPercent.Text = "0%"
            //    End If
            //End Sub
            return true;
        }
        #endregion
    }
}
