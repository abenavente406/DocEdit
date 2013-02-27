namespace DocEdit
{
    partial class frmMain
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.stslblFileLoaded = new System.Windows.Forms.ToolStripStatusLabel();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveAsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.previewToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripThemeItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lightThemeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.darkThemeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkScaleSlides = new System.Windows.Forms.CheckBox();
            this.chkDelSlideNums = new System.Windows.Forms.CheckBox();
            this.btnLoadFile = new System.Windows.Forms.Button();
            this.btnSaveFile = new System.Windows.Forms.Button();
            this.btnPreview = new System.Windows.Forms.Button();
            this.lblPagesInDoc = new System.Windows.Forms.Label();
            this.lblLoadedFile = new System.Windows.Forms.Label();
            this.btnExecute = new System.Windows.Forms.Button();
            this.btnUnload = new System.Windows.Forms.Button();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.pdfReader = new AxAcroPDFLib.AxAcroPDF();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.pBar1 = new PBar.PBar();
            this.minimizeBox = new System.Windows.Forms.PictureBox();
            this.closeFormButtton = new System.Windows.Forms.PictureBox();
            this.statusStrip1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pdfReader)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.minimizeBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.closeFormButtton)).BeginInit();
            this.SuspendLayout();
            // 
            // statusStrip1
            // 
            this.statusStrip1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(246)))), ((int)(((byte)(251)))), ((int)(((byte)(243)))));
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.stslblFileLoaded});
            this.statusStrip1.Location = new System.Drawing.Point(0, 425);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(683, 22);
            this.statusStrip1.TabIndex = 0;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // stslblFileLoaded
            // 
            this.stslblFileLoaded.Font = new System.Drawing.Font("Frutiger Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.stslblFileLoaded.ForeColor = System.Drawing.Color.White;
            this.stslblFileLoaded.Name = "stslblFileLoaded";
            this.stslblFileLoaded.Size = new System.Drawing.Size(104, 17);
            this.stslblFileLoaded.Text = "File Loaded: False";
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(246)))), ((int)(((byte)(251)))), ((int)(((byte)(243)))));
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.viewToolStripMenuItem,
            this.toolStripThemeItem,
            this.helpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(683, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            this.menuStrip1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.menuStrip1_MouseDown);
            this.menuStrip1.MouseMove += new System.Windows.Forms.MouseEventHandler(this.menuStrip1_MouseMove);
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(246)))), ((int)(((byte)(251)))), ((int)(((byte)(243)))));
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem,
            this.saveAsToolStripMenuItem,
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Font = new System.Drawing.Font("Frutiger Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.fileToolStripMenuItem.ForeColor = System.Drawing.Color.Black;
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Padding = new System.Windows.Forms.Padding(0);
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(30, 20);
            this.fileToolStripMenuItem.Text = "&File";
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(246)))), ((int)(((byte)(251)))), ((int)(((byte)(243)))));
            this.openToolStripMenuItem.Font = new System.Drawing.Font("Frutiger Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.openToolStripMenuItem.ForeColor = System.Drawing.Color.Black;
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.Padding = new System.Windows.Forms.Padding(0);
            this.openToolStripMenuItem.Size = new System.Drawing.Size(126, 20);
            this.openToolStripMenuItem.Text = "Open...";
            this.openToolStripMenuItem.Click += new System.EventHandler(this.btnLoadFile_Click);
            // 
            // saveAsToolStripMenuItem
            // 
            this.saveAsToolStripMenuItem.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(246)))), ((int)(((byte)(251)))), ((int)(((byte)(243)))));
            this.saveAsToolStripMenuItem.Font = new System.Drawing.Font("Frutiger Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.saveAsToolStripMenuItem.ForeColor = System.Drawing.Color.Black;
            this.saveAsToolStripMenuItem.Name = "saveAsToolStripMenuItem";
            this.saveAsToolStripMenuItem.Padding = new System.Windows.Forms.Padding(0);
            this.saveAsToolStripMenuItem.Size = new System.Drawing.Size(126, 20);
            this.saveAsToolStripMenuItem.Text = "Save As...";
            this.saveAsToolStripMenuItem.Click += new System.EventHandler(this.btnSaveFile_Click);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(246)))), ((int)(((byte)(251)))), ((int)(((byte)(243)))));
            this.exitToolStripMenuItem.Font = new System.Drawing.Font("Frutiger Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.exitToolStripMenuItem.ForeColor = System.Drawing.Color.Black;
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Padding = new System.Windows.Forms.Padding(0);
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(126, 20);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // viewToolStripMenuItem
            // 
            this.viewToolStripMenuItem.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(246)))), ((int)(((byte)(251)))), ((int)(((byte)(243)))));
            this.viewToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.previewToolStripMenuItem});
            this.viewToolStripMenuItem.Font = new System.Drawing.Font("Frutiger Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.viewToolStripMenuItem.ForeColor = System.Drawing.Color.Black;
            this.viewToolStripMenuItem.Name = "viewToolStripMenuItem";
            this.viewToolStripMenuItem.Padding = new System.Windows.Forms.Padding(0);
            this.viewToolStripMenuItem.Size = new System.Drawing.Size(40, 20);
            this.viewToolStripMenuItem.Text = "&View";
            // 
            // previewToolStripMenuItem
            // 
            this.previewToolStripMenuItem.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(246)))), ((int)(((byte)(251)))), ((int)(((byte)(243)))));
            this.previewToolStripMenuItem.Font = new System.Drawing.Font("Frutiger Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.previewToolStripMenuItem.ForeColor = System.Drawing.Color.Black;
            this.previewToolStripMenuItem.Name = "previewToolStripMenuItem";
            this.previewToolStripMenuItem.Padding = new System.Windows.Forms.Padding(0);
            this.previewToolStripMenuItem.Size = new System.Drawing.Size(120, 20);
            this.previewToolStripMenuItem.Text = "Preview";
            this.previewToolStripMenuItem.Click += new System.EventHandler(this.btnPreview_Click);
            // 
            // toolStripThemeItem
            // 
            this.toolStripThemeItem.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(246)))), ((int)(((byte)(251)))), ((int)(((byte)(243)))));
            this.toolStripThemeItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.lightThemeToolStripMenuItem,
            this.darkThemeToolStripMenuItem});
            this.toolStripThemeItem.Font = new System.Drawing.Font("Frutiger Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.toolStripThemeItem.ForeColor = System.Drawing.Color.Black;
            this.toolStripThemeItem.Name = "toolStripThemeItem";
            this.toolStripThemeItem.Size = new System.Drawing.Size(58, 20);
            this.toolStripThemeItem.Text = "&Theme";
            // 
            // lightThemeToolStripMenuItem
            // 
            this.lightThemeToolStripMenuItem.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(246)))), ((int)(((byte)(251)))), ((int)(((byte)(243)))));
            this.lightThemeToolStripMenuItem.Font = new System.Drawing.Font("Frutiger Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lightThemeToolStripMenuItem.ForeColor = System.Drawing.Color.Black;
            this.lightThemeToolStripMenuItem.Name = "lightThemeToolStripMenuItem";
            this.lightThemeToolStripMenuItem.Size = new System.Drawing.Size(144, 22);
            this.lightThemeToolStripMenuItem.Text = "Light Theme";
            this.lightThemeToolStripMenuItem.Click += new System.EventHandler(this.lightThemeToolStripMenuItem_Click);
            // 
            // darkThemeToolStripMenuItem
            // 
            this.darkThemeToolStripMenuItem.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(246)))), ((int)(((byte)(251)))), ((int)(((byte)(243)))));
            this.darkThemeToolStripMenuItem.ForeColor = System.Drawing.Color.Black;
            this.darkThemeToolStripMenuItem.Name = "darkThemeToolStripMenuItem";
            this.darkThemeToolStripMenuItem.Size = new System.Drawing.Size(144, 22);
            this.darkThemeToolStripMenuItem.Text = "Dark Theme";
            this.darkThemeToolStripMenuItem.Click += new System.EventHandler(this.darkThemeToolStripMenuItem_Click);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(246)))), ((int)(((byte)(251)))), ((int)(((byte)(243)))));
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aboutToolStripMenuItem});
            this.helpToolStripMenuItem.Font = new System.Drawing.Font("Frutiger Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.helpToolStripMenuItem.ForeColor = System.Drawing.Color.Black;
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(45, 20);
            this.helpToolStripMenuItem.Text = "&Help";
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(246)))), ((int)(((byte)(251)))), ((int)(((byte)(243)))));
            this.aboutToolStripMenuItem.Font = new System.Drawing.Font("Frutiger Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.aboutToolStripMenuItem.ForeColor = System.Drawing.Color.Black;
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(118, 22);
            this.aboutToolStripMenuItem.Text = "About...";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chkScaleSlides);
            this.groupBox1.Controls.Add(this.chkDelSlideNums);
            this.groupBox1.Font = new System.Drawing.Font("Frutiger Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.ForeColor = System.Drawing.Color.Black;
            this.groupBox1.Location = new System.Drawing.Point(12, 36);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(346, 90);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Check the tasks you want enabled:";
            // 
            // chkScaleSlides
            // 
            this.chkScaleSlides.AutoSize = true;
            this.chkScaleSlides.Checked = true;
            this.chkScaleSlides.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkScaleSlides.Font = new System.Drawing.Font("Frutiger Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkScaleSlides.ForeColor = System.Drawing.Color.Black;
            this.chkScaleSlides.Location = new System.Drawing.Point(11, 55);
            this.chkScaleSlides.Name = "chkScaleSlides";
            this.chkScaleSlides.Size = new System.Drawing.Size(171, 19);
            this.chkScaleSlides.TabIndex = 1;
            this.chkScaleSlides.Text = "Change Slide Size to 100%";
            this.chkScaleSlides.UseVisualStyleBackColor = true;
            // 
            // chkDelSlideNums
            // 
            this.chkDelSlideNums.AutoSize = true;
            this.chkDelSlideNums.Checked = true;
            this.chkDelSlideNums.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDelSlideNums.Font = new System.Drawing.Font("Frutiger Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkDelSlideNums.ForeColor = System.Drawing.Color.Black;
            this.chkDelSlideNums.Location = new System.Drawing.Point(11, 24);
            this.chkDelSlideNums.Name = "chkDelSlideNums";
            this.chkDelSlideNums.Size = new System.Drawing.Size(147, 19);
            this.chkDelSlideNums.TabIndex = 1;
            this.chkDelSlideNums.Text = "Delete Slide Numbers";
            this.chkDelSlideNums.UseVisualStyleBackColor = true;
            // 
            // btnLoadFile
            // 
            this.btnLoadFile.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(189)))), ((int)(((byte)(178)))), ((int)(((byte)(183)))));
            this.btnLoadFile.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnLoadFile.Font = new System.Drawing.Font("Frutiger Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnLoadFile.Location = new System.Drawing.Point(12, 143);
            this.btnLoadFile.Name = "btnLoadFile";
            this.btnLoadFile.Size = new System.Drawing.Size(113, 24);
            this.btnLoadFile.TabIndex = 1;
            this.btnLoadFile.Text = "Load File";
            this.btnLoadFile.UseVisualStyleBackColor = false;
            this.btnLoadFile.Click += new System.EventHandler(this.btnLoadFile_Click);
            this.btnLoadFile.MouseEnter += new System.EventHandler(this.btn_MouseEnter);
            this.btnLoadFile.MouseLeave += new System.EventHandler(this.btn_MouseLeave);
            // 
            // btnSaveFile
            // 
            this.btnSaveFile.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(189)))), ((int)(((byte)(178)))), ((int)(((byte)(183)))));
            this.btnSaveFile.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnSaveFile.Font = new System.Drawing.Font("Frutiger Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSaveFile.Location = new System.Drawing.Point(131, 143);
            this.btnSaveFile.Name = "btnSaveFile";
            this.btnSaveFile.Size = new System.Drawing.Size(113, 24);
            this.btnSaveFile.TabIndex = 2;
            this.btnSaveFile.Text = "Save File";
            this.btnSaveFile.UseVisualStyleBackColor = false;
            this.btnSaveFile.Click += new System.EventHandler(this.btnSaveFile_Click);
            this.btnSaveFile.MouseEnter += new System.EventHandler(this.btn_MouseEnter);
            this.btnSaveFile.MouseLeave += new System.EventHandler(this.btn_MouseLeave);
            // 
            // btnPreview
            // 
            this.btnPreview.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(189)))), ((int)(((byte)(178)))), ((int)(((byte)(183)))));
            this.btnPreview.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnPreview.Font = new System.Drawing.Font("Frutiger Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnPreview.Location = new System.Drawing.Point(11, 185);
            this.btnPreview.Name = "btnPreview";
            this.btnPreview.Size = new System.Drawing.Size(113, 24);
            this.btnPreview.TabIndex = 4;
            this.btnPreview.Text = "Preview";
            this.btnPreview.UseVisualStyleBackColor = false;
            this.btnPreview.Click += new System.EventHandler(this.btnPreview_Click);
            this.btnPreview.MouseEnter += new System.EventHandler(this.btn_MouseEnter);
            this.btnPreview.MouseLeave += new System.EventHandler(this.btn_MouseLeave);
            // 
            // lblPagesInDoc
            // 
            this.lblPagesInDoc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblPagesInDoc.AutoSize = true;
            this.lblPagesInDoc.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPagesInDoc.ForeColor = System.Drawing.Color.Black;
            this.lblPagesInDoc.Location = new System.Drawing.Point(8, 275);
            this.lblPagesInDoc.Name = "lblPagesInDoc";
            this.lblPagesInDoc.Size = new System.Drawing.Size(127, 13);
            this.lblPagesInDoc.TabIndex = 0;
            this.lblPagesInDoc.Text = "Pages in document: 0";
            // 
            // lblLoadedFile
            // 
            this.lblLoadedFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblLoadedFile.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLoadedFile.ForeColor = System.Drawing.Color.Black;
            this.lblLoadedFile.Location = new System.Drawing.Point(8, 300);
            this.lblLoadedFile.Name = "lblLoadedFile";
            this.lblLoadedFile.Size = new System.Drawing.Size(350, 37);
            this.lblLoadedFile.TabIndex = 0;
            this.lblLoadedFile.Text = "Loaded Doc: None";
            // 
            // btnExecute
            // 
            this.btnExecute.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnExecute.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(189)))), ((int)(((byte)(178)))), ((int)(((byte)(183)))));
            this.btnExecute.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnExecute.Font = new System.Drawing.Font("Frutiger Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExecute.Location = new System.Drawing.Point(11, 387);
            this.btnExecute.Name = "btnExecute";
            this.btnExecute.Size = new System.Drawing.Size(113, 24);
            this.btnExecute.TabIndex = 3;
            this.btnExecute.Text = "Execute";
            this.btnExecute.UseVisualStyleBackColor = false;
            this.btnExecute.Click += new System.EventHandler(this.btnExecute_Click);
            this.btnExecute.MouseEnter += new System.EventHandler(this.btn_MouseEnter);
            this.btnExecute.MouseLeave += new System.EventHandler(this.btn_MouseLeave);
            // 
            // btnUnload
            // 
            this.btnUnload.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(189)))), ((int)(((byte)(178)))), ((int)(((byte)(183)))));
            this.btnUnload.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnUnload.Font = new System.Drawing.Font("Frutiger Linotype", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnUnload.Location = new System.Drawing.Point(12, 226);
            this.btnUnload.Name = "btnUnload";
            this.btnUnload.Size = new System.Drawing.Size(113, 24);
            this.btnUnload.TabIndex = 5;
            this.btnUnload.Text = "Unload File";
            this.toolTip1.SetToolTip(this.btnUnload, "Unstable Feature");
            this.btnUnload.UseVisualStyleBackColor = false;
            this.btnUnload.Click += new System.EventHandler(this.btnUnload_Click);
            this.btnUnload.MouseEnter += new System.EventHandler(this.btn_MouseEnter);
            this.btnUnload.MouseLeave += new System.EventHandler(this.btn_MouseLeave);
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog1";
            this.openFileDialog.Filter = "Microsoft Word Documents (*.docx/*.doc)|*.docx;*.doc|All Files (*.*)|*.*";
            this.openFileDialog.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog_FileOk);
            // 
            // saveFileDialog
            // 
            this.saveFileDialog.Filter = "\"PDF|*.pdf|Word Document|*.doc\"";
            this.saveFileDialog.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog_FileOk);
            // 
            // pdfReader
            // 
            this.pdfReader.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pdfReader.Enabled = true;
            this.pdfReader.Location = new System.Drawing.Point(401, 36);
            this.pdfReader.Name = "pdfReader";
            this.pdfReader.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("pdfReader.OcxState")));
            this.pdfReader.Size = new System.Drawing.Size(270, 374);
            this.pdfReader.TabIndex = 0;
            this.pdfReader.Visible = false;
            // 
            // pBar1
            // 
            this.pBar1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.pBar1.BackColor = System.Drawing.Color.Gainsboro;
            this.pBar1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.pBar1.Location = new System.Drawing.Point(11, 340);
            this.pBar1.Name = "pBar1";
            this.pBar1.Size = new System.Drawing.Size(347, 27);
            this.pBar1.TabIndex = 0;
            this.pBar1.TextColor = System.Drawing.Color.Black;
            this.pBar1.Value = 0F;
            // 
            // minimizeBox
            // 
            this.minimizeBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.minimizeBox.Image = ((System.Drawing.Image)(resources.GetObject("minimizeBox.Image")));
            this.minimizeBox.Location = new System.Drawing.Point(635, 0);
            this.minimizeBox.Name = "minimizeBox";
            this.minimizeBox.Size = new System.Drawing.Size(24, 24);
            this.minimizeBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.minimizeBox.TabIndex = 9;
            this.minimizeBox.TabStop = false;
            this.minimizeBox.Click += new System.EventHandler(this.minimizeBox_Click);
            this.minimizeBox.MouseEnter += new System.EventHandler(this.minimizeBox_MouseEnter);
            this.minimizeBox.MouseLeave += new System.EventHandler(this.minimizeBox_MouseLeave);
            // 
            // closeFormButtton
            // 
            this.closeFormButtton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.closeFormButtton.Image = ((System.Drawing.Image)(resources.GetObject("closeFormButtton.Image")));
            this.closeFormButtton.Location = new System.Drawing.Point(659, 0);
            this.closeFormButtton.Name = "closeFormButtton";
            this.closeFormButtton.Size = new System.Drawing.Size(24, 24);
            this.closeFormButtton.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.closeFormButtton.TabIndex = 9;
            this.closeFormButtton.TabStop = false;
            this.closeFormButtton.Click += new System.EventHandler(this.closeFormButtton_Click);
            this.closeFormButtton.MouseEnter += new System.EventHandler(this.closeFormButtton_MouseEnter);
            this.closeFormButtton.MouseLeave += new System.EventHandler(this.closeFormButtton_MouseLeave);
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(242)))), ((int)(((byte)(242)))), ((int)(((byte)(249)))));
            this.ClientSize = new System.Drawing.Size(683, 447);
            this.ControlBox = false;
            this.Controls.Add(this.pBar1);
            this.Controls.Add(this.minimizeBox);
            this.Controls.Add(this.closeFormButtton);
            this.Controls.Add(this.pdfReader);
            this.Controls.Add(this.lblLoadedFile);
            this.Controls.Add(this.lblPagesInDoc);
            this.Controls.Add(this.btnUnload);
            this.Controls.Add(this.btnExecute);
            this.Controls.Add(this.btnPreview);
            this.Controls.Add(this.btnSaveFile);
            this.Controls.Add(this.btnLoadFile);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pdfReader)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.minimizeBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.closeFormButtton)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel stslblFileLoaded;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox chkScaleSlides;
        private System.Windows.Forms.CheckBox chkDelSlideNums;
        private System.Windows.Forms.Button btnLoadFile;
        private System.Windows.Forms.Button btnSaveFile;
        private System.Windows.Forms.Button btnPreview;
        private System.Windows.Forms.Label lblPagesInDoc;
        private System.Windows.Forms.Label lblLoadedFile;
        private System.Windows.Forms.Button btnExecute;
        private System.Windows.Forms.Button btnUnload;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveAsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem previewToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
        private AxAcroPDFLib.AxAcroPDF pdfReader;
        private System.Windows.Forms.PictureBox closeFormButtton;
        private System.Windows.Forms.PictureBox minimizeBox;
        private System.Windows.Forms.ToolTip toolTip1;
        private PBar.PBar pBar1;
        private System.Windows.Forms.ToolStripMenuItem toolStripThemeItem;
        private System.Windows.Forms.ToolStripMenuItem lightThemeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem darkThemeToolStripMenuItem;
    }
}

