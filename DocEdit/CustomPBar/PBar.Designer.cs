﻿namespace CustomPBar
{
    partial class PBar
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lblValue = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lblValue
            // 
            this.lblValue.BackColor = System.Drawing.Color.Transparent;
            this.lblValue.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblValue.Font = new System.Drawing.Font("Consolas", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblValue.Location = new System.Drawing.Point(0, 0);
            this.lblValue.Name = "lblValue";
            this.lblValue.Size = new System.Drawing.Size(360, 33);
            this.lblValue.TabIndex = 0;
            this.lblValue.Text = "0%";
            this.lblValue.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // PBar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.Controls.Add(this.lblValue);
            this.Name = "PBar";
            this.Size = new System.Drawing.Size(360, 33);
            this.SizeChanged += new System.EventHandler(this.PBar_SizeChanged);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label lblValue;
    }
}
