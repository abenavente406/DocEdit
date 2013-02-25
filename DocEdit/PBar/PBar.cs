using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Drawing2D;

namespace PBar
{
    public partial class PBar : UserControl
    {
        protected float percent = 0.0f;
        protected Color textColor = Color.Black;

        public PBar()
        {
            InitializeComponent();
            lblValue.ForeColor = textColor;
            this.ForeColor = SystemColors.Highlight;
        }

        public float Value
        {
            get { return percent; }
            set
            {
                if (value < 0) value = 0;
                if (value > 100) value = 100;
                percent = value;
                lblValue.Text = value.ToString() + "%";
                this.Invalidate();
            }
        }

        public Color TextColor
        {
            get { return textColor; }
            set 
            { 
                textColor = value;
                lblValue.ForeColor = textColor;
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Brush b = new SolidBrush(this.ForeColor); // Create a brush that will draw the background of the Pbar
            // Create a linear gradient that will be drawn over the background. FromArgb means you can use the Alpha value wich is the transparency
            LinearGradientBrush lb = new LinearGradientBrush(new Rectangle(0, 0, this.Width, this.Height),
                Color.FromArgb(120, Color.White), Color.FromArgb(50, Color.White), LinearGradientMode.ForwardDiagonal);
            // Calculate how much has the Pbar to be filled for "x" %
            int width = (int)((percent / 100) * this.Width);
            e.Graphics.FillRectangle(b, 0, 0, width, this.Height);
            e.Graphics.FillRectangle(lb, 0, 0, width, this.Height);
            b.Dispose(); lb.Dispose();

        }

        private void PBar_SizeChanged(object sender, EventArgs e)
        {
            lblValue.Location = new Point(this.Width / 2 - 21 / 2 - 4, this.Height / 2 - 15 / 2);
        }
    }
}
