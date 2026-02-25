namespace VisualizadorCR.Addon.Forms
{
    partial class CrystalReportViewerForm
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
            this._viewer = new CrystalDecisions.Windows.Forms.CrystalReportViewer();
            this.SuspendLayout();
            // 
            // _viewer
            // 
            this._viewer.ActiveViewIndex = -1;
            this._viewer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this._viewer.Cursor = System.Windows.Forms.Cursors.Default;
            this._viewer.Dock = System.Windows.Forms.DockStyle.Fill;
            this._viewer.Location = new System.Drawing.Point(0, 0);
            this._viewer.Name = "_viewer";
            this._viewer.Size = new System.Drawing.Size(800, 450);
            this._viewer.TabIndex = 0;
            // 
            // CrystalReportViewerForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this._viewer);
            this.Name = "CrystalReportViewerForm";
            this.Text = "CrystalReportViewerForm";
            this.ResumeLayout(false);

        }

        #endregion

        private CrystalDecisions.Windows.Forms.CrystalReportViewer _viewer;
    }
}