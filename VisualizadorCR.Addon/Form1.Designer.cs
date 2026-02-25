
namespace VisualizadorCR.Addon
{
    partial class Form1
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this._viewer2 = new CrystalDecisions.Windows.Forms.CrystalReportViewer();
            this.SuspendLayout();
            // 
            // _viewer2
            // 
            this._viewer2.ActiveViewIndex = -1;
            this._viewer2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this._viewer2.Cursor = System.Windows.Forms.Cursors.Default;
            this._viewer2.Dock = System.Windows.Forms.DockStyle.Fill;
            this._viewer2.Location = new System.Drawing.Point(0, 0);
            this._viewer2.Name = "_viewer2";
            this._viewer2.Size = new System.Drawing.Size(800, 450);
            this._viewer2.TabIndex = 0;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this._viewer2);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private CrystalDecisions.Windows.Forms.CrystalReportViewer _viewer2;
    }
}

