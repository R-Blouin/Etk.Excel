using Etk.Demo.Shops.UI.Common.ViewModels;

namespace Etk.Demo.Shops.UI.Excel.Panels
{
    partial class ShopsPanel
    {
        /// <summary> 
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Nettoyage des ressources utilisées.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur de composants

        /// <summary> 
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas 
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            this.elementHost1 = new System.Windows.Forms.Integration.ElementHost();
            this.shopsControl1 = new Etk.Demo.Shops.UI.Common.Controls.ShopsControl();
            this.SuspendLayout();
            // 
            // elementHost1
            // 
            this.elementHost1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.elementHost1.Location = new System.Drawing.Point(0, 0);
            this.elementHost1.Name = "elementHost1";
            this.elementHost1.Padding = new System.Windows.Forms.Padding(3);
            this.elementHost1.Size = new System.Drawing.Size(491, 299);
            this.elementHost1.TabIndex = 2;
            this.elementHost1.Text = "elementHost";
            this.elementHost1.Child = this.shopsControl1;
            // 
            // ShopsPanel
            // 
            this.AutoSize = true;
            this.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Controls.Add(this.elementHost1);
            this.Name = "ShopsPanel";
            this.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Location = new System.Drawing.Point(0, 0);
            this.ResumeLayout(false);

        }
        #endregion

        private System.Windows.Forms.Integration.ElementHost elementHost1;
        private Common.Controls.ShopsControl shopsControl1;

        public void SetViewModel(ShopsViewModel viewModel)
        {
            shopsControl1.DataContext = viewModel;
        }
    }
}
