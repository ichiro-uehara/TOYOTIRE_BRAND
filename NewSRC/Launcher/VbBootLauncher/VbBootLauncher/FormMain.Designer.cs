
namespace VbBootLauncher
{
    partial class FormMain
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.listBoxCom = new System.Windows.Forms.ListBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.MenuItemSearch = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuItemPrimitiveSearch = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuItemEditingSearch = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuItemRegist = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuItemPrimitiveRegist = new System.Windows.Forms.ToolStripMenuItem();
            this.MenuItemEditingRegist = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // listBoxCom
            // 
            this.listBoxCom.FormattingEnabled = true;
            this.listBoxCom.HorizontalScrollbar = true;
            this.listBoxCom.ItemHeight = 16;
            this.listBoxCom.Location = new System.Drawing.Point(11, 34);
            this.listBoxCom.Name = "listBoxCom";
            this.listBoxCom.Size = new System.Drawing.Size(607, 324);
            this.listBoxCom.TabIndex = 1;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.listBoxCom);
            this.groupBox1.Location = new System.Drawing.Point(12, 70);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(629, 376);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Communication";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.MenuItemSearch,
            this.MenuItemRegist});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(661, 24);
            this.menuStrip1.TabIndex = 7;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // MenuItemSearch
            // 
            this.MenuItemSearch.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.MenuItemPrimitiveSearch,
            this.MenuItemEditingSearch});
            this.MenuItemSearch.Name = "MenuItemSearch";
            this.MenuItemSearch.Size = new System.Drawing.Size(67, 20);
            this.MenuItemSearch.Text = "文字検索";
            // 
            // MenuItemPrimitiveSearch
            // 
            this.MenuItemPrimitiveSearch.Name = "MenuItemPrimitiveSearch";
            this.MenuItemPrimitiveSearch.Size = new System.Drawing.Size(180, 22);
            this.MenuItemPrimitiveSearch.Text = "刻印文字";
            this.MenuItemPrimitiveSearch.Click += new System.EventHandler(this.MenuItemPrimitiveSearch_Click);
            // 
            // MenuItemEditingSearch
            // 
            this.MenuItemEditingSearch.Name = "MenuItemEditingSearch";
            this.MenuItemEditingSearch.Size = new System.Drawing.Size(180, 22);
            this.MenuItemEditingSearch.Text = "編集文字";
            this.MenuItemEditingSearch.Click += new System.EventHandler(this.MenuItemEditingSearch_Click);
            // 
            // MenuItemRegist
            // 
            this.MenuItemRegist.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.MenuItemPrimitiveRegist,
            this.MenuItemEditingRegist});
            this.MenuItemRegist.Name = "MenuItemRegist";
            this.MenuItemRegist.Size = new System.Drawing.Size(67, 20);
            this.MenuItemRegist.Text = "文字登録";
            // 
            // MenuItemPrimitiveRegist
            // 
            this.MenuItemPrimitiveRegist.Name = "MenuItemPrimitiveRegist";
            this.MenuItemPrimitiveRegist.Size = new System.Drawing.Size(180, 22);
            this.MenuItemPrimitiveRegist.Text = "刻印文字";
            this.MenuItemPrimitiveRegist.Click += new System.EventHandler(this.MenuItemPrimitiveRegist_Click);
            // 
            // MenuItemEditingRegist
            // 
            this.MenuItemEditingRegist.Name = "MenuItemEditingRegist";
            this.MenuItemEditingRegist.Size = new System.Drawing.Size(180, 22);
            this.MenuItemEditingRegist.Text = "編集文字";
            this.MenuItemEditingRegist.Click += new System.EventHandler(this.MenuItemEditingRegist_Click);
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(661, 458);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.menuStrip1);
            this.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "FormMain";
            this.Text = "Server";
            this.groupBox1.ResumeLayout(false);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ListBox listBoxCom;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem MenuItemSearch;
        private System.Windows.Forms.ToolStripMenuItem MenuItemPrimitiveSearch;
        private System.Windows.Forms.ToolStripMenuItem MenuItemEditingSearch;
        private System.Windows.Forms.ToolStripMenuItem MenuItemRegist;
        private System.Windows.Forms.ToolStripMenuItem MenuItemPrimitiveRegist;
        private System.Windows.Forms.ToolStripMenuItem MenuItemEditingRegist;
    }
}

