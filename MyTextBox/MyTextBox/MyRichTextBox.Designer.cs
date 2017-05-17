namespace MyTextBox
{
    partial class MyRichTextBox
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
            this.textArea = new MyTextBox.TextArea();
            this.bookMarkMargin = new MyTextBox.BookMarkMargin();
            this.numberMargin = new MyTextBox.NumberMargin();
            this.documentMap = new MyTextBox.DocumentMap();
            this.SuspendLayout();
            // 
            // textArea
            // 
            this.textArea.AcceptsTab = true;
            this.textArea.BasePositionToCheckHighLight = 0;
            this.textArea.BookMarkMargin = this.bookMarkMargin;
            this.textArea.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textArea.CommentBlocksColor = System.Drawing.Color.DarkGreen;
            this.textArea.CommentLinesColor = System.Drawing.Color.Green;
            this.textArea.ControlToFocus = this.numberMargin;
            this.textArea.DefaultColor = System.Drawing.Color.Black;
            this.textArea.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textArea.DocumentMap = this.documentMap;
            this.textArea.FirstVisibleCharIndex = 0;
            this.textArea.FirstVisibleLine = 0;
            this.textArea.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textArea.KeyWordsColor = System.Drawing.Color.Blue;
            this.textArea.Language = "NormalText";
            this.textArea.LastVisibleCharIndex = 0;
            this.textArea.LastVisibleLine = 0;
            this.textArea.Location = new System.Drawing.Point(51, 0);
            this.textArea.Name = "textArea";
            this.textArea.NumberMargin = this.numberMargin;
            this.textArea.PreprocessorsColor = System.Drawing.Color.Gray;
            this.textArea.Size = new System.Drawing.Size(553, 511);
            this.textArea.StringsColor = System.Drawing.Color.Brown;
            this.textArea.TabIndex = 4;
            this.textArea.TabSize = 4;
            this.textArea.StopRecordingUndo();
            this.textArea.Text = "";
            this.textArea.ContinueRecordingUndo();
            this.textArea.WordWrap = false;
            // 
            // bookMarkMargin
            // 
            this.bookMarkMargin.BackColor = System.Drawing.SystemColors.ButtonShadow;
            this.bookMarkMargin.BookMarkFont = new System.Drawing.Font("Webdings", 10F);
            this.bookMarkMargin.Dock = System.Windows.Forms.DockStyle.Left;
            this.bookMarkMargin.IsNeededAutoBookMark = false;
            this.bookMarkMargin.Location = new System.Drawing.Point(25, 0);
            this.bookMarkMargin.Name = "bookMarkMargin";
            this.bookMarkMargin.ParentTextArea = this.textArea;
            this.bookMarkMargin.Size = new System.Drawing.Size(26, 511);
            this.bookMarkMargin.TabIndex = 6;
            this.bookMarkMargin.Text = "bookMarkMargin1";
            this.bookMarkMargin.WidthPadding = 2;
            // 
            // numberMargin
            // 
            this.numberMargin.BackColor = System.Drawing.SystemColors.ControlDark;
            this.numberMargin.Dock = System.Windows.Forms.DockStyle.Left;
            this.numberMargin.IsNeededAutoNumbering = false;
            this.numberMargin.Location = new System.Drawing.Point(0, 0);
            this.numberMargin.Name = "numberMargin";
            this.numberMargin.NumberFont = new System.Drawing.Font("Arial", 10F);
            this.numberMargin.ParentTextArea = this.textArea;
            this.numberMargin.Size = new System.Drawing.Size(25, 511);
            this.numberMargin.TabIndex = 5;
            this.numberMargin.Text = "numberMargin1";
            this.numberMargin.WidthPadding = 4;
            // 
            // documentMap
            // 
            this.documentMap.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.documentMap.Dock = System.Windows.Forms.DockStyle.Right;
            this.documentMap.IsNeededAutoDocumentMap = false;
            this.documentMap.Location = new System.Drawing.Point(604, 0);
            this.documentMap.Name = "documentMap";
            this.documentMap.ParentTextArea = this.textArea;
            this.documentMap.Size = new System.Drawing.Size(165, 511);
            this.documentMap.SizeOfText = 2F;
            this.documentMap.TabIndex = 7;
            this.documentMap.Text = "documentMap1";
            // 
            // MyRichTextBox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.textArea);
            this.Controls.Add(this.documentMap);
            this.Controls.Add(this.bookMarkMargin);
            this.Controls.Add(this.numberMargin);
            this.Name = "MyRichTextBox";
            this.Size = new System.Drawing.Size(769, 511);
            this.ResumeLayout(false);

        }

        #endregion
        private TextArea textArea;
        private NumberMargin numberMargin;
        private BookMarkMargin bookMarkMargin;
        private DocumentMap documentMap;

        public NumberMargin NumberMargin { get { return numberMargin; } set { numberMargin = value; } }
        public TextArea TextArea { get { return textArea; } set { textArea = value; } }
        public BookMarkMargin BookMarkMargin { get { return bookMarkMargin; } set { bookMarkMargin = value; } }
        public DocumentMap DocumentMap { get { return documentMap; } set { documentMap = value; } }
    }
}
