using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;

namespace MyTextBox
{
    public partial class DocumentMap : Control
    {
        public DocumentMap()
        {
            InitializeComponent();
            //Enable DoubleBuffer to avoid flicking 
            SetStyle(ControlStyles.DoubleBuffer| ControlStyles.UserPaint|ControlStyles.AllPaintingInWmPaint, true);
        }

        #region Fields
        //the text area to look after 
        private TextArea parentTextArea = null;
        //the brush to fill the rectangle
        private Brush rectangleBrush = new SolidBrush(Color.FromArgb(100, 255, 0, 0));
        //the size of text in this control
        private float sizeOfText = 2;

        //Enable auto document map or not
        private bool isNeededAutoDocumentMap = false; 

        #endregion


        #region Properties

        [System.ComponentModel.Browsable(true)]
        [System.ComponentModel.DefaultValue(null)]
        [System.ComponentModel.Category("Coder")]
        public TextArea ParentTextArea
        {
            get
            {
                return parentTextArea;
            }
            set
            {
                parentTextArea = value;
            }
        }

        [System.ComponentModel.Browsable(true)]
        [System.ComponentModel.Category("Coder")]
        public Brush RectangleBrush { get { return rectangleBrush; } set { rectangleBrush = value; } }

        [System.ComponentModel.Browsable(true)]
        [System.ComponentModel.Category("Coder")]
        public float SizeOfText { get { return sizeOfText; } set { sizeOfText = value; } }
        public bool IsNeededAutoDocumentMap { get { return isNeededAutoDocumentMap; } set { isNeededAutoDocumentMap = value; } }

        #endregion

        public void AutoDocumentMap(bool autoDocumentMapEnabled)
        {
            if(autoDocumentMapEnabled)
            {
                this.Width = 150;
            }
            else
            {
                this.Width = 0;
            }
            isNeededAutoDocumentMap = autoDocumentMapEnabled;
        }


        protected override void OnPaint(PaintEventArgs e)
        {

            base.OnPaint(e);

            if (parentTextArea == null || isNeededAutoDocumentMap == false) return;

            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighSpeed;

            //Get the font used to draw text
            Font fontToDrawText = new Font(parentTextArea.Font.FontFamily.Name, sizeOfText, FontStyle.Regular);

            //Calculate lineOffset between this control and parentTextArea
            //I do know we have to calculate the linesOffset for some reasons that you will see below
            //But in fact I don't know why we do a calculation like this, but somehow it works quite well I have to say  
            int linesOffset = parentTextArea.Height/fontToDrawText.Height - (int)(parentTextArea.Height/(parentTextArea.Font.Height * parentTextArea.ZoomFactor));

            if (linesOffset < 0)
            {
                fontToDrawText = new Font(parentTextArea.Font.FontFamily.Name, 1, FontStyle.Regular);
                linesOffset = parentTextArea.Height / fontToDrawText.Height - (int)(parentTextArea.Height / (parentTextArea.Font.Height * parentTextArea.ZoomFactor));
            }

            //get the first char index currently visible in the screen
            //int firstVisibleCharIndex = parentTextArea.GetCharIndexFromPosition(new Point(0, 0));
            int firstVisibleCharIndex = parentTextArea.FirstVisibleCharIndex;

            //get the last char index currently visible in the screen
            //int lastVisibleCharIndex = parentTextArea.GetCharIndexFromPosition(new Point(parentTextArea.Width, parentTextArea.Height));
            int lastVisibleCharIndex = parentTextArea.LastVisibleCharIndex;

            //get the first line currently visible in the screen 
            //int firstVisibleLine = parentTextArea.GetLineFromCharIndex(firstVisibleCharIndex);
            int firstVisibleLine = parentTextArea.FirstVisibleLine;

            //Calculate base char index to draw text from //
            //if firstVisibleLine < linesOffset, just get the first character of the text of parentTextArea
            if (firstVisibleLine - linesOffset < 0 || linesOffset<0) linesOffset = firstVisibleLine;
            int baseCharIndex = parentTextArea.GetFirstCharIndexFromLine(firstVisibleLine - linesOffset);
            
            //DRAW TEXT//
            string TextToDraw = parentTextArea.Text.Substring(baseCharIndex, parentTextArea.TextLength - baseCharIndex);
            //Draw the text of the parent text area into this control
            TextRenderer.DrawText(e.Graphics, TextToDraw, fontToDrawText, new Point(0, 0), Color.Black);

            //DRAW RECTANGLE//
            //calculate y Offset of the rectangle
            int yOffset = TextRenderer.MeasureText(e.Graphics, parentTextArea.Text.Substring(baseCharIndex, firstVisibleCharIndex - baseCharIndex), fontToDrawText).Height;
            //calculate the height of the rectangle 
            int rectangleHeight = TextRenderer.MeasureText(e.Graphics, parentTextArea.Text.Substring(firstVisibleCharIndex, lastVisibleCharIndex - firstVisibleCharIndex),fontToDrawText).Height;      
            //Draw rectangle
            e.Graphics.FillRectangle(rectangleBrush, 0f, yOffset, this.Width, rectangleHeight);
            
            //Dispose for sure
            fontToDrawText.Dispose();
        }

    }

}