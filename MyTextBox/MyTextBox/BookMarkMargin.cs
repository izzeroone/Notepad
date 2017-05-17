using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyTextBox
{
    public partial class BookMarkMargin : Control
    {
        public BookMarkMargin()
        {
            InitializeComponent();

            //Enable DoubleBuffer to avoid flicking 
            SetStyle(ControlStyles.DoubleBuffer | ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
        }

        #region Fields

        //textArea
        private TextArea parentTextArea;

        private List<int> bookMarkedLines = null;

        //the font of BookMark
        private Font bookMarkFont = new Font("Webdings", 10);

        //the brush of number
        private Brush bookMarkBrush = Brushes.Red;

        //the left padding
        private int widthPadding = 2;

        private float previousZoomFactor = 1;

        //private int previousNumberOfLines = 0;

        private bool isRefreshedOnResized = false;

        private int previousLinesLength = 0;

        private bool isNeededAutoBookMark = false;

        private int yOffset = 0;


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
        public Font BookMarkFont { get { return bookMarkFont; }   set { bookMarkFont = value; } }

        [System.ComponentModel.Browsable(true)]
        [System.ComponentModel.Category("Coder")]
        public Brush BookMarkBrush { get { return bookMarkBrush; } set { bookMarkBrush = value; } }

        [System.ComponentModel.Browsable(true)]
        [System.ComponentModel.Category("Coder")]
        public int WidthPadding {  get { return widthPadding; } set { widthPadding = value; } }
        public bool IsNeededAutoBookMark { get { return isNeededAutoBookMark; } set { isNeededAutoBookMark = value; } }

        //public int FirstVisibleLine { get => firstVisibleLine; set => firstVisibleLine = value; }
        //public int LastVisibleLine { get => lastVisibleLine; set => lastVisibleLine = value; }


        #endregion


        public void AutoBookMark(bool autoBookMarkEnabled)
        {

            if (autoBookMarkEnabled)
            {
                bookMarkedLines = new List<int>();

                if (parentTextArea != null)
                {

                    // Compute Y offset //
                    Point inTextArea = parentTextArea.PointToScreen(new Point(0, 0));
                    Point inMe = this.PointToScreen(new Point(0, 0));
                    yOffset = inTextArea.Y - inMe.Y;

                }
            }
            else
            {
                this.Width = 0;
            }

            isNeededAutoBookMark = autoBookMarkEnabled;

        }

        protected override void OnClick(EventArgs e)
        {

            base.OnClick(e);

            if (parentTextArea == null || bookMarkedLines == null || isNeededAutoBookMark==false) return;

            //get the real line number nearest the MousePosition
            int charIndex = parentTextArea.GetCharIndexFromPosition(parentTextArea.PointToClient(MousePosition));
            //int line = parentTextArea.Text.Substring(0, charIndex).Split('\n').Length - 1;
            int line = parentTextArea.GetLineFromCharIndex(charIndex);
            
            if (bookMarkedLines.Contains(line) == false)
            {
                bookMarkedLines.Add(line);
                this.Refresh();
            }
            else
            {
                bookMarkedLines.Remove(line);
                this.Refresh();
            }

        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            if (parentTextArea != null && bookMarkedLines != null && isNeededAutoBookMark)
            {
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighSpeed;

                //the zoom factor of parent TextArea
                float zoom = parentTextArea.ZoomFactor;

                //font to Draw Bookmark
                Font fontToDrawBookMark = new Font(bookMarkFont.FontFamily.Name, parentTextArea.Font.Size * zoom);


                //Compute the width of this control
                this.Width = TextRenderer.MeasureText("n", fontToDrawBookMark).Width + widthPadding;

                if (bookMarkedLines.Count > 0)
                {
                    //a temp list to easily store bookmarked lines
                    List<int> cloneBookMarkedLines = new List<int>(bookMarkedLines);

                    ////Get the first and last char index currently visible in the screen
                    //int firstVisibleCharIndex = parentTextArea.GetCharIndexFromPosition(new Point(0, 0));
                    //int lastVisibleCharIndex = parentTextArea.GetCharIndexFromPosition(new Point(parentTextArea.Width, parentTextArea.Height));

                    //Get the first and last line currently visible in the screen
                    //int firstVisibleLine = parentTextArea.GetLineFromCharIndex(firstVisibleCharIndex);
                    //int lastVisibleLine = parentTextArea.GetLineFromCharIndex(lastVisibleCharIndex);

                    int firstVisibleLine = parentTextArea.FirstVisibleLine;
                    int lastVisibleLine = parentTextArea.LastVisibleLine;


                    for (int lineIndex = firstVisibleLine; lineIndex <= lastVisibleLine; lineIndex++)
                    {
                        ////check to see if this line is the real line or virual line 
                        //if (lineIndex != 0)
                        //{
                        //    //if this isn't the real line, just return
                        //    if (parentTextArea.Text[parentTextArea.GetFirstCharIndexFromLine(lineIndex) - 1] != '\n')
                        //    {
                        //        continue;
                        //    }
                        //}

                        ////get the real line number nearest the lineIndex
                        //int charIndex = parentTextArea.GetFirstCharIndexFromLine(lineIndex);
                        //int realLine = parentTextArea.Text.Substring(0, charIndex).Split('\n').Length - 1;
                        
                        //if there is this line in tempList
                        if (cloneBookMarkedLines.Contains(lineIndex))
                        {
                            //Draw bookmark symbol
                            Point rawPosition = parentTextArea.GetPositionFromCharIndex(parentTextArea.GetFirstCharIndexFromLine(lineIndex));
                            e.Graphics.DrawString("n", fontToDrawBookMark, bookMarkBrush, new Point((int)widthPadding / 2, rawPosition.Y + yOffset - 1));

                            //Remove this line from tempList to be able to check the next bookmarked line
                            cloneBookMarkedLines.Remove(lineIndex);
                        }

                        //if cloneBookMarhedLines has nothing left
                        if (cloneBookMarkedLines.Count == 0) break;
                    }

                    //clear for sure
                    cloneBookMarkedLines.Clear();
                }

                //dispose
                fontToDrawBookMark.Dispose();

            }

        }
      
        public void CalculateOnTextAreaTextChanged()
        {
            if (parentTextArea == null || bookMarkedLines == null) return;

            //Calculate the bookmark position
            if (previousLinesLength != parentTextArea.Lines.Length)
            {
                //get the line number nearest the SelectionStart
                int selectedLine = parentTextArea.GetLineFromCharIndex(parentTextArea.SelectionStart);
                
                //get the start line position to calculate from
                int startLineToCalculate = selectedLine - (parentTextArea.Lines.Length - previousLinesLength);

                //get lines are needed to recalculate bookmark position
                List<int> needToReCaculateLines = bookMarkedLines.FindAll(i => i >= startLineToCalculate);

                //remove all the line >= startLineToCalculate in list of bookmarked lines
                bookMarkedLines.RemoveAll(i => i >= startLineToCalculate);

                //recalculate bookmark position
                for (int i = 0; i < needToReCaculateLines.Count; i++)
                {
                    //if startLineToCalculate has a bookmark
                    if (needToReCaculateLines[i] == startLineToCalculate)
                    {
                        //if we cut this line at somewhere in the middle, we shouldn't recalculate the bookmark of this line
                        if (parentTextArea.Lines.Length - previousLinesLength > 0 && 
                            parentTextArea.SelectionStart >= 2 && parentTextArea.Text[parentTextArea.SelectionStart - 2] != '\n')
                        {
                            continue;
                        }
                    }

                    //calculate the new bookmark position
                    needToReCaculateLines[i] += parentTextArea.Lines.Length - previousLinesLength;
                }

                //Add to bookmarked list again
                bookMarkedLines.AddRange(needToReCaculateLines);

                //reset previous lines length
                previousLinesLength = parentTextArea.Lines.Length;

                this.Refresh();
            }

        }

        public void CalculateOnTextAreaVScroll()
        {
            //since in once time, there can be more than just this event was raised, so we need to do some works to prevent this problem
            //actually you can just remove all the things except this.Refresh, but as I say, you will risk the performance if refresh so many times 
            if (isRefreshedOnResized)
            {
                isRefreshedOnResized = false;
                return;
            }

            this.Refresh();
        }

        public void CalculateOnTextAreaContentsResized()
        {
            //since in once time, there can be more than just this event was raised, so we need to do some works to prevent this problem
            //actually you can just remove all the things except this.Refresh, but as I say, you will risk the performance if refresh so many times 
            if (previousZoomFactor == parentTextArea.ZoomFactor) return;
            previousZoomFactor = parentTextArea.ZoomFactor;
            this.Refresh();
        }

        public void CalculateOnTextAreaResize()
        {
            //this bunch of things below just to avoid refreshing many time 
            //since in once time, there can be more than just this event was raised, so we need to do some works to prevent this problem
            //actually you can just remove all the things except this.Refresh, but as I say, you will risk the performance if refresh so many times 
            //if (parentTextArea.WordWrap == false) return;
            //int currentNumberOfLines = parentTextArea.GetLineFromCharIndex(parentTextArea.Text.Length - 1);
            //if (previousNumberOfLines == currentNumberOfLines) return;
            isRefreshedOnResized = true;
            //previousNumberOfLines = currentNumberOfLines;
            this.Refresh();
        }

    }

}

