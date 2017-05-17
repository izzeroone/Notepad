using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Diagnostics;
using System.Windows.Forms;

namespace MyTextBox
{
    public partial class NumberMargin : Control
    {

        /// <summary>
        /// Constructor
        /// </summary>
        public NumberMargin()
        {
            InitializeComponent();

            //Enable DoubleBuffer to avoid flicking 
            SetStyle(ControlStyles.DoubleBuffer |ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
        }

        #region Fields

        //textArea
        private TextArea parentTextArea;

        //the font of number
        private Font numberFont = new Font("Arial",10);

        //the brush of number
        private Brush numberBrush = Brushes.Black;

        //is VScrolled or not?
        //private bool isVScrolledOnTextChanged = false;

        //previous zoom factor of text area
        private float previousZoomFactor = 1;

        //the left padding
        private int widthPadding = 4; 
        
        //previous number of lines (include virtual lines) 
        //private int previousNumberOfLines = 0;

        private bool isRefreshedOnResized = false;

        //the number of real lines
        private int previousLinesLength = 0;

        private bool isNeededAutoNumbering = false;

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
        public Font NumberFont
        {
            get
            {
                return numberFont;
            }
            set
            {
                numberFont = value;
            }
        }

        [System.ComponentModel.Browsable(true)]
        [System.ComponentModel.Category("Coder")]
        public Brush NumberBrush
        {
            get
            {
                return numberBrush;
            }
            set
            {
                numberBrush = value;
            }
        }

        [System.ComponentModel.Browsable(true)]
        [System.ComponentModel.Category("Coder")]
        public int WidthPadding
        {
            get { return widthPadding; }
            set { widthPadding = value; }
        }
        public bool IsNeededAutoNumbering { get { return isNeededAutoNumbering; } set { isNeededAutoNumbering = value; } }

        #endregion

        public void AutoNumbering(bool autoNumberingEnabled)
        {
            if (autoNumberingEnabled)
            {
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


            isNeededAutoNumbering = autoNumberingEnabled;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            //Win32.LockWindowUpdate(Handle);

            if (parentTextArea != null && isNeededAutoNumbering)
            {

                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighSpeed;

                //get the zoom factor of the text Area 
                //this is needed when we zoom in or zoom out the text area
                float zoom = parentTextArea.ZoomFactor;

                //Font to draw number
                //Since we just get the familly name of numberFont, we have to compute the actually size of the font to draw numbers
                Font fontToDrawNumber = new Font(numberFont.FontFamily.Name, parentTextArea.Font.Size * zoom);

                //set the width of this 
                this.Width = TextRenderer.MeasureText((parentTextArea.Lines.Length + 1).ToString(), fontToDrawNumber).Width + widthPadding;

                // Get the first and last line currently visible in the screen
                //int firstVisibleCharIndex = parentTextArea.GetCharIndexFromPosition(new Point(0, 0));
                //int lastVisibleCharIndex = parentTextArea.GetCharIndexFromPosition(new Point(parentTextArea.Width, parentTextArea.Height));

                // Get the first and last line currently visible in the screen
                //we just need to draw the line number of the lines visibly in the screen
                //this makes things more slightly
                //int firstVisibleLine = parentTextArea.GetLineFromCharIndex(firstVisibleCharIndex);
                //int lastVisibleLine = parentTextArea.GetLineFromCharIndex(lastVisibleCharIndex);
                int firstVisibleLine = parentTextArea.FirstVisibleLine;
                int lastVisibleLine = parentTextArea.LastVisibleLine;

                ////get the actually first line Number currently visible within Screen
                ////the firstLine and lastLine in the WordWrap mode are just the virtual line within the screen
                ////so we have to do this to get the actually number of lines from line 0 to the first virtual line currently visible within Screen
                //int lineNumber = parentTextArea.Text.Substring(0, parentTextArea.GetCharIndexFromPosition(new Point(0, 0))).Split('\n').Length;

                ////if the first line is not a real line but just a virtual line, it means the real line contains this virtual line is not visible in the screen
                //if (firstVisibleLine != 0 && parentTextArea.Text[parentTextArea.GetFirstCharIndexFromLine(firstVisibleLine) - 1] != '\n')
                //{
                //    lineNumber++;
                //}

                //save the position of the current line
                //note that this is the location within the textArea
                Point rawPosition;

                //save to position to draw number
                //this is the actually position to draw number within this control
                Point positionToDrawNumber;

                //loop through all the lines currently visible in the screen
                for (int lineIndex = firstVisibleLine; lineIndex <= lastVisibleLine; lineIndex++)
                {

                    ////check to see if this line is the real line or virual line 
                    //if (lineIndex != 0)
                    //{
                    //    //if this isn't the real line, just continue
                    //    if (parentTextArea.Text[parentTextArea.GetFirstCharIndexFromLine(lineIndex) - 1] != '\n')
                    //    {
                    //        continue;
                    //    }
                    //}

                    //get the location of current line
                    //note that this is the location within the textArea
                    rawPosition = parentTextArea.GetPositionFromCharIndex(parentTextArea.GetFirstCharIndexFromLine(lineIndex));

                    //get the position to draw current line number
                    positionToDrawNumber = new Point(widthPadding / 2, rawPosition.Y + yOffset);

                    //draw current line number
                    e.Graphics.DrawString((lineIndex + 1).ToString(), fontToDrawNumber, numberBrush, positionToDrawNumber);
                    //lineNumber++;
                }

                //Draw one more number//
                //get the location of current line
                //note that this is the location within the textArea
                rawPosition = parentTextArea.GetPositionFromCharIndex(parentTextArea.GetFirstCharIndexFromLine(lastVisibleLine));
                //get the position to draw current line number
                positionToDrawNumber = new Point(widthPadding / 2, rawPosition.Y + yOffset + fontToDrawNumber.Height);
                //draw line number
                e.Graphics.DrawString((lastVisibleLine + 2).ToString(), fontToDrawNumber, numberBrush, positionToDrawNumber);

                //Dispose
                fontToDrawNumber.Dispose();
            }
            
            //Win32.LockWindowUpdate((IntPtr)0);
        }

        public void CalculateOnTextAreaTextChanged()
        {


            //calculate current number of lines
            //int currentNumberOfLines = parentTextArea.GetLineFromCharIndex(parentTextArea.Text.Length - 1);

            ////if we are typing at the same line 
            ////This "if" below seems like make no sense but this is actually the best way I can come up with to check if we need to redraw the numbers or not
            //if (previousNumberOfLines == currentNumberOfLines && previousLinesLength == parentTextArea.Lines.Length
            //    && textArea.SelectionStart == textArea.BasePositionToCheckHighLight)
            //{
            //    previousNumberOfLines = currentNumberOfLines;
            //    previousLinesLength = parentTextArea.Lines.Length;
            //    return;
            //}

            //if (isVScrolledOnTextChanged)
            //{
            //    isVScrolledOnTextChanged = false;
            //    return;
            //}

            if (previousLinesLength == parentTextArea.Lines.Length)
            {
                return;
            }

            ////reset previous number of lines
            //previousNumberOfLines = currentNumberOfLines;
            previousLinesLength = parentTextArea.Lines.Length;

            this.Refresh();
        }

        public void CalculateOnTextAreaVScroll()
        {
            //this bunch of things below just to avoid refreshing many time 
            //since in once time, there can be more than just this event was raised, so we need to do some works to prevent this problem
            //actually you can just remove all the things except this.Refresh, but as I say, you will risk the performance if refresh so many times 
            //if (parentTextArea == null) return;
            if (isRefreshedOnResized)
            {
                isRefreshedOnResized = false;
                return;
            }
            //TextArea currentTextArea = (TextArea)parentTextArea;
            //if (currentTextArea.SelectionStart - currentTextArea.BasePositionToCheckHighLight >= 1)
            //{
            //    isVScrolledOnTextChanged = false;
            //    return;
            //}
            //else
            //{
            //    isVScrolledOnTextChanged = true;
            this.Refresh();
            //}
        }


        public void CalculateOnTextAreaContentsResized()
        {
            //since in once time, there can be more than just this event was raised, so we need to do some works to prevent this problem
            //actually you can just remove all the things except this.Refresh, but as I say, you will risk the performance if refresh so many times 
            //if (parentTextArea == null) return;
            if (previousZoomFactor == parentTextArea.ZoomFactor) return;
            previousZoomFactor = parentTextArea.ZoomFactor;
            this.Refresh();
        }



        public void CalculateOnTextAreaResize()
        {
            //this bunch of things below just to avoid refreshing many time 
            //since in once time, there can be more than just this event was raised, so we need to do some works to prevent this problem
            //actually you can just remove all the things except this.Refresh, but as I say, you will risk the performance if refresh so many times 
            //if (parentTextArea == null) return;
            //if (parentTextArea.WordWrap == false) return;
            //int currentNumberOfLines = parentTextArea.GetLineFromCharIndex(parentTextArea.Text.Length - 1);
            //if (previousNumberOfLines == currentNumberOfLines) return;
            isRefreshedOnResized = true;
            //previousNumberOfLines = currentNumberOfLines;    
            this.Refresh();
        }

    }
}
