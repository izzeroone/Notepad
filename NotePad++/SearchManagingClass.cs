using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/// <summary>
/// all the search thing 
/// this code should be rewritten or improved because it uses so many loop things again and again
/// also it should be included highlighting all the similar text  
/// </summary>
namespace NotePad__
{
    class SearchManagingClass
    {
       
        ////the position to search from
        //static int positionToSearchFrom;

        ////a list to hold all the similar text position 
        //static List<int> listOfStartTargets = new List<int>();

        //public static void Setup(Scintilla textArea)
        //{
        //    //set the first position to search from
        //    textArea.TargetStart = 0;

        //    //set the position to search from
        //    positionToSearchFrom = 0;

        //    //clear list for sure
        //    listOfStartTargets.Clear();

        //    //set the last position to search to
        //    textArea.TargetEnd = textArea.TextLength;

        //    //set the search Flag (just set this for sure)
        //    textArea.SearchFlags = SearchFlags.None;
        //}

        //public static void HighlightText(string text, Scintilla textArea)
        //{
        //    textArea.TargetStart = positionToSearchFrom;

        //    //if we find out the text we are searching 
        //    if (textArea.SearchInTarget(text) != -1)
        //    {
        //        //Mark the search results
        //        //note that when we use SearchInTarget, 
        //        //the TargetStart and TargetEnd is automatically set to the position of the text we are finding 
        //        textArea.SetSelection(textArea.TargetStart, textArea.TargetEnd);

        //        //set target end again
        //        textArea.TargetEnd = textArea.TextLength;

        //        //scroll to the current caret if it's not already visible in the current scene
        //        textArea.ScrollCaret();         
        //    }
          
        //}

        ///// <summary>
        ///// search next
        ///// </summary>
        ///// <param name="textToSearch"></param>
        ///// <param name="textArea"></param>
        //public static void SearchNext(string textToSearch, Scintilla textArea)
        //{
        //    if (SearchAllTextArea(textToSearch, textArea) == true)
        //    {
        //        #region Don't care this thing
        //        //I just put this thing here because this code below might be needed sometimes

        //        ////we need to search the text area again to find out the end position of previous search  
        //        ////and set the positionToSearchFrom by textArea.TargetEnd because ..
        //        ////we don't want to start search from the previous text we have already found out
        //        //if (textArea.SearchInTarget(textToSearch) != -1)
        //        //{
        //        //    //set position To Search From
        //        //    positionToSearchFrom = textArea.TargetEnd;

        //        //    //set targetEnd = textLength to search all the remaining text area
        //        //    textArea.TargetEnd = textArea.TextLength;

        //        //    //if HighlightText return true, it means there is still similar text in the remaining text area
        //        //    if (HighlightText(textToSearch, textArea) == true)
        //        //    {
        //        //        //set this to search from the next text we found out, not the previous found text
        //        //        positionToSearchFrom = textArea.TargetStart;

        //        //        //push in the list
        //        //        stackOfStartTargets.Push(positionToSearchFrom);
        //        //    }
        //        //    //if we come to this part, it means we are at the end of of text area 
        //        //    //and so we should come to the very top of text area to search text again
        //        //    else
        //        //    {
        //        //        //set the position to search from 
        //        //        positionToSearchFrom = 0;
        //        //        textArea.TargetStart = 0;

        //        //        //push in the list
        //        //        stackOfStartTargets.Push(positionToSearchFrom);

        //        //        //should do this to immediately highlight the text at the top
        //        //        //if not, we have to click NextSearch again and it's quite painful
        //        //        //that's why let's do highlight again!
        //        //        HighlightText(textToSearch, textArea);
        //        //    }

        //        #endregion

        //        //set position To Search From
        //        positionToSearchFrom = GetRightClosestPosition();

        //        //highlight
        //        HighlightText(textToSearch, textArea);

        //    }
        //}

        ///// <summary>
        ///// search previous
        ///// </summary>
        ///// <param name="textToSearch"></param>
        ///// <param name="textArea"></param>
        //public static void SearchPrevious(string textToSearch, Scintilla textArea)
        //{
        //    if (SearchAllTextArea(textToSearch, textArea) == true)
        //    {
        //        //set position To Search From
        //        positionToSearchFrom = GetLeftClosestPosition();

        //        //highlight
        //        HighlightText(textToSearch, textArea);
        //    }
        //}

        ///// <summary>
        ///// search similar text in the text area
        ///// </summary>
        ///// <param name="textToSearch"></param>
        ///// <param name="textArea"></param>
        ///// <returns></returns>
        //private static bool SearchAllTextArea(string textToSearch, Scintilla textArea)
        //{
        //    //clear
        //    listOfStartTargets.Clear();

        //    //set target start = 0 to search all the text area
        //    textArea.TargetStart = 0;

        //    //loop through all the text area
        //    while (textArea.SearchInTarget(textToSearch) != -1)
        //    {
        //        //add targetStart to queue
        //        listOfStartTargets.Add(textArea.TargetStart);

        //        //set target start and targetEnd  to search all the remaining text area
        //        textArea.TargetStart = textArea.TargetEnd;
        //        textArea.TargetEnd = textArea.TextLength;
        //    }

        //    //if there is similar text 
        //    if (listOfStartTargets.Count > 0) return true;

        //    //if not
        //    return false;
        //}

        ///// <summary>
        /////  Get left closest position related to the positionToSearchFrom
        /////  return -1 if there is nothing in the list
        ///// </summary>
        ///// <returns></returns>
        //private static int GetLeftClosestPosition()
        //{
        //    if (listOfStartTargets.Count > 0)
        //    {
        //        for (int i = 0; i < listOfStartTargets.Count; i++)
        //        {
        //            if (listOfStartTargets[i] >= positionToSearchFrom)
        //            {
        //                if (i > 0)
        //                {
        //                    return listOfStartTargets[i - 1];
        //                }
        //                else //if i=0, it means we are at the very top of text area
        //                {
        //                    //that's why we should come to the very end of text area to search again
        //                    return listOfStartTargets[listOfStartTargets.Count - 1];
        //                }
        //            }
        //        }
        //    }
        //    return -1;
        //}

        ///// <summary>
        ///// Get right closest position related to the positionToSearchFrom
        /////  return -1 if there is nothing in the list
        ///// </summary>
        ///// <returns></returns>
        //private static int GetRightClosestPosition()
        //{
        //    if (listOfStartTargets.Count > 0)
        //    {
        //        for (int i = 0; i < listOfStartTargets.Count; i++)
        //        {
        //            if (listOfStartTargets[i] > positionToSearchFrom)
        //            {
        //                return listOfStartTargets[i];
        //            }
        //        }
        //        //if we can't find the closest right position of positionToSearchFrom
        //        //it means we have already come to the end of text area
        //        //that's why we should go the the top to search again
        //        return listOfStartTargets[0];
        //    }

        //    return -1;
        //}

    }

}
