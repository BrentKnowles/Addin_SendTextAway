using System;
using System.Collections.Generic;
using System.Text;
//TODO: THis file has to bve moved to YOMMARKUP, I THINK< but will wait till I get sendawy up and running
namespace CoreUtilities
{
    /// <summary>
    /// A wrapper (only partially complete) to abstract the keywords and such
    /// that I use for my parsing language I developed
    /// </summary>
    static public class YomParse
    {
        // constants
        const int KEYWORDS = 3;
        // types
        enum KeywordTypes  { SCENE, CENTER_TEXT, LEFT_TEXT };
        // Variables
        static public string prefix = "[[";
        static public string postfix = "]]";
        static public int[] system_id = new int[KEYWORDS] { (int)KeywordTypes.SCENE, 
            (int)KeywordTypes.CENTER_TEXT, 
            (int)KeywordTypes.LEFT_TEXT };
        static public string[] system_keywords = new string[KEYWORDS] { "~scene", "~center", "~left" };


        static public string KEYWORD_SCENE  { get {  return system_keywords[(int)KeywordTypes.SCENE];  }  }

        /// <summary>
        /// returns TRUE if sSourceLine contains a reference to the specified keyword
        /// </summary>
        /// <param name="sSourceLine"></param>
        /// <returns></returns>
        static public bool KeywordMatch(string sSourceLine, string Keyword)
        {
            if (sSourceLine.IndexOf(Keyword) > -1)
            {
                // keyword existed on line
                return true;
            }
            return false;
        }

        // Methods
        /// <summary>
        /// returns true if the supplies string has any of the system_keywords in it.
        /// </summary>
        /// <param name="sString"></param>
        /// <returns></returns>
        static public bool StringContainsASystemKeyword(string sString)
        {
            foreach (string s in system_keywords)
            {
                if (sString.IndexOf(s) > -1)
                {
                    return true;
                }
            }
            return false;
        }

    }
}
