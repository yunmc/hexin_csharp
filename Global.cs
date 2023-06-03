using Microsoft.Office.Interop.PowerPoint;
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace hexin_csharp
{
    public class Global
    {
        static public Dictionary<string, bool> GlobalMergedCellMap = new Dictionary<string, bool>();
        static public Dictionary<string, Shape[]> GlobalParentNodeMap = new Dictionary<string, Shape[]>();
        static public Dictionary<string, Shape[]> GlobalImageTipMap = new Dictionary<string, Shape[]>();
        static public Dictionary<string, Shape> GlobalInlineImageMap = new Dictionary<string, Shape>();
        static public Dictionary<string, Shape> GlobalTableImageMap = new Dictionary<string, Shape>();
        static public Dictionary<string, Shape[]> GlobalBlankMarkIndexMap = new Dictionary<string, Shape[]>();
        static public Dictionary<string, Shape> GlobalAnswerMarkIndexMap = new Dictionary<string, Shape>();
        static public Dictionary<string, List<Shape>> GlobalNodeShapeMap = new Dictionary<string, List<Shape>>();
        static public Dictionary<string, Shape> GlobalWbImageMap = new Dictionary<string, Shape>();
        
        static public double[] GapBetweenTextLine = { -1, -1, -1 };

        static public PowerPoint.Application app;

        static public float slideWidth;
        static public float slideHeight;
        static public float viewTop;
        static public float viewLeft;
        static public float viewBottom;
        static public float viewRight;

        static public float PptViewLeft;
        static public float PptViewRight;

        static public int qcSpaceCount = -1;

        static public VbaConfig config;

        static public string pptProjectId = "-1";
        static public string pptTaskId = "-1";
        static public string pptSubject = "-1";
    }
}
