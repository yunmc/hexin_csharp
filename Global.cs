using Microsoft.Office.Interop.PowerPoint;
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace hexin_csharp
{
    public class Global
    {
        static public Dictionary<string, bool> GlobalMergedCellMap = new Dictionary<string, bool>();
        static public Dictionary<string, List<Shape>> GlobalParentNodeMap = new Dictionary<string, List<Shape>>();
        static public Dictionary<string, Shape[]> GlobalImageTipMap = new Dictionary<string, Shape[]>();
        static public Dictionary<string, Shape> GlobalInlineImageMap = new Dictionary<string, Shape>();
        static public Dictionary<string, Shape> GlobalTableImageMap = new Dictionary<string, Shape>();
        static public Dictionary<string, Shape[]> GlobalBlankMarkIndexMap = new Dictionary<string, Shape[]>();
        static public Dictionary<string, Shape> GlobalAnswerMarkIndexMap = new Dictionary<string, Shape>();
        static public Dictionary<string, List<Shape>> GlobalNodeShapeMap = new Dictionary<string, List<Shape>>();
        static public Dictionary<string, Shape> GlobalWbImageMap = new Dictionary<string, Shape>();

        static public double[] gapBetweenTextLine = { -1, -1, -1 };

        static public PowerPoint.Application app;

        static public float slideWidth;
        static public float slideHeight;
        static public float viewTop;
        static public float viewLeft;
        static public float viewBottom;
        static public float viewRight;

        static public float pptViewLeft;
        static public float pptViewRight;

        static public int qcSpaceCount = -1; // 全局单选题的作答空间宽度

        static public float standardLineHeight = -1;
        static public float firstLineHeight = -1;// 全局首行磅值普通行高
        static public float lastLineHeight = -1;// 全局末行磅值普通行高
        static public float middleLineHeight = -1;// 全局中间行磅值普通行高
        static public float singleLineHeight = -1; // 全局单行磅值普通行高

        static public VbaConfig config;

        static public string pptProjectId = "-1";
        static public string pptTaskId = "-1";
        static public string pptSubject = "-1";
        static public string pptSourceFrom = "-1";
    }
}
