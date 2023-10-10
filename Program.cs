using System;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using System.Linq;


// @description 配置类型结构体
public struct VbaConfig
{
    public double[] SlideView; // 页面视口宽高
    public double[] SlidePadding; // 页面上下左右 Padding
    public object SlideAnimationType; // 页面切换动画
    public object SlideAnimationDuration; // 页面切换动画-时间
    public object ShapeAnimationType; // 内容动画
    public object ShapeAnimationDuration; // 内容动画-时间
    public string AnimationSection; // 动画配置
    public bool AllFontBold; // 字体全为粗体
    public bool AllFontNotItalic; // 字体全不为斜体
    public bool g2ɡ; // g 转化为 ɡ
    public string Suffix; // 后缀类型
    public string WordLocalFolderPath; // 临时文件路径
    public bool WhetherAnimateByParagraph; // 动画是否按段落出
    public string GenerateType; // Ppt 生成类型
    public bool MoveCatalogToFront; // 是否把题号目录从母版移动到前景
    public bool EnableVbaZoom; // 是否允许压缩页面
    public bool CompatibleWithWps; // 是否开启兼容性处理
    public bool Convertanswer2Mark; // 是否把选择题答案转换成对勾形式
}

namespace hexin_csharp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // 详细阅读文档：https://sigmaai.feishu.cn/docs/doccncPV9QaIVlcDYRHs6nEpl8e#ZHEaGa

            bool DEBUG = false;

            string pptxPath, pptxSaveAsPath, pptxImageSavAsPath;

            if (!DEBUG)
            {
                pptxPath = args[0];
                pptxSaveAsPath = args[1];
                pptxImageSavAsPath = args[2];
            }
            else
            {
                pptxPath = "C:\\Users\\17146\\Desktop\\pptx\\集合（学生版）.docx4655e - 副本.pptx";
                pptxSaveAsPath = "C:\\Users\\17146\\Desktop\\pptx\\1（1）.pptx";
                pptxImageSavAsPath = "C:\\Users\\17146\\Desktop\\pptx\\vstopptximages";
            }

            if (!File.Exists(pptxPath))
            {
                Tester.Log("-1001：pptx 文件不存在，直接退出#P00");
                return;
            }

            Global.app = new PowerPoint.Application { Visible = MsoTriState.msoTrue };
            Global.app.Presentations.Open(pptxPath, MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue);

            // 全局变量
            Global.slideWidth = Global.app.ActivePresentation.PageSetup.SlideWidth;
            Global.slideHeight = Global.app.ActivePresentation.PageSetup.SlideHeight;

            Init();
            InitGlobalMap();

            if (DEBUG)
            {
                // 机器质检
                Tester.Test();
            }
            else
            {
                try
                {
                    // 机器质检
                    Tester.Test();
                }
                catch
                {
                    Tester.Log("-1002：系统运行异常#P00");
                }
            }

            // 信息混淆
            ConfuseInformation();

            // 给文件打分
            Tester.Log(Scoring().ToString());

            // 保存文件
            if (!DEBUG)
            {
                Global.app.ActivePresentation.SaveAs(pptxSaveAsPath);
            }


            // 保存图片
            Global.app.ActivePresentation.SaveAs(pptxImageSavAsPath, PpSaveAsFileType.ppSaveAsJPG);

            // 关闭文件
            Global.app.ActivePresentation.Close();
            Global.app.Quit();
            return;
        }

        static public void Init()
        {
            foreach (Design d in Global.app.ActivePresentation.Designs)
            {
                foreach (CustomLayout c in d.SlideMaster.CustomLayouts)
                {
                    if (Regex.IsMatch(c.Name, @"\?subject=([^#]*)#pid=([^#]*)#tid=([^#]*)(#sourcefrom=([^#]*))?"))
                    {
                        Match m = Regex.Match(c.Name, @"\?subject=([^#]*)#pid=([^#]*)#tid=([^#]*)(#sourcefrom=([^#]*))?");
                        Global.pptSubject = m.Groups[1].Value;
                        Global.pptProjectId = m.Groups[2].Value;
                        Global.pptTaskId = m.Groups[3].Value;
                        Global.pptSourceFrom = m.Groups[5].Value;
                        return;
                    }
                }
            }
        }

        static public void InitGlobalMap()
        {
            foreach (Slide slide in Global.app.ActivePresentation.Slides)
            {
                foreach (Shape shape in slide.Shapes)
                {
                    string[] shapeInfo = Utils.GetShapeInfo(shape);
                    string shapeParentNodedId = shapeInfo[4];
                    if (shapeParentNodedId != "-1")
                    {
                        if (Global.GlobalParentNodeMap.ContainsKey(shapeParentNodedId))
                        {
                            Global.GlobalParentNodeMap[shapeParentNodedId].Add(shape);
                        }
                        else
                        {
                            List<Shape> shapes = new List<Shape> { shape };
                            Global.GlobalParentNodeMap.Add(shapeParentNodedId, shapes);
                        }
                    }
                }
            }
        }

        static public void ConfuseInformation()
        {
            foreach (Slide slide in Global.app.ActivePresentation.Slides)
            {
                slide.Name = "hexin slide" + slide.SlideIndex;
                foreach (Shape shape in slide.Shapes)
                {
                    shape.Name = "hexin shape" + shape.Id;
                }
            }
            foreach (Design d in Global.app.ActivePresentation.Designs)
            {
                foreach (CustomLayout c in d.SlideMaster.CustomLayouts)
                {
                    foreach (Shape shape in c.Shapes)
                    {
                        shape.Name = "hexin shape" + shape.Id;
                    }
                }
            }
        }

        static public int Scoring()
        {
            return -1; // @todo：跑通流程，具体打分策略待补充
        }
    }

    internal class Tester
    {

        // @tips：同一页、同类问题，不重复记录。
        static public Dictionary<string, bool> GlobalRecordMap = new Dictionary<string, bool>();

        // -1xxx：系统问题
        // -1001：pptx 文件不存在，直接退出
        // -1002：程序崩溃

        // -2xxx：PPT 问题
        // -2001：存在异常的字符、标记

        // -3xxx：单页问题
        // -3001：内容存在溢出

        // -31xx：分页问题
        // -3101：页面不能太空
        // -3102：页面不能存在单行一页的情况
        // -3103：疑似选择题选项部分被分页

        // -32xx：元素问题
        // -3201：题干中间存在非题干的部分
        // -3202：内容存在重叠
        // -3203：疑似答案回填异常
        // -3204：疑似材料识别异常
        // -3205：试题上面不能有其他被分割的试题节点
        // -3206：若当前试题被分割，则上面不能有其他非父试题节点

        // -33xx：行问题
        // -3301：标点符号不能在行首
        // -3302：单字成行
        // -3303：疑似标题识别异常
        // -3304：左括号不能单独在行末
        // -3305：标题不能在页末
        
        // -34xx：布局问题
        // -3401: 选项布局异常
        
        // -4xxx：docx_html 机器质检问题

        public static void Test()
        {
            foreach (Slide slide in Global.app.ActivePresentation.Slides)
            {
                Global.app.ActiveWindow.View.GotoSlide(slide.SlideIndex);
                TestSlide(slide);
                foreach (Shape shape in slide.Shapes)
                {
                    if (shape.HasTable == MsoTriState.msoTrue)
                    {
                        foreach (Row row in shape.Table.Rows)
                        {
                            foreach (Cell cell in row.Cells)
                            {
                                TestShape(cell.Shape, shape);
                                for (int i = 1; i <= cell.Shape.TextFrame.TextRange.Lines().Count; i++)
                                {
                                    TestLine(cell.Shape.TextFrame.TextRange.Lines(i), i, cell.Shape, shape);
                                }
                            }
                        }
                    }
                    else if (shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        TestShape(shape, shape);
                        for (int i = 1; i <= shape.TextFrame.TextRange.Lines().Count; i++)
                        {
                            TestLine(shape.TextFrame.TextRange.Lines(i), i, shape, shape);
                        }
                    }
                    else
                    {
                        TestShape(shape, shape);
                    }
                }
            }
        }

        public static void TestSlide(Slide slide)
        {
            Slides slides = Global.app.ActivePresentation.Slides;
            List<Shape> sortedShaps = Utils.GetSortedSlideShapes(slide);
            bool isTitlePage = Utils.CheckTitlePage(slide);
            if (sortedShaps.Count == 0)
            {
                return;
            }
            if (!isTitlePage &&
                Utils.ComputeContentBottom(sortedShaps) <= Global.slideHeight * 1 / 3 &&
                Global.pptSourceFrom == "W2PPT")
            {
                bool canipass = false;
                // - 若下一页的开头是标题，则不报
                // - 若下一页的开头是另外一道题，并且和当前页的最后一道题父节点相同，则不报
                // - 若当前页是最后一页，则不报
                // - 若当前页和前后页都是相同父节点的不同节点，并且内容高度无法合并，则不报
                // - 若当前页的前后页都是表格或者图片，并且内容高度无法合并，则不报
                // - 若下一页的开头是另外一道大题，则不报
                // - 若下一页是被分页的其他节点，则不报
                if (slide.SlideIndex < slides.Count)
                {
                    Slide nextSlide = slides[slide.SlideIndex + 1];
                    List<Shape> nextSlideShapes = Utils.GetSortedStaticSlideShapes(nextSlide);
                    if (nextSlideShapes.Count > 0)
                    {
                        Shape firstShape = nextSlideShapes[0];
                        Shape lastShape = sortedShaps[sortedShaps.Count - 1];
                        string[] firstShapeInfo = Utils.GetShapeInfo(firstShape);
                        string[] lastShapeInfo = Utils.GetShapeInfo(lastShape);
                        if (firstShape.Name.StartsWith("C"))
                        {
                            canipass = true;
                        }
                        if (
                            firstShape.Name.StartsWith("Q") &&
                            lastShape.Name.StartsWith("Q") &&
                            firstShapeInfo[0] != lastShapeInfo[0] &&
                            Convert.ToDouble(firstShapeInfo[4]) <= Convert.ToDouble(lastShapeInfo[4]) &&
                            (firstShapeInfo[6] == lastShapeInfo[6] || firstShapeInfo[6] == "BD"))
                        {
                            canipass = true;
                        }
                    }
                }
                if (slide.SlideIndex < slides.Count)
                {
                    if (slides[slide.SlideIndex + 1].Shapes.Count == 0)
                    {
                        canipass = true;
                    }
                }
                if (slide.SlideIndex == slides.Count)
                {
                    canipass = true;
                }
                if (slide.SlideIndex > 1 || slide.SlideIndex < slides.Count)
                {
                    Shape currentShape = sortedShaps[0];
                    string[] currentShapeInfo = Utils.GetShapeInfo(currentShape);
                    double currentContentHeight = Utils.ComputeLogicalNodeHeight(sortedShaps);
                    if (slide.SlideIndex > 1 && slide.SlideIndex < slides.Count)
                    {
                        Slide nextSlide = slides[slide.SlideIndex + 1];
                        Slide prevSlide = slides[slide.SlideIndex - 1];
                        if (nextSlide.Shapes.Count > 0 && prevSlide.Shapes.Count > 0)
                        {
                            Shape lastShape = prevSlide.Shapes[prevSlide.Shapes.Count];
                            Shape firstShape = nextSlide.Shapes[1];
                            string[] lastShapeInfo = Utils.GetShapeInfo(lastShape);
                            string[] firstShapeInfo = Utils.GetShapeInfo(firstShape);
                            double prevContentHeight = Utils.ComputeLogicalNodeHeight(Utils.GetSortedSlideShapes(prevSlide));
                            double nextContentHeight = Utils.ComputeLogicalNodeHeight(Utils.GetSortedSlideShapes(nextSlide));
                            if (lastShapeInfo[0] != firstShapeInfo[0] &&
                                lastShapeInfo[0] != currentShapeInfo[0] &&
                                currentShapeInfo[0] != firstShapeInfo[0] &&
                                (lastShapeInfo[4] == firstShapeInfo[4] || lastShapeInfo[4] == "-1" || firstShapeInfo[4] == "-1") &&
                                (lastShapeInfo[4] == currentShapeInfo[4] || lastShapeInfo[4] == "-1" || currentShapeInfo[4] == "-1") &&
                                (currentShapeInfo[4] == firstShapeInfo[4] || currentShapeInfo[4] == "-1" || firstShapeInfo[4] == "-1") &&
                                prevContentHeight + currentContentHeight > Global.slideHeight &&
                                nextContentHeight + currentContentHeight > Global.slideHeight)
                            {
                                canipass = true;
                            }
                            if (lastShapeInfo[0] == currentShapeInfo[0] &&
                                currentShapeInfo[0] != firstShapeInfo[0] &&
                                prevContentHeight + currentContentHeight > Global.slideHeight)
                            {
                                canipass = true;
                            }
                            if (lastShape.HasTextFrame == MsoTriState.msoFalse &&
                                firstShape.HasTextFrame == MsoTriState.msoFalse &&
                                prevContentHeight + currentContentHeight > Global.slideHeight &&
                                nextContentHeight + currentContentHeight > Global.slideHeight)
                            {
                                canipass = true;
                            }
                        }
                    }
                    else if (slide.SlideIndex == 1)
                    {
                        Slide nextSlide = slides[slide.SlideIndex + 1];
                        if (nextSlide.Shapes.Count > 0)
                        {
                            Shape firstShape = nextSlide.Shapes[1];
                            string[] firstShapeInfo = Utils.GetShapeInfo(firstShape);
                            double nextContentHeight = Utils.ComputeLogicalNodeHeight(Utils.GetSortedSlideShapes(nextSlide));
                            if (currentShapeInfo[0] == firstShapeInfo[0] &&
                                nextContentHeight + currentContentHeight > Global.slideHeight)
                            {
                                canipass = true;
                            }
                        }
                    }
                    else if (slide.SlideIndex == slides.Count)
                    {
                        Slide prevSlide = slides[slide.SlideIndex - 1];
                        if (prevSlide.Shapes.Count > 0)
                        {
                            Shape lastShape = prevSlide.Shapes[prevSlide.Shapes.Count];
                            string[] lastShapeInfo = Utils.GetShapeInfo(lastShape);
                            double prevContentHeight = Utils.ComputeLogicalNodeHeight(Utils.GetSortedSlideShapes(prevSlide));
                            if (lastShapeInfo[0] != currentShapeInfo[0] &&
                                lastShapeInfo[4] == currentShapeInfo[4] &&
                                prevContentHeight + currentContentHeight > Global.slideHeight)
                            {
                                canipass = true;
                            }
                        }
                    }
                }
                if (slide.SlideIndex < slides.Count)
                {
                    Slide nextSlide = slides[slide.SlideIndex + 1];
                    List<Shape> nextShapes = Utils.GetSortedStaticSlideShapes(nextSlide);
                    if (nextShapes.Count > 0)
                    {
                        Shape firstShape = nextShapes[0];
                        if (Utils.CheckHasChildNode(firstShape))
                        {
                            canipass = true;
                        }
                    }
                }
                if (slide.SlideIndex < slides.Count - 1)
                {
                    Slide nextSlide = slides[slide.SlideIndex + 1];
                    Slide nextNextSlide = slides[slide.SlideIndex + 1];
                    List<Shape> nextShapes = Utils.GetSortedStaticSlideShapes(nextSlide);
                    List<Shape> nextNextShapes = Utils.GetSortedStaticSlideShapes(nextNextSlide);
                    if (nextShapes.Count > 0 && nextNextShapes.Count > 0)
                    {
                        string[] nextShapeInfo = Utils.GetShapeInfo(nextShapes[0]);
                        string[] nextNextShapeInfo = Utils.GetShapeInfo(nextNextShapes[0]);
                        string[] shapeInfo = Utils.GetShapeInfo(sortedShaps[0]);
                        if (nextShapeInfo[0] == nextNextShapeInfo[0] &&
                            shapeInfo[0] != nextShapeInfo[0])
                        {
                            canipass = true;
                        }
                    }
                }
                if (!canipass)
                {
                    Log("-3101#" + slide.SlideIndex + "#页面不能太空#P00");
                }
            }
            if (!isTitlePage &&
                sortedShaps.Count == 1 &&
                slide.SlideIndex < slides.Count - 1)
            {
                Shape shape = sortedShaps[0];
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    if (shape.TextFrame.TextRange.Lines().Count == 1)
                    {
                        bool canipass = false;
                        // - 若上一页是一个超大的表格或者图片，下一页是其他非标题节点且没有祖孙关系，则不报
                        // - 若内容高度比较大，则不报（有可能是一个超大的公式
                        if (slide.SlideIndex > 1 &&
                            slides[slide.SlideIndex - 1].Shapes.Count > 0 &&
                            slides[slide.SlideIndex + 1].Shapes.Count > 0)
                        {
                            List<Shape> prevShapes = Utils.GetSortedSlideShapes(slides[slide.SlideIndex - 1]);
                            List<Shape> nextShapes = Utils.GetSortedSlideShapes(slides[slide.SlideIndex + 1]);
                            Shape lastShape = prevShapes[prevShapes.Count - 1];
                            Shape firstShape = nextShapes[0];
                            string[] firstShapeInfo = Utils.GetShapeInfo(firstShape);
                            string[] currentShapeInfo = Utils.GetShapeInfo(shape);
                            if (lastShape.HasTextFrame == MsoTriState.msoFalse &&
                                Utils.ComputeLogicalNodeHeight(prevShapes) + Utils.ComputeShapeHeight(shape) > Global.slideHeight)
                            {
                                if (firstShape.Name.StartsWith("C"))
                                {
                                    canipass = true;
                                }
                                else if (firstShapeInfo[0] != currentShapeInfo[0] &&
                                    firstShapeInfo[4] != currentShapeInfo[0])
                                {
                                    canipass = true;
                                }
                            }
                        }
                        if (Utils.ComputeShapeHeight(shape) > Global.slideHeight / 3)
                        {
                            canipass = true;
                        }
                        if (!canipass)
                        {
                            Log("-3102#" + slide.SlideIndex + "#页面不能存在单行一页的情况#P00");
                        }
                    }
                }
            }
            if (Utils.CheckSlideOverFlow(slide))
            {
                Log("-3001#" + slide.SlideIndex + "#内容存在溢出#P0");
            }
            List<Shape> shapes = Utils.GetSortedStaticSlideShapes(slide);
            for (int i = 1; i < shapes.Count - 1; i++)
            {
                string prevProp = Utils.GetShapeInfo(shapes[i - 1])[6];
                string prevNodeId = Utils.GetShapeInfo(shapes[i - 1])[0];
                string prop = Utils.GetShapeInfo(shapes[i])[6];
                string nextProp = Utils.GetShapeInfo(shapes[i + 1])[6];
                string nextNodeId = Utils.GetShapeInfo(shapes[i + 1])[0];
                if (prop != "-1" &&
                    !shapes[i].Name.Contains("AN") &&
                    prevProp == nextProp &&
                    prevProp != prop &&
                    prevProp == "BD" &&
                    prevNodeId == nextNodeId &&
                    !shapes[i].Name.Contains("hastextimagelayout") &&
                    !shapes[i - 1].Name.Contains("hastextimagelayout") &&
                    !shapes[i + 1].Name.Contains("hastextimagelayout")) // 过滤掉横向布局的情况
                {
                    Log("-3201#" + slide.SlideIndex + "#题干中间存在非题干的部分#P00");
                }
            }
        }

        public static void TestShape(Shape shape, Shape containerShape)
        {
            Slide slide = containerShape.Parent;
            List<Shape> shapes = Utils.GetSortedStaticSlideShapes(slide);
            int shapeIndex = Utils.FindShapeIndex(containerShape, shapes);
            
            // 检查选项布局和标题异常, 只有是 QC 并且不是 AN AS
            if (shape.HasTextFrame == MsoTriState.msoTrue && 
                shape.Name.StartsWith("QC") && 
                !shape.Name.Contains("AN") &&
                !shape.Name.Contains("AS") )
            {
                    List<int> optionCountsPerLine = new List<int>();
                    Regex regexA = new Regex(@"A\..*");
                    Regex regexB = new Regex(@"B\..*");
                    Regex regexC = new Regex(@"C\..*");
                    Regex regexD = new Regex(@"D\..*");
                    TextRange textRange = shape.TextFrame.TextRange;
                    // string bodyType = 
                    int totalOptionsCount = 0;
                    // 收集每一行的选项数量
                    for (int i = 1; i <= textRange.Lines().Count; i++)
                    {
                        try
                        {
                            string lineText = textRange.Lines(i).Text;
                            MatchCollection matchA = regexA.Matches(lineText);
                            MatchCollection matchB = regexB.Matches(lineText);
                            MatchCollection matchC = regexC.Matches(lineText);
                            MatchCollection matchD = regexD.Matches(lineText);
                            if(regexA.IsMatch(lineText))
                            { 
                                totalOptionsCount += matchA.Count;
                            }
                            if(regexB.IsMatch(lineText) )
                            { 
                                totalOptionsCount += matchB.Count;
                            }
                            if(regexC.IsMatch(lineText))
                            { 
                                totalOptionsCount += matchC.Count;
                            }
                            if(regexD.IsMatch(lineText))
                            { 
                                // 收集每一行的选项数量
                                totalOptionsCount += matchD.Count;
                            }
                            optionCountsPerLine.Add(totalOptionsCount);
                            totalOptionsCount = 0;
                        }
                        catch (ArgumentException)
                        {
                            // Log( "无法获取第" + slide.SlideIndex + "页: " + "第 " + i + " 段的文本");
                        }
                    }
                    int total = 0;
                    foreach (var lineCount in optionCountsPerLine)
                    {
                        // 获取当前shape的总的选项个数
                        total += lineCount;
                    }
                    if (total == 4 && 
                        !optionCountsPerLine.Contains(4) && 
                        optionCountsPerLine.Any(count => count != optionCountsPerLine[0]))
                    {
                        Log("-3401#" + slide.SlideIndex + "#选项布局异常#P00");
                    }
            }
            // @tips：
            // docx_html 环节的机器质检信息。
            // 机器质检信息详细参考：https://gitee.com/lawrencekkk/word_to_fbd/blob/master/fbd_task/module_v3/data_collect.py
            if (Regex.IsMatch(containerShape.Name, @"aifcode=(-?[\d]+)"))
            {
                string aifCode = Regex.Match(containerShape.Name, @"aifcode=(-?[\d]+)").Groups[1].Value;
                if (aifCode == "-1")
                {
                    // 无事发生。。。
                }
                else if (aifCode == "101")
                {
                    Log("-4101#" + slide.SlideIndex + "#上下标异常#P0");
                }
                else if (aifCode == "102")
                {
                    Log("-4102#" + slide.SlideIndex + "#题号位置问题#P0");
                }
                else if (aifCode == "103")
                {
                    Log("-4103#" + slide.SlideIndex + "#表格拆分异常#P00");
                }
                else if (aifCode == "104")
                {
                    Log("-4104#" + slide.SlideIndex + "#答案拆分异常#P00");
                }
                else if (aifCode == "105")
                {
                    Log("-4105#" + slide.SlideIndex + "#试题答案拆分异常#P0");
                }
                else if (aifCode == "1061")
                {
                    Log("-41061#" + slide.SlideIndex + "#讲解类试题选择题答案回插异常#P00");
                }
                else if (aifCode == "1062")
                {
                    Log("-41062#" + slide.SlideIndex + "#讲解类试题填空题答案回插异常#P00");
                }
                else if (aifCode == "1063")
                {
                    Log("-41063#" + slide.SlideIndex + "#讲解类试题解答题答案回插异常#P00");
                }
                else if (aifCode == "107")
                {
                    Log("-4107#" + slide.SlideIndex + "#选项多行问题#P0");
                }
                else if (aifCode == "108")
                {
                    Log("-4108#" + slide.SlideIndex + "#公式问题#P0");
                }
                else if (aifCode == "109")
                {
                    Log("-4109#" + slide.SlideIndex + "#异常加粗问题#P1");
                }
                else if (aifCode == "110")
                {
                    Log("-4110#" + slide.SlideIndex + "#材料题识别异常#P1");
                }
                else if (aifCode == "111")
                {
                    Log("-4111#" + slide.SlideIndex + "#答案解析未拆分问题#P0");
                }
            }
            if (!Utils.CheckMatchPositionShape(containerShape) &&
                containerShape.HasTextFrame == MsoTriState.msoTrue &&
                shapeIndex == 0)
            {
                if (Regex.IsMatch(shape.TextFrame.TextRange.Text, @"^[BCD]\."))
                {
                    Log("-3103#" + slide.SlideIndex + "#疑似选择题选项部分被分页#P00");
                }
            }
            if (!Utils.CheckMatchPositionShape(containerShape) &&
                containerShape.HasTextFrame == MsoTriState.msoTrue &&
                containerShape.Name.Contains("AN"))
            {
                if (containerShape.Name.StartsWith("Q") &&
                    Regex.IsMatch(shape.TextFrame.TextRange.Text, @"^【.*?】[ABCDEFG]\s*$"))
                {
                    Log("-3203#" + slide.SlideIndex + "#疑似答案回填异常#P00");
                }
            }
            if (containerShape.HasTextFrame == MsoTriState.msoTrue)
            {
                if (Regex.IsMatch(shape.TextFrame.TextRange.Text, @"■"))
                {
                    Log("-2001#" + slide.SlideIndex + "#存在异常的字符“■”#P00");
                }
                if (Regex.IsMatch(shape.TextFrame.TextRange.Text, @"$[^$]+$"))
                {
                    string mark = Regex.Match(shape.TextFrame.TextRange.Text, @"$[^$]+$").Value;
                    Log("-2001#" + slide.SlideIndex + "#存在异常的标记#P00");
                }
                if (Regex.IsMatch(shape.TextFrame.TextRange.Text, @"HXDOLLAR"))
                {
                    Log("-2001#" + slide.SlideIndex + "#存在异常的标记#P00");
                }
                if (Regex.IsMatch(shape.TextFrame.TextRange.Text, @"\\\s?[a-zA-Z𝑎𝑏𝑐𝑑𝑒𝑓𝑔𝑖𝑗𝑘𝑙𝑚𝑛𝑜𝑝𝑞𝑟𝑠𝑡𝑢𝑣𝑤𝑥𝑦𝑧]+"))
                {
                    Log("-2001#" + slide.SlideIndex + "#存在异常的标记#P00");
                }
                if (Regex.IsMatch(shape.TextFrame.TextRange.Text, @"&[a-zA-Z𝑎𝑏𝑐𝑑𝑒𝑓𝑔𝑖𝑗𝑘𝑙𝑚𝑛𝑜𝑝𝑞𝑟𝑠𝑡𝑢𝑣𝑤𝑥𝑦𝑧]+;"))
                {
                    Log("-2001#" + slide.SlideIndex + "#存在异常的标记#P00");
                }
                if (Regex.IsMatch(shape.TextFrame.TextRange.Text, @"\$\$") ||
                    Regex.IsMatch(shape.TextFrame.TextRange.Text, @"(\{\{)|(\}\})"))
                {
                    Log("-2001#" + slide.SlideIndex + "#存在异常的标记#P00");
                }
            }
            if (shapeIndex >= 0 &&
                shapeIndex == shapes.Count - 1 &&
                shapes[shapeIndex].HasTextFrame == MsoTriState.msoTrue &&
                containerShape.Top >= Global.slideHeight / 2 &&
                !containerShape.Name.StartsWith("C_"))
            {
                if (Regex.IsMatch(shape.TextFrame.TextRange.Text, @"(回答|完成).*?(([\d\、]+)|(\d+\-\d+))小?题"))
                {
                    Log("-3204#" + slide.SlideIndex + "#疑似材料识别异常#P00");
                }
            }
            // @todo：下面的逻辑性能不太好，可以去掉。
            if (!Utils.CheckTitlePage(slide))
            {
                foreach (Shape otherShape in slide.Shapes)
                {

                    int e = 10;
                    if (containerShape.Id != otherShape.Id &&
                        !Utils.CheckMatchPositionShape(containerShape) &&
                        !Utils.CheckMatchPositionShape(otherShape) &&
                        !Utils.CheckHasDiffside(containerShape, otherShape) &&
                        Utils.CheckStrictOverShapes(containerShape, otherShape, e))
                    {
                        Log("-3202#" + slide.SlideIndex + "#内容存在重叠#P00");
                    }
                    if (containerShape.Id != otherShape.Id &&
                        containerShape.Type == MsoShapeType.msoPicture &&
                        otherShape.Type == MsoShapeType.msoPicture &&
                        Utils.CheckStrictOverShapes(containerShape, otherShape, e))
                    {
                        Log("-3202#" + slide.SlideIndex + "#内容存在重叠#P00");
                    }
                    if (containerShape.Id != otherShape.Id &&
                       Utils.CheckMatchPositionShape(containerShape) &&
                       Utils.CheckMatchPositionShape(otherShape) &&
                       Utils.CheckStrictOverShapes(containerShape, otherShape, e))
                    {
                        Log("-3202#" + slide.SlideIndex + "#内容存在重叠#P00");
                    }
                }
            }
            if (containerShape.Name.StartsWith("Q") &&
                Utils.GetShapeInfo(containerShape)[6] == "BD" &&
                shapeIndex > 0)
            {
                Shape prevLastShape = null;
                Shape nextFirstShape = null;
                Shape currentFirstShape = shapes[0];
                if (slide.SlideIndex > 1)
                {
                    Slide prevSlide = Global.app.ActivePresentation.Slides[slide.SlideIndex - 1];
                    if (prevSlide.Shapes.Count > 0)
                    {
                        List<Shape> prevShapes = Utils.GetSortedSlideShapes(prevSlide);
                        prevLastShape = prevShapes[prevShapes.Count - 1];
                    }
                }
                if (slide.SlideIndex < Global.app.ActivePresentation.Slides.Count)
                {
                    Slide nextSlide = Global.app.ActivePresentation.Slides[slide.SlideIndex + 1];
                    if (nextSlide.Shapes.Count > 0)
                    {
                        List<Shape> nextShapes = Utils.GetSortedSlideShapes(nextSlide);
                        nextFirstShape = nextShapes[0];
                    }
                }
                if (prevLastShape != null &&
                    Utils.GetShapeInfo(prevLastShape)[0] == Utils.GetShapeInfo(currentFirstShape)[0] &&
                    Utils.GetShapeInfo(currentFirstShape)[0] != Utils.GetShapeInfo(containerShape)[0] &&
                    currentFirstShape.Name.StartsWith("Q"))
                {
                    Log("-3205#" + slide.SlideIndex + "#试题上面不能有其他被分割的试题节点#P00");
                }
                if (nextFirstShape != null &&
                    Utils.GetShapeInfo(containerShape)[0] == Utils.GetShapeInfo(nextFirstShape)[0] &&
                    Utils.GetShapeInfo(containerShape)[6] == "BD" &&
                    shapes[shapeIndex - 1].Name.StartsWith("Q") &&
                    Utils.GetShapeInfo(containerShape)[0] != Utils.GetShapeInfo(shapes[shapeIndex - 1])[0] &&
                    Utils.GetShapeInfo(containerShape)[4] != Utils.GetShapeInfo(shapes[shapeIndex - 1])[0])
                {
                    Log("-3206#" + slide.SlideIndex + "#若当前试题被分割，则上面不能有其他非父试题节点#P00");
                }
            }
        }

        public static void TestLine(TextRange line, int lineIndex, Shape shape, Shape containerShape)
        {
            Slide slide = containerShape.Parent;
            if (Utils.CheckTitlePage(slide))
            {
                return;
            }
            Slide nextSlide = null;
            if (slide.SlideIndex < Global.app.ActivePresentation.Slides.Count)
            {
                nextSlide = Global.app.ActivePresentation.Slides[slide.SlideIndex + 1];
            }
            List<Shape> shapes = Utils.GetSortedStaticSlideShapes(slide);
            int shapeIndex = Utils.FindShapeIndex(containerShape, shapes);
            // - @disabled：单字成行的问题目前可以忽略，不反馈给用户。@todo：表格里单列文字的情况忽略
            if (line.Length > 1 &&
                !Regex.IsMatch(line.Text, @"(^\.%\d+%)|(^\.&\d+&)") &&
                !Regex.IsMatch(line.Text, @"^...") &&
                !Regex.IsMatch(line.Text, @"^…") &&
                !Regex.IsMatch(line.Text, @"^[\-\—]{2,}") &&
                !Regex.IsMatch(line.Text, @"^\)。") &&
                Regex.IsMatch(line.Text, @"^\s*[!),.:;?\]、。—ˇ¨〃々～‖…’”〕〉》」』〗】∶！＇），．：；？］｀｜｝]"))
            {
                Log("-3301#" + slide.SlideIndex + "#标点符号不能在行首#P0");
            }
            //if (!Utils.CheckMatchPositionShape(containerShape) &&
            //    shape.TextFrame.TextRange.Lines().Count > 1 &&
            //    line.Length == 1)
            //{
            //    Log("-3302#" + slide.SlideIndex + "#单字成行#P1");
            //}
            if (Global.pptSourceFrom == "W2PPT" &&
                containerShape.HasTextFrame == MsoTriState.msoTrue &&
                lineIndex == shape.TextFrame.TextRange.Lines().Count &&
                shape.TextFrame.TextRange.Lines().Count > 1 &&
                !shape.Name.StartsWith("C") &&
                !Regex.IsMatch(line.Text, @"^[\(\（].*?[\)\）]$") &&  // 绕过，括号包裹的大部分是说明性文字
                !Regex.IsMatch(line.Text, @"[，。,.]$") && // 绕过，行末是标点符号
                !Regex.IsMatch(line.Text, @"【答案】")) // 绕过，包含“【答案】”字样的
            { // 检查末行
                TextRange prevLine = shape.TextFrame.TextRange.Lines(lineIndex - 1);
                bool iserror = false;
                bool canipass = false;
                // - 末行居中，并且上一行不居中
                // - 末行形如"【xxx】"
                // - 末行是加粗的，并且上一行不加粗
                // - 末行形如“四、xxxxxx”
                // - 末行形如“第x卷”
                // - 末行形如“第x部分”
                if (prevLine.ParagraphFormat.Alignment != PpParagraphAlignment.ppAlignCenter &&
                    line.ParagraphFormat.Alignment == PpParagraphAlignment.ppAlignCenter &&
                    !Regex.IsMatch(line.Text, @"^_+$"))
                {
                    iserror = true;
                }
                if (Regex.IsMatch(line.Text, @"^【[^【】]+】"))
                {
                    iserror = true;
                    // - 【范文】xxx
                    // - 【答案】xxx
                    // - 【注】xxx
                    // - 【宋】
                    if (Regex.IsMatch(line.Text, @"【（范文|答案|注|[\u4e00-\u9fa5]）】"))
                    {
                        iserror = false;
                    }
                }
                if (prevLine.Font.Bold == MsoTriState.msoFalse &&
                    line.Font.Bold == MsoTriState.msoTrue)
                {
                    iserror = true;
                }
                if (Regex.IsMatch(line.Text, @"^[一二三四五六六七八九]、") ||
                    Regex.IsMatch(line.Text, @"第.*?卷") ||
                    Regex.IsMatch(line.Text, @"第.*?部分"))
                {
                    iserror = true;
                }
                // - 若下一个元素是表格或者图片，可以不报
                // - 若当前行是公式，可以不报
                // - 若当前元素是答案或者解析，并且在当前页的最后一行，并且下一页的第一个元素是其他题干，可以不报
                if (shapeIndex < shapes.Count - 1)
                {
                    if (shapes[shapeIndex + 1].HasTextFrame == MsoTriState.msoFalse)
                    {
                        canipass = true;
                    }
                }
                if (!canipass)
                {
                    foreach (TextRange c in line.Characters())
                    {
                        if (c.Text == " ")
                        {
                            continue;
                        }
                        if (c.Length > 1)
                        { // 用这种比较 trick 的方式判断内容是否是公式
                            canipass = true;
                        }
                        break;
                    }
                }
                if (!canipass)
                {
                    if ((containerShape.Name.Contains("AN") || containerShape.Name.Contains("AS")) &&
                        shapeIndex == shapes.Count - 1 &&
                        nextSlide != null &&
                        nextSlide.Shapes.Count > 0)
                    {
                        Shape firstShape = Utils.GetSortedStaticSlideShapes(nextSlide)[0];
                        string[] firstShapeInfo = Utils.GetShapeInfo(firstShape);
                        if (firstShapeInfo[6] == "BD")
                        {
                            canipass = true;
                        }
                    }
                }
                if (iserror && !canipass)
                {
                    Log("-3303#" + slide.SlideIndex + "#疑似标题识别异常#P00");
                }
            }
            if (Regex.IsMatch(line.Text, @"[\(\（]\s*$"))
            { // 左括号单独成行
                Log("-3304#" + slide.SlideIndex + "#左括号不能单独在行末#P1");
            }
            if (containerShape.HasTextFrame == MsoTriState.msoTrue &&
                shape.TextFrame.TextRange.Lines().Count == lineIndex &&
                Utils.ComputeShapeBottom(shape) + 1 >= Utils.ComputeContentBottom(shapes) &&
                Global.pptSourceFrom == "W2PPT" &&
                nextSlide != null &&
                nextSlide.Shapes.Count > 0 &&
                !containerShape.Name.Contains("AN") &&
                !containerShape.Name.Contains("linknodeid"))
            { // 检查标题，标题不能在页末
                bool iserror = false;
                bool canipass = false;
                // - 末行形如"（4）xxxxxx"、并且字数不多的，在页面的末尾
                // - 末行是居中的，并且上一行不居中
                // - 末行是加粗的，并且上一行不加粗
                // - 末行形如“四、xxxxxx”
                if (Regex.IsMatch(line.Text, @"^[\(\（]\d+[\)\）]"))
                {
                    // - 需要注意不要误伤知识点
                    // - @wip：需要注意不要误伤答案
                    // - 需要注意不要误伤打分
                    if (line.Length > 8)
                    {
                        canipass = true;
                    }
                    if (lineIndex > 1)
                    {
                        if (Regex.IsMatch(shape.TextFrame.TextRange.Lines(lineIndex - 1).Text, @"^[\(\（]\d+[\)\）]"))
                        {
                            canipass = true;
                        }
                    }
                    if (shapeIndex > 0)
                    {
                        Shape prevShape = shapes[shapeIndex - 1];
                        if (prevShape.HasTextFrame == MsoTriState.msoTrue)
                        {
                            TextRange lastLine = prevShape.TextFrame.TextRange.Lines(prevShape.TextFrame.TextRange.Lines().Count);
                            if (Regex.IsMatch(lastLine.Text, @"^[\(\（]\d+[\)\）]"))
                            {
                                canipass = true;
                            }
                        }
                    }
                    if (Regex.IsMatch(line.Text, @"\d+分"))
                    {
                        canipass = true;
                    }
                    iserror = true;
                }
                TextRange prevLine = null;
                if (shape.TextFrame.TextRange.Lines().Count > 1)
                {
                    prevLine = shape.TextFrame.TextRange.Lines(lineIndex - 1);
                }
                if (shape.TextFrame.TextRange.Lines().Count == 1 && shapes.Count > 1)
                {
                    Shape prevShape = shapes[shapeIndex - 1];
                    if (prevShape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        prevLine = prevShape.TextFrame.TextRange.Lines(prevShape.TextFrame.TextRange.Lines().Count);
                    }
                }
                if (prevLine != null &&
                    prevLine.ParagraphFormat.Alignment != PpParagraphAlignment.ppAlignCenter &&
                    line.ParagraphFormat.Alignment == PpParagraphAlignment.ppAlignCenter &&
                    !Regex.IsMatch(line.Text, @"^[_\s]+$") &&
                    !Regex.IsMatch(line.Text, @"^[\(\（].*?[\)\）]$")) // 避免误伤说明性文字
                {
                    iserror = true;
                }
                if (prevLine != null &&
                    prevLine.Font.Bold == MsoTriState.msoFalse &&
                    line.Font.Bold == MsoTriState.msoTrue)
                {
                    iserror = true;
                }
                if (Regex.IsMatch(line.Text, @"^[一二三四五六六七八九]、"))
                {
                    iserror = true;
                }
                if (iserror && !canipass)
                {
                    Log("-3305#" + slide.SlideIndex + "#标题不能在页末#P00");
                }
            }
        }

        public static void Log(string msg)
        {
            if (Regex.IsMatch(msg, @"(\-\d+)[\：\:](\d+)\#[^#]+"))
            {
                string recordType = Regex.Match(msg, @"(\-\d+)[\：\:](\d+)\#[^#]+").Groups[1].Value;
                string slideIndex = Regex.Match(msg, @"(\-\d+)[\：\:](\d+)\#[^#]+").Groups[2].Value;
                if (GlobalRecordMap.ContainsKey(slideIndex + "#" + recordType))
                {
                    return;
                }
                GlobalRecordMap.Add(slideIndex + "#" + recordType, true);
            }
            Console.WriteLine(msg); // @todo：可以支持写入临时文件
        }
    }
}
