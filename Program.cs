using System;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;


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

            // @todo：
            // String pptxPath = args[0];
            String pptxPath = "C:\\Users\\Administrator\\Downloads\\数学学科-演示任务2.pptx";

            if (!File.Exists(pptxPath))
            {
                Console.WriteLine("pptx 文件不存在，直接退出");
                return;
            }

            Global.app = new PowerPoint.Application { Visible = MsoTriState.msoTrue };
            Global.app.Presentations.Open(pptxPath, MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue);

            // 全局变量
            Global.slideWidth = Global.app.ActivePresentation.PageSetup.SlideWidth;
            Global.slideHeight = Global.app.ActivePresentation.PageSetup.SlideHeight;

            Init();

            // ********************************************************************
            // 业务逻辑代码从这里开始！
            // ********************************************************************
            DateTime t = DateTime.Now;

            // 初始化各种全局变量
            // 初始化各种脚本配置和样式配置
            // 记录脚本前的上下文

            foreach (Slide slide in Global.app.ActivePresentation.Slides)
            {
                Global.app.ActiveWindow.View.GotoSlide(slide.SlideIndex);

                // 把页面处理成足够好看的单页
                PageHandler.LayoutSlide(slide);

                // 兼容处理（换行
                // 把页面处理成足够好看的单页
                // 把溢出版心的单页分页
                // 页面的压缩、合并和移动，元素的再生成
                // 场景处理
                // 兼容处理（刷行高
                // 占位元素的位置匹配
                // 加动画
                // 其他不会动版的处理
                // 信息混淆
                // 机器质检

                // foreach (PowerPoint.Shape shape in slide.Shapes)
                // {
                //    Console.WriteLine(shape.Name);
                // }
            }
            Console.WriteLine("运行结束时间：" + DateTime.Now.Subtract(t));

            // 关闭文件
            Global.app.ActivePresentation.Close();
            Global.app.Quit();
            return;
        }

        static private void Init()
        {
            Global.config = InitFormConfig();
            InitProjectIdAndTaskId();
        }

        static private VbaConfig InitFormConfig(
            double width = 33.867,
            double height = 19.05,
            double top = 2,
            double bottom = 2.2,
            double left = 1.45,
            double right = 1.45,
            string slide_animation_type = "ppEffectSplitHorizontalIn",
            string shape_animation_type = "ppEffectAppear",
            double shape_animation_duration = 0.5,
            string animation_section = "notBody",
            bool all_font_bold = false,
            bool all_font_not_italic = false,
            bool g2ɡ = true,
            bool whether_animate_by_paragraph = true,
            string generate_type = "",
            bool move_catalog_to_front = true,
            bool enable_vba_zoom = true,
            bool compatible_with_wps = true,
            bool convert_answer_to_mark = false)
        {
            VbaConfig config = new VbaConfig
            {
                SlideView = new double[2] { Utils.UnitConvert(width), Utils.UnitConvert(height) },
                SlidePadding = new double[4] { Utils.UnitConvert(top), Utils.UnitConvert(bottom), Utils.UnitConvert(left), Utils.UnitConvert(right) },
                SlideAnimationType = slide_animation_type,
                SlideAnimationDuration = 0.5,
                ShapeAnimationType = shape_animation_type,
                ShapeAnimationDuration = shape_animation_duration,
                AnimationSection = animation_section,
                AllFontBold = all_font_bold,
                AllFontNotItalic = all_font_not_italic,
                g2ɡ = g2ɡ,
                Suffix = "pptx",
                WordLocalFolderPath = "C:/doc/",
                WhetherAnimateByParagraph = whether_animate_by_paragraph,
                GenerateType = generate_type,
                MoveCatalogToFront = move_catalog_to_front,
                EnableVbaZoom = enable_vba_zoom,
                CompatibleWithWps = compatible_with_wps,
                Convertanswer2Mark = convert_answer_to_mark
            };
            return config;
        }

        static public void InitProjectIdAndTaskId()
        {
            Regex regex = new Regex(@"\?subject=([^#]*)#pid=([^#]*)#tid=([^#]*)");
            foreach (Design d in Global.app.ActivePresentation.Designs)
            {
                foreach (CustomLayout c in d.SlideMaster.CustomLayouts)
                {
                    if (regex.IsMatch(c.Name))
                    {
                        Global.PptSubject = regex.Match(c.Name).Groups[1].Value;
                        Global.PptProjectId = regex.Match(c.Name).Groups[2].Value;
                        Global.PptTaskId = regex.Match(c.Name).Groups[3].Value;
                        return;
                    }
                }
            }
        }

        static public void InitGlobalVariable()
        {
            double[] view = Global.config.SlideView;
            double[] padding = Global.config.SlidePadding;
            Global.viewLeft = (float)(Global.slideWidth * padding[2] / view[0]);
            Global.viewRight = (float)(Global.slideWidth * (view[0] - padding[3]) / view[0]);
            Global.viewTop = (float)(Global.slideHeight * padding[0] / view[1]);
            Global.viewBottom = (float)(Global.slideHeight * (view[1] - padding[1]) / view[1]);
        }
    }

    internal class PageHandler
    {
        static public void LayoutSlide(Slide slide)
        {
            // 判断当前页是否需要处理
            if (Utils.CheckHasCatalog(slide))
            {
                return;
            }
            List<PowerPoint.Shape> sortedShapes = Utils.GetSortedSlideShapes(slide);
            // - 若存在横向布局的图片和其大题的题干并列，则把图片传递给大题题干
            // - 若存在上下结构的选项图片，则尝试把文本元素拉开，尽量避免渲染差异
            Shape imageShape = null;
            foreach (Shape shape in sortedShapes)
            {
                if (shape.Type == MsoShapeType.msoPicture &&
                    !Utils.CheckMatchPositionShape(shape) &&
                    shape.Name.Contains("Q") &&
                    shape.Name.Contains("hastextimagelayout=1") &&
                    shape.Left > Global.slideWidth / 2)
                {
                    imageShape = shape;
                }
                if (imageShape != null)
                {
                    if (shape.HasTextFrame == MsoTriState.msoTrue &&
                        Utils.GetShapeInfo(shape)[0] == Utils.GetShapeInfo(imageShape)[4] &&
                        shape.Top > imageShape.Top &&
                        imageShape.Top - shape.Top < Global.GapBetweenTextLine[2] &&
                        Utils.CheckYOverShapes(shape, imageShape, slide, slide, 0) &&
                        shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        // 找到了，就是它！！！
                        imageShape.Name = imageShape.Name.Replace(
                            Utils.GetShapeInfo(imageShape)[0],
                            Utils.GetShapeInfo(shape)[0]
                        );
                        imageShape.Name = imageShape.Name.Replace(
                            Utils.GetShapeInfo(imageShape)[4],
                            Utils.GetShapeInfo(shape)[4]
                        );
                        Shape imageTip = Utils.FindImageTipWithImage(imageShape);
                        if (imageTip != null)
                        {
                            imageTip.Name = imageTip.Name.Replace(
                                Utils.GetShapeInfo(imageTip)[0],
                                Utils.GetShapeInfo(shape)[0]
                            );
                            imageTip.Name = imageTip.Name.Replace(
                                Utils.GetShapeInfo(imageTip)[4],
                                Utils.GetShapeInfo(shape)[4]
                            );
                        }
                        break;
                    }
                }
            }
            // 记录移动前的位置
            List<object[]> processedBefore = new List<object[]>();
            foreach (Shape shape in sortedShapes)
            {
                processedBefore.Add(new object[] { shape.Top });
            }
            // 构建需要进行布局处理的文本框的集合
            List<object[]> nodes = GenerateNodeShapes(slide);
            // - 先对逻辑节点内部进行排版处理
            // - 再对逻辑节点进行排版处理
            for (int n = 0; n < nodes.Count; n++)
            {
                float offset;
                bool noNeedProcess = false;
                if (((List<Shape>)nodes[n][0])[0].Name.Substring(0, 2) == "C_")
                {
                    noNeedProcess = true;
                }
                if (!noNeedProcess)
                {
                    List<List<Shape>> blocks = GenerateBlockShapes((List<Shape>)nodes[n][0], (string)nodes[n][1]);
                    List<Shape> prevBlock;
                    float prevBlockLeft;
                    float prevBlockRight;
                    float prevBlockTop;
                    float prevBlockBottom;
                    List<float[]> blocksInfo = new List<float[]>();
                    List<List<Shape>> baseBlocks = new List<List<Shape>>();
                    for (int b = 0; b < blocks.Count; b++)
                    {
                        blocksInfo.Add(new float[] {
                            (float)Utils.ComputeLogicalNodeLeft(blocks[b]),
                            (float)Utils.ComputeLogicalNodeRight(blocks[b]),
                            (float)Utils.ComputeLogicalNodeTop(blocks[b]),
                            (float)Utils.ComputeLogicalNodeBottom(blocks[b])
                        });
                    }
                    // Block 内部的排布
                    for (int b = 0; b < blocks.Count; b++)
                    {
                        bool hasProcessed = false;
                        // 对于横向排列的多个选项，进行文本框的底部对齐
                        if (blocks[b][0].Name.Contains("QC") &&
                            blocks[b][0].HasTextFrame == MsoTriState.msoTrue &&
                            blocks[b][0].Name.Contains("hastextimagelayout=1") &&
                            blocks[b].Count >= 2 &&
                            !blocks[b][0].Name.Contains("imageTipindex"))
                        {
                            // - 这里需要注意，不要误伤图说！！！
                            // - 这里需要注意，不要误伤图文布局的选项！！！
                            bool canimove = false;
                            for (int s = 0; s < blocks[b].Count; s++)
                            {
                                // Tips：若找到选项在版心右侧，则说明是横向排列的多个选项。
                                if (blocks[b][s].Left + Utils.ComputeShapeWidth(blocks[b][s]) / 2 > Global.slideWidth / 2)
                                {
                                    canimove = true;
                                    break;
                                }
                            }
                            if (canimove && !Regex.IsMatch(blocks[b][0].TextFrame.TextRange.Text, "^[ABCDEFG]\\."))
                            {
                                canimove = false;
                            }
                            if (canimove)
                            {
                                float minLeft = 999;
                                float maxRight = -1;
                                float sumWidth = 0;
                                for (int s = 0; s < blocks[b].Count; s++)
                                {
                                    Shape optionShape = blocks[b][s];
                                    if (optionShape.Left < minLeft)
                                    {
                                        minLeft = optionShape.Left;
                                    }
                                    if (optionShape.Left + optionShape.Width > maxRight)
                                    {
                                        maxRight = optionShape.Left + optionShape.Width;
                                    }
                                    sumWidth += optionShape.Width;
                                }
                                int lines = (int)Math.Round(sumWidth / (maxRight - minLeft));
                                if (lines <= 0)
                                {
                                    lines = 1;
                                }
                                int optionsPerLine = (int)Math.Round((double)(blocks[b].Count / lines));
                                List<List<Shape>> options = new List<List<Shape>>();
                                List<Shape> optionsGroup = new List<Shape>();
                                int o = 1;
                                for (int s = 0; s < blocks[b].Count; s++)
                                {
                                    if (o <= optionsPerLine)
                                    {
                                        optionsGroup.Add(blocks[b][s]);
                                    }
                                    else
                                    {
                                        o = 1;
                                        options.Add(optionsGroup);
                                        optionsGroup = new List<Shape>() { blocks[b][s] };
                                    }
                                    o++;
                                }
                                if (o > optionsPerLine)
                                {
                                    options.Add(optionsGroup);
                                }
                                for (o = 0; o < options.Count; o++)
                                {
                                    float maxBottom = -1;
                                    bool hasInlineImage = false;
                                    for (int p = 0; p < options[o].Count; p++)
                                    {
                                        if (Utils.ComputeShapeBottom(options[o][p]) > maxBottom)
                                        {
                                            maxBottom = (float)Utils.ComputeShapeBottom(options[o][p]);
                                        }
                                        if (Regex.IsMatch(options[o][p].TextFrame.TextRange.Text, "&\\d+&"))
                                        {
                                            hasInlineImage = true;
                                        }
                                    }
                                    if (hasInlineImage && maxBottom > -1)
                                    {
                                        for (int p = 0; p < options[o].Count; p++)
                                        {
                                            options[o][p].Top = options[o][p].Top + maxBottom - (float)Utils.ComputeShapeBottom(options[o][p]);
                                        }
                                    }
                                }
                                for (o = 2; o <= options.Count; o++)
                                {
                                    Shape currentOption = null;
                                    float minTop = 99999;
                                    for (int p = 0; p < options[o].Count; p++)
                                    {
                                        if (options[o][p].Top < minTop)
                                        {
                                            minTop = options[o][p].Top;
                                            currentOption = options[o][p];
                                        }
                                    }
                                    Shape prevOption = options[o - 1][0];
                                    offset = (float)Utils.ComputeShapeBottom(prevOption) - currentOption.Top;
                                    offset += (float)HandleTextLineHeight(currentOption, prevOption);
                                    for (int p = 0; p < options[o].Count; p++)
                                    {
                                        options[o][p].Top = options[o][p].Top + offset;
                                    }
                                }
                                hasProcessed = true;
                            }
                            // 题图+答图
                            if (!hasProcessed && ((string)nodes[n][0]).Contains("SUAN"))
                            {
                                for (int s = 0; s < blocks[b].Count; s++)
                                {
                                    // - 图片和文本的距离
                                    // - @todo：表格
                                    // - @todo：文本间
                                    if (s > 1 && blocks[b][s].Type == MsoShapeType.msoPicture)
                                    {
                                        if (blocks[b][s - 1].Type != MsoShapeType.msoPicture)
                                        {
                                            blocks[b][s].Top = (float)Utils.ComputeShapeBottom(blocks[b][s - 1]) + 10;
                                        }
                                    }
                                    if (s > 1 && blocks[b][s].Type != MsoShapeType.msoPicture)
                                    {
                                        if (blocks[b][s - 1].Type == MsoShapeType.msoPicture)
                                        {
                                            blocks[b][s].Top = (float)Utils.ComputeShapeBottom(blocks[b][s - 1]) + 10;
                                        }
                                    }
                                }
                            }
                            // Block 内部的流式布局
                            if (!hasProcessed)
                            {
                                canimove = true;
                                for (int s = 0; s < blocks[b].Count; s++)
                                {
                                    for (int d = s + 1; d < blocks[b].Count; d++)
                                    {
                                        // - 若存在非同侧的则不允许进行干涉
                                        if (Utils.CheckHasDiffside(blocks[b][s], blocks[b][d]))
                                        {
                                            canimove = false;
                                            break;
                                        }
                                    }
                                    if (!canimove)
                                    {
                                        break;
                                    }
                                }
                                if (canimove)
                                {
                                    for (int s = 1; s < blocks[b].Count; s++)
                                    {
                                        canimove = false;
                                        // - 在同侧
                                        if (Utils.CheckIsExpaned(blocks[b][s]) && Utils.CheckIsExpaned(blocks[b][s - 1]))
                                        {
                                            canimove = true;
                                        }
                                        if (!canimove &&
                                            !Utils.CheckIsExpaned(blocks[b][s]) &&
                                            !Utils.CheckIsExpaned(blocks[b][s - 1]) &&
                                            !Utils.CheckHasDiffside(blocks[b][s], blocks[b][s - 1]) &&
                                            !Utils.CheckMatchPositionShape(blocks[b][s]) &&
                                            !Utils.CheckMatchPositionShape(blocks[b][s - 1]))
                                        {
                                            canimove = true;
                                        }
                                        if (canimove)
                                        {
                                            float Offset = (float)Utils.ComputeShapeBottom(blocks[b][s - 1]) - blocks[b][s].Top;
                                            if (blocks[b][s - 1].HasTextFrame == MsoTriState.msoFalse || blocks[b][s].HasTextFrame == MsoTriState.msoFalse)
                                            {
                                                Offset += 10;
                                            }
                                            Offset += (float)HandleTextLineHeight(blocks[b][s - 1], blocks[b][s]);
                                            for (int o = s; o <= blocks[b].Count; o++)
                                            {
                                                blocks[b][o].Top += Offset;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        // 判定 Block 的位移基准
                        for (b = 0; b < blocks.Count; b++)
                        {
                            List<Shape> currentBlock = blocks[b];
                            float blockLeft = blocksInfo[b][0];
                            float blockRight = blocksInfo[b][1];
                            float blockTop = blocksInfo[b][2];
                            List<Shape> moveBaseBlock = null;
                            List<List<Shape>> movebaseBlocks = new List<List<Shape>>();
                            int e = 10;
                            for (int a = b - 1; a >= 0; a--)
                            {
                                prevBlock = blocks[a];
                                prevBlockLeft = blocksInfo[a][0];
                                prevBlockRight = blocksInfo[a][1];
                                prevBlockTop = blocksInfo[a][2];
                                prevBlockBottom = blocksInfo[a][3];
                                bool hasDiffSide = Utils.CheckLogicalNodeHasDiffside(currentBlock, prevBlock);
                                if (prevBlockLeft < Global.slideWidth / 2 &&
                                    prevBlockLeft < blockRight &&
                                    prevBlockRight > blockLeft + e &&
                                    Utils.CheckStrictOverBlocks(currentBlock, prevBlock, e))
                                {
                                    movebaseBlocks.Add(prevBlock);
                                }
                                else if (prevBlockLeft >= Global.slideWidth / 2 &&
                                    prevBlockLeft + e < blockRight &&
                                    Utils.CheckStrictOverBlocks(currentBlock, prevBlock, e))
                                {
                                    movebaseBlocks.Add(prevBlock);
                                }
                                else if (prevBlockBottom < blockTop + e && !hasDiffSide)
                                {
                                    movebaseBlocks.Add(prevBlock);
                                }
                                else if (prevBlockTop < blockTop &&
                                    (Utils.CheckOverBlocks(currentBlock, prevBlock, e) || prevBlockBottom >= blockTop + e) &&
                                    (!hasDiffSide || Utils.CheckStrictOverBlocks(currentBlock, prevBlock, e)))
                                {
                                    movebaseBlocks.Add(prevBlock);
                                }
                                else if (prevBlockBottom < blockTop + e && hasDiffSide && blockRight > prevBlockLeft)
                                {
                                    movebaseBlocks.Add(prevBlock);
                                }
                            }
                            if (!currentBlock[0].Name.Contains("hastextimagelayout=1") && movebaseBlocks.Count >= 1)
                            {
                                if (movebaseBlocks[0] != null)
                                {
                                    if (!movebaseBlocks[0][movebaseBlocks[0].Count - 1].Name.Contains("hastextimagelayout=1"))
                                    {
                                        moveBaseBlock = movebaseBlocks[0];
                                    }
                                }
                            }
                            if (moveBaseBlock == null)
                            {
                                float closedBottom = -1;
                                float closedTop = -1;
                                for (int x = 0; x < movebaseBlocks.Count; x++)
                                {
                                    if (movebaseBlocks[x] != null)
                                    {
                                        float bottom = (float)Utils.ComputeLogicalNodeBottom(movebaseBlocks[x]);
                                        float top = (float)Utils.ComputeLogicalNodeTop(movebaseBlocks[x]);
                                        e = 5;
                                        bool canimove = false;
                                        if (Math.Abs(bottom - closedBottom) < e && top > closedTop)
                                        {
                                            canimove = true;
                                        }
                                        if (bottom > closedBottom + e)
                                        {
                                            canimove = true;
                                        }
                                        if (canimove && moveBaseBlock != null)
                                        {
                                            canimove = false;
                                            if (Utils.CheckLogicalNodeHasDiffside(movebaseBlocks[x], moveBaseBlock) &&
                                                bottom > closedBottom + e)
                                            {
                                                canimove = true;
                                            }
                                            else if (top > closedTop)
                                            {
                                                canimove = true;
                                            }
                                        }
                                        if (canimove)
                                        {
                                            closedBottom = bottom;
                                            closedTop = top;
                                            moveBaseBlock = movebaseBlocks[x];
                                        }
                                    }
                                }
                            }
                            baseBlocks.Add(moveBaseBlock);
                        }
                        // Block 的排布
                        for (b = 0; b < blocks.Count; b++)
                        {
                            List<Shape> currentBlock = blocks[b];
                            List<Shape> moveBaseBlock = baseBlocks[b];
                            float Offset = 0;
                            // Block 间纵向的排布
                            if (moveBaseBlock != null)
                            {
                                // Tips：prevShape 需要取逻辑节点中 Bottom 最大的元素。
                                for (int z = 0; z < moveBaseBlock.Count - 1; z++)
                                {
                                    for (int x = z + 1; x < moveBaseBlock.Count; x++)
                                    {
                                        Shape ts1 = moveBaseBlock[z];
                                        Shape ts2 = moveBaseBlock[x];
                                        if (ts1.Top + Utils.ComputeShapeHeight(ts1) < ts2.Top + Utils.ComputeShapeHeight(ts2))
                                        {
                                            (moveBaseBlock[z], moveBaseBlock[x]) = (moveBaseBlock[x], moveBaseBlock[z]);
                                        }
                                    }
                                }
                                Shape prevShape = moveBaseBlock[0];
                                Shape currentShape = currentBlock[0];
                                float bottom = (float)Math.Round(prevShape.Top + Utils.ComputeShapeHeight(prevShape) + 0.5); // 向上取整，避免精度问题导致被认为是重叠
                                Offset = bottom - currentShape.Top;
                                // 对文本和图片/表格间加 10 间距，更加好看
                                if (currentShape.HasTextFrame == MsoTriState.msoFalse || prevShape.HasTextFrame == MsoTriState.msoFalse)
                                {
                                    Offset += 10;
                                }
                                // 若相邻的两个文本框都是文字，则进行移动时还需要考虑行距
                                Offset += (float)HandleTextLineHeight(currentShape, prevShape);
                                // 需要考虑图文混排的选项部分
                                if (currentShape.Name.Contains("QC") &&
                                    currentShape.HasTextFrame == MsoTriState.msoTrue &&
                                    currentShape.Name.Contains("hastextimagelayout=1") &&
                                    currentBlock.Count >= 2)
                                { // 找到 Top 最小的元素，Offset 设置为 10
                                    Offset = (bottom + 10) - (float)Utils.ComputeLogicalNodeTop(currentBlock);
                                }
                            }
                            // Block 间横向的排布：
                            // - 题干对齐到图片顶部（小心不要误伤小题题图和大题+小题题干横向布局的情况！
                            if (moveBaseBlock == null)
                            {
                                float blockLeft = blocksInfo[b][0];
                                if (
                                    currentBlock[0].Name.Contains("hastextimagelayout=1") &&
                                    blockLeft < Global.slideWidth / 2
                                ) // 左侧的文本内容
                                {
                                    for (int a = b - 1; a >= 0; a--)
                                    {
                                        prevBlock = blocks[a];
                                        prevBlockLeft = blocksInfo[a][0];
                                        if (
                                            prevBlock[0].Type == MsoShapeType.msoPicture &&
                                            prevBlock[0].Name.Contains("hastextimagelayout=1") &&
                                            prevBlockLeft > Global.slideWidth / 2 &&
                                            Utils.CheckYOverShapes(currentBlock[0], prevBlock[0], slide, slide, 0) &&
                                            !Utils.CheckStrictOverBlocks(currentBlock, prevBlock, 0)
                                        ) // 右侧的图片元素
                                        {
                                            Offset = prevBlock[0].Top - 1 - currentBlock[0].Top;
                                        }
                                    }
                                    // - 若当前 block 移动后会和上面的其他节点产生重叠，则取消
                                    float currentBlockTop = currentBlock[0].Top;
                                    currentBlock[0].Top = currentBlock[0].Top + Offset;
                                    for (int a = n - 1; a >= 0; a--)
                                    {
                                        for (int s = 0; s < ((List<Shape>)nodes[a][0]).Count; s++)
                                        {
                                            if (Utils.CheckOverShapes(currentBlock[0], ((List<Shape>)nodes[a][0])[s], 0))
                                            {
                                                Offset = 0;
                                                break;
                                            }
                                        }
                                        if (Offset == 0)
                                        {
                                            break;
                                        }
                                    }
                                    currentBlock[0].Top = currentBlockTop;
                                }
                            }
                            // ********************
                            // 动起来！
                            // ********************
                            for (int s = 0; s < currentBlock.Count; s++)
                            {
                                currentBlock[s].Top = currentBlock[s].Top + Offset;
                            }
                        }
                        // - 题图+答图的节点，底部对齐
                        if (((string)nodes[n][1]).Contains("SUAN") && blocks.Count == 2)
                        {
                            float maxBottom = -1;
                            foreach (List<Shape> block in blocks)
                            {
                                if (Utils.ComputeLogicalNodeBottom(block) > maxBottom)
                                {
                                    maxBottom = (float)Utils.ComputeLogicalNodeBottom(block);
                                }
                            }
                            foreach (List<Shape> block in blocks)
                            {
                                offset = maxBottom - (float)Utils.ComputeLogicalNodeBottom(block);
                                foreach (Shape shape in block)
                                {
                                    shape.Top += offset;
                                }
                            }
                        }
                    }
                }
            }
            for (int n = 1; n < nodes.Count; n++)
            {
                List<Shape> currentNode = (List<Shape>)nodes[n][0];
                Shape currentNodeTop = currentNode[0];
                List<List<Shape>> prevNodes = new List<List<Shape>>(); // 基准节点备选
                List<Shape> prevNode = null;
                for (int k = n - 1; k >= 0; k--)
                {
                    // @tips：
                    // 这里需要注意存在 prevNode 中的部分元素和当前节点是横向并列关系的情况，
                    // e.g. 图文环绕，
                    // 要当做基准节点的备选。
                    if (Utils.ComputeLogicalNodeTop((List<Shape>)nodes[k][0]) < Utils.ComputeLogicalNodeTop(currentNode) &&
                        !(Utils.CheckYOverBlocks(currentNode, (List<Shape>)nodes[k][0], slide, slide, 0) &&
                        !Utils.CheckXOverBlocks(currentNode, (List<Shape>)nodes[k][0], slide, slide, 5)))
                    { // 在当前节点上面的，不是横向并列关系的，过滤掉双栏布局的情况
                        prevNodes.Add((List<Shape>)nodes[k][0]);
                    }
                }
                if (prevNodes.Count > 0)
                {
                    // - 若节点是展开的，则直接找距离当前节点最近的节点即可
                    // - 若节点是非展开的，则找到上方距离节点最近的、同侧的、元素所在的节点
                    // - 若节点是展开/非展开并存的（则大概率是图文环绕的场景），去找到上方距离节点最近的、同侧的、元素所在的节点
                    if (Utils.CheckLogicalNodeIsExpanded(currentNode) == 1)
                    {
                        double MaxBottom = -1;
                        for (int k = 0; k < prevNodes.Count; k++)
                        {
                            double preNodeBottom = Utils.ComputeContentBottom(prevNodes[k]);
                            if (preNodeBottom > MaxBottom)
                            {
                                MaxBottom = preNodeBottom;
                                prevNode = prevNodes[k];
                            }
                        }
                    }
                    else if (Utils.CheckLogicalNodeIsExpanded(currentNode) == 3 &&
                        currentNodeTop.Name.Contains("hassurround"))
                    { // 图文环绕的
                        double MaxBottom = -1;
                        double prevNodeTop = -1;
                        for (int k = 0; k < prevNodes.Count; k++)
                        {
                            for (int l = 0; l < prevNodes[k].Count; l++)
                            { // 若上个节点中存在和当前节点同侧的，则认为可以作为基准备选
                                Shape PreNodeShape = prevNodes[k][l];
                                List<Shape> t1 = new List<Shape>() { PreNodeShape };
                                double PreNodeShapeBottom = Utils.ComputeShapeBottom(PreNodeShape);
                                List<Shape> t2 = new List<Shape>() { currentNodeTop };
                                if (
                                    !Utils.CheckLogicalNodeHasDiffside(t1, t2) &&
                                    PreNodeShapeBottom > MaxBottom &&
                                    ((PreNodeShape.Name.Contains("C") && PreNodeShape.Top > prevNodeTop) ||
                                    (!PreNodeShape.Name.Contains("C")))
                                )
                                {
                                    MaxBottom = PreNodeShapeBottom;
                                    prevNode = prevNodes[k];
                                    prevNodeTop = PreNodeShape.Top;
                                }
                            }
                        }
                    }
                    else
                    { // 非展开的
                        double MaxBottom = -1;
                        double prevNodeTop = -1;
                        for (int k = 0; k < prevNodes.Count; k++)
                        {
                            //  @tips：若上个节点中存在和当前节点同侧的，则认为可以作为基准备选。
                            for (int l = 0; l < prevNodes[k].Count; l++)
                            {
                                Shape PreNodeShape = prevNodes[k][l];
                                List<Shape> t = new List<Shape>() { PreNodeShape };
                                double PreNodeShapeBottom = Utils.ComputeShapeBottom(PreNodeShape);
                                if (
                                    !Utils.CheckLogicalNodeHasDiffside(t, currentNode) &&
                                    PreNodeShapeBottom > MaxBottom &&
                                    ((PreNodeShape.Name.Contains("C") && PreNodeShape.Top > prevNodeTop) ||
                                    (!PreNodeShape.Name.Contains("C")))
                                )
                                {
                                    // @tips：
                                    // 这里尤其需要注意，
                                    // 对于标题节点来说，可能由于配置的原因，或者压缩字号、行高的处理，
                                    // 导致标题节点的 bottom 值大于内容节点的 bottom 值影响这里的判定，
                                    // 故而对于标题节点来说，再加上 top 值的判断。
                                    MaxBottom = PreNodeShapeBottom;
                                    prevNode = prevNodes[k];
                                    prevNodeTop = PreNodeShape.Top;
                                }
                            }
                        }
                    }
                }
                if (prevNode != null)
                {
                    if (prevNode.Count > 1)
                    {
                        for (int z = 0; z < prevNode.Count - 1; z++) // 按照 Bottom 从大到小对 prevNode 集合进行排序
                        {
                            for (int x = z + 1; x < prevNode.Count; x++)
                            {
                                Shape ts1 = prevNode[z];
                                Shape ts2 = prevNode[x];
                                if (ts1.Top + Utils.ComputeShapeHeight(ts1) < ts2.Top + Utils.ComputeShapeHeight(ts2))
                                {
                                    Shape ts = prevNode[x];
                                    prevNode.RemoveAt(x);
                                    prevNode.Insert(z, ts);
                                }
                            }
                        }
                        Shape prevNodeBottom = null;
                        // @tips：
                        // 若上一個節點是當前節點的父節點，
                        // 則取上一個節點中、當前節點上方的（或者横向并列的）、Bottom 最大的元素，
                        // 而且要先尝试找同侧的，找不到再找不同侧的。
                        if (Utils.GetShapeInfo(prevNode[0])[0] == Utils.GetShapeInfo(currentNodeTop)[4] && prevNodeBottom == null)
                        {
                            float MinTop = 999;
                            for (int z = 0; z < prevNode.Count; z++)
                            {
                                double prevNodeShapeBottom = Utils.ComputeShapeBottom(prevNode[z]);
                                double Offset1 = currentNodeTop.Top - prevNodeShapeBottom;
                                foreach (Shape Shape2 in currentNode)
                                {
                                    Shape2.Top -= (float)Offset1;
                                }
                                bool caniuse = true;
                                // @tips：
                                // 判斷一下找到的大題節點能不能用，
                                // 把節點移動過去，看看會不會造成页面內容的重疊。
                                List<Shape> prevNodeShapes = new List<Shape>();
                                // @todo：這裡需要優化一下時間複雜度。
                                for (int i = n; i < nodes.Count; i++)
                                {
                                    object[] tPrevNode = nodes[i];
                                    if (((List<Shape>)tPrevNode[0])[0].Id == currentNodeTop.Id)
                                    {
                                        break;
                                    }
                                    foreach (Shape tprevShape in (List<Shape>)tPrevNode[0])
                                    {
                                        prevNodeShapes.Add(tprevShape);
                                    }
                                }
                                foreach (Shape Shape1 in prevNodeShapes)
                                {
                                    foreach (Shape Shape2 in currentNode)
                                    {
                                        // - 严格相交
                                        // - 甚至同侧且往上跃过
                                        if (Utils.CheckStrictOverShapes(Shape1, Shape2, 5) ||
                                            (!Utils.CheckHasDiffside(Shape1, Shape2) &&
                                                Utils.ComputeShapeBottom(Shape2) < Shape1.Top))
                                        {
                                            caniuse = false;
                                            // @tips：这里需要兼容大小题题干和图片横向布局的情况。
                                            if (Shape1.Type == MsoShapeType.msoPicture &&
                                                Utils.CheckHasDiffside(Shape1, Shape2))
                                            {
                                                caniuse = true;
                                            }
                                            if (!caniuse)
                                            {
                                                break;
                                            }
                                        }
                                    }
                                    if (!caniuse)
                                    {
                                        break;
                                    }
                                }
                                foreach (Shape Shape2 in currentNode) // 恢复
                                {
                                    Shape2.Top += (float)Offset1;
                                }
                                // @todo：不记得这里为什么是判断 MinTop，以后可以尝试重构一下。
                                if (caniuse && prevNodeShapeBottom < MinTop)
                                {
                                    MinTop = (float)prevNodeShapeBottom;
                                    prevNodeBottom = prevNode[z];
                                }
                            }
                        }
                        if (Utils.GetShapeInfo(prevNode[0])[0] == Utils.GetShapeInfo(currentNodeTop)[4] &&
                            prevNodeBottom == null)
                        {
                            // - 若同侧，则命中
                            // - 若不同侧，则继续找
                            for (int z = 0; z < prevNode.Count; z++)
                            {
                                if (!Utils.CheckHasDiffside(currentNodeTop, prevNode[z]))
                                {
                                    prevNodeBottom = prevNode[z];
                                    break;
                                }
                            }
                        }
                        // @tips：
                        // 若上一个节点是当前节点的题干，
                        // 则先尝试找同侧的，找不到再找不同侧的。
                        if (Utils.GetShapeInfo(prevNode[0])[0] == Utils.GetShapeInfo(currentNodeTop)[0])
                        {
                            for (int z = 0; z < prevNode.Count; z++)
                            {
                                for (int x = 0; x < currentNode.Count; x++)
                                {
                                    // @tips：这里需要注意，可能存在答案 Top 元素在左侧，题干 Bottom 元素在右侧，从而命中的情况。
                                    if (!Utils.CheckHasDiffside(currentNode[x], prevNode[z]))
                                    {
                                        prevNodeBottom = prevNode[z];
                                        break;
                                    }
                                }
                                if (prevNodeBottom != null)
                                {
                                    break;
                                }
                            }
                        }
                        // @tips：PrevNodeBottom 需要取逻辑节点中 Bottom 最大的元素。
                        if (prevNodeBottom == null)
                        {
                            prevNodeBottom = prevNode[0];
                        }
                        double Bottom = Math.Round(Utils.ComputeShapeTop(prevNodeBottom) + Utils.ComputeShapeHeight(prevNodeBottom) + 0.5); // 向上取整，避免精度问题导致被认为是重叠
                        // @tips：
                        // 若 Prev Node 是标题节点，则这里使用 BoundHeight，因为标题元素的高度往往比内容高度高出很多，
                        // 以后需要注意标题元素是否包含背景。
                        if (prevNodeBottom.Name.StartsWith("C") && prevNodeBottom.HasTextFrame == MsoTriState.msoTrue)
                        {
                            Bottom = Math.Round(Utils.ComputeShapeTop(prevNodeBottom) + prevNodeBottom.TextFrame.TextRange.BoundHeight + 0.5);
                            if (prevNodeBottom.TextFrame.VerticalAnchor == MsoVerticalAnchor.msoAnchorMiddle)
                            {
                                Bottom = Math.Round(Utils.ComputeShapeTop(prevNodeBottom) + prevNodeBottom.Height + 0.5);
                            }
                        }
                        double Offset2 = Bottom - currentNodeTop.Top;
                        if (currentNodeTop.HasTextFrame == MsoTriState.msoFalse ||
                            prevNodeBottom.HasTextFrame == MsoTriState.msoFalse)
                        {
                            Offset2 += 10;
                        }
                        Offset2 += HandleTextLineHeight(currentNodeTop, prevNodeBottom);
                        for (int i = n; i < nodes.Count; i++)
                        {
                            for (int k = 0; k < ((List<Shape>)nodes[i][0]).Count; k++)
                            {
                                ((List<Shape>)nodes[i][0])[k].Top += (float)Offset2;
                            }
                        }
                    }
                }
            }
            // 部分后处理：
            // - 长文本答案的相对位移
            for (int i = 0; i < sortedShapes.Count; i++)
            {
                Shape longTextanswerShape = sortedShapes[i];
                if (longTextanswerShape.Name.Contains("haslongtextanswer"))
                {
                    for (int j = 0; j < sortedShapes.Count; j++)
                    {
                        Shape shape = sortedShapes[j];
                        if (shape.Id != longTextanswerShape.Id &&
                            shape.Type != MsoShapeType.msoPicture &&
                            Utils.CheckYOverShapes(longTextanswerShape, shape, slide, slide, 0) &&
                            Utils.CheckXOverShapes(longTextanswerShape, shape, slide, slide, 0))
                        {
                            // @todo：这种方案有点不好、再想想！！！
                            float bOffset = (float)processedBefore[i][0] - (float)processedBefore[j][0];
                            float aOffset = longTextanswerShape.Top - shape.Top;
                            longTextanswerShape.Top += (bOffset - aOffset);
                            break;
                        }
                    }
                }
            }
        }

        // @description 对相邻两个文本元素进行位置移动，保证间距准确
        static public double HandleTextLineHeight(Shape targetShape, Shape moveBaseShape)
        {
            double offset = 0;
            // WIP：若相邻的两个文本框都是文字，则进行移动时还需要考虑行距
            if (targetShape.HasTextFrame == MsoTriState.msoFalse || moveBaseShape.HasTextFrame == MsoTriState.msoFalse)
            {
                return offset;
            }
            TextRange lastLine = moveBaseShape.TextFrame.TextRange.Lines(moveBaseShape.TextFrame.TextRange.Lines().Count);
            TextRange firstLine = targetShape.TextFrame.TextRange.Lines(1);
            if (lastLine.Text.Length <= 0 || firstLine.Text.Length <= 0)
            {
                return offset;
            }
            float lastFontSize = (float)Utils.GetShapeTextInfo(moveBaseShape)[0];
            float firstFontSize = (float)Utils.GetShapeTextInfo(targetShape)[0];
            bool needProcess = true;
            // - 若首行或末行的高度大于 2 倍的字号，则不进行移动
            // - 若首行和末行的高度差比较大，则不进行移动
            // - 若当前文本框的首行是纯文本，上一个文本框的末行是公式，则位移 GapBetweenTextLine 距离
            // - 若当前文本框的首行是公式，上一个文本框的末行是公式，则不进行位移
            // = 若首行和末行的高度差不多，则位移 GapBetweenTextLine x 距离
            // - 若当前文本是段落，上一个文本是标题，则位移 GapBetweenTextLine x 距离
            // - 若当前行存在比较大的 inline 的图片，则 Offset + 10 更好看一些
            // - 若末行存在 inline 图片，则 Offset + 10 更好看一些
            float minLineHeight = firstLine.BoundHeight;
            if (lastLine.BoundHeight > 0 && lastLine.BoundHeight < minLineHeight)
            {
                minLineHeight = lastLine.BoundHeight;
            }
            if (firstLine.BoundHeight > 2 * firstFontSize ||
                lastLine.BoundHeight > 2 * lastFontSize ||
                Math.Abs(firstLine.BoundHeight - lastLine.BoundHeight) / minLineHeight > 0.5)
            {
                needProcess = false;
            }
            if (needProcess)
            {
                Regex regex = new Regex("(<m>)|(</m>)");
                if (!regex.IsMatch(firstLine.Text) && regex.IsMatch(lastLine.Text) && Global.GapBetweenTextLine[1] > 0)
                {
                    offset = Global.GapBetweenTextLine[1];
                }
                else if (!regex.IsMatch(firstLine.Text) && !regex.IsMatch(lastLine.Text) && firstLine.BoundHeight > lastLine.BoundHeight)
                {
                    offset = firstLine.BoundHeight - lastLine.BoundHeight;
                    // 若文本进行过兼容处理，直接贴合好像更好看一些
                    if (firstLine.ParagraphFormat.SpaceWithin > 2 || lastLine.ParagraphFormat.SpaceWithin > 2)
                    {
                        offset = 0;
                    }
                }
                else if (Math.Abs(firstLine.BoundHeight - lastLine.BoundHeight) <= 2)
                {
                    offset = Global.GapBetweenTextLine[1];
                }
                if ((!targetShape.Name.StartsWith("C_") && moveBaseShape.Name.StartsWith("C_")) ||
                    (targetShape.Name.StartsWith("C_") && moveBaseShape.Name.StartsWith("C_")))
                {
                    double move = Global.GapBetweenTextLine[1];
                    // 计算标题的间距时，稍微乘点系数
                    if (lastLine.BoundHeight > firstLine.BoundHeight)
                    {
                        move *= (lastLine.BoundHeight / firstLine.BoundHeight);
                    }

                    offset = move;
                }
            }
            if ((Regex.IsMatch(firstLine.Text, "&(\\d+)&") || Regex.IsMatch(lastLine.Text, "&(\\d+)&")) && offset == 0)
            {
                offset = 10;
            }
            if (Regex.IsMatch(targetShape.TextFrame.TextRange.Text, "^\\.?\\s*&(\\d+)&\\s*\\.?$"))
            {
                offset = 10;
            }
            // @tips：
            // 行内图片的位置匹配是底部对齐，
            // 对于某些特殊的场景，可能存在匹配位置后图片高于文本框顶部的情况，
            // 故而这里要考虑图片位置匹配后的高度，+10。
            if (Regex.IsMatch(targetShape.TextFrame.TextRange.Text, "&(\\d+)&"))
            {
                double minTop = 999;
                foreach (Match match in Regex.Matches(targetShape.TextFrame.TextRange.Text, "&(\\d+)&"))
                {
                    int matchStart = match.Index + 1;
                    int matchEnd = matchStart + match.Length - 1;
                    int lineNum = Utils.FindLineNum(matchStart, targetShape, targetShape);
                    if (lineNum == 1)
                    {
                        string markIndex = match.Groups[1].Value;
                        Shape image = Utils.FindInlineImage(markIndex);
                        if (image != null)
                        {
                            TextRange cEnd = targetShape.TextFrame.TextRange.Characters(matchEnd);
                            double imageTop = cEnd.BoundTop + cEnd.BoundHeight - image.Height;
                            if (imageTop < minTop && imageTop < targetShape.Top)
                            {
                                minTop = imageTop;
                            }
                        }
                    }
                }
                if (minTop < 999)
                {
                    offset = targetShape.Top - minTop + 10;
                }
            }
            // 这里需要返回 Offset，用于相对位移当前元素下方的元素
            return offset;
        }


        // @description 构造当前页面的节点数据结构
        static public List<object[]> GenerateNodeShapes(Slide slide)
        {
            List<Shape> sortedShapes = Utils.GetSortedSlideShapes(slide);
            Dictionary<string, List<Shape>> map = new Dictionary<string, List<Shape>>();
            List<object[]> generateNodeShapes = new List<object[]>();
            for (int i = 0; i < sortedShapes.Count; i++)
            {
                Shape shape = sortedShapes[i];
                string[] shapeInfo = Utils.GetShapeInfo(shape);
                string shapeNodeId = shapeInfo[0];
                string shapeLabel = shapeInfo[2];
                bool hasTableimageShape = (shapeLabel == ".table_image");
                bool hasFixed = (shapeLabel == ".fixed");
                bool hasInlineImage = Utils.CheckInlineImage(shape);
                bool hasMatchanswer = Utils.CheckMatchPositionAnswer(shape.Name, shape);
                bool hasChildNode = Utils.CheckHasChildNode(shape);
                bool hasLongTextanswer = (shape.Name.Contains("haslongtextanswer"));
                bool hasWb = (shape.Name.Contains("WB"));
                // @tips：
                // 需要匹配位置的元素，不需要进行排版、分页处理。
                // 这里还有但是！图说需要参与排版和分页的处理！
                bool needMatchPosition =
                    hasTableimageShape ||
                    hasFixed ||
                    hasInlineImage ||
                    hasMatchanswer ||
                    hasLongTextanswer ||
                    hasWb;
                // @tips：
                // 若元素是大题的解析，则有可能通过配置在大小题的最后，
                // 这种情况下当做独立的解析进行处理，
                // 否则会出现内容顺序混乱的问题。
                bool hasCollected = false;
                if (hasChildNode &&
                    (shape.Name.Contains("AS") || shape.Name.Contains("EX")))
                {
                    shapeNodeId = shapeNodeId + "#" + "PASEX";
                    hasCollected = true;
                }
                // @tips：
                // 若元素是大题的答案，则有可能通过配置在解析的后面，
                // 这种情况下当做独立的节点进行处理，
                // 否则会出现内容顺序混乱的问题。
                if (hasChildNode &&
                    map.ContainsKey(shapeNodeId + "#" + "PASEX") &&
                    !shapeNodeId.Contains("PASEX"))
                {
                    shapeNodeId = shapeNodeId + "#" + "PAN";
                    hasCollected = true;
                }
                // @tips：收集题图+答图部分：
                // - 当前试题存在题图和答图各 1 个
                // - 题图和答图是横向布局
                // - 题图没有和题干横向布局
                if (shape.Name.StartsWith("Q") && !needMatchPosition)
                {
                    List<Shape> sImages = new List<Shape>();
                    List<Shape> aImages = new List<Shape>();
                    List<Shape> nodeShapes = Utils.FindNode(shapeNodeId, sortedShapes);
                    foreach (var nodeShape in nodeShapes)
                    {
                        if (!Utils.CheckInlineImage(nodeShape) && nodeShape.Type == MsoShapeType.msoPicture)
                        {
                            if (nodeShape.Name.Contains("AN"))
                            {
                                aImages.Add(nodeShape);
                            }
                            else
                            {
                                bool hit = true;
                                foreach (Shape nodeShape2 in nodeShapes)
                                {
                                    if (nodeShape2.Id != nodeShape.Id &&
                                        nodeShape2.HasTextFrame == MsoTriState.msoTrue &&
                                        !nodeShape2.Name.Contains("AN") &&
                                        Utils.CheckYOverShapes(nodeShape, nodeShape2, slide, slide, 0))
                                    {
                                        hit = false;
                                        break;
                                    }
                                }
                                if (hit)
                                {
                                    sImages.Add(nodeShape);
                                }
                            }
                        }
                    }
                    if (aImages.Count == 1 && sImages.Count == 1)
                    {
                        if (aImages[0].Name.Contains("hastextimagelayout=1") &&
                            sImages[0].Name.Contains("hastextimagelayout=1") &&
                            Utils.CheckYOverShapes(aImages[0], sImages[0], slide, slide, 0))
                        {
                            Shape imageTip = Utils.FindImageTipWithImage(sImages[0]);
                            if (shape.Id == sImages[0].Id)
                            {
                                shapeNodeId = shapeNodeId + "#" + "SUAN";
                                hasCollected = true;
                            }
                            else if (shape.Name.Contains("AN"))
                            {
                                shapeNodeId = shapeNodeId + "#" + "SUAN";
                                hasCollected = true;
                            }
                            else if (imageTip != null)
                            {
                                if (shape.Id == imageTip.Id)
                                {
                                    shapeNodeId = shapeNodeId + "#" + "SUAN";
                                    hasCollected = true;
                                }
                            }
                        }
                    }
                }
                if (map.ContainsKey(shapeNodeId + "#" + "SUAN") &&
                    !shapeNodeId.Contains("SUAN") &&
                    !shapeNodeId.Contains("PASEX") &&
                    (shape.Name.Contains("AS") || shape.Name.Contains("EX")))
                {
                    shapeNodeId = shapeNodeId + "#" + "PASEX";
                    hasCollected = true;
                }
                if (!hasCollected && (shape.Name.Contains("AS") || shape.Name.Contains("EX") || shape.Name.Contains("AN")))
                {
                    Shape lastNodeShape = null;
                    if (map.ContainsKey(shapeNodeId))
                    {
                        lastNodeShape = map[shapeNodeId][map[shapeNodeId].Count - 1];
                    }
                    if (lastNodeShape != null)
                    {
                        if (!lastNodeShape.Name.Contains("AS") &&
                            !lastNodeShape.Name.Contains("EX") &&
                            !lastNodeShape.Name.Contains("AN") &&
                            Utils.CheckIsExpaned(shape) &&
                            !Utils.CheckIsExpaned(lastNodeShape))
                        {
                            shapeNodeId = shapeNodeId + "#" + "UNLAYOUTEDANASEX";
                            hasCollected = true;
                        }
                    }
                }
                if (!hasCollected &&
                    (shape.Name.Contains("AS") || shape.Name.Contains("EX") || shape.Name.Contains("AN")) &&
                    shape.Name.Contains("hastextimagelayout"))
                {
                    shapeNodeId = shapeNodeId + "#" + "LAYOUTEDANASEX";
                    hasCollected = true;
                }
                if (!hasCollected && (shape.Name.Contains("AS") || shape.Name.Contains("EX") || shape.Name.Contains("AN")))
                {
                    foreach (Shape nodeShape in Utils.FindNode(shapeNodeId, sortedShapes))
                    {
                        if (nodeShape.Name.Contains("AS") ||
                            nodeShape.Name.Contains("EX") ||
                            nodeShape.Name.Contains("AN") &&
                            nodeShape.Name.Contains("hastextimagelayout"))
                        {
                            shapeNodeId = shapeNodeId + "#" + "LAYOUTEDANASEX";
                            hasCollected = true;
                            break;
                        }
                    }
                }
                if (!needMatchPosition && shapeNodeId != "-1")
                {
                    if (map.ContainsKey(shapeNodeId))
                    {
                        map[shapeNodeId].Add(shape);
                    }
                    else
                    {
                        List<Shape> tc = new List<Shape>() { shape };
                        generateNodeShapes.Add(new object[2] { tc, shapeNodeId });
                        map.Add(shapeNodeId, tc);
                    }
                }
            }
            return generateNodeShapes;
        }

        // @description 构造当前节点内部的块数据结构
        static public List<List<Shape>> GenerateBlockShapes(List<Shape> shapes, string nodeType)
        {
            // @todo：图文布局的元素可以放在一个 Block 里进行布局！！！
            List<List<Shape>> generateBlockShapes = new List<List<Shape>>();
            Dictionary<string, List<Shape>> map = new Dictionary<string, List<Shape>>();
            for (int i = 0; i < shapes.Count - 1; i++)
            {
                for (int j = i + 1; j < shapes.Count; j++)
                {
                    if (shapes[i].Top > shapes[j].Top)
                    {
                        (shapes[i], shapes[j]) = (shapes[j], shapes[i]);
                    }
                }
            }
            // 收集零散的选项部分，横向排列的多个选项：
            // - 是选择题、文本元素
            // - 是 ABCD 开头的文本
            // - 是存在横向排列的其他元素
            List<Shape> choices = new List<Shape>();
            for (int i = 0; i < shapes.Count; i++)
            {
                if (shapes[i].Name.Contains("QC") &&
                    shapes[i].HasTextFrame == MsoTriState.msoTrue &&
                    shapes[i].Name.Contains("hastextimagelayout=1"))
                {
                    if (Regex.IsMatch(shapes[i].TextFrame.TextRange.Text, "^[ABCDEFGHIJK]\\."))
                    {
                        choices.Add(shapes[i]);
                    }
                }
                if (shapes[i].Name.Contains("QC") &&
                    shapes[i].Type == MsoShapeType.msoPicture &&
                    shapes[i].Name.Contains("choice_image"))
                {
                    choices.Add(shapes[i]);
                }
            }
            if (choices.Count > 0)
            {
                // @tips：这里需要注意，不要误伤图文布局的选项！！！
                bool canimove = false;
                for (int s = 0; s < choices.Count; s++)
                {
                    // @tips：若找到选项在版心右侧，则说明是横向排列的多个选项。
                    if (choices[s].Left + Utils.ComputeShapeWidth(choices[s]) / 2 > Global.slideWidth / 2)
                    {
                        canimove = true;
                        break;
                    }
                }
                if (!canimove)
                {
                    choices = new List<Shape>();
                }
            }
            // 收集多个横向布局的图片
            List<Shape> images = new List<Shape>();
            for (int i = 0; i < shapes.Count; i++)
            {
                if (shapes[i].Type == MsoShapeType.msoPicture &&
                    !Utils.CheckInlineImage(shapes[i]) &&
                    shapes[i].Name.Contains("hastextimagelayout=1") &&
                    !shapes[i].Name.Contains("vbaimageposition") &&
                    !shapes[i].Name.Contains("choice_image"))
                {
                    images.Add(shapes[i]);
                    for (int j = i + 1; j < shapes.Count; j++)
                    {
                        if (
                            shapes[j].Type == MsoShapeType.msoPicture &&
                            !Utils.CheckInlineImage(shapes[j]) &&
                            shapes[j].Name.Contains("hastextimagelayout=1") &&
                            !shapes[j].Name.Contains("vbaimageposition") &&
                            Utils.CheckYOverShapes(shapes[i], shapes[j], shapes[i].Parent, shapes[j].Parent, 0) &&
                            Utils.GetShapeInfo(shapes[i])[6] == Utils.GetShapeInfo(shapes[j])[6] &&
                            !shapes[j].Name.Contains("choice_image")
                        )
                        {
                            images.Add(shapes[j]);
                        }
                    }
                    if (images.Count <= 1)
                    {
                        images = new List<Shape>();
                    }
                    if (images.Count > 1)
                    {
                        break;
                    }
                }
            }
            // @tips：加上图说。
            if (images.Count > 1)
            {
                int count = images.Count;
                for (int i = 0; i < count; i++)
                {
                    Shape imageShape = images[i];
                    Shape imageTip = Utils.FindImageTipWithImage(imageShape);
                    if (imageTip != null)
                    {
                        images.Add(imageTip);
                    }
                }
            }
            // 收集圖文佈局的答案/解析/解題思路
            List<Shape> ans = new List<Shape>();
            List<Shape> ass = new List<Shape>();
            List<Shape> exs = new List<Shape>();
            bool hasLayoutedAn = false;
            bool hasLayoutedAs = false;
            bool hasLayoutedEx = false;
            foreach (Shape shape in shapes)
            {
                if (shape.Name.Contains("AN"))
                {
                    ans.Add(shape);
                    if (shape.Type == MsoShapeType.msoPicture &&
                        shape.Name.Contains("hastextimagelayout"))
                    {
                        hasLayoutedAn = true;
                    }
                }
                else if (shape.Name.Contains("AS"))
                {
                    ass.Add(shape);
                    if (shape.Type == MsoShapeType.msoPicture &&
                        shape.Name.Contains("hastextimagelayout"))
                    {
                        hasLayoutedAs = true;
                    }
                }
                else if (shape.Name.Contains("EX"))
                {
                    exs.Add(shape);
                    if (shape.Type == MsoShapeType.msoPicture &&
                        shape.Name.Contains("hastextimagelayout"))
                    {
                        hasLayoutedEx = true;
                    }
                }
            }
            if (!hasLayoutedAn)
            {
                ans = new List<Shape>();
            }
            if (!hasLayoutedAs)
            {
                ass = new List<Shape>();
            }
            if (!hasLayoutedEx)
            {
                exs = new List<Shape>();
            }
            for (int i = 0; i < shapes.Count; i++)
            {
                string shapeKey = shapes[i].Name + "#" + shapes[i].Id;
                bool hasCollected = false;
                // 题图+答图：题图部分
                if (!hasCollected && nodeType.Contains("SUAN"))
                {
                    if (!shapes[i].Name.Contains("AN") &&
                        !shapes[i].Name.Contains("AS") &&
                        !shapes[i].Name.Contains("EX"))
                    {
                        if (!map.ContainsKey("suansu"))
                        {
                            List<Shape> tc = new List<Shape>() { shapes[i] };
                            map.Add("suansu", tc);
                            generateBlockShapes.Add(tc);
                        }
                        else
                        {
                            map["suansu"].Add(shapes[i]);
                        }
                        hasCollected = true;
                    }
                }
                // 题图+答图：答图部分
                if (!hasCollected && nodeType.Contains("SUAN"))
                {
                    if (
                        shapes[i].Name.Contains("AN") ||
                        shapes[i].Name.Contains("AS") ||
                        shapes[i].Name.Contains("EX")
                    )
                    {
                        if (!map.ContainsKey("suanan"))
                        {
                            List<Shape> tc = new List<Shape>() { shapes[i] };
                            map.Add("suanan", tc);
                            generateBlockShapes.Add(tc);
                        }
                        else
                        {
                            map["suanan"].Add(shapes[i]);
                        }
                        hasCollected = true;
                    }
                }
                // 图片+图说
                if (!hasCollected &&
                    Regex.IsMatch(shapes[i].Name, "imagetipindex=(\\d+)") &&
                    Utils.FindShapeIndex(shapes[i], images) == -1)
                {
                    string ImageTipIndex = Regex.Match(shapes[i].Name, "imagetipindex=(\\d+)").Groups[0].Value;
                    bool canicollect = true;
                    if (shapes[i].Type == MsoShapeType.msoPicture)
                    { // 过滤掉没有图说但是存在标记的情况
                        Shape ImageTip = Utils.FindImageTipWithImage(shapes[i]);
                        if (ImageTip == null)
                        {
                            canicollect = false;
                        }
                    }
                    if (canicollect)
                    {
                        string Key = "imagetipindex=" + ImageTipIndex;
                        if (!map.ContainsKey(Key))
                        {
                            List<Shape> tc = new List<Shape> { shapes[i] };
                            map.Add(Key, tc);
                            generateBlockShapes.Add(tc);
                        }
                        else
                        {
                            map[Key].Add(shapes[i]);
                        }
                        hasCollected = true;
                    }
                }
                // 大括号部件
                if (!hasCollected && Regex.IsMatch(shapes[i].Name, "jbbraceid=([\\d\\w]+)"))
                {
                    string jid = Regex.Match(shapes[i].Name, "jbbraceid=([\\d\\w]+)").Groups[0].Value;
                    string Key = "*jbbraceid=" + jid + "*";
                    if (!map.ContainsKey(Key))
                    {
                        List<Shape> tc = new List<Shape>() { shapes[i] };
                        map.Add(Key, tc);
                        generateBlockShapes.Add(tc);
                    }
                    else
                    {
                        map[Key].Add(shapes[i]);
                    }
                    hasCollected = true;
                }
                // x 个被分割的选项，x 个横向排布的元素
                if (!hasCollected && choices.Count > 0)
                {
                    if (Utils.FindShapeIndex(shapes[i], choices) > -1)
                    {
                        if (!map.ContainsKey("choices"))
                        {
                            List<Shape> tc = new List<Shape>() { shapes[i] };
                            map.Add("choices", tc);
                            generateBlockShapes.Add(tc);
                        }
                        else
                        {
                            map["choices"].Add(shapes[i]);
                        }
                        hasCollected = true;
                    }
                }
                // x 个横向排布的图片
                if (!hasCollected && images.Count > 0)
                {
                    if (Utils.FindShapeIndex(shapes[i], images) > -1)
                    {
                        if (!map.ContainsKey("images"))
                        {
                            List<Shape> tc = new List<Shape>() { shapes[i] };
                            map.Add("images", tc);
                            generateBlockShapes.Add(tc);
                        }
                        else
                        {
                            map["images"].Add(shapes[i]);
                        }
                        hasCollected = true;
                    }
                }
                // - 圖文佈局過的答案/解析/解題思路
                // - 图文环绕切开的文本，根据宽度不同，分别标记 block 范围
                if (!hasCollected && ans.Count > 0)
                {
                    if (Utils.FindShapeIndex(shapes[i], ans) > -1)
                    {
                        string key = "layoutedan" + Math.Round(shapes[i].Width + 0.5);
                        if (!map.ContainsKey(key))
                        {
                            List<Shape> tc = new List<Shape>() { shapes[i] };
                            map.Add(key, tc);
                            generateBlockShapes.Add(tc);
                        }
                        else
                        {
                            map[key].Add(shapes[i]);
                        }
                        hasCollected = true;
                    }
                }
                if (!hasCollected && ass.Count > 0)
                {
                    if (Utils.FindShapeIndex(shapes[i], ass) > -1)
                    {
                        string key = "layoutedas" + Math.Round(shapes[i].Width + 0.5);
                        if (!map.ContainsKey(key))
                        {
                            List<Shape> TC = new List<Shape>() { shapes[i] };
                            map.Add(key, TC);
                            generateBlockShapes.Add(TC);
                        }
                        else
                        {
                            map[key].Add(shapes[i]);
                        }
                        hasCollected = true;
                    }
                }
                if (!hasCollected && exs.Count > 0)
                {
                    if (Utils.FindShapeIndex(shapes[i], exs) > -1)
                    {
                        string key = "layoutedex" + Math.Round(shapes[i].Width + 0.5);
                        if (!map.ContainsKey(key))
                        {
                            List<Shape> TC = new List<Shape>() { shapes[i] };
                            map.Add(key, TC);
                            generateBlockShapes.Add(TC);
                        }
                        else
                        {
                            map[key].Add(shapes[i]);
                        }
                        hasCollected = true;
                    }
                }
                if (!map.ContainsKey(shapeKey) && !hasCollected)
                {
                    List<Shape> tc = new List<Shape>() { shapes[i] };
                    map.Add(shapeKey, tc);
                    generateBlockShapes.Add(tc);
                }
            }
            return generateBlockShapes;
        }
    }
}
