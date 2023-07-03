using System;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using System.Drawing;
using static System.Net.Mime.MediaTypeNames;


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

            // 初始化
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

                // ********************************************************************
                // 把页面处理成足够好看的单页
                // ********************************************************************

                // - ？支持动态标题设置背景颜色

                // 标题页的处理
                // - 解决标题内容重叠的问题

                // 内容页的处理
                // - 文本元素的增加（加续表
                // - 集中处理材料下划线和下横线的问题
                // - 刷文本（g2g，刷省略号，刷公式字体颜色
                // - 刷表格（边框，表头
                // - 刷图片背景颜色
                // - 文本元素的拆分（制表符数量不一致的情况
                // - 刷制表符（选项对齐题干，对折行的选项加换行
                // - 单行的处理（占位，作答空间，元素位置匹配，撑行高，溢出文本框的公式/行内图片、复杂公式拆 P，兼容性换行，解决西文换行配置导致不接排的问题，缓解文末小尾巴的问题，处理标点符号句首的问题，删除多余的作答空间下横线，处理括号单独成行的情况
                // - 支持图文环绕的场景
                // - 刷制表符（选项对齐题干，对折行的选项加换行
                // - 刷界标
                // - 刷图片填空
                // - 刷行高（复杂公式 1.1 倍行高，兼容性刷磅值行高，撑高
                // - 刷文本框的尺寸（兼容性加宽
                // - 其他兼容性处理（公式对齐方式，单行答案换行方式
                // - 排版（紧凑图文布局
                // - 元素位置匹配（ocr_match

                if (Utils.CheckTitlePage(slide))
                {
                    continue;
                }

                PageHandler.HandleLine(slide);

                // - 排版（紧凑图文布局
                PageHandler.LayoutSlide(slide);

                // ********************************************************************
                // 把溢出版心的单页分页
                // ********************************************************************
                // - 尝试压缩页面
                // - 分页
                // - 分页后处理（支持题图跨页复制，处理答案在解析后的情况

                // ********************************************************************
                // 页面的压缩、合并和移动
                // ********************************************************************

                // ********************************************************************
                // 场景处理
                // ********************************************************************
                // - 两端对齐
                // - 长文本答案的处理（接排
                // - 大括号的处理
                // - 选择题答案变对勾
                // - 支持对勾答案
                // - 支持行内框
                // - 支持表格斜线
                // - 支持着重点
                // - 支持页码
                // - 支持超链接
                // - 支持题号目录

                // ********************************************************************
                // 加动画
                // ********************************************************************

                // ********************************************************************
                // 收尾工作
                // ********************************************************************
                // - 删除空文本、空动画、空页面
                // - 删除标记

                // ********************************************************************
                // 机器质检
                // ********************************************************************

                // ********************************************************************
                // 信息混淆
                // ********************************************************************
                // - 母版
                // - 页面
                // - 元素

                // - 保存文件
                // - 另存为 PDF
                // - 上传预览图
                // - 返回


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
            InitGlobalVariable();
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
                        Global.pptSubject = regex.Match(c.Name).Groups[1].Value;
                        Global.pptProjectId = regex.Match(c.Name).Groups[2].Value;
                        Global.pptTaskId = regex.Match(c.Name).Groups[3].Value;
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
            Global.gapBetweenTextLine = Utils.ComputeGapBetweenLines();
            Global.standardLineHeight = (float)(Global.gapBetweenTextLine[0] + Global.gapBetweenTextLine[1] + Global.gapBetweenTextLine[2]);
        }
    }

    internal class PageHandler
    {
        // @description：单行的处理
        static public void HandleLine(Slide slide)
        {
            List<Shape[]> processedShapes = new List<Shape[]>();
            List<Shape> processedMatchedShapes = new List<Shape>();
            foreach (Shape shape in slide.Shapes)
            {
                // 做一些全局的数据结构缓存！！！
                if (shape.Type == MsoShapeType.msoPicture)
                {
                    if (Regex.IsMatch(shape.Name, @"tableimageindex=(\d+)"))
                    {
                        string shapeNodeId = Utils.GetShapeInfo(shape)[0];
                        string index = Regex.Match(shape.Name, @"tableimageindex=(\d+)").Groups[1].Value;
                        string key = shapeNodeId + "#" + index;
                        if (!Global.GlobalTableImageMap.ContainsKey(key))
                        {
                            Global.GlobalTableImageMap.Add(key, shape);
                        }
                    }
                    else if (Regex.IsMatch(shape.Name, @"inlineimagemarkindex=(\d+)"))
                    {
                        string index = Regex.Match(shape.Name, @"inlineimagemarkindex=(\d+)").Groups[1].Value;
                        if (!Global.GlobalInlineImageMap.ContainsKey(index))
                        {
                            Global.GlobalInlineImageMap.Add(index, shape);
                        }
                    }
                }
                if (shape.Type == MsoShapeType.msoPicture && shape.Name.Contains("imagetipindex"))
                {
                    string index = Regex.Match(shape.Name, @"imagetipindex=(\d+)").Groups[1].Value;
                    if (!Global.GlobalImageTipMap.ContainsKey(index))
                    {
                        Global.GlobalImageTipMap.Add(index, new Shape[] { shape, null });
                    }
                    else
                    {
                        Global.GlobalImageTipMap[index] = new Shape[] { shape, Global.GlobalImageTipMap[index][1] };
                    }
                }
                if (shape.HasTextFrame == MsoTriState.msoFalse && shape.Name.Contains("imagetipindex"))
                {
                    string index = Regex.Match(shape.Name, @"imagetipindex=(\d+)").Groups[1].Value;
                    if (!Global.GlobalImageTipMap.ContainsKey(index))
                    {
                        Global.GlobalImageTipMap.Add(index, new Shape[] { null, shape });
                    }
                    else
                    {
                        Global.GlobalImageTipMap[index] = new Shape[] { Global.GlobalImageTipMap[index][0], shape };
                    }
                }
                if (shape.HasTextFrame == MsoTriState.msoFalse && shape.Name.Contains("imagetipindex"))
                {
                    string index = Regex.Match(shape.Name, @"imagetipindex=(\d+)").Groups[1].Value;
                    if (!Global.GlobalImageTipMap.ContainsKey(index))
                    {
                        Global.GlobalImageTipMap.Add(index, new Shape[] { null, shape });
                    }
                    else
                    {
                        Global.GlobalImageTipMap[index] = new Shape[] { Global.GlobalImageTipMap[index][0], shape };
                    }
                }
                if (shape.HasTextFrame == MsoTriState.msoTrue && (shape.Name.Contains("AN") || shape.Name.Contains("_WB")))
                {
                    if (Regex.IsMatch(shape.Name, "vbapositionanswer=(\\d+)"))
                    {
                        string index = Regex.Match(shape.Name, "vbapositionanswer=(\\d+)").Groups[1].Value;
                        if (!Global.GlobalAnswerMarkIndexMap.ContainsKey(index))
                        {
                            Global.GlobalAnswerMarkIndexMap.Add(index, shape);
                        }
                        // @todo：这里需要考虑文本答案作答空间在图片上的情况。
                    }
                }
                if (shape.HasTextFrame == MsoTriState.msoTrue && !shape.Name.Contains("AN"))
                {
                    if (Regex.IsMatch(shape.TextFrame.TextRange.Text, "@(\\d+)@"))
                    {
                        foreach (Match match in Regex.Matches(shape.TextFrame.TextRange.Text, "@(\\d+)@"))
                        {
                            string index = match.Groups[1].Value;
                            if (!Global.GlobalBlankMarkIndexMap.ContainsKey(index))
                            {
                                Global.GlobalBlankMarkIndexMap.Add(index, new Shape[] { shape, shape });
                            }
                        }
                    }
                }
                if (shape.HasTable == MsoTriState.msoTrue)
                {
                    foreach (Row row in shape.Table.Rows)
                    {
                        foreach (Cell cell in row.Cells)
                        {
                            if (Regex.IsMatch(cell.Shape.TextFrame.TextRange.Text, "@(\\d+)@"))
                            {
                                foreach (Match match in Regex.Matches(cell.Shape.TextFrame.TextRange.Text, "@(\\d+)@"))
                                {
                                    string index = match.Groups[1].Value;
                                    if (!Global.GlobalBlankMarkIndexMap.ContainsKey(index))
                                    {
                                        Global.GlobalBlankMarkIndexMap.Add(index, new Shape[] { cell.Shape, shape });
                                    }
                                }
                            }
                        }
                    }
                }
                // 收集一下需要处理的元素
                if (!Utils.CheckMatchPositionShape(shape))
                {
                    if (!shape.Name.StartsWith("C_") && shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        processedShapes.Add(new Shape[] { shape, shape });
                    }
                    else if (shape.HasTable == MsoTriState.msoTrue)
                    {
                        foreach (Row row in shape.Table.Rows)
                        {
                            foreach (Cell cell in row.Cells)
                            {
                                processedShapes.Add(new Shape[] { cell.Shape, shape });
                            }
                        }
                    }
                }
                else
                {
                    processedMatchedShapes.Add(shape);
                }
            }
            foreach (Shape shape in processedMatchedShapes)
            {
                // - 复杂公式有效高度的识别
                if (Utils.CheckMatchPositionAnswer(shape))
                {
                    // @todo
                }
            }
            foreach (Shape[] item in processedShapes)
            {
                Shape shape = item[0];
                Shape containerShape = item[1];
                // @todo：记录一下每个 P 的范围，用于在换行后也能支持动画按段播放。
                for (int l = 1; l <= shape.TextFrame.TextRange.Lines().Count; l++)
                {
                    TextRange line = shape.TextFrame.TextRange.Lines(l);
                    TextRange2 line2 = shape.TextFrame2.TextRange.Lines[l];
                    int formulaBeginLineIndex = -1; // 标记当前行是否存在于跨行的公式里
                    if (Regex.IsMatch(line.Text, "<\\/\\?m>"))
                    {
                        MatchCollection matches = Regex.Matches(line.Text, "<\\/\\?m>");
                        Match firstMatch = matches[0];
                        Match lastMatch = matches[matches.Count - 1];
                        if (lastMatch.Value == "<m>")
                        {
                            formulaBeginLineIndex = l;
                        }
                        if (firstMatch.Value == "<m>" && lastMatch.Value == "</m>")
                        {
                            formulaBeginLineIndex = -1;
                        }
                    }
                    // - 占位（行内图片、表格图片
                    // - 作答空间
                    if (Regex.IsMatch(line.Text, @"(&\d+&)|(%\d+%)|(@\d+@)"))
                    {
                        int offset = 0; // 插东西导致的字符索引位移
                        foreach (Match match in Regex.Matches(line.Text, @"(&\d+&)|(%\d+%)|(@\d+@)"))
                        {
                            if (match.Value.Contains("&"))
                            { // 行内图片，@todo：需要考虑行内图片是图片填空的情况
                                offset += HandleImagePosition(match, line, line2, shape, offset);
                            }
                            else if (match.Value.Contains("%"))
                            { // 表格图片
                                offset += HandleImagePosition(match, line, line2, shape, offset);
                            }
                            else if (match.Value.Contains("@"))
                            { // 作答空间
                                offset += HandleBlankPosition(match, line, line2, shape, containerShape, offset);
                            }
                        }
                    }
                    // - 缓解文末小尾巴的问题
                    // - 兼容性换行
                    HandleLineEnter(l, formulaBeginLineIndex, shape, containerShape, slide);
                    // - 解决包含复杂公式的长文本答案位置匹配不准确的问题
                    // - 解决西文换行配置导致不接排的问题
                    // - 解决标点符号句首的问题
                    // - 删除多余的作答空间下横线
                    // - 处理括号单独成行的情况
                    // - 处理长文本答案的转换
                    // - 刷行高
                    HandleFlushLineHeight(l, shape, containerShape, slide);
                }
            }
            foreach (Shape shape in processedMatchedShapes)
            {
                // - 元素位置匹配
                if (Utils.CheckMatchPositionAnswer(shape))
                {
                    PositionMatchHandler.MatchAnswerPosition(shape);
                }
            }
        }

        static public bool HandleLineEnter(int currentLineIndex, int formulaBeginLineIndex, Shape shape, Shape containerShape, Slide slide)
        {
            bool hasEnter = false; // 当前行是否已经插过换行符
            bool needJump = false; // 当前行是否需要跳过、不插换行符
            if (currentLineIndex == shape.TextFrame.TextRange.Lines().Count)
            {
                return false;
            }
            TextRange line = shape.TextFrame.TextRange.Lines(currentLineIndex);
            TextRange nextLine = shape.TextFrame.TextRange.Lines(currentLineIndex + 1);
            TextRange lines2 = shape.TextFrame.TextRange.Lines(currentLineIndex, 2);
            int currentParagraphIndex = Utils.FindLineParagragh(currentLineIndex, shape)[0];
            int nextParagraphIndex = Utils.FindLineParagragh(currentLineIndex + 1, shape)[0];
            if (string.IsNullOrEmpty(nextLine.Text.Trim()) || nextLine.Text.Trim() == "\r")
            {
                return false;
            }
            // 避免误伤下一行的样式，这里记录一下、在下面进行还原
            PpParagraphAlignment oNextLineAlignment = nextLine.ParagraphFormat.Alignment;
            float nextLineMarginLeft = nextLine.BoundLeft - shape.Left;
            if (line.ParagraphFormat.Alignment == PpParagraphAlignment.ppAlignCenter ||
                line.ParagraphFormat.Alignment == PpParagraphAlignment.ppAlignJustify)
            {
                nextLineMarginLeft = 0;
            }
            // - 在换行过程中，避免误伤大括号转换的 braceleft 标记，保持单独成段
            if (!needJump && !hasEnter && Regex.IsMatch(shape.TextFrame.TextRange.Paragraphs(currentParagraphIndex).Text, @"(braceleft)|(braceright)"))
            { // 大括号
                needJump = true;
            }
            // - 在换行过程中，处理作答空间
            // - 在换行过程中，处理公式
            // - 在换行过程中，处理溢出文本框的空格
            // - @todo：在换行过程中，注意不要误伤标记
            // ********************
            // 处理作答空间
            // ********************
            if (!needJump && !hasEnter && Regex.IsMatch(line.Text, @"_+@(\d+)@\s?_*$") && nextLine.Characters(1).Text == "_")
            {
                Match match = Regex.Match(line.Text, @"_+@(\d+)@\s?_*$");
                string markIndex = match.Groups[1].Value;
                Shape answerShape = Utils.FindAnswerWithMarkIndex(markIndex, slide.SlideIndex);
                if (answerShape != null)
                {
                    if (answerShape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        int linesCount = answerShape.TextFrame.TextRange.Lines().Count;
                        // @tips：在前面 handle_answer_position 的时候进行过答案位置的匹配。
                        if (linesCount <= 2)
                        {
                            float positionLeft = line.Find(match.Value).BoundLeft;
                            float lineBoundTop = line.BoundTop;
                            if (containerShape.HasTable == MsoTriState.msoTrue)
                            {
                                positionLeft += containerShape.Left;
                                lineBoundTop += containerShape.Top;
                            }
                            bool hasOverFlowInlineImage = false;
                            float answerFontSize = (float)Utils.GetShapeTextInfo(answerShape)[0];
                            // - 处理答案从第二行开始的情况
                            // - 处理答案无论是否会折行、但是展开后宽度小于文本框的情况
                            // - 处理答案必然需要折行的情况
                            bool neediprocess = true;
                            bool s = true; // 是否需要继续判定 neediprocess 的结果
                            if (s && answerShape.TextFrame.TextRange.BoundTop > lineBoundTop + answerFontSize / 2)
                            { // 处理答案从第二行开始的情况，不需要在当前行留小尾巴
                                neediprocess = true;
                                s = false;
                            }
                            answerShape.TextFrame.WordWrap = MsoTriState.msoFalse; // 拉开答案
                            float answerContentWidth = (float)Utils.ComputeAnswerFirstLineWidth(answerShape, 2);
                            if (Regex.IsMatch(answerShape.TextFrame.TextRange.Text, "&\\d+&\\s+"))
                            {
                                foreach (Match match1 in Regex.Matches(answerShape.TextFrame.TextRange.Text, "&\\d+&\\s+"))
                                {
                                    TextRange inlineImageRange = answerShape.TextFrame.TextRange.Find(match1.Value);
                                    if (inlineImageRange.BoundLeft < shape.Left + shape.Width &&
                                        inlineImageRange.BoundLeft + inlineImageRange.BoundWidth > shape.Left + shape.Width)
                                    {
                                        hasOverFlowInlineImage = true;
                                        break;
                                    }
                                }
                            }
                            if (s && positionLeft + answerContentWidth < shape.Left + shape.Width)
                            { // 处理答案无论是否会折行、但是展开后宽度小于文本框的情况
                                neediprocess = false;
                                s = false;
                            }
                            if (s && hasOverFlowInlineImage)
                            { // 若答案中存在会溢出作答空间的行内图片
                                neediprocess = true;
                                s = false;
                            }
                            if (s)
                            { // 处理答案必然需要折行的情况
                                if (answerContentWidth <= shape.Width)
                                {
                                    if (Utils.ComputeShapeHeight(answerShape) < Global.standardLineHeight)
                                    { // 若答案没有比作答空间行高高多少，则尝试保留不换行
                                        neediprocess = false;
                                    }
                                    else if (shape.Left + shape.Width - positionLeft < answerFontSize * 1.5)
                                    { // 若答案只会在当前行留下比较短的内容
                                        neediprocess = true;
                                    }
                                }
                                else
                                { // 答案超长，则不把作答空间换行下去，@todo：这里待细化
                                    neediprocess = false;
                                }
                            }
                            answerShape.TextFrame.WordWrap = MsoTriState.msoTrue; // 恢复答案
                            if (neediprocess)
                            {
                                if (linesCount > 1)
                                {
                                    // @tips：
                                    // 换行后删除多余的作答空间，
                                    // 因为公式被换行后展开、总体宽度可能会减小，造成作答空间过长问题，
                                    // 根据长度计算需要删除的下横线个数。
                                    double reduceAnswerWidth = Utils.ComputeAnswerFirstLineWidth(answerShape, 1) -
                                        Utils.ComputeAnswerFirstLineWidth(answerShape, 2);
                                    if (reduceAnswerWidth > 0)
                                    {
                                        float spaceWidth = line.Characters(match.Index + 1).BoundWidth;
                                        int spaceCount = (int)(reduceAnswerWidth / spaceWidth - 0.5);
                                        if (spaceCount >= 3)
                                        {
                                            Match match1 = Regex.Match(shape.TextFrame.TextRange.Text, "_+@(" + markIndex + ")@\\s?_+");
                                            TextRange match1Range = shape.TextFrame.TextRange.Characters(
                                                match1.Index + 1 + (markIndex.Length + 2) + 1 + 1,
                                                spaceCount);
                                            if (Regex.IsMatch(match1Range.Text, "^_+$"))
                                            {
                                                match1Range.Delete();
                                            }
                                        }
                                    }
                                }
                                line.Characters(match.Index).InsertAfter("\r");
                                hasEnter = true;
                            }
                        }
                    }
                }
            }
            // @tricks：对于单元格里，公式后跟着括号内容结束的场景，尝试把括号换行下去，避免渲染差异导致的期望外的折行问题
            if (!needJump && !hasEnter && Regex.IsMatch(line.Text, "<\\/m>.*\\（.*?\\）") && containerShape.HasTable == MsoTriState.msoTrue)
            {
                MatchCollection matches = Regex.Matches(line.Text, "<\\/m>.*\\（.*?\\）");
                Match lastMatch = matches[matches.Count - 1];
                int matchStart = lastMatch.Index + 1;
                string matchValue = lastMatch.Value;
                TextRange matchRange = line.Characters(matchStart, matchValue.Length);
                int e = 15;
                if (Math.Abs((matchRange.BoundLeft + matchRange.BoundWidth) -
                    (shape.TextFrame.TextRange.BoundLeft + shape.TextFrame.TextRange.BoundWidth)) < e)
                {
                    matches = Regex.Matches(line.Text, "\\（.*?\\）");
                    lastMatch = matches[matches.Count - 1];
                    matchStart = lastMatch.Index + 1;
                    matchValue = lastMatch.Value;
                    matchRange = line.Characters(matchStart, matchValue.Length);
                    if (matchRange.BoundWidth < shape.Width / 3)
                    {
                        matchRange.InsertBefore("\r");
                        hasEnter = true;
                    }
                }
            }
            // ********************
            // 处理公式
            // ********************
            if (!needJump && !hasEnter && Regex.IsMatch(line.Text, "<\\/?m>"))
            {
                MatchCollection matches = Regex.Matches(line.Text, "<\\/?m>");
                Match lastMatch = matches[matches.Count - 1];
                Match firstMatch = matches[0];
                // @todo：在超高公式的两边拆 P。
                // - 在折行（但是一行放得下的）公式的前面拆 P
                // - 在跨行公式的前面和行末拆 P
                // - 在跨行公式的后面拆 P
                if (lastMatch.Value == "<m>" && lastMatch.Index > 0)
                {
                    // @tips：下一行是可能不存在标记的，由于多行的超大公式。
                    if (Regex.IsMatch(nextLine.Text, "<\\/?m>"))
                    {
                        Match nextFirstMatch = Regex.Matches(nextLine.Text, "<\\/?m>")[0];
                        if (shape.Left + shape.Width - line.Characters(lastMatch.Index + 1).BoundLeft +
                            nextLine.Characters(nextFirstMatch.Index + 1).BoundLeft - shape.Left < shape.Width)
                        { // 公式的长度小于文本框的宽度
                            line.Characters(lastMatch.Index + 1).InsertBefore("\r");
                            hasEnter = true;
                        }
                    }
                }
                if (!hasEnter &&
                    firstMatch.Value == "</m>" &&
                    firstMatch.Index > 0 &&
                    line.Characters(line.Length).Text != "\r" &&
                    matches.Count % 2 == 1 &&
                    formulaBeginLineIndex > 0 &&
                    shape.TextFrame.WordWrap == MsoTriState.msoTrue &&
                    containerShape.HasTable == MsoTriState.msoFalse)
                {
                    bool caniinsert = true;
                    TextRange formularRange = shape.TextFrame.TextRange.Lines(formulaBeginLineIndex, currentLineIndex + 1 - formulaBeginLineIndex);
                    MatchCollection matches1 = Regex.Matches(formularRange.Text, "<\\/?m>");
                    formularRange.Characters(matches1[0].Index + 1).InsertBefore("\r");
                    formularRange.Characters(matches1[1].Index + 5).InsertAfter("\r");
                    shape.TextFrame.WordWrap = MsoTriState.msoFalse; // 这样的判定方式不适合在表格里使用
                    float formulaLineHeight = formularRange.Lines(2).BoundHeight;
                    shape.TextFrame.WordWrap = MsoTriState.msoTrue;
                    formularRange.Characters(matches1[0].Index + 1).Delete();
                    formularRange.Characters(matches1[1].Index + 5).Delete();
                    // @tips：若跨行公式的每一行都不是很高，则尽量保留行后的内容、避免换行得太碎，后面会刷行高的。
                    if (formulaLineHeight - Global.standardLineHeight > Global.gapBetweenTextLine[2])
                    {
                        caniinsert = false;
                    }
                    if (caniinsert)
                    {
                        line.Characters(line.Length).InsertAfter("\r");
                        hasEnter = true;
                    }
                }
                if (!hasEnter && firstMatch.Value == "</m>" && firstMatch.Index > 0 && line.Characters(line.Length).Text != "\r")
                {
                    // @tips：这里需要避免误伤公式后的标点符号，可以保留些非公式的部分。
                    int punctuationCount = 0;
                    if (line.Characters(firstMatch.Index + 5).Text == " ")
                    {
                        punctuationCount++;
                    }
                    if (Regex.IsMatch(line.Characters(firstMatch.Index + 6).Text, "[、.。！!；;’'”\u201c．,，]"))
                    {
                        punctuationCount++;
                    }
                    line.Characters(firstMatch.Index + 1, 4 + punctuationCount).InsertAfter("\r");
                    hasEnter = true;
                }
            }
            // @tips：若跨行的公式没有闭合，则直接跳过，对于无法在上面正常进行换行处理的复杂公式，在两边切开即可。
            if (!needJump && !hasEnter && formulaBeginLineIndex > 0)
            {
                hasEnter = true;
            }
            // ********************
            // 处理溢出文本框的空格
            // ********************
            // - 需要折行下去的行内图片
            // - 需要折行下去的行内框
            // - 行末的连续空格
            if (!needJump &&
                !hasEnter &&
                line.Characters(line.Length).Text != "\r" &&
                Regex.IsMatch(lines2.Text, @"(&\d+&\s+)|(%\d+%\s+)|(%\d+%_+)") &&
                !(shape.Name.Contains("choices") && shape.Name.Contains("hastextimagelayout")) &&
                !Regex.IsMatch(shape.TextFrame.TextRange.Text, @"^([ABCD]\.?\s*((&\d+&\s+)|(%\d+%\s+)|(%\d+%_+))[\s\t]*)+$"))
            {
                // - 图片选项不需要进行处理，信任工具处理的结果
                // - @todo：答案中存在行内图片的情况待处理，可以放在答案位置匹配里进行
                foreach (Match match in Regex.Matches(lines2.Text, @"(&\d+&\s+)|(%\d+%\s+)|(%\d+%_+)"))
                {
                    // - 溢出
                    // - 标记和撑宽空格发生折行
                    // - 单独成行的图片不处理
                    int matchStart = match.Index + 1;
                    int matchEnd = match.Index + match.Length;
                    TextRange matchRange = lines2.Characters(matchStart, match.Length);
                    double matchLeft = matchRange.BoundLeft;
                    if (containerShape.HasTable == MsoTriState.msoTrue)
                    {
                        matchLeft += containerShape.Left;
                    }
                    int lineNum1 = Utils.FindLineNum(matchStart, shape, containerShape);
                    int lineNum2 = Utils.FindLineNum(matchEnd, shape, containerShape);
                    if (match.Index > 0)
                    {
                        float e; // 可以接受的行内图片溢出阈值
                        if (containerShape.HasTable == MsoTriState.msoTrue)
                        {
                            e = shape.Left + shape.Width;
                        }
                        else
                        {
                            e = Global.slideWidth;
                            if (containerShape.Name.Contains("hastextimagelayout"))
                            {
                                e = containerShape.Left + containerShape.Width;
                            }
                        }
                        if (matchLeft + matchRange.BoundWidth > e || lineNum1 != lineNum2)
                        {
                            lines2.Characters(match.Index).InsertAfter("\r");
                            hasEnter = true;
                        }
                    }
                }
            }
            if (!needJump &&
                !hasEnter &&
                line.Characters(line.Length).Text != "\r" &&
                Regex.IsMatch(line.Text, @"\s+$") &&
                line.BoundLeft + line.BoundWidth > shape.Left + shape.Width)
            { // 行末的连续空格
                for (int c = line.Length; c >= 1; c--)
                {
                    if (line.Characters(c).BoundLeft <= shape.Left + shape.Width && line.Characters(c).Text == " ")
                    {
                        if (c == line.Length)
                        {
                            line.Characters(c).InsertAfter("\r");
                            hasEnter = true;
                            break;
                        }
                        else if (line.Characters(c + 1).Text == " ")
                        {
                            line.Characters(c).InsertAfter("\r");
                            hasEnter = true;
                            break;
                        }
                    }
                }
            }
            // 在行末插入换行符！！！
            if (!hasEnter && line.Characters(line.Length).Text != "\r")
            {
                line.Characters(line.Length).InsertAfter("\r");
                hasEnter = true;
            }
            // @tips：需要注意插入换行符可能对下一行的对齐方式造成影响，这里需要还原一下。
            if (hasEnter)
            {
                shape.TextFrame.TextRange.Paragraphs(currentParagraphIndex + 1).ParagraphFormat.Alignment = oNextLineAlignment;
            }
            if (hasEnter && nextLineMarginLeft > 5 && currentParagraphIndex == nextParagraphIndex)
            { // 还原缩进，@todo：这里看起来有点不对、待完善
                nextLine.IndentLevel = 2;
                shape.TextFrame.Ruler.Levels[2].LeftMargin = nextLineMarginLeft;
            }
            return hasEnter;
        }

        static public void HandleFlushLineHeight(int currentLineIndex, Shape shape, Shape containerShape, Slide slide)
        {
            // @todo：对多行答案的处理，解决答案压线的问题。
            // @tips：这里需要注意，有些段落由于没有拆分的跨行公式，是可能多行的。
            int currentParagraphIndex = Utils.FindLineParagragh(currentLineIndex, shape)[0];
            TextRange textRange = shape.TextFrame.TextRange.Paragraphs(currentParagraphIndex);
            float oldBoundHeight = (float)Math.Round(textRange.BoundHeight - 0.5);
            float aLineHeight = (float)Math.Round(textRange.BoundHeight / textRange.Lines().Count + 0.5);
            bool caniprocess = false;
            // - 这里需要注意避免误伤不需要兼容处理的文本元素，状态上会是一 P 多行的形式
            // - 这里需要注意对于无法换行处理的多行公式，需要进行处理
            // - 这里需要注意若段落已经是磅值行高，则不需要进行处理
            if (textRange.Lines().Count == 1 || Utils.CheckLineFeedFormula(textRange))
            {
                caniprocess = true;
            }
            if (textRange.ParagraphFormat.SpaceWithin > 2)
            {
                caniprocess = false;
            }
            // - 这里需要注意，对超高 P 压缩一下其行高
            // - 这里需要注意若多行公式行高差异明显，则暂时不做处理
            if (caniprocess &&
                !Utils.CheckHasTallSpace(textRange, shape) &&
                Regex.IsMatch(textRange.Text, @"<m>.*</m>") &&
                containerShape.HasTable == MsoTriState.msoFalse)
            {
                MatchCollection matches = Regex.Matches(textRange.Text, @"</?m>");
                // - 超高公式行高设置为 1.1 倍，刷为磅值行高
                // - 较高公式，尝试和普通文本行高一致
                // - 高度参差的公式，不刷行高
                if (Utils.CheckHasLargeFormula(textRange) == 2)
                {
                    textRange.ParagraphFormat.LineRuleWithin = MsoTriState.msoTrue;
                    textRange.ParagraphFormat.SpaceWithin = (float)1.1;
                    oldBoundHeight = (float)Math.Round(textRange.BoundHeight - 0.5);
                    aLineHeight = (float)Math.Round(textRange.BoundHeight / textRange.Lines().Count + 0.5);
                }
                if (textRange.Lines().Count > 1)
                {
                    double e = Global.gapBetweenTextLine[2] * 2;
                    if (Math.Abs(textRange.Characters(matches[0].Index + 1).BoundHeight -
                            textRange.Characters(matches[matches.Count - 1].Index + 1).BoundHeight) > e)
                    {
                        textRange.ParagraphFormat.LineRuleWithin = MsoTriState.msoTrue;
                        textRange.ParagraphFormat.SpaceWithin = (float)1.1;
                        caniprocess = false;
                    }
                    
                }
            }
            if (caniprocess)
            {
                if (Utils.CheckHasLargeFormula(textRange) <= 1 && 
                    textRange.Lines().Count == 1 &&
                    textRange.ParagraphFormat.SpaceWithin > 1.1) 
                {
                    if (shape.TextFrame.TextRange.Lines().Count == 1 && Global.singleLineHeight > 2)
                    {
                        textRange.ParagraphFormat.LineRuleWithin = MsoTriState.msoFalse;
                        textRange.ParagraphFormat.SpaceWithin = Global.singleLineHeight;
                        return;
                    }
                    else if (currentLineIndex == 1 && Global.firstLineHeight > 2)
                    {
                        textRange.ParagraphFormat.LineRuleWithin = MsoTriState.msoFalse;
                        textRange.ParagraphFormat.SpaceWithin = Global.firstLineHeight;
                        return;
                    }
                    else if (currentLineIndex == shape.TextFrame.TextRange.Lines().Count && Global.lastLineHeight > 2)
                    {
                        textRange.ParagraphFormat.LineRuleWithin = MsoTriState.msoFalse;
                        textRange.ParagraphFormat.SpaceWithin = Global.lastLineHeight;
                        return;
                    }
                    else if (Global.middleLineHeight > -1)
                    {
                        textRange.ParagraphFormat.LineRuleWithin = MsoTriState.msoFalse;
                        textRange.ParagraphFormat.SpaceWithin = Global.middleLineHeight;
                        return;
                    }
                }
                textRange.ParagraphFormat.LineRuleWithin = MsoTriState.msoFalse;
                textRange.ParagraphFormat.SpaceWithin = aLineHeight;
                int retryCount = 999;
                // @todo：
                // 单元格内的兼容性处理和普通文本有些差异，
                // 这里暂时没有搞清楚原理，
                // 姑且绕开这个问题，以后有时间再看看。
                if (containerShape.HasTable == MsoTriState.msoTrue)
                {
                    retryCount = -1;
                }
                do
                {
                    textRange.ParagraphFormat.SpaceWithin = textRange.ParagraphFormat.SpaceWithin + 1;
                    retryCount -= 1;
                } while (textRange.BoundHeight < oldBoundHeight && retryCount > 1);
                if (textRange.BoundHeight > oldBoundHeight)
                {
                    textRange.ParagraphFormat.SpaceWithin = textRange.ParagraphFormat.SpaceWithin - 1;
                }
                // @tips：记录一下不同情况下的行高，当做缓存用来提高代码运行的效率。
                if (Utils.CheckHasLargeFormula(textRange) == -1) {
                    if (shape.TextFrame.TextRange.Lines().Count == 1 && Global.singleLineHeight == -1)
                    {
                        Global.singleLineHeight = textRange.ParagraphFormat.SpaceWithin;
                    }
                    else
                    {
                        if (currentLineIndex == 1)
                        {
                            if (Global.firstLineHeight == -1)
                            {
                                Global.firstLineHeight = textRange.ParagraphFormat.SpaceWithin;
                            }
                        }
                        else if (currentLineIndex == shape.TextFrame.TextRange.Lines().Count)
                        {
                            if (Global.lastLineHeight == -1)
                            {
                                Global.lastLineHeight = textRange.ParagraphFormat.SpaceWithin;
                            }
                        }
                        else if (Global.middleLineHeight == -1)
                        {
                            Global.middleLineHeight = textRange.ParagraphFormat.SpaceWithin;
                        }
                    }
                }
            }
        }

        static public int HandleBlankPosition(Match match, TextRange line, TextRange2 line2, Shape shape, Shape containerShape, int offset)
        {
            string index = Regex.Match(match.Value, @"\d+").Value;
            Shape answerShape = null;
            int matchIndex = match.Index + offset;
            if (Global.GlobalAnswerMarkIndexMap.ContainsKey(index))
            {
                answerShape = Global.GlobalAnswerMarkIndexMap[index];
            }
            if (answerShape == null)
            {
                return 0;
            }
            // @tips：
            // 计算前先匹配一下答案位置，因为公式不一定会在哪里产生换行，故而需要匹配位置之后才能准确计算出宽度，
            // 并且，匹配答案位置会给折行的答案前面加空格，需要记录一下添加的空格数量，用来计算答案宽度。
            PositionMatchHandler.MatchAnswerPosition(answerShape, shape, containerShape);
            // 判断填空还是选择，下横线还是空格
            string spaceChar = "_";
            TextRange spaceRange = line.Characters(matchIndex + match.Value.Length + 1);
            if (spaceRange.Text == " ")
            {
                spaceChar = " ";
            }
            else if (spaceRange.Text == "_")
            {
                spaceChar = "_";
            }
            // 计算答案的宽度
            double answerWidth = 0;
            if (answerShape.TextFrame.TextRange.Lines().Count == 1)
            { // 单行答案
                answerWidth = Utils.ComputeAnswerFirstLineWidth(answerShape, 2);
            }
            else
            { // 多行答案
                for (int l = 1; l <= answerShape.TextFrame.TextRange.Lines().Count; l++)
                {
                    if (l == 1)
                    {
                        answerWidth += Utils.ComputeAnswerFirstLineWidth(answerShape, 1);
                    }
                    else if (l == answerShape.TextFrame.TextRange.Lines().Count)
                    {
                        answerWidth += answerShape.TextFrame.TextRange.Lines(l).BoundLeft +
                            answerShape.TextFrame.TextRange.Lines(l).BoundWidth -
                            shape.Left;
                        if (containerShape.HasTable == MsoTriState.msoTrue)
                        {
                            answerWidth += containerShape.Left;
                        }
                    }
                    else
                    {
                        answerWidth += shape.Width;
                    }
                }
            }
            // 找到作答空间 Range 部分
            TextRange matchRange = shape.TextFrame.TextRange.Find(match.Value);
            int rangeStart = matchRange.Start + match.Value.Length;
            int rangeEnd = rangeStart;
            while (rangeEnd <= shape.TextFrame.TextRange.Length)
            {
                if (shape.TextFrame.TextRange.Characters(rangeEnd).Text != spaceChar)
                {
                    rangeEnd--;
                    if (shape.TextFrame.TextRange.Characters(rangeEnd + 1).Text == "@")
                    { // 需要注意多个作答空间连续的情况 _@1@____@2@___
                        rangeEnd++;
                    }
                    break;
                }
                rangeEnd++;
            }
            // 替换
            spaceRange.Font.Name = "SimSun";
            float spaceWidth = spaceRange.BoundWidth;
            int spaceCount = (int)Math.Round(answerWidth / spaceWidth + 0.5);
            if (spaceChar == " " && spaceRange.Font.Underline == MsoTriState.msoFalse)
            { // 选择题、空格的作答空间，多加 2 个空格，看起来好看一些
                spaceCount += 2;
            }
            if (containerShape.Name.Contains("QC") && Regex.IsMatch(answerShape.TextFrame.TextRange.Text, @"^\s*[ABCD]\s*$"))
            { // 全局的选择题答案的宽度，保持一致
                if (Global.qcSpaceCount == -1)
                {
                    Global.qcSpaceCount = spaceCount;
                }
                else
                {
                    spaceCount = Global.qcSpaceCount;
                }
            }
            int i = 1;
            string spaces = spaceChar;
            while (i < spaceCount)
            {
                spaces += spaceChar;
                i++;
            }
            // @todo：需要单独处理带题号的完形填空，e.g. "_1@1@____"。
            MsoTriState hasUnderLine = spaceRange.Font.Underline;
            shape.TextFrame.TextRange.Characters(rangeStart, rangeEnd + 1 - rangeStart).Text = spaces;
            shape.TextFrame.TextRange.Characters(rangeStart, spaces.Length).Font.Underline = hasUnderLine;
            shape.TextFrame.TextRange.Characters(rangeStart, spaces.Length).Font.Name = "SimSun";
            shape.TextFrame2.TextRange.Characters[rangeStart, spaces.Length].Font.Spacing = 0;
            offset = spaces.Length - (rangeEnd + 1 - rangeStart);
            return offset;
        }

        static public int HandleImagePosition(Match match, TextRange line, TextRange2 line2, Shape shape, int offset)
        {
            string index = Regex.Match(match.Value, @"\d+").Value;
            Shape imageShape = null;
            Shape imageTipShape = null;
            float shapeFontSize = (float)Utils.GetShapeTextInfo(shape)[0];
            int matchIndex = match.Index + offset;
            if (match.Value.Contains("&") && Global.GlobalInlineImageMap.ContainsKey(index))
            {
                imageShape = Global.GlobalInlineImageMap[index];
            }
            else if (match.Value.Contains("%") && Global.GlobalTableImageMap.ContainsKey(index))
            {
                imageShape = Global.GlobalTableImageMap[index];
            }
            if (imageShape == null)
            {
                return 0;
            }
            if (Regex.IsMatch(imageShape.Name, @"imagetipindex=(\d+)"))
            {
                string imageTipIndex = Regex.Match(imageShape.Name, @"imagetipindex=(\d+)").Groups[1].Value;
                if (Global.GlobalImageTipMap.ContainsKey(imageTipIndex))
                {
                    imageTipShape = Global.GlobalImageTipMap[imageTipIndex][0];
                }
            }
            float targetHeight = imageShape.Height;
            float targetWidth = imageShape.Width;
            if (imageTipShape != null)
            {
                // @todo：若图说宽度大于图片，则先设置一下宽度。
                targetHeight += imageTipShape.TextFrame.TextRange.BoundHeight + 10;
            }
            // 插空格
            line.Characters(matchIndex + match.Value.Length).InsertAfter(" ");
            TextRange spaceRange = line.Characters(matchIndex + match.Value.Length + 1);
            TextRange2 spaceRange2 = line2.Characters[matchIndex + match.Value.Length + 1];
            spaceRange2.Font.Spacing = -shapeFontSize;
            // 设置字号（高度
            if (spaceRange.BoundHeight < targetHeight)
            {
                float leftHeight = 0;
                float leftFontSize = 1;
                float rightHeight = 999;
                float rightFontSize = 1;
                float targetFontSize = 7;
                double j = shapeFontSize;
                while (j < 999)
                {
                    spaceRange.Font.Size = (float)j;
                    if (spaceRange.BoundHeight == targetHeight)
                    {
                        targetFontSize = (float)j;
                        break;
                    }
                    if (spaceRange.BoundHeight < targetHeight && spaceRange.BoundHeight > leftHeight)
                    {
                        leftHeight = spaceRange.BoundHeight;
                        leftFontSize = (float)j;
                    }
                    if (spaceRange.BoundHeight > targetHeight && spaceRange.BoundHeight < rightHeight)
                    {
                        rightHeight = spaceRange.BoundHeight;
                        rightFontSize = (float)j;
                        break;
                    }
                    j += 0.5;
                }
                if (leftFontSize > 0 && rightHeight > 0 && rightHeight < 999)
                {
                    if (Math.Abs(targetHeight - leftHeight) > Math.Abs(targetHeight - rightHeight))
                    {
                        targetFontSize = rightFontSize;
                    }
                    else
                    {
                        targetFontSize = leftFontSize;
                    }
                }
                else if (targetFontSize < shapeFontSize)
                {
                    targetFontSize = shapeFontSize;
                }
                spaceRange.Font.Size = targetFontSize;
                spaceRange2.Font.Spacing = -targetFontSize;
            }
            // 设置字间距（宽度
            if (spaceRange.BoundWidth < targetWidth)
            {
                float leftWidth = 0;
                float leftSpacing = spaceRange2.Font.Spacing;
                float rightWidth = 999;
                float rightSpacing = 999;
                float targetSpacing = spaceRange2.Font.Spacing;
                float j = spaceRange2.Font.Spacing;
                while (j < 999)
                {
                    spaceRange2.Font.Spacing = j;
                    if (spaceRange.BoundWidth == targetWidth)
                    {
                        targetSpacing = j;
                        break;
                    }
                    if (spaceRange.BoundWidth < targetWidth && spaceRange.BoundWidth > leftWidth)
                    {
                        leftWidth = spaceRange.BoundWidth;
                        leftSpacing = j;
                    }
                    if (spaceRange.BoundWidth > targetWidth && spaceRange.BoundWidth < rightWidth)
                    {
                        rightWidth = spaceRange.BoundWidth;
                        rightSpacing = j;
                        break;
                    }
                    j += 1;
                }
                if (leftWidth > 0 && rightWidth > 0 && rightWidth < 999)
                {
                    if (Math.Abs(targetWidth - leftWidth) > Math.Abs(targetWidth - rightWidth))
                    {
                        targetSpacing = rightSpacing;
                    }
                    else
                    {
                        targetSpacing = leftSpacing;
                    }
                }
                spaceRange2.Font.Spacing = targetSpacing;
            }
            return 1;
        }

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
                        imageShape.Top - shape.Top < Global.gapBetweenTextLine[2] &&
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
                bool noneedProcess = false;
                if (((List<Shape>)nodes[n][0])[0].Name.Substring(0, 2) == "C_")
                {
                    noneedProcess = true;
                }
                if (!noneedProcess)
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
                                // @tips：若找到选项在版心右侧，则说明是横向排列的多个选项。
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
                                    offset += (float)ComputeTextLineHeight(currentOption, prevOption);
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
                                            Offset += (float)ComputeTextLineHeight(blocks[b][s - 1], blocks[b][s]);
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
                                // @tips：prevShape 需要取逻辑节点中 Bottom 最大的元素。
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
                                Offset += (float)ComputeTextLineHeight(currentShape, prevShape);
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
                                    ((PreNodeShape.Name.StartsWith("C") && PreNodeShape.Top > prevNodeTop) ||
                                        (!PreNodeShape.Name.StartsWith("C")))
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
                                    ((PreNodeShape.Name.StartsWith("C") && PreNodeShape.Top > prevNodeTop) ||
                                        (!PreNodeShape.Name.StartsWith("C")))
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
                        Offset2 += ComputeTextLineHeight(currentNodeTop, prevNodeBottom);
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
        static public double ComputeTextLineHeight(Shape targetShape, Shape moveBaseShape)
        {
            double offset = 0;
            // @wip：若相邻的两个文本框都是文字，则进行移动时还需要考虑行距
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
            // - 若当前文本框的首行是纯文本，上一个文本框的末行是公式，则位移 gapBetweenTextLine 距离
            // - 若当前文本框的首行是公式，上一个文本框的末行是公式，则不进行位移
            // = 若首行和末行的高度差不多，则位移 gapBetweenTextLine x 距离
            // - 若当前文本是段落，上一个文本是标题，则位移 gapBetweenTextLine x 距离
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
                if (!regex.IsMatch(firstLine.Text) && regex.IsMatch(lastLine.Text) && Global.gapBetweenTextLine[1] > 0)
                {
                    offset = Global.gapBetweenTextLine[1];
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
                    offset = Global.gapBetweenTextLine[1];
                }
                if ((!targetShape.Name.StartsWith("C_") && moveBaseShape.Name.StartsWith("C_")) ||
                    (targetShape.Name.StartsWith("C_") && moveBaseShape.Name.StartsWith("C_")))
                {
                    double move = Global.gapBetweenTextLine[1];
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
                bool hasMatchanswer = Utils.CheckMatchPositionAnswer(shape);
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
                // @tips：若当前节点是小题、并且进行过大小题图文布局，则（和没有进行图文布局的部分）当作独立的节点处理。
                if (!hasCollected && (shape.Name.Contains("AS") || shape.Name.Contains("EX") || shape.Name.Contains("AN")))
                {
                    // - 展开的，但是题干是非展开的，拆
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
                        else if (map.ContainsKey(shapeNodeId + "#" + "UNLAYOUTEDANASEX"))
                        {
                            shapeNodeId = shapeNodeId + "#" + "UNLAYOUTEDANASEX";
                            hasCollected = true;
                        }
                    }
                }
                // @tips：若答案/解析/解题思路进行过图文布局，则当作独立的节点处理。
                // @todo：若和题干同侧，则这里也可以不拆。
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
                        generateNodeShapes.Add(new object[2] { tc, shapeNodeId }); // 这里把节点的一些语义信息也带上
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

    internal class PositionMatchHandler
    {
        static public void MatchAnswerPosition(Shape answerShape, Shape blankShape = null, Shape containerShape = null)
        {
            string answerIndex = Regex.Match(answerShape.Name, "vbapositionanswer=(\\d+)").Groups[1].Value;
            if (blankShape == null || containerShape == null)
            {
                Shape[] blankShapes = Utils.FindBlankWithMarkIndex(answerIndex, answerShape);
                blankShape = blankShapes[0];
                containerShape = blankShapes[1];
            }
            if (blankShape == null)
            {
                return;
            }
            Match match = Regex.Match(blankShape.TextFrame.TextRange.Text, "@" + answerIndex + "@");
            ExecuteAnswerPosition(
                answerShape,
                blankShape,
                blankShape.TextFrame.TextRange.Find(match.Value),
                containerShape);
        }

        static public void ExecuteAnswerPosition(Shape answerShape, Shape blankShape, TextRange blankMatchRange, Shape containerShape)
        {
            Regex regExp = new Regex("vbscript.regexp");
            // @tips：这里需要注意前面的处理可能造成某些值的变化！！！
            float matchRangeTop = blankMatchRange.BoundTop;
            float matchRangeLeft = blankMatchRange.BoundLeft;
            float matchRangeHeight = blankMatchRange.BoundHeight;
            float blankShapeWidth = blankShape.Width;
            float positionLeft = blankMatchRange.BoundLeft - blankShape.Left;
            float blankFontSize = (float)Utils.GetShapeTextInfo(blankShape)[0];
            if (containerShape.HasTable == MsoTriState.msoTrue)
            {
                matchRangeTop += containerShape.Top;
                matchRangeLeft += containerShape.Left;
                positionLeft = blankMatchRange.BoundLeft - blankShape.TextFrame.TextRange.BoundLeft;
            }
            if (answerShape.HasTextFrame == MsoTriState.msoTrue)
            {
                if (containerShape.HasTable == MsoTriState.msoTrue) // Cell 的宽度、需要考虑两边的留白
                {
                    blankShapeWidth -= 10;
                }
                bool hasInsertSpaceBefore = false;
                string shapeLabel = "";
                if (Regex.IsMatch(answerShape.Name, "([^\\.]+)(\\.[^\\.]+)?#([\\d\\w]+)(\\.\\w+)?"))
                {
                    shapeLabel = Regex.Match(answerShape.Name, "([^\\.]+)(\\.[^\\.]+)?#([\\d\\w]+)(\\.\\w+)?").Groups[4].Value;
                }
                // @tips：清理支持折行时在前面添加的空格。
                // @tips：清理支持折行+溢出行内图换行导致的".\\s+.\\s?"部分。
                int c;
                for (c = 1; c <= answerShape.TextFrame.TextRange.Length; c++)
                {
                    if (answerShape.TextFrame.TextRange.Characters(c).Text != " ")
                    {
                        break;
                    }
                }
                if (c > 1)
                {
                    answerShape.TextFrame.TextRange.Characters(1, c - 1).Delete();
                }
                if (Regex.IsMatch(answerShape.TextFrame.TextRange.Text, "^\\.\\s+\\.\\s+"))
                {
                    Match match = Regex.Match(answerShape.TextFrame.TextRange.Text, "^\\.\\s+\\.\\s+");
                    answerShape.TextFrame.TextRange.Characters(1, match.Length).Delete();
                }
                answerShape.TextFrame.MarginBottom = 0;
                answerShape.TextFrame.MarginRight = 0;
                answerShape.TextFrame.MarginLeft = 0;
                answerShape.TextFrame.MarginTop = 0;
                answerShape.TextFrame.WordWrap = MsoTriState.msoFalse;
                answerShape.Width = answerShape.TextFrame.TextRange.BoundWidth;
                answerShape.TextFrame.Ruler.Levels[1].FirstMargin = 0;
                answerShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                float answerFontSize = (float)Utils.GetShapeTextInfo(answerShape)[0];
                int blankStart = 0;
                int blankEnd = 0;
                int line1 = -1;
                int line2 = -1;
                // 找到整个作答空间
                if ((shapeLabel == ".blank" || shapeLabel == ".bracket") &&
                    answerShape.TextFrame.TextRange.Lines().Count == 1 &&
                    !hasInsertSpaceBefore)
                {
                    TextRange matchRange = blankMatchRange;
                    int matchStart = matchRange.Start;
                    int matchEnd = matchStart + matchRange.Length - 1;
                    blankStart = 0;
                    blankEnd = 0;
                    int i = matchStart - 1;
                    while (i > 0)
                    {
                        if (shapeLabel == ".blank")
                        { // 找到左边的首个_
                            if (blankShape.TextFrame.TextRange.Characters(i).Text != "_" &&
                                !(blankShape.TextFrame.TextRange.Characters(i).Text == " " &&
                                blankShape.TextFrame.TextRange.Characters(i).Font.Underline == MsoTriState.msoTrue))
                            {
                                blankStart = i + 1;
                                break;
                            }
                            else if (blankShape.TextFrame.TextRange.Characters(i).Text == "_" &&
                                blankShape.TextFrame.TextRange.Characters(i).Font.Color.RGB == ColorTranslator.ToOle(Color.FromArgb(255, 255, 255)))
                            { // 需要注意避免误伤表格图片的占位下横线
                                blankStart = i + 1;
                                break;
                            }
                        }
                        if (shapeLabel == ".bracket")
                        { // 找到左边的（括号
                            if (blankShape.TextFrame.TextRange.Characters(i).Text != " ")
                            {
                                blankStart = i;
                                break;
                            }
                        }
                        i--;
                    }
                    if (i == 0)
                    { // 需要考虑 _ 在句首的情况
                        blankStart = 1;
                    }
                    if (i > 1)
                    { // 需要考虑下横线上连续存在多个填空的情况
                        if (blankShape.TextFrame.TextRange.Characters(i).Text == "@" &&
                            Regex.IsMatch(blankShape.TextFrame.TextRange.Characters(i - 1).Text, @"\d+"))
                        {
                            blankStart = matchStart - 1;
                        }
                    }
                    i = matchEnd + 1;
                    while (i <= blankShape.TextFrame.TextRange.Length)
                    {
                        if (blankShape.TextFrame.TextRange.Characters(i).Text == " " &&
                            blankShape.TextFrame.TextRange.Characters(i).Font.Size > (float)Utils.GetShapeTextInfo(blankShape)[0])
                        { // 跳过撑高空格
                        }
                        else if (blankShape.TextFrame.TextRange.Characters(i).Text == "\r")
                        { // 跳过 WPS 兼容用的换行符
                        }
                        else if (shapeLabel == ".blank")
                        { // 找到右边的最后一个 _，可能是下横线、也可能是空格+下划线
                            if (blankShape.TextFrame.TextRange.Characters(i).Text != "_" &&
                                !(blankShape.TextFrame.TextRange.Characters(i).Text == " " &&
                                    blankShape.TextFrame.TextRange.Characters(i).Font.Underline == MsoTriState.msoTrue))
                            {
                                blankEnd = i - 1;
                                if (blankShape.TextFrame.TextRange.Characters(i - 1).Text == "\r")
                                {
                                    blankEnd = i - 2;
                                }
                                break;
                            }
                        }
                        else if (shapeLabel == ".bracket")
                        { // 找到右边的（括号
                            if (blankShape.TextFrame.TextRange.Characters(i).Text != " ")
                            {
                                blankEnd = i;
                                break;
                            }
                        }
                        i++;
                    }
                    if (i == blankShape.TextFrame.TextRange.Length + 1)
                    {
                        blankEnd = blankShape.TextFrame.TextRange.Length;
                    }
                    line1 = Utils.FindLineNum(blankStart, blankShape, containerShape);
                    line2 = Utils.FindLineNum(blankEnd, blankShape, containerShape);
                    // 若答案放置后内容溢出 Shape 宽度，则需要对答案进行换行处理
                    // - 目前这里考虑如果标记距离元素的左侧很近，目前接受 1 个字宽，则不进行文首缩进
                    // - 若作答空间本身只有 1 行，则不需要对答案进行折行
                    // - 若答案和作答空间均只有 1 行，并且其宽度小于作答空间的宽度，则不需要对答案进行折行
                    bool needWrapped = false;
                    if (answerShape.Width + positionLeft > blankShapeWidth && blankShape.TextFrame.TextRange.Lines().Count > 1)
                    {
                        needWrapped = true;
                    }
                    if (needWrapped && blankStart >= 0 && blankEnd > 0)
                    {
                        if (line1 == line2 && blankShape.TextFrame.TextRange.Characters(blankStart, blankEnd - blankStart + 1).BoundWidth > answerShape.Width)
                        {
                            needWrapped = false;
                        }
                    }
                    if (needWrapped)
                    {
                        // @tips：
                        // 在 PPT 里连续空格、连续下横线可能会出现超出元素边界的情况，
                        // 所以这里使用元素的宽度进行判断，
                        // 而不是使用 BoundWidth 进行判断，会大于元素宽度。
                        answerShape.TextFrame.WordWrap = MsoTriState.msoTrue;
                        answerShape.Width = blankShapeWidth;
                        answerShape.Height = answerShape.TextFrame.TextRange.BoundHeight;
                        answerShape.Left = matchRangeLeft - positionLeft;
                        // @tips：限制答案宽度不可超过版心。
                        float viewRight = (float)Utils.GetViewRight();
                        if (answerShape.Left + answerShape.Width > viewRight)
                        {
                            answerShape.Width = viewRight - answerShape.Left;
                        }
                        bool needIndent = true;
                        // @tips：
                        // 若作答空间只包含答案、并且居中，则不需要对答案进行行首的插空格缩进，
                        // 便于在下面的逻辑中对答案进行居中操作。
                        if (Regex.IsMatch(blankShape.TextFrame.TextRange.Text, "^.?[\\@_\\d\\s]+$") &&
                            blankShape.TextFrame.TextRange.ParagraphFormat.Alignment == PpParagraphAlignment.ppAlignCenter)
                        {
                            needIndent = false;
                        }
                        if (needIndent && positionLeft >= answerFontSize)
                        {
                            // @tips：
                            // 使用插空格的方式对齐折行答案和填空的位置，
                            // 替代首行缩进的方式。
                            answerShape.TextFrame2.TextRange.InsertBefore(" ");
                            TextRange2 spaceItem = answerShape.TextFrame2.TextRange.Characters[1];
                            // @tips：
                            // 这里使用答案字号，而不是更小的字号（字号越小越准确），
                            // 为了避免出现 1 行空格、1 行答案的情况，
                            // 行高和填空行高对不齐的情况。
                            // @todo：可以在单独的环节把上述情况专门处理掉。
                            spaceItem.Font.Size = answerFontSize;
                            spaceItem.Font.Name = "SimSun";
                            if (spaceItem.Font.Spacing < 0)
                            {
                                spaceItem.Font.Spacing = spaceItem.Font.Spacing + (-spaceItem.Font.Spacing);
                            }
                            double spaceWidth = spaceItem.BoundWidth;
                            int spaceCount = (int)Math.Round(positionLeft / spaceWidth + 0.5);
                            i = 1;
                            string spaceRange = " ";
                            while (i < spaceCount)
                            {
                                spaceRange = spaceRange + " ";
                                i++;
                            }
                            answerShape.TextFrame2.TextRange.Characters[1].Text = spaceRange;
                            hasInsertSpaceBefore = true;
                            // @tips：
                            // 使用插空格的方式，在某些西文的场景下，会空格单独成行、单词换行，
                            // 导致看起来答案没有对准，
                            // 这种情况下，改成用行首缩进来实现。
                            if (answerShape.TextFrame2.TextRange.Lines.Count > (line2 + 1 - line1))
                            {
                                for (c = 1; c <= answerShape.TextFrame2.TextRange.Length; c++)
                                {
                                    if (answerShape.TextFrame2.TextRange.Characters[c].Text != " ")
                                    {
                                        break;
                                    }
                                }
                                if (c > 1)
                                {
                                    answerShape.TextFrame2.TextRange.Characters[1, c - 1].Delete();
                                }
                                hasInsertSpaceBefore = false;
                            }
                        }
                        if (!hasInsertSpaceBefore) // 若间隙比较小，则使用行首缩进
                        {
                            answerShape.TextFrame.Ruler.Levels[1].FirstMargin = positionLeft;
                        }
                    }
                    else
                    {
                        // @tips：若答案包含复杂公式，并且进行过图像识别，则认为可以调整为单倍行高。
                        if (answerShape.TextFrame.TextRange.ParagraphFormat.SpaceWithin < 2 && answerShape.Name.Contains("rh="))
                        {
                            answerShape.TextFrame.TextRange.ParagraphFormat.LineRuleWithin = MsoTriState.msoTrue;
                            answerShape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1;
                        }
                        answerShape.Width = answerShape.TextFrame.TextRange.BoundWidth;
                        answerShape.Height = answerShape.TextFrame.TextRange.BoundHeight;
                        answerShape.Left = matchRangeLeft;
                    }
                }
                // 动起来！！！
                // 顶部对齐
                answerShape.Top = matchRangeTop;
                // - 若答案折行后的结果是第一行为空，则把答案复原，移动到下一行的行首
                // - 若答案的字号比题干字号大，则在这里设置答案 Top 值时对齐作答空间底部
                if (answerShape.TextFrame.TextRange.Lines().Count == 1 && !hasInsertSpaceBefore && answerFontSize > blankFontSize)
                {
                    answerShape.Top = matchRangeTop + matchRangeHeight - answerShape.Height;
                }
                if (answerShape.TextFrame.TextRange.Lines().Count > 1 &&
                    hasInsertSpaceBefore &&
                    answerShape.TextFrame.TextRange.Lines(1).Text.Trim().Length == 0 &&
                    line1 < blankShape.TextFrame.TextRange.Lines().Count)
                {
                    for (c = 1; c <= answerShape.TextFrame.TextRange.Length; c++)
                    {
                        if (answerShape.TextFrame.TextRange.Characters(c).Text != " ")
                        {
                            break;
                        }
                    }
                    if (c > 1)
                    {
                        answerShape.TextFrame.TextRange.Characters(1, c - 1).Delete();
                    }
                    answerShape.Width = answerShape.TextFrame.TextRange.BoundWidth;
                    float targetTop = blankShape.TextFrame.TextRange.Lines(line1 + 1).BoundTop;
                    float targetLeft = blankShape.TextFrame.TextRange.Lines(line1 + 1).BoundLeft;
                    if (containerShape.HasTable == MsoTriState.msoTrue)
                    {
                        targetTop += containerShape.Top;
                        targetLeft += containerShape.Left;
                    }
                    answerShape.Top = targetTop;
                    answerShape.Left = targetLeft;
                }
                // 对填空的答案进行一些偏移，避免答案压在下横线上
                if (shapeLabel == ".blank")
                {
                    float currentTop = answerShape.Top;
                    // - @todo：对于包含行内图片的答案，观察一下再决定如何处理（在 HandleBlankLineHeight 环节有冗余加些高度
                    // - 对于纯文本的答案，不需要多进行移动，因为行高都是一样的
                    // - 对于单行的，包含 g、p、q、y 的答案，多往上移动一点点
                    // - DISABLED：对于包含公式的答案，多往上移动一点点
                    // - DISABLED：对于处理过行高的答案，多往上移动一点点
                    answerShape.Top -= 3;
                    if (answerShape.TextFrame.TextRange.Lines().Count == 1 &&
                        Regex.IsMatch(answerShape.TextFrame.TextRange.Text, "[gpqy]"))
                    {
                        answerShape.Top = answerShape.Top - 1;
                    }
                }
                // 对答案的位置进行一些偏移，使其可以在空里水平居中
                // @tips：
                // 这里还是需要对折行进行判断，
                // 会出现某些极端的情况，比如说表格里，答案是单行、被认为是折行、但是没有产生折行的情况。
                if ((shapeLabel == ".blank" || shapeLabel == ".bracket") &&
                    answerShape.TextFrame.TextRange.Lines().Count == 1 &&
                    !hasInsertSpaceBefore &&
                    blankStart > 0 && blankEnd > 0 &&
                    blankStart < blankEnd)
                {
                    float blankWidth = blankShape.TextFrame.TextRange.Characters(blankStart, blankEnd - blankStart + 1).BoundWidth;
                    float blankLeft = blankShape.TextFrame.TextRange.Characters(blankStart).BoundLeft;
                    if (containerShape.HasTable == MsoTriState.msoTrue)
                    {
                        blankLeft += containerShape.Left;
                    }
                    // - 需要注意这里可能存在答案没有换行、但是作答空间换行的情况
                    bool needProcess = true;
                    TextRange line = blankShape.TextFrame.TextRange.Lines(line1);
                    if (Regex.IsMatch(line.Text, @"^_@\d+@(_+)$"))
                    {
                        int Length = Regex.Match(line.Text, @"^_@\d+@(_+)$").Groups[0].Value.Length;
                        needProcess = !(Length > 10);
                    }
                    if (line1 == line2 && needProcess)
                    {
                        float newLeft = blankLeft + (float)Math.Round((blankWidth - answerShape.Width) / 2 + 0.5);
                        answerShape.Left = newLeft;
                    }
                }
                // 文本框样式设置为水平居中
                // - 若答案文本没有超过 1 行、却存在溢出，则居左对齐
                // - 若答案文本所在的填空是居中的，则居中对齐
                // - 若答案文本前面存在文首缩进空格，则居左对齐
                // - 若答案只有 1 行，则居中对齐
                if (answerShape.TextFrame.TextRange.Lines().Count == 1 &&
                    answerShape.TextFrame.Ruler.Levels[1].FirstMargin + answerShape.TextFrame.TextRange.BoundWidth > answerShape.Width)
                {
                    answerShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                }
                else if (
                    answerShape.TextFrame.TextRange.Lines().Count > 1 &&
                    blankShape.TextFrame.TextRange.ParagraphFormat.Alignment == PpParagraphAlignment.ppAlignCenter
                )
                {
                    answerShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                }
                else if (hasInsertSpaceBefore)
                {
                    answerShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                }
                else if (answerShape.TextFrame.TextRange.Lines().Count == 1)
                {
                    // @tips：一些特殊情况下，会触发 PPT 的幺蛾子机制导致结果很奇怪。
                    if (!Regex.IsMatch(answerShape.TextFrame.TextRange.Text, @"&\d+&\s+$"))
                    {
                        answerShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                    }
                }
            }
            else if (answerShape.Type == MsoShapeType.msoPicture)
            {
                answerShape.Left = matchRangeLeft;
                answerShape.Top = matchRangeTop;
            }
            // 跨页答案的剪切
            if (answerShape.Parent.SlideIndex != containerShape.Parent.SlideIndex)
            {
                answerShape.Cut();
                // @todo：
                // Utils.Sleep();
                Global.app.ActivePresentation.Slides[(int)containerShape.Parent.SlideIndex].Shapes.Paste();
            }
        }
    }
}
