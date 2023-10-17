using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Threading;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;

public struct VbaReturn
{
    public string ErrorInfo;
    public int PreviousSlidesCount;
    public int CurrentSlidesCount;
    public string FileSuffix;
    public string ErrorType; // 区分 Err 是生成异常 -1、还是机器质检报错 -2
}

namespace hexin_csharp
{
    public class Utils
    {
        static public double UnitConvert(double cm)
        {
            return cm * (480000 / 127) / (400 / 3);
        }

        static public List<Shape> GetSortedSlideShapes(Slide slide)
        {
            List<Shape> sortedShapes = new List<Shape>();
            foreach (Shape shape in slide.Shapes)
            {
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    // - JB 左括号形状不要丢掉
                    //  - 白块不要丢掉
                    if (shape.TextFrame.TextRange.Text.Length > 0 ||
                        shape.AutoShapeType == MsoAutoShapeType.msoShapeLeftBrace ||
                        shape.Name.Contains("WB"))
                    {
                        sortedShapes.Add(shape);
                    }
                }
                else
                {
                    sortedShapes.Add(shape);
                }
            }
            for (int i = 0; i < sortedShapes.Count - 1; i++)
            {
                for (int j = i + 1; j < sortedShapes.Count; j++)
                {
                    if (sortedShapes[i].Top > sortedShapes[j].Top)
                    {
                        Shape t = sortedShapes[j];
                        sortedShapes.RemoveAt(j);
                        sortedShapes.Insert(i, t);
                    }
                }
            }
            return sortedShapes;
        }

        // @description 相比于 GetSortedSlideShapes 函数，当前方法只获取静态的元素
        static public List<Shape> GetSortedStaticSlideShapes(Slide slide)
        {
            List<Shape> sortedShapes = new List<Shape>();
            for (int i = 1; i <= slide.Shapes.Count; i++)
            {
                if (
                    !CheckMatchPositionShape(slide.Shapes[i]) ||
                    CheckHasImageTip(slide.Shapes[i]))
                {
                    if (slide.Shapes[i].HasTextFrame == MsoTriState.msoTrue)
                    {
                        if (
                            slide.Shapes[i].TextFrame.TextRange.Text.Length > 0 ||
                            slide.Shapes[i].AutoShapeType == MsoAutoShapeType.msoShapeLeftBrace
                        )
                        {
                            sortedShapes.Add(slide.Shapes[i]);
                        }
                    }
                    else
                    {
                        sortedShapes.Add(slide.Shapes[i]);
                    }
                }
            }
            for (int i = 0; i < sortedShapes.Count - 1; i++)
            {
                for (int j = i + 1; j < sortedShapes.Count; j++)
                {
                    if (sortedShapes[i].Top > sortedShapes[j].Top)
                    {
                        (sortedShapes[i], sortedShapes[j]) = (sortedShapes[j], sortedShapes[i]);
                    }
                }
            }
            return sortedShapes;
        }

        // @description 获取某个 Shape 里的无标记文本 Runs
        // @todo：尚未兼容占位用的空格作为分隔符的情况。
        static public List<TextRange> GetShapeRuns(Shape shape)
        {
            List<TextRange> runs = new List<TextRange>();
            Regex regex = new Regex("(\\%\\d+\\%\\s*)|(&\\d+&\\s*)|" +
                "(\\@\\d+\\@)|(\\<m\\>)|(\\</m\\>)|(\\<l\\>)|(\\</l\\>)|(\\<zzd\\>)|(\\</zzd\\>)|" +
                "(\\<th\\>)|(\\</th\\>)|(\\<bc\\>)|(\\</bc\\>)|(\\.[\\d\\s\\$\\%]+\\.)|" +
                "(\\<ib\\>)|(\\</ib\\>)");
            if (shape.HasTable == MsoTriState.msoTrue)
            {
                Table table = shape.Table;
                int colCount = table.Columns.Count;
                int rowCount = table.Rows.Count;
                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        int textLength = table.Cell(i, j).Shape.TextFrame.TextRange.Length;
                        if (textLength == 0)
                        {
                            continue;
                        }
                        if (regex.IsMatch(table.Cell(i, j).Shape.TextFrame.TextRange.Text))
                        {
                            int prevMatchEnd = 0;
                            foreach (Match match in regex.Matches(table.Cell(i, j).Shape.TextFrame.TextRange.Text))
                            {
                                int matchRangeStart = match.Index + 1;
                                int matchRangeEnd = matchRangeStart + match.Length - 1;
                                int runRangeStart = prevMatchEnd + 1;
                                int runRangeEnd = matchRangeStart - 1;
                                if (runRangeEnd >= runRangeStart)
                                {
                                    int runRangeLength = runRangeEnd - runRangeStart + 1;
                                    runs.Add(table.Cell(i, j).Shape.TextFrame.TextRange.Characters(runRangeStart, runRangeLength));
                                }
                                int prevMatchStart = matchRangeStart;
                                prevMatchEnd = matchRangeEnd - 1;
                            }
                            if (textLength > prevMatchEnd)
                            {
                                int runRangeStart = prevMatchEnd + 1;
                                int runRangeEnd = textLength;
                                int runRangeLength = runRangeEnd - runRangeStart + 1;
                                runs.Add(table.Cell(i, j).Shape.TextFrame.TextRange.Characters(runRangeStart, runRangeLength));
                            }
                        }
                        else
                        {
                            runs.Add(table.Cell(i, j).Shape.TextFrame.TextRange);
                        }
                    }
                }
            }
            else if (shape.HasTextFrame == MsoTriState.msoTrue)
            {
                if (regex.IsMatch(shape.TextFrame.TextRange.Text))
                {
                    int prevMatchStart;
                    int prevMatchEnd = 0;
                    foreach (Match match in regex.Matches(shape.TextFrame.TextRange.Text))
                    {
                        int matchRangeStart = match.Index + 1;
                        int matchRangeEnd = matchRangeStart + match.Length - 1;
                        int runRangeStart = prevMatchEnd + 1;
                        int runRangeEnd = matchRangeStart - 1;
                        if (runRangeEnd >= runRangeStart)
                        {
                            int runRangeLength = runRangeEnd - runRangeStart + 1;
                            runs.Add(shape.TextFrame.TextRange.Characters(runRangeStart, runRangeLength));
                        }
                        prevMatchStart = matchRangeStart;
                        prevMatchEnd = matchRangeEnd - 1;
                    }
                    if (shape.TextFrame.TextRange.Length > prevMatchEnd)
                    {
                        int runRangeStart = prevMatchEnd + 1;
                        int runRangeEnd = shape.TextFrame.TextRange.Length;
                        int runRangeLength = runRangeEnd - runRangeStart + 1;
                        runs.Add(shape.TextFrame.TextRange.Characters(runRangeStart, runRangeLength));
                    }
                }
                else
                {
                    runs.Add(shape.TextFrame.TextRange);
                }
            }
            return runs;
        }

        static public string[] GetShapeInfo(Shape shape)
        {
            Regex RegExp = new Regex(@"([a-zA-Z]+_[^\.\#]+)(\.[^\.\#]+)?#([\d\w]+)(\.\w+)?(\?[^\?]+)?");
            string[] result = new string[] { "-1", "-1", "-1", "-1", "-1", "-1", "-1" };
            try
            {
                if (!RegExp.IsMatch(shape.Name))
                {
                    return result;
                }
                Match Match = RegExp.Match(shape.Name);
                string ShapeType = Match.Groups[1].Value;
                if (ShapeType.StartsWith("Q") &&
                    !ShapeType.Contains("AN") &&
                    !ShapeType.Contains("AS") &&
                    !ShapeType.Contains("EX") &&
                    !ShapeType.EndsWith("BD"))
                {
                    ShapeType += "_BD";
                }
                string ShapeProp = "BD";
                if (ShapeType.Contains("AN"))
                {
                    ShapeProp = "AN";
                }
                else if (ShapeType.Contains("AS"))
                {
                    ShapeProp = "AS";
                }
                else if (ShapeType.Contains("EX"))
                {
                    ShapeProp = "EX";
                }
                else if (ShapeType.EndsWith("BD"))
                {
                    ShapeProp = "BD";
                }
                string ShapeNodeId = Match.Groups[3].Value;
                string ShapeLabel = Match.Groups[4].Value;
                string ShapeKey = ShapeType + ":" + ShapeNodeId;
                string ShapeParentNodeId = "-1";
                RegExp = new Regex(@"parentnodeid=([\d\w]+)");
                if (RegExp.IsMatch(shape.Name))
                {
                    ShapeParentNodeId = RegExp.Match(shape.Name).Groups[1].Value;
                }
                string ShapeAnimationIndex = Match.Groups[2].Value;
                result = new string[] {
                    ShapeNodeId,
                    ShapeKey,
                    ShapeLabel,
                    ShapeType,
                    ShapeParentNodeId,
                    ShapeAnimationIndex,
                    ShapeProp,
                };
                return result;
            }
            catch
            {
                return result;
            }
        }

        static public double GetTableColumnWidth(Column column)
        {
            try
            {
                return column.Width;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return 99999;
            }
        }

        // @description 获取表格 n 行的高度
        static public double GetTableRowsHeight(int rowIndex, Shape shape)
        {
            if (rowIndex <= 0)
            {
                return 99999;
            }
            else
            {
                double height = 0;
                while (rowIndex > 0)
                {
                    height += GetTableRowHeight(shape.Table.Rows[rowIndex], shape);
                    rowIndex--;
                }
                return height;
            }
        }

        static public double GetTableRowHeight(Row row, Shape tableShape)
        {
            try
            {
                return row.Height;
            }
            catch
            {
                double rowHeight = 99999;
                foreach (Cell cell in row.Cells)
                {
                    if (rowHeight > cell.Shape.Height && !CheckMergedCell(cell, tableShape))
                    {
                        rowHeight = cell.Shape.Height;
                    }
                }
                if (rowHeight < 99999)
                {
                    return rowHeight;
                }
                else
                {
                    return 1;
                }
            }
        }

        static public object[] GetShapeTextInfo(Shape shape)
        {
            float fontSize = 1;
            string fontName = "";
            Color fontColor = Color.FromArgb(255, 255, 255);
            float lineHeight = 1;
            string nameFarEast = "";
            List<TextRange> runs = GetShapeRuns(shape);
            for (int r = 0; r < runs.Count; r++)
            {
                for (int c = 0; c < runs[r].Length; c++)
                {
                    // 需要跳过的字符：
                    // - 标记
                    // - 占位空格、撑高空格
                    // - 占位下横线
                    TextRange character = runs[r].Characters(c);
                    bool caniuse = true;
                    if (character.Font.Size < 10)
                    {
                        caniuse = false;
                    }
                    if (character.Text == " ")
                    {
                        caniuse = false;
                    }
                    if (character.Text == "_")
                    {
                        caniuse = false;
                    }
                    if (caniuse && character.Font.Size > fontSize)
                    {
                        // @tips：字号这里需要注意，可能存在单独给题号设置字号的情况，往往比正文要大一些。
                        fontSize = character.Font.Size;
                        fontName = character.Font.Name;
                        fontColor = ColorTranslator.FromOle(character.Font.Color.RGB);
                        lineHeight = character.ParagraphFormat.SpaceWithin;
                        nameFarEast = character.Font.NameFarEast;
                        break;
                    }
                }
                if (fontSize > 1)
                {
                    break;
                }
            }
            if (fontSize <= 1)
            {
                fontSize = (float)Global.gapBetweenTextLine[2];
            }
            return new object[] {
                fontSize,
                fontName,
                fontColor,
                lineHeight,
                nameFarEast
            };
        }

        public static object GetRangeTextInfo(TextRange Range)
        {
            float fontSize = 1;
            string fontName = "";
            Color fontColor = Color.FromArgb(255, 255, 255);
            float lineHeight = 1;
            for (int c = 1; c <= Range.Length; c++)
            {
                TextRange Characters = Range.Characters(c);
                if (Characters.Text != " " && Characters.Font.Size > fontSize)
                {
                    fontSize = Characters.Font.Size;
                    fontName = Characters.Font.Name;
                    fontColor = ColorTranslator.FromOle(Characters.Font.Color.RGB);
                    lineHeight = Characters.ParagraphFormat.SpaceWithin;
                    break;
                }
            }
            // @tips：对于只包含占位空格的文本，字号读出来可能是不符合预期的，兜底一下。
            if (fontSize < 10 && Global.gapBetweenTextLine[2] > 0)
            {
                fontSize = (float)Global.gapBetweenTextLine[2];
            }
            return new object[] {
                fontSize,
                fontName,
                fontColor,
                lineHeight
            };
        }

        static public Dictionary<string, object[]>[] GetBlankAndAnswerMap(List<Slide> slides)
        {
            Regex regExp = new Regex(@"([^\.]+)(\.[^\.]+)?#([\d\w]+)(\.\w+)?");
            Dictionary<string, object[]> blankMap = new Dictionary<string, object[]>();
            Dictionary<string, object[]> answerMap = new Dictionary<string, object[]>();
            foreach (Slide slide in slides)
            {
                bool hasTable = false;
                foreach (Shape shape in slide.Shapes)
                {
                    if (shape.HasTable == MsoTriState.msoTrue)
                    {
                        hasTable = true;
                        break;
                    }
                }
                if (hasTable)
                {
                    // @tips：获取表格尺寸信息时需要 Focus 在所属页面。
                    Global.app.ActiveWindow.View.GotoSlide(slide.SlideIndex);
                    // @tips：加一个 1s 的休眠，避免执行太快表格获取尺寸信息不准确。
                    Sleep();
                }
                foreach (Shape shape in slide.Shapes)
                {
                    Match match = regExp.Match(shape.Name);
                    string shapeNodeId = "";
                    if (match.Success)
                    {
                        shapeNodeId = match.Groups[3].Value;
                    }
                    match = Regex.Match(shape.Name, "answermarkindex=(\\d+)");
                    if (shape.Type == MsoShapeType.msoTextBox && match.Success)
                    {
                        string answerMarkIndex = match.Groups[1].Value;
                        string shapeKey = shapeNodeId + ":" + answerMarkIndex;
                        if (!answerMap.ContainsKey(shapeKey))
                        {
                            answerMap.Add(shapeKey, new object[] { shape, answerMarkIndex });
                        }
                    }
                    if (shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        match = Regex.Match(shape.TextFrame.TextRange.Text, "@(\\d+)@");
                        if (match.Success)
                        {
                            foreach (Match m in Regex.Matches(shape.TextFrame.TextRange.Text, "@(\\d+)@"))
                            {
                                if (shape.TextFrame.TextRange.Find(m.Value) != null)
                                {
                                    string shapeKey = shapeNodeId + ":" + m.Groups[1].Value;
                                    if (!blankMap.ContainsKey(shapeKey))
                                    {
                                        shape.Height = (float)ComputeShapeHeight(shape);
                                        if (shape.TextFrame.TextRange.ParagraphFormat.LineRuleWithin == MsoTriState.msoTrue)
                                        {
                                            blankMap.Add(shapeKey, new object[] {
                                                slide.SlideIndex,
                                                shape.TextFrame.TextRange.Find(m.Value).BoundTop,
                                                shape.TextFrame.TextRange.Find(m.Value).BoundLeft,
                                                shape.TextFrame.TextRange.Find(m.Value).ParagraphFormat.SpaceWithin * shape.TextFrame.TextRange.Font.Size,
                                                shape.Width,
                                                shape.TextFrame.TextRange.Find(m.Value).BoundLeft - shape.Left, shape, shape.TextFrame.TextRange.Find(m.Value),
                                                shape,
                                            });
                                        }
                                        else
                                        {
                                            blankMap.Add(shapeKey, new object[] {
                                                slide.SlideIndex,
                                                shape.TextFrame.TextRange.Find(m.Value).BoundTop,
                                                shape.TextFrame.TextRange.Find(m.Value).BoundLeft,
                                                shape.TextFrame.TextRange.Find(m.Value).ParagraphFormat.SpaceWithin,
                                                shape.Width,
                                                shape.TextFrame.TextRange.Find(m.Value).BoundLeft - shape.Left,
                                                shape,
                                                shape.TextFrame.TextRange.Find(m.Value),
                                                shape,
                                            });
                                        }
                                    }
                                }
                            }
                        }
                        match = Regex.Match(shape.Name, "vbapositionanswer=(\\d+)");
                        if (match.Success)
                        {
                            foreach (Match m in Regex.Matches(shape.Name, "vbapositionanswer=(\\d+)"))
                            {
                                string shapeKey = shapeNodeId + ":" + m.Groups[1].Value;
                                if (!answerMap.ContainsKey(shapeKey) && shape.Name.Contains("AN"))
                                {
                                    answerMap.Add(shapeKey, new object[] { shape, m.Groups[1].Value });
                                }
                            }
                        }
                    }
                    else if (shape.HasTable == MsoTriState.msoTrue)
                    {
                        foreach (Row row in shape.Table.Rows)
                        {
                            foreach (Cell cell in row.Cells)
                            {
                                match = Regex.Match(cell.Shape.TextFrame.TextRange.Text, "@(\\d+)@");
                                if (match.Success)
                                {
                                    foreach (Match m in Regex.Matches(cell.Shape.TextFrame.TextRange.Text, "@(\\d+)@"))
                                    {
                                        string shapeKey = shapeNodeId + ":" + m.Groups[1].Value;
                                        if (!blankMap.ContainsKey(shapeKey))
                                        {
                                            if (cell.Shape.TextFrame.TextRange.ParagraphFormat.LineRuleWithin == MsoTriState.msoTrue)
                                            {
                                                blankMap.Add(shapeKey, new object[] {
                                                    slide.SlideIndex,
                                                    cell.Shape.TextFrame.TextRange.Find(m.Value).BoundTop + shape.Top,
                                                    cell.Shape.TextFrame.TextRange.Find(m.Value).BoundLeft + shape.Left,
                                                    cell.Shape.TextFrame.TextRange.Find(m.Value).ParagraphFormat.SpaceWithin * cell.Shape.TextFrame.TextRange.Font.Size,
                                                    cell.Shape.Width,
                                                    cell.Shape.TextFrame.TextRange.Find(m.Value).BoundLeft - cell.Shape.TextFrame.TextRange.BoundLeft,
                                                    cell.Shape,
                                                    cell.Shape.TextFrame.TextRange.Find(m.Value),
                                                    shape,
                                                });
                                            }
                                            else
                                            {
                                                blankMap.Add(shapeKey, new object[] {
                                                    slide.SlideIndex,
                                                    cell.Shape.TextFrame.TextRange.Find(m.Value).BoundTop + shape.Top,
                                                    cell.Shape.TextFrame.TextRange.Find(m.Value).BoundLeft + shape.Left,
                                                    cell.Shape.TextFrame.TextRange.Find(m.Value).ParagraphFormat.SpaceWithin,
                                                    cell.Shape.Width,
                                                    cell.Shape.TextFrame.TextRange.Find(m.Value).BoundLeft - cell.Shape.TextFrame.TextRange.BoundLeft,
                                                    cell.Shape,
                                                    cell.Shape.TextFrame.TextRange.Find(m.Value),
                                                    shape,
                                                });
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return new Dictionary<string, object[]>[] { blankMap, answerMap };
        }

        // @description 获取可视区最小 Shape.Left
        static public double GetViewLeft()
        {
            if (Global.pptViewLeft > 0)
            {
                return Global.pptViewLeft;
            }
            double getViewLeft = 99999;
            foreach (Slide slide in Global.app.ActivePresentation.Slides)
            {
                foreach (Shape shape in slide.Shapes)
                {
                    if (
                        !shape.Name.Substring(0, 2).Equals("C_") &&
                        !CheckMatchPositionShape(shape) &&
                        shape.Type != MsoShapeType.msoPlaceholder &&
                        shape.Left < getViewLeft
                    )
                    {
                        getViewLeft = shape.Left;
                    }
                }
            }
            if (getViewLeft > Global.slideWidth)
            {
                getViewLeft = Global.viewLeft;
            }
            Global.pptViewLeft = (float)getViewLeft;
            return getViewLeft;
        }

        // @description 获取可视区最大 Shape.Right
        static public double GetViewRight()
        {
            if (Global.pptViewRight > 0)
            {
                return Global.pptViewRight;
            }
            double getViewRight = 1;
            foreach (Slide slide in Global.app.ActivePresentation.Slides)
            {
                foreach (Shape shape in slide.Shapes)
                {
                    if (
                        !shape.Name.Substring(0, 2).Equals("C_") &&
                        !CheckMatchPositionShape(shape) &&
                        shape.Type != MsoShapeType.msoPlaceholder &&
                        shape.Left + shape.Width > getViewRight
                    )
                    {
                        getViewRight = shape.Left + shape.Width;
                    }
                }
            }
            if (getViewRight > Global.slideWidth || getViewRight < Global.viewRight)
            {
                getViewRight = Global.viewRight;
            }
            Global.pptViewRight = (float)getViewRight;
            return getViewRight;
        }

        static public double GetRealViewTop(Slide slide)
        {
            double getRealViewTop = 0;
            double maxTop = -1;
            foreach (Shape mShape in slide.CustomLayout.Shapes)
            {
                double shapeBottom = mShape.Top + ComputeShapeHeight(mShape);
                if (
                    ComputeShapeRight(mShape) > Global.viewLeft &&
                    mShape.Left < Global.viewRight &&
                    shapeBottom > Global.viewTop &&
                    shapeBottom > maxTop &&
                    shapeBottom < Global.slideHeight / 2
                )
                {
                    maxTop = shapeBottom;
                }
            }
            if (maxTop != -1 && maxTop > Global.viewTop)
            {
                getRealViewTop = maxTop;
            }
            return getRealViewTop;
        }

        // @description 获取大于配置版心的有效版心底部
        static public double GetRealViewBottom(Slide slide)
        {
            double getRealViewBottom = Global.slideHeight; // 期望上内容可以多放在当前页面里
            double minTop = 9999;
            foreach (Shape mShape in slide.CustomLayout.Shapes)
            {
                // - 水平方向、在版心内
                // - 竖直方向，在页面内
                // - 在页面的下半部分
                if (
                    ComputeShapeRight(mShape) > Global.viewLeft &&
                    mShape.Left < Global.viewRight &&
                    mShape.Top < Global.slideHeight &&
                    mShape.Top < minTop &&
                    mShape.Top > Global.slideHeight / 2
                )
                {
                    minTop = mShape.Top;
                }
            }
            if (minTop != 9999 && minTop < getRealViewBottom)
            {
                getRealViewBottom = minTop;
            }
            return getRealViewBottom;
        }

        // @description 对于可能存在多个字号的、返回 Shapes 集合中最小字号
        static public float GetShapesFontSize(List<Shape> sortedShapes)
        {
            float NowMinFontSize = 999;
            for (int n = 0; n < sortedShapes.Count; n++)
            {
                List<TextRange> Runs = GetShapeRuns(sortedShapes[n]);
                if (Runs.Count != 0)
                {
                    for (int j = 0; j < Runs.Count; j++)
                    {
                        if (Runs[j].Text != " " &&
                            Runs[j].Text != "." &&
                            Runs[j].Length > 0 &&
                            Runs[j].Font.Size < NowMinFontSize &&
                            Runs[j].Font.Size > 10 &&
                            Runs[j].Font.Size > 1)
                        {
                            NowMinFontSize = Runs[j].Font.Size;
                        }
                    }
                }
            }
            return NowMinFontSize;
        }

        static public float GetMaxPtShapesLineHeight(List<Shape> sortedShapes)
        {
            float GetMaxPtShapesLineHeight = -1;
            float FontSize = GetShapesFontSize(sortedShapes);
            for (int n = 0; n < sortedShapes.Count; n++)
            {
                if (sortedShapes[n].HasTextFrame == MsoTriState.msoTrue)
                {
                    foreach (PowerPoint.TextRange Paragraph in sortedShapes[n].TextFrame.TextRange.Paragraphs())
                    {
                        float LineHeight = Paragraph.ParagraphFormat.SpaceWithin;
                        if (LineHeight > 2 &&
                            LineHeight > GetMaxPtShapesLineHeight &&
                            LineHeight > FontSize)
                        {
                            GetMaxPtShapesLineHeight = LineHeight;
                        }
                    }
                }
                else if (sortedShapes[n].HasTable == MsoTriState.msoTrue)
                {
                    foreach (PowerPoint.Row Row in sortedShapes[n].Table.Rows)
                    {
                        foreach (PowerPoint.Cell Cell in Row.Cells)
                        {
                            foreach (PowerPoint.TextRange Paragraph in Cell.Shape.TextFrame.TextRange.Paragraphs())
                            {
                                float LineHeight = Paragraph.ParagraphFormat.SpaceWithin;
                                if (LineHeight > 2 &&
                                    LineHeight > GetMaxPtShapesLineHeight &&
                                    LineHeight > FontSize)
                                {
                                    GetMaxPtShapesLineHeight = LineHeight;
                                }
                            }
                        }
                    }
                }
            }
            return GetMaxPtShapesLineHeight;
        }

        static public double GetMaxShapesLineHeight(List<Shape> sortedShapes)
        {
            double GetMaxShapesLineHeight = 1.1;
            for (int n = 0; n < sortedShapes.Count; n++)
            {
                if (sortedShapes[n].HasTextFrame == MsoTriState.msoTrue)
                {
                    foreach (PowerPoint.TextRange Paragraph in sortedShapes[n].TextFrame.TextRange.Paragraphs())
                    {
                        float LineHeight = Paragraph.ParagraphFormat.SpaceWithin;
                        if (LineHeight >= 1 && LineHeight < 2 && LineHeight > GetMaxShapesLineHeight)
                        {
                            GetMaxShapesLineHeight = LineHeight;
                        }
                    }
                }
                else if (sortedShapes[n].HasTable == MsoTriState.msoTrue)
                {
                    foreach (PowerPoint.Row Row in sortedShapes[n].Table.Rows)
                    {
                        foreach (PowerPoint.Cell Cell in Row.Cells)
                        {
                            foreach (PowerPoint.TextRange Paragraph in Cell.Shape.TextFrame.TextRange.Paragraphs())
                            {
                                float LineHeight = Paragraph.ParagraphFormat.SpaceWithin;
                                if (LineHeight >= 1 && LineHeight < 2 && LineHeight > GetMaxShapesLineHeight)
                                {
                                    GetMaxShapesLineHeight = LineHeight;
                                }
                            }
                        }
                    }
                }
            }
            return GetMaxShapesLineHeight;
        }

        // @description 获取元素的真实行数
        static public int GetShapeLines(Shape shape)
        {
            if (shape.HasTextFrame == MsoTriState.msoFalse)
            {
                return 0;
            }
            int lines = 0;
            foreach (TextRange line in shape.TextFrame.TextRange.Lines())
            {
                if (line.BoundHeight > 5)
                {
                    lines++;
                }
            }
            return lines;
        }

        // @description 通过图说 Shape 找到相应的图片 Shape
        static public Shape FindImageWithImageTip(Shape ImageTipShape)
        {
            Regex RegExp = new Regex("imagetipindex=(\\d+)");
            Shape findImageWithImageTip;
            if (ImageTipShape.HasTextFrame == MsoTriState.msoFalse || !RegExp.IsMatch(ImageTipShape.Name))
            {
                return null;
            }
            string ImageTipIndex = RegExp.Match(ImageTipShape.Name).Groups[1].Value;
            findImageWithImageTip = Global.GlobalImageTipMap[ImageTipIndex][0];
            try
            {
                string Name = findImageWithImageTip.Name;
                return findImageWithImageTip;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                foreach (Slide slide in Global.app.ActivePresentation.Slides)
                {
                    foreach (Shape ImageShape in slide.Shapes)
                    {
                        if (ImageShape.Type == MsoShapeType.msoPicture &&
                            (ImageShape.Name.Contains("imagetipindex=" + ImageTipIndex + "[&]") ||
                                ImageShape.Name.Contains("imagetipindex=" + ImageTipIndex)))
                        {
                            findImageWithImageTip = ImageShape;
                            Global.GlobalImageTipMap[ImageTipIndex] = new Shape[] {
                                ImageShape,
                                Global.GlobalImageTipMap[ImageTipIndex][1],
                            };
                            return findImageWithImageTip;
                        }
                    }
                }
            }
            return findImageWithImageTip;
        }

        // @description 通过图片 Shape 找到相应的图说 Shape
        static public Shape FindImageTipWithImage(Shape ImageShape)
        {
            Regex RegExp = new Regex("imagetipindex=(\\d+)");
            if (!RegExp.IsMatch(ImageShape.Name))
            {
                return null;
            }
            string ImageTipIndex = RegExp.Match(ImageShape.Name).Groups[1].Value;
            Shape FindImageTipWithImage = Global.GlobalImageTipMap[ImageTipIndex][1];
            try
            {
                string Name = FindImageTipWithImage.Name;
                string Text = FindImageTipWithImage.TextFrame.TextRange.Text;
                return FindImageTipWithImage;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                foreach (Slide slide in Global.app.ActivePresentation.Slides)
                {
                    foreach (Shape shape in slide.Shapes)
                    {
                        if (shape.HasTextFrame == MsoTriState.msoTrue &&
                            (shape.Name.Contains("imagetipindex=" + ImageTipIndex + "[&]") ||
                                shape.Name.Contains("imagetipindex=" + ImageTipIndex)))
                        {
                            FindImageTipWithImage = shape;
                            Global.GlobalImageTipMap[ImageTipIndex] = new Shape[] {
                                Global.GlobalImageTipMap[ImageTipIndex][0],
                                shape,
                            };
                            return FindImageTipWithImage;
                        }
                    }
                }
                return null;
            }
        }

        // @description 通过 node_id 在 PPT 中找到第一个元素
        static public Shape FindParentShapeWithNodeId(string NodeId)
        {
            List<Shape> TargetShapes = FindNodeWithNodeId(NodeId);
            if (TargetShapes.Count <= 0)
            {
                return null;
            }
            Shape TargetShape = TargetShapes[0];
            string ParentNodeId = GetShapeInfo(TargetShape)[4];
            TargetShapes = FindNodeWithNodeId(ParentNodeId);
            if (TargetShapes.Count <= 0)
            {
                return null;
            }
            return TargetShapes[0];
        }

        // @description 通过 MarkIndex 寻找相应的图片
        static public Shape FindInlineImage(string MarkIndex)
        {
            Regex regExp = new Regex("([^\\?]+)[\\?\\&]inlineimagemarkindex=(\\d+)");
            Shape findInlineImage = Global.GlobalInlineImageMap[MarkIndex];
            try
            {
                string name = findInlineImage.Name;
                return findInlineImage;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                foreach (Slide slide in Global.app.ActivePresentation.Slides)
                {
                    foreach (Shape shape in slide.Shapes)
                    {
                        if (shape.Type == MsoShapeType.msoPicture)
                        {
                            MatchCollection matches = regExp.Matches(shape.Name);
                            if (matches.Count > 0 && matches[0].Groups[2].Value == MarkIndex.ToString())
                            {
                                findInlineImage = shape;
                                Global.GlobalInlineImageMap.Remove(MarkIndex);
                                Global.GlobalInlineImageMap.Add(MarkIndex, shape);
                                return findInlineImage;
                            }
                        }
                    }
                }
                return null;
            }
        }

        // @description 通过 MarkIndex 寻找相应的表格内图片，需要注意表格内图片的 MarkIndex 是局部的、所以需要 ShapeNodeId 进行进一步约束
        static public Shape FindTableImage(int MarkIndex, string ShapeNodeId)
        {
            string key = ShapeNodeId + "#" + MarkIndex;
            Shape findTableImage = Global.GlobalTableImageMap[key];
            try
            {
                string name = findTableImage.Name;
                return findTableImage;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Regex regExp = new Regex("([^\\?]+)[\\?\\&]tableimageindex=(\\d+)");
                foreach (Slide slide in Global.app.ActivePresentation.Slides)
                {
                    foreach (Shape shape in slide.Shapes)
                    {
                        if (shape.Type == MsoShapeType.msoPicture && GetShapeInfo(shape)[0] == ShapeNodeId)
                        {
                            MatchCollection matches = regExp.Matches(shape.Name);
                            if (matches.Count > 0 && matches[0].Groups[2].Value == MarkIndex.ToString())
                            {
                                findTableImage = shape;
                                key = ShapeNodeId + "#" + MarkIndex;
                                Global.GlobalTableImageMap.Remove(key);
                                Global.GlobalTableImageMap.Add(key, shape);
                                return findTableImage;
                            }
                        }
                    }
                }
                return null;
            }
        }

        // @description 通过标记找到答案对应的填空元素
        // @todo：目前没有考虑填空出现在表格里的情况。
        static public Shape[] FindBlankWithMarkIndex(string index, Shape answerShape)
        {
            Shape[] findBlankWithMarkIndex = Global.GlobalBlankMarkIndexMap[index];
            try
            {
                Regex regExp = new Regex(@"@" + index + "@");
                Match match = regExp.Match(findBlankWithMarkIndex[0].TextFrame.TextRange.Text);
                int markIndex = match.Index;
                return findBlankWithMarkIndex;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                findBlankWithMarkIndex = new Shape[2] { null, null };
                int slideIndex = answerShape.Parent.SlideIndex;
                Slides slides = Global.app.ActivePresentation.Slides;
                for (int s = slideIndex; s <= slides.Count; s++)
                {
                    Slide slide = slides[s];
                    foreach (Shape shape in slide.Shapes)
                    {
                        if (shape.HasTextFrame == MsoTriState.msoTrue && !shape.Name.Contains("AN"))
                        {
                            if (shape.TextFrame.TextRange.Text.Contains("@" + index + "@"))
                            {
                                findBlankWithMarkIndex[0] = shape;
                                findBlankWithMarkIndex[1] = shape;
                                Global.GlobalBlankMarkIndexMap.Remove(index);
                                Global.GlobalBlankMarkIndexMap.Add(index, findBlankWithMarkIndex);
                                return findBlankWithMarkIndex;
                            }
                        }
                        else if (shape.HasTable == MsoTriState.msoTrue)
                        {
                            foreach (Row row in shape.Table.Rows)
                            {
                                foreach (Cell cell in row.Cells)
                                {
                                    if (cell.Shape.TextFrame.TextRange.Text.Contains("@" + index + "@"))
                                    {
                                        findBlankWithMarkIndex[0] = cell.Shape;
                                        findBlankWithMarkIndex[1] = shape;
                                        Global.GlobalBlankMarkIndexMap.Remove(index);
                                        Global.GlobalBlankMarkIndexMap.Add(index, findBlankWithMarkIndex);
                                        return findBlankWithMarkIndex;
                                    }
                                }
                            }
                        }
                    }
                }
                if (slideIndex > 1)
                {
                    for (int s = slideIndex - 1; s >= 1; s--)
                    {
                        Slide slide = slides[s];
                        foreach (Shape shape in slide.Shapes)
                        {
                            if (shape.HasTextFrame == MsoTriState.msoTrue && !shape.Name.Contains("AN"))
                            {
                                if (shape.TextFrame.TextRange.Text.Contains("@" + index + "@"))
                                {
                                    findBlankWithMarkIndex[0] = shape;
                                    findBlankWithMarkIndex[1] = shape;
                                    Global.GlobalBlankMarkIndexMap.Remove(index);
                                    Global.GlobalBlankMarkIndexMap.Add(index, findBlankWithMarkIndex);
                                    return findBlankWithMarkIndex;
                                }
                            }
                            else if (shape.HasTable == MsoTriState.msoTrue)
                            {
                                foreach (Row row in shape.Table.Rows)
                                {
                                    foreach (Cell cell in row.Cells)
                                    {
                                        if (cell.Shape.TextFrame.TextRange.Text.Contains("@" + index + "@"))
                                        {
                                            findBlankWithMarkIndex[0] = cell.Shape;
                                            findBlankWithMarkIndex[1] = shape;
                                            Global.GlobalBlankMarkIndexMap.Remove(index);
                                            Global.GlobalBlankMarkIndexMap.Add(index, findBlankWithMarkIndex);
                                            return findBlankWithMarkIndex;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                return findBlankWithMarkIndex;
            }
        }

        // @description 获取在元素列表中指定元素的索引
        static public Shape FindShapeWithId(int Id, List<Shape> Shapes)
        {
            Shape FindShapeWithId = null;
            for (int i = 0; i < Shapes.Count; i++)
            {
                if (Convert.ToInt32(Shapes[i].Id) == Id)
                {
                    FindShapeWithId = Shapes[i];
                    break;
                }
            }
            return FindShapeWithId;
        }

        // @description 通过标记找到答案元素
        // @params answerIndex: 答案标记
        // @params blankShapeSlideIndex: 填空位置所在的页面索引
        static public Shape FindAnswerWithMarkIndex(string answerIndex, int blankShapeSlideIndex)
        {
            Shape findAnswerWithMarkIndex;
            findAnswerWithMarkIndex = Global.GlobalAnswerMarkIndexMap[answerIndex];
            try
            {
                string Text = findAnswerWithMarkIndex.TextFrame.TextRange.Text;
                string Name = findAnswerWithMarkIndex.Name;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                findAnswerWithMarkIndex = null;
                Regex RegExp = new Regex("vbapositionanswer=" + answerIndex);
                Slides slides = Global.app.ActivePresentation.Slides;
                for (int s = blankShapeSlideIndex; s >= 1; s--)
                {
                    Slide Slide = slides[s];
                    foreach (Shape Shape in Slide.Shapes)
                    {
                        if (Shape.HasTextFrame == MsoTriState.msoTrue && (Shape.Name.Contains("AN") || Shape.Name.Contains("_WB")))
                        {
                            if (RegExp.IsMatch(Shape.Name))
                            {
                                findAnswerWithMarkIndex = Shape;
                                Global.GlobalAnswerMarkIndexMap.Remove(answerIndex);
                                Global.GlobalAnswerMarkIndexMap.Add(answerIndex, Shape);
                                return findAnswerWithMarkIndex;
                            }
                        }
                    }
                }
                if (blankShapeSlideIndex < slides.Count)
                {
                    for (int s = blankShapeSlideIndex + 1; s <= slides.Count; s++)
                    {
                        Slide Slide = slides[s];
                        foreach (Shape Shape in Slide.Shapes)
                        {
                            if (Shape.HasTextFrame == MsoTriState.msoTrue && Shape.Name.Contains("AN"))
                            {
                                if (RegExp.IsMatch(Shape.Name))
                                {
                                    findAnswerWithMarkIndex = Shape;
                                    Global.GlobalAnswerMarkIndexMap.Remove(answerIndex);
                                    Global.GlobalAnswerMarkIndexMap.Add(answerIndex, Shape);
                                    return findAnswerWithMarkIndex;
                                }
                            }
                        }
                    }
                }
            }
            return findAnswerWithMarkIndex;
        }

        // @description 通过 node_id 找到元素所在的页索引
        // @tips：目前只有添加超链接的地方在用，可以先不进行缓存
        static public int FindNodeSlide(string targetNodeId)
        {
            foreach (Slide slide in Global.app.ActivePresentation.Slides)
            {
                foreach (Shape shape in slide.Shapes)
                {
                    if (GetShapeInfo(shape)[0] == targetNodeId)
                    {
                        return slide.SlideIndex;
                    }
                }
            }
            return -1;
        }

        // @description 通过 node_id 找到其第一个子元素所在的页索引
        static public int FindFirstChildNodeSlide(string targetNodeId)
        {
            foreach (Slide slide in Global.app.ActivePresentation.Slides)
            {
                foreach (Shape shape in slide.Shapes)
                {
                    if (GetShapeInfo(shape)[4] == targetNodeId)
                    {
                        return slide.SlideIndex;
                    }
                }
            }
            return -1;
        }

        // @description 找到非排序状态下元素在当前页中的索引
        // @tips：目前只有分页批量剪切的地方在用，可以先不进行缓存。
        static public int FindSlideShapeIndex(Shape shape, Slide slide)
        {
            for (int i = 1; i <= slide.Shapes.Count; i++)
            {
                if (shape.Id == slide.Shapes[i].Id &&
                    shape.Name == slide.Shapes[i].Name)
                {
                    return i;
                }
            }
            return -1;
        }

        // @description 获取在元素列表中指定元素的索引
        static public int FindShapeIndex(Shape shape, List<Shape> shapes)
        {
            for (int i = 0; i < shapes.Count; i++)
            {
                if (shape.Name == shapes[i].Name && shape.Id == shapes[i].Id)
                {
                    return i;
                }
            }
            return -1;
        }

        static public List<Shape> FindAnNode(string baseNodeId, Shapes sortedShapes)
        {
            List<Shape> findAnNode = new List<Shape>();
            foreach (Shape shape in sortedShapes)
            {
                if (GetShapeInfo(shape)[0] == baseNodeId && shape.Name.Contains("AN"))
                {
                    findAnNode.Add(shape);
                }
            }
            return findAnNode;
        }

        static public int[] FindLineParagragh(int LineNumber, Shape Shape)
        {
            int[] FindLineParagragh = new int[] { -1, -1, -1, -1 };
            int count = 0;
            int bp = 1;
            int bl = 0;
            for (int p = 1; p <= Shape.TextFrame.TextRange.Paragraphs().Count; p++)
            {
                TextRange Paragraph = Shape.TextFrame.TextRange.Paragraphs(p);
                if (Paragraph.Text.Contains("[#]b[#]"))
                {
                    bp++;
                    bl = -1;
                }
                for (int l = 1; l <= Paragraph.Lines().Count; l++)
                {
                    count++;
                    bl++;
                    if (count >= LineNumber)
                    {
                        FindLineParagragh = new int[] { p, l, bp, bl };
                        return FindLineParagragh;
                    }
                }
            }
            return FindLineParagragh;
        }

        static public TextRange FindLine(int cindex, Shape Shape)
        {
            TextRange findLine = null;
            cindex = Convert.ToInt32(cindex);
            for (int p = 1; p <= Shape.TextFrame.TextRange.Paragraphs().Count; p++)
            {
                TextRange Paragraph = Shape.TextFrame.TextRange.Paragraphs(p);
                if (Paragraph.Start <= cindex && Paragraph.Start + Paragraph.Length >= cindex)
                {
                    for (int l = 1; l <= Paragraph.Lines().Count; l++)
                    {
                        TextRange Line = Paragraph.Lines(l);
                        if (Line.Start <= cindex &&
                            Line.Start + Line.Length >= cindex)
                        {
                            findLine = Line;
                            return findLine;
                        }
                    }
                }
            }
            return findLine;
        }

        static public int FindLineNum(int cindex, Shape shape, Shape containerShape)
        {
            try
            {
                int findLineNum = -1;
                if (shape.HasTextFrame == MsoTriState.msoFalse)
                {
                    return findLineNum;
                }
                if (cindex > shape.TextFrame.TextRange.Characters().Count)
                {
                    return findLineNum;
                }
                float TargetTop = shape.TextFrame.TextRange.Characters(cindex).BoundTop;
                for (int l = 1; l <= shape.TextFrame.TextRange.Lines().Count; l++)
                {
                    if (shape.TextFrame.TextRange.Lines(l).BoundTop == TargetTop)
                    {
                        findLineNum = l;
                        return findLineNum;
                    }
                }
                int Count = 0;
                for (int l = 1; l <= shape.TextFrame.TextRange.Lines().Count; l++)
                {
                    Count += shape.TextFrame.TextRange.Lines(l).Length;
                    if (cindex <= Count)
                    {
                        findLineNum = l;
                        return findLineNum;
                    }
                }
                return findLineNum;
            }
            catch (Exception ex)
            {
                throw new Exception("机器质检：脚本读取文本框的行数异常！", ex);
            }
        }

        // @description 遍历当前页属于相同节点的全部元素
        static public List<Shape> FindNode(string baseNodeId, List<Shape> sortedShapes)
        {
            List<Shape> FindNode = new List<Shape>();
            foreach (Shape shape in sortedShapes)
            {
                if (GetShapeInfo(shape)[0] == baseNodeId)
                {
                    FindNode.Add(shape);
                }
            }
            return FindNode;
        }

        // @description 遍历当前页属于相同节点的全部题干元素
        static public List<Shape> FindSlideSubjectNode(string baseNodeId, List<Shape> sortedShapes)
        {
            List<Shape> FindSlideSubjectNode = new List<Shape>();
            foreach (Shape shape in sortedShapes)
            {
                if (GetShapeInfo(shape)[0] == baseNodeId &&
                    !shape.Name.Contains("AN") &&
                    !shape.Name.Contains("AS") &&
                    !shape.Name.Contains("EX"))
                {
                    FindSlideSubjectNode.Add(shape);
                }
            }
            return FindSlideSubjectNode;
        }

        // @description 遍历当前页属于相同节点的全部答案/解析元素
        static public List<Shape> FindAnAsNode(string baseNodeId, List<Shape> sortedShapes)
        {
            List<Shape> FindAnAsNode = new List<Shape>();
            foreach (Shape shape in sortedShapes)
            {
                if (GetShapeInfo(shape)[0] == baseNodeId && (shape.Name.Contains("AN") || shape.Name.Contains("AS")))
                {
                    FindAnAsNode.Add(shape);
                }
            }
            return FindAnAsNode;
        }

        // @description 通过 node_id 找到节点下的全部元素
        static public List<Shape> FindNodeWithNodeId(string nodeId)
        {
            try
            {
                List<Shape> FindNodeWithNodeId = new List<Shape>();
                if (Global.GlobalNodeShapeMap.ContainsKey(nodeId))
                {
                    if (Global.GlobalNodeShapeMap[nodeId][0].Name.Length > 0 &&
                        GetShapeInfo(Global.GlobalNodeShapeMap[nodeId][0])[0] != "-1")
                    {
                        FindNodeWithNodeId = Global.GlobalNodeShapeMap[nodeId];
                    }
                }
                if (FindNodeWithNodeId.Count <= 0)
                {
                    throw new Exception("No node found with the given node id.");
                }
                foreach (Shape Shape in FindNodeWithNodeId)
                {
                    string ShapeName = Shape.Name;
                }
                return FindNodeWithNodeId;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                List<Shape> FindNodeWithNodeId = new List<Shape>();
                foreach (Slide Slide in Global.app.ActivePresentation.Slides)
                {
                    foreach (Shape Shape in Slide.Shapes)
                    {
                        if (GetShapeInfo(Shape)[0] == nodeId)
                        {
                            FindNodeWithNodeId.Add(Shape);
                        }
                    }
                }
                if (Global.GlobalNodeShapeMap.ContainsKey(nodeId))
                {
                    Global.GlobalNodeShapeMap.Remove(nodeId);
                }
                Global.GlobalNodeShapeMap.Add(nodeId, FindNodeWithNodeId);
                return FindNodeWithNodeId;
            }
        }

        // @description 通过 node_id 找到节点下的题干元素
        static public List<Shape> FindSubjectNodeWithNodeId(string nodeId)
        {
            List<Shape> FindSubjectNodeWithNodeId = new List<Shape>();
            List<Shape> Shapes = FindNodeWithNodeId(nodeId);
            foreach (Shape Shape in Shapes)
            {
                if (!(Shape.Name.Contains("AN") || Shape.Name.Contains("AS") || Shape.Name.Contains("EX")))
                {
                    FindSubjectNodeWithNodeId.Add(Shape);
                }
            }
            return FindSubjectNodeWithNodeId;
        }

        // @description 遍历当前页属于相同大小题的全部元素
        static public List<Shape> FindMoNode(string baseNodeId, List<Shape> sortedShapes)
        {
            List<Shape> FindMoNode = new List<Shape>();
            foreach (Shape Shape in sortedShapes)
            {
                if (GetShapeInfo(Shape)[0] == baseNodeId || Shape.Name.Contains("parentnodeid=" + baseNodeId))
                {
                    FindMoNode.Add(Shape);
                }
            }
            return FindMoNode;
        }

        static public Effect FindEffectWithShapeId(int shapeId, Slide slide)
        {
            Effect findEffectWithShapeId = null;
            foreach (Effect Effect in slide.TimeLine.MainSequence)
            {
                if (Effect.Shape.Id == shapeId)
                {
                    findEffectWithShapeId = Effect;
                    break;
                }
            }
            return findEffectWithShapeId;
        }

        static public int FindMaxGroupIndex()
        {
            Regex regExp = new Regex(@"(\d+)_\d+");
            int findMaxGroupIndex = 0;
            PowerPoint.Slides slides = Global.app.ActivePresentation.Slides;
            for (int s = slides.Count; s >= 1; s--)
            {
                PowerPoint.Slide Slide = slides[s];
                int Max = -1;
                foreach (PowerPoint.Shape Shape in Slide.Shapes)
                {
                    if (regExp.IsMatch(Shape.Name))
                    {
                        Match Match = regExp.Match(Shape.Name);
                        int GroupIndex = int.Parse(Match.Groups[1].Value);
                        if (GroupIndex > Max)
                        {
                            Max = GroupIndex;
                        }
                    }
                }
                if (Max != -1)
                {
                    findMaxGroupIndex = Max;
                    break;
                }
            }
            return findMaxGroupIndex;
        }

        static public PowerPoint.Shape FindImageWithWbIndex(int WbIndex)
        {
            PowerPoint.Shape findImageWithWbIndex = null;
            Regex RegExp;
            if (Global.GlobalWbImageMap.Count <= 0)
            {
                return null;
            }
            WbIndex = int.Parse(WbIndex.ToString());
            try
            {
                RegExp = new Regex(@"(\d+)#(\d+)");
                foreach (string Key in Global.GlobalWbImageMap.Keys)
                {
                    if (RegExp.IsMatch(Key))
                    {
                        Match Match = RegExp.Match(Key);
                        int l = int.Parse(Match.Groups[1].Value);
                        int r = int.Parse(Match.Groups[2].Value);
                        if (WbIndex >= l && WbIndex <= r)
                        {
                            findImageWithWbIndex = Global.GlobalWbImageMap[Key];
                            // @tips：
                            // 经过剪切的元素可以找到，但是调用 API 会报错，
                            // 所以在这里验证一下。
                            float top = findImageWithWbIndex.Top;
                            break;
                        }
                    }
                }
            }
            catch (Exception)
            {
                RegExp = new Regex(@"vbaimageposition=(\d+)#(\d+)");
                foreach (PowerPoint.Slide Slide in Global.app.ActivePresentation.Slides)
                {
                    foreach (PowerPoint.Shape Shape in Slide.Shapes)
                    {
                        if (Shape.Type == MsoShapeType.msoPicture && RegExp.IsMatch(Shape.Name))
                        {
                            Match Match = RegExp.Match(Shape.Name);
                            int l = int.Parse(Match.Groups[1].Value);
                            int r = int.Parse(Match.Groups[2].Value);
                            if (WbIndex >= l && WbIndex <= r)
                            {
                                string Key = l + "#" + r;
                                findImageWithWbIndex = Shape;
                                Global.GlobalWbImageMap.Remove(Key);
                                Global.GlobalWbImageMap.Add(Key, Shape);
                                break;
                            }
                        }
                    }
                }
            }
            return findImageWithWbIndex;
        }

        // @description 计算节点的高度，需要考虑文本框末尾是换空行的情况
        static public double ComputeShapeHeight(Shape shape)
        {
            try
            {
                if (shape.HasTextFrame == MsoTriState.msoFalse ||
                    shape.Name.Contains("WB") ||
                    shape.Name.Contains("jb-part"))
                {
                    return shape.Height;
                }
                // @tips：对于识别过有效高度的元素，返回 real height。
                if (Regex.IsMatch(shape.Name, @"rh=(\d+\.?\d+)"))
                {
                    double h = Convert.ToDouble(Regex.Match(shape.Name, @"rh=(\d+\.?\d+)").Groups[1].Value);
                    // 若 real_height 和当前答案的高度差距很大，则有可能是识别错误，也有可能是空白就这么大
                    if (h < shape.TextFrame.TextRange.BoundHeight &&
                        Math.Abs(h - shape.TextFrame.TextRange.BoundHeight) < Global.gapBetweenTextLine[2])
                    {
                        return h;
                    }
                }
                // @tips：对于应用垂直居中的文本，直接返回高度。
                if (shape.TextFrame.VerticalAnchor == MsoVerticalAnchor.msoAnchorMiddle)
                {
                    double height = shape.Height;
                    if (shape.Height < shape.TextFrame.TextRange.BoundHeight)
                    {
                        height = shape.TextFrame.TextRange.BoundHeight;
                    }
                    return height;
                }
                double shapeHeight = 0;
                // @tips：
                // Lines 中不存在文本框末尾的空换行，所以这样计算才是真实高度，
                // 而直接使用 Shape.TextFrame.TextRange.BoundHeight 会包含末尾空换行的高度。
                for (int i = 1; i <= shape.TextFrame.TextRange.Lines().Count; i++)
                {
                    if (shape.TextFrame.TextRange.Lines(i).Text.Trim().Length > 0)
                    {
                        shapeHeight += shape.TextFrame.TextRange.Lines(i).BoundHeight;
                    }
                }
                double computeShapeHeight = shapeHeight;
                if (shape.TextFrame.TextRange.BoundHeight < shapeHeight)
                {
                    computeShapeHeight = shape.TextFrame.TextRange.BoundHeight;
                }
                computeShapeHeight += shape.TextFrame.MarginTop;
                computeShapeHeight += shape.TextFrame.MarginBottom;
                if (shape.Name.StartsWith("C") && shape.Height > computeShapeHeight)
                {
                    computeShapeHeight = shape.Height;
                }
                return computeShapeHeight;
            }
            catch
            {
                return shape.Height;
            }
        }

        // @description 计算节点的宽度
        static public double ComputeShapeWidth(Shape shape)
        {
            if (shape.HasTextFrame == MsoTriState.msoFalse)
            {
                return shape.Width;
            }
            double width = shape.TextFrame.TextRange.BoundWidth;
            if (shape.TextFrame.TextRange.Lines().Count > 1)
            {
                return width;
            }
            width += shape.TextFrame.MarginLeft;
            width += shape.TextFrame.MarginRight;
            if (shape.TextFrame.Ruler.Levels[1].FirstMargin > 1)
            {
                width += shape.TextFrame.Ruler.Levels[1].FirstMargin;
            }
            if (shape.TextFrame.TextRange.ParagraphFormat.Alignment == PpParagraphAlignment.ppAlignCenter)
            {
                return shape.Width;
            }
            return width;
        }

        static public double ComputeLogicalNodeHeight(List<Shape> LogicalNodeShapes)
        {
            return ComputeLogicalNodeBottom(LogicalNodeShapes) - ComputeLogicalNodeTop(LogicalNodeShapes);
        }

        static public double ComputeLogicalNodeWidth(List<Shape> LogicalNodeShapes)
        {
            return ComputeLogicalNodeRight(LogicalNodeShapes) - ComputeLogicalNodeLeft(LogicalNodeShapes);
        }

        static public double ComputeShapeTop(Shape shape)
        {
            double computeShapeTop = shape.Top;
            if (shape.HasTextFrame == MsoTriState.msoFalse)
            {
                return computeShapeTop;
            }
            if (shape.TextFrame.TextRange.BoundTop < computeShapeTop)
            {
                computeShapeTop = shape.TextFrame.TextRange.BoundTop;
            }
            return computeShapeTop;
        }

        static public double ComputeShapeBottom(Shape shape)
        {
            return ComputeShapeTop(shape) + ComputeShapeHeight(shape);
        }

        static public double ComputeShapeRight(Shape shape)
        {
            return shape.Left + ComputeShapeWidth(shape);
        }

        static public double ComputeLogicalNodeBottom(List<Shape> LogicalNodeShapes)
        {
            double logicalNodeBottom = 0;
            for (int n = 0; n < LogicalNodeShapes.Count; n++)
            {
                double logicalNodeShapeBottom = LogicalNodeShapes[n].Top + ComputeShapeHeight(LogicalNodeShapes[n]);
                if (logicalNodeShapeBottom > logicalNodeBottom)
                {
                    logicalNodeBottom = logicalNodeShapeBottom;
                }
            }
            return logicalNodeBottom;
        }

        static public double ComputeLogicalNodeTop(List<Shape> logicalNodeShapes)
        {
            double logicalNodeTop = 9999;
            for (int n = 0; n < logicalNodeShapes.Count; n++)
            {
                double logicalNodeShapeTop = logicalNodeShapes[n].Top;
                if (logicalNodeShapeTop < logicalNodeTop)
                {
                    logicalNodeTop = logicalNodeShapeTop;
                }
            }
            return logicalNodeTop;
        }

        static public double ComputeLogicalNodeRight(List<Shape> logicalNodeShapes)
        {
            double logicalNodeRight = 0;
            for (int n = 0; n < logicalNodeShapes.Count; n++)
            {
                double logicalNodeShapeRight = logicalNodeShapes[n].Left + logicalNodeShapes[n].Width;
                if (logicalNodeShapeRight > logicalNodeRight)
                {
                    logicalNodeRight = logicalNodeShapeRight;
                }
            }
            return logicalNodeRight;
        }

        static public double ComputeLogicalNodeLeft(List<Shape> logicalNodeShapes)
        {
            double logicalNodeLeft = 9999;
            for (int n = 0; n < logicalNodeShapes.Count; n++)
            {
                double logicalNodeShapeLeft = logicalNodeShapes[n].Left;
                if (logicalNodeShapeLeft < logicalNodeLeft)
                {
                    logicalNodeLeft = logicalNodeShapeLeft;
                }
            }
            return logicalNodeLeft;
        }

        static public double ComputeContentTop(List<Shape> shapes)
        {
            List<Shape> contentShapes = new List<Shape>();
            foreach (Shape shape in shapes)
            {
                contentShapes.Add(shape);
            }
            return ComputeLogicalNodeTop(contentShapes);
        }

        static public double ComputeContentBottom(List<Shape> shapes)
        {
            double computeContentBottom = -1;
            foreach (Shape shape in shapes)
            {
                double height = shape.Height;
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    height = shape.TextFrame.TextRange.BoundHeight;
                }
                if (shape.Top + height > computeContentBottom)
                {
                    computeContentBottom = shape.Top + height;
                }
            }
            return computeContentBottom;
        }

        static public double ComputeContentLeft(List<Shape> shapes)
        {
            double computeContentLeft = 9999;
            foreach (Shape shape in shapes)
            {
                if (shape.Left < computeContentLeft)
                {
                    computeContentLeft = shape.Left;
                }
            }
            return computeContentLeft;
        }

        static public double ComputeContentRight(List<Shape> shapes)
        {
            double computeContentRight = -1;
            foreach (Shape shape in shapes)
            {
                if (ComputeShapeRight(shape) > computeContentRight)
                {
                    computeContentRight = ComputeShapeRight(shape);
                }
            }
            return computeContentRight;
        }

        static public double ComputeRangeWidth(Shape shape, int l, int r)
        {
            double computeRangeWidth = 0;
            TextRange textRange = shape.TextFrame.TextRange;
            int charCount = 0;
            int prevCharCount;
            for (int i = 1; i <= textRange.Lines().Count; i++)
            {
                TextRange line = textRange.Lines(i);
                if (line.BoundWidth > 10)
                {
                    prevCharCount = charCount;
                    charCount += line.Length;
                    if (prevCharCount < l && charCount >= l)
                    {
                        computeRangeWidth = computeRangeWidth + shape.Width - textRange.Characters(l).BoundLeft + shape.Left;
                    }
                    else if (prevCharCount < r && charCount > r)
                    {
                        computeRangeWidth = computeRangeWidth + textRange.Characters(r).BoundLeft - shape.Left;
                    }
                    else if (charCount > l && charCount <= r)
                    {
                        computeRangeWidth += shape.Width;
                    }
                }
            }
            return computeRangeWidth;
        }

        // @description 计算行高设置多少能到目标高度
        static public double ComputeTargetLineHeight(Shape shape, double targetHeight, float leftLineHeight = -1)
        {
            float standardLineHeight = (float)Global.gapBetweenTextLine[0] +
                (float)Global.gapBetweenTextLine[1] +
                (float)Global.gapBetweenTextLine[2];
            float oldSpaceWithin = shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin;
            MsoTriState oldLineRuleWithin = shape.TextFrame.TextRange.ParagraphFormat.LineRuleWithin;
            shape.TextFrame.TextRange.ParagraphFormat.LineRuleWithin = MsoTriState.msoFalse;
            shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = standardLineHeight;
            if (leftLineHeight > 0)
            {
                shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = leftLineHeight;
            }
            int retryCount = 999;
            try
            {
                while (shape.TextFrame.TextRange.BoundHeight < targetHeight && retryCount > 1)
                {
                    shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin += 1;
                    retryCount -= 1;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            double result = shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin;
            shape.TextFrame.TextRange.ParagraphFormat.LineRuleWithin = oldLineRuleWithin;
            shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = oldSpaceWithin;
            return result;
        }

        static public double ComputeLineHeight(Shape shape)
        {
            double computeLineHeight = shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin;
            if (computeLineHeight < 1)
            {
                List<TextRange> runs = GetShapeRuns(shape);
                double maxLineHeight = -1;
                for (int r = 0; r < runs.Count; r++)
                {
                    if (runs[r].ParagraphFormat.SpaceWithin > maxLineHeight)
                    {
                        maxLineHeight = runs[r].ParagraphFormat.SpaceWithin;
                    }
                }
                computeLineHeight = maxLineHeight;
            }
            return computeLineHeight;
        }

        static public double ComputeTableContentHeight(Shape shape)
        {
            double minTop = 99999;
            double maxBottom = 0;
            foreach (Row row in shape.Table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    TextRange textRange = cell.Shape.TextFrame.TextRange;
                    if (textRange.BoundTop < minTop)
                    {
                        minTop = textRange.BoundTop;
                    }
                    if (textRange.BoundTop + textRange.BoundHeight > maxBottom)
                    {
                        maxBottom = textRange.BoundTop + textRange.BoundHeight;
                    }
                }
            }
            return maxBottom - minTop;
        }

        static public double ComputeDistanceBetweenPoints(double[] p1, double[] p2)
        {
            return Math.Sqrt(Math.Pow(p1[0] - p2[0], 2) + Math.Pow(p1[1] - p2[1], 2));
        }

        // @description 计算纯文本 2 行文本之间的间距
        // @tips：（多倍行距的）多行文本中，最后一行文本的高度总是要小于其他行的文本高度的。
        static public double[] ComputeGapBetweenLines()
        {
            PowerPoint.Slides slides = Global.app.ActivePresentation.Slides;
            double[] result = new double[] { 15.5, 0, 1 };
            // @tips：
            // 计算若要对齐两个文本，则需要移动多少距离，
            // 文本+计算出来的距离、减去单倍行距下的文本的高度，就是文本之间需要的间距，
            // 计算时跳过目录页或者标题页。
            foreach (PowerPoint.Slide slide in slides)
            {
                if (!CheckHasCatalog(slide))
                {
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        bool caniuse = false;
                        if (shape.HasTextFrame == MsoTriState.msoTrue)
                        {
                            if (Regex.IsMatch(shape.TextFrame.TextRange.Text, @"<m>"))
                            {
                                caniuse = true;
                            }
                            else if ((shape.Name.Substring(0, 1) == "P" || shape.Name.Substring(0, 1) == "Q") &&
                                shape.TextFrame.TextRange.Length > 0)
                            {
                                caniuse = true;
                            }
                        }
                        if (caniuse)
                        {
                            // 新建一页用于计算，使用后删掉 
                            PowerPoint.Slide newSlide = slides.AddSlide(slide.SlideIndex + 1, slide.CustomLayout);
                            Global.app.ActiveWindow.View.GotoSlide(newSlide.SlideIndex);
                            // 计算字号、字体 
                            float fontSize = 0;
                            string fontNameFarEast = "";
                            for (int f = 1; f <= shape.TextFrame.TextRange.Length; f++)
                            {
                                PowerPoint.TextRange character = shape.TextFrame.TextRange.Characters(f);
                                if (character.Font.Size > fontSize)
                                {
                                    fontSize = character.Font.Size;
                                }
                                if (character.Length == 1 &&
                                    character.Font.Size >= fontSize &&
                                    character.Text != " " &&
                                    character.Text != "_" &&
                                    character.Font.NameFarEast.Length > 0 &&
                                    fontNameFarEast.Length <= 0)
                                {
                                    fontNameFarEast = character.Font.NameFarEast;
                                }
                            }
                            // 计算行高 
                            float spaceWithin = -1;
                            for (int p = 1; p <= shape.TextFrame.TextRange.Paragraphs().Count; p++)
                            {
                                PowerPoint.TextRange paragraph = shape.TextFrame.TextRange.Paragraphs(p);
                                if (paragraph.ParagraphFormat.SpaceWithin > spaceWithin)
                                {
                                    spaceWithin = paragraph.ParagraphFormat.SpaceWithin;
                                }
                            }
                            // 创建 1 个 2 行的纯文本的文本框
                            PowerPoint.Shape textShape = newSlide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 100, 100, 100, 100);
                            textShape.TextFrame.TextRange.Text = "测试";
                            textShape.TextFrame.TextRange.Font.Size = fontSize;
                            textShape.TextFrame.TextRange.Font.Name = "SimSun";
                            textShape.TextFrame.TextRange.Font.NameFarEast = "SimSun";
                            textShape.TextFrame.TextRange.ParagraphFormat.LineRuleWithin = (spaceWithin < 2) ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                            textShape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = spaceWithin;
                            textShape.TextFrame.MarginBottom = 0;
                            textShape.TextFrame.MarginRight = 0;
                            textShape.TextFrame.MarginLeft = 0;
                            textShape.TextFrame.MarginTop = 0;
                            textShape.TextFrame.TextRange.Characters(1).InsertAfter("\r");
                            // 计算若要对齐两个文本，则需要移动多少距离
                            double lineHeight1 = textShape.TextFrame.TextRange.Lines(1).BoundHeight;
                            double lineHeight2 = textShape.TextFrame.TextRange.Lines(2).BoundHeight;
                            if (Math.Abs(lineHeight1 - lineHeight2) < 1)
                            {
                                lineHeight2 = textShape.Height - lineHeight1;
                            }
                            if (lineHeight1 - lineHeight2 > 0 && lineHeight2 > 0)
                            {
                                // 文本+计算出来的距离、减去单倍行距下的文本的高度、除以2，就是文本之间需要的间距  
                                double oldHeight = textShape.TextFrame.TextRange.BoundHeight + lineHeight1 - lineHeight2;
                                textShape.TextFrame.TextRange.ParagraphFormat.LineRuleWithin = MsoTriState.msoTrue;
                                textShape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1;
                                double NewHeight = textShape.TextFrame.TextRange.BoundHeight;
                                result = new double[] { (oldHeight - NewHeight) / 2, lineHeight1 - lineHeight2, fontSize };
                                newSlide.Delete();
                                return result;
                            }
                            newSlide.Delete();
                        }
                    }
                }
            }
            return result;
        }

        // @description 计算 Range 的平均行高
        static public double ComputeRangeAvgLineHeight(Shape shape, int l, int r)
        {
            double result = 0;
            int lineCount = 0;
            try
            {
                TextRange textRange = shape.TextFrame.TextRange;
                int charCount = 0;
                int prevCharCount = 0;
                for (int i = 1; i <= textRange.Lines().Count; i++)
                {
                    TextRange line = textRange.Lines(i);
                    prevCharCount = charCount;
                    charCount += line.Length;
                    if (line.BoundHeight > 10)
                    {
                        if (prevCharCount < l && charCount >= l)
                        {
                            result += line.BoundHeight;
                            lineCount++;
                        }
                        else if (prevCharCount < r && charCount > r)
                        {
                            result += line.BoundHeight;
                            lineCount++;
                        }
                        else if (charCount > l && charCount <= r)
                        {
                            result += line.BoundHeight;
                            lineCount++;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            if (lineCount <= 0)
            {
                return 0;
            }
            return result / lineCount;
        }

        // @description 获取答案第一行宽度，包含两种模式：
        // - 1 获取答案第一个字符到答案文本框右侧的距离
        // - 2 获取答案第一个字符到第一行最后一个字符的距离
        static public double ComputeAnswerFirstLineWidth(Shape shape, int mode)
        {
            try
            {
                if (shape.TextFrame == null)
                {
                    return 0;
                }
                if (shape.TextFrame.TextRange.Length == 0)
                {
                    return 0;
                }
                int spaceCount = 0;
                int i = 1;
                while (i <= shape.TextFrame.TextRange.Length)
                {
                    if (shape.TextFrame.TextRange.Characters(i).Text != " ")
                    {
                        break;
                    }
                    i++;
                }
                spaceCount = i;
                if (mode == 1)
                {
                    return shape.Width -
                        shape.TextFrame.TextRange.Characters(spaceCount).BoundLeft +
                        shape.Left -
                        shape.TextFrame.Ruler.Levels[1].FirstMargin;
                }
                if (mode == 2)
                {
                    return shape.TextFrame.TextRange.Characters(shape.TextFrame.TextRange.Lines(1).Length).BoundLeft +
                        shape.TextFrame.TextRange.Characters(shape.TextFrame.TextRange.Lines(1).Length).BoundWidth -
                        shape.TextFrame.TextRange.Characters(spaceCount).BoundLeft -
                        shape.TextFrame.Ruler.Levels[1].FirstMargin;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return 0;
        }

        static public bool CheckTitlePage(Slide slide)
        {
            foreach (Shape shape in slide.Shapes)
            {
                if (!Regex.IsMatch(shape.Name, @"^C"))
                {
                    return false;
                }
            }
            return true;
        }

        static public int CheckHasLargeFormula(TextRange range)
        {
            if (!Regex.IsMatch(range.Text, @"m>"))
            {
                return -1;
            }
            float aLineHeight = (float)Math.Round(range.BoundHeight / range.Lines().Count + 0.5);
            double e = Global.gapBetweenTextLine[0];
            if (aLineHeight / Global.standardLineHeight > 1.6)
            { // 超高公式，大括号、矩阵，设置为1.1倍行高
                return 2;
            }
            else if (aLineHeight - Global.standardLineHeight > e)
            { // 较高公式，尝试用普通文本的行高来处理
                return 1;
            }
            return -1;
        }

        static public bool CheckHasQnumCatalog(PowerPoint.Shape shape, float viewBottom)
        {
            if (shape.HasTextFrame == MsoTriState.msoFalse)
            {
                return false;
            }
            // - 在页面的下半部分（更准确的判定应该是在底部、版心外
            // - 文本是数字（还需要支持其他的形式
            // - 指向试题节点
            if (shape.Top > Global.slideHeight / 2 &&
                Regex.IsMatch(shape.TextFrame.TextRange.Text, @"^[\(\（]?[\d\w]+[\)\）]?$") &&
                shape.Name.Contains("linknodeid"))
            {
                return true;
            }
            return false;
        }

        static public bool CheckHasImageTextLayout(PowerPoint.Shape shape, PowerPoint.Slide slide)
        {
            if (!shape.Name.Contains("hastextimagelayout") ||
                shape.Type != MsoShapeType.msoPicture ||
                CheckMatchPositionShape(shape))
            {
                return false;
            }
            foreach (PowerPoint.Shape otherShape in slide.Shapes)
            {
                if (otherShape.Type != MsoShapeType.msoPicture &&
                    !CheckMatchPositionShape(otherShape) &&
                    CheckYOverShapes(otherShape, shape, slide, slide, 0))
                {
                    return true;
                }
            }
            return false;
        }

        static public bool CheckFirstSlideShape(Shape targetShape)
        {
            foreach (Shape shape in GetSortedSlideShapes(targetShape.Parent))
            {
                if (!CheckMatchPositionShape(shape))
                {
                    return (targetShape.Id == shape.Id);
                }
            }
            return false;
        }

        static public bool CheckHasTallSpace(TextRange textRange, Shape shape)
        {
            bool result = false;
            float FontSize = (float)GetShapeTextInfo(shape)[0];
            foreach (TextRange c in textRange.Characters())
            {
                if ((c.Text == " " || c.Text == "_") && (c.Font.Size > FontSize + 1.5))
                {
                    result = true;
                    break;
                }
            }
            return result;
        }

        // @description 判定需要位置匹配的答案，是否会折行
        static public bool CheckWrappedAnswerShape(Shape targetAnswerShape)
        {
            bool result = false;
            if (!CheckMatchPositionAnswer(targetAnswerShape))
            {
                return result;
            }
            if (targetAnswerShape.HasTextFrame == MsoTriState.msoFalse)
            {
                return result;
            }
            if (targetAnswerShape.TextFrame.TextRange.Length == 0)
            {
                return result;
            }
            Regex regExp = new Regex("vbapositionanswer=(\\d+)");
            Match match = regExp.Match(targetAnswerShape.Name);
            string answerIndex = match.Groups[1].Value;
            Shape[] blankShapes = FindBlankWithMarkIndex(answerIndex, targetAnswerShape);
            Shape blankShape = blankShapes[0];
            Shape containerShape = blankShapes[1];
            if (blankShape == null)
            {
                return result;
            }
            regExp = new Regex("@" + answerIndex + "@");
            match = regExp.Match(blankShape.TextFrame.TextRange.Text);
            TextRange blankMatchRange = blankShape.TextFrame.TextRange.Find(match.Value);
            float blankShapeWidth = blankShape.Width;
            float positionLeft = blankMatchRange.BoundLeft - blankShape.Left;
            if (containerShape.HasTable == MsoTriState.msoTrue)
            {
                positionLeft = blankMatchRange.BoundLeft - blankShape.TextFrame.TextRange.BoundLeft;
            }
            float oldWidth = targetAnswerShape.Width;
            float oldFirstMargin = targetAnswerShape.TextFrame.Ruler.Levels[1].FirstMargin;
            targetAnswerShape.TextFrame.WordWrap = MsoTriState.msoFalse;
            targetAnswerShape.Width = targetAnswerShape.TextFrame.TextRange.BoundWidth;
            targetAnswerShape.TextFrame.Ruler.Levels[1].FirstMargin = 0;
            int c = 1;
            regExp = new Regex("^\\.\\s+\\.\\s+");
            if (regExp.IsMatch(targetAnswerShape.TextFrame.TextRange.Text))
            {
                match = regExp.Match(targetAnswerShape.TextFrame.TextRange.Text);
                c = match.Length + 1;
            }
            while (c <= targetAnswerShape.TextFrame.TextRange.Length)
            {
                if (targetAnswerShape.TextFrame.TextRange.Characters(c).Text != " ")
                {
                    break;
                }
                c++;
            }
            TextRange contentRange = targetAnswerShape.TextFrame.TextRange.Characters(c, targetAnswerShape.TextFrame.TextRange.Length + 1 - c);
            float contentWidth = contentRange.BoundWidth;
            if (contentWidth + positionLeft > blankShapeWidth && blankShape.TextFrame.TextRange.Lines().Count > 1)
            {
                result = true;
            }
            targetAnswerShape.TextFrame.WordWrap = MsoTriState.msoTrue;
            targetAnswerShape.Width = oldWidth;
            targetAnswerShape.TextFrame.Ruler.Levels[1].FirstMargin = oldFirstMargin;
            return result;
        }

        // - 1 展开的
        // - 2 非展开的
        // - 3 存在非展开元素、也存在展开元素的
        static public int CheckLogicalNodeIsExpanded(List<Shape> shapes)
        {
            int result = -1;
            foreach (Shape shape in shapes)
            {
                if (CheckIsExpaned(shape))
                {
                    if (result == -1)
                    {
                        result = 1;
                    }
                    else if (result == 2)
                    {
                        result = 3;
                        break;
                    }
                }
                else
                {
                    if (result == -1)
                    {
                        result = 2;
                    }
                    else if (result == 1)
                    {
                        result = 3;
                        break;
                    }
                }
            }
            return result;
        }

        static public bool CheckHasCatalog(Slide slide)
        {
            foreach (Shape shape in slide.Shapes)
            {
                if (!shape.Name.Substring(0, 1).Equals("C"))
                {
                    return false;
                }
            }
            return true;
        }

        // @description 判断两个元素是否在 Y 轴上重叠，（除图说外的）浮动元素不进行比较
        static public bool CheckYOverShapes(Shape shape1, Shape shape2, Slide slide1, Slide slide2, int e)
        {
            if (
                (CheckMatchPositionShape(shape1) &&
                    !shape1.Name.Contains("fixed") &&
                    !CheckHasImageTip(shape1)) ||
                (CheckMatchPositionShape(shape2) &&
                    !shape2.Name.Contains("fixed") &&
                    !CheckHasImageTip(shape2))
            )
            {
                return false;
            }
            if (shape1.Name == shape2.Name && shape1.Id == shape2.Id)
            {
                return false;
            }
            if (!(
                (shape1.Top >= shape2.Top + ComputeShapeHeight(shape2) - e) ||
                (shape1.Top + ComputeShapeHeight(shape1) - e <= shape2.Top))
            )
            {
                return true;
            }
            return false;
        }

        // @description 检查元素是否和当前页面中的其他元素，存在 Y 轴上的重叠关系
        static public bool CheckYOverShape(Shape targetShape, Slide slide, int e)
        {
            foreach (Shape shape in slide.Shapes)
            {
                if (CheckYOverShapes(targetShape, shape, slide, slide, e))
                {
                    return true;
                }
            }
            return false;
        }
        static public bool CheckYOverBlocks(List<Shape> block1, List<Shape> block2, Slide slide1, Slide slide2, int e)
        {
            bool checkYOverBlocks = false;
            if (block1 == null || block2 == null)
            {
                return checkYOverBlocks;
            }
            for (int i = 0; i < block1.Count; i++)
            {
                Shape shape1 = block1[i];
                for (int j = 0; j < block2.Count; j++)
                {
                    Shape shape2 = block2[j];
                    if (CheckYOverShapes(shape1, shape2, slide1, slide2, 0))
                    {
                        checkYOverBlocks = true;
                        return checkYOverBlocks;
                    }
                }
            }
            return checkYOverBlocks;
        }

        static public bool CheckXOverShapes(Shape shape1, Shape shape2, Slide slide1, Slide slide2, int e)
        {
            if (CheckMatchPositionAnswer(shape1) || CheckMatchPositionAnswer(shape2))
            {
                return false;
            }
            if (shape1.Name == shape2.Name && shape1.Id == shape2.Id)
            {
                return false;
            }
            // @tips：注意这里使用 .Width 而不是 ComputeShapeWidth(Shape)。
            if (!((shape1.Left >= shape2.Left + shape2.Width - e) || (shape1.Left + shape1.Width - e <= shape2.Left)))
            {
                return true;
            }
            return false;
        }

        static public bool CheckXOverBlocks(List<Shape> block1, List<Shape> block2, Slide slide1, Slide slide2, int e)
        {
            bool checkXOverBlocks = false;
            if (block1 == null || block2 == null)
            {
                return checkXOverBlocks;
            }
            for (int i = 0; i < block1.Count; i++)
            {
                Shape shape1 = block1[i];
                for (int j = 0; j < block2.Count; j++)
                {
                    Shape shape2 = block2[j];
                    if (CheckXOverShapes(shape1, shape2, slide1, slide2, 0))
                    {
                        checkXOverBlocks = true;
                        return checkXOverBlocks;
                    }
                }
            }
            return checkXOverBlocks;
        }

        // @description 检查元素是否和当前页面中的其他元素，存在 X 轴上的重叠关系
        static public bool CheckXOverShape(Shape targetShape, Slide slide, int e)
        {
            foreach (Shape shape in slide.Shapes)
            {
                if (CheckXOverShapes(targetShape, shape, slide, slide, e))
                {
                    return true;
                }
            }
            return false;
        }

        // @description 判断 2 个元素是否重叠
        static public bool CheckOverShapes(Shape shape1, Shape shape2, int e)
        {
            try
            {
                return CheckYOverShapes(shape1, shape2, shape1.Parent, shape2.Parent, e) &&
                    CheckXOverShapes(shape1, shape2, shape1.Parent, shape2.Parent, e);
            }
            catch
            {
                return false;
            }
        }

        static public bool CheckStrictOverShape(Shape targetShape, Slide slide, int e)
        {
            bool checkStrictOverShape = false;
            foreach (Shape shape in slide.Shapes)
            {
                if (shape.Id != targetShape.Id &&
                    !CheckMatchPositionShape(shape) &&
                    CheckStrictOverShapes(targetShape, shape, e))
                {
                    checkStrictOverShape = true;
                    return checkStrictOverShape;
                }
            }
            return checkStrictOverShape;
        }

        // @description 判断 2 个 Block 是否非严格重叠
        static public bool CheckOverBlocks(List<Shape> block1, List<Shape> block2, int e)
        {
            bool checkOverBlocks = false;
            for (int i = 0; i < block1.Count; i++)
            {
                for (int j = 0; j < block2.Count; j++)
                {
                    if (CheckOverShapes(block1[i], block2[j], e))
                    {
                        checkOverBlocks = true;
                        return checkOverBlocks;
                    }
                }
            }
            return checkOverBlocks;
        }

        // @description 判断 2 个 Block 是否严格重叠
        static public bool CheckStrictOverBlocks(List<Shape> block1, List<Shape> block2, int e)
        {
            bool checkStrictOverBlocks = false;
            for (int i = 0; i < block1.Count; i++)
            {
                for (int j = 0; j < block2.Count; j++)
                {
                    if (CheckStrictOverShapes(block1[i], block2[j], e))
                    {
                        checkStrictOverBlocks = true;
                        return checkStrictOverBlocks;
                    }
                }
            }
            return checkStrictOverBlocks;
        }

        // @description 判断 2 个元素是否严格重叠，框内留白的部分不认为是发生重叠
        static public bool CheckStrictOverShapes(Shape shape1, Shape shape2, int e)
        {
            if (!CheckOverShapes(shape1, shape2, e))
            {
                return false;
            }
            // - 都是图片的情况
            if (shape1.HasTextFrame == MsoTriState.msoFalse && shape2.HasTextFrame == MsoTriState.msoFalse)
            {
                double top1 = shape1.Top;
                double left1 = shape1.Left;
                double height1 = shape1.Height;
                double width1 = shape1.Width;
                double top2 = shape2.Top;
                double left2 = shape2.Left;
                double height2 = shape2.Height;
                double width2 = shape2.Width;
                if (!((top1 > top2 + height2 - e) || (top1 + height1 - e < top2)) && !((left1 > left2 + width2 - e) || (left1 + width1 - e < left2)))
                {
                    return true;
                }
            }
            // @tips：其中 1 个是图片的情况，Shape1 为图片、Shape2 为文本
            if ((shape1.HasTextFrame == MsoTriState.msoTrue && shape2.HasTextFrame == MsoTriState.msoFalse) ||
                (shape2.HasTextFrame == MsoTriState.msoTrue && shape1.HasTextFrame == MsoTriState.msoFalse))
            {
                Shape t2 = shape2;
                Shape t1 = shape1;
                if (shape1.HasTextFrame == MsoTriState.msoTrue && shape2.HasTextFrame == MsoTriState.msoFalse)
                {
                    (t1, t2) = (t2, t1);
                }
                for (int l = 1; l <= t2.TextFrame.TextRange.Lines().Count; l++)
                {
                    TextRange textRange2 = t2.TextFrame.TextRange.Lines(l, 1);
                    double top1 = t1.Top;
                    double left1 = t1.Left;
                    double height1 = t1.Height;
                    double width1 = t1.Width;
                    double top2 = textRange2.BoundTop;
                    double left2 = textRange2.BoundLeft;
                    double height2 = textRange2.BoundHeight;
                    double width2 = textRange2.BoundWidth;
                    if (!((top1 > top2 + height2 - e) || (top1 + height1 - e < top2)) && !((left1 > left2 + width2 - e) || (left1 + width1 - e < left2)))
                    {
                        return true;
                    }
                }
            }
            if (shape1.HasTextFrame == MsoTriState.msoTrue && shape2.HasTextFrame == MsoTriState.msoTrue)
            {
                // @tips：都是文本的情况
                for (int k = 1; k <= shape1.TextFrame.TextRange.Lines().Count; k++)
                {
                    TextRange line1 = shape1.TextFrame.TextRange.Lines(k, 1);
                    for (int l = 1; l <= shape2.TextFrame.TextRange.Lines().Count; l++)
                    {
                        TextRange line2 = shape2.TextFrame.TextRange.Lines(l, 1);
                        double top1 = line1.BoundTop;
                        double left1 = line1.BoundLeft;
                        double height1 = line1.BoundHeight;
                        double width1 = line1.BoundWidth;
                        double top2 = line2.BoundTop;
                        double left2 = line2.BoundLeft;
                        double height2 = line2.BoundHeight;
                        double width2 = line2.BoundWidth;
                        if (!((top1 > top2 + height2 - e) || (top1 + height1 - e < top2)) && !((left1 > left2 + width2 - e) || (left1 + width1 - e < left2)))
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        // @description 判断是否是需要位置匹配的元素
        // - 需要匹配位置的答案
        // - inline image
        // - table image
        // - 绝对定位的元素
        // - 图说
        static public bool CheckMatchPositionShape(Shape shape)
        {
            try
            {
                string[] shapeInfo = GetShapeInfo(shape);
                string shapeLabel = shapeInfo[2];
                bool hasTableImageShape = shapeLabel == ".table_image";
                bool hasFixed = shapeLabel == ".fixed";
                bool hasInlineImage = CheckInlineImage(shape);
                bool hasMatchAnswer = CheckMatchPositionAnswer(shape);
                bool hasLongTextAnswer = shape.Name.Contains("haslongtextanswer");
                return hasTableImageShape ||
                    hasFixed ||
                    hasInlineImage ||
                    hasMatchAnswer ||
                    CheckHasImageTip(shape) ||
                    hasLongTextAnswer;
            }
            catch (Exception)
            {
                return false;
            }
        }

        // @description 检查节点是否包含图说
        static public bool CheckHasImageTip(Shape shape)
        {
            try
            {
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    if (shape.Name.Contains("imagetipindex"))
                    {
                        return true;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        // @description 检查元素是否是 inline image
        static public bool CheckInlineImage(Shape shape)
        {
            try
            {
                if (shape.Type == MsoShapeType.msoPicture && Regex.IsMatch(shape.Name, @"inlineimagemarkindex=(\d+)"))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        // @description 判断元素是否在某个节点元素集合的内部，形如大小题结构、并且大题题图和小题题干并列的情况
        static public bool CheckHasWrappedNode(Shape targetShape, List<Shape> shapes)
        {
            Dictionary<string, int> m = new Dictionary<string, int>();
            string targetShapeNodeId = GetShapeInfo(targetShape)[0];
            bool found = false;
            foreach (Shape shape in shapes)
            {
                string shapeNodeId = GetShapeInfo(shape)[0];
                if (shape.Name == targetShape.Name && shape.Id == targetShape.Id)
                {
                    found = true;
                }
                if (found)
                {
                    return targetShapeNodeId != shapeNodeId && m.ContainsKey(shapeNodeId);
                }
                if (!m.ContainsKey(shapeNodeId))
                {
                    m.Add(shapeNodeId, 1);
                }
            }
            return false;
        }

        // @description 检查元素是否是某个节点的子节点
        static public bool CheckHasParentNode(Shape shape)
        {
            return GetShapeInfo(shape)[4] != "-1";
        }

        // @description 检查元素是否是某个元素的父节点
        static public bool CheckHasChildNode(Shape shape)
        {
            return Global.GlobalParentNodeMap.ContainsKey(GetShapeInfo(shape)[0]);
        }

        // @description 判断当前元素上方是否存在其他节点
        // - 接受标题在非标题的节点上方
        // - 若 node_id 不同则是其他元素
        // - 若节点是填空题或者选择题的答案，则不当作独立的节点进行判断
        static public bool CheckHasOtherNode(int index, List<Shape> shapes)
        {
            if (index <= 1)
            {
                return false;
            }
            // - 若节点是填空题或者选择题的答案，则不当作独立的节点进行判断
            string shapeLabel = GetShapeInfo(shapes[index])[2];
            string shapeNodeId = GetShapeInfo(shapes[index])[1];
            if (shapeLabel == ".blank" || (shapeLabel == ".bracket" && shapes[index].Name.Contains("AN")))
            {
                return false;
            }
            string baseNodeId = shapeNodeId;
            int i = index - 1;
            bool hasOtherNode = false;
            try
            {
                while (i > 0)
                {
                    shapeLabel = GetShapeInfo(shapes[i])[2];
                    shapeNodeId = GetShapeInfo(shapes[i])[1];
                    // - 接受标题在非标题的节点上方
                    // - 若 node_id 不同则是其他元素
                    // - 若节点是填空题或者选择题的答案，则不当作独立的节点进行判断
                    // - 若节点上方存在同名节点，则不当做独立的节点进行判断
                    if (shapes[i].Name.Substring(0, 2) == "C_")
                    {
                    }
                    else if (shapeLabel == ".blank" || (shapeLabel == ".bracket" && shapes[i].Name.Contains("AN")))
                    {
                    }
                    else if (shapes[i].Name == shapes[index].Name)
                    {
                        // @tips：
                        // 若节点上方存在同名节点，则不当做独立的节点进行判断，
                        // 因为在上面的节点已经做过分页处理，
                        // 不再进行冗余的判断、避免过度切分。
                        hasOtherNode = false;
                        break;
                    }
                    else if (shapeNodeId != baseNodeId)
                    {
                        hasOtherNode = true;
                        break;
                    }
                    i--;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return hasOtherNode;
        }

        // @description 检查元素是否是上一页分页下来的
        static public bool CheckHasSplittedShape(PowerPoint.Shape shape)
        {
            PowerPoint.Slide slide = shape.Parent;
            PowerPoint.Slides slides = Global.app.ActivePresentation.Slides;
            bool result = false;
            if (slide.SlideIndex <= 1)
            {
                return result;
            }
            PowerPoint.Slide prevSlide = slides[slide.SlideIndex - 1];
            for (int i = 1; i <= prevSlide.Shapes.Count; i++)
            {
                if (prevSlide.Shapes[i].Name == shape.Name)
                {
                    result = true;
                    break;
                }
            }
            return result;
        }

        // @description 检查节点是否是上一页分页下来的
        static public bool CheckHasSplittedNode(PowerPoint.Shape shape)
        {
            PowerPoint.Slide slide = shape.Parent;
            PowerPoint.Slides slides = Global.app.ActivePresentation.Slides;
            bool result = false;
            if (slide.SlideIndex <= 1)
            {
                return result;
            }
            PowerPoint.Slide prevSlide = slides[slide.SlideIndex - 1];
            string shapeNodeId = GetShapeInfo(shape)[0];
            for (int i = 1; i <= prevSlide.Shapes.Count; i++)
            {
                if (GetShapeInfo(prevSlide.Shapes[i])[0] == shapeNodeId)
                {
                    result = true;
                    break;
                }
            }
            return result;
        }

        // @description 检查元素上方是否存在被分页的其他节点
        static public bool CheckHasSplittedOtherNode(PowerPoint.Shape shape, List<PowerPoint.Shape> sortedShapes)
        {
            PowerPoint.Slides slides = Global.app.ActivePresentation.Slides;
            if (shape.Parent.SlideIndex <= 1)
            {
                return false;
            }
            PowerPoint.Slide prevSlide = slides[shape.Parent.SlideIndex - 1];
            if (prevSlide.Shapes.Count < 1)
            {
                return false;
            }
            PowerPoint.Shape firstShape = null;
            for (int i = 0; i < sortedShapes.Count; i++)
            {
                if (GetShapeInfo(sortedShapes[i])[0] == GetShapeInfo(shape)[0])
                {
                    break;
                }
                // @tips：
                // - 无视掉需要匹配位置的答案
                // - 无视掉当前元素的父节点
                if (!CheckMatchPositionAnswer(sortedShapes[i]) &&
                    !shape.Name.Contains("parentnodeid=" + GetShapeInfo(sortedShapes[i])[0]))
                {
                    firstShape = sortedShapes[i];
                    break;
                }
            }
            if (firstShape == null)
            {
                return false;
            }
            PowerPoint.Shape lastShape = prevSlide.Shapes[prevSlide.Shapes.Count];
            // @tips：
            // 目前暂时接受试题节点上方被分页的是段落节点。
            // 需要考虑使用 PPT 预处理工具手动段落内分页的情况。
            return GetShapeInfo(lastShape)[0] == GetShapeInfo(firstShape)[0] && !firstShape.Name.StartsWith("P");
        }


        // @description 判断 Shape 是否是需要位置匹配的答案
        static public bool CheckMatchPositionAnswer(Shape shape)
        {
            return shape.Name.Contains("vbapositionanswer");
        }

        static public bool CheckEmptyTable(Shape shape)
        {
            if (shape.HasTable == MsoTriState.msoFalse)
            {
                return false;
            }
            bool isEmpty = true;
            foreach (Row row in shape.Table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    if (cell.Shape.TextFrame.TextRange.Length > 0)
                    {
                        isEmpty = false;
                        break;
                    }
                }
                if (!isEmpty)
                {
                    break;
                }
            }
            return isEmpty;
        }

        // @description 判断当前元素是否是图说，并且其对应的图片在指定页
        static public bool CheckSlideImageTip(Shape shape, Slide slide)
        {
            bool checkSlideImageTip;
            if (!CheckHasImageTip(shape))
            {
                checkSlideImageTip = false;
                return checkSlideImageTip;
            }
            Shape image = FindImageWithImageTip(shape);
            if (image == null)
            {
                checkSlideImageTip = false;
                return checkSlideImageTip;
            }
            checkSlideImageTip = (image.Parent.SlideIndex == slide.SlideIndex);
            return checkSlideImageTip;
        }

        static public bool CheckSlideOverFlow(Slide slide)
        {
            int e = 50;
            foreach (Shape shape in slide.Shapes)
            {
                if (shape.Top < -1)
                {
                    return true;
                }
                else if (shape.Left < -1)
                {
                    return true;
                }
                else if (ComputeShapeRight(shape) > Global.slideWidth + e)
                {
                    return true;
                }
                else if (ComputeShapeBottom(shape) > Global.slideHeight + e)
                {
                    return true;
                }
            }
            return false;
        }

        // @description 判断单元格是否是合并单元格
        static public bool CheckMergedCell(Cell targetCell, Shape tableShape)
        {
            string key = tableShape.Parent.SlideIndex + "#" +
                tableShape.Id + "#" +
                targetCell.Shape.Left + "," + targetCell.Shape.Top;
            if (Global.GlobalMergedCellMap.ContainsKey(key)) // 这里优先使用缓存
            {
                return Global.GlobalMergedCellMap[key];
            }
            List<double[]> coords = new List<double[]>();
            foreach (Row row in tableShape.Table.Rows) // 若没有找到缓存的结果，则直接重新刷一遍当前表格
            {
                foreach (Cell cell in row.Cells)
                {
                    key = tableShape.Parent.SlideIndex + "#" +
                        tableShape.Id + "#" +
                        cell.Shape.Left + "," + cell.Shape.Top;
                    bool hasMergedCell = false;
                    foreach (double[] coord in coords) // 存在重复坐标的单元格，则判定为合并单元格
                    {
                        if (coord[0] == cell.Shape.Left && coord[1] == cell.Shape.Top)
                        {
                            hasMergedCell = true;
                            break;
                        }
                    }
                    if (!Global.GlobalMergedCellMap.ContainsKey(key))
                    {
                        Global.GlobalMergedCellMap.Add(key, hasMergedCell);
                    }
                    else
                    {
                        Global.GlobalMergedCellMap[key] = hasMergedCell;
                    }
                    coords.Add(new double[] { cell.Shape.Left, cell.Shape.Top });
                }
            }
            key = tableShape.Parent.SlideIndex + "#" +
                tableShape.Id + "#" +
                targetCell.Shape.Left + "," + targetCell.Shape.Top;
            return Global.GlobalMergedCellMap[key];
        }

        // @descrption 判定文本中是否包含复杂公式（has_big_formula
        static public bool CheckHasBFormula(Shape shape, Shape containerShape)
        {
            if (CheckLineFeedFormula(shape.TextFrame.TextRange) || Regex.IsMatch(shape.TextFrame.TextRange.Text, @"^\s*<m>"))
            {
                return true;
            }
            if (!Regex.IsMatch(shape.TextFrame.TextRange.Text, @"<m>.*<\/m>"))
            {
                return false;
            }
            double fontSize = (double)GetShapeTextInfo(shape)[0];
            foreach (Match match in Regex.Matches(shape.TextFrame.TextRange.Text, @"<m>.*<\/m>"))
            {
                TextRange matchRange = shape.TextFrame.TextRange.Characters(match.Index + 1, match.Length);
                if (matchRange.BoundWidth / fontSize > 7)
                {
                    return true;
                }
            }
            return false;
        }

        // @description 判断是否有折行的公式
        static public bool CheckLineFeedFormula(TextRange textRange)
        {
            Regex regex = new Regex("<\\/m>");
            bool checkLineFeedFormula = false;
            for (int i = 0; i < textRange.Lines().Count; i++)
            {
                string line = textRange.Lines(i, 1).Text;
                if (regex.IsMatch(line))
                {
                    MatchCollection matches = regex.Matches(line);
                    Match lastMatch = matches[matches.Count - 1];
                    // - 对于行末只有标记的情况进行豁免
                    if (lastMatch.Value == "<m>" && lastMatch.Index + lastMatch.Length != line.Length)
                    {
                        checkLineFeedFormula = true;
                        break;
                    }
                }
            }
            return checkLineFeedFormula;
        }


        // @description 判断某 LineRange 是否是段落内标题
        // @params index - LineRange 在 TextRange 中的索引
        // @params lineRange - LineRange
        // @params textRange - TextRange
        // @params inline - 是否接受不在文本框的首行或者尾行
        // @@todo：当前分页逻辑待删除。
        public bool CheckLineParagraphTitle(int index, TextRange lineRange, TextRange textRange, bool inline = false)
        {
            bool checkChr13End = lineRange.Characters(lineRange.Length).Text == "\r"; // chr(13)
            // - 单行，并且文本全部加粗
            if (index == 1 && index == textRange.Lines().Count)
            {
                bool hasBold = true;
                for (int c = 1; c <= lineRange.Length; c++)
                {
                    if (lineRange.Characters(c).Font.Bold != MsoTriState.msoTrue)
                    {
                        hasBold = false;
                    }
                }
                if (hasBold)
                {
                    return true;
                }
            }
            // - 以 1. 或者（1）开头
            // - 非 Inline 时是第一行、并且以换行符结束
            // - 非 Inline 时是最后一行
            // - Inline 时以换行符结束
            // - 不包含答案标记
            if (Regex.IsMatch(lineRange.Text, @"@\d+@"))
            {
                return false;
            }
            if (
                Regex.IsMatch(lineRange.Text, @"(^\s*（\d+）.+)|(^\s*\d\..+)") &&
                ((inline && checkChr13End) ||
                (!inline && ((index == 1 && checkChr13End) || index == textRange.Lines().Count)))
            )
            {
                return true;
            }
            return false;
        }

        // @description 检查自 s 开始的 l 位是否匹配 Pattern
        static public bool CheckContentPattern(int s, int l, string TextRange, string Pattern)
        {
            try
            {
                if ((s + l - 1) > TextRange.Length)
                {
                    return false;
                }
                return TextRange.Substring(s - 1, l) == Pattern;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        static public bool CheckIsExpaned(Shape shape)
        {
            // @tips：
            // 判断节点是否是展开的，不能用中心做比较，因为可能存在左右 Padding 配置不一样的情况，
            // 同时也不可以直接用 hastextimage 标记判断，因为容器内可能存在展开的元素，
            // 目前建议用距离版心左右的距离来判断。
            bool checkIsExpaned = false;
            if (CheckHasImageTip(shape))
            {
                checkIsExpaned = false;
                return checkIsExpaned;
            }
            if (!shape.Name.Contains("hastextimagelayout=1"))
            {
                checkIsExpaned = true;
                return checkIsExpaned;
            }
            if (Math.Abs(shape.Left - Global.viewLeft) < 5 &&
                (Math.Abs(shape.Left + shape.Width - Global.viewRight) < 5 ||
                    shape.Left + shape.Width > Global.viewRight))
            {
                checkIsExpaned = true;
            }
            if (Math.Abs(shape.Left + shape.Width / 2 - Global.slideWidth / 2) <= 1)
            {
                checkIsExpaned = true;
            }
            return checkIsExpaned;
        }

        // @description 判断两个元素是否不同侧
        // - 都在左侧，同侧
        // - 都在右侧，同侧
        // - 一个比较窄、一个展开的图文布局的情况，当做是同侧的
        static public bool CheckHasDiffside(Shape shape1, Shape shape2)
        {
            bool checkHasDiffside = true;
            if (shape1.Left + ComputeShapeWidth(shape1) / 2 > Global.slideWidth / 2 &&
                shape2.Left + ComputeShapeWidth(shape2) / 2 > Global.slideWidth / 2)
            { // 都在右侧
                checkHasDiffside = false;
            }
            if (shape1.Left + ComputeShapeWidth(shape1) / 2 < Global.slideWidth / 2 &&
                shape2.Left + ComputeShapeWidth(shape2) / 2 < Global.slideWidth / 2)
            { // 都在左侧
                checkHasDiffside = false;
            }
            if (!(CheckIsExpaned(shape1) && CheckIsExpaned(shape2)))
            { // 需要注意一个比较窄、一个展开的图文布局的情况，当做是同侧的
                if (!((shape1.Left + ComputeShapeWidth(shape1) / 2 > Global.slideWidth / 2 &&
                        shape2.Left + ComputeShapeWidth(shape2) / 2 < Global.slideWidth / 2) ||
                    (shape2.Left + ComputeShapeWidth(shape2) / 2 > Global.slideWidth / 2 &&
                        shape1.Left + ComputeShapeWidth(shape1) / 2 < Global.slideWidth / 2)))
                { // 过滤掉一左一右的情况
                    checkHasDiffside = false;
                }
                if (CheckIsExpaned(shape1) || CheckIsExpaned(shape2))
                {
                    checkHasDiffside = false;
                }
            }
            if (CheckIsExpaned(shape1) && CheckIsExpaned(shape2))
            { // 都是展开的情况
                checkHasDiffside = false;
            }
            return checkHasDiffside;
        }

        static public bool CheckLogicalNodeHasDiffside(List<Shape> Shapes1, List<Shape> Shapes2)
        {
            foreach (Shape Shape1 in Shapes1)
            {
                foreach (Shape Shape2 in Shapes2)
                {
                    if (!CheckHasDiffside(Shape1, Shape2))
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        static public bool Check75Options(Shape Shape)
        {
            bool result = false;
            if (Global.pptSubject == "english" &&
                Shape.Name.StartsWith("QM") &&
                !CheckMatchPositionShape(Shape) &&
                Shape.Left > Global.slideWidth / 2 &&
                Shape.HasTextFrame == MsoTriState.msoTrue)
            {
                result = true;
            }
            if (Shape.HasTextFrame == MsoTriState.msoTrue)
            {
                if (Shape.TextFrame.TextRange.Text.Contains("续表"))
                {
                    result = false;
                }
            }
            return result;
        }

        static public bool CheckChoiceQuestion(Shape shape, Slide slide)
        {
            Regex regex = new Regex("^[ABCD]$");
            bool checkChoiceQuestion = false;
            if (shape.Name.StartsWith("QC"))
            {
                checkChoiceQuestion = true;
                return checkChoiceQuestion;
            }
            string shapeNodeId = GetShapeInfo(shape)[0];
            List<Shape> answers = FindAnNode(shapeNodeId, slide.Shapes);
            if (answers.Count == 1)
            {
                if (answers[0].HasTextFrame == MsoTriState.msoTrue)
                {
                    if (regex.IsMatch(answers[0].TextFrame.TextRange.Text.Trim()))
                    {
                        checkChoiceQuestion = true;
                        return checkChoiceQuestion;
                    }
                }
            }
            return checkChoiceQuestion;
        }

        static public bool CheckRangeHigherAvgLineHeight(Shape shape, int l, int r, float avgLineHeight, int e)
        {
            bool checkRangeHigherAvgLineHeight = false;
            float standardLineHeight = (float)Global.gapBetweenTextLine[0] +
                (float)Global.gapBetweenTextLine[1] +
                (float)Global.gapBetweenTextLine[2];
            int charCount = 0;
            int prevCharCount;
            float fontSize = (float)GetShapeTextInfo(shape)[0];
            for (int i = 1; i <= shape.TextFrame.TextRange.Lines().Count; i++)
            {
                TextRange line = shape.TextFrame.TextRange.Lines(i);
                prevCharCount = charCount;
                charCount += line.Length;
                if ((prevCharCount < l && charCount >= l) ||
                    (prevCharCount < r && charCount > r) ||
                    (charCount > l && charCount <= r))
                {
                    if (line.BoundHeight > avgLineHeight + e)
                    {
                        checkRangeHigherAvgLineHeight = true;
                        return checkRangeHigherAvgLineHeight;
                    }
                    if (line.BoundHeight > (fontSize * 2 + e))
                    {
                        checkRangeHigherAvgLineHeight = true;
                        return checkRangeHigherAvgLineHeight;
                    }
                    if (line.BoundHeight > (standardLineHeight + e))
                    {
                        checkRangeHigherAvgLineHeight = true;
                        return checkRangeHigherAvgLineHeight;
                    }
                }
            }
            return checkRangeHigherAvgLineHeight;
        }

        // @description 判断段尾是否存在不好看的小尾巴
        static public bool CheckParagraphShortText(Shape shape, TextRange paragraph, Shape containerShape)
        {
            Regex regExp = new Regex("");
            bool checkParagraphShortText = false;
            if (paragraph.Length <= 0)
            {
                checkParagraphShortText = false;
                return checkParagraphShortText;
            }
            int linesCount = paragraph.Lines().Count;
            if (linesCount <= 1)
            {
                checkParagraphShortText = false;
                return checkParagraphShortText;
            }
            // - 若某行的字符数小于等于 3 个，则认为是需要处理的文末小尾巴
            // - 若某行的开头是 "__." 的形式，则认为需要缩紧，尽量保证填空横线完整不换行
            // - 若某行存在连续的英文字符，则不进行处理，避免单词看起来很挤
            // - 若某行存在连续的英文字符答案，则不进行处理，避免单词看起来很挤
            for (int i = 1; i <= linesCount; i++)
            {
                if (Regex.IsMatch(paragraph.Lines(i).Text, @"[a-zA-Z]+"))
                {
                    foreach (Match match in Regex.Matches(paragraph.Lines(i).Text, @"[a-zA-Z]+"))
                    {
                        if (match.Length >= 3)
                        {
                            checkParagraphShortText = false;
                            return checkParagraphShortText;
                        }
                    }
                }
                // - 若某行的字符数小于等于 3 个，则认为是需要处理的文末小尾巴
                // - 若某行的开头是 "__." 的形式，则认为需要缩紧，尽量保证填空横线完整不换行
                if (paragraph.Lines(i).Characters().Count <= 3 && i > 1)
                {
                    checkParagraphShortText = true;
                    // @tips：
                    // 若某行只有 1 个空格，那么认为不存在小尾巴，
                    // 留在句首标点的环节去处理。
                    if (Regex.IsMatch(paragraph.Lines(i).Characters(1).Text, @"[!),.:;?\]、。—ˇ¨〃々～‖…’”〕〉》」』〗】∶！＇），．：；？］｀｜｝]"))
                    {
                        checkParagraphShortText = false;
                    }
                }
            }
            if (Regex.IsMatch(paragraph.Text, "@(\\d+)@"))
            {
                foreach (Match m in regExp.Matches(paragraph.Text))
                {
                    string answerIndex = m.Groups[1].Value;
                    Shape answerShape = FindAnswerWithMarkIndex(answerIndex, containerShape.Parent.SlideIndex);
                    // @tips：
                    // 若文本是填空题干，那么由于缩小题干字间距的话，也需要相应缩小其答案的字间距，
                    // 那么答案是否允许缩进，也需要进行判断。
                    if (answerShape != null)
                    {
                        if (answerShape.HasTextFrame == MsoTriState.msoTrue)
                        {
                            regExp = new Regex("[a-zA-Z]+");
                            foreach (Match n in regExp.Matches(answerShape.TextFrame.TextRange.Text))
                            {
                                if (n.Length >= 3)
                                {
                                    checkParagraphShortText = false;
                                    return checkParagraphShortText;
                                }
                            }
                        }
                    }
                }
            }
            return checkParagraphShortText;
        }

        // @description 判断表格的某行是否是表头
        static public bool CheckTableTH(Row row, Shape tableShape)
        {
            bool checkTableTH = false;
            if (GetTableRowHeight(row, tableShape) <= 0)
            {
                return checkTableTH;
            }
            checkTableTH = true;
            for (int c = 1; c <= row.Cells.Count; c++)
            {
                // 当前行全部单元格包含标记，或者为空
                if (row.Cells[c].Shape.TextFrame.TextRange.Text.Trim().Length > 0)
                {
                    if (!Regex.IsMatch(row.Cells[c].Shape.TextFrame.TextRange.Text, @"<th>"))
                    {
                        checkTableTH = false;
                        break;
                    }
                }
            }
            return checkTableTH;
        }

        // @description 找到当前页面中，可以横向排布的、比较窄的填空题
        // - 需要最上面的是标题
        // - 需要页面中其他元素都是宽度比较窄的填空题
        // - 需要溢出版心的填空题数量不超过 2 个
        public List<Shape> CheckShortBlankCollection(Slide slide)
        {
            List<Shape> sortedShapes = new List<Shape>();
            List<Shape> shortBlankCollection = new List<Shape>();
            Regex regExp = new Regex("([^\\.]+)(\\.[^\\.]+)?#([\\d\\w]+)(\\.\\w+)?");
            for (int i = 1; i <= slide.Shapes.Count; i++)
            {
                Match match = regExp.Match(slide.Shapes[i].Name);
                if (match.Success)
                {
                    string shapeLabel = match.Groups[2].Value;
                    if (shapeLabel == ".fixed")
                    {
                        continue;
                    }
                    sortedShapes.Add(slide.Shapes[i]);
                }
            }
            for (int i = 0; i < sortedShapes.Count - 1; i++)
            {
                for (int j = i + 1; j < sortedShapes.Count; j++)
                {
                    if (sortedShapes[i].Top > sortedShapes[j].Top)
                    {
                        (sortedShapes[i], sortedShapes[j]) = (sortedShapes[j], sortedShapes[i]);
                    }
                }
            }
            for (int i = 0; i < sortedShapes.Count; i++)
            {
                Shape currentShape = sortedShapes[i];
                if (!currentShape.Name.StartsWith("C_") && !currentShape.Name.StartsWith("QB"))
                {
                    return new List<Shape>();
                }
                if (currentShape.HasTextFrame == MsoTriState.msoFalse)
                {
                    return new List<Shape>();
                }
                if (currentShape.TextFrame.TextRange.BoundWidth >= Global.slideWidth / 2)
                {
                    return new List<Shape>();
                }
                if (i > 1)
                {
                    Shape prevShape = sortedShapes[i - 1];
                    if (!prevShape.Name.StartsWith("C_") && currentShape.Name.StartsWith("C_"))
                    {
                        return new List<Shape>();
                    }
                }
                if (currentShape.Name.StartsWith("QB"))
                {
                    shortBlankCollection.Add(currentShape);
                }
            }
            List<Shape> tc = new List<Shape>();
            for (int i = 0; i < shortBlankCollection.Count; i++)
            {
                if (!CheckMatchPositionAnswer(shortBlankCollection[i]) &&
                    !shortBlankCollection[i].Name.StartsWith("C_"))
                {
                    tc.Add(shortBlankCollection[i]);
                }
            }
            int k;
            for (k = 0; k < tc.Count; k++)
            {
                if (tc[k].Top + ComputeShapeHeight(tc[k]) > Global.viewBottom)
                {
                    break;
                }
            }
            if (k <= tc.Count && tc.Count <= (k - 1) * 2)
            {
                return shortBlankCollection;
            }
            return new List<Shape>();
        }

        // @description 删除位置标签：
        // - 答案位置
        // - Inline 图片位置
        // - 表格内图片位置
        // - 表头标记
        // - 单元格斜线标记
        // - zzd
        // - inline box
        // - 对勾答案位置
        static public void DeletePositionMark(string shapeName, Shape shape)
        {
            bool hasWpsProcessed = false;
            if (shape.TextFrame.TextRange.Characters(1).ParagraphFormat.SpaceWithin > 2)
            {
                hasWpsProcessed = true;
            }
            if (Regex.IsMatch(shape.TextFrame.TextRange.Text, @"@(\d+)@"))
            {
                foreach (Match match in Regex.Matches(shape.TextFrame.TextRange.Text, @"@(\d+)@"))
                {
                    shape.TextFrame.TextRange.Find(match.Value).Text = "";
                }
            }
            if (Regex.IsMatch(shape.TextFrame.TextRange.Text, @"<\/?th>"))
            {
                foreach (Match match in Regex.Matches(shape.TextFrame.TextRange.Text, @"<\/?th>"))
                {
                    shape.TextFrame.TextRange.Find(match.Value).Text = "";
                }
            }
            if (Regex.IsMatch(shape.TextFrame.TextRange.Text, @"<\/?tl[^<>]*>"))
            {
                foreach (Match match in Regex.Matches(shape.TextFrame.TextRange.Text, @"<\/?tl[^<>]*>"))
                {
                    shape.TextFrame.TextRange.Find(match.Value).Text = "";
                }
            }
            if (Regex.IsMatch(shape.TextFrame.TextRange.Text, @"^[\s_]*%(\d+)%[\s_]*$"))
            {
                shape.TextFrame.TextRange.Text = "";
            }
            if (Regex.IsMatch(shape.TextFrame.TextRange.Text, @"%(\d+)%"))
            {
                foreach (Match match in Regex.Matches(shape.TextFrame.TextRange.Text, @"%(\d+)%"))
                {
                    shape.TextFrame.TextRange.Find(match.Value).Text = "";
                }
            }
            if (hasWpsProcessed && Regex.IsMatch(shape.TextFrame.TextRange.Text, @"&(\d+)&"))
            {
                foreach (Match match in Regex.Matches(shape.TextFrame.TextRange.Text, @"&(\d+)&"))
                {
                    shape.TextFrame.TextRange.Find(match.Value).Text = "";
                }
            }
            if (hasWpsProcessed && Regex.IsMatch(shape.TextFrame.TextRange.Text, @"<\/?zzd>"))
            {
                foreach (Match match in Regex.Matches(shape.TextFrame.TextRange.Text, @"<\/?zzd>"))
                {
                    shape.TextFrame.TextRange.Find(match.Value).Text = "";
                }
            }
            if (hasWpsProcessed && Regex.IsMatch(shape.TextFrame.TextRange.Text, @"<\/?bc>"))
            {
                foreach (Match match in Regex.Matches(shape.TextFrame.TextRange.Text, @"<\/?bc>"))
                {
                    shape.TextFrame.TextRange.Find(match.Value).Text = "";
                }
            }
            if (hasWpsProcessed && Regex.IsMatch(shape.TextFrame.TextRange.Text, @"<\/?ib>"))
            {
                foreach (Match match in Regex.Matches(shape.TextFrame.TextRange.Text, @"<\/?ib>"))
                {
                    shape.TextFrame.TextRange.Find(match.Value).Text = "";
                }
            }
        }

        // @description 删除包裹公式的标签
        // - 删除转成纯文本的公式标记不会有什么影响
        // - 删除答案中的 <m> 标记，可能会导致公式退出 inline mode 字号变大而产生偏移
        // - 删除标记后，只存在公式的文本，可能会有文本对齐方式设置无效的问题
        static public void DeleteMathZoneMark(Slide slide)
        {
            if (!Global.config.CompatibleWithWps)
            {
                return;
            }
            string r1 = @"@(\d+)@";
            string r2 = @"(<m>)|(</m>)|(<mathzone>)|(</mathzone>)";
            foreach (Shape shape in slide.Shapes)
            {
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    PpParagraphAlignment alignment = shape.TextFrame.TextRange.ParagraphFormat.Alignment;
                    if (Regex.IsMatch(shape.TextFrame.TextRange.Text, r1) || shape.Name.Contains("BD")) // 题干
                    {
                        if (Regex.IsMatch(shape.TextFrame.TextRange.Text, r2))
                        {
                            shape.TextFrame.TextRange.Text = Regex.Replace(shape.TextFrame.TextRange.Text, r2, "");
                        }
                    }
                    else if (shape.Name.Contains("AS") || shape.Name.Contains("EX")) // 解析，附加内容
                    {
                        if (Regex.IsMatch(shape.TextFrame.TextRange.Text, r2))
                        {
                            shape.TextFrame.TextRange.Text = Regex.Replace(shape.TextFrame.TextRange.Text, r2, "");
                        }
                    }
                    else if (shape.Name.Contains("AN")) // 答案
                    {
                        if (!CheckMatchPositionAnswer(shape))
                        {
                            if (Regex.IsMatch(shape.TextFrame.TextRange.Text, r2))
                            {
                                shape.TextFrame.TextRange.Text = Regex.Replace(shape.TextFrame.TextRange.Text, r2, "");
                            }
                        }
                    }
                }
                else if (shape.HasTable == MsoTriState.msoTrue)
                {
                    foreach (Row row in shape.Table.Rows)
                    {
                        foreach (Cell cell in row.Cells)
                        {
                            if (Regex.IsMatch(cell.Shape.TextFrame.TextRange.Text, r2))
                            {
                                cell.Shape.TextFrame.TextRange.Text = Regex.Replace(cell.Shape.TextFrame.TextRange.Text, r2, "");
                            }
                        }
                    }
                }
            }
        }

        static public void DeleteEmptySlide(Slide slide)
        {
            if (slide.Shapes.Count == 0)
            {
                if (!(
                    slide.CustomLayout.Name.Contains("Cover") ||
                    slide.CustomLayout.Name.Contains("BackCover") ||
                    slide.CustomLayout.Name.Contains("目录") ||
                    slide.CustomLayout.Name.Contains("宣传")
                )) // 避免误伤封面、封底、目录和宣传页
                {
                    slide.Delete();
                }
            }
        }

        // @@todo：建议全部删除标记相关的函数封装在 DeleteMark 里。
        // @description 删除辅助精排 Ppt 的标记和标签
        static public void DeleteMark(Slide slide)
        {
            foreach (Shape shape in slide.Shapes)
            {
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    DeletePositionMark(shape.Name, shape);
                }
                else if (shape.HasTable == MsoTriState.msoTrue)
                {
                    foreach (Row row in shape.Table.Rows)
                    {
                        foreach (Cell cell in row.Cells)
                        {
                            DeletePositionMark(shape.Name, cell.Shape);
                        }
                    }
                }
            }
        }

        // @description 删除 #b# 标记
        public void DeleteBMark(Slide slide)
        {
            if (!Global.config.CompatibleWithWps)
            {
                return;
            }
            Regex regExp = new Regex("#b#");
            foreach (Shape shape in slide.Shapes)
            {
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    TextRange textRange = shape.TextFrame.TextRange;
                    if (regExp.IsMatch(textRange.Text))
                    {
                        int l = 1;
                        while (l <= textRange.Paragraphs().Count)
                        {
                            if (textRange.Paragraphs(l, 1).Length <= 0 ||
                                textRange.Paragraphs(l, 1).Text == "\r" ||
                                textRange.Paragraphs(l, 1).Text.Contains("#b#"))
                            {
                                float lineHeight = -1;
                                int spaceCount = 0;
                                dynamic alignment = -1;
                                string fontName = "";
                                MsoTriState lineRuleWithin = MsoTriState.msoTrue;
                                MsoTriState fontUnderLine = MsoTriState.msoFalse;
                                if (l < textRange.Paragraphs().Count)
                                {
                                    TextRange nextP = textRange.Paragraphs(l + 1, 1);
                                    lineHeight = nextP.ParagraphFormat.SpaceWithin;
                                    lineRuleWithin = (lineHeight < 2) ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                                }
                                if (l < textRange.Paragraphs().Count)
                                {
                                    fontName = "";
                                    fontUnderLine = MsoTriState.msoFalse;
                                    foreach (TextRange spaceChar in textRange.Paragraphs(l + 1, 1).Characters())
                                    {
                                        if (spaceChar.Text != " ")
                                        {
                                            break;
                                        }
                                        if (spaceChar.Text == " " && fontName == "")
                                        {
                                            fontName = spaceChar.Font.Name;
                                            fontUnderLine = spaceChar.Font.Underline;
                                        }
                                        spaceCount++;
                                    }
                                }
                                if (l < textRange.Paragraphs().Count)
                                {
                                    TextRange nextP = textRange.Paragraphs(l + 1, 1);
                                    alignment = nextP.ParagraphFormat.Alignment;
                                }
                                // 删除 #b# 这一段！！！
                                textRange.Paragraphs(l, 1).Delete();
                                // @tips：这里要注意还原下一行的对齐方式。
                                if (alignment != -1)
                                {
                                    textRange.Paragraphs(l + 1, 1).ParagraphFormat.Alignment = alignment;
                                }
                                // @tips：这里要注意还原下一个 P 的段首空格，不知道为什么会被 PPT 删掉。
                                // @tips：这里要注意除字体外还要还原下划线样式。
                                if (spaceCount > 0 && textRange.Paragraphs(l, 1).Characters(1).Text != " ")
                                {
                                    for (int s = spaceCount; s >= 1; s--)
                                    {
                                        textRange.Paragraphs(l, 1).InsertBefore(" ");
                                    }
                                    // @tips：这里要进行完整的拼写，否则会出现 With 不到最新数据的情况。
                                    textRange.Paragraphs(l, 1).Characters(1, spaceCount).Font.Name = fontName;
                                    textRange.Paragraphs(l, 1).Characters(1, spaceCount).Font.Underline = fontUnderLine;
                                }
                                // @tips：这里要注意避免误伤其他行的行高。
                                if (lineHeight != -1)
                                {
                                    textRange.Paragraphs(l + 1, 1).ParagraphFormat.LineRuleWithin = lineRuleWithin;
                                    textRange.Paragraphs(l + 1, 1).ParagraphFormat.SpaceWithin = lineHeight;
                                }
                            }
                            else
                            {
                                l++;
                            }
                        }
                    }
                }
            }
        }

        // @description 删除没有内容的动画窗格
        static public void DeleteEmptyAnimation(Slide slide)
        {
            int i = 1;
            while (i <= slide.TimeLine.MainSequence.Count)
            {
                Effect e = slide.TimeLine.MainSequence[i];
                if (e.DisplayName.Trim().Length <= 0)
                {
                    e.Delete();
                    i--;
                }
                i++;
            }
        }

        static public void DeleteEmptyShape(Slide slide, bool deleteEmptyInlinePosition = false)
        {
            Dictionary<int, Shape> deleteShapeMap = new Dictionary<int, Shape>();
            Regex regExp = new Regex("^\\.?((&\\d+&)|\\s)+\\.?$");
            foreach (Shape shape in slide.Shapes)
            {
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    bool noNeedRemove = false;
                    // - 避免误删设置过线条样式的文本框
                    if (shape.Line.DashStyle != MsoLineDashStyle.msoLineDashStyleMixed)
                    {
                        noNeedRemove = true;
                    }
                    if (!noNeedRemove && string.IsNullOrWhiteSpace(shape.TextFrame.TextRange.Text))
                    {
                        deleteShapeMap.Add(deleteShapeMap.Count, shape);
                    }
                    // - 这里需要注意填空答案只有图片时，会出现空元素，在最后删除掉
                    // - 这里需要注意只包含行内图片的段落两边可能会有看不见的点
                    if (regExp.IsMatch(shape.TextFrame.TextRange.Text.Trim()) && deleteEmptyInlinePosition)
                    {
                        deleteShapeMap.Add(deleteShapeMap.Count, shape);
                    }
                }
                else if (shape.HasTable == MsoTriState.msoTrue && CheckEmptyTable(shape))
                {
                    deleteShapeMap.Add(deleteShapeMap.Count, shape);
                }
            }
            if (deleteShapeMap.Count > 0)
            {
                foreach (int key in deleteShapeMap.Keys)
                {
                    deleteShapeMap[key].Delete();
                }
            }
        }

        // @description 为 TextRange 调整字符间距
        public void HandleTextRangeSpacing(Shape shape, int RangeStart, int RangeEnd, string OpType, int Offset)
        {
            Regex RegExp = new Regex(@"\s*((@\d+@)|(&\d+&)|(<m>)|(</m>)|(<l>)|(</l>)|(<bc>)|(</bc>)|(<zzd>)|(</zzd>)|(<ib>)|(</ib>))\s*");
            List<TextRange2> Runs = new List<TextRange2>();
            TextRange textRange = shape.TextFrame.TextRange.Characters(RangeStart, RangeEnd - RangeStart + 1);
            TextRange2 textRange2 = shape.TextFrame2.TextRange.Characters[RangeStart, RangeEnd - RangeStart + 1];
            Dictionary<string, object[]> RunMap = new Dictionary<string, object[]>();
            if (textRange2.Font.Spacing < -11) // 避免因为有过 Spacing 的处理误伤导致整个段落的字间距异常
            {
                return;
            }
            // @tips：根据标记进行 Runs 的分割，避免误伤标记的字间距。
            if (RegExp.IsMatch(textRange.Text))
            {
                int PrevMatchStart;
                int PrevMatchEnd = RangeStart;
                foreach (Match Match in RegExp.Matches(textRange.Text))
                {
                    int MatchRangeStart = Match.Index + 1;
                    int MatchRangeEnd = MatchRangeStart + Match.Length - 1;
                    int RunRangeStart = PrevMatchEnd + 1;
                    int RunRangeEnd = MatchRangeStart - 1;
                    if (RunRangeEnd >= RunRangeStart)
                    {
                        int RunRangeLength = RunRangeEnd - RunRangeStart + 1;
                        Runs.Add(textRange2.Characters[RunRangeStart, RunRangeLength]);
                    }
                    PrevMatchStart = MatchRangeStart;
                    PrevMatchEnd = MatchRangeEnd;
                }
                if (textRange.Length > PrevMatchEnd)
                {
                    int RunRangeStart = PrevMatchEnd + 1;
                    int RunRangeEnd = textRange.Length;
                    int RunRangeLength = RunRangeEnd - RunRangeStart + 1;
                    Runs.Add(textRange2.Characters[RunRangeStart, RunRangeLength]);
                }
            }
            if (Runs.Count == 0)
            {
                Runs.Add(textRange2);
            }
            // @tips：先记录一遍每个 Run 的原始数据，避免公式多次加减。
            for (int i = 0; i < Runs.Count; i++)
            {
                for (int j = 0; j < Runs[i].Count; j++)
                {
                    if (Runs[i].Item(j).Font.Spacing > -11)
                    {
                        RunMap.Add(i + "#" + j, new object[] { Runs[i].Item(j), Runs[i].Item(j).Font.Spacing });
                    }
                }
            }
            for (int i = 0; i < Runs.Count; i++)
            {
                for (int j = 0; j < Runs[i].Count; j++)
                {
                    if (Runs[i].Item(j).Font.Spacing > -11)
                    {
                        if (RunMap.ContainsKey(i + "#" + j))
                        {
                            TextRange2 Range = (TextRange2)RunMap[i + "#" + j][0];
                            int Spacing = (int)RunMap[i + "#" + j][1];
                            if (Range.Font.Spacing == Spacing)
                            {
                                if (OpType == "PLUS")
                                {
                                    Range.Font.Spacing = Spacing + Offset;
                                }
                                else if (OpType == "MINUS")
                                {
                                    Range.Font.Spacing = Spacing - Offset;
                                }
                            }
                        }
                    }
                }
            }
        }

        static public string SafeFindShapeName(Shape shape)
        {
            try
            {
                return shape.Name;
            }
            catch (Exception)
            {
                return "";
            }
        }

        static public void AddQuery(Shape shape, string key, string value)
        {
            if (shape.Name.Contains("?"))
            {
                shape.Name += "&" + key + "=" + value;
            }
            else
            {
                shape.Name += "?" + key + "=" + value;
            }
        }
        static public void Sleep()
        {
            int waitTime = 1;
            DateTime start = DateTime.Now;
            while (DateTime.Now < start.AddSeconds(waitTime))
            {
                Thread.Sleep(1);
            }
        }

        static public string[] GetReturn(VbaReturn returnInfo)
        {
            return new string[] {
                returnInfo.ErrorInfo,
                returnInfo.PreviousSlidesCount.ToString(),
                returnInfo.CurrentSlidesCount.ToString(),
                returnInfo.FileSuffix,
                returnInfo.ErrorType.ToString()
            };
        }
        
        static public bool IsShapeOverflowing(Slide slide, Shape shape)
        {
             // 获取版心的边界
             double viewLeft = GetViewLeft();
             double viewTop = GetRealViewTop(slide);
             double viewRight = GetViewRight();
             double viewBottom = GetRealViewBottom(slide);
            
             // 获取shape的边界
             double shapeLeft = shape.Left;
             double shapeTop = shape.Top;
             double shapeRight, shapeBottom;

             // 如果shape是文本框，则获取文本内容的实际宽度和高度
             if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoTrue)
             {
                 TextRange textRange = shape.TextFrame.TextRange;
                 shapeRight = shapeLeft + textRange.BoundWidth;
                 shapeBottom = shapeTop + textRange.BoundHeight;
             }
             else
             {
                 shapeRight = shapeLeft + shape.Width;
                 shapeBottom = shapeTop + shape.Height;
             }
             // 检查shape是否溢出版心
             bool isOverflowingView = 
                 (shapeLeft < viewLeft) ||
                 (shapeTop < viewTop) ||
                 (shapeRight > viewRight) ||
                 (shapeBottom > viewBottom);
            
                return isOverflowingView;
        }
    }
}
