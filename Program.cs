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
using System.Security.Policy;


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
                Tester.Log("-1001：pptx 文件不存在，直接退出");
                return;
            }

            Global.app = new PowerPoint.Application { Visible = MsoTriState.msoTrue };
            Global.app.Presentations.Open(pptxPath, MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue);

            // 全局变量
            Global.slideWidth = Global.app.ActivePresentation.PageSetup.SlideWidth;
            Global.slideHeight = Global.app.ActivePresentation.PageSetup.SlideHeight;

            // 机器质检
            Tester.Test();

            // 关闭文件
            Global.app.ActivePresentation.Close();
            Global.app.Quit();
            return;
        }
    }

    internal class Tester
    {
        // -1xxx：系统问题
        // -1001：pptx 文件不存在，直接退出
        // @todo：-1002：程序崩溃

        // -2xxx：PPT 问题

        // -3xxx：单页问题

        // -31xx：分页问题
        // -3101：页面不能太空
        // -3102：页面不能存在单行一页的情况

        // -32xx：元素问题

        // -33xx：行问题
        // -3301：标点符号不能在行首

        public static void Test()
        {
            foreach (Slide slide in Global.app.ActivePresentation.Slides)
            {
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
                                foreach (TextRange line in cell.Shape.TextFrame.TextRange.Lines())
                                {
                                    TestLine(line, cell.Shape, shape);
                                }
                            }
                        }
                    }
                    else if (shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        TestShape(shape, shape);
                        foreach (TextRange line in shape.TextFrame.TextRange.Lines())
                        {
                            TestLine(line, shape, shape);
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
            List<Shape> sortedShaps = Utils.GetSortedSlideShapes(slide);
            bool isTitlePage = Utils.CheckTitlePage(slide);
            if (!isTitlePage && Utils.ComputeContentBottom(sortedShaps) <= Global.slideHeight * 1 / 3)
            {
                Log("-3101：" + slide.SlideIndex + "#页面不能太空");
            }
            if (!isTitlePage && sortedShaps.Count == 1)
            {
                Shape shape = sortedShaps[0];
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    if (shape.TextFrame.TextRange.Lines().Count == 1)
                    {
                        Log("-3102：" + slide.SlideIndex + "#页面不能存在单行一页的情况");
                    }
                }
            }
        }

        public static void TestShape(Shape shape, Shape containerShape) { }

        public static void TestLine(TextRange line, Shape shape, Shape containerShape)
        {
            Slide slide = containerShape.Parent;
            if (Regex.IsMatch(line.Text, @"^\s*[!),.:;?\]、。—ˇ¨〃々～‖…’”〕〉》」』〗】∶！＇），．：；？］｀｜｝]") &&
                !Regex.IsMatch(line.Text, @"(^\.%\d+%)|(^\.&\d+&)"))
            {
                Log("-3301：" + slide.SlideIndex + "#标点符号不能在行首");
            }
        }

        public static void Log(string msg)
        {
            Console.WriteLine(msg); // @todo：可以支持写入临时文件
        }
    }
}
