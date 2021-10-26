using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SearchTool
{
    /// <summary>
    /// 图片文字识别响应Model
    /// </summary>
    public class OcrResModel
    {
        /// <summary>
        /// 检测到的文本信息，包括文本行内容、置信度、文本行坐标以及文本行旋转纠正后的坐标
        /// </summary>
        public List<TextDetection> TextDetections { get; set; }

        /// <summary>
        /// 图片旋转角度（角度制），文本的水平方向为0°；顺时针为正，逆时针为负
        /// </summary>
        public float Angel { get; set; }

        /// <summary>
        /// 唯一请求 ID，每次请求都会返回
        /// </summary>
        public string RequestId { get; set; }
    }

    public class TextDetection
    {
        /// <summary>
        /// 识别出的文本行内容
        /// </summary>
        public string DetectedText { get; set; }

        /// <summary>
        /// 置信度 0 ~100
        /// </summary>
        public int Confidence { get; set; }

        /// <summary>
        /// 文本行坐标，以四个顶点坐标表示 注意：此字段可能返回 null，表示取不到有效值。
        /// </summary>
        public List<Coord>? Polygon { get; set; }

        /// <summary>
        /// 此字段为扩展字段。GeneralBasicOcr接口返回段落信息Parag，包含ParagNo
        /// </summary>
        public string AdvancedInfo { get; set; }

        /// <summary>
        /// 文本行在旋转纠正之后的图像中的像素坐标，表示为（左上角x, 左上角y，宽width，高height）
        /// </summary>
        public ItemCoord ItemPolygon { get; set; }

        /// <summary>
        /// 识别出来的单字信息包括单字（包括单字Character和单字置信度confidence）
        /// </summary>
        public List<DetectedWords> Words { get; set; }

        /// <summary>
        /// 单字在原图中的四点坐标
        /// </summary>
        public List<DetectedWordCoordPoint> WordCoordPoint { get; set; }
    }

    /// <summary>
    /// 坐标
    /// </summary>
    public class Coord
    {
        /// <summary>
        /// 横坐标
        /// </summary>
        public int X { get; set; }

        /// <summary>
        /// 	纵坐标
        /// </summary>
        public int Y { get; set; }
    }

    /// <summary>
    /// 文本行在旋转纠正之后的图像中的像素坐标，表示为（左上角x, 左上角y，宽width，高height）
    /// </summary>
    public class ItemCoord
    {
        /// <summary>
        /// 左上角x
        /// </summary>
        public int X { get; set; }

        /// <summary>
        /// 左上角y
        /// </summary>
        public int Y { get; set; }

        /// <summary>
        /// 宽width
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// 高height
        /// </summary>
        public int Height { get; set; }
    }

    /// <summary>
    /// 识别出来的单字信息包括单字（包括单字Character和单字置信度confidence）
    /// </summary>
    public class DetectedWords
    {
        /// <summary>
        /// 置信度 0 ~100
        /// </summary>
        public int Confidence { get; set; }

        /// <summary>
        /// 候选字Character
        /// </summary>
        public string Character { get; set; }
    }

    /// <summary>
    /// 单字在原图中的四点坐标
    /// </summary>
    public class DetectedWordCoordPoint
    {
        /// <summary>
        /// 单字在原图中的坐标，以四个顶点坐标表示，以左上角为起点，顺时针返回
        /// </summary>
        public List<Coord> WordCoordinate { get; set; }
    }
}
