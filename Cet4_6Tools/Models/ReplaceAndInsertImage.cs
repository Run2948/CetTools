using Aspose.Words;
using Aspose.Words.Replacing;

namespace Cet4_6Tools.Models
{
    public class ReplaceAndInsertImage : IReplacingCallback
    {
        /// <summary>
        /// 需要插入的图片路径
        /// </summary>
        public string url { get; set; }

        public ReplaceAndInsertImage(string url)
        {
            this.url = url;
        }

        public ReplaceAction Replacing(ReplacingArgs e)
        {
            //获取当前节点
            var node = e.MatchNode;
            //获取当前文档
            var doc = node.Document as Document;
            var builder = new DocumentBuilder(doc);
            //将光标移动到指定节点
            builder.MoveTo(node);
            //插入图片
            var shape = builder.InsertImage(url);
            // 设置x,y坐标和高宽.
            shape.Left = 0;
            shape.Top = 20;
            shape.Width = 60 * 1.2;
            shape.Height = 80 * 1.2;

            return ReplaceAction.Replace;
        }
    }
}