using Aspose.Words;
using Aspose.Words.Saving;
using Cet4_6Tools.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;

namespace Cet4_6Tools.Controllers
{
    public class HistoryController : Controller
    {
        // GET: History

        #region 历史页面
        public ActionResult Index()
        {
            List<Clazz> list = new List<Clazz>();
            DirectoryInfo root = new DirectoryInfo(Server.MapPath("~/Digital"));
            foreach (DirectoryInfo d in root.GetDirectories())
            {
                var folder = d.Name;
                DirectoryInfo son = new DirectoryInfo(Server.MapPath($"~/Digital/{folder}"));
                list.Add(new Clazz(folder,son.GetFiles().Length));
            }
            return View(list);
        }
        #endregion

        #region 导出历史
        [HttpGet]
        public ActionResult Show(string folder)
        {
            string tempPath = Server.MapPath($"~/Templates/Template-{folder}.doc");
            var doc = new Document(tempPath); //载入模板
            List<string> images = new List<string>();
            DirectoryInfo root = new DirectoryInfo(Server.MapPath($"~/Digital/{folder}"));
            foreach (FileInfo f in root.GetFiles())
            {
                images.Add(f.FullName);
            }

            for (int i = 1; i < images.Count + 1; i++)
            {
                doc.Range.Replace($"«Number{i}»", Path.GetFileNameWithoutExtension(images[i - 1]), false, false);
                doc.Range.Replace(new Regex($"Photo{i}&"), new ReplaceAndInsertImage(images[i - 1]), false);
            }

            var docStream = new MemoryStream();
            doc.Save(docStream, SaveOptions.CreateSaveOptions(SaveFormat.Doc));
            return base.File(docStream.ToArray(), "application/msword", $"{folder}班(照片采集)-已完成" + ".doc");
        } 
        #endregion
    }
}