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
    public class HomeController : Controller
    {
        private static string FolderName { get; set; }

        // GET: Home

        #region 系统首页
        public ActionResult Index()
        {
            return View();
        } 
        #endregion

        #region 下载模板
        [HttpGet]
        public ActionResult DownloadTemplate()
        {
            return base.File(Server.MapPath("~/Templates/Template.doc"), "application/vnd.ms-word", "Template.doc");
        }
        #endregion

        #region 上传模板
        [HttpPost]
        public ActionResult UploadTemplate()
        {
            HttpPostedFileBase file = Request.Files["tmp_file"];
            if (file == null)
                return Json(new { status = 0, msg = "上传文件为空！" });
            var fileName = file.FileName;
            FolderName = Path.GetFileNameWithoutExtension(fileName);
            if (string.IsNullOrEmpty(fileName))
                return Json(new { status = 0, msg = "上传文件为空！" });
            file.SaveAs(Path.Combine(Server.MapPath("~/Templates"), $"Template-{fileName}"));
            return Json(new { status = 1, msg = $"/Templates/Template-{fileName}" });
        }
        #endregion

        #region 上传电子照
        [HttpPost]
        public ActionResult UploadImages()
        {
            HttpFileCollectionBase files = HttpContext.Request.Files;
            if (files == null)
                return Json(new { status = 0, msg = "上传文件为空！" });
            if (string.IsNullOrEmpty(FolderName))
                return Json(new { status = 0, msg = "请先上传模板！" });
            try
            {
                foreach (string key in files.Keys)
                {
                    HttpPostedFileBase image = files[key];
                    var basePath = Server.MapPath($"~/Digital/{FolderName}");
                    if (!Directory.Exists(basePath))
                        Directory.CreateDirectory(basePath);
                    image.SaveAs(Path.Combine(basePath, image.FileName));
                }
            }
            catch (Exception e)
            {
                return Json(new { status = 1, msg = e.Message });
            }
            return Json(new { status = 1, msg = "上传成功！" });
        }
        #endregion

        #region 导出文档
        [HttpGet]
        public ActionResult Export()
        {
            string tempPath = Server.MapPath($"~/Templates/Template-{FolderName}.doc");
            var doc = new Document(tempPath); //载入模板
            List<string> images = new List<string>();
            DirectoryInfo root = new DirectoryInfo(Server.MapPath($"~/Digital/{FolderName}"));
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
            return base.File(docStream.ToArray(), "application/msword", $"{FolderName}班(照片采集)-已完成" + ".doc");
        }
        #endregion

    }
}