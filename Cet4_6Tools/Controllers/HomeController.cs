using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Saving;
using Cet4_6Tools.Models;
using Microsoft.AspNetCore.Mvc;
using System.Text.RegularExpressions;

namespace Cet4_6Tools.Controllers
{
    public class HomeController : Controller
    {
        private static string? FolderName { get; set; } = "143821";

        // GET: Home

        #region 系统首页
        public IActionResult Index()
        {
            return View();
        }
        #endregion

        #region 下载模板
        [HttpGet]
        public IActionResult DownloadTemplate()
        {
            return base.File(Server.MapPath("~/Templates/Template.doc"), "application/vnd.ms-word", "Template.doc");
        }
        #endregion

        #region 上传模板
        [HttpPost]
        public async Task<IActionResult> UploadTemplate(IFormFile tmp_file)
        {
            if (tmp_file == null || tmp_file.Length <= 0)
                return Json(new { status = 0, msg = "上传文件为空！" });
            var fileName = tmp_file.FileName.Replace("Template-", "");
            FolderName = Path.GetFileNameWithoutExtension(fileName);
            if (string.IsNullOrEmpty(fileName))
                return Json(new { status = 0, msg = "上传文件为空！" });
            var tempPath = Path.Combine(Server.MapPath("~/Templates"), $"Template-{fileName}");
            using var fileStream = new FileStream(tempPath, FileMode.Create);
            await tmp_file.CopyToAsync(fileStream);
            return Json(new { status = 1, msg = $"/Templates/Template-{fileName}" });
        }
        #endregion

        #region 上传电子照
        [HttpPost]
        public async Task<IActionResult> UploadImages(IFormFileCollection img_files)
        {
            if (img_files == null || img_files.Count <= 0)
                return Json(new { status = 0, msg = "上传文件为空！" });
            if (string.IsNullOrEmpty(FolderName))
                return Json(new { status = 0, msg = "请先上传模板！" });
            try
            {
                var basePath = Server.MapPath($"~/Digital/{FolderName}");
                if (!Directory.Exists(basePath))
                    Directory.CreateDirectory(basePath);

                foreach (var image in img_files)
                {
                    if (image.Length > 0)
                    {
                        var imagePath = Path.Combine(basePath, image.FileName);
                        using var fileStream = new FileStream(imagePath, FileMode.Create);
                        await image.CopyToAsync(fileStream);
                    }
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
        public IActionResult Export()
        {
            string tempPath = Server.MapPath($"~/Templates/Template-{FolderName}.doc");
            var doc = new Document(tempPath); //载入模板
            var images = new List<string>();
            var root = new DirectoryInfo(Server.MapPath($"~/Digital/{FolderName}"));
            foreach (var f in root.GetFiles())
            {
                images.Add(f.FullName);
            }

            for (int i = 1; i < images.Count + 1; i++)
            {
                doc.Range.Replace($"«Number{i}»", Path.GetFileNameWithoutExtension(images[i - 1]), new FindReplaceOptions() { MatchCase = false, FindWholeWordsOnly = false });
                doc.Range.Replace(new Regex($"Photo{i}&"), "", new FindReplaceOptions() { ReplacingCallback = new ReplaceAndInsertImage(images[i - 1]) });
            }

            var docStream = new MemoryStream();
            doc.Save(docStream, SaveOptions.CreateSaveOptions(SaveFormat.Doc));
            return base.File(docStream.ToArray(), "application/msword", $"{FolderName}班(照片采集)-已完成" + ".doc");
        }
        #endregion

    }
}