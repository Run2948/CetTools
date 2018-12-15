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
        // GET: Home
        public ActionResult Index()
        {
            string tempPath = Server.MapPath("~/Templates/Template-163821.doc");
            var doc = new Document(tempPath); //载入模板
            List<string> images = new List<string>();
            DirectoryInfo root = new DirectoryInfo(Server.MapPath("~/Digital/163821"));
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
            return base.File(docStream.ToArray(), "application/msword", "16计本1班(照片采集)-已完成" + ".doc");
        }

        public ActionResult Test()
        {
            return View();
        }


    }
}