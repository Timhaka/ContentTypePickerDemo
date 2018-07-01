using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ContentTypePickerWeb.Controllers
{
    public class DocumentController : Controller
    {
        // GET: Document
        public ActionResult DocumentIndex(string SPHostUrl)
        {
            ViewBag.SPHostUrl = SPHostUrl;
            return View();
        }
        public ActionResult InsertParagraph(string SPHostUrl)
        {
            ViewBag.SPHostUrl = SPHostUrl;
            return View();
        }
        public ActionResult EditHeader(string SPHostUrl)
        {
            ViewBag.SPHostUrl = SPHostUrl;
            return View();
        }
        public ActionResult EditFooter(string SPHostUrl)
        {
            ViewBag.SPHostUrl = SPHostUrl;
            return View();
        }
    }
}