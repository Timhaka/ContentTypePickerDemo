using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ContentTypePickerWeb.Controllers
{
    public class ContentTypeController : Controller
    {
        [SharePointContextFilter]
        // GET: ContentType
        public ActionResult Index(string SPHostUrl)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            using (var ctx = spContext.CreateUserClientContextForSPHost())
            {
                //Hämtar alla content types för listan
                List documentList = ctx.Web.Lists.GetByTitle("MVCDemo");
                ContentTypeCollection collectionOfContentTypes = documentList.ContentTypes;

                ctx.Load(collectionOfContentTypes);
                ctx.ExecuteQuery();
                string s = "hey";

                ViewBag.SPHostUrl = SPHostUrl;
                return View(collectionOfContentTypes);
            }
        }

        public ActionResult GetFields(string ContentTypeId, string SPHostUrl)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            using (var ctx = spContext.CreateUserClientContextForSPHost())
            {
                //Hämtar alla content types för listan
                List list = ctx.Web.Lists.GetByTitle("MVCDemo");
                ContentType contentTypeToEdit = list.ContentTypes.GetById(ContentTypeId);
                ctx.Load(list);
                ctx.Load(contentTypeToEdit, include => include.Fields, include => include.Name);
                ctx.ExecuteQuery();


                TermStore store = ctx.Site.GetDefaultSiteCollectionTermStore();
                TermSet termSet = store.GetTermSet("95AA3BC4-2790-477E-8661-872A83143AE0".ToGuid());

                ctx.Load(store);
                //ctx.Load(group);
                ctx.Load(termSet, include => include.Terms);
                ctx.ExecuteQuery();

                List<string> termList = new List<string>();

                foreach (Term term in termSet.Terms)
                {
                    termList.Add(term.Name);
                }

                ViewBag.Term = termList;

                FieldCollection fields = contentTypeToEdit.Fields;
                ViewBag.SelectedCT = ContentTypeId;
                ViewBag.SPHostUrl = SPHostUrl;
                return View(fields);
            }
        }

        public ActionResult UpdateInformation(string SelectedCT, string SPHostUrl, List<string> newValue)
        {

            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);


            using (var ctx = spContext.CreateUserClientContextForSPHost())
            {
                List list = ctx.Web.Lists.GetByTitle("MVCDemo");
                ContentType contentTypeToEdit = list.ContentTypes.GetById(SelectedCT);
                ctx.Load(list);
                ctx.Load(contentTypeToEdit, include => include.Fields, include => include.Name);
                ctx.ExecuteQuery();

                FieldCollection fields = contentTypeToEdit.Fields;

                

                List<KeyValuePair<string, string>> newProduct = new List<KeyValuePair<string, string>>();


                var a = 0;
                for (int i = 0; i < fields.Count; i++)
                {

                    if (fields[i].Group == "MyCT")
                    {

                        newProduct.Add(new KeyValuePair<string, string>(fields[i].InternalName, newValue[a]));
                        a++;
                    }
                }

                string json = JsonConvert.SerializeObject(newProduct);
                ViewBag.JsonContentTypeList = json;
                ViewBag.SPHostUrl = SPHostUrl;
                return View();
            }
        }














            }
}