//using Microsoft.SharePoint.Client;
//using Microsoft.SharePoint.Client.Taxonomy;
//using Newtonsoft.Json;
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Web;
//using System.Web.Mvc;

//namespace ContentTypePickerWeb.Controllers
//{
//    public class HomeController : Controller
//    {
//        [SharePointContextFilter]
//        public ActionResult Index(string SPHostUrl)
//        {
//            User spUser = null;

//            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

//            using (var ctx = spContext.CreateUserClientContextForSPHost())
//            {
//                if (ctx != null)
//                {
//                    spUser = ctx.Web.CurrentUser;

//                    ctx.Load(spUser, user => user.Title);

//                    ctx.ExecuteQuery();

//                    ViewBag.UserName = spUser.Title;
//                }

//                List list = ctx.Web.Lists.GetByTitle("karlshamn");
//                ContentTypeCollection listOfContentTypes = list.ContentTypes;

//                //ctx.Load(list, include => include.Fields);
//                ctx.Load(listOfContentTypes);
//                ctx.ExecuteQuery();

//                ViewBag.spHost = SPHostUrl;


//                TermStore store = ctx.Site.GetDefaultSiteCollectionTermStore();
//                //TermGroup group = store.GetTermGroupByName("Qian_OfficeAddin");
//                TermSet termSet = store.GetTermSet("08b6e1c6-d9ac-4014-bcbf-6e23987fb5c5".ToGuid());


//                ctx.Load(store);
//                //ctx.Load(group);
//                ctx.Load(termSet, include => include.Terms);
//                ctx.ExecuteQuery();






//                //List<string> fields = new List<string>();

//                //foreach (ContentType item in listOfContentTypes)
//                //{
//                //    if (item.Group == "OfficeAddin CT")
//                //    {
//                //        ctx.Load(item.Fields);
//                //        ctx.ExecuteQuery();




//                //        foreach (Field test in item.Fields)
//                //        {
//                //            if (test.Group == "OfficeAddin Fields")
//                //            {
//                //                fields.Add(test.Title);
//                //            }
//                //        }
//                //    }
//                //}
//                //way to update the content type!
//                //ListItem itemToUpdate = list.GetItemById(8);
//                //ctx.Load(itemToUpdate, include => include["ContentTypeId"]);
//                //ctx.ExecuteQuery();
//                //string var = itemToUpdate["ContentTypeId"].ToString();
//                //itemToUpdate["ContentTypeId"] = listOfContentTypes[0].Id;
//                //string s = listOfContentTypes[0].Name;
//                //itemToUpdate.Update();
//                //ctx.ExecuteQuery();
//                //string var2 = itemToUpdate["ContentTypeId"].ToString();

//                //List <ContentType> ContentTypeList = new List<ContentType>();


//                //foreach (ContentType item in listOfContentTypes)
//                //{
//                //    if (item.Group == "OfficeAddin CT")
//                //    {
//                //        ContentTypeList.Add(item);
//                //    }
//                //}

//                //ViewBag.fieldList = ContentTypeList;


//                // string json = JsonConvert.SerializeObject(ContentTypeList);

//                //ViewBag.JsonContentTypeList = json;

//                return View(listOfContentTypes);
//            }


//        }

//        public ActionResult GetFields(string SelectedCT, string SPHostUrl)
//        {

//            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

//            using (var ctx = spContext.CreateUserClientContextForSPHost())
//            {

//                List list = ctx.Web.Lists.GetByTitle("karlshamn");
//                ContentType contentTypeToEdit = list.ContentTypes.GetById(SelectedCT);
//                ctx.Load(list);
//                ctx.Load(contentTypeToEdit, include => include.Fields, include => include.Name);
//                ctx.ExecuteQuery();

//                TermStore store = ctx.Site.GetDefaultSiteCollectionTermStore();
//                //TermGroup group = store.GetTermGroupByName("Qian_OfficeAddin");
//                TermSet termSet = store.GetTermSet("08b6e1c6-d9ac-4014-bcbf-6e23987fb5c5".ToGuid());


//                ctx.Load(store);
//                //ctx.Load(group);
//                ctx.Load(termSet, include => include.Terms);
//                ctx.ExecuteQuery();


//                List<string> termList = new List<string>();

//                foreach (Term term in termSet.Terms)
//                {
//                    termList.Add(term.Name);
//                }


//                ViewBag.Term = termList;
//                ViewBag.Term1 = "Meh";






//                List<string> fieldList = new List<string>();
//                foreach (Field item in contentTypeToEdit.Fields)
//                {

//                    if (item.Group == "Karlshamn Fields")
//                    {
//                        fieldList.Add(item.Title);
//                    }


//                }

//                FieldCollection fields = contentTypeToEdit.Fields;

//                ViewBag.Title = contentTypeToEdit.Name;

//                ViewBag.SelectedCT = SelectedCT;
//                ViewBag.SPHostUrl = SPHostUrl;

//                return View(fields);
//            }






//        }
//        [HttpPost]
//        public ActionResult GetFields(string SelectedCT, string SPHostUrl, List<string> newValue, string newTax, string GDPR, string newNumber, string newUrl, string newNote)
//        {

//            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

//            using (var ctx = spContext.CreateUserClientContextForSPHost())
//            {
//                List list = ctx.Web.Lists.GetByTitle("karlshamn");
//                ContentTypeCollection listOfContentTypes = list.ContentTypes;
//                ctx.Load(list);
//                ctx.Load(listOfContentTypes);
//                ctx.ExecuteQuery();





//                List<KeyValuePair<string, string>> newProduct = new List<KeyValuePair<string, string>>();

//                string s = listOfContentTypes[0].Id.ToString();
//                string name = listOfContentTypes[0].Name;

//                string n = listOfContentTypes[1].Id.ToString();
//                string name2 = listOfContentTypes[1].Name;

//                //AddinCT
//                if (SelectedCT == "n")
//                {
//                    //Title
//                    //Adress
//                    //logo
//                    //main contact
//                    //office phone


//                    //ctx.ExecuteQuery();
//                    ListItem itemToUpdate = list.GetItemById(33);
//                    ctx.Load(itemToUpdate, include => include["ContentTypeId"]);
//                    //itemToUpdate.File.CheckIn("Checkin", CheckinType.MajorCheckIn);
//                    ctx.ExecuteQuery();

//                    //itemToUpdate["ContentTypeId"] = listOfContentTypes[1].Id;
//                    //itemToUpdate.Update();
//                    //ctx.ExecuteQuery();


//                    FieldUrlValue FieldImage = new FieldUrlValue();

//                    FieldImage.Url = newUrl;
//                    FieldImage.Description = newUrl;

//                    //itemToUpdate["KARLSHAMN_Title"] = newText[0];
//                    itemToUpdate["KARLSHAMN_Adress"] = newNote;
//                    itemToUpdate["KARLSHAMN_Logo"] = FieldImage;
//                    //itemToUpdate["KARLSHAMN_MainContactPerson"] = newText[1];
//                    itemToUpdate["KARLSHAMN_OfficePhone"] = newNumber;

//                    itemToUpdate.Update();
//                    ctx.ExecuteQuery();



//                }
//                //officeAddinCT
//                else if (SelectedCT == s)
//                {

//                    ContentType contentTypeToEdit = list.ContentTypes.GetById(SelectedCT);
//                    ctx.Load(list);
//                    ctx.Load(contentTypeToEdit, include => include.Fields, include => include.Name);
//                    ctx.ExecuteQuery();

//                    FieldCollection fields = contentTypeToEdit.Fields;

//                    var a = 0;
//                    for (int i = 0; i < fields.Count; i++)
//                    {

//                        if (fields[i].Group == "Karlshamn Fields")
//                        {

//                            newProduct.Add(new KeyValuePair<string, string>(fields[i].Title, newValue[a]));
//                            a++;
//                        }
//                    }


//                    Console.WriteLine(newProduct);

//                    string json = JsonConvert.SerializeObject(newProduct);

//                    ViewBag.JsonContentTypeList = json;


//                    //ListItem itemToUpdate = list.GetItemById(38);
//                    //ctx.Load(itemToUpdate, include => include["ContentTypeId"], include => include.File.LockedByUser);
//                    //ctx.ExecuteQuery();
//                    ////itemToUpdate.File.CheckIn("Hey", CheckinType.MajorCheckIn);

//                    //var se = list.FileSavePostProcessingEnabled == true;
//                    //se = true;

//                    ////itemToUpdate["ContentTypeId"] = listOfContentTypes[0].Id;
//                    ////itemToUpdate.Update();
//                    ////ctx.ExecuteQuery();
//                    //File k = itemToUpdate.File;
//                    //ctx.Load(k);
//                    //var s4 = k.ListItemAllFields;
//                    //ctx.Load(s4);
//                    //ctx.ExecuteQuery();
//                    //Console.WriteLine(s4);
//                    //var stst = s4.FieldValues;

//                    //foreach (var item in stst)
//                    //{
//                    //    if (item.Key == "KARLSHAMN_Adress")
//                    //    {
//                    //        var a2 = item.Value;
//                    //    }
//                    //    Console.WriteLine(item.Value);

//                    //}


//                    ////itemToUpdate["KARLSHAMN_Adress"] = newText[0];



//                    //itemToUpdate.Update();
//                    //ctx.ExecuteQuery();

//                }
//                return View();

//                // return RedirectToAction("UpdateInformation", new {newProduct = newProduct, SPHostUrl = SPHostUrl });

//            }
//        }

//        //public ActionResult UpdateInformation(List<KeyValuePair<string, string>> newProduct, string SPHostUrl)
//        //{
//        //    string json = JsonConvert.SerializeObject(newProduct);

//        //    ViewBag.JsonContentTypeList = json;

//        //    return View("UpdateInformation");
//        //}


//        public ActionResult UpdateInformation(string SelectedCT, string SPHostUrl, List<string> newValue)
//        {

//            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);


//            using (var ctx = spContext.CreateUserClientContextForSPHost())
//            {
//                List list = ctx.Web.Lists.GetByTitle("karlshamn");
//                ContentTypeCollection listOfContentTypes = list.ContentTypes;
//                ctx.Load(list);
//                ctx.Load(listOfContentTypes);
//                ctx.ExecuteQuery();


//                List<KeyValuePair<string, string>> newProduct = new List<KeyValuePair<string, string>>();

//                string s = listOfContentTypes[0].Id.ToString();
//                string name = listOfContentTypes[0].Name;

//                string n = listOfContentTypes[1].Id.ToString();
//                string name2 = listOfContentTypes[1].Name;

//                //AddinCT
//                if (SelectedCT == "n")
//                {

//                }
//                //officeAddinCT
//                else if (SelectedCT == s)
//                {

//                    ContentType contentTypeToEdit = list.ContentTypes.GetById(SelectedCT);
//                    ctx.Load(list);
//                    ctx.Load(contentTypeToEdit, include => include.Fields, include => include.Name);
//                    ctx.ExecuteQuery();

//                    FieldCollection fields = contentTypeToEdit.Fields;

//                    var a = 0;
//                    for (int i = 0; i < fields.Count; i++)
//                    {

//                        if (fields[i].Group == "Karlshamn Fields")
//                        {

//                            newProduct.Add(new KeyValuePair<string, string>(fields[i].Title, newValue[a]));
//                            a++;
//                        }
//                    }


//                    Console.WriteLine(newProduct);

//                    string json = JsonConvert.SerializeObject(newProduct);

//                    ViewBag.JsonContentTypeList = json;


//                    //ListItem itemToUpdate = list.GetItemById(38);
//                    //ctx.Load(itemToUpdate, include => include["ContentTypeId"], include => include.File.LockedByUser);
//                    //ctx.ExecuteQuery();
//                    ////itemToUpdate.File.CheckIn("Hey", CheckinType.MajorCheckIn);

//                    //var se = list.FileSavePostProcessingEnabled == true;
//                    //se = true;

//                    ////itemToUpdate["ContentTypeId"] = listOfContentTypes[0].Id;
//                    ////itemToUpdate.Update();
//                    ////ctx.ExecuteQuery();
//                    //File k = itemToUpdate.File;
//                    //ctx.Load(k);
//                    //var s4 = k.ListItemAllFields;
//                    //ctx.Load(s4);
//                    //ctx.ExecuteQuery();
//                    //Console.WriteLine(s4);
//                    //var stst = s4.FieldValues;

//                    //foreach (var item in stst)
//                    //{
//                    //    if (item.Key == "KARLSHAMN_Adress")
//                    //    {
//                    //        var a2 = item.Value;
//                    //    }
//                    //    Console.WriteLine(item.Value);

//                    //}


//                    ////itemToUpdate["KARLSHAMN_Adress"] = newText[0];



//                    //itemToUpdate.Update();
//                    //ctx.ExecuteQuery();

//                }


//                return View();

//            }
//        }

//        public ActionResult ConnectToIssue(string SPHostUrl)
//        {
//            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

//            using (var ctx = spContext.CreateUserClientContextForSPHost())
//            {




//                return View();
//            }
//        }
//        [HttpPost]
//        public ActionResult ConnectToIssue(string issueTitle, string issueDescription, string assignTo, string status, string priority, string SPHostUrl)
//        {

//            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

//            using (var ctx = spContext.CreateUserClientContextForSPHost())
//            {

//                User spUser = null;

//                List list = ctx.Web.Lists.GetByTitle("Karlshamn Issue");
//                ctx.Load(list);
//                ctx.ExecuteQuery();

//                UserCollection users = ctx.Web.SiteUsers;
//                ctx.Load(users);
//                ctx.ExecuteQuery();


//                foreach (User user in users)
//                {
//                    ctx.Load(user, include => include.Email);
//                    ctx.ExecuteQuery();
//                    if (user.Email.ToLower() == assignTo.ToLower())
//                    {
//                        spUser = user;
//                    }

//                }



//                var item = list.AddItem(new ListItemCreationInformation());

//                item["Title"] = issueTitle;
//                item["Comment"] = issueDescription;
//                item["AssignedTo"] = spUser;
//                //item["Status"] = status;
//                //item["Priority"] = priority;


//                item.Update();
//                ctx.ExecuteQuery();

//                var testing = list.GetItemById(1);

//                ctx.Load(testing, include => include["AssignedTo"]);
//                ctx.ExecuteQuery();

//                Console.WriteLine(testing);

//                var usersss = testing["AssignedTo"];

//                return View();
//            }
//        }
//        public ActionResult About()
//        {
//            ViewBag.Message = "Your application description page.";

//            return View();
//        }

//        public ActionResult Contact()
//        {
//            ViewBag.Message = "Your contact page.";

//            return View();
//        }
//    }
//}
