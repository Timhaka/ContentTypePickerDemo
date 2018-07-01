using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ContentTypePickerWeb.Controllers
{
    public class IssueController : Controller
    {
        // GET: Issue
        public ActionResult ConnectToIssue(string SPHostUrl)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            using (var ctx = spContext.CreateUserClientContextForSPHost())
            {

                ViewBag.SPHostUrl = SPHostUrl;

                return View();
            }

        }


        [HttpPost]
        public ActionResult ConnectToIssue(string issueTitle, string issueDescription, string assignTo, string documentName,string documentUrl, string status, string priority, string SPHostUrl)
        {

            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            using (var ctx = spContext.CreateUserClientContextForSPHost())
            {

                User spUser = null;

                List list = ctx.Web.Lists.GetByTitle("MVCDemo Issue");
                List Doclist = ctx.Web.Lists.GetByTitle("MVCDemo");
                ctx.Load(list);
                ctx.Load(Doclist);
                ctx.ExecuteQuery();

                UserCollection users = ctx.Web.SiteUsers;
                ctx.Load(users);
                ctx.ExecuteQuery();

                //var test = list.GetItemById(3);

                //ctx.Load(test, include => include.File, include => include["DocumentIssue"], include => include["Status"], include => include["Priority"]);
                //ctx.ExecuteQuery();

                //var s = test["Status"];
                //var s2 = test["Priority"];

                FieldUrlValue newUrl = new FieldUrlValue();

                newUrl.Description = documentName;
                newUrl.Url = documentUrl;



                foreach (User user in users)
                {
                    ctx.Load(user, include => include.Email);
                    ctx.ExecuteQuery();
                    if (user.Email.ToLower() == assignTo.ToLower())
                    {
                        spUser = user;
                    }

                }

                ListItemCollection AllItems = Doclist.GetItems(CamlQuery.CreateAllItemsQuery());
                ctx.Load(AllItems);
                ctx.ExecuteQuery();

                ListItem itemToUse;

                foreach (ListItem singleItem in AllItems)
                {
                    ctx.Load(singleItem, include => include.File.Name);
                    ctx.ExecuteQuery();
                    if (singleItem.File.Name == documentName)
                    {
                        itemToUse = singleItem;
                        
                    }
                }

               // Attachment at = new Attachment(ctx, );
            
                ListItem item = list.AddItem(new ListItemCreationInformation());

                item["Title"] = issueTitle;
                item["Comment"] = issueDescription;
                item["AssignedTo"] = spUser;
                item["DocumentIssue"] = newUrl;
                item["Status"] = status;
                item["Priority"] = priority;


                item.Update();
                ctx.ExecuteQuery();

                var testing = list.GetItemById(1);

                ctx.Load(testing, include => include["AssignedTo"]);
                ctx.ExecuteQuery();

                Console.WriteLine(testing);

                var usersss = testing["AssignedTo"];

                return View();
            }
        }
    }
}