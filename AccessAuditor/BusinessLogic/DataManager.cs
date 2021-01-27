using AccessAuditor.Models;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace AccessAuditor.BusinessLogic
{
    public class DataManager
    {
        public static SPSite SiteCache { get; set; }


        public static void ClearCache()
        {
            SiteCache = new SPSite();
        }

        public static void LoadSiteDetails(ClientContext clientContext, Web web, SPSite site)
        {

            clientContext.Load(web.Lists);
            clientContext.Load(web.Webs);

            // Execute the query to the server.
            clientContext.ExecuteQuery();

            var lists = new List<SPList>();
            foreach (var item in web.Lists)
            {
                lists.Add(new SPList()
                {
                    Title = item.Title,
                    Permissions = GetSitePermissionDetails(clientContext, web, item.Title)
                });
            }
            site.Lists = lists;
            
            

            var sites = new List<SPSite>();
            foreach (var subWeb in web.Webs)
            {

                var currentSite = new SPSite()
                {
                    Title = subWeb.Title,
                    Permissions = GetSitePermissionDetails(clientContext, subWeb),

                };
                LoadSiteDetails(clientContext, subWeb, currentSite);
                sites.Add(currentSite);
            }
            site.Sites = sites;

        }
        public static void LoadSiteDetails(string spSiteURL, string username, string password)
        {
            SiteCache = new SPSite();
            ClientContext context = GetSPOContext(spSiteURL, username, password);
            context.Load(context.Web);
           
            // Execute the query to the server.
            context.ExecuteQuery();

            // Root site
            SiteCache.Title = context.Web.Title;

            // Root site permissions
            SiteCache.Permissions = GetSitePermissionDetails(context, context.Web);
            LoadSiteDetails(context, context.Web, SiteCache);
            var ab = SiteCache;
            //var lists = new List<SPList>();
            //foreach (var item in context.Web.Lists)
            //{
            //    lists.Add(new SPList()
            //    {
            //        Title = item.Title,
            //        Permissions = GetSitePermissionDetails(context, context.Web, item.Title)
            //    });
            //}

            //var sites = new List<SPSite>();
            //foreach (var web in context.Web.Webs)
            //{
               
            //    sites.Add(new SPSite()
            //    {
            //        Title = web.Title,
            //        Permissions = GetSitePermissionDetails(context, web),

            //    });
            //}
           // SiteCache.Sites = sites;
           // SiteCache.Lists = lists;
        }

        public static void LoadSiteDetails_back(string spSiteURL, string username, string password)
        {
            SiteCache = new SPSite();
            ClientContext context = GetSPOContext(spSiteURL, username, password);
            context.Load(context.Web);
            context.Load(context.Web.Lists);
            context.Load(context.Web.Webs);

            // Execute the query to the server.
            context.ExecuteQuery();

            // Root site
            SiteCache.Title = context.Web.Title;

            // Root site permissions
            SiteCache.Permissions = GetSitePermissionDetails(context, context.Web);
            var lists = new List<SPList>();
            foreach (var item in context.Web.Lists)
            {
                lists.Add(new SPList()
                {
                    Title = item.Title,
                    Permissions = GetSitePermissionDetails(context, context.Web, item.Title)
                });
            }

            var sites = new List<SPSite>();
            foreach (var web in context.Web.Webs)
            {
                sites.Add(new SPSite()
                {
                    Title = web.Title,
                    Permissions = GetSitePermissionDetails(context, web),

                });
            }
            SiteCache.Sites = sites;
            SiteCache.Lists = lists;
        }

        public static Web ValidateCredentials(string spSiteURL, string username, string password)
        {
            ClientContext context = GetSPOContext(spSiteURL, username, password);
            context.Load(context.Web);
            context.ExecuteQuery();

            return context.Web;
        }
        private static ClientContext GetSPOContext(string spSiteURL, string username, string password)
        {

            var secure = new SecureString();
            foreach (char c in password)
            {
                secure.AppendChar(c);
            }
            ClientContext spoContext = new ClientContext(spSiteURL);
            spoContext.Credentials = new SharePointOnlineCredentials(username, secure);
            return spoContext;
        }
        private static IEnumerable<SPPermission> GetSitePermissionDetails(ClientContext clientContext, Web web)
        {
            var roles = clientContext.LoadQuery(web.RoleAssignments.Include(roleAsg => roleAsg.Member,
                                                                              roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name)));
            clientContext.ExecuteQuery();

            var permisionDetails = new List<SPPermission>();
            foreach (RoleAssignment ra in roles)
            {
                var rdc = ra.RoleDefinitionBindings;
                var permissions = new List<string>();
                foreach (var rdbc in rdc) permissions.Add(rdbc.Name);
                permisionDetails.Add(new SPPermission()
                {
                    Member = ra.Member.Title,
                    //RoleBindings = permissions,
                    Permissions = string.Join(", ", permissions)
                });
            }
            return permisionDetails;
        }


        private static IEnumerable<SPPermission> GetSitePermissionDetails(ClientContext clientContext, Web web, string listName)
        {
            var roles = clientContext.LoadQuery(web.Lists.GetByTitle(listName).RoleAssignments.Include(roleAsg => roleAsg.Member,
                                                                              roleAsg => roleAsg.RoleDefinitionBindings.Include(roleDef => roleDef.Name)));
            clientContext.ExecuteQuery();

            var permisionDetails = new List<SPPermission>();
            foreach (RoleAssignment ra in roles)
            {
                var rdc = ra.RoleDefinitionBindings;
                var permissions = new List<string>();
                foreach (var rdbc in rdc) permissions.Add(rdbc.Name);
                permisionDetails.Add(new SPPermission()
                {
                    Member = ra.Member.Title,
                    //RoleBindings = permissions,
                    Permissions = string.Join(", ", permissions)
                });
            }
            return permisionDetails;
        }

    }
}
