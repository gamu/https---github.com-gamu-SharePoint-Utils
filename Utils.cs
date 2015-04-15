using Microsoft.SharePoint;
using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI.WebControls.WebParts;
using System.Xml;
using System.Xml.Linq;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.WebPartPages;
using WebPart = Microsoft.SharePoint.WebPartPages.WebPart;

namespace Resmark.CrmApp
{
    internal static class Utils
    {
        public static void AddEventReceiver(this SPList list, SPEventReceiverType eventType, string Assembly, String ClassFull)
        {
            list.EventReceivers.Add(eventType, Assembly, ClassFull);
            list.Update();
        }

        public static SPContentType AddContentTypeToList(this SPList list, string contentTypeId, bool asDefault=true)
        {
            var currentWeb = list.ParentWeb;
            var ctOrganization = GetContentTypeById(contentTypeId, currentWeb);
            if (ctOrganization != null)
            {
                if (asDefault)
                    list.ContentTypes[0].Delete();
                list.ContentTypes.Add(ctOrganization);
                list.Update();
            }
            list.Update();
            return ctOrganization;
        }

        public static void SetContentTypeDisplayForm(SPContentType contentType, string displayFormUrl)
        {
            if (contentType != null)
            {
                contentType.DisplayFormUrl = displayFormUrl;
                contentType.Update(true);
            }
        }

        public static void AddContentTypeToList(this SPList list, SPContentType contentType, bool asDefault=true)
        {
            if (asDefault)
                list.ContentTypes[0].Delete();
            list.ContentTypes.Add(contentType);
            list.Update();
        }

        public static SPContentType GetContentTypeById(string contentTypeId, SPWeb currentWeb)
        {
            var ctIdOrganization = new SPContentTypeId(contentTypeId);
            var ctOrganization = currentWeb.ContentTypes.Cast<SPContentType>().FirstOrDefault(ct => ct.Id.Equals(ctIdOrganization));
            return ctOrganization;
        }

        public static SPFieldLookup CreateLookupField(SPFieldCollection fieldsCollection, SPList lookupList, string showField, string name, string Group)
        {
            SPFieldLookup lookup=null;
            if (!fieldsCollection.ContainsFieldWithStaticName(name))
            {
                string field = fieldsCollection.AddLookup(name, lookupList.ID, false);
                lookup = new SPFieldLookup(fieldsCollection, field);
                lookup.LookupField = showField;
                lookup.Group = Group.ToString();
                lookup.Update();
            }
            else
            {
                lookup = (SPFieldLookup)fieldsCollection[name];
            }
            return lookup;
        }

        public static void AddFieldToContentType(this SPContentType contentType, SPField field)
        {
            var fieldLink = new SPFieldLink(field);
            var thisField = contentType.FieldLinks.Cast<SPFieldLink>().FirstOrDefault(n => n.Name == field.InternalName);
            if(thisField==null)
            {
                contentType.FieldLinks.Add(fieldLink);
                contentType.Update();
            }
            
        }

        public static void ActionForList(SPList list, Action<SPList> action)
        {
            if (list != null)
            {
                action(list);
            }
        }

        public static void WebpartAction(SPWeb web, string pageUrl, string webPartTitle, Action<WebPart, SPLimitedWebPartManager> func)
        {
            SPLimitedWebPartManager wpManager = web.GetLimitedWebPartManager(web.Url+pageUrl,PersonalizationScope.Shared);
            var webpart = wpManager.WebParts.Cast<WebPart>().FirstOrDefault(wp => wp.Title == webPartTitle);
            func(webpart, wpManager);
            web.Update();
            wpManager.Dispose();
        }

        public static void Seed(this SPList list, Action<SPList> seedMethod)
        {
            if (list != null)
                seedMethod(list);
        }

        public static void AddJSLinkToObject(dynamic jsLinkebleObject, string jslinkFile)
        {
            long assemblyTimeStamp = File.GetCreationTime(Assembly.GetExecutingAssembly().Location).Ticks;
            jsLinkebleObject.JSLink = jslinkFile + "?rev" + assemblyTimeStamp;
            jsLinkebleObject.Update();
        }

        public static void AddJSLinkToObject(XsltListViewWebPart jsLinkebleObject, SPLimitedWebPartManager lwManager,
            string jslinkFile)
        {
            long assemblyTimeStamp = File.GetCreationTime(Assembly.GetExecutingAssembly().Location).Ticks;
            jsLinkebleObject.JSLink = jslinkFile + "?rev" + assemblyTimeStamp;
            lwManager.SaveChanges(jsLinkebleObject);
        }


        
    }
}
