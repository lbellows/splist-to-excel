using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelExport
{
    class Program
    {
        static void Main(string[] args)
        {
            const string fileName = "export.csv";
            const string placeholder = " ";

            string siteUrl = Config.URL;
            string listUrl = Config.LIST_URL;
            AuthenticationManager am = new AuthenticationManager();
            ClientContext cc =
                am.GetSharePointOnlineAuthenticatedContextTenant(siteUrl,
                Config.USER_NAME,
                Config.PASS);

            var list = cc.Web.GetList(listUrl);
            cc.Load(list);
            var fields = list.Fields;
            cc.Load(fields);
            var visFields = fields.Where(f => !f.Hidden);
            var items = list.GetItems(CamlQuery.CreateAllItemsQuery(200)); //new CamlQuery { ViewXml = view.ViewQuery });
            cc.Load(items);
            cc.ExecuteQueryRetry();

            var colList = string.Join("|", visFields.Select(f => f.Title));

            Console.WriteLine("Item count: " + items.Count);
            Console.WriteLine(colList);
            System.IO.File.AppendAllLines(fileName, new string[] { colList });

            var iList = items.ToList();

            iList.ForEach(item =>
            {
                cc.Load(item);
                cc.Load(item.FieldValuesAsText);
                ////////cc.Load(item.FieldValues);
                cc.ExecuteQueryRetry();

                List<string> fieldVals = new List<string>();

                foreach (var f in visFields)
                {
                    try
                    {
                        string text = !string.IsNullOrWhiteSpace(item.FieldValuesAsText[f.InternalName]) ? item.FieldValuesAsText[f.InternalName] : placeholder;
                        text = Regex.Replace(text, @"[\u000A\u000B\u000C\u000D\u2028\u2029\u0085]+", placeholder);
                        fieldVals.Add(text);

                    }
                    catch(Exception e)
                    {
                        fieldVals.Add(placeholder);
                    }
                    
                }

                var aLine = string.Join("|", fieldVals.ToArray());
                Console.WriteLine(aLine);
                System.IO.File.AppendAllLines(fileName, new string[] { aLine });
            });

            Console.WriteLine();
            Console.WriteLine("Done");
            Console.ReadKey();
        }

    }
}
