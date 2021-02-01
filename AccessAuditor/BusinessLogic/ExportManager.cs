using AccessAuditor.Models;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessAuditor.BusinessLogic
{
    public static class ExportManager
    {
        public static int CurrentRow { get; set; }

        public static void ExportData(SPSite siteInfo, List<PermissionSummary> permissionSummaries, string filePath)
        {
            using (XLWorkbook wb = new XLWorkbook())
            {
                CurrentRow = 1;
                //Add DataTable in worksheet  
                var memberPermissionSummarySheet = wb.Worksheets.Add("Member Permission Summary");
                memberPermissionSummarySheet.Cell(CurrentRow, 1).Value = "Member";
                memberPermissionSummarySheet.Cell(CurrentRow, 2).Value = "Site";
                memberPermissionSummarySheet.Cell(CurrentRow, 3).Value = "Lists";
                memberPermissionSummarySheet.Cell(CurrentRow, 4).Value = "Full Control";
                memberPermissionSummarySheet.Cell(CurrentRow, 5).Value = "Edit";
                memberPermissionSummarySheet.Cell(CurrentRow, 6).Value = "Read";
                memberPermissionSummarySheet.Cell(CurrentRow, 7).Value = "Limited Access";

                memberPermissionSummarySheet.Row(CurrentRow).Style.Font.Bold = true;
                memberPermissionSummarySheet.Row(CurrentRow).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                memberPermissionSummarySheet.Row(CurrentRow).Style.Fill.BackgroundColor = XLColor.Yellow;

                foreach (var summary in permissionSummaries)
                {
                    CurrentRow++;
                    memberPermissionSummarySheet.Cell(CurrentRow, 1).Value = summary.Member;
                    memberPermissionSummarySheet.Cell(CurrentRow, 2).Value = summary.Sites;
                    memberPermissionSummarySheet.Cell(CurrentRow, 3).Value = summary.Lists;
                    memberPermissionSummarySheet.Cell(CurrentRow, 4).Value = summary.FullControl;
                    memberPermissionSummarySheet.Cell(CurrentRow, 5).Value = summary.Edit;
                    memberPermissionSummarySheet.Cell(CurrentRow, 6).Value = summary.Read;
                    memberPermissionSummarySheet.Cell(CurrentRow, 7).Value = summary.LimitedAccess;
                }


                memberPermissionSummarySheet.Columns().AdjustToContents();
                memberPermissionSummarySheet.Rows().AdjustToContents();

                CurrentRow = 1;
                var contenWisePermissionSheet = wb.Worksheets.Add("Content Wise Permissions");

                contenWisePermissionSheet.Cell(CurrentRow, 1).Value = "Site";
                contenWisePermissionSheet.Cell(CurrentRow, 2).Value = "Subsite";
                contenWisePermissionSheet.Cell(CurrentRow, 3).Value = "List/Library";
                contenWisePermissionSheet.Cell(CurrentRow, 4).Value = "Members";
                contenWisePermissionSheet.Cell(CurrentRow, 5).Value = "Permission";

                contenWisePermissionSheet.Row(CurrentRow).Style.Font.Bold = true;
                contenWisePermissionSheet.Row(CurrentRow).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                contenWisePermissionSheet.Row(CurrentRow).Style.Fill.BackgroundColor = XLColor.Yellow;

                //first data row will start here
                CurrentRow++;
                contenWisePermissionSheet.Cell(CurrentRow, 1).Value = siteInfo.Title;


                //foreach (var list in siteInfo.Lists)
                //{
                //    foreach (var listPermission in list.Permissions)
                //    {
                //        CurrentRow++;
                //        contenWisePermissionSheet.Cell(CurrentRow, 1).Value = siteInfo.Title;
                //        contenWisePermissionSheet.Cell(CurrentRow, 3).Value = list.Title;
                //        contenWisePermissionSheet.Cell(CurrentRow, 4).Value = listPermission.Member;
                //        contenWisePermissionSheet.Cell(CurrentRow, 5).Value = listPermission.Permissions;

                //    }
                //}
                BuildRows(siteInfo, contenWisePermissionSheet);

               


                contenWisePermissionSheet.Columns().AdjustToContents();
                contenWisePermissionSheet.Rows().AdjustToContents();


                var memberWisePermissionSheet = wb.Worksheets.Add("Members Wise Permission");

                CurrentRow = 1;
                memberWisePermissionSheet.Cell(CurrentRow, 1).Value = "Member";
                memberWisePermissionSheet.Cell(CurrentRow, 2).Value = "Site";
                memberWisePermissionSheet.Cell(CurrentRow, 3).Value = "List";
                memberWisePermissionSheet.Cell(CurrentRow, 4).Value = "Permission";

                memberWisePermissionSheet.Row(CurrentRow).Style.Font.Bold = true;
                memberWisePermissionSheet.Row(CurrentRow).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                memberWisePermissionSheet.Row(CurrentRow).Style.Fill.BackgroundColor = XLColor.Yellow;

                foreach (var summary in permissionSummaries)
                {
                    foreach (var site in summary.SiteCollection)
                    {
                        CurrentRow++;
                        memberWisePermissionSheet.Cell(CurrentRow, 1).Value = summary.Member;
                        memberWisePermissionSheet.Cell(CurrentRow, 2).Value = site.Title;
                        memberWisePermissionSheet.Cell(CurrentRow, 4).Value = site.Permissions;
                    }

                    foreach (var list in summary.ListCollection)
                    {
                        CurrentRow++;
                        memberWisePermissionSheet.Cell(CurrentRow, 1).Value = summary.Member;
                        memberWisePermissionSheet.Cell(CurrentRow, 3).Value = list.Title;
                        memberWisePermissionSheet.Cell(CurrentRow, 4).Value = list.Permissions;
                    }


                }
                memberWisePermissionSheet.Columns().AdjustToContents();
                memberWisePermissionSheet.Rows().AdjustToContents();
                wb.SaveAs(filePath);
            }

        }
        public static void BuildRows(SPSite siteInfo, IXLWorksheet contenWisePermissionSheet)
        {

            foreach (var list in siteInfo.Lists)
            {
                foreach (var listPermission in list.Permissions)
                {
                    CurrentRow++;
                    contenWisePermissionSheet.Cell(CurrentRow, 1).Value = siteInfo.Title;
                    contenWisePermissionSheet.Cell(CurrentRow, 3).Value = list.Title;
                    contenWisePermissionSheet.Cell(CurrentRow, 4).Value = listPermission.Member;
                    contenWisePermissionSheet.Cell(CurrentRow, 5).Value = listPermission.Permissions;

                }
            }

            foreach (var subSite in siteInfo.Sites)
            {

                

                foreach (var subSitePermission in subSite.Permissions)
                {
                    CurrentRow++;
                    contenWisePermissionSheet.Cell(CurrentRow, 1).Value = siteInfo.Title;
                    contenWisePermissionSheet.Cell(CurrentRow, 2).Value = subSite.Title;
                    contenWisePermissionSheet.Cell(CurrentRow, 4).Value = subSitePermission.Member;
                    contenWisePermissionSheet.Cell(CurrentRow, 5).Value = subSitePermission.Permissions;
                }
             

                BuildRows(subSite,contenWisePermissionSheet);


            }
        }
    }
}
