using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessAuditor.BusinessLogic
{
    public static class ExportManager
    {
        //public  static void ExportToExcel()
        //{
        //    var fileName = string.Format(@"Permissions_{0}.xlsx", DateTime.Now.Ticks);
        //    var filePath = AppDomain.CurrentDomain.BaseDirectory + fileName;
        //    using (XLWorkbook wb = new XLWorkbook())
        //    {
        //        //Add DataTable in worksheet  
        //        var contenWisePermissionSheet = wb.Worksheets.Add("Content Wise Permissions");
        //        var siteInfo = DataManager.SiteCache;
        //        var currentRow = 1;
        //        contenWisePermissionSheet.Cell(currentRow, 1).Value = "Site";
        //        contenWisePermissionSheet.Cell(currentRow, 2).Value = "Subsite";
        //        contenWisePermissionSheet.Cell(currentRow, 3).Value = "List/Library";
        //        contenWisePermissionSheet.Cell(currentRow, 4).Value = "Members";
        //        contenWisePermissionSheet.Cell(currentRow, 5).Value = "Permission";

        //        contenWisePermissionSheet.Row(currentRow).Style.Font.Bold = true;
        //        contenWisePermissionSheet.Row(currentRow).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        //        contenWisePermissionSheet.Row(currentRow).Style.Fill.BackgroundColor = XLColor.Yellow;

        //        //first data row will start here
        //        currentRow++;
        //        contenWisePermissionSheet.Cell(currentRow, 1).Value = siteInfo.Title;
        //        foreach (var subSite in siteInfo.Sites)
        //        {


        //            foreach (var subSitePermission in subSite.Permissions)
        //            {
        //                currentRow++;
        //                contenWisePermissionSheet.Cell(currentRow, 2).Value = subSite.Title;
        //                contenWisePermissionSheet.Cell(currentRow, 4).Value = subSitePermission.Member;
        //                contenWisePermissionSheet.Cell(currentRow, 5).Value = subSitePermission.Permissions;
        //            }
        //        }
        //        foreach (var list in siteInfo.Lists)
        //        {
        //            foreach (var listPermission in list.Permissions)
        //            {
        //                currentRow++;
        //                contenWisePermissionSheet.Cell(currentRow, 3).Value = list.Title;
        //                contenWisePermissionSheet.Cell(currentRow, 4).Value = listPermission.Member;
        //                contenWisePermissionSheet.Cell(currentRow, 5).Value = listPermission.Permissions;

        //            }
        //        }


        //        contenWisePermissionSheet.Columns().AdjustToContents();
        //        contenWisePermissionSheet.Rows().AdjustToContents();


        //        var memberWisePermissionSheet = wb.Worksheets.Add("Members Wise Permission");

        //        currentRow = 1;
        //        memberWisePermissionSheet.Cell(currentRow, 1).Value = "Member";
        //        memberWisePermissionSheet.Cell(currentRow, 2).Value = "Site";
        //        memberWisePermissionSheet.Cell(currentRow, 3).Value = "List";
        //        memberWisePermissionSheet.Cell(currentRow, 4).Value = "Permission";

        //        memberWisePermissionSheet.Row(currentRow).Style.Font.Bold = true;
        //        memberWisePermissionSheet.Row(currentRow).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        //        memberWisePermissionSheet.Row(currentRow).Style.Fill.BackgroundColor = XLColor.Yellow;

        //        foreach (var summary in this.permissionSummaries)
        //        {
        //            foreach (var site in summary.SiteCollection)
        //            {
        //                currentRow++;
        //                memberWisePermissionSheet.Cell(currentRow, 1).Value = summary.Member;
        //                memberWisePermissionSheet.Cell(currentRow, 2).Value = site.Title;
        //                memberWisePermissionSheet.Cell(currentRow, 4).Value = site.Permissions;
        //            }

        //            foreach (var list in summary.ListCollection)
        //            {
        //                currentRow++;
        //                memberWisePermissionSheet.Cell(currentRow, 1).Value = summary.Member;
        //                memberWisePermissionSheet.Cell(currentRow, 3).Value = list.Title;
        //                memberWisePermissionSheet.Cell(currentRow, 4).Value = list.Permissions;
        //            }


        //        }
        //        memberWisePermissionSheet.Columns().AdjustToContents();
        //        memberWisePermissionSheet.Rows().AdjustToContents();

        //        var memberPermissionSummarySheet = wb.Worksheets.Add("Member Permission Summary");

        //        currentRow = 1;
        //        memberPermissionSummarySheet.Cell(currentRow, 1).Value = "Member";
        //        memberPermissionSummarySheet.Cell(currentRow, 2).Value = "Site";
        //        memberPermissionSummarySheet.Cell(currentRow, 3).Value = "Lists";
        //        memberPermissionSummarySheet.Cell(currentRow, 4).Value = "Full Control";
        //        memberPermissionSummarySheet.Cell(currentRow, 5).Value = "Edit";
        //        memberPermissionSummarySheet.Cell(currentRow, 6).Value = "Read";
        //        memberPermissionSummarySheet.Cell(currentRow, 7).Value = "Limited Access";

        //        memberPermissionSummarySheet.Row(currentRow).Style.Font.Bold = true;
        //        memberPermissionSummarySheet.Row(currentRow).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        //        memberPermissionSummarySheet.Row(currentRow).Style.Fill.BackgroundColor = XLColor.Yellow;

        //        foreach (var summary in this.permissionSummaries)
        //        {
        //            currentRow++;
        //            memberPermissionSummarySheet.Cell(currentRow, 1).Value = summary.Member;
        //            memberPermissionSummarySheet.Cell(currentRow, 2).Value = summary.Sites;
        //            memberPermissionSummarySheet.Cell(currentRow, 3).Value = summary.Lists;
        //            memberPermissionSummarySheet.Cell(currentRow, 4).Value = summary.FullControl;
        //            memberPermissionSummarySheet.Cell(currentRow, 5).Value = summary.Edit;
        //            memberPermissionSummarySheet.Cell(currentRow, 6).Value = summary.Read;
        //            memberPermissionSummarySheet.Cell(currentRow, 7).Value = summary.LimitedAccess;
        //        }


        //        memberPermissionSummarySheet.Columns().AdjustToContents();
        //        memberPermissionSummarySheet.Rows().AdjustToContents();

        //        //using (MemoryStream memoryStream = SaveWorkbookToMemoryStream(wb))
        //        //{
        //        //    File.WriteAllBytes("permissions.xlsx", memoryStream.ToArray());
        //        //}
        //        wb.SaveAs(filePath);


        //    }
        //}
        
    }
}
