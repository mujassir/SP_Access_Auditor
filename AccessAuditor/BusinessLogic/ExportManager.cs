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

                //site level permissions

                foreach (var sitePermission in siteInfo.Permissions)
                {
                    CurrentRow++;
                    contenWisePermissionSheet.Cell(CurrentRow, 1).Value = siteInfo.Title;
                    contenWisePermissionSheet.Cell(CurrentRow, 4).Value = sitePermission.Member;
                    contenWisePermissionSheet.Cell(CurrentRow, 5).Value = sitePermission.Permissions;

                }

                BuildRowsForContentWiseAccess(siteInfo, contenWisePermissionSheet);

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
                    foreach (var sitePermission in siteInfo.Permissions.Where(p => p.Member == summary.Member))
                    {
                        CurrentRow++;
                        memberWisePermissionSheet.Cell(CurrentRow, 1).Value = summary.Member;
                        memberWisePermissionSheet.Cell(CurrentRow, 2).Value = siteInfo.Title;
                        memberWisePermissionSheet.Cell(CurrentRow, 4).Value = sitePermission.Permissions;
                    }

                    BuildRowsForMemberWiseAccess(siteInfo, memberWisePermissionSheet, summary.Member);
                }
                memberWisePermissionSheet.Columns().AdjustToContents();
                memberWisePermissionSheet.Rows().AdjustToContents();
                wb.SaveAs(filePath);
            }

        }
        private static void BuildRowsForContentWiseAccess(SPSite siteInfo, IXLWorksheet worksheetRow)
        {
            foreach (var list in siteInfo.Lists)
            {
                foreach (var listPermission in list.Permissions)
                {
                    CurrentRow++;
                    worksheetRow.Cell(CurrentRow, 1).Value = siteInfo.Title;
                    worksheetRow.Cell(CurrentRow, 3).Value = list.Title;
                    worksheetRow.Cell(CurrentRow, 4).Value = listPermission.Member;
                    worksheetRow.Cell(CurrentRow, 5).Value = listPermission.Permissions;
                }
            }
            foreach (var subSite in siteInfo.Sites)
            {
                foreach (var subSitePermission in subSite.Permissions)
                {
                    CurrentRow++;
                    worksheetRow.Cell(CurrentRow, 1).Value = siteInfo.Title;
                    worksheetRow.Cell(CurrentRow, 2).Value = subSite.Title;
                    worksheetRow.Cell(CurrentRow, 4).Value = subSitePermission.Member;
                    worksheetRow.Cell(CurrentRow, 5).Value = subSitePermission.Permissions;
                }
                BuildRowsForContentWiseAccess(subSite, worksheetRow);
            }
        }

        private static void BuildRowsForMemberWiseAccess(SPSite siteInfo, IXLWorksheet worksheetRow, string filterByMemberAccess)
        {
            foreach (var list in siteInfo.Lists)
            {
                foreach (var listPermission in list.Permissions.Where(p => p.Member == filterByMemberAccess))
                {
                    CurrentRow++;
                    worksheetRow.Cell(CurrentRow, 1).Value = filterByMemberAccess;
                    worksheetRow.Cell(CurrentRow, 3).Value = list.Title;
                    worksheetRow.Cell(CurrentRow, 4).Value = listPermission.Permissions;
                }
            }
            foreach (var subSite in siteInfo.Sites)
            {
                foreach (var subSitePermission in subSite.Permissions.Where(p => p.Member == filterByMemberAccess))
                {
                    CurrentRow++;
                    worksheetRow.Cell(CurrentRow, 1).Value = filterByMemberAccess;
                    worksheetRow.Cell(CurrentRow, 2).Value = subSite.Title;
                    worksheetRow.Cell(CurrentRow, 4).Value = subSitePermission.Permissions;
                }
                BuildRowsForMemberWiseAccess(subSite, worksheetRow, filterByMemberAccess);
            }
        }
    }
}
