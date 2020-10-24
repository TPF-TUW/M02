﻿using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DBConnection;
using DevExpress.LookAndFeel;
using DevExpress.Utils.Drawing.Helpers;
using DevExpress.Utils.Extensions;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
using MDS00;

namespace M02
{
    public partial class XtraForm1 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private const int HOLIDAY_ROWS = 30;
        private Functionality.Function FUNC = new Functionality.Function();
        public XtraForm1()
        {
            InitializeComponent();
            UserLookAndFeel.Default.StyleChanged += MyStyleChanged;
            iniConfig = new IniFile("Config.ini");
            UserLookAndFeel.Default.SetSkinStyle(iniConfig.Read("SkinName", "DevExpress"), iniConfig.Read("SkinPalette", "DevExpress"));
        }

        private IniFile iniConfig;

        private void MyStyleChanged(object sender, EventArgs e)
        {
            UserLookAndFeel userLookAndFeel = (UserLookAndFeel)sender;
            LookAndFeelChangedEventArgs lookAndFeelChangedEventArgs = (DevExpress.LookAndFeel.LookAndFeelChangedEventArgs)e;
            //MessageBox.Show("MyStyleChanged: " + lookAndFeelChangedEventArgs.Reason.ToString() + ", " + userLookAndFeel.SkinName + ", " + userLookAndFeel.ActiveSvgPaletteName);
            iniConfig.Write("SkinName", userLookAndFeel.SkinName, "DevExpress");
            iniConfig.Write("SkinPalette", userLookAndFeel.ActiveSvgPaletteName, "DevExpress");
        }

        private void XtraForm1_Load(object sender, EventArgs e)
        {
            bbiNew.PerformClick();
            
        }

        private void NewData()
        {
            txteCalendarNo.EditValue = new DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDCALENDAR), '') = '' THEN 1 ELSE MAX(OIDCALENDAR) + 1 END AS NewNo FROM CalendarMaster").getString();
            //txteCalendarNo.EditValue = "";
            speYEAR.Value = Convert.ToInt32(DateTime.Now.ToString("yyyy"));

            cbeComType.EditValue = "";

            cbeComName.EditValue = "";
            cbeComName.Properties.DisplayMember = "";
            cbeComName.Properties.ValueMember = "";

            lblStatus.Text = "* Add Calendar";
            lblStatus.ForeColor = Color.Green;

            txeCREATE.EditValue = "0";
            txeDATE.EditValue = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
        }

        private void LoadData()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT OIDCALENDAR AS CalendarNo, CompanyType, 'Thai Parfun' AS CompanyTypeName, OIDCompany AS CompanyNo, 'Thai Parfun' AS CompanyName, WorkingPerWeek, Year, CreatedBy, CreatedDate ");
            sbSQL.Append("FROM CalendarMaster ");
            sbSQL.Append("WHERE CompanyType = 0 AND Year = '" + speYEAR.Value + "' ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT OIDCALENDAR AS CalendarNo, CompanyType, 'Customer' AS CompanyTypeName, OIDCompany AS CompanyNo, CompanyName, WorkingPerWeek, Year, CreatedBy, CreatedDate ");
            sbSQL.Append("FROM CalendarMaster A ");
            sbSQL.Append("CROSS APPLY(SELECT ShortName AS CompanyName FROM Customer WHERE OIDCUST = A.OIDCompany) B ");
            sbSQL.Append("WHERE CompanyType = 1 AND Year = '" + speYEAR.Value + "' ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT OIDCALENDAR AS CalendarNo, CompanyType, 'Vendor' AS CompanyTypeName, OIDCompany AS CompanyNo, CompanyName, WorkingPerWeek, Year, CreatedBy, CreatedDate ");
            sbSQL.Append("FROM CalendarMaster C ");
            sbSQL.Append("CROSS APPLY(SELECT Name AS CompanyName FROM Vendor WHERE OIDVEND = C.OIDCompany) D ");
            sbSQL.Append("WHERE CompanyType = 2 AND Year = '" + speYEAR.Value + "' ");
            sbSQL.Append("ORDER BY CompanyType, OIDCALENDAR, CompanyName ");
            new ObjDevEx.setGridControl(gcCalendar, gvCalendar, sbSQL).getData(false, false, true, true);

            sbSQL.Clear();
            sbSQL.Append("SELECT '0' AS ID, 'Monday --> Friday' AS WorkingPerWeek ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT '1' AS ID, 'Monday --> Saturday' AS WorkingPerWeek ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT '2' AS ID, 'Sunday --> Saturday' AS WorkingPerWeek ");
            new ObjDevEx.setGridLookUpEdit(lueWorkDay, sbSQL, "WorkingPerWeek", "ID").getData();

            sbSQL.Clear();
            sbSQL.Append("SELECT '0' AS ID, 'Thai Parfun' AS CompanyType ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT '1' AS ID, 'Customer' AS CompanyType ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT '2' AS ID, 'Vendor' AS CompanyType ");
            new ObjDevEx.setGridLookUpEdit(cbeComType, sbSQL, "CompanyType", "ID").getData();
        }

        private void LoadHoliday()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT CONVERT(DATETIME, Holiday, 103) AS Date ");
            sbSQL.Append("FROM CalendarDetail ");
            sbSQL.Append("WHERE(OIDCALENDAR = '" + txteCalendarNo.EditValue.ToString() + "') ");
            sbSQL.Append("ORDER BY Holiday ");

            new ObjDevEx.setGridControl(gcHoliday, gvHoliday).ClearGrid();
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("No", typeof(String));
            dt.Columns.Add("Date", typeof(DateTime));
            dt.Columns[0].ReadOnly = true;
            System.Data.DataTable dtRead = new DBConnection.DBQuery(sbSQL).getDataTable();
            int ii = 0;
            foreach (System.Data.DataRow row in dtRead.Rows)
            {
                ii++;
                DateTime dateX = Convert.ToDateTime(null);
                foreach (var item in row.ItemArray)
                {
                    dateX = Convert.ToDateTime(item);
                }
                dt.Rows.Add(ii, dateX);
            }

            //for (int i = dt.Rows.Count; i < HOLIDAY_ROWS; i++)
            //{
            //    dt.Rows.Add((i + 1).ToString(), null);
            //}

            gcHoliday.DataSource = dt;
            gcHoliday.EndUpdate();
            gcHoliday.ResumeLayout();
            gvHoliday.BestFitColumns();
            gvHoliday.ClearSelection();

            gvHoliday.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Bottom;
        }

        private void LoadCompanyName()
        {

            cbeComName.EditValue = "";
            cbeComName.Properties.DisplayMember = "";
            cbeComName.Properties.ValueMember = "";

            new ObjDevEx.setSearchLookUpEdit(cbeComName).ClearSearchLookUpEdit();
            if (cbeComType.Text.Trim() != "")
            {
                int CompanyID = Convert.ToInt32(cbeComType.EditValue.ToString());
                StringBuilder sbSQL = new StringBuilder();
                if (CompanyID == 0)
                {
                    sbSQL.Append("SELECT '0' AS ID, 'Thai Parfun' AS Name");
                }
                else if (CompanyID == 1)
                {
                    sbSQL.Append("SELECT OIDCUST AS ID, ShortName AS Name FROM Customer ORDER BY ShortName");
                }
                else if (CompanyID == 2)
                {
                    sbSQL.Append("SELECT OIDVEND AS ID, Name FROM Vendor ORDER BY Name");
                }

                if (sbSQL.Length > 0)
                {
                    new ObjDevEx.setSearchLookUpEdit(cbeComName, sbSQL, "Name", "ID").getData(false);
                    try
                    {
                        cbeComName.Properties.View.Columns[0].BestFit();
                    }
                    catch (Exception) { }

                }
            }
        }

        //private int findCompanyType(string Company)
        //{
        //    int return_value = 999;
        //    switch (Company)
        //    {
        //        case "Thai Parfun": 
        //            return_value = 0; 
        //            break;
        //        case "Customer":
        //            return_value = 1;
        //            break;
        //        case "Vendor":
        //            return_value = 2;
        //            break;
        //        default:
        //            return_value = 999;
        //            break;
        //    }
        //    return return_value;
        //}

        //private string findCompanyTypeName(string Company)
        //{
        //    string return_value = "";
        //    switch (Company)
        //    {
        //        case "0":
        //            return_value = "Thai Parfun";
        //            break;
        //        case "1":
        //            return_value = "Customer";
        //            break;
        //        case "2":
        //            return_value = "Vendor";
        //            break;
        //        default:
        //            return_value = "";
        //            break;
        //    }
        //    return return_value;
        //}


        private void cbeComName_Click(object sender, EventArgs e)
        {
            try
            {
                cbeComName.Properties.View.Columns[0].BestFit();
            }
            catch (Exception) { }
        }

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            //MessageBox.Show(cbeComName.Properties.GetDisplayTextByKeyValue(cbeComName.EditValue));
            //MessageBox.Show(cbeComName.Properties.View.GetFocusedRowCellValue(cbeComName.Properties.ValueMember).ToString());
            //MessageBox.Show(cbeComName.EditValue.ToString());
            NewData();
            LoadData();
            LoadHoliday();
            ClearCheckDay();
            gvCalendar.OptionsBehavior.Editable = false;
            gvCalendar.ClearSelection();
            speYEAR.Focus();
        }

        private void lueWorkDay_EditValueChanged(object sender, EventArgs e)
        {
            ShowCheckDay(lueWorkDay.EditValue.ToString());
        }

        private void ClearCheckDay()
        {
            cheSunday.Checked = false;
            cheMonday.Checked = false;
            cheTuesday.Checked = false;
            cheWednesday.Checked = false;
            cheThursday.Checked = false;
            cheFriday.Checked = false;
            cheSaturday.Checked = false;

            cheSunday.Properties.ReadOnly = true;
            cheMonday.Properties.ReadOnly = true;
            cheTuesday.Properties.ReadOnly = true;
            cheWednesday.Properties.ReadOnly = true;
            cheThursday.Properties.ReadOnly = true;
            cheFriday.Properties.ReadOnly = true;
            cheSaturday.Properties.ReadOnly = true;
        }

        private void ShowCheckDay(string strWeekGroup)
        {
            ClearCheckDay();

            if (strWeekGroup != "")
            {
                switch (strWeekGroup)
                {
                    case "0":
                        cheSunday.Checked = false;
                        cheMonday.Checked = true;
                        cheTuesday.Checked = true;
                        cheWednesday.Checked = true;
                        cheThursday.Checked = true;
                        cheFriday.Checked = true;
                        cheSaturday.Checked = false;
                        break;
                    case "1":
                        cheSunday.Checked = false;
                        cheMonday.Checked = true;
                        cheTuesday.Checked = true;
                        cheWednesday.Checked = true;
                        cheThursday.Checked = true;
                        cheFriday.Checked = true;
                        cheSaturday.Checked = true;
                        break;
                    case "2":
                        cheSunday.Checked = true;
                        cheMonday.Checked = true;
                        cheTuesday.Checked = true;
                        cheWednesday.Checked = true;
                        cheThursday.Checked = true;
                        cheFriday.Checked = true;
                        cheSaturday.Checked = true;
                        break;
                    default:
                        ClearCheckDay();
                        break;
                }
            }
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (speYEAR.Value < 1000)
            {
                FUNC.msgWarning("The year must have 4 digits.");
                speYEAR.Focus();
            }
            else if (cbeComType.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input company type.");
                cbeComType.Focus();
            }
            else if (cbeComName.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input company name.");
                cbeComName.Focus();
            }
            else if (lueWorkDay.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select working per week.");
                lueWorkDay.Focus();
            }
            else
            {
                if(FUNC.msgQuiz("Confirm save data ?") == true)
                {
                    StringBuilder sbSQL = new StringBuilder();
                    //CalendarMaster
                    int ComType = Convert.ToInt32(cbeComType.EditValue.ToString());

                    string strCREATE = "0";
                    if (txeCREATE.Text.Trim() != "")
                    {
                        strCREATE = txeCREATE.Text.Trim();
                    }

                    sbSQL.Append("IF NOT EXISTS(SELECT OIDCALENDAR FROM CalendarMaster WHERE OIDCALENDAR = '" + txteCalendarNo.Text.Trim() + "') ");
                    sbSQL.Append(" BEGIN ");
                    sbSQL.Append("  INSERT INTO CalendarMaster(CompanyType, OIDCompany, WorkingPerWeek, Year, CreatedBy, CreatedDate) ");
                    sbSQL.Append("  VALUES('" + ComType.ToString() + "', '" + cbeComName.EditValue.ToString() + "', '" + lueWorkDay.EditValue.ToString() + "', '" + speYEAR.Value.ToString() + "', '" + strCREATE + "', GETDATE()) ");
                    sbSQL.Append(" END ");
                    sbSQL.Append("ELSE ");
                    sbSQL.Append(" BEGIN ");
                    sbSQL.Append("  UPDATE CalendarMaster SET ");
                    sbSQL.Append("      CompanyType = '" + ComType.ToString() + "', OIDCompany = '" + cbeComName.EditValue.ToString() + "', WorkingPerWeek = '" + lueWorkDay.EditValue.ToString() + "', Year = '" + speYEAR.Value.ToString() + "' ");
                    sbSQL.Append("  WHERE(OIDCALENDAR = '" + txteCalendarNo.Text.Trim() + "') ");
                    sbSQL.Append(" END ");

                    //CalendarDetail
                    sbSQL.Append("DELETE FROM CalendarDetail WHERE(OIDCALENDAR = '" + txteCalendarNo.Text.Trim() + "') ");

                    if (gvHoliday.RowCount > 0)
                    {
                        for (int i = 0; i < gvHoliday.RowCount-1; i++)
                        {
                            string xDATE = gvHoliday.GetRowCellValue(i, "Date").ToString();
                            if (xDATE != "")
                            {
                                xDATE = Convert.ToDateTime(xDATE).ToString("yyyy-MM-dd");
                                sbSQL.Append("INSERT INTO CalendarDetail(OIDCALENDAR, Holiday) ");
                                sbSQL.Append("  VALUES('" + txteCalendarNo.Text.Trim() + "', '" + xDATE + "')  ");
                            }
                        }
                    }

                    if (sbSQL.Length > 0)
                    {
                        try
                        {
                            bool chkSAVE = new DBQuery(sbSQL).runSQL();
                            if (chkSAVE == true)
                            {
                                FUNC.msgInfo("Save complete.");
                                bbiNew.PerformClick();
                            }
                        }
                        catch (Exception)
                        { }
                    }
                }
                
            }
            
        }

        private void gvCalendar_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            lblStatus.Text = "* Edit Calendar";
            lblStatus.ForeColor = Color.Red;

            txteCalendarNo.EditValue = gvCalendar.GetFocusedRowCellValue("CalendarNo").ToString();
            LoadHoliday();
            speYEAR.Value = Convert.ToInt32(gvCalendar.GetFocusedRowCellValue("Year").ToString());
            cbeComType.EditValue = gvCalendar.GetFocusedRowCellValue("CompanyType").ToString();
            cbeComName.EditValue = gvCalendar.GetFocusedRowCellValue("CompanyNo").ToString();
            lueWorkDay.EditValue = gvCalendar.GetFocusedRowCellValue("WorkingPerWeek").ToString();

            txeCREATE.EditValue = gvCalendar.GetFocusedRowCellValue("CreatedBy").ToString();
            txeDATE.EditValue = gvCalendar.GetFocusedRowCellValue("CreatedDate").ToString();
        }

        private void lueWorkDay_EditValueChanged_1(object sender, EventArgs e)
        {
            ShowCheckDay(lueWorkDay.EditValue.ToString());
        }

        private void speYEAR_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                cbeComType.Focus();
            }
        }

        private void cbeComName_EditValueChanged(object sender, EventArgs e)
        {
            if (cbeComName.Text.Trim() != "" && lblStatus.Text == "* Add Calendar")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT TOP(1) OIDCALENDAR FROM CalendarMaster WHERE (CompanyType = '" + cbeComType.EditValue.ToString() + "') AND (OIDCompany = '" + cbeComName.EditValue.ToString() + "') ");
                if (new DBQuery(sbSQL).getString() != "")
                {
                    FUNC.msgWarning("Duplicate company. !! Please Change.");
                    cbeComName.EditValue = "";
                    cbeComName.Focus();
                }
                else
                {
                    lueWorkDay.Focus();
                }
            }
            else if (cbeComName.Text.Trim() != "" && lblStatus.Text == "* Edit Calendar")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT TOP(1) OIDCALENDAR ");
                sbSQL.Append("FROM CalendarMaster ");
                sbSQL.Append("WHERE (CompanyType = '" + cbeComType.EditValue.ToString() + "') ");
                sbSQL.Append("AND (OIDCompany = '" + cbeComName.EditValue.ToString() + "') ");
                string strCHK = new DBQuery(sbSQL).getString();
                if (strCHK != "" && strCHK != txteCalendarNo.Text.Trim())
                {
                    FUNC.msgWarning("Duplicate company. !! Please Change.");
                    cbeComName.EditValue = "";
                    cbeComName.Focus();
                }
                else
                {
                    lueWorkDay.Focus();
                }
            }
            else
            {
                lueWorkDay.Focus();
            }
        }


        private void cbeComType_EditValueChanged(object sender, EventArgs e)
        {
            LoadCompanyName();
            cbeComName.Focus();
        }

        private void gvCalendar_RowStyle(object sender, RowStyleEventArgs e)
        {
            if (sender is GridView)
            {
                GridView gView = (GridView)sender;
                if (!gView.IsValidRowHandle(e.RowHandle)) return;
                int parent = gView.GetParentRowHandle(e.RowHandle);
                if (gView.IsGroupRow(parent))
                {
                    for (int i = 0; i < gView.GetChildRowCount(parent); i++)
                    {
                        if (gView.GetChildRowHandle(parent, i) == e.RowHandle)
                        {
                            e.Appearance.BackColor = i % 2 == 0 ? Color.AliceBlue : Color.White;
                        }
                    }
                }
                else
                {
                    e.Appearance.BackColor = e.RowHandle % 2 == 0 ? Color.AliceBlue : Color.White;
                }
            }
        }
    }
}