/*************************
 * 管控对象：月度绩效自评
 * 管控成员：袁文秋
 ************************/

using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using PEMSoft.Client.SYS.WnControl;
using PEMSoft.Client.SYS.WnPlatform;

namespace PEMSoft.Client.IFM.Jxgl.WnDayPerformanceSelfEvaluate
{
    public partial class WnFormMonthTaskEvaluate : WnForm
    {
        #region 变量

        /// <summary>
        /// 存放当前用户的EmployeeName
        /// </summary>
        private string _employeeName = "";

        #endregion

        #region 控件变量及初始化

        private WnToolBar wnToolBarMonthTask;
        private WnTbLabel tsLabelWorkYear;
        private WnCombo tsComboWorkYear;
        private WnTbLabel tsLabelWorkMonth;
        private WnCombo tsComboWorkMonth;
        private WnTbButton tsButtonMonthTaskGet;
        private WnSeparator tsSeparator6;
        private WnTbButton tsButtonMonthTaskSave;
        private WnSeparator tsSeparator7;
        private WnTbButton tsButtonMonthTaskCancel;
        private WnTbLabel tsLabelAudit;
        private WnSeparator tsSeparator9;
        private WnTbButton tsButtonMonthTaskHelp;
        private WnSeparator tsSeparator10;
        private WnTbButton tsButtonMonthTaskExit;
        private WnSplitPanel wnSplitPanelMonthTask;
        private WnSplitPanel wnSplitPanelMonthTaskSummary;
        private WnGrid wnGridMonthTaskSummary;
        private WnGrid wnGridMonthTaskDetail;
        private WnText wnTextMonthSelfEvaluate;

        /// <summary>
        /// 控件变量初始化方法
        /// </summary>
        private void InitializeObject()
        {
            wnToolBarMonthTask = this.AllObjects["wnToolBarMonthTask"] as WnToolBar;
            tsLabelWorkYear = this.AllObjects["tsLabelWorkYear"] as WnTbLabel;
            tsComboWorkYear = this.AllObjects["tsComboWorkYear"] as WnCombo;
            tsLabelWorkMonth = this.AllObjects["tsLabelWorkMonth"] as WnTbLabel;
            tsComboWorkMonth = this.AllObjects["tsComboWorkMonth"] as WnCombo;
            tsButtonMonthTaskGet = this.AllObjects["tsButtonMonthTaskGet"] as WnTbButton;
            tsSeparator6 = this.AllObjects["tsSeparator6"] as WnSeparator;
            tsButtonMonthTaskSave = this.AllObjects["tsButtonMonthTaskSave"] as WnTbButton;
            tsSeparator7 = this.AllObjects["tsSeparator7"] as WnSeparator;
            tsButtonMonthTaskCancel = this.AllObjects["tsButtonMonthTaskCancel"] as WnTbButton;
            tsLabelAudit = this.AllObjects["tsLabelAudit"] as WnTbLabel;
            tsSeparator9 = this.AllObjects["tsSeparator9"] as WnSeparator;
            tsButtonMonthTaskHelp = this.AllObjects["tsButtonMonthTaskHelp"] as WnTbButton;
            tsSeparator10 = this.AllObjects["tsSeparator10"] as WnSeparator;
            tsButtonMonthTaskExit = this.AllObjects["tsButtonMonthTaskExit"] as WnTbButton;
            wnSplitPanelMonthTask = this.AllObjects["wnSplitPanelMonthTask"] as WnSplitPanel;
            wnSplitPanelMonthTaskSummary = this.AllObjects["wnSplitPanelMonthTaskSummary"] as WnSplitPanel;
            wnGridMonthTaskSummary = this.AllObjects["wnGridMonthTaskSummary"] as WnGrid;
            wnGridMonthTaskDetail = this.AllObjects["wnGridMonthTaskDetail"] as WnGrid;
            wnTextMonthSelfEvaluate = this.AllObjects["wnTextMonthSelfEvaluate"] as WnText;
        }

        #endregion

        #region 窗体方法

        /// <summary>
        /// 构造函数
        /// </summary>
        public WnFormMonthTaskEvaluate()
        {
            InitializeComponent();

            /* 该属性设置为true，确保在窗体的控件上按键时，优先触发窗体的按键事件 */
            this.KeyPreview = true;

            /* 注册窗体事件 */
            this.AfterAssemble += WnFormMonthTaskEvaluate_AfterAssemble;
            this.KeyDown += WnFormMonthTaskEvaluate_KeyDown;
            this.FormClosing += WnFormMonthTaskEvaluate_FormClosing;
        }

        /// <summary>
        /// 响应窗体的AfterAssemble事件
        /// </summary>
        private void WnFormMonthTaskEvaluate_AfterAssemble(object sender, CancelEventArgs e)
        {
            /* 初始化控件 */
            InitializeObject();

            /* 注册控件事件 */
            tsComboWorkYear.ValueChanging += tsComboWorkYear_ValueChanging;
            tsComboWorkMonth.ValueChanging += tsComboWorkMonth_ValueChanging;
            tsButtonMonthTaskGet.Click += tsButtonMonthTaskGet_Click;
            tsButtonMonthTaskSave.Click += tsButtonMonthTaskSave_Click;
            tsButtonMonthTaskCancel.Click += tsButtonMonthTaskCancel_Click;
            tsButtonMonthTaskHelp.Click += tsButtonMonthTaskHelp_Click;
            tsButtonMonthTaskExit.Click += tsButtonMonthTaskExit_Click;

            /* 获取当前用户的EmployeeName */
            Hashtable htName = new Hashtable() { { "DataSetName", "GetEmployeeName" } };
            htName["{UserGuid}"] = AppInfo.UserGuid;
            if (this.DataSource.GetDataSet(htName)["IsSuccess"].Equals("0")) { e.Cancel = true; return; }

            /* 如果没有获取到EmployeeName，提示用户后返回 */
            DataTable dtName = (this.DataSource.DataSets["GetEmployeeName"] as DataSet).Tables["T01"];
            if (dtName.Rows.Count <= 0)
            {
                DialogBox.ShowError(string.Format(this.GetCurrLanguageContent("WnFormMain.NoEmployeeName"), AppInfo.UserName)); //Content_CN：未获取到获取当前用户“{0}”的EmployeeName，请联系系统管理员！
                e.Cancel = true;
                return;
            }

            /* 为全局变量赋值 */
            this._employeeName = dtName.Rows[0]["EmployeeName"].ToString();

            /* 获取当前工作日期的核算年月 */
            Hashtable htWorkDate = new Hashtable() { { "DataSetName", "GetWorkPeriod" }, { "TableNames", "Date" } };
            if (this.DataSource.GetDataSet(htWorkDate)["IsSuccess"].Equals("0")) { e.Cancel = true; return; }

            /* 如果没有获取到绩效年月，提示用户后返回 */
            DataTable dtWorkDate = (this.DataSource.DataSets["GetWorkPeriod"] as DataSet).Tables["Date"];
            if (dtWorkDate.Rows.Count <= 0)
            {
                DialogBox.ShowError(this.GetCurrLanguageContent("WnFormMonthTaskEvaluate.NoWorkPeriod")); //Content_CN：未获取到当前工作日期的绩效月，请联系系统管理员！
                e.Cancel = true;
                return;
            }

            /* 设置工作年度控件的数据源，当前绩效年前推三年 */
            StringBuilder workYears = new StringBuilder(dtWorkDate.Rows[0]["Year"].ToString());
            for (int i = 1; i < 3; i++)
            {
                workYears.Append("|" + (Convert.ToInt32(dtWorkDate.Rows[0]["Year"]) - i));
            }
            tsComboWorkYear.ComboString = workYears.ToString();

            /* 设置工作年度、工作月份控件的值 */
            tsComboWorkYear.Value = dtWorkDate.Rows[0]["Year"].ToString();
            tsComboWorkMonth.Value = dtWorkDate.Rows[0]["Month"].ToString();

            /* 设置标准前景色 */
            wnGridMonthTaskSummary.SetForeColor("RowState = '0' || RowState = '10' || RowState = '11'", "", Color.FromArgb(0, 128, 0)); //绿色
            wnGridMonthTaskSummary.SetForeColor("AuditUser is not null", "", Color.FromArgb(230, 100, 0)); //已审核，橙色
            wnGridMonthTaskDetail.SetForeColor("RowState = '0' || RowState = '10' || RowState = '11'", "", Color.FromArgb(0, 128, 0)); //绿色
            wnGridMonthTaskDetail.SetForeColor("AuditUser is not null", "", Color.FromArgb(230, 100, 0)); //已审核，橙色

            /* 获取当前绩效月的月度任务总结数据 */
            GetData();
        }

        /// <summary>
        /// 响应窗体的KeyDown事件
        /// </summary>
        private void WnFormMonthTaskEvaluate_KeyDown(object sender, KeyEventArgs e)
        {
            /* 保存按钮可用时，Ctrl+S：保存数据 */
            if (tsButtonMonthTaskSave.Enabled && e.Control && e.KeyCode == Keys.S)
            {
                SaveData(false);
            }
        }

        /// <summary>
        /// 响应窗体的FormClosing事件
        /// </summary>
        private void WnFormMonthTaskEvaluate_FormClosing(object sender, FormClosingEventArgs e)
        {
            /* 判断是否有未保存的值，如果有，取消关闭 */
            if (!SaveData()) { e.Cancel = true; }
        }

        #endregion

        #region 对象方法

        /// <summary>
        /// 响应tsComboWorkYear的ValueChanging事件
        /// </summary>
        private void tsComboWorkYear_ValueChanging(object sender, ValueChangingArgs e)
        {
            /* 判断是否有未保存的值，如果有，取消切换值 */
            if (!SaveData())
            {
                tsComboWorkYear.Value = e.OldValue;
                return;
            }

            /* 重新获取数据 */
            GetData();
        }

        /// <summary>
        /// 响应tsComboWorkMonth的ValueChanging事件
        /// </summary>
        private void tsComboWorkMonth_ValueChanging(object sender, ValueChangingArgs e)
        {
            /* 判断是否有未保存的值，如果有，取消切换值 */
            if (!SaveData())
            {
                tsComboWorkMonth.Value = e.OldValue;
                return;
            }

            /* 重新获取数据 */
            GetData();
        }

        /// <summary>
        /// 响应“提取”按钮的Click事件
        /// </summary>
        private void tsButtonMonthTaskGet_Click(object sender, EventArgs e)
        {
            /* 提取当前绩效月的数据 */
            HandleData();
        }

        /// <summary>
        /// 响应“保存”按钮的Click事件
        /// </summary>
        private void tsButtonMonthTaskSave_Click(object sender, EventArgs e)
        {
            /* 保存数据 */
            SaveData(false);
        }

        /// <summary>
        /// 响应“撤销”按钮的Click事件
        /// </summary>
        private void tsButtonMonthTaskCancel_Click(object sender, EventArgs e)
        {
            /* 如果当前有行且文本框不为空，向用户确认撤销 */
            if (wnGridMonthTaskSummary.RowCount > 0 && wnTextMonthSelfEvaluate.Value != null && !string.IsNullOrEmpty(wnTextMonthSelfEvaluate.Value.ToString()))
            {
                if (DialogBox.AskYesNoCancel(this.GetCurrLanguageContent("WnFormMonthTaskEvaluate.AskCancel")) != DialogResult.Yes) { return; } //Content_CN：撤销后无法再获取已编辑的个人总结，是否确认撤销本月数据？
            }

            /* 从数据库中删除本月数据 */
            Hashtable htCancel = new Hashtable() { { "DataSetName", "CancelMonthTaskData" } };
            htCancel["{WorkYear}"] = tsComboWorkYear.Value.ToString();
            htCancel["{WorkMonth}"] = tsComboWorkMonth.Value.ToString();
            htCancel["{Executor}"] = this._employeeName;
            if (this.DataSource.ExecSql(htCancel)["IsSuccess"].Equals("0"))
            {
                DialogBox.ShowError(this.GetCurrLanguageContent("WnFormMonthTaskEvaluate.CancelError")); //Content_CN：撤销月度绩效自评数据出错，请重新操作！
                return;
            }

            /* 清空界面两个Grid的所有行 */
            wnGridMonthTaskSummary.DataSource.Clear();
            wnGridMonthTaskDetail.DataSource.Clear();

            /* 撤销成功后 提按钮权限 */
            SetControlRight();
        }

        /// <summary>
        /// 响应“帮助”按钮的Click事件
        /// </summary>
        private void tsButtonMonthTaskHelp_Click(object sender, EventArgs e)
        {
            /* 若具有帮助权限，则弹出帮助窗体且可编辑，否则，只显示帮助窗体 */
            if (this.RightInfo.Contains("EditHelp"))
            {
                HelpInfo.Edit(this.FuncFrameGuid);
            }
            else
            {
                HelpInfo.Show(this.FuncFrameGuid);
            }
        }

        /// <summary>
        /// 响应“退出”按钮的Click事件
        /// </summary>
        private void tsButtonMonthTaskExit_Click(object sender, EventArgs e)
        {
            /* 关闭窗体 */
            this.Close();
        }

        #endregion

        #region 内部方法

        /// <summary>
        /// 获取当前绩效月的月度任务总结数据
        /// </summary>
        private void GetData()
        {
            /* 获取当前账期的月度任务总结数据 */
            Hashtable htSummary = new Hashtable() { { "DataSetName", "GetMonthTaskData" }, { "TableNames", new string[] { "Summary", "Detail" } } };
            htSummary["{EditTime}"] = DateTime.Now.ToString();
            htSummary["{WorkYear}"] = tsComboWorkYear.Value.ToString();
            htSummary["{WorkMonth}"] = tsComboWorkMonth.Value.ToString();
            htSummary["{Executor}"] = this._employeeName;
            if (this.DataSource.GetDataSet(htSummary)["IsSuccess"].Equals("0"))
            {
                /* 获取失败，设置按钮均不可用 */
                tsButtonMonthTaskGet.Enabled = tsButtonMonthTaskSave.Enabled = tsButtonMonthTaskCancel.Enabled = false;
            }

            /* 获取成功后设置控件权限 */
            SetControlRight();
        }

        /// <summary>
        /// 提取当前绩效月的月度任务总结数据
        /// </summary>
        private void HandleData()
        {
            /* 若当前总结记录不为空，直接返回 */
            if (wnGridMonthTaskSummary.RowCount > 0) { return; }

            /* 根据当前日期获取当前绩效月的起止时间 */
            Hashtable htWorkDate = new Hashtable() { { "DataSetName", "GetWorkPeriod" }, { "TableNames", "Period" } };
            if (this.DataSource.GetDataSet(htWorkDate)["IsSuccess"].Equals("0")) { return; }

            /* 如果没有获取到起止时间，提示用户后返回 */
            DataTable dtPeriod = (this.DataSource.DataSets["GetWorkPeriod"] as DataSet).Tables["Period"];
            if (dtPeriod.Rows.Count <= 0)
            {
                DialogBox.ShowError(this.GetCurrLanguageContent("WnFormMonthTaskEvaluate.NoPeriodRange")); //Content_CN：未获取到当前绩效月的起止时间，请联系系统管理员！
                return;
            }

            /* 判断是否存在空的绩效分数据 */
            Hashtable htCheck = new Hashtable() { { "DataSetName", "CheckPerformanceScore" } };
            htCheck["{Executor}"] = this._employeeName;
            htCheck["{StartDate}"] = ((DateTime)dtPeriod.Rows[0]["StartDate"]).ToString("yyyy/MM/dd HH:mm:ss");
            htCheck["{EndDate}"] = ((DateTime)dtPeriod.Rows[0]["EndDate"]).ToString("yyyy/MM/dd HH:mm:ss");
            if (this.DataSource.GetDataSet(htCheck)["IsSuccess"].Equals("0")) { return; }
            DataTable dtCheck = (this.DataSource.DataSets["CheckPerformanceScore"] as DataSet).Tables["T01"];
            if (dtCheck.Rows.Count > 0)
            {
                DialogBox.ShowMessage(this.GetCurrLanguageContent("WnFormMonthTaskEvaluate.NullPerformanceScore")); //Content_CN：当前绩效月存在绩效分为空的任务，请确认本月任务均已评定后重新提取月度数据！
                return;
            }

            /* 执行数据集获取月度任务总结数据 */
            Hashtable htSummary = new Hashtable() { { "DataSetName", "GetMonthTaskSummary" } };
            htSummary["{Executor}"] = this._employeeName;
            htSummary["{StartDate}"] = ((DateTime)dtPeriod.Rows[0]["StartDate"]).ToString("yyyy/MM/dd HH:mm:ss");
            htSummary["{EndDate}"] = ((DateTime)dtPeriod.Rows[0]["EndDate"]).ToString("yyyy/MM/dd HH:mm:ss");
            if (this.DataSource.GetDataSet(htSummary)["IsSuccess"].Equals("0")) { return; }
            DataTable dtSummary = (this.DataSource.DataSets["GetMonthTaskSummary"] as DataSet).Tables["T01"];
            if (dtSummary.Rows[0]["WorkHours"] == DBNull.Value || dtSummary.Rows[0]["PerformanceScore"] == DBNull.Value || dtSummary.Rows[0]["AveWorkEffectRatio"] == DBNull.Value || dtSummary.Rows[0]["UnitHourScore"] == DBNull.Value)
            {
                DialogBox.ShowMessage(this.GetCurrLanguageContent("WnFormMonthTaskEvaluate.NoEvaluateData")); //Content_CN：未获取到当前用户在当前绩效月下的总结数据，请检查数据后重试！
                return;
            }

            /* 新增一行，将数据放到新增行中，并设置行状态 */
            DataRow drSummary = wnGridMonthTaskSummary.AddRow();
            drSummary["WorkYear"] = tsComboWorkYear.Value.ToString();
            drSummary["WorkMonth"] = tsComboWorkMonth.Value.ToString();
            drSummary["Executor"] = this._employeeName;
            drSummary["WorkHours"] = Convert.ToDecimal(dtSummary.Rows[0]["WorkHours"]);
            drSummary["PerformanceScore"] = Convert.ToDecimal(dtSummary.Rows[0]["PerformanceScore"]);
            drSummary["AveWorkEffectRatio"] = Convert.ToDecimal(dtSummary.Rows[0]["AveWorkEffectRatio"]);
            drSummary["UnitHourScore"] = Convert.ToDecimal(dtSummary.Rows[0]["UnitHourScore"]);
            drSummary["EditTime"] = DateTime.Now.ToString();
            drSummary["RowState"] = "10";

            /* 获取子表数据 */
            Hashtable htDetail = new Hashtable() { { "DataSetName", "GetMonthTaskDetail" } };
            htDetail["{Executor}"] = this._employeeName;
            htDetail["{StartDate}"] = ((DateTime)dtPeriod.Rows[0]["StartDate"]).ToString("yyyy/MM/dd HH:mm:ss");
            htDetail["{EndDate}"] = ((DateTime)dtPeriod.Rows[0]["EndDate"]).ToString("yyyy/MM/dd HH:mm:ss");
            if (this.DataSource.GetDataSet(htDetail)["IsSuccess"].Equals("0")) { return; }
            DataTable dtDetail = (this.DataSource.DataSets["GetMonthTaskDetail"] as DataSet).Tables["T01"];
            if (dtDetail.Rows.Count <= 0)
            {
                DialogBox.ShowMessage(this.GetCurrLanguageContent("WnFormMonthTaskEvaluate.NoEvaluateData")); //Content_CN：未获取到当前用户在当前绩效月下的总结数据，请检查数据后重试！
                return;
            }

            foreach (DataRow dr in dtDetail.Rows)
            {
                DataRow drDetail = wnGridMonthTaskDetail.AddRow();
                drDetail["WorkYear"] = tsComboWorkYear.Value.ToString();
                drDetail["WorkMonth"] = tsComboWorkMonth.Value.ToString();
                drDetail["Executor"] = this._employeeName;
                drDetail["WorkThemeGuid"] = dr["WorkThemeGuid"].ToString();
                drDetail["WorkThemeName"] = dr["WorkThemeName"].ToString();
                drDetail["WorkTaskGuid"] = dr["WorkTaskGuid"].ToString();
                drDetail["WorkTaskName"] = dr["WorkTaskName"].ToString();
                drDetail["WorkHours"] = dr["WorkHours"].ToString();
                drDetail["RowState"] = "10";
            }

            /* 插入数据成功后设置按钮权限 */
            SetControlRight();
        }

        /// <summary>
        /// 设置控件权限
        /// </summary>
        private void SetControlRight()
        {
            /* 根据是否有数据判断各个控件的值 */
            tsButtonMonthTaskGet.Enabled = wnGridMonthTaskSummary.RowCount == 0 && wnGridMonthTaskDetail.RowCount == 0;
            tsButtonMonthTaskCancel.Enabled = !tsButtonMonthTaskGet.Enabled;
            wnTextMonthSelfEvaluate.ControlReadOnly = wnGridMonthTaskSummary.RowCount > 0 && !string.IsNullOrEmpty(wnGridMonthTaskSummary.CurrRow["AuditUser"].ToString());
        }

        /// <summary>
        /// 保存月度任务总结数据
        /// </summary>
        private bool SaveData(bool askSave = true)
        {
            /* 判断当前是否有数据，如果没有数据，直接返回 */
            if (wnGridMonthTaskSummary.RowCount == 0 && wnGridMonthTaskDetail.RowCount == 0) { return true; }

            /* 提交数据 */
            Hashtable htSave = new Hashtable() { { "DataSetName", "GetMonthTaskData" }, { "TableNames", new string[] { "Summary", "Detail" } } };
            htSave["ActionStep"] = new List<Act>() { Act.CommitEdit };
            if (this.DataSource.UpdateDataSet(htSave)["IsChanged"].Equals("0")) { return true; }

            /* 询问用户是否保存 */
            if (askSave)
            {
                htSave["ActionStep"] = new List<Act>() { Act.AskIsUpdate };
                Hashtable result = this.DataSource.UpdateDataSet(htSave);
                if (result["IsUpdate"].Equals("-1")) { return false; } //取消
                else if (result["IsUpdate"].Equals("0")) //否
                {
                    wnGridMonthTaskSummary.RejectChanges();
                    return true;
                }
            }

            /* 判断文本框是否有值，没有值则提示用户后返回 */
            if (wnTextMonthSelfEvaluate.Value == null || string.IsNullOrEmpty(wnTextMonthSelfEvaluate.Value.ToString()))
            {
                DialogBox.ShowError(this.GetCurrLanguageContent("WnFormMonthTaskEvaluate.NoSelfEvaluate")); //Content_CN：当前个人总结为空，请编辑个人总结或撤销当前数据后再次操作！
                return false;
            }

            /* 保存数据 */
            htSave["ActionStep"] = new List<Act>() { Act.SaveData };
            if (this.DataSource.UpdateDataSet(htSave)["IsSuccess"].Equals("0"))
            {
                return false;
            }
            return true;
        }

        #endregion
    }
}