/**************************
 * 管控对象：每日绩效自评
 * 管控成员：冯旭
 *************************/

using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid.Views.Grid;
using PEMSoft.Client.SYS.WnControl;
using PEMSoft.Client.SYS.WnPlatform;

namespace PEMSoft.Client.IFM.Jxgl.WnDayPerformanceSelfEvaluate
{
    public partial class WnFormMain : WnForm
    {
        #region 变量

        /// <summary>
        /// 记录是否成功获取数据
        /// </summary>
        private bool _isSuccess;

        /// <summary>
        /// 记录年度控件的值
        /// </summary>
        private string _year;

        /// <summary>
        /// 记录工作日期控件的值
        /// </summary>
        private string _workDate;

        /// <summary>
        /// 记录参数信息的起始日期
        /// </summary>
        private string _startDate;

        /// <summary>
        /// 储存wnGroupEvaluate可编辑的单值控件，用于控制这些控件是否可编辑
        /// </summary>
        private List<ISingleControl> _listEditableControls = new List<ISingleControl>();

        /// <summary>
        /// 记录工作任务列背后的编辑器，展开/关闭下拉列表切换控件数据源时使用
        /// </summary>
        private RepositoryItemGridLookUpEdit _gridLookupCode;

        /// <summary>
        /// 存放当前用户的EmployeeName
        /// </summary>
        private string _employeeName = "";

        /// <summary>
        /// 存放当前用户的EmployeeGuid
        /// </summary>
        private string _employeeGuid = "";

        /// <summary>
        /// 存放当前用户的工作日历
        /// </summary>
        private DataTable _dtCalendar = null;

        #endregion

        #region 控件变量及初始化

        private WnToolBar wnToolBarMain;
        private WnTbLabel tsLabelYear;
        private WnCombo tsComboYear;
        private WnTbLabel tsLabelDate;
        private WnDate tsDateWorkDate;
        private WnTbButton tsButtonLoad;
        private WnSeparator tsSeparator1;
        private WnTbButton tsButtonAppend;
        private WnTbButton tsButtonDelete;
        private WnSeparator tsSeparator2;
        private WnTbButton tsButtonSave;
        private WnSeparator tsSeparator3;
        private WnTbButton tsButtonHelp;
        private WnSeparator tsSeparator4;
        private WnTbButton tsButtonExit;
        private WnTbButton tsButtonMonthTaskEvaluate;
        private WnSeparator tsSeparator8;
        private WnTbLabel tsLabelEvaluate;
        private WnSplitPanel wnSplitPanelMain;
        private WnSplitPanel wnSplitPanelWork;
        private WnGrid wnGridReport;
        private WnGrid wnGridTask;
        private WnToolBar wnToolBarEvaluate;
        private WnTbButton tsButtonHandle;
        private WnGrid wnGridEvaluate;
        private WnGroup wnGroupEvaluate;
        private WnDecimal wnDecimalProcess;
        private WnText wnTextResult;
        private WnText wnTextEvaluate;
        private WnGridLookup wnGridLookupServiceObject;
        private WnText wnTextServiceObjectNote;

        /// <summary>
        /// 控件变量初始化方法
        /// </summary>
        private void InitializeObject()
        {
            wnToolBarMain = this.AllObjects["wnToolBarMain"] as WnToolBar;
            tsLabelYear = this.AllObjects["tsLabelYear"] as WnTbLabel;
            tsComboYear = this.AllObjects["tsComboYear"] as WnCombo;
            tsLabelDate = this.AllObjects["tsLabelDate"] as WnTbLabel;
            tsDateWorkDate = this.AllObjects["tsDateWorkDate"] as WnDate;
            tsButtonLoad = this.AllObjects["tsButtonLoad"] as WnTbButton;
            tsSeparator1 = this.AllObjects["tsSeparator1"] as WnSeparator;
            tsButtonAppend = this.AllObjects["tsButtonAppend"] as WnTbButton;
            tsButtonDelete = this.AllObjects["tsButtonDelete"] as WnTbButton;
            tsSeparator2 = this.AllObjects["tsSeparator2"] as WnSeparator;
            tsButtonSave = this.AllObjects["tsButtonSave"] as WnTbButton;
            tsSeparator3 = this.AllObjects["tsSeparator3"] as WnSeparator;
            tsButtonHelp = this.AllObjects["tsButtonHelp"] as WnTbButton;
            tsSeparator4 = this.AllObjects["tsSeparator4"] as WnSeparator;
            tsButtonExit = this.AllObjects["tsButtonExit"] as WnTbButton;
            tsButtonMonthTaskEvaluate = this.AllObjects["tsButtonMonthTaskEvaluate"] as WnTbButton;
            tsSeparator8 = this.AllObjects["tsSeparator8"] as WnSeparator;
            tsLabelEvaluate = this.AllObjects["tsLabelEvaluate"] as WnTbLabel;
            wnSplitPanelMain = this.AllObjects["wnSplitPanelMain"] as WnSplitPanel;
            wnSplitPanelWork = this.AllObjects["wnSplitPanelWork"] as WnSplitPanel;
            wnGridReport = this.AllObjects["wnGridReport"] as WnGrid;
            wnGridTask = this.AllObjects["wnGridTask"] as WnGrid;
            wnToolBarEvaluate = this.AllObjects["wnToolBarEvaluate"] as WnToolBar;
            tsButtonHandle = this.AllObjects["tsButtonHandle"] as WnTbButton;
            wnGridEvaluate = this.AllObjects["wnGridEvaluate"] as WnGrid;
            wnGroupEvaluate = this.AllObjects["wnGroupEvaluate"] as WnGroup;
            wnDecimalProcess = this.AllObjects["wnDecimalProcess"] as WnDecimal;
            wnTextResult = this.AllObjects["wnTextResult"] as WnText;
            wnTextEvaluate = this.AllObjects["wnTextEvaluate"] as WnText;
            wnGridLookupServiceObject = this.AllObjects["wnGridLookupServiceObject"] as WnGridLookup;
            wnTextServiceObjectNote = this.AllObjects["wnTextServiceObjectNote"] as WnText;
        }

        #endregion

        #region 窗体方法

        /// <summary>
        /// 构造函数
        /// </summary>
        public WnFormMain()
        {
            InitializeComponent();

            /* 设置该属性为true，确保在窗体的控件上按键时，优先触发窗体的按键事件 */
            this.KeyPreview = true;

            /* 注册窗体事件 */
            this.AfterAssemble += WnFormMain_AfterAssemble;
            this.Shown += WnFormMain_Shown;
            this.KeyDown += WnFormMain_KeyDown;
            this.FormClosing += WnFormMain_FormClosing;
            this.SizeChanged += WnFormMain_SizeChanged;
        }

        /// <summary>
        /// 响应窗体的AfterAssemble事件
        /// </summary>
        private void WnFormMain_AfterAssemble(object sender, CancelEventArgs e)
        {
            /* 初始化控件 */
            InitializeObject();

            /* 注册控件事件 */
            tsComboYear.ValueChanging += tsComboYear_ValueChanging;
            tsDateWorkDate.ValueChanging += tsDateWorkDate_ValueChanging;
            tsButtonLoad.Click += tsButtonLoad_Click;
            tsButtonAppend.Click += tsButtonAppend_Click;
            tsButtonDelete.Click += tsButtonDelete_Click;
            tsButtonSave.Click += tsButtonSave_Click;
            tsButtonHelp.Click += tsButtonHelp_Click;
            tsButtonExit.Click += tsButtonExit_Click;
            tsButtonMonthTaskEvaluate.Click += tsButtonMonthTaskEvaluate_Click;
            tsButtonHandle.Click += tsButtonHandle_Click;
            wnGridTask.GridView.ShownEditor += wnGridTask_ShownEditor;
            wnGridReport.CellValueChanged += wnGridReport_CellValueChanged;
            wnGridReport.CurrRowChanged += wnGridReport_CurrRowChanged;
            wnGridEvaluate.CurrRowChanged += wnGridEvaluate_CurrRowChanged;
            ((wnTextResult.OriginControl as WnTextEditorBase).Controls[0] as RichTextBox).DoubleClick += wnTextResult_DoubleClick;

            /* 获取部门运行效率 */
            Hashtable htDept = new Hashtable();
            htDept["DataSetName"] = "GetDeptEfficiency";
            if (this.DataSource.GetDataSet(htDept)["IsSuccess"].Equals("0")) { e.Cancel = true; return; }
            DataTable dtDept = (this.DataSource.DataSets["GetDeptEfficiency"] as DataSet).Tables["T01"];
            if (dtDept.Rows.Count == 0)
            {
                /* 没有获取到数据，弹出提示，取消打开窗体并返回 */
                DialogBox.ShowWarning(this.GetCurrLanguageContent("WnFormMain.GetDeptEfficiencyError")); //Content_CN：没有获取到部门运行效率，请联系管理员！
                e.Cancel = true;
                return;
            }

            /* 获取当前用户的EmployeeName */
            Hashtable htName = new Hashtable() { { "DataSetName", "GetEmployeeName" } };
            //htName["UserGuid"] = AppInfo.UserGuid;
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
            this._employeeGuid = dtName.Rows[0]["EmployeeGuid"].ToString();

            /* 获取当前用户的岗位信息及能力等级 */
            Hashtable htPost = new Hashtable();
            htPost["DataSetName"] = "GetPost";
            htPost["{EmployeeName}"] = this._employeeName;
            if (this.DataSource.GetDataSet(htPost)["IsSuccess"].Equals("0")) { e.Cancel = true; return; }
            DataTable dtPost = (this.DataSource.DataSets["GetPost"] as DataSet).Tables["T01"];
            if (dtPost.Rows.Count == 0)
            {
                /* 没有获取到数据，弹出提示，取消打开窗体并返回 */
                DialogBox.ShowWarning(this.GetCurrLanguageContent("WnFormMain.GetPostError")); //Content_CN：没有获取到当前用户的岗位信息，请联系管理员！
                e.Cancel = true;
                return;
            }

            /* 获取当前日期，若获取失败则取消打开窗体 */
            object date = AppServer.GetDate();
            if (date == null) { e.Cancel = true; return; }

            /* 获取系统参数设置的日期，若没有设置则使用默认值27号 */
            _startDate = this.ParamInfo["StartDate"] + "";
            if (_startDate == "") { _startDate = "27"; }

            /* 获取年、月、日的值 */
            int nowYear = ((DateTime)date).Year;
            int nowMonth = ((DateTime)date).Month;
            int nowDay = ((DateTime)date).Day;

            /* 根据系统参数设置的日期更新年份的值 */
            nowYear = nowMonth == 12 ? (nowDay >= Convert.ToInt16(_startDate) ? (nowYear + 1) : nowYear) : nowYear;

            /* 设置年度控件的可选年份和默认值 */
            tsComboYear.ComboString = (nowYear - 1) + "|" + nowYear + "|" + (nowYear + 1);
            tsComboYear.Value = nowYear.ToString();

            /* 设置工作日期控件的可选范围和默认值 */
            tsDateWorkDate.MinValue = Convert.ToDateTime((nowYear - 1) + "-12-" + _startDate);
            tsDateWorkDate.MaxValue = Convert.ToDateTime(nowYear + "-12-" + (Convert.ToInt16(_startDate) - 1));
            tsDateWorkDate.Value = (DateTime)date;

            /* 设置wnGroup中工作进度单值控件的最小值和最大值 */
            wnDecimalProcess.MinValue = 0;
            wnDecimalProcess.MaxValue = 1;

            /* 获取可编辑的单值控件，用来控制单值控件是否可编辑 */
            foreach (Control control in wnGroupEvaluate.Controls)
            {
                if (control is ISingleControl && !(control as ISingleControl).ControlReadOnly)
                {
                    _listEditableControls.Add(control as ISingleControl);
                }
            }

            /* 加载数据 */
            GetData();

            /* 设置“每日任务汇报”和“每日任务评价”新增行、新增修改行、修改行的前景色 */
            wnGridReport.SetForeColor("RowState = '0' || RowState = '10' || RowState = '11'", "", Color.FromArgb(0, 128, 0));
            wnGridEvaluate.SetForeColor("RowState = '11'", "", Color.FromArgb(0, 128, 0));

            /* 设置“每日任务汇报”、“工作任务”、“每日任务评价”已完结行的前景色 */
            wnGridReport.SetForeColor("SelfFinishTime is not null", "", Color.FromArgb(230, 100, 0));
            wnGridTask.SetForeColor("SelfFinishTime is not null", "", Color.FromArgb(230, 100, 0));
            wnGridEvaluate.SetForeColor("SelfFinishTime is not null", "", Color.FromArgb(230, 100, 0));
        }

        /// <summary>
        /// 响应窗体的Shown事件
        /// </summary>
        private void WnFormMain_Shown(object sender, EventArgs e)
        {

            /* 获取当前工作日期的核算年月 */
            Hashtable htWorkDate = new Hashtable() { { "DataSetName", "GetWorkPeriod" }, { "TableNames", "Date" } };
            if (this.DataSource.GetDataSet(htWorkDate)["IsSuccess"].Equals("0")) { return; }

            /* 如果没有获取到绩效年月，提示用户后返回 */
            DataTable dtWorkDate = (this.DataSource.DataSets["GetWorkPeriod"] as DataSet).Tables["Date"];
            if (dtWorkDate.Rows.Count <= 0)
            {
                DialogBox.ShowError(this.GetCurrLanguageContent("WnFormMain.NoWorkDate")); //Content_CN：未获取到当前工作日期的绩效月，请联系系统管理员！
                return;
            }

            /* 获取当前月份的起止时间 */
            Hashtable htPeriod = new Hashtable() { { "DataSetName", "GetWorkPeriod" }, { "TableNames", "Period" }, { "{WorkYear}", dtWorkDate.Rows[0]["Year"].ToString() }, { "{WorkMonth}", dtWorkDate.Rows[0]["Month"].ToString() } };
            if (this.DataSource.GetDataSet(htPeriod)["IsSuccess"].Equals("0")) { return; }

            /* 如果没有获取到起止时间，提示用户后返回 */
            DataTable dtPeriod = (this.DataSource.DataSets["GetWorkPeriod"] as DataSet).Tables["Period"];
            if (dtPeriod.Rows.Count <= 0)
            {
                DialogBox.ShowError(this.GetCurrLanguageContent("WnFormMain.NoWorkPeriod")); //Content_CN：未获取到当前工作日期的起止时间，请联系系统管理员！
                return;
            }

            /* 检查当前月份是否存在未汇报且未请假的时间段 */
            Hashtable htCheck = new Hashtable();
            htCheck["DataSetName"] = "CheckMonthTaskData";
            htCheck["{Year}"] = dtWorkDate.Rows[0]["Year"].ToString();
            htCheck["{Month}"] = dtWorkDate.Rows[0]["Month"].ToString();
            htCheck["{StartDate}"] = dtPeriod.Rows[0]["StartDate"].ToString();
            htCheck["{EndDate}"] = dtPeriod.Rows[0]["EndDate"].ToString();
            htCheck["{YgGuid}"] = _employeeGuid;
            if (this.DataSource.GetDataSet(htCheck)["IsSuccess"].Equals("0")) { return; }
            DataTable dtCheck = (this.DataSource.DataSets["CheckMonthTaskData"] as DataSet).Tables["T01"];
            if (dtCheck.Rows.Count > 0)
            {
                string workDate = "";
                foreach (DataRow dr in dtCheck.Rows)
                {
                    workDate += "【" + dr["WorkDate"].ToString() + "】,";
                }
                DialogBox.ShowMessage(string.Format(this.GetCurrLanguageContent("WnFormMain.WorkTimeLack"), workDate.TrimEnd(','))); //Content_CN：工作日期：【{0}】的工作存在异常，请前往该日期汇报工作或填写请假单！
            }
        }

        /// <summary>
        /// 响应窗体的KeyDown事件
        /// </summary>
        private void WnFormMain_KeyDown(object sender, KeyEventArgs e)
        {
            /* 保存按钮可用并且用户按下Ctrl+S时执行保存 */
            if (tsButtonSave.Enabled && e.Control && e.KeyCode == Keys.S) { tsButtonSave_Click(null, null); }

            /* 用户按下Alt+V时隐藏或显示绩效自评窗体 */
            if (e.Alt && e.KeyCode == Keys.V) { wnGroupEvaluate.Visible = !wnGroupEvaluate.Visible; }
        }

        /// <summary>
        /// 响应窗体的FormClosing事件
        /// </summary>
        private void WnFormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            /* 当保存按钮可用时，检查是否有未保存的数据，若取消保存或保存失败，则取消关闭窗体 */
            if (tsButtonSave.Enabled && !SaveData(new List<Act> { Act.AskIsUpdate, Act.SaveData })) { e.Cancel = true; }
        }

        /// <summary>
        /// 响应窗体的SizeChanged事件
        /// </summary>
        private void WnFormMain_SizeChanged(object sender, EventArgs e)
        {
            /* 设置wnGroup居中显示 */
            if (wnGroupEvaluate != null)
            {
                PEMSoft.Client.SYS.WnPlatform.ControlHelper.CenterShow(wnGroupEvaluate);
            }
        }

        #endregion

        #region 对象方法

        /// <summary>
        /// 响应tsComboYear的ValueChanging事件
        /// </summary>
        private void tsComboYear_ValueChanging(object sender, ValueChangingArgs e)
        {
            /* 重新设置工作日期控件的默认值并清空工作日期控件的值 */
            tsDateWorkDate.MinValue = Convert.ToDateTime((Convert.ToInt16(e.NewValue.ToString()) - 1) + "-12-" + _startDate);
            tsDateWorkDate.MaxValue = Convert.ToDateTime(e.NewValue.ToString() + "-12-" + (Convert.ToInt16(_startDate) - 1));
            tsDateWorkDate.Value = "";

            /* 控制按钮权限 */
            SetControlRight();
        }

        /// <summary>
        /// 响应tsDateWorkDate的ValueChanging事件
        /// </summary>
        private void tsDateWorkDate_ValueChanging(object sender, ValueChangingArgs e)
        {
            /* 控制按钮权限 */
            SetControlRight();
        }

        /// <summary>
        /// 响应“加载”按钮的Click事件
        /// </summary>
        private void tsButtonLoad_Click(object sender, EventArgs e)
        {
            /* 当保存按钮可用时，检查是否有未保存的数据，若取消保存或保存失败，则返回 */
            if (tsButtonSave.Enabled && !SaveData(new List<Act> { Act.AskIsUpdate, Act.SaveData })) { return; }

            /* 清空过滤行 */
            wnGridReport.FilterString = wnGridEvaluate.FilterString = wnGridTask.FilterString = "";

            /* 加载数据 */
            GetData();
        }

        /// <summary>
        /// 响应“新增”按钮的Click事件
        /// </summary>
        private void tsButtonAppend_Click(object sender, EventArgs e)
        {
            /* 清空过滤行 */
            wnGridReport.FilterString = wnGridEvaluate.FilterString = wnGridTask.FilterString = "";

            /* 新增一行 */
            wnGridReport.AddRow();
        }

        /// <summary>
        /// 响应“删除”按钮的Click事件
        /// </summary>
        private void tsButtonDelete_Click(object sender, EventArgs e)
        {
            /* 获取当前行，若为空则返回 */
            DataRow drCurr = wnGridReport.CurrRow;
            if (drCurr == null) { return; }

            /* 当前行为新增行时，直接删除 */
            if (drCurr.RowState == DataRowState.Added)
            {
                wnGridReport.DeleteRow(drCurr);
                return;
            }

            /* 若当前行已完结，则不可删除 */
            if (!string.IsNullOrEmpty(drCurr["SelfFinishTime"].ToString()))
            {
                DialogBox.ShowMessage(this.GetCurrLanguageContent("WnFormMain.CheckDeleteHandle")); //Content_CN：当前数据已完结，不可删除！
                return;
            }

            /* 询问是否确定要删除数据 */
            if (DialogBox.AskYesNo(this.GetCurrLanguageContent("WnFormMain.AskDelete")) == DialogResult.No) { return; } //Content_CN：确定删除当前数据吗？

            /* 定义变量，存储当前行的内码和工作任务内码，若当前行为修改行，则存储其原始值 */
            string guid = drCurr["Guid"].ToString();
            string workTaskGuid = drCurr.RowState == DataRowState.Modified ? drCurr["WorkTaskGuid", DataRowVersion.Original].ToString() : drCurr["WorkTaskGuid"].ToString();

            /* 存储“每日任务汇报”中与当前行工作任务相同的记录 */
            DataRow[] workTaskArray = wnGridReport.DataSource.Select("WorkTaskGuid = '" + workTaskGuid + "' and Guid <> '" + guid + "'");

            /* 定义变量，记录更新或删除操作。若数组中没有数据，则应删除“每日任务评价”的相应数据；若有数据，则应更新“每日任务评价”相应数据的工作时长 */
            string action = workTaskArray.Count() == 0 ? "Delete" : "Update";

            /* 删除“每日任务汇报”的数据，并更新或删除“每日任务评价”的数据 */
            Hashtable htDel = new Hashtable();
            htDel["DataSetName"] = "DeleteData";
            htDel["{Guid}"] = drCurr["Guid"].ToString();
            htDel["{Action}"] = action;
            htDel["{WorkHours}"] = drCurr["WorkHours"].ToString();
            htDel["{WorkDate}"] = ((DateTime)tsDateWorkDate.Value).ToString("yyyy/MM/dd");
            htDel["{WorkTaskGuid}"] = workTaskGuid;
            if (this.DataSource.ExecSql(htDel)["IsSuccess"].Equals("0")) { return; }

            /* 获取“每日任务评价”中对应的行 */
            DataRow[] drEvaluate = wnGridEvaluate.DataSource.Select("WorkTaskGuid = '" + workTaskGuid + "'");

            /* “每日任务评价”界面中对应的行应有且只有1条 */
            if (drEvaluate.Count() == 1)
            {
                /* 当操作为删除时，删除“每日任务评价”的数据 */
                if (action == "Delete")
                {
                    wnGridEvaluate.DeleteRow(drEvaluate[0]);
                    drEvaluate[0].AcceptChanges();
                }

                /* 当操作为更新时，更新“每日任务评价”的工作时长 */
                else if (action == "Update")
                {
                    drEvaluate[0]["WorkHours"] = Convert.ToDecimal(drEvaluate[0]["WorkHours"]) - Convert.ToDecimal(drCurr["WorkHours"]);
                    drEvaluate[0].AcceptChanges();
                }
            }

            /* 删除“每日任务汇报”界面当前行 */
            wnGridReport.DeleteRow(drCurr);
            drCurr.AcceptChanges();
        }

        /// <summary>
        /// 响应“保存”按钮的Click事件
        /// </summary>
        private void tsButtonSave_Click(object sender, EventArgs e)
        {
            /* 保存数据 */
            SaveData(new List<Act> { Act.SaveData });
        }

        /// <summary>
        /// 响应“帮助”按钮的Click事件
        /// </summary>
        private void tsButtonHelp_Click(object sender, EventArgs e)
        {
            /* 若具有编辑帮助权限，则弹出帮助窗体且可编辑；否则，只显示帮助窗体，不可编辑 */
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
        private void tsButtonExit_Click(object sender, EventArgs e)
        {
            /* 关闭窗体 */
            this.Close();
        }

        /// <summary>
        /// 响应“月度绩效自评”按钮的Click事件
        /// </summary>
        private void tsButtonMonthTaskEvaluate_Click(object sender, EventArgs e)
        {
            /* 打开月度绩效自评窗体 */
            WnFormMonthTaskEvaluate wnFormMonthTaskEvaluate = new WnFormMonthTaskEvaluate();
            if (!wnFormMonthTaskEvaluate.Assemble(this, "WnFormMonthTaskEvaluate")) { return; }
            wnFormMonthTaskEvaluate.ShowDialog();
        }

        /// <summary>
        /// 响应“完结”按钮的Click事件
        /// </summary>
        private void tsButtonHandle_Click(object sender, EventArgs e)
        {
            /* 当保存按钮可用时，检查是否有未保存的数据，若取消保存或保存失败，则返回 */
            if (tsButtonSave.Enabled && !SaveData(new List<Act> { Act.AskIsUpdate, Act.SaveData })) { return; }

            /* 如果没有勾选行且当前行为空，则返回；若当前行不为空，则勾选当前行 */
            if (wnGridEvaluate.SelectedRowsArray.Length == 0)
            {
                /* 获取当前行 */
                DataRow drCurr = wnGridEvaluate.CurrRow;
                if (drCurr == null) { return; }

                /* 勾选当前行 */
                wnGridEvaluate.SetRowSelected(drCurr);
            }

            /* 检查界面是否有需要完结的数据 */
            if (wnGridEvaluate.SelectedRowsArray.AsEnumerable().Cast<DataRow>().Where(x => x["SelfFinishTime"] == DBNull.Value).Count() == 0)
            {
                /* 弹出提示，取消打钩 */
                DialogBox.ShowWarning(this.GetCurrLanguageContent("WnFormMain.NoHandle")); //Content_CN：没有需要完结的数据！
                wnGridEvaluate.UnSelectAll();
                return;
            }

            ///* 检查界面是否有未填写服务对象的数据 */
            //DataRow[] drResultArray = wnGridEvaluate.SelectedRowsArray.AsEnumerable().Cast<DataRow>().Where(x => x["SelfFinishTime"] == DBNull.Value && x["ServiceObject"] == DBNull.Value).ToArray();
            //if (drResultArray.Count() > 0)
            //{
            //    /* 弹出提示，定位到对应行 */
            //    DialogBox.ShowError(string.Format(this.GetCurrLanguageContent("WnFormMain.NoRequiredFileds"), "服务对象*")); //Content_CN：“{0}”列不能为空！
            //    wnGridEvaluate.SetGridFocus(wnGridEvaluate.GetRowIndex(drResultArray[0]), 0);
            //    return;
            //}

            ///* 检查界面是否有未填写量化成果的数据 */
            //drResultArray = wnGridEvaluate.SelectedRowsArray.AsEnumerable().Cast<DataRow>().Where(x => x["SelfFinishTime"] == DBNull.Value && x["MeasurableResult"] == DBNull.Value).ToArray();
            //if (drResultArray.Count() > 0)
            //{
            //    /* 弹出提示，定位到对应行 */
            //    DialogBox.ShowError(string.Format(this.GetCurrLanguageContent("WnFormMain.NoRequiredFileds"), "量化成果*")); //Content_CN：“{0}”列不能为空！
            //    wnGridEvaluate.SetGridFocus(wnGridEvaluate.GetRowIndex(drResultArray[0]), 0);
            //    return;
            //}

            /* 询问是否确定要进行完结，若不完结则取消打钩并返回 */
            if (DialogBox.AskYesNo(this.GetCurrLanguageContent("WnFormMain.AskHandle")) == DialogResult.No) { wnGridEvaluate.UnSelectAll(); return; } //Content_CN：确定完结选择的数据吗？

            /* 获取当前系统时间 */
            Object currTime = AppServer.GetDateTime();
            if (currTime == null) { return; }

            /* 创建集合存储工作任务、工作日期、自评完结时间 */
            List<string> listWorkTask = new List<string>();
            List<string> listWorkDate = new List<string>();
            List<string> listSelfFinishTime = new List<string>();

            /* 根据所选行，将自评完结时间为空的行的工作任务、工作日期、自评完结时间存入集合 */
            foreach (DataRow dr in wnGridEvaluate.SelectedRowsArray.AsEnumerable().Cast<DataRow>().Where(x => x["SelfFinishTime"] == DBNull.Value))
            {
                listWorkTask.Add(dr["WorkTaskGuid"].ToString());
                listWorkDate.Add(((DateTime)dr["WorkDate"]).ToString("yyyy/MM/dd"));
                listSelfFinishTime.Add(((DateTime)currTime).ToString("yyyy/MM/dd HH:mm:ss"));
            }

            /* 完结选中的数据 */
            Hashtable htHandle = new Hashtable();
            htHandle["DataSetName"] = "HandleData";
            htHandle["{WorkTaskGuid}"] = listWorkTask;
            htHandle["{WorkDate}"] = listWorkDate;
            htHandle["{SelfFinishTime}"] = listSelfFinishTime;
            if (this.DataSource.ExecSql(htHandle)["IsSuccess"].Equals("0"))
            {
                /* 弹出提示，取消勾选数据并返回 */
                DialogBox.ShowWarning(this.GetCurrLanguageContent("WnFormMain.HandleError")); //Content_CN：完结失败！
                wnGridEvaluate.UnSelectAll();
                return;
            }

            /* 完结成功后，循环完结的数据，修改“每日任务评价”的数据 */
            foreach (DataRow dr in wnGridEvaluate.SelectedRowsArray)
            {
                if (dr["SelfFinishTime"] == DBNull.Value)
                {
                    dr["SelfFinishTime"] = currTime;
                    dr.AcceptChanges();
                }
            }

            /* 完结成功后，根据存储的工作任务内码，更新每日任务汇报界面的完结标记 */
            foreach (DataRow dr in wnGridReport.DataSource.Select("SelfFinishTime is null"))
            {
                for (int i = 0; i < listWorkTask.Count; i++)
                {
                    if (dr["WorkTaskGuid"].ToString() == listWorkTask[i])
                    {
                        dr["SelfFinishTime"] = currTime;
                        dr.AcceptChanges();
                    }
                }
            }

            /* 完结成功后，根据存储的工作任务内码，更新工作任务界面的完结标记 */
            foreach (DataRow dr in wnGridTask.DataSource.Select("SelfFinishTime is null"))
            {
                for (int i = 0; i < listWorkTask.Count; i++)
                {
                    if (dr["Guid"].ToString() == listWorkTask[i])
                    {
                        dr["SelfFinishTime"] = currTime;
                        dr.AcceptChanges();
                    }
                }
            }

            /* 取消打钩 */
            wnGridEvaluate.UnSelectAll();

            /* 设置已完结的行不可编辑 */
            SetControlRight();
        }

        /// <summary>
        /// 响应wnGridTask的ShownEditor事件
        /// </summary>
        private void wnGridTask_ShownEditor(object sender, EventArgs e)
        {
            /* 为GridView在可编辑状态下注册双击事件 */
            (sender as GridView).ActiveEditor.DoubleClick += wnGridTask_DoubleClick;
        }

        /// <summary>
        /// 响应wnGridTask的GridDoubleClick事件
        /// </summary>
        private void wnGridTask_DoubleClick(object sender, EventArgs e)
        {
            /* 若新增按钮不可用，则返回 */
            if (!tsButtonAppend.Enabled) { return; }

            /* 获取当前行，若为空则返回 */
            DataRow drCurr = wnGridTask.CurrRow;
            if (drCurr == null) { return; }

            /* 获取“每日任务评价”的完结数据，判断选择的工作任务是否已完结，若完结则不能选择 */
            DataRow[] drHandleArray = wnGridEvaluate.DataSource.Select("SelfFinishTime is not null and WorkTaskGuid = '" + drCurr["Guid"].ToString() + "'");
            if (drHandleArray.Count() > 0)
            {
                /* 弹出提示并返回 */
                DialogBox.ShowMessage(this.GetCurrLanguageContent("WnFormMain.CheckSelectWorkTask")); //Content_CN：当前工作任务已完结，无法选择！
                return;
            }

            /* 在“每日任务汇报”界面新增一行，并将当前行的工作任务赋值给新增行 */
            DataRow drAdd = wnGridReport.AddRow();
            drAdd["WorkTaskGuid"] = drCurr["Guid"];
        }

        /// <summary>
        /// 响应wnGridReport的CellValueChanged事件
        /// </summary>
        private void wnGridReport_CellValueChanged(object sender, CellValueChangedArgs e)
        {
            /* 获取当前行，若为空则返回 */
            DataRow drCurr = wnGridReport.CurrRow;
            if (drCurr == null) { return; }

            /* 如果修改的是开始时间列或结束时间列 */
            if (e.CurrColumnName == "StartTime" || e.CurrColumnName == "EndTime")
            {
                /* 如果当前行的开始时间或结束时间没填写，将工作时长赋值为0并返回 */
                if (drCurr["StartTime"] == DBNull.Value || drCurr["EndTime"] == DBNull.Value) { drCurr["WorkHours"] = 0; return; }

                /* 获取当前行的开始时间和结束时间 */
                TimeSpan startTime = ((DateTime)drCurr["StartTime"]).TimeOfDay;
                TimeSpan endTime = ((DateTime)drCurr["EndTime"]).TimeOfDay;

                /* 如果结束时间小于等于开始时间，则将工作时长赋值为0并返回 */
                if (startTime >= endTime) { drCurr["WorkHours"] = 0; return; }

                /* 计算工作时长，并将sum列的结果显示到界面的汇总列 */
                drCurr["WorkHours"] = (endTime - startTime).TotalHours;
                wnGridReport.CommitEdit();
            }

            /* 如果修改的是工作任务列 */
            if (e.CurrColumnName == "WorkTaskGuid")
            {
                /* 获取“每日任务评价”的完结数据，判断选择的工作任务是否已完结，若完结则不能选择 */
                DataRow[] drHandleArray = wnGridEvaluate.DataSource.Select("SelfFinishTime is not null and WorkTaskGuid = '" + e.NewData.ToString() + "'");
                if (drHandleArray.Count() > 0)
                {
                    /* 弹出提示，清空当前行工作任务的值并返回 */
                    DialogBox.ShowMessage(this.GetCurrLanguageContent("WnFormMain.CheckSelectWorkTask")); //Content_CN：当前工作任务已完结，无法选择！
                    drCurr["WorkTaskGuid"] = DBNull.Value;
                    return;
                }

                /* 获取“每日任务评价”中对应的行，并定位到该行 */
                if (wnGridEvaluate.DataSource == null) { return; }
                DataRow[] drIndexArray = wnGridEvaluate.DataSource.Select("WorkTaskGuid = '" + drCurr["WorkTaskGuid"] + "'");
                if (drIndexArray.Count() > 0) { wnGridEvaluate.SetGridFocus(wnGridEvaluate.GetRowIndex(drIndexArray[0]), 0); }
            }
        }

        /// <summary>
        /// 响应wnGridReport的CurrRowChanged事件
        /// </summary>
        private void wnGridReport_CurrRowChanged(object sender, CurrRowChangedArgs e)
        {
            /* 设置当前行是否可编辑，若已完结则不可编辑，反之则可编辑 */
            SetControlRight();

            /* 获取当前行，若为空则返回 */
            DataRow drCurr = wnGridReport.CurrRow;
            if (drCurr == null) { return; }

            /* 获取“每日任务评价”中对应的行，并定位到该行 */
            if (wnGridEvaluate.DataSource == null) { return; }
            DataRow[] drIndexArray = wnGridEvaluate.DataSource.Select("WorkTaskGuid = '" + drCurr["WorkTaskGuid"] + "'");
            if (drIndexArray.Count() > 0) { wnGridEvaluate.SetGridFocus(wnGridEvaluate.GetRowIndex(drIndexArray[0]), 0); }
        }

        /// <summary>
        /// 响应wnGridEvaluate的CurrRowChanged事件
        /// </summary>
        private void wnGridEvaluate_CurrRowChanged(object sender, CurrRowChangedArgs e)
        {
            /* 设置当前行是否可编辑，若已完结则不可编辑，反之则可编辑 */
            SetControlRight();
        }

        /// <summary>
        /// 响应_gridLookupCode的QueryPopUp事件
        /// </summary>
        private void _gridLookupCode_QueryPopUp(object sender, CancelEventArgs e)
        {
            /* 设置wnGridReport工作任务列的Grid查询数据源为展开时的数据源 */
            SetGridLookupData("QueryPopUp");
            (sender as GridLookUpEdit).Properties.DataSource = _gridLookupCode.DataSource as DataTable;
        }

        /// <summary>
        /// 响应_gridLookupCode的Closed事件
        /// </summary>
        private void _gridLookupCode_Closed(object sender, ClosedEventArgs e)
        {
            /* 设置wnGridReport工作任务列的Grid查询数据源为关闭时的数据源 */
            SetGridLookupData("Closed");
            (sender as GridLookUpEdit).Properties.DataSource = _gridLookupCode.DataSource as DataTable;
        }

        /// <summary>
        /// 响应wnTextResult的DoubleClick事件
        /// </summary>
        private void wnTextResult_DoubleClick(object sender, EventArgs e)
        {
            /* 将Group移动到窗体顶部 */
            Point originalLocation = wnGroupEvaluate.Location;
            int groupX = (this.Width - wnGroupEvaluate.Width) / 2;
            wnGroupEvaluate.Location = new Point(groupX, 0);

            /* 判断量化成果文本框有没有值 */
            bool isEmpty = wnTextResult.Value == null || string.IsNullOrEmpty(wnTextResult.Value.ToString());
            if (isEmpty)
            {
                /* 打开量化成果模板窗体 */
                WnFormTemplet wnFormTemplet = new WnFormTemplet();
                if (!wnFormTemplet.Assemble(this, "WnFormTemplet")) { return; }
                wnFormTemplet.ShowDialog();

                /* 设置group的位置返回到原始位置 */
                wnGroupEvaluate.Location = originalLocation;

                /* 为量化成果文本框赋值 */
                wnTextResult.Value = wnFormTemplet._result;
                wnTextResult.SyncData();
            }
            else
            {
                /* 打开量化成果模板窗体 */
                WnFormTemplet wnFormTemplet = new WnFormTemplet(wnTextResult.Value.ToString());
                if (!wnFormTemplet.Assemble(this, "WnFormTemplet")) { return; }
                wnFormTemplet.ShowDialog();

                /* 设置group的位置返回到原始位置 */
                wnGroupEvaluate.Location = originalLocation;

                /* 为量化成果文本框赋值 */
                if (!string.IsNullOrEmpty(wnFormTemplet._result))
                {
                    AppendTempletContent(wnFormTemplet._result);
                }
            }
        }

        #endregion

        #region 内部方法

        /// <summary>
        /// 获取数据
        /// </summary>
        private bool GetData()
        {
            /* 准备参数，获取数据 */
            Hashtable ht = new Hashtable();
            ht["DataSetName"] = "GetData";
            ht["{Year}"] = tsComboYear.Value + "";
            ht["{WorkDate}"] = tsDateWorkDate.Value + "";
            ht["{Executor}"] = this._employeeName;

            /* 如果获取数据成功，则相关按钮可用，否则相关按钮不可用 */
            if (this.DataSource.GetDataSet(ht)["IsSuccess"].Equals("0"))
            {
                /* 获取数据失败，则按钮不可用 */
                _isSuccess = false;
                SetControlRight();
                return false;
            }

            /* 获取当前员工当前工作日期的工作日历，如果没有获取到工作日历，提示用户后返回 */
            Hashtable htCalendar = new Hashtable() { { "DataSetName", "GetStaffCalendar" } };
            htCalendar["{YgGuid}"] = _employeeGuid;
            htCalendar["{Gzrq}"] = tsDateWorkDate.Value.ToString();
            if (this.DataSource.GetDataSet(htCalendar)["IsSuccess"].Equals("0")) { _isSuccess = false; SetControlRight(); return false; }
            _dtCalendar = (this.DataSource.DataSets["GetStaffCalendar"] as DataSet).Tables["T01"];
            if (_dtCalendar.Rows.Count <= 0)
            {
                DialogBox.ShowWarning(string.Format(this.GetCurrLanguageContent("WnFormMain.NoStaffCalendar"), ((DateTime)tsDateWorkDate.Value).ToString("yyyy/MM/dd"))); //Content_CN：未获取到获取当前用户“{0}”的工作日历，请联系系统管理员！
                _isSuccess = false;
                SetControlRight();
                return false;
            }

            /* 取数成功时，记录取数状态、年度、工作日期作为控制按钮是否可用的条件 */
            _isSuccess = true;
            _year = tsComboYear.Value + "";
            _workDate = tsDateWorkDate.Value + "";

            /* 控制按钮权限 */
            SetControlRight();

            /* 设置时间列的格式，目的是在编辑工作时间的时候禁止选择日期 */
            RepositoryItemTimeEdit repositoryItemTimeEdit = new RepositoryItemTimeEdit();
            repositoryItemTimeEdit.EditMask = "HH:mm";
            this.wnGridReport.GridView.Columns["StartTime"].ColumnEdit = this.wnGridReport.GridView.Columns["EndTime"].ColumnEdit = repositoryItemTimeEdit;

            /* 绑定“每日任务汇报”的工作任务列的Grid查询数据源 */
            SetGridLookupData("Closed");

            /* 绑定“每日任务评价”的工作任务列的Grid查询数据源 */
            wnGridEvaluate.Columns["WorkTaskGuid"].KeyValues["{Condition}"] = "Executor = '" + this._employeeName + "'";
            wnGridEvaluate.GridLookupBindData("WorkTaskGuid");

            /* 绑定“加班申请编号”的Grid查询数据源 */
            wnGridReport.Columns["OTApplyNO"].KeyValues["{WorkDate}"] = tsDateWorkDate.Value.ToString();
            wnGridReport.Columns["OTApplyNO"].KeyValues["{OTUser}"] = _employeeName;
            wnGridReport.GridLookupBindData("OTApplyNO");

            /* 将“每日任务汇报”中工作任务列的编辑器赋值给_gridLookupCode，同时注册展开/关闭查询时的事件 */
            _gridLookupCode = wnGridReport.Columns["WorkTaskGuid"].RealColumnEdit as RepositoryItemGridLookUpEdit;
            _gridLookupCode.QueryPopUp += _gridLookupCode_QueryPopUp;
            _gridLookupCode.Closed += _gridLookupCode_Closed;

            return true;
        }

        /// <summary>
        /// 保存数据
        /// </summary>
        private bool SaveData(List<Act> actionStep)
        {
            /* 提交数据，检查是否有需要保存的数据 */
            Hashtable htSave = new Hashtable() { { "DataSetName", "GetData" } };
            htSave["DataSetName"] = "GetData";
            htSave["TableNames"] = new string[] { "T01", "T03" };
            htSave["ActionStep"] = new List<Act>() { Act.CommitEdit };
            if (this.DataSource.UpdateDataSet(htSave)["IsChanged"].Equals("0")) { return true; }

            /* 询问是否保存 */
            if (actionStep.Contains(Act.AskIsUpdate))
            {
                htSave["ActionStep"] = new List<Act> { Act.AskIsUpdate };
                Hashtable result = this.DataSource.UpdateDataSet(htSave);
                if (result["IsUpdate"].Equals("-1")) { return false; } //取消
                if (result["IsUpdate"].Equals("0")) //否
                {
                    /* 回滚数据 */
                    wnGridReport.RejectChanges();
                    wnGridEvaluate.RejectChanges();
                    return true;
                }
                actionStep.Remove(Act.AskIsUpdate);
            }

            /* 检查界面必填项不能为空 */
            htSave["ActionStep"] = new List<Act>() { Act.CheckData };
            if (this.DataSource.UpdateDataSet(htSave)["IsEmpty"].Equals("1")) { return false; }

            /* 若在“每日任务汇报”存在修改的数据，则需要检查汇报的时间段，并重新汇总“每日任务汇报”的数据到“每日任务评价”中 */
            if (wnGridReport.DataSource.Select("RowState = '10' or RowState = '11'").Count() > 0)
            {
                #region 检查开始时间和结束时间是否存在重复时间段

                /* 检查工作时数不能为0 */
                if (wnGridReport.DataSource.Select("WorkHours = 0").Count() > 0)
                {
                    DialogBox.ShowWarning(this.GetCurrLanguageContent("WnFormMain.CheckWorkHours")); //Content_CN：工作时数不能为0，请修改后保存！
                    return false;
                }

                /* 按照开始时间将界面数据排序，并存入到数组中 */
                List<DataRow> listDate = wnGridReport.DataSource.AsEnumerable().Cast<DataRow>().OrderBy(x => Convert.ToDateTime(x["StartTime"]).TimeOfDay).ToList();

                /* 如果只有1条记录，则不需要检查重复时间段 */
                if (listDate.Count() > 1)
                {
                    /* 定义一个集合，用于存储重复的时间段 */
                    List<string> listTime = new List<string>();

                    /* 循环判断是否存在重复的时间段 */
                    for (int i = 0; i < listDate.Count() - 1; i++)
                    {
                        /* 将重复的时间段存入集合，用于提示 */
                        if (((DateTime)listDate[i]["EndTime"]).TimeOfDay > ((DateTime)listDate[i + 1]["StartTime"]).TimeOfDay)
                        {
                            /* 将该时间段添加进集合 */
                            string timeSpan = ((DateTime)listDate[i]["StartTime"]).ToString("HH:mm") + "-" +
                                              ((DateTime)listDate[i]["EndTime"]).ToString("HH:mm") + "与" +
                                              ((DateTime)listDate[i + 1]["StartTime"]).ToString("HH:mm") + "-" +
                                              ((DateTime)listDate[i + 1]["EndTime"]).ToString("HH:mm");
                            listTime.Add(timeSpan);
                        }
                    }

                    /* 如果有重复时间段 */
                    if (listTime.Count > 0)
                    {
                        /* 获取数据重复的时间段，提示后返回 */
                        string timeSpans = string.Join(",", listTime);
                        DialogBox.ShowWarning(string.Format(this.GetCurrLanguageContent("WnFormMain.CheckRepeat"), timeSpans)); //Content_CN：工作时间段“{0}”存在重复！请修改后保存！
                        return false;
                    }
                }

                #endregion

                #region 检查汇报时间是否符合要求

                /* 获取员工工作日历 */
                string sfgzr = _dtCalendar.Rows[0]["Sfgzr"].ToString();
                DateTime zcsbsj = Convert.ToDateTime(_dtCalendar.Rows[0]["Zcsbsj"].ToString());
                DateTime zcxbsj = Convert.ToDateTime(_dtCalendar.Rows[0]["Zcxbsj"].ToString());
                DateTime xwsbsj = Convert.ToDateTime(_dtCalendar.Rows[0]["Xwsbsj"].ToString());
                DateTime xwxbsj = Convert.ToDateTime(_dtCalendar.Rows[0]["Xwxbsj"].ToString());

                #region 工作日与休息日单独检查

                string otApplyNO = "";
                string noonOTApplyNO = "";
                if (sfgzr == "1")
                {
                    /* 循环判断是否有跨时间段汇报 */
                    for (int i = 0; i < listDate.Count(); i++)
                    {
                        if ((((DateTime)listDate[i]["StartTime"]).TimeOfDay < zcsbsj.TimeOfDay && ((DateTime)listDate[i]["EndTime"]).TimeOfDay > zcsbsj.TimeOfDay) || (((DateTime)listDate[i]["StartTime"]).TimeOfDay < zcxbsj.TimeOfDay && ((DateTime)listDate[i]["EndTime"]).TimeOfDay > zcxbsj.TimeOfDay) || (xwsbsj != xwxbsj && (((DateTime)listDate[i]["StartTime"]).TimeOfDay < xwsbsj.TimeOfDay && ((DateTime)listDate[i]["EndTime"]).TimeOfDay > xwsbsj.TimeOfDay || ((DateTime)listDate[i]["StartTime"]).TimeOfDay < xwxbsj.TimeOfDay && ((DateTime)listDate[i]["EndTime"]).TimeOfDay > xwxbsj.TimeOfDay)) || (xwsbsj == xwxbsj && ((DateTime)listDate[i]["StartTime"]).TimeOfDay < zcxbsj.TimeOfDay && ((DateTime)listDate[i]["EndTime"]).TimeOfDay > zcxbsj.TimeOfDay))
                        {
                            //DialogBox.ShowWarning(this.GetCurrLanguageContent("WnFormMain.IsAcrossTime")); //Content_CN：加班工作需单独汇报，请修改后保存！
                            //return false;
                        }
                    }

                    /* 定义集合，用于存储不同时间段 */
                    List<DataRow> listOTDate = listDate.Where(x => (Convert.ToDateTime(x["StartTime"]).TimeOfDay < zcsbsj.TimeOfDay && Convert.ToDateTime(x["EndTime"]).TimeOfDay <= zcsbsj.TimeOfDay) || (xwsbsj != xwxbsj && Convert.ToDateTime(x["StartTime"]).TimeOfDay >= xwxbsj.TimeOfDay) || (xwsbsj == xwxbsj && Convert.ToDateTime(x["StartTime"]).TimeOfDay >= zcxbsj.TimeOfDay)).ToList();
                    List<DataRow> listZwOTDate = listDate.Where(x => (Convert.ToDateTime(x["StartTime"]).TimeOfDay >= zcxbsj.TimeOfDay && Convert.ToDateTime(x["EndTime"]).TimeOfDay <= xwsbsj.TimeOfDay)).ToList();
                    double jbsj = 0.0;
                    string no = "";

                    /* 检查上午、下午加班汇报情况 */
                    if (listOTDate.Count() > 0)
                    {
                        for (int i = 0; i < listOTDate.Count(); i++)
                        {
                            if (String.IsNullOrWhiteSpace(listOTDate[i]["OTApplyNO"].ToString()))
                            {
                                //DialogBox.ShowWarning(string.Format(this.GetCurrLanguageContent("WnFormMain.NoOTApplyNO"), ((DateTime)listOTDate[i]["StartTime"]).ToString("HH:mm"), ((DateTime)listOTDate[i]["EndTime"]).ToString("HH:mm"))); //Content_CN：{0}-{1}工作需填写加班申请编号，请修改后保存！
                                //return false;
                            }
                            if (listOTDate.Count() == 1 && ((DateTime)listOTDate[0]["EndTime"] - (DateTime)listOTDate[0]["StartTime"]).TotalMinutes < 30)
                            {
                                otApplyNO += "'" + listOTDate[0]["OTApplyNO"].ToString() + "',";
                            }
                            if (listOTDate.Count() != 1)
                            {
                                if (i < listOTDate.Count() - 1)
                                {
                                    if (((DateTime)listOTDate[i]["EndTime"]).TimeOfDay == ((DateTime)listOTDate[i + 1]["StartTime"]).TimeOfDay)
                                    {
                                        jbsj += ((DateTime)listOTDate[i + 1]["EndTime"] - (DateTime)listOTDate[i]["StartTime"]).TotalMinutes;
                                    }
                                    else
                                    {
                                        jbsj += ((DateTime)listOTDate[i]["EndTime"] - (DateTime)listOTDate[i]["StartTime"]).TotalMinutes;
                                        if (jbsj < 30) { no += "'" + listOTDate[i]["OTApplyNO"].ToString() + "',"; }
                                        else { no = ""; jbsj = 0; }
                                    }
                                }
                                else
                                {
                                    if (((DateTime)listOTDate[i]["EndTime"] - (DateTime)listOTDate[i]["StartTime"]).TotalMinutes < 30 && ((DateTime)listOTDate[i]["StartTime"]).TimeOfDay != ((DateTime)listOTDate[i - 1]["EndTime"]).TimeOfDay)
                                    {
                                        otApplyNO += "'" + listOTDate[i]["OTApplyNO"].ToString() + "',";
                                    }
                                }
                            }
                        }
                        otApplyNO += no;
                    }

                    /* 检查中午加班汇报情况 */
                    if (listZwOTDate.Count() > 0)
                    {
                        for (int i = 0; i < listZwOTDate.Count(); i++)
                        {
                            if (String.IsNullOrWhiteSpace(listZwOTDate[i]["OTApplyNO"].ToString()))
                            {
                                //DialogBox.ShowWarning(string.Format(this.GetCurrLanguageContent("WnFormMain.NoOTApplyNO"), ((DateTime)listZwOTDate[i]["StartTime"]).ToString("HH:mm"), ((DateTime)listZwOTDate[i]["EndTime"]).ToString("HH:mm"))); //Content_CN：{0}-{1}工作需填写加班申请编号，请修改后保存！
                                //return false;
                            }
                            if (xwxbsj - xwsbsj != TimeSpan.Zero)
                            {
                                noonOTApplyNO += "'" + listZwOTDate[i]["OTApplyNO"].ToString() + "',";
                            }
                            else
                            {
                                if (listZwOTDate.Count() == 1 && ((DateTime)listZwOTDate[0]["EndTime"] - (DateTime)listZwOTDate[0]["StartTime"]).TotalMinutes < 30)
                                {
                                    otApplyNO += "'" + listZwOTDate[0]["OTApplyNO"].ToString() + "',";
                                }
                                if (listZwOTDate.Count() != 1)
                                {
                                    if (i < listZwOTDate.Count() - 1)
                                    {
                                        if (((DateTime)listZwOTDate[i]["EndTime"]).TimeOfDay == ((DateTime)listZwOTDate[i + 1]["StartTime"]).TimeOfDay)
                                        {
                                            jbsj += ((DateTime)listZwOTDate[i + 1]["EndTime"] - (DateTime)listZwOTDate[i]["StartTime"]).TotalMinutes;
                                        }
                                        else
                                        {
                                            jbsj += ((DateTime)listOTDate[i]["EndTime"] - (DateTime)listOTDate[i]["StartTime"]).TotalMinutes;
                                            if (jbsj < 30) { no += "'" + listZwOTDate[i]["OTApplyNO"].ToString() + "',"; }
                                            else { no = ""; jbsj = 0; }
                                        }
                                    }
                                    else
                                    {
                                        if (((DateTime)listZwOTDate[i]["EndTime"] - (DateTime)listZwOTDate[i]["StartTime"]).TotalMinutes < 30 && ((DateTime)listZwOTDate[i]["StartTime"]).TimeOfDay != ((DateTime)listZwOTDate[i - 1]["EndTime"]).TimeOfDay)
                                        {
                                            otApplyNO += "'" + listZwOTDate[i]["OTApplyNO"].ToString() + "',";
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < listDate.Count(); i++)
                    {
                        if (String.IsNullOrWhiteSpace(listDate[i]["OTApplyNO"].ToString()))
                        {
                            //DialogBox.ShowWarning(string.Format(this.GetCurrLanguageContent("WnFormMain.NoOTApplyNO"), ((DateTime)listDate[i]["StartTime"]).ToString("HH:mm"), ((DateTime)listDate[i]["EndTime"]).ToString("HH:mm"))); //Content_CN：{0}-{1}工作需填写加班申请编号，请修改后保存！
                            //return false;
                        }
                    }
                }

                #endregion

                #region 特殊汇报情况，需检查该汇报是否符合加班申请要求

                DataRow[] drs = wnGridReport.DataSource.Select("OTApplyNO is not null");
                string otApplyNOs = "";
                if (drs.Count() > 0)
                {
                    Hashtable htOTApply = new Hashtable() { { "DataSetName", "GetOTApply" } };
                    htOTApply["{OTUser}"] = _employeeName;
                    htOTApply["{WorkDate}"] = tsDateWorkDate.Value.ToString();
                    if (this.DataSource.GetDataSet(htOTApply)["IsSuccess"].Equals("0")) { return false; }
                    DataTable dtOTApply = (this.DataSource.DataSets["GetOTApply"] as DataSet).Tables["T01"];
                    foreach (DataRow dr in drs)
                    {
                        otApplyNOs += "'" + dr["OTApplyNO"] + "',";
                    }

                    /* 如果界面上总工作时数大于11小时，检查加班申请记录中是否允许超过11小时 */
                    if (Convert.ToDecimal(wnGridReport.Columns["WorkHours"].GetSummaryValue()) > 11)
                    {
                        DataRow[] drOverEleven = dtOTApply.Select("IsAllowOverEleven = 1 and ApplyNO in (" + otApplyNOs.TrimEnd(',') + ")");
                        if (drOverEleven.Count() == 0)
                        {
                            //DialogBox.ShowWarning(this.GetCurrLanguageContent("WnFormMain.IsNotAllowOverEleven")); //Content_CN：工作时数不允许超过11小时，请修改后保存！
                            //return false;
                        }
                    }
                    if (otApplyNO != "")
                    {
                        if (dtOTApply.Select("IsAllowLessHalfHour = 1 and ApplyNO in (" + otApplyNO.TrimEnd(',') + ")").Count() == 0)
                        {
                            DialogBox.ShowWarning(this.GetCurrLanguageContent("WnFormMain.IsNotAllowLessHalfHour")); //Content_CN：工作时数不足半小时不允许汇报，请修改后保存！
                            return false;
                        }
                    }
                    if (noonOTApplyNO != "")
                    {
                        if (dtOTApply.Select("IsAllowNoonOT = 1 and ApplyNO in (" + noonOTApplyNO.TrimEnd(',') + ")").Count() == 0)
                        {
                            DialogBox.ShowWarning(this.GetCurrLanguageContent("WnFormMain.IsNotAllowNoonOT")); //Content_CN：中午不允许汇报工作，请修改后保存！
                            return false;
                        }
                    }
                }

                #endregion

                #endregion

                #region 汇总处理“每日任务评价”界面未完结的数据

                /* 将岗位信息和部门运行效率存储到DataRow中 */
                DataRow drPost = (this.DataSource.DataSets["GetPost"] as DataSet).Tables["T01"].Rows[0];
                DataRow drDept = (this.DataSource.DataSets["GetDeptEfficiency"] as DataSet).Tables["T01"].Rows[0];

                /* 获取“每日任务汇报”上的新增修改行和修改行的任务内码作为linq的where条件，在汇总处理数据时，只针对与修改行任务内码相同的数据 */
                List<string> listWhereTask = new List<string>();
                foreach (DataRow drWhere in wnGridReport.DataSource.Select("RowState = '10' or RowState = '11'"))
                {
                    /* 如果为新增行，则存入当前值 */
                    if (drWhere.RowState == DataRowState.Added && !listWhereTask.Contains(drWhere["WorkTaskGuid"].ToString()))
                    {
                        listWhereTask.Add(drWhere["WorkTaskGuid"].ToString());
                    }

                    /* 如果为修改行，则存入当前值和原始值 */
                    if (drWhere.RowState == DataRowState.Modified && !listWhereTask.Contains(drWhere["WorkTaskGuid"].ToString()))
                    {
                        listWhereTask.Add(drWhere["WorkTaskGuid"].ToString());
                    }
                    if (drWhere.RowState == DataRowState.Modified && drWhere["WorkTaskGuid", DataRowVersion.Original].ToString() != drWhere["WorkTaskGuid"].ToString() && !listWhereTask.Contains(drWhere["WorkTaskGuid", DataRowVersion.Original].ToString()))
                    {
                        listWhereTask.Add(drWhere["WorkTaskGuid", DataRowVersion.Original].ToString());
                    }
                }

                /* 汇总“每日任务汇报”中发生过修改的数据 */
                var reportGrid = from p in wnGridReport.DataSource.AsEnumerable().Cast<DataRow>().Where(x => listWhereTask.Contains(x["WorkTaskGuid"].ToString()))
                                 group p by new { workTaskGuid = p.Field<string>("WorkTaskGuid") } into g
                                 select new
                                 {
                                     workTaskGuid = g.Key.workTaskGuid,
                                     workHours = g.Sum(m => m.Field<decimal>("WorkHours"))
                                 };

                /* 获取“每日任务评价”发生过修改的数据，并将工作任务内码存储到list中 */
                List<string> listWorkTask = wnGridEvaluate.DataSource.AsEnumerable().Cast<DataRow>().Where(x => listWhereTask.Contains(x["WorkTaskGuid"].ToString())).Select(x => x["WorkTaskGuid"].ToString()).ToList();

                /* 检查“每日任务汇报”中的工作任务是否存在于“每日任务评价”中 */
                foreach (var item in reportGrid)
                {
                    /* 定义变量，存储工作内容的汇总值 */
                    string workContent = string.Join("；", wnGridReport.DataSource.AsEnumerable().Cast<DataRow>().Where(x => x["WorkTaskGuid"].ToString() == item.workTaskGuid.ToString()).Select(x => x["WorkContent"].ToString()).Distinct()) + "。";

                    /* 如果在“每日任务评价”中存在，则需要更新其工作时长列，反之则需要新增 */
                    if (listWorkTask.Contains(item.workTaskGuid.ToString()))
                    {
                        /* 定位到相对应的行，用于更新工作时长并替换工作内容 */
                        DataRow drUpdate = wnGridEvaluate.DataSource.Select("WorkTaskGuid = '" + item.workTaskGuid.ToString() + "'")[0];

                        /* 分割工作成效评价，用来更新工作成效评价中的工作内容 */
                        string workEffect = drUpdate["WorkEffectEvaluate"].ToString();
                        string oldWorkEffect = workEffect.IndexOf("\n") < 0 ? "\n【进度调整】：\n【协同配合】：\n【个性化程度】：\n【复杂程度】：\n【知识共享】：" : workEffect.Substring(workEffect.IndexOf("\n"), workEffect.Length - workEffect.IndexOf("\n"));

                        /* 更新相应的列 */
                        drUpdate["WorkEffectEvaluate"] = "【工作内容】：" + workContent + oldWorkEffect; //工作成效评价
                        drUpdate["WorkHours"] = item.workHours.ToString(); //工作时数
                        drUpdate["DeleteFlag"] = "0"; //删除标记
                        drUpdate["RowState"] = "11";
                    }
                    else
                    {
                        /* 新增一行，并赋值默认值 */
                        DataRow drAdd = wnGridEvaluate.AddRow();
                        drAdd["WorkThemeName"] = wnGridTask.DataSource.Select("Guid = '" + item.workTaskGuid.ToString() + "'")[0]["WorkThemeName"]; //工作主题
                        drAdd["WorkTaskGuid"] = item.workTaskGuid.ToString(); //工作任务
                        drAdd["WorkHours"] = item.workHours.ToString(); //工作时数
                        drAdd["WorkEffectEvaluate"] = "【工作内容】：" + workContent + "\n【进度调整】：\n【协同配合】：\n【个性化程度】：\n【复杂程度】：\n【知识共享】："; //工作成效评价
                        drAdd["PostName"] = drPost["PostName"].ToString(); //岗位名称
                        drAdd["PostGuid"] = drPost["PostGuid"].ToString(); //岗位内码
                        drAdd["PostRatio"] = drPost["PostRatio"].ToString(); //岗位系数
                        drAdd["AbilityLevel"] = drPost["AbilityLevel"].ToString(); //能力等级
                        drAdd["DeptEfficiency"] = drDept["ParamValue"].ToString(); //部门运行效率
                        drAdd["DeleteFlag"] = "0"; //删除标记
                        drAdd["RowState"] = "10";
                    }
                }

                /* 将删除标记为1的列删除，只删除界面数据，利用UpdateDataSet方法实现删除 */
                foreach (DataRow drDel in wnGridEvaluate.DataSource.Select("DeleteFlag = '1' and WorkTaskGuid in ('" + string.Join("','", listWhereTask) + "')"))
                {
                    wnGridEvaluate.DeleteRow(drDel);
                }

                /* 重置“每日任务评价”中的删除标记列 */
                foreach (DataRow drReset in wnGridEvaluate.DataSource.Select("WorkTaskGuid in ('" + string.Join("','", listWhereTask) + "')"))
                {
                    if (drReset.RowState == DataRowState.Deleted) { continue; }
                    drReset["DeleteFlag"] = "1"; //删除标记
                }

                #endregion
            }

            /* 若在“每日绩效自评”存在修改的数据，则需检查是否存在“服务对象”为“公共”，“服务对象备注”为空的数据 */
            //DataRow[] drEvaluates = wnGridEvaluate.DataSource.Select("RowState = '10' or RowState = '11'");
            //if (drEvaluates.Count() > 0)
            //{
            //    foreach (DataRow dr in drEvaluates)
            //    {
            //        if (dr["ServiceObject"].ToString() == "GG" && string.IsNullOrWhiteSpace(dr["ServiceObjectNote"].ToString()))
            //        {
            //            DialogBox.ShowError(this.GetCurrLanguageContent("WnFormMain.NoServiceObjectNote")); //Content_CN：“服务对象”为“公共”，“服务对象备注”列不能为空！
            //            wnGridEvaluate.SetGridFocus(wnGridEvaluate.GetRowIndex(dr), 0);
            //            return false;
            //        }
            //    }
            //}

            /* 获取时间 */
            object currTime = AppServer.GetDateTime();
            if (currTime == null) { return false; }

            /* 为汇报时间、工作日期赋值，同时修改开始时间和结束时间的日期部分 */
            foreach (DataRow dr in wnGridReport.DataSource.Select("RowState = '10' or RowState = '11'"))
            {
                dr["ReportTime"] = ((DateTime)currTime).ToString("yyyy/MM/dd HH:mm:ss");

                /* 如果服务器时间与当前选择的日期不同，则以当前日期为准更新开始时间和结束时间的日期部分 */
                if (((DateTime)currTime).Date != (DateTime)dr["WorkDate"])
                {
                    /* 为开始时间和结束时间重新赋值 */
                    dr["StartTime"] = ((DateTime)dr["WorkDate"]).Add(((DateTime)dr["StartTime"]).TimeOfDay);
                    dr["EndTime"] = ((DateTime)dr["WorkDate"]).Add(((DateTime)dr["EndTime"]).TimeOfDay);
                }
            }

            /* 获取“每日任务评价”的修改行，检查工作进度列和工作成效评价列的值 */
            foreach (DataRow dr in wnGridEvaluate.DataSource.Select("RowState = '10' or RowState = '11'"))
            {
                /* 如果工作进度列的值为空或值小于0，则赋值为0；若值大于1，则赋值为1 */
                if (dr["WorkTaskProcess"] == DBNull.Value || Convert.ToDecimal(dr["WorkTaskProcess"]) < 0) { dr["WorkTaskProcess"] = 0; }
                if (Convert.ToDecimal(dr["WorkTaskProcess"]) > 1) { dr["WorkTaskProcess"] = 1; }

                /* 如果工作成效评价的值为空，则赋值默认格式 */
                if (dr["WorkEffectEvaluate"] == DBNull.Value) { dr["WorkEffectEvaluate"] = "【工作内容】：\n【进度调整】：\n【协同配合】：\n【个性化程度】：\n【复杂程度】：\n【知识共享】："; }
            }

            /* 重新设置通信 */
            wnGridEvaluate.SyncData();

            /* 保存数据 */
            htSave["ActionStep"] = actionStep;
            if (this.DataSource.UpdateDataSet(htSave)["IsSuccess"].Equals("0")) { return false; }

            /* 手动触发CurrRowChanged事件，定位到对应的行 */
            wnGridReport_CurrRowChanged(null, null);
            return true;
        }

        /// <summary>
        /// 设置按钮权限
        /// </summary>
        private void SetControlRight()
        {
            /* 根据取数是否成功、控件值与界面数据是否一致、用户拥有的权限、工作任务界面是否有工作任务来控制按钮可用性 */
            bool flag = _isSuccess && (tsComboYear.Value + "" == _year) && (tsDateWorkDate.Value + "" == _workDate) && !string.IsNullOrEmpty(tsDateWorkDate.Value + "") && this.RightInfo.Contains("EditRight") && (wnGridTask.DataSource != null) && wnGridTask.DataSource.Rows.Count > 0;
            tsButtonAppend.Enabled = tsButtonDelete.Enabled = tsButtonSave.Enabled = tsButtonMonthTaskEvaluate.Enabled = tsButtonHandle.Enabled = wnGridReport.IsAllowEdit = flag;

            /* 控制单值控件的权限 */
            SetSingleControlRight(flag);

            /* 获取“每日任务汇报”的当前行，若为不为空则根据是否完结设置权限 */
            if (wnGridReport.CurrRow != null)
            {
                wnGridReport.IsAllowEdit = flag && string.IsNullOrEmpty(wnGridReport.CurrRow["SelfFinishTime"].ToString());
            }

            /* 获取“每日任务评价”的当前行，若为不为空则根据是否完结设置权限 */
            if (wnGridEvaluate.CurrRow != null)
            {
                /* 设置当前行的单值控件是否可编辑，若已完结则不可编辑，反之则可编辑 */
                SetSingleControlRight(flag && string.IsNullOrEmpty(wnGridEvaluate.CurrRow["SelfFinishTime"].ToString()));
            }
        }

        /// <summary>
        /// 控制可编辑的单值控件是否可编辑
        /// </summary>
        /// <param name="editable">传入的是false，不可编辑，否则可编辑</param>
        private void SetSingleControlRight(bool editable)
        {
            foreach (ISingleControl singleControl in _listEditableControls)
            {
                singleControl.ControlReadOnly = !editable;
            }
        }

        /// <summary>
        /// 设置“每日任务汇报”工作任务列Grid查询的数据源
        /// </summary>
        private void SetGridLookupData(String action)
        {
            /* 定义变量用于存储关键字的值 */
            string condition = "";

            /* 展开下拉框，替换关键字 */
            if (action.Equals("QueryPopUp"))
            {
                condition = "a.Executor = '" + this._employeeName + "' and a.ExecuteState = '执行中' and b.Year = '" + tsComboYear.Value + "'";
            }

            /* 关闭下拉框，替换关键字 */
            if (action.Equals("Closed"))
            {
                condition = "a.Executor = '" + this._employeeName + "'";
            }

            /* 替换关键字并绑定数据源 */
            wnGridReport.Columns["WorkTaskGuid"].KeyValues["{Condition}"] = condition;
            wnGridReport.GridLookupBindData("WorkTaskGuid");
        }

        /// <summary>
        /// 将量化成果模板窗体返回的字符串追加到量化成果中
        /// </summary>
        private void AppendTempletContent(string result)
        {
            /* 把结果字符串的最后一个“。”替换为“；”，方便拼接 */
            result = result.Replace("。", "；");
            string[] newClassifyContent = result.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
            Hashtable htNewClassifies = new Hashtable();
            foreach (string s in newClassifyContent)
            {
                int index = s.IndexOf("：");
                string classify = s.Substring(0, index);
                string content = s.Substring(index + 1);
                htNewClassifies[classify] = content;
            }

            /* 把原始字符串的最后一个“。”替换为“；” */
            string originalStr = wnTextResult.Value.ToString();
            originalStr = originalStr.Replace("。", "；");

            /* 用分号分割原始字符串 */
            string[] oldClassifyContent = originalStr.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);

            /* 将新添加的模板追加到原始字符串分割后的对应模板分类中 */
            for (int i = 0; i < oldClassifyContent.Length; i++)
            {
                int index = oldClassifyContent[i].IndexOf("：");

                /* 当原始字符串格式不同时，直接拼接原始字符串和新的模板数据 */
                if (index == -1)
                {
                    string result0 = originalStr + "\n" + result;
                    result0 = result0.Substring(0, result0.Length - 1) + "。";
                    wnTextResult.Value = result0;
                    wnTextResult.SyncData();
                    return;
                }

                string oldClassify = oldClassifyContent[i].Substring(0, index);
                string oldcontent = oldClassifyContent[i].Substring(index + 1);

                if (htNewClassifies.Contains(oldClassify))
                {
                    oldClassifyContent[i] += htNewClassifies[oldClassify];
                    htNewClassifies.Remove(oldClassify);
                }
            }

            /* 拼接新的结果字符串 */
            string oldModified = string.Join("\n", oldClassifyContent);
            StringBuilder newTempletContetnt = new StringBuilder();
            foreach (string ss in htNewClassifies.Keys)
            {
                newTempletContetnt.Append("\n" + ss + "：" + htNewClassifies[ss]);
            }
            string newResult = oldModified + newTempletContetnt;
            newResult = newResult.Substring(0, newResult.Length - 1) + "。";

            /* 为文本框赋值，并同步数据 */
            wnTextResult.Value = newResult;
            wnTextResult.SyncData();
        }

        #endregion
    }
}