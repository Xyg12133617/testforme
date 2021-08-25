/*************************
* 管控对象：量化成果模板
* 管控成员：袁文秋
*************************/

using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraTreeList;
using DevExpress.XtraTreeList.Nodes;
using PEMSoft.Client.SYS.WnControl;
using PEMSoft.Client.SYS.WnPlatform;

namespace PEMSoft.Client.IFM.Jxgl.WnDayPerformanceSelfEvaluate
{
    public partial class WnFormTemplet : WnForm
    {
        #region 变量

        /// <summary>
        /// 存放处理后的量化成果字符串
        /// </summary>
        public string _result = "";

        /// <summary>
        /// 存放主窗体的原始字符串
        /// </summary>
        private string _originalResult = "";

        /// <summary>
        /// 存放选中量化成果模板的内容
        /// </summary>
        private string _templetContent = "";

        /// <summary>
        /// 标志关闭本窗体时是否询问用户
        /// </summary>
        private bool _askClose = true;

        #endregion

        #region 控件变量及初始化

        private WnToolBar wnToolBarTemplet;
        private WnTbButton tsButtonOK;
        private WnSeparator tsSeparator5;
        private WnTbButton tsButtonCancel;
        private WnSplitPanel wnSplitPanelTemplet;
        private WnSplitPanel wnSplitPanelChooseTemplet;
        private WnTreeGrid wnTreeGridMain;
        private WnPanel wnPanelChoose;
        private WnImage wnImageRight;
        private WnImage wnImageLeft;
        private WnImage wnImageAllRight;
        private WnImage wnImageAllLeft;
        private WnSplitPanel wnSplitPanelTempletContent;
        private WnText wnTextContent;
        private WnGroup wnGroupParam;
        private WnGrid wnGridParam;

        /// <summary>
        /// 控件变量初始化方法
        /// </summary>
        private void InitializeObject()
        {
            wnToolBarTemplet = this.AllObjects["wnToolBarTemplet"] as WnToolBar;
            tsButtonOK = this.AllObjects["tsButtonOK"] as WnTbButton;
            tsSeparator5 = this.AllObjects["tsSeparator5"] as WnSeparator;
            tsButtonCancel = this.AllObjects["tsButtonCancel"] as WnTbButton;
            wnSplitPanelTemplet = this.AllObjects["wnSplitPanelTemplet"] as WnSplitPanel;
            wnSplitPanelChooseTemplet = this.AllObjects["wnSplitPanelChooseTemplet"] as WnSplitPanel;
            wnTreeGridMain = this.AllObjects["wnTreeGridMain"] as WnTreeGrid;
            wnPanelChoose = this.AllObjects["wnPanelChoose"] as WnPanel;
            wnImageRight = this.AllObjects["wnImageRight"] as WnImage;
            wnImageLeft = this.AllObjects["wnImageLeft"] as WnImage;
            wnImageAllRight = this.AllObjects["wnImageAllRight"] as WnImage;
            wnImageAllLeft = this.AllObjects["wnImageAllLeft"] as WnImage;
            wnSplitPanelTempletContent = this.AllObjects["wnSplitPanelTempletContent"] as WnSplitPanel;
            wnTextContent = this.AllObjects["wnTextContent"] as WnText;
            wnGroupParam = this.AllObjects["wnGroupParam"] as WnGroup;
            wnGridParam = this.AllObjects["wnGridParam"] as WnGrid;
        }

        #endregion

        #region 窗体方法

        /// <summary>
        /// 构造函数
        /// </summary>
        public WnFormTemplet()
        {
            InitializeComponent();

            /* 注册窗体事件 */
            this.AfterAssemble += WnFormTemplet_AfterAssemble;
            this.Shown += WnFormTemplet_Shown;
            this.FormClosing += WnFormTemplet_FormClosing;
        }

        /// <summary>
        /// 带参构造函数
        /// </summary>
        public WnFormTemplet(string originalResult)
        {
            InitializeComponent();

            /* 注册窗体事件 */
            this.AfterAssemble += WnFormTemplet_AfterAssemble;
            this.Shown += WnFormTemplet_Shown;
            this.FormClosing += WnFormTemplet_FormClosing;

            /* 为全局变量模板参数赋值 */
            this._originalResult = originalResult;
        }

        /// <summary>
        /// 响应窗体的AfterAssemble事件
        /// </summary>
        private void WnFormTemplet_AfterAssemble(object sender, CancelEventArgs e)
        {
            /* 初始化控件 */
            InitializeObject();

            #region 设置图片数据源，注册图片事件

            /* 设置图片控件的初始图片 */
            wnImageRight.Value = Properties.Resources.GrayRight1_32x22;
            wnImageLeft.Value = Properties.Resources.GrayLeft1_32x22;
            wnImageAllRight.Value = Properties.Resources.GrayRight2_32x22;
            wnImageAllLeft.Value = Properties.Resources.GrayLeft2_32x22;

            /* 为图片控件绑定相关事件 */
            (wnImageRight.OriginControl as WnImageEditorBase).MouseEnter += wnImageRight_MouseEnter;
            (wnImageRight.OriginControl as WnImageEditorBase).MouseLeave += wnImageRight_MouseLeave;
            wnImageRight.MouseDown += wnImageRight_MouseDown;
            (wnImageLeft.OriginControl as WnImageEditorBase).MouseEnter += wnImageLeft_MouseEnter;
            (wnImageLeft.OriginControl as WnImageEditorBase).MouseLeave += wnImageLeft_MouseLeave;
            wnImageLeft.MouseDown += wnImageLeft_MouseDown;
            (wnImageAllRight.OriginControl as WnImageEditorBase).MouseEnter += wnImageAllRight_MouseEnter;
            (wnImageAllRight.OriginControl as WnImageEditorBase).MouseLeave += wnImageAllRight_MouseLeave;
            wnImageAllRight.MouseDown += wnImageAllRight_MouseDown;
            (wnImageAllLeft.OriginControl as WnImageEditorBase).MouseEnter += wnImageAllLeft_MouseEnter;
            (wnImageAllLeft.OriginControl as WnImageEditorBase).MouseLeave += wnImageAllLeft_MouseLeave;
            wnImageAllLeft.MouseDown += wnImageAllLeft_MouseDown;

            #endregion

            /* 注册控件事件 */
            tsButtonOK.Click += tsButtonOK_Click;
            tsButtonCancel.Click += tsButtonCancel_Click;
            wnTreeGridMain.AfterCheckNode += wnTreeGridMain_AfterCheckNode;
            wnGridParam.CellValueChanging += wnGridParam_CellValueChanging;

            /* 获取模板数据 */
            Hashtable htTemplet = new Hashtable() { { "DataSetName", "GetTemplet" } };
            if (this.DataSource.GetDataSet(htTemplet)["IsSuccess"].Equals("0")) { e.Cancel = true; return; }

            /* 如果没有模板，提示后返回 */
            DataTable dtTemplet = (this.DataSource.DataSets["GetTemplet"] as DataSet).Tables["T01"];
            if (dtTemplet.Rows.Count <= 0)
            {
                DialogBox.ShowError(this.GetCurrLanguageContent("WnFormTemplet.NoTemplet")); //Content_CN：未获取到量化成果模板，请联系系统管理员！
                e.Cancel = true;
                return;
            }

            /* 确认打开窗体后，设置窗体的位置 */
            this.StartPosition = FormStartPosition.Manual;
            Rectangle screenArea = System.Windows.Forms.Screen.GetWorkingArea(this);
            int screenWidth = screenArea.Width;
            int screenHeight = screenArea.Height;
            int x = (screenWidth - this.Width) / 2;
            int y = screenHeight - this.Height;
            this.Location = new Point(x, y);

            /* 展开默认分类模板，并移动至最前 */
            HandleDfTempClassify();

            /* 设置参数信息的数据结构 */
            DataTable dtColumn = new DataTable();
            dtColumn.Columns.Add("Guid", typeof(string));
            dtColumn.Columns.Add("TempletGuid", typeof(string));
            dtColumn.Columns.Add("Templet", typeof(string));
            dtColumn.Columns.Add("Param", typeof(string));
            dtColumn.Columns.Add("Value", typeof(string));
            dtColumn.Columns.Add("OrderIndex", typeof(int));
            dtColumn.Columns["Guid"].Caption = "参数内码";
            dtColumn.Columns["TempletGuid"].Caption = "模板内码";
            dtColumn.Columns["Templet"].Caption = "模板";
            dtColumn.Columns["Param"].Caption = "参数名称";
            dtColumn.Columns["Value"].Caption = "参数值";
            dtColumn.Columns["OrderIndex"].Caption = "模板显示顺序";
            wnGridParam.DataSource = dtColumn;

            /* 设置wnGridParam列的格式 */
            wnGridParam.Columns["Guid"].Visible = false;
            wnGridParam.Columns["Guid"].IsAllowEdit = false;
            wnGridParam.Columns["TempletGuid"].Visible = false;
            wnGridParam.Columns["TempletGuid"].IsAllowEdit = false;
            wnGridParam.Columns["Templet"].IsAllowEdit = false;
            wnGridParam.Columns["Param"].IsAllowEdit = false;
            wnGridParam.Columns["Value"].IsAllowEdit = true;
            wnGridParam.Columns["Value"].Width = 2 * wnGridParam.Columns["Param"].Width;
            wnGridParam.Columns["OrderIndex"].Visible = false;
            wnGridParam.Columns["OrderIndex"].IsAllowEdit = false;

            /* 设置wnGridParam允许多选 */
            wnGridParam.IsMultiRowSelect = true;
        }

        /// <summary>
        /// 响应窗体的Shown事件
        /// </summary>
        private void WnFormTemplet_Shown(object sender, EventArgs e)
        {
            /* 当WnFormMain的量化成果文本框有值时 */
            if (!string.IsNullOrEmpty(_originalResult))
            {
                /* 提示用户添加的模板数据，会追加到原来的字符串后面 */
                DialogBox.ShowMessage(this.GetCurrLanguageContent("WnFormTemplet.HintAppend")); //Content_CN：提示：处理后的模板数据会追加到原字符串末尾！
            }
        }

        /// <summary>
        /// 响应窗体的FormClosing事件
        /// </summary>
        private void WnFormTemplet_FormClosing(object sender, FormClosingEventArgs e)
        {
            /* 若不询问用户，直接关闭窗体 */
            if (!_askClose) { return; }

            /* 当“内容”文本框内有值、或者“参数信息”列表上有行时，询问用户是否确定取消 */
            if (wnTextContent.Value != null || (wnGridParam.DataSource != null && wnGridParam.RowCount > 0))
            {
                DialogResult result = DialogBox.AskYesNoCancel(this.GetCurrLanguageContent("WnFormTemplet.AskSave")); //Content_CN：量化成果内容已改变，是否保存已编辑的内容？
                if (result == DialogResult.Yes) //是
                {
                    /* 关闭窗体 */
                    return;
                }
                else if (result == DialogResult.No) //否
                {
                    /* 将全局变量结果字符串置为空 */
                    this._result = "";

                    /* 关闭窗体 */
                    return;
                }
                else if (result == DialogResult.Cancel) //取消
                {
                    /* 不关闭窗体 */
                    e.Cancel = true;
                }
            }
        }

        #endregion

        #region 对象方法

        /// <summary>
        /// 响应“确定”按钮的Click事件
        /// </summary>
        private void tsButtonOK_Click(object sender, EventArgs e)
        {
            /* 关闭窗体，将结果字符串和参数信息传递给WnFormMain */
            this._askClose = false;
            this.Close();
        }

        /// <summary>
        /// 响应“取消”按钮的Click事件
        /// </summary>
        private void tsButtonCancel_Click(object sender, EventArgs e)
        {
            /* 询问是否确定关闭窗体 */
            this.Close();
        }

        /// <summary>
        /// 响应wnTreeGridMain的AfterCheckNode事件
        /// </summary>
        private void wnTreeGridMain_AfterCheckNode(object sender, NodeEventArgs e)
        {
            /* 若当前节点不是父节点，直接返回 */
            if (e.Node.Nodes.Count <= 0) { return; }

            /* 勾选或取消勾选父节点的同时，全选或取消全选当前节点的子节点 */
            foreach (TreeListNode node in e.Node.Nodes)
            {
                node.Checked = e.Node.Checked;
            }
        }

        #region 图片切换事件

        /// <summary>
        /// 响应wnImageRight的MouseEnter事件
        /// </summary>
        private void wnImageRight_MouseEnter(object sender, EventArgs e)
        {
            /* 切换为蓝色图标 */
            (sender as WnImageEditorBase).Image = Properties.Resources.BlueRight1_32x22;
        }

        /// <summary>
        /// 响应wnImageRight的MouseLeave事件
        /// </summary>
        private void wnImageRight_MouseLeave(object sender, EventArgs e)
        {
            /* 切换回灰色图标 */
            (sender as WnImageEditorBase).Image = Properties.Resources.GrayRight1_32x22;
        }

        /// <summary>
        /// 响应wnImageLeft的MouseEnter事件
        /// </summary>
        private void wnImageLeft_MouseEnter(object sender, EventArgs e)
        {
            /* 切换为蓝色图标 */
            (sender as WnImageEditorBase).Image = Properties.Resources.BlueLeft1_32x22;
        }

        /// <summary>
        /// 响应wnImageLeft的MouseLeave事件
        /// </summary>
        private void wnImageLeft_MouseLeave(object sender, EventArgs e)
        {
            /* 切换回灰色图标 */
            (sender as WnImageEditorBase).Image = Properties.Resources.GrayLeft1_32x22;
        }

        /// <summary>
        /// 响应wnImageAllRight的MouseEnter事件
        /// </summary>
        private void wnImageAllRight_MouseEnter(object sender, EventArgs e)
        {
            /* 切换为蓝色图标 */
            (sender as WnImageEditorBase).Image = Properties.Resources.BlueRight2_32x22;
        }

        /// <summary>
        /// 响应wnImageAllRight的MouseLeave事件
        /// </summary>
        private void wnImageAllRight_MouseLeave(object sender, EventArgs e)
        {
            /* 切换回灰色图标 */
            (sender as WnImageEditorBase).Image = Properties.Resources.GrayRight2_32x22;
        }

        /// <summary>
        /// 响应wnImageAllLeft的MouseEnter事件
        /// </summary>
        private void wnImageAllLeft_MouseEnter(object sender, EventArgs e)
        {
            /* 切换为蓝色图标 */
            (sender as WnImageEditorBase).Image = Properties.Resources.BlueLeft2_32x22;
        }

        /// <summary>
        /// 响应wnImageAllLeft的MouseLeave事件
        /// </summary>
        private void wnImageAllLeft_MouseLeave(object sender, EventArgs e)
        {
            /* 切换回灰色图标 */
            (sender as WnImageEditorBase).Image = Properties.Resources.GrayLeft2_32x22;
        }

        #endregion

        /// <summary>
        /// 响应wnImageRight的MouseDown事件
        /// </summary>
        private void wnImageRight_MouseDown(object sender, MouseEventArgs e)
        {
            /* 若没有勾选节点，直接退出 */
            if (wnTreeGridMain.CheckedNodes.Count <= 0) { return; }

            /* 获取原始数据 */
            Hashtable htOriginal = GetOriginalData();
            Hashtable htReplaceInfo = null;
            DataTable dtOriginalParam = null;
            if (htOriginal != null)
            {
                htReplaceInfo = htOriginal["ReplaceInfo"] as Hashtable;
                dtOriginalParam = htOriginal["OriginalParams"] as DataTable;
            }

            /* 清空“内容”文本框 */
            wnTextContent.Value = DBNull.Value;

            /* 设置“内容”文本框和“参数信息”列表的值 */
            SetTempletContent(wnTreeGridMain.CheckedNodes, htReplaceInfo);
            SetWnGridParam(dtOriginalParam);
        }

        /// <summary>
        /// 响应wnImageLeft的MouseDown事件
        /// </summary>
        private void wnImageLeft_MouseDown(object sender, MouseEventArgs e)
        {
            /* 如果没有勾选数据，选中当前行 */
            if (wnGridParam.SelectedRowsArray.Length == 0)
            {
                wnGridParam.SetRowSelected(wnGridParam.CurrRow);
            }

            /* 每选中一条参数记录，选中对应模板的所有参数记录 */
            foreach (DataRow dr in wnGridParam.SelectedRowsArray)
            {
                DataRow[] drs = wnGridParam.DataSource.Select("TempletGuid = '" + dr["TempletGuid"] + "'");
                for (int i = 0; i < drs.Length; i++)
                {
                    wnGridParam.SetRowSelected(drs[i]);
                }
            }

            /* 选中记录的节点取消勾选 */
            foreach (DataRow dr in wnGridParam.SelectedRowsArray)
            {
                TreeListNode node = wnTreeGridMain.FindNodeByFieldValue("Guid", dr["TempletGuid"].ToString());
                node.Checked = false;

                /* 从Grid中删除勾选记录 */
                wnGridParam.DeleteRow(dr);
            }

            /* 获取所有父节点 */
            List<TreeListNode> parentNodes = new List<TreeListNode>();
            foreach (TreeListNode node in wnTreeGridMain.CheckedNodes)
            {
                if (node.Nodes.Count > 0)
                {
                    parentNodes.Add(node);
                }
            }

            /* 判断父节点是否勾选 */
            foreach (TreeListNode node in parentNodes)
            {
                /* 判断父节点下是否还存在未勾选的子节点 */
                bool flag = false;
                for (int i = 0; i < node.Nodes.Count; i++)
                {
                    if (!node.Nodes[i].Checked)
                    {
                        flag = true;
                        break;
                    }
                }

                /* 如果存在未勾选的子节点，取消勾选父节点 */
                node.Checked = !flag;
            }

            /* 获取原始数据 */
            Hashtable htOriginalData = GetOriginalData();
            Hashtable htReplaceInfo = null;
            if (htOriginalData != null)
            {
                htReplaceInfo = htOriginalData["ReplaceInfo"] as Hashtable;
            }

            /* 为“内容”文本框赋值 */
            SetTempletContent(wnTreeGridMain.CheckedNodes, htReplaceInfo);
        }

        /// <summary>
        /// 响应wnImageAllRight的MouseDown事件
        /// </summary>
        private void wnImageAllRight_MouseDown(object sender, MouseEventArgs e)
        {
            /* 获取原始数据 */
            Hashtable htOriginal = GetOriginalData();
            Hashtable htReplaceInfo = null;
            DataTable dtOriginal = null;
            if (htOriginal != null)
            {
                htReplaceInfo = htOriginal["ReplaceInfo"] as Hashtable;
                dtOriginal = htOriginal["OriginalParams"] as DataTable;
            }

            /* 清空文本框和界面数据 */
            wnTextContent.Value = DBNull.Value;
            wnGridParam.DataSource.Clear();

            /* 全选节点 */
            wnTreeGridMain.CheckAll();

            /* 设置“内容”文本框和“参数信息”列表的值 */
            SetTempletContent(wnTreeGridMain.CheckedNodes, htReplaceInfo);
            SetWnGridParam(dtOriginal);
        }

        /// <summary>
        /// 响应wnImageAllLeft的MouseDown事件
        /// </summary>
        private void wnImageAllLeft_MouseDown(object sender, MouseEventArgs e)
        {
            /* 取消勾选模板 */
            wnTreeGridMain.UncheckAll();

            /* 清空“内容”文本框和“参数信息”列表的行 */
            wnTextContent.Value = DBNull.Value;
            wnGridParam.DataSource.Clear();

            /* 更新全局变量 */
            this._result = "";
            this._templetContent = "";
        }

        /// <summary>
        /// 响应wnGridParam的CellValueChanging事件
        /// </summary>
        private void wnGridParam_CellValueChanging(object sender, CellValueChangingArgs e)
        {
            /* 获取当前已经编辑和正在编辑的模板信息 */
            Hashtable htReplaceInfo = new Hashtable();
            foreach (DataRow dr in wnGridParam.DataSource.Rows)
            {
                /* 判断当前行对应的模板是否存在多个参数 */
                DataRow[] drs = wnGridParam.DataSource.Select("TempletGuid = '" + dr["TempletGuid"] + "'");
                if (drs.Length == 1)
                {
                    if (e.CurrRow == dr)
                    {
                        /* 如果当前行是正在编辑的行，直接替换正在编辑的值 */
                        string originalTemplet = wnTreeGridMain.FindNodeByFieldValue("Guid", dr["TempletGuid"].ToString()).GetValue("Content").ToString();
                        string modifiedTemplet = originalTemplet.Replace("{" + dr["Param"] + "}", e.CurrEditingValue.ToString());

                        /* 加到Hashtable中 */
                        htReplaceInfo[originalTemplet] = modifiedTemplet;
                    }
                    else
                    {
                        /* 不是正在编辑行，如果当前参数值为空，跳过 */
                        if (string.IsNullOrEmpty(dr["Value"].ToString())) { continue; }

                        /* 如果模板只有一个参数，且有参数值，直接替换已有的值 */
                        string originalTemplet = wnTreeGridMain.FindNodeByFieldValue("Guid", dr["TempletGuid"].ToString()).GetValue("Content").ToString();
                        string modifiedTemplet = originalTemplet.Replace("{" + dr["Param"] + "}", dr["Value"].ToString());

                        /* 加到Hashtable中 */
                        htReplaceInfo[originalTemplet] = modifiedTemplet;
                    }
                }
                else
                {
                    /* 获取原始模板 */
                    string originalTemplet = wnTreeGridMain.FindNodeByFieldValue("Guid", dr["TempletGuid"].ToString()).GetValue("Content").ToString();
                    string modifiedTemplet = originalTemplet;

                    /* 如果模板有多个参数，遍历本模板的所有参数，每个参数值都替换 */
                    for (int i = 0; i < drs.Length; i++)
                    {
                        /* 如果当前行是正在编辑行，用正在编辑的行替换模板参数 */
                        if (e.CurrRow == drs[i])
                        {
                            modifiedTemplet = modifiedTemplet.Replace("{" + drs[i]["Param"] + "}", e.CurrEditingValue.ToString());
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(drs[i]["Value"].ToString())) { continue; }

                            /* 如果当前行不是正在编辑行，且当前行参数值不为空，替换参数 */
                            modifiedTemplet = modifiedTemplet.Replace("{" + drs[i]["Param"] + "}", drs[i]["Value"].ToString());
                        }
                    }

                    /* 加到Hashtable中 */
                    htReplaceInfo[originalTemplet] = modifiedTemplet;
                }
            }

            /* 将替换后的模板数据替换到原始模板中 */
            string result = _templetContent;
            foreach (string origi in htReplaceInfo.Keys)
            {
                result = result.Replace(origi, htReplaceInfo[origi].ToString());
            }

            /* 为"内容"文本框赋值 */
            string referenceStr = "模板：\n" + _templetContent + "\n\n==========\n内容：\n";
            wnTextContent.Value = referenceStr + result;

            /* 文本框滚动到插入字符的位置 */
            RichTextBox text = ((wnTextContent.OriginControl as WnTextEditorBase).Controls[0] as RichTextBox);
            text.Select(text.TextLength, 0);
            text.ScrollToCaret();

            /* 更新全局变量结果字符串 */
            this._result = result;
        }

        #endregion

        #region 内部方法

        /// <summary>
        /// 默认展开通用分类模板和用户设置的常用量化成果分类模板，并移动至最前
        /// </summary>
        private void HandleDfTempClassify()
        {
            /* 默认收缩全部分类 */
            wnTreeGridMain.CollapseAll();

            /* 展开第一个分类（通用分类） */
            wnTreeGridMain.Nodes.FirstNode.ExpandAll();

            /* 获取当前用户设置的默认展开的成果模板分类 */
            Hashtable htDfTempClassify = new Hashtable() { { "DataSetName", "GetResultTempClassify" } };
            htDfTempClassify["{EmployeeName}"] = AppInfo.UserName;
            if (this.DataSource.GetDataSet(htDfTempClassify)["IsSuccess"].Equals("0")) { return; }

            /* 若没有获取到当前用户设置的默认模板分类的Guid，提示用户后返回 */
            DataTable dtDfTempClassify = (this.DataSource.DataSets["GetResultTempClassify"] as DataSet).Tables["T01"];
            if (dtDfTempClassify.Rows.Count <= 0)
            {
                DialogBox.ShowError(this.GetCurrLanguageContent("WnFormTemplet.NoDefaultTempClassify")); //Content_CN：当前用户没有设置默认模板分类，只展开通用分类模板！
                return;
            }

            /* 如果设置的常用模板分类是通用，直接返回 */
            string common = wnTreeGridMain.Nodes.FirstNode.GetValue("Guid").ToString();
            string tempClassify = dtDfTempClassify.Rows[0][0].ToString();
            if (common.Equals(tempClassify)) { return; }

            /* 展开默认模板分类节点，并移动至通用节点之后 */
            TreeListNode dfNode = wnTreeGridMain.FindNodeByFieldValue("Guid", tempClassify);
            dfNode.ExpandAll();
            wnTreeGridMain.SetNodeIndex(dfNode, 1);
        }

        /// <summary>
        /// 获取初始模板信息
        /// </summary>
        private Hashtable GetOriginalData()
        {
            /* 当WnFormTemplet上没有数据时，直接返回 */
            if (wnTreeGridMain.CheckedNodes.Count == 0 && (wnTextContent.Value == null || !string.IsNullOrEmpty(wnTextContent.Value.ToString())) && wnGridParam.RowCount == 0) { return null; }

            /* 判断是否存在值不为空的参数 */
            bool isChanged = false;
            foreach (DataRow dr in wnGridParam.DataSource.Rows)
            {
                if (!string.IsNullOrEmpty(dr["Value"].ToString()))
                {
                    isChanged = true;
                    break;
                }
            }

            /* 如果不存在，直接返回 */
            if (!isChanged) { return null; }

            /* 记录初始参数值 */
            DataTable dtOriginalParam = wnGridParam.DataSource.Copy();

            /* 记录初始参数模板信息 */
            Hashtable htReplaceInfo = new Hashtable();
            foreach (DataRow dr in wnGridParam.DataSource.Rows)
            {
                /* 获取当前行对应模板 */
                string originalTemplet = wnTreeGridMain.FindNodeByFieldValue("Guid", dr["TempletGuid"].ToString()).GetValue("Content").ToString();

                /* 判断当前行对应的模板是否存在多个参数 */
                DataRow[] drs = wnGridParam.DataSource.Select("TempletGuid = '" + dr["TempletGuid"] + "'");
                if (drs.Length == 1)
                {
                    /* 如果模板只有一个参数，且当前行没有值，直接跳过 */
                    if (string.IsNullOrEmpty(dr["Value"].ToString())) { continue; }

                    /* 替换模板参数 */
                    string modifiedTemplet = originalTemplet.Replace("{" + dr["Param"] + "}", dr["Value"].ToString());
                    htReplaceInfo[originalTemplet] = modifiedTemplet;

                    /* 加到Hashtable中 */
                    htReplaceInfo[originalTemplet] = modifiedTemplet;
                }
                else
                {
                    /* 如果模板有多个参数，遍历本模板的所有参数，每个参数值都替换 */
                    string modifiedTemplet = originalTemplet;
                    for (int i = 0; i < drs.Length; i++)
                    {
                        /* 如果模板的当前参数没有值，直接跳过，处理模板的下一个参数 */
                        if (string.IsNullOrEmpty(drs[i]["Value"].ToString())) { continue; }

                        /* 如果当前行参数值不为空，替换参数 */
                        modifiedTemplet = modifiedTemplet.Replace("{" + drs[i]["Param"] + "}", drs[i]["Value"].ToString());
                    }

                    /* 加到Hashtable中 */
                    htReplaceInfo[originalTemplet] = modifiedTemplet;
                }
            }

            /* 返回初始信息 */
            Hashtable ht = new Hashtable();
            ht["OriginalParams"] = dtOriginalParam;
            ht["ReplaceInfo"] = htReplaceInfo;
            return ht;
        }

        /// <summary>
        /// 根据选择节点设置“内容”文本框的值
        /// </summary>
        private void SetTempletContent(List<TreeListNode> checkedNodes, Hashtable htReplaceInfo)
        {
            /* 根据选中模板拼接“内容”文本框的值 */
            StringBuilder templetContent = new StringBuilder();
            Hashtable htTempName = new Hashtable();

            /* 选中节点按照OrderIndex列进行正序排序 */
            checkedNodes.Sort((TreeListNode node1, TreeListNode node2) =>
            {
                int index1 = Convert.ToInt32(node1.GetValue("OrderIndex"));
                int index2 = Convert.ToInt32(node2.GetValue("OrderIndex"));
                if (index1 < index2)
                {
                    return -1;
                }
                else if (index1 > index2)
                {
                    return 1;
                }
                return 0;
            });

            foreach (TreeListNode node in checkedNodes)
            {
                /* 若当前节点的模板内容为空，跳过该节点 */
                if (string.IsNullOrEmpty(node.GetValue("Content").ToString())) { continue; }

                /* 拼接勾选模板的模板内容 */
                string tempName = node.GetValue("TempName").ToString();
                if (htTempName.Contains(tempName))
                {
                    /* 相同模板名称的模板在同一行，以分号分隔 */
                    templetContent.Append(node.GetValue("Content") + "；");
                }
                else
                {
                    /* 不同模板名称的模板内容在不同的行 */
                    templetContent.Append("\n" + node.GetValue("TempName") + "：" + node.GetValue("Content") + "；");

                    /* 记录当前模板名称 */
                    htTempName[tempName] = tempName;
                }
            }

            if (templetContent.Length > 0)
            {
                templetContent.Remove(templetContent.Length - 1, 1);
                templetContent.Append("。");
            }

            /* 给“内容”文本框赋值 */
            wnTextContent.Value = "模板：" + templetContent + "\n\n==========\n内容：" + templetContent;

            /* 更新全局变量模板内容的值 */
            _templetContent = templetContent.ToString().Trim();
            this._result = _templetContent;

            /* 若WnFormTemplet的初始数据，则直接返回 */
            if (htReplaceInfo == null) { return; }

            /* 获取初始参考内容和初始模板内容 */
            string referenceStr = "模板：\n" + _templetContent + "\n\n==========\n内容：\n";
            string originalStr = _templetContent;

            /* 将原始参数更新到“内容”文本框 */
            string result = originalStr;
            foreach (string origi in htReplaceInfo.Keys)
            {
                /* 遍历全部原始参数，替换原始模板内容 */
                result = result.Replace(origi, htReplaceInfo[origi].ToString());
            }
            wnTextContent.Value = referenceStr + result;
            this._result = result;
        }

        /// <summary>
        /// 设置WnGridParam的数据
        /// </summary>
        private void SetWnGridParam(DataTable dtOriginal)
        {
            /* 根据原始参数列表获取初始选择的节点 */
            List<TreeListNode> origiCheckedNodes = new List<TreeListNode>();
            if (dtOriginal != null)
            {
                foreach (DataRow dr in dtOriginal.Rows)
                {
                    TreeListNode ndoe = wnTreeGridMain.FindNodeByFieldValue("Guid", dr["TempletGuid"].ToString());
                    origiCheckedNodes.Add(ndoe);
                }
            }

            /* 清空“参数信息”列表 */
            wnGridParam.DataSource.Clear();

            /* 将原始数据复制到WnGridParam中 */
            if (dtOriginal != null)
            {
                wnGridParam.DataSource = dtOriginal;
            }

            /* 选中模板的数据添加到“参数信息”列表 */
            foreach (TreeListNode node in wnTreeGridMain.CheckedNodes)
            {
                /* 跳过原始勾选节点,不在重复添加参数 */
                if (origiCheckedNodes.Contains(node))
                {
                    continue;
                }

                /* 如果当前模板没有内容，直接跳过 */
                if (string.IsNullOrEmpty(node.GetValue("ParamList").ToString()))
                {
                    continue;
                }

                /* 新增一行 */
                string[] tempParams = node.GetValue("ParamList").ToString().Split(new char[] { ',' });
                string templetGuid = node.GetValue("Guid").ToString();
                string templetName = node.GetValue("TempName").ToString();
                int orderIndex = Convert.ToInt32(node.GetValue("OrderIndex"));
                for (int i = 0; i < tempParams.Length; i++)
                {
                    DataRow dr = wnGridParam.DataSource.NewRow();
                    dr["Guid"] = Guid.NewGuid().ToString();
                    dr["TempletGuid"] = templetGuid;
                    dr["Templet"] = templetName;
                    dr["Param"] = tempParams[i];
                    dr["OrderIndex"] = orderIndex;
                    wnGridParam.DataSource.Rows.Add(dr);
                }
            }

            /* 按OrderIndex列正序排序 */
            wnGridParam.Columns["OrderIndex"].SortOrder = DevExpress.Data.ColumnSortOrder.Ascending;
        }

        #endregion
    }
}