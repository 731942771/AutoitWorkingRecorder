#Region ;**** 由 AccAu3Wrapper_GUI 创建指令 ****
#AccAu3Wrapper_Icon=favicon.ico
#AccAu3Wrapper_OutFile=C:\Documents and Settings\Administrator\桌面\vigiles.exe
#AccAu3Wrapper_Compression=4
#AccAu3Wrapper_Res_Comment=威格灵博客
#AccAu3Wrapper_Res_Description=www.cuiweiyou.com
#AccAu3Wrapper_Res_Fileversion=8.8.8.8
#AccAu3Wrapper_Res_ProductVersion=9.9.9.9
#AccAu3Wrapper_Res_LegalCopyright=vigiles
#AccAu3Wrapper_Res_Language=2052
#AccAu3Wrapper_Res_requestedExecutionLevel=None
#AccAu3Wrapper_Res_Field=OriginalFilename|崔维友
#AccAu3Wrapper_Res_Field=ProductName|崔维友
#AccAu3Wrapper_Res_Field=ProductVersion|V1.0
#AccAu3Wrapper_Res_Field=InternalName|崔维友
#AccAu3Wrapper_Res_Field=FileDescription|崔维友
#AccAu3Wrapper_Res_Field=Comments|崔维友
#AccAu3Wrapper_Res_Field=LegalTrademarks|cuiweiyou.com
#AccAu3Wrapper_Res_Field=CompanyName|cuiweiyou.com
#AccAu3Wrapper_Run_AU3Check=n
#AutoIt3Wrapper_UseX64=n
#Tidy_Parameters=/sfc/rel
#AccAu3Wrapper_Tidy_Stop_OnError=n
#EndRegion ;**** 由 AccAu3Wrapper_GUI 创建指令 ****

#include <GUIConstantsEx.au3>
#include <GUIEdit.au3>
#include <GuiImageList.au3>
#include <GuiListView.au3>
#include <GuiMenu.au3>
#include <GuiTab.au3>
#include <GuiToolbar.au3>
#include <GuiToolTip.au3>
#include <StaticConstants.au3>
#include <WinAPI.au3>
#Include <WinAPIEx.au3>
#include <WindowsConstants.au3>

;======================================================================
Opt("GUIOnEventMode", 1)

;======================================================================
; 全局的数据库路径
Global $db_path = @ScriptDir & "\mdbZhongHui.mdb"
Global $db_pswd = ""	; 此密码是mdb文件的密码

; 用到的表单：          公司人员                   学校信息                      公司产品                  资产管理                     合作伙伴                       工程管理                   工作日志                   用户                     元数据              操作日志
Global $tb_workers = "tb_workers", $tb_schools = "tb_schools", $tb_products = "tb_products", $tb_accets = "tb_accets", $tb_partners = "tb_partners", $tb_projects = "tb_projects", $tb_journal = "tb_journal", $tb_users = "tb_users", $tb_source = "tb_source", $tb_log = "tb_log"

;-------------------- 源数据操作 -------------------------------
If FileExists ( $db_path ) = 0 Then
	;If MsgBox(52, "警告", "未检测到数据库文件，是否重新创建？") = 6 Then
		FuncCreateDb ( )
	;EndIf
EndIf

;======================================================================
;-------------------- 用户登录 --------------------------------


;======================================================================
; 单击标签切换LV时用到的
Global $itemInToolbar, $idOfTabItem	; 添加一个全局变量$idOfTabItem记录被右击的标签索引
; 主窗体宽高
Global $WidthOfWindow = 1100, $HeightOfWindow = 600
; 主窗体，工具条，工具条提示器
Global $guiMainWindow, $toolbarInMainWindow, $hToolTip
; 工具条上的按钮既定ID
Global Enum $id_Toolbar_New = 1000, $id_Toolbar_Save, $id_Toolbar_Delete, $id_Toolbar_Find, $id_Toolbar_Help
; 数据库读取标识。读过的True
Global $hasReadedDbTbWorkers = False, $hasReadedDbTbSchools = False, $hasReadedDbTbProducts = False, $hasReadedDbTbAccets = False, $hasReadedDbTbPartners = False, $hasReadedDbTbProjects = False, $hasReadedDbTbUsers = False, $hasReadedDbTbSource = False, $hasReadedDbTbLog = False, $hasReadedDbTbJournal = False
; 双击、右击单元格时创建编辑框用到的
Global $hEdit, $hBrush, $hDC, $hItemRow, $hItemColumn	; 双击LV进行修改用到了
; 右击LV的项目（行）时右键菜单的ID
Global $mouseMenuLV, $mouseMenuItemDelLV, $mouseMenuItemUpdateLV, $mouseMenuItemCopyLV, $dataIndex
;
Global $hEnableListView
;
Global Enum $id_menu_lv_del = 2000, $id_menu_lv_update, $id_menu_lv_copy
;
Global $hLvInTabJournal, $hLvInTabAccets, $hLvInTabPartners, $hLvInTabProducts, $hLvInTabProjects, $hLvInTabSchools, $hLvInTabSources, $hLvInTabWorkers, $hLvInTabUsers, $hlvintabLog
; 双击单元格修改数据，保存日志到Log表时用到的。修改的数据属于哪个表，属于哪个LV，对应的列名，修改前数据
Global $columnOfTable, $columnOfLV, $columnName, $columnOldData
; 子窗体，子窗体的确定按钮，取消按钮。
Global $popupWindow, $btnOKInPopWin, $btnNOInPopWin
; 查询时保存查询条件。插入数据也用到了
Global $strArgsOfSql, $strArgsOfLVArr[15], $strArgsOfTableArr[15]	; 根据本例需要

;======================================================================
;-------------------- GUI ---------------------------------
$guiMainWindow = GUICreate("威格灵", $WidthOfWindow, $HeightOfWindow)
	GUISetOnEvent($GUI_EVENT_CLOSE, "Func_GUI_EVENT_CLOSE")
	GUISetIcon( @ScriptDir & "\favicon.ico")	; 设置程序图标为脚本文件同目录中的Logo.ico

	$menuFile = GUICtrlCreateMenu ( "文件 &F")
		$itemOpenInMenuFile = GUICtrlCreateMenuItem("打开", $menuFile)
		$itemSaveInMenuFile = GUICtrlCreateMenuItem("保存", $menuFile)
		GUICtrlCreateMenuItem("", $menuFile) ; 分隔线
		$itemRecentfilesInMenuFile = GUICtrlCreateMenu("最近的文件", $menuFile)
		GUICtrlCreateMenuItem("", $menuFile)
		$itemExitInMenuFile = GUICtrlCreateMenuItem("退出", $menuFile)
			GUICtrlSetOnEvent($itemExitInMenuFile, "Func_GUI_EVENT_CLOSE")

	$menuTab = GUICtrlCreateMenu ( "窗口 &W")
		$menuItemJournal = GUICtrlCreateMenuItem("√ 工作日志", $menuTab)
			GUICtrlSetOnEvent($menuItemJournal, "Func_ShowTab_ByMenu")
		$menuItemProducts = GUICtrlCreateMenuItem("  公司产品", $menuTab)
			GUICtrlSetOnEvent($menuItemProducts, "Func_ShowTab_ByMenu")
		$menuItemSchools = GUICtrlCreateMenuItem("  学校信息", $menuTab)
			GUICtrlSetOnEvent($menuItemSchools, "Func_ShowTab_ByMenu")
		$menuItemPartners = GUICtrlCreateMenuItem("  合作伙伴", $menuTab)
			GUICtrlSetOnEvent($menuItemPartners, "Func_ShowTab_ByMenu")
		GUICtrlCreateMenuItem("", $menuTab) ; 分隔线
		$menuItemProjects = GUICtrlCreateMenuItem("  工程管理", $menuTab)
			GUICtrlSetOnEvent($menuItemProjects, "Func_ShowTab_ByMenu")
		$menuItemUsers = GUICtrlCreateMenuItem("  用户管理", $menuTab)
			GUICtrlSetOnEvent($menuItemUsers, "Func_ShowTab_ByMenu")
		$menuItemAccets = GUICtrlCreateMenuItem("  资产管理", $menuTab)
			GUICtrlSetOnEvent($menuItemAccets, "Func_ShowTab_ByMenu")
		$menuItemWorkers = GUICtrlCreateMenuItem("  人员管理", $menuTab)
			GUICtrlSetOnEvent($menuItemWorkers, "Func_ShowTab_ByMenu")
		GUICtrlCreateMenuItem("", $menuTab) ; 分隔线
		$menuItemLog = GUICtrlCreateMenuItem("  操作日志", $menuTab)
			GUICtrlSetOnEvent($menuItemLog, "Func_ShowTab_ByMenu")
		$menuItemSources = GUICtrlCreateMenuItem("  元 数 据", $menuTab)
			GUICtrlSetOnEvent($menuItemSources, "Func_ShowTab_ByMenu")

	$menuHelp = GUICtrlCreateMenu ( "帮助 &H")
		$itemGuideInMenuHelp = GUICtrlCreateMenuItem("使用指南", $menuHelp)
			GUICtrlSetOnEvent($itemGuideInMenuHelp, "Func_MenuHelp")	; 打开”公司产品”标签卡
		$itemAboutInMenuHelp = GUICtrlCreateMenuItem("关于", $menuHelp)
			GUICtrlSetOnEvent($itemAboutInMenuHelp, "Func_MenuHelp")	; 打开”学校信息”标签卡

	$toolbarInMainWindow = _GUICtrlToolbar_Create($guiMainWindow)
		; 创建工具提示控件，然后才能在$WM_NOTIFY中获取消息响应
		$hToolTip = _GUIToolTip_Create($toolbarInMainWindow)
		_GUICtrlToolbar_SetToolTips($toolbarInMainWindow, $hToolTip)

		_GUICtrlToolbar_AddBitmap($toolbarInMainWindow, 1, -1, $IDB_STD_SMALL_COLOR)	; 添加图标
		_GUICtrlToolbar_AddButton($toolbarInMainWindow, $id_Toolbar_New, $STD_FILENEW)				; 新建
		_GUICtrlToolbar_AddButton($toolbarInMainWindow, $id_Toolbar_Save, $STD_FILESAVE)			; 保存
		_GUICtrlToolbar_AddButton($toolbarInMainWindow, $id_Toolbar_Delete, $STD_DELETE)			; 删除
		_GUICtrlToolbar_AddButtonSep($toolbarInMainWindow)
		_GUICtrlToolbar_AddButton($toolbarInMainWindow, $id_Toolbar_Find, $STD_FIND)				; 查找
		_GUICtrlToolbar_AddButtonSep($toolbarInMainWindow)
		_GUICtrlToolbar_AddButton($toolbarInMainWindow, $id_Toolbar_Help, $STD_HELP)				; 帮助

	$tabInMainWindow = GUICtrlCreateTab ( 1, 28, $WidthOfWindow - 1, $HeightOfWindow - 70)
		GUICtrlSetOnEvent($tabInMainWindow, "Func_Tab_Click")

		$imgList = _GUIImageList_Create()
			_GUIImageList_Add($imgList, _WinAPI_CreateSolidBitmap($guiMainWindow, 0x990033, 16, 16))	; 0
			_GUIImageList_Add($imgList, _WinAPI_CreateSolidBitmap($guiMainWindow, 0x00FF00, 16, 16))
			_GUIImageList_Add($imgList, _WinAPI_CreateSolidBitmap($guiMainWindow, 0x0000FF, 16, 16))
			_GUIImageList_Add($imgList, _WinAPI_CreateSolidBitmap($guiMainWindow, 0xFF3399, 16, 16))
			_GUIImageList_Add($imgList, _WinAPI_CreateSolidBitmap($guiMainWindow, 0x007700, 16, 16))
			_GUIImageList_Add($imgList, _WinAPI_CreateSolidBitmap($guiMainWindow, 0xFF6600, 16, 16))	; 5
			_GUIImageList_Add($imgList, _WinAPI_CreateSolidBitmap($guiMainWindow, 0x663399, 16, 16))
			_GUIImageList_Add($imgList, _WinAPI_CreateSolidBitmap($guiMainWindow, 0x337777, 16, 16))
			_GUIImageList_Add($imgList, _WinAPI_CreateSolidBitmap($guiMainWindow, 0x7700FF, 16, 16))
			_GUIImageList_Add($imgList, _WinAPI_CreateSolidBitmap($guiMainWindow, 0xCC33CC, 16, 16))	; 9
			_GUICtrlTab_SetImageList($tabInMainWindow, $imgList)

		_GUICtrlTab_InsertItem ( $tabInMainWindow, 0, "工作日志", 0)
		;_GUICtrlTab_InsertItem ( $tabInMainWindow, 1, "公司产品", 1)
		;_GUICtrlTab_InsertItem ( $tabInMainWindow, 2, "学校信息", 2)
		;_GUICtrlTab_InsertItem ( $tabInMainWindow, 3, "合作伙伴", 3)
		;_GUICtrlTab_InsertItem ( $tabInMainWindow, 4, "工程管理", 4)
		;_GUICtrlTab_InsertItem ( $tabInMainWindow, 5, "用户管理", 5)
		;_GUICtrlTab_InsertItem ( $tabInMainWindow, 6, "资产管理", 6)
		;_GUICtrlTab_InsertItem ( $tabInMainWindow, 7, "人员管理", 7)
		;_GUICtrlTab_InsertItem ( $tabInMainWindow, 8, "操作日志", 8)
		;_GUICtrlTab_InsertItem ( $tabInMainWindow, 9, "元 数 据", 9)

	$mouseMenuTab = GUICtrlCreateContextMenu($tabInMainWindow)
		$mouseMenuItemClose = GUICtrlCreateMenuItem("关闭", $mouseMenuTab)
			GUICtrlSetOnEvent($mouseMenuItemClose, "Func_MouseMenuItem")
		GUICtrlCreateMenuItem("", $mouseMenuTab)
		$mouseMenuItemSaveAs = GUICtrlCreateMenuItem("另存为", $mouseMenuTab)
			GUICtrlSetOnEvent($mouseMenuItemSaveAs, "Func_MouseMenuItem")
		GUICtrlCreateMenuItem("", $mouseMenuTab)
		$mouseMenuItemPrint = GUICtrlCreateMenuItem("打印", $mouseMenuTab)
			GUICtrlSetOnEvent($mouseMenuItemPrint, "Func_MouseMenuItem")

	$strInfoLbl = GUICtrlCreateLabel("", 2, $HeightOfWindow - 38, $WidthOfWindow - 5, 20)

	$lvInTabJournal = GUICtrlCreateListView ( "编号|人员|日期|地点|交通|食宿|工作描述|备注", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)		; 工作日志表
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabJournal, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))			; 统计设置多种扩展样式
		GUICtrlSetBkColor($lvInTabJournal, 0xffffff)					;设置listview的背景色
		GUICtrlSetBkColor($lvInTabJournal, $GUI_BKCOLOR_LV_ALTERNATE)	;奇数行为listview的背景色，偶数行为listviewitem的背景色
		GUICtrlSetState($lvInTabJournal, $GUI_SHOW)
		$hLvInTabJournal = GUICtrlGetHandle($lvInTabJournal)
	$lvInTabProducts = GUICtrlCreateListView ( "编号|产品|类型|设计|配置|造价|备注", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)		; 公司产品表
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabProducts, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabProducts, 0xffffff)
		GUICtrlSetBkColor($lvInTabProducts, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabProducts, $GUI_HIDE)
		$hLvInTabProducts = GUICtrlGetHandle($lvInTabProducts)
	$lvInTabSchools  = GUICtrlCreateListView ( "编号|学校|联系人|地址|电话|邮箱|备注", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)	; 学校信息表
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabSchools, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabSchools, 0xffffff)
		GUICtrlSetBkColor($lvInTabSchools, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabSchools, $GUI_HIDE)
		$hLvInTabSchools = GUICtrlGetHandle($lvInTabSchools)
	$lvInTabPartners = GUICtrlCreateListView ( "编号|名称|类型|地址|电话|邮箱|业务|备注", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)	; 合作伙伴表
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabPartners, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabPartners, 0xffffff)
		GUICtrlSetBkColor($lvInTabPartners, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabPartners, $GUI_HIDE)
		$hLvInTabPartners = GUICtrlGetHandle($lvInTabPartners)
	$lvInTabProjects = GUICtrlCreateListView ( "编号|名称|学校|产品|合作者|我司负责人|细则|起始日期|状态|结束日期|结算记录|备注", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabProjects, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabProjects, 0xffffff)
		GUICtrlSetBkColor($lvInTabProjects, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabProjects, $GUI_HIDE)
		$hLvInTabProjects = GUICtrlGetHandle($lvInTabProjects)
	; 用户管理
	$lvInTabUsers = GUICtrlCreateListView ( "编号|用户名|密码|权限|备注", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabUsers, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabUsers, 0xffffff)
		GUICtrlSetBkColor($lvInTabUsers, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabUsers, $GUI_HIDE)
		$hLvInTabUsers = GUICtrlGetHandle($lvInTabUsers)
	; 资产管理
	$lvInTabAccets = GUICtrlCreateListView ( "编号|名称|串号|单位|类型|购入日期|单价|所属部门|经销商|是否报废|备注", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabAccets, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabAccets, 0xffffff)
		GUICtrlSetBkColor($lvInTabAccets, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabAccets, $GUI_HIDE)
		$hLvInTabAccets = GUICtrlGetHandle($lvInTabAccets)
	; 人员管理
	$lvInTabWorkers = GUICtrlCreateListView ( "编号|姓名|身份证号|部门|职务|入职日期|转正日期|月薪资|性别|生日|电话|邮箱|住址|家庭信息|备注", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabWorkers, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabWorkers, 0xffffff)
		GUICtrlSetBkColor($lvInTabWorkers, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabWorkers, $GUI_HIDE)
		$hLvInTabWorkers = GUICtrlGetHandle($lvInTabWorkers)
	; 操作日志
	$lvInTabLog = GUICtrlCreateListView ( "编号|用户|操作|表单|列名|旧数据|新数据|日期|备注", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabLog, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabLog, 0xffffff)
		GUICtrlSetBkColor($lvInTabLog, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabLog, $GUI_HIDE)
		$hlvintabLog = GUICtrlGetHandle($lvInTabLog)
	; 元数据
	$lvInTabSources = GUICtrlCreateListView ( "编号|表单|列|值|描述", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabSources, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabSources, 0xffffff)
		GUICtrlSetBkColor($lvInTabSources, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabSources, $GUI_HIDE)
		$hLvInTabSources = GUICtrlGetHandle($lvInTabSources)

GUISetState(@SW_SHOW, $guiMainWindow)

; #cs
;	初始在“日志”数据库查询，在“日志”列表插入数据
; #ce
FuncReadDb( $tb_journal, $lvInTabJournal )
$hasReadedDbTbJournal = True

;======================================================================
GUIRegisterMsg($WM_NOTIFY, "_WM_NOTIFY")
GUIRegisterMsg($WM_COMMAND, '_WM_COMMAND')
GUIRegisterMsg($WM_CONTEXTMENU, "_WM_CONTEXTMENU")

;======================================================================
While 1
	Sleep(200)
WEnd

;======================================================================
Func Func_GUI_EVENT_CLOSE ()
	Exit
EndFunc

; #cs
; 鼠标左键点击标签卡项
; 切换标签的选中状态
; #ce
Func Func_Tab_Click ()
	Local $ctrlId = GUICtrlRead (@GUI_CtrlId)
	Local $itemText = _GUICtrlTab_GetItemText(@GUI_CtrlId, $ctrlId)

	GUICtrlSetState($lvInTabJournal, $GUI_HIDE)
	GUICtrlSetState($lvInTabProducts, $GUI_HIDE)
	GUICtrlSetState($lvInTabSchools, $GUI_HIDE)
	GUICtrlSetState($lvInTabProjects, $GUI_HIDE)
	GUICtrlSetState($lvInTabPartners, $GUI_HIDE)

	GUICtrlSetState($lvInTabUsers, $GUI_HIDE)
	GUICtrlSetState($lvInTabAccets, $GUI_HIDE)
	GUICtrlSetState($lvInTabWorkers, $GUI_HIDE)
	GUICtrlSetState($lvInTabLog, $GUI_HIDE)
	GUICtrlSetState($lvInTabSources, $GUI_HIDE)

	Switch $itemText
		Case "工作日志"
			GUICtrlSetState($lvInTabJournal, $GUI_SHOW)
			GUICtrlSetData($strInfoLbl, "列表《工作日志》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabJournal ))
		Case "公司产品"
			GUICtrlSetState($lvInTabProducts, $GUI_SHOW)
			GUICtrlSetData($strInfoLbl, "列表《公司产品》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabProducts ))
		Case "学校信息"
			GUICtrlSetState($lvInTabSchools, $GUI_SHOW)
			GUICtrlSetData($strInfoLbl, "列表《学校信息》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabSchools ))
		Case "工程管理"
			GUICtrlSetState($lvInTabProjects, $GUI_SHOW)
			GUICtrlSetData($strInfoLbl, "列表《工程管理》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabProjects ))
		Case "合作伙伴"
			GUICtrlSetState($lvInTabPartners, $GUI_SHOW)
			GUICtrlSetData($strInfoLbl, "列表《合作伙伴》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabPartners ))

		Case "用户管理"
			GUICtrlSetState($lvInTabUsers, $GUI_SHOW)
			GUICtrlSetData($strInfoLbl, "列表《用户管理》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabUsers ))
		Case "资产管理"
			GUICtrlSetState($lvInTabAccets, $GUI_SHOW)
			GUICtrlSetData($strInfoLbl, "列表《资产管理》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabAccets ))
		Case "人员管理"
			GUICtrlSetState($lvInTabWorkers, $GUI_SHOW)
			GUICtrlSetData($strInfoLbl, "列表《人员管理》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabWorkers ))
		Case "操作日志"
			GUICtrlSetState($lvInTabLog, $GUI_SHOW)
			GUICtrlSetData($strInfoLbl, "列表《操作日志》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabLog ))
		Case "元 数 据"
			GUICtrlSetState($lvInTabSources, $GUI_SHOW)
			GUICtrlSetData($strInfoLbl, "列表《元 数 据》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabSources ))
	EndSwitch
EndFunc

; 消息监听
; 监听UDF创建的工具条，添加提示，点击互动
; 监听标签卡右键点击事件
; 监听...
Func _WM_NOTIFY($hWndGUI, $MsgID, $wParam, $lParam)
	Local $tNMHDR, $hwndFrom, $code, $i_idOld, $i_idNew
    Local $tNMTBHOTITEM

    $tNMHDR = DllStructCreate($tagNMHDR, $lParam)
    $hwndFrom = DllStructGetData($tNMHDR, "hWndFrom")
    $code = DllStructGetData($tNMHDR, "Code")

	Local $tInfo, $iID, $iCode
	$tInfo = DllStructCreate($tagNMTTDISPINFO, $lParam)
    $iCode = DllStructGetData($tInfo, "Code")
    If $iCode = $TTN_GETDISPINFOW Then
        $iID = DllStructGetData($tInfo, "IDFrom")
        Switch $iID
			Case $id_Toolbar_New
				DllStructSetData($tInfo, "aText", "在当前数据表插入新数据")
			Case $id_Toolbar_Save
				DllStructSetData($tInfo, "aText", "保存当前数据表到Excel")
			Case $id_Toolbar_Delete
				DllStructSetData($tInfo, "aText", "删除当前数据表选择的行")
			Case $id_Toolbar_Find
				DllStructSetData($tInfo, "aText", "在当前数据表查找数据")
			Case $id_Toolbar_Help
				DllStructSetData($tInfo, "aText", "帮助")
		EndSwitch
	EndIf

    Switch $hwndFrom	; 控件
        Case $toolbarInMainWindow	;----------- 工具栏
            Switch $code	; 事件
				Case $TBN_HOTITEMCHANGE
					$tNMTBHOTITEM = DllStructCreate($tagNMTBHOTITEM, $lParam)
					$i_idOld = DllStructGetData($tNMTBHOTITEM, "idOld")
					$i_idNew = DllStructGetData($tNMTBHOTITEM, "idNew")
					$itemInToolbar = $i_idNew

				Case $NM_CLICK	; 左键点击
					Switch $itemInToolbar
						Case $id_Toolbar_New
							$itemInToolbar = -1

							FuncInsertItemToListView()	; 插入条目到ListView

						Case $id_Toolbar_Save
							$itemInToolbar = -1

						Case $id_Toolbar_Delete
							$itemInToolbar = -1

							FuncDeleteItemFromListView()

						Case $id_Toolbar_Find
							$itemInToolbar = -1

							FuncFindData()	; 在当前显示的LV对应表查找

						Case $id_Toolbar_Help
							$itemInToolbar = -1

					EndSwitch

			EndSwitch

		Case GUICtrlGetHandle($tabInMainWindow)	;-------------------- 标签页
            Switch $code
				Case $NM_RCLICK	; 右击

					Local $x, $y, $aHit
					$x = _WinAPI_GetMousePosX(True, GUICtrlGetHandle($tabInMainWindow))
					$y = _WinAPI_GetMousePosY(True, GUICtrlGetHandle($tabInMainWindow))
					$aHit = _GUICtrlTab_HitTest($tabInMainWindow, $x, $y)

					$idOfTabItem = $aHit[0]

					_GUICtrlTab_SetCurSel($tabInMainWindow, $aHit[0])

					GUICtrlSetState($lvInTabJournal, $GUI_HIDE)
					GUICtrlSetState($lvInTabProducts, $GUI_HIDE)
					GUICtrlSetState($lvInTabSchools, $GUI_HIDE)
					GUICtrlSetState($lvInTabProjects, $GUI_HIDE)
					GUICtrlSetState($lvInTabPartners, $GUI_HIDE)

					GUICtrlSetState($lvInTabUsers, $GUI_HIDE)
					GUICtrlSetState($lvInTabAccets, $GUI_HIDE)
					GUICtrlSetState($lvInTabWorkers, $GUI_HIDE)
					GUICtrlSetState($lvInTabLog, $GUI_HIDE)
					GUICtrlSetState($lvInTabSources, $GUI_HIDE)

					Switch _GUICtrlTab_GetItemText ( $tabInMainWindow, $aHit[0] )
						Case "工作日志"
							GUICtrlSetState($lvInTabJournal, $GUI_SHOW)
							GUICtrlSetData($strInfoLbl, "列表《工作日志》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabJournal ))
						Case "公司产品"
							GUICtrlSetState($lvInTabProducts, $GUI_SHOW)
							GUICtrlSetData($strInfoLbl, "列表《公司产品》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabProducts ))
						Case "学校信息"
							GUICtrlSetState($lvInTabSchools, $GUI_SHOW)
							GUICtrlSetData($strInfoLbl, "列表《学校信息》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabSchools ))
						Case "工程管理"
							GUICtrlSetState($lvInTabProjects, $GUI_SHOW)
							GUICtrlSetData($strInfoLbl, "列表《工程管理》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabProjects ))
						Case "合作伙伴"
							GUICtrlSetState($lvInTabPartners, $GUI_SHOW)
							GUICtrlSetData($strInfoLbl, "列表《合作伙伴》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabPartners ))
						Case "用户管理"
							GUICtrlSetState($lvInTabUsers, $GUI_SHOW)
							GUICtrlSetData($strInfoLbl, "列表《用户管理》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabUsers ))
						Case "资产管理"
							GUICtrlSetState($lvInTabAccets, $GUI_SHOW)
							GUICtrlSetData($strInfoLbl, "列表《资产管理》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabAccets ))
						Case "人员管理"
							GUICtrlSetState($lvInTabWorkers, $GUI_SHOW)
							GUICtrlSetData($strInfoLbl, "列表《人员管理》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabWorkers ))
						Case "操作日志"
							GUICtrlSetState($lvInTabLog, $GUI_SHOW)
							GUICtrlSetData($strInfoLbl, "列表《操作日志》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabLog ))
						Case "元 数 据"
							GUICtrlSetState($lvInTabSources, $GUI_SHOW)
							GUICtrlSetData($strInfoLbl, "列表《元 数 据》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabSources ))
					EndSwitch
			EndSwitch

		Case GUICtrlGetHandle($lvInTabJournal)	;--------------------- 列表控件：工作日志
            Switch $code
				Case $NM_DBLCLK	; 双击
					$columnOfTable = $tb_journal
					$columnOfLV = $lvInTabJournal

					FuncCreateEditRecForColumn("工作日志", $lvInTabJournal, $lParam )

				Case $NM_CLICK	; 左键点击
					Local $ts = _GUICtrlListView_GetItemTextString ( $lvInTabJournal )
					If StringReplace ( $ts, "|", "") <> "" Then
						GUICtrlSetData($strInfoLbl, "列表《工作日志》中选中的行：" & $ts )
					Else
						GUICtrlSetData($strInfoLbl, "" )
					EndIf
				#cs
				Case $NM_RCLICK	; 右击
					$columnOfTable = $tb_journal
					$columnOfLV = $lvInTabJournal
					$columnOldData = _GUICtrlListView_GetItemTextString ( $lvInTabJournal )

					Local $tInfo = DllStructCreate($tagNMITEMACTIVATE, $lParam)
					$dataIndex = DllStructGetData($tInfo, "Index")		; 点击的行。始于0
					Local $dataSubItem = DllStructGetData($tInfo, "SubItem")	; 点击的列。始于0
					ConsoleWrite("右击----------行：" & $dataIndex & "，列：" & $dataSubItem & @CRLF)

					If $dataIndex = -1 Then
						_GUICtrlMenu_DestroyMenu(GUICtrlGetHandle($mouseMenuLV))
					Else
						$mouseMenuLV = GUICtrlCreateContextMenu($lvInTabJournal)
							$mouseMenuItemDelLV = GUICtrlCreateMenuItem("删除", $mouseMenuLV)
								GUICtrlSetOnEvent($mouseMenuItemDelLV, "Func_MouseMenuItem_LV")
							GUICtrlCreateMenuItem("", $mouseMenuLV)
							$mouseMenuItemUpdateLV = GUICtrlCreateMenuItem("修改", $mouseMenuLV)
								GUICtrlSetOnEvent($mouseMenuItemUpdateLV, "Func_MouseMenuItem_LV")
							GUICtrlCreateMenuItem("", $mouseMenuLV)
							$mouseMenuItemCopyLV = GUICtrlCreateMenuItem("复制", $mouseMenuLV)
								GUICtrlSetOnEvent($mouseMenuItemCopyLV, "Func_MouseMenuItem_LV")
				EndIf
				#ce
			EndSwitch
		Case GUICtrlGetHandle($lvInTabProducts)
            Switch $code
				Case $NM_DBLCLK
					$columnOfTable = $tb_products
					$columnOfLV = $lvInTabProducts
					FuncCreateEditRecForColumn("公司产品", $lvInTabProducts, $lParam )
				Case $NM_CLICK
					Local $ts = _GUICtrlListView_GetItemTextString ( $lvInTabProducts )
					If StringReplace ( $ts, "|", "") <> "" Then
						GUICtrlSetData($strInfoLbl, "列表《公司产品》中选中的行：" & $ts )
					Else
						GUICtrlSetData($strInfoLbl, "" )
					EndIf
			EndSwitch
		Case GUICtrlGetHandle($lvInTabProjects)
            Switch $code
				Case $NM_DBLCLK
					$columnOfTable = $tb_projects
					$columnOfLV = $lvInTabProjects
					FuncCreateEditRecForColumn("工程管理", $lvInTabProjects, $lParam )
				Case $NM_CLICK
					Local $ts = _GUICtrlListView_GetItemTextString ( $lvInTabProjects )
					If StringReplace ( $ts, "|", "") <> "" Then
						GUICtrlSetData($strInfoLbl, "列表《工程管理》中选中的行：" & $ts )
					Else
						GUICtrlSetData($strInfoLbl, "" )
					EndIf
			EndSwitch
		Case GUICtrlGetHandle($lvInTabPartners)
            Switch $code
				Case $NM_DBLCLK
					$columnOfTable = $tb_partners
					$columnOfLV = $lvInTabPartners
					FuncCreateEditRecForColumn("合作伙伴", $lvInTabPartners, $lParam )
				Case $NM_CLICK
					Local $ts = _GUICtrlListView_GetItemTextString ( $lvInTabPartners )
					If StringReplace ( $ts, "|", "") <> "" Then
						GUICtrlSetData($strInfoLbl, "列表《合作伙伴》中选中的行：" & $ts )
					Else
						GUICtrlSetData($strInfoLbl, "" )
					EndIf
			EndSwitch
		Case GUICtrlGetHandle($lvInTabSchools)
            Switch $code
				Case $NM_DBLCLK
					$columnOfTable = $tb_schools
					$columnOfLV = $lvInTabSchools
					FuncCreateEditRecForColumn("学校信息", $lvInTabSchools, $lParam )
				Case $NM_CLICK
					Local $ts = _GUICtrlListView_GetItemTextString ( $lvInTabSchools )
					If StringReplace ( $ts, "|", "") <> "" Then
						GUICtrlSetData($strInfoLbl, "列表《学校信息》中选中的行：" & $ts )
					Else
						GUICtrlSetData($strInfoLbl, "" )
					EndIf
			EndSwitch

		Case GUICtrlGetHandle($lvInTabUsers)
            Switch $code
				Case $NM_DBLCLK
					$columnOfTable = $tb_users
					$columnOfLV = $lvInTabUsers
					FuncCreateEditRecForColumn("用户管理", $lvInTabUsers, $lParam )
				Case $NM_CLICK
					Local $ts = _GUICtrlListView_GetItemTextString ( $lvInTabUsers )
					If StringReplace ( $ts, "|", "") <> "" Then
						GUICtrlSetData($strInfoLbl, "列表《用户管理》中选中的行：" & $ts )
					Else
						GUICtrlSetData($strInfoLbl, "" )
					EndIf
			EndSwitch
		Case GUICtrlGetHandle($lvInTabAccets)
            Switch $code
				Case $NM_DBLCLK
					$columnOfTable = $tb_accets
					$columnOfLV = $lvInTabAccets
					FuncCreateEditRecForColumn("资产管理", $lvInTabAccets, $lParam )
				Case $NM_CLICK
					Local $ts = _GUICtrlListView_GetItemTextString ( $lvInTabAccets )
					If StringReplace ( $ts, "|", "") <> "" Then
						GUICtrlSetData($strInfoLbl, "列表《资产管理》中选中的行：" & $ts )
					Else
						GUICtrlSetData($strInfoLbl, "" )
					EndIf
			EndSwitch
		Case GUICtrlGetHandle($lvInTabWorkers)
            Switch $code
				Case $NM_DBLCLK
					$columnOfTable = $tb_workers
					$columnOfLV = $lvInTabWorkers
					FuncCreateEditRecForColumn("人员管理", $lvInTabWorkers, $lParam )
				Case $NM_CLICK
					Local $ts = _GUICtrlListView_GetItemTextString ( $lvInTabWorkers )
					If StringReplace ( $ts, "|", "") <> "" Then
						GUICtrlSetData($strInfoLbl, "列表《人员管理》中选中的行：" & $ts )
					Else
						GUICtrlSetData($strInfoLbl, "" )
					EndIf
			EndSwitch
		Case GUICtrlGetHandle($lvInTabLog)
            Switch $code
				Case $NM_DBLCLK
					$columnOfTable = $tb_log
					$columnOfLV = $lvInTabLog
					FuncCreateEditRecForColumn("操作日志", $lvInTabLog, $lParam )
				Case $NM_CLICK
					Local $ts = _GUICtrlListView_GetItemTextString ( $lvInTabLog )
					If StringReplace ( $ts, "|", "") <> "" Then
						GUICtrlSetData($strInfoLbl, "列表《操作日志》中选中的行：" & $ts )
					Else
						GUICtrlSetData($strInfoLbl, "" )
					EndIf
			EndSwitch
		Case GUICtrlGetHandle($lvInTabSources)
            Switch $code
				Case $NM_DBLCLK
					$columnOfTable = $tb_source
					$columnOfLV = $lvInTabSources
					FuncCreateEditRecForColumn("元 数 据", $lvInTabSources, $lParam )
				Case $NM_CLICK
					Local $ts = _GUICtrlListView_GetItemTextString ( $lvInTabSources )
					If StringReplace ( $ts, "|", "") <> "" Then
						GUICtrlSetData($strInfoLbl, "列表《元 数 据》中选中的行：" & $ts )
					Else
						GUICtrlSetData($strInfoLbl, "" )
					EndIf
			EndSwitch

	EndSwitch

    Return $GUI_RUNDEFMSG
EndFunc

; 在LV上的右键菜单
Func Func_MouseMenuItem_LV ()
	Switch @GUI_CtrlId
		Case $mouseMenuItemDelLV
			; 使用全局变量
			; $columnOfTable：	表名称
			; $columnOfLV：		列表ID
			; $columnOldData：	整行的内容
			FuncDeleteItemFromAccessAndListView ( $columnOfTable, $columnOfLV, $columnOldData)
	EndSwitch

	; $mouseMenuLV, $mouseMenuItemDelLV, $mouseMenuItemUpdateLV, $mouseMenuItemCopyLV
	GUICtrlDelete($mouseMenuItemDelLV)
	GUICtrlDelete($mouseMenuItemUpdateLV)
	GUICtrlDelete($mouseMenuItemCopyLV)
	GUICtrlDelete($mouseMenuLV)
EndFunc

;#cs
; 在当前列表中查询数据
; 累加不同的字段集合为查询条件
; 在$strArgsOfLVArr中保存关键字控件的ID
; 在$strArgsOfTableArr中保存表、列
;#ce
Func FuncFindData()
	Local $strLvName
	$strLvName = _GUICtrlTab_GetItemText ( $tabInMainWindow, _GUICtrlTab_GetCurSel ( $tabInMainWindow ))

	Local $popWinWidth = 500, $popWinHeight = 180

	$popupWindow = GUICreate("" , 500, 175, Default, Default, $WS_CAPTION, BitOR($WS_EX_TOPMOST, $GUI_WS_EX_PARENTDRAG))
		WinSetTitle($popupWindow, "", "在《" & $strLvName & "》中查询数据" )

		GUICtrlCreateGroup("在需要的条件里输入关键字。多个词用空格分开", 10, 10, $popWinWidth - 20, $popWinHeight - 45)

		; 获取当前激活的LV
		Switch $strLvName
			Case "工作日志"		;  人员 日期 地点 交通 食宿 工作描述。id, j_name, j_date, j_address, j_traffic, j_board, j_content, j_note, j_record, j_date_record
				$strArgsOfLVArr[0] = $lvInTabJournal	; LV
				$strArgsOfTableArr[0] = $tb_journal		; Table

				GUICtrlCreateLabel("人员:", 20, 45, 30, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 52, 40, 180, 20)
				$strArgsOfTableArr[1] = "j_name"
				GUICtrlCreateLabel("日期:", 20, 75, 30, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 52, 70, 180, 20)
				$strArgsOfTableArr[2] = "j_date"
				GUICtrlCreateLabel("地点:", 20, 105, 30, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 52, 100, 180, 20)
				$strArgsOfTableArr[3] = "j_address"

				GUICtrlCreateLabel("交通:", 260, 45, 30, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 292, 40, 180, 20)
				$strArgsOfTableArr[4] = "j_traffic"
				GUICtrlCreateLabel("食宿:", 260, 75, 30, 20 )
				$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 292, 70, 180, 20)
				$strArgsOfTableArr[5] = "j_board"
				GUICtrlCreateLabel("工作:", 260, 105, 30, 20 )
				$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 292, 100, 180, 20)
				$strArgsOfTableArr[6] = "j_content"

			Case "公司产品"		; 产品 类型 设计 配置 造价。id, pd_name, pd_type, pd_desiger, pd_configuration, pd_cost, pd_note
				$strArgsOfLVArr[0] = $lvInTabProducts
				$strArgsOfTableArr[0] = $tb_products

				GUICtrlCreateLabel("产品:", 20, 45, 30, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 52, 40, 180, 20)
				$strArgsOfTableArr[1] = "pd_name"
				GUICtrlCreateLabel("类型:", 20, 75, 30, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 52, 70, 180, 20)
				$strArgsOfTableArr[2] = "pd_type"
				GUICtrlCreateLabel("设计:", 20, 105, 30, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 52, 100, 180, 20)
				$strArgsOfTableArr[3] = "pd_desiger"

				GUICtrlCreateLabel("配置:", 260, 45, 30, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 292, 40, 180, 20)
				$strArgsOfTableArr[4] = "pd_configuration"
				GUICtrlCreateLabel("造价:", 260, 75, 30, 20 )
				$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 292, 70, 180, 20)
				$strArgsOfTableArr[5] = "pd_cost"
			Case "学校信息"		; 学校 联系人 地址 电话 邮箱。id, s_name, s_contact, s_address, s_phone, s_email, s_note
				$strArgsOfLVArr[0] = $lvInTabSchools
				$strArgsOfTableArr[0] = $tb_schools

				GUICtrlCreateLabel("学  校:", 20, 45, 45, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 67, 40, 170, 20)
				$strArgsOfTableArr[1] = "s_name"
				GUICtrlCreateLabel("联系人:", 20, 75, 45, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 67, 70, 170, 20)
				$strArgsOfTableArr[2] = "s_contact"
				GUICtrlCreateLabel("地  址:", 20, 105, 45, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 67, 100, 170, 20)
				$strArgsOfTableArr[3] = "s_address"

				GUICtrlCreateLabel("电  话:", 260, 45, 45, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 307, 40, 170, 20)
				$strArgsOfTableArr[4] = "s_phone"
				GUICtrlCreateLabel("邮  箱:", 260, 75, 45, 20 )
				$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 307, 70, 170, 20)
				$strArgsOfTableArr[5] = "s_email"
			Case "工程管理"		; 名称 学校 产品 合作者 状态。id, pj_name, pj_s_name, pj_pd_name, pj_pt_name, pj_w_name, pj_content, pj_date_start, pj_state, pj_date_finish, pj_account, pj_note
				$strArgsOfLVArr[0] = $lvInTabProjects
				$strArgsOfTableArr[0] = $tb_projects

				GUICtrlCreateLabel("名  称:", 20, 45, 45, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 67, 40, 170, 20)
				$strArgsOfTableArr[1] = "pj_name"
				GUICtrlCreateLabel("学  校:", 20, 75, 45, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 67, 70, 170, 20)
				$strArgsOfTableArr[2] = "pj_s_name"
				GUICtrlCreateLabel("产  品:", 20, 105, 45, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 67, 100, 170, 20)
				$strArgsOfTableArr[3] = "pj_pt_name"

				GUICtrlCreateLabel("合作者:", 260, 45, 45, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 307, 40, 170, 20)
				$strArgsOfTableArr[4] = "pj_content"
				GUICtrlCreateLabel("状  态:", 260, 75, 45, 20 )
				$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 307, 70, 170, 20)
				$strArgsOfTableArr[5] = "pj_state"
			Case "合作伙伴"		; 名称 类型 地址 电话 邮箱 业务。id, pt_name, pt_type, pt_address, pt_phone, pt_email, pt_business, pt_note
				$strArgsOfLVArr[0] = $lvInTabPartners
				$strArgsOfTableArr[0] = $tb_partners

				GUICtrlCreateLabel("名称:", 20, 45, 30, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 52, 40, 180, 20)
				$strArgsOfTableArr[1] = "pt_name"
				GUICtrlCreateLabel("类型:", 20, 75, 30, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 52, 70, 180, 20)
				$strArgsOfTableArr[2] = "pt_type"
				GUICtrlCreateLabel("地址:", 20, 105, 30, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 52, 100, 180, 20)
				$strArgsOfTableArr[3] = "pt_address"

				GUICtrlCreateLabel("电话:", 260, 45, 30, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 292, 40, 180, 20)
				$strArgsOfTableArr[4] = "pt_phone"
				GUICtrlCreateLabel("邮箱:", 260, 75, 30, 20 )
				$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 292, 70, 180, 20)
				$strArgsOfTableArr[5] = "pt_email"
				GUICtrlCreateLabel("业务:", 260, 105, 30, 20 )
				$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 292, 100, 180, 20)
				$strArgsOfTableArr[6] = "pt_business"

			Case "用户管理"
				;编号|用户名|密码|权限|备注
				; id, u_name, u_pswd, u_authority, u_note
				$strArgsOfLVArr[0] = $lvInTabUsers
				$strArgsOfTableArr[0] = $tb_users

				GUICtrlCreateLabel("用户名:", 20, 35, 45, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 72, 30, 180, 20)
				$strArgsOfTableArr[1] = "u_name"
				GUICtrlCreateLabel("密  码:", 20, 60, 45, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 72, 55, 180, 20)
				$strArgsOfTableArr[2] = "u_pswd"
				GUICtrlCreateLabel("权  限:", 20, 90, 45, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 72, 85, 180, 20)
				$strArgsOfTableArr[3] = "u_authority"
				GUICtrlCreateLabel("备  注:", 20, 120, 45, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 72, 115, 180, 20)
				$strArgsOfTableArr[4] = "u_note"

			Case "资产管理"
				;编号|名称|串号|单位|类型|购入日期|单价|所属部门|经销商|是否报废|备注
				; id, a_name, a_serial_number, a_unit, a_type, a_date_bought, a_price, a_deportment, a_dealer, a_scrap, a_note
				$strArgsOfLVArr[0] = $lvInTabAccets	; LV
				$strArgsOfTableArr[0] = $tb_accets		; Table

				GUIDelete ( $popupWindow )

				$popWinHeight = 210

				$popupWindow = GUICreate("" , $popWinWidth, $popWinHeight, Default, Default, $WS_CAPTION, BitOR($WS_EX_TOPMOST, $GUI_WS_EX_PARENTDRAG))
					WinSetTitle($popupWindow, "", "在《" & $strLvName & "》中插入数据" )

					GUICtrlCreateGroup("在需要的条件里输入关键字。多个词用空格分开", 10, 5, 480, $popWinHeight - 45)

					GUICtrlCreateLabel("名    称:", 20, 25, 60, 20 )
					$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 82, 20, 150, 20)
					$strArgsOfTableArr[1] = "a_name"
					GUICtrlCreateLabel("串    号:", 20, 55, 60, 20 )
					$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 82, 50, 150, 20)
					$strArgsOfTableArr[2] = "a_serial_number"
					GUICtrlCreateLabel("单    位:", 20, 85, 60, 20 )
					$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 82, 80, 150, 20)
					$strArgsOfTableArr[3] = "a_unit"
					GUICtrlCreateLabel("类    型:", 20, 115, 60, 20 )
					$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 82, 110, 150, 20)
					$strArgsOfTableArr[4] = "a_type"
					GUICtrlCreateLabel("购入日期:", 20, 145, 60, 20 )
					$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 82, 140, 150, 20)
					$strArgsOfTableArr[5] = "a_date_bought"

					GUICtrlCreateLabel("单    价:", 260, 25, 60, 20 )
					$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 322, 20, 150, 20)
					$strArgsOfTableArr[6] = "a_price"
					GUICtrlCreateLabel("所属部门:", 260, 55, 60, 20 )
					$strArgsOfLVArr[7] = GUICtrlCreateInput ( "", 322, 50, 150, 20)
					$strArgsOfTableArr[7] = "a_deportment"
					GUICtrlCreateLabel("经 销 商:", 260, 85, 60, 20 )
					$strArgsOfLVArr[8] = GUICtrlCreateInput ( "", 322, 80, 150, 20)
					$strArgsOfTableArr[8] = "a_dealer"
					GUICtrlCreateLabel("是否报废:", 260, 115, 60, 20 )
					$strArgsOfLVArr[9] = GUICtrlCreateInput ( "", 322, 110, 150, 20)
					$strArgsOfTableArr[9] = "a_scrap"
					GUICtrlCreateLabel("备    注:", 260, 145, 60, 20 )
					$strArgsOfLVArr[10] = GUICtrlCreateInput ( "", 322, 140, 150, 20)
					$strArgsOfTableArr[10] = "a_note"

			Case "人员管理"
				; 姓名|身份证号|部门|职务|入职日期|转正日期|月薪资|性别|生日|电话|邮箱|住址|家庭信息|备注
				; w_name, w_identity_number, w_deportment, w_position, w_date_of_entry, w_date_regular, w_month_wage, w_sex, w_birthday, w_phone, w_email, w_address, w_family, w_note
				$strArgsOfLVArr[0] = $lvInTabWorkers
				$strArgsOfTableArr[0] = $tb_workers

				GUIDelete ( $popupWindow )

				$popWinHeight = 270

				$popupWindow = GUICreate("" , $popWinWidth, $popWinHeight, Default, Default, $WS_CAPTION, BitOR($WS_EX_TOPMOST, $GUI_WS_EX_PARENTDRAG))
					WinSetTitle($popupWindow, "", "在《" & $strLvName & "》中插入数据" )

					GUICtrlCreateGroup("在需要的条件里输入关键字。多个词用空格分开", 10, 5, 480, $popWinHeight - 45)

					GUICtrlCreateLabel("姓    名:", 20, 25, 60, 20 )
					$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 82, 20, 150, 20)
					$strArgsOfTableArr[1] = "w_name"
					GUICtrlCreateLabel("身份证号:", 20, 55, 60, 20 )
					$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 82, 50, 150, 20)
					$strArgsOfTableArr[2] = "w_identity_number"
					GUICtrlCreateLabel("部    门:", 20, 85, 60, 20 )
					$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 82, 80, 150, 20)
					$strArgsOfTableArr[3] = "w_deportment"
					GUICtrlCreateLabel("职    务:", 20, 115, 60, 20 )
					$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 82, 110, 150, 20)
					$strArgsOfTableArr[4] = "w_position"
					GUICtrlCreateLabel("入职日期:", 20, 145, 60, 20 )
					$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 82, 140, 150, 20)
					$strArgsOfTableArr[5] = "w_date_of_entry"
					GUICtrlCreateLabel("转正日期:", 20, 175, 60, 20 )
					$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 82, 170, 150, 20)
					$strArgsOfTableArr[6] = "w_date_regular"
					GUICtrlCreateLabel("月 薪 资:", 20, 205, 60, 20 )
					$strArgsOfLVArr[7] = GUICtrlCreateInput ( "", 82, 200, 150, 20)
					$strArgsOfTableArr[7] = "w_month_wage"

					GUICtrlCreateLabel("性    别:", 260, 25, 60, 20 )
					$strArgsOfLVArr[8] = GUICtrlCreateInput ( "", 322, 20, 150, 20)
					$strArgsOfTableArr[8] = "w_sex"
					GUICtrlCreateLabel("生    日:", 260, 55, 60, 20 )
					$strArgsOfLVArr[9] = GUICtrlCreateInput ( "", 322, 50, 150, 20)
					$strArgsOfTableArr[9] = "w_birthday"
					GUICtrlCreateLabel("电    话:", 260, 85, 60, 20 )
					$strArgsOfLVArr[10] = GUICtrlCreateInput ( "", 322, 80, 150, 20)
					$strArgsOfTableArr[10] = "w_phone"
					GUICtrlCreateLabel("邮    箱:", 260, 115, 60, 20 )
					$strArgsOfLVArr[11] = GUICtrlCreateInput ( "", 322, 110, 150, 20)
					$strArgsOfTableArr[11] = "w_email"
					GUICtrlCreateLabel("住    址:", 260, 145, 60, 20 )
					$strArgsOfLVArr[12] = GUICtrlCreateInput ( "", 322, 140, 150, 20)
					$strArgsOfTableArr[12] = "w_address"
					GUICtrlCreateLabel("家庭信息:", 260, 175, 60, 20 )
					$strArgsOfLVArr[13] = GUICtrlCreateInput ( "", 322, 170, 150, 20)
					$strArgsOfTableArr[13] = "w_family"
					GUICtrlCreateLabel("备    注:", 260, 205, 60, 20 )
					$strArgsOfLVArr[14] = GUICtrlCreateInput ( "", 322, 200, 150, 20)
					$strArgsOfTableArr[14] = "w_note"

			Case "操作日志"
				; 用户|登陆退出增删改|表|列|旧数据|新数据|日期|备注 8
				; l_name, l_operate, l_table, l_column, l_old_data, l_new_data, l_date, l_note
				$strArgsOfLVArr[0] = $lvInTabLog
				$strArgsOfTableArr[0] = $tb_log

				GUICtrlCreateLabel("用户:", 20, 30, 30, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 52, 25, 180, 20)
				$strArgsOfTableArr[1] = "l_name"
				GUICtrlCreateLabel("操作:", 20, 60, 30, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 52, 55, 180, 20)
				$strArgsOfTableArr[2] = "l_operate"
				GUICtrlCreateLabel("表单:", 20, 90, 30, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 52, 85, 180, 20)
				$strArgsOfTableArr[3] = "l_table"
				GUICtrlCreateLabel("列名:", 20, 120, 30, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 52, 115, 180, 20)
				$strArgsOfTableArr[4] = "l_column"

				GUICtrlCreateLabel("旧值:", 260, 30, 30, 20 )
				$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 292, 25, 180, 20)
				$strArgsOfTableArr[5] = "l_old_data"
				GUICtrlCreateLabel("新值:", 260, 60, 30, 20 )
				$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 292, 55, 180, 20)
				$strArgsOfTableArr[6] = "l_new_data"
				GUICtrlCreateLabel("日期:", 260, 90, 30, 20 )
				$strArgsOfLVArr[7] = GUICtrlCreateInput ( "", 292, 85, 180, 20)
				$strArgsOfTableArr[7] = "l_date"
				GUICtrlCreateLabel("备注:", 260, 120, 30, 20 )
				$strArgsOfLVArr[8] = GUICtrlCreateInput ( "", 292, 115, 180, 20)
				$strArgsOfTableArr[8] = "l_note"

			Case "元 数 据"
				; 表单|列|值|描述 4
				; sr_tb_name, sr_column, sr_value, sr_note
				$strArgsOfLVArr[0] = $lvInTabSources
				$strArgsOfTableArr[0] = $tb_source

				GUICtrlCreateLabel("表单:", 20, 30, 30, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 52, 25, 180, 20)
				$strArgsOfTableArr[1] = "sr_tb_name"
				GUICtrlCreateLabel("列名:", 20, 60, 30, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 52, 55, 180, 20)
				$strArgsOfTableArr[2] = "sr_column"
				GUICtrlCreateLabel("选值:", 20, 90, 30, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 52, 85, 180, 20)
				$strArgsOfTableArr[3] = "sr_value"
				GUICtrlCreateLabel("备注:", 20, 120, 30, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 52, 115, 180, 20)
				$strArgsOfTableArr[4] = "sr_note"

		EndSwitch

		GUICtrlCreateGroup("", -99, -99, 1, 1)

		$btnOKInPopWin = GUICtrlCreateButton("确定", 10, $popWinHeight - 30, 60, 22)
			GUICtrlSetOnEvent($btnOKInPopWin, "FuncFindDataByArgs")
		$btnNOInPopWin = GUICtrlCreateButton("取消", 428, $popWinHeight - 30, 60, 22)
			GUICtrlSetOnEvent($btnNOInPopWin, "FuncFindDataByArgs")

	GUISetState(@SW_SHOW, $popupWindow)
	GUISetState(@SW_DISABLE, $guiMainWindow)
EndFunc

Func FuncFindDataByArgs ()
	If @GUI_CtrlId = $btnOKInPopWin Then
		Local $isSqlStrEnable = False	; 查询语句是否可用。如果全部参数都是空的，不会设为true

		$strArgsOfSql = "SELECT * FROM " & $strArgsOfTableArr[0] & " WHERE "

		; 遍历控件数组，拼接语句
		For $iiii = 1 To 11 Step 1	; 0索引保存的是lV或Access名称。剩余 1-11 保存控件

			Local $strWidget = GUICtrlRead($strArgsOfLVArr[$iiii])

			If $strWidget <> "" And StringIsSpace($strWidget) <> 1 Then

				; 读控件内的文本，去除首、尾空格、连续多个空格合为1个
				Local $keywords = StringStripWS($strWidget, $STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES )
				; 根据关键词文本之间的空格进行拆组
				Local $keywordArr = StringSplit ($keywords, " ")
				; 只要有词语，至少能有一个元素[1]
				If $keywordArr[0] > 0 Then	; $keywordArr[0]保存有效元素数量
					; 如果有多个词，拼接
					For $j = 1 To $keywordArr[0]
						; 列 like %值%
						$strArgsOfSql &= $strArgsOfTableArr[$iiii] & " LIKE '%" & $keywordArr[$j] & "%' AND "
					Next
				EndIf

				$isSqlStrEnable = True	; 至少有一个有效参数，才可用
			EndIf
		Next

		$strArgsOfSql &= "1 = 1"	; 拼接sql语句常用结尾。否则删除最后一个 AND比较麻烦

		If $isSqlStrEnable Then
			; 执行数据库查询
			Local $adoCon = ObjCreate("ADODB.Connection")
			$adoCon.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & $db_path & "; Jet OLEDB:Database Password='" & $db_pswd & "'")
			Local $adoRec = ObjCreate("ADODB.Recordset")
			$adoRec.ActiveConnection = $adoCon

			$adoRec.Open($strArgsOfSql)	; 从表中查到的结果集

			Local $fieldsCount = $adoRec.fields.count	; 执行上一步后，得到表的字段数量
			If $strArgsOfTableArr[0] = $tb_journal Then
				$fieldsCount = $fieldsCount - 2
			EndIf

			; 清空列表
			_GUICtrlListView_DeleteAllItems ($strArgsOfLVArr[0])

			_GUICtrlListView_BeginUpdate($strArgsOfLVArr[0])
			While Not $adoRec.Eof And Not $adoRec.Bof	; 遍历结果集每一行
				If @error = 1 Then
					ExitLoop
				Else
					; 这里最好是创建一个新的标签、LV来展示查询结果

					Local $strResultDataForLv = ""

					For $i = 0 To $fieldsCount - 1 Step 1

						Local $strTmpS = $adoRec.fields( $i ).value
						If $strArgsOfTableArr[0] = "tb_log" Then
							$strTmpS = StringReplace ( $strTmpS, "|", "/")
						EndIf

						$strResultDataForLv &= $strTmpS & "|"
					Next

					GUICtrlCreateListViewItem($strResultDataForLv, $strArgsOfLVArr[0])
						GUICtrlSetBkColor (-1, 0xffa500 );设置listviewitem的背景色

					$adoRec.Movenext	; 结果集的下一行
				EndIf
			WEnd
			_GUICtrlListView_EndUpdate($strArgsOfLVArr[0])
			_GUICtrlListView_Scroll($strArgsOfLVArr[0], 0, _GUICtrlListView_GetItemCount($strArgsOfLVArr[0])*10)

			GUICtrlSetData($strInfoLbl, "表《" & $strArgsOfTableArr[0] & "》中查询到的记录数量：" & _GUICtrlListView_GetItemCount ( $strArgsOfLVArr[0] ))

			$adoRec.Close
			$adoCon.Close
		EndIf

	EndIf

	GUIDelete ( $popupWindow )
	GUISetState(@SW_ENABLE, $guiMainWindow)
	GUISetState(@SW_RESTORE, $guiMainWindow)

	; 复位。清空数组
	$strArgsOfSql = ""
	For $i = 0 To 14
		$strArgsOfLVArr[$i] = ""
		$strArgsOfTableArr[$i] = ""
	Next
EndFunc

; #cs
; 	双击LV的单元格后，创建一个输入框
; 	对原数据进行修改
; 	参数 $strLvInTab:要操作的列表ID
; 	参数 $lParam:系统消息
; #ce
Func FuncCreateEditRecForColumn($tmplvname, $strLvInTab, $lParam )
	; 创建结构体数据（结构体，提供数据的指针）
	Local $tInfo = DllStructCreate($tagNMITEMACTIVATE, $lParam)
	$hItemRow = DllStructGetData($tInfo, "Index")		; 点击的行。始于0
	$hItemColumn = DllStructGetData($tInfo, "SubItem")	; 点击的列。始于0

	Local $oADO = ObjCreate("ADODB.Connection")
	$oADO.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & $db_path & "; Jet OLEDB:Database Password='" & $db_pswd & "'")
	Local $oREC = ObjCreate("ADODB.Recordset")
	$oREC.ActiveConnection = $oADO
	$oREC.Open($columnOfTable, $oADO, 3, 3)	; 检索被双击的LV所属的表
	Local $fieldsCount = $oREC.fields.count	; 获取表单的列数，总共有多少字段
	; $oREC.fields[i].name	; 列名

	; 全局变量，_WM_COMMAND中用到
	$columnName = $oREC.fields($hItemColumn).name		; 得到单元格的列名。注意是小括号
	; $oREC.fields(1).value	; 列值

	$oREC.Close
	$oADO.Close

	Local $subItemRect = _GUICtrlListView_GetSubItemRect($strLvInTab, $hItemRow, $hItemColumn)
	Local $lvPosInWindow = ControlGetPos($guiMainWindow, '', $strLvInTab)
	$columnOldData = _GUICtrlListView_GetItemText($strLvInTab, $hItemRow, $hItemColumn)
	Local $iStyle = BitOR($WS_CHILD, $WS_VISIBLE, $ES_AUTOHSCROLL, $ES_LEFT)
	$hEdit = _GUICtrlEdit_Create($guiMainWindow, $columnOldData, $lvPosInWindow[0] + $subItemRect[0], $lvPosInWindow[1] + $subItemRect[1], _GUICtrlListView_GetColumnWidth($strLvInTab, $hItemColumn), 17, $iStyle)
	_GUICtrlEdit_SetSel($hEdit, 0, -1)
	_WinAPI_BringWindowToTop($hEdit)
	_WinAPI_SetFocus($hEdit)
	$hDC = _WinAPI_GetWindowDC($hEdit)
	$hBrush = _WinAPI_CreateSolidBrush(0xFF0000)

	Local $stRect = DllStructCreate('int;int;int;int')
	DllStructSetData($stRect, 1, 0)
	DllStructSetData($stRect, 2, 0)
	DllStructSetData($stRect, 3, _GUICtrlListView_GetColumnWidth($strLvInTab, $hItemColumn))
	DllStructSetData($stRect, 4, 22)

	_WinAPI_FrameRect($hDC, DllStructGetPtr($stRect), $hBrush)

	GUICtrlSetData($strInfoLbl, "列表《" & $tmplvname & "》中选中的行编号：" & _GUICtrlListView_GetItemText($strLvInTab, $hItemRow) & "，选中的列：" & _GUICtrlListView_GetColumn($strLvInTab, $hItemColumn)[5] & "，原数据：" & $columnOldData)
EndFunc

; 右键菜单 WM_CONTEXTMENU 消息
Func _WM_CONTEXTMENU($hWnd, $iMsg, $iwParam, $ilParam)
	Switch $iwParam
		Case $hLvInTabJournal, $hLvInTabAccets, $hLvInTabPartners, $hLvInTabProducts, $hLvInTabProjects, $hLvInTabSchools, $hLvInTabSources, $hLvInTabWorkers, $hLvInTabUsers, $hlvintabLog

			; 获取右键点击的有效索引
			If _GUICtrlListView_GetSelectedIndices ( $iwParam ) <> "" Then
				$columnOldData = _GUICtrlListView_GetItemTextString ($iwParam, _GUICtrlListView_GetSelectedIndices ( $iwParam ))
				Local $hMenu; $id_menu_lv_del = 2000, $id_menu_lv_update, $id_menu_lv_copy
				$hMenu = _GUICtrlMenu_CreatePopup()
				_GUICtrlMenu_InsertMenuItem($hMenu, 0, "删除", $id_menu_lv_del)
				_GUICtrlMenu_InsertMenuItem($hMenu, 1, "", 0)
				_GUICtrlMenu_InsertMenuItem($hMenu, 2, "修改", $id_menu_lv_update)
				_GUICtrlMenu_InsertMenuItem($hMenu, 3, "", 0)
				_GUICtrlMenu_InsertMenuItem($hMenu, 4, "复制", $id_menu_lv_copy)
				_GUICtrlMenu_TrackPopupMenu($hMenu, $hWnd)
				_GUICtrlMenu_DestroyMenu($hMenu)
				$hEnableListView = $iwParam
			EndIf

			Return True
	EndSwitch
EndFunc

; 控件失去焦点
; LV的右键菜单项处理
; 单元格新数据保存
Func _WM_COMMAND($hWnd, $msg, $wParam, $lParam)

	; 匹配是在哪个LV上右键的菜单
	Switch $hEnableListView
		Case $hLvInTabJournal
			; 匹配菜单项
			Switch $wParam
				Case $id_menu_lv_del
					; $strTbName： 表名称
					; $strLvInTab：列表ID
					; $columnOldData： 整行的内容
					FuncDeleteItemFromAccessAndListView ( "tb_journal", $lvInTabJournal, $columnOldData )

				Case $id_menu_lv_update

				Case $id_menu_lv_copy

			EndSwitch
		Case $hLvInTabAccets
			Switch $wParam
				Case $id_menu_lv_del
					FuncDeleteItemFromAccessAndListView ( "", , $columnOldData )
				Case $id_menu_lv_update

				Case $id_menu_lv_copy

			EndSwitch
		Case $hLvInTabPartners
			Switch $wParam
				Case $id_menu_lv_del
					FuncDeleteItemFromAccessAndListView ( "", , $columnOldData )
				Case $id_menu_lv_update

				Case $id_menu_lv_copy

			EndSwitch

		Case $hLvInTabProducts
			Switch $wParam
				Case $id_menu_lv_del
					FuncDeleteItemFromAccessAndListView ( "", , $columnOldData )
				Case $id_menu_lv_update

				Case $id_menu_lv_copy

			EndSwitch

		Case $hLvInTabProjects
			Switch $wParam
				Case $id_menu_lv_del
					FuncDeleteItemFromAccessAndListView ( "", , $columnOldData )
				Case $id_menu_lv_update

				Case $id_menu_lv_copy

			EndSwitch

		Case $hLvInTabSchools
			Switch $wParam
				Case $id_menu_lv_del
					FuncDeleteItemFromAccessAndListView ( "", , $columnOldData )
				Case $id_menu_lv_update

				Case $id_menu_lv_copy

			EndSwitch

		Case $hLvInTabWorkers
			Switch $wParam
				Case $id_menu_lv_del
					FuncDeleteItemFromAccessAndListView ( "", , $columnOldData )
				Case $id_menu_lv_update

				Case $id_menu_lv_copy

			EndSwitch

		Case $hLvInTabSources
			Switch $wParam
				Case $id_menu_lv_del
					FuncDeleteItemFromAccessAndListView ( "", , $columnOldData )
				Case $id_menu_lv_update

				Case $id_menu_lv_copy

			EndSwitch

		Case $hLvInTabUsers
			Switch $wParam
				Case $id_menu_lv_del
					FuncDeleteItemFromAccessAndListView ( "", , $columnOldData )
				Case $id_menu_lv_update

				Case $id_menu_lv_copy

			EndSwitch

		Case $hlvintabLog
			Switch $wParam
				Case $id_menu_lv_del
					FuncDeleteItemFromAccessAndListView ( "", , $columnOldData )
				Case $id_menu_lv_update

				Case $id_menu_lv_copy

			EndSwitch

	EndSwitch

	Local $iCode = BitShift($wParam, 16)
	Switch $lParam
		Case $hEdit
			Switch $iCode
				Case $EN_KILLFOCUS
					Local $iText = _GUICtrlEdit_GetText($hEdit)
					_GUICtrlListView_SetItemText($columnOfLV, $hItemRow, $iText, $hItemColumn)
					_WinAPI_DeleteObject($hBrush)
					_WinAPI_ReleaseDC($hEdit, $hDC)
					_WinAPI_DestroyWindow($hEdit)
					_GUICtrlEdit_Destroy($hEdit)

					; 如果输入的是新数据
					If $iText <> $columnOldData Then

						Local $tmpIndexOfDClick = _GUICtrlListView_GetItemText ( $columnOfLV, $hItemRow )

						;                               表                             列                新值             条件
						Local $strUpdate = "UPDATE " & $columnOfTable & " SET " & $columnName & "='" & $iText & "' WHERE id=" & $tmpIndexOfDClick
						Local $adoCon = ObjCreate("ADODB.Connection")
						$adoCon.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & $db_path & ";Jet Oledb:Database Password='" & $db_pswd & "'")
						$adoCon.execute($strUpdate)
						;$adoCon.close

						; 日志表
						Local $strLvItem = "'Eminem', 'update', '" & $columnOfTable & "', '" & $columnName & "', '" & $columnOldData & "', '" & $iText & "', '" & @YEAR & "-" & @MON & "-" & @MDAY & "(" & @WDAY - 1 & ")" & @HOUR & ":" & @MIN & ":" & @SEC & "`" & @MSEC & "', '_none'"	; 数据
						$strLvItem = "insert into " & $tb_log & " (l_name, l_operate, l_table, l_column, l_old_data, l_new_data, l_date, l_note) values ( " & $strLvItem & " )"				; 插入语句
						$adoCon.Execute($strLvItem)
						$adoCon.Close

					EndIf

					$hItemRow = -1
					$hItemColumn = 0
					$columnOldData = ""
					$columnOfTable = ""
					$columnOfLV = ""
					$columnName = ""

			EndSwitch
	EndSwitch
	Return $GUI_RUNDEFMSG
EndFunc

; #cs 删除当前列表中选择的行
; 获取标签索引，LV中行索引
; 删除这一行
; 在对于数据库中删除记录，在Log表中记录
; #ce
Func FuncDeleteItemFromListView()
	Local $strLvItem, $strLvInTab, $strTbName
	; 获取当前激活的LV
	Switch _GUICtrlTab_GetItemText ( $tabInMainWindow, _GUICtrlTab_GetCurSel ( $tabInMainWindow ))
		Case "工作日志"
			; 获取选中行的内容
			;$row = _GUICtrlListView_GetSelectedIndices ( $lvInTabJournal )			; 选中的条目的索引
			$strLvItem = _GUICtrlListView_GetItemTextString ( $lvInTabJournal, -1 )	; 选中的条目的整行内容

			$strLvInTab = $lvInTabJournal	; LV
			$strTbName = $tb_journal		; Table
		Case "公司产品"
			$strLvItem = _GUICtrlListView_GetItemTextString ( $lvInTabProducts, -1 )

			$strLvInTab = $lvInTabProducts
			$strTbName = $tb_products
		Case "学校信息"
			$strLvItem = _GUICtrlListView_GetItemTextString ( $lvInTabSchools, -1 )

			$strLvInTab = $lvInTabSchools
			$strTbName = $tb_schools
		Case "工程管理"
			$strLvItem = _GUICtrlListView_GetItemTextString ( $lvInTabProjects, -1 )

			$strLvInTab = $lvInTabProjects
			$strTbName = $tb_projects
		Case "合作伙伴"
			$strLvItem = _GUICtrlListView_GetItemTextString ( $lvInTabPartners, -1 )

			$strLvInTab = $lvInTabPartners
			$strTbName = $tb_partners

		Case "用户管理"
			$strLvItem = _GUICtrlListView_GetItemTextString ( $lvInTabUsers, -1 )

			$strLvInTab = $lvInTabUsers
			$strTbName = $tb_users

		Case "资产管理"
			$strLvItem = _GUICtrlListView_GetItemTextString ( $lvInTabAccets, -1 )

			$strLvInTab = $lvInTabAccets
			$strTbName = $tb_accets

		Case "人员管理"
			$strLvItem = _GUICtrlListView_GetItemTextString ( $lvInTabWorkers, -1 )

			$strLvInTab = $lvInTabWorkers
			$strTbName = $tb_workers

		Case "操作日志"
			$strLvItem = _GUICtrlListView_GetItemTextString ( $lvInTabLog, -1 )

			$strLvInTab = $lvInTabLog
			$strTbName = $tb_log

		Case "元 数 据"
			$strLvItem = _GUICtrlListView_GetItemTextString ( $lvInTabSources, -1 )

			$strLvInTab = $lvInTabSources
			$strTbName = $tb_source

	EndSwitch


	If _GUICtrlListView_GetSelectedCount ( $strLvInTab ) < 1 Then Return

	FuncDeleteItemFromAccessAndListView ( $strTbName, $strLvInTab, $strLvItem )
EndFunc

; 删除选中的行
; $strTbName： 表名称
; $strLvInTab：列表ID
; $strLvItem： 整行的内容
; 同步到数据库
Func FuncDeleteItemFromAccessAndListView ( $strTbName, $strLvInTab, $strLvItem )
	Local $indexOfDelItem = StringSplit($strLvItem, "|")

	Local $confirmDel = MsgBox(1 + 16, "命令确认", "你确定删除所选条目？")

	If $confirmDel = 1 Then
		_GUICtrlListView_DeleteItemsSelected ($strLvInTab)

		Local $dataAdodbConnectionDel = ObjCreate("ADODB.Connection")
		$dataAdodbConnectionDel.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & $db_path & ";Jet Oledb:Database Password='" & $db_pswd & "'")

		Local $strDel = "DELETE FROM " & $strTbName & " IN '" & $db_path & "' WHERE id" & " = " & $indexOfDelItem[1]

		$dataAdodbConnectionDel.execute($strDel)

		; 如果是在“日志表”中删除数据，就是绝对的清除记录。不应该有次操作
		If $strTbName <> "tb_log" Then	; 操作日志表应该是有很高权限的
			; 日志表
			$strLvItem = "'Eminem', 'del', '" & $strTbName & "', '_item', '" & StringTrimLeft ($strLvItem, StringInStr( $strLvItem, "|")) & "', '_none', '" & @YEAR & "-" & @MON & "-" & @MDAY & "(" & @WDAY - 1 & ")" & @HOUR & ":" & @MIN & ":" & @SEC & "`" & @MSEC & "', '_none'"	; 数据
			$strLvItem = "insert into " & $tb_log & " (l_name, l_operate, l_table, l_column, l_old_data, l_new_data, l_date, l_note) values ( " & $strLvItem & " )"				; 插入语句
			$dataAdodbConnectionDel.Execute($strLvItem)
		EndIf

		$dataAdodbConnectionDel.Close

		GUICtrlSetData($strInfoLbl, "列表《" & $strTbName & "》中记录数量：" & _GUICtrlListView_GetItemCount ( $strLvInTab ))
	EndIf

	$columnOfTable = ""
	$columnOfLV = ""
	$columnOldData = ""
EndFunc

; #cs
; 响应工具条“新建”按钮。插入条目到ListView
; 首先判断当前激活的标签-哪一个ListView处于显示状态
; 然后插入条目，更新条目位置
; #ce
Func FuncInsertItemToListView ()
	Local $strLvName = _GUICtrlTab_GetItemText ( $tabInMainWindow, _GUICtrlTab_GetCurSel ( $tabInMainWindow ))
	Local $popWinWidth = 500, $popWinHeight = 180

	$popupWindow = GUICreate("" , $popWinWidth, $popWinHeight, Default, Default, $WS_CAPTION, BitOR($WS_EX_TOPMOST, $GUI_WS_EX_PARENTDRAG))
		WinSetTitle($popupWindow, "", "在《" & $strLvName & "》中插入数据" )

		GUICtrlCreateGroup("", 10, 5, 480, 135)

		; 获取当前激活的LV
		Switch $strLvName
			Case "工作日志"
				; 人员 日期 地点 交通 食宿 工作描述 备注 7
				;j_name, j_date, j_address, j_traffic, j_board, j_content, j_note, j_record, j_date_record
				$strArgsOfLVArr[0] = $lvInTabJournal	; LV
				$strArgsOfTableArr[0] = $tb_journal		; Table

				GUICtrlCreateLabel("人员:", 20, 25, 30, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 52, 20, 180, 20)
				$strArgsOfTableArr[1] = "j_name"
				GUICtrlCreateLabel("日期:", 20, 55, 30, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 52, 50, 180, 20)
				$strArgsOfTableArr[2] = "j_date"
				GUICtrlCreateLabel("地点:", 20, 85, 30, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 52, 80, 180, 20)
				$strArgsOfTableArr[3] = "j_address"

				GUICtrlCreateLabel("备注:", 20, 115, 30, 20 )
				$strArgsOfLVArr[7] = GUICtrlCreateInput ( "", 52, 110, 420, 20)
				$strArgsOfTableArr[7] = "j_note"

				GUICtrlCreateLabel("交通:", 260, 25, 30, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 292, 20, 180, 20)
				$strArgsOfTableArr[4] = "j_traffic"
				GUICtrlCreateLabel("食宿:", 260, 55, 30, 20 )
				$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 292, 50, 180, 20)
				$strArgsOfTableArr[5] = "j_board"
				GUICtrlCreateLabel("工作:", 260, 85, 30, 20 )
				$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 292, 80, 180, 20)
				$strArgsOfTableArr[6] = "j_content"

			Case "公司产品"
				;产品 类型 设计 配置 造价 备注 6
				;pd_name, pd_type, pd_desiger, pd_configuration, pd_cost, pd_note
				$strArgsOfLVArr[0] = $lvInTabProducts	; LV
				$strArgsOfTableArr[0] = $tb_products		; Table

				GUICtrlCreateLabel("产品:", 20, 45, 30, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 52, 40, 180, 20)
				$strArgsOfTableArr[1] = "pd_name"
				GUICtrlCreateLabel("类型:", 20, 75, 30, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 52, 70, 180, 20)
				$strArgsOfTableArr[2] = "pd_type"
				GUICtrlCreateLabel("设计:", 20, 105, 30, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 52, 100, 180, 20)
				$strArgsOfTableArr[3] = "pd_desiger"
				GUICtrlCreateLabel("配置:", 260, 45, 30, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 292, 40, 180, 20)
				$strArgsOfTableArr[4] = "pd_configuration"
				GUICtrlCreateLabel("造价:", 260, 75, 30, 20 )
				$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 292, 70, 180, 20)
				$strArgsOfTableArr[5] = "pd_cost"
				GUICtrlCreateLabel("备注:", 260, 105, 30, 20 )
				$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 292, 100, 180, 20)
				$strArgsOfTableArr[6] = "pd_note"

			Case "学校信息"
				;学校 联系人 地址 电话 邮箱 备注 6
				;s_name, s_contact, s_address, s_phone, s_email, s_note
				$strArgsOfLVArr[0] = $lvInTabSchools	; LV
				$strArgsOfTableArr[0] = $tb_schools		; Table

				GUICtrlCreateLabel("学  校:", 20, 45, 45, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 67, 40, 170, 20)
				$strArgsOfTableArr[1] = "s_name"
				GUICtrlCreateLabel("联系人:", 20, 75, 45, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 67, 70, 170, 20)
				$strArgsOfTableArr[2] = "s_contact"
				GUICtrlCreateLabel("地  址:", 20, 105, 45, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 67, 100, 170, 20)
				$strArgsOfTableArr[3] = "s_address"

				GUICtrlCreateLabel("电  话:", 260, 45, 45, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 307, 40, 170, 20)
				$strArgsOfTableArr[4] = "s_phone"
				GUICtrlCreateLabel("邮  箱:", 260, 75, 45, 20 )
				$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 307, 70, 170, 20)
				$strArgsOfTableArr[5] = "s_email"
				GUICtrlCreateLabel("备  注:", 260, 105, 45, 20 )
				$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 307, 100, 170, 20)
				$strArgsOfTableArr[6] = "s_note"

			Case "工程管理"
				;名称 学校 产品 合作者 我司负责人 细则（照片） 起始日期 状态 结束日期 结算记录 备注 11
				;pj_name, pj_s_name, pj_pd_name, pj_pt_name, pj_w_name, pj_content, pj_date_start, pj_state, pj_date_finish, pj_account, pj_note
				$strArgsOfLVArr[0] = $lvInTabProjects	; LV
				$strArgsOfTableArr[0] = $tb_projects		; Table

				GUIDelete ( $popupWindow )

				$popWinHeight = 300

				$popupWindow = GUICreate("" , $popWinWidth, $popWinHeight, Default, Default, $WS_CAPTION, BitOR($WS_EX_TOPMOST, $GUI_WS_EX_PARENTDRAG))
					WinSetTitle($popupWindow, "", "在《" & $strLvName & "》中插入数据" )

					GUICtrlCreateGroup("", 10, 5, 480, $popWinHeight - 45)

					GUICtrlCreateLabel("名    称:", 20, 25, 60, 20 )
					$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 82, 20, 150, 20)
					$strArgsOfTableArr[1] = "pj_name"
					GUICtrlCreateLabel("学    校:", 20, 55, 60, 20 )
					$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 82, 50, 150, 20)
					$strArgsOfTableArr[2] = "pj_s_name"
					GUICtrlCreateLabel("产    品:", 20, 85, 60, 20 )
					$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 82, 80, 150, 20)
					$strArgsOfTableArr[3] = "pj_pd_name"
					GUICtrlCreateLabel("合 作 者:", 20, 115, 60, 20 )
					$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 82, 110, 150, 20)
					$strArgsOfTableArr[4] = "pj_pt_name"
					GUICtrlCreateLabel("我负责人:", 20, 145, 60, 20 )
					$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 82, 140, 150, 20)
					$strArgsOfTableArr[5] = "pj_w_name"

					GUICtrlCreateLabel("备    注:", 20, 175, 60, 20 )
					$strArgsOfLVArr[11] = GUICtrlCreateInput ( "", 82, 170, 390, 80, BitOR($ES_MULTILINE, $ES_WANTRETURN, $ES_AUTOVSCROLL))
					$strArgsOfTableArr[11] = "pj_note"

					GUICtrlCreateLabel("协议细则:", 260, 25, 60, 20 )
					$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 322, 20, 150, 20)
					$strArgsOfTableArr[6] = "pj_content"
					GUICtrlCreateLabel("开工日期:", 260, 55, 60, 20 )
					$strArgsOfLVArr[7] = GUICtrlCreateInput ( "", 322, 50, 150, 20)
					$strArgsOfTableArr[7] = "pj_date_start"
					GUICtrlCreateLabel("进展状态:", 260, 85, 60, 20 )
					$strArgsOfLVArr[8] = GUICtrlCreateInput ( "", 322, 80, 150, 20)
					$strArgsOfTableArr[8] = "pj_state"
					GUICtrlCreateLabel("完工日期:", 260, 115, 60, 20 )
					$strArgsOfLVArr[9] = GUICtrlCreateInput ( "", 322, 110, 150, 20)
					$strArgsOfTableArr[9] = "pj_date_finish"
					GUICtrlCreateLabel("结算记录:", 260, 145, 60, 20 )
					$strArgsOfLVArr[10] = GUICtrlCreateInput ( "", 322, 140, 150, 20)
					$strArgsOfTableArr[10] = "pj_account"

			Case "合作伙伴"
				;名称 类型 地址 电话 邮箱 业务 备注 7
				;pt_name, pt_type, pt_address, pt_phone, pt_email, pt_business, pt_note
				$strArgsOfLVArr[0] = $lvInTabPartners	; LV
				$strArgsOfTableArr[0] = $tb_partners	; Table

				GUICtrlCreateLabel("名称:", 20, 25, 30, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 52, 20, 180, 20)
				$strArgsOfTableArr[1] = "pt_name"
				GUICtrlCreateLabel("类型:", 20, 55, 30, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 52, 50, 180, 20)
				$strArgsOfTableArr[2] = "pt_type"
				GUICtrlCreateLabel("地址:", 20, 85, 30, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 52, 80, 180, 20)
				$strArgsOfTableArr[3] = "pt_address"

				GUICtrlCreateLabel("电话:", 260, 25, 30, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 292, 20, 180, 20)
				$strArgsOfTableArr[4] = "pt_phone"
				GUICtrlCreateLabel("邮箱:", 260, 55, 30, 20 )
				$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 292, 50, 180, 20)
				$strArgsOfTableArr[5] = "pt_email"
				GUICtrlCreateLabel("业务:", 260, 85, 30, 20 )
				$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 292, 80, 180, 20)
				$strArgsOfTableArr[6] = "pt_business"

				GUICtrlCreateLabel("备注:", 20, 115, 30, 20 )
				$strArgsOfLVArr[7] = GUICtrlCreateInput ( "", 52, 110, 420, 20)
				$strArgsOfTableArr[7] = "pt_note"

			Case "用户管理"
				; 用户名|密码|权限|备注 4
				; u_name, u_pswd, u_authority, u_note
				$strArgsOfLVArr[0] = $lvInTabUsers	; LV
				$strArgsOfTableArr[0] = $tb_users	; Table

				GUICtrlCreateLabel("用户:", 20, 25, 30, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 52, 20, 180, 20)
				$strArgsOfTableArr[1] = "u_name"
				GUICtrlCreateLabel("密码:", 20, 55, 30, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 52, 50, 180, 20)
				$strArgsOfTableArr[2] = "u_pswd"
				GUICtrlCreateLabel("权限:", 20, 85, 30, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 52, 80, 180, 20)
				$strArgsOfTableArr[3] = "u_authority"
				GUICtrlCreateLabel("备注:", 20, 115, 30, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 52, 110, 180, 20)
				$strArgsOfTableArr[4] = "u_note"


			Case "资产管理"
				; 名称|串号|单位|类型|购入日期|单价|所属部门|经销商|是否报废|备注 10
				; a_name, a_serial_number, a_unit, a_type, a_date_bought, a_price, a_deportment, a_dealer, a_scrap, a_note
				$strArgsOfLVArr[0] = $lvInTabAccets	; LV
				$strArgsOfTableArr[0] = $tb_accets		; Table

				GUIDelete ( $popupWindow )

				$popWinHeight = 210

				$popupWindow = GUICreate("" , $popWinWidth, $popWinHeight, Default, Default, $WS_CAPTION, BitOR($WS_EX_TOPMOST, $GUI_WS_EX_PARENTDRAG))
					WinSetTitle($popupWindow, "", "在《" & $strLvName & "》中插入数据" )

					GUICtrlCreateGroup("", 10, 5, 480, $popWinHeight - 45)

					GUICtrlCreateLabel("名    称:", 20, 25, 60, 20 )
					$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 82, 20, 150, 20)
					$strArgsOfTableArr[1] = "a_name"
					GUICtrlCreateLabel("串    号:", 20, 55, 60, 20 )
					$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 82, 50, 150, 20)
					$strArgsOfTableArr[2] = "a_serial_number"
					GUICtrlCreateLabel("单    位:", 20, 85, 60, 20 )
					$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 82, 80, 150, 20)
					$strArgsOfTableArr[3] = "a_unit"
					GUICtrlCreateLabel("类    型:", 20, 115, 60, 20 )
					$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 82, 110, 150, 20)
					$strArgsOfTableArr[4] = "a_type"
					GUICtrlCreateLabel("购入日期:", 20, 145, 60, 20 )
					$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 82, 140, 150, 20)
					$strArgsOfTableArr[5] = "a_date_bought"

					GUICtrlCreateLabel("单    价:", 260, 25, 60, 20 )
					$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 322, 20, 150, 20)
					$strArgsOfTableArr[6] = "a_price"
					GUICtrlCreateLabel("所属部门:", 260, 55, 60, 20 )
					$strArgsOfLVArr[7] = GUICtrlCreateInput ( "", 322, 50, 150, 20)
					$strArgsOfTableArr[7] = "a_deportment"
					GUICtrlCreateLabel("经 销 商:", 260, 85, 60, 20 )
					$strArgsOfLVArr[8] = GUICtrlCreateInput ( "", 322, 80, 150, 20)
					$strArgsOfTableArr[8] = "a_dealer"
					GUICtrlCreateLabel("是否报废:", 260, 115, 60, 20 )
					$strArgsOfLVArr[9] = GUICtrlCreateInput ( "", 322, 110, 150, 20)
					$strArgsOfTableArr[9] = "a_scrap"
					GUICtrlCreateLabel("备    注:", 260, 145, 60, 20 )
					$strArgsOfLVArr[10] = GUICtrlCreateInput ( "", 322, 140, 150, 20)
					$strArgsOfTableArr[10] = "a_note"

			Case "人员管理"
				; 姓名|身份证号|部门|职务|入职日期|转正日期|月薪资|性别|生日|电话|邮箱|住址|家庭信息|备注
				; w_name, w_identity_number, w_deportment, w_position, w_date_of_entry, w_date_regular, w_month_wage, w_sex, w_birthday, w_phone, w_email, w_address, w_family, w_note
				$strArgsOfLVArr[0] = $lvInTabWorkers
				$strArgsOfTableArr[0] = $tb_workers

				GUIDelete ( $popupWindow )

				$popWinHeight = 270

				$popupWindow = GUICreate("" , $popWinWidth, $popWinHeight, Default, Default, $WS_CAPTION, BitOR($WS_EX_TOPMOST, $GUI_WS_EX_PARENTDRAG))
					WinSetTitle($popupWindow, "", "在《" & $strLvName & "》中插入数据" )

					GUICtrlCreateGroup("", 10, 5, 480, $popWinHeight - 45)

					GUICtrlCreateLabel("姓    名:", 20, 25, 60, 20 )
					$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 82, 20, 150, 20)
					$strArgsOfTableArr[1] = "w_name"
					GUICtrlCreateLabel("身份证号:", 20, 55, 60, 20 )
					$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 82, 50, 150, 20)
					$strArgsOfTableArr[2] = "w_identity_number"
					GUICtrlCreateLabel("部    门:", 20, 85, 60, 20 )
					$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 82, 80, 150, 20)
					$strArgsOfTableArr[3] = "w_deportment"
					GUICtrlCreateLabel("职    务:", 20, 115, 60, 20 )
					$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 82, 110, 150, 20)
					$strArgsOfTableArr[4] = "w_position"
					GUICtrlCreateLabel("入职日期:", 20, 145, 60, 20 )
					$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 82, 140, 150, 20)
					$strArgsOfTableArr[5] = "w_date_of_entry"
					GUICtrlCreateLabel("转正日期:", 20, 175, 60, 20 )
					$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 82, 170, 150, 20)
					$strArgsOfTableArr[6] = "w_date_regular"
					GUICtrlCreateLabel("月 薪 资:", 20, 205, 60, 20 )
					$strArgsOfLVArr[7] = GUICtrlCreateInput ( "", 82, 200, 150, 20)
					$strArgsOfTableArr[7] = "w_month_wage"

					GUICtrlCreateLabel("性    别:", 260, 25, 60, 20 )
					$strArgsOfLVArr[8] = GUICtrlCreateInput ( "", 322, 20, 150, 20)
					$strArgsOfTableArr[8] = "w_sex"
					GUICtrlCreateLabel("生    日:", 260, 55, 60, 20 )
					$strArgsOfLVArr[9] = GUICtrlCreateInput ( "", 322, 50, 150, 20)
					$strArgsOfTableArr[9] = "w_birthday"
					GUICtrlCreateLabel("电    话:", 260, 85, 60, 20 )
					$strArgsOfLVArr[10] = GUICtrlCreateInput ( "", 322, 80, 150, 20)
					$strArgsOfTableArr[10] = "w_phone"
					GUICtrlCreateLabel("邮    箱:", 260, 115, 60, 20 )
					$strArgsOfLVArr[11] = GUICtrlCreateInput ( "", 322, 110, 150, 20)
					$strArgsOfTableArr[11] = "w_email"
					GUICtrlCreateLabel("住    址:", 260, 145, 60, 20 )
					$strArgsOfLVArr[12] = GUICtrlCreateInput ( "", 322, 140, 150, 20)
					$strArgsOfTableArr[12] = "w_address"
					GUICtrlCreateLabel("家庭信息:", 260, 175, 60, 20 )
					$strArgsOfLVArr[13] = GUICtrlCreateInput ( "", 322, 170, 150, 20)
					$strArgsOfTableArr[13] = "w_family"
					GUICtrlCreateLabel("备    注:", 260, 205, 60, 20 )
					$strArgsOfLVArr[14] = GUICtrlCreateInput ( "", 322, 200, 150, 20)
					$strArgsOfTableArr[14] = "w_note"

			Case "操作日志"
				; 用户|登陆退出增删改|表|列|旧数据|新数据|日期|备注 8
				; l_name, l_operate, l_table, l_column, l_old_data, l_new_data, l_date, l_note
				$strArgsOfLVArr[0] = $lvInTabLog
				$strArgsOfTableArr[0] = $tb_log

				GUICtrlCreateLabel("用户:", 20, 25, 30, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 52, 20, 180, 20)
				$strArgsOfTableArr[1] = "l_name"
				GUICtrlCreateLabel("操作:", 20, 55, 30, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 52, 50, 180, 20)
				$strArgsOfTableArr[2] = "l_operate"
				GUICtrlCreateLabel("表单:", 20, 85, 30, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 52, 80, 180, 20)
				$strArgsOfTableArr[3] = "l_table"
				GUICtrlCreateLabel("列名:", 20, 115, 30, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 52, 110, 180, 20)
				$strArgsOfTableArr[4] = "l_column"

				GUICtrlCreateLabel("旧值:", 260, 25, 30, 20 )
				$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 292, 20, 180, 20)
				$strArgsOfTableArr[5] = "l_old_data"
				GUICtrlCreateLabel("新值:", 260, 55, 30, 20 )
				$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 292, 50, 180, 20)
				$strArgsOfTableArr[6] = "l_new_data"
				GUICtrlCreateLabel("日期:", 260, 85, 30, 20 )
				$strArgsOfLVArr[7] = GUICtrlCreateInput ( "", 292, 80, 180, 20)
				$strArgsOfTableArr[7] = "l_date"
				GUICtrlCreateLabel("备注:", 260, 115, 30, 20 )
				$strArgsOfLVArr[8] = GUICtrlCreateInput ( "", 292, 110, 180, 20)
				$strArgsOfTableArr[8] = "l_note"

			Case "元 数 据"
				; 表单|列|值|描述 4
				; sr_tb_name, sr_column, sr_value, sr_note
				$strArgsOfLVArr[0] = $lvInTabSources
				$strArgsOfTableArr[0] = $tb_source

				GUICtrlCreateLabel("表单:", 20, 25, 30, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 52, 20, 180, 20)
				$strArgsOfTableArr[1] = "sr_tb_name"
				GUICtrlCreateLabel("列名:", 20, 55, 30, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 52, 50, 180, 20)
				$strArgsOfTableArr[2] = "sr_column"
				GUICtrlCreateLabel("选值:", 20, 85, 30, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 52, 80, 180, 20)
				$strArgsOfTableArr[3] = "sr_value"
				GUICtrlCreateLabel("备注:", 20, 115, 30, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 52, 110, 180, 20)
				$strArgsOfTableArr[4] = "sr_note"

	EndSwitch

	GUICtrlCreateGroup("", -99, -99, 1, 1)

	$btnOKInPopWin = GUICtrlCreateButton("确定", 10, $popWinHeight - 30, 60, 22)
		GUICtrlSetOnEvent($btnOKInPopWin, "FuncInsertDatas")
	$btnNOInPopWin = GUICtrlCreateButton("取消", 428, $popWinHeight - 30, 60, 22)
		GUICtrlSetOnEvent($btnNOInPopWin, "FuncInsertDatas")

	GUISetState(@SW_SHOW, $popupWindow)
	GUISetState(@SW_DISABLE, $guiMainWindow)
EndFunc

; #cs
; 向数据库中写数据条目
; #ce
Func FuncInsertDatas ()

	If @GUI_CtrlId = $btnOKInPopWin Then
		Local $strColumns	; 匹配操作的列
		Local $isSqlStrEnable = False	; 查询语句是否可用。如果全部参数都是空的，不会设为true
		Local $tmpStep	; 控制循环长度，匹配表列数
		Local $strTempLVName

		Switch $strArgsOfTableArr[0]	; 匹配数据库表
			Case $tb_products
				$strTempLVName = "公司产品"
				$tmpStep = 6
				$strColumns = " (pd_name, pd_type, pd_desiger, pd_configuration, pd_cost, pd_note) values ('"
			Case $tb_partners
				$strTempLVName = "合作伙伴"
				$tmpStep = 7
				$strColumns = " (pt_name, pt_type, pt_address, pt_phone, pt_email, pt_business, pt_note) values ('"
			Case $tb_projects
				$strTempLVName = "工程管理"
				$tmpStep = 11
				$strColumns = " (pj_name, pj_s_name, pj_pd_name, pj_pt_name, pj_w_name, pj_content, pj_date_start, pj_state, pj_date_finish, pj_account, pj_note) values ('"
			Case $tb_journal
				$strTempLVName = "工作日志"
				$tmpStep = 7
				$strColumns = " (j_name, j_date, j_address, j_traffic, j_board, j_content, j_note) values ('"	; , j_record, j_date_record
			Case $tb_schools
				$strTempLVName = "学校信息"
				$tmpStep = 6
				$strColumns = " (s_name, s_contact, s_address, s_phone, s_email, s_note) values ('"

			Case $tb_users
				$strTempLVName = "用户管理"
				$tmpStep = 4
				$strColumns = " (u_name, u_pswd, u_authority, u_note) values ('"

			Case $tb_accets
				$strTempLVName = "资产管理"
				$tmpStep = 10
				$strColumns = " (a_name, a_serial_number, a_unit, a_type, a_date_bought, a_price, a_deportment, a_dealer, a_scrap, a_note) values ('"

			Case $tb_workers
				$strTempLVName = "人员管理"
				$tmpStep = 14
				$strColumns = " (w_name, w_identity_number, w_deportment, w_position, w_date_of_entry, w_date_regular, w_month_wage, w_sex, w_birthday, w_phone, w_email, w_address, w_family, w_note) values ('"

			Case $tb_log
				$strTempLVName = "操作日志"
				$tmpStep = 8
				$strColumns = " (l_name, l_operate, l_table, l_column, l_old_data, l_new_data, l_date, l_note) values ('"

			Case $tb_source
				$strTempLVName = "元 数 据"
				$tmpStep = 4
				$strColumns = " (sr_tb_name, sr_column, sr_value, sr_note) values ('"
				;------------------------------------------------------ todo --------------------------------------
		EndSwitch

		$strArgsOfSql = ""	; 数据拼接

		; 遍历控件数组，拼接输入的数据
		For $iiii = 1 To $tmpStep Step 1	; 0索引保存的是lV或Access名称。剩余 1-11 保存控件

			Local $strWidget = GUICtrlRead($strArgsOfLVArr[$iiii])

			If $strWidget <> "" And StringIsSpace($strWidget) <> 1 Then	; 如果输入非全空格的数据

				; 读控件内的文本，去除首、尾空格、连续多个空格合为1个
				Local $keywords = StringStripWS($strWidget, $STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES )
				$strArgsOfSql &= $keywords & "', '"

				$isSqlStrEnable = True	; 至少有一个有效参数，才可用

			ElseIf $strWidget = "" Or StringIsSpace($strWidget) = 1 Then ; 如果输入框留空，或仅输入的空格
				$strArgsOfSql &= " " & "', '"

			EndIf
		Next

		$strArgsOfSql = StringTrimRight ( $strArgsOfSql, 4 )	; 清除最有一个逗号和空格和一对单引号 ', '

		If $isSqlStrEnable Then

			Local $dataAdodbConnectionIn = ObjCreate("ADODB.Connection")
			$dataAdodbConnectionIn.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & $db_path & ";Jet Oledb:Database Password='" & $db_pswd & "'")
			Local $strSqlInsert = "insert into " & $strArgsOfTableArr[0] & $strColumns & $strArgsOfSql & "')"
			$dataAdodbConnectionIn.Execute($strSqlInsert)

			$strArgsOfSql = StringReplace ( $strArgsOfSql, "', '", "|" )

			; 日志表
			Local $strTmpSqlInsert = "'Eminem', 'add', '" & $strArgsOfTableArr[0] & "', '_item', '_none', '" & $strArgsOfSql & "', '" & @YEAR & "-" & @MON & "-" & @MDAY & "(" & @WDAY - 1 & ")" & @HOUR & ":" & @MIN & ":" & @SEC & "`" & @MSEC & "', '_none'"	; 数据
			$strTmpSqlInsert = "insert into " & $tb_log & " (l_name, l_operate, l_table, l_column, l_old_data, l_new_data, l_date, l_note) values ( " & $strTmpSqlInsert & " )"	; 插入语句
			$dataAdodbConnectionIn.Execute($strTmpSqlInsert)

			Local $adoRec = ObjCreate("ADODB.Recordset")
			$adoRec.ActiveConnection = $dataAdodbConnectionIn
			$adoRec.open("SELECT TOP 1 id FROM " & $strArgsOfTableArr[0] & " ORDER BY id DESC")	; 查询最大的id。就算已经删除了记录，ID仍然保持增长
			Local $recordcount = $adoRec.fields(0).value	; 只有一条记录，所以直接取。这个ID值已经 + 1

			$adoRec.Close
			$dataAdodbConnectionIn.Close

			; 这里还要添加一个“编号”列到LV
			; 这个编号是数据库中最大的ID+1。这里得到的值已经+1
			$strArgsOfSql = $recordcount & "|" & $strArgsOfSql

			Local $tmpLvItem = GUICtrlCreateListViewItem( $strArgsOfSql, $strArgsOfLVArr[0])
			GUICtrlSetBkColor ($tmpLvItem, 0xff9900 )

			_GUICtrlListView_Scroll($strArgsOfLVArr[0], 0, _GUICtrlListView_GetItemCount($strArgsOfLVArr[0])*15)

			GUICtrlSetData($strInfoLbl, "列表《" & $strTempLVName & "》中记录数量：" & _GUICtrlListView_GetItemCount ( $strArgsOfLVArr[0] ))
		EndIf
	EndIf

	GUIDelete ( $popupWindow )
	GUISetState(@SW_ENABLE, $guiMainWindow)
	GUISetState(@SW_RESTORE, $guiMainWindow)

	; 复位。清空数组
	$strArgsOfSql = ""
	For $i = 0 To 14
		$strArgsOfLVArr[$i] = ""
		$strArgsOfTableArr[$i] = ""
	Next
EndFunc

; #cs
; 右键点击标签卡弹出菜单
; 对右键菜单项的响应
; #ce
Func Func_MouseMenuItem()
	Switch @GUI_CtrlId
		Case $mouseMenuItemClose

			GUICtrlSetState($lvInTabJournal, $GUI_HIDE)
			GUICtrlSetState($lvInTabProducts, $GUI_HIDE)
			GUICtrlSetState($lvInTabSchools, $GUI_HIDE)
			GUICtrlSetState($lvInTabProjects, $GUI_HIDE)
			GUICtrlSetState($lvInTabPartners, $GUI_HIDE)

			GUICtrlSetState($lvInTabUsers, $GUI_HIDE)
			GUICtrlSetState($lvInTabAccets, $GUI_HIDE)
			GUICtrlSetState($lvInTabWorkers, $GUI_HIDE)
			GUICtrlSetState($lvInTabLog, $GUI_HIDE)
			GUICtrlSetState($lvInTabSources, $GUI_HIDE)

			If $idOfTabItem = 0 Then
				_GUICtrlTab_SetCurSel($tabInMainWindow, 1)
			Else
				_GUICtrlTab_SetCurSel($tabInMainWindow, $idOfTabItem - 1)
			EndIf

			Switch _GUICtrlTab_GetItemText($tabInMainWindow, $idOfTabItem)
				Case "工作日志"
					GUICtrlSetData($menuItemJournal, "　工作日志")
				Case "公司产品"
					GUICtrlSetData($menuItemProducts, "　公司产品")
				Case "学校信息"
					GUICtrlSetData($menuItemSchools, "　学校信息")
				Case "工程管理"
					GUICtrlSetData($menuItemProjects, "　工程管理")
				Case "合作伙伴"
					GUICtrlSetData($menuItemPartners, "　合作伙伴")

				Case "用户管理"
					GUICtrlSetData($menuItemUsers, "  用户管理")
				Case "资产管理"
					GUICtrlSetData($menuItemAccets, "  资产管理")
				Case "人员管理"
					GUICtrlSetData($menuItemWorkers, "  人员管理")
				Case "操作日志"
					GUICtrlSetData($menuItemLog, "  操作日志")
				Case "元 数 据"
					GUICtrlSetData($menuItemSources, "  元 数 据")
			EndSwitch

			_GUICtrlTab_DeleteItem ( $tabInMainWindow, $idOfTabItem )

			$strTabItemText = _GUICtrlTab_GetItemText ( $tabInMainWindow, _GUICtrlTab_GetCurSel ( $tabInMainWindow ) )

			Switch $strTabItemText
				Case "工作日志"
					GUICtrlSetState($lvInTabJournal, $GUI_SHOW)
					GUICtrlSetData($strInfoLbl, "列表《工作日志》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabJournal ))
				Case "公司产品"
					GUICtrlSetState($lvInTabProducts, $GUI_SHOW)
					GUICtrlSetData($strInfoLbl, "列表《公司产品》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabProducts ))
				Case "学校信息"
					GUICtrlSetState($lvInTabSchools, $GUI_SHOW)
					GUICtrlSetData($strInfoLbl, "列表《学校信息》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabSchools ))
				Case "工程管理"
					GUICtrlSetState($lvInTabProjects, $GUI_SHOW)
					GUICtrlSetData($strInfoLbl, "列表《工程管理》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabProjects ))
				Case "合作伙伴"
					GUICtrlSetState($lvInTabPartners, $GUI_SHOW)
					GUICtrlSetData($strInfoLbl, "列表《合作伙伴》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabPartners ))

				Case "用户管理"
					GUICtrlSetState($lvInTabUsers, $GUI_SHOW)
					GUICtrlSetData($strInfoLbl, "列表《用户管理》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabUsers ))
				Case "资产管理"
					GUICtrlSetState($lvInTabAccets, $GUI_SHOW)
					GUICtrlSetData($strInfoLbl, "列表《资产管理》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabAccets ))
				Case "人员管理"
					GUICtrlSetState($lvInTabWorkers, $GUI_SHOW)
					GUICtrlSetData($strInfoLbl, "列表《人员管理》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabWorkers ))
				Case "操作日志"
					GUICtrlSetState($lvInTabLog, $GUI_SHOW)
					GUICtrlSetData($strInfoLbl, "列表《操作日志》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabLog ))
				Case "元 数 据"
					GUICtrlSetState($lvInTabSources, $GUI_SHOW)
					GUICtrlSetData($strInfoLbl, "列表《元 数 据》中记录数量：" & _GUICtrlListView_GetItemCount ( $lvInTabSources ))
			EndSwitch

			$idOfTabItem = -1

		Case $mouseMenuItemSaveAs

		Case $mouseMenuItemPrint

	EndSwitch
EndFunc

#cs
	通过菜单栏控制标签的显示和隐藏
#ce
Func Func_ShowTab_ByMenu ()
	Switch @GUI_CtrlId
		Case $menuItemJournal
			Func_ShowTab_ByText ("工作日志")
		Case $menuItemProducts
			Func_ShowTab_ByText ("公司产品")
		Case $menuItemSchools
			Func_ShowTab_ByText ("学校信息")
		Case $menuItemProjects
			Func_ShowTab_ByText ("工程管理")
		Case $menuItemPartners
			Func_ShowTab_ByText ("合作伙伴")
		Case $menuItemUsers
			Func_ShowTab_ByText ("用户管理")
		Case $menuItemAccets
			Func_ShowTab_ByText ("资产管理")
		Case $menuItemWorkers
			Func_ShowTab_ByText ("人员管理")
		Case $menuItemLog
			Func_ShowTab_ByText ("操作日志")
		Case $menuItemSources
			Func_ShowTab_ByText ("元 数 据")
	EndSwitch
EndFunc

#cs 通过传入的字符串，和标签头进行匹配。控制显示或隐藏
#ce
Func Func_ShowTab_ByText ( $strItemName )

	GUICtrlSetState($lvInTabJournal, $GUI_HIDE)
	GUICtrlSetState($lvInTabProducts, $GUI_HIDE)
	GUICtrlSetState($lvInTabSchools, $GUI_HIDE)
	GUICtrlSetState($lvInTabProjects, $GUI_HIDE)
	GUICtrlSetState($lvInTabPartners, $GUI_HIDE)

	GUICtrlSetState($lvInTabUsers, $GUI_HIDE)
	GUICtrlSetState($lvInTabAccets, $GUI_HIDE)
	GUICtrlSetState($lvInTabWorkers, $GUI_HIDE)
	GUICtrlSetState($lvInTabLog, $GUI_HIDE)
	GUICtrlSetState($lvInTabSources, $GUI_HIDE)

	$id_item = _GUICtrlTab_FindTab ( $tabInMainWindow, $strItemName, True)

	If $id_item <> -1 Then

		If _GUICtrlTab_GetItemState ( $tabInMainWindow, $id_item ) = $TCIS_BUTTONPRESSED Then
			If $id_item = 0  Then
				_GUICtrlTab_SetCurSel($tabInMainWindow, 1)
			Else
				_GUICtrlTab_SetCurSel($tabInMainWindow, 0)
			EndIf
		EndIf

		Switch _GUICtrlTab_GetItemText($tabInMainWindow, $id_item)
			Case "工作日志"
				GUICtrlSetData($menuItemJournal, "  工作日志")
			Case "公司产品"
				GUICtrlSetData($menuItemProducts, "  公司产品")
			Case "学校信息"
				GUICtrlSetData($menuItemSchools, "  学校信息")
			Case "工程管理"
				GUICtrlSetData($menuItemProjects, "  工程管理")
			Case "合作伙伴"
				GUICtrlSetData($menuItemPartners, "  合作伙伴")

			Case "用户管理"
				GUICtrlSetData($menuItemUsers, "  用户管理")
			Case "资产管理"
				GUICtrlSetData($menuItemAccets, "  资产管理")
			Case "人员管理"
				GUICtrlSetData($menuItemWorkers, "  人员管理")
			Case "操作日志"
				GUICtrlSetData($menuItemLog, "  操作日志")
			Case "元 数 据"
				GUICtrlSetData($menuItemSources, "  元 数 据")
		EndSwitch

		_GUICtrlTab_DeleteItem ( $tabInMainWindow, $id_item )

		$strTabItemText = _GUICtrlTab_GetItemText ( $tabInMainWindow, _GUICtrlTab_GetCurSel ( $tabInMainWindow ) )
		Switch $strTabItemText
			Case "工作日志"
				If $hasReadedDbTbJournal = False Then
					FuncReadDb($tb_journal, $lvInTabJournal)
					$hasReadedDbTbJournal = True
				EndIf

				GUICtrlSetState($lvInTabJournal, $GUI_SHOW)
				GUICtrlSetData($menuItemJournal, "√ 工作日志")
			Case "公司产品"
				If $hasReadedDbTbProducts = False Then
					FuncReadDb($tb_products, $lvInTabProducts)
					$hasReadedDbTbProducts = True
				EndIf

				GUICtrlSetState($lvInTabProducts, $GUI_SHOW)
				GUICtrlSetData($menuItemProducts, "√ 公司产品")
			Case "学校信息"
				If $hasReadedDbTbSchools = False Then
					FuncReadDb($tb_schools, $lvInTabSchools)
					$hasReadedDbTbSchools = True
				EndIf

				GUICtrlSetState($lvInTabSchools, $GUI_SHOW)
				GUICtrlSetData($menuItemSchools, "√ 学校信息")
			Case "工程管理"
				If $hasReadedDbTbProjects = False Then
					FuncReadDb($tb_projects, $lvInTabProjects)
					$hasReadedDbTbProjects = True
				EndIf

				GUICtrlSetState($lvInTabProjects, $GUI_SHOW)
				GUICtrlSetData($menuItemProjects, "√ 工程管理")
			Case "合作伙伴"
				If $hasReadedDbTbPartners = False Then
					FuncReadDb($tb_partners, $lvInTabPartners)
					$hasReadedDbTbPartners = True
				EndIf

				GUICtrlSetState($lvInTabPartners, $GUI_SHOW)
				GUICtrlSetData($menuItemPartners, "√ 合作伙伴")

			Case "用户管理"
				If $hasReadedDbTbUsers= False Then
					FuncReadDb($tb_users, $lvInTabUsers)
					$hasReadedDbTbUsers = True
				EndIf

				GUICtrlSetState($lvInTabUsers, $GUI_SHOW)
				GUICtrlSetData($menuItemUsers, "√ 用户管理")
			Case "资产管理"
				If $hasReadedDbTbAccets = False Then
					FuncReadDb($tb_accets, $lvInTabAccets)
					$hasReadedDbTbAccets = True
				EndIf

				GUICtrlSetState($lvInTabAccets, $GUI_SHOW)
				GUICtrlSetData($menuItemAccets, "√ 资产管理")
			Case "人员管理"
				If $hasReadedDbTbWorkers = False Then
					FuncReadDb($tb_workers, $lvInTabWorkers)
					$hasReadedDbTbWorkers = True
				EndIf

				GUICtrlSetState($lvInTabWorkers, $GUI_SHOW)
				GUICtrlSetData($menuItemWorkers, "√ 人员管理")
			Case "操作日志"
				If $hasReadedDbTbLog = False Then
					FuncReadDb($tb_log, $lvInTabLog)
					$hasReadedDbTbLog = True
				EndIf

				GUICtrlSetState($lvInTabLog, $GUI_SHOW)
				GUICtrlSetData($menuItemLog, "√ 操作日志")
			Case "元 数 据"
				If $hasReadedDbTbSource = False Then
					FuncReadDb($tb_source, $lvInTabSources)
					$hasReadedDbTbSource = True
				EndIf

				GUICtrlSetState($lvInTabSources, $GUI_SHOW)
				GUICtrlSetData($menuItemSources, "√ 元 数 据")
		EndSwitch

	Else
		Switch $strItemName
			Case "工作日志"
				$id_item = 0
			Case "公司产品"
				$id_item = 1
			Case "学校信息"
				$id_item = 2
			Case "工程管理"
				$id_item = 3
			Case "合作伙伴"
				$id_item = 4

			Case "用户管理"
				$id_item = 5
			Case "资产管理"
				$id_item = 6
			Case "人员管理"
				$id_item = 7
			Case "操作日志"
				$id_item = 8
			Case "元 数 据"
				$id_item = 9
		EndSwitch

		_GUICtrlTab_InsertItem ( $tabInMainWindow, $id_item, $strItemName, $id_item)

		For $i = 0 To _GUICtrlTab_GetItemCount ( $tabInMainWindow ) - 1 Step 1
			If $strItemName = _GUICtrlTab_GetItemText ( $tabInMainWindow, $i ) Then
				_GUICtrlTab_SetCurSel($tabInMainWindow, $i)

				GUICtrlSetState($lvInTabJournal, $GUI_HIDE)
				GUICtrlSetState($lvInTabProducts, $GUI_HIDE)
				GUICtrlSetState($lvInTabSchools, $GUI_HIDE)
				GUICtrlSetState($lvInTabProjects, $GUI_HIDE)
				GUICtrlSetState($lvInTabPartners, $GUI_HIDE)

				GUICtrlSetState($lvInTabUsers, $GUI_HIDE)
				GUICtrlSetState($lvInTabAccets, $GUI_HIDE)
				GUICtrlSetState($lvInTabWorkers, $GUI_HIDE)
				GUICtrlSetState($lvInTabLog, $GUI_HIDE)
				GUICtrlSetState($lvInTabSources, $GUI_HIDE)

				Switch $strItemName
					Case "工作日志"
						If $hasReadedDbTbJournal = False Then
							FuncReadDb($tb_journal, $lvInTabJournal)
							$hasReadedDbTbJournal = True
						EndIf

						GUICtrlSetState($lvInTabJournal, $GUI_SHOW)
						GUICtrlSetData($menuItemJournal, "√ 工作日志")
					Case "公司产品"
						If $hasReadedDbTbProducts = False Then
							FuncReadDb($tb_products, $lvInTabProducts)
							$hasReadedDbTbProducts = True
						EndIf

						GUICtrlSetState($lvInTabProducts, $GUI_SHOW)
						GUICtrlSetData($menuItemProducts, "√ 公司产品")
					Case "学校信息"
						If $hasReadedDbTbSchools = False Then
							FuncReadDb($tb_schools, $lvInTabSchools)
							$hasReadedDbTbSchools = True
						EndIf

						GUICtrlSetState($lvInTabSchools, $GUI_SHOW)
						GUICtrlSetData($menuItemSchools, "√ 学校信息")
					Case "工程管理"
						If $hasReadedDbTbProjects = False Then
							FuncReadDb($tb_projects, $lvInTabProjects)
							$hasReadedDbTbProjects = True
						EndIf

						GUICtrlSetState($lvInTabProjects, $GUI_SHOW)
						GUICtrlSetData($menuItemProjects, "√ 工程管理")
					Case "合作伙伴"
						If $hasReadedDbTbPartners = False Then
							FuncReadDb($tb_partners, $lvInTabPartners)
							$hasReadedDbTbPartners = True
						EndIf

						GUICtrlSetState($lvInTabPartners, $GUI_SHOW)
						GUICtrlSetData($menuItemPartners, "√ 合作伙伴")

					Case "用户管理"
						If $hasReadedDbTbUsers= False Then
							FuncReadDb($tb_users, $lvInTabUsers)
							$hasReadedDbTbUsers = True
						EndIf

						GUICtrlSetState($lvInTabUsers, $GUI_SHOW)
						GUICtrlSetData($menuItemUsers, "√ 用户管理")
					Case "资产管理"
						If $hasReadedDbTbAccets = False Then
							FuncReadDb($tb_accets, $lvInTabAccets)
							$hasReadedDbTbAccets = True
						EndIf

						GUICtrlSetState($lvInTabAccets, $GUI_SHOW)
						GUICtrlSetData($menuItemAccets, "√ 资产管理")
					Case "人员管理"
						If $hasReadedDbTbWorkers = False Then
							FuncReadDb($tb_workers, $lvInTabWorkers)
							$hasReadedDbTbWorkers = True
						EndIf

						GUICtrlSetState($lvInTabWorkers, $GUI_SHOW)
						GUICtrlSetData($menuItemWorkers, "√ 人员管理")
					Case "操作日志"
						If $hasReadedDbTbLog = False Then
							FuncReadDb($tb_log, $lvInTabLog)
							$hasReadedDbTbLog = True
						EndIf

						GUICtrlSetState($lvInTabLog, $GUI_SHOW)
						GUICtrlSetData($menuItemLog, "√ 操作日志")
					Case "元 数 据"
						If $hasReadedDbTbSource = False Then
							FuncReadDb($tb_source, $lvInTabSources)
							$hasReadedDbTbSource = True
						EndIf

						GUICtrlSetState($lvInTabSources, $GUI_SHOW)
						GUICtrlSetData($menuItemSources, "√ 元 数 据")
				EndSwitch
			EndIf
		Next
	EndIf
EndFunc

Func Func_MenuHelp ()
EndFunc

;#cs 读取表的数据，显示到ListView上
;	$strTbName：要查询的表名称
;	$strListView：要显示数据的列表
;#ce
Func FuncReadDb( $strTbName, $strListView )

	Local $obj_adodb_connection = ObjCreate("ADODB.Connection")
	$obj_adodb_connection.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & $db_path & "; Jet OLEDB:Database Password='" & $db_pswd & "'")
	Local $obj_adodb_recordset = ObjCreate("ADODB.Recordset")
	$obj_adodb_recordset.ActiveConnection = $obj_adodb_connection

	Local $oRec = $obj_adodb_recordset
	$oRec.Open($strTbName, $obj_adodb_connection, 3, 3)
	Local $fieldsCount = $oRec.fields.count	; 获取表单的列数，总共有多少字段

	$obj_adodb_recordset.Open("SELECT * FROM " & $strTbName)
	;
	; 工作日志表特殊：最后两个字段隐藏
	;
	If $strTbName = $tb_journal Then
		$fieldsCount = $fieldsCount - 2
	EndIf

	_GUICtrlListView_BeginUpdate($strListView)
	; 遍历 行
	While Not $obj_adodb_recordset.eof And Not $obj_adodb_recordset.bof
		If @error = 1 Then ExitLoop
		;
		Local $strFields = ""
		For $i = 0 To $fieldsCount - 1 Step 1

			Local $strTmpS = $obj_adodb_recordset.Fields( $i ).value
			If $strTbName = "tb_log" Then
				$strTmpS = StringReplace ( $strTmpS, "|", "/")
			EndIf

			$strFields &= $strTmpS & "|"
		Next
		$strFields = StringTrimRight ( $strFields, 1 )	; 删除末尾的 |

		; 呈现 列
		GUICtrlCreateListViewItem( $strFields, $strListView)
			GUICtrlSetBkColor (-1, 0xffa500 );设置listviewitem的背景色
		$obj_adodb_recordset.movenext
	WEnd
	_GUICtrlListView_EndUpdate($strListView)

	GUICtrlSetData($strInfoLbl, "表《" & $strTbName & "》中记录数量：" & $obj_adodb_recordset.recordcount)

	$obj_adodb_recordset.close
	$obj_adodb_connection.Close
	$oRec.close

EndFunc

;#cs 创建数据库
;	$strDbPath：文件路径
;	$strDbPswd：访问密码
;#ce
Func FuncCreateDb()
	; 表单的列和ListView中对应

	;=== 公司人员表 tb_workers ：14
	;                              编号。主键自动增长|                   姓名。文本|  身份证号|               部门|              职务|            入职日期|             转正日期|            月薪资|            性别|       生日|            电话|         邮箱|         住址，默认50字符|    家庭信息|          备注
	Local $strColsForTbWorders = "id integer identity(1,1) primary key, w_name text, w_identity_number text, w_deportment text, w_position text, w_date_of_entry text, w_date_regular text, w_month_wage text, w_sex text, w_birthday text, w_phone text, w_email text, w_address text(250), w_family text(254), w_note memo"

	;=== 学校信息表 tb_schools ：6
	;                              编号|                                 学校|        联系人|         地址|                电话|         邮箱|        备注
	Local $strColsForTbSchools = "id integer identity(1,1) primary key, s_name text, s_contact text, s_address text(250), s_phone text, s_email text, s_note memo"

	;=== 公司产品表 tb_products ：6
	;                               编号|                                 产品|         类型|        设计|             配置|                       造价|        备注
	Local $strColsForTbProducts = "id integer identity(1,1) primary key, pd_name text, pd_type text, pd_desiger text, pd_configuration text(250), pd_cost text, pd_note memo"

	;=== 资产管理表 tb_accets ：10
	;                             编号|                                 名称|        串号|                 单位|        类型|        购入日期|           单价|         所属部门|          经销商|        是否报废|     备注
	Local $strColsForTbAccets = "id integer identity(1,1) primary key, a_name text, a_serial_number text, a_unit text, a_type text, a_date_bought text, a_price text, a_deportment text, a_dealer text, a_scrap text, a_note text(250)"

	;=== 合作伙伴表 tb_partners ：7
	;                               编号|                                 名称|         类型|         地址|                 电话|          邮箱|          业务|             备注
	Local $strColsForTbPartners = "id integer identity(1,1) primary key, pt_name text, pt_type text, pt_address text(250), pt_phone text, pt_email text, pt_business text, pt_note memo"

	;=== 工程管理表 tb_projects ：10
	;                               编号|                                 名称|         学校|           产品|            合作者|          我司负责人|     细则（照片）|         起始日期|           状态|          结束日期|            结算记录|             备注
	Local $strColsForTbProjects = "id integer identity(1,1) primary key, pj_name text, pj_s_name text, pj_pd_name text, pj_pt_name text, pj_w_name text, pj_content text(250), pj_date_start text, pj_state text, pj_date_finish text, pj_account text(250), pj_note memo"

	;=== 工作日志表 tb_journal ：7
	;                              编号|                                 人员|        日期|        地点|                交通|                食宿|              工作描述|            备注|        记录人(隐藏)|        记录日期(隐藏)
	Local $strColsForTbJournal = "id integer identity(1,1) primary key, j_name text, j_date text, j_address text(250), j_traffic text(250), j_board text(250), j_content text(250), j_note memo, j_record text, j_date_record text"

	;=== 用户表 tb_users ：4
	;                            编号|                                 用户名|      密码|        权限|             备注
	Local $strColsForTbUsers = "id integer identity(1,1) primary key, u_name text, u_pswd text, u_authority text, u_note text(250)"

	;=== 源数据表 tb_source ：4
	;                             编号|                                 表单|            列|             值|            描述
	Local $strColsForTbSource = "id integer identity(1,1) primary key, sr_tb_name text, sr_column text, sr_value text, sr_note text(250)"

	;=== 操作日志表 tb_log ：
	;                          编号|                                 用户|        登退增删改|     表|            列|            旧数据|          新数据|          日期|        备注
	Local $strColsForTbLog = "id integer identity(1,1) primary key, l_name text, l_operate text, l_table text, l_column text,  l_old_data memo, l_new_data memo, l_date text, l_note text(250)"

	; 创建数据库文件
	Local $oADO = ObjCreate("ADOX.Catalog")
	$oADO.Create("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & $db_path & "; Jet OLEDB:Database Password='" & $db_pswd & "'")
	$oADO.ActiveConnection.Close

	; 创建表单 $tb_workers, $tb_schools, $tb_products, $tb_accets, $tb_partners, $tb_projects, $tb_journal, $tb_users, $tb_source, $tb_log
	$oADO = ObjCreate("ADODB.Connection")
	$oADO.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & $db_path & "; Jet OLEDB:Database Password='" & $db_pswd & "'")
	$oADO.Execute("CREATE TABLE " & $tb_workers)	; 人员管理'
	$oADO.Execute("CREATE TABLE " & $tb_schools)	; 学校信息
	$oADO.Execute("CREATE TABLE " & $tb_products)	; 公司产品
	$oADO.Execute("CREATE TABLE " & $tb_accets)		; 资产管理
	$oADO.Execute("CREATE TABLE " & $tb_partners)	; 合作伙伴
	$oADO.Execute("CREATE TABLE " & $tb_projects)	; 工程管理
	$oADO.Execute("CREATE TABLE " & $tb_journal)	; 主界面，工作日志
	$oADO.Execute("CREATE TABLE " & $tb_users)		; 用户管理
	$oADO.Execute("CREATE TABLE " & $tb_source)		; 元数据
	$oADO.Execute("CREATE TABLE " & $tb_log)		; 操作日志

	$oADO.Execute("ALTER TABLE " & $tb_workers & " ADD " & $strColsForTbWorders)
	$oADO.Execute("ALTER TABLE " & $tb_schools  & " ADD " & $strColsForTbSchools)
	$oADO.Execute("ALTER TABLE " & $tb_products  & " ADD " & $strColsForTbProducts)
	$oADO.Execute("ALTER TABLE " & $tb_accets & " ADD " & $strColsForTbAccets)
	$oADO.Execute("ALTER TABLE " & $tb_partners & " ADD " & $strColsForTbPartners)
	$oADO.Execute("ALTER TABLE " & $tb_projects & " ADD " & $strColsForTbProjects)
	$oADO.Execute("ALTER TABLE " & $tb_journal & " ADD " & $strColsForTbJournal)
	$oADO.Execute("ALTER TABLE " & $tb_users & " ADD " & $strColsForTbUsers)
	$oADO.Execute("ALTER TABLE " & $tb_source & " ADD " & $strColsForTbSource)
	$oADO.Execute("ALTER TABLE " & $tb_log & " ADD " & $strColsForTbLog)
	$oADO.Close
EndFunc
