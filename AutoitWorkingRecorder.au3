; www.cuiweiyou.com
; 威格灵|vigiles
; http://www.cuiweiyou.com/tag/autoit%E4%B8%8A%E8%B7%AF

#AutoIt3Wrapper_UseX64=n		; 64位win7

#include <GUIConstantsEx.au3>
#include <GUIEdit.au3>
#include <GuiImageList.au3>
#include <GuiListView.au3>
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
	If MsgBox(52, "警告", "未检测到数据库文件，是否重新创建？") = 6 Then
		FuncCreateDb ( )
	EndIf
EndIf

;======================================================================
;-------------------- 用户登录 --------------------------------


;======================================================================
; 单击标签切换LV时用到的
Global $itemInToolbar, $idOfTabItem	; 添加一个全局变量$idOfTabItem记录被右击的标签索引
; 主窗体宽高
Global $WidthOfWindow = 800, $HeightOfWindow = 600
; 主窗体，工具条，工具条提示器
Global $guiMainWindow, $toolbarInMainWindow, $hToolTip 
; 工具条上的按钮既定ID
Global Enum $id_Toolbar_New = 1000, $id_Toolbar_Save, $id_Toolbar_Delete, $id_Toolbar_Find, $id_Toolbar_Help
; 数据库读取标识。读过的True
Global $hasReadedDbTbWorkers = False, $hasReadedDbTbSchools = False, $hasReadedDbTbProducts = False, $hasReadedDbTbAccets = False, $hasReadedDbTbPartners = False, $hasReadedDbTbProjects = False, $hasReadedDbTbUsers = False, $hasReadedDbTbSource = False, $hasReadedDbTbLog = False, $hasReadedDbTbJournal = False
; 双击单元格时创建编辑框用到的
Global $hEdit, $hBrush, $hDC, $hItemRow, $hItemColumn	; 双击LV进行修改用到了
; 双击单元格修改数据，保存日志到Log表时用到的。修改的数据属于哪个表，属于哪个LV，对应的列名，修改前数据
Global $columnOfTable, $columnOfLV, $columnName, $columnOldData
; 子窗体，子窗体的确定按钮，取消按钮。
Global $popupWindow, $btnOKInPopWin, $btnNOInPopWin
; 查询时保存查询条件
Global $strArgsOfQuery, $strArgsOfLVArr[7], $strArgsOfTableArr[7]	; 根据本例需要，长度设为7

;======================================================================
;-------------------- GUI ---------------------------------
$guiMainWindow = GUICreate("中徽教育|威格灵 www.cuiweiyou.com", $WidthOfWindow, $HeightOfWindow) 
	GUISetOnEvent($GUI_EVENT_CLOSE, "Func_GUI_EVENT_CLOSE")
	GUISetIcon( @ScriptDir & "\Logo.ico")	; 设置程序图标为脚本文件同目录中的Logo.ico
	
	$menuFile = GUICtrlCreateMenu ( "文件 &F")
		$itemOpenInMenuFile = GUICtrlCreateMenuItem("打开", $menuFile)
		$itemSaveInMenuFile = GUICtrlCreateMenuItem("保存", $menuFile)
		GUICtrlCreateMenuItem("", $menuFile) ; 分隔线
		$itemRecentfilesInMenuFile = GUICtrlCreateMenu("最近的文件", $menuFile)
		GUICtrlCreateMenuItem("", $menuFile) 
		$itemExitInMenuFile = GUICtrlCreateMenuItem("退出", $menuFile)
			GUICtrlSetOnEvent($itemExitInMenuFile, "Func_GUI_EVENT_CLOSE")
	
	$menuTab = GUICtrlCreateMenu ( "窗口 &W")
		$itemJournalOfTabInMenu = GUICtrlCreateMenuItem("√ 工作日志", $menuTab)
			GUICtrlSetOnEvent($itemJournalOfTabInMenu, "Func_ShowTab_ByMenu")
		$itemProductsOfTabInMenu = GUICtrlCreateMenuItem("  公司产品", $menuTab)
			GUICtrlSetOnEvent($itemProductsOfTabInMenu, "Func_ShowTab_ByMenu")
		$itemSchoolsOfTabInMenu = GUICtrlCreateMenuItem("  学校信息", $menuTab)
			GUICtrlSetOnEvent($itemSchoolsOfTabInMenu, "Func_ShowTab_ByMenu")
		$itemPartnersOfTabInMenu = GUICtrlCreateMenuItem("  合作伙伴", $menuTab)
			GUICtrlSetOnEvent($itemPartnersOfTabInMenu, "Func_ShowTab_ByMenu")
		$itemProjectsOfTabInMenu = GUICtrlCreateMenuItem("  工程管理", $menuTab)
			GUICtrlSetOnEvent($itemProjectsOfTabInMenu, "Func_ShowTab_ByMenu")
	
	$menuHelp = GUICtrlCreateMenu ( "帮助 &H")
		$itemGuideInMenuHelp = GUICtrlCreateMenuItem("使用指南", $menuHelp)
			GUICtrlSetOnEvent($itemGuideInMenuHelp, "Func_MenuItemGuide_In_MenuHelp")	; 打开”公司产品”标签卡
		$itemAboutInMenuHelp = GUICtrlCreateMenuItem("关于", $menuHelp)
			GUICtrlSetOnEvent($itemAboutInMenuHelp, "Func_MenuItemAbout_In_MenuHelp")	; 打开”学校信息”标签卡
	
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
			_GUIImageList_Add($imgList, _WinAPI_CreateSolidBitmap($guiMainWindow, 0x990033, 16, 16))
			_GUIImageList_Add($imgList, _WinAPI_CreateSolidBitmap($guiMainWindow, 0x00FF00, 16, 16))
			_GUIImageList_Add($imgList, _WinAPI_CreateSolidBitmap($guiMainWindow, 0x0000FF, 16, 16))
			_GUIImageList_Add($imgList, _WinAPI_CreateSolidBitmap($guiMainWindow, 0xFF3399, 16, 16))
			_GUIImageList_Add($imgList, _WinAPI_CreateSolidBitmap($guiMainWindow, 0x007700, 16, 16))
			_GUIImageList_Add($imgList, _WinAPI_CreateSolidBitmap($guiMainWindow, 0xFF6600, 16, 16))	; 查询结果
			_GUICtrlTab_SetImageList($tabInMainWindow, $imgList)
	
		_GUICtrlTab_InsertItem ( $tabInMainWindow, 0, "工作日志", 0)
	
	$mouseMenuTab = GUICtrlCreateContextMenu($tabInMainWindow)
		$mouseMenuItemClose = GUICtrlCreateMenuItem("关闭", $mouseMenuTab)
			GUICtrlSetOnEvent($mouseMenuItemClose, "Func_MouseMenuItem")
		GUICtrlCreateMenuItem("", $mouseMenuTab)
		$mouseMenuItemSaveAs = GUICtrlCreateMenuItem("另存为", $mouseMenuTab)
			GUICtrlSetOnEvent($mouseMenuItemSaveAs, "Func_MouseMenuItem")
		GUICtrlCreateMenuItem("", $mouseMenuTab)
		$mouseMenuItemPrint = GUICtrlCreateMenuItem("打印", $mouseMenuTab)
			GUICtrlSetOnEvent($mouseMenuItemPrint, "Func_MouseMenuItem")

	GUICtrlCreateLabel("文本 3", 2, $HeightOfWindow - 38, 50, 20)
	
	$lvInTabJournal = GUICtrlCreateListView ( "编号|人员|日期|地点|交通|食宿|工作描述|备注", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)	; 工作日志表
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabJournal, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))	; 统计设置多种扩展样式
		GUICtrlSetBkColor($lvInTabJournal, 0xffffff)					;设置listview的背景色
		GUICtrlSetBkColor($lvInTabJournal, $GUI_BKCOLOR_LV_ALTERNATE)	;奇数行为listview的背景色，偶数行为listviewitem的背景色
		GUICtrlSetState($lvInTabJournal, $GUI_SHOW)
	$lvInTabProducts = GUICtrlCreateListView ( "编号|产品|类型|设计|配置|造价|备注", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)	; 公司产品表
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabProducts, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabProducts, 0xffffff)
		GUICtrlSetBkColor($lvInTabProducts, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabProducts, $GUI_HIDE)
	$lvInTabSchools  = GUICtrlCreateListView ( "编号|学校|联系人|地址|电话|邮箱|备注", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)	; 学校信息表
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabSchools, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabSchools, 0xffffff)
		GUICtrlSetBkColor($lvInTabSchools, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabSchools, $GUI_HIDE)
	$lvInTabProjects = GUICtrlCreateListView ( "编号|名称|学校|产品|合作者|我司负责人|细则|起始日期|状态|结束日期|结算记录|备注", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)	; 工程管理表
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabProjects, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabProjects, 0xffffff)
		GUICtrlSetBkColor($lvInTabProjects, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabProjects, $GUI_HIDE)
	$lvInTabPartners = GUICtrlCreateListView ( "编号|名称|类型|地址|电话|邮箱|业务|备注", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)	; 合作伙伴表
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabPartners, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabPartners, 0xffffff)
		GUICtrlSetBkColor($lvInTabPartners, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabPartners, $GUI_HIDE)
#cs
	$lvInTabAssets   = GUICtrlCreateListView ( "编号|名称|串号|单位|类型|购入日期|单价|所属部门|经销商|是否报废|备注", 3, 51, $WidthOfWindow - 6, $HeightOfWindow - 97)	; 资产管理表
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabAssets, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabAssets, 0xffffff)
		GUICtrlSetBkColor($lvInTabAssets, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabAssets, $GUI_HIDE)
#ce

GUISetState(@SW_SHOW, $guiMainWindow)

; #cs
;	初始在“日志”数据库查询，在“日志”列表插入数据
; #ce
FuncReadDb( $tb_journal, $lvInTabJournal )
$hasReadedDbTbJournal = True

;======================================================================
GUIRegisterMsg($WM_NOTIFY, "_WM_NOTIFY")
GUIRegisterMsg($WM_COMMAND, '_WM_COMMAND')

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
	
	Switch $itemText
		Case "工作日志"
			GUICtrlSetState($lvInTabJournal, $GUI_SHOW)
		Case "公司产品"
			GUICtrlSetState($lvInTabProducts, $GUI_SHOW)
		Case "学校信息"
			GUICtrlSetState($lvInTabSchools, $GUI_SHOW)
		Case "工程管理"
			GUICtrlSetState($lvInTabProjects, $GUI_SHOW)
		Case "合作伙伴"
			GUICtrlSetState($lvInTabPartners, $GUI_SHOW)
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
        Case $toolbarInMainWindow
            Switch $code	; 事件
				Case $TBN_HOTITEMCHANGE
					$tNMTBHOTITEM = DllStructCreate($tagNMTBHOTITEM, $lParam)
					$i_idOld = DllStructGetData($tNMTBHOTITEM, "idOld")
					$i_idNew = DllStructGetData($tNMTBHOTITEM, "idNew")
					$itemInToolbar = $i_idNew
 
				Case $NM_CLICK
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
		
		Case GUICtrlGetHandle($tabInMainWindow)
            Switch $code
				Case $NM_RCLICK
					
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
						
						Switch _GUICtrlTab_GetItemText ( $tabInMainWindow, $aHit[0] )
							Case "工作日志"
								GUICtrlSetState($lvInTabJournal, $GUI_SHOW)
							Case "公司产品"
								GUICtrlSetState($lvInTabProducts, $GUI_SHOW)
							Case "学校信息"
								GUICtrlSetState($lvInTabSchools, $GUI_SHOW)
							Case "工程管理"
								GUICtrlSetState($lvInTabProjects, $GUI_SHOW)
							Case "合作伙伴"
								GUICtrlSetState($lvInTabPartners, $GUI_SHOW)
						EndSwitch
			EndSwitch
		
		Case GUICtrlGetHandle($lvInTabJournal)
            Switch $code
				Case $NM_DBLCLK
					$columnOfTable = $tb_journal
					$columnOfLV = $lvInTabJournal
					FuncCreateEditRecForColumn( $lvInTabJournal, $lParam )
			EndSwitch
		Case GUICtrlGetHandle($lvInTabProducts)
            Switch $code
				Case $NM_DBLCLK
					$columnOfTable = $tb_products
					$columnOfLV = $lvInTabProducts
					FuncCreateEditRecForColumn( $lvInTabProducts, $lParam )
			EndSwitch
		Case GUICtrlGetHandle($lvInTabProjects)
            Switch $code
				Case $NM_DBLCLK
					$columnOfTable = $tb_projects
					$columnOfLV = $lvInTabProjects
					FuncCreateEditRecForColumn( $lvInTabProjects, $lParam )
			EndSwitch
		Case GUICtrlGetHandle($lvInTabPartners)
            Switch $code
				Case $NM_DBLCLK	
					$columnOfTable = $tb_partners
					$columnOfLV = $lvInTabPartners
					FuncCreateEditRecForColumn( $lvInTabPartners, $lParam )
			EndSwitch
		Case GUICtrlGetHandle($lvInTabSchools)
            Switch $code
				Case $NM_DBLCLK
					$columnOfTable = $tb_schools
					$columnOfLV = $lvInTabSchools
					FuncCreateEditRecForColumn( $lvInTabSchools, $lParam )
			EndSwitch
			
	EndSwitch
	
    Return $GUI_RUNDEFMSG
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
	
	$popupWindow = GUICreate("" , 500, 175, Default, Default, $WS_CAPTION, BitOR($WS_EX_TOPMOST, $GUI_WS_EX_PARENTDRAG))
		WinSetTitle($popupWindow, "", "在《" & $strLvName & "》中查询数据" )
		
		GUICtrlCreateGroup("在需要的条件里输入关键字。多个词用空格分开", 10, 10, 480, 125)

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
		EndSwitch
		
		GUICtrlCreateGroup("", -99, -99, 1, 1) 
		
		$btnOKInPopWin = GUICtrlCreateButton("确定", 10, 145, 60, 22)
			GUICtrlSetOnEvent($btnOKInPopWin, "FuncFindDataByArgs")
		$btnNOInPopWin = GUICtrlCreateButton("取消", 428, 145, 60, 22)
			GUICtrlSetOnEvent($btnNOInPopWin, "FuncFindDataByArgs")
		
	GUISetState(@SW_SHOW, $popupWindow)
	GUISetState(@SW_DISABLE, $guiMainWindow) 
EndFunc

Func FuncFindDataByArgs ()
	If @GUI_CtrlId = $btnOKInPopWin Then
		Local $isSqlStrEnable = False	; 查询语句是否可用。如果全部参数都是空的，不会设为true
		
		$strArgsOfQuery = "SELECT * FROM " & $strArgsOfTableArr[0] & " WHERE "
		
		; 遍历控件数组，拼接语句
		For $iiii = 1 To 6 Step 1	; 0索引保存的是lV或Access名称。剩余1-6保存控件
			
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
						$strArgsOfQuery &= $strArgsOfTableArr[$iiii] & " LIKE '%" & $keywordArr[$j] & "%' AND "
					Next
				EndIf
				
				$isSqlStrEnable = True	; 至少有一个有效参数，才可用
			EndIf
		Next
		
		$strArgsOfQuery &= "1 = 1"	; 拼接sql语句常用结尾。否则删除最后一个 AND比较麻烦
		
		If $isSqlStrEnable Then
			; 执行数据库查询
			Local $adoCon = ObjCreate("ADODB.Connection")
			$adoCon.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & $db_path & "; Jet OLEDB:Database Password='" & $db_pswd & "'")
			Local $adoRec = ObjCreate("ADODB.Recordset")
			$adoRec.ActiveConnection = $adoCon
			
			$adoRec.Open($strArgsOfQuery)	; 从表中查到的结果集
			
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
						$strResultDataForLv &= $adoRec.fields($i).value & "|" 
					Next
					
					GUICtrlCreateListViewItem($strResultDataForLv, $strArgsOfLVArr[0])
						GUICtrlSetBkColor (-1, 0xffa500 );设置listviewitem的背景色
					
					$adoRec.Movenext	; 结果集的下一行
				EndIf
			WEnd
			_GUICtrlListView_EndUpdate($strArgsOfLVArr[0])
			_GUICtrlListView_Scroll($strArgsOfLVArr[0], 0, _GUICtrlListView_GetItemCount($strArgsOfLVArr[0])*10)

			$adoRec.Close
			$adoCon.Close
		EndIf
		
	EndIf
	
	GUIDelete ( $popupWindow )
	GUISetState(@SW_ENABLE, $guiMainWindow) 
	GUISetState(@SW_RESTORE, $guiMainWindow)
	
	; 复位。清空数组
	$strArgsOfQuery = ""
	For $i = 0 To 6
		$strArgsOfLVArr[$i] = ""
		$strArgsOfTableArr[$i] = ""
	Next
EndFunc

; #cs
; 双击LV的单元格后，创建一个输入框
; 对原数据进行修改
; 参数 $strLvInTab:要操作的列表
; 参数 $lParam:系统消息
; #ce
Func FuncCreateEditRecForColumn( $strLvInTab, $lParam )
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
EndFunc

; 控件失去焦点
; 单元格新数据保存
Func _WM_COMMAND($hWnd, $msg, $wParam, $lParam)
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
						$strLvItem = "'Eminem', 'update', '" & $columnOfTable & "', '" & $columnName & "', '" & $columnOldData & "', '" & $iText & "', '" & @YEAR & "-" & @MON & "-" & @MDAY & "(" & @WDAY - 1 & ")" & @HOUR & ":" & @MIN & ":" & @SEC & "`" & @MSEC & "', '_none'"	; 数据
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
	EndSwitch
	
	FuncDeleteItemFromAccessAndListView ( $strTbName, $strLvInTab, $strLvItem )
EndFunc

; 删除选中的行
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
		
		; 日志表
		$strLvItem = "'Eminem', 'del', '" & $strTbName & "', '_item', '" & StringTrimLeft ($strLvItem, StringInStr( $strLvItem, "|")) & "', '_none', '" & @YEAR & "-" & @MON & "-" & @MDAY & "(" & @WDAY - 1 & ")" & @HOUR & ":" & @MIN & ":" & @SEC & "`" & @MSEC & "', '_none'"	; 数据
		$strLvItem = "insert into " & $tb_log & " (l_name, l_operate, l_table, l_column, l_old_data, l_new_data, l_date, l_note) values ( " & $strLvItem & " )"				; 插入语句
		$dataAdodbConnectionDel.Execute($strLvItem)
		
		$dataAdodbConnectionDel.Close
	EndIf
EndFunc

; #cs
; 响应工具条“新建”按钮。插入条目到ListView
; 首先判断当前激活的标签-哪一个ListView处于显示状态
; 然后插入条目，更新条目位置
; #ce
Func FuncInsertItemToListView ()
	Local $tmpLvItem, $strLvItem, $strLvInTab, $strTbName
	Switch _GUICtrlTab_GetItemText ( $tabInMainWindow, _GUICtrlTab_GetCurSel ( $tabInMainWindow ))
		Case "工作日志"
			$strLvItem = "编号|GGGG人员|日期|地点|交通|食宿|工作描述|备注"	; 插入到LV中的条目
			$strLvInTab = $lvInTabJournal	; LV
			$strTbName = $tb_journal		; Table
		Case "公司产品"
			$strLvItem = "编号|GSGS产品|类型|设计|配置|造价|备注"
			
			$strLvInTab = $lvInTabProducts
			$strTbName = $tb_products
		Case "学校信息"
			$strLvItem = "编号|XXXX学校|联系人|地址|电话|邮箱|备注"
			
			$strLvInTab = $lvInTabSchools
			$strTbName = $tb_schools
		Case "工程管理"
			$strLvItem = "编号|GCGC名称|学校|产品|合作者|我司负责人|细则（照片）|起始日期| 状态|结束日期|结算记录|备注"
			
			$strLvInTab = $lvInTabProjects
			$strTbName = $tb_projects
		Case "合作伙伴"
			$strLvItem = "编号|HZHZ名称|类型|地址|电话|邮箱|业务|备注"
			
			$strLvInTab = $lvInTabPartners
			$strTbName = $tb_partners
	EndSwitch
	
	$tmpLvItem = GUICtrlCreateListViewItem( $strLvItem, $strLvInTab)
	GUICtrlSetBkColor ($tmpLvItem, 0xff9900 )
	_GUICtrlListView_Scroll($strLvInTab, 0, _GUICtrlListView_GetItemCount($strLvInTab)*15)
	
	; 第一列是“编号”，去掉
	$strLvItem = "'" & StringReplace ( StringTrimLeft ($strLvItem, StringInStr( $strLvItem, "|")), "|", "', '" ) & "'"	; 转为既定格式的数据库可用字串
	
	FuncInsertItemToAccessAndListView ( $strTbName, $strLvItem )
EndFunc

; #cs
; 向数据库中写数据条目
; $strTbName：表名称
; $strLvItem：数据
; #ce
Func FuncInsertItemToAccessAndListView ( $strTbName, $strLvItem )
	Local $strColumns
	Switch $strTbName
		Case $tb_products	; 产品（）
			$strColumns = " (pd_name, pd_type, pd_desiger, pd_configuration, pd_cost, pd_note) values ("
		Case $tb_partners	; 伙伴
			$strColumns = " (pt_name, pt_type, pt_address, pt_phone, pt_email, pt_business, pt_note) values ("
		Case $tb_projects	; 工程
			$strColumns = " (pj_name, pj_s_name, pj_pd_name, pj_pt_name, pj_w_name, pj_content, pj_date_start, pj_state, pj_date_finish, pj_account, pj_note) values ("
		Case $tb_journal	; 日志
			$strColumns = " (j_name, j_date, j_address, j_traffic, j_board, j_content, j_note) values ("	; , j_record, j_date_record
		Case $tb_schools	; 学校"
			$strColumns = " (s_name, s_contact, s_address, s_phone, s_email, s_note) values ("
	EndSwitch
	Local $dataAdodbConnectionIn = ObjCreate("ADODB.Connection")
	$dataAdodbConnectionIn.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & $db_path & ";Jet Oledb:Database Password='" & $db_pswd & "'")
	$dataAdodbConnectionIn.Execute("insert into " & $strTbName & $strColumns & $strLvItem & ")")
	; 日志表
	$strLvItem = "'Eminem', 'add', '" & $strTbName & "', '_item', '_none', '" & StringReplace($strLvItem, "'", "") & "', '" & @YEAR & "-" & @MON & "-" & @MDAY & "(" & @WDAY - 1 & ")" & @HOUR & ":" & @MIN & ":" & @SEC & "`" & @MSEC & "', '_none'"	; 数据
	$strLvItem = "insert into " & $tb_log & " (l_name, l_operate, l_table, l_column, l_old_data, l_new_data, l_date, l_note) values ( " & $strLvItem & " )"				; 插入语句
	$dataAdodbConnectionIn.Execute($strLvItem)
	$dataAdodbConnectionIn.Close

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
			
			If $idOfTabItem = 0 Then
				_GUICtrlTab_SetCurSel($tabInMainWindow, 1)
			Else 
				_GUICtrlTab_SetCurSel($tabInMainWindow, $idOfTabItem - 1)
			EndIf
			
			Switch _GUICtrlTab_GetItemText($tabInMainWindow, $idOfTabItem)
				Case "工作日志"
					GUICtrlSetData($itemJournalOfTabInMenu, "　工作日志")
				Case "公司产品"
					GUICtrlSetData($itemProductsOfTabInMenu, "　公司产品")
				Case "学校信息"
					GUICtrlSetData($itemSchoolsOfTabInMenu, "　学校信息")
				Case "工程管理"
					GUICtrlSetData($itemProjectsOfTabInMenu, "　工程管理")
				Case "合作伙伴"
					GUICtrlSetData($itemPartnersOfTabInMenu, "　合作伙伴")
			EndSwitch
			
			_GUICtrlTab_DeleteItem ( $tabInMainWindow, $idOfTabItem )
			
			$strTabItemText = _GUICtrlTab_GetItemText ( $tabInMainWindow, _GUICtrlTab_GetCurSel ( $tabInMainWindow ) )
			
			Switch $strTabItemText
				Case "工作日志"
					GUICtrlSetState($lvInTabJournal, $GUI_SHOW)
				Case "公司产品"
					GUICtrlSetState($lvInTabProducts, $GUI_SHOW)
				Case "学校信息"
					GUICtrlSetState($lvInTabSchools, $GUI_SHOW)
				Case "工程管理"
					GUICtrlSetState($lvInTabProjects, $GUI_SHOW)
				Case "合作伙伴"
					GUICtrlSetState($lvInTabPartners, $GUI_SHOW)
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
		Case $itemJournalOfTabInMenu
			Func_ShowTab_ByText ("工作日志")
		Case $itemProductsOfTabInMenu
			Func_ShowTab_ByText ("公司产品")
		Case $itemSchoolsOfTabInMenu
			Func_ShowTab_ByText ("学校信息")
		Case $itemProjectsOfTabInMenu
			Func_ShowTab_ByText ("工程管理")
		Case $itemPartnersOfTabInMenu
			Func_ShowTab_ByText("合作伙伴")
	EndSwitch
EndFunc

#cs 通过传入的字符串，和标签头进行匹配。控制显示或隐藏
#ce
Func Func_ShowTab_ByText ( $strItemName )
	; $itemJournalOfTabInMenu $itemProductsOfTabInMenu $itemSchoolsOfTabInMenu $itemProjectsOfTabInMenu $itemPartnersOfTabInMenu "　公司产品"
	GUICtrlSetState($lvInTabJournal, $GUI_HIDE)
	GUICtrlSetState($lvInTabProducts, $GUI_HIDE)
	GUICtrlSetState($lvInTabSchools, $GUI_HIDE)
	GUICtrlSetState($lvInTabProjects, $GUI_HIDE)
	GUICtrlSetState($lvInTabPartners, $GUI_HIDE)
	
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
				GUICtrlSetData($itemJournalOfTabInMenu, "  工作日志")
			Case "公司产品"
				GUICtrlSetData($itemProductsOfTabInMenu, "  公司产品")
			Case "学校信息"
				GUICtrlSetData($itemSchoolsOfTabInMenu, "  学校信息")
			Case "工程管理"
				GUICtrlSetData($itemProjectsOfTabInMenu, "  工程管理")
			Case "合作伙伴"
				GUICtrlSetData($itemPartnersOfTabInMenu, "  合作伙伴")
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
					GUICtrlSetData($itemJournalOfTabInMenu, "√ 工作日志")
				Case "公司产品"
					If $hasReadedDbTbProducts = False Then
						FuncReadDb($tb_products, $lvInTabProducts)
						$hasReadedDbTbProducts = True
					EndIf
					
					GUICtrlSetState($lvInTabProducts, $GUI_SHOW)
					GUICtrlSetData($itemProductsOfTabInMenu, "√ 公司产品")
				Case "学校信息"
					If $hasReadedDbTbSchools = False Then
						FuncReadDb($tb_schools, $lvInTabSchools)
						$hasReadedDbTbSchools = True
					EndIf
					
					GUICtrlSetState($lvInTabSchools, $GUI_SHOW)
					GUICtrlSetData($itemSchoolsOfTabInMenu, "√ 学校信息")
				Case "工程管理"
					If $hasReadedDbTbProjects = False Then
						FuncReadDb($tb_projects, $lvInTabProjects)
						$hasReadedDbTbProjects = True
					EndIf
					
					GUICtrlSetState($lvInTabProjects, $GUI_SHOW)
					GUICtrlSetData($itemProjectsOfTabInMenu, "√ 工程管理")
				Case "合作伙伴"
					If $hasReadedDbTbPartners = False Then
						FuncReadDb($tb_partners, $lvInTabPartners)
						$hasReadedDbTbPartners = True
					EndIf
					
					GUICtrlSetState($lvInTabPartners, $GUI_SHOW)
					GUICtrlSetData($itemPartnersOfTabInMenu, "√ 合作伙伴")
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
		EndSwitch
		
		_GUICtrlTab_InsertItem ( $tabInMainWindow, $id_item, $strItemName, $id_item)
		
		For $i = 0 To _GUICtrlTab_GetItemCount ( $tabInMainWindow ) - 1 Step 1
			If $strItemName = _GUICtrlTab_GetItemText ( $tabInMainWindow, $i ) Then
				_GUICtrlTab_SetCurSel($tabInMainWindow, $i)
				
				; $itemJournalOfTabInMenu $itemProductsOfTabInMenu $itemSchoolsOfTabInMenu $itemProjectsOfTabInMenu $itemPartnersOfTabInMenu "　公司产品"
				GUICtrlSetState($lvInTabJournal, $GUI_HIDE)
				GUICtrlSetState($lvInTabProducts, $GUI_HIDE)
				GUICtrlSetState($lvInTabSchools, $GUI_HIDE)
				GUICtrlSetState($lvInTabProjects, $GUI_HIDE)
				GUICtrlSetState($lvInTabPartners, $GUI_HIDE)
				
				Switch $strItemName
					Case "工作日志"
						If $hasReadedDbTbJournal = False Then
							FuncReadDb($tb_journal, $lvInTabJournal)
							$hasReadedDbTbJournal = True
						EndIf
						
						GUICtrlSetState($lvInTabJournal, $GUI_SHOW)
						GUICtrlSetData($itemJournalOfTabInMenu, "√ 工作日志")
					Case "公司产品"
						If $hasReadedDbTbProducts = False Then
							FuncReadDb($tb_products, $lvInTabProducts)
							$hasReadedDbTbProducts = True
						EndIf
						
						GUICtrlSetState($lvInTabProducts, $GUI_SHOW)
						GUICtrlSetData($itemProductsOfTabInMenu, "√ 公司产品")
					Case "学校信息"
						If $hasReadedDbTbSchools = False Then
							FuncReadDb($tb_schools, $lvInTabSchools)
							$hasReadedDbTbSchools = True
						EndIf
						
						GUICtrlSetState($lvInTabSchools, $GUI_SHOW)
						GUICtrlSetData($itemSchoolsOfTabInMenu, "√ 学校信息")
					Case "工程管理"
						If $hasReadedDbTbProjects = False Then
							FuncReadDb($tb_projects, $lvInTabProjects)
							$hasReadedDbTbProjects = True
						EndIf
						
						GUICtrlSetState($lvInTabProjects, $GUI_SHOW)
						GUICtrlSetData($itemProjectsOfTabInMenu, "√ 工程管理")
					Case "合作伙伴"
						If $hasReadedDbTbPartners = False Then
							FuncReadDb($tb_partners, $lvInTabPartners)
							$hasReadedDbTbPartners = True
						EndIf
						
						GUICtrlSetState($lvInTabPartners, $GUI_SHOW)
						GUICtrlSetData($itemPartnersOfTabInMenu, "√ 合作伙伴")
				EndSwitch
			EndIf
		Next
	EndIf
EndFunc

Func Func_MenuItemGuide_In_MenuHelp ()
EndFunc

Func Func_MenuItemAbout_In_MenuHelp ()
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
	; 日志表特殊：最后两个字段隐藏
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
			$strFields &= $obj_adodb_recordset.Fields( $i ).value & "|"
		Next
		$strFields = StringTrimRight ( $strFields, 1 )	; 删除末尾的 | 
		
		; 呈现 列
		GUICtrlCreateListViewItem( $strFields, $strListView)
			GUICtrlSetBkColor (-1, 0xffa500 );设置listviewitem的背景色
		$obj_adodb_recordset.movenext
	WEnd
	_GUICtrlListView_EndUpdate($strListView)
	
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

	;=== 公司人员表 tb_workers ：
	;                              编号。主键自动增长|                   姓名。文本|  身份证号|               部门|              职务|            入职日期|             转正日期|            月薪资|            性别|       生日|            电话|         邮箱|         住址，默认50字符|    家庭信息|          备注
	Local $strColsForTbWorders = "id integer identity(1,1) primary key, w_name text, w_identity_number text, w_deportment text, w_position text, w_date_of_entry text, w_date_regular text, w_month_wage text, w_sex text, w_birthday text, w_phone text, w_email text, w_address text(250), w_family text(254), w_note memo"

	;=== 学校信息表 tb_schools ：
	;                              编号|                                 学校|        联系人|         地址|                电话|         邮箱|        备注
	Local $strColsForTbSchools = "id integer identity(1,1) primary key, s_name text, s_contact text, s_address text(250), s_phone text, s_email text, s_note memo"

	;=== 公司产品表 tb_products ：
	;                               编号|                                 产品|         类型|        设计|             配置|                       造价|        备注
	Local $strColsForTbProducts = "id integer identity(1,1) primary key, pd_name text, pd_type text, pd_desiger text, pd_configuration text(250), pd_cost text, pd_note memo"

	;=== 资产管理表 tb_accets ：
	;                             编号|                                 名称|        串号|                 单位|        类型|        购入日期|           单价|         所属部门|          经销商|        是否报废|     备注
	Local $strColsForTbAccets = "id integer identity(1,1) primary key, a_name text, a_serial_number text, a_unit text, a_type text, a_date_bought text, a_price text, a_deportment text, a_dealer text, a_scrap text, a_note text(250)"

	;=== 合作伙伴表 tb_partners ：
	;                               编号|                                 名称|         类型|         地址|                 电话|          邮箱|          业务|             备注
	Local $strColsForTbPartners = "id integer identity(1,1) primary key, pt_name text, pt_type text, pt_address text(250), pt_phone text, pt_email text, pt_business text, pt_note memo"

	;=== 工程管理表 tb_projects ：
	;                               编号|                                 名称|         学校|           产品|            合作者|          我司负责人|     细则（照片）|         起始日期|           状态|          结束日期|            结算记录|             备注
	Local $strColsForTbProjects = "id integer identity(1,1) primary key, pj_name text, pj_s_name text, pj_pd_name text, pj_pt_name text, pj_w_name text, pj_content text(250), pj_date_start text, pj_state text, pj_date_finish text, pj_account text(250), pj_note memo"

	;=== 工作日志表 tb_journal ：
	;                              编号|                                 人员|        日期|        地点|                交通|                食宿|              工作描述|            备注|        记录人(隐藏)|        记录日期(隐藏)
	Local $strColsForTbJournal = "id integer identity(1,1) primary key, j_name text, j_date text, j_address text(250), j_traffic text(250), j_board text(250), j_content text(250), j_note memo, j_record text, j_date_record text"

	;=== 用户表 tb_users ：
	;                            编号|                                 用户名|      密码|        权限|             备注
	Local $strColsForTbUsers = "id integer identity(1,1) primary key, u_name text, u_pswd text, u_authority text, u_note text(250)"

	;=== 源数据表 tb_source ：
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
