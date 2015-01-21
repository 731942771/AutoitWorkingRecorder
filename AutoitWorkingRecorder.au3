#Region ;**** �� AccAu3Wrapper_GUI ����ָ�� ****
#AccAu3Wrapper_Icon=favicon.ico
#AccAu3Wrapper_OutFile=C:\Documents and Settings\Administrator\����\vigiles.exe
#AccAu3Wrapper_Compression=4
#AccAu3Wrapper_Res_Comment=�����鲩��
#AccAu3Wrapper_Res_Description=www.cuiweiyou.com
#AccAu3Wrapper_Res_Fileversion=8.8.8.8
#AccAu3Wrapper_Res_ProductVersion=9.9.9.9
#AccAu3Wrapper_Res_LegalCopyright=vigiles
#AccAu3Wrapper_Res_Language=2052
#AccAu3Wrapper_Res_requestedExecutionLevel=None
#AccAu3Wrapper_Res_Field=OriginalFilename|�лս���-��ά��
#AccAu3Wrapper_Res_Field=ProductName|�лս���-��ά��
#AccAu3Wrapper_Res_Field=ProductVersion|V1.0
#AccAu3Wrapper_Res_Field=InternalName|�лս���-��ά��
#AccAu3Wrapper_Res_Field=FileDescription|�лս���-��ά��
#AccAu3Wrapper_Res_Field=Comments|�лս���-��ά��
#AccAu3Wrapper_Res_Field=LegalTrademarks|cuiweiyou.com
#AccAu3Wrapper_Res_Field=CompanyName|cuiweiyou.com
#AccAu3Wrapper_Run_AU3Check=n
#AutoIt3Wrapper_UseX64=n
#Tidy_Parameters=/sfc/rel
#AccAu3Wrapper_Tidy_Stop_OnError=n
#EndRegion ;**** �� AccAu3Wrapper_GUI ����ָ�� ****

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
; ȫ�ֵ����ݿ�·��
Global $db_path = @ScriptDir & "\mdbZhongHui.mdb"
Global $db_pswd = ""	; ��������mdb�ļ�������

; �õ��ı���          ��˾��Ա                   ѧУ��Ϣ                      ��˾��Ʒ                  �ʲ�����                     �������                       ���̹���                   ������־                   �û�                     Ԫ����              ������־
Global $tb_workers = "tb_workers", $tb_schools = "tb_schools", $tb_products = "tb_products", $tb_accets = "tb_accets", $tb_partners = "tb_partners", $tb_projects = "tb_projects", $tb_journal = "tb_journal", $tb_users = "tb_users", $tb_source = "tb_source", $tb_log = "tb_log"

;-------------------- Դ���ݲ��� -------------------------------
If FileExists ( $db_path ) = 0 Then
	;If MsgBox(52, "����", "δ��⵽���ݿ��ļ����Ƿ����´�����") = 6 Then
		FuncCreateDb ( )
	;EndIf
EndIf

;======================================================================
;-------------------- �û���¼ --------------------------------


;======================================================================
; ������ǩ�л�LVʱ�õ���
Global $itemInToolbar, $idOfTabItem	; ���һ��ȫ�ֱ���$idOfTabItem��¼���һ��ı�ǩ����
; ��������
Global $WidthOfWindow = 1100, $HeightOfWindow = 600
; �����壬����������������ʾ��
Global $guiMainWindow, $toolbarInMainWindow, $hToolTip
; �������ϵİ�ť�ȶ�ID
Global Enum $id_Toolbar_New = 1000, $id_Toolbar_Save, $id_Toolbar_Delete, $id_Toolbar_Find, $id_Toolbar_Help
; ���ݿ��ȡ��ʶ��������True
Global $hasReadedDbTbWorkers = False, $hasReadedDbTbSchools = False, $hasReadedDbTbProducts = False, $hasReadedDbTbAccets = False, $hasReadedDbTbPartners = False, $hasReadedDbTbProjects = False, $hasReadedDbTbUsers = False, $hasReadedDbTbSource = False, $hasReadedDbTbLog = False, $hasReadedDbTbJournal = False
; ˫�����һ���Ԫ��ʱ�����༭���õ���
Global $hEdit, $hBrush, $hDC, $hItemRow, $hItemColumn	; ˫��LV�����޸��õ���
; �һ�LV����Ŀ���У�ʱ�Ҽ��˵���ID
Global $mouseMenuLV, $mouseMenuItemDelLV, $mouseMenuItemUpdateLV, $mouseMenuItemCopyLV, $dataIndex
;
Global $hEnableListView
;
Global Enum $id_menu_lv_del = 2000, $id_menu_lv_update, $id_menu_lv_copy
;
Global $hLvInTabJournal, $hLvInTabAccets, $hLvInTabPartners, $hLvInTabProducts, $hLvInTabProjects, $hLvInTabSchools, $hLvInTabSources, $hLvInTabWorkers, $hLvInTabUsers, $hlvintabLog
; ˫����Ԫ���޸����ݣ�������־��Log��ʱ�õ��ġ��޸ĵ����������ĸ��������ĸ�LV����Ӧ���������޸�ǰ����
Global $columnOfTable, $columnOfLV, $columnName, $columnOldData
; �Ӵ��壬�Ӵ����ȷ����ť��ȡ����ť��
Global $popupWindow, $btnOKInPopWin, $btnNOInPopWin
; ��ѯʱ�����ѯ��������������Ҳ�õ���
Global $strArgsOfSql, $strArgsOfLVArr[15], $strArgsOfTableArr[15]	; ���ݱ�����Ҫ

;======================================================================
;-------------------- GUI ---------------------------------
$guiMainWindow = GUICreate("�лս���|������", $WidthOfWindow, $HeightOfWindow)
	GUISetOnEvent($GUI_EVENT_CLOSE, "Func_GUI_EVENT_CLOSE")
	GUISetIcon( @ScriptDir & "\favicon.ico")	; ���ó���ͼ��Ϊ�ű��ļ�ͬĿ¼�е�Logo.ico

	$menuFile = GUICtrlCreateMenu ( "�ļ� &F")
		$itemOpenInMenuFile = GUICtrlCreateMenuItem("��", $menuFile)
		$itemSaveInMenuFile = GUICtrlCreateMenuItem("����", $menuFile)
		GUICtrlCreateMenuItem("", $menuFile) ; �ָ���
		$itemRecentfilesInMenuFile = GUICtrlCreateMenu("������ļ�", $menuFile)
		GUICtrlCreateMenuItem("", $menuFile)
		$itemExitInMenuFile = GUICtrlCreateMenuItem("�˳�", $menuFile)
			GUICtrlSetOnEvent($itemExitInMenuFile, "Func_GUI_EVENT_CLOSE")

	$menuTab = GUICtrlCreateMenu ( "���� &W")
		$menuItemJournal = GUICtrlCreateMenuItem("�� ������־", $menuTab)
			GUICtrlSetOnEvent($menuItemJournal, "Func_ShowTab_ByMenu")
		$menuItemProducts = GUICtrlCreateMenuItem("  ��˾��Ʒ", $menuTab)
			GUICtrlSetOnEvent($menuItemProducts, "Func_ShowTab_ByMenu")
		$menuItemSchools = GUICtrlCreateMenuItem("  ѧУ��Ϣ", $menuTab)
			GUICtrlSetOnEvent($menuItemSchools, "Func_ShowTab_ByMenu")
		$menuItemPartners = GUICtrlCreateMenuItem("  �������", $menuTab)
			GUICtrlSetOnEvent($menuItemPartners, "Func_ShowTab_ByMenu")
		GUICtrlCreateMenuItem("", $menuTab) ; �ָ���
		$menuItemProjects = GUICtrlCreateMenuItem("  ���̹���", $menuTab)
			GUICtrlSetOnEvent($menuItemProjects, "Func_ShowTab_ByMenu")
		$menuItemUsers = GUICtrlCreateMenuItem("  �û�����", $menuTab)
			GUICtrlSetOnEvent($menuItemUsers, "Func_ShowTab_ByMenu")
		$menuItemAccets = GUICtrlCreateMenuItem("  �ʲ�����", $menuTab)
			GUICtrlSetOnEvent($menuItemAccets, "Func_ShowTab_ByMenu")
		$menuItemWorkers = GUICtrlCreateMenuItem("  ��Ա����", $menuTab)
			GUICtrlSetOnEvent($menuItemWorkers, "Func_ShowTab_ByMenu")
		GUICtrlCreateMenuItem("", $menuTab) ; �ָ���
		$menuItemLog = GUICtrlCreateMenuItem("  ������־", $menuTab)
			GUICtrlSetOnEvent($menuItemLog, "Func_ShowTab_ByMenu")
		$menuItemSources = GUICtrlCreateMenuItem("  Ԫ �� ��", $menuTab)
			GUICtrlSetOnEvent($menuItemSources, "Func_ShowTab_ByMenu")

	$menuHelp = GUICtrlCreateMenu ( "���� &H")
		$itemGuideInMenuHelp = GUICtrlCreateMenuItem("ʹ��ָ��", $menuHelp)
			GUICtrlSetOnEvent($itemGuideInMenuHelp, "Func_MenuHelp")	; �򿪡���˾��Ʒ����ǩ��
		$itemAboutInMenuHelp = GUICtrlCreateMenuItem("����", $menuHelp)
			GUICtrlSetOnEvent($itemAboutInMenuHelp, "Func_MenuHelp")	; �򿪡�ѧУ��Ϣ����ǩ��

	$toolbarInMainWindow = _GUICtrlToolbar_Create($guiMainWindow)
		; ����������ʾ�ؼ���Ȼ�������$WM_NOTIFY�л�ȡ��Ϣ��Ӧ
		$hToolTip = _GUIToolTip_Create($toolbarInMainWindow)
		_GUICtrlToolbar_SetToolTips($toolbarInMainWindow, $hToolTip)

		_GUICtrlToolbar_AddBitmap($toolbarInMainWindow, 1, -1, $IDB_STD_SMALL_COLOR)	; ���ͼ��
		_GUICtrlToolbar_AddButton($toolbarInMainWindow, $id_Toolbar_New, $STD_FILENEW)				; �½�
		_GUICtrlToolbar_AddButton($toolbarInMainWindow, $id_Toolbar_Save, $STD_FILESAVE)			; ����
		_GUICtrlToolbar_AddButton($toolbarInMainWindow, $id_Toolbar_Delete, $STD_DELETE)			; ɾ��
		_GUICtrlToolbar_AddButtonSep($toolbarInMainWindow)
		_GUICtrlToolbar_AddButton($toolbarInMainWindow, $id_Toolbar_Find, $STD_FIND)				; ����
		_GUICtrlToolbar_AddButtonSep($toolbarInMainWindow)
		_GUICtrlToolbar_AddButton($toolbarInMainWindow, $id_Toolbar_Help, $STD_HELP)				; ����

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

		_GUICtrlTab_InsertItem ( $tabInMainWindow, 0, "������־", 0)
		;_GUICtrlTab_InsertItem ( $tabInMainWindow, 1, "��˾��Ʒ", 1)
		;_GUICtrlTab_InsertItem ( $tabInMainWindow, 2, "ѧУ��Ϣ", 2)
		;_GUICtrlTab_InsertItem ( $tabInMainWindow, 3, "�������", 3)
		;_GUICtrlTab_InsertItem ( $tabInMainWindow, 4, "���̹���", 4)
		;_GUICtrlTab_InsertItem ( $tabInMainWindow, 5, "�û�����", 5)
		;_GUICtrlTab_InsertItem ( $tabInMainWindow, 6, "�ʲ�����", 6)
		;_GUICtrlTab_InsertItem ( $tabInMainWindow, 7, "��Ա����", 7)
		;_GUICtrlTab_InsertItem ( $tabInMainWindow, 8, "������־", 8)
		;_GUICtrlTab_InsertItem ( $tabInMainWindow, 9, "Ԫ �� ��", 9)

	$mouseMenuTab = GUICtrlCreateContextMenu($tabInMainWindow)
		$mouseMenuItemClose = GUICtrlCreateMenuItem("�ر�", $mouseMenuTab)
			GUICtrlSetOnEvent($mouseMenuItemClose, "Func_MouseMenuItem")
		GUICtrlCreateMenuItem("", $mouseMenuTab)
		$mouseMenuItemSaveAs = GUICtrlCreateMenuItem("���Ϊ", $mouseMenuTab)
			GUICtrlSetOnEvent($mouseMenuItemSaveAs, "Func_MouseMenuItem")
		GUICtrlCreateMenuItem("", $mouseMenuTab)
		$mouseMenuItemPrint = GUICtrlCreateMenuItem("��ӡ", $mouseMenuTab)
			GUICtrlSetOnEvent($mouseMenuItemPrint, "Func_MouseMenuItem")

	$strInfoLbl = GUICtrlCreateLabel("", 2, $HeightOfWindow - 38, $WidthOfWindow - 5, 20)

	$lvInTabJournal = GUICtrlCreateListView ( "���|��Ա|����|�ص�|��ͨ|ʳ��|��������|��ע", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)		; ������־��
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabJournal, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))			; ͳ�����ö�����չ��ʽ
		GUICtrlSetBkColor($lvInTabJournal, 0xffffff)					;����listview�ı���ɫ
		GUICtrlSetBkColor($lvInTabJournal, $GUI_BKCOLOR_LV_ALTERNATE)	;������Ϊlistview�ı���ɫ��ż����Ϊlistviewitem�ı���ɫ
		GUICtrlSetState($lvInTabJournal, $GUI_SHOW)
		$hLvInTabJournal = GUICtrlGetHandle($lvInTabJournal)
	$lvInTabProducts = GUICtrlCreateListView ( "���|��Ʒ|����|���|����|���|��ע", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)		; ��˾��Ʒ��
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabProducts, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabProducts, 0xffffff)
		GUICtrlSetBkColor($lvInTabProducts, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabProducts, $GUI_HIDE)
		$hLvInTabProducts = GUICtrlGetHandle($lvInTabProducts)
	$lvInTabSchools  = GUICtrlCreateListView ( "���|ѧУ|��ϵ��|��ַ|�绰|����|��ע", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)	; ѧУ��Ϣ��
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabSchools, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabSchools, 0xffffff)
		GUICtrlSetBkColor($lvInTabSchools, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabSchools, $GUI_HIDE)
		$hLvInTabSchools = GUICtrlGetHandle($lvInTabSchools)
	$lvInTabPartners = GUICtrlCreateListView ( "���|����|����|��ַ|�绰|����|ҵ��|��ע", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)	; ��������
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabPartners, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabPartners, 0xffffff)
		GUICtrlSetBkColor($lvInTabPartners, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabPartners, $GUI_HIDE)
		$hLvInTabPartners = GUICtrlGetHandle($lvInTabPartners)
	$lvInTabProjects = GUICtrlCreateListView ( "���|����|ѧУ|��Ʒ|������|��˾������|ϸ��|��ʼ����|״̬|��������|�����¼|��ע", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabProjects, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabProjects, 0xffffff)
		GUICtrlSetBkColor($lvInTabProjects, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabProjects, $GUI_HIDE)
		$hLvInTabProjects = GUICtrlGetHandle($lvInTabProjects)
	; �û�����
	$lvInTabUsers = GUICtrlCreateListView ( "���|�û���|����|Ȩ��|��ע", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabUsers, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabUsers, 0xffffff)
		GUICtrlSetBkColor($lvInTabUsers, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabUsers, $GUI_HIDE)
		$hLvInTabUsers = GUICtrlGetHandle($lvInTabUsers)
	; �ʲ�����
	$lvInTabAccets = GUICtrlCreateListView ( "���|����|����|��λ|����|��������|����|��������|������|�Ƿ񱨷�|��ע", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabAccets, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabAccets, 0xffffff)
		GUICtrlSetBkColor($lvInTabAccets, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabAccets, $GUI_HIDE)
		$hLvInTabAccets = GUICtrlGetHandle($lvInTabAccets)
	; ��Ա����
	$lvInTabWorkers = GUICtrlCreateListView ( "���|����|���֤��|����|ְ��|��ְ����|ת������|��н��|�Ա�|����|�绰|����|סַ|��ͥ��Ϣ|��ע", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabWorkers, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabWorkers, 0xffffff)
		GUICtrlSetBkColor($lvInTabWorkers, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabWorkers, $GUI_HIDE)
		$hLvInTabWorkers = GUICtrlGetHandle($lvInTabWorkers)
	; ������־
	$lvInTabLog = GUICtrlCreateListView ( "���|�û�|����|��|����|������|������|����|��ע", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabLog, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabLog, 0xffffff)
		GUICtrlSetBkColor($lvInTabLog, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabLog, $GUI_HIDE)
		$hlvintabLog = GUICtrlGetHandle($lvInTabLog)
	; Ԫ����
	$lvInTabSources = GUICtrlCreateListView ( "���|��|��|ֵ|����", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabSources, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabSources, 0xffffff)
		GUICtrlSetBkColor($lvInTabSources, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabSources, $GUI_HIDE)
		$hLvInTabSources = GUICtrlGetHandle($lvInTabSources)

GUISetState(@SW_SHOW, $guiMainWindow)

; #cs
;	��ʼ�ڡ���־�����ݿ��ѯ���ڡ���־���б��������
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
; �����������ǩ����
; �л���ǩ��ѡ��״̬
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
		Case "������־"
			GUICtrlSetState($lvInTabJournal, $GUI_SHOW)
			GUICtrlSetData($strInfoLbl, "�б�������־���м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabJournal ))
		Case "��˾��Ʒ"
			GUICtrlSetState($lvInTabProducts, $GUI_SHOW)
			GUICtrlSetData($strInfoLbl, "�б���˾��Ʒ���м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabProducts ))
		Case "ѧУ��Ϣ"
			GUICtrlSetState($lvInTabSchools, $GUI_SHOW)
			GUICtrlSetData($strInfoLbl, "�б�ѧУ��Ϣ���м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabSchools ))
		Case "���̹���"
			GUICtrlSetState($lvInTabProjects, $GUI_SHOW)
			GUICtrlSetData($strInfoLbl, "�б����̹����м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabProjects ))
		Case "�������"
			GUICtrlSetState($lvInTabPartners, $GUI_SHOW)
			GUICtrlSetData($strInfoLbl, "�б�������顷�м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabPartners ))

		Case "�û�����"
			GUICtrlSetState($lvInTabUsers, $GUI_SHOW)
			GUICtrlSetData($strInfoLbl, "�б��û������м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabUsers ))
		Case "�ʲ�����"
			GUICtrlSetState($lvInTabAccets, $GUI_SHOW)
			GUICtrlSetData($strInfoLbl, "�б��ʲ������м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabAccets ))
		Case "��Ա����"
			GUICtrlSetState($lvInTabWorkers, $GUI_SHOW)
			GUICtrlSetData($strInfoLbl, "�б���Ա�����м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabWorkers ))
		Case "������־"
			GUICtrlSetState($lvInTabLog, $GUI_SHOW)
			GUICtrlSetData($strInfoLbl, "�б�������־���м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabLog ))
		Case "Ԫ �� ��"
			GUICtrlSetState($lvInTabSources, $GUI_SHOW)
			GUICtrlSetData($strInfoLbl, "�б�Ԫ �� �ݡ��м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabSources ))
	EndSwitch
EndFunc

; ��Ϣ����
; ����UDF�����Ĺ������������ʾ���������
; ������ǩ���Ҽ�����¼�
; ����...
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
				DllStructSetData($tInfo, "aText", "�ڵ�ǰ���ݱ����������")
			Case $id_Toolbar_Save
				DllStructSetData($tInfo, "aText", "���浱ǰ���ݱ�Excel")
			Case $id_Toolbar_Delete
				DllStructSetData($tInfo, "aText", "ɾ����ǰ���ݱ�ѡ�����")
			Case $id_Toolbar_Find
				DllStructSetData($tInfo, "aText", "�ڵ�ǰ���ݱ��������")
			Case $id_Toolbar_Help
				DllStructSetData($tInfo, "aText", "����")
		EndSwitch
	EndIf

    Switch $hwndFrom	; �ؼ�
        Case $toolbarInMainWindow	;----------- ������
            Switch $code	; �¼�
				Case $TBN_HOTITEMCHANGE
					$tNMTBHOTITEM = DllStructCreate($tagNMTBHOTITEM, $lParam)
					$i_idOld = DllStructGetData($tNMTBHOTITEM, "idOld")
					$i_idNew = DllStructGetData($tNMTBHOTITEM, "idNew")
					$itemInToolbar = $i_idNew

				Case $NM_CLICK	; ������
					Switch $itemInToolbar
						Case $id_Toolbar_New
							$itemInToolbar = -1

							FuncInsertItemToListView()	; ������Ŀ��ListView

						Case $id_Toolbar_Save
							$itemInToolbar = -1

						Case $id_Toolbar_Delete
							$itemInToolbar = -1

							FuncDeleteItemFromListView()

						Case $id_Toolbar_Find
							$itemInToolbar = -1

							FuncFindData()	; �ڵ�ǰ��ʾ��LV��Ӧ�����

						Case $id_Toolbar_Help
							$itemInToolbar = -1

					EndSwitch

			EndSwitch

		Case GUICtrlGetHandle($tabInMainWindow)	;-------------------- ��ǩҳ
            Switch $code
				Case $NM_RCLICK	; �һ�

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
						Case "������־"
							GUICtrlSetState($lvInTabJournal, $GUI_SHOW)
							GUICtrlSetData($strInfoLbl, "�б�������־���м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabJournal ))
						Case "��˾��Ʒ"
							GUICtrlSetState($lvInTabProducts, $GUI_SHOW)
							GUICtrlSetData($strInfoLbl, "�б���˾��Ʒ���м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabProducts ))
						Case "ѧУ��Ϣ"
							GUICtrlSetState($lvInTabSchools, $GUI_SHOW)
							GUICtrlSetData($strInfoLbl, "�б�ѧУ��Ϣ���м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabSchools ))
						Case "���̹���"
							GUICtrlSetState($lvInTabProjects, $GUI_SHOW)
							GUICtrlSetData($strInfoLbl, "�б����̹����м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabProjects ))
						Case "�������"
							GUICtrlSetState($lvInTabPartners, $GUI_SHOW)
							GUICtrlSetData($strInfoLbl, "�б�������顷�м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabPartners ))
						Case "�û�����"
							GUICtrlSetState($lvInTabUsers, $GUI_SHOW)
							GUICtrlSetData($strInfoLbl, "�б��û������м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabUsers ))
						Case "�ʲ�����"
							GUICtrlSetState($lvInTabAccets, $GUI_SHOW)
							GUICtrlSetData($strInfoLbl, "�б��ʲ������м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabAccets ))
						Case "��Ա����"
							GUICtrlSetState($lvInTabWorkers, $GUI_SHOW)
							GUICtrlSetData($strInfoLbl, "�б���Ա�����м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabWorkers ))
						Case "������־"
							GUICtrlSetState($lvInTabLog, $GUI_SHOW)
							GUICtrlSetData($strInfoLbl, "�б�������־���м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabLog ))
						Case "Ԫ �� ��"
							GUICtrlSetState($lvInTabSources, $GUI_SHOW)
							GUICtrlSetData($strInfoLbl, "�б�Ԫ �� �ݡ��м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabSources ))
					EndSwitch
			EndSwitch

		Case GUICtrlGetHandle($lvInTabJournal)	;--------------------- �б�ؼ���������־
            Switch $code
				Case $NM_DBLCLK	; ˫��
					$columnOfTable = $tb_journal
					$columnOfLV = $lvInTabJournal

					FuncCreateEditRecForColumn("������־", $lvInTabJournal, $lParam )

				Case $NM_CLICK	; ������
					Local $ts = _GUICtrlListView_GetItemTextString ( $lvInTabJournal )
					If StringReplace ( $ts, "|", "") <> "" Then
						GUICtrlSetData($strInfoLbl, "�б�������־����ѡ�е��У�" & $ts )
					Else
						GUICtrlSetData($strInfoLbl, "" )
					EndIf
				#cs
				Case $NM_RCLICK	; �һ�
					$columnOfTable = $tb_journal
					$columnOfLV = $lvInTabJournal
					$columnOldData = _GUICtrlListView_GetItemTextString ( $lvInTabJournal )

					Local $tInfo = DllStructCreate($tagNMITEMACTIVATE, $lParam)
					$dataIndex = DllStructGetData($tInfo, "Index")		; ������С�ʼ��0
					Local $dataSubItem = DllStructGetData($tInfo, "SubItem")	; ������С�ʼ��0
					ConsoleWrite("�һ�----------�У�" & $dataIndex & "���У�" & $dataSubItem & @CRLF)

					If $dataIndex = -1 Then
						_GUICtrlMenu_DestroyMenu(GUICtrlGetHandle($mouseMenuLV))
					Else
						$mouseMenuLV = GUICtrlCreateContextMenu($lvInTabJournal)
							$mouseMenuItemDelLV = GUICtrlCreateMenuItem("ɾ��", $mouseMenuLV)
								GUICtrlSetOnEvent($mouseMenuItemDelLV, "Func_MouseMenuItem_LV")
							GUICtrlCreateMenuItem("", $mouseMenuLV)
							$mouseMenuItemUpdateLV = GUICtrlCreateMenuItem("�޸�", $mouseMenuLV)
								GUICtrlSetOnEvent($mouseMenuItemUpdateLV, "Func_MouseMenuItem_LV")
							GUICtrlCreateMenuItem("", $mouseMenuLV)
							$mouseMenuItemCopyLV = GUICtrlCreateMenuItem("����", $mouseMenuLV)
								GUICtrlSetOnEvent($mouseMenuItemCopyLV, "Func_MouseMenuItem_LV")
				EndIf
				#ce
			EndSwitch
		Case GUICtrlGetHandle($lvInTabProducts)
            Switch $code
				Case $NM_DBLCLK
					$columnOfTable = $tb_products
					$columnOfLV = $lvInTabProducts
					FuncCreateEditRecForColumn("��˾��Ʒ", $lvInTabProducts, $lParam )
				Case $NM_CLICK
					Local $ts = _GUICtrlListView_GetItemTextString ( $lvInTabProducts )
					If StringReplace ( $ts, "|", "") <> "" Then
						GUICtrlSetData($strInfoLbl, "�б���˾��Ʒ����ѡ�е��У�" & $ts )
					Else
						GUICtrlSetData($strInfoLbl, "" )
					EndIf
			EndSwitch
		Case GUICtrlGetHandle($lvInTabProjects)
            Switch $code
				Case $NM_DBLCLK
					$columnOfTable = $tb_projects
					$columnOfLV = $lvInTabProjects
					FuncCreateEditRecForColumn("���̹���", $lvInTabProjects, $lParam )
				Case $NM_CLICK
					Local $ts = _GUICtrlListView_GetItemTextString ( $lvInTabProjects )
					If StringReplace ( $ts, "|", "") <> "" Then
						GUICtrlSetData($strInfoLbl, "�б����̹�����ѡ�е��У�" & $ts )
					Else
						GUICtrlSetData($strInfoLbl, "" )
					EndIf
			EndSwitch
		Case GUICtrlGetHandle($lvInTabPartners)
            Switch $code
				Case $NM_DBLCLK
					$columnOfTable = $tb_partners
					$columnOfLV = $lvInTabPartners
					FuncCreateEditRecForColumn("�������", $lvInTabPartners, $lParam )
				Case $NM_CLICK
					Local $ts = _GUICtrlListView_GetItemTextString ( $lvInTabPartners )
					If StringReplace ( $ts, "|", "") <> "" Then
						GUICtrlSetData($strInfoLbl, "�б�������顷��ѡ�е��У�" & $ts )
					Else
						GUICtrlSetData($strInfoLbl, "" )
					EndIf
			EndSwitch
		Case GUICtrlGetHandle($lvInTabSchools)
            Switch $code
				Case $NM_DBLCLK
					$columnOfTable = $tb_schools
					$columnOfLV = $lvInTabSchools
					FuncCreateEditRecForColumn("ѧУ��Ϣ", $lvInTabSchools, $lParam )
				Case $NM_CLICK
					Local $ts = _GUICtrlListView_GetItemTextString ( $lvInTabSchools )
					If StringReplace ( $ts, "|", "") <> "" Then
						GUICtrlSetData($strInfoLbl, "�б�ѧУ��Ϣ����ѡ�е��У�" & $ts )
					Else
						GUICtrlSetData($strInfoLbl, "" )
					EndIf
			EndSwitch

		Case GUICtrlGetHandle($lvInTabUsers)
            Switch $code
				Case $NM_DBLCLK
					$columnOfTable = $tb_users
					$columnOfLV = $lvInTabUsers
					FuncCreateEditRecForColumn("�û�����", $lvInTabUsers, $lParam )
				Case $NM_CLICK
					Local $ts = _GUICtrlListView_GetItemTextString ( $lvInTabUsers )
					If StringReplace ( $ts, "|", "") <> "" Then
						GUICtrlSetData($strInfoLbl, "�б��û�������ѡ�е��У�" & $ts )
					Else
						GUICtrlSetData($strInfoLbl, "" )
					EndIf
			EndSwitch
		Case GUICtrlGetHandle($lvInTabAccets)
            Switch $code
				Case $NM_DBLCLK
					$columnOfTable = $tb_accets
					$columnOfLV = $lvInTabAccets
					FuncCreateEditRecForColumn("�ʲ�����", $lvInTabAccets, $lParam )
				Case $NM_CLICK
					Local $ts = _GUICtrlListView_GetItemTextString ( $lvInTabAccets )
					If StringReplace ( $ts, "|", "") <> "" Then
						GUICtrlSetData($strInfoLbl, "�б��ʲ�������ѡ�е��У�" & $ts )
					Else
						GUICtrlSetData($strInfoLbl, "" )
					EndIf
			EndSwitch
		Case GUICtrlGetHandle($lvInTabWorkers)
            Switch $code
				Case $NM_DBLCLK
					$columnOfTable = $tb_workers
					$columnOfLV = $lvInTabWorkers
					FuncCreateEditRecForColumn("��Ա����", $lvInTabWorkers, $lParam )
				Case $NM_CLICK
					Local $ts = _GUICtrlListView_GetItemTextString ( $lvInTabWorkers )
					If StringReplace ( $ts, "|", "") <> "" Then
						GUICtrlSetData($strInfoLbl, "�б���Ա������ѡ�е��У�" & $ts )
					Else
						GUICtrlSetData($strInfoLbl, "" )
					EndIf
			EndSwitch
		Case GUICtrlGetHandle($lvInTabLog)
            Switch $code
				Case $NM_DBLCLK
					$columnOfTable = $tb_log
					$columnOfLV = $lvInTabLog
					FuncCreateEditRecForColumn("������־", $lvInTabLog, $lParam )
				Case $NM_CLICK
					Local $ts = _GUICtrlListView_GetItemTextString ( $lvInTabLog )
					If StringReplace ( $ts, "|", "") <> "" Then
						GUICtrlSetData($strInfoLbl, "�б�������־����ѡ�е��У�" & $ts )
					Else
						GUICtrlSetData($strInfoLbl, "" )
					EndIf
			EndSwitch
		Case GUICtrlGetHandle($lvInTabSources)
            Switch $code
				Case $NM_DBLCLK
					$columnOfTable = $tb_source
					$columnOfLV = $lvInTabSources
					FuncCreateEditRecForColumn("Ԫ �� ��", $lvInTabSources, $lParam )
				Case $NM_CLICK
					Local $ts = _GUICtrlListView_GetItemTextString ( $lvInTabSources )
					If StringReplace ( $ts, "|", "") <> "" Then
						GUICtrlSetData($strInfoLbl, "�б�Ԫ �� �ݡ���ѡ�е��У�" & $ts )
					Else
						GUICtrlSetData($strInfoLbl, "" )
					EndIf
			EndSwitch

	EndSwitch

    Return $GUI_RUNDEFMSG
EndFunc

; ��LV�ϵ��Ҽ��˵�
Func Func_MouseMenuItem_LV ()
	Switch @GUI_CtrlId
		Case $mouseMenuItemDelLV
			; ʹ��ȫ�ֱ���
			; $columnOfTable��	������
			; $columnOfLV��		�б�ID
			; $columnOldData��	���е�����
			FuncDeleteItemFromAccessAndListView ( $columnOfTable, $columnOfLV, $columnOldData)
	EndSwitch

	; $mouseMenuLV, $mouseMenuItemDelLV, $mouseMenuItemUpdateLV, $mouseMenuItemCopyLV
	GUICtrlDelete($mouseMenuItemDelLV)
	GUICtrlDelete($mouseMenuItemUpdateLV)
	GUICtrlDelete($mouseMenuItemCopyLV)
	GUICtrlDelete($mouseMenuLV)
EndFunc

;#cs
; �ڵ�ǰ�б��в�ѯ����
; �ۼӲ�ͬ���ֶμ���Ϊ��ѯ����
; ��$strArgsOfLVArr�б���ؼ��ֿؼ���ID
; ��$strArgsOfTableArr�б������
;#ce
Func FuncFindData()
	Local $strLvName
	$strLvName = _GUICtrlTab_GetItemText ( $tabInMainWindow, _GUICtrlTab_GetCurSel ( $tabInMainWindow ))

	Local $popWinWidth = 500, $popWinHeight = 180

	$popupWindow = GUICreate("" , 500, 175, Default, Default, $WS_CAPTION, BitOR($WS_EX_TOPMOST, $GUI_WS_EX_PARENTDRAG))
		WinSetTitle($popupWindow, "", "�ڡ�" & $strLvName & "���в�ѯ����" )

		GUICtrlCreateGroup("����Ҫ������������ؼ��֡�������ÿո�ֿ�", 10, 10, $popWinWidth - 20, $popWinHeight - 45)

		; ��ȡ��ǰ�����LV
		Switch $strLvName
			Case "������־"		;  ��Ա ���� �ص� ��ͨ ʳ�� ����������id, j_name, j_date, j_address, j_traffic, j_board, j_content, j_note, j_record, j_date_record
				$strArgsOfLVArr[0] = $lvInTabJournal	; LV
				$strArgsOfTableArr[0] = $tb_journal		; Table

				GUICtrlCreateLabel("��Ա:", 20, 45, 30, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 52, 40, 180, 20)
				$strArgsOfTableArr[1] = "j_name"
				GUICtrlCreateLabel("����:", 20, 75, 30, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 52, 70, 180, 20)
				$strArgsOfTableArr[2] = "j_date"
				GUICtrlCreateLabel("�ص�:", 20, 105, 30, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 52, 100, 180, 20)
				$strArgsOfTableArr[3] = "j_address"

				GUICtrlCreateLabel("��ͨ:", 260, 45, 30, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 292, 40, 180, 20)
				$strArgsOfTableArr[4] = "j_traffic"
				GUICtrlCreateLabel("ʳ��:", 260, 75, 30, 20 )
				$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 292, 70, 180, 20)
				$strArgsOfTableArr[5] = "j_board"
				GUICtrlCreateLabel("����:", 260, 105, 30, 20 )
				$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 292, 100, 180, 20)
				$strArgsOfTableArr[6] = "j_content"

			Case "��˾��Ʒ"		; ��Ʒ ���� ��� ���� ��ۡ�id, pd_name, pd_type, pd_desiger, pd_configuration, pd_cost, pd_note
				$strArgsOfLVArr[0] = $lvInTabProducts
				$strArgsOfTableArr[0] = $tb_products

				GUICtrlCreateLabel("��Ʒ:", 20, 45, 30, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 52, 40, 180, 20)
				$strArgsOfTableArr[1] = "pd_name"
				GUICtrlCreateLabel("����:", 20, 75, 30, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 52, 70, 180, 20)
				$strArgsOfTableArr[2] = "pd_type"
				GUICtrlCreateLabel("���:", 20, 105, 30, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 52, 100, 180, 20)
				$strArgsOfTableArr[3] = "pd_desiger"

				GUICtrlCreateLabel("����:", 260, 45, 30, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 292, 40, 180, 20)
				$strArgsOfTableArr[4] = "pd_configuration"
				GUICtrlCreateLabel("���:", 260, 75, 30, 20 )
				$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 292, 70, 180, 20)
				$strArgsOfTableArr[5] = "pd_cost"
			Case "ѧУ��Ϣ"		; ѧУ ��ϵ�� ��ַ �绰 ���䡣id, s_name, s_contact, s_address, s_phone, s_email, s_note
				$strArgsOfLVArr[0] = $lvInTabSchools
				$strArgsOfTableArr[0] = $tb_schools

				GUICtrlCreateLabel("ѧ  У:", 20, 45, 45, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 67, 40, 170, 20)
				$strArgsOfTableArr[1] = "s_name"
				GUICtrlCreateLabel("��ϵ��:", 20, 75, 45, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 67, 70, 170, 20)
				$strArgsOfTableArr[2] = "s_contact"
				GUICtrlCreateLabel("��  ַ:", 20, 105, 45, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 67, 100, 170, 20)
				$strArgsOfTableArr[3] = "s_address"

				GUICtrlCreateLabel("��  ��:", 260, 45, 45, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 307, 40, 170, 20)
				$strArgsOfTableArr[4] = "s_phone"
				GUICtrlCreateLabel("��  ��:", 260, 75, 45, 20 )
				$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 307, 70, 170, 20)
				$strArgsOfTableArr[5] = "s_email"
			Case "���̹���"		; ���� ѧУ ��Ʒ ������ ״̬��id, pj_name, pj_s_name, pj_pd_name, pj_pt_name, pj_w_name, pj_content, pj_date_start, pj_state, pj_date_finish, pj_account, pj_note
				$strArgsOfLVArr[0] = $lvInTabProjects
				$strArgsOfTableArr[0] = $tb_projects

				GUICtrlCreateLabel("��  ��:", 20, 45, 45, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 67, 40, 170, 20)
				$strArgsOfTableArr[1] = "pj_name"
				GUICtrlCreateLabel("ѧ  У:", 20, 75, 45, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 67, 70, 170, 20)
				$strArgsOfTableArr[2] = "pj_s_name"
				GUICtrlCreateLabel("��  Ʒ:", 20, 105, 45, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 67, 100, 170, 20)
				$strArgsOfTableArr[3] = "pj_pt_name"

				GUICtrlCreateLabel("������:", 260, 45, 45, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 307, 40, 170, 20)
				$strArgsOfTableArr[4] = "pj_content"
				GUICtrlCreateLabel("״  ̬:", 260, 75, 45, 20 )
				$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 307, 70, 170, 20)
				$strArgsOfTableArr[5] = "pj_state"
			Case "�������"		; ���� ���� ��ַ �绰 ���� ҵ��id, pt_name, pt_type, pt_address, pt_phone, pt_email, pt_business, pt_note
				$strArgsOfLVArr[0] = $lvInTabPartners
				$strArgsOfTableArr[0] = $tb_partners

				GUICtrlCreateLabel("����:", 20, 45, 30, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 52, 40, 180, 20)
				$strArgsOfTableArr[1] = "pt_name"
				GUICtrlCreateLabel("����:", 20, 75, 30, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 52, 70, 180, 20)
				$strArgsOfTableArr[2] = "pt_type"
				GUICtrlCreateLabel("��ַ:", 20, 105, 30, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 52, 100, 180, 20)
				$strArgsOfTableArr[3] = "pt_address"

				GUICtrlCreateLabel("�绰:", 260, 45, 30, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 292, 40, 180, 20)
				$strArgsOfTableArr[4] = "pt_phone"
				GUICtrlCreateLabel("����:", 260, 75, 30, 20 )
				$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 292, 70, 180, 20)
				$strArgsOfTableArr[5] = "pt_email"
				GUICtrlCreateLabel("ҵ��:", 260, 105, 30, 20 )
				$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 292, 100, 180, 20)
				$strArgsOfTableArr[6] = "pt_business"

			Case "�û�����"
				;���|�û���|����|Ȩ��|��ע
				; id, u_name, u_pswd, u_authority, u_note
				$strArgsOfLVArr[0] = $lvInTabUsers
				$strArgsOfTableArr[0] = $tb_users

				GUICtrlCreateLabel("�û���:", 20, 35, 45, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 72, 30, 180, 20)
				$strArgsOfTableArr[1] = "u_name"
				GUICtrlCreateLabel("��  ��:", 20, 60, 45, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 72, 55, 180, 20)
				$strArgsOfTableArr[2] = "u_pswd"
				GUICtrlCreateLabel("Ȩ  ��:", 20, 90, 45, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 72, 85, 180, 20)
				$strArgsOfTableArr[3] = "u_authority"
				GUICtrlCreateLabel("��  ע:", 20, 120, 45, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 72, 115, 180, 20)
				$strArgsOfTableArr[4] = "u_note"

			Case "�ʲ�����"
				;���|����|����|��λ|����|��������|����|��������|������|�Ƿ񱨷�|��ע
				; id, a_name, a_serial_number, a_unit, a_type, a_date_bought, a_price, a_deportment, a_dealer, a_scrap, a_note
				$strArgsOfLVArr[0] = $lvInTabAccets	; LV
				$strArgsOfTableArr[0] = $tb_accets		; Table

				GUIDelete ( $popupWindow )

				$popWinHeight = 210

				$popupWindow = GUICreate("" , $popWinWidth, $popWinHeight, Default, Default, $WS_CAPTION, BitOR($WS_EX_TOPMOST, $GUI_WS_EX_PARENTDRAG))
					WinSetTitle($popupWindow, "", "�ڡ�" & $strLvName & "���в�������" )

					GUICtrlCreateGroup("����Ҫ������������ؼ��֡�������ÿո�ֿ�", 10, 5, 480, $popWinHeight - 45)

					GUICtrlCreateLabel("��    ��:", 20, 25, 60, 20 )
					$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 82, 20, 150, 20)
					$strArgsOfTableArr[1] = "a_name"
					GUICtrlCreateLabel("��    ��:", 20, 55, 60, 20 )
					$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 82, 50, 150, 20)
					$strArgsOfTableArr[2] = "a_serial_number"
					GUICtrlCreateLabel("��    λ:", 20, 85, 60, 20 )
					$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 82, 80, 150, 20)
					$strArgsOfTableArr[3] = "a_unit"
					GUICtrlCreateLabel("��    ��:", 20, 115, 60, 20 )
					$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 82, 110, 150, 20)
					$strArgsOfTableArr[4] = "a_type"
					GUICtrlCreateLabel("��������:", 20, 145, 60, 20 )
					$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 82, 140, 150, 20)
					$strArgsOfTableArr[5] = "a_date_bought"

					GUICtrlCreateLabel("��    ��:", 260, 25, 60, 20 )
					$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 322, 20, 150, 20)
					$strArgsOfTableArr[6] = "a_price"
					GUICtrlCreateLabel("��������:", 260, 55, 60, 20 )
					$strArgsOfLVArr[7] = GUICtrlCreateInput ( "", 322, 50, 150, 20)
					$strArgsOfTableArr[7] = "a_deportment"
					GUICtrlCreateLabel("�� �� ��:", 260, 85, 60, 20 )
					$strArgsOfLVArr[8] = GUICtrlCreateInput ( "", 322, 80, 150, 20)
					$strArgsOfTableArr[8] = "a_dealer"
					GUICtrlCreateLabel("�Ƿ񱨷�:", 260, 115, 60, 20 )
					$strArgsOfLVArr[9] = GUICtrlCreateInput ( "", 322, 110, 150, 20)
					$strArgsOfTableArr[9] = "a_scrap"
					GUICtrlCreateLabel("��    ע:", 260, 145, 60, 20 )
					$strArgsOfLVArr[10] = GUICtrlCreateInput ( "", 322, 140, 150, 20)
					$strArgsOfTableArr[10] = "a_note"

			Case "��Ա����"
				; ����|���֤��|����|ְ��|��ְ����|ת������|��н��|�Ա�|����|�绰|����|סַ|��ͥ��Ϣ|��ע
				; w_name, w_identity_number, w_deportment, w_position, w_date_of_entry, w_date_regular, w_month_wage, w_sex, w_birthday, w_phone, w_email, w_address, w_family, w_note
				$strArgsOfLVArr[0] = $lvInTabWorkers
				$strArgsOfTableArr[0] = $tb_workers

				GUIDelete ( $popupWindow )

				$popWinHeight = 270

				$popupWindow = GUICreate("" , $popWinWidth, $popWinHeight, Default, Default, $WS_CAPTION, BitOR($WS_EX_TOPMOST, $GUI_WS_EX_PARENTDRAG))
					WinSetTitle($popupWindow, "", "�ڡ�" & $strLvName & "���в�������" )

					GUICtrlCreateGroup("����Ҫ������������ؼ��֡�������ÿո�ֿ�", 10, 5, 480, $popWinHeight - 45)

					GUICtrlCreateLabel("��    ��:", 20, 25, 60, 20 )
					$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 82, 20, 150, 20)
					$strArgsOfTableArr[1] = "w_name"
					GUICtrlCreateLabel("���֤��:", 20, 55, 60, 20 )
					$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 82, 50, 150, 20)
					$strArgsOfTableArr[2] = "w_identity_number"
					GUICtrlCreateLabel("��    ��:", 20, 85, 60, 20 )
					$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 82, 80, 150, 20)
					$strArgsOfTableArr[3] = "w_deportment"
					GUICtrlCreateLabel("ְ    ��:", 20, 115, 60, 20 )
					$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 82, 110, 150, 20)
					$strArgsOfTableArr[4] = "w_position"
					GUICtrlCreateLabel("��ְ����:", 20, 145, 60, 20 )
					$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 82, 140, 150, 20)
					$strArgsOfTableArr[5] = "w_date_of_entry"
					GUICtrlCreateLabel("ת������:", 20, 175, 60, 20 )
					$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 82, 170, 150, 20)
					$strArgsOfTableArr[6] = "w_date_regular"
					GUICtrlCreateLabel("�� н ��:", 20, 205, 60, 20 )
					$strArgsOfLVArr[7] = GUICtrlCreateInput ( "", 82, 200, 150, 20)
					$strArgsOfTableArr[7] = "w_month_wage"

					GUICtrlCreateLabel("��    ��:", 260, 25, 60, 20 )
					$strArgsOfLVArr[8] = GUICtrlCreateInput ( "", 322, 20, 150, 20)
					$strArgsOfTableArr[8] = "w_sex"
					GUICtrlCreateLabel("��    ��:", 260, 55, 60, 20 )
					$strArgsOfLVArr[9] = GUICtrlCreateInput ( "", 322, 50, 150, 20)
					$strArgsOfTableArr[9] = "w_birthday"
					GUICtrlCreateLabel("��    ��:", 260, 85, 60, 20 )
					$strArgsOfLVArr[10] = GUICtrlCreateInput ( "", 322, 80, 150, 20)
					$strArgsOfTableArr[10] = "w_phone"
					GUICtrlCreateLabel("��    ��:", 260, 115, 60, 20 )
					$strArgsOfLVArr[11] = GUICtrlCreateInput ( "", 322, 110, 150, 20)
					$strArgsOfTableArr[11] = "w_email"
					GUICtrlCreateLabel("ס    ַ:", 260, 145, 60, 20 )
					$strArgsOfLVArr[12] = GUICtrlCreateInput ( "", 322, 140, 150, 20)
					$strArgsOfTableArr[12] = "w_address"
					GUICtrlCreateLabel("��ͥ��Ϣ:", 260, 175, 60, 20 )
					$strArgsOfLVArr[13] = GUICtrlCreateInput ( "", 322, 170, 150, 20)
					$strArgsOfTableArr[13] = "w_family"
					GUICtrlCreateLabel("��    ע:", 260, 205, 60, 20 )
					$strArgsOfLVArr[14] = GUICtrlCreateInput ( "", 322, 200, 150, 20)
					$strArgsOfTableArr[14] = "w_note"

			Case "������־"
				; �û�|��½�˳���ɾ��|��|��|������|������|����|��ע 8
				; l_name, l_operate, l_table, l_column, l_old_data, l_new_data, l_date, l_note
				$strArgsOfLVArr[0] = $lvInTabLog
				$strArgsOfTableArr[0] = $tb_log

				GUICtrlCreateLabel("�û�:", 20, 30, 30, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 52, 25, 180, 20)
				$strArgsOfTableArr[1] = "l_name"
				GUICtrlCreateLabel("����:", 20, 60, 30, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 52, 55, 180, 20)
				$strArgsOfTableArr[2] = "l_operate"
				GUICtrlCreateLabel("��:", 20, 90, 30, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 52, 85, 180, 20)
				$strArgsOfTableArr[3] = "l_table"
				GUICtrlCreateLabel("����:", 20, 120, 30, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 52, 115, 180, 20)
				$strArgsOfTableArr[4] = "l_column"

				GUICtrlCreateLabel("��ֵ:", 260, 30, 30, 20 )
				$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 292, 25, 180, 20)
				$strArgsOfTableArr[5] = "l_old_data"
				GUICtrlCreateLabel("��ֵ:", 260, 60, 30, 20 )
				$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 292, 55, 180, 20)
				$strArgsOfTableArr[6] = "l_new_data"
				GUICtrlCreateLabel("����:", 260, 90, 30, 20 )
				$strArgsOfLVArr[7] = GUICtrlCreateInput ( "", 292, 85, 180, 20)
				$strArgsOfTableArr[7] = "l_date"
				GUICtrlCreateLabel("��ע:", 260, 120, 30, 20 )
				$strArgsOfLVArr[8] = GUICtrlCreateInput ( "", 292, 115, 180, 20)
				$strArgsOfTableArr[8] = "l_note"

			Case "Ԫ �� ��"
				; ��|��|ֵ|���� 4
				; sr_tb_name, sr_column, sr_value, sr_note
				$strArgsOfLVArr[0] = $lvInTabSources
				$strArgsOfTableArr[0] = $tb_source

				GUICtrlCreateLabel("��:", 20, 30, 30, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 52, 25, 180, 20)
				$strArgsOfTableArr[1] = "sr_tb_name"
				GUICtrlCreateLabel("����:", 20, 60, 30, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 52, 55, 180, 20)
				$strArgsOfTableArr[2] = "sr_column"
				GUICtrlCreateLabel("ѡֵ:", 20, 90, 30, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 52, 85, 180, 20)
				$strArgsOfTableArr[3] = "sr_value"
				GUICtrlCreateLabel("��ע:", 20, 120, 30, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 52, 115, 180, 20)
				$strArgsOfTableArr[4] = "sr_note"

		EndSwitch

		GUICtrlCreateGroup("", -99, -99, 1, 1)

		$btnOKInPopWin = GUICtrlCreateButton("ȷ��", 10, $popWinHeight - 30, 60, 22)
			GUICtrlSetOnEvent($btnOKInPopWin, "FuncFindDataByArgs")
		$btnNOInPopWin = GUICtrlCreateButton("ȡ��", 428, $popWinHeight - 30, 60, 22)
			GUICtrlSetOnEvent($btnNOInPopWin, "FuncFindDataByArgs")

	GUISetState(@SW_SHOW, $popupWindow)
	GUISetState(@SW_DISABLE, $guiMainWindow)
EndFunc

Func FuncFindDataByArgs ()
	If @GUI_CtrlId = $btnOKInPopWin Then
		Local $isSqlStrEnable = False	; ��ѯ����Ƿ���á����ȫ���������ǿյģ�������Ϊtrue

		$strArgsOfSql = "SELECT * FROM " & $strArgsOfTableArr[0] & " WHERE "

		; �����ؼ����飬ƴ�����
		For $iiii = 1 To 11 Step 1	; 0�����������lV��Access���ơ�ʣ�� 1-11 ����ؼ�

			Local $strWidget = GUICtrlRead($strArgsOfLVArr[$iiii])

			If $strWidget <> "" And StringIsSpace($strWidget) <> 1 Then

				; ���ؼ��ڵ��ı���ȥ���ס�β�ո���������ո��Ϊ1��
				Local $keywords = StringStripWS($strWidget, $STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES )
				; ���ݹؼ����ı�֮��Ŀո���в���
				Local $keywordArr = StringSplit ($keywords, " ")
				; ֻҪ�д����������һ��Ԫ��[1]
				If $keywordArr[0] > 0 Then	; $keywordArr[0]������ЧԪ������
					; ����ж���ʣ�ƴ��
					For $j = 1 To $keywordArr[0]
						; �� like %ֵ%
						$strArgsOfSql &= $strArgsOfTableArr[$iiii] & " LIKE '%" & $keywordArr[$j] & "%' AND "
					Next
				EndIf

				$isSqlStrEnable = True	; ������һ����Ч�������ſ���
			EndIf
		Next

		$strArgsOfSql &= "1 = 1"	; ƴ��sql��䳣�ý�β������ɾ�����һ�� AND�Ƚ��鷳

		If $isSqlStrEnable Then
			; ִ�����ݿ��ѯ
			Local $adoCon = ObjCreate("ADODB.Connection")
			$adoCon.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & $db_path & "; Jet OLEDB:Database Password='" & $db_pswd & "'")
			Local $adoRec = ObjCreate("ADODB.Recordset")
			$adoRec.ActiveConnection = $adoCon

			$adoRec.Open($strArgsOfSql)	; �ӱ��в鵽�Ľ����

			Local $fieldsCount = $adoRec.fields.count	; ִ����һ���󣬵õ�����ֶ�����
			If $strArgsOfTableArr[0] = $tb_journal Then
				$fieldsCount = $fieldsCount - 2
			EndIf

			; ����б�
			_GUICtrlListView_DeleteAllItems ($strArgsOfLVArr[0])

			_GUICtrlListView_BeginUpdate($strArgsOfLVArr[0])
			While Not $adoRec.Eof And Not $adoRec.Bof	; ���������ÿһ��
				If @error = 1 Then
					ExitLoop
				Else
					; ��������Ǵ���һ���µı�ǩ��LV��չʾ��ѯ���

					Local $strResultDataForLv = ""

					For $i = 0 To $fieldsCount - 1 Step 1

						Local $strTmpS = $adoRec.fields( $i ).value
						If $strArgsOfTableArr[0] = "tb_log" Then
							$strTmpS = StringReplace ( $strTmpS, "|", "/")
						EndIf

						$strResultDataForLv &= $strTmpS & "|"
					Next

					GUICtrlCreateListViewItem($strResultDataForLv, $strArgsOfLVArr[0])
						GUICtrlSetBkColor (-1, 0xffa500 );����listviewitem�ı���ɫ

					$adoRec.Movenext	; ���������һ��
				EndIf
			WEnd
			_GUICtrlListView_EndUpdate($strArgsOfLVArr[0])
			_GUICtrlListView_Scroll($strArgsOfLVArr[0], 0, _GUICtrlListView_GetItemCount($strArgsOfLVArr[0])*10)

			GUICtrlSetData($strInfoLbl, "��" & $strArgsOfTableArr[0] & "���в�ѯ���ļ�¼������" & _GUICtrlListView_GetItemCount ( $strArgsOfLVArr[0] ))

			$adoRec.Close
			$adoCon.Close
		EndIf

	EndIf

	GUIDelete ( $popupWindow )
	GUISetState(@SW_ENABLE, $guiMainWindow)
	GUISetState(@SW_RESTORE, $guiMainWindow)

	; ��λ���������
	$strArgsOfSql = ""
	For $i = 0 To 14
		$strArgsOfLVArr[$i] = ""
		$strArgsOfTableArr[$i] = ""
	Next
EndFunc

; #cs
; 	˫��LV�ĵ�Ԫ��󣬴���һ�������
; 	��ԭ���ݽ����޸�
; 	���� $strLvInTab:Ҫ�������б�ID
; 	���� $lParam:ϵͳ��Ϣ
; #ce
Func FuncCreateEditRecForColumn($tmplvname, $strLvInTab, $lParam )
	; �����ṹ�����ݣ��ṹ�壬�ṩ���ݵ�ָ�룩
	Local $tInfo = DllStructCreate($tagNMITEMACTIVATE, $lParam)
	$hItemRow = DllStructGetData($tInfo, "Index")		; ������С�ʼ��0
	$hItemColumn = DllStructGetData($tInfo, "SubItem")	; ������С�ʼ��0

	Local $oADO = ObjCreate("ADODB.Connection")
	$oADO.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & $db_path & "; Jet OLEDB:Database Password='" & $db_pswd & "'")
	Local $oREC = ObjCreate("ADODB.Recordset")
	$oREC.ActiveConnection = $oADO
	$oREC.Open($columnOfTable, $oADO, 3, 3)	; ������˫����LV�����ı�
	Local $fieldsCount = $oREC.fields.count	; ��ȡ�����������ܹ��ж����ֶ�
	; $oREC.fields[i].name	; ����

	; ȫ�ֱ�����_WM_COMMAND���õ�
	$columnName = $oREC.fields($hItemColumn).name		; �õ���Ԫ���������ע����С����
	; $oREC.fields(1).value	; ��ֵ

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

	GUICtrlSetData($strInfoLbl, "�б�" & $tmplvname & "����ѡ�е��б�ţ�" & _GUICtrlListView_GetItemText($strLvInTab, $hItemRow) & "��ѡ�е��У�" & _GUICtrlListView_GetColumn($strLvInTab, $hItemColumn)[5] & "��ԭ���ݣ�" & $columnOldData)
EndFunc

; �Ҽ��˵� WM_CONTEXTMENU ��Ϣ
Func _WM_CONTEXTMENU($hWnd, $iMsg, $iwParam, $ilParam)
	Switch $iwParam
		Case $hLvInTabJournal, $hLvInTabAccets, $hLvInTabPartners, $hLvInTabProducts, $hLvInTabProjects, $hLvInTabSchools, $hLvInTabSources, $hLvInTabWorkers, $hLvInTabUsers, $hlvintabLog

			; ��ȡ�Ҽ��������Ч����
			If _GUICtrlListView_GetSelectedIndices ( $iwParam ) <> "" Then
				$columnOldData = _GUICtrlListView_GetItemTextString ($iwParam, _GUICtrlListView_GetSelectedIndices ( $iwParam ))
				Local $hMenu; $id_menu_lv_del = 2000, $id_menu_lv_update, $id_menu_lv_copy
				$hMenu = _GUICtrlMenu_CreatePopup()
				_GUICtrlMenu_InsertMenuItem($hMenu, 0, "ɾ��", $id_menu_lv_del)
				_GUICtrlMenu_InsertMenuItem($hMenu, 1, "", 0)
				_GUICtrlMenu_InsertMenuItem($hMenu, 2, "�޸�", $id_menu_lv_update)
				_GUICtrlMenu_InsertMenuItem($hMenu, 3, "", 0)
				_GUICtrlMenu_InsertMenuItem($hMenu, 4, "����", $id_menu_lv_copy)
				_GUICtrlMenu_TrackPopupMenu($hMenu, $hWnd)
				_GUICtrlMenu_DestroyMenu($hMenu)
				$hEnableListView = $iwParam
			EndIf

			Return True
	EndSwitch
EndFunc

; �ؼ�ʧȥ����
; LV���Ҽ��˵����
; ��Ԫ�������ݱ���
Func _WM_COMMAND($hWnd, $msg, $wParam, $lParam)

	; ƥ�������ĸ�LV���Ҽ��Ĳ˵�
	Switch $hEnableListView
		Case $hLvInTabJournal
			; ƥ��˵���
			Switch $wParam
				Case $id_menu_lv_del
					; $strTbName�� ������
					; $strLvInTab���б�ID
					; $columnOldData�� ���е�����
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

					; ����������������
					If $iText <> $columnOldData Then

						Local $tmpIndexOfDClick = _GUICtrlListView_GetItemText ( $columnOfLV, $hItemRow )

						;                               ��                             ��                ��ֵ             ����
						Local $strUpdate = "UPDATE " & $columnOfTable & " SET " & $columnName & "='" & $iText & "' WHERE id=" & $tmpIndexOfDClick
						Local $adoCon = ObjCreate("ADODB.Connection")
						$adoCon.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & $db_path & ";Jet Oledb:Database Password='" & $db_pswd & "'")
						$adoCon.execute($strUpdate)
						;$adoCon.close

						; ��־��
						Local $strLvItem = "'Eminem', 'update', '" & $columnOfTable & "', '" & $columnName & "', '" & $columnOldData & "', '" & $iText & "', '" & @YEAR & "-" & @MON & "-" & @MDAY & "(" & @WDAY - 1 & ")" & @HOUR & ":" & @MIN & ":" & @SEC & "`" & @MSEC & "', '_none'"	; ����
						$strLvItem = "insert into " & $tb_log & " (l_name, l_operate, l_table, l_column, l_old_data, l_new_data, l_date, l_note) values ( " & $strLvItem & " )"				; �������
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

; #cs ɾ����ǰ�б���ѡ�����
; ��ȡ��ǩ������LV��������
; ɾ����һ��
; �ڶ������ݿ���ɾ����¼����Log���м�¼
; #ce
Func FuncDeleteItemFromListView()
	Local $strLvItem, $strLvInTab, $strTbName
	; ��ȡ��ǰ�����LV
	Switch _GUICtrlTab_GetItemText ( $tabInMainWindow, _GUICtrlTab_GetCurSel ( $tabInMainWindow ))
		Case "������־"
			; ��ȡѡ���е�����
			;$row = _GUICtrlListView_GetSelectedIndices ( $lvInTabJournal )			; ѡ�е���Ŀ������
			$strLvItem = _GUICtrlListView_GetItemTextString ( $lvInTabJournal, -1 )	; ѡ�е���Ŀ����������

			$strLvInTab = $lvInTabJournal	; LV
			$strTbName = $tb_journal		; Table
		Case "��˾��Ʒ"
			$strLvItem = _GUICtrlListView_GetItemTextString ( $lvInTabProducts, -1 )

			$strLvInTab = $lvInTabProducts
			$strTbName = $tb_products
		Case "ѧУ��Ϣ"
			$strLvItem = _GUICtrlListView_GetItemTextString ( $lvInTabSchools, -1 )

			$strLvInTab = $lvInTabSchools
			$strTbName = $tb_schools
		Case "���̹���"
			$strLvItem = _GUICtrlListView_GetItemTextString ( $lvInTabProjects, -1 )

			$strLvInTab = $lvInTabProjects
			$strTbName = $tb_projects
		Case "�������"
			$strLvItem = _GUICtrlListView_GetItemTextString ( $lvInTabPartners, -1 )

			$strLvInTab = $lvInTabPartners
			$strTbName = $tb_partners

		Case "�û�����"
			$strLvItem = _GUICtrlListView_GetItemTextString ( $lvInTabUsers, -1 )

			$strLvInTab = $lvInTabUsers
			$strTbName = $tb_users

		Case "�ʲ�����"
			$strLvItem = _GUICtrlListView_GetItemTextString ( $lvInTabAccets, -1 )

			$strLvInTab = $lvInTabAccets
			$strTbName = $tb_accets

		Case "��Ա����"
			$strLvItem = _GUICtrlListView_GetItemTextString ( $lvInTabWorkers, -1 )

			$strLvInTab = $lvInTabWorkers
			$strTbName = $tb_workers

		Case "������־"
			$strLvItem = _GUICtrlListView_GetItemTextString ( $lvInTabLog, -1 )

			$strLvInTab = $lvInTabLog
			$strTbName = $tb_log

		Case "Ԫ �� ��"
			$strLvItem = _GUICtrlListView_GetItemTextString ( $lvInTabSources, -1 )

			$strLvInTab = $lvInTabSources
			$strTbName = $tb_source

	EndSwitch


	If _GUICtrlListView_GetSelectedCount ( $strLvInTab ) < 1 Then Return

	FuncDeleteItemFromAccessAndListView ( $strTbName, $strLvInTab, $strLvItem )
EndFunc

; ɾ��ѡ�е���
; $strTbName�� ������
; $strLvInTab���б�ID
; $strLvItem�� ���е�����
; ͬ�������ݿ�
Func FuncDeleteItemFromAccessAndListView ( $strTbName, $strLvInTab, $strLvItem )
	Local $indexOfDelItem = StringSplit($strLvItem, "|")

	Local $confirmDel = MsgBox(1 + 16, "����ȷ��", "��ȷ��ɾ����ѡ��Ŀ��")

	If $confirmDel = 1 Then
		_GUICtrlListView_DeleteItemsSelected ($strLvInTab)

		Local $dataAdodbConnectionDel = ObjCreate("ADODB.Connection")
		$dataAdodbConnectionDel.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & $db_path & ";Jet Oledb:Database Password='" & $db_pswd & "'")

		Local $strDel = "DELETE FROM " & $strTbName & " IN '" & $db_path & "' WHERE id" & " = " & $indexOfDelItem[1]

		$dataAdodbConnectionDel.execute($strDel)

		; ������ڡ���־����ɾ�����ݣ����Ǿ��Ե������¼����Ӧ���дβ���
		If $strTbName <> "tb_log" Then	; ������־��Ӧ�����кܸ�Ȩ�޵�
			; ��־��
			$strLvItem = "'Eminem', 'del', '" & $strTbName & "', '_item', '" & StringTrimLeft ($strLvItem, StringInStr( $strLvItem, "|")) & "', '_none', '" & @YEAR & "-" & @MON & "-" & @MDAY & "(" & @WDAY - 1 & ")" & @HOUR & ":" & @MIN & ":" & @SEC & "`" & @MSEC & "', '_none'"	; ����
			$strLvItem = "insert into " & $tb_log & " (l_name, l_operate, l_table, l_column, l_old_data, l_new_data, l_date, l_note) values ( " & $strLvItem & " )"				; �������
			$dataAdodbConnectionDel.Execute($strLvItem)
		EndIf

		$dataAdodbConnectionDel.Close

		GUICtrlSetData($strInfoLbl, "�б�" & $strTbName & "���м�¼������" & _GUICtrlListView_GetItemCount ( $strLvInTab ))
	EndIf

	$columnOfTable = ""
	$columnOfLV = ""
	$columnOldData = ""
EndFunc

; #cs
; ��Ӧ���������½�����ť��������Ŀ��ListView
; �����жϵ�ǰ����ı�ǩ-��һ��ListView������ʾ״̬
; Ȼ�������Ŀ��������Ŀλ��
; #ce
Func FuncInsertItemToListView ()
	Local $strLvName = _GUICtrlTab_GetItemText ( $tabInMainWindow, _GUICtrlTab_GetCurSel ( $tabInMainWindow ))
	Local $popWinWidth = 500, $popWinHeight = 180

	$popupWindow = GUICreate("" , $popWinWidth, $popWinHeight, Default, Default, $WS_CAPTION, BitOR($WS_EX_TOPMOST, $GUI_WS_EX_PARENTDRAG))
		WinSetTitle($popupWindow, "", "�ڡ�" & $strLvName & "���в�������" )

		GUICtrlCreateGroup("", 10, 5, 480, 135)

		; ��ȡ��ǰ�����LV
		Switch $strLvName
			Case "������־"
				; ��Ա ���� �ص� ��ͨ ʳ�� �������� ��ע 7
				;j_name, j_date, j_address, j_traffic, j_board, j_content, j_note, j_record, j_date_record
				$strArgsOfLVArr[0] = $lvInTabJournal	; LV
				$strArgsOfTableArr[0] = $tb_journal		; Table

				GUICtrlCreateLabel("��Ա:", 20, 25, 30, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 52, 20, 180, 20)
				$strArgsOfTableArr[1] = "j_name"
				GUICtrlCreateLabel("����:", 20, 55, 30, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 52, 50, 180, 20)
				$strArgsOfTableArr[2] = "j_date"
				GUICtrlCreateLabel("�ص�:", 20, 85, 30, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 52, 80, 180, 20)
				$strArgsOfTableArr[3] = "j_address"

				GUICtrlCreateLabel("��ע:", 20, 115, 30, 20 )
				$strArgsOfLVArr[7] = GUICtrlCreateInput ( "", 52, 110, 420, 20)
				$strArgsOfTableArr[7] = "j_note"

				GUICtrlCreateLabel("��ͨ:", 260, 25, 30, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 292, 20, 180, 20)
				$strArgsOfTableArr[4] = "j_traffic"
				GUICtrlCreateLabel("ʳ��:", 260, 55, 30, 20 )
				$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 292, 50, 180, 20)
				$strArgsOfTableArr[5] = "j_board"
				GUICtrlCreateLabel("����:", 260, 85, 30, 20 )
				$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 292, 80, 180, 20)
				$strArgsOfTableArr[6] = "j_content"

			Case "��˾��Ʒ"
				;��Ʒ ���� ��� ���� ��� ��ע 6
				;pd_name, pd_type, pd_desiger, pd_configuration, pd_cost, pd_note
				$strArgsOfLVArr[0] = $lvInTabProducts	; LV
				$strArgsOfTableArr[0] = $tb_products		; Table

				GUICtrlCreateLabel("��Ʒ:", 20, 45, 30, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 52, 40, 180, 20)
				$strArgsOfTableArr[1] = "pd_name"
				GUICtrlCreateLabel("����:", 20, 75, 30, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 52, 70, 180, 20)
				$strArgsOfTableArr[2] = "pd_type"
				GUICtrlCreateLabel("���:", 20, 105, 30, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 52, 100, 180, 20)
				$strArgsOfTableArr[3] = "pd_desiger"
				GUICtrlCreateLabel("����:", 260, 45, 30, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 292, 40, 180, 20)
				$strArgsOfTableArr[4] = "pd_configuration"
				GUICtrlCreateLabel("���:", 260, 75, 30, 20 )
				$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 292, 70, 180, 20)
				$strArgsOfTableArr[5] = "pd_cost"
				GUICtrlCreateLabel("��ע:", 260, 105, 30, 20 )
				$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 292, 100, 180, 20)
				$strArgsOfTableArr[6] = "pd_note"

			Case "ѧУ��Ϣ"
				;ѧУ ��ϵ�� ��ַ �绰 ���� ��ע 6
				;s_name, s_contact, s_address, s_phone, s_email, s_note
				$strArgsOfLVArr[0] = $lvInTabSchools	; LV
				$strArgsOfTableArr[0] = $tb_schools		; Table

				GUICtrlCreateLabel("ѧ  У:", 20, 45, 45, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 67, 40, 170, 20)
				$strArgsOfTableArr[1] = "s_name"
				GUICtrlCreateLabel("��ϵ��:", 20, 75, 45, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 67, 70, 170, 20)
				$strArgsOfTableArr[2] = "s_contact"
				GUICtrlCreateLabel("��  ַ:", 20, 105, 45, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 67, 100, 170, 20)
				$strArgsOfTableArr[3] = "s_address"

				GUICtrlCreateLabel("��  ��:", 260, 45, 45, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 307, 40, 170, 20)
				$strArgsOfTableArr[4] = "s_phone"
				GUICtrlCreateLabel("��  ��:", 260, 75, 45, 20 )
				$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 307, 70, 170, 20)
				$strArgsOfTableArr[5] = "s_email"
				GUICtrlCreateLabel("��  ע:", 260, 105, 45, 20 )
				$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 307, 100, 170, 20)
				$strArgsOfTableArr[6] = "s_note"

			Case "���̹���"
				;���� ѧУ ��Ʒ ������ ��˾������ ϸ����Ƭ�� ��ʼ���� ״̬ �������� �����¼ ��ע 11
				;pj_name, pj_s_name, pj_pd_name, pj_pt_name, pj_w_name, pj_content, pj_date_start, pj_state, pj_date_finish, pj_account, pj_note
				$strArgsOfLVArr[0] = $lvInTabProjects	; LV
				$strArgsOfTableArr[0] = $tb_projects		; Table

				GUIDelete ( $popupWindow )

				$popWinHeight = 300

				$popupWindow = GUICreate("" , $popWinWidth, $popWinHeight, Default, Default, $WS_CAPTION, BitOR($WS_EX_TOPMOST, $GUI_WS_EX_PARENTDRAG))
					WinSetTitle($popupWindow, "", "�ڡ�" & $strLvName & "���в�������" )

					GUICtrlCreateGroup("", 10, 5, 480, $popWinHeight - 45)

					GUICtrlCreateLabel("��    ��:", 20, 25, 60, 20 )
					$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 82, 20, 150, 20)
					$strArgsOfTableArr[1] = "pj_name"
					GUICtrlCreateLabel("ѧ    У:", 20, 55, 60, 20 )
					$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 82, 50, 150, 20)
					$strArgsOfTableArr[2] = "pj_s_name"
					GUICtrlCreateLabel("��    Ʒ:", 20, 85, 60, 20 )
					$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 82, 80, 150, 20)
					$strArgsOfTableArr[3] = "pj_pd_name"
					GUICtrlCreateLabel("�� �� ��:", 20, 115, 60, 20 )
					$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 82, 110, 150, 20)
					$strArgsOfTableArr[4] = "pj_pt_name"
					GUICtrlCreateLabel("�Ҹ�����:", 20, 145, 60, 20 )
					$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 82, 140, 150, 20)
					$strArgsOfTableArr[5] = "pj_w_name"

					GUICtrlCreateLabel("��    ע:", 20, 175, 60, 20 )
					$strArgsOfLVArr[11] = GUICtrlCreateInput ( "", 82, 170, 390, 80, BitOR($ES_MULTILINE, $ES_WANTRETURN, $ES_AUTOVSCROLL))
					$strArgsOfTableArr[11] = "pj_note"

					GUICtrlCreateLabel("Э��ϸ��:", 260, 25, 60, 20 )
					$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 322, 20, 150, 20)
					$strArgsOfTableArr[6] = "pj_content"
					GUICtrlCreateLabel("��������:", 260, 55, 60, 20 )
					$strArgsOfLVArr[7] = GUICtrlCreateInput ( "", 322, 50, 150, 20)
					$strArgsOfTableArr[7] = "pj_date_start"
					GUICtrlCreateLabel("��չ״̬:", 260, 85, 60, 20 )
					$strArgsOfLVArr[8] = GUICtrlCreateInput ( "", 322, 80, 150, 20)
					$strArgsOfTableArr[8] = "pj_state"
					GUICtrlCreateLabel("�깤����:", 260, 115, 60, 20 )
					$strArgsOfLVArr[9] = GUICtrlCreateInput ( "", 322, 110, 150, 20)
					$strArgsOfTableArr[9] = "pj_date_finish"
					GUICtrlCreateLabel("�����¼:", 260, 145, 60, 20 )
					$strArgsOfLVArr[10] = GUICtrlCreateInput ( "", 322, 140, 150, 20)
					$strArgsOfTableArr[10] = "pj_account"

			Case "�������"
				;���� ���� ��ַ �绰 ���� ҵ�� ��ע 7
				;pt_name, pt_type, pt_address, pt_phone, pt_email, pt_business, pt_note
				$strArgsOfLVArr[0] = $lvInTabPartners	; LV
				$strArgsOfTableArr[0] = $tb_partners	; Table

				GUICtrlCreateLabel("����:", 20, 25, 30, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 52, 20, 180, 20)
				$strArgsOfTableArr[1] = "pt_name"
				GUICtrlCreateLabel("����:", 20, 55, 30, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 52, 50, 180, 20)
				$strArgsOfTableArr[2] = "pt_type"
				GUICtrlCreateLabel("��ַ:", 20, 85, 30, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 52, 80, 180, 20)
				$strArgsOfTableArr[3] = "pt_address"

				GUICtrlCreateLabel("�绰:", 260, 25, 30, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 292, 20, 180, 20)
				$strArgsOfTableArr[4] = "pt_phone"
				GUICtrlCreateLabel("����:", 260, 55, 30, 20 )
				$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 292, 50, 180, 20)
				$strArgsOfTableArr[5] = "pt_email"
				GUICtrlCreateLabel("ҵ��:", 260, 85, 30, 20 )
				$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 292, 80, 180, 20)
				$strArgsOfTableArr[6] = "pt_business"

				GUICtrlCreateLabel("��ע:", 20, 115, 30, 20 )
				$strArgsOfLVArr[7] = GUICtrlCreateInput ( "", 52, 110, 420, 20)
				$strArgsOfTableArr[7] = "pt_note"

			Case "�û�����"
				; �û���|����|Ȩ��|��ע 4
				; u_name, u_pswd, u_authority, u_note
				$strArgsOfLVArr[0] = $lvInTabUsers	; LV
				$strArgsOfTableArr[0] = $tb_users	; Table

				GUICtrlCreateLabel("�û�:", 20, 25, 30, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 52, 20, 180, 20)
				$strArgsOfTableArr[1] = "u_name"
				GUICtrlCreateLabel("����:", 20, 55, 30, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 52, 50, 180, 20)
				$strArgsOfTableArr[2] = "u_pswd"
				GUICtrlCreateLabel("Ȩ��:", 20, 85, 30, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 52, 80, 180, 20)
				$strArgsOfTableArr[3] = "u_authority"
				GUICtrlCreateLabel("��ע:", 20, 115, 30, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 52, 110, 180, 20)
				$strArgsOfTableArr[4] = "u_note"


			Case "�ʲ�����"
				; ����|����|��λ|����|��������|����|��������|������|�Ƿ񱨷�|��ע 10
				; a_name, a_serial_number, a_unit, a_type, a_date_bought, a_price, a_deportment, a_dealer, a_scrap, a_note
				$strArgsOfLVArr[0] = $lvInTabAccets	; LV
				$strArgsOfTableArr[0] = $tb_accets		; Table

				GUIDelete ( $popupWindow )

				$popWinHeight = 210

				$popupWindow = GUICreate("" , $popWinWidth, $popWinHeight, Default, Default, $WS_CAPTION, BitOR($WS_EX_TOPMOST, $GUI_WS_EX_PARENTDRAG))
					WinSetTitle($popupWindow, "", "�ڡ�" & $strLvName & "���в�������" )

					GUICtrlCreateGroup("", 10, 5, 480, $popWinHeight - 45)

					GUICtrlCreateLabel("��    ��:", 20, 25, 60, 20 )
					$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 82, 20, 150, 20)
					$strArgsOfTableArr[1] = "a_name"
					GUICtrlCreateLabel("��    ��:", 20, 55, 60, 20 )
					$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 82, 50, 150, 20)
					$strArgsOfTableArr[2] = "a_serial_number"
					GUICtrlCreateLabel("��    λ:", 20, 85, 60, 20 )
					$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 82, 80, 150, 20)
					$strArgsOfTableArr[3] = "a_unit"
					GUICtrlCreateLabel("��    ��:", 20, 115, 60, 20 )
					$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 82, 110, 150, 20)
					$strArgsOfTableArr[4] = "a_type"
					GUICtrlCreateLabel("��������:", 20, 145, 60, 20 )
					$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 82, 140, 150, 20)
					$strArgsOfTableArr[5] = "a_date_bought"

					GUICtrlCreateLabel("��    ��:", 260, 25, 60, 20 )
					$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 322, 20, 150, 20)
					$strArgsOfTableArr[6] = "a_price"
					GUICtrlCreateLabel("��������:", 260, 55, 60, 20 )
					$strArgsOfLVArr[7] = GUICtrlCreateInput ( "", 322, 50, 150, 20)
					$strArgsOfTableArr[7] = "a_deportment"
					GUICtrlCreateLabel("�� �� ��:", 260, 85, 60, 20 )
					$strArgsOfLVArr[8] = GUICtrlCreateInput ( "", 322, 80, 150, 20)
					$strArgsOfTableArr[8] = "a_dealer"
					GUICtrlCreateLabel("�Ƿ񱨷�:", 260, 115, 60, 20 )
					$strArgsOfLVArr[9] = GUICtrlCreateInput ( "", 322, 110, 150, 20)
					$strArgsOfTableArr[9] = "a_scrap"
					GUICtrlCreateLabel("��    ע:", 260, 145, 60, 20 )
					$strArgsOfLVArr[10] = GUICtrlCreateInput ( "", 322, 140, 150, 20)
					$strArgsOfTableArr[10] = "a_note"

			Case "��Ա����"
				; ����|���֤��|����|ְ��|��ְ����|ת������|��н��|�Ա�|����|�绰|����|סַ|��ͥ��Ϣ|��ע
				; w_name, w_identity_number, w_deportment, w_position, w_date_of_entry, w_date_regular, w_month_wage, w_sex, w_birthday, w_phone, w_email, w_address, w_family, w_note
				$strArgsOfLVArr[0] = $lvInTabWorkers
				$strArgsOfTableArr[0] = $tb_workers

				GUIDelete ( $popupWindow )

				$popWinHeight = 270

				$popupWindow = GUICreate("" , $popWinWidth, $popWinHeight, Default, Default, $WS_CAPTION, BitOR($WS_EX_TOPMOST, $GUI_WS_EX_PARENTDRAG))
					WinSetTitle($popupWindow, "", "�ڡ�" & $strLvName & "���в�������" )

					GUICtrlCreateGroup("", 10, 5, 480, $popWinHeight - 45)

					GUICtrlCreateLabel("��    ��:", 20, 25, 60, 20 )
					$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 82, 20, 150, 20)
					$strArgsOfTableArr[1] = "w_name"
					GUICtrlCreateLabel("���֤��:", 20, 55, 60, 20 )
					$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 82, 50, 150, 20)
					$strArgsOfTableArr[2] = "w_identity_number"
					GUICtrlCreateLabel("��    ��:", 20, 85, 60, 20 )
					$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 82, 80, 150, 20)
					$strArgsOfTableArr[3] = "w_deportment"
					GUICtrlCreateLabel("ְ    ��:", 20, 115, 60, 20 )
					$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 82, 110, 150, 20)
					$strArgsOfTableArr[4] = "w_position"
					GUICtrlCreateLabel("��ְ����:", 20, 145, 60, 20 )
					$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 82, 140, 150, 20)
					$strArgsOfTableArr[5] = "w_date_of_entry"
					GUICtrlCreateLabel("ת������:", 20, 175, 60, 20 )
					$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 82, 170, 150, 20)
					$strArgsOfTableArr[6] = "w_date_regular"
					GUICtrlCreateLabel("�� н ��:", 20, 205, 60, 20 )
					$strArgsOfLVArr[7] = GUICtrlCreateInput ( "", 82, 200, 150, 20)
					$strArgsOfTableArr[7] = "w_month_wage"

					GUICtrlCreateLabel("��    ��:", 260, 25, 60, 20 )
					$strArgsOfLVArr[8] = GUICtrlCreateInput ( "", 322, 20, 150, 20)
					$strArgsOfTableArr[8] = "w_sex"
					GUICtrlCreateLabel("��    ��:", 260, 55, 60, 20 )
					$strArgsOfLVArr[9] = GUICtrlCreateInput ( "", 322, 50, 150, 20)
					$strArgsOfTableArr[9] = "w_birthday"
					GUICtrlCreateLabel("��    ��:", 260, 85, 60, 20 )
					$strArgsOfLVArr[10] = GUICtrlCreateInput ( "", 322, 80, 150, 20)
					$strArgsOfTableArr[10] = "w_phone"
					GUICtrlCreateLabel("��    ��:", 260, 115, 60, 20 )
					$strArgsOfLVArr[11] = GUICtrlCreateInput ( "", 322, 110, 150, 20)
					$strArgsOfTableArr[11] = "w_email"
					GUICtrlCreateLabel("ס    ַ:", 260, 145, 60, 20 )
					$strArgsOfLVArr[12] = GUICtrlCreateInput ( "", 322, 140, 150, 20)
					$strArgsOfTableArr[12] = "w_address"
					GUICtrlCreateLabel("��ͥ��Ϣ:", 260, 175, 60, 20 )
					$strArgsOfLVArr[13] = GUICtrlCreateInput ( "", 322, 170, 150, 20)
					$strArgsOfTableArr[13] = "w_family"
					GUICtrlCreateLabel("��    ע:", 260, 205, 60, 20 )
					$strArgsOfLVArr[14] = GUICtrlCreateInput ( "", 322, 200, 150, 20)
					$strArgsOfTableArr[14] = "w_note"

			Case "������־"
				; �û�|��½�˳���ɾ��|��|��|������|������|����|��ע 8
				; l_name, l_operate, l_table, l_column, l_old_data, l_new_data, l_date, l_note
				$strArgsOfLVArr[0] = $lvInTabLog
				$strArgsOfTableArr[0] = $tb_log

				GUICtrlCreateLabel("�û�:", 20, 25, 30, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 52, 20, 180, 20)
				$strArgsOfTableArr[1] = "l_name"
				GUICtrlCreateLabel("����:", 20, 55, 30, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 52, 50, 180, 20)
				$strArgsOfTableArr[2] = "l_operate"
				GUICtrlCreateLabel("��:", 20, 85, 30, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 52, 80, 180, 20)
				$strArgsOfTableArr[3] = "l_table"
				GUICtrlCreateLabel("����:", 20, 115, 30, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 52, 110, 180, 20)
				$strArgsOfTableArr[4] = "l_column"

				GUICtrlCreateLabel("��ֵ:", 260, 25, 30, 20 )
				$strArgsOfLVArr[5] = GUICtrlCreateInput ( "", 292, 20, 180, 20)
				$strArgsOfTableArr[5] = "l_old_data"
				GUICtrlCreateLabel("��ֵ:", 260, 55, 30, 20 )
				$strArgsOfLVArr[6] = GUICtrlCreateInput ( "", 292, 50, 180, 20)
				$strArgsOfTableArr[6] = "l_new_data"
				GUICtrlCreateLabel("����:", 260, 85, 30, 20 )
				$strArgsOfLVArr[7] = GUICtrlCreateInput ( "", 292, 80, 180, 20)
				$strArgsOfTableArr[7] = "l_date"
				GUICtrlCreateLabel("��ע:", 260, 115, 30, 20 )
				$strArgsOfLVArr[8] = GUICtrlCreateInput ( "", 292, 110, 180, 20)
				$strArgsOfTableArr[8] = "l_note"

			Case "Ԫ �� ��"
				; ��|��|ֵ|���� 4
				; sr_tb_name, sr_column, sr_value, sr_note
				$strArgsOfLVArr[0] = $lvInTabSources
				$strArgsOfTableArr[0] = $tb_source

				GUICtrlCreateLabel("��:", 20, 25, 30, 20 )
				$strArgsOfLVArr[1] = GUICtrlCreateInput ( "", 52, 20, 180, 20)
				$strArgsOfTableArr[1] = "sr_tb_name"
				GUICtrlCreateLabel("����:", 20, 55, 30, 20 )
				$strArgsOfLVArr[2] = GUICtrlCreateInput ( "", 52, 50, 180, 20)
				$strArgsOfTableArr[2] = "sr_column"
				GUICtrlCreateLabel("ѡֵ:", 20, 85, 30, 20 )
				$strArgsOfLVArr[3] = GUICtrlCreateInput ( "", 52, 80, 180, 20)
				$strArgsOfTableArr[3] = "sr_value"
				GUICtrlCreateLabel("��ע:", 20, 115, 30, 20 )
				$strArgsOfLVArr[4] = GUICtrlCreateInput ( "", 52, 110, 180, 20)
				$strArgsOfTableArr[4] = "sr_note"

	EndSwitch

	GUICtrlCreateGroup("", -99, -99, 1, 1)

	$btnOKInPopWin = GUICtrlCreateButton("ȷ��", 10, $popWinHeight - 30, 60, 22)
		GUICtrlSetOnEvent($btnOKInPopWin, "FuncInsertDatas")
	$btnNOInPopWin = GUICtrlCreateButton("ȡ��", 428, $popWinHeight - 30, 60, 22)
		GUICtrlSetOnEvent($btnNOInPopWin, "FuncInsertDatas")

	GUISetState(@SW_SHOW, $popupWindow)
	GUISetState(@SW_DISABLE, $guiMainWindow)
EndFunc

; #cs
; �����ݿ���д������Ŀ
; #ce
Func FuncInsertDatas ()

	If @GUI_CtrlId = $btnOKInPopWin Then
		Local $strColumns	; ƥ���������
		Local $isSqlStrEnable = False	; ��ѯ����Ƿ���á����ȫ���������ǿյģ�������Ϊtrue
		Local $tmpStep	; ����ѭ�����ȣ�ƥ�������
		Local $strTempLVName

		Switch $strArgsOfTableArr[0]	; ƥ�����ݿ��
			Case $tb_products
				$strTempLVName = "��˾��Ʒ"
				$tmpStep = 6
				$strColumns = " (pd_name, pd_type, pd_desiger, pd_configuration, pd_cost, pd_note) values ('"
			Case $tb_partners
				$strTempLVName = "�������"
				$tmpStep = 7
				$strColumns = " (pt_name, pt_type, pt_address, pt_phone, pt_email, pt_business, pt_note) values ('"
			Case $tb_projects
				$strTempLVName = "���̹���"
				$tmpStep = 11
				$strColumns = " (pj_name, pj_s_name, pj_pd_name, pj_pt_name, pj_w_name, pj_content, pj_date_start, pj_state, pj_date_finish, pj_account, pj_note) values ('"
			Case $tb_journal
				$strTempLVName = "������־"
				$tmpStep = 7
				$strColumns = " (j_name, j_date, j_address, j_traffic, j_board, j_content, j_note) values ('"	; , j_record, j_date_record
			Case $tb_schools
				$strTempLVName = "ѧУ��Ϣ"
				$tmpStep = 6
				$strColumns = " (s_name, s_contact, s_address, s_phone, s_email, s_note) values ('"

			Case $tb_users
				$strTempLVName = "�û�����"
				$tmpStep = 4
				$strColumns = " (u_name, u_pswd, u_authority, u_note) values ('"

			Case $tb_accets
				$strTempLVName = "�ʲ�����"
				$tmpStep = 10
				$strColumns = " (a_name, a_serial_number, a_unit, a_type, a_date_bought, a_price, a_deportment, a_dealer, a_scrap, a_note) values ('"

			Case $tb_workers
				$strTempLVName = "��Ա����"
				$tmpStep = 14
				$strColumns = " (w_name, w_identity_number, w_deportment, w_position, w_date_of_entry, w_date_regular, w_month_wage, w_sex, w_birthday, w_phone, w_email, w_address, w_family, w_note) values ('"

			Case $tb_log
				$strTempLVName = "������־"
				$tmpStep = 8
				$strColumns = " (l_name, l_operate, l_table, l_column, l_old_data, l_new_data, l_date, l_note) values ('"

			Case $tb_source
				$strTempLVName = "Ԫ �� ��"
				$tmpStep = 4
				$strColumns = " (sr_tb_name, sr_column, sr_value, sr_note) values ('"
				;------------------------------------------------------ todo --------------------------------------
		EndSwitch

		$strArgsOfSql = ""	; ����ƴ��

		; �����ؼ����飬ƴ�����������
		For $iiii = 1 To $tmpStep Step 1	; 0�����������lV��Access���ơ�ʣ�� 1-11 ����ؼ�

			Local $strWidget = GUICtrlRead($strArgsOfLVArr[$iiii])

			If $strWidget <> "" And StringIsSpace($strWidget) <> 1 Then	; ��������ȫ�ո������

				; ���ؼ��ڵ��ı���ȥ���ס�β�ո���������ո��Ϊ1��
				Local $keywords = StringStripWS($strWidget, $STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES )
				$strArgsOfSql &= $keywords & "', '"

				$isSqlStrEnable = True	; ������һ����Ч�������ſ���

			ElseIf $strWidget = "" Or StringIsSpace($strWidget) = 1 Then ; �����������գ��������Ŀո�
				$strArgsOfSql &= " " & "', '"

			EndIf
		Next

		$strArgsOfSql = StringTrimRight ( $strArgsOfSql, 4 )	; �������һ�����źͿո��һ�Ե����� ', '

		If $isSqlStrEnable Then

			Local $dataAdodbConnectionIn = ObjCreate("ADODB.Connection")
			$dataAdodbConnectionIn.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & $db_path & ";Jet Oledb:Database Password='" & $db_pswd & "'")
			Local $strSqlInsert = "insert into " & $strArgsOfTableArr[0] & $strColumns & $strArgsOfSql & "')"
			$dataAdodbConnectionIn.Execute($strSqlInsert)

			$strArgsOfSql = StringReplace ( $strArgsOfSql, "', '", "|" )

			; ��־��
			Local $strTmpSqlInsert = "'Eminem', 'add', '" & $strArgsOfTableArr[0] & "', '_item', '_none', '" & $strArgsOfSql & "', '" & @YEAR & "-" & @MON & "-" & @MDAY & "(" & @WDAY - 1 & ")" & @HOUR & ":" & @MIN & ":" & @SEC & "`" & @MSEC & "', '_none'"	; ����
			$strTmpSqlInsert = "insert into " & $tb_log & " (l_name, l_operate, l_table, l_column, l_old_data, l_new_data, l_date, l_note) values ( " & $strTmpSqlInsert & " )"	; �������
			$dataAdodbConnectionIn.Execute($strTmpSqlInsert)

			Local $adoRec = ObjCreate("ADODB.Recordset")
			$adoRec.ActiveConnection = $dataAdodbConnectionIn
			$adoRec.open("SELECT TOP 1 id FROM " & $strArgsOfTableArr[0] & " ORDER BY id DESC")	; ��ѯ����id�������Ѿ�ɾ���˼�¼��ID��Ȼ��������
			Local $recordcount = $adoRec.fields(0).value	; ֻ��һ����¼������ֱ��ȡ�����IDֵ�Ѿ� + 1

			$adoRec.Close
			$dataAdodbConnectionIn.Close

			; ���ﻹҪ���һ������š��е�LV
			; �����������ݿ�������ID+1������õ���ֵ�Ѿ�+1
			$strArgsOfSql = $recordcount & "|" & $strArgsOfSql

			Local $tmpLvItem = GUICtrlCreateListViewItem( $strArgsOfSql, $strArgsOfLVArr[0])
			GUICtrlSetBkColor ($tmpLvItem, 0xff9900 )

			_GUICtrlListView_Scroll($strArgsOfLVArr[0], 0, _GUICtrlListView_GetItemCount($strArgsOfLVArr[0])*15)

			GUICtrlSetData($strInfoLbl, "�б�" & $strTempLVName & "���м�¼������" & _GUICtrlListView_GetItemCount ( $strArgsOfLVArr[0] ))
		EndIf
	EndIf

	GUIDelete ( $popupWindow )
	GUISetState(@SW_ENABLE, $guiMainWindow)
	GUISetState(@SW_RESTORE, $guiMainWindow)

	; ��λ���������
	$strArgsOfSql = ""
	For $i = 0 To 14
		$strArgsOfLVArr[$i] = ""
		$strArgsOfTableArr[$i] = ""
	Next
EndFunc

; #cs
; �Ҽ������ǩ�������˵�
; ���Ҽ��˵������Ӧ
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
				Case "������־"
					GUICtrlSetData($menuItemJournal, "��������־")
				Case "��˾��Ʒ"
					GUICtrlSetData($menuItemProducts, "����˾��Ʒ")
				Case "ѧУ��Ϣ"
					GUICtrlSetData($menuItemSchools, "��ѧУ��Ϣ")
				Case "���̹���"
					GUICtrlSetData($menuItemProjects, "�����̹���")
				Case "�������"
					GUICtrlSetData($menuItemPartners, "���������")

				Case "�û�����"
					GUICtrlSetData($menuItemUsers, "  �û�����")
				Case "�ʲ�����"
					GUICtrlSetData($menuItemAccets, "  �ʲ�����")
				Case "��Ա����"
					GUICtrlSetData($menuItemWorkers, "  ��Ա����")
				Case "������־"
					GUICtrlSetData($menuItemLog, "  ������־")
				Case "Ԫ �� ��"
					GUICtrlSetData($menuItemSources, "  Ԫ �� ��")
			EndSwitch

			_GUICtrlTab_DeleteItem ( $tabInMainWindow, $idOfTabItem )

			$strTabItemText = _GUICtrlTab_GetItemText ( $tabInMainWindow, _GUICtrlTab_GetCurSel ( $tabInMainWindow ) )

			Switch $strTabItemText
				Case "������־"
					GUICtrlSetState($lvInTabJournal, $GUI_SHOW)
					GUICtrlSetData($strInfoLbl, "�б�������־���м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabJournal ))
				Case "��˾��Ʒ"
					GUICtrlSetState($lvInTabProducts, $GUI_SHOW)
					GUICtrlSetData($strInfoLbl, "�б���˾��Ʒ���м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabProducts ))
				Case "ѧУ��Ϣ"
					GUICtrlSetState($lvInTabSchools, $GUI_SHOW)
					GUICtrlSetData($strInfoLbl, "�б�ѧУ��Ϣ���м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabSchools ))
				Case "���̹���"
					GUICtrlSetState($lvInTabProjects, $GUI_SHOW)
					GUICtrlSetData($strInfoLbl, "�б����̹����м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabProjects ))
				Case "�������"
					GUICtrlSetState($lvInTabPartners, $GUI_SHOW)
					GUICtrlSetData($strInfoLbl, "�б�������顷�м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabPartners ))

				Case "�û�����"
					GUICtrlSetState($lvInTabUsers, $GUI_SHOW)
					GUICtrlSetData($strInfoLbl, "�б��û������м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabUsers ))
				Case "�ʲ�����"
					GUICtrlSetState($lvInTabAccets, $GUI_SHOW)
					GUICtrlSetData($strInfoLbl, "�б��ʲ������м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabAccets ))
				Case "��Ա����"
					GUICtrlSetState($lvInTabWorkers, $GUI_SHOW)
					GUICtrlSetData($strInfoLbl, "�б���Ա�����м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabWorkers ))
				Case "������־"
					GUICtrlSetState($lvInTabLog, $GUI_SHOW)
					GUICtrlSetData($strInfoLbl, "�б�������־���м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabLog ))
				Case "Ԫ �� ��"
					GUICtrlSetState($lvInTabSources, $GUI_SHOW)
					GUICtrlSetData($strInfoLbl, "�б�Ԫ �� �ݡ��м�¼������" & _GUICtrlListView_GetItemCount ( $lvInTabSources ))
			EndSwitch

			$idOfTabItem = -1

		Case $mouseMenuItemSaveAs

		Case $mouseMenuItemPrint

	EndSwitch
EndFunc

#cs
	ͨ���˵������Ʊ�ǩ����ʾ������
#ce
Func Func_ShowTab_ByMenu ()
	Switch @GUI_CtrlId
		Case $menuItemJournal
			Func_ShowTab_ByText ("������־")
		Case $menuItemProducts
			Func_ShowTab_ByText ("��˾��Ʒ")
		Case $menuItemSchools
			Func_ShowTab_ByText ("ѧУ��Ϣ")
		Case $menuItemProjects
			Func_ShowTab_ByText ("���̹���")
		Case $menuItemPartners
			Func_ShowTab_ByText ("�������")
		Case $menuItemUsers
			Func_ShowTab_ByText ("�û�����")
		Case $menuItemAccets
			Func_ShowTab_ByText ("�ʲ�����")
		Case $menuItemWorkers
			Func_ShowTab_ByText ("��Ա����")
		Case $menuItemLog
			Func_ShowTab_ByText ("������־")
		Case $menuItemSources
			Func_ShowTab_ByText ("Ԫ �� ��")
	EndSwitch
EndFunc

#cs ͨ��������ַ������ͱ�ǩͷ����ƥ�䡣������ʾ������
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
			Case "������־"
				GUICtrlSetData($menuItemJournal, "  ������־")
			Case "��˾��Ʒ"
				GUICtrlSetData($menuItemProducts, "  ��˾��Ʒ")
			Case "ѧУ��Ϣ"
				GUICtrlSetData($menuItemSchools, "  ѧУ��Ϣ")
			Case "���̹���"
				GUICtrlSetData($menuItemProjects, "  ���̹���")
			Case "�������"
				GUICtrlSetData($menuItemPartners, "  �������")

			Case "�û�����"
				GUICtrlSetData($menuItemUsers, "  �û�����")
			Case "�ʲ�����"
				GUICtrlSetData($menuItemAccets, "  �ʲ�����")
			Case "��Ա����"
				GUICtrlSetData($menuItemWorkers, "  ��Ա����")
			Case "������־"
				GUICtrlSetData($menuItemLog, "  ������־")
			Case "Ԫ �� ��"
				GUICtrlSetData($menuItemSources, "  Ԫ �� ��")
		EndSwitch

		_GUICtrlTab_DeleteItem ( $tabInMainWindow, $id_item )

		$strTabItemText = _GUICtrlTab_GetItemText ( $tabInMainWindow, _GUICtrlTab_GetCurSel ( $tabInMainWindow ) )
		Switch $strTabItemText
			Case "������־"
				If $hasReadedDbTbJournal = False Then
					FuncReadDb($tb_journal, $lvInTabJournal)
					$hasReadedDbTbJournal = True
				EndIf

				GUICtrlSetState($lvInTabJournal, $GUI_SHOW)
				GUICtrlSetData($menuItemJournal, "�� ������־")
			Case "��˾��Ʒ"
				If $hasReadedDbTbProducts = False Then
					FuncReadDb($tb_products, $lvInTabProducts)
					$hasReadedDbTbProducts = True
				EndIf

				GUICtrlSetState($lvInTabProducts, $GUI_SHOW)
				GUICtrlSetData($menuItemProducts, "�� ��˾��Ʒ")
			Case "ѧУ��Ϣ"
				If $hasReadedDbTbSchools = False Then
					FuncReadDb($tb_schools, $lvInTabSchools)
					$hasReadedDbTbSchools = True
				EndIf

				GUICtrlSetState($lvInTabSchools, $GUI_SHOW)
				GUICtrlSetData($menuItemSchools, "�� ѧУ��Ϣ")
			Case "���̹���"
				If $hasReadedDbTbProjects = False Then
					FuncReadDb($tb_projects, $lvInTabProjects)
					$hasReadedDbTbProjects = True
				EndIf

				GUICtrlSetState($lvInTabProjects, $GUI_SHOW)
				GUICtrlSetData($menuItemProjects, "�� ���̹���")
			Case "�������"
				If $hasReadedDbTbPartners = False Then
					FuncReadDb($tb_partners, $lvInTabPartners)
					$hasReadedDbTbPartners = True
				EndIf

				GUICtrlSetState($lvInTabPartners, $GUI_SHOW)
				GUICtrlSetData($menuItemPartners, "�� �������")

			Case "�û�����"
				If $hasReadedDbTbUsers= False Then
					FuncReadDb($tb_users, $lvInTabUsers)
					$hasReadedDbTbUsers = True
				EndIf

				GUICtrlSetState($lvInTabUsers, $GUI_SHOW)
				GUICtrlSetData($menuItemUsers, "�� �û�����")
			Case "�ʲ�����"
				If $hasReadedDbTbAccets = False Then
					FuncReadDb($tb_accets, $lvInTabAccets)
					$hasReadedDbTbAccets = True
				EndIf

				GUICtrlSetState($lvInTabAccets, $GUI_SHOW)
				GUICtrlSetData($menuItemAccets, "�� �ʲ�����")
			Case "��Ա����"
				If $hasReadedDbTbWorkers = False Then
					FuncReadDb($tb_workers, $lvInTabWorkers)
					$hasReadedDbTbWorkers = True
				EndIf

				GUICtrlSetState($lvInTabWorkers, $GUI_SHOW)
				GUICtrlSetData($menuItemWorkers, "�� ��Ա����")
			Case "������־"
				If $hasReadedDbTbLog = False Then
					FuncReadDb($tb_log, $lvInTabLog)
					$hasReadedDbTbLog = True
				EndIf

				GUICtrlSetState($lvInTabLog, $GUI_SHOW)
				GUICtrlSetData($menuItemLog, "�� ������־")
			Case "Ԫ �� ��"
				If $hasReadedDbTbSource = False Then
					FuncReadDb($tb_source, $lvInTabSources)
					$hasReadedDbTbSource = True
				EndIf

				GUICtrlSetState($lvInTabSources, $GUI_SHOW)
				GUICtrlSetData($menuItemSources, "�� Ԫ �� ��")
		EndSwitch

	Else
		Switch $strItemName
			Case "������־"
				$id_item = 0
			Case "��˾��Ʒ"
				$id_item = 1
			Case "ѧУ��Ϣ"
				$id_item = 2
			Case "���̹���"
				$id_item = 3
			Case "�������"
				$id_item = 4

			Case "�û�����"
				$id_item = 5
			Case "�ʲ�����"
				$id_item = 6
			Case "��Ա����"
				$id_item = 7
			Case "������־"
				$id_item = 8
			Case "Ԫ �� ��"
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
					Case "������־"
						If $hasReadedDbTbJournal = False Then
							FuncReadDb($tb_journal, $lvInTabJournal)
							$hasReadedDbTbJournal = True
						EndIf

						GUICtrlSetState($lvInTabJournal, $GUI_SHOW)
						GUICtrlSetData($menuItemJournal, "�� ������־")
					Case "��˾��Ʒ"
						If $hasReadedDbTbProducts = False Then
							FuncReadDb($tb_products, $lvInTabProducts)
							$hasReadedDbTbProducts = True
						EndIf

						GUICtrlSetState($lvInTabProducts, $GUI_SHOW)
						GUICtrlSetData($menuItemProducts, "�� ��˾��Ʒ")
					Case "ѧУ��Ϣ"
						If $hasReadedDbTbSchools = False Then
							FuncReadDb($tb_schools, $lvInTabSchools)
							$hasReadedDbTbSchools = True
						EndIf

						GUICtrlSetState($lvInTabSchools, $GUI_SHOW)
						GUICtrlSetData($menuItemSchools, "�� ѧУ��Ϣ")
					Case "���̹���"
						If $hasReadedDbTbProjects = False Then
							FuncReadDb($tb_projects, $lvInTabProjects)
							$hasReadedDbTbProjects = True
						EndIf

						GUICtrlSetState($lvInTabProjects, $GUI_SHOW)
						GUICtrlSetData($menuItemProjects, "�� ���̹���")
					Case "�������"
						If $hasReadedDbTbPartners = False Then
							FuncReadDb($tb_partners, $lvInTabPartners)
							$hasReadedDbTbPartners = True
						EndIf

						GUICtrlSetState($lvInTabPartners, $GUI_SHOW)
						GUICtrlSetData($menuItemPartners, "�� �������")

					Case "�û�����"
						If $hasReadedDbTbUsers= False Then
							FuncReadDb($tb_users, $lvInTabUsers)
							$hasReadedDbTbUsers = True
						EndIf

						GUICtrlSetState($lvInTabUsers, $GUI_SHOW)
						GUICtrlSetData($menuItemUsers, "�� �û�����")
					Case "�ʲ�����"
						If $hasReadedDbTbAccets = False Then
							FuncReadDb($tb_accets, $lvInTabAccets)
							$hasReadedDbTbAccets = True
						EndIf

						GUICtrlSetState($lvInTabAccets, $GUI_SHOW)
						GUICtrlSetData($menuItemAccets, "�� �ʲ�����")
					Case "��Ա����"
						If $hasReadedDbTbWorkers = False Then
							FuncReadDb($tb_workers, $lvInTabWorkers)
							$hasReadedDbTbWorkers = True
						EndIf

						GUICtrlSetState($lvInTabWorkers, $GUI_SHOW)
						GUICtrlSetData($menuItemWorkers, "�� ��Ա����")
					Case "������־"
						If $hasReadedDbTbLog = False Then
							FuncReadDb($tb_log, $lvInTabLog)
							$hasReadedDbTbLog = True
						EndIf

						GUICtrlSetState($lvInTabLog, $GUI_SHOW)
						GUICtrlSetData($menuItemLog, "�� ������־")
					Case "Ԫ �� ��"
						If $hasReadedDbTbSource = False Then
							FuncReadDb($tb_source, $lvInTabSources)
							$hasReadedDbTbSource = True
						EndIf

						GUICtrlSetState($lvInTabSources, $GUI_SHOW)
						GUICtrlSetData($menuItemSources, "�� Ԫ �� ��")
				EndSwitch
			EndIf
		Next
	EndIf
EndFunc

Func Func_MenuHelp ()
EndFunc

;#cs ��ȡ������ݣ���ʾ��ListView��
;	$strTbName��Ҫ��ѯ�ı�����
;	$strListView��Ҫ��ʾ���ݵ��б�
;#ce
Func FuncReadDb( $strTbName, $strListView )

	Local $obj_adodb_connection = ObjCreate("ADODB.Connection")
	$obj_adodb_connection.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & $db_path & "; Jet OLEDB:Database Password='" & $db_pswd & "'")
	Local $obj_adodb_recordset = ObjCreate("ADODB.Recordset")
	$obj_adodb_recordset.ActiveConnection = $obj_adodb_connection

	Local $oRec = $obj_adodb_recordset
	$oRec.Open($strTbName, $obj_adodb_connection, 3, 3)
	Local $fieldsCount = $oRec.fields.count	; ��ȡ�����������ܹ��ж����ֶ�

	$obj_adodb_recordset.Open("SELECT * FROM " & $strTbName)
	;
	; ������־�����⣺��������ֶ�����
	;
	If $strTbName = $tb_journal Then
		$fieldsCount = $fieldsCount - 2
	EndIf

	_GUICtrlListView_BeginUpdate($strListView)
	; ���� ��
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
		$strFields = StringTrimRight ( $strFields, 1 )	; ɾ��ĩβ�� |

		; ���� ��
		GUICtrlCreateListViewItem( $strFields, $strListView)
			GUICtrlSetBkColor (-1, 0xffa500 );����listviewitem�ı���ɫ
		$obj_adodb_recordset.movenext
	WEnd
	_GUICtrlListView_EndUpdate($strListView)

	GUICtrlSetData($strInfoLbl, "��" & $strTbName & "���м�¼������" & $obj_adodb_recordset.recordcount)

	$obj_adodb_recordset.close
	$obj_adodb_connection.Close
	$oRec.close

EndFunc

;#cs �������ݿ�
;	$strDbPath���ļ�·��
;	$strDbPswd����������
;#ce
Func FuncCreateDb()
	; �����к�ListView�ж�Ӧ

	;=== ��˾��Ա�� tb_workers ��14
	;                              ��š������Զ�����|                   �������ı�|  ���֤��|               ����|              ְ��|            ��ְ����|             ת������|            ��н��|            �Ա�|       ����|            �绰|         ����|         סַ��Ĭ��50�ַ�|    ��ͥ��Ϣ|          ��ע
	Local $strColsForTbWorders = "id integer identity(1,1) primary key, w_name text, w_identity_number text, w_deportment text, w_position text, w_date_of_entry text, w_date_regular text, w_month_wage text, w_sex text, w_birthday text, w_phone text, w_email text, w_address text(250), w_family text(254), w_note memo"

	;=== ѧУ��Ϣ�� tb_schools ��6
	;                              ���|                                 ѧУ|        ��ϵ��|         ��ַ|                �绰|         ����|        ��ע
	Local $strColsForTbSchools = "id integer identity(1,1) primary key, s_name text, s_contact text, s_address text(250), s_phone text, s_email text, s_note memo"

	;=== ��˾��Ʒ�� tb_products ��6
	;                               ���|                                 ��Ʒ|         ����|        ���|             ����|                       ���|        ��ע
	Local $strColsForTbProducts = "id integer identity(1,1) primary key, pd_name text, pd_type text, pd_desiger text, pd_configuration text(250), pd_cost text, pd_note memo"

	;=== �ʲ������ tb_accets ��10
	;                             ���|                                 ����|        ����|                 ��λ|        ����|        ��������|           ����|         ��������|          ������|        �Ƿ񱨷�|     ��ע
	Local $strColsForTbAccets = "id integer identity(1,1) primary key, a_name text, a_serial_number text, a_unit text, a_type text, a_date_bought text, a_price text, a_deportment text, a_dealer text, a_scrap text, a_note text(250)"

	;=== �������� tb_partners ��7
	;                               ���|                                 ����|         ����|         ��ַ|                 �绰|          ����|          ҵ��|             ��ע
	Local $strColsForTbPartners = "id integer identity(1,1) primary key, pt_name text, pt_type text, pt_address text(250), pt_phone text, pt_email text, pt_business text, pt_note memo"

	;=== ���̹���� tb_projects ��10
	;                               ���|                                 ����|         ѧУ|           ��Ʒ|            ������|          ��˾������|     ϸ����Ƭ��|         ��ʼ����|           ״̬|          ��������|            �����¼|             ��ע
	Local $strColsForTbProjects = "id integer identity(1,1) primary key, pj_name text, pj_s_name text, pj_pd_name text, pj_pt_name text, pj_w_name text, pj_content text(250), pj_date_start text, pj_state text, pj_date_finish text, pj_account text(250), pj_note memo"

	;=== ������־�� tb_journal ��7
	;                              ���|                                 ��Ա|        ����|        �ص�|                ��ͨ|                ʳ��|              ��������|            ��ע|        ��¼��(����)|        ��¼����(����)
	Local $strColsForTbJournal = "id integer identity(1,1) primary key, j_name text, j_date text, j_address text(250), j_traffic text(250), j_board text(250), j_content text(250), j_note memo, j_record text, j_date_record text"

	;=== �û��� tb_users ��4
	;                            ���|                                 �û���|      ����|        Ȩ��|             ��ע
	Local $strColsForTbUsers = "id integer identity(1,1) primary key, u_name text, u_pswd text, u_authority text, u_note text(250)"

	;=== Դ���ݱ� tb_source ��4
	;                             ���|                                 ��|            ��|             ֵ|            ����
	Local $strColsForTbSource = "id integer identity(1,1) primary key, sr_tb_name text, sr_column text, sr_value text, sr_note text(250)"

	;=== ������־�� tb_log ��
	;                          ���|                                 �û�|        ������ɾ��|     ��|            ��|            ������|          ������|          ����|        ��ע
	Local $strColsForTbLog = "id integer identity(1,1) primary key, l_name text, l_operate text, l_table text, l_column text,  l_old_data memo, l_new_data memo, l_date text, l_note text(250)"

	; �������ݿ��ļ�
	Local $oADO = ObjCreate("ADOX.Catalog")
	$oADO.Create("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & $db_path & "; Jet OLEDB:Database Password='" & $db_pswd & "'")
	$oADO.ActiveConnection.Close

	; ������ $tb_workers, $tb_schools, $tb_products, $tb_accets, $tb_partners, $tb_projects, $tb_journal, $tb_users, $tb_source, $tb_log
	$oADO = ObjCreate("ADODB.Connection")
	$oADO.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & $db_path & "; Jet OLEDB:Database Password='" & $db_pswd & "'")
	$oADO.Execute("CREATE TABLE " & $tb_workers)	; ��Ա����'
	$oADO.Execute("CREATE TABLE " & $tb_schools)	; ѧУ��Ϣ
	$oADO.Execute("CREATE TABLE " & $tb_products)	; ��˾��Ʒ
	$oADO.Execute("CREATE TABLE " & $tb_accets)		; �ʲ�����
	$oADO.Execute("CREATE TABLE " & $tb_partners)	; �������
	$oADO.Execute("CREATE TABLE " & $tb_projects)	; ���̹���
	$oADO.Execute("CREATE TABLE " & $tb_journal)	; �����棬������־
	$oADO.Execute("CREATE TABLE " & $tb_users)		; �û�����
	$oADO.Execute("CREATE TABLE " & $tb_source)		; Ԫ����
	$oADO.Execute("CREATE TABLE " & $tb_log)		; ������־

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
