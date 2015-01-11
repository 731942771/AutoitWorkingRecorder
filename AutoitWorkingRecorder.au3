; www.cuiweiyou.com
; ������|vigiles
; http://www.cuiweiyou.com/tag/autoit%E4%B8%8A%E8%B7%AF

#AutoIt3Wrapper_UseX64=n		; 64λwin7

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
; ȫ�ֵ����ݿ�·��
Global $db_path = @ScriptDir & "\mdbZhongHui.mdb"
Global $db_pswd = ""	; ��������mdb�ļ�������

; �õ��ı���          ��˾��Ա                   ѧУ��Ϣ                      ��˾��Ʒ                  �ʲ�����                     �������                       ���̹���                   ������־                   �û�                     Ԫ����              ������־
Global $tb_workers = "tb_workers", $tb_schools = "tb_schools", $tb_products = "tb_products", $tb_accets = "tb_accets", $tb_partners = "tb_partners", $tb_projects = "tb_projects", $tb_journal = "tb_journal", $tb_users = "tb_users", $tb_source = "tb_source", $tb_log = "tb_log"

;-------------------- Դ���ݲ��� -------------------------------

If FileExists ( $db_path ) = 0 Then
	If MsgBox(52, "����", "δ��⵽���ݿ��ļ����Ƿ����´�����") = 6 Then
		FuncCreateDb ( )
	EndIf
EndIf

;======================================================================
;-------------------- �û���¼ --------------------------------


;======================================================================
; ������ǩ�л�LVʱ�õ���
Global $itemInToolbar, $idOfTabItem	; ���һ��ȫ�ֱ���$idOfTabItem��¼���һ��ı�ǩ����
; ��������
Global $WidthOfWindow = 800, $HeightOfWindow = 600
; �����壬����������������ʾ��
Global $guiMainWindow, $toolbarInMainWindow, $hToolTip 
; �������ϵİ�ť�ȶ�ID
Global Enum $id_Toolbar_New = 1000, $id_Toolbar_Save, $id_Toolbar_Delete, $id_Toolbar_Find, $id_Toolbar_Help
; ���ݿ��ȡ��ʶ��������True
Global $hasReadedDbTbWorkers = False, $hasReadedDbTbSchools = False, $hasReadedDbTbProducts = False, $hasReadedDbTbAccets = False, $hasReadedDbTbPartners = False, $hasReadedDbTbProjects = False, $hasReadedDbTbUsers = False, $hasReadedDbTbSource = False, $hasReadedDbTbLog = False, $hasReadedDbTbJournal = False
; ˫����Ԫ��ʱ�����༭���õ���
Global $hEdit, $hBrush, $hDC, $hItemRow, $hItemColumn	; ˫��LV�����޸��õ���
; ˫����Ԫ���޸����ݣ�������־��Log��ʱ�õ��ġ��޸ĵ����������ĸ��������ĸ�LV����Ӧ���������޸�ǰ����
Global $columnOfTable, $columnOfLV, $columnName, $columnOldData
; �Ӵ��壬�Ӵ����ȷ����ť��ȡ����ť��
Global $popupWindow, $btnOKInPopWin, $btnNOInPopWin
; ��ѯʱ�����ѯ����
Global $strArgsOfQuery, $strArgsOfLVArr[7], $strArgsOfTableArr[7]	; ���ݱ�����Ҫ��������Ϊ7

;======================================================================
;-------------------- GUI ---------------------------------
$guiMainWindow = GUICreate("�лս���|������ www.cuiweiyou.com", $WidthOfWindow, $HeightOfWindow) 
	GUISetOnEvent($GUI_EVENT_CLOSE, "Func_GUI_EVENT_CLOSE")
	GUISetIcon( @ScriptDir & "\Logo.ico")	; ���ó���ͼ��Ϊ�ű��ļ�ͬĿ¼�е�Logo.ico
	
	$menuFile = GUICtrlCreateMenu ( "�ļ� &F")
		$itemOpenInMenuFile = GUICtrlCreateMenuItem("��", $menuFile)
		$itemSaveInMenuFile = GUICtrlCreateMenuItem("����", $menuFile)
		GUICtrlCreateMenuItem("", $menuFile) ; �ָ���
		$itemRecentfilesInMenuFile = GUICtrlCreateMenu("������ļ�", $menuFile)
		GUICtrlCreateMenuItem("", $menuFile) 
		$itemExitInMenuFile = GUICtrlCreateMenuItem("�˳�", $menuFile)
			GUICtrlSetOnEvent($itemExitInMenuFile, "Func_GUI_EVENT_CLOSE")
	
	$menuTab = GUICtrlCreateMenu ( "���� &W")
		$itemJournalOfTabInMenu = GUICtrlCreateMenuItem("�� ������־", $menuTab)
			GUICtrlSetOnEvent($itemJournalOfTabInMenu, "Func_ShowTab_ByMenu")
		$itemProductsOfTabInMenu = GUICtrlCreateMenuItem("  ��˾��Ʒ", $menuTab)
			GUICtrlSetOnEvent($itemProductsOfTabInMenu, "Func_ShowTab_ByMenu")
		$itemSchoolsOfTabInMenu = GUICtrlCreateMenuItem("  ѧУ��Ϣ", $menuTab)
			GUICtrlSetOnEvent($itemSchoolsOfTabInMenu, "Func_ShowTab_ByMenu")
		$itemPartnersOfTabInMenu = GUICtrlCreateMenuItem("  �������", $menuTab)
			GUICtrlSetOnEvent($itemPartnersOfTabInMenu, "Func_ShowTab_ByMenu")
		$itemProjectsOfTabInMenu = GUICtrlCreateMenuItem("  ���̹���", $menuTab)
			GUICtrlSetOnEvent($itemProjectsOfTabInMenu, "Func_ShowTab_ByMenu")
	
	$menuHelp = GUICtrlCreateMenu ( "���� &H")
		$itemGuideInMenuHelp = GUICtrlCreateMenuItem("ʹ��ָ��", $menuHelp)
			GUICtrlSetOnEvent($itemGuideInMenuHelp, "Func_MenuItemGuide_In_MenuHelp")	; �򿪡���˾��Ʒ����ǩ��
		$itemAboutInMenuHelp = GUICtrlCreateMenuItem("����", $menuHelp)
			GUICtrlSetOnEvent($itemAboutInMenuHelp, "Func_MenuItemAbout_In_MenuHelp")	; �򿪡�ѧУ��Ϣ����ǩ��
	
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
			_GUIImageList_Add($imgList, _WinAPI_CreateSolidBitmap($guiMainWindow, 0x990033, 16, 16))
			_GUIImageList_Add($imgList, _WinAPI_CreateSolidBitmap($guiMainWindow, 0x00FF00, 16, 16))
			_GUIImageList_Add($imgList, _WinAPI_CreateSolidBitmap($guiMainWindow, 0x0000FF, 16, 16))
			_GUIImageList_Add($imgList, _WinAPI_CreateSolidBitmap($guiMainWindow, 0xFF3399, 16, 16))
			_GUIImageList_Add($imgList, _WinAPI_CreateSolidBitmap($guiMainWindow, 0x007700, 16, 16))
			_GUIImageList_Add($imgList, _WinAPI_CreateSolidBitmap($guiMainWindow, 0xFF6600, 16, 16))	; ��ѯ���
			_GUICtrlTab_SetImageList($tabInMainWindow, $imgList)
	
		_GUICtrlTab_InsertItem ( $tabInMainWindow, 0, "������־", 0)
	
	$mouseMenuTab = GUICtrlCreateContextMenu($tabInMainWindow)
		$mouseMenuItemClose = GUICtrlCreateMenuItem("�ر�", $mouseMenuTab)
			GUICtrlSetOnEvent($mouseMenuItemClose, "Func_MouseMenuItem")
		GUICtrlCreateMenuItem("", $mouseMenuTab)
		$mouseMenuItemSaveAs = GUICtrlCreateMenuItem("���Ϊ", $mouseMenuTab)
			GUICtrlSetOnEvent($mouseMenuItemSaveAs, "Func_MouseMenuItem")
		GUICtrlCreateMenuItem("", $mouseMenuTab)
		$mouseMenuItemPrint = GUICtrlCreateMenuItem("��ӡ", $mouseMenuTab)
			GUICtrlSetOnEvent($mouseMenuItemPrint, "Func_MouseMenuItem")

	GUICtrlCreateLabel("�ı� 3", 2, $HeightOfWindow - 38, 50, 20)
	
	$lvInTabJournal = GUICtrlCreateListView ( "���|��Ա|����|�ص�|��ͨ|ʳ��|��������|��ע", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)	; ������־��
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabJournal, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))	; ͳ�����ö�����չ��ʽ
		GUICtrlSetBkColor($lvInTabJournal, 0xffffff)					;����listview�ı���ɫ
		GUICtrlSetBkColor($lvInTabJournal, $GUI_BKCOLOR_LV_ALTERNATE)	;������Ϊlistview�ı���ɫ��ż����Ϊlistviewitem�ı���ɫ
		GUICtrlSetState($lvInTabJournal, $GUI_SHOW)
	$lvInTabProducts = GUICtrlCreateListView ( "���|��Ʒ|����|���|����|���|��ע", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)	; ��˾��Ʒ��
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabProducts, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabProducts, 0xffffff)
		GUICtrlSetBkColor($lvInTabProducts, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabProducts, $GUI_HIDE)
	$lvInTabSchools  = GUICtrlCreateListView ( "���|ѧУ|��ϵ��|��ַ|�绰|����|��ע", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)	; ѧУ��Ϣ��
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabSchools, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabSchools, 0xffffff)
		GUICtrlSetBkColor($lvInTabSchools, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabSchools, $GUI_HIDE)
	$lvInTabProjects = GUICtrlCreateListView ( "���|����|ѧУ|��Ʒ|������|��˾������|ϸ��|��ʼ����|״̬|��������|�����¼|��ע", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)	; ���̹����
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabProjects, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabProjects, 0xffffff)
		GUICtrlSetBkColor($lvInTabProjects, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabProjects, $GUI_HIDE)
	$lvInTabPartners = GUICtrlCreateListView ( "���|����|����|��ַ|�绰|����|ҵ��|��ע", 3, 51, $WidthOfWindow - 5, $HeightOfWindow - 95)	; ��������
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabPartners, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabPartners, 0xffffff)
		GUICtrlSetBkColor($lvInTabPartners, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabPartners, $GUI_HIDE)
#cs
	$lvInTabAssets   = GUICtrlCreateListView ( "���|����|����|��λ|����|��������|����|��������|������|�Ƿ񱨷�|��ע", 3, 51, $WidthOfWindow - 6, $HeightOfWindow - 97)	; �ʲ������
		_GUICtrlListView_SetExtendedListViewStyle($lvInTabAssets, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
		GUICtrlSetBkColor($lvInTabAssets, 0xffffff)
		GUICtrlSetBkColor($lvInTabAssets, $GUI_BKCOLOR_LV_ALTERNATE)
		GUICtrlSetState($lvInTabAssets, $GUI_HIDE)
#ce

GUISetState(@SW_SHOW, $guiMainWindow)

; #cs
;	��ʼ�ڡ���־�����ݿ��ѯ���ڡ���־���б��������
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
	
	Switch $itemText
		Case "������־"
			GUICtrlSetState($lvInTabJournal, $GUI_SHOW)
		Case "��˾��Ʒ"
			GUICtrlSetState($lvInTabProducts, $GUI_SHOW)
		Case "ѧУ��Ϣ"
			GUICtrlSetState($lvInTabSchools, $GUI_SHOW)
		Case "���̹���"
			GUICtrlSetState($lvInTabProjects, $GUI_SHOW)
		Case "�������"
			GUICtrlSetState($lvInTabPartners, $GUI_SHOW)
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
        Case $toolbarInMainWindow
            Switch $code	; �¼�
				Case $TBN_HOTITEMCHANGE
					$tNMTBHOTITEM = DllStructCreate($tagNMTBHOTITEM, $lParam)
					$i_idOld = DllStructGetData($tNMTBHOTITEM, "idOld")
					$i_idNew = DllStructGetData($tNMTBHOTITEM, "idNew")
					$itemInToolbar = $i_idNew
 
				Case $NM_CLICK
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
							Case "������־"
								GUICtrlSetState($lvInTabJournal, $GUI_SHOW)
							Case "��˾��Ʒ"
								GUICtrlSetState($lvInTabProducts, $GUI_SHOW)
							Case "ѧУ��Ϣ"
								GUICtrlSetState($lvInTabSchools, $GUI_SHOW)
							Case "���̹���"
								GUICtrlSetState($lvInTabProjects, $GUI_SHOW)
							Case "�������"
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
; �ڵ�ǰ�б��в�ѯ����
; �ۼӲ�ͬ���ֶμ���Ϊ��ѯ����
; ��$strArgsOfLVArr�б���ؼ��ֿؼ���ID
; ��$strArgsOfTableArr�б������
;#ce
Func FuncFindData()
	Local $strLvName
	$strLvName = _GUICtrlTab_GetItemText ( $tabInMainWindow, _GUICtrlTab_GetCurSel ( $tabInMainWindow ))
	
	$popupWindow = GUICreate("" , 500, 175, Default, Default, $WS_CAPTION, BitOR($WS_EX_TOPMOST, $GUI_WS_EX_PARENTDRAG))
		WinSetTitle($popupWindow, "", "�ڡ�" & $strLvName & "���в�ѯ����" )
		
		GUICtrlCreateGroup("����Ҫ������������ؼ��֡�������ÿո�ֿ�", 10, 10, 480, 125)

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
		EndSwitch
		
		GUICtrlCreateGroup("", -99, -99, 1, 1) 
		
		$btnOKInPopWin = GUICtrlCreateButton("ȷ��", 10, 145, 60, 22)
			GUICtrlSetOnEvent($btnOKInPopWin, "FuncFindDataByArgs")
		$btnNOInPopWin = GUICtrlCreateButton("ȡ��", 428, 145, 60, 22)
			GUICtrlSetOnEvent($btnNOInPopWin, "FuncFindDataByArgs")
		
	GUISetState(@SW_SHOW, $popupWindow)
	GUISetState(@SW_DISABLE, $guiMainWindow) 
EndFunc

Func FuncFindDataByArgs ()
	If @GUI_CtrlId = $btnOKInPopWin Then
		Local $isSqlStrEnable = False	; ��ѯ����Ƿ���á����ȫ���������ǿյģ�������Ϊtrue
		
		$strArgsOfQuery = "SELECT * FROM " & $strArgsOfTableArr[0] & " WHERE "
		
		; �����ؼ����飬ƴ�����
		For $iiii = 1 To 6 Step 1	; 0�����������lV��Access���ơ�ʣ��1-6����ؼ�
			
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
						$strArgsOfQuery &= $strArgsOfTableArr[$iiii] & " LIKE '%" & $keywordArr[$j] & "%' AND "
					Next
				EndIf
				
				$isSqlStrEnable = True	; ������һ����Ч�������ſ���
			EndIf
		Next
		
		$strArgsOfQuery &= "1 = 1"	; ƴ��sql��䳣�ý�β������ɾ�����һ�� AND�Ƚ��鷳
		
		If $isSqlStrEnable Then
			; ִ�����ݿ��ѯ
			Local $adoCon = ObjCreate("ADODB.Connection")
			$adoCon.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & $db_path & "; Jet OLEDB:Database Password='" & $db_pswd & "'")
			Local $adoRec = ObjCreate("ADODB.Recordset")
			$adoRec.ActiveConnection = $adoCon
			
			$adoRec.Open($strArgsOfQuery)	; �ӱ��в鵽�Ľ����
			
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
						$strResultDataForLv &= $adoRec.fields($i).value & "|" 
					Next
					
					GUICtrlCreateListViewItem($strResultDataForLv, $strArgsOfLVArr[0])
						GUICtrlSetBkColor (-1, 0xffa500 );����listviewitem�ı���ɫ
					
					$adoRec.Movenext	; ���������һ��
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
	
	; ��λ���������
	$strArgsOfQuery = ""
	For $i = 0 To 6
		$strArgsOfLVArr[$i] = ""
		$strArgsOfTableArr[$i] = ""
	Next
EndFunc

; #cs
; ˫��LV�ĵ�Ԫ��󣬴���һ�������
; ��ԭ���ݽ����޸�
; ���� $strLvInTab:Ҫ�������б�
; ���� $lParam:ϵͳ��Ϣ
; #ce
Func FuncCreateEditRecForColumn( $strLvInTab, $lParam )
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
EndFunc

; �ؼ�ʧȥ����
; ��Ԫ�������ݱ���
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
						$strLvItem = "'Eminem', 'update', '" & $columnOfTable & "', '" & $columnName & "', '" & $columnOldData & "', '" & $iText & "', '" & @YEAR & "-" & @MON & "-" & @MDAY & "(" & @WDAY - 1 & ")" & @HOUR & ":" & @MIN & ":" & @SEC & "`" & @MSEC & "', '_none'"	; ����
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
	EndSwitch
	
	FuncDeleteItemFromAccessAndListView ( $strTbName, $strLvInTab, $strLvItem )
EndFunc

; ɾ��ѡ�е���
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
		
		; ��־��
		$strLvItem = "'Eminem', 'del', '" & $strTbName & "', '_item', '" & StringTrimLeft ($strLvItem, StringInStr( $strLvItem, "|")) & "', '_none', '" & @YEAR & "-" & @MON & "-" & @MDAY & "(" & @WDAY - 1 & ")" & @HOUR & ":" & @MIN & ":" & @SEC & "`" & @MSEC & "', '_none'"	; ����
		$strLvItem = "insert into " & $tb_log & " (l_name, l_operate, l_table, l_column, l_old_data, l_new_data, l_date, l_note) values ( " & $strLvItem & " )"				; �������
		$dataAdodbConnectionDel.Execute($strLvItem)
		
		$dataAdodbConnectionDel.Close
	EndIf
EndFunc

; #cs
; ��Ӧ���������½�����ť��������Ŀ��ListView
; �����жϵ�ǰ����ı�ǩ-��һ��ListView������ʾ״̬
; Ȼ�������Ŀ��������Ŀλ��
; #ce
Func FuncInsertItemToListView ()
	Local $tmpLvItem, $strLvItem, $strLvInTab, $strTbName
	Switch _GUICtrlTab_GetItemText ( $tabInMainWindow, _GUICtrlTab_GetCurSel ( $tabInMainWindow ))
		Case "������־"
			$strLvItem = "���|GGGG��Ա|����|�ص�|��ͨ|ʳ��|��������|��ע"	; ���뵽LV�е���Ŀ
			$strLvInTab = $lvInTabJournal	; LV
			$strTbName = $tb_journal		; Table
		Case "��˾��Ʒ"
			$strLvItem = "���|GSGS��Ʒ|����|���|����|���|��ע"
			
			$strLvInTab = $lvInTabProducts
			$strTbName = $tb_products
		Case "ѧУ��Ϣ"
			$strLvItem = "���|XXXXѧУ|��ϵ��|��ַ|�绰|����|��ע"
			
			$strLvInTab = $lvInTabSchools
			$strTbName = $tb_schools
		Case "���̹���"
			$strLvItem = "���|GCGC����|ѧУ|��Ʒ|������|��˾������|ϸ����Ƭ��|��ʼ����| ״̬|��������|�����¼|��ע"
			
			$strLvInTab = $lvInTabProjects
			$strTbName = $tb_projects
		Case "�������"
			$strLvItem = "���|HZHZ����|����|��ַ|�绰|����|ҵ��|��ע"
			
			$strLvInTab = $lvInTabPartners
			$strTbName = $tb_partners
	EndSwitch
	
	$tmpLvItem = GUICtrlCreateListViewItem( $strLvItem, $strLvInTab)
	GUICtrlSetBkColor ($tmpLvItem, 0xff9900 )
	_GUICtrlListView_Scroll($strLvInTab, 0, _GUICtrlListView_GetItemCount($strLvInTab)*15)
	
	; ��һ���ǡ���š���ȥ��
	$strLvItem = "'" & StringReplace ( StringTrimLeft ($strLvItem, StringInStr( $strLvItem, "|")), "|", "', '" ) & "'"	; תΪ�ȶ���ʽ�����ݿ�����ִ�
	
	FuncInsertItemToAccessAndListView ( $strTbName, $strLvItem )
EndFunc

; #cs
; �����ݿ���д������Ŀ
; $strTbName��������
; $strLvItem������
; #ce
Func FuncInsertItemToAccessAndListView ( $strTbName, $strLvItem )
	Local $strColumns
	Switch $strTbName
		Case $tb_products	; ��Ʒ����
			$strColumns = " (pd_name, pd_type, pd_desiger, pd_configuration, pd_cost, pd_note) values ("
		Case $tb_partners	; ���
			$strColumns = " (pt_name, pt_type, pt_address, pt_phone, pt_email, pt_business, pt_note) values ("
		Case $tb_projects	; ����
			$strColumns = " (pj_name, pj_s_name, pj_pd_name, pj_pt_name, pj_w_name, pj_content, pj_date_start, pj_state, pj_date_finish, pj_account, pj_note) values ("
		Case $tb_journal	; ��־
			$strColumns = " (j_name, j_date, j_address, j_traffic, j_board, j_content, j_note) values ("	; , j_record, j_date_record
		Case $tb_schools	; ѧУ"
			$strColumns = " (s_name, s_contact, s_address, s_phone, s_email, s_note) values ("
	EndSwitch
	Local $dataAdodbConnectionIn = ObjCreate("ADODB.Connection")
	$dataAdodbConnectionIn.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & $db_path & ";Jet Oledb:Database Password='" & $db_pswd & "'")
	$dataAdodbConnectionIn.Execute("insert into " & $strTbName & $strColumns & $strLvItem & ")")
	; ��־��
	$strLvItem = "'Eminem', 'add', '" & $strTbName & "', '_item', '_none', '" & StringReplace($strLvItem, "'", "") & "', '" & @YEAR & "-" & @MON & "-" & @MDAY & "(" & @WDAY - 1 & ")" & @HOUR & ":" & @MIN & ":" & @SEC & "`" & @MSEC & "', '_none'"	; ����
	$strLvItem = "insert into " & $tb_log & " (l_name, l_operate, l_table, l_column, l_old_data, l_new_data, l_date, l_note) values ( " & $strLvItem & " )"				; �������
	$dataAdodbConnectionIn.Execute($strLvItem)
	$dataAdodbConnectionIn.Close

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
			
			If $idOfTabItem = 0 Then
				_GUICtrlTab_SetCurSel($tabInMainWindow, 1)
			Else 
				_GUICtrlTab_SetCurSel($tabInMainWindow, $idOfTabItem - 1)
			EndIf
			
			Switch _GUICtrlTab_GetItemText($tabInMainWindow, $idOfTabItem)
				Case "������־"
					GUICtrlSetData($itemJournalOfTabInMenu, "��������־")
				Case "��˾��Ʒ"
					GUICtrlSetData($itemProductsOfTabInMenu, "����˾��Ʒ")
				Case "ѧУ��Ϣ"
					GUICtrlSetData($itemSchoolsOfTabInMenu, "��ѧУ��Ϣ")
				Case "���̹���"
					GUICtrlSetData($itemProjectsOfTabInMenu, "�����̹���")
				Case "�������"
					GUICtrlSetData($itemPartnersOfTabInMenu, "���������")
			EndSwitch
			
			_GUICtrlTab_DeleteItem ( $tabInMainWindow, $idOfTabItem )
			
			$strTabItemText = _GUICtrlTab_GetItemText ( $tabInMainWindow, _GUICtrlTab_GetCurSel ( $tabInMainWindow ) )
			
			Switch $strTabItemText
				Case "������־"
					GUICtrlSetState($lvInTabJournal, $GUI_SHOW)
				Case "��˾��Ʒ"
					GUICtrlSetState($lvInTabProducts, $GUI_SHOW)
				Case "ѧУ��Ϣ"
					GUICtrlSetState($lvInTabSchools, $GUI_SHOW)
				Case "���̹���"
					GUICtrlSetState($lvInTabProjects, $GUI_SHOW)
				Case "�������"
					GUICtrlSetState($lvInTabPartners, $GUI_SHOW)
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
		Case $itemJournalOfTabInMenu
			Func_ShowTab_ByText ("������־")
		Case $itemProductsOfTabInMenu
			Func_ShowTab_ByText ("��˾��Ʒ")
		Case $itemSchoolsOfTabInMenu
			Func_ShowTab_ByText ("ѧУ��Ϣ")
		Case $itemProjectsOfTabInMenu
			Func_ShowTab_ByText ("���̹���")
		Case $itemPartnersOfTabInMenu
			Func_ShowTab_ByText("�������")
	EndSwitch
EndFunc

#cs ͨ��������ַ������ͱ�ǩͷ����ƥ�䡣������ʾ������
#ce
Func Func_ShowTab_ByText ( $strItemName )
	; $itemJournalOfTabInMenu $itemProductsOfTabInMenu $itemSchoolsOfTabInMenu $itemProjectsOfTabInMenu $itemPartnersOfTabInMenu "����˾��Ʒ"
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
			Case "������־"
				GUICtrlSetData($itemJournalOfTabInMenu, "  ������־")
			Case "��˾��Ʒ"
				GUICtrlSetData($itemProductsOfTabInMenu, "  ��˾��Ʒ")
			Case "ѧУ��Ϣ"
				GUICtrlSetData($itemSchoolsOfTabInMenu, "  ѧУ��Ϣ")
			Case "���̹���"
				GUICtrlSetData($itemProjectsOfTabInMenu, "  ���̹���")
			Case "�������"
				GUICtrlSetData($itemPartnersOfTabInMenu, "  �������")
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
					GUICtrlSetData($itemJournalOfTabInMenu, "�� ������־")
				Case "��˾��Ʒ"
					If $hasReadedDbTbProducts = False Then
						FuncReadDb($tb_products, $lvInTabProducts)
						$hasReadedDbTbProducts = True
					EndIf
					
					GUICtrlSetState($lvInTabProducts, $GUI_SHOW)
					GUICtrlSetData($itemProductsOfTabInMenu, "�� ��˾��Ʒ")
				Case "ѧУ��Ϣ"
					If $hasReadedDbTbSchools = False Then
						FuncReadDb($tb_schools, $lvInTabSchools)
						$hasReadedDbTbSchools = True
					EndIf
					
					GUICtrlSetState($lvInTabSchools, $GUI_SHOW)
					GUICtrlSetData($itemSchoolsOfTabInMenu, "�� ѧУ��Ϣ")
				Case "���̹���"
					If $hasReadedDbTbProjects = False Then
						FuncReadDb($tb_projects, $lvInTabProjects)
						$hasReadedDbTbProjects = True
					EndIf
					
					GUICtrlSetState($lvInTabProjects, $GUI_SHOW)
					GUICtrlSetData($itemProjectsOfTabInMenu, "�� ���̹���")
				Case "�������"
					If $hasReadedDbTbPartners = False Then
						FuncReadDb($tb_partners, $lvInTabPartners)
						$hasReadedDbTbPartners = True
					EndIf
					
					GUICtrlSetState($lvInTabPartners, $GUI_SHOW)
					GUICtrlSetData($itemPartnersOfTabInMenu, "�� �������")
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
		EndSwitch
		
		_GUICtrlTab_InsertItem ( $tabInMainWindow, $id_item, $strItemName, $id_item)
		
		For $i = 0 To _GUICtrlTab_GetItemCount ( $tabInMainWindow ) - 1 Step 1
			If $strItemName = _GUICtrlTab_GetItemText ( $tabInMainWindow, $i ) Then
				_GUICtrlTab_SetCurSel($tabInMainWindow, $i)
				
				; $itemJournalOfTabInMenu $itemProductsOfTabInMenu $itemSchoolsOfTabInMenu $itemProjectsOfTabInMenu $itemPartnersOfTabInMenu "����˾��Ʒ"
				GUICtrlSetState($lvInTabJournal, $GUI_HIDE)
				GUICtrlSetState($lvInTabProducts, $GUI_HIDE)
				GUICtrlSetState($lvInTabSchools, $GUI_HIDE)
				GUICtrlSetState($lvInTabProjects, $GUI_HIDE)
				GUICtrlSetState($lvInTabPartners, $GUI_HIDE)
				
				Switch $strItemName
					Case "������־"
						If $hasReadedDbTbJournal = False Then
							FuncReadDb($tb_journal, $lvInTabJournal)
							$hasReadedDbTbJournal = True
						EndIf
						
						GUICtrlSetState($lvInTabJournal, $GUI_SHOW)
						GUICtrlSetData($itemJournalOfTabInMenu, "�� ������־")
					Case "��˾��Ʒ"
						If $hasReadedDbTbProducts = False Then
							FuncReadDb($tb_products, $lvInTabProducts)
							$hasReadedDbTbProducts = True
						EndIf
						
						GUICtrlSetState($lvInTabProducts, $GUI_SHOW)
						GUICtrlSetData($itemProductsOfTabInMenu, "�� ��˾��Ʒ")
					Case "ѧУ��Ϣ"
						If $hasReadedDbTbSchools = False Then
							FuncReadDb($tb_schools, $lvInTabSchools)
							$hasReadedDbTbSchools = True
						EndIf
						
						GUICtrlSetState($lvInTabSchools, $GUI_SHOW)
						GUICtrlSetData($itemSchoolsOfTabInMenu, "�� ѧУ��Ϣ")
					Case "���̹���"
						If $hasReadedDbTbProjects = False Then
							FuncReadDb($tb_projects, $lvInTabProjects)
							$hasReadedDbTbProjects = True
						EndIf
						
						GUICtrlSetState($lvInTabProjects, $GUI_SHOW)
						GUICtrlSetData($itemProjectsOfTabInMenu, "�� ���̹���")
					Case "�������"
						If $hasReadedDbTbPartners = False Then
							FuncReadDb($tb_partners, $lvInTabPartners)
							$hasReadedDbTbPartners = True
						EndIf
						
						GUICtrlSetState($lvInTabPartners, $GUI_SHOW)
						GUICtrlSetData($itemPartnersOfTabInMenu, "�� �������")
				EndSwitch
			EndIf
		Next
	EndIf
EndFunc

Func Func_MenuItemGuide_In_MenuHelp ()
EndFunc

Func Func_MenuItemAbout_In_MenuHelp ()
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
	; ��־�����⣺��������ֶ�����
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
			$strFields &= $obj_adodb_recordset.Fields( $i ).value & "|"
		Next
		$strFields = StringTrimRight ( $strFields, 1 )	; ɾ��ĩβ�� | 
		
		; ���� ��
		GUICtrlCreateListViewItem( $strFields, $strListView)
			GUICtrlSetBkColor (-1, 0xffa500 );����listviewitem�ı���ɫ
		$obj_adodb_recordset.movenext
	WEnd
	_GUICtrlListView_EndUpdate($strListView)
	
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

	;=== ��˾��Ա�� tb_workers ��
	;                              ��š������Զ�����|                   �������ı�|  ���֤��|               ����|              ְ��|            ��ְ����|             ת������|            ��н��|            �Ա�|       ����|            �绰|         ����|         סַ��Ĭ��50�ַ�|    ��ͥ��Ϣ|          ��ע
	Local $strColsForTbWorders = "id integer identity(1,1) primary key, w_name text, w_identity_number text, w_deportment text, w_position text, w_date_of_entry text, w_date_regular text, w_month_wage text, w_sex text, w_birthday text, w_phone text, w_email text, w_address text(250), w_family text(254), w_note memo"

	;=== ѧУ��Ϣ�� tb_schools ��
	;                              ���|                                 ѧУ|        ��ϵ��|         ��ַ|                �绰|         ����|        ��ע
	Local $strColsForTbSchools = "id integer identity(1,1) primary key, s_name text, s_contact text, s_address text(250), s_phone text, s_email text, s_note memo"

	;=== ��˾��Ʒ�� tb_products ��
	;                               ���|                                 ��Ʒ|         ����|        ���|             ����|                       ���|        ��ע
	Local $strColsForTbProducts = "id integer identity(1,1) primary key, pd_name text, pd_type text, pd_desiger text, pd_configuration text(250), pd_cost text, pd_note memo"

	;=== �ʲ������ tb_accets ��
	;                             ���|                                 ����|        ����|                 ��λ|        ����|        ��������|           ����|         ��������|          ������|        �Ƿ񱨷�|     ��ע
	Local $strColsForTbAccets = "id integer identity(1,1) primary key, a_name text, a_serial_number text, a_unit text, a_type text, a_date_bought text, a_price text, a_deportment text, a_dealer text, a_scrap text, a_note text(250)"

	;=== �������� tb_partners ��
	;                               ���|                                 ����|         ����|         ��ַ|                 �绰|          ����|          ҵ��|             ��ע
	Local $strColsForTbPartners = "id integer identity(1,1) primary key, pt_name text, pt_type text, pt_address text(250), pt_phone text, pt_email text, pt_business text, pt_note memo"

	;=== ���̹���� tb_projects ��
	;                               ���|                                 ����|         ѧУ|           ��Ʒ|            ������|          ��˾������|     ϸ����Ƭ��|         ��ʼ����|           ״̬|          ��������|            �����¼|             ��ע
	Local $strColsForTbProjects = "id integer identity(1,1) primary key, pj_name text, pj_s_name text, pj_pd_name text, pj_pt_name text, pj_w_name text, pj_content text(250), pj_date_start text, pj_state text, pj_date_finish text, pj_account text(250), pj_note memo"

	;=== ������־�� tb_journal ��
	;                              ���|                                 ��Ա|        ����|        �ص�|                ��ͨ|                ʳ��|              ��������|            ��ע|        ��¼��(����)|        ��¼����(����)
	Local $strColsForTbJournal = "id integer identity(1,1) primary key, j_name text, j_date text, j_address text(250), j_traffic text(250), j_board text(250), j_content text(250), j_note memo, j_record text, j_date_record text"

	;=== �û��� tb_users ��
	;                            ���|                                 �û���|      ����|        Ȩ��|             ��ע
	Local $strColsForTbUsers = "id integer identity(1,1) primary key, u_name text, u_pswd text, u_authority text, u_note text(250)"

	;=== Դ���ݱ� tb_source ��
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
