
'	Name for RC module
' can hold a complete path, or only the name with extension
_:[CSED_COMPIL_RC]:   AxSuite3.rc   ' with ou without double quotes !

'	Name for module to compile :  EXE ; DLL ; LIB (A) ; or OCX ; or O (Obj)
' can hold a complete path, or only the name with extension, or only extension
_:[CSED_COMPIL_NAME]:   AxSuite3_4.exe  ' with ou without double quotes !


'_:[CSED_OPTIMIZE]: 'Private procs , dead code Removal


_:[CSED_CLEAN]:  'Private procs , dead code Removal and Cleaned commented lines

'=========================================================================
'Include files of control
'=========================================================================
#include once "windows.bi"
#include once "win/richedit.bi"
#Include Once "win/ocidl.bi"
#Include Once "win/shellapi.bi"
#include once "win/commctrl.bi"
#include once "win/commdlg.bi"

#Include once "inc_files.bi"
'=========================================================================
'Defined Constants
'=========================================================================
Const szNULL = !"\0"
Const szOpen = "File type : Exe ; Dll ; Ocx  (*.exe ,*.dll ,ocx)" & szNULL & "*.exe;*.dll;*.ocx" & szNULL & "File type : Tlb or Olb    (*.tlb ,*.olb)" & szNULL & "*.tlb;*.olb" & szNULL
Const szSave = "Include File : bi (*.bi)" & szNULL
Const szDll = "File : dll (*.dll)" & szNULL

Const crlf = Chr$(13, 10)
Const dq = Chr$(34)
Const Tb = Chr$(9)
Const d_bc = Chr$(47,39)
Const f_bc = Chr$(39,47)

'=========================================================================
'Defined Types
'=========================================================================
Type TV_HITTESTINFO
   pt                              As POINT
   flags                           As UINT
   hitem                           As HTREEITEM
End Type

Type LPTV_HITTESTINFO As TV_HITTESTINFO ptr


'=========================================================================
'Include file from  original axhelper.bi + strwrap.bi + lv.bi + tv.bi
'=========================================================================
#Include Once "Ax_Utils.bi"                   ' axhelper + strwrap + lv + tv


'=========================================================================
'Defined constant in .rc file.
'Please delete this section if you have included exported names of resource.
'=========================================================================
#define dlgMain 1000
#define IDC_TAB1 1001
#define IDC_STC3 1002
#define IDC_EDT1 1003
#define IDC_SBR1 1004

#define dlgList 1100
#define IDC_LSV1 1101
#define dlgCode 1200
#define IDC_RED1 1201
#Define dlgTree 1300
#define IDC_TRV1 1301
#define dlgAbout 1400
#define IDC_STC1 1402

#define dlg_Ask 1410
#define IDC_BTN21 1412
#define IDC_EDT11 1413
#define IDC_BTN11 1414
#define IDC_STC11 1415

#define IDC_STC31 1430
#define IDC_EDT31 1432


#define IDR_MENU1 10000
#define IDM_TypeLibs 10001
#define tlb_reged 10002
#define tlb_libfile 10003
#define tlb_exit 10004
#define IDM_Code 10005
#define code_const 10006
#define code_Module 10007
#define code_vtable 10008
#define code_Invoke 10009
#define code_event 10010
#define IDM_Help 10011
#define hlp_About 10012
#Define hlp_axsuite 10013
#Define code_init 10014
#Define code_full 10015
#define code_codegen 10016
#define code_save 10017
#Define code_full2 10018
#define code_ident 10019
#Define uOcx1 10020
#Define uOcx2 10021
#Define codll 10022
#Define codll71 10023
#Define aodll 10024
#Define aodll71 10025
#Define regf 10026
#Define unregl 10027
#Define code_Disp 10028
#Define code_Disp2 10029
#Define code_mkDisp2 10030
'#Define Static3 10031
'#Define Static4 10032
#Define Static5 10033
#Define Static6 10034
#define Search_F2 10035
#define Search_F3 10036
#define Search_F4 10037

#define code_Cop 10041
#define code_All 10042
#define code_Not 10043
#define code_Copall 10044


#define As_AxSuite 10060
#define As_DispHelper 10061

#define Copy_GUID 10071
#define Copy_NAME 10072
#define Copy_PATH 10073
#define code_search 1480
#define code_search1 1481
#define code_search2 1482
'=========================================================================
'Global variable
'=========================================================================
dim iccex AS InitCommonControlsEx
Dim Shared as HINSTANCE hinstance
Dim Shared As hwnd htab0, htab1, htab2, hList, hCode, hTree, hmain
Dim Shared as HMENU hPopupMenu
Dim shared hIDC_RED1 as HWND
Dim Shared As Integer lviewsort
dim Shared selpath as zstring*max_path, prefix As ZString*11
Dim shared as string clsids, tldesc,OcxName
Dim Shared hParent(7) As Long, rtab As rect
dim shared As integer ABOUT_TIMER
Dim Shared hdll As Any Ptr
Dim Shared sreport As String, savepath As String
Dim Shared ptli as lptypelib
Dim Shared EvtList As String
Dim Shared searching As String
Dim Shared HWND_FORM1_LISTVIEW1 as hwnd
Dim shared As integer pos0,pos0bak
Dim shared as integer nt1,nowin , icoclass ,isevent ,nook
Dim shared as integer gene1,xret1,xret2,xfull,notsure
Dim Shared searching2 As String, iscod as integer
dim shared as string coclass(),sevent,sinfcode
'=========================================================================
'declarations
'=========================================================================
declare Function FF_ClipboardSetText(ByVal TheText As String) As Long
Declare Sub showinfo()
declare sub dlgMain_IDC_TAB1_Sample(byval hWin as hwnd, byval CTLID as integer)
declare Function dlgMain_Init(Byval hWin as hwnd, Byval wParam as wParam, Byval lParam as lParam) as Integer
declare Function dlgMain_OnSize(Byval hWin as hwnd, Byval lState as long, Byval cx as long, Byval cy as long) as Integer
declare Function dlgMain_OnClose(Byval hWin as hwnd) as Integer
declare Function dlgMain_OnDestroy(Byval hWin as hwnd) as Integer
declare Function dlgMain_OnNotify(Byval hWin as hwnd, Byval lCtlID as long, Byval lpNMHDR as NMHDR ptr, Byref lresult as long) as Integer
declare Function dlgMain_OnTimer(Byval hWin as hwnd, Byval wTimerID as word) as Integer
declare Function dlgMain_IDC_TAB1_SelChange(byval hWin as hwnd, byval lpNMHDR as NMHDR ptr, Byref lresult as long) as long
declare Function dlgMain_IDC_TAB1_SelChanging(byval hWin as hwnd, byval lpNMHDR as NMHDR ptr, Byref lresult as long) as long
declare Function dlgMain_IDR_MENU1_menu(byval hwin as hwnd, byval id as long) as long
declare sub dlgList_IDC_LSV1_Sample(byval hWin as hwnd, byval CTLID as integer)
declare Function dlgList_Init(Byval hWin as hwnd, Byval wParam as wParam, Byval lParam as lParam) as Integer
declare Function dlgList_OnSize(Byval hWin as hwnd, Byval lState as long, Byval cx as long, Byval cy as long) as Integer
declare Function dlgList_OnClose(Byval hWin as hwnd) as Integer
declare Function dlgList_OnDestroy(Byval hWin as hwnd) as Integer
declare Function dlgList_OnNotify(Byval hWin as hwnd, Byval lCtlID as long, Byval lpNMHDR as NMHDR ptr, Byref lresult as long) as Integer
declare Function dlgList_OnTimer(Byval hWin as hwnd, Byval wTimerID as word) as Integer
declare function dlgList_IDC_LSV1_Clicked(Byval hWin as hwnd, Byval pNMLV as NMLISTVIEW ptr) as long
declare Function dlgCode_Init(Byval hWin as hwnd, Byval wParam as wParam, Byval lParam as lParam) as Integer
declare Function dlgCode_OnSize(Byval hWin as hwnd, Byval lState as long, Byval cx as long, Byval cy as long) as Integer
declare Function dlgCode_OnClose(Byval hWin as hwnd) as Integer
declare Function dlgCode_OnDestroy(Byval hWin as hwnd) as Integer
declare Function dlgCode_OnNotify(Byval hWin as hwnd, Byval lCtlID as long, Byval lpNMHDR as NMHDR ptr, Byref lresult as long) as Integer
declare Function dlgCode_OnTimer(Byval hWin as hwnd, Byval wTimerID as word) as Integer
declare Function dlgTree_Init(Byval hWin as hwnd, Byval wParam as wParam, Byval lParam as lParam) as Integer
declare Function dlgTree_OnSize(Byval hWin as hwnd, Byval lState as long, Byval cx as long, Byval cy as long) as Integer
declare Function dlgTree_OnClose(Byval hWin as hwnd) as Integer
declare Function dlgTree_OnDestroy(Byval hWin as hwnd) as Integer
declare Function dlgTree_OnNotify(Byval hWin as hwnd, Byval lCtlID as long, Byval lpNMHDR as NMHDR ptr, Byref lresult as long) as Integer
declare Function dlgTree_OnTimer(Byval hWin as hwnd, Byval wTimerID as word) as Integer
declare Function dlgAbout_Init(Byval hWin as hwnd, Byval wParam as wParam, Byval lParam as lParam) as Integer
declare Function dlgAbout_OnSize(Byval hWin as hwnd, Byval lState as long, Byval cx as long, Byval cy as long) as Integer
declare Function dlgAbout_OnClose(Byval hWin as hwnd) as Integer
declare Function dlgAbout_OnDestroy(Byval hWin as hwnd) as Integer
declare Function dlgAbout_OnNotify(Byval hWin as hwnd, Byval lCtlID as long, Byval lpNMHDR as NMHDR ptr, Byref lresult as long) as Integer
declare Function dlgAbout_OnTimer(Byval hWin as hwnd, Byval wTimerID as word) as Integer
declare Function dlgMain_DlgProc(byval hWin as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM) as integer
declare Function dlgList_DlgProc(byval hWin as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM) as integer
declare Function dlgCode_DlgProc(byval hWin as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM) as integer
declare Function dlgTree_DlgProc(byval hWin as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM) as Integer
declare Function dlgAbout_DlgProc(byval hWin as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM) as Integer
declare Function dlgAsk_DlgProc(byval hWin as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM) as Integer
declare Function AskMain(ByVal hwParent As hwnd) As Integer
declare Function AskMain2(ByVal hwParent As hwnd) As Integer
declare Function StringToBSTR1(cnv_string As String) As BSTR
declare Function TreeView_GetItemText1(ByVal hTreeView As Long, ByVal hItem As Long) As String
declare sub GetCopy1(posi1 as integer)
declare function tab_pos(hWin1 as hwnd) as integer
declare function FF_ListView_SetSelectedItem(byval hlist as hwnd, byval i as integer) as integer
declare function dlgList_IDC_LSV1_Click1(Byval hWin as hwnd, Byval pNMLV as NMLISTVIEW ptr) as long
declare function force_click(hWin as hwnd) as integer
declare Function FF_ListView_GetItemText(ByVal hWndControl As hWnd, ByVal iRow As Integer, ByVal iColumn As Integer ) As String
declare Function FF_ListView_GetSelectedItem(ByVal hWndControl As hWnd) As Integer
declare sub search1()
'=========================================================================
'Include files of code
'=========================================================================
#Include Once "tlbcode1.bi"
#Include Once "tlib1.bi"
#Include once "dlgMain1.bi"

'=========================================================================


Function dlgCode_Init(Byval hWin as hwnd, Byval wParam as wParam, Byval lParam as lParam) as Integer
   hcode = getdlgitem(hwin, idc_red1)
   sendmessage hcode, wm_settext, 0, 0
   function = 0
End Function

Function dlgCode_OnSize(Byval hWin as hwnd, Byval lState as long, Byval cx as long, Byval cy as long) as Integer
   movewindow(getdlgitem(hwin, idc_red1), 0, 0, cx, cy, true)
   function = 0
End Function

Function dlgCode_OnClose(Byval hWin as hwnd) as Integer
   function = 0
End Function

Function dlgCode_OnDestroy(Byval hWin as hwnd) as Integer
   function = 0
End Function

Function dlgCode_OnNotify(Byval hWin as hwnd, Byval lCtlID as long, Byval lpNMHDR as NMHDR ptr, Byref lresult as long) as Integer
   function = 0
End Function

Function dlgCode_OnTimer(Byval hWin as hwnd, Byval wTimerID as word) as Integer
   function = 0
End Function

Function dlgTree_Init(Byval hWin as hwnd, Byval wParam as wParam, Byval lParam as lParam) as Integer
   htree = getdlgitem(hwin, idc_trv1)
   function = 0
End Function

Function dlgTree_OnSize(Byval hWin as hwnd, Byval lState as long, Byval cx as long, ByVal cy as long) as Integer
   movewindow(htree, 0, 0, cx, cy, true)
   function = 0
End Function

Function dlgTree_OnClose(Byval hWin as hwnd) as Integer
   function = 0
End Function

Function dlgTree_OnDestroy(Byval hWin as hwnd) as Integer
   function = 0
End Function



function IDD_DLG1_idc_trv1_chk(Byval hWin as hwnd, ByVal hitem as integer) as Long
   Dim tstate As Integer
   tstate = TreeView_GetCheckState(cast(long, hwin), hitem)
   Treeview_ChangeChildState(cast(long, hwin), hitem, tstate)
   Treeview_ChangeParentState(cast(long, hwin), hitem, 0)
   function = 0
End Function

Function dlgTree_OnNotify(Byval hWin as hwnd, Byval lCtlID as long, Byval lpNMHDR as NMHDR ptr, Byref lresult as long) as Integer
   Dim As Long event, itmp
   Dim ht As TV_HITTESTINFO
   Dim As dword dwpos
   event = lpnmhdr -> code
   if lctlid = idc_trv1 Then
      If event = NM_CLICK Then
         dwpos = GetMessagePos()
         iTmp = LoWord(dwpos)
         ht.pt.x = iTmp
         iTmp = HiWord(dwpos)
         ht.pt.y = iTmp
         MapWindowPoints(HWND_DESKTOP, lpnmhdr -> hWndFrom, @ht.pt, 1)
         TreeView_HitTest(lpnmhdr -> hwndFrom, @ht)
         If ht.flags = TVHT_ONITEMSTATEICON Then
            IDD_DLG1_idc_trv1_chk(lpnmhdr -> hWndFrom, cast(long, ht.hItem))
				'IDD_DLG1_idc_trv1_chk(lpnmhdr -> hWndFrom,ht.hItem)
         End If
      End If
   End If
   function = 0
End Function

Function dlgTree_OnTimer(Byval hWin as hwnd, Byval wTimerID as word) as Integer
   function = 0
End Function

sub dlgList_IDC_LSV1_Sample(byval hWin as hwnd, byval CTLID as integer)
   'dim hCTLID as hwnd
   dim dwStyle as dword
   hlist = getdlgitem(hwin, CTLID)
   dwStyle = SendMessage(hlist, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
   dwStyle = dwStyle OR LVS_REPORT or LVS_EX_FULLROWSELECT
   sendmessage hlist, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, dwStyle
   lvsetmaxcols(cast(long, hlist), 4)

   lvsetheader(cast(long, hlist), "GUID|Description|File|Path")
End sub

FUNCTION compare(BYVAL lParam1 AS LONG, BYVAL lParam2 AS LONG, BYVAL lParamsort AS LONG) AS Integer
   dim value1 AS zstring * 300
   dim value2 AS zstring * 300
   dim j as integer

   FUNCTION = 0

   value1 = ucase(LVGetValue(cast(long, hlist), lParam1, lParamsort))
   value2 = ucase(LVGetValue(cast(long, hlist), lParam2, lParamsort))
   SELECT CASE lviewSort
      CASE 0                                     'ascending
         j = 1
      CASE ELSE                                  'descending
         j = - 1
   END SELECT
   IF value1 < value2 THEN
      FUNCTION = - 1 * j
   ELSEIF value1 = value2 THEN
      FUNCTION = 0
   ELSE
      FUNCTION = 1 * j
   END IF
END Function

SUB ResetLParam(ByRef hList AS LONG)
   dim i AS LONG
   dim recs AS LONG
   dim lvi AS LV_ITEM
   dim x AS LONG

   lvi.mask = LVIF_PARAM
   lvi.iSubItem = 0
   recs = sendmessage(cast(hwnd, hlist), lvm_GetItemCount, 0, 0)
   FOR i = 0 TO recs - 1
      lvi.iItem = i
      x = SendMessage(cast(hwnd, hlist), LVM_GetItem, 0, cast(lparam, @lvi))
      lvi.lParam = lvi.iItem
      x = SendMessage(cast(hwnd, hlist), LVM_SetItem, 0, cast(lparam, @lvi))
   NEXT
END SUB

Function dlgList_Init(Byval hWin as hwnd, Byval wParam as wParam, Byval lParam as lParam) as Integer
   dlgList_IDC_LSV1_Sample(hWin, IDC_LSV1)
   function = 0
End Function

Function dlgList_OnSize(Byval hWin as hwnd, Byval lState as long, Byval cx as long, Byval cy as long) as Integer
   movewindow(getdlgitem(hwin, idc_lsv1), 0, 0, cx, cy, true)
   function = 0
End Function

Function dlgList_OnClose(Byval hWin as hwnd) as Integer
   function = 0
End Function

Function dlgList_OnDestroy(Byval hWin as hwnd) as Integer
   function = 0
End Function

Function dlgList_OnNotify(Byval hWin as hwnd, Byval lCtlID as long, Byval lpNMHDR as NMHDR ptr, Byref lresult as long) as Integer
   Dim pnmlv As NMLISTVIEW Ptr
   Dim item As Integer
   SELECT CASE lpNMHdr -> idFrom
      CASE idc_LSV1
         pNMLV = cast(NMLISTVIEW Ptr, lpnmhdr)
         SELECT CASE pNmlv -> hdr.Code
            Case LVN_COLUMNCLICK
               IF pNmLv -> iSubItem <> - 1 THEN
                  'toggle ascending/descending
                  lviewSort = NOT lviewSort
                  ListView_SortItems(hList, procPTR(compare), pNmLv -> iSubItem)
                  ResetLParam cast(long, hList)
               END If
         End select
   End select
   function = 0
End Function

Function dlgList_OnTimer(Byval hWin as hwnd, Byval wTimerID as word) as Integer
   function = 0
End Function

function dlgList_IDC_LSV1_Clicked(Byval hWin as hwnd, Byval pNMLV as NMLISTVIEW ptr) as long
   dim hIDC_LSV1 as hwnd
   dim item as Long, lret As Long
	dim as string act1
   hIDC_LSV1 = GetDlgItem(hWin, IDC_LSV1)
   item = pnmlv -> iitem
   selpath = lvgetvalue(cast(long, hlist), item, 3)
   clsids = LVGetvalue(cast(long, hlist), item, 0)
   tldesc = LVGetvalue(cast(long, hlist), item, 1)
	act1=LVGetvalue(cast(long, hlist), item, 2)
   IF LEN(SelPath) Then
      cleartli
      lret = LoadtypeLib(stringtobstr(selpath), @ptli)
      IF lret Then
         MessageBox BYVAL NULL, "Unable to load " & SelPath, _
               "Error", MB_OK OR MB_ICONERROR OR MB_TASKMODAL
      Else
			'act1=" Selected :  " & act1 & "    |    Path :  " & selpath
         'sendmessage GetDlgItem(hmain, IDC_SBR1),WM_SETTEXT ,0,cast(lparam,strptr(act1))
			setwindowtext GetDlgItem(hmain, IDC_SBR1)," Selected :  " & act1 & "    |    Path :  " & selpath
			LoadTLI
         chgtab(1)
         treeview_select(htree, hparent(0), TVGN_CARET)
      End If
   End If
   gene1 = 0
	searching=""
   function = 0
end function


function dlgList_IDC_LSV1_Click1(Byval hWin as hwnd, Byval pNMLV as NMLISTVIEW ptr) as long
   dim hIDC_LSV1 as hwnd
   dim item as Long, lret As Long
	dim as string act1
   hIDC_LSV1 = GetDlgItem(hWin, IDC_LSV1)
   item = pnmlv -> iitem
   selpath = lvgetvalue(cast(long, hlist), item, 3)
   clsids = LVGetvalue(cast(long, hlist), item, 0)
   tldesc = LVGetvalue(cast(long, hlist), item, 1)
	act1=LVGetvalue(cast(long, hlist), item, 2)
   IF LEN(SelPath) Then
      cleartli
      lret = LoadtypeLib(stringtobstr(selpath), @ptli)
      IF lret Then
         MessageBox BYVAL NULL, "Unable to load " & SelPath, _
               "Error", MB_OK OR MB_ICONERROR OR MB_TASKMODAL
      Else
			'act1=" Selected :  " & act1 & "    |    Path :  " & selpath
         'sendmessage GetDlgItem(hmain, IDC_SBR1),WM_SETTEXT ,0,cast(lparam,strptr(act1))
			setwindowtext GetDlgItem(hmain, IDC_SBR1)," Selected :  " & act1 & "    |    Path :  " & selpath
         LoadTLI
         chgtab1()
         treeview_select(htree, hparent(0), TVGN_CARET)
         'chgtab(0)

      End If
   End If
   gene1 = 0
	searching=""
   function = 0
end function


function makefont(byval lCyPixels as long, sFntName as string, byval fFntSize as long, byval lWeight as long, byval lUnderlined as long, byval lItalic as long, byval lStrike as long, byval lCharSet as long) as long
   dim lf as LOGFONT
   lF.lfHeight = - (fFntSize * lCyPixels) \ 69
   lF.lfFaceName = sFntName
   IF lWeight = FW_DONTCARE THEN
      lf.lfWeight = FW_NORMAL
   ELSE
      lf.lfWeight = lWeight
   END IF
   lf.lfUnderline = lUnderlined
   lf.lfItalic = lItalic
   lf.lfStrikeOut = lStrike
   lf.lfCharSet = lCharSet
   'create font and assign handle to a global 'variable
   function = cast(long, CreateFontIndirect(BYVAL VARPTR(lF)))
end Function

Function dlgAbout_Init(Byval hWin as hwnd, Byval wParam as wParam, Byval lParam as lParam) as Integer
   Dim text As String
   Dim As Integer hIcon, hFont, hDC

   SetTimer hwin, ABOUT_TIMER, 5000, NULL

   hIcon = cast(integer, LoadIcon(0, IDI_ASTERISK))
   SendMessage hwin, WM_SETICON, ICON_BIG, hIcon

   hDC = cast(integer, GetDC(0))
   hfont = MakeFont(GetDeviceCaps(cast(hdc, hDC), LOGPIXELSY), "Times New Roman", 20, fw_bold, false, false, FALSE, ANSI_CHARSET)
   ReleaseDC 0, cast(hdc, hDC)
   SendMessage GetDlgItem(hwin, idc_stc1), WM_SETFONT, hFont, 0
   function = 0
End Function

Function dlgAbout_OnSize(Byval hWin as hwnd, Byval lState as long, Byval cx as long, Byval cy as long) as Integer
   function = 0
End Function

Function dlgAbout_OnClose(Byval hWin as hwnd) as Integer
   function = 0
End Function

Function dlgAbout_OnDestroy(Byval hWin as hwnd) as Integer
   function = 0
End Function






'=========================================================================
'Initial code of control
'=========================================================================
iccex.dwSize = SIZEOF(iccex)
iccex.dwICC = ICC_TAB_CLASSES
InitCommonControlsEx(@iccex)
iccex.dwSize = SIZEOF(iccex)
iccex.dwICC = ICC_LISTVIEW_CLASSES
InitCommonControlsEx(@iccex)
iccex.dwSize = SIZEOF(iccex)
iccex.dwICC = ICC_TREEVIEW_CLASSES
InitCommonControlsEx(@iccex)
hIDC_RED1 = dylibload( "riched20.dll")

DialogBoxParam(GetModuleHandle(NULL), Cast(zstring ptr, dlgMain), NULL, @dlgMain_DlgProc, NULL)

DylibFree(hdll)
ExitProcess(0)
end

'=========================================================================
'Dialog callback procedure
'=========================================================================
Function dlgMain_DlgProc(byval hWin as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM) as integer
   dim as long id, event, lresult
   dim lpNMHDR as NMHDR ptr
   dim nBuffer AS string = space(250)
   dim itab0 as integer

   select case uMsg
      case WM_INITDIALOG
         if dlgMain_Init(hWin, wparam, lparam) then
            function = false
            exit function
         end if
      case WM_COMMAND
         HWND_FORM1_LISTVIEW1 = hlist
         id = loword(wParam)
         event = hiword(wParam)
         select CASE id
            case IDC_STC3
               'messagebox 0, "lParam = " & lparam ,"prefix", mb_ok
               force_click(hWin)
               exit function
            case IDC_STC31
               if tab_pos(hWin) <> 0 THEN exit function
               'messagebox 0, "search" ,"info", mb_ok
               GetWindowText GetDlgItem(hWin, IDC_EDT31), nBuffer, 250
               nBuffer = trim(nBuffer)
               if searching <> nBuffer and nBuffer <> "" THEN
                  searching = nBuffer
                  search1()
               else
                  next1()
               END IF
               exit function
            case Search_F2
					'messagebox 0 ,"Search","DDD" ,0
					chgtab(0)
               'if tab_pos(hWin) <> 0 THEN exit function
               AskMain(hWin)
               SetWindowText GetDlgItem(hWin, IDC_EDT31), searching
               if searching <> "" and nook=0 THEN
                  search1()
               END IF
					nook=0
               exit function
            case Search_F3
					chgtab(0)
               'if tab_pos(hWin) <> 0 THEN exit function
               GetWindowText GetDlgItem(hWin, IDC_EDT31), nBuffer, 250
               nBuffer = trim(nBuffer)
               if searching <> nBuffer and nBuffer <> "" THEN
                  searching = nBuffer
                  search1()
                elseif nBuffer <> "" THEN
                  next1()
               END IF
               exit function
            case Search_F4
					chgtab(0)
               'if tab_pos(hWin) <> 0 THEN exit function
               GetWindowText GetDlgItem(hWin, IDC_EDT31), nBuffer, 250
               nBuffer = trim(nBuffer)
               if searching <> nBuffer and nBuffer <> "" THEN
                  searching = nBuffer
                  search1()
               elseif nBuffer <> "" THEN
                  previous1()
               END IF
               exit function

         END SELECT
         If dlgMain_IDR_MENU1_menu(hwin, id) Then
            function = false
            exit Function
         End If

      case WM_SIZE
         if dlgMain_OnSize(hWin, wparam, loword(lparam), hiword(lparam)) then
            function = false
            exit function
         end if

      case WM_NOTIFY
         lpNMHDR = cast(NMHDR ptr, lparam)
         id = lpnmhdr -> idfrom
         event = lpnmhdr -> code
         if id = IDC_TAB1 then
            if event = TCN_SELCHANGE then
               'A tab selection is made
               if dlgMain_IDC_TAB1_SelChange(hWin, lpNMHDR, lresult) then
                  function = lresult
                  exit function
               end if
            end if
            if event = TCN_SELCHANGING then
               if dlgMain_IDC_TAB1_SelChanging(hWin, lpNMHDR, lresult) then
                  function = lresult
                  exit function
               end if
            end if
         end if
         if dlgMain_OnNotify(hWin, wparam, cast(NMHDR ptr, lparam), lresult) then
            function = lresult
            exit function
         end if
      case WM_TIMER
         if dlgMain_OnTimer(hWin, wparam) then
            function = false
            exit function
         end if
      case WM_CLOSE
         if dlgMain_OnClose(hWin) then
            function = false
            exit function
         end if
         EndDialog(hWin, 0)
      case WM_DESTROY
         dlgMain_OnDestroy(hWin)
         function = false
         exit function
      case else
         return FALSE
   end select
   return TRUE
End Function

Function dlgList_DlgProc(byval hWin as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM) as integer
   dim as long id, event, lresult
   dim lpNMHDR as NMHDR ptr
   dim pNMLV AS NMLISTVIEW ptr
   dim as string mess1, mess2
   dim nBuffer AS string = space(250)

   select case uMsg
      case WM_INITDIALOG
         if dlgList_Init(hWin, wparam, lparam) then
            function = false
            exit function
         end if
      case WM_COMMAND
         id = loword(wParam)
         event = hiword(wParam)
         HWND_FORM1_LISTVIEW1 = hList
         SELECT CASE id
            case Search_F2
               AskMain(hWin)
               SetWindowText GetDlgItem(hmain, IDC_EDT31), searching
               if searching <> "" and nook=0 THEN
                  search1()
               END IF
					nook=0
               exit function
            case Search_F3
               GetWindowText GetDlgItem(hmain, IDC_EDT31), nBuffer, 250
               nBuffer = trim(nBuffer)
               if searching <> nBuffer and nBuffer <> "" THEN
                  searching = nBuffer
                  search1()
               else
                  next1()
               END IF
               exit function
            case Search_F4
               GetWindowText GetDlgItem(hmain, IDC_EDT31), nBuffer, 250
               nBuffer = trim(nBuffer)
               if searching <> nBuffer and nBuffer <> "" THEN
                  searching = nBuffer
                  search1()
               else
                  previous1()
               END IF
               exit function
            case Copy_GUID
               GetCopy1 0
               exit function
            case Copy_NAME
               GetCopy1 2
               exit function
            case Copy_PATH
               GetCopy1 3
               exit function
         end select
      case WM_SIZE
         if dlgList_OnSize(hWin, wparam, loword(lparam), hiword(lparam)) then
            function = false
            exit function
         end if
      case WM_CONTEXTMENU
         If hPopupMenu Then DestroyMenu hPopupMenu
         hPopupMenu = CreatePopupMenu()
         AppendMenu hPopupMenu, MF_STRING, Search_F2, "New Search... "
         AppendMenu hPopupMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopupMenu, MF_STRING, Search_F3, "Find Next "
         AppendMenu hPopupMenu, MF_STRING, Search_F4, "Find Previous "
         AppendMenu hPopupMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopupMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopupMenu, MF_STRING, Copy_GUID, "Copy GUID "
         AppendMenu hPopupMenu, MF_STRING, Copy_NAME, "Copy File Name "
         AppendMenu hPopupMenu, MF_STRING, Copy_PATH, "Copy Full Path "
         AppendMenu hPopupMenu, MF_SEPARATOR, 0, ""
         TrackPopupMenu hPopupMenu, TPM_LEFTALIGN Or TPM_LEFTBUTTON, loword(lparam), hiword(lparam), 0, hWin, 0
		case WM_NOTIFY
         lpNMHDR = cast(NMHDR ptr, lparam)
         id = lpnmhdr -> idfrom
         event = lpnmhdr -> code
         if id = IDC_LSV1 then
            pNMLV = cast(NMLISTVIEW ptr, lpNMHDR)
            if pNMLV -> hdr.code = NM_DBLCLK then
               if dlgList_IDC_LSV1_Clicked(hWin, cast(NMLISTVIEW ptr, lparam)) then
                  function = false
                  exit function
               end if
            elseif pNMLV -> hdr.code = NM_CLICK then
               if dlgList_IDC_LSV1_Click1(hWin, cast(NMLISTVIEW ptr, lparam)) then
                  function = false
                  exit function
               end if
				elseif pNMLV -> hdr.code = NM_RCLICK then
               if dlgList_IDC_LSV1_Click1(hWin, cast(NMLISTVIEW ptr, lparam)) then
                  function = false
                  exit function
               end if
            end if
         end if
         if dlgList_OnNotify(hWin, wparam, cast(NMHDR ptr, lparam), lresult) then
            function = lresult
            exit function
         end if
      case WM_TIMER
         if dlgList_OnTimer(hWin, wparam) then
            function = false
            exit function
         end if
      case WM_CLOSE
         if dlgList_OnClose(hWin) then
            function = false
            exit function
         end if
         EndDialog(hWin, 0)
			exit function
      case WM_DESTROY

         dlgList_OnDestroy(hWin)
         function = false
         exit function
      case else
         return FALSE
   end select
   return TRUE
End Function

sub cop_all()
   SendMessage hcode, EM_SETSEL, 0, - 1
   SendMessage hcode, WM_COPY, 0, 0
   SendMessage hcode, EM_SETSEL, - 1, 0
   SendMessage hcode, EM_SETSEL, 0, 0
END SUB

sub cod21(byval iv as integer)

	dim D as GetTextLengthEX, T as GETTEXTEX, iLengthText as integer ,buf as string
	D.flags = GTL_DEFAULT
	D.codepage = CP_ACP
   'return number of characters
   iLengthText = SendMessage(hcode, EM_GetTextLengthEX, cast(wparam,VarPTR(D)),0)
	'print " iLengthText = " ;iLengthText & "  len(sreport) = "; len(sreport)
	buf = Space(iLengthText)

	T.cb = iLengthText + 1
	T.flags = GT_DEFAULT
	T.codepage = CP_ACP
	T.lpDefaultChar = NULL
	T.lpUsedDefChar = NULL
	SendMessage hcode, EM_GETTEXTEX, cast(wparam,varptr(T)), cast(lparam,StrPTR(buf))
'   iStartPos = 0
'   iStopPos = iLengthText - 1
'   iStartLine& = 0
'   iStopLine& = SendMessage(hEdit, %EM_GetLineCount, 0,0) - 1

'   T.Chrg.cpmin = 0: T.Chrg.cpmax = iLengthText - 1
'   SendMessage(hcode, EM_GetTextEX, cast(wparam,VarPTR(T)), cast(lparam,StrPTR(buf)))   'returns all chars in buf
	'print buf

	if iv = 0 THEN
		iscod= Instr(iscod +1 , ucase(buf), ucase(searching2))
		'print " + " ;iscod
	elseif iv = 1 THEN
		iscod= Instrrev( ucase(buf), ucase(searching2),iscod -1)
		'print " - " ;iscod
   END IF
   If iscod Then
		SendMessage hcode, EM_SetSel , iscod-1, iscod + Len(searching2)-1  '+/-1 to convert to zero-based
		SendMessage hcode, EM_SCROLLCARET, 0, 0
	else
		messagebox 0, searching2 & "  : no more to find ! " , "Warning", mb_ok
	END IF
END SUB

Function dlgCode_DlgProc(byval hWin as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM) as integer
   dim as long id, event, lresult
   dim lpNMHDR as NMHDR ptr
   dim nBuffer AS string = space(250)

   select case uMsg
      case WM_INITDIALOG
         if dlgCode_Init(hWin, wparam, lparam) then
            function = false
            exit function
         end if
      case WM_COMMAND
         id = loword(wParam)
         event = hiword(wParam)
         SELECT CASE id
            case code_Cop
               SendMessage hcode, WM_COPY, 0, 0
               exit function
            case code_Copall
               cop_all()
               exit function
            case code_All
               SendMessage hcode, EM_SETSEL, 0, - 1
               exit function
            case code_save
               dlgMain_IDR_MENU1_menu(hwin, code_save)
               exit function
            case code_Not
               SendMessage hcode, EM_SETSEL, - 1, 0
               exit function
				case code_Search
					AskMain2(hWin)
					if searching2 <> "" and searching2 <>"Searching Item" and nook=0 THEN cod21(0)
					nook=0
					exit function
				case code_Search1
					if searching2 <> "" and searching2 <>"Searching Item" THEN
						cod21(0)
					else
						AskMain2(hWin)
						if searching2 <> "" and searching2 <>"Searching Item" and nook=0 THEN	cod21(0)
						nook=0
               END IF
					exit function
				case code_Search2
					if searching2 <> "" and searching2 <>"Searching Item" THEN
						'print "reverse"
						cod21(1)
					else
						AskMain2(hWin)
						if searching2 <> "" and searching2 <>"Searching Item" and nook=0 THEN	cod21(0)
						nook=0
               END IF
					exit function
         end select
      case WM_SIZE
         if dlgCode_OnSize(hWin, wparam, loword(lparam), hiword(lparam)) then
            function = false
            exit function
         end if
      case WM_CONTEXTMENU
         If hPopupMenu Then DestroyMenu hPopupMenu
         hPopupMenu = CreatePopupMenu()
         AppendMenu hPopupMenu, MF_STRING, code_Cop, "Copy Selection   / Ctrl-C"
         AppendMenu hPopupMenu, MF_STRING, code_Copall, "Copy All"
         AppendMenu hPopupMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopupMenu, MF_STRING, code_save, "Save All to File"
         AppendMenu hPopupMenu, MF_SEPARATOR, 0, ""
         AppendMenu hPopupMenu, MF_STRING, code_All, "Select All            / Ctrl-A"
         AppendMenu hPopupMenu, MF_STRING, code_Not, "Unselect"
			AppendMenu hPopupMenu, MF_SEPARATOR, 0, ""
			AppendMenu hPopupMenu, MF_STRING, code_Search, "New Search..."
			AppendMenu hPopupMenu, MF_STRING, code_Search1, "Next"
			AppendMenu hPopupMenu, MF_STRING, code_Search2, "Previous"
         TrackPopupMenu hPopupMenu, TPM_LEFTALIGN Or TPM_LEFTBUTTON, loword(lparam), hiword(lparam), 0, hWin, 0
      case WM_NOTIFY
         lpNMHDR = cast(NMHDR ptr, lparam)
         id = lpnmhdr -> idfrom
         event = lpnmhdr -> code

         if dlgCode_OnNotify(hWin, wparam, cast(NMHDR ptr, lparam), lresult) then
            function = lresult
            exit function
         end if
      case WM_TIMER
         if dlgCode_OnTimer(hWin, wparam) then
            function = false
            exit function
         end if
      case WM_CLOSE
         if dlgCode_OnClose(hWin) then
            function = false
            exit function
         end if
         EndDialog(hWin, 0)
      case WM_DESTROY
         'dylibfree(hIDC_RED1)

         dlgCode_OnDestroy(hWin)
         function = false
         exit function
      case else
         return FALSE
   end select
   return TRUE
End Function


Function dlgTree_DlgProc(byval hWin as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM) as integer
   dim as long id, event, lresult
   dim lpNMHDR as NMHDR ptr

   select case uMsg
      case WM_INITDIALOG
         if dlgTree_Init(hWin, wparam, lparam) then
            function = false
            exit function
         end if
      case WM_COMMAND
         id = loword(wParam)
         event = hiword(wParam)
      case WM_SIZE
         if dlgTree_OnSize(hWin, wparam, loword(lparam), hiword(lparam)) then
            function = false
            exit function
         end if
      case WM_NOTIFY
         lpNMHDR = cast(NMHDR ptr, lparam)
         id = lpnmhdr -> idfrom
         event = lpnmhdr -> code
         if dlgTree_OnNotify(hWin, wparam, cast(NMHDR ptr, lparam), lresult) then
            function = lresult
            exit function
         end if
      case WM_TIMER
         if dlgTree_OnTimer(hWin, wparam) then
            function = false
            exit function
         end if
      case WM_CLOSE
         if dlgTree_OnClose(hWin) then
            function = false
            exit function
         end if
         EndDialog(hWin, 0)
      case WM_DESTROY
         dlgTree_OnDestroy(hWin)
         function = false
         exit function
      case else
         return FALSE
   end select
   return TRUE
End Function


Function dlgAbout_DlgProc(byval hWin as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM) as integer
   dim as long id, event, lresult
   dim lpNMHDR as NMHDR ptr

   select case uMsg
      case WM_INITDIALOG
         if dlgAbout_Init(hWin, wparam, lparam) then
            function = false
            exit function
         end if
      case WM_CLOSE
         if dlgAbout_OnClose(hWin) then
            function = false
            exit function
         end if
         EndDialog(hWin, 0)
      case WM_DESTROY
         dlgAbout_OnDestroy(hWin)
         function = false
         exit function
      case else
         return FALSE
   end select
   return TRUE
End Function


function tab_pos(hWin1 as hwnd) as integer
   dim TabID as integer
   dim hIDC_TAB1 as hwnd
   hIDC_TAB1 = getdlgitem(hWin1, IDC_TAB1)
   TabID = SendMessage(hIDC_TAB1, TCM_GETCURSEL, 0, 0)
   'messagebox 0 , "Tab = " & str(TabID),"Tab",mb_ok
   function = TabID
END FUNCTION



sub GetCopy1(posi1 as integer)
   dim as string sitem
   dim as integer irow = FF_ListView_GetSelectedItem(hList)
   'messagebox 0 , "selected pos = " & str(irow),"Warning",mb_ok
   if irow > - 1 THEN
      sitem = FF_ListView_GetItemText(hList, irow, posi1)
      FF_ClipboardSetText(sitem)
   else
      messagebox 0, "Nothing selected !", "Warning", mb_ok
   END IF
end sub

Function AskMain(ByVal hwParent As hwnd) As Integer
   if tab_pos(hwParent) = 0 then
      DialogBoxParam(GetModuleHandle(NULL), Cast(ZString Ptr, dlg_Ask), hwParent, @dlgAsk_DlgProc, 0)
      'messagebox 0,searching,"retour",mb_ok
      Return True
   Else
      messagebox 0, "Not in the list view tab !", "Warning", mb_ok
   End if
   Return False
End Function

Function dlgAsk_DlgProc(ByVal hWin As hWnd, ByVal uMsg As Uinteger, _
         ByVal wParam1 As wParam, _
         ByVal lParam1 As lParam) As Integer
   Dim AS Integer id
   Dim AS Integer Event1
	nook=0
   dim nBuffer AS string = space(250)
   dim info as string
   Select Case uMsg
      Case WM_INITDIALOG
         SelPath = ""
         GetWindowText GetDlgItem(hmain, IDC_EDT31), nBuffer, 250
         nBuffer = trim(nBuffer)
         if searching <> nBuffer and nBuffer <> "" THEN searching = nBuffer
         info = searching
         if info = "" THEN info = "Searching Item"
         SetWindowText GetDlgItem(hWin, IDC_EDT11), info
         SendMessage GetDlgItem(hWin, IDC_EDT11), EM_SETSEL, 0, - 1
			setfocus(GetDlgItem(hWin, IDC_EDT11))
			exit function
      Case WM_CLOSE
         SendMessage getdlgitem(hmain, idc_edt31), EM_SETSEL, 0, - 1
         setfocus getdlgitem(hmain, idc_edt31)
         EndDialog(hWin, 0)
			exit function
      Case WM_COMMAND
         id = LoWord(wParam1)
         Event1 = HiWord(wParam1)
         Select Case id
            Case IDC_BTN11
               GetWindowText GetDlgItem(hWin, IDC_EDT11), nBuffer, 250
               searching = trim(nBuffer)
               SendMessage getdlgitem(hmain, idc_edt31), EM_SETSEL, 0, - 1
               setfocus getdlgitem(hmain, idc_edt31)
               EndDialog(hWin, 0)
					exit function
            Case IDC_BTN21
               SendMessage getdlgitem(hmain, idc_edt31), EM_SETSEL, 0, - 1
               setfocus getdlgitem(hmain, idc_edt31)
					nook=1
               EndDialog(hWin, 0)
					exit function
         End Select
      Case Else
         Return FALSE
   End Select
   Return TRUE
End Function

Function dlgAsk_DlgProc2(ByVal hWin As hWnd, ByVal uMsg As Uinteger, _
         ByVal wParam1 As wParam, _
         ByVal lParam1 As lParam) As Integer
   Dim AS Integer id
   Dim AS Integer Event1

   dim nBuffer AS string = space(250)
   dim info as string
   Select Case uMsg
      Case WM_INITDIALOG
         info = searching2
         if info = "" THEN info = "Searching Item"
         SetWindowText GetDlgItem(hWin, IDC_EDT11), info
         SendMessage GetDlgItem(hWin, IDC_EDT11), EM_SETSEL, 0, - 1
			setfocus(GetDlgItem(hWin, IDC_EDT11))
			exit function
      Case WM_CLOSE
         EndDialog(hWin, 0)
			setfocus(hcode)
			exit function
      Case WM_COMMAND
         id = LoWord(wParam1)
         Event1 = HiWord(wParam1)
         Select Case id
            Case IDC_BTN11
               GetWindowText GetDlgItem(hWin, IDC_EDT11), nBuffer, 250
               searching2 = trim(nBuffer)
               EndDialog(hWin, 0)
					setfocus(hcode)
					exit function
            Case IDC_BTN21
               EndDialog(hWin, 0)
					setfocus(hcode)
					nook=1
					exit function
         End Select
      Case Else
         Return FALSE
   End Select
   Return TRUE
End Function

Function AskMain2(ByVal hwParent As hwnd) As Integer
		iscod=1
		nook=0
      DialogBoxParam(GetModuleHandle(NULL), Cast(ZString Ptr, dlg_Ask), hwParent, @dlgAsk_DlgProc2, 0)
      'messagebox 0,searching,"retour",mb_ok
      Return True
'   Else
'      messagebox 0, "Not in the code view tab ! " & tab_pos(hwParent), "Warning", mb_ok
'   End if
'   Return False
End Function