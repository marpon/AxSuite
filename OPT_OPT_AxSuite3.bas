'CSED_FB : New Generated Optimized / Concentrated Code File
REM 'CSED_FB : New Generated Optimized / Concentrated / Cleaned Code File

_:[CSED_COMPIL_RC]:   AxSuite3.rc   ' with ou without double quotes !

_:[CSED_COMPIL_NAME]:   AxSuite3_23.exe  ' with ou without double quotes !

'_:[CSED_OPTIMIZE]: 'Private procs , dead code Removal

#include once "windows.bi"
#include once "win/richedit.bi"
#Include Once "win/ocidl.bi"
#Include Once "win/shellapi.bi"
#include once "win/commctrl.bi"
#include once "win/commdlg.bi"

#IFNDEF INCFILE_PLUS
   #DEFINE INCFILE_PLUS

   Private Function MakeFileImport_(TabU1 as ubyte ptr, IncFil_len As Uinteger, NewFile As String) As integer
      Dim lFile AS short
      lFile = FreeFile                           'find next open file number
      Open NewFile For Binary As #lFile          'create the new file
      Put #lFile,, TabU1[0], IncFil_len          'write the buffer to the file
      Close #lFile                               'close the file
      Function = 1                               'worked so return a 1
   End Function

	Private Function mem_ClipboardSetText(TabU1 as ubyte ptr, IncFil_len As Uinteger) As Long
		dim pstr as zstring ptr
		pstr=TabU1
		Var hGlobalClip = GlobalAlloc(GMEM_MOVEABLE Or GMEM_SHARE, IncFil_len + 1)
		If OpenClipboard(0) Then
			EmptyClipboard()
			Var lpMem = GlobalLock(hGlobalClip)
			If lpMem Then
				CopyMemory(lpMem, pstr, IncFil_len)
				GlobalUnlock(lpMem)
				Function = cast(long, SetClipboardData(CF_TEXT, hGlobalClip))
			End If
			CloseClipboard()
		else
			Function = 0
		End If
	End Function

'¤¤¤'   #MACRO IncFilePlus(label, file, sectionName, attr)
'¤¤¤'      #if __FUNCTION__ <> "__FB_MAINPROC__"
'¤¤¤'         #IfnDef Var__##label##__Import__
'¤¤¤'            Dim label As UByte Ptr = Any
'¤¤¤'            Dim label##_len As UInteger = Any
'¤¤¤'         #endif
'¤¤¤'         #if __FB_DEBUG__
'¤¤¤'            asm jmp .LT_END_OF_FILE_##label##_DEBUG_JMP
'¤¤¤'         #else
'¤¤¤'            #if attr = ""
'¤¤¤'               asm .section sectionName
'¤¤¤'            #else
'¤¤¤'               asm .section sectionName , attr
'¤¤¤'            #endif
'¤¤¤'         #endif
'¤¤¤'
'¤¤¤'         asm .LT_START_OF_FILE_##label#:
'¤¤¤'         asm __##label##__start = .
'¤¤¤'         asm .incbin ##file
'¤¤¤'         asm __##label##__len = . - __##label##__start
'¤¤¤'         asm .LT_END_OF_FILE_##label:
'¤¤¤'         asm .LONG 0
'¤¤¤'         #if __FB_DEBUG__
'¤¤¤'            asm .LT_END_OF_FILE_##label##_DEBUG_JMP:
'¤¤¤'         #else
'¤¤¤'            asm .section .text
'¤¤¤'            asm .balign 16
'¤¤¤'         #endif
'¤¤¤'         asm .LT_SKIP_FILE_##label:
'¤¤¤'         asm mov dword ptr [label ] , offset .LT_START_OF_FILE_##label#
'¤¤¤'         asm mov dword ptr [label##_len ] , offset __##label##__len
'¤¤¤'
'¤¤¤'      #else
'¤¤¤'         #IfnDef Var__##label##__Import__
'¤¤¤'            Extern "c"
'¤¤¤'               Extern label As UByte Ptr
'¤¤¤'               Extern label##_len As UInteger
'¤¤¤'            End Extern
'¤¤¤'         #endif
'¤¤¤'         #if __FB_DEBUG__
'¤¤¤'            asm jmp .LT_END_OF_FILE_##label##_DEBUG_JMP
'¤¤¤'         #else
'¤¤¤'            #if attr = ""
'¤¤¤'               asm .section sectionName
'¤¤¤'            #else
'¤¤¤'               asm .section sectionName , attr
'¤¤¤'            #endif
'¤¤¤'         #endif
'¤¤¤'         asm .LT_START_OF_FILE_##label#:
'¤¤¤'         asm __##label##__start = .
'¤¤¤'         asm .incbin ##file
'¤¤¤'         asm __##label##__len = . - __##label##__start
'¤¤¤'         asm .LT_END_OF_FILE_##label:
'¤¤¤'         asm .LONG 0
'¤¤¤'         asm label:
'¤¤¤'         asm .int .LT_START_OF_FILE_##label#
'¤¤¤'         asm label##_len:
'¤¤¤'         asm .int __##label##__len
'¤¤¤'         #if __FB_DEBUG__
'¤¤¤'            asm .LT_END_OF_FILE_##label##_DEBUG_JMP:
'¤¤¤'         #else
'¤¤¤'            asm .section .text
'¤¤¤'            asm .balign 16
'¤¤¤'         #endif
'¤¤¤'         asm .LT_SKIP_FILE_##label:
'¤¤¤'      #endif
'¤¤¤'
'¤¤¤'   #endmacro

'¤¤¤'   #macro IncFileMem(label, file)
'¤¤¤'      IncFilePlus(label, file,.data, "")         'Use the .data (storage) section (SHARED)
'¤¤¤'   #endmacro

#endif

IncFileMem(inc1,"ATL.DLL")			'MakeFile(lpFile_inc1, "new path", is1_inc1)
IncFileMem(inc2,"ATL71.DLL")			'MakeFile(lpFile_inc2, "new path", is1_inc2)
IncFileMem(inc5,"atlCtl.dll")
IncFileMem(inc6,"atl71Ctl.dll")

IncFileMem(inc9,"Ax_Lite.bi")		'Mak

Private Function RegisterServer(hWnd As hwnd, filePath As String, Mode As integer=1)as integer
  Dim hLib As HMODULE
  Dim ProcAd As any ptr
  Dim sMod As String
  Dim res As integer

  if mode=0 THEN
	  sMod = "DllUnregisterServer"
  else
	  sMod = "DllRegisterServer"
  END IF

  hLib = LoadLibrary(filePath)
  ProcAd = GetProcAddress(hLib, sMod)

  res = CallWindowProc(ProcAd, hwnd, 0, 0, 0)

  FreeLibrary hLib
  Function = cast(integer,res)

End Function

Const szNULL = !"\0"
Const szOpen = "File type : Exe ; Dll ; Ocx  (*.exe ,*.dll ,ocx)" & szNULL & "*.exe;*.dll;*.ocx" & szNULL & "File type : Tlb or Olb    (*.tlb ,*.olb)" & szNULL & "*.tlb;*.olb" & szNULL
Const szSave = "Include File : bi (*.bi)" & szNULL
Const szDll = "File : dll (*.dll)" & szNULL

Const crlf = Chr$(13, 10)
Const dq = Chr$(34)
Const Tb = Chr$(9)
Const d_bc = Chr$(47,39)
Const f_bc = Chr$(39,47)

Type TV_HITTESTINFO
   pt                              As POINT
   flags                           As UINT
   hitem                           As HTREEITEM
End Type

Type LPTV_HITTESTINFO As TV_HITTESTINFO ptr

const LVM_SORTITEMSEX = LVM_FIRST + 81

TYPE LV_DISPINFO
   hdr                             AS NMHDR
   item                            AS LV_ITEM
END TYPE

Private Function StringToBSTR(cnv_string As String) As BSTR
   Dim sb As BSTR
   Dim As Integer n
   n = (MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, cnv_string, - 1, NULL, 0)) - 1
   sb = SysAllocStringLen(sb, n)
   MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, cnv_string, - 1, sb, n)
   Return sb
End Function

Private Function Variantv(ByRef v As variant) As Double
   Dim vvar As variant
   VariantChangeTypeEx(@vvar, @v, NULL, VARIANT_NOVALUEPROP, VT_R8)
   Return vvar.dblval
End Function

'¤¤¤'Private Function str_remove(ByRef source As String, byref match As String) As String
'¤¤¤'   dim as long s = 1, lmt
'¤¤¤'   Dim As String txt
'¤¤¤'   txt = source
'¤¤¤'   lmt = Len(match)
'¤¤¤'   Do
'¤¤¤'      s = InStr(s, txt, match)
'¤¤¤'      If s Then txt = Left(txt, s - 1) + Right(txt, Len(txt) - lmt - (s - 1))
'¤¤¤'   Loop While s
'¤¤¤'   Function = txt
'¤¤¤'end function

Private Function Str_removeany(ByRef source As String, ByRef match As String) As String
   dim as long s = 1
   Dim As String txt
   txt = source
   Do
      s = Instr(s, txt, Any match)
      If s Then txt = Left(txt, s - 1) + Right(txt, Len(txt) - 1 - (s - 1))
   Loop While s
   Function = txt
end Function

Private Function str_numparse(ByRef source as string, ByRef delimiter as string) as long
   Dim As Long s = 1, c, l
   l = Len(delimiter)
   Do
      s = instr(s, source, delimiter)
      If s then
         c += 1
         s += l
      end if
   Loop While s
   Function = c + 1
end function

Private Function str_parse( source As String,  delimiter As String, ByVal idx As integer) As String
   Dim As integer s = 1, c, l
   l = Len(delimiter)
   Do
      If c = idx - 1 then Return mid(source, s, instr(s, source, delimiter) - s)
      s = instr(s, source, delimiter)
      If s then
         c += 1
         s += l
      end if
   Loop While s
End Function

Private Function str_tally(ByRef source as string, byref delimiter as string) as long
   Dim As Long s = 1, c, l = Len(delimiter)
   Do
      s = instr(s, source, delimiter)
      If s then
         c += 1
         s += l
      end if
   Loop While s
   Function = c
End function

'¤¤¤'Private Function FF_Replace( ByRef sText        As String, _
'¤¤¤'                     ByRef sLookFor     As String, _
'¤¤¤'                     ByRef sReplaceWith As String _
'¤¤¤'                     ) As String
'¤¤¤'
'¤¤¤'   Dim sTemp As String  = sText
'¤¤¤'   Dim f     As Integer = 1
'¤¤¤'   If sLookFor = sReplaceWith Then
'¤¤¤'
'¤¤¤'      Function = sTemp
'¤¤¤'      Exit Function
'¤¤¤'   EndIf
'¤¤¤'   If InStr(sReplaceWith,sLookFor) Then
'¤¤¤'        Do Until f = 0
'¤¤¤'           f = InStr( sTemp, sLookFor )
'¤¤¤'           If f Then
'¤¤¤'              sTemp = Left(sTemp, f-1) & "#§§~¿~§§#" & Mid(sTemp, f + Len(sLookFor))
'¤¤¤'           End If
'¤¤¤'        Loop
'¤¤¤'        f = 1
'¤¤¤'        Do Until f = 0
'¤¤¤'          f = InStr( sTemp, "#§§~¿~§§#" )
'¤¤¤'          If f Then
'¤¤¤'             sTemp = Left(sTemp, f-1) & sReplaceWith & Mid(sTemp, f + 9)
'¤¤¤'          End If
'¤¤¤'        Loop
'¤¤¤'
'¤¤¤'        Function = sTemp
'¤¤¤'        Exit Function
'¤¤¤'   EndIf
'¤¤¤'   Do Until f = 0
'¤¤¤'      f = InStr( sTemp, sLookFor )
'¤¤¤'      If f Then
'¤¤¤'         sTemp = Left(sTemp, f-1) & sReplaceWith & Mid(sTemp, f + Len(sLookFor))
'¤¤¤'      End If
'¤¤¤'   Loop
'¤¤¤'
'¤¤¤'   Function = sTemp
'¤¤¤'
'¤¤¤'End Function

Private Function str_replace(ByRef match As String, Byref sreplace As String, Byref source As String) As String
   Dim As Long s = 1, c = len(sreplace), l = len(match)
   Dim As String text
   text = source
   Do
      s = InStr(s, text, match)
      If s then
         text = Mid$(text, 1, s - 1) + sreplace + Mid$(text, s + l, - 1)
         s += c
      end if
   Loop While s
   Function = text
End Function

Private Function str_ReplaceAll(ByRef match As String, ByRef sreplace As String, ByRef source As String) As String
   Dim text As String
   text = source
   While InStr(text, match)
      text = str_replace(match, sreplace, text)
   Wend
   Function = text
End Function

Private Sub LVSetHeader(BYVAL lvg AS uinteger, BYVAL strItem AS STRING)
   dim szItem AS zstring * MAX_PATH              ' working variable
   dim tlvc AS LVCOLUMN                          ' specifies or receives the attributes of a listview column
   dim tlvi AS LVITEM                            ' specifies or receives the attributes of a listview item
   dim iItem AS LONG
   dim nItem AS LONG

   nItem = str_numparse(stritem, "|")

   FOR iItem = 1 TO nItem
      szItem = str_parse(stritem, "|", iItem)
      tlvc.mask = LVCF_FMT OR LVCF_WIDTH OR LVCF_TEXT OR LVCF_SUBITEM
      tlvc.fmt = LVCFMT_LEFT
      tlvc.cx = 96
      tlvc.pszText = VARPTR(szItem)
      tlvc.iSubItem = iitem - 1
      SendMessage cast(hwnd, lvg), LVM_SETCOLUMN, iitem - 1, cast(LPARAM, VARPTR(tlvc))
   NEXT
END Sub

Private Sub LVSetMaxCols(BYVAL rr AS uinteger, BYVAL cols AS LONG)
   dim tlvc AS LVCOLUMN
   dim col AS LONG

   FOR col = 0 TO cols - 1 :SendMessage cast(hwnd, rr), LVM_INSERTCOLUMN, 0, cast(LPARAM, VARPTR(tlvc)) :NEXT
END SUB

Private Function LVGetValue(BYVAL rr AS uinteger, BYVAL item AS LONG, BYVAL column AS LONG) AS STRING
   dim ms_lvi AS LV_ITEM
   dim psztext as zstring*max_path
   ms_lvi.iSubItem = Column
   ms_lvi.cchTextMax = max_path
   ms_lvi.pszText = @pszText
   SendMessage cast(hwnd, rr), LVM_GETITEMTEXT, item, cast(LPARAM, @ms_lvi)
   FUNCTION = psztext
END FUNCTION

Private Function TreeViewInsertItem(BYVAL hTreeView AS long, BYVAL hParent AS long, sItem AS STRING) AS LONG
   Dim tTVInsert AS TV_INSERTSTRUCT
   Dim tTVItem AS TV_ITEM

   If hParent THEN
      tTVItem.mask = TVIF_CHILDREN OR TVIF_HANDLE
      tTVItem.hItem = cast(HTREEITEM, hParent)
      tTVItem.cchildren = 1
      TreeView_SetItem(cast(hwnd, hTreeView), @tTVItem)
   END IF

   tTVInsert.hParent = cast(HTREEITEM, hParent)
   tTVInsert.Item.mask = TVIF_TEXT Or TVIF_STATE
   tTVInsert.Item.state = TVIS_EXPANDED
   tTVInsert.Item.pszText = STRPTR(sItem)
   tTVInsert.Item.cchTextMax = LEN(sItem)
   FUNCTION = cast(long, TreeView_InsertItem(cast(hwnd, hTreeView), @tTVInsert))
END Function

Private Function TreeView_GetItemText(ByVal hTreeView As Long, ByVal hItem As Long) As String
   Dim ItemText As zstring*32765
   Dim Item As TV_ITEM
   Item.hItem = cast(HTREEITEM, hItem)
   Item.Mask = TVIF_TEXT
   Item.cchTextMax = 32765
   Item.pszText = @ItemText
   SendMessage cast(hwnd, hTreeView), TVM_GETITEM, 0, cast(LPARAM, @Item)
   Function = ItemText
End Function

Private Function TreeView_ChangeChildState(ByVal hTreeView As Long, ByVal hItem As Long, ByVal iState As Long) As Long
   Dim hTmpItem As HTREEITEM
   Dim TVITEM As TV_ITEM
   hTmpItem = TreeView_GetChild(cast(hwnd, hTreeView), hItem)
   While hTmpItem
      TVITEM.mask = TVIF_STATE Or TVIF_HANDLE
      TVITEM.hItem = hTmpItem
      TVITEM.stateMask = TVIS_STATEIMAGEMASK
      TVITEM.state = (iState + 1) Shl 12
      sendmessage cast(hwnd, hTreeView), TVM_SetItem, 0, cast(LPARAM, @TVITEM)
      TreeView_ChangeChildState hTreeView, cast(long, hTmpItem), iState
      hTmpItem = treeview_getnextitem(cast(hwnd, htreeview), hTmpItem, TVGN_NEXT)
   wend
   function = 1
End Function

Private Function TreeView_ChangeParentState(ByVal hTreeView As Long, ByVal hItem As Long, ByVal iState As Long) As Long
   Dim hTmpItem As HTREEITEM
   Dim TVITEM As TV_ITEM
   hTmpItem = TreeView_GetParent(cast(hwnd, hTreeView), hItem)
   If hTmpItem then
      TVITEM.mask = TVIF_STATE Or TVIF_HANDLE
      TVITEM.hItem = hTmpItem
      TVITEM.stateMask = TVIS_STATEIMAGEMASK
      TVITEM.state = (iState + 1) Shl 12
      sendmessage cast(hwnd, hTreeView), TVM_SetItem, 0, cast(LPARAM, @TVITEM)
      TreeView_ChangeParentState hTreeView, cast(long, hTmpItem), 0
   End If
   function = 1
End Function

Private Function TreeView_SetCheckState(ByVal hTreeView As Long, ByVal hItem As Long, ByVal iState As Long) As Long
   Dim TVITEM As TV_ITEM
   TVITEM.mask = TVIF_STATE Or TVIF_HANDLE
   TVITEM.hItem = cast(HTREEITEM, hItem)
   TVITEM.stateMask = TVIS_STATEIMAGEMASK
   TVITEM.state = (iState + 1) Shl 12
   sendmessage cast(hwnd, hTreeView), TVM_SetItem, 0, cast(LPARAM, @TVITEM)
   function = 1
End Function

Private Function TreeView_GetCheckState(ByVal hWnd1 As Long, ByVal hItem As Long) As Long
   Dim tvItem As TV_ITEM
   tvItem.mask = TVIF_HANDLE Or TVIF_STATE
   tvItem.hItem = cast(HTREEITEM, hItem)
   tvItem.stateMask = TVIS_STATEIMAGEMASK
   sendmessage cast(hwnd, hwnd1), TVM_GetItem, 0, cast(LPARAM, @tvItem)
   tvItem.state Shr= 12
   tvItem.state -= 1
   Function = tvItem.state Xor 1
End Function

Const TV_UNCHECKED as integer = 1  ' for treeview checked
Const TV_CHECKED as integer = 2

Private Function FF_FileName (Src As String) As String
    Dim x As Integer
    x = InStrRev( Src, Any ":/\")
    If x Then
        Function = Mid(Src, x + 1)
    Else
        Function = Src
    End If
End Function

'¤¤¤'#define dlgMain 1000
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

#define IDC_STC31 1430
#define IDC_EDT31 1432

#define tlb_reged 10002
#define tlb_libfile 10003
#define tlb_exit 10004
#define code_const 10006
#define code_Module 10007
#define code_vtable 10008
#define code_Invoke 10009
#define code_event 10010
#define hlp_About 10012
#Define hlp_axsuite 10013
#Define code_init 10014
#Define code_full 10015
#define code_codegen 10016
#define code_save 10017
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

#define Copy_GUID 10071
#define Copy_NAME 10072
#define Copy_PATH 10073
#define code_search 1480
#define code_search1 1481
#define code_search2 1482
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

Private Sub GetPrefix()
	sendmessage getdlgitem(hmain, idc_edt1), wm_gettext, 10, cast(lparam, @prefix)
END sub

Private Function TreeView_GetCheckState1( ByVal hWnd As Long, ByVal hItem As long ) As uinteger
	Dim tvItem As TV_ITEM
	dim as uinteger st1
	tvItem.mask         = TVIS_SELECTED'TVIF_HANDLE Or TVIF_STATE
	tvItem.hItem        = cast(HTREEITEM, hItem)
	tvItem.stateMask    = TVIS_SELECTED'TVIS_STATEIMAGEMASK
	If TreeView_GetItem (cast(hwnd,hWnd), @tvItem ) = 0 Then
		Return 0
	Else
		st1=tvItem.state
		if st1 > 4193 THEN
			Return 1
		else
			Return 0
      END IF
	End If
End Function

Private Function TreeView_GetItemText1(ByVal hTreeView As Long, ByVal hItem As Long) As String
   Dim ItemText As zstring*32765
	Dim Item As TV_ITEM
	dim as integer gi1= TreeView_GetCheckState1(hTreeView,hItem)
		Item.hItem = cast(HTREEITEM, hItem)
		Item.Mask = TVIF_TEXT
		Item.cchTextMax = 32765
		Item.pszText = @ItemText
		SendMessage cast(hwnd,hTreeView), TVM_GETITEM, 0, cast(LPARAM,@Item)
	if gi1=0 and xfull = 0 THEN
		function= ItemText & "          'NOT'" ' 10 spaces
	else
		Function = ItemText
	end if
End Function

Private Sub putEvtList Cdecl Alias "putEvtList"(evt As String)
   EvtList &= evt & "|"
End Sub

Private Sub clrEvtList Cdecl Alias "clrEvtList"()
   EvtList = ""
End Sub

'¤¤¤'Private Function inEvtList(evt As String) As Long
'¤¤¤'   For i As Short = 1 To str_numparse(evtlist, "|")
'¤¤¤'      If evt = str_parse(evtlist, "|", i) Then Return TRUE
'¤¤¤'   Next
'¤¤¤'End Function

'¤¤¤'Private Function EnumvTable cdecl Alias "EnumvTable"(hTree As Long, hParent() As long) As zString Ptr
'¤¤¤'   Dim Parents() As Long
'¤¤¤'   Dim hNextItem As Long
'¤¤¤'   Dim HasChilds As Long
'¤¤¤'   Dim hItem1 as Long,virg as integer,xvir as integer,in1 as integer,bevent as integer
'¤¤¤'   Dim token As String, s0 As String, tldesc As String, cor1 as string, cor2 as string, cor3 as string,sp1 as string
'¤¤¤'   Dim xface As String, memid As Integer, member As String, offs As Integer
'¤¤¤'	redim Parents(0)
'¤¤¤'   Dim s as String,us as string,xface0 as string,svt as string,s1 as string
'¤¤¤'	dim as string str1, strclasses, interf, siid
'¤¤¤'	GetPrefix()
'¤¤¤'	str1 = trim(prefix, any "-_ ")
'¤¤¤'   hItem1 = hparent(tkind_interface)
'¤¤¤'   tldesc = TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem1, TVGN_PARENT)))
'¤¤¤'	haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem1))
'¤¤¤'   If haschilds Then
'¤¤¤'      svt &= "'" & String(80, "=") & crlf
'¤¤¤'      svt &= "'vTable - " & tldesc & crlf
'¤¤¤'      svt &= "'" & String(80, "=") & crlf
'¤¤¤'      hnextitem = haschilds
'¤¤¤'      While hnextitem
'¤¤¤'         token = TreeView_GetItemText(hTree, hNextItem)
'¤¤¤'         haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hnextitem))
'¤¤¤'         If haschilds <> 0 And UBound(parents) = 0 Then
'¤¤¤'				interf = str_parse(token, "?", 2)
'¤¤¤'
'¤¤¤'				xface = str_parse(token, "?", 1)
'¤¤¤'				xface0=xface
'¤¤¤'				xface= str1 & xface
'¤¤¤'            bevent = inevtlist(xface0)
'¤¤¤'            If bevent=0 THEN
'¤¤¤'					if instr(ucase(xface0),"EVENT") THEN bevent= 2
'¤¤¤'				end if
'¤¤¤'				s1 = "'" & String(80, "=") & crlf
'¤¤¤'				s1 &= "'Interface " &  xface & "     ?" & str_parse(token, "?", 3) & crlf
'¤¤¤'				s1 &= "'Const IID_" &  xface & "=" & dq & str_parse(token, "?", 2) & dq & crlf
'¤¤¤'				s1 &= "'" & String(80, "=") & crlf
'¤¤¤'				s1 &= "Type " &  xface & "vTbl" & crlf
'¤¤¤'
'¤¤¤'         ElseIf (UBound(parents) <> 0) Then
'¤¤¤'            memid = Val(str_parse(token, "?", 1))
'¤¤¤'            member = str_parse(token, "?", 3) & "     '" & str_parse(token, "?", 4)
'¤¤¤'            If offs = 0 Then
'¤¤¤'               If memid > 8 Then
'¤¤¤'                  s1 &= tb & "QueryInterface As Function (Byval pThis As Any ptr,Byref riid As GUID,Byref ppvObj As Dword) As hResult" & crlf
'¤¤¤'                  s1 &= tb & "AddRef As Function (Byval pThis As Any ptr) As Ulong" & crlf
'¤¤¤'                  s1 &= tb & "Release As Function (Byval pThis As Any ptr) As Ulong" & crlf
'¤¤¤'                  offs = 8
'¤¤¤'               End If
'¤¤¤'               If memid > 24 Then
'¤¤¤'                  s1 &= tb & "GetTypeInfoCount As Function (Byval pThis As Any ptr,Byref pctinfo As Uinteger) As hResult" & crlf
'¤¤¤'                  s1 &= tb & "GetTypeInfo As Function (Byval pThis As Any ptr,Byval itinfo As Uinteger,Byval lcid As Uinteger,Byref pptinfo As Any Ptr) As hResult" & crlf
'¤¤¤'                  s1 &= tb & "GetIDsOfNames As Function (Byval pThis As Any ptr,Byval riid As GUID,Byval rgszNames As Byte,Byval cNames As Uinteger,Byval lcid As Uinteger,Byref rgdispid As Integer) As hResult" & crlf
'¤¤¤'                  s1 &= tb & "Invoke As Function (Byval pThis As Any ptr,Byval dispidMember As Integer,Byval riid As GUID,Byval lcid As Uinteger,Byval wFlags As Ushort,Byval pdispparams As DISPPARAMS,Byref pvarResult As Variant,Byref pexcepinfo As EXCEPINFO,Byref puArgErr As Uinteger) As hResult" & crlf
'¤¤¤'                  offs = 24
'¤¤¤'               End If
'¤¤¤'            End If
'¤¤¤'            s0 = ""
'¤¤¤'            offs = memid - offs - 4
'¤¤¤'            If offs > 0 Then s0 = "(" &(offs - 1) & ") As Byte" & crlf
'¤¤¤'            offs = memid
'¤¤¤'            If Len(s0) Then s1 &= tb & "Offset" & offs & s0
'¤¤¤'
'¤¤¤'            If InStr(member, "()") and InStr(member, " As Sub()")>0 Then
'¤¤¤'					s1 &= tb & str_parse(member, "(", 1) & "(ByVal pThis As Any Ptr)" & crlf
'¤¤¤'				elseIf InStr(member, "()") and InStr(member, " As Function()")>0 Then
'¤¤¤'               s1 &= tb & str_parse(member, "(", 1) & "(ByVal pThis As Any Ptr" & Mid(member, Len(str_parse(member, "(", 1)) + 2, - 1) & crlf
'¤¤¤'            ElseIf InStr(member, "()")= 0 then
'¤¤¤'					cor1= str_parse(member, "(", 1) & "(ByVal pThis As Any Ptr," & Mid(member, Len(str_parse(member, "(", 1)) + 2, - 1) & crlf
'¤¤¤'
'¤¤¤'					sp1= ") As " & str_parse(cor1,") As ",str_numparse(cor1, ") As " ))
'¤¤¤'					cor1= str_parse(cor1,") As ",1)
'¤¤¤'					virg = str_numparse(cor1, "," )
'¤¤¤'					for xvir = 2 to virg
'¤¤¤'						cor2 = str_parse(cor1, "," , xvir)
'¤¤¤'						if instr(cor2," As ") THEN
'¤¤¤'							cor3 = "ByVal " & cor2
'¤¤¤'							cor1 = FF_Replace( cor1 ,cor2,cor3)
'¤¤¤'                  END IF
'¤¤¤'               NEXT
'¤¤¤'
'¤¤¤'					cor1 = FF_Replace( cor1 ,",ByVal  As",",ByVal VarAny As")
'¤¤¤'				   s1 &= tb & cor1 & sp1
'¤¤¤'            End If
'¤¤¤'
'¤¤¤'         End If
'¤¤¤'
'¤¤¤'         If haschilds Then
'¤¤¤'            ReDim Preserve parents(UBound(parents) + 1)
'¤¤¤'            parents(UBound(parents)) = hnextitem
'¤¤¤'            hnextitem = haschilds
'¤¤¤'         Else
'¤¤¤'            hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
'¤¤¤'            If hnextitem = 0 And UBound(parents) > 0 Then
'¤¤¤'               hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), parents(UBound(parents)), TVGN_NEXT))
'¤¤¤'               ReDim Preserve parents(UBound(parents) - 1)
'¤¤¤'               s1 &= "End Type      '" &  xface & "vTbl" & crlf & crlf
'¤¤¤'               s1 &= "Type " &  xface & "_"& crlf
'¤¤¤'               s1 &= tb & "lpvtbl As " &  xface & "vTbl Ptr" & crlf
'¤¤¤'               s1 &= "End Type" & crlf & crlf
'¤¤¤'					if bevent=0 then us &= tb & "Type  " & xface & "    as  " & xface & "_     ' Interface :  " &  xface  & crlf
'¤¤¤'					s1 &= tb & "'     best way to do ...            Dim Shared As " & xface & " ptr  " & str1 & "pVTI  " & crlf
'¤¤¤'					s1 &= tb & "'     ex:  "  & str1 & "pVTI->lpvtbl->SetColor( " & str1 & "pVTI ,...)" & crlf
'¤¤¤'					s1 &= tb & "'     or    Ax_Vt ( " & str1 & "pVTI , SetColor ,...)" & crlf
'¤¤¤'					s1 &= tb & "'     or    Ax_Vt0 ( " & str1 & "pVTI , About )" & crlf & crlf
'¤¤¤'					in1=0
'¤¤¤'					for xvir = 1 to icoclass
'¤¤¤'						if coclass(xvir,5) = interf THEN
'¤¤¤'							in1=xvir
'¤¤¤'							exit for
'¤¤¤'                  END IF
'¤¤¤'               NEXT
'¤¤¤'					if in1 THEN
'¤¤¤'						s1 &= "Function Create_" & xface & "() As any ptr" & crlf
'¤¤¤'						s1 &= tb & "Function = AxCreate_object( """ & coclass(in1,3) & """ , """ & coclass(in1,5)& """ )" & crlf
'¤¤¤'						s1 &= "End Function" & crlf & crlf
'¤¤¤'						s1 &= " ' use  normaly    like that  :      " & str1 & "pVTI = " & str1 & "Obj_Ptr" & crlf
'¤¤¤'						s1 &= " ' but if problem try         :      " & str1 & "pVTI = Create_" & xface & "()"  & crlf & crlf & crlf
'¤¤¤'					else
'¤¤¤'						s1 &= " ' use   " & str1 & "pVTI = " & str1 & "Obj_Ptr" & crlf & crlf & crlf
'¤¤¤'               END IF
'¤¤¤'					if bevent=0 then  s &= s1
'¤¤¤'               offs = 0
'¤¤¤'            End If
'¤¤¤'         End If
'¤¤¤'      Wend
'¤¤¤'
'¤¤¤'   End If
'¤¤¤'	if s <> "" THEN
'¤¤¤'		s = "'" & String(80, "*") & crlf & "'  Interface vTable Types :" & crlf & crlf & us & crlf & crlf & "'" & String(80, "*") & crlf & crlf & s
'¤¤¤'		s = svt & crlf & crlf & s
'¤¤¤'	END IF
'¤¤¤'
'¤¤¤'	if instr(s , " _Collection Ptr" ) THEN
'¤¤¤'		us= "'================================================================================" & crlf & crlf
'¤¤¤'		us &= "'Interface _Collection  , Standard Vb6/Vba collection vTable needed here" & crlf & crlf
'¤¤¤'		us &= "Type _CollectionvTbl" & crlf
'¤¤¤'		us &= tb & "QueryInterface As Function (Byval pThis As Any ptr,Byref riid As GUID,Byref ppvObj As Dword) As hResult" & crlf
'¤¤¤'		us &= tb & "AddRef As Function (Byval pThis As Any ptr) As Ulong" & crlf
'¤¤¤'		us &= tb & "Release As Function (Byval pThis As Any ptr) As Ulong" & crlf
'¤¤¤'		us &= tb & "GetTypeInfoCount As Function (Byval pThis As Any ptr,Byref pctinfo As Uinteger) As hResult" & crlf
'¤¤¤'		us &= tb & "GetTypeInfo As Function (Byval pThis As Any ptr,Byval itinfo As Uinteger,Byval lcid As Uinteger,Byref pptinfo As Any Ptr) As hResult" & crlf
'¤¤¤'		us &= tb & "GetIDsOfNames As Function (Byval pThis As Any ptr,Byval riid As GUID,Byval rgszNames As Byte,Byval cNames As Uinteger,Byval lcid As Uinteger,Byref rgdispid As Integer) As hResult" & crlf
'¤¤¤'		us &= tb & "Invoke As Function (Byval pThis As Any ptr,Byval dispidMember As Integer,Byval riid As GUID,Byval lcid As Uinteger,Byval wFlags As Ushort,Byval pdispparams As DISPPARAMS,Byref pvarResult As Variant,Byref pexcepinfo As EXCEPINFO,Byref puArgErr As Uinteger) As hResult" & crlf
'¤¤¤'		us &= tb & "Item As Function(ByVal pThis As Any Ptr,ByVal Index As VARIANT Ptr,ByVal pvarRet As VARIANT Ptr) As HRESULT      '" & crlf
'¤¤¤'		us &= tb & "Add As Function(ByVal pThis As Any Ptr,ByVal Item As VARIANT Ptr,ByVal Key As VARIANT Ptr=0,ByVal Before As VARIANT Ptr=0,ByVal After As VARIANT Ptr=0) As HRESULT     '" & crlf
'¤¤¤'		us &= tb & "Count As Function(ByVal pThis As Any Ptr,ByVal pi4 As Integer Ptr) As HRESULT     '" & crlf
'¤¤¤'		us &= tb & "Remove As Function(ByVal pThis As Any Ptr,ByVal Index As VARIANT Ptr) As HRESULT     '" & crlf
'¤¤¤'		us &= tb & "_NewEnum As Function(ByVal pThis As Any Ptr,ByVal ppunk As LPUNKNOWN Ptr) As HRESULT     '" & crlf
'¤¤¤'		us &= "End Type      '_CollectionvTbl" & crlf & crlf
'¤¤¤'		us &= "Type _Collection" & crlf
'¤¤¤'		us &= tb & "lpvtbl As _CollectionvTbl Ptr" & crlf
'¤¤¤'		us &= "End Type" & crlf
'¤¤¤'		us &= "'================================================================================" & crlf
'¤¤¤'		s = us & crlf & s
'¤¤¤'   END IF
'¤¤¤'
'¤¤¤'	Return cast(zstring ptr, sysallocstringbytelen(s, len(s)))
'¤¤¤'End Function

'¤¤¤'Private Function EnumInvoke cdecl Alias "EnumInvoke"(hTree As Long, hParent() As long) As zString Ptr
'¤¤¤'   Dim Parents() As Long
'¤¤¤'   Dim hNextItem As Long
'¤¤¤'   Dim HasChilds As Long
'¤¤¤'
'¤¤¤'   Dim token As String, tldesc As String
'¤¤¤'   Dim hItem as Long
'¤¤¤'   Dim xface As String, memid As String, member As String,s1 as string,xface0 as string,sinv as string
'¤¤¤'   Dim As Integer cpar,  bevent
'¤¤¤'   Dim s as String
'¤¤¤'	dim as string str1
'¤¤¤'	redim Parents(0)
'¤¤¤'	GetPrefix()
'¤¤¤'	str1 = trim(prefix, any "-_ ")
'¤¤¤'   hItem = hparent(tkind_dispatch)
'¤¤¤'   tldesc =  TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem, TVGN_PARENT)))
'¤¤¤'   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
'¤¤¤'   If haschilds Then
'¤¤¤'      sinv = "'" & String(80, "=") & crlf
'¤¤¤'      sinv &= "'Invoke - " & tldesc & crlf
'¤¤¤'      sinv &= "'" & String(80, "=") & crlf
'¤¤¤'      hnextitem = haschilds
'¤¤¤'      While hnextitem
'¤¤¤'         token = TreeView_GetItemText(hTree, hNextItem)
'¤¤¤'         haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hnextitem))
'¤¤¤'         If (haschilds <> 0) And (UBound(parents) = 0) Then
'¤¤¤'				xface = str_parse(token, "?", 1)
'¤¤¤'				xface0=xface
'¤¤¤'				xface= str1 & xface
'¤¤¤'            bevent = inevtlist(xface0)
'¤¤¤'				If bevent=0 THEN
'¤¤¤'					if instr(ucase(xface0),"EVENT") THEN bevent= 2
'¤¤¤'				end if
'¤¤¤'            s1 = "'" & String(80, "=") & crlf
'¤¤¤'            s1 &= "'Dispatch " & xface & "     ?" & str_parse(token, "?", 3) & crlf
'¤¤¤'            s1 &= "'Const IID_" & xface & "=" & dq & str_parse(token, "?", 2) & dq & crlf
'¤¤¤'            s1 &= "'" & String(80, "=") & crlf
'¤¤¤'            s1 &= "Type " & xface & crlf
'¤¤¤'         ElseIf (UBound(parents) <> 0) then
'¤¤¤'            memid = str_parse(token, "?", 1)
'¤¤¤'            member = str_parse(token, "?", 3)
'¤¤¤'				member = FF_Replace( member , "Const ", "")
'¤¤¤'				member=rtrim(member,any"'= " & dq)
'¤¤¤'				if instr(member,")")=0 and instr(member,"(")THEN member=member & ")"
'¤¤¤'
'¤¤¤'            cpar = str_numparse(member, ",")
'¤¤¤'            If Len(str_parse(str_parse(member, "(", 2), ")", 1)) = 0 Then cpar = 0
'¤¤¤'				s1 &=tb & str_parse(member," As ",1) & " As tMember = ("
'¤¤¤'				s1 &= memid & ",2," & cpar & "," & str_parse(token,"?",2) & ")    '" & str_remove(member,str_parse(member," As ",1))  & crlf
'¤¤¤'         End If
'¤¤¤'
'¤¤¤'         If haschilds Then
'¤¤¤'            ReDim Preserve parents(UBound(parents) + 1)
'¤¤¤'            parents(UBound(parents)) = hnextitem
'¤¤¤'            hnextitem = haschilds
'¤¤¤'         Else
'¤¤¤'            hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
'¤¤¤'            If hnextitem = 0 And UBound(parents) > 0 Then
'¤¤¤'               hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), parents(UBound(parents)), TVGN_NEXT))
'¤¤¤'               ReDim Preserve parents(UBound(parents) - 1)
'¤¤¤'					s1 &= tb & "pMark As Integer = -1" & crlf
'¤¤¤'					s1 &= tb & "pThis As Integer" & crlf
'¤¤¤'               s1 &= "End Type        ' " & xface & crlf & crlf
'¤¤¤'					s1 &= "	'Use like that to use these dispach/invoke functions" & crlf
'¤¤¤'					s1 &= "	'  Dim Shared As " & xface & " " & str1 & "Obj_Disp" & crlf
'¤¤¤'					s1 &= "	'  SetObj ( @" & str1 & "Obj_Disp , " & str1 & "Obj_Ptr ) ' connect to object"  & crlf
'¤¤¤'					s1 &= "	'  ex : AxCall " & str1 & "Obj_Disp.putMonth,vptr(05)"  & crlf & crlf & crlf
'¤¤¤'					if bevent =0 THEN s &=s1
'¤¤¤'				End If
'¤¤¤'
'¤¤¤'         End If
'¤¤¤'
'¤¤¤'      Wend
'¤¤¤'		if s <> "" THEN s = sinv & crlf & crlf & s
'¤¤¤'   End If
'¤¤¤'   Return cast(zstring ptr, sysallocstringbytelen(s, len(s)))
'¤¤¤'End Function

'¤¤¤'Private Function EnumEvent2 (hTree As Long, hParent() As long, hItem as Long )as String
'¤¤¤'	Dim Parents() As Long
'¤¤¤'   Dim hNextItem As Long
'¤¤¤'   Dim HasChilds As Long
'¤¤¤'
'¤¤¤'   Dim token As String, tldesc As String
'¤¤¤'
'¤¤¤'   Dim As String xface, member, help, temp ,xface0,us,memberF,memberV, temp2
'¤¤¤'   Dim As Integer  bevent,numF,x1,itemp
'¤¤¤'
'¤¤¤'   Dim s as String
'¤¤¤'	dim as string str1,st_evt(),st_M(),sieven  ',st_F()
'¤¤¤'	redim Parents(0)
'¤¤¤'	redim st_evt(0)
'¤¤¤'	redim st_M(0)
'¤¤¤'	numF = 0
'¤¤¤'
'¤¤¤'	GetPrefix()
'¤¤¤'	str1 = trim(prefix, any "-_ ")
'¤¤¤'
'¤¤¤'   tldesc = str1 & TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem, TVGN_PARENT)))
'¤¤¤'   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
'¤¤¤'   If haschilds Then
'¤¤¤'
'¤¤¤'      us = "'" & String(80, "=") & crlf
'¤¤¤'      us &= "'Event - " & tldesc & crlf
'¤¤¤'      us &= "'" & String(80, "=") & crlf
'¤¤¤'      hnextitem = haschilds
'¤¤¤'      While hnextitem
'¤¤¤'
'¤¤¤'         token = TreeView_GetItemText(hTree, hNextItem)
'¤¤¤'         haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hnextitem))
'¤¤¤'         If (haschilds <> 0) And (UBound(parents) = 0) Then
'¤¤¤'				xface = str_parse(token, "?", 1)
'¤¤¤'				xface0=xface
'¤¤¤'				xface= str1 & xface
'¤¤¤'            bevent = inevtlist(xface0)
'¤¤¤'            If bevent Then
'¤¤¤'					sieven = str_parse(token, "?", 2)
'¤¤¤'					if sevent = "" or sieven = sevent THEN
'¤¤¤'						s &= "'" & String(80, "=") & crlf
'¤¤¤'						s &= "'Event vTable - " & xface & "     ?" & str_parse(token, "?", 3) & crlf
'¤¤¤'						s &= "Const " & xface & "_Ev_IID_CP = " & dq & str_parse(token, "?", 2) & dq & crlf
'¤¤¤'						s &= "'" & String(80, "=") & crlf & crlf & crlf
'¤¤¤'               END IF
'¤¤¤'            End If
'¤¤¤'         ElseIf (UBound(parents) <> 0) then
'¤¤¤'            numF +=1
'¤¤¤'            memberF = str_parse(token, "?", 3)
'¤¤¤'				memberF = FF_Replace(memberF,"   "," ")
'¤¤¤'				memberF = FF_Replace(memberF,"  "," ")
'¤¤¤'				memberF = FF_Replace(memberF,"  "," ")
'¤¤¤'				memberF = FF_Replace(memberF,"  "," ")
'¤¤¤'				memberF = trim (memberF)
'¤¤¤'				temp2 = memberF
'¤¤¤'				memberF = lcase(memberF)
'¤¤¤'
'¤¤¤'            help = str_parse(token, "?", 4)
'¤¤¤'
'¤¤¤'            redim preserve st_evt(1 to numF )
'¤¤¤'				redim preserve st_M(1 to numF)
'¤¤¤'				itemp = instr(memberF, " as ")
'¤¤¤'
'¤¤¤'				member = left(temp2,itemp -1)
'¤¤¤'
'¤¤¤'				st_M(numF)= xface & "_" & member
'¤¤¤'				If InStr(memberF, " as sub()")>0 or InStr(memberF, " as sub ()")> 0 _
'¤¤¤'							or	InStr(memberF, " as sub ( )")>0 or	InStr(memberF, " as sub( )")>0 Then
'¤¤¤'					memberF= member & "     As Sub ( ByVal pThis As Any Ptr )"
'¤¤¤'				elseIf InStr(memberF, " as function()")>0 or InStr(memberF, " as function ()")> 0 _
'¤¤¤'							or	InStr(memberF, " as function ( )")>0 or	InStr(memberF, " as function( )")>0 Then
'¤¤¤'               memberF= member & "     As Function ( ByVal pThis As Any Ptr )" & str_parse(temp2, ")", 2)
'¤¤¤'				elseIf InStr(memberF, " as function(")>0 or InStr(memberF, " as function (")> 0 then
'¤¤¤'					memberF= member & "     As Function ( ByVal pThis As Any Ptr , " & str_parse(temp2, "(", 2)
'¤¤¤'				elseIf InStr(memberF, " as sub(")>0 or InStr(memberF, " as sub (")> 0 then
'¤¤¤'					memberF= member & "     As Sub ( ByVal pThis As Any Ptr , " & str_parse(temp2, "(", 2)
'¤¤¤'				end if
'¤¤¤'
'¤¤¤'				st_evt(numF )= memberF
'¤¤¤'
'¤¤¤'				temp = mid(memberF, instr(memberF," As ") +3)
'¤¤¤'				temp = trim(temp)
'¤¤¤'
'¤¤¤'				if left(temp,3)="Sub" THEN
'¤¤¤'					temp = "Sub " &  st_M(numF) & "( " & str_parse(temp, "(", 2)
'¤¤¤'
'¤¤¤'				else
'¤¤¤'					temp = "Function " &  st_M(numF) & "( " & str_parse(temp, "(", 2)
'¤¤¤'            END IF
'¤¤¤'
'¤¤¤'            If bevent and (sevent = "" or sieven = sevent) Then
'¤¤¤'               s &= "'" & String(80, "=") & crlf
'¤¤¤'               s &= "'Event # " & str(numF) & "   " & member & "   ? " & help & crlf
'¤¤¤'               s &= "'" & String(80, "=") & crlf
'¤¤¤'					memberV = temp
'¤¤¤'               s &= memberV & crlf & crlf
'¤¤¤'               s &= tb & "'*** Put your code here ***" & crlf & crlf & crlf
'¤¤¤'					if left(ucase(memberV),3)="SUB" THEN
'¤¤¤'						 s &= "End Sub" & crlf & crlf
'¤¤¤'					else
'¤¤¤'						s &= tb & "Function = S_OK   ' change if needed" & crlf
'¤¤¤'						s &= "End Function" & crlf & crlf
'¤¤¤'               END IF
'¤¤¤'            End If
'¤¤¤'         End If
'¤¤¤'         If haschilds Then
'¤¤¤'            ReDim Preserve parents(UBound(parents) + 1)
'¤¤¤'            parents(UBound(parents)) = hnextitem
'¤¤¤'            hnextitem = haschilds
'¤¤¤'         Else
'¤¤¤'            hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
'¤¤¤'            If hnextitem = 0 And UBound(parents) > 0 Then
'¤¤¤'               hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), parents(UBound(parents)), TVGN_NEXT))
'¤¤¤'               ReDim Preserve parents(UBound(parents) - 1)
'¤¤¤'
'¤¤¤'               If bevent and (sevent = "" or sieven = sevent) and numF >0 Then
'¤¤¤'						s &= "Type " & xface & "Vtbl_Ev" & crlf
'¤¤¤'						s &= tb & "QueryInterface   As Function ( ByVal pThis As Any Ptr, ByRef riid As GUID, ByRef ppvObj As Dword ) As hResult" & crlf
'¤¤¤'						s &= tb & "AddRef           As Function ( ByVal pThis As Any Ptr ) As Ulong" & crlf
'¤¤¤'						s &= tb & "Release          As Function ( ByVal pThis As Any Ptr ) As Ulong" & crlf
'¤¤¤'						for x1 = 1 to numF
'¤¤¤'							s &= tb & st_evt(x1) & crlf
'¤¤¤'						NEXT
'¤¤¤'						s &= "End type" & crlf & crlf
'¤¤¤'						s &= "Type " & xface & "_Ev" & crlf
'¤¤¤'						s &= tb & "lpVtbl           As " & xface & "Vtbl_Ev Ptr" & crlf
'¤¤¤'						s &= "End type" & crlf & crlf
'¤¤¤'						s &= "Dim Shared " & xface & "_EvTable As " & xface & "Vtbl_Ev = ( @ Ev_Vtbl_QueryInterface, _" & crlf
'¤¤¤'						s &= tb & tb & tb & tb & tb & "@ Ev_Vtbl_AddRef, _" & crlf
'¤¤¤'						s &= tb & tb & tb & tb & tb & "@ Ev_Vtbl_Release"
'¤¤¤'						for x1 = 1 to numF
'¤¤¤'							s &= ", _" & crlf & tb & tb & tb & tb & tb & "@ " & st_M (x1)
'¤¤¤'						NEXT
'¤¤¤'						s &= ")"	 & crlf & crlf
'¤¤¤'						s &= "Dim Shared " & xface & "pSEvent As " & xface & "_Ev" & crlf & crlf
'¤¤¤'						s &= tb & tb & xface & "pSEvent.lpvtbl =  @ " & xface & "_EvTable" & crlf & crlf
'¤¤¤'						s &= "Function " & xface & "_Events_Connect ( ByVal pUnk As Any Ptr, ByRef EvdwCookie As Dword ) As Integer" & crlf
'¤¤¤'						s &= tb & "EvdwCookie = Ev_Vtbl_Cp ( pUnk , " & xface & "_Ev_IID_CP, cast( Dword, @ " & xface & "pSEvent ))" & crlf
'¤¤¤'						s &= tb & "Function = cast ( integer, EvdwCookie )" & crlf
'¤¤¤'						s &= "End Function" & crlf & crlf
'¤¤¤'						s &= "Function " & xface & "_Events_Disconnect ( ByVal pUnk As Any Ptr, ByVal EvdwCookie As Dword ) As Integer " & crlf
'¤¤¤'						s &= tb & "Function = cast ( integer, Ev_Vtbl_Cp ( pUnk, " & xface & "_Ev_IID_CP, cast( Dword, @ " & xface & "pSEvent ), EvdwCookie ) )" & crlf
'¤¤¤'						s &= "End Function" & crlf & crlf
'¤¤¤'               End If
'¤¤¤'					numF = 0
'¤¤¤'					redim st_evt(0)
'¤¤¤'					redim st_M(0)
'¤¤¤'            End If
'¤¤¤'         End If
'¤¤¤'      Wend
'¤¤¤'		function=s
'¤¤¤'   End If
'¤¤¤'
'¤¤¤'end function

'¤¤¤'Private Function EnumEvent cdecl Alias "EnumEvent"(hTree As Long, hParent() As long) As zString Ptr
'¤¤¤'   Dim Parents() As Long
'¤¤¤'   Dim hNextItem As Long
'¤¤¤'   Dim HasChilds As Long
'¤¤¤'
'¤¤¤'   Dim token As String, tldesc As String
'¤¤¤'   Dim hItem as Long
'¤¤¤'   Dim As String xface, memid, member, par, help, temp,xface0,us
'¤¤¤'   Dim As Integer cpar, bevent,pass
'¤¤¤'
'¤¤¤'   Dim s as String
'¤¤¤'	dim as string str1,st_as(),sieven
'¤¤¤'	dim as integer i_as()
'¤¤¤'	redim Parents(0)
'¤¤¤'	GetPrefix()
'¤¤¤'	str1 = trim(prefix, any "-_ ")
'¤¤¤'
'¤¤¤'   hItem = hparent(tkind_dispatch)
'¤¤¤'
'¤¤¤'   tldesc = str1 & TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem, TVGN_PARENT)))
'¤¤¤'   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
'¤¤¤'   If haschilds Then
'¤¤¤'      us = "'" & String(80, "=") & crlf
'¤¤¤'      us &= "'Event - " & tldesc & crlf
'¤¤¤'      us &= "'" & String(80, "=") & crlf
'¤¤¤'      hnextitem = haschilds
'¤¤¤'      While hnextitem
'¤¤¤'         token = TreeView_GetItemText(hTree, hNextItem)
'¤¤¤'         haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hnextitem))
'¤¤¤'         If (haschilds <> 0) And (UBound(parents) = 0) Then
'¤¤¤'				xface = str_parse(token, "?", 1)
'¤¤¤'				xface0=xface
'¤¤¤'				xface= str1 & xface
'¤¤¤'            bevent = inevtlist(xface0)
'¤¤¤'            If bevent Then
'¤¤¤'					sieven = str_parse(token, "?", 2)
'¤¤¤'					if sevent = "" or sieven = sevent THEN
'¤¤¤'						s &= "'" & String(80, "=") & crlf
'¤¤¤'						s &= "'Event Dispatch - " & xface & "     ?" & str_parse(token, "?", 3) & crlf
'¤¤¤'						s &= "Const " & xface & "_IID_CP" & "=" & dq & str_parse(token, "?", 2) & dq & crlf
'¤¤¤'						s &= "'" & String(80, "=") & crlf
'¤¤¤'               END IF
'¤¤¤'            End If
'¤¤¤'         ElseIf (UBound(parents) <> 0) then
'¤¤¤'            memid = str_parse(token, "?", 1)
'¤¤¤'            member = str_parse(token, "?", 3)
'¤¤¤'            par = str_parse(str_parse(member, "(", 2), ")", 1)
'¤¤¤'            member = str_parse(member, " As", 1)
'¤¤¤'            help = str_parse(token, "?", 4)
'¤¤¤'            cpar = str_numparse(par, ",")
'¤¤¤'				if ucase(member)<> "QUERYINTERFACE"  and ucase(member)<>"ADDREF"  _
'¤¤¤'						and ucase(member)<>"RELEASE" and ucase(member)<>"GETTYPEINFOCOUNT" _
'¤¤¤'						and ucase(member)<>"GETTYPEINFO"  and ucase(member)<>"GETIDSOFNAMES" _
'¤¤¤'						and ucase(member)<>"INVOKE" THEN
'¤¤¤'					If Len(par) = 0 Then cpar = 0
'¤¤¤'					If bevent and (sevent = "" or sieven = sevent) Then
'¤¤¤'						s &= "'" & String(80, "=") & crlf
'¤¤¤'						s &= "'Event=" & member & ", ID=" & memid & " ?" & help & crlf
'¤¤¤'						s &= "'" & String(80, "=") & crlf
'¤¤¤'						s &= "Function " & xface & "_" & member & "(ByVal pCookie As Events_IDispatchVtbl Ptr, ByVal pdispparams As DISPPARAMS Ptr) As HRESULT" & crlf
'¤¤¤'						s &= tb & "Dim pThis As Dword=>pCookie->pThis" & crlf
'¤¤¤'						s &= tb & "Dim IdEvent As Dword = cast(Dword, pCookie)" & crlf
'¤¤¤'						temp &= tb & tb & tb & "Case " & memid & " '" & help & crlf
'¤¤¤'						temp &= tb & tb & tb & tb & "Function=" & xface & "_" & member & "(cast(Any Ptr,pUnk), pdispparams)" & crlf
'¤¤¤'						If cpar Then s &= tb & "Dim pv As Variant Ptr = pdispparams->rgVArg" & crlf
'¤¤¤'						redim i_as(1 to cpar)
'¤¤¤'						redim st_as(1 to cpar)
'¤¤¤'						For i As Short = 1 To cpar
'¤¤¤'							st_as(i)= trim(str_parse(par, ",", i))
'¤¤¤'							if left(st_as(i),5)= "Data " THEN st_as(i) = "Data1 " & mid (st_as(i),6)
'¤¤¤'							if instr(st_as(i)," DataObject Ptr") then
'¤¤¤'								st_as(i)=FF_Replace( st_as(i), " DataObject Ptr"," IDispatch Ptr    ' DataObject Ptr" )
'¤¤¤'							elseif instr(st_as(i)," DataObject") then
'¤¤¤'								st_as(i)=FF_Replace( st_as(i), " DataObject"," IDispatch    ' DataObject" )
'¤¤¤'							end if
'¤¤¤'							if right(st_as(i),3)= "Ptr" THEN st_as(i) = left(st_as(i) ,len(st_as(i))- 4)
'¤¤¤'							s &= tb & "Dim " & st_as(i)  & crlf
'¤¤¤'							if right(st_as(i),3)= "Ptr" THEN i_as(i)=1
'¤¤¤'						Next
'¤¤¤'						s &= crlf
'¤¤¤'						For i As Short = 1 To cpar
'¤¤¤'							if i_as(i)=1 THEN
'¤¤¤'								s &= tb & str_parse(st_as(i), " As", 1) & "=cast(Any Ptr,CUInt(VariantV(pv[" & cpar - i & "])))" & crlf
'¤¤¤'							else
'¤¤¤'								s &= tb & str_parse(st_as(i), " As", 1) & "=VariantV(pv[" & cpar - i & "])" & crlf
'¤¤¤'							end if
'¤¤¤'						Next
'¤¤¤'						s &= tb & "'*** Put your code here ***" & crlf & crlf
'¤¤¤'						s &= tb & "Function=S_OK" & crlf
'¤¤¤'						s &= "End Function" & crlf & crlf
'¤¤¤'					End If
'¤¤¤'				End If
'¤¤¤'         End If
'¤¤¤'         If haschilds Then
'¤¤¤'            ReDim Preserve parents(UBound(parents) + 1)
'¤¤¤'            parents(UBound(parents)) = hnextitem
'¤¤¤'            hnextitem = haschilds
'¤¤¤'         Else
'¤¤¤'            hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
'¤¤¤'            If hnextitem = 0 And UBound(parents) > 0 Then
'¤¤¤'               hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), parents(UBound(parents)), TVGN_NEXT))
'¤¤¤'               ReDim Preserve parents(UBound(parents) - 1)
'¤¤¤'               If bevent and (sevent = "" or sieven = sevent) Then
'¤¤¤'                  temp = "Function " & xface & "_IDispatch_Invoke(ByVal pUnk As IDispatch Ptr, ByVal dispidMember As DispID, ByVal riid As IID Ptr, _" & crlf & _
'¤¤¤'                        "  ByVal lcid As LCID, ByVal wFlags As Ushort, ByVal pdispparams As DISPPARAMS Ptr, Byval pvarResult As Variant Ptr, _" & crlf & _
'¤¤¤'                        "  ByVal pexcepinfo As EXCEPINFO Ptr, ByVal puArgErr As Uint Ptr) As Hresult" & crlf & crlf & _
'¤¤¤'                        tb & "If Varptr(pdispparams) Then" & crlf & _
'¤¤¤'                        tb & tb & "Select Case dispidMember" & crlf & _
'¤¤¤'                        temp & _
'¤¤¤'                        tb & tb & tb & "Case Else" & crlf & _
'¤¤¤'                        tb & tb & tb & tb & "Function = DISP_E_MEMBERNOTFOUND" & crlf & _
'¤¤¤'                        tb & tb & "End Select" & crlf & _
'¤¤¤'                        tb & "End If" & crlf & _
'¤¤¤'                        "End Function" & crlf & crlf
'¤¤¤'                  s &= temp
'¤¤¤'                  temp = ""
'¤¤¤'
'¤¤¤'                  s &= "' --------------------------------------------------------------------------------------------" & CRLF
'¤¤¤'                  s &= "' " & xface & " Events Connection function" & crlf
'¤¤¤'                  s &= "' --------------------------------------------------------------------------------------------" & CRLF
'¤¤¤'                  s &= "Function " & xface & "_Events_Connect ( ByVal pUnk As Any Ptr, ByRef dwcookie As Dword ) As Integer" & crlf
'¤¤¤'						s &= tb & "Function = Ax_Events_cnx ( " & xface & "_IID_CP , pUnk , dwcookie, ProcPtr ( " & xface & "_IDispatch_Invoke ))" & crlf
'¤¤¤'                  s &= "End Function" & crlf & crlf
'¤¤¤'                  s &= "Function " & xface & "_Events_Disconnect ( ByVal pUnk As Any Ptr, ByVal dwcookie As Dword ) As Integer" & crlf
'¤¤¤'						s &= tb & "Function = Ax_Events_cnx ( " & xface & "_IID_CP , pUnk , dwcookie )" & crlf
'¤¤¤'                  s &= "End Function" & crlf & crlf
'¤¤¤'               End If
'¤¤¤'            End If
'¤¤¤'         End If
'¤¤¤'      Wend
'¤¤¤'		if s<>"" THEN s= us & crlf & s
'¤¤¤'   End If
'¤¤¤'	if s="" THEN
'¤¤¤'		if pass =1 THEN
'¤¤¤'			s=" No Event "
'¤¤¤'		else
'¤¤¤'			hItem = hparent(tkind_interface)
'¤¤¤'			s="" : us = ""
'¤¤¤'			pass =1
'¤¤¤'			s=EnumEvent2 (hTree , hParent(), hItem )
'¤¤¤'			if s="" THEN
'¤¤¤'				s= "'  No Events ! "
'¤¤¤'			else
'¤¤¤'				s= "'       vTable Events ! " & crlf & crlf & crlf & s
'¤¤¤'         END IF
'¤¤¤'      END IF
'¤¤¤'	else
'¤¤¤'		s= "'       Dispach Events ! " & crlf & crlf & crlf & s
'¤¤¤'	END IF
'¤¤¤'   Return cast(zstring ptr, sysallocstringbytelen(s, len(s)))
'¤¤¤'End Function

'¤¤¤'Private Function EnumEnum cdecl Alias "EnumEnum"(hTree As Long, hParent() As long) As zString Ptr
'¤¤¤'   Dim Parents() As Long
'¤¤¤'   Dim hNextItem As Long
'¤¤¤'   Dim HasChilds As Long
'¤¤¤'
'¤¤¤'   Dim token As String, tldesc As String
'¤¤¤'   Dim hItem as Long
'¤¤¤'
'¤¤¤'   Dim s as String
'¤¤¤'
'¤¤¤'	redim Parents(0)
'¤¤¤'   hItem = hparent(tkind_enum)
'¤¤¤'   tldesc = TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem, TVGN_PARENT)))
'¤¤¤'   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
'¤¤¤'   If haschilds Then
'¤¤¤'      s &= "'" & String(80, "=") & crlf
'¤¤¤'      s &= "'Enum -  " & tldesc & crlf
'¤¤¤'      s &= "'" & String(80, "=") & crlf
'¤¤¤'      hnextitem = haschilds
'¤¤¤'      While hnextitem
'¤¤¤'         token = TreeView_GetItemText(hTree, hNextItem)
'¤¤¤'         haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hnextitem))
'¤¤¤'         If UBound(parents) = 0 Then
'¤¤¤'            s &= "'" & String(80, "=") & crlf
'¤¤¤'            s &= "Type " & str_parse(token, "?", 1) & " As Integer     '" & str_parse(token, "?", 2) & crlf
'¤¤¤'            s &= "'" & String(80, "=") & crlf
'¤¤¤'         Else
'¤¤¤'            s &= str_parse(token, "?", 1) & "     '" & str_parse(token, "?", 1) & crlf
'¤¤¤'         End If
'¤¤¤'         If haschilds Then
'¤¤¤'            ReDim Preserve parents(UBound(parents) + 1)
'¤¤¤'            parents(UBound(parents)) = hnextitem
'¤¤¤'            hnextitem = haschilds
'¤¤¤'         Else
'¤¤¤'            hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
'¤¤¤'            If hnextitem = 0 And UBound(parents) > 0 Then
'¤¤¤'               hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), parents(UBound(parents)), TVGN_NEXT))
'¤¤¤'               ReDim Preserve parents(UBound(parents) - 1)
'¤¤¤'               s &= crlf
'¤¤¤'            End If
'¤¤¤'         end if
'¤¤¤'      Wend
'¤¤¤'   End If
'¤¤¤'   Return cast(zstring ptr, sysallocstringbytelen(s, len(s)))
'¤¤¤'End Function

'¤¤¤'Private Function EnumRecord cdecl Alias "EnumRecord"(hTree As Long, hParent() As long) As zString Ptr
'¤¤¤'   Dim Parents() As Long
'¤¤¤'   Dim hNextItem As Long
'¤¤¤'   Dim HasChilds As Long
'¤¤¤'   Dim token As String
'¤¤¤'   Dim hItem as Long
'¤¤¤'   Dim s as String
'¤¤¤'	redim Parents(0)
'¤¤¤'   hItem = hparent(tkind_record)
'¤¤¤'   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
'¤¤¤'   If haschilds Then
'¤¤¤'      hnextitem = haschilds
'¤¤¤'      While hnextitem
'¤¤¤'         token = TreeView_GetItemText(hTree, hNextItem)
'¤¤¤'         haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hnextitem))
'¤¤¤'         If haschilds <> 0 And UBound(parents) = 0 Then
'¤¤¤'            s &= "'" & String(80, "=") & crlf
'¤¤¤'            s &= "'Record " & str_parse(token, "?", 1) & "     '" & str_parse(token, "?", 2) & crlf
'¤¤¤'            s &= "'" & String(80, "=") & crlf
'¤¤¤'            s &= "Type " & str_parse(token, "?", 1) & crlf
'¤¤¤'         ElseIf UBound(parents) <> 0 then
'¤¤¤'            s &= tb & token & crlf
'¤¤¤'         End If
'¤¤¤'
'¤¤¤'         If haschilds Then
'¤¤¤'            ReDim Preserve parents(UBound(parents) + 1)
'¤¤¤'            parents(UBound(parents)) = hnextitem
'¤¤¤'            hnextitem = haschilds
'¤¤¤'         Else
'¤¤¤'            hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
'¤¤¤'            If hnextitem = 0 And UBound(parents) > 0 Then
'¤¤¤'               hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), parents(UBound(parents)), TVGN_NEXT))
'¤¤¤'               ReDim Preserve parents(UBound(parents) - 1)
'¤¤¤'               s &= "End Type" & crlf
'¤¤¤'            End If
'¤¤¤'         End If
'¤¤¤'      Wend
'¤¤¤'   End If
'¤¤¤'   Return cast(zstring ptr, sysallocstringbytelen(s, len(s)))
'¤¤¤'End Function

'¤¤¤'Private Function EnumUnion cdecl Alias "EnumUnion"(hTree As Long, hParent() As long) As zString Ptr
'¤¤¤'   Dim Parents() As Long
'¤¤¤'   Dim hNextItem As Long
'¤¤¤'   Dim HasChilds As Long
'¤¤¤'   Dim token As String
'¤¤¤'   Dim hItem as Long
'¤¤¤'
'¤¤¤'   Dim s as String
'¤¤¤'
'¤¤¤'	redim Parents(0)
'¤¤¤'   hItem = hparent(tkind_union)
'¤¤¤'   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
'¤¤¤'   If haschilds Then
'¤¤¤'      hnextitem = haschilds
'¤¤¤'      While hnextitem
'¤¤¤'         token = TreeView_GetItemText(hTree, hNextItem)
'¤¤¤'         haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hnextitem))
'¤¤¤'         If haschilds <> 0 And UBound(parents) = 0 Then
'¤¤¤'            s &= "'" & String(80, "=") & crlf
'¤¤¤'            s &= "'Union " & str_parse(token, "?", 1) & "     '" & str_parse(token, "?", 2) & crlf
'¤¤¤'            s &= "'" & String(80, "=") & crlf
'¤¤¤'            s &= "Union " & str_parse(token, "?", 1) & crlf
'¤¤¤'         Else
'¤¤¤'            s &= tb & token & crlf
'¤¤¤'         End If
'¤¤¤'
'¤¤¤'         If haschilds Then
'¤¤¤'            ReDim Preserve parents(UBound(parents) + 1)
'¤¤¤'            parents(UBound(parents)) = hnextitem
'¤¤¤'            hnextitem = haschilds
'¤¤¤'         Else
'¤¤¤'            hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
'¤¤¤'            If hnextitem = 0 And UBound(parents) > 0 Then
'¤¤¤'               hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), parents(UBound(parents)), TVGN_NEXT))
'¤¤¤'               ReDim Preserve parents(UBound(parents) - 1)
'¤¤¤'               s &= "End Union" & crlf
'¤¤¤'            End If
'¤¤¤'         End If
'¤¤¤'      Wend
'¤¤¤'   End If
'¤¤¤'   Return cast(zstring ptr, sysallocstringbytelen(s, len(s)))
'¤¤¤'End Function

'¤¤¤'Private Function EnumModule cdecl Alias "EnumModule"(hTree As Long,hParent() As long)As zString Ptr
'¤¤¤'	Dim Parents() As Long
'¤¤¤'	Dim hNextItem As Long
'¤¤¤'	Dim HasChilds As Long
'¤¤¤'	Dim token As String,tldesc As String
'¤¤¤'	Dim hItem as Long
'¤¤¤'	Dim as String s,func,dlls,aliass,ord,us,d1,d2
'¤¤¤'	redim Parents(0)
'¤¤¤'	hItem=hparent(tkind_module)
'¤¤¤'	tldesc = TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem, TVGN_PARENT)))
'¤¤¤'   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
'¤¤¤'	If haschilds Then
'¤¤¤'	s &="'" & String(80,"=") & crlf
'¤¤¤'	s &="'Module - " & tldesc & crlf
'¤¤¤'	s &="'" & String(80,"=") & crlf
'¤¤¤'	hnextitem=haschilds
'¤¤¤'	While hnextitem
'¤¤¤'
'¤¤¤'		token=TreeView_GetItemText(hTree,hNextItem)
'¤¤¤'		If UBound(parents)=0 Then
'¤¤¤'			s &="'" & String(80,"=") & crlf
'¤¤¤'			s &="'" & str_parse(token,"?",1) & "     ?" & str_parse(token,"?",2) & crlf
'¤¤¤'			s &="'" & String(80,"=") & crlf
'¤¤¤'		Else
'¤¤¤'			If Left(token,5)<>"Const" Then
'¤¤¤'				dlls=str_parse(token,"?",1):aliass=str_parse(token,"?",2)
'¤¤¤'				ord=str_parse(token,"?",3):func=str_parse(token,"?",4)
'¤¤¤'				s &="Dim shared " & func & "     '" & str_parse(token,"?",5) & crlf
'¤¤¤'				If Len(aliass) Then
'¤¤¤'					s &=str_parse(func," As",1) & " = DylibSymbol(h" & str_parse(dlls,".",1) & "," & dq & aliass & dq & ")"
'¤¤¤'				Else
'¤¤¤'					s &=str_parse(func," As",1) & " = DylibSymbol(h" & str_parse(dlls,".",1) & "," & ord & ")"
'¤¤¤'				EndIf
'¤¤¤'				s &="     'h" & str_parse(dlls,".",1) & " = DylibLoad(" & dq & dlls & dq & ")" & crlf
'¤¤¤'				d1="h" & str_parse(dlls,".",1) & " = DylibLoad(" & dq & dlls & dq & ")"
'¤¤¤'				if instr(ucase(us),ucase("Dim Shared h" & str_parse(dlls,".",1)))= 0 THEN
'¤¤¤'					d2=selpath
'¤¤¤'					if instr( d2,dlls)THEN
'¤¤¤'						d1= ff_replace(d1,dlls,d2)
'¤¤¤'               END IF
'¤¤¤'					us &= "Dim Shared h" & str_parse(dlls,".",1) & " As Any Ptr " & crlf & d1 & crlf & crlf
'¤¤¤'            END IF
'¤¤¤'			Else
'¤¤¤'				s &=str_parse(token,"?",1) & "     '" & str_parse(token,"?",2) & crlf
'¤¤¤'			EndIf
'¤¤¤'		EndIf
'¤¤¤'
'¤¤¤'		haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hnextitem))
'¤¤¤'		If haschilds Then
'¤¤¤'			ReDim Preserve parents(UBound(parents)+1)
'¤¤¤'			parents(UBound(parents))=hnextitem
'¤¤¤'			hnextitem=haschilds
'¤¤¤'		Else
'¤¤¤'			hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
'¤¤¤'			If hnextitem=0 And UBound(parents)>0 Then
'¤¤¤'				hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), parents(UBound(parents)), TVGN_NEXT))
'¤¤¤'				ReDim Preserve parents(UBound(parents)-1)
'¤¤¤'				s &=crlf
'¤¤¤'			EndIf
'¤¤¤'		endif
'¤¤¤'	Wend
'¤¤¤'	EndIf
'¤¤¤'	s= us & crlf & crlf & s
'¤¤¤'	Return cast(zstring ptr, sysallocstringbytelen(s, len(s)))
'¤¤¤'End Function

'¤¤¤'Private Function EnumAlias cdecl Alias "EnumAlias"(hTree As Long, hParent() As long) As zString Ptr
'¤¤¤'   Dim Parents() As Long
'¤¤¤'   Dim hNextItem As Long
'¤¤¤'   Dim HasChilds As Long
'¤¤¤'   Dim token As String
'¤¤¤'   Dim hItem as Long
'¤¤¤'   Dim as String s, tldesc
'¤¤¤'	redim Parents(0)
'¤¤¤'   hItem = hparent(tkind_alias)
'¤¤¤'   tldesc = TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem, TVGN_PARENT)))
'¤¤¤'   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
'¤¤¤'   If haschilds Then
'¤¤¤'      s &= "'" & String(80, "=") & crlf
'¤¤¤'      s &= "'Alias - " & tldesc & crlf
'¤¤¤'      s &= "'" & String(80, "=") & crlf
'¤¤¤'      hnextitem = haschilds
'¤¤¤'      While hnextitem
'¤¤¤'         token = TreeView_GetItemText(hTree, hNextItem)
'¤¤¤'         s &= token & crlf
'¤¤¤'         hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
'¤¤¤'      Wend
'¤¤¤'   End If
'¤¤¤'   Return cast(zstring ptr, sysallocstringbytelen(s, len(s)))
'¤¤¤'End Function

'¤¤¤'Private Function EnumCoClass cdecl Alias "EnumCoClass"(hTree As Long, hParent() As long) As zString Ptr
'¤¤¤'   Dim Parents() As Long
'¤¤¤'   Dim hNextItem As Long
'¤¤¤'   Dim HasChilds As Long
'¤¤¤'   Dim token As String, progid As WString Ptr, clssid As clsid
'¤¤¤'   Dim As String clsids, progids, tldesc
'¤¤¤'   Dim hItem as long
'¤¤¤'
'¤¤¤'   Dim s as string, str1 as string
'¤¤¤'	redim Parents(0)
'¤¤¤'	GetPrefix()
'¤¤¤'	str1 = trim(prefix, any "-_ ")
'¤¤¤'
'¤¤¤'	hItem = hparent(tkind_coclass)
'¤¤¤'   tldesc = TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem, TVGN_PARENT)))
'¤¤¤'   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
'¤¤¤'   If haschilds Then
'¤¤¤'      hnextitem = haschilds
'¤¤¤'      While hnextitem
'¤¤¤'         token = TreeView_GetItemText(hTree, hNextItem)
'¤¤¤'         clsids &= "   Const CLSID_" & str1 & str_parse(token, "?", 1) & "=" & dq & str_parse(token, "?", 2) & dq & "     '" & str_parse(token, "?", 3) & crlf
'¤¤¤'         progid = Stringtobstr(str_parse(token, "?", 2))
'¤¤¤'         clsidfromString(progid, @clssid)
'¤¤¤'         sysfreeString(progid)
'¤¤¤'         progidfromclsid(@clssid, @progid)
'¤¤¤'			If Len(*progid) Then progids &="Const ProgID_" & str_parse(token,"?",1) & "=" & dq & *progid & dq & crlf
'¤¤¤'         sysfreeString(progid)
'¤¤¤'         hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
'¤¤¤'      Wend
'¤¤¤'   End If
'¤¤¤'	s &= "' Warning : Do not use vTable and Invoke/Dispatch methods on the same program !" & crlf
'¤¤¤'	s &= "'           It will be duplicated constants and type confusion "& crlf
'¤¤¤'	s &= "'           or use different prefix to differentiate them ..." & crlf & crlf
'¤¤¤'   s &= "'" & String(80, "=") & crlf
'¤¤¤'   s &= "'CLSID - " & tldesc & crlf
'¤¤¤'   s &= "'" & String(80, "=") & crlf
'¤¤¤'   s &= clsids & crlf
'¤¤¤'   s &= "'" & String(80, "=") & crlf
'¤¤¤'   s &= "'ProgID - " & tldesc & crlf
'¤¤¤'   s &= "'" & String(80, "=") & crlf
'¤¤¤'   s &= progids & crlf
'¤¤¤'   Return cast(zstring ptr, SysAllocStringByteLen(s, len(s)))
'¤¤¤'End Function

'¤¤¤'Private Function Full_control(xface as string, s_connect as string, s_disc as string, L1 as integer) as string
'¤¤¤'   dim as string s
'¤¤¤'	if L1 THEN
'¤¤¤'		if notsure = 1 THEN
'¤¤¤'			s = crlf  & tb & "' The events are not connected ....." & crlf
'¤¤¤'			s &= tb & "' not sure : more than 1 connection to connect, check yourself ... if ok uncomment or modify " & crlf
'¤¤¤'			s &= tb & "'   Declare " & s_connect & "(ByVal As any Ptr,ByRef As Dword) As Integer" & crlf
'¤¤¤'			s &= tb & "'   Declare " & s_disc & "(ByVal As any Ptr, ByVal As Dword) As Integer"  & crlf & crlf & crlf
'¤¤¤'		else
'¤¤¤'			s = crlf &  tb & "' The events can be connected ....." & crlf
'¤¤¤'			s &= tb & "#Define Decl_Ev1_" & xface & " Declare Function " & s_connect & "(ByVal As any Ptr, ByRef As dword) As Integer" & crlf
'¤¤¤'			s &= tb & "#Define Decl_Ev2_" & xface & " Declare Function " & s_disc & "(ByVal As any Ptr, ByVal As dword) As Integer"  & crlf
'¤¤¤'			s &= tb & "Decl_Ev1_" & xface & "      ' indirect declare to avoid FireFly duplication " & crlf
'¤¤¤'			s &= tb & "Decl_Ev2_" & xface & "      ' indirect declare to avoid FireFly duplication " & crlf
'¤¤¤'		END IF
'¤¤¤'	END IF
'¤¤¤'	s &= "#EndIf" & crlf & crlf & crlf
'¤¤¤'	s &= "Sub " & xface & "Call_Init()			' be called from initialization of the control form" & crlf
'¤¤¤'   s &= tb & xface & "Obj_Ptr = AxCreate_Object( " & xface & "OcxHwnd  )  			             'get object control address with control hwnd" & crlf & crlf
'¤¤¤'	s &= tb & "'    if use of vTable functions ...  Dim Shared " & xface & "pVTI As ... with the right TYPE see in vtable code" & crlf
'¤¤¤'	s &= tb & "'    " & xface & "pVTI = " & xface & "Obj_Ptr   ':  " & xface & "pVTI->lpvtbl->SetColor( " & xface & "pVTI ,... )" & crlf
'¤¤¤'   s &= tb & "'                            or Ax_Vt ( " & xface & "pVTI, SetColor, arg... )  or  Ax_Vt0 ( " & xface & "pVTI, About )" & crlf & crlf
'¤¤¤'   s &= tb & "'    if use of Invoke/dispatch functions ...  Dim Shared " & xface & "pDisp As ... with the right TYPE see in Invoke code" & crlf
'¤¤¤'	s &= tb & "'    SetObj ( @" & xface & "pDisp , " & xface & "Obj_Ptr )   ':   AxCall " & xface & "pDisp.PutValue , vptr(x) , vptr(y) , ..." & crlf & crlf
'¤¤¤'
'¤¤¤'	if L1 THEN
'¤¤¤'		if notsure = 1 THEN
'¤¤¤'			s &= tb & "' not sure : more than 1 connection to connect, check yourself ... if ok uncomment or modify " & crlf
'¤¤¤'			s &= tb & "'  " & s_connect & "( " & xface & "Obj_Ptr , " & xface & "Obj_Event ) 'connect object with its event" & crlf & crlf
'¤¤¤'		else
'¤¤¤'			s &= tb &"' The events are now connected ....." & crlf
'¤¤¤'			s &= tb & s_connect & "( " & xface & "Obj_Ptr , " & xface & "Obj_Event ) 'connect object with its event" & crlf & crlf
'¤¤¤'		end if
'¤¤¤'	END IF
'¤¤¤'   s &= tb & xface & "Call_Sett() 'initial settings if you want some" & crlf
'¤¤¤'   s &= "End Sub" & crlf & crlf & crlf
'¤¤¤'   s &= "Sub " & xface & "Call_OnClose() ' normaly be called from close form command" & crlf
'¤¤¤'   if L1 THEN
'¤¤¤'		if notsure = 1 THEN
'¤¤¤'			s &= tb & "' not sure : more than 1 connection to connect, check yourself ... if ok uncomment or modify " & crlf
'¤¤¤'			s &= tb & "'  " & s_disc & "( " & xface & "Obj_Ptr , " & xface & "Obj_Event ) 'disconnect  event from object" & crlf
'¤¤¤'		else
'¤¤¤'			s &= tb &"' The events are now disconnected ....." & crlf
'¤¤¤'			s &= tb & s_disc & " ( " & xface & "Obj_Ptr , " & xface & "Obj_Event ) 'disconnect  event from object" & crlf
'¤¤¤'		end if
'¤¤¤'	END IF
'¤¤¤'   s &= tb & " AxRelease_Object(" & xface & "Obj_Ptr)     'release object" & crlf
'¤¤¤'   s &= tb & "'AxStop()         'only one by project, better on the WM_close of last Form" & crlf
'¤¤¤'   s &= "End Sub" & crlf & crlf & crlf
'¤¤¤'   s &= "Sub " & xface & "Call_Sett()   'initial settings here" & crlf
'¤¤¤'   s &= tb & "'ex put/ get /call values  " & crlf
'¤¤¤'   s &= "End Sub" & crlf
'¤¤¤'   s &= "'================================================================================" & crlf
'¤¤¤'	if isevent=0 THEN
'¤¤¤'			s &= "'  No Events for that ActiveX Control interface " & crlf
'¤¤¤'	elseif L1 =0 and isevent=1 THEN
'¤¤¤'			s &= "'  Events exist for that ActiveX Control interface, if needed select Template with events !" & crlf
'¤¤¤'	end if
'¤¤¤'	s &= "'================================================================================" & crlf
'¤¤¤'   function = s
'¤¤¤'END FUNCTION

'¤¤¤'Private Function Simple_control(xface as string, ret1 as string, tldesc as string, progids as string, l1 as integer) as string
'¤¤¤'   dim as string s
'¤¤¤'   s = "'" & String(80, "=") & crlf
'¤¤¤'   s &= "'ProgID - " & tldesc & crlf
'¤¤¤'   s &= "'" & String(80, "=") & crlf
'¤¤¤'   s &= progids & crlf
'¤¤¤'
'¤¤¤'   if ret1 = " ProgID " then
'¤¤¤'      s &= "'    not unique ProgID   - Select only 1 CoClass in tree to get the one needed to create object " & crlf & crlf
'¤¤¤'		function = s
'¤¤¤'		if xfull = 0 THEN exit function
'¤¤¤'	elseif ret1 = " No ProgID " then
'¤¤¤'		s &= "'    No ProgID   - Select 1 CoClass in tree to get the one needed to create object " & crlf & crlf
'¤¤¤'		function = s
'¤¤¤'		if xfull = 0 THEN exit function
'¤¤¤'	end if
'¤¤¤'   s &= "'" & String(80, "=") & crlf & crlf
'¤¤¤'   s &= "'   Use Prefix in AxSuite3 to differenciate the functions names / variables names" & crlf & crlf
'¤¤¤'   s &= "'================================================================================" & crlf
'¤¤¤'   s &= "'for that control :  use  variable Classname > Var_ATL_Win_     and Caption = " & ret1  & crlf
'¤¤¤'   s &= "'================================================================================" & crlf
'¤¤¤'   s &= "Dim Shared As any ptr      " & xface & "Obj_Ptr      ' object Ptr   " & crlf
'¤¤¤'   if l1 = 1 then s &= "Dim Shared As Dword        " & xface & "Obj_Event    ' cookie for object events" & crlf
'¤¤¤'	s &= "Dim Shared As hwnd         " & xface & "OcxHwnd      ' Ocx handle " & crlf & crlf
'¤¤¤'   s &=  "' If use of Firefly Custom Control for OCX,  this following function is not needed, you can comment it... " & crlf
'¤¤¤'	s &=  "Function " & xface & "WinOcx  ( hparent as hwnd ) as hwnd 		' make window for control"& crlf
'¤¤¤'	s &= tb & "Dim as integer x= 0 								' x left position , change for your need" & crlf
'¤¤¤'	s &= tb & "Dim as integer y= 0 								' y top position , change for your need" & crlf
'¤¤¤'	s &= tb & "Dim as integer w= 200 							' w width  , change for your need" & crlf
'¤¤¤'	s &= tb & "Dim as integer h= 150 							' h heigth  , change for your need" & crlf
'¤¤¤'	s &= tb & "dim as string N = """ & xface & " WinOcxName1""  					' Name of the window , change for your need" & crlf
'¤¤¤'	s &= tb & "dim as string P = """ & ret1 & """  			' prodId of the window , change for your need" & crlf
'¤¤¤'	s &= tb & xface & "OcxHwnd = AxWinChild( hparent , N , P , x , y , w , h ) 		' adapt to your needs  see Axlite.bi" & crlf
'¤¤¤'	s &= tb & xface & "'OcxHwnd = AxWinTool( hparent , N , P , x , y , w , h ) 		' or try this for Toll floating window" & crlf
'¤¤¤'	s &= tb & xface & "'OcxHwnd = AxWinFull( hparent , N , P , x , y , w , h ) 		' or try this for Full floating window" & crlf
'¤¤¤'	s &= tb & "Function = " & xface & "OcxHwnd" & crlf
'¤¤¤'	s &= "End Function " & crlf & crlf
'¤¤¤'	s &= "#IfnDef _PARSED_MOD_FF3_         'Defined in FireFly to avoid duplications " & crlf & crlf
'¤¤¤'	s &= tb & "#Define Decl_Set_" & xface & " Declare Sub " & xface & "Call_Sett()" & crlf
'¤¤¤'	s &= tb & "Decl_Set_" & xface & "      ' indirect declare to avoid FireFly duplication " & crlf & crlf
'¤¤¤'   function = s
'¤¤¤'END FUNCTION

'¤¤¤'Private Function Full_nowin(xface as string, s_connect as string, s_disc as string, L1 as integer,ret1 as string) as string
'¤¤¤'   dim as string s
'¤¤¤'	if L1 THEN
'¤¤¤'		if notsure = 1 THEN
'¤¤¤'			s = crlf  & tb & "' The events are not connected ....." & crlf
'¤¤¤'			s &= tb & "' not sure : more than 1 connection to connect, check yourself ... if ok uncomment or modify " & crlf
'¤¤¤'			s &= tb & "'   Declare " & s_connect & "(ByVal As any Ptr, ByRef As Dword) As Integer" & crlf
'¤¤¤'			s &= tb & "'   Declare " & s_disc & "(ByVal As any Ptr, ByVal As Dword) As Integer"  & crlf & crlf & crlf
'¤¤¤'		else
'¤¤¤'			s = crlf & tb & "' The events can be connected ....." & crlf
'¤¤¤'			s &= tb & "#Define Decl_Ev1_" & xface & " Declare Function " & s_connect & "(ByVal As any Ptr, ByRef As dword) As Integer" & crlf
'¤¤¤'			s &= tb & "#Define Decl_Ev2_" & xface & " Declare Function " & s_disc & "(ByVal As any Ptr, ByVal As dword) As Integer"  & crlf
'¤¤¤'			s &= tb & "Decl_Ev1_" & xface & "      ' indirect declare to avoid FireFly duplication " & crlf
'¤¤¤'			s &= tb & "Decl_Ev2_" & xface & "      ' indirect declare to avoid FireFly duplication " & crlf
'¤¤¤'		END IF
'¤¤¤'	END IF
'¤¤¤'	s &= "#EndIf" & crlf & crlf & crlf
'¤¤¤'	s &= "Sub " & xface & "Call_Init()				' be called from initialization " & crlf
'¤¤¤'   s &= tb & xface & "Obj_Ptr = AxCreate_Object( " & ret1 & " )  			             'get object control address with prodid" & crlf & crlf
'¤¤¤'	s &= tb & "'    if use of vTable functions ...  Dim Shared " & xface & "pVTI As ... with the right TYPE see in vtable code" & crlf
'¤¤¤'	s &= tb & "'    " & xface & "pVTI = " & xface & "Obj_Ptr   ':  " & xface & "pVTI->lpvtbl->SetColor( " & xface & "pVTI ,... )" & crlf
'¤¤¤'   s &= tb & "'                            or Ax_Vt ( " & xface & "pVTI, SetColor, arg... )  or  Ax_Vt0 ( " & xface & "pVTI, About )" & crlf & crlf
'¤¤¤'	s &= tb & "'    if use of Invoke/dispatch functions ...  Dim Shared " & xface & "pDisp As ... with the right TYPE see in Invoke code" & crlf
'¤¤¤'	s &= tb & "'    SetObj ( @" & xface & "pDisp , " & xface & "Obj_Ptr )   ':   AxCall " & xface & "pDisp.PutValue , vptr(x) , vptr(y) , ..." & crlf & crlf
'¤¤¤'
'¤¤¤'	if L1 THEN
'¤¤¤'		if notsure = 1 THEN
'¤¤¤'			s &= tb & "' not sure : more than 1 connection to connect, check yourself ... if ok uncomment or modify " & crlf
'¤¤¤'			s &= tb & "'  " & s_connect & " ( " & xface & "Obj_Ptr , " & xface & "Obj_Event ) 'connect object with its event" & crlf & crlf
'¤¤¤'		else
'¤¤¤'			s &= tb &"' The events are now connected ....." & crlf
'¤¤¤'			s &= tb & s_connect & " (  " & xface & "Obj_Ptr , " & xface & "Obj_Event ) 'connect object with its event" & crlf & crlf
'¤¤¤'		end if
'¤¤¤'	END IF
'¤¤¤'   s &= tb & xface & "Call_Sett() 'initial settings if you want some" & crlf
'¤¤¤'   s &= "End Sub" & crlf & crlf & crlf
'¤¤¤'   s &= "Sub " & xface & "Call_OnClose() ' normaly be called from close form command" & crlf
'¤¤¤'   if L1 THEN
'¤¤¤'		if notsure = 1 THEN
'¤¤¤'			s &= tb & "' not sure : more than 1 connection to connect, check yourself ... if ok uncomment or modify " & crlf
'¤¤¤'			s &= tb & "'  " & s_disc & "( " & xface & "Obj_Ptr , " & xface & "Obj_Event ) 'disconnect  event from object" & crlf
'¤¤¤'		else
'¤¤¤'			s &= tb &"' The events are now disconnected ....." & crlf
'¤¤¤'			s &= tb & s_disc & "( " & xface & "Obj_Ptr , " & xface & "Obj_Event ) 'disconnect  event from object" & crlf
'¤¤¤'		end if
'¤¤¤'	END IF
'¤¤¤'   s &= tb & " AxRelease_Object( " & xface & "Obj_Ptr )     'release object" & crlf
'¤¤¤'   s &= tb & "'AxStop()         'only one by project, better on the WM_close of last Form" & crlf
'¤¤¤'   s &= "End Sub" & crlf & crlf & crlf
'¤¤¤'   s &= "Sub " & xface & "Call_Sett()   'initial settings here" & crlf
'¤¤¤'   s &= tb & "'ex put/ get /call values  " & crlf
'¤¤¤'   s &= "End Sub" & crlf
'¤¤¤'   s &= "'================================================================================" & crlf
'¤¤¤'	if isevent=0 THEN
'¤¤¤'			s &= "'  No Events for that ActiveX Control interface " & crlf
'¤¤¤'	elseif L1 =0 and isevent=1 THEN
'¤¤¤'			s &= "'  Events exist for that ActiveX Control interface, if needed select Template with events !" & crlf
'¤¤¤'	end if
'¤¤¤'	s &= "'================================================================================" & crlf
'¤¤¤'   function = s
'¤¤¤'END FUNCTION

'¤¤¤'Private Function Simple_nowin(xface as string, ret1 as string, tldesc as string, progids as string, l1 as integer) as string
'¤¤¤'   dim as string s
'¤¤¤'   s = "'" & String(80, "=") & crlf
'¤¤¤'   s &= "'ProgID - " & tldesc & crlf
'¤¤¤'   s &= "'" & String(80, "=") & crlf
'¤¤¤'   s &= progids & crlf
'¤¤¤'
'¤¤¤'   if ret1 = " ProgID " then
'¤¤¤'      s &= "'    not unique ProgID   - Select only 1 CoClass in tree to get the one needed to create object " & crlf & crlf
'¤¤¤'		function = s
'¤¤¤'		if xfull = 0 THEN exit function
'¤¤¤'	elseif ret1 = " No ProgID " then
'¤¤¤'		s &= "'    No ProgID   - Select 1 CoClass in tree to get the one needed to create object " & crlf & crlf
'¤¤¤'		function = s
'¤¤¤'		if xfull = 0 THEN exit function
'¤¤¤'	end if
'¤¤¤'   s &= "'" & String(80, "=") & crlf & crlf
'¤¤¤'   s &= "'   Use Prefix in AxSuite3 to differenciate the functions names / variables names" & crlf & crlf
'¤¤¤'   s &= "'================================================================================" & crlf
'¤¤¤'   s &= "Dim Shared As any ptr      " & xface & "Obj_Ptr      ' object Ptr   " & crlf
'¤¤¤'   if l1 = 1 then s &= "Dim Shared As Dword        " & xface & "Obj_Event    ' cookie for object events" & crlf & crlf & crlf
'¤¤¤'   s &= "#IfnDef _PARSED_MOD_FF3_         'Defined in FireFly to avoid duplications " & crlf & crlf
'¤¤¤'	s &= tb & "#Define Decl_Set_" & xface & " Declare Sub " & xface & "Call_Sett()" & crlf
'¤¤¤'	s &= tb & "Decl_Set_" & xface & "      ' indirect declare to avoid FireFly duplication " & crlf & crlf
'¤¤¤'   function = s
'¤¤¤'END FUNCTION

Private Function Ini_templ() as string
   dim as string s
   s = "'" & String(80, "=") & crlf
   s &= "'Template to use OCX / ActiveX / COM " & crlf
   s &= "'" & String(80, "=") & crlf & crlf
	if nowin<>0  THEN
		s &= "'#Define Ax_NoAtl        				'to use Ax_Lite.bi without atl.dll , when no control window ( reduce size of exe)"& crlf & crlf
	END IF
	s &= "#Include Once ""Windows.bi""        ' Windows specific, if not defined yet" & crlf
	s &= "#Include Once ""Ax_lite.bi""        ' this one is needed" & crlf & crlf

	if nowin=0  THEN
		s &= "'Select_ATL(71)   '   to use ATL71.dll  uncomment,  else commented use of ATL.dll" & crlf
		s &= "'                     Select_ATL instruction must be put before AxInit , after it has no effect" & crlf & crlf
		s &= "' Global COM instance initialization , only one by project" & crlf
		s &= "AxInit(True)      '   True : if at least 1 Ocx control is a Visual control (use Atl) , else False" & crlf & crlf
	else
		s &= "' Global COM instance initialization , only one by project" & crlf
		s &= "AxInit(False)     '   False : if all Ocx controls are Not-Visual controls , else True" & crlf & crlf
   END IF
   s &= "' AxStop()       'use only one by project to stop COM instance , put at the end of program" & crlf & crlf
   s &= "'" & String(80, "=") & crlf
   function = s
END FUNCTION

Private Sub EnumCoClassB (hTree As Long, hParent() As long)
   Dim Parents() As Long
   Dim hNextItem As Long
   Dim HasChilds As Long
   Dim token As String, progid As WString Ptr, clssid As clsid
   Dim As String  ret1, ret2
   Dim hItem as long
   Dim nitem as long
   Dim u1 as integer, n1 as integer,inot as integer, x as integer
   Dim s as string,iseul as string,clsids as string,progids as string,tldesc as string
	redim Parents(0)
	hItem = hparent(tkind_coclass)

   tldesc = TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem, TVGN_PARENT)))
   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
   If haschilds Then
      hnextitem = haschilds
      While hnextitem
         token = TreeView_GetItemText1(hTree, hNextItem)
			if instr(token,"          'NOT'") THEN
				token=left(token , len(token)-len("          'NOT'"))
				inot=1
			else
				inot=0
			end if
				if inot=0 THEN
					clsids &= "  ' Const CLSID_" & str_parse(token, "?", 1) & "=" & dq & str_parse(token, "?", 2) & dq & "     '" & str_parse(token, "?", 3) & crlf
				end if
				progid = Stringtobstr(str_parse(token, "?", 2))
				clsidfromString(progid, @clssid)
				sysfreeString(progid)
				progidfromclsid(@clssid, @progid)
				ret1 = *progid
				if inot=1 THEN ret1 = ""
				ret2 = ret1
				if ret1 <> "" THEN
					n1 = 0 :u1 = 0
					do
						u1 = instr(u1 + 1, ret1, ".")
						if u1 > 0 THEN
							n1 = n1 + 1
							if n1 = 2 THEN
								ret1 = left(ret1, u1 - 1)
								progids &= "  ' Const I_ProgID_" & str_parse(token, "?", 1) & "=" & dq & ret1 & dq & _
										"                 '*** Version independent ProgID" & crlf
								if nitem = 0 THEN
									iseul = ret1
								end if
								exit do
							END IF
						END IF
					LOOP UNTIL u1 = 0
				END IF
				if ret2 <> ret1 THEN
					progids &= "  ' Const ProgID_" & str_parse(token, "?", 1) & "=" & dq & *progid & dq & _
							"                                       'Version dependent ProgID" & crlf & crlf
				else
					if *progid="" THEN
					elseif ret1 = "" THEN
						if inot=0 THEN progids &= "  ' Const ProgID_" & str_parse(token, "?", 1) & "=" & dq & *progid & dq & crlf & crlf
					else
						progids &= "  ' Const ProgID_" & str_parse(token, "?", 1) & "=" & dq & *progid & dq & _
								"                 '*** Version independent ProgID" & crlf & crlf
						if nitem = 0 THEN
							iseul = *progid
						end if
					END IF
				END IF
				if inot=0 and *progid <> "" THEN nitem += 1
				sysfreeString(progid)
				hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
      Wend
   End If
	sevent =""
   if nitem = 1 then
		ret1 =  iseul
		for x= 1 to icoclass
			if coclass(x,2)= ret1 then
				sevent = coclass(x,7)
				exit for
			end if
      NEXT
   end if

End sub

Private Function EnumCoClassA (hTree As Long, hParent() As long,byRef clsids as string,byRef progids as string,Byref tldesc as string) As String
   Dim Parents() As Long
   Dim hNextItem As Long
   Dim HasChilds As Long
   Dim token As String, progid As WString Ptr, clssid As clsid
   Dim As String  ret1, ret2
   Dim hItem as long
   Dim nitem as long
   Dim u1 as integer, n1 as integer,inot as integer, x as integer
   Dim s as string,iseul as string,iseul2 as string
	redim Parents(0)
	hItem = hparent(tkind_coclass)
	clsids = ""
   tldesc = TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem, TVGN_PARENT)))
   iseul2 = tldesc
	haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
   If haschilds Then
      hnextitem = haschilds
      While hnextitem
         token = TreeView_GetItemText1(hTree, hNextItem)
			if instr(token,"          'NOT'") THEN
				token=left(token , len(token)-len("          'NOT'"))
				inot=1
			else
				inot=0
			end if
				if inot=0 THEN
					if clsids = "" THEN
						clsids =  dq & str_parse(token, "?", 2) & dq
					else
						iseul2 =""
               END IF

				end if
				progid = Stringtobstr(str_parse(token, "?", 2))
				clsidfromString(progid, @clssid)
				sysfreeString(progid)
				progidfromclsid(@clssid, @progid)
				ret1 = *progid
				if inot=1 THEN ret1 = ""
				ret2 = ret1
				if ret1 <> "" THEN
					n1 = 0 :u1 = 0
					do
						u1 = instr(u1 + 1, ret1, ".")
						if u1 > 0 THEN
							n1 = n1 + 1
							if n1 = 2 THEN
								ret1 = left(ret1, u1 - 1)
								progids &= "  ' Const I_ProgID_" & str_parse(token, "?", 1) & "=" & dq & ret1 & dq & _
										"                 '*** Version independent ProgID" & crlf
								if nitem = 0 THEN
									iseul = ret1
								end if
								exit do
							END IF
						END IF
					LOOP UNTIL u1 = 0
				END IF
				if ret2 <> ret1 THEN
					progids &= "  ' Const ProgID_" & str_parse(token, "?", 1) & "=" & dq & *progid & dq & _
							"                                       'Version dependent ProgID" & crlf & crlf
				else
					if *progid="" THEN
					elseif ret1 = "" THEN
						if inot=0 THEN progids &= "  ' Const ProgID_" & str_parse(token, "?", 1) & "=" & dq & *progid & dq & crlf & crlf
					else
						progids &= "  ' Const ProgID_" & str_parse(token, "?", 1) & "=" & dq & *progid & dq & _
								"                 '*** Version independent ProgID" & crlf & crlf
						if nitem = 0 THEN
							iseul = *progid
						end if
					END IF
				END IF
				if inot=0 and *progid <> "" THEN nitem += 1
				sysfreeString(progid)
				hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
      Wend
   End If
	sevent =""
   if nitem = 1 then
      FF_ClipboardSetText(iseul)
		ret1 =  iseul
	elseif nitem = 0 then
		FF_ClipboardSetText( "")
		ret1 = " No ProgID "
   else
      FF_ClipboardSetText( "")
      ret1 = " ProgID "
   end if
	tldesc = iseul2
	function = ret1
End Function

'¤¤¤'Private Function EnumCoClass1 cdecl Alias "EnumCoClass1"(hTree As Long, hParent() As long, l1 as integer, _
'¤¤¤'         s_conect as string, s_disc as string) As zString Ptr
'¤¤¤'	Dim As String clsids, progids, tldesc , ret1
'¤¤¤'   Dim s as string,iseul as string
'¤¤¤'   Dim xface as string, ss22 as string
'¤¤¤'	Dim x as integer
'¤¤¤'	GetPrefix()
'¤¤¤'	xface = trim(prefix, any "-_ ")
'¤¤¤'   if xface <> "" THEN xface = xface & "_"
'¤¤¤'	isevent=0
'¤¤¤'
'¤¤¤'	ret1=EnumCoClassA(hTree , hParent(),clsids, progids, tldesc )
'¤¤¤'	if nowin=0 THEN
'¤¤¤'		s = Simple_control(xface, ret1, tldesc, progids, l1)
'¤¤¤'		if (ret1 <> " ProgID " and ret1 <> " No ProgID " ) or xfull =1 THEN
'¤¤¤'			xret1=0
'¤¤¤'
'¤¤¤'			for x= 1 to icoclass
'¤¤¤'				if coclass(x,2)=ret1 THEN
'¤¤¤'					if coclass(x,7)<>"" then isevent=1
'¤¤¤'					exit for
'¤¤¤'            END IF
'¤¤¤'         NEXT
'¤¤¤'
'¤¤¤'			s &= Full_control(xface,s_conect, s_disc, l1)
'¤¤¤'		else
'¤¤¤'			xret1=1
'¤¤¤'			s="'   " & ret1
'¤¤¤'			messagebox(getactivewindow()," ProgId not defined... Select only 1 Coclass on the tree view ( with ProgId)."  ,"Error",MB_ICONERROR)
'¤¤¤'			goto ddetour
'¤¤¤'		END IF
'¤¤¤'	else
'¤¤¤'		s = Simple_nowin(xface, ret1, tldesc, progids, l1)
'¤¤¤'		if (ret1 <> " ProgID " and ret1 <> " No ProgID " ) or xfull =1 THEN
'¤¤¤'			xret1=0
'¤¤¤'			for x= 1 to icoclass
'¤¤¤'				if coclass(x,2)=ret1 THEN
'¤¤¤'					ret1 = dq & coclass(x,2) & dq & " , " & dq & coclass(x,5)& dq
'¤¤¤'					if coclass(x,7)<>"" then isevent=1
'¤¤¤'					exit for
'¤¤¤'            END IF
'¤¤¤'         NEXT
'¤¤¤'			s &= Full_nowin(xface,s_conect, s_disc, l1,ret1)
'¤¤¤'		else
'¤¤¤'			xret1=1
'¤¤¤'			s="'   " & ret1
'¤¤¤'			messagebox(getactivewindow()," ProgId not defined... Select 1 Coclass only on the tree view ( with ProgId)."  ,"Error",MB_ICONERROR)
'¤¤¤'			goto ddetour
'¤¤¤'		END IF
'¤¤¤'   END IF
'¤¤¤'ddetour:
'¤¤¤'   Return cast(zstring ptr, SysAllocStringByteLen(s, len(s)))
'¤¤¤'
'¤¤¤'End Function

Private Function EnumCoClass3 (hTree As Long, hParent() As long) As zString Ptr
	Dim As String clsids, progids, tldesc , ret1
   Dim s as string
   Dim xface as string
	Dim x as integer
	GetPrefix()
	xface = trim(prefix, any "-_ ")
   if xface <> "" THEN xface = xface & "_"

	ret1=EnumCoClassA(hTree , hParent(),clsids, progids, tldesc )
		if (clsids <> "" and tldesc <> "" ) THEN
			for x= 1 to icoclass
				if clsids = dq & coclass(x,3) & dq THEN
					s = "'" & String(80, "=") & crlf
					s &= "'   CoClass     " & coclass(x,1) &   "      " & clsids & crlf
					s &= "'   Interface   " & coclass(x,4) &   "      " & dq & coclass(x,5)& dq & crlf
					s &= "'" & String(80, "=") & crlf & crlf & crlf
					s &= "  'Dim Shared As any ptr      " & xface & "uObj_Ptr   ' unRegistered Object Ptr   " & crlf & crlf
					s &= "  'Dim Shared As Hmodule      " & xface & "hdll       ' Ocx/Dll library   " & crlf
					s &= "  '" & xface & "hdll = LoadLibrary ( """ & OcxName & """ )    ' adapt to your target path" & crlf & crlf

					if nowin=0 THEN
						s &= "  'Dim Shared As hwnd         " & xface & "hWndOcx	  ' Win handle for Ocx" & crlf & crlf & crlf
						s &= "  ' Put the following code to initialize the control where it is needed " & crlf
						s &= "  '" & xface & "hWndOcx = AxWinUnreg( Hparent , 350 , 300 , 250 , 170 ) ' define your needs for control window , Hparent (hwnd) for parent window"	& crlf & crlf
						s &= "  '" & xface & "uObj_Ptr = AxCreate_Unreg( " & xface & "hdll , " & dq & coclass(x,3) & dq & " , " & dq & coclass(x,5)& dq & " , " & xface & "hWndOcx )"& crlf & crlf
               else
						s &= "  ' Put the following code to initialize the control where it is needed " & crlf
						s &= "  '" & xface & "uObj_Ptr = AxCreate_Unreg( " & xface & "hdll , " & dq & coclass(x,3) & dq & " , " & dq & coclass(x,5)& dq & " )"& crlf & crlf
					END IF
					s &= "  ' Put the following code to Free the control after the Release command " & crlf
					s &= "  ' FreeLibrary ( " & xface & "hdll )" & crlf
					exit for
            END IF
         NEXT
			s= "'   Mini Template for unRegistered Ocx : remind, use prefix to differentiate variables " & crlf & crlf  & s
		else
			s= "'   CoClass not defined... Select only 1 Coclass on the tree view ."
			messagebox(getactivewindow()," CoClass not defined... Select only 1 Coclass on the tree view ."  ,"Error",MB_ICONERROR)
		END IF

   Return cast(zstring ptr, SysAllocStringByteLen(s, len(s)))

End Function

Private Function RegSearchWin32(tszKey AS string) AS STRING
   Dim szKeyName AS zstring * max_PATH
   Dim szKey AS zstring * max_PATH
   Dim szClass AS zstring * max_PATH
   dim ft AS FILETIME
   Dim hKey AS HKEY
   Dim dwIdx AS uinteger
   Dim hr AS uinteger
   Dim dwname as uinteger = max_path
   Dim dwclass as uinteger = max_path
   szkey = tszkey
   DO
      hr = RegOpenKeyEx(hkey_CLASSES_ROOT, @szKey, 0, KEY_READ, @hKey)
      IF hr = ERROR_NO_MORE_ITEMS THEN EXIT FUNCTION
      IF hKey = 0 THEN EXIT FUNCTION
      dwname = max_path
      dwclass = max_path

      hr = RegEnumKeyEx(hKey, dwIdx, @szKeyName, @dwname, 0, @szClass, @dwclass, @ft)
      IF hr <> 0 THEN EXIT DO
      IF UCASE$(szKeyName) = "WIN32" THEN EXIT DO
      dwIdx += 1
   LOOP WHILE hr = 0
   RegCloseKey hKey
   IF hr <> 0 OR szKeyName = "" THEN EXIT FUNCTION
   FUNCTION = szKey
END FUNCTION

Private Function RegEnumDirectory(sKey AS STRING) AS STRING
   Dim szKey AS zstring * max_PATH
   Dim szKeyName AS zstring * max_PATH
   Dim szClass AS zstring * max_PATH
   Dim ft AS FILETIME
   Dim hKey AS HKEY
   Dim dwIdx AS uinteger
   Dim hr AS uinteger
   dim sSubkey AS STRING
   Dim dwname as uinteger
   Dim dwclass as uinteger

   szKey = sKey
   hr = RegOpenKeyEx(hkey_CLASSES_ROOT, @szKey, 0, KEY_READ, @hKey)
   IF hr <> ERROR_SUCCESS THEN EXIT FUNCTION
   IF hKey = 0 THEN EXIT FUNCTION
   dwIdx = 0
   DO
      dwname = max_path
      dwclass = max_path
      hr = RegEnumKeyEx(hKey, dwIdx, @szKeyName, @dwname, 0, @szClass, @dwclass, @ft)
      IF hr = ERROR_NO_MORE_ITEMS THEN EXIT DO
      sSubkey = RegSearchWin32(szKey & "\" & szKeyName)
      IF LEN(sSubkey) THEN EXIT DO
      dwIdx += 1
   LOOP
   RegCloseKey hKey
   IF hr <> 0 OR sSubkey = "" THEN EXIT FUNCTION

   dim szValueName AS zstring * max_PATH
   Dim KeyType AS DWORD
   Dim szKeyValue AS zstring * max_PATH
   Dim cValueName AS DWORD
   Dim cbData AS DWORD

   dwIdx = 0
   cValueName = max_PATH
   cbData = max_PATH
   szKey = sSubkey & "\" & "win32"
   hr = RegOpenKeyEx(hkey_CLASSES_ROOT, @szKey, 0, KEY_READ, @hKey)
   IF hr <> ERROR_SUCCESS THEN EXIT FUNCTION
   hr = RegEnumValue(hKey, dwIdx, @szValueName, @cValueName, BYVAL NULL, @KeyType, @szKeyValue, @cbData)
   RegCloseKey hKey
   FUNCTION = szKeyValue
END FUNCTION

Private Function CheckEnvVar(BYVAL sTextIn AS STRING) AS STRING
   Dim iPos AS LONG
   dim iSafety AS LONG
   Dim sTemp AS STRING
   Dim sNew AS String
   Dim ttxt As String

   ttxt = stextin
   DO UNTIL str_TALLY(ttxt, "%") < 2
      iPos = INSTR(1, ttxt, "%")                 'get first position
      sTemp = MID$(ttxt, iPos + 1,(INSTR(iPos + 1, ttxt, "%") - (iPos + 1)))
      sNew = ENVIRON$(sTemp)
      ttxt = str_replace( "%" + sTemp + "%", sNew, ttxt)
      iSafety += 1 :IF iSafety > 5 THEN EXIT do  'if we find more than 5 vars...  something is wrong
   LOOP
   FUNCTION = ttxt
End FUNCTION

Private Function RegEnumVersions(sCLSID AS STRING) AS LONG
   Dim szKey AS zstring * max_PATH
   dim szKeyName AS zstring * max_PATH
   dim szClass AS zstring * max_PATH
   dim ft AS FILETIME
   dim hKey AS HKEY
   dim dwIdx AS DWORD
   dim hr AS DWORD
   dim idx AS LONG
   dim sPath AS STRING
   dim PathPos AS LONG
   dim sFile AS STRING
   dim i AS LONG
   dim lvi as lv_item

   dim szValueName AS zstring * max_PATH
   dim KeyType AS DWORD
   dim szKeyValue AS zstring * max_PATH
   dim tsz AS zstring * max_PATH
   dim tsz0 AS zstring * max_PATH
   dim tsz1 AS zstring * max_PATH

   dim cValueName AS DWORD
   dim cbData AS DWORD
   dim hVerKey AS HKEY
   dim verIdx AS DWORD
   dim vmax_path as uinteger
   dim dwname as uinteger
   dim dwclass as uinteger

   cValueName = max_PATH
   cbData = max_PATH
   szKey = "TypeLib\" & sCLSID

   hr = RegOpenKeyEx(hkey_CLASSES_ROOT, @szKey, 0, KEY_READ, @hKey)
   IF hr <> ERROR_SUCCESS THEN EXIT Function
   IF hKey = 0 THEN EXIT FUNCTION
   dwIdx = 0
   DO
      dwname = max_path
      dwclass = max_path
      hr = RegEnumKeyEx(hKey, dwIdx, @szKeyName, @dwname, 0, @szClass, @dwclass, @ft)
      IF hr = ERROR_NO_MORE_ITEMS THEN EXIT DO
      tsz0 = UCASE$(sCLSID)

      verIdx = 0
      cvaluename = max_path
      cbdata = max_path
      tsz = szKey & "\" & szKeyName
      hr = RegOpenKeyEx(hkey_CLASSES_ROOT, @tsz, 0, KEY_READ, @hVerKey)
      hr = RegEnumValue(hVerKey, verIdx, @szValueName, @cValueName, NULL, @KeyType, @szKeyValue, @cbData)
      RegCloseKey hVerKey
      IF szValueName = "" THEN
         tsz1 = szKeyvalue
      Else
         tsz1 = szValueName
      End IF

      sPath = RegEnumDirectory(szKey & "\" & szKeyName)
      IF INSTR(sPath, "%") THEN
         sPath = CheckEnvVar(sPath)              '-- added in July 25th, 2003
      End If
      spath = trim(spath, any "" " ")
		spath = trim(spath, dq)
      sfile = str_parse(spath, "\", str_numparse(spath, "\"))
      If InStr(sfile, ".") = 0 Then sfile = str_parse(spath, "\", str_numparse(spath, "\") - 1)
      sfile = ucase(sfile)
      if tsz0 <> "" and tsz1 <> "" and sfile <> "" then ' and spath <> "" THEN

         lvi.iItem = 0
         lvi.mask = LVIF_TEXT
         lvi.iSubItem = 0
         lvi.pszText = VARPTR(tsz0)
         SendMessage hlist, LVM_INSERTITEM, 0, cast(lParam, @lvi)

         lvi.iSubItem = 1
         lvi.pszText = VARPTR(tsz1)
         SendMessage hlist, LVM_SETITEM, 0, cast(lparam, @lvi)

         lvi.iSubItem = 2
         lvi.pszText = strptr(sfile)
         SendMessage hlist, LVM_SETITEM, 0, cast(lparam, @lvi)

         lvi.iSubItem = 3
         lvi.pszText = strptr(spath)
         SendMessage hlist, LVM_SETITEM, 0, cast(lparam, @lvi)

         nt1 += 1                                'count libs
      END IF
      dwIdx += 1											 'loop registry
   Loop
   RegCloseKey hKey
   function = 1
END FUNCTION

Private Sub RegEnumTypeLibs
   Dim szKey AS zstring*max_path
   dim szKeyName AS zstring*max_path
   dim szClass AS zstring*max_path
   dim ft AS FILETIME
   dim hKey AS hkey
   dim dwIdx AS uinteger
   dim hr AS uinteger
   dim ttxt as string
   dim dwname as dword
   dim dwclass as dword

   szKey = "TypeLib"
   hr = RegOpenKeyEx(HKEY_CLASSES_ROOT, @szKey, 0, key_read, @hKey)
   IF hr <> ERROR_SUCCESS THEN EXIT SUB
   IF hKey = 0 THEN EXIT SUB
   dwIdx = 0
   Do
      dwname = MAX_PATH :dwclass = max_path
      hr = RegEnumKeyEx(hKey, dwIdx, @szKeyName, @dwname, 0, @szClass, @dwclass, @ft)
      IF hr = ERROR_NO_MORE_ITEMS THEN EXIT DO
      dwIdx += 1
      RegEnumVersions(left(szKeyName, dwname))
   LOOP
   RegCloseKey(hKey)
End SUB

Private Function SearchTypeLibs() AS LONG
   Dim i AS LONG
   dim ttxt as string
   SendMessage hlist, LVM_DELETEALLITEMS, 0, 0
   nt1 = 0
   RegEnumTypeLibs
   SetWindowText(hmain, "AxSuite3           ....  " & str(nt1) & " registered libs")
   FF_ListView_SetSelectedItem(hlist, 0)
   function = 1
End Function

Private Function CTStr(reftype As hreftype, pti As lptypeinfo) As String
   Dim As lptypeinfo pcti
   Dim As WString Ptr iname, ihelp
   Dim hr As hresult

   hr = pTI -> lpvtbl -> GetRefTypeInfo(pti, refType, @pCTI)
   If hr Then return "UnknownCustomType"
   hr = pCTI -> lpvtbl -> GetDocumentation(pcti, - 1, @iname, 0, 0, 0)
   If hr then return "UnknownCustomType"
   Return " As " & *iname
End Function

Private Function TDStr(ptD As typedesc ptr, pTI As lptypeinfo) As String
   Dim oss As String
   Dim s2 As safearraybound Ptr
   Dim u2 As UShort Ptr
   Dim st1 As String

   If ptD -> vt = VT_PTR Then
      oss = TDStr(ptD -> lptdesc, pTI) & " Ptr"
      If oss = " As void Ptr" Then Return " As lpvoid" Else Return oss
   End If

   If ptd -> vt = VT_SAFEARRAY Then
		st1="_SafeArray" & tdstr(ptd->lptdesc, pTI)
		st1 = Str_removeany(st1,"()")
		st1 = Str_ReplaceAll(" ", "_", st1)
		oss=  st1 & " As Dword"
      Return oss
   End If

   If ptd -> vt = vt_carray Then
      oss = "("
      For i As UShort = 0 To ptd -> lpadesc -> cdims - 1
         oss &= ptd -> lpadesc -> rgbounds(i).llbound & " To "
         oss &= ptd -> lpadesc -> rgbounds(i).llbound + ptd -> lpadesc -> rgbounds(i).celements - 1 & ","
      Next
      oss = Trim(oss, ",")
      oss &= ")"
      oss &= tdstr(@ptd -> lpadesc -> tdescElem, pTI)
      Return oss
   End If

   If ptd -> vt = VT_USERDEFINED Then
      oss = CTstr(ptd -> hreftype, pTI)
      Return oss
   End If

   Select Case ptd -> vt  ' VARIANT/VARIANTARG compatible types
      Case VT_I2 :Return " As Short"
      Case VT_I4 :Return " As Integer"
      Case VT_R4 :Return " As Single"
      Case VT_R8 :Return " As Double"
      Case VT_CY :Return " As CY"
      Case VT_DATE :Return " As DATE"
      Case VT_BSTR :Return " As BSTR"
      Case VT_DISPATCH :Return " As LPDISPATCH"
      Case VT_ERROR :Return " As SCODE"
      Case VT_BOOL :Return " As BOOL"
      Case VT_VARIANT :Return " As VARIANT"
      Case VT_UNKNOWN :Return " As LPUNKNOWN"
      Case VT_UI1 :Return " As Ubyte"
      Case VT_DECIMAL :Return " As DECIMAL"
      Case VT_I1 :Return " As Byte"
      Case VT_UI2 :Return " As Ushort"
      Case VT_UI4 :Return " As Uinteger"
      Case VT_I8 :Return " As LongInt"
      Case VT_UI8 :Return " As UlongInt"
      Case VT_INT :Return " As Integer"
      Case VT_UINT :Return " As Uinteger"
      Case VT_HRESULT :Return " As HRESULT"
      Case VT_VOID :Return " As any ptr"
      Case VT_LPSTR :Return " As Zstring ptr"
      Case VT_LPWSTR :Return " As Wstring ptr"
		case else : Return " As any ptr"

   End Select
End Function

Private Function VDStr(ByVal pvd As lpVARDESC, byval pti As lptypeinfo, tk As Integer) As String
   Dim As WString Ptr iname, ihelp
   Dim hr As hresult
   Dim oss As String
   Dim vvar As variant

   If (tk = TKIND_ENUM) or (tk = TKIND_MODULE) Then tk = var_const Else tk = pvd -> varkind
   If (tk And var_const) = var_const Then oss = "Const "
   hr = pti -> lpvtbl -> GetDocumentation(pti, pvd -> memid, @iName, @ihelp, 0, 0)
   If hr Then return "UnknownName"
   oss &= *iName & " "
   oss &= tdstr(@pvd -> elemdescVar.tdesc, pti)
   If (tk and var_const) <> var_const Then Return oss
   oss &= "="
   hr = VariantChangeType(@vvar, pvd -> lpvarValue, 0, VT_BSTR)
   If hr Then
      oss &= "'''" & "?" & *ihelp
   Else
      oss &= *Cast(WString Ptr, vvar.bstrval) & "?" & *ihelp
   End If
   If tdstr(@pvd -> elemdescVar.tdesc, pti) = " As BSTR" Then
      oss = str_replace( " As BSTR", " As String", oss)
      oss = str_replace( "=", "=" & dq, oss)
      oss = str_replace( "?", dq & "?", oss)
   End If
   SysFreeString(iname)
   SysFreeString(ihelp)
   Return oss
End function

Private Sub mkConst(pti As lptypeinfo, pta As lptypeattr, tk As typekind)
   Dim As WString ptr iname, ihelp
   Dim hroot As Long, hr As Long, pvd As lpVarDesc

   pTI -> lpvtbl -> GetDocumentation(pti, - 1, @iname, @ihelp, 0, 0)
   hroot = TreeViewInsertItem(cast(long, htree), hParent(tk), *iname & "?" & *ihelp)
   sysfreestring(iname)
   SysFreeString(ihelp)
   For j As Integer = 0 To pta -> cVars - 1
      hr = pTI -> lpvtbl -> GetVarDesc(pti, j, @pvd)
      TreeViewInsertItem(cast(long, htree), hroot, VDStr(pvd, pTI, tk))
      pTI -> lpvtbl -> ReleaseVarDesc(pti, pvd)
   Next
End Sub

Private Sub mkInterface(pti As lptypeinfo, pta As lptypeattr, tk As typekind)
   Dim As WString ptr iname, ihelp, dlls, aliass, iids
   Dim As String oss, prop, sk, ok, par
   Dim hroot As Long, hr As Long
   Dim pvd As lpVarDesc, pfd As lpfuncdesc, pd As paramdesc
   Dim names() As WString Ptr, cnames As Integer, ord As Short
   Dim otli As lptypelib

   hr = pTI -> lpvtbl -> GetDocumentation(pti, - 1, @iname, @ihelp, 0, 0)
   If tk = tkind_module Then
      hroot = TreeViewInsertItem(cast(long, htree), hParent(tk), *iname & "?" & *ihelp)
   Else
      stringfromiid(@pta -> guid, @iids)
		hroot = TreeViewInsertItem(cast(long, htree), hParent(tk), *iname & "?" & *iids & "?" & *ihelp)
      SysFreeString(iids)
   End If
   sysfreestring(iname)
   SysFreeString(ihelp)
   For ifunc As Integer = 0 To pta -> cvars - 1
      pTI -> lpvtbl -> GetvarDesc(pti, ifunc, @pvd)
      pTI -> lpvtbl -> GetDocumentation(pti, ifunc, @iname, @ihelp, 0, 0)
      If tk = tkind_dispatch Then ok = Str(pvd -> memid) Else ok = Str(pvd -> oInst)
      oss = ok & "?2" & "?get" & str_parse(VDStr(pvd, pti, tk), " As ", 1) & " As Function()As " & str_parse(VDStr(pvd, pti, tk), " As ", 2) & "?" & *ihelp
      TreeViewInsertItem(cast(long, htree), hroot, oss)
      oss = ok & "?4" & "?put" & str_parse(VDStr(pvd, pti, tk), " As ", 1) & " As Sub(" & VDStr(pvd, pti, tk) & ")?" & *ihelp
      TreeViewInsertItem(cast(long, htree), hroot, oss)
      pTI -> lpvtbl -> ReleasevarDesc(pti, pvd)
      SysFreeString(iname)
      SysFreeString(ihelp)
   Next
   For ifunc As Integer = 0 To pta -> cfuncs - 1
      pTI -> lpvtbl -> GetfuncDesc(pti, ifunc, @pfd)
      pTI -> lpvtbl -> GetDocumentation(pti, pfd -> memid, @iname, @ihelp, 0, 0)
      ReDim names(pfd -> cparams)
      pti -> lpvtbl -> getnames(pti, pfd -> memid, @names(0), 1 + pfd -> cparams, @cnames)
      sysfreestring(names(0))
      If pfD -> cParams = 0 Then
         oss = "()"
      Else
         oss = "("
         For ipar As Integer = 1 To pfD -> cParams
            oss &= *names(ipar) :
            sysfreestring(names(ipar))
            par = TDStr(@pfd -> lprgelemdescParam[iPar - 1].tdesc, pti)
            pd = pfd -> lprgelemdescParam[iPar - 1].paramdesc

            If (pd.wParamFlags and PARAMFLAG_FHASDEFAULT) = PARAMFLAG_FHASDEFAULT Then
               If par = " As VARIANT" Then
                  oss &= par & "=Type(" & pd.pparamdescex -> varDefaultValue.vt & ",0,0,0," & variantv(pd.pparamdescex -> varDefaultValue) & ")"
               Else
                  oss &= par & "=" & variantv(pd.pparamdescex -> varDefaultValue)
               End If
            ElseIf (pd.wParamFlags and PARAMFLAG_FOPT) = PARAMFLAG_FOPT Then
               If par = " As VARIANT" Then oss &= par & "=Type(0,0,0,0,0)" Else oss &= par & "=0"
            Else
               oss &= par
            End If
            If ipar < pfD -> cParams Then oss &= ","
         Next
         oss &= ")"
      End If
      Select Case pfd -> invkind
         Case INVOKE_PROPERTYGET
            prop = "get"
         Case INVOKE_PROPERTYPUT
            prop = "put"
         Case INVOKE_PROPERTYPUTREF
            prop = "set"
         Case Else
            prop = ""
      End Select
      If tk = tkind_dispatch Then ok = Str(pfd -> memid) Else ok = Str(pfd -> oVft)
      If tk = tkind_module Then sk = "" Else sk = ok & "?" & pfd -> invkind & "?"
      If UCase(TDStr(@pfd -> elemdescfunc.tdesc, pti)) = " AS VOID" Then
         oss = sk & prop & *iname & " As Sub" & oss & "?" & *ihelp
      Else
         oss = sk & prop & *iname & " As Function" & oss & TDStr(@pfd -> elemdescfunc.tdesc, pti) & "?" & *ihelp
      End If
      If tk = tkind_module Then
         pti -> lpvtbl -> GetDllEntry(pti, pfd -> memid, pfd -> invkind, @dlls, @aliass, @ord)
         prop = *dlls & "?" & *aliass & "?" & ord & "?"
         oss = prop & oss
         sysfreestring(dlls)
         sysfreestring(aliass)
      End If
      TreeViewInsertItem(cast(long, htree), hroot, oss)
      SysFreeString(iname)
      SysFreeString(ihelp)
      pTI -> lpvtbl -> ReleasefuncDesc(pti, pfd)
   Next
End Sub

Private Sub fs_ENUM(ByVal i As uint)
   Dim oss As String
   Dim pti As lpTypeInfo, pta As lpTypeAttr
   pTLi -> lpvtbl -> GetTypeInfo(ptli, i, @pti)
   pTI -> lpvtbl -> GetTypeAttr(pti, @ptA)
   mkConst(pti, pta, tkind_enum)
End Sub

Private Sub fs_RECORD(ByVal i As uint)
   Dim As WString ptr iname, ihelp
   Dim attr As UShort
   Dim hroot As Long, hr As Long
   Dim pti As lpTypeInfo, pta As lpTypeAttr, pvd As lpVarDesc, pv As lpVariant

   hr = pTLi -> lpvtbl -> GetTypeInfo(ptli, i, @pti)
   hr = pTI -> lpvtbl -> GetTypeAttr(pti, @ptA)
   hr = pTI -> lpvtbl -> GetDocumentation(pti, - 1, @iname, @ihelp, 0, 0)
   hroot = TreeViewInsertItem(cast(long, htree), hParent(TKIND_RECORD), *iname & "?" & *ihelp)
   sysfreestring(iname)
   SysFreeString(ihelp)
   For j As Integer = 0 To pta -> cVars - 1
      hr = pTI -> lpvtbl -> GetVarDesc(pti, j, @pvd)
      TreeViewInsertItem(cast(long, htree), hroot, VDStr(pvd, pTI, tkind_record))
      pTI -> lpvtbl -> ReleaseVarDesc(pti, pvd)
   Next
   pTI -> lpvtbl -> ReleaseTypeAttr(pti, ptA)
End Sub

Private Sub fs_MODULE(ByVal i As uint)
   Dim hr As Long
   Dim pti As lpTypeInfo, pta As lpTypeAttr

   hr = pTLi -> lpvtbl -> GetTypeInfo(ptli, i, @pti)
   hr = pTI -> lpvtbl -> GetTypeAttr(pti, @ptA)
   If pta -> cvars Then mkConst(pti, pta, tkind_module)
   If pta -> cfuncs Then mkInterface(pti, pta, tkind_module)
End Sub

Private Sub fs_INTERFACE(ByVal i As uint)
   Dim hr As Long
   Dim pti As lpTypeInfo, pta As lpTypeAttr, hrt As HREFTYPE, prti As lpTypeInfo

   hr = pTLi -> lpvtbl -> GetTypeInfo(ptli, i, @pti)
   hr = pTI -> lpvtbl -> GetTypeAttr(pti, @ptA)
   mkInterface(pti, pta, tkind_interface)

   If (pta -> wTypeFlags And TYPEFLAG_FDUAL) = TYPEFLAG_FDUAL Then
      hr = pti -> lpvtbl -> GetRefTypeOfImplType(pti, - 1, @hrt)
      hr = pti -> lpvtbl -> GetRefTypeInfo(pti, hrt, @prti)
      pTI -> lpvtbl -> ReleaseTypeAttr(pti, ptA)
      hr = prTI -> lpvtbl -> GetTypeAttr(prti, @ptA)
      mkInterface(prti, pta, tkind_dispatch)
      prTI -> lpvtbl -> ReleaseTypeAttr(prti, ptA)
   Else
      pTI -> lpvtbl -> ReleaseTypeAttr(pti, ptA)
   End If
End Sub

Private Sub fs_DISPATCH(ByVal i As uint)
   Dim hr As Long
   Dim pti As lpTypeInfo, pta As lpTypeAttr, hrt As HREFTYPE, prti As lpTypeInfo

   hr = pTLi -> lpvtbl -> GetTypeInfo(ptli, i, @pti)
   hr = pTI -> lpvtbl -> GetTypeAttr(pti, @ptA)
   mkInterface(pti, pta, tkind_dispatch)
   If (pta -> wTypeFlags And TYPEFLAG_FDUAL) = TYPEFLAG_FDUAL Then
      hr = pti -> lpvtbl -> GetRefTypeOfImplType(pti, - 1, @hrt)
      hr = pti -> lpvtbl -> GetRefTypeInfo(pti, hrt, @prti)
      pTI -> lpvtbl -> ReleaseTypeAttr(pti, ptA)
      hr = prTI -> lpvtbl -> GetTypeAttr(prti, @ptA)
      mkInterface(prti, pta, tkind_interface)
      prTI -> lpvtbl -> ReleaseTypeAttr(prti, ptA)
   Else
      pTI -> lpvtbl -> ReleaseTypeAttr(pti, ptA)
   End If
End Sub

Private Sub fs_COCLASS(ByVal i As uint)
   Dim As WString ptr iname, ihelp
   Dim hroot As Long, hr As Long, cit As Short, litf As memberid, x As Short, prt As hreftype, s As String
   Dim pti As lpTypeInfo, pIti As lpTypeInfo, pta As lpTypeAttr, Clsids As wString Ptr, progid As WString ptr
	Dim ptB as lpTypeAttr, interf As wString Ptr ',ClassID As CLSID,wcl as LPOLESTR
	Dim as integer ievent,iclass,ipar
	Dim as string s0,s1

   pTLi -> lpvtbl -> GetTypeInfo(ptli, i, @pti)
   hr = pTI -> lpvtbl -> GetDocumentation(pti, - 1, @iname, @ihelp, 0, 0)
   hr = pTI -> lpvtbl -> GetTypeAttr(pti, @ptA)
   stringfromiid(@pta -> guid, @clsids)
   progidfromclsid(@pta -> guid, @progid)
   hroot = TreeViewInsertItem(cast(long, htree), hParent(TKIND_COCLASS), *iname & "?" & *clsids & "?" & *ihelp)
	s1= *progid
	ipar=str_numparse(s1,".")
	if ipar > 2 then
		s1= str_parse(s1,".",1) & "." & str_parse(s1,".",2)
	end if
	TreeViewInsertItem(cast(long, htree), hroot,"ProgID Version Independant >>> " & s1)
	s0=*iname
	if ubound(coclass)= icoclass then redim preserve coclass(1 to icoclass + 25 , 1 to 7) as string
	icoclass += 1
	coclass(icoclass,1)= s0
	coclass(icoclass,2)= s1
	coclass(icoclass,3)= *clsids : coclass(icoclass,3)=trim(coclass(icoclass,3),any " " & chr(0) & chr(9)& chr(10)& chr(13) )
   sysfreestring(iname)
   sysfreestring(clsids)
   sysfreestring(ihelp)
   cIT = pTa -> cImplTypes
   For x = 0 TO cIT - 1
      lITF = 0
      prt = 0
      hr = pTI -> lpvtbl -> GetImplTypeFlags(pti, x, @lITF)
      hr = pTI -> lpvtbl -> GetRefTypeOfImplType(pti, x, @prt)
      hr = pTI -> lpvtbl -> GetRefTypeInfo(pti, prt, @pITI)

      If piti Then
         hr = pITI -> lpvtbl -> GetDocumentation(piti, - 1, @iname, @ihelp, 0, 0)
			hr = pITI -> lpvtbl -> GetTypeAttr(piti, @ptB)
			stringfromiid(@ptB -> guid, @interf)
         IF lITF = 2 OR lITF = 3 Then            ' // Events interface / Default events interface
            s = *iname
            putevtlist(s)
					TreeViewInsertItem(cast(long, htree), hParent(8), s0 & " ( " & s & " )" )
					if lITF = 3 THEN
						TreeViewInsertItem(cast(long, htree), hroot, "Default Event Interface >>> " & *iname  & " >>> " & *interf )
						coclass(icoclass,6)= *iname
						coclass(icoclass,7)= *interf : coclass(icoclass,7)=trim(coclass(icoclass,7),any " " & chr(0) & chr(9)& chr(10)& chr(13))
						ievent=1
					else
						TreeViewInsertItem(cast(long, htree), hroot, "Event Interface >>> " & *iname  & " >>> " & *interf )
					END IF
			elseif lITF = 1 Then ' //  Default interface
					TreeViewInsertItem(cast(long, htree), hroot,"Default Interface >>> " & *iname  & " >>> " & *interf )
					coclass(icoclass,4)= *iname
					coclass(icoclass,5)= *interf : coclass(icoclass,5)=trim(coclass(icoclass,5),any " " & chr(0) & chr(9)& chr(10)& chr(13))
					iclass= 1
			elseif lITF = 4 Then ' //  Restricted
					TreeViewInsertItem(cast(long, htree), hroot,"Restricted Interface >>> " & *iname  & " >>> " & *interf )
			elseif lITF = 8 Then ' //  Default vTable
					TreeViewInsertItem(cast(long, htree), hroot,"Default vTable Interface >>> " & *iname  & " >>> " & *interf )
         End If
         sysfreestring(iname)
         sysfreestring(ihelp)
			sysfreestring(interf)
      End If

   Next
End Sub

Private Sub fs_ALIAS(ByVal i As uint)
   Dim As WString ptr iname, ihelp
   Dim hroot As Long, hr As Long, oss As String
   Dim pti As lpTypeInfo, pta As lpTypeAttr, pvd As lpVarDesc, pv As lpVariant

   pTLi -> lpvtbl -> GetTypeInfo(ptli, i, @pti)
   pTI -> lpvtbl -> GetTypeAttr(pti, @pta)
   hr = pTI -> lpvtbl -> GetDocumentation(pti, - 1, @iname, @ihelp, 0, 0)
   If hr then oss &= "UnknownTypedefName" Else oss = "Type " & *iname
   oss &= tdstr(@pta -> tdescAlias, pTI)
   hroot = TreeViewInsertItem(cast(long, htree), hparent(tkind_alias), oss & "     '" & *ihelp)
   pTI -> lpvtbl -> ReleaseTypeAttr(pti, pta)
   sysfreestring(iname)
   SysFreeString(ihelp)
End Sub

Private Sub fs_UNION(ByVal i As uint)
   Dim As WString ptr iname, ihelp
   Dim attr As UShort
   Dim hroot As Long, hr As Long
   Dim pti As lpTypeInfo, pta As lpTypeAttr, pvd As lpVarDesc, pv As lpVariant

   hr = pTLi -> lpvtbl -> GetTypeInfo(ptli, i, @pti)
   hr = pTI -> lpvtbl -> GetTypeAttr(pti, @ptA)
   hr = pTI -> lpvtbl -> GetDocumentation(pti, - 1, @iname, @ihelp, 0, 0)
   hroot = TreeViewInsertItem(cast(long, htree), hParent(TKIND_UNION), *iname & "?" & *ihelp)
   sysfreestring(iname)
   SysFreeString(ihelp)
   For j As Integer = 0 To pta -> cVars - 1
      hr = pTI -> lpvtbl -> GetVarDesc(pti, j, @pvd)
      TreeViewInsertItem(cast(long, htree), hroot, VDStr(pvd, pTI, tkind_union))
      pTI -> lpvtbl -> ReleaseVarDesc(pti, pvd)
   Next
   pTI -> lpvtbl -> ReleaseTypeAttr(pti, ptA)
End Sub

Private Sub LoadTLI
   Dim As Integer cti, i
   Dim tk As typekind
   sreport = ""
	icoclass=0
	dim as string act1
	redim coclass(1 to 50, 1 to 7) as string
	act1=selpath
	act1=FF_FileName (act1)
	OcxName=act1
	act1= " Selected File :  " & act1 & "    |    Path :  " & selpath
	sendmessage GetDlgItem(hmain, IDC_SBR1),WM_SETTEXT ,0,cast(lparam,strptr(act1))
   sendmessage hcode, wm_settext, 0, 0
   sendmessage getdlgitem(hmain, idc_edt1), wm_gettext, 10, cast(lparam, @prefix)
	InvalidateRect hmain, ByVal 0, True
	UpdateWindow hmain
   cti = pTLI -> lpvtbl -> gettypeinfocount(pTLI)
   If cti = 0 then exit sub
   for i = 0 To cti - 1
      pTLI -> lpvtbl -> GetTypeInfoType(pTLI, i, @tk)
      select case tk
         case TKIND_ENUM
            fs_ENUM(i)
         case TKIND_RECORD
            fs_RECORD(i)
         case TKIND_MODULE
            fs_MODULE(i)
         case TKIND_INTERFACE
            fs_INTERFACE(i)
         case TKIND_DISPATCH
            fs_DISPATCH(i)
         case TKIND_COCLASS
            fs_COCLASS(i)
         Case TKIND_ALIAS
            fs_ALIAS(i)
         case TKIND_UNION
            fs_UNION(i)
      End select
   Next
End Sub

Private Sub cleartli
   Dim hRoot As Long
   clrevtlist()
   TreeView_DeleteAllItems(htree)
   hRoot = TreeViewInsertItem(cast(long, htree), 0, tldesc & "?" & clsids)
	hParent(5) = TreeViewInsertItem(cast(long, htree), hRoot, "CoClass   : progID , CLSID , Default Interface , Default Event Interface")
	hParent(3) = TreeViewInsertItem(cast(long, htree), hRoot, "VTable : Interface")
	hParent(4) = TreeViewInsertItem(cast(long, htree), hRoot, "Invoke : Dispatch")
   hParent(0) = TreeViewInsertItem(cast(long, htree), hRoot, "Enum")
   hParent(1) = TreeViewInsertItem(cast(long, htree), hRoot, "Record")
   hParent(2) = TreeViewInsertItem(cast(long, htree), hRoot, "Module")
   hParent(6) = TreeViewInsertItem(cast(long, htree), hRoot, "Alias")
   hParent(7) = TreeViewInsertItem(cast(long, htree), hRoot, "Union")
	hParent(8) = TreeViewInsertItem(cast(long, htree), hRoot, "Events")
End Sub

Private Sub showinfo()
	sreport= "'   " &  sinfcode & crlf & crlf & sreport
   sendmessage hcode, wm_settext, 0, cast(lparam, StrPtr(sreport))
end Sub

Private Function FileFound(mes As String) As Integer
     Dim  As Integer  myHandle , fileSize, result
     If mes = "" Then Return 0
     myHandle = Freefile()
     result = Open (mes For Binary Access Read As #myHandle )
     If result <> 0 Then
        Close #myHandle
        Function = 0
        Exit Function
     End If
     fileSize = LOF(myHandle)
	  Close #myHandle
     If fileSize=0 Then
       Function = 0
		 kill mes
		 Sleep 10
       Exit Function
     End If
	  Sleep 50
     Function = 1
End Function

Private Function idem_prefix(hwin as hwnd) as integer
   dim prefix1 as zstring *11,prefix0 as zstring *11
	dim as integer i1=0
   getwindowtext getdlgitem(hwin, idc_edt1), prefix1, 10
	prefix0=prefix
	prefix1=trim(prefix1)
   if LEN(SelPath) > 0 Then
		if prefix1 = prefix0 THEN i1 = 1
	end if
	prefix=prefix1
	function= i1
END FUNCTION

Private Function FF_ClipboardSetText(ByVal Text As String) As Long
   Var hGlobalClip = GlobalAlloc(GMEM_MOVEABLE Or GMEM_SHARE, Len(Text) + 1)
   If OpenClipboard(0) Then
      EmptyClipboard()
      Var lpMem = GlobalLock(hGlobalClip)
      If lpMem Then
         CopyMemory(lpMem, StrPtr(Text), Len(text))
         GlobalUnlock(lpMem)
         Function = cast(long, SetClipboardData(CF_TEXT, hGlobalClip))
      End If
      CloseClipboard()
   else
      Function = 0
   End If
End Function

Private Sub mkInit()
	sevent=""
   sreport = Ini_templ()
   showinfo
END SUB

Private Sub mkvident
   Dim b As zString Ptr
	sevent=""
	sreport= Ini_templ()
   b = EnumCoClass1(cast(long, htree), hparent(), 0, "", "")
   sreport &=  *b :sysfreestring(cast(BSTR, b))
	if xret1=0 then
	end if
	xret1=0
	xfull=0
   showinfo
End Sub

Private Sub mkuOcx2
   Dim b As zString Ptr
	sevent=""
   b = EnumCoClass3(cast(long, htree), hparent())
   sreport =  *b :sysfreestring(cast(BSTR, b))
	xret1=0
	xfull=0
   showinfo
End Sub

Private Sub mkConstant
   Dim b As zString Ptr
	sevent=""
   b = EnumEnum(cast(long, htree), hparent())
   sreport = *b :sysfreestring(cast(BSTR, b))
   b = EnumUnion(cast(long, htree), hparent())
   sreport &= *b :sysfreestring(cast(BSTR, b))
   b = EnumAlias(cast(long, htree), hparent())
   sreport &= *b :sysfreestring(cast(BSTR, b))
   b = EnumRecord(cast(long, htree), hparent())
   sreport &= *b :sysfreestring(cast(BSTR, b))
   showinfo
End Sub

Private Function verif(s1 as string) as string
   dim as integer i1, i2
	notsure = 0
   i1 = instr(s1, "_Events_Disconnect")
   if i1 = 0 THEN
      function = ""
      exit function
   END IF
   i2 = InStrRev(s1, "Function ", i1)
   if i2 = 0 THEN
      function = ""
      exit function
   END IF
   if instr(i1 + 18, s1, "_Events_Disconnect") THEN
      function =  mid(s1, i2 + 9, i1 - (i2 + 9))
		notsure = 1
      exit function
   END IF
   function = mid(s1, i2 + 9, i1 - (i2 + 9))
END Function

Private Sub mkFull
   Dim b As zString Ptr
   dim as string s1, s2, s3
	sevent=""
	EnumCoClassB (cast(long, htree), hparent())
	sreport= Ini_templ()
	b = EnumEvent(cast(long, htree), hparent())
	s1 = *b :sysfreestring(cast(BSTR, b))
	s2 = verif(s1)

   if s1 = "" or s2 = "" THEN
      s1 = ""
      b = EnumCoClass1(cast(long, htree), hparent(), 0, "", "")
   else
		if xret2=0 then
			b = EnumCoClass1(cast(long, htree), hparent(), 1 , s2 & "_Events_Connect", s2 & "_Events_Disconnect")
		else
			b = EnumCoClass1(cast(long, htree), hparent(), 0, "", "")
		end if
	END IF
   sreport &=  *b :sysfreestring(cast(BSTR, b))
	if xret1=0 then
		sreport &= s3
		if s2<>"" and xret2=0 then
			sreport &= s1
		elseif s2<>"" then
			sreport &= crlf & "'  Events are not shown, you must select the only Interface : CoClass you need"	& crlf & crlf
		end if
	end if
	xret2=0
	xret1=0
   showinfo
End sub

Private Sub mkModule
   Dim b As zString Ptr
	sevent=""
   b = EnumModule(cast(long, htree), hparent())
   sreport = *b :sysfreestring(cast(BSTR, b))
   showinfo
end sub

Private Sub kDisp
   Dim b As zString Ptr
	sevent=""
	sreport= Ini_templ()
   b = EnumCoClass1(cast(long, htree), hparent(), 0, "", "")
   sreport &= *b :sysfreestring(cast(BSTR, b))
	if xret1=0 then
	end if
	xret1=0
	xfull=0
   showinfo
End Sub

Private Sub kDisp2
   Dim b As zString Ptr
   dim as string s1, s2, s3
	sevent=""
	EnumCoClassB (cast(long, htree), hparent())
	sreport= Ini_templ()
	b = EnumEvent(cast(long, htree), hparent())
	s1 = *b :sysfreestring(cast(BSTR, b))
	s2 = verif(s1)

   if s1 = "" or s2 = "" THEN
      s1 = ""
      b = EnumCoClass1(cast(long, htree), hparent(), 0, "", "")
   else
		if xret2=0 then
			b = EnumCoClass1(cast(long, htree), hparent(), 1 , s2 & "_Events_Connect", s2 & "_Events_Disconnect")
		else
			b = EnumCoClass1(cast(long, htree), hparent(), 0, "", "")
		end if
	END IF
   sreport &=  *b :sysfreestring(cast(BSTR, b))
	if xret1=0 then
		sreport &= s3
		if s2<>"" and xret2=0 then
			sreport &= s1
		elseif s2<>"" then
			sreport &= crlf & "'  Events are not shown, you must select the only Interface : CoClass you need"	& crlf & crlf
		end if
	end if
	xret2=0
	xret1=0
   showinfo
End sub

Private Sub mkvtable
   Dim b As zString Ptr
	sevent=""
   b = EnumCoClass(cast(long, htree), hparent())
   sreport = *b :sysfreestring(cast(BSTR, b))
   b = EnumvTable(cast(long, htree), hparent())
   sreport &= *b :sysfreestring(cast(BSTR, b))
   showinfo
End Sub

Private Sub mkInvoke
   Dim b As zString Ptr
	sevent=""
   b = EnumCoClass(cast(long, htree), hparent())
   sreport = *b :sysfreestring(cast(BSTR, b))
   b = EnumInvoke(cast(long, htree), hparent())
   sreport &= *b :sysfreestring(cast(BSTR, b))
   showinfo
End Sub

Private Sub mkEvent
   Dim b As zString Ptr
	sevent=""
   b = EnumEvent(cast(long, htree), hparent())
   sreport = *b :sysfreestring(cast(BSTR, b))
   showinfo
End Sub

Private Sub dlgMain_IDC_TAB1_Sample(byval hWin as hwnd, byval CTLID as integer)
   dim hCTLID as hwnd
   dim ts as TCITEM
   hCTLID = GetDlgItem(hWin, CTLID)
   ts.mask = TCIF_TEXT
   ts.pszText = StrPtr( "Registered")
   SendMessage(hCTLID, TCM_INSERTITEM, 0, cast(lparam, @ts))
   ts.pszText = StrPtr( "Tree ")
   SendMessage(hCTLID, TCM_INSERTITEM, 1, cast(lparam, @ts))
   ts.pszText = StrPtr( "Code ")
   SendMessage(hCTLID, TCM_INSERTITEM, 2, cast(lparam, @ts))
   hTab0 = CreateDialogParam(hinstance, cast(LPCTSTR, dlgList), hCTLID, @dlgList_DlgProc, 0)
   hTab1 = CreateDialogParam(hinstance, cast(LPCTSTR, dlgTree), hCTLID, @dlgTree_DlgProc, 0)
   hTab2 = CreateDialogParam(hinstance, cast(LPCTSTR, dlgCode), hCTLID, @dlgCode_dlgProc, 0)
   ShowWindow(hTab0, SW_SHOW)
   ShowWindow(hTab1, SW_HIDE)
   ShowWindow(hTab2, SW_HIDE)
end sub

Private Function dlgMain_Init(Byval hWin as hwnd, Byval wParam as wParam, Byval lParam as lParam) as Integer
   Dim hicon As Integer
   hinstance = cast(hinstance, getwindowlong(hwin, gwl_hInstance))
   hmain = hwin
   hIcon = cast(integer, LoadIcon(hinstance, cast(LPCTSTR, 100)))
   SendMessage hwin, WM_SETICON, ICON_SMALL, hIcon
   dlgMain_IDC_TAB1_Sample(hWin, IDC_TAB1)
   TabCtrl_GetItemRect(getdlgitem(hwin, idc_tab1), 2, @rtab)
   movewindow(getdlgitem(hwin, idc_stc3), rtab.right + 25, rtab.top, 40, 18, TRUE)
   movewindow(getdlgitem(hwin, idc_edt1), rtab.right + 68, rtab.top, 100, 18, TRUE)
   SearchTypeLibs
   listview_setcolumnwidth(hlist, 1, LVSCW_AUTOSIZE)

   function = 0
End Function

Private Function dlgMain_OnSize(Byval hWin as hwnd, Byval lState as long, Byval cx as long, Byval cy as long) as Integer
   movewindow(getdlgitem(hwin, idc_tab1), 0, 0, cx, cy, TRUE)
   movewindow(htab0, 0, 22, cx, cy - 43, TRUE)
   movewindow(htab1, 0, 22, cx, cy - 43, TRUE)
   movewindow(htab2, 0, 22, cx, cy - 43, TRUE)
	movewindow(getdlgitem(hwin, IDC_SBR1), 6 , cy-12, cx-20, 12, TRUE)
   function = 0
End Function

Private Function dlgMain_OnClose(Byval hWin as hwnd) as Integer
   function = 0
End Function

Private Function dlgMain_OnDestroy(Byval hWin as hwnd) as Integer
   function = 0
End Function

Private Function dlgMain_OnNotify(Byval hWin as hwnd, Byval lCtlID as long, Byval lpNMHDR as NMHDR ptr, Byref lresult as long) as Integer
   function = 0
End Function

Private Function dlgMain_OnTimer(Byval hWin as hwnd, Byval wTimerID as word) as Integer
   function = 0
End Function

Private Function dlgMain_IDC_TAB1_SelChange(byval hWin as hwnd, byval lpNMHDR as NMHDR ptr, Byref lresult as long) as long
   dim TabID as Long
   dim hIDC_TAB1 as hwnd
   hIDC_TAB1 = getdlgitem(hWin, IDC_TAB1)
   TabID = SendMessage(hIDC_TAB1, TCM_GETCURSEL, 0, 0)
   select case TabID
      case 0
         ShowWindow(hTab0, SW_SHOW)
         ShowWindow(hTab1, SW_HIDE)
         ShowWindow(hTab2, SW_HIDE)
			ShowWindow(GetDlgItem(hWin, IDC_STC31), SW_SHOW)
         ShowWindow(GetDlgItem(hWin, IDC_EDT31), SW_SHOW)
         setfocus hlist
			exit function
      Case 1
         ShowWindow(hTab0, SW_HIDE)
         ShowWindow(hTab1, SW_SHOW)
         ShowWindow(hTab2, SW_HIDE)

			ShowWindow(GetDlgItem(hWin, IDC_STC31), SW_HIDE)
         ShowWindow(GetDlgItem(hWin, IDC_EDT31), SW_HIDE)
         setfocus htree
			exit function
      Case 2
         ShowWindow(hTab0, SW_HIDE)
         ShowWindow(hTab1, SW_HIDE)
         ShowWindow(hTab2, SW_SHOW)
			ShowWindow(GetDlgItem(hWin, IDC_STC31), SW_HIDE)
         ShowWindow(GetDlgItem(hWin, IDC_EDT31), SW_HIDE)

         setfocus hcode
			exit function
   end select
   function = 0
End function

Private Function dlgMain_IDC_TAB1_SelChanging(byval hWin as hwnd, byval lpNMHDR as NMHDR ptr, Byref lresult as long) as long
   dim TabID as Long
   dim hIDC_TAB1 as hwnd
   hIDC_TAB1 = getdlgitem(hWin, IDC_TAB1)
   TabID = SendMessage(hIDC_TAB1, TCM_GETCURSEL, 0, 0)
   select case TabID
      case 0
      case 1
   end select
   function = 0
End Function

Private Sub chgtab(ByVal p As Integer)
   Dim nmh As nmhdr
   nmh.hwndfrom = getdlgitem(hmain, idc_tab1)
   nmh.idfrom = idc_tab1
   nmh.code = TCN_SELCHANGE
   sendmessage getdlgitem(hmain, idc_tab1), TCM_SETCURSEL, p, 0
   sendmessage getdlgitem(hmain, idc_tab1), WM_NOTIFY, idc_tab1, cast(lparam, @nmh)
   If p = 1 Then
      treeview_setcheckstate(cast(long, htree), cast(long, TreeView_GetChild(hTree, 0)), 1)
      treeview_changechildstate(cast(long, htree), cast(long, TreeView_GetChild(hTree, 0)), 1)
   End If
End Sub

Private Sub chgtab1()
   treeview_setcheckstate(cast(long, htree), cast(long, TreeView_GetChild(hTree, 0)), 1)
   treeview_changechildstate(cast(long, htree), cast(long, TreeView_GetChild(hTree, 0)), 1)
End Sub

Private Function dlgMain_IDR_MENU1_menu(byval hwin as hwnd, byval id as long) as Long
   dim as OPENFILENAME ofn
   dim as zstring*max_path pathfilename
	dim as zstring * max_path fstatic1,fstatic2
	dim as string fsave
   Dim dllpath As String
   dim as string str1
	dim as string sitem,Newvar
   dim as integer r1

   Select Case id
		Case regf
			ofn.lStructSize = sizeof(OPENFILENAME)
         ofn.hwndOwner = hwin
         ofn.hInstance = cast(hinstance, GetWindowLong(hwin, GWL_HINSTANCE))
         ofn.lpstrFile = @pathFileName
         ofn.nMaxFile = 260
         ofn.lpstrFilter = @szOpen
         ofn.Flags = OFN_FILEMUSTEXIST or OFN_HIDEREADONLY or OFN_PATHMUSTEXIST
         if GetOpenFileName(@ofn) Then sitem = pathfilename
         IF LEN(sitem) Then
				r1=RegisterServer(hwin, sitem, 1)
            if r1 <> 0 then
					MessageBox getactiveWindow(), "Unable to Register " & crlf & crlf  & sitem, _
                     "Error", MB_OK OR MB_ICONERROR OR MB_TASKMODAL
				else
					MessageBox getactiveWindow(), "Registered !  " & crlf & crlf & sitem, _
                     "Done", MB_OK OR MB_ICONINFORMATION OR MB_TASKMODAL
					SearchTypeLibs
					searching = str_parse(sitem, "\", str_numparse(sitem, "\"))
					SetWindowText GetDlgItem(hWin, IDC_EDT31),searching
					if  searching <> "" THEN search1()
					setfocus GetDlgItem(hWin, IDC_EDT31)
				End If
         End If
			exit function
		Case unregl
			chgtab(0)
			dim as integer irow = FF_ListView_GetSelectedItem(hList)
			if irow > - 1 THEN
				sitem = FF_ListView_GetItemText(hList, irow, 3)
				IF LEN(sitem) Then
					r1=MessageBox (getactiveWindow(), "Are you sure, you want to unregister : " & crlf & crlf  _
							& sitem, "Confirmation", MB_YESNO OR MB_ICONWARNING OR MB_TASKMODAL	)
					if r1= IDYES THEN
						r1=RegisterServer(hwin, sitem, 0)
						if r1 <> 0 then
							MessageBox getactiveWindow(), "Unable to Unregister " & crlf & crlf  & sitem, _
									"Error", MB_OK OR MB_ICONERROR OR MB_TASKMODAL
						else
							MessageBox getactiveWindow(), "Unregistered !  " & crlf & crlf & sitem, _
									"Done", MB_OK OR MB_ICONINFORMATION OR MB_TASKMODAL
							cleartli
							selpath=""
							SearchTypeLibs
							searching=""
							sendmessage GetDlgItem(hmain, IDC_SBR1),WM_SETTEXT ,0,cast(lparam,strptr("Nothing selected"))
							SetWindowText GetDlgItem(hWin, IDC_EDT31), ""
							setfocus GetDlgItem(hWin, IDC_EDT31)
						End If
					end if
				End If
			else
				MessageBox getactiveWindow(), "Nothing selected " & crlf & crlf  & "Selec a line, please" , _
									"Error", MB_OK OR MB_ICONERROR OR MB_TASKMODAL
			End If
			exit function
      case tlb_reged
         SearchTypeLibs
			exit function
      Case tlb_libfile
         ofn.lStructSize = sizeof(OPENFILENAME)
         ofn.hwndOwner = hwin
         ofn.hInstance = cast(hinstance, GetWindowLong(hwin, GWL_HINSTANCE))
         ofn.lpstrFile = @pathFileName
         ofn.nMaxFile = 260
         ofn.lpstrFilter = @szOpen
         ofn.Flags = OFN_FILEMUSTEXIST or OFN_HIDEREADONLY or OFN_PATHMUSTEXIST
         if GetOpenFileName(@ofn) Then selpath = pathfilename
         IF LEN(SelPath) Then
            clsids = ""
            tldesc = ""
            cleartli
            IF LoadtypeLib(stringtobstr(selpath), @ptli) Then
               MessageBox BYVAL NULL, "Unable to load " & SelPath, _
                     "Error", MB_OK OR MB_ICONERROR OR MB_TASKMODAL
            Else
               LoadTLI

               chgtab(1)
               treeview_select(htree, hparent(0), TVGN_CARET)
            End If
         End If
			exit function
      case tlb_exit
         EndDialog(hWin, 0)
			exit function
      Case code_init
         savepath = "Init"
			sinfcode = "Init Code"
         mkInit()
         chgtab(2)
         gene1 = code_init
			exit function
		Case uOcx1 '10020
			iscod=0
         if idem_prefix(hWin) = 0 then
            r1 = force_click(hWin)
            if r1 = 0 THEN exit Function
         END IF
         savepath = "Unreg_Win"
			sinfcode = "Unregistered Visual Control"
			nowin=0
			mkuOcx2
			chgtab(2)
         gene1 = uOcx1
			exit function
		Case uOcx2 '10021
			iscod=0
         if idem_prefix(hWin) = 0 then
            r1 = force_click(hWin)
            if r1 = 0 THEN exit Function
         END IF
         savepath = "Unreg_NoWin"
			sinfcode = "Unregistered Not-Visual Control"
			nowin=1
			mkuOcx2
			chgtab(2)
         gene1 = uOcx2
			exit function
      Case code_ident
			iscod=0
         if idem_prefix(hWin) = 0 then
            r1 = force_click(hWin)
            if r1 = 0 THEN exit Function
         END IF
         savepath = "Win_noEv"
			sinfcode = "Visual Control"
			nowin=0
         mkvident
         chgtab(2)
         gene1 = code_ident
			exit function
      case code_const
			iscod=0
         if idem_prefix(hWin) = 0 then
            r1 = force_click(hWin)
            if r1 = 0 THEN exit Function
         END IF
         savepath = "Constant"
			sinfcode = "Constants Code"
         mkconstant
         chgtab(2)
         gene1 = code_const
			exit function
		Case code_Disp
			iscod=0
			if idem_prefix(hWin) = 0 then
            r1 = force_click(hWin)
            if r1 = 0 THEN exit Function
         END IF
         savepath = "noWin_noEv"
			sinfcode = "Not-Visual Control"
			nowin=1
         kDisp
         chgtab(2)
         gene1 = code_Disp
			exit function
		Case code_Disp2
			iscod=0
			if idem_prefix(hWin) = 0 then
            r1 = force_click(hWin)
            if r1 = 0 THEN exit Function
         END IF
         savepath = "noWin_Ev"
			sinfcode = "Not-Visual Control + Events"
			nowin=1
         kDisp2
         chgtab(2)
         gene1 = code_Disp2
			exit function
      Case code_full
			iscod=0
         if idem_prefix(hWin) = 0 then
            r1 = force_click(hWin)
            if r1 = 0 THEN exit Function
         END IF
         savepath = "Win_Ev"
			sinfcode = "Visual Control + Events"
			nowin=0
         mkFull
         chgtab(2)
         gene1 = code_full
			exit function
      Case code_Module
			iscod=0
         if idem_prefix(hWin) = 0 then
            r1 = force_click(hWin)
            if r1 = 0 THEN exit Function
         END IF
         savepath = "Module"
			sinfcode = "Module Code"
         mkModule
         chgtab(2)
         gene1 = code_Module
			exit function
      Case code_vTable
			iscod=0
         if idem_prefix(hWin) = 0 then
            r1 = force_click(hWin)
            if r1 = 0 THEN exit Function
         END IF
         savepath = "vTable"
			sinfcode = "vTable Code"
         mkvtable
         chgtab(2)
         gene1 = code_vTable
			exit function
      case code_Invoke
			iscod=0
         if idem_prefix(hWin) = 0 then
            r1 = force_click(hWin)
            if r1 = 0 THEN exit Function
         END IF
         savepath = "Invoke"
			sinfcode = "Invoke Code"
         mkInvoke
         chgtab(2)
         gene1 = code_Invoke
			exit function
      case code_event
			iscod=0
         if idem_prefix(hWin) = 0 then
            r1 = force_click(hWin)
            if r1 = 0 THEN exit Function
         END IF
         savepath = "Event"
			sinfcode = "Events Code"
         mkEvent
         chgtab(2)
         gene1 = code_event
			exit function
      Case code_save
         if sreport <> "" then
            str1 = trim(prefix, any "-_ ")
            if str1 <> "" THEN
               str1 = trim(str1, any "-_ ")
               if str1 <> "" THEN str1 = str1 & "_"
            END IF
            pathfilename = str1 & str_parse(str_parse(selpath, "\", str_numparse(selpath, "\")), ".", 1)
            pathfilename &= "_" & savepath & ".bi"
            ofn.lStructSize = sizeof(OPENFILENAME)
            ofn.hwndOwner = hwin
            ofn.hInstance = cast(hinstance, GetWindowLong(hwin, GWL_HINSTANCE))
            ofn.lpstrFile = @pathFileName
            ofn.nMaxFile = 260
            ofn.lpstrFilter = @szSave
            ofn.Flags = OFN_FILEMUSTEXIST or OFN_HIDEREADONLY or OFN_PATHMUSTEXIST
            if GetSaveFileName(@ofn) Then savepath = pathfilename
            IF LEN(SavePath) Then
               Open savepath for Binary As #1
               Put #1,, sreport
               Close #1
            End If
         End if
			exit function
		Case Static5
			fstatic1="Ax_Lite.bi"
			fstatic2="Include file   Ax_Lite" & szNULL & "*.bi" & szNULL
			ofn.lStructSize = sizeof(OPENFILENAME)
			ofn.hwndOwner = hwin
			ofn.hInstance = cast(hinstance, GetWindowLong(hwin, GWL_HINSTANCE))
			ofn.lpstrFile = @fstatic1
			ofn.nMaxFile = 260
			ofn.lpstrFilter = @fstatic2
			ofn.Flags = OFN_FILEMUSTEXIST or OFN_HIDEREADONLY or OFN_PATHMUSTEXIST
			if GetSaveFileName(@ofn) Then fsave=fstatic1
			IF LEN(fsave) Then
				MakeFileImport_(inc9, inc9_len, fsave)
			End If
			exit function
		Case Static6
			mem_ClipboardSetText(inc9, inc9_len)
			exit function

		Case codll
			fstatic1="ATL.DLL"
			fstatic2="Atl Dll file " & szNULL & "*.dll" & szNULL
			ofn.lStructSize = sizeof(OPENFILENAME)
			ofn.hwndOwner = hwin
			ofn.hInstance = cast(hinstance, GetWindowLong(hwin, GWL_HINSTANCE))
			ofn.lpstrFile = @fstatic1
			ofn.nMaxFile = 260
			ofn.lpstrFilter = @fstatic2
			ofn.Flags = OFN_FILEMUSTEXIST or OFN_HIDEREADONLY or OFN_PATHMUSTEXIST
			if GetSaveFileName(@ofn) Then fsave=fstatic1
			IF LEN(fsave) Then
				MakeFileImport_(inc1, inc1_len, fsave)
			End If
			exit function
		Case codll71
			fstatic1="ATL71.DLL"
			fstatic2="Atl71 Dll file " & szNULL & "*.dll" & szNULL
			ofn.lStructSize = sizeof(OPENFILENAME)
			ofn.hwndOwner = hwin
			ofn.hInstance = cast(hinstance, GetWindowLong(hwin, GWL_HINSTANCE))
			ofn.lpstrFile = @fstatic1
			ofn.nMaxFile = 260
			ofn.lpstrFilter = @fstatic2
			ofn.Flags = OFN_FILEMUSTEXIST or OFN_HIDEREADONLY or OFN_PATHMUSTEXIST
			if GetSaveFileName(@ofn) Then fsave=fstatic1
			IF LEN(fsave) Then
				MakeFileImport_(inc2, inc2_len, fsave)
			End If
			exit function
		Case aodll
			fstatic1="atlCtl.dll"
			fstatic2="Atl FbEdit Addin " & szNULL & "*.dll" & szNULL
			ofn.lStructSize = sizeof(OPENFILENAME)
			ofn.hwndOwner = hwin
			ofn.hInstance = cast(hinstance, GetWindowLong(hwin, GWL_HINSTANCE))
			ofn.lpstrFile = @fstatic1
			ofn.nMaxFile = 260
			ofn.lpstrFilter = @fstatic2
			ofn.Flags = OFN_FILEMUSTEXIST or OFN_HIDEREADONLY or OFN_PATHMUSTEXIST
			if GetSaveFileName(@ofn) Then fsave=fstatic1
			IF LEN(fsave) Then
				MakeFileImport_(inc5, inc5_len, fsave)
			End If
			exit function
		Case aodll71
			fstatic1="atl71Ctl.dll"
			fstatic2="Atl71 FbEdit Addin " & szNULL & "*.dll" & szNULL
			ofn.lStructSize = sizeof(OPENFILENAME)
			ofn.hwndOwner = hwin
			ofn.hInstance = cast(hinstance, GetWindowLong(hwin, GWL_HINSTANCE))
			ofn.lpstrFile = @fstatic1
			ofn.nMaxFile = 260
			ofn.lpstrFilter = @fstatic2
			ofn.Flags = OFN_FILEMUSTEXIST or OFN_HIDEREADONLY or OFN_PATHMUSTEXIST
			if GetSaveFileName(@ofn) Then fsave=fstatic1
			IF LEN(fsave) Then
				MakeFileImport_(inc6, inc6_len, fsave)
			End If
			exit function
      Case code_codegen
         if sreport <> "" then
            FF_ClipboardSetText(sreport)
         End if
			exit function
      case hlp_axsuite
         pathfilename = ExePath & "\axsuite3.chm"
			if FileFound(pathfilename)= 0 then
				messagebox (getactivewindow(), "File : AxSuite3.Chm not found ", "Error",MB_ICONERROR or MB_SYSTEMMODAL)
			else
				shellexecute 0, strptr( "Open"), @pathfilename, NULL, NULL, SW_SHOWNORMAL
			end if
			exit function
      case hlp_About
         DialogBoxParam(GetModuleHandle(NULL), cast(LPCTSTR, dlgAbout), hwin, @dlgAbout_DlgProc, NULL)
			exit function
      case As_AxSuite
         pathfilename = ExePath & "\axsuite3.pdf"
			if FileFound(pathfilename)= 0 then
				messagebox (getactivewindow(), "File : AxSuite3.Pdf not found ", "Error",MB_ICONERROR or MB_SYSTEMMODAL)
			else
				shellexecute 0, strptr( "Open"), @pathfilename, NULL, NULL, SW_SHOWNORMAL
			end if
			exit function
   End Select
   function = 0
End Function

Private Function FF_ListView_GetSelectedItem(ByVal hWndControl As hWnd) As Integer
   If IsWindow(hWndControl) Then
      Function = SendMessage(hWndControl, LVM_GETSELECTIONMARK, 0, 0)
   End If

End Function

Private Function FF_ListView_GetItemText(ByVal hWndControl As hWnd, ByVal iRow As Integer, ByVal iColumn As Integer ) As String
   Dim tlv_item As LV_ITEM
   Dim zText As ZString * MAX_PATH
   Dim Text1 as string
   If IsWindow(hWndControl) Then
      tlv_item.mask = LVIF_TEXT
      tlv_item.iItem = iRow
      tlv_item.iSubItem = iColumn
      tlv_item.pszText = VarPtr(zText)
      tlv_item.cchTextMax = SizeOf(zText)

      If SendMessage(hWndControl, LVM_GETITEM, 0, Cast(LPARAM, VarPtr(tlv_item))) Then
         Text1 = zText
         Function = Text1
      End If
   End If

End Function

Private Function FF_ListView_SetSelectedItem(ByVal hWndControl As hWnd, ByVal nIndex As Integer ) As Integer

   If IsWindow(hWndControl) Then
      Function = SendMessage(hWndControl, LVM_SETSELECTIONMARK, 0, nIndex)

      Dim tItem As LVITEM
      tItem.mask = LVIF_STATE
      tItem.iItem = nIndex
      tItem.iSubItem = 0
      tItem.State = LVIS_FOCUSED Or LVIS_SELECTED
      tItem.statemask = LVIS_FOCUSED Or LVIS_SELECTED
      SendMessage hWndControl, LVM_SETITEMSTATE, nIndex, Cast(LPARAM, VarPtr(tItem))
   End If

End Function

Private Function FF_ListView_GetTopIndex(ByVal hWndControl As hWnd) As Integer

   If IsWindow(hWndControl) Then

      Function = SendMessage(hWndControl, LVM_GETTOPINDEX, 0, 0)

   End If

End Function

Private Sub FF_Control_Redraw(ByVal hWndControl As hWnd)

   If IsWindow(hWndControl) Then

      InvalidateRect hWndControl, ByVal 0, True
      UpdateWindow hWndControl

   End If
End Sub

Private Sub show_line(ret1 As integer)
	dim as integer nScroll,pos1,i
   FF_ListView_SetSelectedItem(HWND_FORM1_LISTVIEW1, ret1)
	pos1 = FF_ListView_GetTopIndex(HWND_FORM1_LISTVIEW1)

	nScroll = pos1 - ret1
	If nScroll > 0 Then
		For i = 1 To nScroll
			SendMessage HWND_FORM1_LISTVIEW1, WM_VSCROLL, SB_LINEUP, 0
		Next
	Else
		For i = 1 To - nScroll
			SendMessage HWND_FORM1_LISTVIEW1, WM_VSCROLL, SB_LINEDOWN, 0
		Next
	End If

	SendMessage HWND_FORM1_LISTVIEW1, WM_SETREDRAW, TRUE, 0
	FF_Control_Redraw HWND_FORM1_LISTVIEW1

END sub

Private Function getsearch(str1 As String, ini As integer) As integer
   dim x As integer
   dim y As integer
   dim as string mess1, mess2
   mess2 = ucase(str1)
   For x = ini To nt1 - 1
      For y = 0 To 3
         mess1 = UCase(FF_ListView_GetItemText(HWND_FORM1_LISTVIEW1, x, y))
         If Instr(mess1, mess2) > 0 Then
            Function = x
            Exit Function
         End If
      Next
   Next
   If ini = 0 Then
      messagebox 0, "Search not found  ! " & Chr(13) & Chr(13) & str1, "Information", MB_ICONINFORMATION
   Else
		if ini >1 then
			show_line(ini -1)
		else
			show_line(0)
		end if
      messagebox 0, "No more occurrence ! " & Chr(13) & Chr(13) & str1, "Information", MB_ICONINFORMATION
   End If
   Function = - 1
End Function

Private Sub search1()

   Dim ret1 As integer, i As integer
   Dim pos1 As integer
   Dim nScroll As integer
   SelPath = ""
   pos0 = 0

   If searching <> "" Then
      ret1 = getsearch(searching, 0)
      If ret1 > - 1 Then
         pos0 = ret1
         FF_ListView_SetSelectedItem(HWND_FORM1_LISTVIEW1, ret1)
         pos1 = FF_ListView_GetTopIndex(HWND_FORM1_LISTVIEW1)

         nScroll = pos1 - ret1
         If nScroll > 0 Then
            For i = 1 To nScroll
               SendMessage HWND_FORM1_LISTVIEW1, WM_VSCROLL, SB_LINEUP, 0
            Next
         Else
            For i = 1 To - nScroll
               SendMessage HWND_FORM1_LISTVIEW1, WM_VSCROLL, SB_LINEDOWN, 0
            Next
         End If

         SendMessage HWND_FORM1_LISTVIEW1, WM_SETREDRAW, TRUE, 0
         FF_Control_Redraw HWND_FORM1_LISTVIEW1

      End If
   End If
   SendMessage getdlgitem(hmain, idc_edt31), EM_SETSEL, 0, - 1
   setfocus getdlgitem(hmain, idc_edt31)
End sub

Private Sub next1()

   dim ret1 As integer, i As integer
   dim pos1 As integer
   dim nScroll As integer
   SelPath = ""
   If searching <> "" And pos0 < nt1 - 1 Then
      ret1 = getsearch(searching, pos0 + 1)
      If ret1 > - 1 Then
         pos0 = ret1
         FF_ListView_SetSelectedItem(HWND_FORM1_LISTVIEW1, ret1)
         pos1 = FF_ListView_GetTopIndex(HWND_FORM1_LISTVIEW1)

         nScroll = pos1 - ret1
         If nScroll > 0 Then
            For i = 1 To nScroll
               SendMessage HWND_FORM1_LISTVIEW1, WM_VSCROLL, SB_LINEUP, 0
            Next
         Else
            For i = 1 To - nScroll
               SendMessage HWND_FORM1_LISTVIEW1, WM_VSCROLL, SB_LINEDOWN, 0
            Next
         End If

         SendMessage HWND_FORM1_LISTVIEW1, WM_SETREDRAW, TRUE, 0
         FF_Control_Redraw HWND_FORM1_LISTVIEW1
      End If
   elseIf searching <> "" And pos0 = nt1 - 1 Then

		messagebox 0, "No more occurrence ! " & Chr(13) & Chr(13) & searching, "Information", MB_ICONINFORMATION
   elseIf searching = "" Then
      messagebox 0, "Nothing to search ! ", "Information", MB_ICONINFORMATION
   End If
   SendMessage getdlgitem(hmain, idc_edt31), EM_SETSEL, 0, - 1
   setfocus getdlgitem(hmain, idc_edt31)
End sub

Private Function revsearch(str1 As String, ini As integer) As integer
   dim x As integer
   dim y As integer
   dim as string mess1, mess2

   mess2 = ucase(str1)
   For x = ini To 0 Step - 1
      For y = 0 To 3
         mess1 = UCase(FF_ListView_GetItemText(HWND_FORM1_LISTVIEW1, x, y))
         If Instr(mess1, mess2) > 0 Then
            Function = x
            Exit Function
         End If
      Next
   Next
	show_line(ini + 1)
   messagebox 0, "No more occurrence ! " & Chr(13) & Chr(13) & str1, "Information", MB_ICONINFORMATION
   Function = - 1
End Function

Private Sub previous1()

   dim ret1 As integer, i As integer
   dim pos1 As integer
   dim nScroll As integer
   SelPath = ""
   If searching <> "" And pos0 > 0 Then
      ret1 = revsearch(searching, pos0 - 1)
      If ret1 > - 1 Then
         pos0 = ret1

         FF_ListView_SetSelectedItem(HWND_FORM1_LISTVIEW1, ret1)
         pos1 = FF_ListView_GetTopIndex(HWND_FORM1_LISTVIEW1)

         nScroll = pos1 - ret1
         If nScroll > 0 Then
            For i = 1 To nScroll
               SendMessage HWND_FORM1_LISTVIEW1, WM_VSCROLL, SB_LINEUP, 0
            Next
         Else
            For i = 1 To - nScroll
               SendMessage HWND_FORM1_LISTVIEW1, WM_VSCROLL, SB_LINEDOWN, 0
            Next
         End If
         SendMessage HWND_FORM1_LISTVIEW1, WM_SETREDRAW, TRUE, 0
         FF_Control_Redraw HWND_FORM1_LISTVIEW1

      End If
   elseIf searching <> "" And pos0 = 0 Then

      messagebox 0, "No more occurrence ! " & Chr(13) & Chr(13) & searching, "Information", MB_ICONINFORMATION
   elseIf searching = "" Then
      messagebox 0, "Nothing to search ! ", "Information", MB_ICONINFORMATION
   End If
   SendMessage getdlgitem(hmain, idc_edt31), EM_SETSEL, 0, - 1
   setfocus getdlgitem(hmain, idc_edt31)
End sub

Private Function force_click(hWin as hwnd) as integer
   dim as integer itab0, irow0, i
   dim lret As Long
   function = 0
   itab0 = tab_pos(hWin)
   getwindowtext getdlgitem(hWin, idc_edt1), prefix, 10
   IF LEN(SelPath) Then
   else
      if tab_pos(hWin) = 2 then
         sreport = ""
         showinfo
      end if
      chgtab(0)
      MessageBox BYVAL NULL, "Nothing selected !" & crlf & crlf & _
            "Click to select one on registered list or open a lib file.", _
            "Error", MB_OK OR MB_ICONERROR OR MB_TASKMODAL
      SendMessage getdlgitem(hWin, idc_edt31), EM_SETSEL, 0, - 1
      setfocus getdlgitem(hWin, idc_edt31)
      exit function

   End If
   function = 1
   if gene1 > 0 then sendmessage hWin, WM_COMMAND, cast(wparam, gene1), 0
END function

Private Function dlgCode_Init(Byval hWin as hwnd, Byval wParam as wParam, Byval lParam as lParam) as Integer
   hcode = getdlgitem(hwin, idc_red1)
   sendmessage hcode, wm_settext, 0, 0
   function = 0
End Function

Private Function dlgCode_OnSize(Byval hWin as hwnd, Byval lState as long, Byval cx as long, Byval cy as long) as Integer
   movewindow(getdlgitem(hwin, idc_red1), 0, 0, cx, cy, true)
   function = 0
End Function

Private Function dlgCode_OnClose(Byval hWin as hwnd) as Integer
   function = 0
End Function

Private Function dlgCode_OnDestroy(Byval hWin as hwnd) as Integer
   function = 0
End Function

Private Function dlgCode_OnNotify(Byval hWin as hwnd, Byval lCtlID as long, Byval lpNMHDR as NMHDR ptr, Byref lresult as long) as Integer
   function = 0
End Function

Private Function dlgCode_OnTimer(Byval hWin as hwnd, Byval wTimerID as word) as Integer
   function = 0
End Function

Private Function dlgTree_Init(Byval hWin as hwnd, Byval wParam as wParam, Byval lParam as lParam) as Integer
   htree = getdlgitem(hwin, idc_trv1)
   function = 0
End Function

Private Function dlgTree_OnSize(Byval hWin as hwnd, Byval lState as long, Byval cx as long, ByVal cy as long) as Integer
   movewindow(htree, 0, 0, cx, cy, true)
   function = 0
End Function

Private Function dlgTree_OnClose(Byval hWin as hwnd) as Integer
   function = 0
End Function

Private Function dlgTree_OnDestroy(Byval hWin as hwnd) as Integer
   function = 0
End Function

Private Function IDD_DLG1_idc_trv1_chk(Byval hWin as hwnd, ByVal hitem as integer) as Long
   Dim tstate As Integer
   tstate = TreeView_GetCheckState(cast(long, hwin), hitem)
   Treeview_ChangeChildState(cast(long, hwin), hitem, tstate)
   Treeview_ChangeParentState(cast(long, hwin), hitem, 0)
   function = 0
End Function

Private Function dlgTree_OnNotify(Byval hWin as hwnd, Byval lCtlID as long, Byval lpNMHDR as NMHDR ptr, Byref lresult as long) as Integer
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
         End If
      End If
   End If
   function = 0
End Function

Private Function dlgTree_OnTimer(Byval hWin as hwnd, Byval wTimerID as word) as Integer
   function = 0
End Function

Private Sub dlgList_IDC_LSV1_Sample(byval hWin as hwnd, byval CTLID as integer)
   dim dwStyle as dword
   hlist = getdlgitem(hwin, CTLID)
   dwStyle = SendMessage(hlist, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
   dwStyle = dwStyle OR LVS_REPORT or LVS_EX_FULLROWSELECT
   sendmessage hlist, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, dwStyle
   lvsetmaxcols(cast(long, hlist), 4)

   lvsetheader(cast(long, hlist), "GUID|Description|File|Path")
End sub

Private Function compare(BYVAL lParam1 AS LONG, BYVAL lParam2 AS LONG, BYVAL lParamsort AS LONG) AS Integer
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

Private Sub ResetLParam(ByRef hList AS LONG)
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

Private Function dlgList_Init(Byval hWin as hwnd, Byval wParam as wParam, Byval lParam as lParam) as Integer
   dlgList_IDC_LSV1_Sample(hWin, IDC_LSV1)
   function = 0
End Function

Private Function dlgList_OnSize(Byval hWin as hwnd, Byval lState as long, Byval cx as long, Byval cy as long) as Integer
   movewindow(getdlgitem(hwin, idc_lsv1), 0, 0, cx, cy, true)
   function = 0
End Function

Private Function dlgList_OnClose(Byval hWin as hwnd) as Integer
   function = 0
End Function

Private Function dlgList_OnDestroy(Byval hWin as hwnd) as Integer
   function = 0
End Function

Private Function dlgList_OnNotify(Byval hWin as hwnd, Byval lCtlID as long, Byval lpNMHDR as NMHDR ptr, Byref lresult as long) as Integer
   Dim pnmlv As NMLISTVIEW Ptr
   Dim item As Integer
   SELECT CASE lpNMHdr -> idFrom
      CASE idc_LSV1
         pNMLV = cast(NMLISTVIEW Ptr, lpnmhdr)
         SELECT CASE pNmlv -> hdr.Code
            Case LVN_COLUMNCLICK
               IF pNmLv -> iSubItem <> - 1 THEN
                  lviewSort = NOT lviewSort
                  ListView_SortItems(hList, procPTR(compare), pNmLv -> iSubItem)
                  ResetLParam cast(long, hList)
               END If
         End select
   End select
   function = 0
End Function

Private Function dlgList_OnTimer(Byval hWin as hwnd, Byval wTimerID as word) as Integer
   function = 0
End Function

Private Function dlgList_IDC_LSV1_Clicked(Byval hWin as hwnd, Byval pNMLV as NMLISTVIEW ptr) as long
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

Private Function dlgList_IDC_LSV1_Click1(Byval hWin as hwnd, Byval pNMLV as NMLISTVIEW ptr) as long
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
			setwindowtext GetDlgItem(hmain, IDC_SBR1)," Selected :  " & act1 & "    |    Path :  " & selpath
         LoadTLI
         chgtab1()
         treeview_select(htree, hparent(0), TVGN_CARET)

      End If
   End If
   gene1 = 0
	searching=""
   function = 0
end function

Private Function makefont(byval lCyPixels as long, sFntName as string, byval fFntSize as long, byval lWeight as long, byval lUnderlined as long, byval lItalic as long, byval lStrike as long, byval lCharSet as long) as long
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
   function = cast(long, CreateFontIndirect(BYVAL VARPTR(lF)))
end Function

Private Function dlgAbout_Init(Byval hWin as hwnd, Byval wParam as wParam, Byval lParam as lParam) as Integer
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

Private Function dlgAbout_OnClose(Byval hWin as hwnd) as Integer
   function = 0
End Function

Private Function dlgAbout_OnDestroy(Byval hWin as hwnd) as Integer
   function = 0
End Function

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

Private Function dlgMain_DlgProc(byval hWin as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM) as integer
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
               force_click(hWin)
               exit function
            case IDC_STC31
               if tab_pos(hWin) <> 0 THEN exit function
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
					chgtab(0)
               AskMain(hWin)
               SetWindowText GetDlgItem(hWin, IDC_EDT31), searching
               if searching <> "" and nook=0 THEN
                  search1()
               END IF
					nook=0
               exit function
            case Search_F3
					chgtab(0)
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

Private Function dlgList_DlgProc(byval hWin as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM) as integer
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

Private Sub cop_all()
   SendMessage hcode, EM_SETSEL, 0, - 1
   SendMessage hcode, WM_COPY, 0, 0
   SendMessage hcode, EM_SETSEL, - 1, 0
   SendMessage hcode, EM_SETSEL, 0, 0
END SUB

Private Sub cod21(byval iv as integer)

	dim D as GetTextLengthEX, T as GETTEXTEX, iLengthText as integer ,buf as string
	D.flags = GTL_DEFAULT
	D.codepage = CP_ACP
   iLengthText = SendMessage(hcode, EM_GetTextLengthEX, cast(wparam,VarPTR(D)),0)
	buf = Space(iLengthText)

	T.cb = iLengthText + 1
	T.flags = GT_DEFAULT
	T.codepage = CP_ACP
	T.lpDefaultChar = NULL
	T.lpUsedDefChar = NULL
	SendMessage hcode, EM_GETTEXTEX, cast(wparam,varptr(T)), cast(lparam,StrPTR(buf))

	if iv = 0 THEN
		iscod= Instr(iscod +1 , ucase(buf), ucase(searching2))
	elseif iv = 1 THEN
		iscod= Instrrev( ucase(buf), ucase(searching2),iscod -1)
   END IF
   If iscod Then
		SendMessage hcode, EM_SetSel , iscod-1, iscod + Len(searching2)-1  '+/-1 to convert to zero-based
		SendMessage hcode, EM_SCROLLCARET, 0, 0
	else
		messagebox 0, searching2 & "  : no more to find ! " , "Warning", mb_ok
	END IF
END SUB

Private Function dlgCode_DlgProc(byval hWin as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM) as integer
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

         dlgCode_OnDestroy(hWin)
         function = false
         exit function
      case else
         return FALSE
   end select
   return TRUE
End Function

Private Function dlgTree_DlgProc(byval hWin as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM) as integer
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

Private Function dlgAbout_DlgProc(byval hWin as HWND, byval uMsg as UINT, byval wParam as WPARAM, byval lParam as LPARAM) as integer
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

Private Function tab_pos(hWin1 as hwnd) as integer
   dim TabID as integer
   dim hIDC_TAB1 as hwnd
   hIDC_TAB1 = getdlgitem(hWin1, IDC_TAB1)
   TabID = SendMessage(hIDC_TAB1, TCM_GETCURSEL, 0, 0)
   function = TabID
END FUNCTION

Private Sub GetCopy1(posi1 as integer)
   dim as string sitem
   dim as integer irow = FF_ListView_GetSelectedItem(hList)
   if irow > - 1 THEN
      sitem = FF_ListView_GetItemText(hList, irow, posi1)
      FF_ClipboardSetText(sitem)
   else
      messagebox 0, "Nothing selected !", "Warning", mb_ok
   END IF
end sub

Private Function AskMain(ByVal hwParent As hwnd) As Integer
   if tab_pos(hwParent) = 0 then
      DialogBoxParam(GetModuleHandle(NULL), Cast(ZString Ptr, dlg_Ask), hwParent, @dlgAsk_DlgProc, 0)
      Return True
   Else
      messagebox 0, "Not in the list view tab !", "Warning", mb_ok
   End if
   Return False
End Function

Private Function dlgAsk_DlgProc(ByVal hWin As hWnd, ByVal uMsg As Uinteger, _
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

Private Function dlgAsk_DlgProc2(ByVal hWin As hWnd, ByVal uMsg As Uinteger, _
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

Private Function AskMain2(ByVal hwParent As hwnd) As Integer
		iscod=1
		nook=0
      DialogBoxParam(GetModuleHandle(NULL), Cast(ZString Ptr, dlg_Ask), hwParent, @dlgAsk_DlgProc2, 0)
      Return True
End Function

