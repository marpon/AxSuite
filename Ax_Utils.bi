'from lv.bi
const LVM_SORTITEMSEX = LVM_FIRST + 81

TYPE LV_DISPINFO
   hdr                             AS NMHDR
   item                            AS LV_ITEM
END TYPE


'from axhelper.bi

'convert string to bstr
'please follow with sysfreestring(bstr) after use to avoid memory leak
Function StringToBSTR(cnv_string As String) As BSTR
   Dim sb As BSTR
   Dim As Integer n
   n = (MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, cnv_string, - 1, NULL, 0)) - 1
   sb = SysAllocStringLen(sb, n)
   MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, cnv_string, - 1, sb, n)
   Return sb
End Function

Function FromBSTR(ByVal szW As BSTR) As String
      dim as string sret
      dim as integer L1
      If szW = Null Then Return ""
      L1 = WideCharToMultiByte(CP_ACP, 0, SzW, - 1, Null, 0, Null, Null) - 1
      sret = space(L1)
      WideCharToMultiByte(CP_ACP, 0, SzW, L1, sret, L1, Null, Null)
      Return sret
   End Function



'return numeric value from variant
Function Variantv(ByRef v As variant) As Double
   Dim vvar As variant
   VariantChangeTypeEx(@vvar, @v, NULL, VARIANT_NOVALUEPROP, VT_R8)
   Return vvar.dblval
End Function



'from strwrap.bi
'******************************************
'remove match string from source string
'******************************************
Function str_remove(ByRef source As String, byref match As String) As String
   dim as long s = 1, lmt
   Dim As String txt
   txt = source
   lmt = Len(match)
   Do
      s = InStr(s, txt, match)
      If s Then txt = Left(txt, s - 1) + Right(txt, Len(txt) - lmt - (s - 1))
   Loop While s
   Function = txt
end function

'***********************************************************************
'remove any character in Match string from stxt string
'***********************************************************************
Function Str_removeany(ByRef source As String, ByRef match As String) As String
   dim as long s = 1
   Dim As String txt
   txt = source
   Do
      s = Instr(s, txt, Any match)
      If s Then txt = Left(txt, s - 1) + Right(txt, Len(txt) - 1 - (s - 1))
   Loop While s
   Function = txt
end Function

'***************************************************************
'count number of parse string, separated by delimiter of source
'***************************************************************
Function str_numparse(ByRef source as string, ByRef delimiter as string) as long
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

'************************************************************
'parse source string, indexed by delimiter string at idx
'************************************************************
Function str_parse( source As String,  delimiter As String, ByVal idx As integer) As String
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

'**********************************************************
'find index of match string as parse source by delimiter
'**********************************************************
Function str_parsepos(ByRef match As String, byref source As String, ByRef delimiter As String) As Integer
   Dim As Integer i
   For i = 1 To str_numparse(source, delimiter)
      If str_parse(source, delimiter, i) = match Then Return i
   Next
   Function = 0
End Function

'************************************************************************
'count number of parse string, separated by any character in delimiter string
'************************************************************************
Function str_numparseany(ByRef source As String, Byref delimiter As String) As Long
   Dim As Long s = 1, c
   Do
      s = Instr(s, source, Any delimiter)
      If s Then
         c += 1
         s += 1
      End If
   Loop While s
   Function = c + 1
End Function

'********************************************************************
'parse source string, separated by any character in delimiter string
'********************************************************************
Function str_parseany(ByRef source As String, Byref delimiter As String, ByVal idx As Long) As String
   Dim As Long s = 1, c
   Do
      If c = idx - 1 Then Return Mid(source, s, Instr(s, source, Any delimiter) - s)
      s = Instr(s, source, Any delimiter)
      If s Then
         c += 1
         s += 1
      End If
   Loop While s
End Function

'**********************************************************
'find index of match string as parse source by delimiter
'**********************************************************
Function str_parseanypos(ByRef match As String, byref source As String, ByRef delimiter As String) As Integer
   Dim As Integer i
   For i = 1 To str_numparseany(source, delimiter)
      If str_parseany(source, delimiter, i) = match Then Return i
   Next
   Function = 0
End Function

'*****************************************************
'count occurance of delimiter string in source string
'*****************************************************
function str_tally(ByRef source as string, byref delimiter as string) as long
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

'**************************************************************
'count occurance of any character in del string in cmd string
'**************************************************************
function str_tallyany(ByRef source as string, byref delimiter as string) as long
   Dim As Long s = 1, c
   Do
      s = InStr(s, source, any delimiter)
      If s then
         c += 1
         s += 1
      end if
   Loop While s
   Function = c
End function


Function FF_Replace( ByRef sText        As String, _
                     ByRef sLookFor     As String, _
                     ByRef sReplaceWith As String _
                     ) As String

   Dim sTemp As String  = sText
   Dim f     As Integer = 1
   If sLookFor = sReplaceWith Then

      Function = sTemp
      Exit Function
   EndIf
   If InStr(sReplaceWith,sLookFor) Then
        Do Until f = 0
           f = InStr( sTemp, sLookFor )
           If f Then
              sTemp = Left(sTemp, f-1) & "#§§~¿~§§#" & Mid(sTemp, f + Len(sLookFor))
           End If
        Loop
        f = 1
        Do Until f = 0
          f = InStr( sTemp, "#§§~¿~§§#" )
          If f Then
             sTemp = Left(sTemp, f-1) & sReplaceWith & Mid(sTemp, f + 9)
          End If
        Loop

        Function = sTemp
        Exit Function
   EndIf
   Do Until f = 0
      f = InStr( sTemp, sLookFor )
      If f Then
         sTemp = Left(sTemp, f-1) & sReplaceWith & Mid(sTemp, f + Len(sLookFor))
      End If
   Loop

   Function = sTemp

End Function



'***********************************************************
'replace first occurance of match string with replace string in source string
'***********************************************************
Function str_replace(ByRef match As String, Byref sreplace As String, Byref source As String) As String
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

'****************************************************************
'replace first occurance of any character in del string with r string in cmd string
'*****************************************************************
Function Str_replaceany(ByRef match As String, Byref sreplace As String, Byref source As String) As String
   Dim As Long s = 1, c = len(sreplace)
   Dim As String text
   text = source
   Do
      s = InStr(s, text, Any match)
      If s then
         text = Mid$(text, 1, s - 1) + sreplace + Mid$(text, s + 1, - 1)
         s += c
      end if
   Loop While s
   Function = text
End Function

'****************************************************************
'replace all occurance of any character in del string with r string in cmd string
'*****************************************************************
Function str_ReplaceAll(ByRef match As String, ByRef sreplace As String, ByRef source As String) As String
   Dim text As String
   text = source
   While InStr(text, match)
      text = str_replace(match, sreplace, text)
   Wend
   Function = text
End Function

'****************************************************************
'replace all occurance of any character in del string with r string in cmd string
'*****************************************************************
Function str_ReplaceAnyAll(ByRef match As String, ByRef sreplace As String, ByRef source As String) As String
   Dim text As String
   text = source
   While InStr(text, Any match)
      text = str_replaceany(match, sreplace, text)
   Wend
   Function = text
End Function


'from lv.bi

Sub LvInit
   dim iccex AS InitCommonControlsEx
   iccex.dwSize = SIZEOF(iccex)
   iccex.dwICC = ICC_LISTVIEW_CLASSES
   InitCommonControlsEx(@iccex)
End Sub

SUB LVSetHeader(BYVAL lvg AS uinteger, BYVAL strItem AS STRING)
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

Sub LVSetColFmt(BYVAL lvg AS uinteger, ByVal col As LONG, BYVAL fmt AS LONG)
   dim szItem AS zstring * MAX_PATH              ' working variable
   dim tlvc AS LVCOLUMN                          ' specifies or receives the attributes of a listview column
   'dim tlvi          AS LV_ITEM               ' specifies or receives the attributes of a listview item

   tlvc.mask = LVCF_FMT OR LVCF_SUBITEM
   tlvc.fmt = fmt
   tlvc.iSubItem = col
   SendMessage cast(hwnd, lvg), LVM_SETCOLUMN, col, cast(LPARAM, VARPTR(tlvc))

End Sub

SUB LVSetMaxCols(BYVAL rr AS uinteger, BYVAL cols AS LONG)
   dim tlvc AS LVCOLUMN
   dim col AS LONG

   FOR col = 0 TO cols - 1 :SendMessage cast(hwnd, rr), LVM_INSERTCOLUMN, 0, cast(LPARAM, VARPTR(tlvc)) :NEXT
END SUB

SUB LVSetMaxRows(BYVAL rr as uinteger, BYVAL rows AS LONG)
   dim irow AS LONG
   dim tlvi AS LVITEM
   dim glvnumrows as long
   glvnumrows = sendmessage(cast(hwnd, rr), LVM_GETITEMCOUNT, 0, 0)
   IF glvnumrows <= rows THEN
      FOR irow = glvnumrows TO rows - 1
         tlvi.iItem = irow
         SendMessage cast(hwnd, rr), LVM_INSERTITEM, irow, cast(LPARAM, VARPTR(tlvi))
      NEXT
   ELSE
      FOR irow = rows TO glvnumrows - 1
         SendMessage cast(hwnd, rr), LVM_DELETEITEM, 0, cast(LPARAM, rows)
      NEXT
   END IF
END SUB

SUB LVSetValue(BYVAL rr AS uinteger, BYVAL item AS LONG, BYVAL column AS LONG, BYVAL strs AS STRING)
   dim ms_lvi AS LVITEM
   dim psztext as zstring*max_path
   psztext = strs
   ms_lvi.iitem = item
   ms_lvi.iSubItem = column
   ms_lvi.pszText = @pszText
   ms_lvi.lparam = item
   SendMessage cast(hwnd, rr), LVM_SETITEMTEXT, item, cast(LPARAM, VARPTR(ms_lvi))
END SUB

FUNCTION LVGetValue(BYVAL rr AS uinteger, BYVAL item AS LONG, BYVAL column AS LONG) AS STRING
   dim ms_lvi AS LV_ITEM
   dim psztext as zstring*max_path
   ms_lvi.iSubItem = Column
   ms_lvi.cchTextMax = max_path
   ms_lvi.pszText = @pszText
   SendMessage cast(hwnd, rr), LVM_GETITEMTEXT, item, cast(LPARAM, @ms_lvi)
   FUNCTION = psztext
END FUNCTION

function lvFindItem(hList AS LONG, Value AS STRING) as integer
   DIM lvf AS LVFINDINFO
   DIM i AS integer
   DIM szStr AS zstring * 300
   DIM iStatus AS INTEGER

   lvf.flags = LVFI_STRING OR LVFI_Partial
   szStr = ucase$(Value)

   lvf.psz = @szStr
   i = SendMessage(cast(hwnd, hList), LVM_FINDITEM, - 1, cast(LPARAM, @lvf))
   IF i >= 0 THEN
      istatus = ListView_EnsureVisible(cast(hwnd, hList), i, true)
   END IF
   function = i
END function

SUB LVSetItem(hList AS LONG, byval item as long, Rec AS STRING)
   DIM z AS INTEGER
   DIM iStatus AS INTEGER
   DIM szStr AS zstring* 300
   DIM lvi AS LV_ITEM
   dim x AS LONG

   'this will be the next record

   lvi.iItem = item
   lvi.mask = LVIF_TEXT
   lvi.stateMask = LVIS_FOCUSED                  '%LVIS_SELECTED
   lvi.pszText = @szStr

   FOR z = 0 TO str_numparse(Rec, "|") - 1
      szStr = str_parse(Rec, "|", z + 1)
      lvi.iSubItem = z

      lvi.lParam = lvi.iItem
      IF z = 0 THEN
         lvi.mask = LVIF_TEXT OR LVIF_PARAM OR LVIF_STATE
         iStatus = ListView_setItem(cast(hwnd, hList), @lvi)
      ELSE
         lvi.mask = LVIF_TEXT
         iStatus = ListView_SetItem(cast(hwnd, hList), @lvi)
      END IF
   NEXT
END SUB


SUB LVAddItem(hList AS LONG, Rec AS STRING)
   DIM z AS INTEGER
   DIM iStatus AS INTEGER
   DIM szStr AS zstring* 300
   DIM lvi AS LV_ITEM
   dim x AS LONG

   'this will be the next record

   'ListView_GetItemCount(hList) '+ 1
   lvi.iItem = sendmessage(cast(hwnd, hlist), lvm_GetItemCount, 0, 0)

   lvi.mask = LVIF_TEXT
   lvi.stateMask = LVIS_FOCUSED                  '%LVIS_SELECTED
   lvi.pszText = @szStr

   FOR z = 0 TO str_numparse(Rec, "|") - 1
      szStr = str_parse(Rec, "|", z + 1)
      lvi.iSubItem = z

      lvi.lParam = lvi.iItem
      IF z = 0 THEN
         lvi.mask = LVIF_TEXT OR LVIF_PARAM OR LVIF_STATE
         iStatus = ListView_InsertItem(cast(hwnd, hList), @lvi)
      ELSE
         lvi.mask = LVIF_TEXT
         iStatus = ListView_SetItem(cast(hwnd, hList), @lvi)
      END IF
   NEXT
END SUB


'from tv.bi
Function TreeViewInsertItem(BYVAL hTreeView AS long, BYVAL hParent AS long, sItem AS STRING) AS LONG
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

Function TreeView_GetItemText(ByVal hTreeView As Long, ByVal hItem As Long) As String
   Dim ItemText As zstring*32765
   Dim Item As TV_ITEM
   Item.hItem = cast(HTREEITEM, hItem)
   Item.Mask = TVIF_TEXT
   Item.cchTextMax = 32765
   Item.pszText = @ItemText
   SendMessage cast(hwnd, hTreeView), TVM_GETITEM, 0, cast(LPARAM, @Item)
   Function = ItemText
End Function

Function TreeView_ChangeChildState(ByVal hTreeView As Long, ByVal hItem As Long, ByVal iState As Long) As Long
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

Function TreeView_ChangeParentState(ByVal hTreeView As Long, ByVal hItem As Long, ByVal iState As Long) As Long
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

Function TreeView_SetCheckState(ByVal hTreeView As Long, ByVal hItem As Long, ByVal iState As Long) As Long
   Dim TVITEM As TV_ITEM
   TVITEM.mask = TVIF_STATE Or TVIF_HANDLE
   TVITEM.hItem = cast(HTREEITEM, hItem)
   TVITEM.stateMask = TVIS_STATEIMAGEMASK
   TVITEM.state = (iState + 1) Shl 12
   sendmessage cast(hwnd, hTreeView), TVM_SetItem, 0, cast(LPARAM, @TVITEM)
   function = 1
End Function

Function TreeView_GetCheckState(ByVal hWnd1 As Long, ByVal hItem As Long) As Long
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

'---------------------------------------------------
'Determines if the current state image of the
'specified treeview item is set to the checked
'checkbox image index.
'
'hwndTV - treeview window handle
'hItem - item's handle whose checkbox state is to be to returned
'
'Returns true if the item's state image is
'set to the checked checkbox index, returns
'false otherwise.
'---------------------------------------------------
Function IsTVItemChecked(hwndTV as Long, hItem as Long) as integer
	Dim tvi as TV_ITEM
	dim i1 as integer
	'Initialize the struct and get the item's state value.
	With tvi
	.mask = TVIF_STATE
	.hItem = cast(HTREEITEM, hItem)
	.stateMask = TVIS_STATEIMAGEMASK
	End With
	sendmessage cast(hwnd, hwndTV), TVM_GetItem, 0, cast(LPARAM, @tvi)
	'We have to test to see if the treeview
	'checked state image *is* set since the logical
	'And test on the unchecked image (1) will
	'evaluate to true when either checkbox image
	'is set.
	i1=(tvi.state And INDEXTOSTATEIMAGEMASK(TV_CHECKED))
	i1 Shr= 13
	function = i1

End Function


Function FF_FileName (Src As String) As String
    Dim x As Integer
    x = InStrRev( Src, Any ":/\")
    If x Then
        Function = Mid(Src, x + 1)
    Else
        Function = Src
    End If
End Function