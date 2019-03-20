sub dlgList_IDC_LSV1_Sample(byval hWin as hwnd,byval CTLID as integer)
	'dim hCTLID as hwnd
	dim dwStyle as dword
	hlist=getdlgitem(hwin,CTLID)
	dwStyle=SendMessage(hlist,LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
	dwStyle = dwStyle OR LVS_REPORT or LVS_EX_FULLROWSELECT
	sendmessage hlist, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, dwStyle
	lvsetmaxcols(hlist,4)
	lvsetheader(hlist,"GUID|Description|File|Path")

End sub

FUNCTION compare(BYVAL lParam1 AS LONG, BYVAL lParam2 AS LONG, BYVAL lParamsort AS LONG) AS Integer
    dim value1 AS zstring * 300
    dim value2 AS zstring * 300
    dim j as integer

    FUNCTION = 0

    value1 = LVGetValue(hList, lParam1, lParamsort)
    value2 = LVGetValue(hList, lParam2, lParamsort)
    SELECT CASE lviewSort
        CASE 0 'ascending
        j = 1
        CASE ELSE 'descending
        j = -1
    END SELECT
    IF value1 < value2 THEN
        FUNCTION = -1 * j
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
    recs = sendmessage(hlist,lvm_GetItemCount,0,0)
    FOR i = 0 TO recs - 1
        lvi.iItem = i
        x = SendMessage(hList,LVM_GetItem,0,@lvi)
        lvi.lParam = lvi.iItem
        x = SendMessage(hList,LVM_SetItem,0,@lvi)
    NEXT
END SUB

Function dlgList_Init(Byval hWin as hwnd,Byval wParam as wParam,Byval lParam as lParam)as Integer
	dlgList_IDC_LSV1_Sample(hWin,IDC_LSV1)

End Function
Function dlgList_OnSize(Byval hWin as hwnd,Byval lState as long,Byval cx as long,Byval cy as long)as Integer
	movewindow(getdlgitem(hwin,idc_lsv1),0,0,cx,cy,true)
End Function
Function dlgList_OnClose(Byval hWin as hwnd)as Integer
End Function
Function dlgList_OnDestroy(Byval hWin as hwnd)as Integer
End Function
Function dlgList_OnNotify(Byval hWin as hwnd,Byval lCtlID as long,Byval lpNMHDR as NMHDR ptr,Byref lresult as long)as Integer
	Dim pnmlv As NMLISTVIEW Ptr
	Dim item As Integer
	SELECT CASE lpNMHdr->idFrom
		CASE idc_LSV1
		pNMLV=lpnmhdr
      SELECT CASE pNmlv->hdr.Code
			Case LVN_COLUMNCLICK
				IF pNmLv->iSubItem <> -1 THEN
      			'toggle ascending/descending
      			lviewSort = NOT lviewSort
      			ListView_SortItems(hList,procPTR(compare), pNmLv->iSubItem)
      			ResetLParam hList
				END If
  		End select
	End select
End Function
Function dlgList_OnTimer(Byval hWin as hwnd,Byval wTimerID as word)as Integer
End Function
function dlgList_IDC_LSV1_Clicked(Byval hWin as hwnd,Byval pNMLV as NMLISTVIEW ptr)as long
	dim hIDC_LSV1 as hwnd
	dim item as Long, lret As Long

	hIDC_LSV1=GetDlgItem(hWin,IDC_LSV1)
	item=pnmlv->iitem
   selpath=lvgetvalue(hlist,item,3)
   clsids=LVGetvalue(hlist,item,0)
   tldesc=LVGetvalue(hlist,item,1)
   IF LEN(SelPath) Then
   	cleartli
      lret=LoadtypeLib(stringtobstr(selpath),@ptli)
      IF lret Then
      	MessageBox BYVAL NULL, "Unable to load " & SelPath, _
         "Error", MB_OK OR MB_ICONERROR OR MB_TASKMODAL
      Else
			LoadTLI
      	chgtab(1)
      	treeview_select(htree,hparent(0),TVGN_CARET)
      End If
   End If
end function


