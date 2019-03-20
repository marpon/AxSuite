Function dlgTree_Init(Byval hWin as hwnd,Byval wParam as wParam,Byval lParam as lParam)as Integer	
	htree=getdlgitem(hwin,idc_trv1)
End Function
Function dlgTree_OnSize(Byval hWin as hwnd,Byval lState as long,Byval cx as long,ByVal cy as long)as Integer
	movewindow(htree,0,0,cx,cy,true)
End Function
Function dlgTree_OnClose(Byval hWin as hwnd)as Integer
End Function
Function dlgTree_OnDestroy(Byval hWin as hwnd)as Integer
End Function

Type TV_HITTESTINFO
    pt As POINT 
    flags As UINT 
    hitem As HTREEITEM 	
End Type

Type LPTV_HITTESTINFO As TV_HITTESTINFO ptr

function IDD_DLG1_idc_trv1_chk(Byval hWin as hwnd,ByVal hitem as integer)as Long
	Dim tstate As Integer
	tstate=TreeView_GetCheckState(hwin,hitem)
	Treeview_ChangeChildState(hwin,hitem,tstate)
	Treeview_ChangeParentState(hwin,hitem,0)	 
End Function

Function dlgTree_OnNotify(Byval hWin as hwnd,Byval lCtlID as long,Byval lpNMHDR as NMHDR ptr,Byref lresult as long)as Integer
	Dim As Long event,itmp
	Dim ht As TV_HITTESTINFO	
	Dim As dword dwpos	
	event=lpnmhdr->code
	if lctlid=idc_trv1 Then
		If event=NM_CLICK Then
			dwpos = GetMessagePos()
			iTmp = LoWord(dwpos)
			ht.pt.x = iTmp
			iTmp = HiWord(dwpos)
			ht.pt.y = iTmp
			MapWindowPoints(HWND_DESKTOP, lpnmhdr->hWndFrom, @ht.pt, 1)
			TreeView_HitTest(lpnmhdr->hwndFrom, @ht)
			If ht.flags = TVHT_ONITEMSTATEICON Then
				IDD_DLG1_idc_trv1_chk(lpnmhdr->hWndFrom, ht.hItem)
			End If
		EndIf
	EndIf
End Function

Function dlgTree_OnTimer(Byval hWin as hwnd,Byval wTimerID as word)as Integer
End Function

