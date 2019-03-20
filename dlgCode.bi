

Function dlgCode_Init(Byval hWin as hwnd,Byval wParam as wParam,Byval lParam as lParam)as Integer
	hcode=getdlgitem(hwin,idc_red1)
	sendmessage hcode,wm_settext,0,0
End Function
Function dlgCode_OnSize(Byval hWin as hwnd,Byval lState as long,Byval cx as long,Byval cy as long)as Integer
	movewindow(getdlgitem(hwin,idc_red1),0,0,cx,cy,true)
End Function
Function dlgCode_OnClose(Byval hWin as hwnd)as Integer
End Function
Function dlgCode_OnDestroy(Byval hWin as hwnd)as Integer
End Function
Function dlgCode_OnNotify(Byval hWin as hwnd,Byval lCtlID as long,Byval lpNMHDR as NMHDR ptr,Byref lresult as long)as Integer
End Function
Function dlgCode_OnTimer(Byval hWin as hwnd,Byval wTimerID as word)as Integer
End Function
