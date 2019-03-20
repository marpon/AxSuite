

function makefont(byval lCyPixels as long, sFntName as string,byval fFntSize as long, byval lWeight as long, byval lUnderlined as long,byval lItalic as long,byval lStrike as long,byval lCharSet as long)as long
	dim lf as LOGFONT
	lF.lfHeight = -(fFntSize * lCyPixels) \ 69
	lF.lfFaceName = sFntName	
   IF lWeight = FW_DONTCARE THEN
   	lf.lfWeight = FW_NORMAL
   ELSE
   	lf.lfWeight = lWeight
   END IF 
	lf.lfUnderline  = lUnderlined
	lf.lfItalic     = lItalic
	lf.lfStrikeOut  = lStrike
	lf.lfCharSet    = lCharSet      
	'create font and assign handle to a global 'variable
	function = CreateFontIndirect( BYVAL VARPTR(lF))	
end Function

Function dlgAbout_Init(Byval hWin as hwnd,Byval wParam as wParam,Byval lParam as lParam)as Integer
	Dim text As String
	Dim As Integer hIcon, hFont, hDC
	
	SetTimer hwin,ABOUT_TIMER,5000,NULL
	
	hIcon = LoadIcon(0,IDI_ASTERISK)			
	SendMessage hwin,WM_SETICON, ICON_BIG, hIcon

	hDC = GetDC(0)
	hfont=MakeFont(GetDeviceCaps(hDC, LOGPIXELSY),"Times New Roman",20,fw_bold,false, false, FALSE, ANSI_CHARSET)                   
	ReleaseDC 0, hDC
	SendMessage GetDlgItem(hwin,idc_stc1),WM_SETFONT,hFont,0
End Function
Function dlgAbout_OnSize(Byval hWin as hwnd,Byval lState as long,Byval cx as long,Byval cy as long)as Integer
End Function
Function dlgAbout_OnClose(Byval hWin as hwnd)as Integer
End Function
Function dlgAbout_OnDestroy(Byval hWin as hwnd)as Integer
End Function
Function dlgAbout_OnNotify(Byval hWin as hwnd,Byval lCtlID as long,Byval lpNMHDR as NMHDR ptr,Byref lresult as long)as Integer
End Function
Function dlgAbout_OnTimer(Byval hWin as hwnd,Byval wTimerID as word)as Integer
	If (wTimerID=ABOUT_TIMER) then
		EndDialog(hwin, 0)
	end If
End Function