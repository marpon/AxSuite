
sub dlgMain_IDC_TAB1_Sample(byval hWin as hwnd,byval CTLID as integer)
	dim hCTLID as hwnd
	dim ts as TCITEM
	'Creating Tab
	hCTLID=GetDlgItem(hWin,CTLID)
	ts.mask=TCIF_TEXT
	ts.pszText=StrPtr("Registered")
	SendMessage(hCTLID,TCM_INSERTITEM,0,@ts)
	ts.pszText=StrPtr("Tree")
	SendMessage(hCTLID,TCM_INSERTITEM,1,@ts)
	ts.pszText=StrPtr("Code")
	SendMessage(hCTLID,TCM_INSERTITEM,2,@ts)
	hTab0=CreateDialogParam(hinstance,dlgList,hCTLID,@dlgList_DlgProc,0)
	hTab1=CreateDialogParam(hinstance,dlgTree,hCTLID,@dlgTree_DlgProc,0)
	hTab2=CreateDialogParam(hinstance,dlgCode,hCTLID,@dlgCode_dlgProc,0)
	ShowWindow(hTab0,SW_SHOW)
	ShowWindow(hTab1,SW_HIDE)
	ShowWindow(hTab2,SW_HIDE)	
end sub

Function dlgMain_Init(Byval hWin as hwnd,Byval wParam as wParam,Byval lParam as lParam)as Integer
	Dim hicon As Integer
	hinstance=getwindowlong(hwin,gwl_hInstance)
	hmain=hwin
	hIcon = LoadIcon(hinstance,100)
	SendMessage hwin,WM_SETICON,ICON_SMALL, hIcon	
	dlgMain_IDC_TAB1_Sample(hWin,IDC_TAB1)
	TabCtrl_GetItemRect(getdlgitem(hwin,idc_tab1),2,@rtab)
	movewindow(getdlgitem(hwin,idc_stc3),rtab.right+5,rtab.top,40,18,TRUE)
	movewindow(getdlgitem(hwin,idc_edt1),rtab.right+45,rtab.top,100,18,TRUE)
	SearchTypeLibs
	listview_setcolumnwidth(hlist,1,LVSCW_AUTOSIZE)		
End Function

Function dlgMain_OnSize(Byval hWin as hwnd,Byval lState as long,Byval cx as long,Byval cy as long)as Integer
	movewindow(getdlgitem(hwin,idc_tab1),0,0,cx,cy,TRUE)
	movewindow(htab0,0,22,cx,cy-22,TRUE)
	movewindow(htab1,0,22,cx,cy-22,TRUE)
	movewindow(htab2,0,22,cx,cy-22,TRUE)
	
End Function
Function dlgMain_OnClose(Byval hWin as hwnd)as Integer
End Function
Function dlgMain_OnDestroy(Byval hWin as hwnd)as Integer
End Function
Function dlgMain_OnNotify(Byval hWin as hwnd,Byval lCtlID as long,Byval lpNMHDR as NMHDR ptr,Byref lresult as long)as Integer
End Function
Function dlgMain_OnTimer(Byval hWin as hwnd,Byval wTimerID as word)as Integer
End Function
Function dlgMain_IDC_TAB1_SelChange(byval hWin as hwnd,byval lpNMHDR as NMHDR ptr,Byref lresult as long)as long
	dim TabID as Long
	dim hIDC_TAB1 as hwnd
	hIDC_TAB1=getdlgitem(hWin,IDC_TAB1)
	TabID=SendMessage(hIDC_TAB1,TCM_GETCURSEL,0,0)
	select case TabID
		case 0
			ShowWindow(hTab0,SW_SHOW)
			ShowWindow(hTab1,SW_HIDE)
			ShowWindow(hTab2,SW_HIDE)
			setfocus hlist
		Case 1	
			ShowWindow(hTab0,SW_HIDE)
			ShowWindow(hTab1,SW_SHOW)
			ShowWindow(hTab2,SW_HIDE)
			setfocus htree
		Case 2	
			ShowWindow(hTab0,SW_HIDE)
			ShowWindow(hTab1,SW_HIDE)
			ShowWindow(hTab2,SW_SHOW)
			setfocus hcode			
	end select
End function
Function dlgMain_IDC_TAB1_SelChanging(byval hWin as hwnd,byval lpNMHDR as NMHDR ptr,Byref lresult as long)as long
	dim TabID as Long
	dim hIDC_TAB1 as hwnd
	hIDC_TAB1=getdlgitem(hWin,IDC_TAB1)
	TabID=SendMessage(hIDC_TAB1,TCM_GETCURSEL,0,0)
	select case TabID
		case 0
			'ShowWindow(hTabDlg0,SW_HIDE)
		case 1	
			'ShowWindow(hTabDlg1,SW_HIDE)
	end select
End Function

Sub chgtab(ByVal p As Integer)
	Dim nmh As nmhdr
	nmh.hwndfrom=getdlgitem(hmain,idc_tab1)
	nmh.idfrom=idc_tab1
	nmh.code=TCN_SELCHANGE
	sendmessage getdlgitem(hmain,idc_tab1),TCM_SETCURSEL,p,0
	sendmessage getdlgitem(hmain,idc_tab1),WM_NOTIFY,idc_tab1,@nmh
	If p=1 Then 
		treeview_setcheckstate(htree,TreeView_GetChild(hTree,0),1)
		treeview_changechildstate(htree,TreeView_GetChild(hTree,0),1)
	EndIf
End Sub

Sub dllcodegen(dlls As String)
	DylibFree(hdll)
	hdll=DylibLoad(dlls)
	putEvtList=DylibSymbol(hdll,"putEvtList")
	clrEvtList=DylibSymbol(hdll,"clrEvtList")
	EnumvTable=dylibsymbol(hdll,"EnumvTable")
	EnumInvoke=dylibsymbol(hdll,"EnumInvoke")
	EnumEvent=dylibsymbol(hdll,"EnumEvent")
	EnumEnum=dylibsymbol(hdll,"EnumEnum")
	EnumRecord=dylibsymbol(hdll,"EnumRecord")
	EnumUnion=dylibsymbol(hdll,"EnumUnion")
	EnumModule=dylibsymbol(hdll,"EnumModule")
	EnumAlias=dylibsymbol(hdll,"EnumAlias")
	EnumCoClass=dylibsymbol(hdll,"EnumCoClass")	
End Sub

Const szNULL=!"\0"
Const szOpen	= "tlb;olb (*.tlb;*.olb)" & szNULL & "*.tlb;*.olb" & szNULL & "exe;dll;ocx (*.exe;*.dll;*.ocx)" & szNULL & "*.exe;*.dll;*.ocx" & szNULL
Const szSave	= "bi (*.bi)" & szNULL
Const szDll	= "dll (*.dll)" & szNULL
Function dlgMain_IDR_MENU1_menu(byval hwin as hwnd,byval id as long) as Long
	dim as OPENFILENAME ofn
	dim as zstring*max_path pathfilename
	Dim dllpath As String
	
	Select Case id		
		case tlb_reged
			SearchTypeLibs
		Case tlb_libfile
			ofn.lStructSize=sizeof(OPENFILENAME)
			ofn.hwndOwner=hwin
			ofn.hInstance=GetWindowLong(hwin,GWL_HINSTANCE)
			ofn.lpstrFile=@pathFileName
			ofn.nMaxFile=260
			ofn.lpstrFilter=@szOpen
			ofn.Flags=OFN_FILEMUSTEXIST or OFN_HIDEREADONLY or OFN_PATHMUSTEXIST
			if GetOpenFileName(@ofn) Then selpath=pathfilename			
		   IF LEN(SelPath) Then
		   	clsids=""
		   	tldesc=""
		   	cleartli
		      IF LoadtypeLib(stringtobstr(selpath),@ptli) Then
		      	MessageBox BYVAL NULL, "Unable to load " & SelPath, _
		         "Error", MB_OK OR MB_ICONERROR OR MB_TASKMODAL
		      Else      	
					LoadTLI
		      	chgtab(1)
		      	treeview_select(htree,hparent(0),TVGN_CARET)
		      End If
		   End If
		case tlb_exit
			EndDialog(hWin,0)
		Case code_codegen
			pathfilename=ExePath & "\*.dll"
			ofn.lStructSize=sizeof(OPENFILENAME)
			ofn.hwndOwner=hwin
			ofn.hInstance=GetWindowLong(hwin,GWL_HINSTANCE)
			ofn.lpstrFile=@pathFileName
			ofn.nMaxFile=260
			ofn.lpstrFilter=@szDll
			ofn.Flags=OFN_FILEMUSTEXIST or OFN_HIDEREADONLY or OFN_PATHMUSTEXIST
			if GetOpenFileName(@ofn) Then dllpath=pathfilename Else dllpath=""			
		   IF LEN(dllPath) Then dllcodegen(dllpath)
		case code_const
			savepath="Constant"
			mkconstant
			chgtab(2)	
		Case code_Module
			savepath="Module"
			mkModule
			chgtab(2)
		Case code_vTable
			savepath="vTable"
			mkvtable
			chgtab(2)
		case code_Invoke
			savepath="Invoke"
			mkInvoke
			chgtab(2)
		case code_event
			savepath="Event"
			mkEvent
			chgtab(2)
		Case code_save
			pathfilename=str_parse(str_parse(selpath,"\",str_numparse(selpath,"\")),".",1)
			pathfilename &="_" & savepath & ".bi"
			ofn.lStructSize=sizeof(OPENFILENAME)
			ofn.hwndOwner=hwin
			ofn.hInstance=GetWindowLong(hwin,GWL_HINSTANCE)
			ofn.lpstrFile=@pathFileName
			ofn.nMaxFile=260
			ofn.lpstrFilter=@szSave
			ofn.Flags=OFN_FILEMUSTEXIST or OFN_HIDEREADONLY or OFN_PATHMUSTEXIST
			if GetSaveFileName(@ofn) Then savepath=pathfilename			
		   IF LEN(SavePath) Then
				Open savepath for Binary As #1
					Put #1,,sreport
				Close #1
		   End If
		case hlp_axsuite
			pathfilename=ExePath & "\axsuite.pdf"
			shellexecute 0,strptr("Open"),@pathfilename,NULL,NULL,SW_SHOWNORMAL
		case hlp_About
			DialogBoxParam(GetModuleHandle(NULL), dlgAbout, hwin, @dlgAbout_DlgProc, NULL)
	End Select
End Function
