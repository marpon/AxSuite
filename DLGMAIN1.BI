Function FileFound(mes As String) As Integer
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


function idem_prefix(hwin as hwnd) as integer
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
	'print "function= " & i1 & "    prefix0  >" & prefix0 & "<   prefix  >" & prefix1 & "<"
END FUNCTION

Function FF_ClipboardSetText(ByVal Text As String) As Long
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

Sub mkInit()
	sevent=""
   sreport = Ini_templ()
   showinfo
END SUB

Sub mkvident
   Dim b As zString Ptr
	sevent=""
	sreport= Ini_templ()
   b = EnumCoClass1(cast(long, htree), hparent(), 0, "", "")
   sreport &=  *b :sysfreestring(cast(BSTR, b))
	if xret1=0 then
'		b = EnumModel(cast(long, htree), hparent())
'		sreport &= *b :sysfreestring(cast(BSTR, b))
	end if
	xret1=0
	xfull=0
   showinfo
End Sub

Sub mkuOcx1
   Dim b As zString Ptr
	sevent=""
   b = EnumCoClass3(cast(long, htree), hparent())
   sreport =  *b :sysfreestring(cast(BSTR, b))
	xret1=0
	xfull=0
   showinfo
End Sub

Sub mkuOcx2
   Dim b As zString Ptr
	sevent=""
   b = EnumCoClass3(cast(long, htree), hparent())
   sreport =  *b :sysfreestring(cast(BSTR, b))
	xret1=0
	xfull=0
   showinfo
End Sub

Sub mkConstant
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

function verif(s1 as string) as string
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

'Sub mkFull2
'   Dim b As zString Ptr
'   dim as string s1, s2, s3
'	xfull = 1
'	'print "EnumEvent xfull = " & xfull
'	b = EnumEvent(cast(long, htree), hparent())
'	s1 = *b :sysfreestring(cast(BSTR, b))
'	s2 = verif(s1)
'	'print "EnumModel"
'	'b = EnumModel(cast(long, htree), hparent())
'	's3= *b :sysfreestring(cast(BSTR, b))

'   if s1 = "" or s2 = "" THEN
'      s1 = ""
'		'print "EnumCoClass1 _1"
'      b = EnumCoClass1(cast(long, htree), hparent(), 0, "", "")
'   else
'		'print "EnumCoClass1 _2"
'		b = EnumCoClass1(cast(long, htree), hparent(), 1 , s2 & "_Events_Connect", s2 & "_Events_Disconnect")
'	END IF
'   sreport = "'=====   classe(s)" & chr(13, 10) & chr(13, 10) & *b :sysfreestring(cast(BSTR, b))
'	sreport &= s3
'	sreport &= s1
'   showinfo
'	xfull = 0
'End sub

Sub mkFull
   Dim b As zString Ptr
   dim as string s1, s2, s3
	sevent=""
	EnumCoClassB (cast(long, htree), hparent())
	sreport= Ini_templ()
	b = EnumEvent(cast(long, htree), hparent())
	s1 = *b :sysfreestring(cast(BSTR, b))
	s2 = verif(s1)

	'b = EnumModel(cast(long, htree), hparent())
	's3= *b :sysfreestring(cast(BSTR, b))

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

Sub mkModule
   Dim b As zString Ptr
	sevent=""
   b = EnumModule(cast(long, htree), hparent())
   sreport = *b :sysfreestring(cast(BSTR, b))
   showinfo
end sub

Sub kDisp
   Dim b As zString Ptr
	sevent=""
	sreport= Ini_templ()
   b = EnumCoClass1(cast(long, htree), hparent(), 0, "", "")
   sreport &= *b :sysfreestring(cast(BSTR, b))
	if xret1=0 then
		'b = EnumInvoke1(cast(long, htree), hparent())
		'sreport &= *b :sysfreestring(cast(BSTR, b))
	end if
	xret1=0
	xfull=0
   showinfo
End Sub

Sub kDisp2
   Dim b As zString Ptr
   dim as string s1, s2, s3
	sevent=""
	EnumCoClassB (cast(long, htree), hparent())
	sreport= Ini_templ()
	b = EnumEvent(cast(long, htree), hparent())
	s1 = *b :sysfreestring(cast(BSTR, b))
	s2 = verif(s1)

	'b = EnumInvoke1(cast(long, htree), hparent())
	's3= *b :sysfreestring(cast(BSTR, b))

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

'Sub mkDisp2
'	Dim b As zString Ptr
'   dim as string s1, s2, s3
'	xfull = 1
'	b = EnumEvent(cast(long, htree), hparent())
'	s1 = *b :sysfreestring(cast(BSTR, b))
'	s2 = verif(s1)
'	'b = EnumInvoke1(cast(long, htree), hparent())
'	's3= *b :sysfreestring(cast(BSTR, b))
'   if s1 = "" or s2 = "" THEN
'      s1 = ""
'		'print "EnumCoClass1 _1"
'      b = EnumCoClass1(cast(long, htree), hparent(), 0, "", "")
'   else
'		'print "EnumCoClass1 _2"
'		b = EnumCoClass1(cast(long, htree), hparent(), 1 , s2 & "_Events_Connect", s2 & "_Events_Disconnect")
'	END IF
'   sreport = "'=====   classe(s)" & chr(13, 10) & chr(13, 10) & *b :sysfreestring(cast(BSTR, b))
'	sreport &= s3
'	sreport &= s1
'   showinfo
'	xfull = 0
'end sub

Sub mkvtable
   Dim b As zString Ptr
	sevent=""
   b = EnumCoClass(cast(long, htree), hparent())
   sreport = *b :sysfreestring(cast(BSTR, b))
   b = EnumvTable(cast(long, htree), hparent())
   sreport &= *b :sysfreestring(cast(BSTR, b))
   showinfo
End Sub

Sub mkInvoke
   Dim b As zString Ptr
	sevent=""
   b = EnumCoClass(cast(long, htree), hparent())
   sreport = *b :sysfreestring(cast(BSTR, b))
   b = EnumInvoke(cast(long, htree), hparent())
   sreport &= *b :sysfreestring(cast(BSTR, b))
   showinfo
End Sub

Sub mkEvent
   Dim b As zString Ptr
	sevent=""
   b = EnumEvent(cast(long, htree), hparent())
   sreport = *b :sysfreestring(cast(BSTR, b))
   showinfo
End Sub

sub dlgMain_IDC_TAB1_Sample(byval hWin as hwnd, byval CTLID as integer)
   dim hCTLID as hwnd
   dim ts as TCITEM
   'Creating Tab
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

Function dlgMain_Init(Byval hWin as hwnd, Byval wParam as wParam, Byval lParam as lParam) as Integer
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
'	InvalidateRect hmain, ByVal 0, True
'	UpdateWindow hmain
   listview_setcolumnwidth(hlist, 1, LVSCW_AUTOSIZE)

   function = 0
End Function

Function dlgMain_OnSize(Byval hWin as hwnd, Byval lState as long, Byval cx as long, Byval cy as long) as Integer
   movewindow(getdlgitem(hwin, idc_tab1), 0, 0, cx, cy, TRUE)
   movewindow(htab0, 0, 22, cx, cy - 43, TRUE)
   movewindow(htab1, 0, 22, cx, cy - 43, TRUE)
   movewindow(htab2, 0, 22, cx, cy - 43, TRUE)
	movewindow(getdlgitem(hwin, IDC_SBR1), 6 , cy-12, cx-20, 12, TRUE)
   function = 0
End Function

Function dlgMain_OnClose(Byval hWin as hwnd) as Integer
   function = 0
End Function

Function dlgMain_OnDestroy(Byval hWin as hwnd) as Integer
   function = 0
End Function

Function dlgMain_OnNotify(Byval hWin as hwnd, Byval lCtlID as long, Byval lpNMHDR as NMHDR ptr, Byref lresult as long) as Integer
   function = 0
End Function

Function dlgMain_OnTimer(Byval hWin as hwnd, Byval wTimerID as word) as Integer
   function = 0
End Function

Function dlgMain_IDC_TAB1_SelChange(byval hWin as hwnd, byval lpNMHDR as NMHDR ptr, Byref lresult as long) as long
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

Function dlgMain_IDC_TAB1_SelChanging(byval hWin as hwnd, byval lpNMHDR as NMHDR ptr, Byref lresult as long) as long
   dim TabID as Long
   dim hIDC_TAB1 as hwnd
   hIDC_TAB1 = getdlgitem(hWin, IDC_TAB1)
   TabID = SendMessage(hIDC_TAB1, TCM_GETCURSEL, 0, 0)
   select case TabID
      case 0
         'ShowWindow(hTabDlg0,SW_HIDE)
      case 1
         'ShowWindow(hTabDlg1,SW_HIDE)
   end select
   function = 0
End Function

Sub chgtab(ByVal p As Integer)
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

Sub chgtab1()
   treeview_setcheckstate(cast(long, htree), cast(long, TreeView_GetChild(hTree, 0)), 1)
   treeview_changechildstate(cast(long, htree), cast(long, TreeView_GetChild(hTree, 0)), 1)
End Sub

Function dlgMain_IDR_MENU1_menu(byval hwin as hwnd, byval id as long) as Long
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
			'messagebox 0 , "selected pos = " & str(irow),"Warning",mb_ok
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
'			InvalidateRect hmain, ByVal 0, True
'			UpdateWindow hmain
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
         '        if idem_prefix(hWin)= 0 then
         '           r1=force_click(hWin)
         '           if r1=0 THEN exit Function
         '         END IF
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
'		Case code_mkDisp2
'			if idem_prefix(hWin) = 0 then
'            r1 = force_click(hWin)
'            if r1 = 0 THEN exit Function
'         END IF
'         savepath = "Inv-Full_disp"
'         mkDisp2
'         chgtab(2)
'         gene1 = code_mkDisp2
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
'		Case code_full2
'         if idem_prefix(hWin) = 0 then
'            r1 = force_click(hWin)
'            if r1 = 0 THEN exit Function
'         END IF
'         savepath = "vt_Full_disp"
'         mkFull2
'         chgtab(2)
'         gene1 = code_full2
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
'		Case Static1
'			fstatic1="libAx_Disp.a"
'			fstatic2="Static lib  Ax_Disp" & szNULL & "*.a" & szNULL
'			ofn.lStructSize = sizeof(OPENFILENAME)
'			ofn.hwndOwner = hwin
'			ofn.hInstance = cast(hinstance, GetWindowLong(hwin, GWL_HINSTANCE))
'			ofn.lpstrFile = @fstatic1
'			ofn.nMaxFile = 260
'			ofn.lpstrFilter = @fstatic2
'			ofn.Flags = OFN_FILEMUSTEXIST or OFN_HIDEREADONLY or OFN_PATHMUSTEXIST
'			if GetSaveFileName(@ofn) Then fsave=fstatic1
'			'print " fsave "; fsave
'			IF LEN(fsave) Then
'				'print " MakeFile "; fsave
'				MakeFileImport_(inc4, inc4_len, fsave)
'			End If
'		Case Static2
'			fstatic1="Ax_Disp_lib.bi"
'			fstatic2="Include file   Ax_Disp_Lib" & szNULL & "*.bi" & szNULL
'			ofn.lStructSize = sizeof(OPENFILENAME)
'			ofn.hwndOwner = hwin
'			ofn.hInstance = cast(hinstance, GetWindowLong(hwin, GWL_HINSTANCE))
'			ofn.lpstrFile = @fstatic1
'			ofn.nMaxFile = 260
'			ofn.lpstrFilter = @fstatic2
'			ofn.Flags = OFN_FILEMUSTEXIST or OFN_HIDEREADONLY or OFN_PATHMUSTEXIST
'			if GetSaveFileName(@ofn) Then fsave=fstatic1
'			'print " fsave "; fsave
'			IF LEN(fsave) Then
'				'print " MakeFile "; fsave
'				MakeFileImport_(inc3, inc3_len, fsave)
'			End If
'		Case Static3
'			fstatic1="libAx_Lite.a"
'			fstatic2="Static lib  Ax_Lite" & szNULL & "*.a" & szNULL
'			ofn.lStructSize = sizeof(OPENFILENAME)
'			ofn.hwndOwner = hwin
'			ofn.hInstance = cast(hinstance, GetWindowLong(hwin, GWL_HINSTANCE))
'			ofn.lpstrFile = @fstatic1
'			ofn.nMaxFile = 260
'			ofn.lpstrFilter = @fstatic2
'			ofn.Flags = OFN_FILEMUSTEXIST or OFN_HIDEREADONLY or OFN_PATHMUSTEXIST
'			if GetSaveFileName(@ofn) Then fsave=fstatic1
'			'print " fsave "; fsave
'			IF LEN(fsave) Then
'				'print " MakeFile "; fsave
'				MakeFileImport_(inc8, inc8_len, fsave)
'			End If
'		Case Static4
'			fstatic1="Ax_Lite_Lib.bi"
'			fstatic2="Include file   Ax_Lite_Lib" & szNULL & "*.bi" & szNULL
'			ofn.lStructSize = sizeof(OPENFILENAME)
'			ofn.hwndOwner = hwin
'			ofn.hInstance = cast(hinstance, GetWindowLong(hwin, GWL_HINSTANCE))
'			ofn.lpstrFile = @fstatic1
'			ofn.nMaxFile = 260
'			ofn.lpstrFilter = @fstatic2
'			ofn.Flags = OFN_FILEMUSTEXIST or OFN_HIDEREADONLY or OFN_PATHMUSTEXIST
'			if GetSaveFileName(@ofn) Then fsave=fstatic1
'			'print " fsave "; fsave
'			IF LEN(fsave) Then
'				'print " MakeFile "; fsave
'				MakeFileImport_(inc7, inc7_len, fsave)
'			End If
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
			'print " fsave "; fsave
			IF LEN(fsave) Then
				'print " MakeFile "; fsave
				MakeFileImport_(inc9, inc9_len, fsave)
			End If
			exit function
		Case Static6
			mem_ClipboardSetText(inc9, inc9_len)
			exit function
'		Case Static6
'			fstatic1="libAtl.dll.a"
'			fstatic2="Static lib  libAtl.dll.a" & szNULL & "*.a" & szNULL
'			ofn.lStructSize = sizeof(OPENFILENAME)
'			ofn.hwndOwner = hwin
'			ofn.hInstance = cast(hinstance, GetWindowLong(hwin, GWL_HINSTANCE))
'			ofn.lpstrFile = @fstatic1
'			ofn.nMaxFile = 260
'			ofn.lpstrFilter = @fstatic2
'			ofn.Flags = OFN_FILEMUSTEXIST or OFN_HIDEREADONLY or OFN_PATHMUSTEXIST
'			if GetSaveFileName(@ofn) Then fsave=fstatic1
'			'print " fsave "; fsave
'			IF LEN(fsave) Then
'				'print " MakeFile "; fsave
'				MakeFileImport_(inc10, inc10_len, fsave)
'				fsave=left(fsave,len(fsave)-6) & "71.dll.a"
'				MakeFileImport_(inc11, inc11_len, fsave)
'			End If

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
			'print " fsave "; fsave
			IF LEN(fsave) Then
				'print " MakeFile "; fsave
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
			'print " fsave "; fsave
			IF LEN(fsave) Then
				'print " MakeFile "; fsave
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
			'print " fsave "; fsave
			IF LEN(fsave) Then
				'print " MakeFile "; fsave
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
			'print " fsave "; fsave
			IF LEN(fsave) Then
				'print " MakeFile "; fsave
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
'      case As_DispHelper
'         pathfilename = ExePath & "\DispHelper.pdf"
'         shellexecute 0, strptr( "Open"), @pathfilename, NULL, NULL, SW_SHOWNORMAL
   End Select
   function = 0
End Function

Function FF_ListView_GetSelectedItem(ByVal hWndControl As hWnd) As Integer
   ' Do a check to ensure that this is actually a window handle
   If IsWindow(hWndControl) Then
      Function = SendMessage(hWndControl, LVM_GETSELECTIONMARK, 0, 0)
   End If

End Function

Function FF_ListView_GetItemText(ByVal hWndControl As hWnd, ByVal iRow As Integer, ByVal iColumn As Integer ) As String
   Dim tlv_item As LV_ITEM
   Dim zText As ZString * MAX_PATH
   Dim Text1 as string
   ' Do a check to ensure that this is actually a window handle
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

Function FF_ListView_SetSelectedItem(ByVal hWndControl As hWnd, ByVal nIndex As Integer ) As Integer

   ' Do a check to ensure that this is actually a window handle
   If IsWindow(hWndControl) Then
      Function = SendMessage(hWndControl, LVM_SETSELECTIONMARK, 0, nIndex)

      'ensure that the item is set visually as well
      Dim tItem As LVITEM
      tItem.mask = LVIF_STATE
      tItem.iItem = nIndex
      tItem.iSubItem = 0
      tItem.State = LVIS_FOCUSED Or LVIS_SELECTED
      tItem.statemask = LVIS_FOCUSED Or LVIS_SELECTED
      SendMessage hWndControl, LVM_SETITEMSTATE, nIndex, Cast(LPARAM, VarPtr(tItem))
   End If

End Function

Function FF_ListView_GetTopIndex(ByVal hWndControl As hWnd) As Integer

   ' Do a check to ensure that this is actually a window handle
   If IsWindow(hWndControl) Then

      Function = SendMessage(hWndControl, LVM_GETTOPINDEX, 0, 0)

   End If

End Function

Sub FF_Control_Redraw(ByVal hWndControl As hWnd)

   ' Do a check to ensure that this is actually a window handle
   If IsWindow(hWndControl) Then

      ' Signal the Control to be redrawn.
      InvalidateRect hWndControl, ByVal 0, True
      UpdateWindow hWndControl

   End If
End Sub

sub show_line(ret1 As integer)
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



Function getsearch(str1 As String, ini As integer) As integer
   dim x As integer
   dim y As integer
   dim as string mess1, mess2
   'messagebox 0 , "nt1=  " & str(nt1) & "   str1 = " & str1  & "   ini = " & str( ini) , _
   '"HWND_FORM1_LISTVIEW1 = " & str(HWND_FORM1_LISTVIEW1),MB_ICONINFORMATION
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

sub search1()

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


sub next1()

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



Function revsearch(str1 As String, ini As integer) As integer
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


sub previous1()

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

function force_click(hWin as hwnd) as integer
   dim as integer itab0, irow0, i
   dim lret As Long
   function = 0
   itab0 = tab_pos(hWin)
   getwindowtext getdlgitem(hWin, idc_edt1), prefix, 10
   IF LEN(SelPath) Then
'      cleartli
'      lret = LoadtypeLib(stringtobstr(selpath), @ptli)
'      IF lret Then
'         MessageBox BYVAL NULL, "Unable to load " & SelPath, _
'               "Error", MB_OK OR MB_ICONERROR OR MB_TASKMODAL
'         exit function
'      Else
'         LoadTLI
'      End If
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


