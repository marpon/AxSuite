

'Const crlf = Chr$(13, 10)
'Const dq = Chr$(34)
'Const Tb = Chr$(9)

'declare Function FF_ClipboardSetText( ByVal TheText As String ) As Long
'Dim Shared EvtList As String

Function StringToBSTR1(cnv_string As String) As BSTR
    Dim sb As BSTR
    Dim As Integer n
    n = (MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, cnv_string, -1, NULL, 0))-1
    sb=SysAllocStringLen(sb,n)
    MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, cnv_string, -1, sb, n)
    Return sb
End Function

Function GetImage1 (ByVal hWndControl As HWND, ByVal hItem As integer) As Integer
	Dim ti As TV_ITEM
	function= 0
	If IsWindow(hWndControl) Then
		ti.hItem = hItem
		ti.mask = TVIF_HANDLE Or TVIF_IMAGE
		If TreeView_GetItem (hWndControl, @ti ) <> 0 Then
			function= cast(integer,ti.iImage)
		End If
	End If
	print "image"
End Function


Function TreeView_GetItemText1(ByVal hTreeView As Long, ByVal hItem As Long) As String
   Dim ItemText As zstring*32765
	Dim Item As TV_ITEM
	dim as integer gi1= GetImage1(cast (hwnd,hTreeView),hItem)
	if gi1=0 THEN
		function= ""
	else
		Item.hItem = hItem
		Item.Mask = TVIF_TEXT
		Item.cchTextMax = 32765
		Item.pszText = @ItemText
		SendMessage hTreeView, TVM_GETITEM, 0, @Item
		Function = ItemText
	end if
	print "image = " & str(gi1) & "     texte  = " ;ItemText
End Function

Sub putEvtList Cdecl Alias "putEvtList"(evt As String)
   EvtList &= evt & "|"
End Sub

Sub clrEvtList Cdecl Alias "clrEvtList"()
   EvtList = ""
End Sub

Function inEvtList(evt As String) As Long
   For i As Short = 1 To str_numparse(evtlist, "|")
      If evt = str_parse(evtlist, "|", i) Then Return TRUE
   Next
End Function

Function EnumvTable cdecl Alias "EnumvTable"(hTree As Long, hParent() As long) As zString Ptr
   Dim Parents() As Long
   Dim hNextItem As Long
   Dim HasChilds As Long
   Dim hItem as Long
   Dim token As String, s0 As String, tldesc As String
   Dim xface As String, memid As Integer, member As String, offs As Integer

   Dim s as String
   hItem = hparent(tkind_interface)
   tldesc = TreeView_GetItemText1(hTree, treeview_getnextitem(htree, hitem, TVGN_PARENT))
   haschilds = treeview_getchild(htree, hitem)
   If haschilds Then
      s &= "'" & String(80, "=") & crlf
      s &= "'vTable - " & tldesc & crlf
      s &= "'" & String(80, "=") & crlf
      hnextitem = haschilds
      While hnextitem
         token = TreeView_GetItemText1(hTree, hNextItem)
         haschilds = treeview_getchild(htree, hnextitem)
         If haschilds <> 0 And UBound(parents) = 0 Then
            xface = str_parse(token, "?", 1)
            s &= "'" & String(80, "=") & crlf
            s &= "'Interface " & xface & "     ?" & str_parse(token, "?", 3) & crlf
            s &= "Const IID_" & xface & "=" & dq & str_parse(token, "?", 2) & dq & crlf
            s &= "'" & String(80, "=") & crlf
            s &= "Type " & xface & "vTbl" & crlf
         ElseIf (UBound(parents) <> 0) Then
            memid = Val(str_parse(token, "?", 1))
            member = str_parse(token, "?", 3) & "     '" & str_parse(token, "?", 4)

            If offs = 0 Then
               If memid > 8 Then
                  s &= tb & "QueryInterface As Function (Byval pThis As Any ptr,Byref riid As GUID,Byref ppvObj As Dword) As hResult" & crlf
                  s &= tb & "AddRef As Function (Byval pThis As Any ptr) As hResult" & crlf
                  s &= tb & "Release As Function (Byval pThis As Any ptr) As hResult" & crlf
                  offs = 8
               End If
               If memid > 24 Then
                  s &= tb & "GetTypeInfoCount As Function (Byval pThis As Any ptr,Byref pctinfo As Uinteger) As hResult" & crlf
                  s &= tb & "GetTypeInfo As Function (Byval pThis As Any ptr,Byval itinfo As Uinteger,Byval lcid As Uinteger,Byref pptinfo As Any Ptr) As hResult" & crlf
                  s &= tb & "GetIDsOfNames As Function (Byval pThis As Any ptr,Byval riid As GUID,Byval rgszNames As Byte,Byval cNames As Uinteger,Byval lcid As Uinteger,Byref rgdispid As Integer) As hResult" & crlf
                  s &= tb & "Invoke As Function (Byval pThis As Any ptr,Byval dispidMember As Integer,Byval riid As GUID,Byval lcid As Uinteger,Byval wFlags As Ushort,Byval pdispparams As DISPPARAMS,Byref pvarResult As Variant,Byref pexcepinfo As EXCEPINFO,Byref puArgErr As Uinteger) As hResult" & crlf
                  offs = 24
               End If
            End If
            s0 = ""
            offs = memid - offs - 4
            If offs > 0 Then s0 = "(" &(offs - 1) & ") As Byte" & crlf
            offs = memid
            If Len(s0) Then s &= tb & "Offset" & offs & s0

            If InStr(member, "()") Then
               s &= tb & str_parse(member, "(", 1) & "(pThis As Any Ptr)" & crlf
            Else
               s &= tb & str_parse(member, "(", 1) & "(pThis As Any Ptr," & Mid(member, Len(str_parse(member, "(", 1)) + 2, - 1) & crlf
            End If

         End If

         If haschilds Then
            ReDim Preserve parents(UBound(parents) + 1)
            parents(UBound(parents)) = hnextitem
            hnextitem = haschilds
         Else
            hnextitem = treeview_getnextitem(htree, hnextitem, TVGN_NEXT)
            If hnextitem = 0 And UBound(parents) > 0 Then
               hnextitem = treeview_getnextitem(htree, parents(UBound(parents)), TVGN_NEXT)
               ReDim Preserve parents(UBound(parents) - 1)
               s &= "End Type" & crlf & crlf
               s &= "Type " & xface & "_" & crlf
               s &= tb & "lpvtbl As " & xface & "vTbl Ptr" & crlf
               s &= "End Type" & crlf & crlf
               offs = 0
            End If
         End If
      Wend
   End If
   Return sysallocstringbytelen(s, len(s))
End Function

Function EnumInvoke cdecl Alias "EnumInvoke"(hTree As Long, hParent() As long) As zString Ptr
   Dim Parents() As Long
   Dim hNextItem As Long
   Dim HasChilds As Long

   Dim token As String, tldesc As String
   Dim hItem as Long
   Dim xface As String, memid As String, member As String
   Dim As Integer cpar

   Dim s as String
   hItem = hparent(tkind_dispatch)
   tldesc = TreeView_GetItemText1(hTree, treeview_getnextitem(htree, hitem, TVGN_PARENT))
   haschilds = treeview_getchild(htree, hitem)
   If haschilds Then
      s &= "'" & String(80, "=") & crlf
      s &= "'Invoke - " & tldesc & crlf
      s &= "'" & String(80, "=") & crlf
      hnextitem = haschilds
      While hnextitem
         token = TreeView_GetItemText1(hTree, hNextItem)
         haschilds = treeview_getchild(htree, hnextitem)
         If (haschilds <> 0) And (UBound(parents) = 0) Then
            xface = str_parse(token, "?", 1)
            s &= "'" & String(80, "=") & crlf
            s &= "'Dispatch " & xface & "     ?" & str_parse(token, "?", 3) & crlf
            s &= "Const IID_" & xface & "=" & dq & str_parse(token, "?", 2) & dq & crlf
            s &= "'" & String(80, "=") & crlf
            s &= "Type " & xface & crlf
         ElseIf (UBound(parents) <> 0) then
            memid = str_parse(token, "?", 1)
            member = str_parse(token, "?", 3)
            cpar = str_numparse(member, ",")
            If Len(str_parse(str_parse(member, "(", 2), ")", 1)) = 0 Then cpar = 0
			s &=tb & str_parse(member," As ",1) & " As tMember/'"
			s &=str_remove(member,str_parse(member," As ",1)) & "'/=(" & memid & ",2," & cpar & "," & str_parse(token,"?",2) & ")" & crlf
         End If

         If haschilds Then
            ReDim Preserve parents(UBound(parents) + 1)
            parents(UBound(parents)) = hnextitem
            hnextitem = haschilds
         Else
            hnextitem = treeview_getnextitem(htree, hnextitem, TVGN_NEXT)
            If hnextitem = 0 And UBound(parents) > 0 Then
               hnextitem = treeview_getnextitem(htree, parents(UBound(parents)), TVGN_NEXT)
               ReDim Preserve parents(UBound(parents) - 1)
               s &= "End Type" & crlf & crlf
            End If
         End If
      Wend
   End If
   Return sysallocstringbytelen(s, len(s))
End Function

Function EnumEvent cdecl Alias "EnumEvent"(hTree As Long, hParent() As long) As zString Ptr
   Dim Parents() As Long
   Dim hNextItem As Long
   Dim HasChilds As Long

   Dim token As String, tldesc As String
   Dim hItem as Long
   Dim As String xface, memid, member, par, help, temp
   Dim As Integer cpar, bevent

   Dim s as String
   hItem = hparent(tkind_dispatch)
   tldesc = TreeView_GetItemText1(hTree, treeview_getnextitem(htree, hitem, TVGN_PARENT))
   haschilds = treeview_getchild(htree, hitem)
   If haschilds Then
      s &= "'" & String(80, "=") & crlf
      s &= "'Event - " & tldesc & crlf
      s &= "'" & String(80, "=") & crlf
      hnextitem = haschilds
      While hnextitem
         token = TreeView_GetItemText1(hTree, hNextItem)
         haschilds = treeview_getchild(htree, hnextitem)
         If (haschilds <> 0) And (UBound(parents) = 0) Then
            xface = str_parse(token, "?", 1)
            bevent = inevtlist(xface)
            If bevent Then
               s &= "'" & String(80, "=") & crlf
               s &= "'Event Dispatch - " & xface & "     ?" & str_parse(token, "?", 3) & crlf
               s &= "Const " & xface & "_IID_CP" & "=" & dq & str_parse(token, "?", 2) & dq & crlf
               s &= "'" & String(80, "=") & crlf
            End If
         ElseIf (UBound(parents) <> 0) then
            memid = str_parse(token, "?", 1)
            member = str_parse(token, "?", 3)
            par = str_parse(str_parse(member, "(", 2), ")", 1)
            member = str_parse(member, " As", 1)
            help = str_parse(token, "?", 4)
            cpar = str_numparse(par, ",")
            If Len(par) = 0 Then cpar = 0
            If bevent Then
               s &= "'" & String(80, "=") & crlf
               s &= "'Event=" & member & ", ID=" & memid & " ?" & help & crlf
               s &= "'" & String(80, "=") & crlf
               s &= "Function " & xface & "_" & member & "(Byval pCookie AS Events_IDispatchVtbl ptr, Byval pdispparams As DISPPARAMS ptr) As HRESULT" & crlf
               s &= tb & "Dim pthis As Dword=>pCookie->pthis" & crlf
               temp &= tb & tb & tb & "Case " & memid & " '" & help & crlf
               temp &= tb & tb & tb & tb & "Function=" & xface & "_" & member & "(pUnk, pdispparams)" & crlf
               If cpar Then s &= tb & "Dim pv As Variant Ptr" & crlf
               For i As Short = 1 To cpar
                  s &= tb & "Dim " & str_parse(par, ",", i) & crlf
               Next
               s &= crlf
               For i As Short = 1 To cpar
                  s &= tb & str_parse(str_parse(par, ",", i), " As", 1) & "=Variantv(pv[" & cpar - i & "])" & crlf
               Next
               s &= tb & "'*** Put your code here ***" & crlf & crlf
               s &= tb & "Function=S_OK" & crlf
               s &= "End Function" & crlf & crlf
            End If
         End If
         If haschilds Then
            ReDim Preserve parents(UBound(parents) + 1)
            parents(UBound(parents)) = hnextitem
            hnextitem = haschilds
         Else
            hnextitem = treeview_getnextitem(htree, hnextitem, TVGN_NEXT)
            If hnextitem = 0 And UBound(parents) > 0 Then
               hnextitem = treeview_getnextitem(htree, parents(UBound(parents)), TVGN_NEXT)
               ReDim Preserve parents(UBound(parents) - 1)
               If bevent Then
                  temp = "Function " & xface & "_IDispatch_Invoke(Byval pUnk As IDispatch ptr, Byval dispidMember As DispID, Byval riid As IID ptr, _" & crlf & _
                        "  Byval lcid As LCID, Byval wFlags As Ushort, Byval pdispparams As DISPPARAMS ptr, Byval pvarResult AS Variant ptr, _" & crlf & _
                        "  Byval pexcepinfo As EXCEPINFO ptr, Byval puArgErr AS Uint ptr) As Hresult" & crlf & crlf & _
                        tb & "If Varptr(pdispparams) Then" & crlf & _
                        tb & tb & "Select Case dispidMember" & crlf & _
                        temp & _
                        tb & tb & tb & "Case Else" & crlf & _
                        tb & tb & tb & tb & "Function = DISP_E_MEMBERNOTFOUND" & crlf & _
                        tb & tb & "End Select" & crlf & _
                        tb & "End If" & crlf & _
                        "End Function" & crlf & crlf
                  s &= temp
                  temp = ""
                  s &= "Function " & xface & "_IDispatch_QueryInterface (byval pCookie as Events_IDispatchVtbl PTR, byval riid as IID ptr, byval ppVObj as PVOID ptr) as HRESULT" & crlf
                  s &= tb & "Dim riids As lpolestr" & crlf & crlf
                  s &= tb & "StringFromIID(riid,@riids)" & crlf
                  s &= tb & "If (*riids=" & dq & "{00000000-0000-0000-C000-000000000046}" & dq & ") OR _" & crlf
                  s &= tb & "   (*riids=" & dq & "{00020400-0000-0000-C000-000000000046}" & dq & ") OR _" & crlf
                  s &= tb & "   (*riids=" & xface & "_IID_CP) then" & crlf
                  s &= tb & tb & "*ppvObj = pCookie" & crlf
                  s &= tb & tb & "Events_AddRef(pCookie)" & crlf
                  s &= tb & tb & "Function = S_OK" & crlf
                  s &= tb & "Else" & crlf
                  s &= tb & tb & "*ppvObj = NULL" & crlf
                  s &= tb & tb & "Function = E_NOINTERFACE" & crlf
                  s &= tb & "End IF" & crlf
                  s &= "End Function" & crlf & crlf
                  s &= "' --------------------------------------------------------------------------------------------" & CRLF
                  s &= "' " & xface & " Events Connection function" & crlf
                  s &= "' --------------------------------------------------------------------------------------------" & CRLF
                  s &= "Function " & xface & "_Events_Connect (BYVAL pUnk AS IConnectionPointContainer ptr,dwcookie As dword) AS DWORD" & crlf
                  s &= tb & "Dim pCPC AS IConnectionPointContainer ptr" & crlf
                  s &= tb & "Dim pCP AS IConnectionPoint ptr" & crlf
                  s &= tb & "Dim IID_CPC AS IID      ' IID_IConnectionPointContainer" & crlf
                  s &= tb & "Dim IID_CP AS IID       ' IID_IConnectionPoint" & crlf
                  s &= tb & "Dim pSink AS lpunknown  ' Pointer to our sink interface" & crlf
                  s &= tb & "Dim pdwCookie AS DWORD  ' Returned token" & crlf
                  s &= crlf
                  s &= tb & "IIDFromString(" & dq & "{B196B284-BAB4-101A-B69C-00AA00341D07}" & dq & ",@IID_CPC)" & crlf
                  s &= tb & "IIDFromString(" & xface & "_IID_CP,@IID_CP)" & crlf
                  s &= tb & "AxScode = IConnectionPointContainer_QueryInterface(pUnk, @IID_CPC, @pCPC)" & crlf
                  s &= tb & "If AxScode <> S_OK THEN Function = AxScode:Exit Function" & crlf
                  s &= tb & "AxScode = IConnectionPointContainer_FindConnectionPoint(pCPC, @IID_CP, @pCP)" & crlf
                  s &= tb & "IConnectionPointContainer_release(pCPC)" & crlf
                  s &= tb & "If AxScode <> S_OK Then Function = AxScode:Exit Function" & crlf
                  s &= tb & "psink=Events_BuildVTbl(punk,ProcPtr(" & xface & "_IDispatch_QueryInterface), ProcPtr(" & xface & "_IDispatch_Invoke))" & crlf
                  s &= tb & "AxScode = IConnectionPoint_Advise(pCP, pSink, @pdwCookie)" & crlf
                  s &= tb & "dwcookie=pdwcookie" & crlf
                  s &= tb & "IConnectionPoint_Release(pCP)" & crlf
                  s &= tb & "Function = AxScode" & crlf
                  s &= "End Function" & crlf
                  s &= crlf
                  s &= "Function " & xface & "_Events_Disconnect (BYVAL pUnk AS IConnectionPointContainer ptr, byval dwcookie As dword) AS LONG" & crlf
                  s &= tb & "Dim pCPC AS IConnectionPointContainer ptr" & crlf
                  s &= tb & "Dim pCP AS IConnectionPoint ptr" & crlf
                  s &= tb & "Dim IID_CPC AS IID    ' IID_IConnectionPointContainer" & crlf
                  s &= tb & "Dim IID_CP AS IID     ' IID_IConnectionPoint" & crlf
                  s &= crlf
                  s &= tb & "If pUnk = 0 Then Function = E_POINTER : Exit Function" & crlf
                  s &= tb & "IIDFromString(" & dq & "{B196B284-BAB4-101A-B69C-00AA00341D07}" & dq & ",@IID_CPC)" & crlf
                  s &= tb & "IIDFromString(" & xface & "_IID_CP,@IID_CP)" & crlf
                  s &= tb & "AxScode = IConnectionPointContainer_QueryInterface(pUnk, @IID_CPC, @pCPC)" & crlf
                  s &= tb & "If AxScode <> S_OK Then Function = AxScode : Exit Function" & crlf
                  s &= tb & "AxScode = IConnectionPointContainer_FindConnectionPoint(pCPC, @IID_CP, @pCP)" & crlf
                  s &= tb & "IConnectionPointContainer_release(pCPC)" & crlf
                  s &= tb & "If AxScode <> S_OK Then Function = AxScode : Exit Function" & crlf
                  s &= tb & "AxScode = IConnectionPoint_Unadvise(pCP, dwcookie)" & crlf
                  s &= tb & "IConnectionPoint_Release(pCP)" & crlf
                  s &= tb & "Function = AxScode" & crlf
                  s &= "End Function" & crlf & crlf
               End If
            End If
         End If
      Wend
   End If
   Return sysallocstringbytelen(s, len(s))
End Function

Function EnumEnum cdecl Alias "EnumEnum"(hTree As Long, hParent() As long) As zString Ptr
   Dim Parents() As Long
   Dim hNextItem As Long
   Dim HasChilds As Long

   Dim token As String, tldesc As String
   Dim hItem as Long

   Dim s as String
   hItem = hparent(tkind_enum)
   tldesc = TreeView_GetItemText1(hTree, treeview_getnextitem(htree, hitem, TVGN_PARENT))
   haschilds = treeview_getchild(htree, hitem)
   If haschilds Then
      s &= "'" & String(80, "=") & crlf
      s &= "Enum -  " & tldesc & crlf
      s &= "'" & String(80, "=") & crlf
      hnextitem = haschilds
      While hnextitem
         token = TreeView_GetItemText1(hTree, hNextItem)
         haschilds = treeview_getchild(htree, hnextitem)
         If UBound(parents) = 0 Then
            s &= "'" & String(80, "=") & crlf
            s &= "Type " & str_parse(token, "?", 1) & " As Integer     '" & str_parse(token, "?", 2) & crlf
            s &= "'" & String(80, "=") & crlf
         Else
            s &= str_parse(token, "?", 1) & "     '" & str_parse(token, "?", 1) & crlf
         End If
         If haschilds Then
            ReDim Preserve parents(UBound(parents) + 1)
            parents(UBound(parents)) = hnextitem
            hnextitem = haschilds
         Else
            hnextitem = treeview_getnextitem(htree, hnextitem, TVGN_NEXT)
            If hnextitem = 0 And UBound(parents) > 0 Then
               hnextitem = treeview_getnextitem(htree, parents(UBound(parents)), TVGN_NEXT)
               ReDim Preserve parents(UBound(parents) - 1)
               s &= crlf
            End If
         end if
      Wend
   End If
   Return sysallocstringbytelen(s, len(s))
End Function

Function EnumRecord cdecl Alias "EnumRecord"(hTree As Long, hParent() As long) As zString Ptr
   Dim Parents() As Long
   Dim hNextItem As Long
   Dim HasChilds As Long
   Dim token As String
   Dim hItem as Long
   Dim s as String

   hItem = hparent(tkind_record)
   haschilds = treeview_getchild(htree, hitem)
   If haschilds Then
      hnextitem = haschilds
      While hnextitem
         token = TreeView_GetItemText1(hTree, hNextItem)
         haschilds = treeview_getchild(htree, hnextitem)
         If haschilds <> 0 And UBound(parents) = 0 Then
            s &= "'" & String(80, "=") & crlf
            s &= "'Record " & str_parse(token, "?", 1) & "     '" & str_parse(token, "?", 2) & crlf
            s &= "'" & String(80, "=") & crlf
            s &= "Type " & str_parse(token, "?", 1) & crlf
         ElseIf UBound(parents) <> 0 then
            s &= tb & token & crlf
         End If

         If haschilds Then
            ReDim Preserve parents(UBound(parents) + 1)
            parents(UBound(parents)) = hnextitem
            hnextitem = haschilds
         Else
            hnextitem = treeview_getnextitem(htree, hnextitem, TVGN_NEXT)
            If hnextitem = 0 And UBound(parents) > 0 Then
               hnextitem = treeview_getnextitem(htree, parents(UBound(parents)), TVGN_NEXT)
               ReDim Preserve parents(UBound(parents) - 1)
               s &= "End Type" & crlf
            End If
         End If
      Wend
   End If
   Return sysallocstringbytelen(s, len(s))
End Function

Function EnumUnion cdecl Alias "EnumUnion"(hTree As Long, hParent() As long) As zString Ptr
   Dim Parents() As Long
   Dim hNextItem As Long
   Dim HasChilds As Long
   Dim token As String
   Dim hItem as Long

   Dim s as String
   hItem = hparent(tkind_union)
   haschilds = treeview_getchild(htree, hitem)
   If haschilds Then
      hnextitem = haschilds
      While hnextitem
         token = TreeView_GetItemText1(hTree, hNextItem)
         haschilds = treeview_getchild(htree, hnextitem)
         If haschilds <> 0 And UBound(parents) = 0 Then
            s &= "'" & String(80, "=") & crlf
            s &= "'Union " & str_parse(token, "?", 1) & "     '" & str_parse(token, "?", 2) & crlf
            s &= "'" & String(80, "=") & crlf
            s &= "Union " & str_parse(token, "?", 1) & crlf
         Else
            s &= tb & token & crlf
         End If

         If haschilds Then
            ReDim Preserve parents(UBound(parents) + 1)
            parents(UBound(parents)) = hnextitem
            hnextitem = haschilds
         Else
            hnextitem = treeview_getnextitem(htree, hnextitem, TVGN_NEXT)
            If hnextitem = 0 And UBound(parents) > 0 Then
               hnextitem = treeview_getnextitem(htree, parents(UBound(parents)), TVGN_NEXT)
               ReDim Preserve parents(UBound(parents) - 1)
               s &= "End Union" & crlf
            End If
         End If
      Wend
   End If
   Return sysallocstringbytelen(s, len(s))
End Function

Function EnumModule cdecl Alias "EnumModule"(hTree As Long, hParent() As long) As zString Ptr
   Dim Parents() As Long
   Dim hNextItem As Long
   Dim HasChilds As Long
   Dim token As String, tldesc As String
   Dim hItem as Long

   Dim s as String
   hItem = hparent(tkind_module)
   tldesc = TreeView_GetItemText1(hTree, treeview_getnextitem(htree, hitem, TVGN_PARENT))
   haschilds = treeview_getchild(htree, hitem)
   If haschilds Then
      s &= "'" & String(80, "=") & crlf
      s &= "'Module - " & tldesc & crlf
      s &= "'" & String(80, "=") & crlf
      hnextitem = haschilds
      While hnextitem

         token = TreeView_GetItemText1(hTree, hNextItem)
         If UBound(parents) = 0 Then
            s &= "'" & String(80, "=") & crlf
            s &= "'" & str_parse(token, "?", 1) & "     ?" & str_parse(token, "?", 2) & crlf
            s &= "'" & String(80, "=") & crlf
         Else
            s &= str_parse(token, "?", 1) & "     '" & str_parse(token, "?", 2) & crlf
         End If

         haschilds = treeview_getchild(htree, hnextitem)
         If haschilds Then
            ReDim Preserve parents(UBound(parents) + 1)
            parents(UBound(parents)) = hnextitem
            hnextitem = haschilds
         Else
            hnextitem = treeview_getnextitem(htree, hnextitem, TVGN_NEXT)
            If hnextitem = 0 And UBound(parents) > 0 Then
               hnextitem = treeview_getnextitem(htree, parents(UBound(parents)), TVGN_NEXT)
               ReDim Preserve parents(UBound(parents) - 1)
               s &= crlf
            End If
         end if
      Wend
   End If
   Return sysallocstringbytelen(s, len(s))
End Function


Function EnumAlias cdecl Alias "EnumAlias"(hTree As Long, hParent() As long) As zString Ptr
   Dim Parents() As Long
   Dim hNextItem As Long
   Dim HasChilds As Long
   Dim token As String
   Dim hItem as Long
   Dim as String s, tldesc

   hItem = hparent(tkind_alias)
   tldesc = TreeView_GetItemText1(hTree, treeview_getnextitem(htree, hitem, TVGN_PARENT))
   haschilds = treeview_getchild(htree, hitem)
   If haschilds Then
      s &= "'" & String(80, "=") & crlf
      s &= "'Alias - " & tldesc & crlf
      s &= "'" & String(80, "=") & crlf
      hnextitem = haschilds
      While hnextitem
         token = TreeView_GetItemText1(hTree, hNextItem)
         s &= token & crlf
         hnextitem = treeview_getnextitem(htree, hnextitem, TVGN_NEXT)
      Wend
   End If
   Return sysallocstringbytelen(s, len(s))
End Function

Function EnumCoClass cdecl Alias "EnumCoClass"(hTree As Long, hParent() As long) As zString Ptr
   Dim Parents() As Long
   Dim hNextItem As Long
   Dim HasChilds As Long
   Dim token As String, progid As WString Ptr, clssid As clsid
   Dim As String clsids, progids, tldesc
   Dim hItem as long

   Dim s as string
   hItem = hparent(tkind_coclass)
   tldesc = TreeView_GetItemText1(hTree, treeview_getnextitem(htree, hitem, TVGN_PARENT))
   haschilds = treeview_getchild(htree, hitem)
   If haschilds Then
      hnextitem = haschilds
      While hnextitem
         token = TreeView_GetItemText1(hTree, hNextItem)
         clsids &= "Const CLSID_" & str_parse(token, "?", 1) & "=" & dq & str_parse(token, "?", 2) & dq & "     '" & str_parse(token, "?", 3) & crlf
         progid = Stringtobstr1(str_parse(token, "?", 2))
         clsidfromString(progid, @clssid)
         sysfreeString(progid)
         progidfromclsid(@clssid, @progid)
         progids &= "Const ProgID_" & str_parse(token, "?", 1) & "=" & dq & *progid & dq & crlf
         sysfreeString(progid)
         hnextitem = treeview_getnextitem(htree, hnextitem, TVGN_NEXT)
      Wend
   End If
   s &= "'" & String(80, "=") & crlf
   s &= "'CLSID - " & tldesc & crlf
   s &= "'" & String(80, "=") & crlf
   s &= clsids & crlf
   s &= "'" & String(80, "=") & crlf
   s &= "'ProgID - " & tldesc & crlf
   s &= "'" & String(80, "=") & crlf
   s &= progids & crlf
   Return SysAllocStringByteLen(s, len(s))
End Function


Function EnumCoClass1 cdecl Alias "EnumCoClass1"(hTree As Long, hParent() As long) As zString Ptr
   Dim Parents() As Long
   Dim hNextItem As Long
   Dim HasChilds As Long
   Dim token As String, progid As WString Ptr, clssid As clsid
   Dim As String clsids, progids, tldesc, ret1
   Dim hItem as long
	Dim nitem as long

   Dim s as string
   hItem = hparent(tkind_coclass)
   tldesc = TreeView_GetItemText1(hTree, treeview_getnextitem(htree, hitem, TVGN_PARENT))
   haschilds = treeview_getchild(htree, hitem)
   If haschilds Then
      hnextitem = haschilds

      While hnextitem
         token = TreeView_GetItemText1(hTree, hNextItem)
         clsids &= "Const CLSID_" & str_parse(token, "?", 1) & "=" & dq & str_parse(token, "?", 2) & dq & "     '" & str_parse(token, "?", 3) & crlf
         progid = Stringtobstr1(str_parse(token, "?", 2))
         clsidfromString(progid, @clssid)
         sysfreeString(progid)
         progidfromclsid(@clssid, @progid)
         progids &= "Const ProgID_" & str_parse(token, "?", 1) & "=" & dq & *progid & dq & crlf
			ret1= *progid
         sysfreeString(progid)
         hnextitem = treeview_getnextitem(htree, hnextitem, TVGN_NEXT)
			nitem +=1
      Wend
   End If
'   s &= "'" & String(80, "=") & crlf
'   s &= "'CLSID - " & tldesc & crlf
'   s &= "'" & String(80, "=") & crlf
'   s &= clsids & crlf
   s &= "'" & String(80, "=") & crlf
   s &= "'ProgID - " & tldesc & crlf
   s &= "'" & String(80, "=") & crlf
   s &= progids & crlf
	if nitem =1 then
		FF_ClipboardSetText( ret1)
	else
		FF_ClipboardSetText( "")
	endif
   Return SysAllocStringByteLen(s, len(s))

End Function


