sub GetPrefix()
	sendmessage getdlgitem(hmain, idc_edt1), wm_gettext, 10, cast(lparam, @prefix)
END sub

Function TreeView_GetCheckState1( ByVal hWnd As Long, ByVal hItem As long ) As uinteger
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


Function TreeView_GetItemText1(ByVal hTreeView As Long, ByVal hItem As Long) As String
   Dim ItemText As zstring*32765
	Dim Item As TV_ITEM
	dim as integer gi1= TreeView_GetCheckState1(hTreeView,hItem)
	'dim as integer gi2= IsTVItemChecked(hTreeView,hItem)
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
	'print "State = " & str(gi1) & crlf & ">" ;ItemText :print "State2 = " & str(gi2) :print
End Function

'function trait(st2 as string, ar() as string) as integer
'   dim as integer x,n0
'   dim as string s1, s2,st1,s3
'	st1=rtrim (st2,any "' ")
'	'st1=left(st1,len(st1)-1)
'	st1=Str_removeany(st1, "0123456789=-")
'	st1=str_remove(st1, "Type(,,,,)")


'	'print st1
'   dim i1 as integer = str_numparse(st1, ",")
'   'print "st1",st1
'   'print "i1",i1
'   redim ar(1 to i1, 1 to 2) As string
'   x=0
'   for n0 = 1 to i1
'      s1 = str_parse(st1, ",", n0)
'		s1=rtrim(s1,any "') ")
'		if s1 <> "" and instr(s1, " As ") > 0 THEN
'			x += 1
'			s2 = trim(str_parse(ucase(s1), " AS ", 1))
'			if s2 = "" THEN s2 = "ANY_VAR"
'			ar(x, 2) = s2
'			s3=trim(str_parse(ucase(s1), " AS ", 2))
'			select CASE s3
'				CASE "LONG", "INTEGER", "INTEGER PTR", "LONG PTR"
'					ar(x, 1) = "%d"
'				CASE "ULONG", "UINTEGER", "ULONG PTR", "UINTEGER PTR"
'					ar(x, 1) = "%u"
'				CASE "SHORT", "SHORT PTR"
'					ar(x, 1) = "%i"
'				CASE "USHORT", "USHORT PTR"
'					ar(x, 1) = "%I"
'				CASE "LONGINT", "LONGINT PTR"
'					ar(x, 1) = "%z"
'				CASE "ULONGINT", "ULONGINT PTR"
'					ar(x, 1) = "%Z"
'				CASE "BSTR", "BSTR PTR"
'					ar(x, 1) = "%B"
'				CASE "STRING"
'					ar(x, 1) = "%s"
'				CASE "ZSTRING", "ZSTRING PTR"
'					ar(x, 1) = "%s"
'				CASE "WSTRING", "WSTRING PTR"
'					ar(x, 1) = "%S"
'				CASE "DOUBLE", "DOUBLE PTR" ,"FLOAT","FLOAT PTR","SINGLE","SINGLE PTR"
'					ar(x, 1) = "%e"
'				CASE "BOOL", "BOOL PTR"
'					ar(x, 1) = "%b"
'				CASE "VARIANT", "VARIANT PTR"
'					ar(x, 1) = "%v"
'				CASE "LPDISPATCH", "LPDISPATCH PTR", "IDISPATCH", "IDISPATCH PTR"
'					ar(x, 1) = "%o"
'				CASE "LPUNKNOWN", "LPUNKNOWN PTR", "IUNKNOWN PTR", "IUNKNOWN"
'					ar(x, 1) = "%0"
'	'         CASE "DATE_", "DATE_ PTR"
'	'            ar(x, 1) = "%t"
'				CASE "DATE", "DATE PTR"
'					ar(x, 1) = "%D"
'				CASE ELSE
'					ar(x, 1) = "%p"
'			END SELECT
'			if right(s3,4)=" PTR" THEN
'				s3=left(s3,len(s3)-4)
'				ar(x, 2) ="@"& ar(x, 2)
'				if right(s3,4)=" PTR" THEN ar(x, 2) ="@"& ar(x, 2)
'         END IF
'		end if
'   NEXT

'   'print "fin",i1
'   function = x
'END function

Sub putEvtList Cdecl Alias "putEvtList"(evt As String)
   EvtList &= evt & "|"
	'print EvtList
End Sub

Sub clrEvtList Cdecl Alias "clrEvtList"()
   EvtList = ""
End Sub

Function inEvtList(evt As String) As Long
   For i As Short = 1 To str_numparse(evtlist, "|")
      If evt = str_parse(evtlist, "|", i) Then Return TRUE
   Next
End Function

sub inEvtList2()
   For i As Short = 1 To str_numparse(evtlist, "|")
      'print str(i) & "  :  " &  str_parse(evtlist, "|", i)
   Next
End sub

'Function EnumInvoke1 cdecl (hTree As Long, hParent() As long) As zString Ptr
'   Dim Parents() As Long
'   Dim hNextItem As Long
'   Dim HasChilds As Long

'   Dim token As String, tldesc As String
'   Dim hItem as Long
'   Dim xface As String, memid As String, member As String
'   Dim As Integer cpar
'   '  Dim s22 as String
'   Dim s as String,niveau1 as string,debut as string,niv3 as string
'	dim as string str1
'	dim inot as integer,icall as integer
'	Dim s1 as String, s2 as string, prop as string,su as string ,sf as string
'   dim i1 as integer, x1 as integer, x91 as integer
'   Dim ar() As string
'	GetPrefix()
'	str1 = trim(prefix, any "-_ ")

'   hItem = hparent(tkind_dispatch)
'   tldesc =  TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem, TVGN_PARENT)))
'   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
'   If haschilds Then
'      niveau1 = "'" & String(80, "=") & crlf
'      niveau1 &= "'Invoke - " & tldesc & crlf
'      niveau1 &= "'" & String(80, "=") & crlf
'		s=""
'      hnextitem = haschilds
'      While hnextitem
'         token = TreeView_GetItemText1(hTree, hNextItem)
'			if instr(token,"          'NOT'") THEN
'				token=left(token , len(token)-len("          'NOT'"))
'				inot=1
'			else
'				inot=0
'			end if
'         haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hnextitem))
'         If (haschilds <> 0) And (UBound(parents) = 0) Then
'            xface = str1 & str_parse(token, "?", 1)
'            debut = "'" & String(80, "=") & crlf
'            debut &= "'Dispatch " & xface & "     ?" & str_parse(token, "?", 3) & crlf
'            debut &= "'" & String(80, "=") & crlf
'            debut &= d_bc & "   Begin Model List : Disphelper syntax for  " & xface & crlf & crlf
'				niv3=""
'         ElseIf (UBound(parents) <> 0) and inot=0 then
'            memid = str_parse(token, "?", 1)
'            member = str_parse(token, "?", 3)
'				member = FF_Replace( member , "Const ", "")
'				member=rtrim(member,any"'= " & dq)
'				if instr(member,")")=0 and instr(member,"(")THEN member=member & ")"
'				member = FF_Replace( member ,")As ",") As ")
'				member = FF_Replace( member ," PTR"," Ptr")
'				member = FF_Replace( member ," ptr"," Ptr")
'				sf=""
'				if right(member,12)= ") As any Ptr" then
'					sf=" As any Ptr"
'					member=left(member,len(member)- 11)
'				end if
'				member = FF_Replace( member , " As VARIANT", " As VARIANT Ptr")
'				member = FF_Replace( member , " As VARIANT Ptr Ptr", " As VARIANT Ptr")
'				i1=instr(ucase(member),"AS FUNCTION(")
'				if i1>0 THEN member = left(member,i1 -1) & mid(member,i1 + 11)
'				i1=instr(ucase(member),"AS SUB(")
'				if i1>0 THEN member = left(member,i1 -1) & mid(member,i1 + 6)
'				if instr(member,"QueryInterface (") =0 and instr(member,"AddRef (") =0 and _
'						instr(member,"Release (") =0 and instr(member,"GetTypeInfoCount (") =0 and _
'						instr(member,"GetTypeInfo (") =0 and instr(member,"GetIDsOfNames (") =0 and _
'						instr(member,"Invoke (") =0  then
'					'print member
'					prop=trim(str_parse(member, "(", 1),any " '")
'					'print prop
'					member = "(" & str_parse(member, "(", 2)
'					icall=0
'					If InStr(member, "()") Then icall=1
'					if left(ucase(prop),3)="SET" or left(ucase(prop),3)="PUT" THEN
'						if sf="" THEN
'							sf=str_parse(member, ")", 2)
'							sf = FF_Replace( sf , " As VARIANT Ptr", " As VARIANT")
'							if trim (sf,any " '" )<>"" and instr(sf,"As ") THEN
'								if icall=1 THEN
'									member=str_parse(member, ")", 1) &  sf & " Ptr)"
'									'member=FF_Replace( member ,"()","(" & sf & " Ptr)" )
'									icall=0
'								else
'									member=str_parse(member, ")", 1) & "," &  sf & " Ptr)"
'									'member=FF_Replace( member ,")","," & sf & " Ptr)" )
'                        END IF
'							else
'								member=str_parse(member, ")", 1) & ")"
'                     END IF
'						else
'							member=str_parse(member, ")", 1) & ")"
'                  END IF
'						'print "set or put "; member
'						su=left(ucase(prop),3)
'						prop=mid(prop,4)
'					elseif left(ucase(prop),3)="GET" THEN
'						su="GET"
'						prop=mid(prop,4)
'						sf=str_parse(member, ")", 2)
'						sf = FF_Replace( sf , " As VARIANT Ptr", " As VARIANT")
'						if icall=1 THEN
'							member=str_parse(member, ")", 1) &  sf & " Ptr)"
'							'member=FF_Replace( member ,"()","(" & sf & " Ptr)" )
'							icall=0
'						else
'							member=str_parse(member, ")", 1) & "," &  sf & " Ptr)"
'							'member=FF_Replace( member ,")","," & sf & " Ptr)" )
'						END IF
'						'print "get "; member
'					else
'						if sf="" THEN
'							sf=str_parse(member, ")", 2)
'							sf = FF_Replace( sf , " As VARIANT Ptr", " As VARIANT")
'							if trim (sf,any " '" )<>"" and instr(sf,"As ") THEN
'								if icall=1 THEN
'									member=str_parse(member, ")", 1) &  sf & " Ptr)"
'									'member=FF_Replace( member ,"()","(" & sf & " Ptr)" )
'									icall=0
'								else
'									member=str_parse(member, ")",1) & "," &  sf & " Ptr)"
'									'member=FF_Replace( member ,")","," & sf & " Ptr)" )
'                        END IF
'							else
'								member=str_parse(member, ")", 1) & ")"
'                     END IF
'						else
'							member=str_parse(member, ")", 1) & ")"
'                  END IF
'						'print "call "; member
'						su=""
'               END IF
'					'print prop
'					'print su
'					if su="PUT" or su ="SET"  THEN
'						s2=""
'						s1=""
'						if icall=0 then
'							s1=mid(member,len(str_parse(member, "(", 1))+2)
'							i1=trait(s1,ar())
'							if i1>1 THEN
'								s1="( "
'								for x1 = 1 to i1-1
'									if x1< i1-1 then
'										s1 &= ar(x1,1)& " , "
'									else
'										s1 &= ar(x1,1)& " ) "
'									end if
'									s2 &= " , " & ar(x1,2)
'								NEXT
'							end if
'							s1 &= "=( " & ar(i1,1)& ")"
'							s2 &= " , " & ar(i1,2)
'						end if
'						if su="PUT" THEN
'							member = "Ax_Put " & str1 & "Obj_Ptr , ""." & prop & s1 & """" & s2 & " ' " & member
'						elseif su="SET" THEN
'							member = "Ax_Set " & str1 & "Obj_Ptr , ""." & prop & s1 & """" & s2 & " ' " & member
'						end if
'					elseif su="GET"  THEN
'						s1=mid(member,len(str_parse(member, "(", 1))+2)
'						s1=FF_Replace( s1 ,")",",")
'						's1=str_parse(member, "(", 2)
'						's1=rtrim (str_parse(member, "(", 2),any ") '")
'						i1=trait(s1,ar())
'						s2=""
'						s1=""
'						if i1>1 THEN
'							s1="( "
'							for x1 = 1 to i1-1
'								if x1< i1-1 then
'									s1 &= ar(x1,1)& " , "
'								else
'									s1 &= ar(x1,1)& " ) "
'								end if
'								s2 &= " , "& ar(x1,2)
'							NEXT
'						end if
'						member = "Ax_Get """ & ar(i1,1) & """ , " & ar(i1,2)& " , " & str1 & "Obj_Ptr , ""." & prop & s1 & """" & s2 & " ' " & member
'					else
'						if icall= 1 THEN
'							member = """." & prop & """"
'						else
'							s1=mid(member,len(str_parse(member, "(", 1))+2)
'							's1=str_parse(member, "(", 2)
'							's1=rtrim (str_parse(member, "(", 2),any ") '")
'							i1=trait(s1,ar())
'							s2=""
'							s1="( "
'							for x1 = 1 to i1
'								if x1 < i1 then
'									s1 &= ar(x1,1)& " , "
'								else
'									s1 &= ar(x1,1)& " ) "
'								end if
'								s2 &= " , "& ar(x1,2)
'							NEXT
'							member = """." & prop & s1 & """" & s2 & " ' " & member
'						end if
'						member = "Ax_Call " & str1 & "Obj_Ptr , " & member
'					END IF
'					niv3 &= tb & member & crlf
'					's &=tb & prop & member & crlf
'				End If
'         End If

'         If haschilds Then
'            ReDim Preserve parents(UBound(parents) + 1)
'            parents(UBound(parents)) = hnextitem
'            hnextitem = haschilds
'         Else
'            hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
'            If hnextitem = 0 And UBound(parents) > 0 Then
'               hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), parents(UBound(parents)), TVGN_NEXT))
'               ReDim Preserve parents(UBound(parents) - 1)

'					if niv3<>"" then
'						x91 +=1
'						niv3 &= crlf  & "     End Model List  for   " & xface & "  " & f_bc & crlf & crlf & crlf
'						niv3= debut & niv3
'						s &= niv3
'						niv3=""
'						debut = ""
'					end if
'					's &= tb & "pMark As Integer = -1" & crlf
'					's &= tb & "pThis As Integer" & crlf
'               's &= "End Type        ' " & xface & crlf & crlf
'            End If
'         End If

'      Wend
'		if s<>"" and xfull =1 then
'			s = niveau1 & s
'		elseif s<>"" then
'			s = niveau1 & s
'			if x91 >1 THEN
'				s= crlf & crlf & "'More than 1 Invoke Dispach item selected on the tree , reselect the only one you need" & crlf & crlf & s
'				xret2=1
'			END IF
'		else
'				s= crlf & crlf & "'No Invoke Dispach item selected on the tree , reselect the functions you need" & crlf & crlf
'				xret2=1
'		end if
'   End If
'   'print s22
'   Return cast(zstring ptr, sysallocstringbytelen(s, len(s)))
'End Function


'Function EnumModel cdecl(hTree As Long, hParent() As long) As zString Ptr
'   Dim Parents() As Long
'   Dim hNextItem As Long
'   Dim HasChilds As Long
'   Dim hItem1 as Long
'   Dim token As String, s0 As String, tldesc As String,niveau1 as string,debut as string,niv3 as string
'   Dim xface As String, memid As Integer, member As String, offs As Integer
'	dim inot as integer,icall as integer
'   Dim s as String
'   Dim s1 as String, s2 as string, prop as string
'   dim i1 as integer, x1 as integer, x91 as integer
'   Dim ar() As string
'   dim as string str1
'	GetPrefix()
'	str1 = trim(prefix, any "-_ ")
'   if str1 <> "" THEN str1 = str1 & "_"

'   hItem1 = hparent(tkind_interface)

'   tldesc = TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem1, TVGN_PARENT)))
'	'print "tldesc" : print tldesc : print
'	haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem1))
'   If haschilds Then
'      niveau1 = "'" & String(80, "=") & crlf
'      niveau1 &= "'Template   DispHelper code  -     " & tldesc & crlf
'      niveau1 &= "'" & String(80, "=") & crlf
'		s=""
'      hnextitem = haschilds
'      While hnextitem
'         token = TreeView_GetItemText1(hTree, hNextItem)
'			'print "token" : print token : print
'			if instr(token,"          'NOT'") THEN
'				token=left(token , len(token)-len("          'NOT'"))
'				inot=1
'			else
'				inot=0
'			end if
'         haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hnextitem))
'         If haschilds <> 0 And UBound(parents) = 0 Then
'				xface = str_parse(token, "?", 1)
'				debut = "'" & String(80, "=") & crlf
'				debut &= "'vTable Interface :  " & xface & crlf
'				debut &= "'" & String(80, "=") & crlf & crlf
'				debut &= d_bc & "  Begin Model List : Disphelper syntax for  " & xface & crlf & crlf
'				niv3=""
'         ElseIf (UBound(parents) <> 0) and inot=0 Then
'            memid = Val(str_parse(token, "?", 1))
'            member = str_parse(token, "?", 3) & "     '" & str_parse(token, "?", 4)
'				member = FF_Replace( member , " As VARIANT", " As VARIANT Ptr")
'				member = FF_Replace( member , " As VARIANT Ptr Ptr", " As VARIANT Ptr")
'				'member=str_ReplaceAll(" As VARIANT ", " As VARIANT XXX", member)
'				'member=str_ReplaceAll(" As VARIANT XXX Ptr", " As VARIANT Ptr    ", member)
'				'member=str_ReplaceAll(" As VARIANT XXX", " As VARIANT ptr", member)
'				'member=str_ReplaceAll(" As VARIANT Ptr    ", "_ret As VARIANT Ptr", member)
'				i1=instr(ucase(member),"AS FUNCTION(")
'				if i1>0 THEN
'					member = left(member,i1 -1) & mid(member,i1 + 11)
'            END IF
'				i1=instr(ucase(member),"AS SUB(")
'				if i1>0 THEN member = left(member,i1 -1) & mid(member,i1 + 6)
'				i1=instr(ucase(member),"AS HRESULT")
'				if i1>0 THEN member = left(member,i1 -1)
'            If offs = 0 Then
'               If memid > 8 Then
'                  offs = 8
'               End If
'               If memid > 24 Then
'                  offs = 24
'               End If
'            End If
'            s0 = ""
'            'offs = memid - offs - 4
'            offs = memid
'				prop=trim(str_parse(member, "(", 1),any " '")
'				icall=0
'            If InStr(member, "()") Then
'					member= left(member,InStr(member, "()")-4)
'					icall=1
'            Else
'					member= str_parse(member, "(", 1) & Mid(member, Len(str_parse(member, "(", 1))+1)
'            End If
'				if (ucase(left(member,3))="PUT" or ucase(left(member,3))="SET") and icall= 0 THEN
'					prop=mid(prop,4)
'					s1=mid(member,len(str_parse(member, "(", 1))+2)
'					's1=str_parse(member, "(", 2)
'					's1=rtrim (str_parse(member, "(", 2),any ") '")
'					i1=trait(s1,ar())
'					s2=""
'					s1=""
'					if i1>1 THEN
'						s1="( "
'						for x1 = 1 to i1-1
'							if x1< i1-1 then
'								s1 &= ar(x1,1)& " , "
'							else
'								s1 &= ar(x1,1)& " ) "
'							end if
'							s2 &= " , " & ar(x1,2)
'						NEXT
'					end if
'					s1 &= "=( " & ar(i1,1)& ")"
'					s2 &= " , " & ar(i1,2)
'					if ucase(left(member,3))="PUT" THEN
'						member = "Ax_Put " & str1 & "Obj_Ptr , ""." & prop & s1 & """" & s2 & " ' " & member
'					elseif ucase(left(member,3))="SET" THEN
'						member = "Ax_Set " & str1 & "Obj_Ptr , ""." & prop & s1 & """" & s2 & " ' " & member
'					end if
'				elseif ucase(left(member,3))="GET" and icall= 0 THEN
'					prop=mid(prop,4)
'					s1=mid(member,len(str_parse(member, "(", 1))+2)
'					's1=str_parse(member, "(", 2)
'					's1=rtrim (str_parse(member, "(", 2),any ") '")
'					i1=trait(s1,ar())
'					s2=""
'					s1=""
'					if i1>1 THEN
'						s1="( "
'						for x1 = 1 to i1-1
'							if x1< i1-1 then
'								s1 &= ar(x1,1)& " , "
'							else
'								s1 &= ar(x1,1)& " ) "
'							end if
'							s2 &= " , "& ar(x1,2)
'						NEXT
'					end if
'					member = "Ax_Get """ & ar(i1,1) & """ , " & ar(i1,2)& " , " & str1 & "Obj_Ptr , ""." & prop & s1 & """" & s2 & " ' " & member
'				else
'					if instr (member ,"(")=0 THEN
'						member = """." & prop & """"
'					else
'						s1=mid(member,len(str_parse(member, "(", 1))+2)
'						's1=str_parse(member, "(", 2)
'						's1=rtrim (str_parse(member, "(", 2),any ") '")
'						i1=trait(s1,ar())
'						s2=""
'						s1="( "
'						for x1 = 1 to i1
'							if x1 < i1 then
'								s1 &= ar(x1,1)& " , "
'							else
'								s1 &= ar(x1,1)& " ) "
'							end if
'							s2 &= " , "& ar(x1,2)
'						NEXT
'						member = """." & prop & s1 & """" & s2 & " ' " & member
'					end if
'					member = "Ax_Call " & str1 & "Obj_Ptr , " & member
'            END IF
'				niv3 &= tb & member & crlf
'         End If
'         If haschilds Then
'            ReDim Preserve parents(UBound(parents) + 1)
'            parents(UBound(parents)) = hnextitem
'            hnextitem = haschilds

'         Else
'            hnextitem = cast(long,treeview_getnextitem(cast(hwnd,htree), hnextitem, TVGN_NEXT))
'            If hnextitem = 0 And UBound(parents) > 0 Then
'               hnextitem = cast(long,treeview_getnextitem(cast(hwnd,htree), parents(UBound(parents)), TVGN_NEXT))
'               ReDim Preserve parents(UBound(parents) - 1)
'					if niv3<>"" then
'						x91 +=1
'						niv3 &= crlf  & "     End Model List  for   " & xface & "  " & f_bc & crlf & crlf & crlf
'						niv3= debut & niv3
'						s &= niv3
'						niv3=""
'						debut = ""
'					end if
'					offs = 0
'				End If
'			End If
'		Wend
'		if s<>"" and xfull =1 then
'			s = niveau1 & s
'		elseif s<>"" then
'			s = niveau1 & s
'			if x91 >1 THEN
'				s= crlf & crlf & "'More than 1 vTable Interface item selected on the tree , reselect the only one you need" & crlf & crlf & s
'				xret2=1
'			END IF
'		else
'				s= crlf & crlf & "'No vTable Interface item selected on the tree , reselect the functions you need" & crlf & crlf
'				xret2=1
'		end if
'	End If
'Return cast(zstring ptr, sysallocstringbytelen(s, len(s)))
'End Function

Function EnumCoClass2 (hTree As Long, hParent() As long) As String
   Dim Parents() As Long
   Dim hNextItem As Long
   Dim HasChilds As Long
   Dim token As String
   Dim As String clsids,  tldesc
   Dim hItem as long

   Dim s as string
	redim Parents(0)
	hItem = hparent(tkind_coclass)
   tldesc = TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem, TVGN_PARENT)))
   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
   If haschilds Then
      hnextitem = haschilds
      While hnextitem
         token = TreeView_GetItemText(hTree, hNextItem)
         clsids &=  str_parse(token, "?", 1) & "=" & dq & str_parse(token, "?", 2) & dq & crlf
          hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
      Wend
   End If

   's &= "'CLSID - " & tldesc & crlf
   s &= clsids & crlf

   Return s
End Function


Function EnumvTable cdecl Alias "EnumvTable"(hTree As Long, hParent() As long) As zString Ptr
   Dim Parents() As Long
   Dim hNextItem As Long
   Dim HasChilds As Long
   Dim hItem1 as Long,virg as integer,xvir as integer,in1 as integer,bevent as integer
   Dim token As String, s0 As String, tldesc As String, cor1 as string, cor2 as string, cor3 as string,sp1 as string
   Dim xface As String, memid As Integer, member As String, offs As Integer
	redim Parents(0)
   Dim s as String,us as string,xface0 as string,svt as string,s1 as string
	dim as string str1, strclasses, interf, siid
	GetPrefix()
	'strclasses =EnumCoClass2 (hTree , hParent())
	's &=strclasses
	str1 = trim(prefix, any "-_ ")
   hItem1 = hparent(tkind_interface)
   tldesc = TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem1, TVGN_PARENT)))
   'print "tldesc = " ; tldesc
	haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem1))
	'print "haschilds = ";haschilds
   If haschilds Then
      svt &= "'" & String(80, "=") & crlf
      svt &= "'vTable - " & tldesc & crlf
      svt &= "'" & String(80, "=") & crlf
      hnextitem = haschilds
      While hnextitem
         token = TreeView_GetItemText(hTree, hNextItem)
			'print "token = ";token
         haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hnextitem))
			'haschilds = treeview_getchild( htree, hnextitem)
			'print "haschilds =";haschilds ; "  UBound(parents)= " ;UBound(parents)
         If haschilds <> 0 And UBound(parents) = 0 Then
			'If haschilds <> 0 And UBound(parents) < 1 Then
				interf = str_parse(token, "?", 2)
				'print "interf >" ; interf ; "<"
            'xface = str1 & str_parse(token, "?", 1)

				xface = str_parse(token, "?", 1)
				xface0=xface
				xface= str1 & xface
            bevent = inevtlist(xface0)
            If bevent=0 THEN
					if instr(ucase(xface0),"EVENT") THEN bevent= 2
				end if
				s1 = "'" & String(80, "=") & crlf
				s1 &= "'Interface " &  xface & "     ?" & str_parse(token, "?", 3) & crlf
				s1 &= "'Const IID_" &  xface & "=" & dq & str_parse(token, "?", 2) & dq & crlf
				s1 &= "'" & String(80, "=") & crlf
				s1 &= "Type " &  xface & "vTbl" & crlf

         ElseIf (UBound(parents) <> 0) Then
			'ElseIf (UBound(parents) > 0) Then
            memid = Val(str_parse(token, "?", 1))
            member = str_parse(token, "?", 3) & "     '" & str_parse(token, "?", 4)
				'print " memid = " & memid
				'print "member = " & member
            If offs = 0 Then
               If memid > 8 Then
                  s1 &= tb & "QueryInterface As Function (Byval pThis As Any ptr,Byref riid As GUID,Byref ppvObj As Dword) As hResult" & crlf
                  s1 &= tb & "AddRef As Function (Byval pThis As Any ptr) As Ulong" & crlf
                  s1 &= tb & "Release As Function (Byval pThis As Any ptr) As Ulong" & crlf
                  offs = 8
               End If
               If memid > 24 Then
                  s1 &= tb & "GetTypeInfoCount As Function (Byval pThis As Any ptr,Byref pctinfo As Uinteger) As hResult" & crlf
                  s1 &= tb & "GetTypeInfo As Function (Byval pThis As Any ptr,Byval itinfo As Uinteger,Byval lcid As Uinteger,Byref pptinfo As Any Ptr) As hResult" & crlf
                  s1 &= tb & "GetIDsOfNames As Function (Byval pThis As Any ptr,Byval riid As GUID,Byval rgszNames As Byte,Byval cNames As Uinteger,Byval lcid As Uinteger,Byref rgdispid As Integer) As hResult" & crlf
                  s1 &= tb & "Invoke As Function (Byval pThis As Any ptr,Byval dispidMember As Integer,Byval riid As GUID,Byval lcid As Uinteger,Byval wFlags As Ushort,Byval pdispparams As DISPPARAMS,Byref pvarResult As Variant,Byref pexcepinfo As EXCEPINFO,Byref puArgErr As Uinteger) As hResult" & crlf
                  offs = 24
               End If
            End If
            s0 = ""
            offs = memid - offs - 4
            If offs > 0 Then s0 = "(" &(offs - 1) & ") As Byte" & crlf
            offs = memid
            If Len(s0) Then s1 &= tb & "Offset" & offs & s0

'				if instr(ucase(member),")AS HRESULT ") = 0 and instr(ucase(member),") AS HRESULT ") = 0 then
'					if instr(member," As Function(")>0 then
'						member=FF_replace(member," As Function("," As Sub(")
'					end if
'					'print member
'				end if
            If InStr(member, "()") and InStr(member, " As Sub()")>0 Then
					s1 &= tb & str_parse(member, "(", 1) & "(ByVal pThis As Any Ptr)" & crlf
				elseIf InStr(member, "()") and InStr(member, " As Function()")>0 Then
               s1 &= tb & str_parse(member, "(", 1) & "(ByVal pThis As Any Ptr" & Mid(member, Len(str_parse(member, "(", 1)) + 2, - 1) & crlf
            ElseIf InStr(member, "()")= 0 then
					cor1= str_parse(member, "(", 1) & "(ByVal pThis As Any Ptr," & Mid(member, Len(str_parse(member, "(", 1)) + 2, - 1) & crlf

					sp1= ") As " & str_parse(cor1,") As ",str_numparse(cor1, ") As " ))
					cor1= str_parse(cor1,") As ",1)
					virg = str_numparse(cor1, "," )
					for xvir = 2 to virg
						cor2 = str_parse(cor1, "," , xvir)
						if instr(cor2," As ") THEN
							cor3 = "ByVal " & cor2
							cor1 = FF_Replace( cor1 ,cor2,cor3)
                  END IF
               NEXT

					cor1 = FF_Replace( cor1 ,",ByVal  As",",ByVal VarAny As")
				   s1 &= tb & cor1 & sp1
            End If

         End If

         If haschilds Then
            ReDim Preserve parents(UBound(parents) + 1)
            parents(UBound(parents)) = hnextitem
            hnextitem = haschilds
         Else
            hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
            If hnextitem = 0 And UBound(parents) > 0 Then
               hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), parents(UBound(parents)), TVGN_NEXT))
               ReDim Preserve parents(UBound(parents) - 1)
               s1 &= "End Type      '" &  xface & "vTbl" & crlf & crlf
               s1 &= "Type " &  xface & "_"& crlf
               s1 &= tb & "lpvtbl As " &  xface & "vTbl Ptr" & crlf
               s1 &= "End Type" & crlf & crlf
					if bevent=0 then us &= tb & "Type  " & xface & "    as  " & xface & "_     ' Interface :  " &  xface  & crlf
					s1 &= tb & "'     best way to do ...            Dim Shared As " & xface & " ptr  " & str1 & "pVTI  " & crlf
					s1 &= tb & "'     ex:  "  & str1 & "pVTI->lpvtbl->SetColor( " & str1 & "pVTI ,...)" & crlf
					s1 &= tb & "'     or    Ax_Vt ( " & str1 & "pVTI , SetColor ,...)" & crlf
					s1 &= tb & "'     or    Ax_Vt0 ( " & str1 & "pVTI , About )" & crlf & crlf
					in1=0
					for xvir = 1 to icoclass
						if coclass(xvir,5) = interf THEN
							'print coclass(xvir,4)
							in1=xvir
							exit for
                  END IF
               NEXT
					if in1 THEN
						s1 &= "Function Create_" & xface & "() As any ptr" & crlf
						s1 &= tb & "Function = AxCreate_object( """ & coclass(in1,3) & """ , """ & coclass(in1,5)& """ )" & crlf
						s1 &= "End Function" & crlf & crlf
						s1 &= " ' use  normaly    like that  :      " & str1 & "pVTI = " & str1 & "Obj_Ptr" & crlf
						s1 &= " ' but if problem try         :      " & str1 & "pVTI = Create_" & xface & "()"  & crlf & crlf & crlf
					else
						s1 &= " ' use   " & str1 & "pVTI = " & str1 & "Obj_Ptr" & crlf & crlf & crlf
               END IF
					if bevent=0 then  s &= s1
               offs = 0
            End If
         End If
      Wend

   End If
	if s <> "" THEN
		s = "'" & String(80, "*") & crlf & "'  Interface vTable Types :" & crlf & crlf & us & crlf & crlf & "'" & String(80, "*") & crlf & crlf & s
		s = svt & crlf & crlf & s
	END IF

   'print s

	if instr(s , " _Collection Ptr" ) THEN
		us= "'================================================================================" & crlf & crlf
		us &= "'Interface _Collection  , Standard Vb6/Vba collection vTable needed here" & crlf & crlf
		us &= "Type _CollectionvTbl" & crlf
		us &= tb & "QueryInterface As Function (Byval pThis As Any ptr,Byref riid As GUID,Byref ppvObj As Dword) As hResult" & crlf
		us &= tb & "AddRef As Function (Byval pThis As Any ptr) As Ulong" & crlf
		us &= tb & "Release As Function (Byval pThis As Any ptr) As Ulong" & crlf
		us &= tb & "GetTypeInfoCount As Function (Byval pThis As Any ptr,Byref pctinfo As Uinteger) As hResult" & crlf
		us &= tb & "GetTypeInfo As Function (Byval pThis As Any ptr,Byval itinfo As Uinteger,Byval lcid As Uinteger,Byref pptinfo As Any Ptr) As hResult" & crlf
		us &= tb & "GetIDsOfNames As Function (Byval pThis As Any ptr,Byval riid As GUID,Byval rgszNames As Byte,Byval cNames As Uinteger,Byval lcid As Uinteger,Byref rgdispid As Integer) As hResult" & crlf
		us &= tb & "Invoke As Function (Byval pThis As Any ptr,Byval dispidMember As Integer,Byval riid As GUID,Byval lcid As Uinteger,Byval wFlags As Ushort,Byval pdispparams As DISPPARAMS,Byref pvarResult As Variant,Byref pexcepinfo As EXCEPINFO,Byref puArgErr As Uinteger) As hResult" & crlf
		us &= tb & "Item As Function(ByVal pThis As Any Ptr,ByVal Index As VARIANT Ptr,ByVal pvarRet As VARIANT Ptr) As HRESULT      '" & crlf
		us &= tb & "Add As Function(ByVal pThis As Any Ptr,ByVal Item As VARIANT Ptr,ByVal Key As VARIANT Ptr=0,ByVal Before As VARIANT Ptr=0,ByVal After As VARIANT Ptr=0) As HRESULT     '" & crlf
		us &= tb & "Count As Function(ByVal pThis As Any Ptr,ByVal pi4 As Integer Ptr) As HRESULT     '" & crlf
		us &= tb & "Remove As Function(ByVal pThis As Any Ptr,ByVal Index As VARIANT Ptr) As HRESULT     '" & crlf
		us &= tb & "_NewEnum As Function(ByVal pThis As Any Ptr,ByVal ppunk As LPUNKNOWN Ptr) As HRESULT     '" & crlf
		us &= "End Type      '_CollectionvTbl" & crlf & crlf
		us &= "Type _Collection" & crlf
		us &= tb & "lpvtbl As _CollectionvTbl Ptr" & crlf
		us &= "End Type" & crlf
		us &= "'================================================================================" & crlf
		s = us & crlf & s
   END IF

	Return cast(zstring ptr, sysallocstringbytelen(s, len(s)))
End Function

Function EnumInvoke cdecl Alias "EnumInvoke"(hTree As Long, hParent() As long) As zString Ptr
   Dim Parents() As Long
   Dim hNextItem As Long
   Dim HasChilds As Long

   Dim token As String, tldesc As String
   Dim hItem as Long
   Dim xface As String, memid As String, member As String,s1 as string,xface0 as string,sinv as string
   Dim As Integer cpar,  bevent
   '  Dim s22 as String
   Dim s as String
	dim as string str1
	redim Parents(0)
	GetPrefix()
	str1 = trim(prefix, any "-_ ")
	'inEvtList2()
   hItem = hparent(tkind_dispatch)
   tldesc =  TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem, TVGN_PARENT)))
   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
   If haschilds Then
      sinv = "'" & String(80, "=") & crlf
      sinv &= "'Invoke - " & tldesc & crlf
      sinv &= "'" & String(80, "=") & crlf
      hnextitem = haschilds
      While hnextitem
         token = TreeView_GetItemText(hTree, hNextItem)
         haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hnextitem))
         If (haschilds <> 0) And (UBound(parents) = 0) Then
            'xface = str1 & str_parse(token, "?", 1)
				xface = str_parse(token, "?", 1)
				xface0=xface
				xface= str1 & xface
            bevent = inevtlist(xface0)
				If bevent=0 THEN
					if instr(ucase(xface0),"EVENT") THEN bevent= 2
				end if
				'print xface0 & "  = " & str(bevent)
            s1 = "'" & String(80, "=") & crlf
            s1 &= "'Dispatch " & xface & "     ?" & str_parse(token, "?", 3) & crlf
            s1 &= "'Const IID_" & xface & "=" & dq & str_parse(token, "?", 2) & dq & crlf
            s1 &= "'" & String(80, "=") & crlf
            s1 &= "Type " & xface & crlf
         ElseIf (UBound(parents) <> 0) then
            memid = str_parse(token, "?", 1)
            member = str_parse(token, "?", 3)
				member = FF_Replace( member , "Const ", "")
				member=rtrim(member,any"'= " & dq)
				if instr(member,")")=0 and instr(member,"(")THEN member=member & ")"

            cpar = str_numparse(member, ",")
            If Len(str_parse(str_parse(member, "(", 2), ")", 1)) = 0 Then cpar = 0
				s1 &=tb & str_parse(member," As ",1) & " As tMember = ("
				's22 &=tb & str_parse(member," As ",1)
				s1 &= memid & ",2," & cpar & "," & str_parse(token,"?",2) & ")    '" & str_remove(member,str_parse(member," As ",1))  & crlf
            ' s22 &=str_remove(member,str_parse(member," As ",1))  & crlf
         End If

         If haschilds Then
            ReDim Preserve parents(UBound(parents) + 1)
            parents(UBound(parents)) = hnextitem
            hnextitem = haschilds
         Else
            hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
            If hnextitem = 0 And UBound(parents) > 0 Then
               hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), parents(UBound(parents)), TVGN_NEXT))
               ReDim Preserve parents(UBound(parents) - 1)
					s1 &= tb & "pMark As Integer = -1" & crlf
					s1 &= tb & "pThis As Integer" & crlf
               s1 &= "End Type        ' " & xface & crlf & crlf
					s1 &= "	'Use like that to use these dispach/invoke functions" & crlf
					s1 &= "	'  Dim Shared As " & xface & " " & str1 & "Obj_Disp" & crlf
					s1 &= "	'  SetObj ( @" & str1 & "Obj_Disp , " & str1 & "Obj_Ptr ) ' connect to object"  & crlf
					s1 &= "	'  ex : AxCall " & str1 & "Obj_Disp.putMonth,vptr(05)"  & crlf & crlf & crlf
					if bevent =0 THEN s &=s1
				End If

         End If

      Wend
		if s <> "" THEN s = sinv & crlf & crlf & s
   End If
   'print s22
   Return cast(zstring ptr, sysallocstringbytelen(s, len(s)))
End Function

Function EnumEvent2 (hTree As Long, hParent() As long, hItem as Long )as String
	Dim Parents() As Long
   Dim hNextItem As Long
   Dim HasChilds As Long

   Dim token As String, tldesc As String

   Dim As String xface, member, help, temp ,xface0,us,memberF,memberV, temp2
   Dim As Integer  bevent,numF,x1,itemp

   Dim s as String
	dim as string str1,st_evt(),st_M(),sieven  ',st_F()
	redim Parents(0)
	redim st_evt(0)
	'redim st_F(0)
	redim st_M(0)
	numF = 0

	GetPrefix()
	str1 = trim(prefix, any "-_ ")

	'print "sevent = " ; sevent
   'hItem = hparent(tkind_interface)

   tldesc = str1 & TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem, TVGN_PARENT)))
   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
   If haschilds Then


      us = "'" & String(80, "=") & crlf
      us &= "'Event - " & tldesc & crlf
      us &= "'" & String(80, "=") & crlf
      hnextitem = haschilds
      While hnextitem

         token = TreeView_GetItemText(hTree, hNextItem)
         haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hnextitem))
         If (haschilds <> 0) And (UBound(parents) = 0) Then
				xface = str_parse(token, "?", 1)
				xface0=xface
				xface= str1 & xface
            bevent = inevtlist(xface0)
            If bevent Then
					sieven = str_parse(token, "?", 2)
					'print "sieven = " ; sieven
					if sevent = "" or sieven = sevent THEN
						s &= "'" & String(80, "=") & crlf
						s &= "'Event vTable - " & xface & "     ?" & str_parse(token, "?", 3) & crlf
						s &= "Const " & xface & "_Ev_IID_CP = " & dq & str_parse(token, "?", 2) & dq & crlf
						s &= "'" & String(80, "=") & crlf & crlf & crlf
               END IF
            End If
         ElseIf (UBound(parents) <> 0) then
            numF +=1
            memberF = str_parse(token, "?", 3)
				memberF = FF_Replace(memberF,"   "," ")
				memberF = FF_Replace(memberF,"  "," ")
				memberF = FF_Replace(memberF,"  "," ")
				memberF = FF_Replace(memberF,"  "," ")
				memberF = trim (memberF)
				temp2 = memberF
				memberF = lcase(memberF)

            help = str_parse(token, "?", 4)

            redim preserve st_evt(1 to numF )
				'redim preserve st_F(1 to numF)
				redim preserve st_M(1 to numF)
				itemp = instr(memberF, " as ")

				member = left(temp2,itemp -1)

				st_M(numF)= xface & "_" & member
				If InStr(memberF, " as sub()")>0 or InStr(memberF, " as sub ()")> 0 _
							or	InStr(memberF, " as sub ( )")>0 or	InStr(memberF, " as sub( )")>0 Then
					memberF= member & "     As Sub ( ByVal pThis As Any Ptr )"
				elseIf InStr(memberF, " as function()")>0 or InStr(memberF, " as function ()")> 0 _
							or	InStr(memberF, " as function ( )")>0 or	InStr(memberF, " as function( )")>0 Then
               memberF= member & "     As Function ( ByVal pThis As Any Ptr )" & str_parse(temp2, ")", 2)
				elseIf InStr(memberF, " as function(")>0 or InStr(memberF, " as function (")> 0 then
					memberF= member & "     As Function ( ByVal pThis As Any Ptr , " & str_parse(temp2, "(", 2)
				elseIf InStr(memberF, " as sub(")>0 or InStr(memberF, " as sub (")> 0 then
					memberF= member & "     As Sub ( ByVal pThis As Any Ptr , " & str_parse(temp2, "(", 2)
				end if



				st_evt(numF )= memberF

				'print memberF

				temp = mid(memberF, instr(memberF," As ") +3)
				temp = trim(temp)

				'print " temp = " & temp & "   numf = " & str(numF)

				if left(temp,3)="Sub" THEN
					temp = "Sub " &  st_M(numF) & "( " & str_parse(temp, "(", 2)

				else
					temp = "Function " &  st_M(numF) & "( " & str_parse(temp, "(", 2)
            END IF


            If bevent and (sevent = "" or sieven = sevent) Then
               s &= "'" & String(80, "=") & crlf
               s &= "'Event # " & str(numF) & "   " & member & "   ? " & help & crlf
               s &= "'" & String(80, "=") & crlf
					memberV = temp
               s &= memberV & crlf & crlf
               s &= tb & "'*** Put your code here ***" & crlf & crlf & crlf
					if left(ucase(memberV),3)="SUB" THEN
						 s &= "End Sub" & crlf & crlf
					else
						s &= tb & "Function = S_OK   ' change if needed" & crlf
						s &= "End Function" & crlf & crlf
               END IF
            End If
         End If
         If haschilds Then
            ReDim Preserve parents(UBound(parents) + 1)
            parents(UBound(parents)) = hnextitem
            hnextitem = haschilds
         Else
            hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
            If hnextitem = 0 And UBound(parents) > 0 Then
               hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), parents(UBound(parents)), TVGN_NEXT))
               ReDim Preserve parents(UBound(parents) - 1)

               If bevent and (sevent = "" or sieven = sevent) and numF >0 Then
						s &= "Type " & xface & "Vtbl_Ev" & crlf
						s &= tb & "QueryInterface   As Function ( ByVal pThis As Any Ptr, ByRef riid As GUID, ByRef ppvObj As Dword ) As hResult" & crlf
						s &= tb & "AddRef           As Function ( ByVal pThis As Any Ptr ) As Ulong" & crlf
						s &= tb & "Release          As Function ( ByVal pThis As Any Ptr ) As Ulong" & crlf
						for x1 = 1 to numF
							s &= tb & st_evt(x1) & crlf
						NEXT
						s &= "End type" & crlf & crlf
						s &= "Type " & xface & "_Ev" & crlf
						s &= tb & "lpVtbl           As " & xface & "Vtbl_Ev Ptr" & crlf
						s &= "End type" & crlf & crlf
'						for x1 = 1 to numF
'							s &= "Declare " & st_F & crlf
'						NEXT
						s &= "Dim Shared " & xface & "_EvTable As " & xface & "Vtbl_Ev = ( @ Ev_Vtbl_QueryInterface, _" & crlf
						s &= tb & tb & tb & tb & tb & "@ Ev_Vtbl_AddRef, _" & crlf
						s &= tb & tb & tb & tb & tb & "@ Ev_Vtbl_Release"
						for x1 = 1 to numF
							s &= ", _" & crlf & tb & tb & tb & tb & tb & "@ " & st_M (x1)
						NEXT
						s &= ")"	 & crlf & crlf
						s &= "Dim Shared " & xface & "pSEvent As " & xface & "_Ev" & crlf & crlf
						s &= tb & tb & xface & "pSEvent.lpvtbl =  @ " & xface & "_EvTable" & crlf & crlf
						s &= "Function " & xface & "_Events_Connect ( ByVal pUnk As Any Ptr, ByRef EvdwCookie As Dword ) As Integer" & crlf
						s &= tb & "EvdwCookie = Ev_Vtbl_Cp ( pUnk , " & xface & "_Ev_IID_CP, cast( Dword, @ " & xface & "pSEvent ))" & crlf
						s &= tb & "Function = cast ( integer, EvdwCookie )" & crlf
						s &= "End Function" & crlf & crlf
						s &= "Function " & xface & "_Events_Disconnect ( ByVal pUnk As Any Ptr, ByVal EvdwCookie As Dword ) As Integer " & crlf
						s &= tb & "Function = cast ( integer, Ev_Vtbl_Cp ( pUnk, " & xface & "_Ev_IID_CP, cast( Dword, @ " & xface & "pSEvent ), EvdwCookie ) )" & crlf
						s &= "End Function" & crlf & crlf
               End If
					numF = 0
					redim st_evt(0)
					'redim st_F(0)
					redim st_M(0)
            End If
         End If
      Wend
		'if s<>"" THEN s= us & crlf & s
		function=s
   End If



end function

Function EnumEvent cdecl Alias "EnumEvent"(hTree As Long, hParent() As long) As zString Ptr
   Dim Parents() As Long
   Dim hNextItem As Long
   Dim HasChilds As Long

   Dim token As String, tldesc As String
   Dim hItem as Long
   Dim As String xface, memid, member, par, help, temp,xface0,us
   Dim As Integer cpar, bevent,pass

   Dim s as String
	dim as string str1,st_as(),sieven
	dim as integer i_as()
	redim Parents(0)
	GetPrefix()
	str1 = trim(prefix, any "-_ ")

	'print "sevent = " ; sevent
   hItem = hparent(tkind_dispatch)

   tldesc = str1 & TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem, TVGN_PARENT)))
   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
   If haschilds Then
      us = "'" & String(80, "=") & crlf
      us &= "'Event - " & tldesc & crlf
      us &= "'" & String(80, "=") & crlf
      hnextitem = haschilds
      While hnextitem
         token = TreeView_GetItemText(hTree, hNextItem)
         haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hnextitem))
         If (haschilds <> 0) And (UBound(parents) = 0) Then
				xface = str_parse(token, "?", 1)
				xface0=xface
				xface= str1 & xface
            bevent = inevtlist(xface0)
            If bevent Then
					sieven = str_parse(token, "?", 2)
					'print "sieven = " ; sieven
					if sevent = "" or sieven = sevent THEN
						s &= "'" & String(80, "=") & crlf
						s &= "'Event Dispatch - " & xface & "     ?" & str_parse(token, "?", 3) & crlf
						s &= "Const " & xface & "_IID_CP" & "=" & dq & str_parse(token, "?", 2) & dq & crlf
						s &= "'" & String(80, "=") & crlf
               END IF
            End If
         ElseIf (UBound(parents) <> 0) then
            memid = str_parse(token, "?", 1)
            member = str_parse(token, "?", 3)
            par = str_parse(str_parse(member, "(", 2), ")", 1)
            member = str_parse(member, " As", 1)
            help = str_parse(token, "?", 4)
            cpar = str_numparse(par, ",")
				if ucase(member)<> "QUERYINTERFACE"  and ucase(member)<>"ADDREF"  _
						and ucase(member)<>"RELEASE" and ucase(member)<>"GETTYPEINFOCOUNT" _
						and ucase(member)<>"GETTYPEINFO"  and ucase(member)<>"GETIDSOFNAMES" _
						and ucase(member)<>"INVOKE" THEN
					If Len(par) = 0 Then cpar = 0
					If bevent and (sevent = "" or sieven = sevent) Then
						s &= "'" & String(80, "=") & crlf
						s &= "'Event=" & member & ", ID=" & memid & " ?" & help & crlf
						s &= "'" & String(80, "=") & crlf
						s &= "Function " & xface & "_" & member & "(ByVal pCookie As Events_IDispatchVtbl Ptr, ByVal pdispparams As DISPPARAMS Ptr) As HRESULT" & crlf
						s &= tb & "Dim pThis As Dword=>pCookie->pThis" & crlf
						s &= tb & "Dim IdEvent As Dword = cast(Dword, pCookie)" & crlf
						temp &= tb & tb & tb & "Case " & memid & " '" & help & crlf
						temp &= tb & tb & tb & tb & "Function=" & xface & "_" & member & "(cast(Any Ptr,pUnk), pdispparams)" & crlf
						If cpar Then s &= tb & "Dim pv As Variant Ptr = pdispparams->rgVArg" & crlf
						redim i_as(1 to cpar)
						redim st_as(1 to cpar)
						For i As Short = 1 To cpar
							st_as(i)= trim(str_parse(par, ",", i))
							if left(st_as(i),5)= "Data " THEN st_as(i) = "Data1 " & mid (st_as(i),6)
							if instr(st_as(i)," DataObject Ptr") then
								st_as(i)=FF_Replace( st_as(i), " DataObject Ptr"," IDispatch Ptr    ' DataObject Ptr" )
							elseif instr(st_as(i)," DataObject") then
								st_as(i)=FF_Replace( st_as(i), " DataObject"," IDispatch    ' DataObject" )
							end if
							if right(st_as(i),3)= "Ptr" THEN st_as(i) = left(st_as(i) ,len(st_as(i))- 4)
							s &= tb & "Dim " & st_as(i)  & crlf
							if right(st_as(i),3)= "Ptr" THEN i_as(i)=1
						Next
						s &= crlf
						For i As Short = 1 To cpar
							if i_as(i)=1 THEN
								s &= tb & str_parse(st_as(i), " As", 1) & "=cast(Any Ptr,CUInt(VariantV(pv[" & cpar - i & "])))" & crlf
							else
								s &= tb & str_parse(st_as(i), " As", 1) & "=VariantV(pv[" & cpar - i & "])" & crlf
							end if
						Next
						s &= tb & "'*** Put your code here ***" & crlf & crlf
						s &= tb & "Function=S_OK" & crlf
						s &= "End Function" & crlf & crlf
					End If
				End If
         End If
         If haschilds Then
            ReDim Preserve parents(UBound(parents) + 1)
            parents(UBound(parents)) = hnextitem
            hnextitem = haschilds
         Else
            hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
            If hnextitem = 0 And UBound(parents) > 0 Then
               hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), parents(UBound(parents)), TVGN_NEXT))
               ReDim Preserve parents(UBound(parents) - 1)
               If bevent and (sevent = "" or sieven = sevent) Then
                  temp = "Function " & xface & "_IDispatch_Invoke(ByVal pUnk As IDispatch Ptr, ByVal dispidMember As DispID, ByVal riid As IID Ptr, _" & crlf & _
                        "  ByVal lcid As LCID, ByVal wFlags As Ushort, ByVal pdispparams As DISPPARAMS Ptr, Byval pvarResult As Variant Ptr, _" & crlf & _
                        "  ByVal pexcepinfo As EXCEPINFO Ptr, ByVal puArgErr As Uint Ptr) As Hresult" & crlf & crlf & _
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

                  s &= "' --------------------------------------------------------------------------------------------" & CRLF
                  s &= "' " & xface & " Events Connection function" & crlf
                  s &= "' --------------------------------------------------------------------------------------------" & CRLF
                  s &= "Function " & xface & "_Events_Connect ( ByVal pUnk As Any Ptr, ByRef dwcookie As Dword ) As Integer" & crlf
						s &= tb & "Function = Ax_Events_cnx ( " & xface & "_IID_CP , pUnk , dwcookie, ProcPtr ( " & xface & "_IDispatch_Invoke ))" & crlf
                  s &= "End Function" & crlf & crlf
                  s &= "Function " & xface & "_Events_Disconnect ( ByVal pUnk As Any Ptr, ByVal dwcookie As Dword ) As Integer" & crlf
						s &= tb & "Function = Ax_Events_cnx ( " & xface & "_IID_CP , pUnk , dwcookie )" & crlf
                  s &= "End Function" & crlf & crlf
               End If
            End If
         End If
      Wend
		if s<>"" THEN s= us & crlf & s
   End If
	if s="" THEN
		if pass =1 THEN
			s=" No Event "
		else
			hItem = hparent(tkind_interface)
			s="" : us = ""
			pass =1
			s=EnumEvent2 (hTree , hParent(), hItem )
			if s="" THEN
				s= "'  No Events ! "
			else
				s= "'       vTable Events ! " & crlf & crlf & crlf & s
         END IF
      END IF
	else
		s= "'       Dispach Events ! " & crlf & crlf & crlf & s
	END IF
   Return cast(zstring ptr, sysallocstringbytelen(s, len(s)))
End Function

Function EnumEnum cdecl Alias "EnumEnum"(hTree As Long, hParent() As long) As zString Ptr
   Dim Parents() As Long
   Dim hNextItem As Long
   Dim HasChilds As Long

   Dim token As String, tldesc As String
   Dim hItem as Long

   Dim s as String

	redim Parents(0)
   hItem = hparent(tkind_enum)
   tldesc = TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem, TVGN_PARENT)))
   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
   If haschilds Then
      s &= "'" & String(80, "=") & crlf
      s &= "'Enum -  " & tldesc & crlf
      s &= "'" & String(80, "=") & crlf
      hnextitem = haschilds
      While hnextitem
         token = TreeView_GetItemText(hTree, hNextItem)
         haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hnextitem))
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
            hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
            If hnextitem = 0 And UBound(parents) > 0 Then
               hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), parents(UBound(parents)), TVGN_NEXT))
               ReDim Preserve parents(UBound(parents) - 1)
               s &= crlf
            End If
         end if
      Wend
   End If
   Return cast(zstring ptr, sysallocstringbytelen(s, len(s)))
End Function

Function EnumRecord cdecl Alias "EnumRecord"(hTree As Long, hParent() As long) As zString Ptr
   Dim Parents() As Long
   Dim hNextItem As Long
   Dim HasChilds As Long
   Dim token As String
   Dim hItem as Long
   Dim s as String
	redim Parents(0)
   hItem = hparent(tkind_record)
   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
   If haschilds Then
      hnextitem = haschilds
      While hnextitem
         token = TreeView_GetItemText(hTree, hNextItem)
         haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hnextitem))
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
            hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
            If hnextitem = 0 And UBound(parents) > 0 Then
               hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), parents(UBound(parents)), TVGN_NEXT))
               ReDim Preserve parents(UBound(parents) - 1)
               s &= "End Type" & crlf
            End If
         End If
      Wend
   End If
   Return cast(zstring ptr, sysallocstringbytelen(s, len(s)))
End Function

Function EnumUnion cdecl Alias "EnumUnion"(hTree As Long, hParent() As long) As zString Ptr
   Dim Parents() As Long
   Dim hNextItem As Long
   Dim HasChilds As Long
   Dim token As String
   Dim hItem as Long

   Dim s as String

	redim Parents(0)
   hItem = hparent(tkind_union)
   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
   If haschilds Then
      hnextitem = haschilds
      While hnextitem
         token = TreeView_GetItemText(hTree, hNextItem)
         haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hnextitem))
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
            hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
            If hnextitem = 0 And UBound(parents) > 0 Then
               hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), parents(UBound(parents)), TVGN_NEXT))
               ReDim Preserve parents(UBound(parents) - 1)
               s &= "End Union" & crlf
            End If
         End If
      Wend
   End If
   Return cast(zstring ptr, sysallocstringbytelen(s, len(s)))
End Function

Function EnumModule cdecl Alias "EnumModule"(hTree As Long,hParent() As long)As zString Ptr
	Dim Parents() As Long
	Dim hNextItem As Long
	Dim HasChilds As Long
	Dim token As String,tldesc As String
	Dim hItem as Long
	Dim as String s,func,dlls,aliass,ord,us,d1,d2
	redim Parents(0)
	hItem=hparent(tkind_module)
	'tldesc=TreeView_GetItemText(hTree,treeview_getnextitem(htree,hitem,TVGN_PARENT))
	'haschilds=treeview_getchild(htree,hitem)
	tldesc = TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem, TVGN_PARENT)))
   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
	If haschilds Then
	s &="'" & String(80,"=") & crlf
	s &="'Module - " & tldesc & crlf
	s &="'" & String(80,"=") & crlf
	hnextitem=haschilds
	While hnextitem
'		While treeview_getcheckstate(htree,hnextitem)=1
'			'hnextitem=treeview_getnextitem(htree,hnextitem,TVGN_NEXT)
'			hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
'			If hnextitem=0 Then Return cast(zstring ptr, sysallocstringbytelen(s, len(s)))
'		Wend

		token=TreeView_GetItemText(hTree,hNextItem)
		If UBound(parents)=0 Then
			s &="'" & String(80,"=") & crlf
			s &="'" & str_parse(token,"?",1) & "     ?" & str_parse(token,"?",2) & crlf
			s &="'" & String(80,"=") & crlf
		Else
			If Left(token,5)<>"Const" Then
				dlls=str_parse(token,"?",1):aliass=str_parse(token,"?",2)
				ord=str_parse(token,"?",3):func=str_parse(token,"?",4)
				s &="Dim shared " & func & "     '" & str_parse(token,"?",5) & crlf
				If Len(aliass) Then
					s &=str_parse(func," As",1) & " = DylibSymbol(h" & str_parse(dlls,".",1) & "," & dq & aliass & dq & ")"
				Else
					s &=str_parse(func," As",1) & " = DylibSymbol(h" & str_parse(dlls,".",1) & "," & ord & ")"
				EndIf
				s &="     'h" & str_parse(dlls,".",1) & " = DylibLoad(" & dq & dlls & dq & ")" & crlf
				d1="h" & str_parse(dlls,".",1) & " = DylibLoad(" & dq & dlls & dq & ")"
				if instr(ucase(us),ucase("Dim Shared h" & str_parse(dlls,".",1)))= 0 THEN
					'print selpath
					d2=selpath
					if instr( d2,dlls)THEN
						d1= ff_replace(d1,dlls,d2)
               END IF
					us &= "Dim Shared h" & str_parse(dlls,".",1) & " As Any Ptr " & crlf & d1 & crlf & crlf
            END IF
			Else
				s &=str_parse(token,"?",1) & "     '" & str_parse(token,"?",2) & crlf
			EndIf
		EndIf

		'haschilds=treeview_getchild(htree,hnextitem)
		haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hnextitem))
		If haschilds Then
			ReDim Preserve parents(UBound(parents)+1)
			parents(UBound(parents))=hnextitem
			hnextitem=haschilds
		Else
			'hnextitem=treeview_getnextitem(htree,hnextitem,TVGN_NEXT)
			hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
			If hnextitem=0 And UBound(parents)>0 Then
				'hnextitem=treeview_getnextitem(htree,parents(UBound(parents)),TVGN_NEXT)
				hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), parents(UBound(parents)), TVGN_NEXT))
				ReDim Preserve parents(UBound(parents)-1)
				s &=crlf
			EndIf
		endif
	Wend
	EndIf
	s= us & crlf & crlf & s
	'Return sysallocstringbytelen(s,len(s))
	Return cast(zstring ptr, sysallocstringbytelen(s, len(s)))
End Function


Function EnumModule2 cdecl Alias "EnumModule2"(hTree As Long, hParent() As long) As zString Ptr
   Dim Parents() As Long
   Dim hNextItem As Long
   Dim HasChilds As Long
   Dim token As String, tldesc As String
   Dim hItem as Long

   Dim s as String,us as string
	redim Parents(0)
   hItem = hparent(tkind_module)
   tldesc = TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem, TVGN_PARENT)))
   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
   If haschilds Then
      s &= "'" & String(80, "=") & crlf
      s &= "'Module - " & tldesc & crlf
      s &= "'" & String(80, "=") & crlf
      hnextitem = haschilds
      While hnextitem

         token = TreeView_GetItemText(hTree, hNextItem)
         If UBound(parents) = 0 Then
            s &= "'" & String(80, "=") & crlf
            s &= "'" & str_parse(token, "?", 1) & "     ?" & str_parse(token, "?", 2) & crlf
            s &= "'" & String(80, "=") & crlf
         Else
            s &= str_parse(token, "?", 1) & "     '" & str_parse(token, "?", 2) & crlf
         End If

         haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hnextitem))
         If haschilds Then
            ReDim Preserve parents(UBound(parents) + 1)
            parents(UBound(parents)) = hnextitem
            hnextitem = haschilds
         Else
            hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
            If hnextitem = 0 And UBound(parents) > 0 Then
               hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), parents(UBound(parents)), TVGN_NEXT))
               ReDim Preserve parents(UBound(parents) - 1)
               s &= crlf
            End If
         end if
      Wend
   End If
   Return cast(zstring ptr, sysallocstringbytelen(s, len(s)))
End Function


Function EnumAlias cdecl Alias "EnumAlias"(hTree As Long, hParent() As long) As zString Ptr
   Dim Parents() As Long
   Dim hNextItem As Long
   Dim HasChilds As Long
   Dim token As String
   Dim hItem as Long
   Dim as String s, tldesc
	redim Parents(0)
   hItem = hparent(tkind_alias)
   tldesc = TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem, TVGN_PARENT)))
   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
   If haschilds Then
      s &= "'" & String(80, "=") & crlf
      s &= "'Alias - " & tldesc & crlf
      s &= "'" & String(80, "=") & crlf
      hnextitem = haschilds
      While hnextitem
         token = TreeView_GetItemText(hTree, hNextItem)
         s &= token & crlf
         hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
      Wend
   End If
   Return cast(zstring ptr, sysallocstringbytelen(s, len(s)))
End Function

Function EnumCoClass cdecl Alias "EnumCoClass"(hTree As Long, hParent() As long) As zString Ptr
   Dim Parents() As Long
   Dim hNextItem As Long
   Dim HasChilds As Long
   Dim token As String, progid As WString Ptr, clssid As clsid
   Dim As String clsids, progids, tldesc
   Dim hItem as long

   Dim s as string, str1 as string
	redim Parents(0)
	GetPrefix()
	str1 = trim(prefix, any "-_ ")

	hItem = hparent(tkind_coclass)
   tldesc = TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem, TVGN_PARENT)))
   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
   If haschilds Then
      hnextitem = haschilds
      While hnextitem
         token = TreeView_GetItemText(hTree, hNextItem)
         clsids &= "   Const CLSID_" & str1 & str_parse(token, "?", 1) & "=" & dq & str_parse(token, "?", 2) & dq & "     '" & str_parse(token, "?", 3) & crlf
         progid = Stringtobstr(str_parse(token, "?", 2))
         clsidfromString(progid, @clssid)
         sysfreeString(progid)
         progidfromclsid(@clssid, @progid)
			If Len(*progid) Then progids &="Const ProgID_" & str_parse(token,"?",1) & "=" & dq & *progid & dq & crlf
         sysfreeString(progid)
         hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
      Wend
   End If
	s &= "' Warning : Do not use vTable and Invoke/Dispatch methods on the same program !" & crlf
	s &= "'           It will be duplicated constants and type confusion "& crlf
	s &= "'           or use different prefix to differentiate them ..." & crlf & crlf
   s &= "'" & String(80, "=") & crlf
   s &= "'CLSID - " & tldesc & crlf
   s &= "'" & String(80, "=") & crlf
   s &= clsids & crlf
   s &= "'" & String(80, "=") & crlf
   s &= "'ProgID - " & tldesc & crlf
   s &= "'" & String(80, "=") & crlf
   s &= progids & crlf
   Return cast(zstring ptr, SysAllocStringByteLen(s, len(s)))
End Function





Function Full_control(xface as string, s_connect as string, s_disc as string, L1 as integer) as string
   dim as string s
	if L1 THEN
		if notsure = 1 THEN
			s = crlf  & tb & "' The events are not connected ....." & crlf
			s &= tb & "' not sure : more than 1 connection to connect, check yourself ... if ok uncomment or modify " & crlf
			s &= tb & "'   Declare " & s_connect & "(ByVal As any Ptr,ByRef As Dword) As Integer" & crlf
			s &= tb & "'   Declare " & s_disc & "(ByVal As any Ptr, ByVal As Dword) As Integer"  & crlf & crlf & crlf
		else
			s = crlf &  tb & "' The events can be connected ....." & crlf
			s &= tb & "#Define Decl_Ev1_" & xface & " Declare Function " & s_connect & "(ByVal As any Ptr, ByRef As dword) As Integer" & crlf
			s &= tb & "#Define Decl_Ev2_" & xface & " Declare Function " & s_disc & "(ByVal As any Ptr, ByVal As dword) As Integer"  & crlf
			s &= tb & "Decl_Ev1_" & xface & "      ' indirect declare to avoid FireFly duplication " & crlf
			s &= tb & "Decl_Ev2_" & xface & "      ' indirect declare to avoid FireFly duplication " & crlf
		END IF
	END IF
	s &= "#EndIf" & crlf & crlf & crlf
	s &= "Sub " & xface & "Call_Init()			' be called from initialization of the control form" & crlf
   s &= tb & xface & "Obj_Ptr = AxCreate_Object( " & xface & "OcxHwnd  )  			             'get object control address with control hwnd" & crlf & crlf
	s &= tb & "'    if use of vTable functions ...  Dim Shared " & xface & "pVTI As ... with the right TYPE see in vtable code" & crlf
	s &= tb & "'    " & xface & "pVTI = " & xface & "Obj_Ptr   ':  " & xface & "pVTI->lpvtbl->SetColor( " & xface & "pVTI ,... )" & crlf
   s &= tb & "'                            or Ax_Vt ( " & xface & "pVTI, SetColor, arg... )  or  Ax_Vt0 ( " & xface & "pVTI, About )" & crlf & crlf
   s &= tb & "'    if use of Invoke/dispatch functions ...  Dim Shared " & xface & "pDisp As ... with the right TYPE see in Invoke code" & crlf
	s &= tb & "'    SetObj ( @" & xface & "pDisp , " & xface & "Obj_Ptr )   ':   AxCall " & xface & "pDisp.PutValue , vptr(x) , vptr(y) , ..." & crlf & crlf

	if L1 THEN
		if notsure = 1 THEN
			s &= tb & "' not sure : more than 1 connection to connect, check yourself ... if ok uncomment or modify " & crlf
			s &= tb & "'  " & s_connect & "( " & xface & "Obj_Ptr , " & xface & "Obj_Event ) 'connect object with its event" & crlf & crlf
		else
			s &= tb &"' The events are now connected ....." & crlf
			s &= tb & s_connect & "( " & xface & "Obj_Ptr , " & xface & "Obj_Event ) 'connect object with its event" & crlf & crlf
		end if
	END IF
   s &= tb & xface & "Call_Sett() 'initial settings if you want some" & crlf
   s &= "End Sub" & crlf & crlf & crlf
   s &= "Sub " & xface & "Call_OnClose() ' normaly be called from close form command" & crlf
   if L1 THEN
		if notsure = 1 THEN
			s &= tb & "' not sure : more than 1 connection to connect, check yourself ... if ok uncomment or modify " & crlf
			s &= tb & "'  " & s_disc & "( " & xface & "Obj_Ptr , " & xface & "Obj_Event ) 'disconnect  event from object" & crlf
		else
			s &= tb &"' The events are now disconnected ....." & crlf
			s &= tb & s_disc & " ( " & xface & "Obj_Ptr , " & xface & "Obj_Event ) 'disconnect  event from object" & crlf
		end if
	END IF
   s &= tb & " AxRelease_Object(" & xface & "Obj_Ptr)     'release object" & crlf
   s &= tb & "'AxStop()         'only one by project, better on the WM_close of last Form" & crlf
   s &= "End Sub" & crlf & crlf & crlf
   s &= "Sub " & xface & "Call_Sett()   'initial settings here" & crlf
   s &= tb & "'ex put/ get /call values  " & crlf
   s &= "End Sub" & crlf
   s &= "'================================================================================" & crlf
	if isevent=0 THEN
			s &= "'  No Events for that ActiveX Control interface " & crlf
	elseif L1 =0 and isevent=1 THEN
			s &= "'  Events exist for that ActiveX Control interface, if needed select Template with events !" & crlf
	end if
	s &= "'================================================================================" & crlf
   function = s
END FUNCTION

Function Simple_control(xface as string, ret1 as string, tldesc as string, progids as string, l1 as integer) as string
   dim as string s
   s = "'" & String(80, "=") & crlf
   s &= "'ProgID - " & tldesc & crlf
   s &= "'" & String(80, "=") & crlf
   s &= progids & crlf

   if ret1 = " ProgID " then
      s &= "'    not unique ProgID   - Select only 1 CoClass in tree to get the one needed to create object " & crlf & crlf
		function = s
		if xfull = 0 THEN exit function
	elseif ret1 = " No ProgID " then
		s &= "'    No ProgID   - Select 1 CoClass in tree to get the one needed to create object " & crlf & crlf
		function = s
		if xfull = 0 THEN exit function
	end if
   s &= "'" & String(80, "=") & crlf & crlf
   s &= "'   Use Prefix in AxSuite3 to differenciate the functions names / variables names" & crlf & crlf
   s &= "'================================================================================" & crlf
   s &= "'for that control :  use  variable Classname > Var_ATL_Win_     and Caption = " & ret1  & crlf
   s &= "'================================================================================" & crlf
   s &= "Dim Shared As any ptr      " & xface & "Obj_Ptr      ' object Ptr   " & crlf
   if l1 = 1 then s &= "Dim Shared As Dword        " & xface & "Obj_Event    ' cookie for object events" & crlf
	s &= "Dim Shared As hwnd         " & xface & "OcxHwnd      ' Ocx handle " & crlf & crlf
   s &=  "' If use of Firefly Custom Control for OCX,  this following function is not needed, you can comment it... " & crlf
	s &=  "Function " & xface & "WinOcx  ( hparent as hwnd ) as hwnd 		' make window for control"& crlf
	s &= tb & "Dim as integer x= 0 								' x left position , change for your need" & crlf
	s &= tb & "Dim as integer y= 0 								' y top position , change for your need" & crlf
	s &= tb & "Dim as integer w= 200 							' w width  , change for your need" & crlf
	s &= tb & "Dim as integer h= 150 							' h heigth  , change for your need" & crlf
	s &= tb & "dim as string N = """ & xface & " WinOcxName1""  					' Name of the window , change for your need" & crlf
	s &= tb & "dim as string P = """ & ret1 & """  			' prodId of the window , change for your need" & crlf
	s &= tb & xface & "OcxHwnd = AxWinChild( hparent , N , P , x , y , w , h ) 		' adapt to your needs  see Axlite.bi" & crlf
	s &= tb & xface & "'OcxHwnd = AxWinTool( hparent , N , P , x , y , w , h ) 		' or try this for Toll floating window" & crlf
	s &= tb & xface & "'OcxHwnd = AxWinFull( hparent , N , P , x , y , w , h ) 		' or try this for Full floating window" & crlf
	s &= tb & "Function = " & xface & "OcxHwnd" & crlf
	s &= "End Function " & crlf & crlf
	s &= "#IfnDef _PARSED_MOD_FF3_         'Defined in FireFly to avoid duplications " & crlf & crlf
	s &= tb & "#Define Decl_Set_" & xface & " Declare Sub " & xface & "Call_Sett()" & crlf
	s &= tb & "Decl_Set_" & xface & "      ' indirect declare to avoid FireFly duplication " & crlf & crlf
   function = s
END FUNCTION

Function Full_nowin(xface as string, s_connect as string, s_disc as string, L1 as integer,ret1 as string) as string
   dim as string s
	if L1 THEN
		if notsure = 1 THEN
			s = crlf  & tb & "' The events are not connected ....." & crlf
			s &= tb & "' not sure : more than 1 connection to connect, check yourself ... if ok uncomment or modify " & crlf
			s &= tb & "'   Declare " & s_connect & "(ByVal As any Ptr, ByRef As Dword) As Integer" & crlf
			s &= tb & "'   Declare " & s_disc & "(ByVal As any Ptr, ByVal As Dword) As Integer"  & crlf & crlf & crlf
		else
			s = crlf & tb & "' The events can be connected ....." & crlf
			s &= tb & "#Define Decl_Ev1_" & xface & " Declare Function " & s_connect & "(ByVal As any Ptr, ByRef As dword) As Integer" & crlf
			s &= tb & "#Define Decl_Ev2_" & xface & " Declare Function " & s_disc & "(ByVal As any Ptr, ByVal As dword) As Integer"  & crlf
			s &= tb & "Decl_Ev1_" & xface & "      ' indirect declare to avoid FireFly duplication " & crlf
			s &= tb & "Decl_Ev2_" & xface & "      ' indirect declare to avoid FireFly duplication " & crlf
		END IF
	END IF
	s &= "#EndIf" & crlf & crlf & crlf
	s &= "Sub " & xface & "Call_Init()				' be called from initialization " & crlf
   s &= tb & xface & "Obj_Ptr = AxCreate_Object( " & ret1 & " )  			             'get object control address with prodid" & crlf & crlf
	s &= tb & "'    if use of vTable functions ...  Dim Shared " & xface & "pVTI As ... with the right TYPE see in vtable code" & crlf
	s &= tb & "'    " & xface & "pVTI = " & xface & "Obj_Ptr   ':  " & xface & "pVTI->lpvtbl->SetColor( " & xface & "pVTI ,... )" & crlf
   s &= tb & "'                            or Ax_Vt ( " & xface & "pVTI, SetColor, arg... )  or  Ax_Vt0 ( " & xface & "pVTI, About )" & crlf & crlf
	s &= tb & "'    if use of Invoke/dispatch functions ...  Dim Shared " & xface & "pDisp As ... with the right TYPE see in Invoke code" & crlf
	s &= tb & "'    SetObj ( @" & xface & "pDisp , " & xface & "Obj_Ptr )   ':   AxCall " & xface & "pDisp.PutValue , vptr(x) , vptr(y) , ..." & crlf & crlf

	if L1 THEN
		if notsure = 1 THEN
			s &= tb & "' not sure : more than 1 connection to connect, check yourself ... if ok uncomment or modify " & crlf
			s &= tb & "'  " & s_connect & " ( " & xface & "Obj_Ptr , " & xface & "Obj_Event ) 'connect object with its event" & crlf & crlf
		else
			s &= tb &"' The events are now connected ....." & crlf
			s &= tb & s_connect & " (  " & xface & "Obj_Ptr , " & xface & "Obj_Event ) 'connect object with its event" & crlf & crlf
		end if
	END IF
   s &= tb & xface & "Call_Sett() 'initial settings if you want some" & crlf
   s &= "End Sub" & crlf & crlf & crlf
   s &= "Sub " & xface & "Call_OnClose() ' normaly be called from close form command" & crlf
   if L1 THEN
		if notsure = 1 THEN
			s &= tb & "' not sure : more than 1 connection to connect, check yourself ... if ok uncomment or modify " & crlf
			s &= tb & "'  " & s_disc & "( " & xface & "Obj_Ptr , " & xface & "Obj_Event ) 'disconnect  event from object" & crlf
		else
			s &= tb &"' The events are now disconnected ....." & crlf
			s &= tb & s_disc & "( " & xface & "Obj_Ptr , " & xface & "Obj_Event ) 'disconnect  event from object" & crlf
		end if
	END IF
   s &= tb & " AxRelease_Object( " & xface & "Obj_Ptr )     'release object" & crlf
   s &= tb & "'AxStop()         'only one by project, better on the WM_close of last Form" & crlf
   s &= "End Sub" & crlf & crlf & crlf
   s &= "Sub " & xface & "Call_Sett()   'initial settings here" & crlf
   s &= tb & "'ex put/ get /call values  " & crlf
   s &= "End Sub" & crlf
   s &= "'================================================================================" & crlf
	if isevent=0 THEN
			s &= "'  No Events for that ActiveX Control interface " & crlf
	elseif L1 =0 and isevent=1 THEN
			s &= "'  Events exist for that ActiveX Control interface, if needed select Template with events !" & crlf
	end if
	s &= "'================================================================================" & crlf
   function = s
END FUNCTION


Function Simple_nowin(xface as string, ret1 as string, tldesc as string, progids as string, l1 as integer) as string
   dim as string s
   s = "'" & String(80, "=") & crlf
   s &= "'ProgID - " & tldesc & crlf
   s &= "'" & String(80, "=") & crlf
   s &= progids & crlf

   if ret1 = " ProgID " then
      s &= "'    not unique ProgID   - Select only 1 CoClass in tree to get the one needed to create object " & crlf & crlf
		function = s
		if xfull = 0 THEN exit function
	elseif ret1 = " No ProgID " then
		s &= "'    No ProgID   - Select 1 CoClass in tree to get the one needed to create object " & crlf & crlf
		function = s
		if xfull = 0 THEN exit function
	end if
   s &= "'" & String(80, "=") & crlf & crlf
   s &= "'   Use Prefix in AxSuite3 to differenciate the functions names / variables names" & crlf & crlf
   s &= "'================================================================================" & crlf
   s &= "Dim Shared As any ptr      " & xface & "Obj_Ptr      ' object Ptr   " & crlf
   if l1 = 1 then s &= "Dim Shared As Dword        " & xface & "Obj_Event    ' cookie for object events" & crlf & crlf & crlf
   s &= "#IfnDef _PARSED_MOD_FF3_         'Defined in FireFly to avoid duplications " & crlf & crlf
	s &= tb & "#Define Decl_Set_" & xface & " Declare Sub " & xface & "Call_Sett()" & crlf
	s &= tb & "Decl_Set_" & xface & "      ' indirect declare to avoid FireFly duplication " & crlf & crlf
   function = s
END FUNCTION

Function Ini_templ() as string
   dim as string s
   s = "'" & String(80, "=") & crlf
   s &= "'Template to use OCX / ActiveX / COM " & crlf
   s &= "'" & String(80, "=") & crlf & crlf
'	if nowin=0  THEN
'		s &= "'#Define useATL71             'to use ATL71.dll  uncomment,  else commented use of ATL.dll"& crlf & crlf
'  END IF
	if nowin<>0  THEN
		s &= "'#Define Ax_NoAtl        				'to use Ax_Lite.bi without atl.dll , when no control window ( reduce size of exe)"& crlf & crlf
	END IF
	s &= "#Include Once ""Windows.bi""        ' Windows specific, if not defined yet" & crlf
	s &= "#Include Once ""Ax_lite.bi""        ' this one is needed" & crlf & crlf
	's &= "'#Include Once ""Ax_lite_lib.bi""    ' use libAx_lite.a static lib   ( reduce size of exe)" & crlf

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




sub EnumCoClassB (hTree As Long, hParent() As long)
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
'   Dim xface as string, ss22 as string
'	GetPrefix()
'	xface = trim(prefix, any "-_ ")
'   if xface <> "" THEN xface = xface & "_"
	hItem = hparent(tkind_coclass)

   tldesc = TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem, TVGN_PARENT)))
   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
	'tldesc = TreeView_GetItemText(hTree, treeview_getnextitem (htree, hitem, TVGN_PARENT))
   'haschilds = treeview_getchild(htree, hitem)
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
'					ss22= dq & str_parse(token, "?", 2) & dq
'					clsids &="'ss22 = " & ss22 & crlf
'					s &= clsids
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
									'ss22= dq & str_parse(token, "?", 2) & dq
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
							'ss22 = dq & str_parse(token, "?", 2) & dq
						end if
					END IF
				END IF
				if inot=0 and *progid <> "" THEN nitem += 1
				sysfreeString(progid)
				hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
				'hnextitem = treeview_getnextitem( htree, hnextitem, TVGN_NEXT)
      Wend
   End If
	sevent =""
   if nitem = 1 then
		ret1 =  iseul
		for x= 1 to icoclass
			if coclass(x,2)= ret1 then
				sevent = coclass(x,7)
				'print "sevent" ; sevent
				exit for
			end if
      NEXT
   end if

End sub


Function EnumCoClassA (hTree As Long, hParent() As long,byRef clsids as string,byRef progids as string,Byref tldesc as string) As String
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
'   Dim xface as string, ss22 as string
'	GetPrefix()
'	xface = trim(prefix, any "-_ ")
'   if xface <> "" THEN xface = xface & "_"
	hItem = hparent(tkind_coclass)
	clsids = ""
   tldesc = TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem, TVGN_PARENT)))
   iseul2 = tldesc
	haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
	'tldesc = TreeView_GetItemText(hTree, treeview_getnextitem (htree, hitem, TVGN_PARENT))
   'haschilds = treeview_getchild(htree, hitem)
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

'					ss22= dq & str_parse(token, "?", 2) & dq
'					clsids &="'ss22 = " & ss22 & crlf
'					s &= clsids
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
									'ss22= dq & str_parse(token, "?", 2) & dq
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
							'ss22 = dq & str_parse(token, "?", 2) & dq
						end if
					END IF
				END IF
				if inot=0 and *progid <> "" THEN nitem += 1
				sysfreeString(progid)
				hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
				'hnextitem = treeview_getnextitem( htree, hnextitem, TVGN_NEXT)
      Wend
   End If
	sevent =""
   if nitem = 1 then
      FF_ClipboardSetText(iseul)
		ret1 =  iseul
'		for x= 1 to icoclass
'			if coclass(x,2)= ret1 then
'				sevent = coclass(x,7)
'				print "sevent" ; sevent
'			end if
'      NEXT
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


Function EnumCoClass1 cdecl Alias "EnumCoClass1"(hTree As Long, hParent() As long, l1 as integer, _
         s_conect as string, s_disc as string) As zString Ptr
'   Dim Parents() As Long
'   Dim hNextItem As Long
'   Dim HasChilds As Long
'   Dim token As String, progid As WString Ptr, clssid As clsid
'   Dim As String ret2
'   Dim hItem as long
'   Dim nitem as long
'   Dim u1 as integer, n1 as integer,inot as integer
	Dim As String clsids, progids, tldesc , ret1
   Dim s as string,iseul as string
   Dim xface as string, ss22 as string
	Dim x as integer
	GetPrefix()
	xface = trim(prefix, any "-_ ")
   if xface <> "" THEN xface = xface & "_"
	isevent=0
'	hItem = hparent(tkind_coclass)

'   tldesc = TreeView_GetItemText(hTree, cast(long, treeview_getnextitem(cast(hwnd, htree), hitem, TVGN_PARENT)))
'   haschilds = cast(long, treeview_getchild(cast(hwnd, htree), hitem))
'	'tldesc = TreeView_GetItemText(hTree, treeview_getnextitem (htree, hitem, TVGN_PARENT))
'   'haschilds = treeview_getchild(htree, hitem)
'   If haschilds Then
'      hnextitem = haschilds
'      While hnextitem
'         token = TreeView_GetItemText1(hTree, hNextItem)
'			if instr(token,"          'NOT'") THEN
'				token=left(token , len(token)-len("          'NOT'"))
'				inot=1
'			else
'				inot=0
'			end if
'				if inot=0 THEN
'					clsids &= "  ' Const CLSID_" & str_parse(token, "?", 1) & "=" & dq & str_parse(token, "?", 2) & dq & "     '" & str_parse(token, "?", 3) & crlf
''					ss22= dq & str_parse(token, "?", 2) & dq
''					clsids &="'ss22 = " & ss22 & crlf
''					s &= clsids
'				end if
'				progid = Stringtobstr(str_parse(token, "?", 2))
'				clsidfromString(progid, @clssid)
'				sysfreeString(progid)
'				progidfromclsid(@clssid, @progid)
'				ret1 = *progid
'				if inot=1 THEN ret1 = ""
'				ret2 = ret1
'				if ret1 <> "" THEN
'					n1 = 0 :u1 = 0
'					do
'						u1 = instr(u1 + 1, ret1, ".")
'						if u1 > 0 THEN
'							n1 = n1 + 1
'							if n1 = 2 THEN
'								ret1 = left(ret1, u1 - 1)
'								progids &= "  ' Const I_ProgID_" & str_parse(token, "?", 1) & "=" & dq & ret1 & dq & _
'										"                 '*** Version independent ProgID" & crlf
'								if nitem = 0 THEN
'									iseul = ret1
'									'ss22= dq & str_parse(token, "?", 2) & dq
'								end if
'								exit do
'							END IF
'						END IF
'					LOOP UNTIL u1 = 0
'				END IF
'				if ret2 <> ret1 THEN
'					progids &= "  ' Const ProgID_" & str_parse(token, "?", 1) & "=" & dq & *progid & dq & _
'							"                                       'Version dependent ProgID" & crlf & crlf
'				else
'					if *progid="" THEN
'					elseif ret1 = "" THEN
'						if inot=0 THEN progids &= "  ' Const ProgID_" & str_parse(token, "?", 1) & "=" & dq & *progid & dq & crlf & crlf
'					else
'						progids &= "  ' Const ProgID_" & str_parse(token, "?", 1) & "=" & dq & *progid & dq & _
'								"                 '*** Version independent ProgID" & crlf & crlf
'						if nitem = 0 THEN
'							iseul = *progid
'							'ss22 = dq & str_parse(token, "?", 2) & dq
'						end if
'					END IF
'				END IF
'				if inot=0 and *progid <> "" THEN nitem += 1
'				sysfreeString(progid)
'				hnextitem = cast(long, treeview_getnextitem(cast(hwnd, htree), hnextitem, TVGN_NEXT))
'				'hnextitem = treeview_getnextitem( htree, hnextitem, TVGN_NEXT)
'      Wend
'   End If
'	sevent =""
'   if nitem = 1 then
'      FF_ClipboardSetText(iseul)
'		ret1 =  iseul
'		for x= 1 to icoclass
'			if coclass(x,2)= ret1 then
'				sevent = coclass(x,7)
'				print "sevent" ; sevent
'			end if
'      NEXT
'	elseif nitem = 0 then
'		ret1 = " No ProgID "
'   else
'      FF_ClipboardSetText( "")
'      ret1 = " ProgID "
'   end if

	ret1=EnumCoClassA(hTree , hParent(),clsids, progids, tldesc )
	if nowin=0 THEN
		s = Simple_control(xface, ret1, tldesc, progids, l1)
		if (ret1 <> " ProgID " and ret1 <> " No ProgID " ) or xfull =1 THEN
			xret1=0

			for x= 1 to icoclass
				if coclass(x,2)=ret1 THEN
					'ret1 = dq & coclass(x,2) & dq & " , " & dq & coclass(x,5)& dq
					if coclass(x,7)<>"" then isevent=1
					exit for
            END IF
         NEXT

			s &= Full_control(xface,s_conect, s_disc, l1)
		else
			xret1=1
			s="'   " & ret1
			messagebox(getactivewindow()," ProgId not defined... Select only 1 Coclass on the tree view ( with ProgId)."  ,"Error",MB_ICONERROR)
			goto ddetour
		END IF
	else
		s = Simple_nowin(xface, ret1, tldesc, progids, l1)
		if (ret1 <> " ProgID " and ret1 <> " No ProgID " ) or xfull =1 THEN
			xret1=0
			'coclass(i,2)= s1
			for x= 1 to icoclass
				if coclass(x,2)=ret1 THEN
					ret1 = dq & coclass(x,2) & dq & " , " & dq & coclass(x,5)& dq
					if coclass(x,7)<>"" then isevent=1
					exit for
            END IF
         NEXT
			s &= Full_nowin(xface,s_conect, s_disc, l1,ret1)
		else
			xret1=1
			s="'   " & ret1
			messagebox(getactivewindow()," ProgId not defined... Select 1 Coclass only on the tree view ( with ProgId)."  ,"Error",MB_ICONERROR)
			goto ddetour
		END IF
   END IF
ddetour:
   Return cast(zstring ptr, SysAllocStringByteLen(s, len(s)))

End Function


Function EnumCoClass3 (hTree As Long, hParent() As long) As zString Ptr
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

