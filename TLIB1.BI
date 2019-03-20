

Function RegSearchWin32(tszKey AS string) AS STRING
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

FUNCTION RegEnumDirectory(sKey AS STRING) AS STRING
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

   ' Searches the HKEY_CLASSES_ROOT\TypeLib\<LIBID> node.
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

   ' win32 node
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

Function CheckEnvVar(BYVAL sTextIn AS STRING) AS STRING
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

Function RegEnumVersions(sCLSID AS STRING) AS LONG
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
   ' // Searches the HKEY_CLASSES_ROOT\TypeLib\<LIBID> node.
   szKey = "TypeLib\" & sCLSID

   hr = RegOpenKeyEx(hkey_CLASSES_ROOT, @szKey, 0, KEY_READ, @hKey)
   IF hr <> ERROR_SUCCESS THEN EXIT Function
   IF hKey = 0 THEN EXIT FUNCTION
   dwIdx = 0
   ' // Open the subtrees of the different versions of the TypeLib library
   ' // and store the filenames, descriptions, paths and CLSIDs in a 4D array.
   DO
      dwname = max_path
      dwclass = max_path
      hr = RegEnumKeyEx(hKey, dwIdx, @szKeyName, @dwname, 0, @szClass, @dwclass, @ft)
      IF hr = ERROR_NO_MORE_ITEMS THEN EXIT DO
      tsz0 = UCASE$(sCLSID)

      ' -- Get default value -------------------------------------------
      verIdx = 0
      cvaluename = max_path
      cbdata = max_path
      tsz = szKey & "\" & szKeyName
      hr = RegOpenKeyEx(hkey_CLASSES_ROOT, @tsz, 0, KEY_READ, @hVerKey)
      'Default value
      hr = RegEnumValue(hVerKey, verIdx, @szValueName, @cValueName, NULL, @KeyType, @szKeyValue, @cbData)
      RegCloseKey hVerKey
      IF szValueName = "" THEN
         tsz1 = szKeyvalue
      Else
         tsz1 = szValueName
      End IF

      '-----------------------------------------------------------------
      sPath = RegEnumDirectory(szKey & "\" & szKeyName)
      ' Check for environment variables
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
   ' // Close the registry
   RegCloseKey hKey
   function = 1
END FUNCTION

' *********************************************************************************************
' Enumerates all the typelibs.
' *********************************************************************************************
SUB RegEnumTypeLibs
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


   ' // Open the HKEY_CLASSES_ROOT\TypeLib subtree.
   szKey = "TypeLib"
   hr = RegOpenKeyEx(HKEY_CLASSES_ROOT, @szKey, 0, key_read, @hKey)
   IF hr <> ERROR_SUCCESS THEN EXIT SUB
   IF hKey = 0 THEN EXIT SUB
   ' // Parse all the TypeLib subtree and get the CLSIDs of all the TypeLibs.
   dwIdx = 0
   Do
      dwname = MAX_PATH :dwclass = max_path
      hr = RegEnumKeyEx(hKey, dwIdx, @szKeyName, @dwname, 0, @szClass, @dwclass, @ft)
      IF hr = ERROR_NO_MORE_ITEMS THEN EXIT DO
      dwIdx += 1
      RegEnumVersions(left(szKeyName, dwname))
   LOOP
   ' // Close the registry
   RegCloseKey(hKey)
End SUB

#define ListView_SetItemCountEx(l, c, f) SendMessage(cast(hwnd, L), LVM_SETITEMCOUNT, c, F)
FUNCTION SearchTypeLibs() AS LONG
   Dim i AS LONG
   dim ttxt as string
   SendMessage hlist, LVM_DELETEALLITEMS, 0, 0
   nt1 = 0
   RegEnumTypeLibs
   SetWindowText(hmain, "AxSuite3           ....  " & str(nt1) & " registered libs")
   FF_ListView_SetSelectedItem(hlist, 0)
   'lvsetheader(hlist,"GUID|Description                  ....  listed : " & str(nt1) & " libs |File|Path")
   function = 1
End Function

'Custom Type string
Function CTStr(reftype As hreftype, pti As lptypeinfo) As String
   Dim As lptypeinfo pcti
   Dim As WString Ptr iname, ihelp
   Dim hr As hresult

   hr = pTI -> lpvtbl -> GetRefTypeInfo(pti, refType, @pCTI)
   If hr Then return "UnknownCustomType"
   hr = pCTI -> lpvtbl -> GetDocumentation(pcti, - 1, @iname, 0, 0, 0)
   If hr then return "UnknownCustomType"
   Return " As " & *iname
End Function

'typedesc string
Function TDStr(ptD As typedesc ptr, pTI As lptypeinfo) As String
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
'		if instr(lcase(oss),"collection") THEN
'			print "collection : " & oss
'      END IF
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
		'Case VT_LPTSTR :Return " As LPTSTR"
		case else : Return " As any ptr"

   End Select
End Function

'vardesc string
Function VDStr(ByVal pvd As lpVARDESC, byval pti As lptypeinfo, tk As Integer) As String
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

Sub mkConst(pti As lptypeinfo, pta As lptypeattr, tk As typekind)
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

Sub mkInterface(pti As lptypeinfo, pta As lptypeattr, tk As typekind)
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
      'hroot = TreeViewInsertItem(cast(long, htree), hParent(tk), prefix & *iname & "?" & *iids & "?" & *ihelp)
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

Sub fs_ENUM(ByVal i As uint)
   Dim oss As String
   Dim pti As lpTypeInfo, pta As lpTypeAttr
   pTLi -> lpvtbl -> GetTypeInfo(ptli, i, @pti)
   pTI -> lpvtbl -> GetTypeAttr(pti, @ptA)
   mkConst(pti, pta, tkind_enum)
End Sub

Sub fs_RECORD(ByVal i As uint)
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

Sub fs_MODULE(ByVal i As uint)
   Dim hr As Long
   Dim pti As lpTypeInfo, pta As lpTypeAttr

   hr = pTLi -> lpvtbl -> GetTypeInfo(ptli, i, @pti)
   hr = pTI -> lpvtbl -> GetTypeAttr(pti, @ptA)
   If pta -> cvars Then mkConst(pti, pta, tkind_module)
   If pta -> cfuncs Then mkInterface(pti, pta, tkind_module)
End Sub

Sub fs_INTERFACE(ByVal i As uint)
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

Sub fs_DISPATCH(ByVal i As uint)
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

sub fs_COCLASS(ByVal i As uint)
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
'	CLSIDFromProgID(StringToBSTR(s1), @ClassID)
'	print "CLSIDFromProgID", s1
'	StringFromCLSID(@ClassID,@wcl)
'	print "StringFromCLSID"
	TreeViewInsertItem(cast(long, htree), hroot,"ProgID Version Independant >>> " & s1)
	s0=*iname
	if ubound(coclass)= icoclass then redim preserve coclass(1 to icoclass + 25 , 1 to 7) as string
	icoclass += 1
	'print "icoclass = "; icoclass
	coclass(icoclass,1)= s0
	coclass(icoclass,2)= s1
	coclass(icoclass,3)= *clsids : coclass(icoclass,3)=trim(coclass(icoclass,3),any " " & chr(0) & chr(9)& chr(10)& chr(13) )
   sysfreestring(iname)
   sysfreestring(clsids)
   sysfreestring(ihelp)
   cIT = pTa -> cImplTypes
	'print " cIT = " ; cIT
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
            ' Store the name of the event interface for later use
            s = *iname
				's = prefix & *iname
            putevtlist(s)
				'if ievent= 0 THEN
					TreeViewInsertItem(cast(long, htree), hParent(8), s0 & " ( " & s & " )" )
					if lITF = 3 THEN
						TreeViewInsertItem(cast(long, htree), hroot, "Default Event Interface >>> " & *iname  & " >>> " & *interf )
						coclass(icoclass,6)= *iname
						coclass(icoclass,7)= *interf : coclass(icoclass,7)=trim(coclass(icoclass,7),any " " & chr(0) & chr(9)& chr(10)& chr(13))
						ievent=1
					else
						TreeViewInsertItem(cast(long, htree), hroot, "Event Interface >>> " & *iname  & " >>> " & *interf )
					END IF
'					coclass(icoclass,6)= *iname
'					coclass(icoclass,7)= *interf : coclass(icoclass,7)=trim(coclass(icoclass,7),any " " & chr(0) & chr(9)& chr(10)& chr(13))
'					ievent=1
					'print "icoclass= " ; icoclass,s0,*iname,*interf
            'END IF
			elseif lITF = 1 Then ' //  Default interface
				'if iclass= 0 THEN
					TreeViewInsertItem(cast(long, htree), hroot,"Default Interface >>> " & *iname  & " >>> " & *interf )
					coclass(icoclass,4)= *iname
					coclass(icoclass,5)= *interf : coclass(icoclass,5)=trim(coclass(icoclass,5),any " " & chr(0) & chr(9)& chr(10)& chr(13))
					iclass= 1
					'print "icoclass= " ; icoclass, s0,*iname,*interf
				'END IF
			elseif lITF = 4 Then ' //  Restricted
					TreeViewInsertItem(cast(long, htree), hroot,"Restricted Interface >>> " & *iname  & " >>> " & *interf )
'					coclass(icoclass,4)= *iname
'					coclass(icoclass,5)= *interf : coclass(icoclass,5)=trim(coclass(icoclass,5),any " " & chr(0) & chr(9)& chr(10)& chr(13))
'					iclass= 1
			elseif lITF = 8 Then ' //  Default vTable
					TreeViewInsertItem(cast(long, htree), hroot,"Default vTable Interface >>> " & *iname  & " >>> " & *interf )
         End If
         sysfreestring(iname)
         sysfreestring(ihelp)
			sysfreestring(interf)
      End If

   Next
End Sub



sub fs_ALIAS(ByVal i As uint)
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

sub fs_UNION(ByVal i As uint)
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

Sub LoadTLI
   Dim As Integer cti, i
   Dim tk As typekind
   sreport = ""
	icoclass=0
	dim as string act1
	redim coclass(1 to 50, 1 to 7) as string
	act1=selpath
	act1=FF_FileName (act1)
	OcxName=act1
	'setwindowtext GetDlgItem(hmain, IDC_SBR1)," Selected :  " & act1 & "    |    Path :  " & selpath
	act1= " Selected File :  " & act1 & "    |    Path :  " & selpath
	sendmessage GetDlgItem(hmain, IDC_SBR1),WM_SETTEXT ,0,cast(lparam,strptr(act1))
   sendmessage hcode, wm_settext, 0, 0
   sendmessage getdlgitem(hmain, idc_edt1), wm_gettext, 10, cast(lparam, @prefix)
	InvalidateRect hmain, ByVal 0, True
	UpdateWindow hmain
   cti = pTLI -> lpvtbl -> gettypeinfocount(pTLI)
   If cti = 0 then exit sub
	'print "cti = " ; cti
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

Sub cleartli
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

Sub showinfo()
	'sreport= ff_replace(sreport,chr(9),"   ")
	sreport= "'   " &  sinfcode & crlf & crlf & sreport
   sendmessage hcode, wm_settext, 0, cast(lparam, StrPtr(sreport))
end Sub
