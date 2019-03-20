#INCLUDE ONCE "windows.bi"                       'if not included earlier

#include once "win/olectl.bi"
#include once "win/ole2.bi"
#include once "win/objbase.bi"

'#Define Ax_WindowLess			'to use without atl.dll , when no control window
'#Define useATL71             'to use ATL71.dll  uncomment,  else commented use of ATL.dll
'#Define _USE_DISPHELPER_     'to use Dishelper lib  uncomment,  else commented  not used


#ifndef __AX_LITE__
   #define __AX_LITE__
	#print ====
   #print ==== info ====> Compiling with Ax_lite.bi <====
   #print ====
   #ifdef _USE_DISPHELPER_
      #ifndef __disphelper_bi__
         #ifndef UNICODE
            #define UNICODE
            #print ==== info ====> #define UNICODE needed for disphelper : is it true ? <====
         #endif

         #include "disphelper/disphelper.bi"
			#print ====
         #print ==== info ====> Compiling with disphelper.bi <====
			#print ====

         #define Ax_DimObjPtr(objName) 	dim as any ptr objName = NULL
         #define Ax_Call 						dhCallMethod
         #define Ax_Put 						dhPutValue
         #define Ax_Set 						dhPutRef
         #define Ax_Get 						dhGetValue

      #endif
   #ENDIF

   #define Ax_FreeStr(bs) 		SysFreeString(cptr(BSTR, bs))
   #define Kill_Str(bs) 		Ax_FreeStr(bs) : bs = NULL



   Type tMember
      DispID                       As dispid
      cDummy                       As UINT
      cArgs                        As UINT
      tKind                        As UINT
   End Type

   '************************************************************************************
   'Event sink common procedure & constants
   '************************************************************************************
   TYPE Events_IDispatchVtbl
      QueryInterface               AS DWORD      ' Returns pointers to supported interfaces
      AddRef                       AS DWORD      ' Increments reference count
      Release                      AS DWORD      ' Decrements reference count
      GetTypeInfoCount             AS DWORD      ' Retrieves the number of type descriptions
      GetTypeInfo                  AS DWORD      ' Retrieves a description of object's programmable interface
      GetIDsOfNames                AS DWORD      ' Maps name of method or property to DispId
      Invoke                       AS DWORD      ' Calls one of the object's methods, or gets/sets one of its properties
      pVtblAddr                    AS DWORD      ' Address of the virtual table
      cRef                         AS DWORD      ' Reference counter
      pthis                        AS DWORD      ' IUnknown or IDispatch of the control that fires the events
   END Type


   dim shared AxScode as scode
   dim shared AxPexcepinfo as excepinfo
   dim shared AxPuArgErr AS uinteger


   Declare Function Variantv(ByRef v As variant) As Double
   Declare Function ToBSTR(cnv_string As String) As BSTR
	Declare Function AxCreate_Object overload(strProgID AS string, ByVal clsctx As Integer = 21) as any ptr


	#Ifndef Ax_WindowLess
		Declare Function AxCreate_Object overload(BYVAL hWndControl AS hwnd) as any ptr
		dim shared as any ptr hLib
		#Ifdef useATL71
			hLib = DylibLoad( "atl71.dll" )
			if hLib = 0 then
				MessageBox( 0, "ATL71.DLL :    is missing !", "Error, exit Program", MB_ICONERROR )
				end
			end if
			function impose()as string
				function = "AtlAxWin71"
			end function
		#Else
			hLib = DylibLoad( "atl.dll" )
			if hLib = 0 then
				MessageBox( 0, "ATL.DLL :    is missing !", "Error, exit Program", MB_ICONERROR )
				end
			end if
			function impose()as string
				function = "AtlAxWin"
			end function
		#EndIf

		dim shared AtlAxWinInit as function ()as integer
		dim shared AtlAxGetControl as function (BYVAL hWnd AS hwnd, Byval pp AS uint ptr) AS uinteger
		dim shared AtlAxAttachControl as function (BYVAL pControl AS any ptr, _
					BYVAL hWnd AS hwnd, ByVal ppUnkContainer AS lpunknown) AS UInteger
		AtlAxWinInit = DylibSymbol( hLib, "AtlAxWinInit" )
		AtlAxGetControl = DylibSymbol( hLib, "AtlAxGetControl")
		AtlAxAttachControl = DylibSymbol( hLib,"AtlAxAttachControl")

		FUNCTION AxWinChild(byVal h_parent as hwnd, name1 as string, progid as string, _
					x as integer, y as integer, w as integer, h as integer, _
					style as integer = WS_visible or WS_child or WS_border, exstyle as integer = 0) as hwnd
			Dim as hwnd h1
			h1 = CreateWindowEx(exstyle, impose(), progid, style, x, y, w, h, _
					h_parent, NULL, GetmoduleHandle(0), NULL)
			setwindowtext h1, name1
			function = h1
		END FUNCTION

		FUNCTION AxWinTool(byVal h_parent as hwnd, name1 as string, progid as string, _
					x as integer, y as integer, w as integer, h as integer, _
					style as integer = WS_visible, exstyle as integer = WS_EX_TOOLWINDOW) as hwnd
			Dim as hwnd h1
			h1 = CreateWindowEx(exstyle, impose(), progid, style, x, y, w, h, _
					h_parent, NULL, GetmoduleHandle(0), NULL)
			setwindowtext h1, name1
			function = h1
		END FUNCTION

		FUNCTION AxWinFull(byVal h_parent as hwnd, name1 as string, progid as string, _
					x as integer, y as integer, w as integer, h as integer, _
					style as integer = WS_visible or WS_OVERLAPPEDWINDOW, exstyle as integer = 0) as hwnd
			Dim as hwnd h1
			h1 = CreateWindowEx(exstyle, impose(), progid, style, x, y, w, h, _
					h_parent, NULL, GetmoduleHandle(0), NULL)
			setwindowtext h1, name1
			function = h1
		END FUNCTION


		Sub AxWinKill(byVal h_Control as hwnd)
			DestroyWindow(h_Control)
		END SUB

		Sub AxWinHide(byVal h_Control as hwnd, byVal h_Parent as hwnd = 0)
			ShowWindow(h_Control, SW_HIDE)
			if h_Parent THEN
				InvalidateRect h_Parent, ByVal 0, True
				UpdateWindow h_Parent
			end if
		END SUB

		Sub AxWinShow(byVal h_Control as hwnd, byVal h_Parent as hwnd = 0)
			ShowWindow(h_Control, SW_SHOW)
			InvalidateRect h_Control, ByVal 0, True
			UpdateWindow h_Control
			if h_Parent THEN
				InvalidateRect h_Parent, ByVal 0, True
				UpdateWindow h_Parent
			end if
		END SUB
		'
		' ****************************************************************************************
		' functions for guid
		' ****************************************************************************************
		'Function s2guid(txt As String) As guid
		'	Static oGuid As guid
		'	iidfromstring(wstr(txt), @oGuid)
		'	Return oGuid
		'End Function
		'
		'Function guid2s(iguid As guid) As string
		'	dim oGuids As WString ptr
		'	stringfromiid(@iGuid, Cast(LPOLESTR ptr, @oguids))
		'	Return *oGuids
		'End Function
		'
		' ****************************************************************************************
		' Retrieves the interface of the ActiveX control given the handle of its ATL container
		' ****************************************************************************************
		SUB AtlAxGetDispatch(BYVAL hWndControl AS hwnd, BYREF ppvObj AS lpvoid)
			Dim ppUnk AS lpunknown
			dim ppDispatch as pvoid
			'dim IID_IDispatch as IID

			' Get the IUnknown of the OCX hosted in the control
			AxScode = AtlAxGetControl(hWndControl, cast(uint ptr, @ppUnk))
			IF AxScode <> 0 OR ppUnk = 0 THEN EXIT SUB
			' Query for the existence of the dispatch interface
			'IIDFromString("{00020400-0000-0000-c000-000000000046}",@IID_IDispatch)
			AxScode = IUnknown_QueryInterface(ppUnk, @IID_IDispatch, @ppDispatch)
			' If not found, return the IUnknown of the control
			IF AxScode <> 0 OR ppDispatch = 0 THEN
				ppvObj = ppUnk
				EXIT SUB
			END IF
			' Release the IUnknown of the control
			IUnknown_Release(ppUnk)
			' Return the retrieved address
			ppvObj = ppDispatch
		End SUB

		Function AxCreate_Object(BYVAL hWndControl AS hwnd) as any ptr
			dim ppvObj AS lpvoid
			AtlAxGetDispatch( hWndControl, ppvObj )
			function=ppvObj
		end function

	   #define IClassFactory2_CreateInstanceLic(T, u, r, i, s, o)(T) -> lpVtbl -> CreateInstanceLic(T, u, r, i, s, o)
		#define IClassFactory2_GetLicInfo(T, u)(T) -> lpVtbl -> GetLicInfo(T, u)
		#define IClassFactory2_RequestLicKey(T, u, r)(T) -> lpVtbl -> RequestLicKey(T, u, r)
		#define IClassFactory2_Release(T)(T) -> lpVtbl -> Release(T)

		' ****************************************************************************************
		' Creates a licensed instance of a visual control (OCX) and attaches it to a window.
		' StrProgID can be the ProgID or the ClsID. If you pass a version dependent ProgID or a ClsID,
		' it will work only with this particular version.
		' hWndControl is the handle of the window and strLicKey the license key.
		' ****************************************************************************************
		FUNCTION AxCreateControlLic(BYVAL strProgID AS LPOLESTR, byval hWndControl AS uinteger, _
					byval strLicKey AS lpwstr) AS LONG
			DIM ppUnknown AS lpunknown                 ' IUnknown pointer
			DIM ppDispatch AS lpdispatch               ' IDispatch pointer
			DIM ppObj AS lpvoid                        ' Dispatch interface of the control
			' IClassFactory2 pointer
			DIM ppClassFactory2 AS IClassFactory2 ptr
			DIM ppUnkContainer AS lpunknown            ' IUnknown of the container
			'DIM IID_NULL as IID               ' Null GUID
			'DIM IID_IUnknown as IID           ' Iunknown GUID
			'DIM IID_IDispatch as IID          ' IDispatch GUID
			'DIM IID_IClassFactory2 as IID     ' IClassFactory2 GUID
			DIM ClassID AS clsid                       ' CLSID

			' Standard interface GUIDs
			'IIDFromString("{00000000-0000-0000-0000-000000000000}",@IID_NULL)
			'IIDFromString("{00000000-0000-0000-C000-000000000046}",@IID_IUnknown)
			'IIDFromString("{00020400-0000-0000-C000-000000000046}",@IID_IDispatch)
			'IIDFromString("{b196b28f-bab4-101a-b69c-00aa00341d07}",@IID_IClassFactory2)
			' Exit if strProgID is a null string
			IF *strProgID = "" THEN
				FUNCTION = E_INVALIDARG
				EXIT FUNCTION
			END If
			' Convert the ProgID in a CLSID
			AxScode = CLSIDFromProgID(strProgID, @ClassID)
			' If it fails, see if it is a CLSID
			IF AxScode <> 0 THEN AxScode = IIDFromString(strProgID, @ClassID)
			' If not a valid ProgID or CLSID return an error
			IF AxScode <> 0 THEN
				FUNCTION = E_INVALIDARG
				EXIT FUNCTION
			END If
			' Get a reference to the IClassFactory2 interface of the control
			' Context: &H17 (%CLSCTX_ALL) =
			' %CLSCTX_INPROC_SERVER OR %CLSCTX_INPROC_HANDLER OR _
			' %CLSCTX_LOCAL_SERVER OR %CLSCTX_REMOTE_SERVER
			AxScode = CoGetClassObject(@ClassID, &H17, null, @IID_IClassFactory2, @ppClassFactory2)
			IF AxScode <> 0 THEN
				FUNCTION = AxScode
				EXIT FUNCTION
			END If
			' Create a licensed instance of the control
			AxScode = IClassFactory2_CreateInstanceLic(ppClassFactory2, NULL, NULL, @IID_IUnknown, strlickey, @ppUnknown)
			DeAllocate(strLicKey)
			' First release the IClassFactory2 interface
			IClassFactory2_Release(ppClassFactory2)
			IF AxScode <> 0 OR ppUnknown = 0 Then
				FUNCTION = AxScode
				EXIT FUNCTION
			END If
			' Ask for the dispatch interface of the control
			AxScode = IUnknown_QueryInterface(ppUnknown, @IID_IDispatch, @ppDispatch)
			' If it fails, use the IUnknown of the control, else use IDispatch
			IF AxScode <> 0 OR ppDispatch = 0 THEN
				ppObj = ppUnknown
			Else
				' Release the IUnknown interface
				IUnknown_Release(ppUnknown)
				ppObj = ppDispatch
			END If
			' Attach the control to the window
			AxScode = AtlAxAttachControl(ppObj, cast(hWnd, hwndcontrol), cast(lpunknown, @ppunkcontainer))
			' Note: Do not release ppObj or your application will GPF when it ends because
			' ATL will release it when the window that hosts the control is destroyed.
			FUNCTION = AxScode
		END Function
	#else
		function atlaxwininit()as scode
			function = AxScode
		end function
	#endif   '#Ifndef Ax_WindowLess


	'only one by project , true if control with ATL , else false
	Function AxInit(ByVal host As Integer = false) As Integer
		AxScode = CoInitialize(null)
		If host Then AxScode = atlaxwininit()
		Function = AxScode
	End Function

	Sub AxStop()                                  'only one by project
		CoUninitialize
	End Sub

   'CLSCTX_INPROC_SERVER   = 1    ' The code that creates and manages objects of this class is a DLL that runs in the same process as the caller of the function specifying the class context.
   'CLSCTX_INPROC_HANDLER  = 2    ' The code that manages objects of this class is an in-process handler.
   'CLSCTX_LOCAL_SERVER    = 4    ' The EXE code that creates and manages objects of this class runs on same machine but is loaded in a separate process space.
   'CLSCTX_REMOTE_SERVER   = 16   ' A remote machine context.
   'CLSCTX_SERVER          = 21   ' CLSCTX_INPROC_SERVER OR CLSCTX_LOCAL_SERVER OR CLSCTX_REMOTE_SERVER
   'CLSCTX_ALL             = 23   ' CLSCTX_INPROC_HANDLER OR CLSCTX_SERVER
   SUB AXCreateObject(BYVAL strProgID AS LPOLESTR, byref ppv as lpvoid, ByVal clsctx As Integer = 21)
      Dim pUnknown AS lpunknown                  ' IUnknown pointer
      dim pDispatch AS lpdispatch                ' IDispatch pointer
      'dim IID_NULL as IID               ' Null GUID
      'dim IID_IUnknown as IID           ' Iunknown GUID
      'Dim IID_IDispatch as IID          ' IDispatch GUID
      dim ClassID AS CLSID                       ' CLSID

      IF *strProgID = "" Then
         AxScode = E_INVALIDARG
         EXIT SUB
      END IF
      ' Standard interface GUIDs
      'IIDFromString("{00000000-0000-0000-0000-000000000000}",@IID_NULL)
      'IIDFromString("{00000000-0000-0000-c000-000000000046}",@IID_IUnknown)
      'IIDFromString("{00020400-0000-0000-c000-000000000046}",@IID_IDispatch)
      ' Exit if strProgID is a null string
      ' Convert the ProgID in a CLSID
      AxScode = CLSIDFromProgID(strProgID, @ClassID)
      ' If it fails, see if it is a CLSID
      IF AxScode <> 0 THEN AxScode = IIDFromString(strProgID, @ClassID)
      ' If not a valid ProgID or CLSID return an error
      IF AxScode <> 0 Then
         AxScode = E_INVALIDARG
         EXIT SUB
      END IF
      ' Create an instance of the object
      AxScode = CoCreateInstance(@ClassID, null, clsctx, @IID_IUnknown, @pUnknown)
      IF AxScode <> 0 OR pUnknown = 0 THEN EXIT Sub

      ' Ask for the dispatch interface
      AxScode = IUnknown_QueryInterface(pUnknown, @IID_IDispatch, @pDispatch)
      ' If it fails, return the Iunknown interface
      IF AxScode <> 0 OR pDispatch = 0 Then
         ppv = pUnknown
         AxScode = S_OK
         EXIT SUB
      END IF
      ' Release the IUnknown interface
      IUnknown_Release(pUnknown)
      ' Return a pointer to the dispatch interface
      ppv = pDispatch
      AxScode = S_OK
   END Sub



   Function AxCreate_Object(str1 as string, ByVal clsctx As Integer = 21) as any ptr
      dim ppv as lpvoid
		dim strProgID AS LPOLESTR= tobstr(str1)
		AXCreateObject strProgID, ppv, clsctx
		function = ppv
		Kill_Str(strProgID)
   end function

   Sub AxRelease_Object(byVal ppUnk as any ptr)
      Dim obj As lpunknown
      if (ppUnk) then
         obj = ppUnk
         (obj) -> lpVtbl -> Release(obj)
         obj = NULL
      end if
      'if ppUnk THEN Ax_Call ppUnk, "Release"
   end sub

   Function AxDllGetClassObject(ByVal hdll As Integer, byval CLSIDS As string, byval IIDS As string, _
            byref pObj as PVOID ptr) as HRESULT
      dim fDllGetClassObject As Function(byval as CLSID ptr, byval as IID ptr, byval as PVOID ptr) as HRESULT
      Dim ClassID As CLSID
      Dim InterfaceID As IID
      Dim picf As iclassfactory Ptr
      Dim punk As lpunknown

      fDllGetClassObject = cast(any ptr, GetProcAddress(cast(hModule, hDll), "DllGetClassObject"))
      CLSIDFromString(clsids, @ClassID)
      IIDFromString(iids, @InterfaceID)
      axscode = fDllGetClassObject(@ClassID, @IID_IClassFactory, @picf)
      If axscode = s_ok Then
         axscode = picf -> lpvtbl -> CreateInstance(picf, NULL, @InterfaceID, @pObj)
         picf -> lpvtbl -> release(picf)
      End If
      Function = axscode
   End Function



   'CONST DISPATCH_METHOD         = 1  ' The member is called using a normal function invocation syntax.
   'CONST DISPATCH_PROPERTYGET    = 2  ' The function is invoked using a normal property-access syntax.
   'CONST DISPATCH_PROPERTYPUT    = 4  ' The function is invoked using a property value assignment syntax.
   'CONST DISPATCH_PROPERTYPUTREF = 8  ' The function is invoked using a property reference assignment syntax.

   #Define IDispatch_GetIDsOfNames(T, i, s, u, l, d)(T) -> lpVtbl -> GetIDsOfNames(T, i, s, u, l, d)
   #define IDispatch_Invoke(T, d, i, l, w, p, v, e, u)(T) -> lpVtbl -> Invoke(T, d, i, l, w, p, v, e, u)

   Sub AxInvoke(BYVAL pthis AS lpdispatch, BYVAL callType AS long, byval vName AS string, _
            byval dispid AS dispid, byval nparams as long, vArgs() AS VARIANT, ByRef vResult AS VARIANT)
      Dim as DISPID dipp = DISPID_PROPERTYPUT
      DIM AS DISPPARAMS udt_DispParams
      Dim pws As WString Ptr
      dim strname as lpcolestr

      ' Check for null pointer
      IF pthis = 0 THEN AXscode = - 1 :EXIT SUB
      If Len(vname) Then
         pws = callocate(len(vName) *len(wstring))
         *pws = WStr(vName)
         strname = pws
         ' Get the DispID
         Axscode = IDispatch_GetIDsOfNames(pthis, @IID_NULL, cast(any ptr, cptr(lpolestr, @strname)), _
               1, LOCALE_USER_DEFAULT, @DispID)
         DeAllocate(strname)
         If Axscode THEN EXIT Sub
      end if
      If nparams Then
         udt_DispParams.cargs = nparams
         udt_DispParams.rgvarg = @vargs(0)
      end If
      IF CallType = 4 OR CallType = 8 THEN
         udt_DispParams.rgdispidNamedArgs = VARPTR(dipp)
         udt_DispParams.cNamedArgs = 1
      END IF
      Axscode = IDispatch_Invoke(pthis, DispID, @IID_NULL, LOCALE_SYSTEM_DEFAULT, _
            CallType, @udt_DispParams, @vresult, @Axpexcepinfo, @AxpuArgErr)
   End Sub

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
   Function str_parse(ByRef source As String, Byref delimiter As String, ByVal idx As Long) As String
      Dim As Long s = 1, c, l
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


   'set obj with pointer
   'sub setObj(byval pxface as UInteger Ptr, ByVal pThis as uinteger)
   sub setObj(byval pxface as any Ptr, ByVal paThis as any ptr)
      dim pthis as uinteger = cuint(paThis)
      If pxface = 0 then exit sub
      Asm
            mov edx , [pxface ]
         othis:
            Xor ecx , ecx
            mov cx , [edx + 6 ]
            Shl ecx , 2
            add edx , ecx
            add edx , 8
            mov eax , [edx ]
            cmp eax , - 1
            jne othis
            mov eax , [pthis ]
            mov [edx + 4 ] , eax
      End Asm
   End Sub

   'set obj with variant [pdispatch]
   sub setVObj(byval pxface as uinteger ptr, byval vvar as variant)
      Dim pthis As LPDISPATCH
      If (vvar.vt = vt_dispatch) and (pxface <> 0) Then
         AxScode = S_OK
         pthis = vvar.pdispval
      Else
         AxScode = E_NoInterface
         pthis = 0
      End If
      Asm
            mov edx , [pxface ]
         vthis:
            Xor ecx , ecx
            mov cx , [edx + 6 ]
            Shl ecx , 2
            add edx , ecx
            add edx , 8
            mov eax , [edx ]
            cmp eax , - 1
            jne vthis
            mov eax , [pthis ]
            mov [edx + 4 ] , eax
      End Asm
   End Sub


   'Function fthis(byval pxface As UInteger) As UInteger
   Function fthis(byval pxface As any ptr) As any ptr
      Asm
            mov edx , [pxface ]
         getpthis:
            Xor ecx , ecx
            mov cx , [edx + 6 ]
            Shl ecx , 2
            add edx , ecx
            add edx , 8
            mov eax , [edx ]
            cmp eax , - 1
            jne getpthis
            mov eax , [edx + 4 ]
            mov [function ] , eax
      End Asm
   End Function

   Sub AxCall cdecl(ByRef pmember as tmember,...)
      Dim vresult as variant
      dim as long i
      dim ARG as any ptr
      dim pv as variant ptr
      dim vargs() as variant
      dim as any ptr pxFace, proc
      Dim pthis As lpdispatch
      DIM AS DISPPARAMS udt_DispParams
      Dim as DISPID dipp = DISPID_PROPERTYPUT

      pxFace = @pmember
      pthis = fthis(pxface)
      If pmember.cargs > 0 then
         ReDim vargs(pmember.cargs - 1) as variant
         ARG = VA_FIRST()
         FOR i = pmember.cargs - 1 to 0 step - 1
            pv = VA_ARG(ARG, any ptr)
            If pv -> vt = vt_empty then
               vargs(i).vt = vt_error
               vargs(i).scode = DISP_E_PARAMNOTFOUND
            else
               vargs(i) = *pv
            End if
            ARG = VA_NEXT(ARG, uinteger)
         NEXT i
      End if
      AXInvoke(pthis, pmember.tkind, "", pMember.dispid, pmember.cargs, vargs(), vresult)
   End sub

   FUNCTION AxGet cdecl(ByRef pmember as tmember,...) as variant
      Dim vresult as variant
      dim as long i
      dim ARG as any ptr
      dim pv as variant ptr
      dim vargs() as variant
      dim as any ptr pxFace, proc
      Dim pthis As lpdispatch
      DIM AS DISPPARAMS udt_DispParams
      Dim as DISPID dipp = DISPID_PROPERTYPUT

      pxFace = @pmember
      pthis = fthis(pxface)
      If pmember.cargs > 0 then
         redim vargs(pmember.cargs - 1) as variant
         ARG = VA_FIRST()
         FOR i = pmember.cargs - 1 to 0 step - 1
            pv = VA_ARG(ARG, variant Ptr)
            If pv -> vt = vt_empty then
               vargs(i).vt = vt_error
               vargs(i).scode = DISP_E_PARAMNOTFOUND
            Else
               vargs(i) = *pv
            End if
            ARG = VA_NEXT(ARG, variant Ptr)
         NEXT i
      End if
      AXInvoke(pthis, pmember.tkind, "", pMember.dispid, pmember.cargs, vargs(), vresult)
      function = vresult
   End Function



   Sub ObjCall Cdecl(pThis As Any Ptr, Script As String,...)
      Dim As String member
      Dim As Integer cargs, cmember, ctype
      Dim ARG as any ptr
      Dim pv as variant Ptr
      Dim as variant vargs(), vresult

      ARG = VA_FIRST()
      cmember = str_numparse(script, ".")
      For i As UShort = 1 To cmember
         member = str_parse(script, ".", i)
         cargs = Val(str_parse(member, "@", 2))
         If cargs Then
            ReDim vargs(cargs - 1) As variant
            For j As Integer = cargs - 1 To 0 Step - 1
               pv = VA_ARG(ARG, variant Ptr)
               If pv -> vt = vt_empty then
                  vargs(j).vt = vt_error
                  vargs(j).scode = DISP_E_PARAMNOTFOUND
               Else
                  vargs(j) = *pv
               End if
               ARG = VA_NEXT(ARG, variant Ptr)
            Next
         Else
            Erase vargs
         End If
         If i <> cmember Then
            If cargs Then ctype = 3 Else ctype = 2
            AXInvoke(pthis, ctype, str_parse(member, "@", 1), 0, cargs, vargs(), vresult)
            pThis = vresult.pdispval
         Else
            ctype = 1
            AXInvoke(pthis, ctype, str_parse(member, "@", 1), 0, cargs, vargs(), vresult)
         End If
      Next
   End Sub


   Sub ObjPut Cdecl(pThis As Any Ptr, Script As String,...)
      Dim As String member
      Dim As Integer cargs, cmember, ctype
      Dim ARG as any ptr
      Dim pv as variant Ptr
      Dim as variant vargs(), vresult

      ARG = VA_FIRST()
      cmember = str_numparse(script, ".")
      For i As UShort = 1 To cmember
         member = str_parse(script, ".", i)
         cargs = Val(str_parse(member, "@", 2))
         If cargs Then
            ReDim vargs(cargs - 1) As variant
            For j As Integer = cargs - 1 To 0 Step - 1
               pv = VA_ARG(ARG, variant Ptr)
               If pv -> vt = vt_empty then
                  vargs(j).vt = vt_error
                  vargs(j).scode = DISP_E_PARAMNOTFOUND
               Else
                  vargs(j) = *pv
               End if
               ARG = VA_NEXT(ARG, variant Ptr)
            Next
         Else
            Erase vargs
         End If
         If i <> cmember Then
            If cargs Then ctype = 3 Else ctype = 2
            AXInvoke(pthis, ctype, str_parse(member, "@", 1), 0, cargs, vargs(), vresult)
            pThis = vresult.pdispval
         Else
            If cargs > 1 Then ctype = 5 Else ctype = 4
            AXInvoke(pthis, ctype, str_parse(member, "@", 1), 0, cargs, vargs(), vresult)
         End If
      Next
   End Sub

   Sub ObjSet Cdecl(pThis As Any Ptr, Script As String,...)
      Dim As String member
      Dim As Integer cargs, cmember, ctype
      Dim ARG as any ptr
      Dim pv as variant Ptr
      Dim as variant vargs(), vresult

      ARG = VA_FIRST()
      cmember = str_numparse(script, ".")
      For i As UShort = 1 To cmember
         member = str_parse(script, ".", i)
         cargs = Val(str_parse(member, "@", 2))
         If cargs Then
            ReDim vargs(cargs - 1)
            For j As Integer = cargs - 1 To 0 Step - 1
               pv = VA_ARG(ARG, variant Ptr)
               If pv -> vt = vt_empty then
                  vargs(j).vt = vt_error
                  vargs(j).scode = DISP_E_PARAMNOTFOUND
               Else
                  vargs(j) = *pv
               End if
               ARG = VA_NEXT(ARG, variant Ptr)
            Next
         Else
            Erase vargs
         End If
         If i <> cmember Then
            If cargs Then ctype = 3 Else ctype = 2
            AXInvoke(pthis, ctype, str_parse(member, "@", 1), 0, cargs, vargs(), vresult)
            pThis = vresult.pdispval
         Else
            If cargs > 1 Then ctype = 9 Else ctype = 8
            AXInvoke(pthis, ctype, str_parse(member, "@", 1), 0, cargs, vargs(), vresult)
         End If
      Next
   End Sub


   function ObjGet Cdecl(pThis As Any Ptr, Script As String,...) As Variant Ptr
      Dim As String member
      Dim As Integer cargs, cmember, ctype
      Dim ARG as any ptr
      Dim pv as variant Ptr
      Dim as variant vargs()
      Static vresult As variant

      ARG = VA_FIRST()
      cmember = str_numparse(script, ".")
      For i As UShort = 1 To cmember
         member = str_parse(script, ".", i)
         cargs = Val(str_parse(member, "@", 2))
         If cargs Then
            ReDim vargs(cargs - 1)
            For j As Integer = cargs - 1 To 0 Step - 1
               pv = VA_ARG(ARG, variant Ptr)
               If pv -> vt = vt_empty then
                  vargs(j).vt = vt_error
                  vargs(j).scode = DISP_E_PARAMNOTFOUND
               Else
                  vargs(j) = *pv
               End if
               ARG = VA_NEXT(ARG, variant Ptr)
            Next
         Else
            Erase vargs
         End If
         If cargs Then ctype = 3 Else ctype = 2
         AXInvoke(pthis, ctype, str_parse(member, "@", 1), 0, cargs, vargs(), vresult)
         If i <> cmember Then pThis = vresult.pdispval Else Return @vresult
      Next
   End Function



   '************************************************************
   ' Basic vTable Call
   ' Syntax: vtCall interface.member,arg1,arg2,arg_n
   ' Note:
   ' BYREF arg -> BYVAL @arg
   ' For Function/Property get -> [retValue] as BYREF last arg
   '************************************************************
   sub vtCall cdecl(ByRef pmember as uinteger,...)
      Dim as any ptr pthis, pxFace, proc
      pxFace = @pmember
      pthis = fthis(pxface)
      If (pthis = 0) then axscode = E_NOINTERFACE :exit Sub
      proc = cast(any ptr, pmember)
      Asm
            mov edx , [pxFace ]
            mov ecx , [edx + 4 ]
            add edx , 8
            mov ebx , 12
            add bx , cx
            Shr ecx , 16
            jecxz cproc
         carg:
            mov eax , [edx ]
            cmp eax , 1
            jne cv2
            Sub ebx , 1
            Xor ax , ax
            mov al , [ebp + ebx ]
            push ax
            jmp cresume
         cv2:
            cmp eax , 2
            jne cv8
            Sub ebx , 2
            mov ax , [ebp + ebx ]
            push ax
            jmp cresume
         cv8:
            cmp eax , 8
            jne cvother
            Sub ebx , 8
            mov eax , [ebp + ebx + 4 ]
            push eax
            mov eax , [ebp + ebx ]
            push eax
            jmp cresume
         cvother:
            Sub ebx , 4
            mov eax , [ebp + ebx ]
            push eax
            jmp cresume
         cresume:
            add edx , 4
            loop carg
         cproc:
            mov eax , [pthis ]
            push eax
            mov eax , [eax ]
            mov edx , [proc ]
            call [eax + edx ]
            mov [axscode ] , eax
      End Asm
   End Sub

   'Exist for backup, vtCall for TLB6
   Sub vtCall2 cdecl(ByRef pmember as uinteger,...)
      Dim as any ptr pthis, pxFace, proc
      pxFace = @pmember
      Asm
            mov edx , [pxface ]
            add edx , 4
         getpthis3:
            Xor eax , eax
            mov ax , [edx + 2 ]
            Shl eax , 2
            add edx , eax
            mov eax , [edx + 4 ]
            add edx , 8
            cmp eax , - 1
            jne getpthis3
            mov eax , [edx ]
            mov [pthis ] , eax
      End Asm
      If (pthis = 0) then axscode = E_NOINTERFACE :exit Sub
      proc = cast(any ptr, pmember)
      Asm
            mov edx , [pxFace ]
            mov ecx , [edx + 4 ]
            add edx , 8
            mov ebx , 12
            add bx , cx
            Shr ecx , 16
            jcxz nextproc
         getarg:
            mov eax , [edx ]
            cmp eax , vt_i2
            je byte2
            cmp eax , vt_ui2
            jne ti8
         byte2:
            Sub ebx , 2
            mov ax , [ebp + ebx ]
            push ax
            jmp resumeloop
         ti8:
            cmp eax , vt_i8
            je byte8
            cmp eax , vt_ui8
            je byte8
            cmp eax , vt_cy
            je byte8
            cmp eax , vt_date
            je byte8
            cmp eax , vt_r8
            jne tother
         byte8:
            Sub ebx , 8
            mov eax , [ebp + ebx + 4 ]
            push eax
            mov eax , [ebp + ebx ]
            push eax
            jmp resumeloop
         tother:
            Sub ebx , 4
            mov eax , [ebp + ebx ]
            push eax
            jmp resumeloop
         resumeloop:
            add edx , 4
            loop getarg
         nextproc:
            mov eax , [pthis ]
            push eax
            mov eax , [eax ]
            mov edx , [proc ]
            call [eax + edx ]
            mov [axscode ] , eax
      End Asm
   end Sub


   'helper function for return value/error code
   Function Scodes(hr As Integer) As String
      function = str(hr)
   End Function




   ' ****************************************************************************************
   ' UI4 AddRef()
   ' Increments the reference counter.
   ' ****************************************************************************************
   FUNCTION Events_AddRef(BYVAL pCookie AS Events_IDispatchVtbl PTR) AS DWORD
      pCookie -> cRef += 1
      FUNCTION = pCookie -> cRef
   END FUNCTION

   ' ****************************************************************************************
   ' UI4 Release()
   ' Releases our class if there is only a reference to him and decrements the reference counter.
   ' ****************************************************************************************
   FUNCTION Events_Release(BYVAL pCookie AS Events_IDispatchVtbl PTR) AS DWORD
      Dim pVtblAddr AS DWORD

      If pCookie -> cRef = 1 THEN
         pVtblAddr = pCookie -> pVtblAddr
         IF HeapFree(GetProcessHeap(), 0, byval cast(lpvoid, pVtblAddr)) THEN
            FUNCTION = 0
            EXIT Function
         ELSE
            FUNCTION = pCookie -> cRef
            EXIT FUNCTION
         END IF
      End IF
      pCookie -> cRef -= 1
      Function = pCookie -> cRef
   End FUNCTION

   ' ****************************************************************************************
   ' HRESULT GetTypeInfoCount([out] *UINT pctinfo)
   ' ****************************************************************************************
   FUNCTION Events_GetTypeInfoCount(BYVAL pCookie AS Events_IDispatchVtbl PTR, BYREF pctInfo AS DWORD) AS LONG
      pctInfo = 0
      Function = S_OK
   END FUNCTION

   ' ****************************************************************************************
   ' HRESULT GetTypeInfo([in] UINT itinfo, [in] UI4 lcid, [out] **VOID pptinfo)
   ' ****************************************************************************************
   FUNCTION Events_GetTypeInfo(BYVAL pCookie AS Events_IDispatchVtbl PTR, _
            BYVAL itinfo AS DWORD, BYVAL lcid AS DWORD, BYREF pptinfo AS DWORD) AS LONG
      FUNCTION = E_NOTIMPL
   END Function

   ' ****************************************************************************************
   ' HRESULT GetTypeInfo([in] UINT itinfo, [in] UI4 lcid, [out] **VOID pptinfo)
   ' ****************************************************************************************
   FUNCTION Events_TypeInfo(BYVAL pCookie AS Events_IDispatchVtbl PTR, _
            BYVAL itinfo AS DWORD, BYVAL lcid AS DWORD, BYREF pptinfo AS DWORD) AS LONG
      FUNCTION = E_NOTIMPL
   END FUNCTION

   ' ****************************************************************************************
   ' HRESULT GetIDsOfNames([in] *GUID riid, [in] **I1 rgszNames, [in] UINT cNames, [in] UI4 lcid, [out] *I4 rgdispid)
   ' ****************************************************************************************
   Function Events_GetIDsOfNames(BYVAL pCookie AS Events_IDispatchVtbl PTR, _
            BYREF riid as IID, BYVAL rgszNames AS DWORD, BYVAL cNames AS DWORD, _
            BYVAL lcid AS DWORD, BYREF rgdispid AS LONG) AS LONG
      FUNCTION = E_NOTIMPL
   End Function

   ' ****************************************************************************************
   ' Builds the IDispatch Virtual Table
   ' ****************************************************************************************
   Function Events_BuildVtbl(BYVAL pthis AS any ptr, byval qryptr As any ptr, ByVal invptr As any ptr) AS DWORD
      'Function Events_BuildVtbl(BYVAL pthis AS DWORD, byval qryptr As dword, ByVal invptr As dword) AS DWORD
      DIM pVtbl AS Events_IDispatchVtbl PTR
      DIM pUnk AS Events_IDispatchVtbl PTR

      pVtbl = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, SIZEOF(*pVtbl))
      IF pVtbl = 0 THEN EXIT FUNCTION
      pVtbl -> QueryInterface = cast(DWORD, QryPtr)
      pVtbl -> AddRef = cast(DWORD, ProcPTR(Events_AddRef))
      pVtbl -> Release = cast(DWORD, ProcPTR(Events_Release))
      pVtbl -> GetTypeInfoCount = cast(DWORD, ProcPTR(Events_GetTypeInfoCount))
      pVtbl -> GetTypeInfo = cast(DWORD, ProcPTR(Events_GetTypeInfo))
      pVtbl -> GetIDsOfNames = cast(DWORD, ProcPTR(Events_GetIDsOfNames))
      pVtbl -> Invoke = cast(DWORD, InvPtr)
      pVtbl -> pVtblAddr = cast(DWORD, pVtbl)
      pVtbl -> pthis = cast(DWORD, pthis)
      pUnk = cast(any ptr, VARPTR(pVtbl -> pVtblAddr))
      FUNCTION = cast(Dword, pUnk)
   End Function

   'convert bstr to string
   'please follow with Ax_FreeStr(bstr) to clean bstr after use to avoid memory leak
   Function FromBSTR(ByVal szW As BSTR) As String
      dim as string sret
      dim as integer L1
      If szW = Null Then Return ""
      L1 = WideCharToMultiByte(CP_ACP, 0, SzW, - 1, Null, 0, Null, Null) - 1
      sret = space(L1)
      WideCharToMultiByte(CP_ACP, 0, SzW, L1, sret, L1, Null, Null)
      Return sret
   End Function


   'convert string to bstr
   'please follow with Ax_FreeStr(bstr) to clean bstr after use to avoid memory leak
   Function ToBSTR(cnv_string As String) As BSTR
      Dim sb As BSTR
      Dim As Integer n = len(cnv_string)
      if n < 1 THEN
         sb = Null
         return sb
      END IF
      n = (MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, cnv_string, - 1, NULL, 0)) - 1
      sb = SysAllocStringLen(sb, n)
      MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, cnv_string, - 1, sb, n)
      Return sb
   End Function

   #Define VariantD VariantV                     ' D for double
   'return numeric double from variant
   Function VariantV(ByRef v As variant) As Double
      Dim vvar As variant
      VariantChangeTypeEx(@vvar, @v, NULL, VARIANT_NOVALUEPROP, VT_R8)
      Return vvar.dblval
   End Function

   'return numeric integer from variant
   Function VariantI(ByRef v As variant) As Integer
      Dim vvar As variant
      VariantChangeTypeEx(@vvar, @v, NULL, VARIANT_NOVALUEPROP, VT_I4)
      Return vvar.lVal
   End Function

   'return numeric uinteger from variant
   Function VariantUI(ByRef v As variant) As Uinteger
      Dim vvar As variant
      VariantChangeTypeEx(@vvar, @v, NULL, VARIANT_NOVALUEPROP, VT_UI4)
      Return vvar.ulVal
   End Function

   'return numeric short from variant
   Function VariantSI(ByRef v As variant) As Short
      Dim vvar As variant
      VariantChangeTypeEx(@vvar, @v, NULL, VARIANT_NOVALUEPROP, VT_I2)
      Return vvar.iVal
   End Function

   'return numeric ushort from variant
   Function VariantUSI(ByRef v As variant) As Ushort
      Dim vvar As variant
      VariantChangeTypeEx(@vvar, @v, NULL, VARIANT_NOVALUEPROP, VT_UI2)
      Return vvar.uiVal
   End Function

   'return numeric longinteger from variant
   Function VariantLI(ByRef v As variant) As longInt
      Dim vvar As variant
      VariantChangeTypeEx(@vvar, @v, NULL, VARIANT_NOVALUEPROP, VT_I8)
      Return vvar.llVal
   End Function

   'return numeric ulonginteger from variant
   Function VariantULI(ByRef v As variant) As UlongInt
      Dim vvar As variant
      VariantChangeTypeEx(@vvar, @v, NULL, VARIANT_NOVALUEPROP, VT_UI8)
      Return vvar.ullVal
   End Function

   'return string value from variant
   Function VariantS(ByRef v As variant) As String
      Dim vvar As variant
      VariantChangeTypeEx(@vvar, @v, NULL, VARIANT_NOVALUEPROP, VT_BSTR)
      'Function = *Cast(wstring Ptr, vvar.bstrval)
      Function = fromBSTR(vvar.bstrval)
   End Function

   'return bstr value from variant
   Function VariantB(ByRef v As variant) As BSTR
      Dim vvar As variant
      VariantChangeTypeEx(@vvar, @v, NULL, VARIANT_NOVALUEPROP, VT_BSTR)
      Function = vvar.bstrval
   End Function


   'assign value to variant ptr
   Declare function vptr OverLoad(x As variant) As Variant Ptr
   Declare Function vptr OverLoad(x As string) As Variant Ptr
   Declare Function vptr OverLoad(x As Longint) As Variant Ptr
   Declare Function vptr OverLoad(x As Ulongint) As Variant Ptr
   Declare Function vptr OverLoad(x As Integer) As Variant Ptr
   Declare Function vptr OverLoad(x As UInteger) As Variant Ptr
   Declare Function vptr OverLoad(x As Short) As Variant Ptr
   Declare Function vptr OverLoad(x As UShort) As Variant Ptr
   Declare Function vptr OverLoad(x As Double) As Variant Ptr
   Declare Function vptr OverLoad(x As Single) As Variant Ptr
   Declare Function vptr OverLoad(x As Byte) As Variant Ptr
   Declare Function vptr OverLoad(x As UByte) As Variant Ptr
   Declare Function vptr OverLoad(x As BSTR) As Variant Ptr
   Declare Function vptr OverLoad(x As any ptr) As Variant Ptr


   Function vptr(x As any ptr) As Variant Ptr
      static v As variant
      v.vt = VT_PTR :v.ulVal = cast(dword, x)
      Return @v
   End Function

   Function vptr(x As variant) As Variant Ptr
      static v As variant
      v.vt = vt_variant :v.pvarval = @x
      Return @v
   End Function

   Function vptr(x As BSTR) As Variant Ptr
      Static v As variant
      v.vt = vt_bstr :v.bstrval = x
      Return @v
   End Function

   Function vptr(x As string) As Variant Ptr
      Static v As variant
      v.vt = vt_bstr :v.bstrval = tobstr(x)
      Return @v
   End Function

   Function vptr(x As LongInt) As Variant Ptr
      Static v As variant
      v.vt = vt_i8 :v.llval = x
      Return @v
   End Function

   Function vptr(x As ULongInt) As Variant Ptr
      Static v As variant
      v.vt = vt_ui8 :v.ullval = x
      Return @v
   End Function

   Function vptr(x As Integer) As Variant Ptr
      Static v As variant
      v.vt = vt_i4 :v.lVal = x
      Return @v
   End Function

   Function vptr(x As UInteger) As Variant Ptr
      Static v As variant
      v.vt = vt_ui4 :v.ulVal = x
      Return @v
   End Function

   Function vptr(x As Short) As Variant Ptr
      Static v As variant
      v.vt = vt_i2 :v.iVal = x
      Return @v
   End Function

   Function vptr(x As UShort) As Variant Ptr
      Static v As variant
      v.vt = vt_ui2 :v.uiVal = x
      Return @v
   End Function

   Function vptr(x As Byte) As Variant Ptr
      Static v As variant
      v.vt = vt_i1 :v.cVal = x
      Return @v
   End Function

   Function vptr(x As UByte) As Variant Ptr
      Static v As variant
      v.vt = vt_ui1 :v.bVal = x
      Return @v
   End Function

   Function vptr(x As double) As Variant Ptr
      Static v As variant
      v.vt = vt_r8 :v.dblval = x
      Return @v
   End Function

   Function vptr(x As single) As Variant Ptr
      Static v As variant
      v.vt = vt_r4 :v.fltVal = x
      Return @v
   End Function


   'assign variant with value
   Declare sub vlet OverLoad(v As variant, x As variant)
   Declare sub vlet OverLoad(v As variant, x As string)
   Declare sub vlet OverLoad(v As variant, x As Longint)
   Declare sub vlet OverLoad(v As variant, x As Ulongint)
   Declare sub vlet OverLoad(v As variant, x As Double)
   Declare sub vlet OverLoad(v As variant, x As single)
   Declare sub vlet OverLoad(v As variant, x As BSTR)
   Declare sub vlet OverLoad(v As variant, x As integer)
   Declare sub vlet OverLoad(v As variant, x As Uinteger)

   sub vlet(v As variant, x As variant)
      v.vt = vt_variant :v.pvarval = @x
   End sub

   sub vlet(v As variant, x As string)
      v.vt = vt_bstr :v.bstrval = tobstr(x)
   End sub

   sub vlet(v As variant, x As BSTR)
      v.vt = vt_bstr :v.bstrval = x
   End sub

   sub vlet(v As variant, x As Integer)
      v.vt = vt_i4 :v.llval = x
   End Sub

   sub vlet(v As variant, x As UInteger)
      v.vt = vt_ui4 :v.ullval = x
   End Sub

   sub vlet(v As variant, x As LongInt)
      v.vt = vt_i8 :v.llval = x
   End Sub

   sub vlet(v As variant, x As ULongInt)
      v.vt = vt_ui8 :v.ullval = x
   End Sub

   sub vlet(v As variant, x As single)
      v.vt = vt_r4 :v.fltval = x
   End Sub

   Sub vlet(v As variant, x As double)
      v.vt = vt_r8 :v.dblval = x
   End Sub

   'assign variant as pointer to a variable
   Declare sub vplet OverLoad(byref v As variant, Byref x As variant)
   Declare sub vplet OverLoad(byref v As variant, Byref x As bstr)
   Declare sub vplet OverLoad(ByRef v As variant, Byref x As byte)
   Declare sub vplet OverLoad(ByRef v As variant, Byref x As ubyte)
   Declare sub vplet OverLoad(ByRef v As variant, Byref x As short)
   Declare sub vplet OverLoad(ByRef v As variant, Byref x As ushort)
   Declare sub vplet OverLoad(ByRef v As variant, Byref x As Integer)
   Declare sub vplet OverLoad(ByRef v As variant, Byref x As Uinteger)
   'Declare sub vplet OverLoad(ByRef v As variant,Byref x As Longint)
   'Declare sub vplet OverLoad(ByRef v As variant,Byref x As Ulongint)
   Declare sub vplet OverLoad(ByRef v As variant, Byref x As single)
   Declare sub vplet OverLoad(ByRef v As variant, Byref x As Double)


   sub vplet(ByRef v As variant, Byref x As variant)
      v.vt = vt_byref Or vt_variant
      v.pvarval = @x
   End sub

   sub vplet(ByRef v As variant, Byref x As bstr)
      v.vt = vt_byref Or vt_bstr :v.pbstrval = @x
   End sub

   sub vplet(ByRef v As variant, ByRef x As Byte)
      v.vt = vt_byref Or vt_i1 :v.pbval = @x
   End Sub

   sub vplet(ByRef v As variant, ByRef x As UByte)
      v.vt = vt_byref Or vt_ui1 :v.pcVal = @x
   End Sub

   sub vplet(ByRef v As variant, ByRef x As Short)
      v.vt = vt_byref Or vt_i2 :v.pival = @x
   End Sub

   sub vplet(ByRef v As variant, ByRef x As UShort)
      v.vt = vt_byref Or vt_ui2 :v.puival = @x
   End Sub

   sub vplet(ByRef v As variant, Byref x As Integer)
      v.vt = vt_byref Or vt_i4 :v.plval = @x
   End Sub

   sub vplet(ByRef v As variant, Byref x As UInteger)
      v.vt = vt_byref Or vt_ui4 :v.pulval = @x
   End Sub
		/'
		sub vplet(ByRef v As variant,Byref x As LongInt)
			v.vt=vt_byref Or vt_i8:v.pllval=@x
		End Sub

		sub vplet(ByRef v As variant,Byref x As ULongInt)
			v.vt=vt_byref Or vt_ui8:v.pullval=@x
		End Sub
		'/
   sub vplet(ByRef v As variant, Byref x As single)
      v.vt = vt_byref Or vt_r4 :v.pfltval = @x
   End Sub

   Sub vplet(ByRef v As variant, Byref x As double)
      v.vt = vt_byref Or vt_r8 :v.pdblval = @x
   End Sub

#endif












