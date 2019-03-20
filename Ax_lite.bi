' for AxSuite version 3.2.0.0  date: 27-Mar-2015
'		       .bi file updated "Private Sub/functions", total compatibility but exe smaller

#INCLUDE ONCE "windows.bi"          'if not included earlier

#include once "win/olectl.bi"

'#Define Ax_NoAtl        				'to use Ax_Lite.bi without atl.dll , when no control window ( reduce size of exe)
												' same alternative  #Define Ax_WindowLess
'#Define useATL71             		'to use ATL71.dll  uncomment,  else commented use of ATL.dll
												' not enabled when Ax_NoAtl (or Ax_WindowLess) defined


#ifndef __AX_LITE__
   #define __AX_LITE__
	#ifdef Ax_WindowLess
		#define Ax_NoAtl
   #ENDIF

	#print ====
   #print ==== info ====> Compiling with Ax_Lite.bi  with Private procedures <====
   #print ====


   #define Ax_FreeStr(bs) 				SysFreeString(cptr(BSTR, bs))
   #define Kill_Bstr(bs) 				Ax_FreeStr(bs) : bs = NULL


	#Define toVariant(x)					*vptr(x)
	#Define Vlet(x,y)						x = toVariant(y)					' compatibility axsuite2

	#Define ObjPut             		Ax_Put								' compatibility axsuite2
	#Define ObjCall            		Ax_Call								' compatibility axsuite2
	#Define ObjSet             		Ax_Set								' compatibility axsuite2
	#Define ObjGet        				Ax_Get								' compatibility axsuite2
	#Define Ax_GetStr(a, arg...) 		VariantS(*ax_get(a,arg))
	#Define Ax_GetVal(a, arg...)  	VariantV(*ax_get(a,arg))
	#Define Ax_GetBstr(a, arg...)  	VariantB(*ax_get(a,arg))
	#Define Ax_GetObj(a, arg...)  	Ax_Get(a,arg)->pdispval
	#Define Ax_Vt(a, b, arg...)		a->lpvtbl->b(a ,arg)				' easiest way to adress vtable function with argument
	#Define Ax_Vt0(a, b)					a->lpvtbl->b(a)					' same but without argument
	#Define AxWinNoAtl					AxWinUnreg

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



	dim shared Init_Ax_Var_ AS integer

   dim shared AxScode as scode
   dim shared AxPexcepinfo as excepinfo
   dim shared AxPuArgErr AS uinteger





   Declare Function VariantS(ByRef v As variant) As String
   Declare Function VariantV(ByRef v As variant) As Double
   Declare Function ToBSTR(cnv_string As String) As BSTR
   Declare Function AxCreate_Object overload(strProgID AS string, strIID AS string = "") as any ptr



	Private Function Get_Ax_Stat() as Integer
		Function = Init_Ax_Var_
	End Function

	Private Sub Put_Ax_Stat(tr1 as Integer)
		Init_Ax_Var_ = tr1
	End sub

   #Ifndef Ax_NoAtl
		dim shared Var_ATL_Win_ AS string ' AtlAxWin or  AtlAxWin71
		dim shared as any ptr hLib_ax_atl
		dim shared AtlAxWinInit as function() as integer
		dim shared AtlAxGetControl as function(ByVal hWnd AS hwnd, Byval pp AS UInteger ptr) As uinteger
		dim shared AtlAxAttachControl as function(ByVal pControl As any ptr, _
											ByVal hWnd AS hwnd, ByVal ppUnkContainer AS lpunknown) As UInteger

		Declare Function AxCreate_Object overload(BYVAL hWndControl AS hwnd) as any ptr

		Private Function Get_Atl_Cls() as string
			Function = Var_ATL_Win_
		End Function

		Private Sub Put_Atl_Cls(str1 as string)
			Var_ATL_Win_ = str1
		End sub

		#Ifdef useATL71
			Put_Atl_Cls("AtlAxWin71")
		#Else
			Put_Atl_Cls("AtlAxWin")
		#EndIf

		Private Sub Select_ATL(nver as integer = -1)
			dim as zstring *10 zver
			if hLib_ax_atl = 0 then
				if nver = 0 THEN
					Put_Atl_Cls("AtlAxWin")
				elseif nver = 71 THEN
					Put_Atl_Cls("AtlAxWin71")
				END IF
				zver="Atl.dll"
				if Get_Atl_Cls() = "AtlAxWin71" THEN zver = "Atl71.dll"
				hLib_ax_atl = DylibLoad( zver )
				if hLib_ax_atl = 0 then
					MessageBox(0, zver & "  :    is missing !", "Error, exit Program", MB_ICONERROR)
					end
				end if
				AtlAxWinInit = DylibSymbol(hLib_ax_atl, "AtlAxWinInit")
				AtlAxGetControl = DylibSymbol(hLib_ax_atl, "AtlAxGetControl")
				AtlAxAttachControl = DylibSymbol(hLib_ax_atl, "AtlAxAttachControl")
			end if
      END SUB

		Private sub AtlAxWinStop ()
			UnregisterClass ( Get_Atl_Cls() , GetModuleHandle(byVal 0))
		END SUB

		Private FUNCTION AxWinChild(byVal h_parent as hwnd, name1 as string, progid as string, _
               x as integer, y as integer, w as integer, h as integer, _
               style as integer = WS_visible or WS_child or WS_border, exstyle as integer = 0) as hwnd
         Dim as hwnd h1
         h1 = CreateWindowEx(exstyle, Get_Atl_Cls() , progid, style, x, y, w, h, _
               h_parent, NULL, GetmoduleHandle(0), NULL)
         setwindowtext h1, name1
         function = h1
      END FUNCTION

      Private FUNCTION AxWinTool(byVal h_parent as hwnd, name1 as string, progid as string, _
               x as integer, y as integer, w as integer, h as integer, _
               style as integer = WS_visible, exstyle as integer = WS_EX_TOOLWINDOW) as hwnd
         Dim as hwnd h1
         h1 = CreateWindowEx(exstyle, Get_Atl_Cls() , progid, style, x, y, w, h, _
               h_parent, NULL, GetmoduleHandle(0), NULL)
         setwindowtext h1, name1
         function = h1
      END FUNCTION

      Private FUNCTION AxWinFull(byVal h_parent as hwnd, name1 as string, progid as string, _
               x as integer, y as integer, w as integer, h as integer, _
               style as integer = WS_visible or WS_OVERLAPPEDWINDOW, exstyle as integer = 0) as hwnd
         Dim as hwnd h1
         h1 = CreateWindowEx(exstyle, Get_Atl_Cls() , progid, style, x, y, w, h, _
               h_parent, NULL, GetmoduleHandle(0), NULL)
         setwindowtext h1, name1
         function = h1
      END FUNCTION



      ' ****************************************************************************************
      ' Retrieves the interface of the ActiveX control given the handle of its ATL container
      ' ****************************************************************************************
      Private SUB AtlAxGetDispatch(BYVAL hWndControl AS hwnd, BYREF ppvObj AS lpvoid)
         Dim ppUnk AS lpunknown
         dim ppDispatch as pvoid
         'dim IID_IDispatch as IID

         ' Get the IUnknown of the OCX hosted in the control
         AxScode = AtlAxGetControl(hWndControl, cast(uinteger ptr, @ppUnk))
         IF AxScode <> 0 OR ppUnk = 0 THEN EXIT SUB
         ' Query for the existence of the dispatch interface
         'IIDFromString("{00020400-0000-0000-c000-000000000046}",@IID_IDispatch)
         AxScode = IUnknown_QueryInterface(ppUnk, @IID_IDispatch, @ppDispatch)
         ' If not found, return the IUnknown of the control
         IF AxScode <> 0 OR ppDispatch = 0 THEN
            'print "unknown"
            ppvObj = ppUnk
            EXIT SUB
         END IF
         'print "dispach"
         ' Release the IUnknown of the control
         IUnknown_Release(ppUnk)
         ' Return the retrieved address
         ppvObj = ppDispatch
      End SUB

      Private Function AxCreate_Object(BYVAL hWndControl AS hwnd) as any ptr
         dim ppvObj AS lpvoid
         AtlAxGetDispatch(hWndControl, ppvObj)
         function = ppvObj
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
'      FUNCTION AxCreateControlLic(BYVAL strProgID AS LPOLESTR, byval hWndControl AS uinteger, _
'               byval strLicKey AS lpwstr) AS LONG
		Private FUNCTION AxCreateControlLic(BYVAL strProgID AS LPOLESTR, byval hWndControl AS hwnd, _
               byval strLicKey AS string ) AS LONG
         DIM ppUnknown AS lpunknown              ' IUnknown pointer
         DIM ppDispatch AS lpdispatch            ' IDispatch pointer
         DIM ppObj AS lpvoid                     ' Dispatch interface of the control
         ' IClassFactory2 pointer
         DIM ppClassFactory2 AS IClassFactory2 ptr
         DIM ppUnkContainer AS lpunknown         ' IUnknown of the container
         'DIM IID_NULL as IID               ' Null GUID
         'DIM IID_IUnknown as IID           ' Iunknown GUID
         'DIM IID_IDispatch as IID          ' IDispatch GUID
         'DIM IID_IClassFactory2 as IID     ' IClassFactory2 GUID
         DIM ClassID AS clsid                    ' CLSID
			Dim as Wstring ptr wstrLicKey = tobstr(strLicKey)

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
         AxScode = IClassFactory2_CreateInstanceLic(ppClassFactory2, NULL, NULL, @IID_IUnknown, wstrlickey, @ppUnknown)
         Kill_Bstr(wstrLicKey)
			'DeAllocate(wstrLicKey)
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
			AxScode = AtlAxAttachControl(ppObj,  hwndcontrol, ppunkcontainer)
         'AxScode = AtlAxAttachControl(ppObj, cast(hWnd, hwndcontrol), cast(lpunknown, @ppunkcontainer))
         ' Note: Do not release ppObj or your application will GPF when it ends because
         ' ATL will release it when the window that hosts the control is destroyed.
         FUNCTION = AxScode
      END Function
   #else
      Private function atlaxwininit() as scode
         function = AxScode
      end function

		Private sub AtlAxWinStop()
      end sub

		Private sub Select_ATL()
      end sub
   #endif                                        '#Ifndef Ax_NoAtl


   'only one by project , true if control with ATL , else false
   Private Function AxInit(ByVal host As Integer = false) As Integer
		if Get_Ax_Stat() = 0 then
			AxScode = CoInitialize(null)
			Put_Ax_Stat(1)
			If host Then
				Select_ATL()
				AxScode = atlaxwininit()
				Function = AxScode
				Put_Ax_Stat(2)
			End if
		Elseif Get_Ax_Stat() = 1 then
			If host Then
				Select_ATL()
				AxScode = atlaxwininit()
				Function = AxScode
				Put_Ax_Stat(2)
			End if
		Else
			Function = 0
		End If
   End Function

   Private Sub AxStop()                                  'only one by project
      if Get_Ax_Stat() then CoUninitialize
		if Get_Ax_Stat() = 2 then AtlAxWinStop()
		Put_Ax_Stat(0)
   End Sub

	Private FUNCTION AxWinUnreg(byVal h_parent as hwnd, _
				x as integer, y as integer, w as integer, h as integer, strclass as string = "" , _
				style as integer = WS_visible or WS_child or WS_border, exstyle as integer = 0) as hwnd
		Dim as hwnd h1
		if strclass = "" THEN strclass = "#32770"
		h1 = CreateWindowEx(exstyle, strclass, "Ax_Container", style, x, y, w, h, h_parent, NULL, GetmoduleHandle(0), NULL)
		function = h1
	END FUNCTION

	Private Sub AxWinKill(byVal h_Control as hwnd)
		DestroyWindow(h_Control)
	END SUB

	Private Sub AxWinHide(byVal h_Control as hwnd, byVal h_Parent as hwnd = 0)
		ShowWindow(h_Control, SW_HIDE)
		if h_Parent THEN
			InvalidateRect h_Parent, ByVal 0, True
			UpdateWindow h_Parent
		end if
	END SUB

	Private Sub AxWinShow(byVal h_Control as hwnd, byVal h_Parent as hwnd = 0)
		ShowWindow(h_Control, SW_SHOW)
		InvalidateRect h_Control, ByVal 0, True
		UpdateWindow h_Control
		if h_Parent THEN
			InvalidateRect h_Parent, ByVal 0, True
			UpdateWindow h_Parent
		end if
	END SUB

   'CLSCTX_INPROC_SERVER   = 1    ' The code that creates and manages objects of this class is a DLL that runs in the same process as the caller of the function specifying the class context.
   'CLSCTX_INPROC_HANDLER  = 2    ' The code that manages objects of this class is an in-process handler.
   'CLSCTX_LOCAL_SERVER    = 4    ' The EXE code that creates and manages objects of this class runs on same machine but is loaded in a separate process space.
   'CLSCTX_REMOTE_SERVER   = 16   ' A remote machine context.
   'CLSCTX_SERVER          = 21   ' CLSCTX_INPROC_SERVER OR CLSCTX_LOCAL_SERVER OR CLSCTX_REMOTE_SERVER
   'CLSCTX_ALL             = 23   ' CLSCTX_INPROC_HANDLER OR CLSCTX_SERVER
   Private SUB AXCreateObject(BYVAL strProgID AS LPOLESTR, byref ppv as lpvoid, ByVal clsctx As Integer = 21)
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
         'print "unknown"
         ppv = pUnknown
         AxScode = S_OK
         EXIT SUB
      END IF
      'print "dispatch"
      ' Release the IUnknown interface
      IUnknown_Release(pUnknown)
      ' Return a pointer to the dispatch interface
      ppv = pDispatch
      AxScode = S_OK
   END Sub

   Private Function AxCreate_Object(strProgID1 AS string, strIID1 AS string = "") as any ptr
      Dim pUnknown AS lpunknown                  ' IUnknown pointer
      dim ClassID AS CLSID                       ' CLSID
      dim IID_IUnknown1 as IID
      function = 0
      IF strProgID1 = "" Then
         AxScode = E_INVALIDARG
         EXIT function
      END IF
      ' Exit if strProgID1  null string
      dim strProgID AS LPOLESTR = tobstr(strProgID1)
      if strIID1 = "" THEN
         dim ppv as lpvoid
         AXCreateObject strProgID, ppv, 21
         function = ppv
         Kill_bStr(strProgID)
         exit function
      END IF
      ' Convert the ProgID in a CLSID
      AxScode = CLSIDFromProgID(strProgID, @ClassID)
      ' If it fails, see if it is a CLSID
      IF AxScode <> 0 THEN AxScode = IIDFromString(strProgID, @ClassID)
      ' If not a valid ProgID or CLSID return an error
      dim strIID AS LPOLESTR = tobstr(strIID1)
      AxScode = IIDFromString(strIID, @IID_IUnknown1)
      Kill_bStr(strIID)
      IF AxScode <> 0 Then
         AxScode = E_INVALIDARG
         EXIT function
      END IF
      AxScode = CoCreateInstance(@ClassID, null, 21, @IID_IUnknown1, @pUnknown)
      IF AxScode = S_OK then
         function = pUnknown
      end if
   END function

   Private Sub AxRelease_Object(byVal ppUnk as any ptr)
      Dim obj As lpunknown
      if (ppUnk) then
         obj = ppUnk
         obj -> lpVtbl -> Release(obj)
         obj = NULL
      end if
      'if ppUnk THEN Ax_Call ppUnk, "Release"
   end sub

   Private Function AxDllGetClassObject(ByVal hdll As HMODULE, byval CLSIDS As string, byval IIDS As string, _
            byref pObj as PVOID ptr) as HRESULT
      dim fDllGetClassObject As Function(byval as CLSID ptr, byval as IID ptr, byval as PVOID ptr) as HRESULT
      Dim ClassID As CLSID
      Dim InterfaceID As IID
      Dim picf As iclassfactory Ptr
      Dim punk As lpunknown

      fDllGetClassObject = cast(any ptr, GetProcAddress(hDll, "DllGetClassObject"))
      CLSIDFromString(clsids, @ClassID)
      IIDFromString(iids, @InterfaceID)
      axscode = fDllGetClassObject(@ClassID, @IID_IClassFactory, @picf)
      If axscode = s_ok Then
         axscode = picf -> lpvtbl -> CreateInstance(picf, NULL, @InterfaceID, @pObj)
         picf -> lpvtbl -> release(picf)
      End If
      Function = axscode
   End Function

   'ex dim shared hdll as integer  : hdll=LoadLibrary("NTGraph.ocx")
   'ex: CLSIDS="{C9FE01C2-2746-479B-96AB-E0BE9931B018}"
   'ex: IID_IS="{AC90A107-78E8-4ED8-995A-3AE8BB3044A7}"
   Private function AxCreate_Unreg(ByVal hdll As HMODULE, byval CLSIDS As string, byval IIDS As string, _
            ByVal hWndControl AS hwnd = 0) as any ptr
      DIM ppUnknown AS lpunknown                 ' IUnknown pointer
      DIM ppDispatch AS lpdispatch               ' IDispatch pointer
      DIM ppObj AS lpvoid                        ' Dispatch interface of the control
		DIM UnkContainer AS iunknown            ' IUnknown of the container
      dim pObj as any ptr
		dim zclass as zstring *255,sclass as string
      if AxDllGetClassObject(hdll, CLSIDS, IIDS, pObj) = s_ok Then
         'print "pObj= ";pObj
			ppUnknown = pObj
			'print "hWndControl = ";hWndControl
			' Ask for the dispatch interface of the control
			AxScode = IUnknown_QueryInterface(ppUnknown, @IID_IDispatch, @ppDispatch)
			' If it fails, use the IUnknown of the control, else use IDispatch
			IF AxScode <> 0 OR ppDispatch = 0 THEN
				ppObj = ppUnknown
				'print "ppUnknown"
			Else
				' Release the IUnknown interface
				IUnknown_Release(ppUnknown)
				ppObj = ppDispatch
				'print "ppDispatch"
			END If
			function = ppObj
			'print " ppObj= " ; str(ppObj)
	#Ifndef Ax_NoAtl
			if hWndControl <> 0 and hLib_ax_atl <> 0 THEN
'				GetClassName(hWndControl,zclass,255)
'				sclass = zclass
'				if sclass <> Get_Atl_Cls() THEN exit function
				' Attach the control to the window the control must exist before in a UnkContainer ?
				AxScode = AtlAxAttachControl(ppObj,  hwndcontrol, @UnkContainer)
				'print " @UnkContainer= " ; str(@UnkContainer)
				'AxScode = AtlAxAttachControl(ppObj, hwndcontrol, 0)
			END IF
   #Endif          'Ax_NoAtl
      else
         function = Null
      end if
   END FUNCTION



   'CONST DISPATCH_METHOD         = 1  ' The member is called using a normal function invocation syntax.
   'CONST DISPATCH_PROPERTYGET    = 2  ' The function is invoked using a normal property-access syntax.
   'CONST DISPATCH_PROPERTYPUT    = 4  ' The function is invoked using a property value assignment syntax.
   'CONST DISPATCH_PROPERTYPUTREF = 8  ' The function is invoked using a property reference assignment syntax.

   #Define IDispatch_GetIDsOfNames(T, i, s, u, l, d)(T) -> lpVtbl -> GetIDsOfNames(T, i, s, u, l, d)
   #define IDispatch_Invoke(T, d, i, l, w, p, v, e, u)(T) -> lpVtbl -> Invoke(T, d, i, l, w, p, v, e, u)

   Private Sub AxInvoke(BYVAL pthis AS lpdispatch, BYVAL callType AS long, byval vName AS string, _
            byval dispid AS dispid, byval nparams as long, vArgs() AS VARIANT, ByRef vResult AS VARIANT)
      Dim as DISPID dipp = DISPID_PROPERTYPUT
      DIM AS DISPPARAMS udt_DispParams
      Dim pws As WString Ptr
      dim strname as lpcolestr

      ' Check for null pointer
      IF pthis = 0 THEN AXscode = - 1 :EXIT SUB
      If Len(vname) Then
         'print : print "vname = " & vname : print
         pws = callocate(len(vName) *len(wstring))
         *pws = WStr(vName)
         strname = pws
         ' Get the DispID
         'Axscode = IDispatch_GetIDsOfNames(pthis, @IID_NULL, cptr(ushort ptr ptr, @strname), _
         Axscode = IDispatch_GetIDsOfNames(pthis, @IID_NULL, cast(any ptr, cptr(lpolestr, @strname)), _
               1, LOCALE_USER_DEFAULT, @DispID)
         DeAllocate(strname)
         'print : print "DispID = " & DispID : print
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
      '     Axscode = IDispatch_Invoke(pthis, DispID, @IID_NULL, LOCALE_USER_DEFAULT, _
      '            CallType, @udt_DispParams, @vresult, @Axpexcepinfo, @AxpuArgErr)
   End Sub

   '***************************************************************
   'count number of parse string, separated by delimiter of source
   '***************************************************************
   Private Function str_numparse(ByRef source as string, ByRef delimiter as string) as long
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
   Private Function str_parse(ByRef source As String, Byref delimiter As String, ByVal idx As Long) As String
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

   Private sub free_variant_bstr(byval pv as Variant ptr)
      if pv THEN
         If pv -> vt = vt_bstr then
            SysFreeString(cptr(BSTR, pv -> bstrval))
            pv -> bstrval = Null
         end if
         deallocate(pv)
         pv = Null
      END IF
   END sub



   'set obj with pointer
   'sub setObj(byval pxface as UInteger Ptr, ByVal pThis as uinteger)
   Private sub setObj(byval pxface as any Ptr, ByVal paThis as any ptr)
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
   'sub setVObj(byval pxface as uinteger ptr, byval vvar as variant)
	Private sub setVObj(byval pxface as any ptr, byval vvar as variant)
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
   Private Function fthis(byval pxface As any ptr) As any ptr
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


   'parameters ... as vptr(var)
   Private Sub AxCall cdecl(ByRef pmember as tmember,...)
      Dim vresult as variant
      dim as Integer i
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
            free_variant_bstr(pv)
            ARG = VA_NEXT(ARG, uinteger)
         NEXT i
      End if
      AXInvoke(pthis, pmember.tkind, "", pMember.dispid, pmember.cargs, vargs(), vresult)
   End sub

   'parameters ... as vptr(var)
   Private FUNCTION AxGet cdecl(ByRef pmember as tmember,...) as variant
      Dim vresult as variant
      dim as Integer i
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
            free_variant_bstr(pv)
            ARG = VA_NEXT(ARG, variant Ptr)
         NEXT i
      End if
      AXInvoke(pthis, pmember.tkind, "", pMember.dispid, pmember.cargs, vargs(), vresult)
      function = vresult
   End Function


   'parameters ... as vptr(var)
   Private Sub Ax_Call Cdecl(pThis As Any Ptr, Script As String,...)
      Dim As String member
      Dim As Integer cargs, cmember, ctype
      Dim ARG as any ptr
      Dim pv as variant Ptr
      Dim as variant vargs(), vresult

      ARG = VA_FIRST()
      if left(script, 1) = "." THEN script = mid(script, 2)
      cmember = str_numparse(script, ".")
      For i As Integer = 1 To cmember
         member = str_parse(script, ".", i)
         cargs = Val(str_parse(member, "@", 2))
         if cargs = 0 and instr(member, "@") THEN cargs = 1
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
               free_variant_bstr(pv)
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

   'parameters ... as vptr(var)
   Private Sub Ax_Put Cdecl(pThis As Any Ptr, Script As String,...)
      Dim As String member
      Dim As Integer cargs, cmember, ctype, e
      Dim ARG as any ptr
      Dim pv as variant Ptr
      Dim as variant vargs(), vresult

      ARG = VA_FIRST()
      if left(script, 1) = "." THEN script = mid(script, 2)
      cmember = str_numparse(script, ".")
      For i As Integer = 1 To cmember
         member = str_parse(script, ".", i)
         cargs = Val(str_parse(member, "@", 2))
         if cargs = 0 and instr(member, "@") THEN cargs = 1
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
               free_variant_bstr(pv)
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
            If cargs < 1 Then ctype = 5 Else ctype = 4
            AXInvoke(pthis, ctype, str_parse(member, "@", 1), 0, cargs, vargs(), vresult)
         End If
      Next
   End Sub

   'parameters ... as vptr(var)
   Private Sub Ax_Set Cdecl(pThis As Any Ptr, Script As String,...)
      Dim As String member
      Dim As Integer cargs, cmember, ctype
      Dim ARG as any ptr
      Dim pv as variant Ptr
      Dim as variant vargs(), vresult

      ARG = VA_FIRST()
      if left(script, 1) = "." THEN script = mid(script, 2)
      cmember = str_numparse(script, ".")
      For i As Integer = 1 To cmember
         member = str_parse(script, ".", i)
         cargs = Val(str_parse(member, "@", 2))
         if cargs = 0 and instr(member, "@") THEN cargs = 1
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
               free_variant_bstr(pv)
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
            If cargs < 1 Then ctype = 9 Else ctype = 8
            AXInvoke(pthis, ctype, str_parse(member, "@", 1), 0, cargs, vargs(), vresult)
         End If
      Next
   End Sub


   'parameters ... as vptr(var)
   Private function Ax_Get Cdecl(pThis As Any Ptr, Script As String,...) As Variant ptr
      Dim As String member
      Dim As Integer cargs, cmember, ctype
      Dim ARG as any ptr
      Dim pv as variant Ptr
      Dim as variant vargs()
      Static vresult As variant

      ARG = VA_FIRST()
      if left(script, 1) = "." THEN script = mid(script, 2)
      cmember = str_numparse(script, ".")
      For i As Integer = 1 To cmember
         member = str_parse(script, ".", i)
         cargs = Val(str_parse(member, "@", 2))
         if cargs = 0 and instr(member, "@") THEN cargs = 1
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
               free_variant_bstr(pv)
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


   ' ****************************************************************************************
   ' UI4 AddRef()
   ' Increments the reference counter.
   ' ****************************************************************************************
   Private FUNCTION Events_AddRef(BYVAL pCookie AS Events_IDispatchVtbl PTR) AS DWORD
      pCookie -> cRef += 1
      FUNCTION = pCookie -> cRef
   END FUNCTION

   ' ****************************************************************************************
   ' UI4 Release()
   ' Releases our class if there is only a reference to him and decrements the reference counter.
   ' ****************************************************************************************
   Private FUNCTION Events_Release(BYVAL pCookie AS Events_IDispatchVtbl PTR) AS DWORD
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
   Private FUNCTION Events_GetTypeInfoCount(BYVAL pCookie AS Events_IDispatchVtbl PTR, BYREF pctInfo AS DWORD) AS LONG
      pctInfo = 0
      Function = S_OK
   END FUNCTION

   ' ****************************************************************************************
   ' HRESULT GetTypeInfo([in] UINT itinfo, [in] UI4 lcid, [out] **VOID pptinfo)
   ' ****************************************************************************************
   Private FUNCTION Events_GetTypeInfo(BYVAL pCookie AS Events_IDispatchVtbl PTR, _
            BYVAL itinfo AS DWORD, BYVAL lcid AS DWORD, BYREF pptinfo AS DWORD) AS LONG
      FUNCTION = E_NOTIMPL
   END Function

   ' ****************************************************************************************
   ' HRESULT GetTypeInfo([in] UINT itinfo, [in] UI4 lcid, [out] **VOID pptinfo)
   ' ****************************************************************************************
'   Private FUNCTION Events_TypeInfo(BYVAL pCookie AS Events_IDispatchVtbl PTR, _
'            BYVAL itinfo AS DWORD, BYVAL lcid AS DWORD, BYREF pptinfo AS DWORD) AS LONG
'      FUNCTION = E_NOTIMPL
'   END FUNCTION

   ' ****************************************************************************************
   ' HRESULT GetIDsOfNames([in] *GUID riid, [in] **I1 rgszNames, [in] UINT cNames, [in] UI4 lcid, [out] *I4 rgdispid)
   ' ****************************************************************************************
   Private Function Events_GetIDsOfNames(BYVAL pCookie AS Events_IDispatchVtbl PTR, _
            BYREF riid as IID, BYVAL rgszNames AS DWORD, BYVAL cNames AS DWORD, _
            BYVAL lcid AS DWORD, BYREF rgdispid AS LONG) AS LONG
      FUNCTION = E_NOTIMPL
   End Function

   ' ****************************************************************************************
   ' Builds the IDispatch Virtual Table
   ' ****************************************************************************************
   Private Function Events_BuildVtbl(BYVAL pthis AS any ptr, byval qryptr As any ptr, ByVal invptr As any ptr  ) AS DWORD
      'Function Events_BuildVtbl(BYVAL pthis AS DWORD, byval qryptr As dword, ByVal invptr As dword) AS DWORD
      DIM pVtbl AS Events_IDispatchVtbl PTR
      DIM pUnk AS Events_IDispatchVtbl PTR
      FUNCTION = 0
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

	Private Function stock_ev(ByRef pstr As String = "")As String
		Static dstr As String
		If pstr <> "" Then 'write
			dstr = pstr
		ElseIf pstr<> "0" Then 'clean
			dstr = ""
		End If
		Function = dstr
	End Function

	Private Function Events_IDispatch_QueryInterface (ByVal pCookie As Events_IDispatchVtbl Ptr, ByVal riid As IID Ptr, ByVal ppVObj As PVOID Ptr) As HRESULT
		Dim riids As lpolestr
		Dim As String EV_IID = stock_ev()
		StringFromIID(riid,@riids)
		If (*riids = "{00000000-0000-0000-C000-000000000046}") Or _
			(*riids = "{00020400-0000-0000-C000-000000000046}") Or _
			(*riids = EV_IID) Then
			*ppvObj = pCookie
			Events_AddRef(pCookie)
			Function = S_OK
		Else
			*ppvObj = Null
			Function = E_NOINTERFACE
		End If
	End Function

	Private Function  Ax_Events_cnx (strEvent As String, ByVal pIUnk As Any Ptr, _
										 ByRef ev_wcookie As dword, ByVal pProc As Any Ptr = 0)As Integer
		Dim pUnk As IConnectionPointContainer Ptr = pIUnk
		Dim pCPC As IConnectionPointContainer Ptr
		Dim pCP As IConnectionPoint Ptr
		Dim IID_CPC As IID      ' IID_IConnectionPointContainer
		Dim IID_CP As IID       ' IID_IConnectionPoint
		Dim pSink As lpunknown  ' Pointer to our sink interface
		Dim strdumm As String = stock_ev ( strEvent )
		print "pIUnk = "& str(pIUnk)
		If pUnk = 0 Then Function = E_POINTER : Exit Function
		IIDFromString("{B196B284-BAB4-101A-B69C-00AA00341D07}",@IID_CPC)
		IIDFromString(strEvent,@IID_CP)
		AxScode = IConnectionPointContainer_QueryInterface(pUnk, @IID_CPC, @pCPC)
		print "pCPC = "& str(pCPC)
		If AxScode <> S_OK Then Function = AxScode:Exit Function
		AxScode = IConnectionPointContainer_FindConnectionPoint(pCPC, @IID_CP, @pCP)
		IConnectionPointContainer_release(pCPC)
		If AxScode <> S_OK Then Function = AxScode:Exit Function
		print "pCP = "& str(pCP)
		If pProc Then
			pSink=Cast ( Any Ptr,Events_BuildVTbl ( punk, ProcPtr ( Events_IDispatch_QueryInterface ), pProc ) )
			AxScode = IConnectionPoint_Advise ( pCP, pSink, @ev_wcookie )
			print "ev_wcookie = "& str(ev_wcookie)
		Else
			AxScode = IConnectionPoint_Unadvise ( pCP, ev_wcookie )
		End If
		Function = AxScode
		IConnectionPoint_Release( pCP )
	End Function


	'convert bstr to string
   'please follow with Ax_FreeStr(bstr) to clean bstr after use to avoid memory leak
   Private Function FromBSTR(ByVal szW As BSTR) As String
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
   Private Function ToBSTR(cnv_string As String) As BSTR
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


   'return numeric double from variant
   Private Function VariantV(ByRef v As variant) As Double
      Dim vvar As variant
      VariantChangeTypeEx(@vvar, @v, NULL, VARIANT_NOVALUEPROP, VT_R8)
      function = vvar.dblval
   End Function


   'return string value from variant
   Private Function VariantS(ByRef v As variant) As String
      Dim vvar As variant
      VariantChangeTypeEx(@vvar, @v, NULL, VARIANT_NOVALUEPROP, VT_BSTR)
      '     Function = *Cast(wstring Ptr, vvar.bstrval)
      Function = fromBSTR(vvar.bstrval)
      '     Kill_Bstr(vvar.bstrval) ' free_variant_bstr(@v)
   End Function

   'return bstr value from variant
   Private Function VariantB(ByRef v As variant) As BSTR
      Dim vvar As variant
      VariantChangeTypeEx(@vvar, @v, NULL, VARIANT_NOVALUEPROP, VT_BSTR)
      Function = vvar.bstrval
      Kill_Bstr(vvar.bstrval)                    '  free_variant_bstr(@v)
   End Function



   'assign variant with different types
   Declare  Function vptr OverLoad(x As variant) As Variant Ptr
   Declare  Function vptr OverLoad(x As string) As Variant Ptr
   Declare  Function vptr OverLoad(x As Longint) As Variant Ptr
   Declare  Function vptr OverLoad(x As Ulongint) As Variant Ptr
   Declare  Function vptr OverLoad(x As Integer) As Variant Ptr
   Declare  Function vptr OverLoad(x As UInteger) As Variant Ptr
   Declare  Function vptr OverLoad(x As Short) As Variant Ptr
   Declare  Function vptr OverLoad(x As UShort) As Variant Ptr
   Declare  Function vptr OverLoad(x As Double) As Variant Ptr
   Declare  Function vptr OverLoad(x As Single) As Variant Ptr
   Declare  Function vptr OverLoad(x As Byte) As Variant Ptr
   Declare  Function vptr OverLoad(x As UByte) As Variant Ptr
   Declare  Function vptr OverLoad(x As BSTR) As Variant Ptr


   Private Function vptr(x As variant) As Variant Ptr
      Return @x
   End Function

   Private Function vptr(x As BSTR) As Variant Ptr
      dim pv as Variant ptr = callocate(1, sizeof(Variant))
      pv -> vt = vt_bstr :pv -> bstrval = x
      Return pv
   End Function

   Private Function vptr(x As string) As Variant Ptr
      dim pv as Variant ptr = callocate(1, sizeof(Variant))
      dim bs as bstr = tobstr(x)
      pv -> vt = vt_bstr :pv -> bstrval = bs
      Return pv
   End Function

   Private Function vptr(x As LongInt) As Variant Ptr
      dim pv as Variant ptr = callocate(1, sizeof(Variant))
      pv -> vt = vt_i8 :pv -> llval = x
      Return pv
   End Function

   Private Function vptr(x As ULongInt) As Variant Ptr
      dim pv as Variant ptr = callocate(1, sizeof(Variant))
      pv -> vt = vt_ui8 :pv -> ullval = x
      Return pv
   End Function

   Private Function vptr(x As Integer) As Variant Ptr
      dim pv as Variant ptr = callocate(1, sizeof(Variant))
      pv -> vt = vt_i4 :pv -> lVal = x
      Return pv
   End Function

   Private Function vptr(x As UInteger) As Variant Ptr
      dim pv as Variant ptr = callocate(1, sizeof(Variant))
      pv -> vt = vt_ui4 :pv -> ulVal = x
      Return pv
   End Function

   Private Function vptr(x As Short) As Variant Ptr
      dim pv as Variant ptr = callocate(1, sizeof(Variant))
      pv -> vt = vt_i2 :pv -> iVal = x
      Return pv
   End Function

   Private Function vptr(x As UShort) As Variant Ptr
      dim pv as Variant ptr = callocate(1, sizeof(Variant))
      pv -> vt = vt_ui2 :pv -> uiVal = x
      Return pv
   End Function

   Private Function vptr(x As Byte) As Variant Ptr
      dim pv as Variant ptr = callocate(1, sizeof(Variant))
      pv -> vt = vt_i1 :pv -> cVal = x
      Return pv
   End Function

   Private Function vptr(x As UByte) As Variant Ptr
      dim pv as Variant ptr = callocate(1, sizeof(Variant))
      pv -> vt = vt_ui1 :pv -> bVal = x
      Return pv
   End Function

   Private Function vptr(x As double) As Variant Ptr
      dim pv as Variant ptr = callocate(1, sizeof(Variant))
      pv -> vt = vt_r8 :pv -> dblval = x
      Return pv
   End Function

   Private Function vptr(x As single) As Variant Ptr
      dim pv as Variant ptr = callocate(1, sizeof(Variant))
      pv -> vt = vt_r4 :pv -> fltVal = x
      Return pv
   End Function

	' to register or unregister the ocx dynamicly
	Private Function AxRegMode( filePath As String, Mode As integer=1)as integer
	  Dim hLib1 As HMODULE
	  Dim ProcAd As any ptr
	  Dim sMod As String
	  Dim res As integer
	  Dim hWnd1 As hwnd
	  hWnd1 = CreateWindowEx(0, "reg_form1", "reg_form1", WS_POPUP or WS_SYSMENU, _
									 0, 0, 20, 20, HWND_DESKTOP, NULL, GetmoduleHandle(0), NULL)
	  if mode=0 THEN
		  sMod = "DllUnregisterServer"
	  else
		  sMod = "DllRegisterServer"
	  END IF
	  hLib1 = LoadLibrary(filePath)
	  ProcAd = GetProcAddress(hLib1, sMod)
	  res = CallWindowProc(ProcAd, hWnd1, 0, 0, 0)
	  FreeLibrary hLib1
	  DestroyWindow(hWnd1)
	  Function = cast(integer,res)
	End Function

	'***************************************************************
   'Functions/subs to use vtable Events
   '***************************************************************

	Private Function Ev_Vtbl_Release(Byval pThis As Any ptr) AS Ulong
		Function = 0
	End Function

	Private Function Ev_Vtbl_AddRef(Byval pThis As Any ptr) AS Ulong
		Function = 1
	End Function

	Private sub stock_var(byref pstr as string , byref pvtb as dword, byval rw as integer=0)
		static dstr as string
		static dvtb as dword

		if rw=1 THEN 'write
			dstr = pstr
			dvtb = pvtb
		elseif rw=2 THEN 'read
			pstr = dstr
			pvtb = dvtb
		elseif rw=0 THEN 'clean
			dstr = ""
			dvtb = 0
		END IF
	END sub

	Private Function Ev_Vtbl_Cp(ByVal punk as any ptr, str_ev as string ,byval pSEvent as dword, byval dwconn as dword = 0) as Dword
		dim As IConnectionPointContainer ptr cpc
		dim As IConnectionPoint ptr cp
		Dim As IID IID_Event
		dim As HRESULT hr
		dim as Dword  EvdwCookie
		if punk = 0 THEN
			MessageBox(GetActiveWindow(), "No Object", "Event Error", MB_ICONERROR)
			Function = 0
			exit Function
		END IF
		stock_var (str_ev, pSEvent,1)
		hr = (cast(IUnknown ptr, punk)) -> lpVtbl -> QueryInterface((cast(IUnknown ptr, punk)), @IID_IConnectionPointContainer, cast(LPVOID ptr, @cpc))
		if (FAILED(hr)) then
			MessageBox(GetActiveWindow(), "Container failed", "Event Error", MB_ICONERROR)
			Function = 0
			exit Function
		End if
		IIDFromString(str_ev, @IID_Event)
		hr = cpc -> lpVtbl -> FindConnectionPoint(cpc, @IID_Event, @cp)
		cpc -> lpVtbl -> Release(cpc)
		if (FAILED(hr)) then
			MessageBox(GetActiveWindow(), "Connection point failed", "Event Error", MB_ICONERROR)
			Function = 0
			exit Function
		End if
		if dwconn = 0 THEN
			hr = cp -> lpVtbl -> Advise(cp, cast(IUnknown ptr, pSEvent), @EvdwCookie)
			function = EvdwCookie
      else
			hr = cp -> lpVtbl -> Unadvise(cp, dwconn)
			Function = 1
		END IF
		cp -> lpVtbl -> Release(cp)
		if (FAILED(hr)) then Function = 0
	End Function

	Private FUNCTION Ev_Vtbl_QueryInterface(Byval pThis As Any ptr, Byref riid As GUID, Byref ppvObj As Dword) As HRESULT
		Dim riids As lpolestr
		Dim EventsIIDCP as string, pSEvent as Dword

		StringFromIID(@riid, @riids)
		stock_var(EventsIIDCP , pSEvent,2 )
		If *riids = "{00000000-0000-0000-C000-000000000046}" then  ' @IID_IUnknown
			ppvObj = cast(Dword, pThis)
		elseif *riids = EventsIIDCP then		' interface
			ppvObj = pSEvent
		else
			ppvObj = NULL
			return E_NOINTERFACE
		End if
		return S_OK
	END FUNCTION
#endif
