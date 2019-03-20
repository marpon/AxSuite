
#INCLUDE ONCE "windows.bi"
#include once "win/olectl.bi"
#include once "win/ole2.bi"
#include once "win/objbase.bi"

#ifndef UNICODE
	#define UNICODE
	#print ==== info ====> #define UNICODE needed for disphelper :why ? <====
#endif

'#Define Ax_WindowLess			'to use without atl.dll , when no control window
'#Define useATL71  				'to use ATL71.dll  uncomment it,  if commented using  ATL.dll


#ifndef __AX_DISP__
   #define __AX_DISP__

	#inclib "Ax_Disp"

   #print ==== info ====> Compiling with Ax_Disp_lib.bi <====
	#print ====

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


	#define IClassFactory2_CreateInstanceLic(T, u, r, i, s, o)(T) -> lpVtbl -> CreateInstanceLic(T, u, r, i, s, o)
	#define IClassFactory2_GetLicInfo(T, u)(T) -> lpVtbl -> GetLicInfo(T, u)
	#define IClassFactory2_RequestLicKey(T, u, r)(T) -> lpVtbl -> RequestLicKey(T, u, r)
	#define IClassFactory2_Release(T)(T) -> lpVtbl -> Release(T)
	#Define IDispatch_GetIDsOfNames(T, i, s, u, l, d)(T) -> lpVtbl -> GetIDsOfNames(T, i, s, u, l, d)
	#define IDispatch_Invoke(T, d, i, l, w, p, v, e, u)(T) -> lpVtbl -> Invoke(T, d, i, l, w, p, v, e, u)

	#Define VariantD VariantV

	#define Ax_DimObjPtr(objName) dim as any ptr objName = NULL
   #define Ax_FreeStr(bs) SysFreeString(cptr(BSTR, bs))
   #define Kill_Str(bs) dhFreeString(bs) : bs = NULL
   #define Ax_Call dhCallMethod
   #define Ax_Put dhPutValue
   #define Ax_Set dhPutRef
   #define Ax_Get dhGetValue

	#Ifndef Ax_WindowLess
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

		Declare FUNCTION AxWinChild(byVal h_parent as hwnd, name1 as string, progid as string, _
						x as integer, y as integer, w as integer, h as integer, _
						style as integer = WS_visible or WS_child or WS_border, exstyle as integer = 0) as hwnd
		Declare FUNCTION AxWinTool(byVal h_parent as hwnd, name1 as string, progid as string, _
						x as integer, y as integer, w as integer, h as integer, _
						style as integer = WS_visible, exstyle as integer = WS_EX_TOOLWINDOW) as hwnd
		Declare FUNCTION AxWinFull(byVal h_parent as hwnd, name1 as string, progid as string, _
						x as integer, y as integer, w as integer, h as integer, _
						style as integer = WS_visible or WS_OVERLAPPEDWINDOW  , exstyle as integer = 0) as hwnd
		Declare Sub AxWinKill(byVal h_Control as hwnd )
		Declare Sub AxWinHide(byVal h_Control as hwnd, byVal h_Parent as hwnd = 0)
		Declare Sub AxWinShow(byVal h_Control as hwnd, byVal h_Parent as hwnd = 0)
		Declare SUB AtlAxGetDispatch(BYVAL hWndControl AS hwnd, BYREF ppvObj AS lpvoid)
		Declare Function AxCreate_Object overload(BYVAL hWndControl AS hwnd) as any ptr
		Declare FUNCTION AxCreateControlLic(BYVAL strProgID AS LPOLESTR, byval hWndControl AS uinteger, _
						byval strLicKey AS lpwstr) AS LONG


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

	#Else
      function atlaxwininit() as scode
         function = AxScode
      end function
   #endif                                        '#Ifndef Ax_WindowLess

	Declare Function AxInit(ByVal host As Integer = false) As Integer
	Declare Sub AxStop()

	Declare SUB AXCreateObject(BYVAL strProgID AS LPOLESTR, byref ppv as lpvoid, ByVal clsctx As Integer = 21)
	Declare Function AxCreate_Object overload(strProgID AS string, ByVal clsctx As Integer = 21) as any ptr
	Declare Sub AxRelease_Object(byVal ppUnk as any ptr)
	Declare Function AxDllGetClassObject(ByVal hdll As Integer, byval CLSIDS As string, byval IIDS As string, _
					byref pObj as PVOID ptr) as HRESULT
	Declare Sub AxInvoke(BYVAL pthis AS lpdispatch, BYVAL callType AS long, byval vName AS string, _
					byval dispid AS dispid, byval nparams as long, vArgs() AS VARIANT, ByRef vResult AS VARIANT)

	Declare Function FromBSTR(ByVal szW As BSTR) As String
	Declare Function ToBSTR(cnv_string As String) As BSTR
	Declare Function VariantV(ByRef v As variant) As Double
	Declare Function VariantI(ByRef v As variant) As Integer
	Declare Function VariantUI(ByRef v As variant) As Uinteger
	Declare Function VariantSI(ByRef v As variant) As Short
	Declare Function VariantUSI(ByRef v As variant) As Ushort
	Declare Function VariantULI(ByRef v As variant) As UlongInt
	Declare Function VariantS(ByRef v As variant) As String
	Declare Function VariantB(ByRef v As variant) As BSTR
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
	Declare sub setObj(byval pxface as any Ptr, ByVal paThis as any ptr)
	Declare sub setVObj(byval pxface as uinteger ptr, byval vvar as variant)
	Declare Function fthis(byval pxface As any ptr) As any ptr
	Declare Sub AxCall cdecl(ByRef pmember as tmember,...)
	Declare FUNCTION AxGet cdecl(ByRef pmember as tmember,...) as variant
	Declare sub vlet OverLoad(v As variant, x As variant)
	Declare sub vlet OverLoad(v As variant, x As string)
	Declare sub vlet OverLoad(v As variant, x As Longint)
	Declare sub vlet OverLoad(v As variant, x As Ulongint)
	Declare sub vlet OverLoad(v As variant, x As Double)
	Declare sub vlet OverLoad(v As variant, x As single)
	Declare sub vlet OverLoad(v As variant, x As BSTR)
	Declare sub vlet OverLoad(v As variant, x As integer)
	Declare sub vlet OverLoad(v As variant, x As Uinteger)
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
	Declare Function str_numparse(ByRef source as string, ByRef delimiter as string) as long
	Declare Function str_parse(ByRef source As String, Byref delimiter As String, ByVal idx As Long) As String
	Declare Sub ObjCall Cdecl(pThis As Any Ptr, Script As String,...)
	Declare Sub ObjPut Cdecl(pThis As Any Ptr, Script As String,...)
	Declare Sub ObjSet Cdecl(pThis As Any Ptr, Script As String,...)
	Declare function ObjGet Cdecl(pThis As Any Ptr, Script As String,...) As Variant Ptr
	Declare sub vtCall cdecl(ByRef pmember as uinteger,...)
	Declare Sub vtCall2 cdecl(ByRef pmember as uinteger,...)
	Declare Function Scodes(hr As Integer) As String
	Declare FUNCTION Events_AddRef(BYVAL pCookie AS Events_IDispatchVtbl PTR) AS DWORD
	Declare FUNCTION Events_Release(BYVAL pCookie AS Events_IDispatchVtbl PTR) AS DWORD
	Declare FUNCTION Events_GetTypeInfoCount(BYVAL pCookie AS Events_IDispatchVtbl PTR, BYREF pctInfo AS DWORD) AS LONG
	Declare FUNCTION Events_GetTypeInfo(BYVAL pCookie AS Events_IDispatchVtbl PTR, _
					BYVAL itinfo AS DWORD, BYVAL lcid AS DWORD, BYREF pptinfo AS DWORD) AS LONG
	Declare FUNCTION Events_TypeInfo(BYVAL pCookie AS Events_IDispatchVtbl PTR, _
					BYVAL itinfo AS DWORD, BYVAL lcid AS DWORD, BYREF pptinfo AS DWORD) AS LONG
	Declare Function Events_GetIDsOfNames(BYVAL pCookie AS Events_IDispatchVtbl PTR, _
					BYREF riid as IID, BYVAL rgszNames AS DWORD, BYVAL cNames AS DWORD, _
					BYVAL lcid AS DWORD, BYREF rgdispid AS LONG) AS LONG
	Declare Function Events_BuildVtbl(BYVAL pthis AS any ptr, byval qryptr As any ptr, ByVal invptr As any ptr) AS DWORD

	'only one by project , true if control with ATL , else false
   Function AxInit(ByVal host As Integer = false) As Integer
      AxScode = CoInitialize(null)
      If host Then AxScode = atlaxwininit()
      Function = AxScode
   End Function

#endif

#ifndef __disphelper_bi__
	#define __disphelper_bi__

	#define dhInitializeA(bInitializeCOM) dhInitializeImp(bInitializeCOM, FALSE)
	#define dhInitializeW(bInitializeCOM) dhInitializeImp(bInitializeCOM, TRUE)

	#ifdef UNICODE
		#define dhInitialize dhInitializeW
	#else
		#define dhInitialize dhInitializeA
	#endif

	#define AutoWrap dhAutoWrap
	#define DISPATCH_OBJ(objName) dim as IDispatch ptr objName = NULL
	#define dhFreeString(s) SysFreeString(cptr(BSTR,s))

	#ifndef SAFE_RELEASE
		#define SAFE_RELEASE(pObj) if( pObj ) then (pObj)->lpVtbl->Release(pObj): pObj = NULL : end if
	#endif

	#define SAFE_FREE_STRING(s) dhFreeString(s): s = NULL

	#define WITH0(objName, pDisp, szMember) _
		scope : _
			DISPATCH_OBJ(objName) : _
			if( SUCCEEDED(dhGetValue("%o", @objName, pDisp, szMember)) ) then

	#define WITH1(objName, pDisp, szMember, arg1) _
		scope : _
			DISPATCH_OBJ(objName) : _
			if( SUCCEEDED(dhGetValue("%o", @objName, pDisp, szMember, arg1)) ) then

	#define WITH2(objName, pDisp, szMember, arg1, arg2) _
		scope	: _
			DISPATCH_OBJ(objName) : _
			if( SUCCEEDED(dhGetValue("%o", @objName, pDisp, szMember, arg1, arg2)) ) then

	#define WITH3(objName, pDisp, szMember, arg1, arg2, arg3) _
		scope : _
			DISPATCH_OBJ(objName) : _
			if( SUCCEEDED(dhGetValue("%o", @objName, pDisp, szMember, arg1, arg2, arg3)) ) then

	#define WITH4(objName, pDisp, szMember, arg1, arg2, arg3, arg4) _
		scope	: _
			DISPATCH_OBJ(objName) : _
			if( SUCCEEDED(dhGetValue("%o", @objName, pDisp, szMember, arg1, arg2, arg3, arg4)) ) then

	#define ON_WITH_ERROR(objName) _
			else

	#define END_WITH(objName) _
			end if  : _
			SAFE_RELEASE(objName) : _
		end scope


	#define FOR_EACH0(objName, pDisp, szMember) _
		scope  : _
			dim as IEnumVARIANT ptr xx_pEnum_xx = NULL : _
			DISPATCH_OBJ(objName) : _
			if (SUCCEEDED(dhEnumBegin(@xx_pEnum_xx, pDisp, szMember))) then : _
				do while(dhEnumNextObject(xx_pEnum_xx, @objName) = NOERROR)

	#define FOR_EACH1(objName, pDisp, szMember, arg1) _
		scope  : _
			dim as IEnumVARIANT ptr xx_pEnum_xx = NULL  : _
			DISPATCH_OBJ(objName) : _
			if (SUCCEEDED(dhEnumBegin(@xx_pEnum_xx, pDisp, szMember, arg1))) then : _
				do while(dhEnumNextObject(xx_pEnum_xx, @objName) = NOERROR)

	#define FOR_EACH2(objName, pDisp, szMember, arg1, arg2) _
		scope  : _
			dim as IEnumVARIANT ptr xx_pEnum_xx = NULL : _
			DISPATCH_OBJ(objName) : _
			if (SUCCEEDED(dhEnumBegin(@xx_pEnum_xx, pDisp, szMember, arg1, arg2))) then : _
				do while(dhEnumNextObject(xx_pEnum_xx, @objName) = NOERROR)

	#define FOR_EACH3(objName, pDisp, szMember, arg1, arg2, arg3) _
		scope  : _
			dim as IEnumVARIANT ptr xx_pEnum_xx = NULL  : _
			DISPATCH_OBJ(objName) : _
			if (SUCCEEDED(dhEnumBegin(@xx_pEnum_xx, pDisp, szMember, arg1, arg2, arg3))) then : _
				do while(dhEnumNextObject(xx_pEnum_xx, @objName) = NOERROR)

	#define FOR_EACH4(objName, pDisp, szMember, arg1, arg2, arg3, arg4) _
		scope : _
			dim as IEnumVARIANT ptr xx_pEnum_xx = NULL  : _
			DISPATCH_OBJ(objName)  : _
			if (SUCCEEDED(dhEnumBegin(@xx_pEnum_xx, pDisp, szMember, arg1, arg2, arg3, arg4))) then : _
				do while(dhEnumNextObject(xx_pEnum_xx, @objName) = NOERROR)

	#define ON_FOR_EACH_ERROR(objName) _
					SAFE_RELEASE(objName) : _
				loop : _
			else  : _
				do while 0

	#define NEXT_(objName) _
					SAFE_RELEASE(objName) : _
				loop : _
			end if : _
			SAFE_RELEASE(objName) : _
			SAFE_RELEASE(xx_pEnum_xx) : _
		end scope



	type DH_EXCEPTION
		szInitialFunction as LPCWSTR
		szErrorFunction as LPCWSTR
		hr as HRESULT
		szMember(0 to 64-1) as WCHAR
		szCompleteMember(0 to 256-1) as WCHAR
		swCode as UINT
		szDescription as LPWSTR
		szSource as LPWSTR
		szHelpFile as LPWSTR
		dwHelpContext as DWORD
		iArgError as UINT
		bDispatchError as BOOL
	end type

	type PDH_EXCEPTION as DH_EXCEPTION ptr
	type DH_EXCEPTION_CALLBACK as sub cdecl(byval as PDH_EXCEPTION)

	type DH_EXCEPTION_OPTIONS
		hwnd as HWND
		szAppName as LPCWSTR
		bShowExceptions as BOOL
		bDisableRecordExceptions as BOOL
		pfnExceptionCallback as DH_EXCEPTION_CALLBACK
	end type

	type PDH_EXCEPTION_OPTIONS as DH_EXCEPTION_OPTIONS ptr

	extern "c"
		declare function dhCreateObject (byval szProgId as LPCOLESTR, byval szMachine as LPCWSTR, byval ppDisp as IDispatch ptr ptr) as HRESULT
		declare function dhGetObject (byval szFile as LPCOLESTR, byval szProgId as LPCOLESTR, byval ppDisp as IDispatch ptr ptr) as HRESULT
		declare function dhCreateObjectEx (byval szProgId as LPCOLESTR, byval riid as IID ptr, byval dwClsContext as DWORD, byval pServerInfo as COSERVERINFO ptr, byval ppv as any ptr ptr) as HRESULT
		declare function dhGetObjectEx (byval szFile as LPCOLESTR, byval szProgId as LPCOLESTR, byval riid as IID ptr, byval dwClsContext as DWORD, byval lpvReserved as LPVOID, byval ppv as any ptr ptr) as HRESULT
		declare function dhCallMethod (byval pDisp as IDispatch ptr, byval szMember as LPCOLESTR, ...) as HRESULT
		declare function dhPutValue (byval pDisp as IDispatch ptr, byval szMember as LPCOLESTR, ...) as HRESULT
		declare function dhPutRef (byval pDisp as IDispatch ptr, byval szMember as LPCOLESTR, ...) as HRESULT
		declare function dhGetValue (byval szIdentifier as LPCWSTR, byval pResult as any ptr, byval pDisp as IDispatch ptr, byval szMember as LPCOLESTR, ...) as HRESULT
		declare function ax_geti alias "ax_geti" (byval pDisp as IDispatch ptr, byval szMember as LPCOLESTR, ...) as integer
		declare function ax_getui alias "ax_getui" (byval pDisp as IDispatch ptr, byval szMember as LPCOLESTR, ...) as uinteger
		declare function ax_getd alias "ax_getd" (byval pDisp as IDispatch ptr, byval szMember as LPCOLESTR, ...) as double
		declare function ax_gets alias "ax_gets" (byval pDisp as IDispatch ptr, byval szMember as LPCOLESTR, ...) as string
		declare function ax_getvar alias "ax_getvar" (byval pDisp as IDispatch ptr, byval szMember as LPCOLESTR, ...) as VARIANT
		declare function dhInvoke (byval invokeType as integer, byval returnType as VARTYPE, byval pvResult as VARIANT ptr, byval pDisp as IDispatch ptr, byval szMember as LPCOLESTR, ...) as HRESULT
		declare function dhInvokeArray (byval invokeType as integer, byval pvResult as VARIANT ptr, byval cArgs as UINT, byval pDisp as IDispatch ptr, byval szMember as LPCOLESTR, byval pArgs as VARIANT ptr) as HRESULT
		''''''' declare function dhCallMethodV (byval pDisp as IDispatch ptr, byval szMember as LPCOLESTR, byval marker as va_list ptr) as HRESULT
		''''''' declare function dhPutValueV (byval pDisp as IDispatch ptr, byval szMember as LPCOLESTR, byval marker as va_list ptr) as HRESULT
		''''''' declare function dhPutRefV (byval pDisp as IDispatch ptr, byval szMember as LPCOLESTR, byval marker as va_list ptr) as HRESULT
		''''''' declare function dhGetValueV (byval szIdentifier as LPCWSTR, byval pResult as any ptr, byval pDisp as IDispatch ptr, byval szMember as LPCOLESTR, byval marker as va_list ptr) as HRESULT
		''''''' declare function dhInvokeV (byval invokeType as integer, byval returnType as VARTYPE, byval pvResult as VARIANT ptr, byval pDisp as IDispatch ptr, byval szMember as LPCOLESTR, byval marker as va_list ptr) as HRESULT
		declare function dhAutoWrap (byval invokeType as integer, byval pvResult as VARIANT ptr, byval pDisp as IDispatch ptr, byval szMember as LPCOLESTR, byval cArgs as UINT, ...) as HRESULT
		declare function dhParseProperties (byval pDisp as IDispatch ptr, byval szProperties as LPCWSTR, byval lpcPropsSet as UINT ptr) as HRESULT
		declare function dhEnumBegin (byval ppEnum as IEnumVARIANT ptr ptr, byval pDisp as IDispatch ptr, byval szMember as LPCOLESTR, ...) as HRESULT
		''''''' declare function dhEnumBeginV (byval ppEnum as IEnumVARIANT ptr ptr, byval pDisp as IDispatch ptr, byval szMember as LPCOLESTR, byval marker as va_list ptr) as HRESULT
		declare function dhEnumNextObject (byval pEnum as IEnumVARIANT ptr, byval ppDisp as IDispatch ptr ptr) as HRESULT
		declare function dhEnumNextVariant (byval pEnum as IEnumVARIANT ptr, byval pvResult as VARIANT ptr) as HRESULT
		declare function dhInitializeImp (byval bInitializeCOM as BOOL, byval bUnicode as BOOL) as HRESULT
		declare sub dhUninitialize (byval bUninitializeCOM as BOOL)
		declare function dhToggleExceptions (byval bShow as BOOL) as HRESULT
		declare function dhSetExceptionOptions (byval pExceptionOptions as PDH_EXCEPTION_OPTIONS) as HRESULT
		declare function dhGetExceptionOptions (byval pExceptionOptions as PDH_EXCEPTION_OPTIONS) as HRESULT
		declare function dhShowException (byval pException as PDH_EXCEPTION) as HRESULT
		declare function dhGetLastException (byval pException as PDH_EXCEPTION ptr) as HRESULT

		declare function ConvertFileTimeToVariantTime(byval pft as FILETIME ptr, byval pDate as DATE_ ptr) as HRESULT
      declare function ConvertVariantTimeToFileTime(byval date as DATE_, byval pft as FILETIME ptr) as HRESULT
      declare function ConvertVariantTimeToSystemTime(byval date as DATE_, byval pSystemTime as SYSTEMTIME ptr) as HRESULT
      declare function ConvertSystemTimeToVariantTime(byval pSystemTime as SYSTEMTIME ptr, byval pDate as DATE_ ptr) as HRESULT
		#ifdef UNICODE
			declare function dhFormatException alias "dhFormatExceptionW" (byval pException as PDH_EXCEPTION, byval szBuffer as LPWSTR, byval cchBufferSize as UINT, byval bFixedFont as BOOL) as HRESULT
		#else
			declare function dhFormatException alias "dhFormatExceptionA" (byval pException as PDH_EXCEPTION, byval szBuffer as LPSTR, byval cchBufferSize as UINT, byval bFixedFont as BOOL) as HRESULT
		#endif
	end extern

#endif
