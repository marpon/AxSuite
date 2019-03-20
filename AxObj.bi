#Include once"win/ocidl.bi"
#Include Once "strwrap.bi"
#Include Once "axhelper.bi"

#Ifdef useATL71
	Declare FUNCTION AtlAxWinInit LIB "ATL71.DLL" ALIAS "AtlAxWinInit" () AS LONG
	Declare FUNCTION AtlAxGetControl LIB "ATL71.DLL" ALIAS "AtlAxGetControl" (BYVAL hWnd AS hwnd, Byval pp AS uint ptr) AS uinteger
	Declare FUNCTION AtlAxAttachControl LIB "ATL71.DLL" ALIAS "AtlAxAttachControl" (BYVAL pControl AS any ptr, BYVAL hWnd AS hwnd, ByVal ppUnkContainer AS lpunknown) AS UInteger
#Else
	Declare FUNCTION AtlAxWinInit LIB "ATL.DLL" ALIAS "AtlAxWinInit" () AS LONG
	Declare FUNCTION AtlAxGetControl LIB "ATL.DLL" ALIAS "AtlAxGetControl" (BYVAL hWnd AS hwnd, Byval pp AS uint ptr) AS uinteger
	Declare FUNCTION AtlAxAttachControl LIB "ATL.DLL" ALIAS "AtlAxAttachControl" (BYVAL pControl AS any ptr, BYVAL hWnd AS hwnd, ByVal ppUnkContainer AS lpunknown) AS UInteger
#EndIf

Type tMember
	DispID As dispid
	cDummy As UINT
	cArgs As UINT
	tKind As UINT
End Type

dim shared AxScode as scode
dim shared AxPexcepinfo as excepinfo
dim shared AxPuArgErr AS uinteger

Function AxInit(ByVal host As Integer=false)As Integer
	AxScode=CoInitialize(null)	
	If host Then AxScode=atlaxwininit
	Function=AxScode
End Function

' ****************************************************************************************
' Retrieves the interface of the ActiveX control given the handle of its ATL container
' ****************************************************************************************
SUB AtlAxGetDispatch (BYVAL hWndControl AS hwnd, BYREF ppvObj AS lpvoid)
	Dim ppUnk AS lpunknown
	dim ppDispatch as pvoid
	'dim IID_IDispatch as IID
	
	' Get the IUnknown of the OCX hosted in the control
	AxScode = AtlAxGetControl(hWndControl, @ppUnk)
	IF AxScode<>0 OR ppUnk=0 THEN EXIT SUB
	' Query for the existence of the dispatch interface
	'IIDFromString("{00020400-0000-0000-c000-000000000046}",@IID_IDispatch)
	AxScode=IUnknown_QueryInterface(ppUnk, @IID_IDispatch, @ppDispatch)
	' If not found, return the IUnknown of the control
	IF AxScode<>0 OR ppDispatch=0 THEN
		ppvObj = ppUnk
		EXIT SUB
	END IF
	' Release the IUnknown of the control
	IUnknown_Release(ppUnk)
	' Return the retrieved address
	ppvObj = ppDispatch
End SUB

'CLSCTX_INPROC_SERVER   = 1    ' The code that creates and manages objects of this class is a DLL that runs in the same process as the caller of the function specifying the class context.
'CLSCTX_INPROC_HANDLER  = 2    ' The code that manages objects of this class is an in-process handler.
'CLSCTX_LOCAL_SERVER    = 4    ' The EXE code that creates and manages objects of this class runs on same machine but is loaded in a separate process space.
'CLSCTX_REMOTE_SERVER   = 16   ' A remote machine context.
'CLSCTX_SERVER          = 21   ' CLSCTX_INPROC_SERVER OR CLSCTX_LOCAL_SERVER OR CLSCTX_REMOTE_SERVER
'CLSCTX_ALL             = 23   ' CLSCTX_INPROC_HANDLER OR CLSCTX_SERVER
SUB AXCreateObject (BYVAL strProgID AS LPOLESTR,byref ppv as lpvoid,ByVal clsctx As Integer=21)
	Dim pUnknown AS lpunknown         ' IUnknown pointer
	dim pDispatch AS lpdispatch       ' IDispatch pointer
	'dim IID_NULL as IID               ' Null GUID
	'dim IID_IUnknown as IID           ' Iunknown GUID
	'Dim IID_IDispatch as IID          ' IDispatch GUID
	dim ClassID AS CLSID         	  ' CLSID
	
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
	AxScode=CLSIDFromProgID(strProgID,@ClassID)
	' If it fails, see if it is a CLSID
	IF AxScode<>0 THEN AxScode=IIDFromString(strProgID,@ClassID)
	' If not a valid ProgID or CLSID return an error
	IF AxScode<>0 Then		
		AxScode = E_INVALIDARG
		EXIT SUB
	END IF
	' Create an instance of the object
	AxScode = CoCreateInstance(@ClassID,null,clsctx, @IID_IUnknown, @pUnknown)
	IF AxScode<>0 OR pUnknown=0 THEN EXIT Sub
	
	' Ask for the dispatch interface
	AxScode = IUnknown_QueryInterface(pUnknown, @IID_IDispatch, @pDispatch)
	' If it fails, return the Iunknown interface
	IF AxScode<>0 OR pDispatch=0 Then		
		ppv = pUnknown
		AxScode = S_OK
		EXIT SUB
	END IF
	' Release the IUnknown interface
	IUnknown_Release(pUnknown)
	' Return a pointer to the dispatch interface
	ppv = pDispatch
	AxScode=S_OK
END Sub

Function AxDllGetClassObject(ByVal hdll As Integer,byval CLSIDS As string,byval IIDS As string,byref pObj as PVOID ptr) as HRESULT
	dim fDllGetClassObject As Function(byval as CLSID ptr, byval as IID ptr, byval as PVOID ptr) as HRESULT
	Dim ClassID As CLSID
	Dim InterfaceID As IID
	Dim picf As iclassfactory Ptr
	Dim punk As lpunknown
	
	fDllGetClassObject=GetProcAddress(hDll,"DllGetClassObject")
	CLSIDFromString(clsids,@ClassID)
	IIDFromString(iids,@InterfaceID)
	axscode=fDllGetClassObject(@ClassID,@IID_IClassFactory,@picf)
	If axscode=s_ok Then	
		axscode=picf->lpvtbl->CreateInstance(picf,NULL,@InterfaceID,@pObj)
		picf->lpvtbl->release(picf)
	EndIf 	
	Function=axscode
End Function 

#define IClassFactory2_CreateInstanceLic(T,u,r,i,s,o) (T)->lpVtbl->CreateInstanceLic(T,u,r,i,s,o)
#define IClassFactory2_GetLicInfo(T,u) (T)->lpVtbl->GetLicInfo(T,u)
#define IClassFactory2_RequestLicKey(T,u,r) (T)->lpVtbl->RequestLicKey(T,u,r)
#define IClassFactory2_Release(T) (T)->lpVtbl->Release(T)
' ****************************************************************************************
' Creates a licensed instance of a visual control (OCX) and attaches it to a window.
' StrProgID can be the ProgID or the ClsID. If you pass a version dependent ProgID or a ClsID,
' it will work only with this particular version.
' hWndControl is the handle of the window and strLicKey the license key.
' ****************************************************************************************
FUNCTION AxCreateControlLic (BYVAL strProgID AS LPOLESTR, byval hWndControl AS uinteger, byval strLicKey AS lpwstr) AS LONG
	DIM ppUnknown AS lpunknown        ' IUnknown pointer
	DIM ppDispatch AS lpdispatch      ' IDispatch pointer
	DIM ppObj AS lpvoid               ' Dispatch interface of the control
	DIM ppClassFactory2 AS IClassFactory2 ptr  ' IClassFactory2 pointer
	DIM ppUnkContainer AS lpunknown   ' IUnknown of the container
	'DIM IID_NULL as IID               ' Null GUID
	'DIM IID_IUnknown as IID           ' Iunknown GUID
	'DIM IID_IDispatch as IID          ' IDispatch GUID
	'DIM IID_IClassFactory2 as IID     ' IClassFactory2 GUID
	DIM ClassID AS clsid              ' CLSID

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
	AxScode=CLSIDFromProgID(strProgID,@ClassID)
	' If it fails, see if it is a CLSID
	IF AxScode<>0 THEN AxScode=IIDFromString(strProgID,@ClassID)
	' If not a valid ProgID or CLSID return an error
	IF AxScode<>0 THEN    	
		FUNCTION = E_INVALIDARG
		EXIT FUNCTION
	END If	
	' Get a reference to the IClassFactory2 interface of the control
	' Context: &H17 (%CLSCTX_ALL) =
	' %CLSCTX_INPROC_SERVER OR %CLSCTX_INPROC_HANDLER OR _
	' %CLSCTX_LOCAL_SERVER OR %CLSCTX_REMOTE_SERVER    	
	AxScode = CoGetClassObject(@ClassID,&H17,null,@IID_IClassFactory2,@ppClassFactory2)
	IF AxScode<>0 THEN
		FUNCTION = AxScode
		EXIT FUNCTION
	END If
	' Create a licensed instance of the control
	AxScode=IClassFactory2_CreateInstanceLic(ppClassFactory2,NULL,NULL,@IID_IUnknown,strlickey,@ppUnknown)
	DeAllocate(strLicKey)
	' First release the IClassFactory2 interface
	IClassFactory2_Release(ppClassFactory2)
	IF AxScode<>0  OR ppUnknown=0 Then
		FUNCTION = AxScode
		EXIT FUNCTION
	END If
	' Ask for the dispatch interface of the control
	AxScode = IUnknown_QueryInterface(ppUnknown, @IID_IDispatch, @ppDispatch)
	' If it fails, use the IUnknown of the control, else use IDispatch
	IF AxScode<>0 OR ppDispatch=0 THEN
		ppObj = ppUnknown
	Else
		' Release the IUnknown interface
		IUnknown_Release(ppUnknown)
		ppObj = ppDispatch
	END If
	' Attach the control to the window
	AxScode=AtlAxAttachControl(ppObj, hwndcontrol, @ppunkcontainer) 
	' Note: Do not release ppObj or your application will GPF when it ends because
	' ATL will release it when the window that hosts the control is destroyed.
	FUNCTION = AxScode
END Function

'CONST DISPATCH_METHOD         = 1  ' The member is called using a normal function invocation syntax.
'CONST DISPATCH_PROPERTYGET    = 2  ' The function is invoked using a normal property-access syntax.
'CONST DISPATCH_PROPERTYPUT    = 4  ' The function is invoked using a property value assignment syntax.
'CONST DISPATCH_PROPERTYPUTREF = 8  ' The function is invoked using a property reference assignment syntax.
#Define IDispatch_GetIDsOfNames(T,i,s,u,l,d) (T)->lpVtbl->GetIDsOfNames(T,i,s,u,l,d)
#define IDispatch_Invoke(T,d,i,l,w,p,v,e,u) (T)->lpVtbl->Invoke(T,d,i,l,w,p,v,e,u)
Sub AxInvoke(BYVAL pthis AS lpdispatch,BYVAL callType AS long,byval vName AS string,byval dispid AS dispid,byval nparams as long,vArgs() AS VARIANT,ByRef vResult AS VARIANT)   
	Dim as DISPID dipp=DISPID_PROPERTYPUT 
	DIM AS DISPPARAMS udt_DispParams
	Dim pws As WString Ptr
	dim strname as lpcolestr
	
	' Check for null pointer
	IF pthis = 0 THEN AXscode = -1 : EXIT SUB	
	If Len(vname) Then
		pws=callocate(len(vName)*len(wstring))
		*pws=WStr(vName)
		strname=pws
		' Get the DispID  	 
		Axscode = IDispatch_GetIDsOfNames(pthis, @IID_NULL,cptr(lpolestr,@strname),1,LOCALE_USER_DEFAULT, @DispID)
		DeAllocate(strname)
		If Axscode THEN EXIT Sub
	endif	
	If nparams Then
		udt_DispParams.cargs=nparams
		udt_DispParams.rgvarg=@vargs(0)
	end If
	IF CallType = 4 OR CallType = 8 THEN
		udt_DispParams.rgdispidNamedArgs= VARPTR(dipp)
		udt_DispParams.cNamedArgs=1
	END IF 
	Axscode = IDispatch_Invoke(pthis, DispID, @IID_NULL, LOCALE_SYSTEM_DEFAULT, CallType, @udt_DispParams, @vresult, @Axpexcepinfo, @AxpuArgErr)
End Sub

Function fthis(byval pxface As UInteger)As UInteger
	Asm		
		mov edx,[pxface]
	getpthis:
		Xor ecx,ecx
		mov cx,[edx+6]
		Shl ecx,2
		add edx,ecx
		add edx,8
		mov eax,[edx]
		cmp eax,-1			
		jne getpthis
		mov eax,[edx+4]
		mov [function],eax
	End Asm
End Function

Sub AxCall cdecl(ByRef pmember as tmember,...)
	Dim vresult as variant 
	dim as long i 
	dim ARG as any ptr
	dim pv as variant ptr
	dim vargs()as variant
	dim as uinteger ptr pxFace,proc
	Dim pthis As lpdispatch
	DIM AS DISPPARAMS udt_DispParams
	Dim as DISPID dipp=DISPID_PROPERTYPUT
	
	pxFace=@pmember
	pthis=fthis(pxface)
	If pmember.cargs>0 then
		ReDim vargs(pmember.cargs-1)as variant
		ARG = VA_FIRST()
		FOR i = pmember.cargs-1 to 0 step -1
			pv=VA_ARG(ARG,uinteger)
			If pv->vt=vt_empty then
				vargs(i).vt=vt_error
				vargs(i).scode=DISP_E_PARAMNOTFOUND
			else
				vargs(i)=*pv
			End if   		
			ARG = VA_NEXT(ARG,uinteger)
		NEXT i
	End if
	AXInvoke(pthis,pmember.tkind,"",pMember.dispid,pmember.cargs,vargs(),vresult)
End sub

FUNCTION AxGet cdecl(ByRef pmember as tmember,...)as variant
	Dim vresult as variant
	dim as long i 
	dim ARG as any ptr
	dim pv as variant ptr
	dim vargs()as variant
	dim as uinteger ptr pxFace,proc
	Dim pthis As lpdispatch
	DIM AS DISPPARAMS udt_DispParams
	Dim as DISPID dipp=DISPID_PROPERTYPUT
	
	pxFace=@pmember
	pthis=fthis(pxface)	
	If pmember.cargs>0 then
		redim vargs(pmember.cargs-1)as variant
		ARG = VA_FIRST()
		FOR i = pmember.cargs-1 to 0 step -1
			pv=VA_ARG(ARG,variant Ptr)
			If pv->vt=vt_empty then
				vargs(i).vt=vt_error
				vargs(i).scode=DISP_E_PARAMNOTFOUND
			Else
				vargs(i)=*pv
			End if   		
			ARG = VA_NEXT(ARG,variant Ptr)
		NEXT i
	End if
	AXInvoke(pthis,pmember.tkind,"",pMember.dispid,pmember.cargs,vargs(),vresult)
	function=vresult        
End Function

Sub ObjCall Cdecl(pThis As Any Ptr,Script As String,...)
	Dim As String member
	Dim As Integer cargs,cmember,ctype
	Dim ARG as any ptr
	Dim pv as variant Ptr	
	Dim as variant vargs(),vresult
	
	ARG = VA_FIRST()
	cmember=str_numparse(script,".")
	For i As UShort=1 To cmember
		member=str_parse(script,".",i)
		cargs=Val(str_parse(member,"@",2))		
		If cargs Then
			ReDim vargs(cargs-1)As variant
			For j As Integer=cargs-1 To 0 Step -1
				pv=VA_ARG(ARG,variant Ptr)
				If pv->vt=vt_empty then
					vargs(j).vt=vt_error
					vargs(j).scode=DISP_E_PARAMNOTFOUND
				Else
					vargs(j)=*pv
				End if   		
				ARG = VA_NEXT(ARG,variant Ptr)
			Next
		Else
			Erase vargs
		EndIf
		If i<>cmember Then
			If cargs Then ctype=3 Else ctype=2
			AXInvoke(pthis,ctype,str_parse(member,"@",1),0,cargs,vargs(),vresult)
			pThis=vresult.pdispval
		Else
			ctype=1
			AXInvoke(pthis,ctype,str_parse(member,"@",1),0,cargs,vargs(),vresult)
		EndIf
	Next
End Sub

Sub ObjPut Cdecl(pThis As Any Ptr,Script As String,...)
	Dim As String member
	Dim As Integer cargs,cmember,ctype
	Dim ARG as any ptr
	Dim pv as variant Ptr	
	Dim as variant vargs(),vresult
	
	ARG = VA_FIRST()
	cmember=str_numparse(script,".")
	For i As UShort=1 To cmember
		member=str_parse(script,".",i)
		cargs=Val(str_parse(member,"@",2))
		If cargs Then
			ReDim vargs(cargs-1)As variant
			For j As Integer=cargs-1 To 0 Step -1
				pv=VA_ARG(ARG,variant Ptr)
				If pv->vt=vt_empty then
					vargs(j).vt=vt_error
					vargs(j).scode=DISP_E_PARAMNOTFOUND
				Else
					vargs(j)=*pv
				End if   		
				ARG = VA_NEXT(ARG,variant Ptr)
			Next
		Else
			Erase vargs
		EndIf
		If i<>cmember Then
			If cargs Then ctype=3 Else ctype=2
			AXInvoke(pthis,ctype,str_parse(member,"@",1),0,cargs,vargs(),vresult)
			pThis=vresult.pdispval
		Else
			If cargs>1 Then ctype=5 Else ctype=4
			AXInvoke(pthis,ctype,str_parse(member,"@",1),0,cargs,vargs(),vresult)
		EndIf
	Next
End Sub

Sub ObjSet Cdecl(pThis As Any Ptr,Script As String,...)
	Dim As String member
	Dim As Integer cargs,cmember,ctype
	Dim ARG as any ptr
	Dim pv as variant Ptr	
	Dim as variant vargs(),vresult
	
	ARG = VA_FIRST()
	cmember=str_numparse(script,".")
	For i As UShort=1 To cmember
		member=str_parse(script,".",i)
		cargs=Val(str_parse(member,"@",2))
		If cargs Then
			ReDim vargs(cargs-1)
			For j As Integer=cargs-1 To 0 Step -1
				pv=VA_ARG(ARG,variant Ptr)
				If pv->vt=vt_empty then
					vargs(j).vt=vt_error
					vargs(j).scode=DISP_E_PARAMNOTFOUND
				Else
					vargs(j)=*pv
				End if   		
				ARG = VA_NEXT(ARG,variant Ptr)
			Next
		Else
			Erase vargs
		EndIf
		If i<>cmember Then
			If cargs Then ctype=3 Else ctype=2
			AXInvoke(pthis,ctype,str_parse(member,"@",1),0,cargs,vargs(),vresult)
			pThis=vresult.pdispval
		Else
			If cargs>1 Then ctype=9 Else ctype=8
			AXInvoke(pthis,ctype,str_parse(member,"@",1),0,cargs,vargs(),vresult)
		EndIf
	Next
End Sub

function ObjGet Cdecl(pThis As Any Ptr,Script As String,...)As Variant Ptr
	Dim As String member
	Dim As Integer cargs,cmember,ctype
	Dim ARG as any ptr
	Dim pv as variant Ptr	
	Dim as variant vargs()
	Static vresult As variant
	
	ARG = VA_FIRST()
	cmember=str_numparse(script,".")
	For i As UShort=1 To cmember
		member=str_parse(script,".",i)
		cargs=Val(str_parse(member,"@",2))
		If cargs Then
			ReDim vargs(cargs-1)
			For j As Integer=cargs-1 To 0 Step -1
				pv=VA_ARG(ARG,variant Ptr)
				If pv->vt=vt_empty then
					vargs(j).vt=vt_error
					vargs(j).scode=DISP_E_PARAMNOTFOUND
				Else
					vargs(j)=*pv
				End if   		
				ARG = VA_NEXT(ARG,variant Ptr)
			Next
		Else
			Erase vargs
		EndIf
		If cargs Then ctype=3 Else ctype=2
		AXInvoke(pthis,ctype,str_parse(member,"@",1),0,cargs,vargs(),vresult)
		If i<>cmember Then pThis=vresult.pdispval Else Return @vresult
	Next
End Function

'set obj with pointer
sub setObj(byval pxface as UInteger Ptr,ByVal pThis as uinteger)
	If pxface=0 then exit sub
	Asm		
		mov edx,[pxface]
	othis:
		Xor ecx,ecx
		mov cx,[edx+6]
		Shl ecx,2			
		add edx,ecx
		add edx,8
		mov eax,[edx]
		cmp eax,-1			
		jne othis		
		mov eax,[pthis]
		mov [edx+4],eax
	End Asm	
End Sub

'set obj with variant [pdispatch]
sub setVObj(byval pxface as uinteger ptr,byval vvar as variant)
	Dim pthis As uint
	If (vvar.vt=vt_dispatch) and (pxface<>0) Then 
		AxScode=S_OK
		pthis=vvar.pdispval
	Else
		AxScode=E_NoInterface
		pthis=0
	EndIf
	Asm		
		mov edx,[pxface]
	vthis:
		Xor ecx,ecx
		mov cx,[edx+6]
		Shl ecx,2			
		add edx,ecx
		add edx,8
		mov eax,[edx]
		cmp eax,-1			
		jne vthis		
		mov eax,[pthis]
		mov [edx+4],eax
	End Asm	
End Sub

'************************************************************
' Basic vTable Call
' Syntax: vtCall interface.member,arg1,arg2,arg_n
' Note: 
' BYREF arg -> BYVAL @arg
' For Function/Property get -> [retValue] as BYREF last arg
'************************************************************
sub vtCall cdecl(ByRef pmember as uinteger,...)
	Dim as long ptr pthis,pxFace,proc
	pxFace=@pmember
	pthis=fthis(pxface)
	If (pthis=0)then axscode=E_NOINTERFACE:exit Sub
	proc=pmember
	Asm		
		mov edx,[pxFace]
		mov ecx,[edx+4]		
		add edx,8
		mov ebx,12
		add bx,cx
		Shr ecx,16
		jecxz cproc
	carg:
		mov eax,[edx]
		cmp eax,1
		jne cv2
		Sub ebx,1
		Xor ax,ax
		mov al,[ebp+ebx]
		push ax
		jmp cresume
	cv2:
		cmp eax,2
		jne cv8				
		Sub ebx,2		
		mov ax,[ebp+ebx]
		push ax
		jmp cresume		
	cv8:	
		cmp eax,8
		jne cvother				
		Sub ebx,8			
		mov eax,[ebp+ebx+4]
		push eax
		mov eax,[ebp+ebx]
		push eax		
		jmp cresume
	cvother:
		Sub ebx,4
		mov eax,[ebp+ebx]
		push eax
		jmp cresume	
	cresume:
		add edx,4
		loop carg	
	cproc:	
		mov eax,[pthis]
		push eax
		mov eax,[eax]
		mov edx,[proc]
		call [eax+edx]
		mov [axscode],eax
	End Asm
End Sub

'Exist for backup, vtCall for TLB6  
Sub vtCall2 cdecl(ByRef pmember as uinteger,...)
	Dim as uinteger ptr pthis,pxFace,proc
	pxFace=@pmember
	Asm
		mov edx,[pxface]
		add edx,4
	getpthis3:
		Xor eax,eax		
		mov ax,[edx+2]
		Shl eax,2
		add edx,eax
		mov eax,[edx+4]
		add edx,8		
		cmp eax,-1
		jne getpthis3
		mov eax,[edx]
		mov [pthis],eax
	End Asm
	If (pthis=0)then axscode=E_NOINTERFACE:exit Sub
	proc=pmember
	Asm		
		mov edx,[pxFace]
		mov ecx,[edx+4]		
		add edx,8
		mov ebx,12
		add bx,cx
		Shr ecx,16
		jcxz nextproc
	getarg:
		mov eax,[edx]
		cmp eax,vt_i2
		je byte2
		cmp eax,vt_ui2
		jne ti8				
	byte2:
		Sub ebx,2		
		mov ax,[ebp+ebx]
		push ax
		jmp resumeloop			
	ti8:	
		cmp eax,vt_i8
		je byte8
		cmp eax,vt_ui8
		je byte8
		cmp eax,vt_cy
		je byte8
		cmp eax,vt_date
		je byte8				
		cmp eax,vt_r8
		jne tother				
	byte8:
		Sub ebx,8			
		mov eax,[ebp+ebx+4]
		push eax
		mov eax,[ebp+ebx]
		push eax
		jmp resumeloop
	tother:
		Sub ebx,4
		mov eax,[ebp+ebx]
		push eax
		jmp resumeloop	
	resumeloop:
		add edx,4
		loop getarg	
	nextproc:	
		mov eax,[pthis]
		push eax
		mov eax,[eax]
		mov edx,[proc]
		call [eax+edx]
		mov [axscode],eax
	End Asm
end Sub

'helper function for return value/error code
Function Scodes(hr As Integer) As String
	Select Case hr
		case -2147418113
		 return "E_UNEXPECTED "
		case -2147467263
		 return "E_NOTIMPL "
		case -2147024882
		 return "E_OUTOFMEMORY "
		case -2147024809
		 return "E_INVALIDARG "
		case -2147467262
		 return "E_NOINTERFACE "
		case -2147467261
		 return "E_POINTER "
		case -2147024890
		 return "E_HANDLE "
		case -2147467260
		 return "E_ABORT "
		case -2147467259
		 return "E_FAIL "
		case -2147024891
		 return "E_ACCESSDENIED "
		case -2147483638
		 return "E_PENDING "
		case -2147467258
		 return "CO_E_INIT_TLS "
		case -2147467257
		 return "CO_E_INIT_SHARED_ALLOCATOR "
		case -2147467256
		 return "CO_E_INIT_MEMORY_ALLOCATOR "
		case -2147467255
		 return "CO_E_INIT_CLASS_CACHE "
		case -2147467254
		 return "CO_E_INIT_RPC_CHANNEL "
		case -2147467253
		 return "CO_E_INIT_TLS_SET_CHANNEL_CONTROL "
		case -2147467252
		 return "CO_E_INIT_TLS_CHANNEL_CONTROL "
		case -2147467251
		 return "CO_E_INIT_UNACCEPTED_USER_ALLOCATOR "
		case -2147467250
		 return "CO_E_INIT_SCM_MUTEX_EXISTS "
		case -2147467249
		 return "CO_E_INIT_SCM_FILE_MAPPING_EXISTS "
		case -2147467248
		 return "CO_E_INIT_SCM_MAP_VIEW_OF_FILE "
		case -2147467247
		 return "CO_E_INIT_SCM_EXEC_FAILURE "
		case -2147467246
		 return "CO_E_INIT_ONLY_SINGLE_THREADED "
		case -2147467245
		 return "CO_E_CANT_REMOTE "
		case -2147467244
		 return "CO_E_BAD_SERVER_NAME "
		case -2147467243
		 return "CO_E_WRONG_SERVER_IDENTITY "
		case -2147467242
		 return "CO_E_OLE1DDE_DISABLED "
		case -2147467241
		 return "CO_E_RUNAS_SYNTAX "
		case -2147467240
		 return "CO_E_CREATEPROCESS_FAILURE "
		case -2147467239
		 return "CO_E_RUNAS_CREATEPROCESS_FAILURE "
		case -2147467238
		 return "CO_E_RUNAS_LOGON_FAILURE "
		case -2147467237
		 return "CO_E_LAUNCH_PERMSSION_DENIED "
		case -2147467236
		 return "CO_E_START_SERVICE_FAILURE "
		case -2147467235
		 return "CO_E_REMOTE_COMMUNICATION_FAILURE "
		case -2147467234
		 return "CO_E_SERVER_START_TIMEOUT "
		case -2147467233
		 return "CO_E_CLSREG_INCONSISTENT "
		case -2147467232
		 return "CO_E_IIDREG_INCONSISTENT "
		case -2147467231
		 return "CO_E_NOT_SUPPORTED "
		case -2147467230
		 return "CO_E_RELOAD_DLL "
		case -2147467229
		 return "CO_E_MSI_ERROR "
		case -2147467228
		 return "CO_E_ATTEMPT_TO_CREATE_OUTSIDE_CLIENT_CONTEXT "
		case -2147467227
		 return "CO_E_SERVER_PAUSED "
		case -2147467226
		 return "CO_E_SERVER_NOT_PAUSED "
		case -2147467225
		 return "CO_E_CLASS_DISABLED "
		case -2147467224
		 return "CO_E_CLRNOTAVAILABLE "
		case -2147467223
		 return "CO_E_ASYNC_WORK_REJECTED "
		case -2147467222
		 return "CO_E_SERVER_INIT_TIMEOUT "
		case -2147467221
		 return "CO_E_NO_SECCTX_IN_ACTIVATE "
		case -2147467216
		 return "CO_E_TRACKER_CONFIG "
		case -2147467215
		 return "CO_E_THREADPOOL_CONFIG "
		case -2147467214
		 return "CO_E_SXS_CONFIG "
		case -2147467213
		 return "CO_E_MALFORMED_SPN "
		case 0
		 return "S_OK "
		case 1
		 return "S_FALSE "
		case -2147221504
		 return "OLE_E_FIRST "
		case -2147221249
		 return "OLE_E_LAST "
		case 262144
		 return "OLE_S_FIRST "
		case 262399
		 return "OLE_S_LAST "
		case -2147221504
		 return "OLE_E_OLEVERB "
		case -2147221503
		 return "OLE_E_ADVF "
		case -2147221502
		 return "OLE_E_ENUM_NOMORE "
		case -2147221501
		 return "OLE_E_ADVISENOTSUPPORTED "
		case -2147221500
		 return "OLE_E_NOCONNECTION "
		case -2147221499
		 return "OLE_E_NOTRUNNING "
		case -2147221498
		 return "OLE_E_NOCACHE "
		case -2147221497
		 return "OLE_E_BLANK "
		case -2147221496
		 return "OLE_E_CLASSDIFF "
		case -2147221495
		 return "OLE_E_CANT_GETMONIKER "
		case -2147221494
		 return "OLE_E_CANT_BINDTOSOURCE "
		case -2147221493
		 return "OLE_E_STATIC "
		case -2147221492
		 return "OLE_E_PROMPTSAVECANCELLED "
		case -2147221491
		 return "OLE_E_INVALIDRECT "
		case -2147221490
		 return "OLE_E_WRONGCOMPOBJ "
		case -2147221489
		 return "OLE_E_INVALIDHWND "
		case -2147221488
		 return "OLE_E_NOT_INPLACEACTIVE "
		case -2147221487
		 return "OLE_E_CANTCONVERT "
		case -2147221486
		 return "OLE_E_NOSTORAGE "
		case -2147221404
		 return "DV_E_FORMATETC "
		case -2147221403
		 return "DV_E_DVTARGETDEVICE "
		case -2147221402
		 return "DV_E_STGMEDIUM "
		case -2147221401
		 return "DV_E_STATDATA "
		case -2147221400
		 return "DV_E_LINDEX "
		case -2147221399
		 return "DV_E_TYMED "
		case -2147221398
		 return "DV_E_CLIPFORMAT "
		case -2147221397
		 return "DV_E_DVASPECT "
		case -2147221396
		 return "DV_E_DVTARGETDEVICE_SIZE "
		case -2147221395
		 return "DV_E_NOIVIEWOBJECT "
		case -2147221248
		 return "DRAGDROP_E_FIRST "
		case -2147221233
		 return "DRAGDROP_E_LAST "
		case 262400
		 return "DRAGDROP_S_FIRST "
		case 262415
		 return "DRAGDROP_S_LAST "
		case -2147221248
		 return "DRAGDROP_E_NOTREGISTERED "
		case -2147221247
		 return "DRAGDROP_E_ALREADYREGISTERED "
		case -2147221246
		 return "DRAGDROP_E_INVALIDHWND "
		case -2147221232
		 return "CLASSFACTORY_E_FIRST "
		case -2147221217
		 return "CLASSFACTORY_E_LAST "
		case 262416
		 return "CLASSFACTORY_S_FIRST "
		case 262431
		 return "CLASSFACTORY_S_LAST "
		case -2147221232
		 return "CLASS_E_NOAGGREGATION "
		case -2147221231
		 return "CLASS_E_CLASSNOTAVAILABLE "
		case -2147221230
		 return "CLASS_E_NOTLICENSED "
		case -2147221216
		 return "MARSHAL_E_FIRST "
		case -2147221201
		 return "MARSHAL_E_LAST "
		case 262432
		 return "MARSHAL_S_FIRST "
		case 262447
		 return "MARSHAL_S_LAST "
		case -2147221200
		 return "DATA_E_FIRST "
		case -2147221185
		 return "DATA_E_LAST "
		case 262448
		 return "DATA_S_FIRST "
		case 262463
		 return "DATA_S_LAST "
		case -2147221184
		 return "VIEW_E_FIRST "
		case -2147221169
		 return "VIEW_E_LAST "
		case 262464
		 return "VIEW_S_FIRST "
		case 262479
		 return "VIEW_S_LAST "
		case -2147221184
		 return "VIEW_E_DRAW "
		case -2147221168
		 return "REGDB_E_FIRST "
		case -2147221153
		 return "REGDB_E_LAST "
		case 262480
		 return "REGDB_S_FIRST "
		case 262495
		 return "REGDB_S_LAST "
		case -2147221168
		 return "REGDB_E_READREGDB "
		case -2147221167
		 return "REGDB_E_WRITEREGDB "
		case -2147221166
		 return "REGDB_E_KEYMISSING "
		case -2147221165
		 return "REGDB_E_INVALIDVALUE "
		case -2147221164
		 return "REGDB_E_CLASSNOTREG "
		case -2147221163
		 return "REGDB_E_IIDNOTREG "
		case -2147221162
		 return "REGDB_E_BADTHREADINGMODEL "
		case -2147221152
		 return "CAT_E_FIRST "
		case -2147221151
		 return "CAT_E_LAST "
		case -2147221152
		 return "CAT_E_CATIDNOEXIST "
		case -2147221151
		 return "CAT_E_NODESCRIPTION "
		case -2147221148
		 return "CS_E_FIRST "
		case -2147221137
		 return "CS_E_LAST "
		case -2147221148
		 return "CS_E_PACKAGE_NOTFOUND "
		case -2147221147
		 return "CS_E_NOT_DELETABLE "
		case -2147221146
		 return "CS_E_CLASS_NOTFOUND "
		case -2147221145
		 return "CS_E_INVALID_VERSION "
		case -2147221144
		 return "CS_E_NO_CLASSSTORE "
		case -2147221143
		 return "CS_E_OBJECT_NOTFOUND "
		case -2147221142
		 return "CS_E_OBJECT_ALREADY_EXISTS "
		case -2147221141
		 return "CS_E_INVALID_PATH "
		case -2147221140
		 return "CS_E_NETWORK_ERROR "
		case -2147221139
		 return "CS_E_ADMIN_LIMIT_EXCEEDED "
		case -2147221138
		 return "CS_E_SCHEMA_MISMATCH "
		case -2147221137
		 return "CS_E_INTERNAL_ERROR "
		case -2147221136
		 return "CACHE_E_FIRST "
		case -2147221121
		 return "CACHE_E_LAST "
		case 262512
		 return "CACHE_S_FIRST "
		case 262527
		 return "CACHE_S_LAST "
		case -2147221136
		 return "CACHE_E_NOCACHE_UPDATED "
		case -2147221120
		 return "OLEOBJ_E_FIRST "
		case -2147221105
		 return "OLEOBJ_E_LAST "
		case 262528
		 return "OLEOBJ_S_FIRST "
		case 262543
		 return "OLEOBJ_S_LAST "
		case -2147221120
		 return "OLEOBJ_E_NOVERBS "
		case -2147221119
		 return "OLEOBJ_E_INVALIDVERB "
		case -2147221104
		 return "CLIENTSITE_E_FIRST "
		case -2147221089
		 return "CLIENTSITE_E_LAST "
		case 262544
		 return "CLIENTSITE_S_FIRST "
		case 262559
		 return "CLIENTSITE_S_LAST "
		case -2147221088
		 return "INPLACE_E_NOTUNDOABLE "
		case -2147221087
		 return "INPLACE_E_NOTOOLSPACE "
		case -2147221088
		 return "INPLACE_E_FIRST "
		case -2147221073
		 return "INPLACE_E_LAST "
		case 262560
		 return "INPLACE_S_FIRST "
		case 262575
		 return "INPLACE_S_LAST "
		case -2147221072
		 return "ENUM_E_FIRST "
		case -2147221057
		 return "ENUM_E_LAST "
		case 262576
		 return "ENUM_S_FIRST "
		case 262591
		 return "ENUM_S_LAST "
		case -2147221056
		 return "CONVERT10_E_FIRST "
		case -2147221041
		 return "CONVERT10_E_LAST "
		case 262592
		 return "CONVERT10_S_FIRST "
		case 262607
		 return "CONVERT10_S_LAST "
		case -2147221056
		 return "CONVERT10_E_OLESTREAM_GET "
		case -2147221055
		 return "CONVERT10_E_OLESTREAM_PUT "
		case -2147221054
		 return "CONVERT10_E_OLESTREAM_FMT "
		case -2147221053
		 return "CONVERT10_E_OLESTREAM_BITMAP_TO_DIB "
		case -2147221052
		 return "CONVERT10_E_STG_FMT "
		case -2147221051
		 return "CONVERT10_E_STG_NO_STD_STREAM "
		case -2147221050
		 return "CONVERT10_E_STG_DIB_TO_BITMAP "
		case -2147221040
		 return "CLIPBRD_E_FIRST "
		case -2147221025
		 return "CLIPBRD_E_LAST "
		case 262608
		 return "CLIPBRD_S_FIRST "
		case 262623
		 return "CLIPBRD_S_LAST "
		case -2147221040
		 return "CLIPBRD_E_CANT_OPEN "
		case -2147221039
		 return "CLIPBRD_E_CANT_EMPTY "
		case -2147221038
		 return "CLIPBRD_E_CANT_SET "
		case -2147221037
		 return "CLIPBRD_E_BAD_DATA "
		case -2147221036
		 return "CLIPBRD_E_CANT_CLOSE "
		case -2147221024
		 return "MK_E_FIRST "
		case -2147221009
		 return "MK_E_LAST "
		case 262624
		 return "MK_S_FIRST "
		case 262639
		 return "MK_S_LAST "
		case -2147221024
		 return "MK_E_CONNECTMANUALLY "
		case -2147221023
		 return "MK_E_EXCEEDEDDEADLINE "
		case -2147221022
		 return "MK_E_NEEDGENERIC "
		case -2147221021
		 return "MK_E_UNAVAILABLE "
		case -2147221020
		 return "MK_E_SYNTAX "
		case -2147221019
		 return "MK_E_NOOBJECT "
		case -2147221018
		 return "MK_E_INVALIDEXTENSION "
		case -2147221017
		 return "MK_E_INTERMEDIATEINTERFACENOTSUPPORTED "
		case -2147221016
		 return "MK_E_NOTBINDABLE "
		case -2147221015
		 return "MK_E_NOTBOUND "
		case -2147221014
		 return "MK_E_CANTOPENFILE "
		case -2147221013
		 return "MK_E_MUSTBOTHERUSER "
		case -2147221012
		 return "MK_E_NOINVERSE "
		case -2147221011
		 return "MK_E_NOSTORAGE "
		case -2147221010
		 return "MK_E_NOPREFIX "
		case -2147221009
		 return "MK_E_ENUMERATION_FAILED "
		case -2147221008
		 return "CO_E_FIRST "
		case -2147220993
		 return "CO_E_LAST "
		case 262640
		 return "CO_S_FIRST "
		case 262655
		 return "CO_S_LAST "
		case -2147221008
		 return "CO_E_NOTINITIALIZED "
		case -2147221007
		 return "CO_E_ALREADYINITIALIZED "
		case -2147221006
		 return "CO_E_CANTDETERMINECLASS "
		case -2147221005
		 return "CO_E_CLASSSTRING "
		case -2147221004
		 return "CO_E_IIDSTRING "
		case -2147221003
		 return "CO_E_APPNOTFOUND "
		case -2147221002
		 return "CO_E_APPSINGLEUSE "
		case -2147221001
		 return "CO_E_ERRORINAPP "
		case -2147221000
		 return "CO_E_DLLNOTFOUND "
		case -2147220999
		 return "CO_E_ERRORINDLL "
		case -2147220998
		 return "CO_E_WRONGOSFORAPP "
		case -2147220997
		 return "CO_E_OBJNOTREG "
		case -2147220996
		 return "CO_E_OBJISREG "
		case -2147220995
		 return "CO_E_OBJNOTCONNECTED "
		case -2147220994
		 return "CO_E_APPDIDNTREG "
		case -2147220993
		 return "CO_E_RELEASED "
		case -2147220992
		 return "EVENT_E_FIRST "
		case -2147220961
		 return "EVENT_E_LAST "
		case 262656
		 return "EVENT_S_FIRST "
		case 262687
		 return "EVENT_S_LAST "
		case 262656
		 return "EVENT_S_SOME_SUBSCRIBERS_FAILED "
		case -2147220991
		 return "EVENT_E_ALL_SUBSCRIBERS_FAILED "
		case 262658
		 return "EVENT_S_NOSUBSCRIBERS "
		case -2147220989
		 return "EVENT_E_QUERYSYNTAX "
		case -2147220988
		 return "EVENT_E_QUERYFIELD "
		case -2147220987
		 return "EVENT_E_INTERNALEXCEPTION "
		case -2147220986
		 return "EVENT_E_INTERNALERROR "
		case -2147220985
		 return "EVENT_E_INVALID_PER_USER_SID "
		case -2147220984
		 return "EVENT_E_USER_EXCEPTION "
		case -2147220983
		 return "EVENT_E_TOO_MANY_METHODS "
		case -2147220982
		 return "EVENT_E_MISSING_EVENTCLASS "
		case -2147220981
		 return "EVENT_E_NOT_ALL_REMOVED "
		case -2147220980
		 return "EVENT_E_COMPLUS_NOT_INSTALLED "
		case -2147220979
		 return "EVENT_E_CANT_MODIFY_OR_DELETE_UNCONFIGURED_OBJECT "
		case -2147220978
		 return "EVENT_E_CANT_MODIFY_OR_DELETE_CONFIGURED_OBJECT "
		case -2147220977
		 return "EVENT_E_INVALID_EVENT_CLASS_PARTITION "
		case -2147220976
		 return "EVENT_E_PER_USER_SID_NOT_LOGGED_ON "
		case -2147168256
		 return "XACT_E_FIRST "
		case -2147168215
		 return "XACT_E_LAST "
		case 315392
		 return "XACT_S_FIRST "
		case 315408
		 return "XACT_S_LAST "
		case -2147168256
		 return "XACT_E_ALREADYOTHERSINGLEPHASE "
		case -2147168255
		 return "XACT_E_CANTRETAIN "
		case -2147168254
		 return "XACT_E_COMMITFAILED "
		case -2147168253
		 return "XACT_E_COMMITPREVENTED "
		case -2147168252
		 return "XACT_E_HEURISTICABORT "
		case -2147168251
		 return "XACT_E_HEURISTICCOMMIT "
		case -2147168250
		 return "XACT_E_HEURISTICDAMAGE "
		case -2147168249
		 return "XACT_E_HEURISTICDANGER "
		case -2147168248
		 return "XACT_E_ISOLATIONLEVEL "
		case -2147168247
		 return "XACT_E_NOASYNC "
		case -2147168246
		 return "XACT_E_NOENLIST "
		case -2147168245
		 return "XACT_E_NOISORETAIN "
		case -2147168244
		 return "XACT_E_NORESOURCE "
		case -2147168243
		 return "XACT_E_NOTCURRENT "
		case -2147168242
		 return "XACT_E_NOTRANSACTION "
		case -2147168241
		 return "XACT_E_NOTSUPPORTED "
		case -2147168240
		 return "XACT_E_UNKNOWNRMGRID "
		case -2147168239
		 return "XACT_E_WRONGSTATE "
		case -2147168238
		 return "XACT_E_WRONGUOW "
		case -2147168237
		 return "XACT_E_XTIONEXISTS "
		case -2147168236
		 return "XACT_E_NOIMPORTOBJECT "
		case -2147168235
		 return "XACT_E_INVALIDCOOKIE "
		case -2147168234
		 return "XACT_E_INDOUBT "
		case -2147168233
		 return "XACT_E_NOTIMEOUT "
		case -2147168232
		 return "XACT_E_ALREADYINPROGRESS "
		case -2147168231
		 return "XACT_E_ABORTED "
		case -2147168230
		 return "XACT_E_LOGFULL "
		case -2147168229
		 return "XACT_E_TMNOTAVAILABLE "
		case -2147168228
		 return "XACT_E_CONNECTION_DOWN "
		case -2147168227
		 return "XACT_E_CONNECTION_DENIED "
		case -2147168226
		 return "XACT_E_REENLISTTIMEOUT "
		case -2147168225
		 return "XACT_E_TIP_CONNECT_FAILED "
		case -2147168224
		 return "XACT_E_TIP_PROTOCOL_ERROR "
		case -2147168223
		 return "XACT_E_TIP_PULL_FAILED "
		case -2147168222
		 return "XACT_E_DEST_TMNOTAVAILABLE "
		case -2147168221
		 return "XACT_E_TIP_DISABLED "
		case -2147168220
		 return "XACT_E_NETWORK_TX_DISABLED "
		case -2147168219
		 return "XACT_E_PARTNER_NETWORK_TX_DISABLED "
		case -2147168218
		 return "XACT_E_XA_TX_DISABLED "
		case -2147168217
		 return "XACT_E_UNABLE_TO_READ_DTC_CONFIG "
		case -2147168216
		 return "XACT_E_UNABLE_TO_LOAD_DTC_PROXY "
		case -2147168215
		 return "XACT_E_ABORTING "
		case -2147168128
		 return "XACT_E_CLERKNOTFOUND "
		case -2147168127
		 return "XACT_E_CLERKEXISTS "
		case -2147168126
		 return "XACT_E_RECOVERYINPROGRESS "
		case -2147168125
		 return "XACT_E_TRANSACTIONCLOSED "
		case -2147168124
		 return "XACT_E_INVALIDLSN "
		case -2147168123
		 return "XACT_E_REPLAYREQUEST "
		case 315392
		 return "XACT_S_ASYNC "
		case 315393
		 return "XACT_S_DEFECT "
		case 315394
		 return "XACT_S_READONLY "
		case 315395
		 return "XACT_S_SOMENORETAIN "
		case 315396
		 return "XACT_S_OKINFORM "
		case 315397
		 return "XACT_S_MADECHANGESCONTENT "
		case 315398
		 return "XACT_S_MADECHANGESINFORM "
		case 315399
		 return "XACT_S_ALLNORETAIN "
		case 315400
		 return "XACT_S_ABORTING "
		case 315401
		 return "XACT_S_SINGLEPHASE "
		case 315402
		 return "XACT_S_LOCALLY_OK "
		case 315408
		 return "XACT_S_LASTRESOURCEMANAGER "
		case -2147164160
		 return "CONTEXT_E_FIRST "
		case -2147164113
		 return "CONTEXT_E_LAST "
		case 319488
		 return "CONTEXT_S_FIRST "
		case 319535
		 return "CONTEXT_S_LAST "
		case -2147164158
		 return "CONTEXT_E_ABORTED "
		case -2147164157
		 return "CONTEXT_E_ABORTING "
		case -2147164156
		 return "CONTEXT_E_NOCONTEXT "
		case -2147164155
		 return "CONTEXT_E_WOULD_DEADLOCK "
		case -2147164154
		 return "CONTEXT_E_SYNCH_TIMEOUT "
		case -2147164153
		 return "CONTEXT_E_OLDREF "
		case -2147164148
		 return "CONTEXT_E_ROLENOTFOUND "
		case -2147164145
		 return "CONTEXT_E_TMNOTAVAILABLE "
		case -2147164127
		 return "CO_E_ACTIVATIONFAILED "
		case -2147164126
		 return "CO_E_ACTIVATIONFAILED_EVENTLOGGED "
		case -2147164125
		 return "CO_E_ACTIVATIONFAILED_CATALOGERROR "
		case -2147164124
		 return "CO_E_ACTIVATIONFAILED_TIMEOUT "
		case -2147164123
		 return "CO_E_INITIALIZATIONFAILED "
		case -2147164122
		 return "CONTEXT_E_NOJIT "
		case -2147164121
		 return "CONTEXT_E_NOTRANSACTION "
		case -2147164120
		 return "CO_E_THREADINGMODEL_CHANGED "
		case -2147164119
		 return "CO_E_NOIISINTRINSICS "
		case -2147164118
		 return "CO_E_NOCOOKIES "
		case -2147164117
		 return "CO_E_DBERROR "
		case -2147164116
		 return "CO_E_NOTPOOLED "
		case -2147164115
		 return "CO_E_NOTCONSTRUCTED "
		case -2147164114
		 return "CO_E_NOSYNCHRONIZATION "
		case -2147164113
		 return "CO_E_ISOLEVELMISMATCH "
		case 262144
		 return "OLE_S_USEREG "
		case 262145
		 return "OLE_S_STATIC "
		case 262146
		 return "OLE_S_MAC_CLIPFORMAT "
		case 262400
		 return "DRAGDROP_S_DROP "
		case 262401
		 return "DRAGDROP_S_CANCEL "
		case 262402
		 return "DRAGDROP_S_USEDEFAULTCURSORS "
		case 262448
		 return "DATA_S_SAMEFORMATETC "
		case 262464
		 return "VIEW_S_ALREADY_FROZEN "
		case 262512
		 return "CACHE_S_FORMATETC_NOTSUPPORTED "
		case 262513
		 return "CACHE_S_SAMECACHE "
		case 262514
		 return "CACHE_S_SOMECACHES_NOTUPDATED "
		case 262528
		 return "OLEOBJ_S_INVALIDVERB "
		case 262529
		 return "OLEOBJ_S_CANNOT_DOVERB_NOW "
		case 262530
		 return "OLEOBJ_S_INVALIDHWND "
		case 262560
		 return "INPLACE_S_TRUNCATED "
		case 262592
		 return "CONVERT10_S_NO_PRESENTATION "
		case 262626
		 return "MK_S_REDUCED_TO_SELF "
		case 262628
		 return "MK_S_ME "
		case 262629
		 return "MK_S_HIM "
		case 262630
		 return "MK_S_US "
		case 262631
		 return "MK_S_MONIKERALREADYREGISTERED "
		case 267008
		 return "SCHED_S_TASK_READY "
		case 267009
		 return "SCHED_S_TASK_RUNNING "
		case 267010
		 return "SCHED_S_TASK_DISABLED "
		case 267011
		 return "SCHED_S_TASK_HAS_NOT_RUN "
		case 267012
		 return "SCHED_S_TASK_NO_MORE_RUNS "
		case 267013
		 return "SCHED_S_TASK_NOT_SCHEDULED "
		case 267014
		 return "SCHED_S_TASK_TERMINATED "
		case 267015
		 return "SCHED_S_TASK_NO_VALID_TRIGGERS "
		case 267016
		 return "SCHED_S_EVENT_TRIGGER "
		case -2147216631
		 return "SCHED_E_TRIGGER_NOT_FOUND "
		case -2147216630
		 return "SCHED_E_TASK_NOT_READY "
		case -2147216629
		 return "SCHED_E_TASK_NOT_RUNNING "
		case -2147216628
		 return "SCHED_E_SERVICE_NOT_INSTALLED "
		case -2147216627
		 return "SCHED_E_CANNOT_OPEN_TASK "
		case -2147216626
		 return "SCHED_E_INVALID_TASK "
		case -2147216625
		 return "SCHED_E_ACCOUNT_INFORMATION_NOT_SET "
		case -2147216624
		 return "SCHED_E_ACCOUNT_NAME_NOT_FOUND "
		case -2147216623
		 return "SCHED_E_ACCOUNT_DBASE_CORRUPT "
		case -2147216622
		 return "SCHED_E_NO_SECURITY_SERVICES "
		case -2147216621
		 return "SCHED_E_UNKNOWN_OBJECT_VERSION "
		case -2147216620
		 return "SCHED_E_UNSUPPORTED_ACCOUNT_OPTION "
		case -2147216619
		 return "SCHED_E_SERVICE_NOT_RUNNING "
		case -2146959359
		 return "CO_E_CLASS_CREATE_FAILED "
		case -2146959358
		 return "CO_E_SCM_ERROR "
		case -2146959357
		 return "CO_E_SCM_RPC_FAILURE "
		case -2146959356
		 return "CO_E_BAD_PATH "
		case -2146959355
		 return "CO_E_SERVER_EXEC_FAILURE "
		case -2146959354
		 return "CO_E_OBJSRV_RPC_FAILURE "
		case -2146959353
		 return "MK_E_NO_NORMALIZED "
		case -2146959352
		 return "CO_E_SERVER_STOPPING "
		case -2146959351
		 return "MEM_E_INVALID_ROOT "
		case -2146959344
		 return "MEM_E_INVALID_LINK "
		case -2146959343
		 return "MEM_E_INVALID_SIZE "
		case 524306
		 return "CO_S_NOTALLINTERFACES "
		case 524307
		 return "CO_S_MACHINENAMENOTFOUND "
		case -2147352575
		 return "DISP_E_UNKNOWNINTERFACE "
		case -2147352573
		 return "DISP_E_MEMBERNOTFOUND "
		case -2147352572
		 return "DISP_E_PARAMNOTFOUND "
		case -2147352571
		 return "DISP_E_TYPEMISMATCH "
		case -2147352570
		 return "DISP_E_UNKNOWNNAME "
		case -2147352569
		 return "DISP_E_NONAMEDARGS "
		case -2147352568
		 return "DISP_E_BADVARTYPE "
		case -2147352567
		 return "DISP_E_EXCEPTION "
		case -2147352566
		 return "DISP_E_OVERFLOW "
		case -2147352565
		 return "DISP_E_BADINDEX "
		case -2147352564
		 return "DISP_E_UNKNOWNLCID "
		case -2147352563
		 return "DISP_E_ARRAYISLOCKED "
		case -2147352562
		 return "DISP_E_BADPARAMCOUNT "
		case -2147352561
		 return "DISP_E_PARAMNOTOPTIONAL "
		case -2147352560
		 return "DISP_E_BADCALLEE "
		case -2147352559
		 return "DISP_E_NOTACOLLECTION "
		case -2147352558
		 return "DISP_E_DIVBYZERO "
		case -2147352557
		 return "DISP_E_BUFFERTOOSMALL "
		case -2147319786
		 return "TYPE_E_BUFFERTOOSMALL "
		case -2147319785
		 return "TYPE_E_FIELDNOTFOUND "
		case -2147319784
		 return "TYPE_E_INVDATAREAD "
		case -2147319783
		 return "TYPE_E_UNSUPFORMAT "
		case -2147319780
		 return "TYPE_E_REGISTRYACCESS "
		case -2147319779
		 return "TYPE_E_LIBNOTREGISTERED "
		case -2147319769
		 return "TYPE_E_UNDEFINEDTYPE "
		case -2147319768
		 return "TYPE_E_QUALIFIEDNAMEDISALLOWED "
		case -2147319767
		 return "TYPE_E_INVALIDSTATE "
		case -2147319766
		 return "TYPE_E_WRONGTYPEKIND "
		case -2147319765
		 return "TYPE_E_ELEMENTNOTFOUND "
		case -2147319764
		 return "TYPE_E_AMBIGUOUSNAME "
		case -2147319763
		 return "TYPE_E_NAMECONFLICT "
		case -2147319762
		 return "TYPE_E_UNKNOWNLCID "
		case -2147319761
		 return "TYPE_E_DLLFUNCTIONNOTFOUND "
		case -2147317571
		 return "TYPE_E_BADMODULEKIND "
		case -2147317563
		 return "TYPE_E_SIZETOOBIG "
		case -2147317562
		 return "TYPE_E_DUPLICATEID "
		case -2147317553
		 return "TYPE_E_INVALIDID "
		case -2147316576
		 return "TYPE_E_TYPEMISMATCH "
		case -2147316575
		 return "TYPE_E_OUTOFBOUNDS "
		case -2147316574
		 return "TYPE_E_IOERROR "
		case -2147316573
		 return "TYPE_E_CANTCREATETMPFILE "
		case -2147312566
		 return "TYPE_E_CANTLOADLIBRARY "
		case -2147312509
		 return "TYPE_E_INCONSISTENTPROPFUNCS "
		case -2147312508
		 return "TYPE_E_CIRCULARTYPE "
		case -2147287039
		 return "STG_E_INVALIDFUNCTION "
		case -2147287038
		 return "STG_E_FILENOTFOUND "
		case -2147287037
		 return "STG_E_PATHNOTFOUND "
		case -2147287036
		 return "STG_E_TOOMANYOPENFILES "
		case -2147287035
		 return "STG_E_ACCESSDENIED "
		case -2147287034
		 return "STG_E_INVALIDHANDLE "
		case -2147287032
		 return "STG_E_INSUFFICIENTMEMORY "
		case -2147287031
		 return "STG_E_INVALIDPOINTER "
		case -2147287022
		 return "STG_E_NOMOREFILES "
		case -2147287021
		 return "STG_E_DISKISWRITEPROTECTED "
		case -2147287015
		 return "STG_E_SEEKERROR "
		case -2147287011
		 return "STG_E_WRITEFAULT "
		case -2147287010
		 return "STG_E_READFAULT "
		case -2147287008
		 return "STG_E_SHAREVIOLATION "
		case -2147287007
		 return "STG_E_LOCKVIOLATION "
		case -2147286960
		 return "STG_E_FILEALREADYEXISTS "
		case -2147286953
		 return "STG_E_INVALIDPARAMETER "
		case -2147286928
		 return "STG_E_MEDIUMFULL "
		case -2147286800
		 return "STG_E_PROPSETMISMATCHED "
		case -2147286790
		 return "STG_E_ABNORMALAPIEXIT "
		case -2147286789
		 return "STG_E_INVALIDHEADER "
		case -2147286788
		 return "STG_E_INVALIDNAME "
		case -2147286787
		 return "STG_E_UNKNOWN "
		case -2147286786
		 return "STG_E_UNIMPLEMENTEDFUNCTION "
		case -2147286785
		 return "STG_E_INVALIDFLAG "
		case -2147286784
		 return "STG_E_INUSE "
		case -2147286783
		 return "STG_E_NOTCURRENT "
		case -2147286782
		 return "STG_E_REVERTED "
		case -2147286781
		 return "STG_E_CANTSAVE "
		case -2147286780
		 return "STG_E_OLDFORMAT "
		case -2147286779
		 return "STG_E_OLDDLL "
		case -2147286778
		 return "STG_E_SHAREREQUIRED "
		case -2147286777
		 return "STG_E_NOTFILEBASEDSTORAGE "
		case -2147286776
		 return "STG_E_EXTANTMARSHALLINGS "
		case -2147286775
		 return "STG_E_DOCFILECORRUPT "
		case -2147286768
		 return "STG_E_BADBASEADDRESS "
		case -2147286767
		 return "STG_E_DOCFILETOOLARGE "
		case -2147286766
		 return "STG_E_NOTSIMPLEFORMAT "
		case -2147286527
		 return "STG_E_INCOMPLETE "
		case -2147286526
		 return "STG_E_TERMINATED "
		case 197120
		 return "STG_S_CONVERTED "
		case 197121
		 return "STG_S_BLOCK "
		case 197122
		 return "STG_S_RETRYNOW "
		case 197123
		 return "STG_S_MONITORING "
		case 197124
		 return "STG_S_MULTIPLEOPENS "
		case 197125
		 return "STG_S_CONSOLIDATIONFAILED "
		case 197126
		 return "STG_S_CANNOTCONSOLIDATE "
		case -2147286267
		 return "STG_E_STATUS_COPY_PROTECTION_FAILURE "
		case -2147286266
		 return "STG_E_CSS_AUTHENTICATION_FAILURE "
		case -2147286265
		 return "STG_E_CSS_KEY_NOT_PRESENT "
		case -2147286264
		 return "STG_E_CSS_KEY_NOT_ESTABLISHED "
		case -2147286263
		 return "STG_E_CSS_SCRAMBLED_SECTOR "
		case -2147286262
		 return "STG_E_CSS_REGION_MISMATCH "
		case -2147286261
		 return "STG_E_RESETS_EXHAUSTED "
		case -2147418111
		 return "RPC_E_CALL_REJECTED "
		case -2147418110
		 return "RPC_E_CALL_CANCELED "
		case -2147418109
		 return "RPC_E_CANTPOST_INSENDCALL "
		case -2147418108
		 return "RPC_E_CANTCALLOUT_INASYNCCALL "
		case -2147418107
		 return "RPC_E_CANTCALLOUT_INEXTERNALCALL "
		case -2147418106
		 return "RPC_E_CONNECTION_TERMINATED "
		case -2147418105
		 return "RPC_E_SERVER_DIED "
		case -2147418104
		 return "RPC_E_CLIENT_DIED "
		case -2147418103
		 return "RPC_E_INVALID_DATAPACKET "
		case -2147418102
		 return "RPC_E_CANTTRANSMIT_CALL "
		case -2147418101
		 return "RPC_E_CLIENT_CANTMARSHAL_DATA "
		case -2147418100
		 return "RPC_E_CLIENT_CANTUNMARSHAL_DATA "
		case -2147418099
		 return "RPC_E_SERVER_CANTMARSHAL_DATA "
		case -2147418098
		 return "RPC_E_SERVER_CANTUNMARSHAL_DATA "
		case -2147418097
		 return "RPC_E_INVALID_DATA "
		case -2147418096
		 return "RPC_E_INVALID_PARAMETER "
		case -2147418095
		 return "RPC_E_CANTCALLOUT_AGAIN "
		case -2147418094
		 return "RPC_E_SERVER_DIED_DNE "
		case -2147417856
		 return "RPC_E_SYS_CALL_FAILED "
		case -2147417855
		 return "RPC_E_OUT_OF_RESOURCES "
		case -2147417854
		 return "RPC_E_ATTEMPTED_MULTITHREAD "
		case -2147417853
		 return "RPC_E_NOT_REGISTERED "
		case -2147417852
		 return "RPC_E_FAULT "
		case -2147417851
		 return "RPC_E_SERVERFAULT "
		case -2147417850
		 return "RPC_E_CHANGED_MODE "
		case -2147417849
		 return "RPC_E_INVALIDMETHOD "
		case -2147417848
		 return "RPC_E_DISCONNECTED "
		case -2147417847
		 return "RPC_E_RETRY "
		case -2147417846
		 return "RPC_E_SERVERCALL_RETRYLATER "
		case -2147417845
		 return "RPC_E_SERVERCALL_REJECTED "
		case -2147417844
		 return "RPC_E_INVALID_CALLDATA "
		case -2147417843
		 return "RPC_E_CANTCALLOUT_ININPUTSYNCCALL "
		case -2147417842
		 return "RPC_E_WRONG_THREAD "
		case -2147417841
		 return "RPC_E_THREAD_NOT_INIT "
		case -2147417840
		 return "RPC_E_VERSION_MISMATCH "
		case -2147417839
		 return "RPC_E_INVALID_HEADER "
		case -2147417838
		 return "RPC_E_INVALID_EXTENSION "
		case -2147417837
		 return "RPC_E_INVALID_IPID "
		case -2147417836
		 return "RPC_E_INVALID_OBJECT "
		case -2147417835
		 return "RPC_S_CALLPENDING "
		case -2147417834
		 return "RPC_S_WAITONTIMER "
		case -2147417833
		 return "RPC_E_CALL_COMPLETE "
		case -2147417832
		 return "RPC_E_UNSECURE_CALL "
		case -2147417831
		 return "RPC_E_TOO_LATE "
		case -2147417830
		 return "RPC_E_NO_GOOD_SECURITY_PACKAGES "
		case -2147417829
		 return "RPC_E_ACCESS_DENIED "
		case -2147417828
		 return "RPC_E_REMOTE_DISABLED "
		case -2147417827
		 return "RPC_E_INVALID_OBJREF "
		case -2147417826
		 return "RPC_E_NO_CONTEXT "
		case -2147417825
		 return "RPC_E_TIMEOUT "
		case -2147417824
		 return "RPC_E_NO_SYNC "
		case -2147417823
		 return "RPC_E_FULLSIC_REQUIRED "
		case -2147417822
		 return "RPC_E_INVALID_STD_NAME "
		case -2147417821
		 return "CO_E_FAILEDTOIMPERSONATE "
		case -2147417820
		 return "CO_E_FAILEDTOGETSECCTX "
		case -2147417819
		 return "CO_E_FAILEDTOOPENTHREADTOKEN "
		case -2147417818
		 return "CO_E_FAILEDTOGETTOKENINFO "
		case -2147417817
		 return "CO_E_TRUSTEEDOESNTMATCHCLIENT "
		case -2147417816
		 return "CO_E_FAILEDTOQUERYCLIENTBLANKET "
		case -2147417815
		 return "CO_E_FAILEDTOSETDACL "
		case -2147417814
		 return "CO_E_ACCESSCHECKFAILED "
		case -2147417813
		 return "CO_E_NETACCESSAPIFAILED "
		case -2147417812
		 return "CO_E_WRONGTRUSTEENAMESYNTAX "
		case -2147417811
		 return "CO_E_INVALIDSID "
		case -2147417810
		 return "CO_E_CONVERSIONFAILED "
		case -2147417809
		 return "CO_E_NOMATCHINGSIDFOUND "
		case -2147417808
		 return "CO_E_LOOKUPACCSIDFAILED "
		case -2147417807
		 return "CO_E_NOMATCHINGNAMEFOUND "
		case -2147417806
		 return "CO_E_LOOKUPACCNAMEFAILED "
		case -2147417805
		 return "CO_E_SETSERLHNDLFAILED "
		case -2147417804
		 return "CO_E_FAILEDTOGETWINDIR "
		case -2147417803
		 return "CO_E_PATHTOOLONG "
		case -2147417802
		 return "CO_E_FAILEDTOGENUUID "
		case -2147417801
		 return "CO_E_FAILEDTOCREATEFILE "
		case -2147417800
		 return "CO_E_FAILEDTOCLOSEHANDLE "
		case -2147417799
		 return "CO_E_EXCEEDSYSACLLIMIT "
		case -2147417798
		 return "CO_E_ACESINWRONGORDER "
		case -2147417797
		 return "CO_E_INCOMPATIBLESTREAMVERSION "
		case -2147417796
		 return "CO_E_FAILEDTOOPENPROCESSTOKEN "
		case -2147417795
		 return "CO_E_DECODEFAILED "
		case -2147417793
		 return "CO_E_ACNOTINITIALIZED "
		case -2147417792
		 return "CO_E_CANCEL_DISABLED "
		case -2147352577
		 return "RPC_E_UNEXPECTED "
		case -1073151999
		 return "ERROR_AUDITING_DISABLED "
		case -1073151998
		 return "ERROR_ALL_SIDS_FILTERED "
		case -2146893823
		 return "NTE_BAD_UID "
		case -2146893822
		 return "NTE_BAD_HASH "
		case -2146893821
		 return "NTE_BAD_KEY "
		case -2146893820
		 return "NTE_BAD_LEN "
		case -2146893819
		 return "NTE_BAD_DATA "
		case -2146893818
		 return "NTE_BAD_SIGNATURE "
		case -2146893817
		 return "NTE_BAD_VER "
		case -2146893816
		 return "NTE_BAD_ALGID "
		case -2146893815
		 return "NTE_BAD_FLAGS "
		case -2146893814
		 return "NTE_BAD_TYPE "
		case -2146893813
		 return "NTE_BAD_KEY_STATE "
		case -2146893812
		 return "NTE_BAD_HASH_STATE "
		case -2146893811
		 return "NTE_NO_KEY "
		case -2146893810
		 return "NTE_NO_MEMORY "
		case -2146893809
		 return "NTE_EXISTS "
		case -2146893808
		 return "NTE_PERM "
		case -2146893807
		 return "NTE_NOT_FOUND "
		case -2146893806
		 return "NTE_DOUBLE_ENCRYPT "
		case -2146893805
		 return "NTE_BAD_PROVIDER "
		case -2146893804
		 return "NTE_BAD_PROV_TYPE "
		case -2146893803
		 return "NTE_BAD_PUBLIC_KEY "
		case -2146893802
		 return "NTE_BAD_KEYSET "
		case -2146893801
		 return "NTE_PROV_TYPE_NOT_DEF "
		case -2146893800
		 return "NTE_PROV_TYPE_ENTRY_BAD "
		case -2146893799
		 return "NTE_KEYSET_NOT_DEF "
		case -2146893798
		 return "NTE_KEYSET_ENTRY_BAD "
		case -2146893797
		 return "NTE_PROV_TYPE_NO_MATCH "
		case -2146893796
		 return "NTE_SIGNATURE_FILE_BAD "
		case -2146893795
		 return "NTE_PROVIDER_DLL_FAIL "
		case -2146893794
		 return "NTE_PROV_DLL_NOT_FOUND "
		case -2146893793
		 return "NTE_BAD_KEYSET_PARAM "
		case -2146893792
		 return "NTE_FAIL "
		case -2146893791
		 return "NTE_SYS_ERR "
		case -2146893790
		 return "NTE_SILENT_CONTEXT "
		case -2146893789
		 return "NTE_TOKEN_KEYSET_STORAGE_FULL "
		case -2146893788
		 return "NTE_TEMPORARY_PROFILE "
		case -2146893787
		 return "NTE_FIXEDPARAMETER "
		case -2146893056
		 return "SEC_E_INSUFFICIENT_MEMORY "
		case -2146893055
		 return "SEC_E_INVALID_HANDLE "
		case -2146893054
		 return "SEC_E_UNSUPPORTED_FUNCTION "
		case -2146893053
		 return "SEC_E_TARGET_UNKNOWN "
		case -2146893052
		 return "SEC_E_INTERNAL_ERROR "
		case -2146893051
		 return "SEC_E_SECPKG_NOT_FOUND "
		case -2146893050
		 return "SEC_E_NOT_OWNER "
		case -2146893049
		 return "SEC_E_CANNOT_INSTALL "
		case -2146893048
		 return "SEC_E_INVALID_TOKEN "
		case -2146893047
		 return "SEC_E_CANNOT_PACK "
		case -2146893046
		 return "SEC_E_QOP_NOT_SUPPORTED "
		case -2146893045
		 return "SEC_E_NO_IMPERSONATION "
		case -2146893044
		 return "SEC_E_LOGON_DENIED "
		case -2146893043
		 return "SEC_E_UNKNOWN_CREDENTIALS "
		case -2146893042
		 return "SEC_E_NO_CREDENTIALS "
		case -2146893041
		 return "SEC_E_MESSAGE_ALTERED "
		case -2146893040
		 return "SEC_E_OUT_OF_SEQUENCE "
		case -2146893039
		 return "SEC_E_NO_AUTHENTICATING_AUTHORITY "
		case 590610
		 return "SEC_I_CONTINUE_NEEDED "
		case 590611
		 return "SEC_I_COMPLETE_NEEDED "
		case 590612
		 return "SEC_I_COMPLETE_AND_CONTINUE "
		case 590613
		 return "SEC_I_LOCAL_LOGON "
		case -2146893034
		 return "SEC_E_BAD_PKGID "
		case -2146893033
		 return "SEC_E_CONTEXT_EXPIRED "
		case 590615
		 return "SEC_I_CONTEXT_EXPIRED "
		case -2146893032
		 return "SEC_E_INCOMPLETE_MESSAGE "
		case -2146893024
		 return "SEC_E_INCOMPLETE_CREDENTIALS "
		case -2146893023
		 return "SEC_E_BUFFER_TOO_SMALL "
		case 590624
		 return "SEC_I_INCOMPLETE_CREDENTIALS "
		case 590625
		 return "SEC_I_RENEGOTIATE "
		case -2146893022
		 return "SEC_E_WRONG_PRINCIPAL "
		case 590627
		 return "SEC_I_NO_LSA_CONTEXT "
		case -2146893020
		 return "SEC_E_TIME_SKEW "
		case -2146893019
		 return "SEC_E_UNTRUSTED_ROOT "
		case -2146893018
		 return "SEC_E_ILLEGAL_MESSAGE "
		case -2146893017
		 return "SEC_E_CERT_UNKNOWN "
		case -2146893016
		 return "SEC_E_CERT_EXPIRED "
		case -2146893015
		 return "SEC_E_ENCRYPT_FAILURE "
		case -2146893008
		 return "SEC_E_DECRYPT_FAILURE "
		case -2146893007
		 return "SEC_E_ALGORITHM_MISMATCH "
		case -2146893006
		 return "SEC_E_SECURITY_QOS_FAILED "
		case -2146893005
		 return "SEC_E_UNFINISHED_CONTEXT_DELETED "
		case -2146893004
		 return "SEC_E_NO_TGT_REPLY "
		case -2146893003
		 return "SEC_E_NO_IP_ADDRESSES "
		case -2146893002
		 return "SEC_E_WRONG_CREDENTIAL_HANDLE "
		case -2146893001
		 return "SEC_E_CRYPTO_SYSTEM_INVALID "
		case -2146893000
		 return "SEC_E_MAX_REFERRALS_EXCEEDED "
		case -2146892999
		 return "SEC_E_MUST_BE_KDC "
		case -2146892998
		 return "SEC_E_STRONG_CRYPTO_NOT_SUPPORTED "
		case -2146892997
		 return "SEC_E_TOO_MANY_PRINCIPALS "
		case -2146892996
		 return "SEC_E_NO_PA_DATA "
		case -2146892995
		 return "SEC_E_PKINIT_NAME_MISMATCH "
		case -2146892994
		 return "SEC_E_SMARTCARD_LOGON_REQUIRED "
		case -2146892993
		 return "SEC_E_SHUTDOWN_IN_PROGRESS "
		case -2146892992
		 return "SEC_E_KDC_INVALID_REQUEST "
		case -2146892991
		 return "SEC_E_KDC_UNABLE_TO_REFER "
		case -2146892990
		 return "SEC_E_KDC_UNKNOWN_ETYPE "
		case -2146892989
		 return "SEC_E_UNSUPPORTED_PREAUTH "
		case -2146892987
		 return "SEC_E_DELEGATION_REQUIRED "
		case -2146892986
		 return "SEC_E_BAD_BINDINGS "
		case -2146892985
		 return "SEC_E_MULTIPLE_ACCOUNTS "
		case -2146892984
		 return "SEC_E_NO_KERB_KEY "
		case -2146892983
		 return "SEC_E_CERT_WRONG_USAGE "
		case -2146892976
		 return "SEC_E_DOWNGRADE_DETECTED "
		case -2146892975
		 return "SEC_E_SMARTCARD_CERT_REVOKED "
		case -2146892974
		 return "SEC_E_ISSUING_CA_UNTRUSTED "
		case -2146892973
		 return "SEC_E_REVOCATION_OFFLINE_C "
		case -2146892972
		 return "SEC_E_PKINIT_CLIENT_FAILURE "
		case -2146892971
		 return "SEC_E_SMARTCARD_CERT_EXPIRED "
		case -2146892970
		 return "SEC_E_NO_S4U_PROT_SUPPORT "
		case -2146892969
		 return "SEC_E_CROSSREALM_DELEGATION_FAILURE "
		case -2146892968
		 return "SEC_E_REVOCATION_OFFLINE_KDC "
		case -2146892967
		 return "SEC_E_ISSUING_CA_UNTRUSTED_KDC "
		case -2146892966
		 return "SEC_E_KDC_CERT_EXPIRED "
		case -2146892965
		 return "SEC_E_KDC_CERT_REVOKED "
		case -2146889727
		 return "CRYPT_E_MSG_ERROR "
		case -2146889726
		 return "CRYPT_E_UNKNOWN_ALGO "
		case -2146889725
		 return "CRYPT_E_OID_FORMAT "
		case -2146889724
		 return "CRYPT_E_INVALID_MSG_TYPE "
		case -2146889723
		 return "CRYPT_E_UNEXPECTED_ENCODING "
		case -2146889722
		 return "CRYPT_E_AUTH_ATTR_MISSING "
		case -2146889721
		 return "CRYPT_E_HASH_VALUE "
		case -2146889720
		 return "CRYPT_E_INVALID_INDEX "
		case -2146889719
		 return "CRYPT_E_ALREADY_DECRYPTED "
		case -2146889718
		 return "CRYPT_E_NOT_DECRYPTED "
		case -2146889717
		 return "CRYPT_E_RECIPIENT_NOT_FOUND "
		case -2146889716
		 return "CRYPT_E_CONTROL_TYPE "
		case -2146889715
		 return "CRYPT_E_ISSUER_SERIALNUMBER "
		case -2146889714
		 return "CRYPT_E_SIGNER_NOT_FOUND "
		case -2146889713
		 return "CRYPT_E_ATTRIBUTES_MISSING "
		case -2146889712
		 return "CRYPT_E_STREAM_MSG_NOT_READY "
		case -2146889711
		 return "CRYPT_E_STREAM_INSUFFICIENT_DATA "
		case 593938
		 return "CRYPT_I_NEW_PROTECTION_REQUIRED "
		case -2146885631
		 return "CRYPT_E_BAD_LEN "
		case -2146885630
		 return "CRYPT_E_BAD_ENCODE "
		case -2146885629
		 return "CRYPT_E_FILE_ERROR "
		case -2146885628
		 return "CRYPT_E_NOT_FOUND "
		case -2146885627
		 return "CRYPT_E_EXISTS "
		case -2146885626
		 return "CRYPT_E_NO_PROVIDER "
		case -2146885625
		 return "CRYPT_E_SELF_SIGNED "
		case -2146885624
		 return "CRYPT_E_DELETED_PREV "
		case -2146885623
		 return "CRYPT_E_NO_MATCH "
		case -2146885622
		 return "CRYPT_E_UNEXPECTED_MSG_TYPE "
		case -2146885621
		 return "CRYPT_E_NO_KEY_PROPERTY "
		case -2146885620
		 return "CRYPT_E_NO_DECRYPT_CERT "
		case -2146885619
		 return "CRYPT_E_BAD_MSG "
		case -2146885618
		 return "CRYPT_E_NO_SIGNER "
		case -2146885617
		 return "CRYPT_E_PENDING_CLOSE "
		case -2146885616
		 return "CRYPT_E_REVOKED "
		case -2146885615
		 return "CRYPT_E_NO_REVOCATION_DLL "
		case -2146885614
		 return "CRYPT_E_NO_REVOCATION_CHECK "
		case -2146885613
		 return "CRYPT_E_REVOCATION_OFFLINE "
		case -2146885612
		 return "CRYPT_E_NOT_IN_REVOCATION_DATABASE "
		case -2146885600
		 return "CRYPT_E_INVALID_NUMERIC_STRING "
		case -2146885599
		 return "CRYPT_E_INVALID_RETURNABLE_STRING "
		case -2146885598
		 return "CRYPT_E_INVALID_IA5_STRING "
		case -2146885597
		 return "CRYPT_E_INVALID_X500_STRING "
		case -2146885596
		 return "CRYPT_E_NOT_CHAR_STRING "
		case -2146885595
		 return "CRYPT_E_FILERESIZED "
		case -2146885594
		 return "CRYPT_E_SECURITY_SETTINGS "
		case -2146885593
		 return "CRYPT_E_NO_VERIFY_USAGE_DLL "
		case -2146885592
		 return "CRYPT_E_NO_VERIFY_USAGE_CHECK "
		case -2146885591
		 return "CRYPT_E_VERIFY_USAGE_OFFLINE "
		case -2146885590
		 return "CRYPT_E_NOT_IN_CTL "
		case -2146885589
		 return "CRYPT_E_NO_TRUSTED_SIGNER "
		case -2146885588
		 return "CRYPT_E_MISSING_PUBKEY_PARA "
		case -2146881536
		 return "CRYPT_E_OSS_ERROR "
		case -2146881535
		 return "OSS_MORE_BUF "
		case -2146881534
		 return "OSS_NEGATIVE_UINTEGER "
		case -2146881533
		 return "OSS_PDU_RANGE "
		case -2146881532
		 return "OSS_MORE_INPUT "
		case -2146881531
		 return "OSS_DATA_ERROR "
		case -2146881530
		 return "OSS_BAD_ARG "
		case -2146881529
		 return "OSS_BAD_VERSION "
		case -2146881528
		 return "OSS_OUT_MEMORY "
		case -2146881527
		 return "OSS_PDU_MISMATCH "
		case -2146881526
		 return "OSS_LIMITED "
		case -2146881525
		 return "OSS_BAD_PTR "
		case -2146881524
		 return "OSS_BAD_TIME "
		case -2146881523
		 return "OSS_INDEFINITE_NOT_SUPPORTED "
		case -2146881522
		 return "OSS_MEM_ERROR "
		case -2146881521
		 return "OSS_BAD_TABLE "
		case -2146881520
		 return "OSS_TOO_LONG "
		case -2146881519
		 return "OSS_CONSTRAINT_VIOLATED "
		case -2146881518
		 return "OSS_FATAL_ERROR "
		case -2146881517
		 return "OSS_ACCESS_SERIALIZATION_ERROR "
		case -2146881516
		 return "OSS_NULL_TBL "
		case -2146881515
		 return "OSS_NULL_FCN "
		case -2146881514
		 return "OSS_BAD_ENCRULES "
		case -2146881513
		 return "OSS_UNAVAIL_ENCRULES "
		case -2146881512
		 return "OSS_CANT_OPEN_TRACE_WINDOW "
		case -2146881511
		 return "OSS_UNIMPLEMENTED "
		case -2146881510
		 return "OSS_OID_DLL_NOT_LINKED "
		case -2146881509
		 return "OSS_CANT_OPEN_TRACE_FILE "
		case -2146881508
		 return "OSS_TRACE_FILE_ALREADY_OPEN "
		case -2146881507
		 return "OSS_TABLE_MISMATCH "
		case -2146881506
		 return "OSS_TYPE_NOT_SUPPORTED "
		case -2146881505
		 return "OSS_REAL_DLL_NOT_LINKED "
		case -2146881504
		 return "OSS_REAL_CODE_NOT_LINKED "
		case -2146881503
		 return "OSS_OUT_OF_RANGE "
		case -2146881502
		 return "OSS_COPIER_DLL_NOT_LINKED "
		case -2146881501
		 return "OSS_CONSTRAINT_DLL_NOT_LINKED "
		case -2146881500
		 return "OSS_COMPARATOR_DLL_NOT_LINKED "
		case -2146881499
		 return "OSS_COMPARATOR_CODE_NOT_LINKED "
		case -2146881498
		 return "OSS_MEM_MGR_DLL_NOT_LINKED "
		case -2146881497
		 return "OSS_PDV_DLL_NOT_LINKED "
		case -2146881496
		 return "OSS_PDV_CODE_NOT_LINKED "
		case -2146881495
		 return "OSS_API_DLL_NOT_LINKED "
		case -2146881494
		 return "OSS_BERDER_DLL_NOT_LINKED "
		case -2146881493
		 return "OSS_PER_DLL_NOT_LINKED "
		case -2146881492
		 return "OSS_OPEN_TYPE_ERROR "
		case -2146881491
		 return "OSS_MUTEX_NOT_CREATED "
		case -2146881490
		 return "OSS_CANT_CLOSE_TRACE_FILE "
		case -2146881280
		 return "CRYPT_E_ASN1_ERROR "
		case -2146881279
		 return "CRYPT_E_ASN1_INTERNAL "
		case -2146881278
		 return "CRYPT_E_ASN1_EOD "
		case -2146881277
		 return "CRYPT_E_ASN1_CORRUPT "
		case -2146881276
		 return "CRYPT_E_ASN1_LARGE "
		case -2146881275
		 return "CRYPT_E_ASN1_CONSTRAINT "
		case -2146881274
		 return "CRYPT_E_ASN1_MEMORY "
		case -2146881273
		 return "CRYPT_E_ASN1_OVERFLOW "
		case -2146881272
		 return "CRYPT_E_ASN1_BADPDU "
		case -2146881271
		 return "CRYPT_E_ASN1_BADARGS "
		case -2146881270
		 return "CRYPT_E_ASN1_BADREAL "
		case -2146881269
		 return "CRYPT_E_ASN1_BADTAG "
		case -2146881268
		 return "CRYPT_E_ASN1_CHOICE "
		case -2146881267
		 return "CRYPT_E_ASN1_RULE "
		case -2146881266
		 return "CRYPT_E_ASN1_UTF8 "
		case -2146881229
		 return "CRYPT_E_ASN1_PDU_TYPE "
		case -2146881228
		 return "CRYPT_E_ASN1_NYI "
		case -2146881023
		 return "CRYPT_E_ASN1_EXTENDED "
		case -2146881022
		 return "CRYPT_E_ASN1_NOEOD "
		case -2146877439
		 return "CERTSRV_E_BAD_REQUESTSUBJECT "
		case -2146877438
		 return "CERTSRV_E_NO_REQUEST "
		case -2146877437
		 return "CERTSRV_E_BAD_REQUESTSTATUS "
		case -2146877436
		 return "CERTSRV_E_PROPERTY_EMPTY "
		case -2146877435
		 return "CERTSRV_E_INVALID_CA_CERTIFICATE "
		case -2146877434
		 return "CERTSRV_E_SERVER_SUSPENDED "
		case -2146877433
		 return "CERTSRV_E_ENCODING_LENGTH "
		case -2146877432
		 return "CERTSRV_E_ROLECONFLICT "
		case -2146877431
		 return "CERTSRV_E_RESTRICTEDOFFICER "
		case -2146877430
		 return "CERTSRV_E_KEY_ARCHIVAL_NOT_CONFIGURED "
		case -2146877429
		 return "CERTSRV_E_NO_VALID_KRA "
		case -2146877428
		 return "CERTSRV_E_BAD_REQUEST_KEY_ARCHIVAL "
		case -2146877427
		 return "CERTSRV_E_NO_CAADMIN_DEFINED "
		case -2146877426
		 return "CERTSRV_E_BAD_RENEWAL_CERT_ATTRIBUTE "
		case -2146877425
		 return "CERTSRV_E_NO_DB_SESSIONS "
		case -2146877424
		 return "CERTSRV_E_ALIGNMENT_FAULT "
		case -2146877423
		 return "CERTSRV_E_ENROLL_DENIED "
		case -2146877422
		 return "CERTSRV_E_TEMPLATE_DENIED "
		case -2146877421
		 return "CERTSRV_E_DOWNLEVEL_DC_SSL_OR_UPGRADE "
		case -2146875392
		 return "CERTSRV_E_UNSUPPORTED_CERT_TYPE "
		case -2146875391
		 return "CERTSRV_E_NO_CERT_TYPE "
		case -2146875390
		 return "CERTSRV_E_TEMPLATE_CONFLICT "
		case -2146875389
		 return "CERTSRV_E_SUBJECT_ALT_NAME_REQUIRED "
		case -2146875388
		 return "CERTSRV_E_ARCHIVED_KEY_REQUIRED "
		case -2146875387
		 return "CERTSRV_E_SMIME_REQUIRED "
		case -2146875386
		 return "CERTSRV_E_BAD_RENEWAL_SUBJECT "
		case -2146875385
		 return "CERTSRV_E_BAD_TEMPLATE_VERSION "
		case -2146875384
		 return "CERTSRV_E_TEMPLATE_POLICY_REQUIRED "
		case -2146875383
		 return "CERTSRV_E_SIGNATURE_POLICY_REQUIRED "
		case -2146875382
		 return "CERTSRV_E_SIGNATURE_COUNT "
		case -2146875381
		 return "CERTSRV_E_SIGNATURE_REJECTED "
		case -2146875380
		 return "CERTSRV_E_ISSUANCE_POLICY_REQUIRED "
		case -2146875379
		 return "CERTSRV_E_SUBJECT_UPN_REQUIRED "
		case -2146875378
		 return "CERTSRV_E_SUBJECT_DIRECTORY_GUID_REQUIRED "
		case -2146875377
		 return "CERTSRV_E_SUBJECT_DNS_REQUIRED "
		case -2146875376
		 return "CERTSRV_E_ARCHIVED_KEY_UNEXPECTED "
		case -2146875375
		 return "CERTSRV_E_KEY_LENGTH "
		case -2146875374
		 return "CERTSRV_E_SUBJECT_EMAIL_REQUIRED "
		case -2146875373
		 return "CERTSRV_E_UNKNOWN_CERT_TYPE "
		case -2146875372
		 return "CERTSRV_E_CERT_TYPE_OVERLAP "
		case -2146873344
		 return "XENROLL_E_KEY_NOT_EXPORTABLE "
		case -2146873343
		 return "XENROLL_E_CANNOT_ADD_ROOT_CERT "
		case -2146873342
		 return "XENROLL_E_RESPONSE_KA_HASH_NOT_FOUND "
		case -2146873341
		 return "XENROLL_E_RESPONSE_UNEXPECTED_KA_HASH "
		case -2146873340
		 return "XENROLL_E_RESPONSE_KA_HASH_MISMATCH "
		case -2146873339
		 return "XENROLL_E_KEYSPEC_SMIME_MISMATCH "
		case -2146869247
		 return "TRUST_E_SYSTEM_ERROR "
		case -2146869246
		 return "TRUST_E_NO_SIGNER_CERT "
		case -2146869245
		 return "TRUST_E_COUNTER_SIGNER "
		case -2146869244
		 return "TRUST_E_CERT_SIGNATURE "
		case -2146869243
		 return "TRUST_E_TIME_STAMP "
		case -2146869232
		 return "TRUST_E_BAD_DIGEST "
		case -2146869223
		 return "TRUST_E_BASIC_CONSTRAINTS "
		case -2146869218
		 return "TRUST_E_FINANCIAL_CRITERIA "
		case -2146865151
		 return "MSSIPOTF_E_OUTOFMEMRANGE "
		case -2146865150
		 return "MSSIPOTF_E_CANTGETOBJECT "
		case -2146865149
		 return "MSSIPOTF_E_NOHEADTABLE "
		case -2146865148
		 return "MSSIPOTF_E_BAD_MAGICNUMBER "
		case -2146865147
		 return "MSSIPOTF_E_BAD_OFFSET_TABLE "
		case -2146865146
		 return "MSSIPOTF_E_TABLE_TAGORDER "
		case -2146865145
		 return "MSSIPOTF_E_TABLE_LONGWORD "
		case -2146865144
		 return "MSSIPOTF_E_BAD_FIRST_TABLE_PLACEMENT "
		case -2146865143
		 return "MSSIPOTF_E_TABLES_OVERLAP "
		case -2146865142
		 return "MSSIPOTF_E_TABLE_PADBYTES "
		case -2146865141
		 return "MSSIPOTF_E_FILETOOSMALL "
		case -2146865140
		 return "MSSIPOTF_E_TABLE_CHECKSUM "
		case -2146865139
		 return "MSSIPOTF_E_FILE_CHECKSUM "
		case -2146865136
		 return "MSSIPOTF_E_FAILED_POLICY "
		case -2146865135
		 return "MSSIPOTF_E_FAILED_HINTS_CHECK "
		case -2146865134
		 return "MSSIPOTF_E_NOT_OPENTYPE "
		case -2146865133
		 return "MSSIPOTF_E_FILE "
		case -2146865132
		 return "MSSIPOTF_E_CRYPT "
		case -2146865131
		 return "MSSIPOTF_E_BADVERSION "
		case -2146865130
		 return "MSSIPOTF_E_DSIG_STRUCTURE "
		case -2146865129
		 return "MSSIPOTF_E_PCONST_CHECK "
		case -2146865128
		 return "MSSIPOTF_E_STRUCTURE "
		case -2146762751
		 return "TRUST_E_PROVIDER_UNKNOWN "
		case -2146762750
		 return "TRUST_E_ACTION_UNKNOWN "
		case -2146762749
		 return "TRUST_E_SUBJECT_FORM_UNKNOWN "
		case -2146762748
		 return "TRUST_E_SUBJECT_NOT_TRUSTED "
		case -2146762747
		 return "DIGSIG_E_ENCODE "
		case -2146762746
		 return "DIGSIG_E_DECODE "
		case -2146762745
		 return "DIGSIG_E_EXTENSIBILITY "
		case -2146762744
		 return "DIGSIG_E_CRYPTO "
		case -2146762743
		 return "PERSIST_E_SIZEDEFINITE "
		case -2146762742
		 return "PERSIST_E_SIZEINDEFINITE "
		case -2146762741
		 return "PERSIST_E_NOTSELFSIZING "
		case -2146762496
		 return "TRUST_E_NOSIGNATURE "
		case -2146762495
		 return "CERT_E_EXPIRED "
		case -2146762494
		 return "CERT_E_VALIDITYPERIODNESTING "
		case -2146762493
		 return "CERT_E_ROLE "
		case -2146762492
		 return "CERT_E_PATHLENCONST "
		case -2146762491
		 return "CERT_E_CRITICAL "
		case -2146762490
		 return "CERT_E_PURPOSE "
		case -2146762489
		 return "CERT_E_ISSUERCHAINING "
		case -2146762488
		 return "CERT_E_MALFORMED "
		case -2146762487
		 return "CERT_E_UNTRUSTEDROOT "
		case -2146762486
		 return "CERT_E_CHAINING "
		case -2146762485
		 return "TRUST_E_FAIL "
		case -2146762484
		 return "CERT_E_REVOKED "
		case -2146762483
		 return "CERT_E_UNTRUSTEDTESTROOT "
		case -2146762482
		 return "CERT_E_REVOCATION_FAILURE "
		case -2146762481
		 return "CERT_E_CN_NO_MATCH "
		case -2146762480
		 return "CERT_E_WRONG_USAGE "
		case -2146762479
		 return "TRUST_E_EXPLICIT_DISTRUST "
		case -2146762478
		 return "CERT_E_UNTRUSTEDCA "
		case -2146762477
		 return "CERT_E_INVALID_POLICY "
		case -2146762476
		 return "CERT_E_INVALID_NAME "
		case -2146500608
		 return "SPAPI_E_EXPECTED_SECTION_NAME "
		case -2146500607
		 return "SPAPI_E_BAD_SECTION_NAME_LINE "
		case -2146500606
		 return "SPAPI_E_SECTION_NAME_TOO_LONG "
		case -2146500605
		 return "SPAPI_E_GENERAL_SYNTAX "
		case -2146500352
		 return "SPAPI_E_WRONG_INF_STYLE "
		case -2146500351
		 return "SPAPI_E_SECTION_NOT_FOUND "
		case -2146500350
		 return "SPAPI_E_LINE_NOT_FOUND "
		case -2146500349
		 return "SPAPI_E_NO_BACKUP "
		case -2146500096
		 return "SPAPI_E_NO_ASSOCIATED_CLASS "
		case -2146500095
		 return "SPAPI_E_CLASS_MISMATCH "
		case -2146500094
		 return "SPAPI_E_DUPLICATE_FOUND "
		case -2146500093
		 return "SPAPI_E_NO_DRIVER_SELECTED "
		case -2146500092
		 return "SPAPI_E_KEY_DOES_NOT_EXIST "
		case -2146500091
		 return "SPAPI_E_INVALID_DEVINST_NAME "
		case -2146500090
		 return "SPAPI_E_INVALID_CLASS "
		case -2146500089
		 return "SPAPI_E_DEVINST_ALREADY_EXISTS "
		case -2146500088
		 return "SPAPI_E_DEVINFO_NOT_REGISTERED "
		case -2146500087
		 return "SPAPI_E_INVALID_REG_PROPERTY "
		case -2146500086
		 return "SPAPI_E_NO_INF "
		case -2146500085
		 return "SPAPI_E_NO_SUCH_DEVINST "
		case -2146500084
		 return "SPAPI_E_CANT_LOAD_CLASS_ICON "
		case -2146500083
		 return "SPAPI_E_INVALID_CLASS_INSTALLER "
		case -2146500082
		 return "SPAPI_E_DI_DO_DEFAULT "
		case -2146500081
		 return "SPAPI_E_DI_NOFILECOPY "
		case -2146500080
		 return "SPAPI_E_INVALID_HWPROFILE "
		case -2146500079
		 return "SPAPI_E_NO_DEVICE_SELECTED "
		case -2146500078
		 return "SPAPI_E_DEVINFO_LIST_LOCKED "
		case -2146500077
		 return "SPAPI_E_DEVINFO_DATA_LOCKED "
		case -2146500076
		 return "SPAPI_E_DI_BAD_PATH "
		case -2146500075
		 return "SPAPI_E_NO_CLASSINSTALL_PARAMS "
		case -2146500074
		 return "SPAPI_E_FILEQUEUE_LOCKED "
		case -2146500073
		 return "SPAPI_E_BAD_SERVICE_INSTALLSECT "
		case -2146500072
		 return "SPAPI_E_NO_CLASS_DRIVER_LIST "
		case -2146500071
		 return "SPAPI_E_NO_ASSOCIATED_SERVICE "
		case -2146500070
		 return "SPAPI_E_NO_DEFAULT_DEVICE_INTERFACE "
		case -2146500069
		 return "SPAPI_E_DEVICE_INTERFACE_ACTIVE "
		case -2146500068
		 return "SPAPI_E_DEVICE_INTERFACE_REMOVED "
		case -2146500067
		 return "SPAPI_E_BAD_INTERFACE_INSTALLSECT "
		case -2146500066
		 return "SPAPI_E_NO_SUCH_INTERFACE_CLASS "
		case -2146500065
		 return "SPAPI_E_INVALID_REFERENCE_STRING "
		case -2146500064
		 return "SPAPI_E_INVALID_MACHINENAME "
		case -2146500063
		 return "SPAPI_E_REMOTE_COMM_FAILURE "
		case -2146500062
		 return "SPAPI_E_MACHINE_UNAVAILABLE "
		case -2146500061
		 return "SPAPI_E_NO_CONFIGMGR_SERVICES "
		case -2146500060
		 return "SPAPI_E_INVALID_PROPPAGE_PROVIDER "
		case -2146500059
		 return "SPAPI_E_NO_SUCH_DEVICE_INTERFACE "
		case -2146500058
		 return "SPAPI_E_DI_POSTPROCESSING_REQUIRED "
		case -2146500057
		 return "SPAPI_E_INVALID_COINSTALLER "
		case -2146500056
		 return "SPAPI_E_NO_COMPAT_DRIVERS "
		case -2146500055
		 return "SPAPI_E_NO_DEVICE_ICON "
		case -2146500054
		 return "SPAPI_E_INVALID_INF_LOGCONFIG "
		case -2146500053
		 return "SPAPI_E_DI_DONT_INSTALL "
		case -2146500052
		 return "SPAPI_E_INVALID_FILTER_DRIVER "
		case -2146500051
		 return "SPAPI_E_NON_WINDOWS_NT_DRIVER "
		case -2146500050
		 return "SPAPI_E_NON_WINDOWS_DRIVER "
		case -2146500049
		 return "SPAPI_E_NO_CATALOG_FOR_OEM_INF "
		case -2146500048
		 return "SPAPI_E_DEVINSTALL_QUEUE_NONNATIVE "
		case -2146500047
		 return "SPAPI_E_NOT_DISABLEABLE "
		case -2146500046
		 return "SPAPI_E_CANT_REMOVE_DEVINST "
		case -2146500045
		 return "SPAPI_E_INVALID_TARGET "
		case -2146500044
		 return "SPAPI_E_DRIVER_NONNATIVE "
		case -2146500043
		 return "SPAPI_E_IN_WOW64 "
		case -2146500042
		 return "SPAPI_E_SET_SYSTEM_RESTORE_POINT "
		case -2146500041
		 return "SPAPI_E_INCORRECTLY_COPIED_INF "
		case -2146500040
		 return "SPAPI_E_SCE_DISABLED "
		case -2146500039
		 return "SPAPI_E_UNKNOWN_EXCEPTION "
		case -2146500038
		 return "SPAPI_E_PNP_REGISTRY_ERROR "
		case -2146500037
		 return "SPAPI_E_REMOTE_REQUEST_UNSUPPORTED "
		case -2146500036
		 return "SPAPI_E_NOT_AN_INSTALLED_OEM_INF "
		case -2146500035
		 return "SPAPI_E_INF_IN_USE_BY_DEVICES "
		case -2146500034
		 return "SPAPI_E_DI_FUNCTION_OBSOLETE "
		case -2146500033
		 return "SPAPI_E_NO_AUTHENTICODE_CATALOG "
		case -2146500032
		 return "SPAPI_E_AUTHENTICODE_DISALLOWED "
		case -2146500031
		 return "SPAPI_E_AUTHENTICODE_TRUSTED_PUBLISHER "
		case -2146500030
		 return "SPAPI_E_AUTHENTICODE_TRUST_NOT_ESTABLISHED "
		case -2146500029
		 return "SPAPI_E_AUTHENTICODE_PUBLISHER_NOT_TRUSTED "
		case -2146500028
		 return "SPAPI_E_SIGNATURE_OSATTRIBUTE_MISMATCH "
		case -2146500027
		 return "SPAPI_E_ONLY_VALIDATE_VIA_AUTHENTICODE "
		case -2146499840
		 return "SPAPI_E_UNRECOVERABLE_STACK_OVERFLOW "
		case -2146496512
		 return "SPAPI_E_ERROR_NOT_INSTALLED "
		case -2146435071
		 return "SCARD_F_INTERNAL_ERROR "
		case -2146435070
		 return "SCARD_E_CANCELLED "
		case -2146435069
		 return "SCARD_E_INVALID_HANDLE "
		case -2146435068
		 return "SCARD_E_INVALID_PARAMETER "
		case -2146435067
		 return "SCARD_E_INVALID_TARGET "
		case -2146435066
		 return "SCARD_E_NO_MEMORY "
		case -2146435065
		 return "SCARD_F_WAITED_TOO_LONG "
		case -2146435064
		 return "SCARD_E_INSUFFICIENT_BUFFER "
		case -2146435063
		 return "SCARD_E_UNKNOWN_READER "
		case -2146435062
		 return "SCARD_E_TIMEOUT "
		case -2146435061
		 return "SCARD_E_SHARING_VIOLATION "
		case -2146435060
		 return "SCARD_E_NO_SMARTCARD "
		case -2146435059
		 return "SCARD_E_UNKNOWN_CARD "
		case -2146435058
		 return "SCARD_E_CANT_DISPOSE "
		case -2146435057
		 return "SCARD_E_PROTO_MISMATCH "
		case -2146435056
		 return "SCARD_E_NOT_READY "
		case -2146435055
		 return "SCARD_E_INVALID_VALUE "
		case -2146435054
		 return "SCARD_E_SYSTEM_CANCELLED "
		case -2146435053
		 return "SCARD_F_COMM_ERROR "
		case -2146435052
		 return "SCARD_F_UNKNOWN_ERROR "
		case -2146435051
		 return "SCARD_E_INVALID_ATR "
		case -2146435050
		 return "SCARD_E_NOT_TRANSACTED "
		case -2146435049
		 return "SCARD_E_READER_UNAVAILABLE "
		case -2146435048
		 return "SCARD_P_SHUTDOWN "
		case -2146435047
		 return "SCARD_E_PCI_TOO_SMALL "
		case -2146435046
		 return "SCARD_E_READER_UNSUPPORTED "
		case -2146435045
		 return "SCARD_E_DUPLICATE_READER "
		case -2146435044
		 return "SCARD_E_CARD_UNSUPPORTED "
		case -2146435043
		 return "SCARD_E_NO_SERVICE "
		case -2146435042
		 return "SCARD_E_SERVICE_STOPPED "
		case -2146435041
		 return "SCARD_E_UNEXPECTED "
		case -2146435040
		 return "SCARD_E_ICC_INSTALLATION "
		case -2146435039
		 return "SCARD_E_ICC_CREATEORDER "
		case -2146435038
		 return "SCARD_E_UNSUPPORTED_FEATURE "
		case -2146435037
		 return "SCARD_E_DIR_NOT_FOUND "
		case -2146435036
		 return "SCARD_E_FILE_NOT_FOUND "
		case -2146435035
		 return "SCARD_E_NO_DIR "
		case -2146435034
		 return "SCARD_E_NO_FILE "
		case -2146435033
		 return "SCARD_E_NO_ACCESS "
		case -2146435032
		 return "SCARD_E_WRITE_TOO_MANY "
		case -2146435031
		 return "SCARD_E_BAD_SEEK "
		case -2146435030
		 return "SCARD_E_INVALID_CHV "
		case -2146435029
		 return "SCARD_E_UNKNOWN_RES_MNG "
		case -2146435028
		 return "SCARD_E_NO_SUCH_CERTIFICATE "
		case -2146435027
		 return "SCARD_E_CERTIFICATE_UNAVAILABLE "
		case -2146435026
		 return "SCARD_E_NO_READERS_AVAILABLE "
		case -2146435025
		 return "SCARD_E_COMM_DATA_LOST "
		case -2146435024
		 return "SCARD_E_NO_KEY_CONTAINER "
		case -2146435023
		 return "SCARD_E_SERVER_TOO_BUSY "
		case -2146434971
		 return "SCARD_W_UNSUPPORTED_CARD "
		case -2146434970
		 return "SCARD_W_UNRESPONSIVE_CARD "
		case -2146434969
		 return "SCARD_W_UNPOWERED_CARD "
		case -2146434968
		 return "SCARD_W_RESET_CARD "
		case -2146434967
		 return "SCARD_W_REMOVED_CARD "
		case -2146434966
		 return "SCARD_W_SECURITY_VIOLATION "
		case -2146434965
		 return "SCARD_W_WRONG_CHV "
		case -2146434964
		 return "SCARD_W_CHV_BLOCKED "
		case -2146434963
		 return "SCARD_W_EOF "
		case -2146434962
		 return "SCARD_W_CANCELLED_BY_USER "
		case -2146434961
		 return "SCARD_W_CARD_NOT_AUTHENTICATED "
		case -2146368511
		 return "COMADMIN_E_OBJECTERRORS "
		case -2146368510
		 return "COMADMIN_E_OBJECTINVALID "
		case -2146368509
		 return "COMADMIN_E_KEYMISSING "
		case -2146368508
		 return "COMADMIN_E_ALREADYINSTALLED "
		case -2146368505
		 return "COMADMIN_E_APP_FILE_WRITEFAIL "
		case -2146368504
		 return "COMADMIN_E_APP_FILE_READFAIL "
		case -2146368503
		 return "COMADMIN_E_APP_FILE_VERSION "
		case -2146368502
		 return "COMADMIN_E_BADPATH "
		case -2146368501
		 return "COMADMIN_E_APPLICATIONEXISTS "
		case -2146368500
		 return "COMADMIN_E_ROLEEXISTS "
		case -2146368499
		 return "COMADMIN_E_CANTCOPYFILE "
		case -2146368497
		 return "COMADMIN_E_NOUSER "
		case -2146368496
		 return "COMADMIN_E_INVALIDUSERIDS "
		case -2146368495
		 return "COMADMIN_E_NOREGISTRYCLSID "
		case -2146368494
		 return "COMADMIN_E_BADREGISTRYPROGID "
		case -2146368493
		 return "COMADMIN_E_AUTHENTICATIONLEVEL "
		case -2146368492
		 return "COMADMIN_E_USERPASSWDNOTVALID "
		case -2146368488
		 return "COMADMIN_E_CLSIDORIIDMISMATCH "
		case -2146368487
		 return "COMADMIN_E_REMOTEINTERFACE "
		case -2146368486
		 return "COMADMIN_E_DLLREGISTERSERVER "
		case -2146368485
		 return "COMADMIN_E_NOSERVERSHARE "
		case -2146368483
		 return "COMADMIN_E_DLLLOADFAILED "
		case -2146368482
		 return "COMADMIN_E_BADREGISTRYLIBID "
		case -2146368481
		 return "COMADMIN_E_APPDIRNOTFOUND "
		case -2146368477
		 return "COMADMIN_E_REGISTRARFAILED "
		case -2146368476
		 return "COMADMIN_E_COMPFILE_DOESNOTEXIST "
		case -2146368475
		 return "COMADMIN_E_COMPFILE_LOADDLLFAIL "
		case -2146368474
		 return "COMADMIN_E_COMPFILE_GETCLASSOBJ "
		case -2146368473
		 return "COMADMIN_E_COMPFILE_CLASSNOTAVAIL "
		case -2146368472
		 return "COMADMIN_E_COMPFILE_BADTLB "
		case -2146368471
		 return "COMADMIN_E_COMPFILE_NOTINSTALLABLE "
		case -2146368470
		 return "COMADMIN_E_NOTCHANGEABLE "
		case -2146368469
		 return "COMADMIN_E_NOTDELETEABLE "
		case -2146368468
		 return "COMADMIN_E_SESSION "
		case -2146368467
		 return "COMADMIN_E_COMP_MOVE_LOCKED "
		case -2146368466
		 return "COMADMIN_E_COMP_MOVE_BAD_DEST "
		case -2146368464
		 return "COMADMIN_E_REGISTERTLB "
		case -2146368461
		 return "COMADMIN_E_SYSTEMAPP "
		case -2146368460
		 return "COMADMIN_E_COMPFILE_NOREGISTRAR "
		case -2146368459
		 return "COMADMIN_E_COREQCOMPINSTALLED "
		case -2146368458
		 return "COMADMIN_E_SERVICENOTINSTALLED "
		case -2146368457
		 return "COMADMIN_E_PROPERTYSAVEFAILED "
		case -2146368456
		 return "COMADMIN_E_OBJECTEXISTS "
		case -2146368455
		 return "COMADMIN_E_COMPONENTEXISTS "
		case -2146368453
		 return "COMADMIN_E_REGFILE_CORRUPT "
		case -2146368452
		 return "COMADMIN_E_PROPERTY_OVERFLOW "
		case -2146368450
		 return "COMADMIN_E_NOTINREGISTRY "
		case -2146368449
		 return "COMADMIN_E_OBJECTNOTPOOLABLE "
		case -2146368442
		 return "COMADMIN_E_APPLID_MATCHES_CLSID "
		case -2146368441
		 return "COMADMIN_E_ROLE_DOES_NOT_EXIST "
		case -2146368440
		 return "COMADMIN_E_START_APP_NEEDS_COMPONENTS "
		case -2146368439
		 return "COMADMIN_E_REQUIRES_DIFFERENT_PLATFORM "
		case -2146368438
		 return "COMADMIN_E_CAN_NOT_EXPORT_APP_PROXY "
		case -2146368437
		 return "COMADMIN_E_CAN_NOT_START_APP "
		case -2146368436
		 return "COMADMIN_E_CAN_NOT_EXPORT_SYS_APP "
		case -2146368435
		 return "COMADMIN_E_CANT_SUBSCRIBE_TO_COMPONENT "
		case -2146368434
		 return "COMADMIN_E_EVENTCLASS_CANT_BE_SUBSCRIBER "
		case -2146368433
		 return "COMADMIN_E_LIB_APP_PROXY_INCOMPATIBLE "
		case -2146368432
		 return "COMADMIN_E_BASE_PARTITION_ONLY "
		case -2146368431
		 return "COMADMIN_E_START_APP_DISABLED "
		case -2146368425
		 return "COMADMIN_E_CAT_DUPLICATE_PARTITION_NAME "
		case -2146368424
		 return "COMADMIN_E_CAT_INVALID_PARTITION_NAME "
		case -2146368423
		 return "COMADMIN_E_CAT_PARTITION_IN_USE "
		case -2146368422
		 return "COMADMIN_E_FILE_PARTITION_DUPLICATE_FILES "
		case -2146368421
		 return "COMADMIN_E_CAT_IMPORTED_COMPONENTS_NOT_ALLOWED "
		case -2146368420
		 return "COMADMIN_E_AMBIGUOUS_APPLICATION_NAME "
		case -2146368419
		 return "COMADMIN_E_AMBIGUOUS_PARTITION_NAME "
		case -2146368398
		 return "COMADMIN_E_REGDB_NOTINITIALIZED "
		case -2146368397
		 return "COMADMIN_E_REGDB_NOTOPEN "
		case -2146368396
		 return "COMADMIN_E_REGDB_SYSTEMERR "
		case -2146368395
		 return "COMADMIN_E_REGDB_ALREADYRUNNING "
		case -2146368384
		 return "COMADMIN_E_MIG_VERSIONNOTSUPPORTED "
		case -2146368383
		 return "COMADMIN_E_MIG_SCHEMANOTFOUND "
		case -2146368382
		 return "COMADMIN_E_CAT_BITNESSMISMATCH "
		case -2146368381
		 return "COMADMIN_E_CAT_UNACCEPTABLEBITNESS "
		case -2146368380
		 return "COMADMIN_E_CAT_WRONGAPPBITNESS "
		case -2146368379
		 return "COMADMIN_E_CAT_PAUSE_RESUME_NOT_SUPPORTED "
		case -2146368378
		 return "COMADMIN_E_CAT_SERVERFAULT "
		case -2146368000
		 return "COMQC_E_APPLICATION_NOT_QUEUED "
		case -2146367999
		 return "COMQC_E_NO_QUEUEABLE_INTERFACES "
		case -2146367998
		 return "COMQC_E_QUEUING_SERVICE_NOT_AVAILABLE "
		case -2146367997
		 return "COMQC_E_NO_IPERSISTSTREAM "
		case -2146367996
		 return "COMQC_E_BAD_MESSAGE "
		case -2146367995
		 return "COMQC_E_UNAUTHENTICATED "
		case -2146367994
		 return "COMQC_E_UNTRUSTED_ENQUEUER "
		case -2146367743
		 return "MSDTC_E_DUPLICATE_RESOURCE "
		case -2146367480
		 return "COMADMIN_E_OBJECT_PARENT_MISSING "
		case -2146367479
		 return "COMADMIN_E_OBJECT_DOES_NOT_EXIST "
		case -2146367478
		 return "COMADMIN_E_APP_NOT_RUNNING "
		case -2146367477
		 return "COMADMIN_E_INVALID_PARTITION "
		case -2146367475
		 return "COMADMIN_E_SVCAPP_NOT_POOLABLE_OR_RECYCLABLE "
		case -2146367474
		 return "COMADMIN_E_USER_IN_SET "
		case -2146367473
		 return "COMADMIN_E_CANTRECYCLELIBRARYAPPS "
		case -2146367471
		 return "COMADMIN_E_CANTRECYCLESERVICEAPPS "
		case -2146367470
		 return "COMADMIN_E_PROCESSALREADYRECYCLED "
		case -2146367469
		 return "COMADMIN_E_PAUSEDPROCESSMAYNOTBERECYCLED "
		case -2146367468
		 return "COMADMIN_E_CANTMAKEINPROCSERVICE "
		case -2146367467
		 return "COMADMIN_E_PROGIDINUSEBYCLSID "
		case -2146367466
		 return "COMADMIN_E_DEFAULT_PARTITION_NOT_IN_SET "
		case -2146367465
		 return "COMADMIN_E_RECYCLEDPROCESSMAYNOTBEPAUSED "
		case -2146367464
		 return "COMADMIN_E_PARTITION_ACCESSDENIED "
		case -2146367463
		 return "COMADMIN_E_PARTITION_MSI_ONLY "
		case -2146367462
		 return "COMADMIN_E_LEGACYCOMPS_NOT_ALLOWED_IN_1_0_FORMAT "
		case -2146367461
		 return "COMADMIN_E_LEGACYCOMPS_NOT_ALLOWED_IN_NONBASE_PARTITIONS "
		case -2146367460
		 return "COMADMIN_E_COMP_MOVE_SOURCE "
		case -2146367459
		 return "COMADMIN_E_COMP_MOVE_DEST "
		case -2146367458
		 return "COMADMIN_E_COMP_MOVE_PRIVATE "
		case -2146367457
		 return "COMADMIN_E_BASEPARTITION_REQUIRED_IN_SET "
		case -2146367456
		 return "COMADMIN_E_CANNOT_ALIAS_EVENTCLASS "
		case -2146367455
		 return "COMADMIN_E_PRIVATE_ACCESSDENIED "
		case -2146367454
		 return "COMADMIN_E_SAFERINVALID "
		case -2146367453
		 return "COMADMIN_E_REGISTRY_ACCESSDENIED "
		case -2146367452
		 return "COMADMIN_E_PARTITIONS_DISABLED "
		case 0
		 return "// Imported HRESULTs from CorError.h "
		case -2146234368
		 return "CEE_E_ENTRYPOINT "
		case -2146234367
		 return "CEE_E_CVTRES_NOT_FOUND "
		case -2146234352
		 return "MSEE_E_LOADLIBFAILED "
		case -2146234351
		 return "MSEE_E_GETPROCFAILED "
		case -2146234350
		 return "MSEE_E_MULTCOPIESLOADED "
		case -2146234348
		 return "COR_E_APPDOMAINUNLOADED "
		case -2146234347
		 return "COR_E_CANNOTUNLOADAPPDOMAIN "
		case -2146234346
		 return "MSEE_E_ASSEMBLYLOADINPROGRESS "
		case -2146234345
		 return "MSEE_E_CANNOTCREATEAPPDOMAIN "
		case -2146234343
		 return "COR_E_FIXUPSINEXE "
		case -2146234342
		 return "COR_E_NO_LOADLIBRARY_ALLOWED "
		case -2146234341
		 return "COR_E_NEWER_RUNTIME "
		case -2146234336
		 return "HOST_E_DEADLOCK "
		case -2146234335
		 return "HOST_E_INTERRUPTED "
		case -2146234334
		 return "HOST_E_INVALIDOPERATION "
		case -2146234333
		 return "HOST_E_CLRNOTAVAILABLE "
		case -2146234332
		 return "HOST_E_TIMEOUT "
		case -2146234331
		 return "HOST_E_NOT_OWNER "
		case -2146234330
		 return "HOST_E_ABANDONED "
		case -2146234329
		 return "HOST_E_EXITPROCESS_THREADABORT "
		case -2146234328
		 return "HOST_E_EXITPROCESS_ADUNLOAD "
		case -2146234327
		 return "HOST_E_EXITPROCESS_TIMEOUT "
		case -2146234326
		 return "HOST_E_EXITPROCESS_OUTOFMEMORY "
		case -2146234325
		 return "HOST_E_EXITPROCESS_STACKOVERFLOW "
		case -2146234311
		 return "COR_E_MODULE_HASH_CHECK_FAILED "
		case -2146234304
		 return "FUSION_E_REF_DEF_MISMATCH "
		case -2146234303
		 return "FUSION_E_INVALID_PRIVATE_ASM_LOCATION "
		case -2146234302
		 return "FUSION_E_ASM_MODULE_MISSING "
		case -2146234301
		 return "FUSION_E_UNEXPECTED_MODULE_FOUND "
		case -2146234300
		 return "FUSION_E_PRIVATE_ASM_DISALLOWED "
		case -2146234299
		 return "FUSION_E_SIGNATURE_CHECK_FAILED "
		case -2146234298
		 return "FUSION_E_DATABASE_ERROR "
		case -2146234297
		 return "FUSION_E_INVALID_NAME "
		case -2146234296
		 return "FUSION_E_CODE_DOWNLOAD_DISABLED "
		case -2146234295
		 return "FUSION_E_UNINSTALL_DISALLOWED "
		case -2146234288
		 return "FUSION_E_HOST_GAC_ASM_MISMATCH "
		case -2146234112
		 return "CLDB_E_FILE_BADREAD "
		case -2146234111
		 return "CLDB_E_FILE_BADWRITE "
		case -2146234109
		 return "CLDB_E_FILE_READONLY "
		case -2146234107
		 return "CLDB_E_NAME_ERROR "
		case 1249542
		 return "CLDB_S_TRUNCATION "
		case -2146234106
		 return "CLDB_E_TRUNCATION "
		case -2146234105
		 return "CLDB_E_FILE_OLDVER "
		case -2146234104
		 return "CLDB_E_RELOCATED "
		case 1249545
		 return "CLDB_S_NULL "
		case -2146234102
		 return "CLDB_E_SMDUPLICATE "
		case -2146234101
		 return "CLDB_E_NO_DATA "
		case -2146234100
		 return "CLDB_E_READONLY "
		case -2146234099
		 return "CLDB_E_INCOMPATIBLE "
		case -2146234098
		 return "CLDB_E_FILE_CORRUPT "
		case -2146234097
		 return "CLDB_E_SCHEMA_VERNOTFOUND "
		case -2146234096
		 return "CLDB_E_BADUPDATEMODE "
		case -2146234079
		 return "CLDB_E_INDEX_NONULLKEYS "
		case -2146234078
		 return "CLDB_E_INDEX_DUPLICATE "
		case -2146234077
		 return "CLDB_E_INDEX_BADTYPE "
		case -2146234076
		 return "CLDB_E_INDEX_NOTFOUND "
		case 1249573
		 return "CLDB_S_INDEX_TABLESCANREQUIRED "
		case -2146234064
		 return "CLDB_E_RECORD_NOTFOUND "
		case -2146234063
		 return "CLDB_E_RECORD_OVERFLOW "
		case -2146234062
		 return "CLDB_E_RECORD_DUPLICATE "
		case -2146234061
		 return "CLDB_E_RECORD_PKREQUIRED "
		case -2146234060
		 return "CLDB_E_RECORD_DELETED "
		case -2146234059
		 return "CLDB_E_RECORD_OUTOFORDER "
		case -2146234048
		 return "CLDB_E_COLUMN_OVERFLOW "
		case -2146234047
		 return "CLDB_E_COLUMN_READONLY "
		case -2146234046
		 return "CLDB_E_COLUMN_SPECIALCOL "
		case -2146234045
		 return "CLDB_E_COLUMN_PKNONULLS "
		case -2146234032
		 return "CLDB_E_TABLE_CANTDROP "
		case -2146234031
		 return "CLDB_E_OBJECT_NOTFOUND "
		case -2146234030
		 return "CLDB_E_OBJECT_COLNOTFOUND "
		case -2146234029
		 return "CLDB_E_VECTOR_BADINDEX "
		case -2146234028
		 return "CLDB_E_TOO_BIG "
		case -2146234017
		 return "META_E_INVALID_TOKEN_TYPE "
		case -2146234016
		 return "TLBX_E_INVALID_TYPEINFO "
		case -2146234015
		 return "TLBX_E_INVALID_TYPEINFO_UNNAMED "
		case -2146234014
		 return "TLBX_E_CTX_NESTED "
		case -2146234013
		 return "TLBX_E_ERROR_MESSAGE "
		case -2146234012
		 return "TLBX_E_CANT_SAVE "
		case -2146234011
		 return "TLBX_W_LIBNOTREGISTERED "
		case -2146234010
		 return "TLBX_E_CANTLOADLIBRARY "
		case -2146234009
		 return "TLBX_E_BAD_VT_TYPE "
		case -2146234008
		 return "TLBX_E_NO_MSCOREE_TLB "
		case -2146234007
		 return "TLBX_E_BAD_MSCOREE_TLB "
		case -2146234006
		 return "TLBX_E_TLB_EXCEPTION "
		case -2146234005
		 return "TLBX_E_MULTIPLE_LCIDS "
		case 1249644
		 return "TLBX_I_TYPEINFO_IMPORTED "
		case -2146234003
		 return "TLBX_E_AMBIGUOUS_RETURN "
		case -2146234002
		 return "TLBX_E_DUPLICATE_TYPE_NAME "
		case 1249647
		 return "TLBX_I_USEIUNKNOWN "
		case 1249648
		 return "TLBX_I_UNCONVERTABLE_ARGS "
		case 1249649
		 return "TLBX_I_UNCONVERTABLE_FIELD "
		case -2146233998
		 return "TLBX_I_NONSEQUENTIALSTRUCT "
		case 1249651
		 return "TLBX_W_WARNING_MESSAGE "
		case -2146233996
		 return "TLBX_I_RESOLVEREFFAILED "
		case -2146233995
		 return "TLBX_E_ASANY "
		case -2146233994
		 return "TLBX_E_INVALIDLCIDPARAM "
		case -2146233993
		 return "TLBX_E_LCIDONDISPONLYITF "
		case -2146233992
		 return "TLBX_E_NONPUBLIC_FIELD "
		case 1249657
		 return "TLBX_I_TYPE_EXPORTED "
		case 1249658
		 return "TLBX_I_DUPLICATE_DISPID "
		case -2146233989
		 return "TLBX_E_BAD_NAMES "
		case 1249660
		 return "TLBX_I_REF_TYPE_AS_STRUCT "
		case -2146233987
		 return "TLBX_E_GENERICINST_SIGNATURE "
		case -2146233986
		 return "TLBX_E_GENERICPAR_SIGNATURE "
		case 1249663
		 return "TLBX_I_GENERIC_TYPE "
		case -2146233984
		 return "META_E_DUPLICATE "
		case -2146233983
		 return "META_E_GUID_REQUIRED "
		case -2146233982
		 return "META_E_TYPEDEF_MISMATCH "
		case -2146233981
		 return "META_E_MERGE_COLLISION "
		case 1249668
		 return "TLBX_W_NON_INTEGRAL_CA_TYPE "
		case 1249669
		 return "TLBX_W_IENUM_CA_ON_IUNK "
		case -2146233978
		 return "TLBX_E_NO_SAFEHANDLE_ARRAYS "
		case -2146233977
		 return "META_E_METHD_NOT_FOUND "
		case -2146233976
		 return "META_E_FIELD_NOT_FOUND "
		case 1249673
		 return "META_S_PARAM_MISMATCH "
		case -2146233975
		 return "META_E_PARAM_MISMATCH "
		case -2146233974
		 return "META_E_BADMETADATA "
		case -2146233973
		 return "META_E_INTFCEIMPL_NOT_FOUND "
		case -2146233972
		 return "TLBX_E_NO_CRITICALHANDLE_ARRAYS "
		case -2146233971
		 return "META_E_CLASS_LAYOUT_INCONSISTENT "
		case -2146233970
		 return "META_E_FIELD_MARSHAL_NOT_FOUND "
		case -2146233969
		 return "META_E_METHODSEM_NOT_FOUND "
		case -2146233968
		 return "META_E_EVENT_NOT_FOUND "
		case -2146233967
		 return "META_E_PROP_NOT_FOUND "
		case -2146233966
		 return "META_E_BAD_SIGNATURE "
		case -2146233965
		 return "META_E_BAD_INPUT_PARAMETER "
		case -2146233964
		 return "META_E_METHDIMPL_INCONSISTENT "
		case -2146233963
		 return "META_E_MD_INCONSISTENCY "
		case -2146233962
		 return "META_E_CANNOTRESOLVETYPEREF "
		case 1249687
		 return "META_S_DUPLICATE "
		case -2146233960
		 return "META_E_STRINGSPACE_FULL "
		case -2146233959
		 return "META_E_UNEXPECTED_REMAP "
		case -2146233958
		 return "META_E_HAS_UNMARKALL "
		case -2146233957
		 return "META_E_MUST_CALL_UNMARKALL "
		case -2146233956
		 return "META_E_GENERICPARAM_INCONSISTENT "
		case -2146233955
		 return "META_E_EVENT_COUNTS "
		case -2146233954
		 return "META_E_PROPERTY_COUNTS "
		case -2146233953
		 return "META_E_TYPEDEF_MISSING "
		case -2146233952
		 return "TLBX_E_CANT_LOAD_MODULE "
		case -2146233951
		 return "TLBX_E_CANT_LOAD_CLASS "
		case -2146233950
		 return "TLBX_E_NULL_MODULE "
		case -2146233949
		 return "TLBX_E_NO_CLSID_KEY "
		case -2146233948
		 return "TLBX_E_CIRCULAR_EXPORT "
		case -2146233947
		 return "TLBX_E_CIRCULAR_IMPORT "
		case -2146233946
		 return "TLBX_E_BAD_NATIVETYPE "
		case -2146233945
		 return "TLBX_E_BAD_VTABLE "
		case -2146233944
		 return "TLBX_E_CRM_NON_STATIC "
		case -2146233943
		 return "TLBX_E_CRM_INVALID_SIG "
		case -2146233942
		 return "TLBX_E_CLASS_LOAD_EXCEPTION "
		case -2146233941
		 return "TLBX_E_UNKNOWN_SIGNATURE "
		case -2146233940
		 return "TLBX_E_REFERENCED_TYPELIB "
		case 1249708
		 return "TLBX_S_REFERENCED_TYPELIB "
		case -2146233939
		 return "TLBX_E_INVALID_NAMESPACE "
		case -2146233938
		 return "TLBX_E_LAYOUT_ERROR "
		case -2146233937
		 return "TLBX_E_NOTIUNKNOWN "
		case -2146233936
		 return "TLBX_E_NONVISIBLEVALUECLASS "
		case -2146233935
		 return "TLBX_E_LPTSTR_NOT_ALLOWED "
		case -2146233934
		 return "TLBX_E_AUTO_CS_NOT_ALLOWED "
		case 1249715
		 return "TLBX_S_NOSTDINTERFACE "
		case 1249716
		 return "TLBX_S_DUPLICATE_DISPID "
		case -2146233931
		 return "TLBX_E_ENUM_VALUE_INVALID "
		case -2146233930
		 return "TLBX_E_DUPLICATE_IID "
		case -2146233929
		 return "TLBX_E_NO_NESTED_ARRAYS "
		case -2146233928
		 return "TLBX_E_PARAM_ERROR_NAMED "
		case -2146233927
		 return "TLBX_E_PARAM_ERROR_UNNAMED "
		case -2146233926
		 return "TLBX_E_AGNOST_SIGNATURE "
		case -2146233925
		 return "TLBX_E_CONVERT_FAIL "
		case -2146233924
		 return "TLBX_W_DUAL_NOT_DISPATCH "
		case -2146233923
		 return "TLBX_E_BAD_SIGNATURE "
		case -2146233922
		 return "TLBX_E_ARRAY_NEEDS_NT_FIXED "
		case -2146233921
		 return "TLBX_E_CLASS_NEEDS_NT_INTF "
		case -2146233920
		 return "META_E_CA_INVALID_TARGET "
		case -2146233919
		 return "META_E_CA_INVALID_VALUE "
		case -2146233918
		 return "META_E_CA_INVALID_BLOB "
		case -2146233917
		 return "META_E_CA_REPEATED_ARG "
		case -2146233916
		 return "META_E_CA_UNKNOWN_ARGUMENT "
		case -2146233915
		 return "META_E_CA_VARIANT_NYI "
		case -2146233914
		 return "META_E_CA_ARRAY_NYI "
		case -2146233913
		 return "META_E_CA_UNEXPECTED_TYPE "
		case -2146233912
		 return "META_E_CA_INVALID_ARGTYPE "
		case -2146233911
		 return "META_E_CA_INVALID_ARG_FOR_TYPE "
		case -2146233910
		 return "META_E_CA_INVALID_UUID "
		case -2146233909
		 return "META_E_CA_INVALID_MARSHALAS_FIELDS "
		case -2146233908
		 return "META_E_CA_NT_FIELDONLY "
		case -2146233907
		 return "META_E_CA_NEGATIVE_PARAMINDEX "
		case -2146233906
		 return "META_E_CA_NEGATIVE_MULTIPLIER "
		case -2146233905
		 return "META_E_CA_NEGATIVE_CONSTSIZE "
		case -2146233904
		 return "META_E_CA_FIXEDSTR_SIZE_REQUIRED "
		case -2146233903
		 return "META_E_CA_CUSTMARSH_TYPE_REQUIRED "
		case -2146233902
		 return "META_E_CA_FILENAME_REQUIRED "
		case -2146233901
		 return "TLBX_W_NO_PROPS_IN_EVENTS "
		case -2146233900
		 return "META_E_NOT_IN_ENC_MODE "
		case 1249749
		 return "TLBX_W_ENUM_VALUE_TOOBIG "
		case -2146233898
		 return "META_E_METHOD_COUNTS "
		case -2146233897
		 return "META_E_FIELD_COUNTS "
		case -2146233896
		 return "META_E_PARAM_COUNTS "
		case 1249753
		 return "TLBX_W_EXPORTING_AUTO_LAYOUT "
		case -2146233894
		 return "TLBX_E_TYPED_REF "
		case 1249755
		 return "TLBX_W_DEFAULT_INTF_NOT_VISIBLE "
		case 1249758
		 return "TLBX_W_BAD_SAFEARRAYFIELD_NO_ELEMENTVT "
		case 1249759
		 return "TLBX_W_LAYOUTCLASS_AS_INTERFACE "
		case 1249760
		 return "TLBX_I_GENERIC_BASE_TYPE "
		case -2146233887
		 return "TLBX_E_BITNESS_MISMATCH "
		case 1249792
		 return "VLDTR_S_WRN "
		case 1249793
		 return "VLDTR_S_ERR "
		case 1249794
		 return "VLDTR_S_WRNERR "
		case -2146233853
		 return "VLDTR_E_RID_OUTOFRANGE "
		case -2146233852
		 return "VLDTR_E_CDTKN_OUTOFRANGE "
		case -2146233851
		 return "VLDTR_E_CDRID_OUTOFRANGE "
		case -2146233850
		 return "VLDTR_E_STRING_INVALID "
		case -2146233849
		 return "VLDTR_E_GUID_INVALID "
		case -2146233848
		 return "VLDTR_E_BLOB_INVALID "
		case -2146233847
		 return "VLDTR_E_MOD_MULTI "
		case -2146233846
		 return "VLDTR_E_MOD_NULLMVID "
		case -2146233845
		 return "VLDTR_E_TR_NAMENULL "
		case -2146233844
		 return "VLDTR_E_TR_DUP "
		case -2146233843
		 return "VLDTR_E_TD_NAMENULL "
		case -2146233842
		 return "VLDTR_E_TD_DUPNAME "
		case -2146233841
		 return "VLDTR_E_TD_DUPGUID "
		case -2146233840
		 return "VLDTR_E_TD_NOTIFACEOBJEXTNULL "
		case -2146233839
		 return "VLDTR_E_TD_OBJEXTENDSNONNULL "
		case -2146233838
		 return "VLDTR_E_TD_EXTENDSSEALED "
		case -2146233837
		 return "VLDTR_E_TD_DLTNORTSPCL "
		case -2146233836
		 return "VLDTR_E_TD_RTSPCLNOTDLT "
		case -2146233835
		 return "VLDTR_E_MI_DECLPRIV "
		case -2146233834
		 return "VLDTR_E_AS_BADNAME "
		case -2146233833
		 return "VLDTR_E_FILE_SYSNAME "
		case -2146233832
		 return "VLDTR_E_MI_BODYSTATIC "
		case -2146233831
		 return "VLDTR_E_TD_IFACENOTABS "
		case -2146233830
		 return "VLDTR_E_TD_IFACEPARNOTNIL "
		case -2146233829
		 return "VLDTR_E_TD_IFACEGUIDNULL "
		case -2146233828
		 return "VLDTR_E_MI_DECLFINAL "
		case -2146233827
		 return "VLDTR_E_TD_VTNOTSEAL "
		case -2146233826
		 return "VLDTR_E_PD_BADFLAGS "
		case -2146233825
		 return "VLDTR_E_IFACE_DUP "
		case -2146233824
		 return "VLDTR_E_MR_NAMENULL "
		case -2146233823
		 return "VLDTR_E_MR_VTBLNAME "
		case -2146233822
		 return "VLDTR_E_MR_DELNAME "
		case -2146233821
		 return "VLDTR_E_MR_PARNIL "
		case -2146233820
		 return "VLDTR_E_MR_BADCALLINGCONV "
		case -2146233819
		 return "VLDTR_E_MR_NOTVARARG "
		case -2146233818
		 return "VLDTR_E_MR_NAMEDIFF "
		case -2146233817
		 return "VLDTR_E_MR_SIGDIFF "
		case -2146233816
		 return "VLDTR_E_MR_DUP "
		case -2146233815
		 return "VLDTR_E_CL_TDAUTO "
		case -2146233814
		 return "VLDTR_E_CL_BADPCKSZ "
		case -2146233813
		 return "VLDTR_E_CL_DUP "
		case -2146233812
		 return "VLDTR_E_FL_BADOFFSET "
		case -2146233811
		 return "VLDTR_E_FL_TDNIL "
		case -2146233810
		 return "VLDTR_E_FL_NOCL "
		case -2146233809
		 return "VLDTR_E_FL_TDNOTEXPLCT "
		case -2146233808
		 return "VLDTR_E_FL_FLDSTATIC "
		case -2146233807
		 return "VLDTR_E_FL_DUP "
		case -2146233806
		 return "VLDTR_E_MODREF_NAMENULL "
		case -2146233805
		 return "VLDTR_E_MODREF_DUP "
		case -2146233804
		 return "VLDTR_E_TR_BADSCOPE "
		case -2146233803
		 return "VLDTR_E_TD_NESTEDNOENCL "
		case -2146233802
		 return "VLDTR_E_TD_EXTTRRES "
		case -2146233801
		 return "VLDTR_E_SIGNULL "
		case -2146233800
		 return "VLDTR_E_SIGNODATA "
		case -2146233799
		 return "VLDTR_E_MD_BADCALLINGCONV "
		case -2146233798
		 return "VLDTR_E_MD_THISSTATIC "
		case -2146233797
		 return "VLDTR_E_MD_NOTTHISNOTSTATIC "
		case -2146233796
		 return "VLDTR_E_MD_NOARGCNT "
		case -2146233795
		 return "VLDTR_E_SIG_MISSELTYPE "
		case -2146233794
		 return "VLDTR_E_SIG_MISSTKN "
		case -2146233793
		 return "VLDTR_E_SIG_TKNBAD "
		case -2146233792
		 return "VLDTR_E_SIG_MISSFPTR "
		case -2146233791
		 return "VLDTR_E_SIG_MISSFPTRARGCNT "
		case -2146233790
		 return "VLDTR_E_SIG_MISSRANK "
		case -2146233789
		 return "VLDTR_E_SIG_MISSNSIZE "
		case -2146233788
		 return "VLDTR_E_SIG_MISSSIZE "
		case -2146233787
		 return "VLDTR_E_SIG_MISSNLBND "
		case -2146233786
		 return "VLDTR_E_SIG_MISSLBND "
		case -2146233785
		 return "VLDTR_E_SIG_BADELTYPE "
		case -2146233784
		 return "VLDTR_E_SIG_MISSVASIZE "
		case -2146233783
		 return "VLDTR_E_FD_BADCALLINGCONV "
		case -2146233782
		 return "VLDTR_E_MD_NAMENULL "
		case -2146233781
		 return "VLDTR_E_MD_PARNIL "
		case -2146233780
		 return "VLDTR_E_MD_DUP "
		case -2146233779
		 return "VLDTR_E_FD_NAMENULL "
		case -2146233778
		 return "VLDTR_E_FD_PARNIL "
		case -2146233777
		 return "VLDTR_E_FD_DUP "
		case -2146233776
		 return "VLDTR_E_AS_MULTI "
		case -2146233775
		 return "VLDTR_E_AS_NAMENULL "
		case -2146233774
		 return "VLDTR_E_SIG_TOKTYPEMISMATCH "
		case -2146233773
		 return "VLDTR_E_CL_TDINTF "
		case -2146233772
		 return "VLDTR_E_ASOS_OSPLTFRMIDINVAL "
		case -2146233771
		 return "VLDTR_E_AR_NAMENULL "
		case -2146233770
		 return "VLDTR_E_TD_ENCLNOTNESTED "
		case -2146233769
		 return "VLDTR_E_AROS_OSPLTFRMIDINVAL "
		case -2146233768
		 return "VLDTR_E_FILE_NAMENULL "
		case -2146233767
		 return "VLDTR_E_CT_NAMENULL "
		case -2146233766
		 return "VLDTR_E_TD_EXTENDSCHILD "
		case -2146233765
		 return "VLDTR_E_MAR_NAMENULL "
		case -2146233764
		 return "VLDTR_E_FILE_DUP "
		case -2146233763
		 return "VLDTR_E_FILE_NAMEFULLQLFD "
		case -2146233762
		 return "VLDTR_E_CT_DUP "
		case -2146233761
		 return "VLDTR_E_MAR_DUP "
		case -2146233760
		 return "VLDTR_E_MAR_NOTPUBPRIV "
		case -2146233759
		 return "VLDTR_E_TD_ENUMNOVALUE "
		case -2146233758
		 return "VLDTR_E_TD_ENUMVALSTATIC "
		case -2146233757
		 return "VLDTR_E_TD_ENUMVALNOTSN "
		case -2146233756
		 return "VLDTR_E_TD_ENUMFLDNOTST "
		case -2146233755
		 return "VLDTR_E_TD_ENUMFLDNOTLIT "
		case -2146233754
		 return "VLDTR_E_TD_ENUMNOLITFLDS "
		case -2146233753
		 return "VLDTR_E_TD_ENUMFLDSIGMISMATCH "
		case -2146233752
		 return "VLDTR_E_TD_ENUMVALNOT1ST "
		case -2146233751
		 return "VLDTR_E_FD_NOTVALUERTSN "
		case -2146233750
		 return "VLDTR_E_FD_VALUEPARNOTENUM "
		case -2146233749
		 return "VLDTR_E_FD_INSTINIFACE "
		case -2146233748
		 return "VLDTR_E_FD_NOTPUBINIFACE "
		case -2146233747
		 return "VLDTR_E_FMD_GLOBALNOTPUBPRIVSC "
		case -2146233746
		 return "VLDTR_E_FMD_GLOBALNOTSTATIC "
		case -2146233745
		 return "VLDTR_E_FD_GLOBALNORVA "
		case -2146233744
		 return "VLDTR_E_MD_CTORZERORVA "
		case -2146233743
		 return "VLDTR_E_FD_MARKEDNOMARSHAL "
		case -2146233742
		 return "VLDTR_E_FD_MARSHALNOTMARKED "
		case -2146233741
		 return "VLDTR_E_FD_MARKEDNODEFLT "
		case -2146233740
		 return "VLDTR_E_FD_DEFLTNOTMARKED "
		case -2146233739
		 return "VLDTR_E_FMD_MARKEDNOSECUR "
		case -2146233738
		 return "VLDTR_E_FMD_SECURNOTMARKED "
		case -2146233737
		 return "VLDTR_E_FMD_PINVOKENOTSTATIC "
		case -2146233736
		 return "VLDTR_E_FMD_MARKEDNOPINVOKE "
		case -2146233735
		 return "VLDTR_E_FMD_PINVOKENOTMARKED "
		case -2146233734
		 return "VLDTR_E_FMD_BADIMPLMAP "
		case -2146233733
		 return "VLDTR_E_IMAP_BADMODREF "
		case -2146233732
		 return "VLDTR_E_IMAP_BADMEMBER "
		case -2146233731
		 return "VLDTR_E_IMAP_BADIMPORTNAME "
		case -2146233730
		 return "VLDTR_E_IMAP_BADCALLCONV "
		case -2146233729
		 return "VLDTR_E_FMD_BADACCESSFLAG "
		case -2146233728
		 return "VLDTR_E_FD_INITONLYANDLITERAL "
		case -2146233727
		 return "VLDTR_E_FD_LITERALNOTSTATIC "
		case -2146233726
		 return "VLDTR_E_FMD_RTSNNOTSN "
		case -2146233725
		 return "VLDTR_E_MD_ABSTPARNOTABST "
		case -2146233724
		 return "VLDTR_E_MD_NOTSTATABSTININTF "
		case -2146233723
		 return "VLDTR_E_MD_NOTPUBININTF "
		case -2146233722
		 return "VLDTR_E_MD_CTORININTF "
		case -2146233721
		 return "VLDTR_E_MD_GLOBALCTORCCTOR "
		case -2146233720
		 return "VLDTR_E_MD_CTORSTATIC "
		case -2146233719
		 return "VLDTR_E_MD_CTORNOTSNRTSN "
		case -2146233718
		 return "VLDTR_E_MD_CTORVIRT "
		case -2146233717
		 return "VLDTR_E_MD_CTORABST "
		case -2146233716
		 return "VLDTR_E_MD_CCTORNOTSTATIC "
		case -2146233715
		 return "VLDTR_E_MD_ZERORVA "
		case -2146233714
		 return "VLDTR_E_MD_FINNOTVIRT "
		case -2146233713
		 return "VLDTR_E_MD_STATANDFINORVIRT "
		case -2146233712
		 return "VLDTR_E_MD_ABSTANDFINAL "
		case -2146233711
		 return "VLDTR_E_MD_ABSTANDIMPL "
		case -2146233710
		 return "VLDTR_E_MD_ABSTANDPINVOKE "
		case -2146233709
		 return "VLDTR_E_MD_ABSTNOTVIRT "
		case -2146233708
		 return "VLDTR_E_MD_NOTABSTNOTIMPL "
		case -2146233707
		 return "VLDTR_E_MD_NOTABSTBADFLAGSRVA "
		case -2146233706
		 return "VLDTR_E_MD_PRIVSCOPENORVA "
		case -2146233705
		 return "VLDTR_E_MD_GLOBALABSTORVIRT "
		case -2146233704
		 return "VLDTR_E_SIG_LONGFORM "
		case -2146233703
		 return "VLDTR_E_MD_MULTIPLESEMANTICS "
		case -2146233702
		 return "VLDTR_E_MD_INVALIDSEMANTICS "
		case -2146233701
		 return "VLDTR_E_MD_SEMANTICSNOTEXIST "
		case -2146233700
		 return "VLDTR_E_MI_DECLNOTVIRT "
		case -2146233699
		 return "VLDTR_E_FMD_GLOBALITEM "
		case -2146233698
		 return "VLDTR_E_MD_MULTSEMANTICFLAGS "
		case -2146233697
		 return "VLDTR_E_MD_NOSEMANTICFLAGS "
		case -2146233696
		 return "VLDTR_E_FD_FLDINIFACE "
		case -2146233695
		 return "VLDTR_E_AS_HASHALGID "
		case -2146233694
		 return "VLDTR_E_AS_PROCID "
		case -2146233693
		 return "VLDTR_E_AR_PROCID "
		case -2146233692
		 return "VLDTR_E_CN_PARENTRANGE "
		case -2146233691
		 return "VLDTR_E_AS_BADFLAGS "
		case -2146233690
		 return "VLDTR_E_TR_HASTYPEDEF "
		case -2146233689
		 return "VLDTR_E_IFACE_BADIMPL "
		case -2146233688
		 return "VLDTR_E_IFACE_BADIFACE "
		case -2146233687
		 return "VLDTR_E_TD_SECURNOTMARKED "
		case -2146233686
		 return "VLDTR_E_TD_MARKEDNOSECUR "
		case -2146233685
		 return "VLDTR_E_MD_CCTORHASARGS "
		case -2146233684
		 return "VLDTR_E_CT_BADIMPL "
		case -2146233683
		 return "VLDTR_E_MI_ALIENBODY "
		case -2146233682
		 return "VLDTR_E_MD_CCTORCALLCONV "
		case -2146233681
		 return "VLDTR_E_MI_BADCLASS "
		case -2146233680
		 return "VLDTR_E_MI_CLASSISINTF "
		case -2146233679
		 return "VLDTR_E_MI_BADDECL "
		case -2146233678
		 return "VLDTR_E_MI_BADBODY "
		case -2146233677
		 return "VLDTR_E_MI_DUP "
		case -2146233676
		 return "VLDTR_E_FD_BADPARENT "
		case -2146233675
		 return "VLDTR_E_MD_PARAMOUTOFSEQ "
		case -2146233674
		 return "VLDTR_E_MD_PARASEQTOOBIG "
		case -2146233673
		 return "VLDTR_E_MD_PARMMARKEDNOMARSHAL "
		case -2146233672
		 return "VLDTR_E_MD_PARMMARSHALNOTMARKED "
		case -2146233670
		 return "VLDTR_E_MD_PARMMARKEDNODEFLT "
		case -2146233669
		 return "VLDTR_E_MD_PARMDEFLTNOTMARKED "
		case -2146233668
		 return "VLDTR_E_PR_BADSCOPE "
		case -2146233667
		 return "VLDTR_E_PR_NONAME "
		case -2146233666
		 return "VLDTR_E_PR_NOSIG "
		case -2146233665
		 return "VLDTR_E_PR_DUP "
		case -2146233664
		 return "VLDTR_E_PR_BADCALLINGCONV "
		case -2146233663
		 return "VLDTR_E_PR_MARKEDNODEFLT "
		case -2146233662
		 return "VLDTR_E_PR_DEFLTNOTMARKED "
		case -2146233661
		 return "VLDTR_E_PR_BADSEMANTICS "
		case -2146233660
		 return "VLDTR_E_PR_BADMETHOD "
		case -2146233659
		 return "VLDTR_E_PR_ALIENMETHOD "
		case -2146233658
		 return "VLDTR_E_CN_BLOBNOTNULL "
		case -2146233657
		 return "VLDTR_E_CN_BLOBNULL "
		case -2146233656
		 return "VLDTR_E_EV_BADSCOPE "
		case -2146233654
		 return "VLDTR_E_EV_NONAME "
		case -2146233653
		 return "VLDTR_E_EV_DUP "
		case -2146233652
		 return "VLDTR_E_EV_BADEVTYPE "
		case -2146233651
		 return "VLDTR_E_EV_EVTYPENOTCLASS "
		case -2146233650
		 return "VLDTR_E_EV_BADSEMANTICS "
		case -2146233649
		 return "VLDTR_E_EV_BADMETHOD "
		case -2146233648
		 return "VLDTR_E_EV_ALIENMETHOD "
		case -2146233647
		 return "VLDTR_E_EV_NOADDON "
		case -2146233646
		 return "VLDTR_E_EV_NOREMOVEON "
		case -2146233645
		 return "VLDTR_E_CT_DUPTDNAME "
		case -2146233644
		 return "VLDTR_E_MAR_BADOFFSET "
		case -2146233643
		 return "VLDTR_E_DS_BADOWNER "
		case -2146233642
		 return "VLDTR_E_DS_BADFLAGS "
		case -2146233641
		 return "VLDTR_E_DS_NOBLOB "
		case -2146233640
		 return "VLDTR_E_MAR_BADIMPL "
		case -2146233638
		 return "VLDTR_E_MR_VARARGCALLINGCONV "
		case -2146233637
		 return "VLDTR_E_MD_CTORNOTVOID "
		case -2146233636
		 return "VLDTR_E_EV_FIRENOTVOID "
		case -2146233635
		 return "VLDTR_E_AS_BADLOCALE "
		case -2146233634
		 return "VLDTR_E_CN_PARENTTYPE "
		case -2146233633
		 return "VLDTR_E_SIG_SENTINMETHODDEF "
		case -2146233632
		 return "VLDTR_E_SIG_SENTMUSTVARARG "
		case -2146233631
		 return "VLDTR_E_SIG_MULTSENTINELS "
		case -2146233630
		 return "VLDTR_E_SIG_LASTSENTINEL "
		case -2146233629
		 return "VLDTR_E_SIG_MISSARG "
		case -2146233628
		 return "VLDTR_E_SIG_BYREFINFIELD "
		case -2146233627
		 return "VLDTR_E_MD_SYNCMETHODINVTYPE "
		case -2146233626
		 return "VLDTR_E_TD_NAMETOOLONG "
		case -2146233625
		 return "VLDTR_E_AS_PROCDUP "
		case -2146233624
		 return "VLDTR_E_ASOS_DUP "
		case -2146233623
		 return "VLDTR_E_MAR_BADFLAGS "
		case -2146233622
		 return "VLDTR_E_CT_NOTYPEDEFID "
		case -2146233621
		 return "VLDTR_E_FILE_BADFLAGS "
		case -2146233620
		 return "VLDTR_E_FILE_NULLHASH "
		case -2146233619
		 return "VLDTR_E_MOD_NONAME "
		case -2146233618
		 return "VLDTR_E_MOD_NAMEFULLQLFD "
		case -2146233617
		 return "VLDTR_E_TD_RTSPCLNOTSPCL "
		case -2146233616
		 return "VLDTR_E_TD_EXTENDSIFACE "
		case -2146233615
		 return "VLDTR_E_MD_CTORPINVOKE "
		case -2146233614
		 return "VLDTR_E_TD_SYSENUMNOTCLASS "
		case -2146233613
		 return "VLDTR_E_TD_SYSENUMNOTEXTVTYPE "
		case -2146233612
		 return "VLDTR_E_MI_SIGMISMATCH "
		case -2146233611
		 return "VLDTR_E_TD_ENUMHASMETHODS "
		case -2146233610
		 return "VLDTR_E_TD_ENUMIMPLIFACE "
		case -2146233609
		 return "VLDTR_E_TD_ENUMHASPROP "
		case -2146233608
		 return "VLDTR_E_TD_ENUMHASEVENT "
		case -2146233607
		 return "VLDTR_E_TD_BADMETHODLST "
		case -2146233606
		 return "VLDTR_E_TD_BADFIELDLST "
		case -2146233605
		 return "VLDTR_E_CN_BADTYPE "
		case -2146233604
		 return "VLDTR_E_TD_ENUMNOINSTFLD "
		case -2146233603
		 return "VLDTR_E_TD_ENUMMULINSTFLD "
		case -2146233602
		 return "VLDTR_E_INTERRUPTED "
		case -2146233601
		 return "VLDTR_E_NOTINIT "
		case -2146231552
		 return "VLDTR_E_IFACE_NOTIFACE "
		case -2146231551
		 return "VLDTR_E_FD_RVAHASNORVA "
		case -2146231550
		 return "VLDTR_E_FD_RVAHASZERORVA "
		case -2146231549
		 return "VLDTR_E_MD_RVAANDIMPLMAP "
		case -2146231548
		 return "VLDTR_E_TD_EXTRAFLAGS "
		case -2146231547
		 return "VLDTR_E_TD_EXTENDSITSELF "
		case -2146231546
		 return "VLDTR_E_TD_SYSVTNOTEXTOBJ "
		case -2146231545
		 return "VLDTR_E_TD_EXTTYPESPEC "
		case -2146231543
		 return "VLDTR_E_TD_VTNOSIZE "
		case -2146231542
		 return "VLDTR_E_TD_IFACESEALED "
		case -2146231541
		 return "VLDTR_E_NC_BADNESTED "
		case -2146231540
		 return "VLDTR_E_NC_BADENCLOSER "
		case -2146231539
		 return "VLDTR_E_NC_DUP "
		case -2146231538
		 return "VLDTR_E_NC_DUPENCLOSER "
		case -2146231537
		 return "VLDTR_E_FRVA_ZERORVA "
		case -2146231536
		 return "VLDTR_E_FRVA_BADFIELD "
		case -2146231535
		 return "VLDTR_E_FRVA_DUPRVA "
		case -2146231534
		 return "VLDTR_E_FRVA_DUPFIELD "
		case -2146231533
		 return "VLDTR_E_EP_BADTOKEN "
		case -2146231532
		 return "VLDTR_E_EP_INSTANCE "
		case -2146231531
		 return "VLDTR_E_TD_ENUMFLDBADTYPE "
		case -2146231530
		 return "VLDTR_E_MD_BADRVA "
		case -2146231529
		 return "VLDTR_E_FD_LITERALNODEFAULT "
		case -2146231528
		 return "VLDTR_E_IFACE_METHNOTIMPL "
		case -2146231527
		 return "VLDTR_E_CA_BADPARENT "
		case -2146231526
		 return "VLDTR_E_CA_BADTYPE "
		case -2146231525
		 return "VLDTR_E_CA_NOTCTOR "
		case -2146231524
		 return "VLDTR_E_CA_BADSIG "
		case -2146231523
		 return "VLDTR_E_CA_NOSIG "
		case -2146231522
		 return "VLDTR_E_CA_BADPROLOG "
		case -2146231521
		 return "VLDTR_E_MD_BADLOCALSIGTOK "
		case -2146231520
		 return "VLDTR_E_MD_BADHEADER "
		case -2146231519
		 return "VLDTR_E_EP_TOOMANYARGS "
		case -2146231518
		 return "VLDTR_E_EP_BADRET "
		case -2146231517
		 return "VLDTR_E_EP_BADARG "
		case -2146231516
		 return "VLDTR_E_SIG_BADVOID "
		case -2146231515
		 return "VLDTR_E_IFACE_METHMULTIMPL "
		case -2146231514
		 return "VLDTR_E_GP_NAMENULL "
		case -2146231513
		 return "VLDTR_E_GP_OWNERNIL "
		case -2146231512
		 return "VLDTR_E_GP_DUPNAME "
		case -2146231511
		 return "VLDTR_E_GP_DUPNUMBER "
		case -2146231510
		 return "VLDTR_E_GP_NONSEQ_BY_OWNER "
		case -2146231509
		 return "VLDTR_E_GP_NONSEQ_BY_NUMBER "
		case -2146231508
		 return "VLDTR_E_GP_UNEXPECTED_OWNER_FOR_VARIANT_VAR "
		case -2146231507
		 return "VLDTR_E_GP_ILLEGAL_VARIANT_MVAR "
		case -2146231506
		 return "VLDTR_E_GP_ILLEGAL_VARIANCE_FLAGS "
		case -2146231505
		 return "VLDTR_E_GP_REFANDVALUETYPE "
		case -2146231504
		 return "VLDTR_E_GPC_OWNERNIL "
		case -2146231503
		 return "VLDTR_E_GPC_DUP "
		case -2146231502
		 return "VLDTR_E_GPC_NONCONTIGUOUS "
		case -2146231501
		 return "VLDTR_E_MS_METHODNIL "
		case -2146231500
		 return "VLDTR_E_MS_DUP "
		case -2146231499
		 return "VLDTR_E_MS_BADCALLINGCONV "
		case -2146231498
		 return "VLDTR_E_MS_MISSARITY "
		case -2146231497
		 return "VLDTR_E_MS_MISSARG "
		case -2146231496
		 return "VLDTR_E_MS_ARITYMISMATCH "
		case -2146231495
		 return "VLDTR_E_MS_METHODNOTGENERIC "
		case -2146231494
		 return "VLDTR_E_SIG_MISSARITY "
		case -2146231493
		 return "VLDTR_E_SIG_ARITYMISMATCH "
		case -2146231492
		 return "VLDTR_E_MD_GENERIC_CCTOR "
		case -2146231491
		 return "VLDTR_E_MD_GENERIC_CTOR "
		case -2146231490
		 return "VLDTR_E_MD_GENERIC_IMPORT "
		case -2146231489
		 return "VLDTR_E_MD_GENERIC_BADCALLCONV "
		case -2146231488
		 return "VLDTR_E_MD_GENERIC_GLOBAL "
		case -2146231487
		 return "VLDTR_E_EP_GENERIC_METHOD "
		case -2146231486
		 return "VLDTR_E_MD_MISSARITY "
		case -2146231485
		 return "VLDTR_E_MD_ARITYZERO "
		case -2146231484
		 return "VLDTR_E_SIG_ARITYZERO "
		case -2146231483
		 return "VLDTR_E_MS_ARITYZERO "
		case -2146231482
		 return "VLDTR_E_MD_GPMISMATCH "
		case -2146231481
		 return "VLDTR_E_EP_GENERIC_TYPE "
		case -2146231480
		 return "VLDTR_E_MI_DECLNOTGENERIC "
		case -2146231479
		 return "VLDTR_E_MI_IMPLNOTGENERIC "
		case -2146231478
		 return "VLDTR_E_MI_ARITYMISMATCH "
		case -2146231477
		 return "VLDTR_E_TD_EXTBADTYPESPEC "
		case -2146231476
		 return "VLDTR_E_SIG_BYREFINST "
		case -2146231475
		 return "VLDTR_E_MS_BYREFINST "
		case -2146231474
		 return "VLDTR_E_TS_EMPTY "
		case -2146231473
		 return "VLDTR_E_TS_HASSENTINALS "
		case -2146231472
		 return "VLDTR_E_TD_GENERICHASEXPLAYOUT "
		case -2146231471
		 return "VLDTR_E_SIG_BADTOKTYPE "
		case -2146231470
		 return "VLDTR_E_IFACE_METHNOTIMPLTHISMOD "
		case -2146233600
		 return "CORDBG_E_UNRECOVERABLE_ERROR "
		case -2146233599
		 return "CORDBG_E_PROCESS_TERMINATED "
		case -2146233598
		 return "CORDBG_E_PROCESS_NOT_SYNCHRONIZED "
		case -2146233597
		 return "CORDBG_E_CLASS_NOT_LOADED "
		case -2146233596
		 return "CORDBG_E_IL_VAR_NOT_AVAILABLE "
		case -2146233595
		 return "CORDBG_E_BAD_REFERENCE_VALUE "
		case -2146233594
		 return "CORDBG_E_FIELD_NOT_AVAILABLE "
		case -2146233593
		 return "CORDBG_E_NON_NATIVE_FRAME "
		case -2146233592
		 return "CORDBG_E_NONCONTINUABLE_EXCEPTION "
		case -2146233591
		 return "CORDBG_E_CODE_NOT_AVAILABLE "
		case -2146233590
		 return "CORDBG_E_FUNCTION_NOT_IL "
		case 1250059
		 return "CORDBG_S_BAD_START_SEQUENCE_POINT "
		case 1250060
		 return "CORDBG_S_BAD_END_SEQUENCE_POINT "
		case 1250061
		 return "CORDBG_S_INSUFFICIENT_INFO_FOR_SET_IP "
		case -2146233586
		 return "CORDBG_E_CANT_SET_IP_INTO_FINALLY "
		case -2146233585
		 return "CORDBG_E_CANT_SET_IP_OUT_OF_FINALLY "
		case -2146233584
		 return "CORDBG_E_CANT_SET_IP_INTO_CATCH "
		case -2146233583
		 return "CORDBG_E_SET_IP_NOT_ALLOWED_ON_NONLEAF_FRAME "
		case -2146233582
		 return "CORDBG_E_SET_IP_IMPOSSIBLE "
		case -2146233581
		 return "CORDBG_E_FUNC_EVAL_BAD_START_POINT "
		case -2146233580
		 return "CORDBG_E_INVALID_OBJECT "
		case -2146233579
		 return "CORDBG_E_FUNC_EVAL_NOT_COMPLETE "
		case 1250070
		 return "CORDBG_S_FUNC_EVAL_HAS_NO_RESULT "
		case 1250071
		 return "CORDBG_S_VALUE_POINTS_TO_VOID "
		case -2146233576
		 return "CORDBG_E_INPROC_NOT_IMPL "
		case 1250073
		 return "CORDBG_S_FUNC_EVAL_ABORTED "
		case -2146233574
		 return "CORDBG_E_STATIC_VAR_NOT_AVAILABLE "
		case -2146233573
		 return "CORDBG_E_OBJECT_IS_NOT_COPYABLE_VALUE_CLASS "
		case -2146233572
		 return "CORDBG_E_CANT_SETIP_INTO_OR_OUT_OF_FILTER "
		case -2146233571
		 return "CORDBG_E_CANT_CHANGE_JIT_SETTING_FOR_ZAP_MODULE "
		case -2146233570
		 return "CORDBG_E_CANT_SET_IP_OUT_OF_FINALLY_ON_WIN64 "
		case -2146233569
		 return "CORDBG_E_CANT_SET_IP_OUT_OF_CATCH_ON_WIN64 "
		case -2146233568
		 return "CORDBG_E_REMOTE_CONNECTION_CONN_RESET "
		case -2146233567
		 return "CORDBG_E_REMOTE_CONNECTION_KEEP_ALIVE "
		case -2146233566
		 return "CORDBG_E_REMOTE_CONNECTION_FATAL_ERROR "
		case -2146233565
		 return "CORDBG_E_CANT_SET_TO_JMC "
		case -2146233555
		 return "CORDBG_E_BAD_THREAD_STATE "
		case -2146233554
		 return "CORDBG_E_DEBUGGER_ALREADY_ATTACHED "
		case -2146233553
		 return "CORDBG_E_SUPERFLOUS_CONTINUE "
		case -2146233552
		 return "CORDBG_E_SET_VALUE_NOT_ALLOWED_ON_NONLEAF_FRAME "
		case -2146233551
		 return "CORDBG_E_ENC_EH_MAX_NESTING_LEVEL_CANT_INCREASE "
		case -2146233550
		 return "CORDBG_E_ENC_MODULE_NOT_ENC_ENABLED "
		case -2146233549
		 return "CORDBG_E_SET_IP_NOT_ALLOWED_ON_EXCEPTION "
		case -2146233548
		 return "CORDBG_E_VARIABLE_IS_ACTUALLY_LITERAL "
		case -2146233547
		 return "CORDBG_E_PROCESS_DETACHED "
		case -2146233546
		 return "CORDBG_E_ENC_METHOD_SIG_CHANGED "
		case -2146233545
		 return "CORDBG_E_ENC_METHOD_NO_LOCAL_SIG "
		case -2146233544
		 return "CORDBG_E_ENC_CANT_ADD_FIELD_TO_VALUE_OR_LAYOUT_CLASS "
		case -2146233543
		 return "CORDBG_E_ENC_CANT_CHANGE_FIELD "
		case -2146233542
		 return "CORDBG_E_ENC_CANT_ADD_NON_PRIVATE_MEMBER "
		case -2146233541
		 return "CORDBG_E_FIELD_NOT_STATIC "
		case -2146233540
		 return "CORDBG_E_FIELD_NOT_INSTANCE "
		case -2146233539
		 return "CORDBG_E_ENC_ZAPPED_WITHOUT_ENC "
		case -2146233538
		 return "CORDBG_E_ENC_BAD_METHOD_INFO "
		case -2146233537
		 return "CORDBG_E_ENC_JIT_CANT_UPDATE "
		case -2146233536
		 return "CORDBG_E_ENC_MISSING_CLASS "
		case -2146233535
		 return "CORDBG_E_ENC_INTERNAL_ERROR "
		case -2146233534
		 return "CORDBG_E_ENC_HANGING_FIELD "
		case -2146233533
		 return "CORDBG_E_MODULE_NOT_LOADED "
		case -2146233532
		 return "CORDBG_E_ENC_CANT_CHANGE_SUPERCLASS "
		case -2146233531
		 return "CORDBG_E_UNABLE_TO_SET_BREAKPOINT "
		case -2146233530
		 return "CORDBG_E_DEBUGGING_NOT_POSSIBLE "
		case -2146233529
		 return "CORDBG_E_KERNEL_DEBUGGER_ENABLED "
		case -2146233528
		 return "CORDBG_E_KERNEL_DEBUGGER_PRESENT "
		case -2146233527
		 return "CORDBG_E_HELPER_THREAD_DEAD "
		case -2146233526
		 return "CORDBG_E_INTERFACE_INHERITANCE_CANT_CHANGE "
		case -2146233525
		 return "CORDBG_E_INCOMPATIBLE_PROTOCOL "
		case -2146233524
		 return "CORDBG_E_TOO_MANY_PROCESSES "
		case -2146233523
		 return "CORDBG_E_INTEROP_NOT_SUPPORTED "
		case -2146233522
		 return "CORDBG_E_NO_REMAP_BREAKPIONT "
		case -2146233521
		 return "CORDBG_E_OBJECT_NEUTERED "
		case -2146233520
		 return "CORPROF_E_FUNCTION_NOT_COMPILED "
		case -2146233519
		 return "CORPROF_E_DATAINCOMPLETE "
		case -2146233518
		 return "CORPROF_E_NOT_REJITABLE_METHODS "
		case -2146233517
		 return "CORPROF_E_CANNOT_UPDATE_METHOD "
		case -2146233516
		 return "CORPROF_E_FUNCTION_NOT_IL "
		case -2146233515
		 return "CORPROF_E_NOT_MANAGED_THREAD "
		case -2146233514
		 return "CORPROF_E_CALL_ONLY_FROM_INIT "
		case -2146233513
		 return "CORPROF_E_INPROC_NOT_ENABLED "
		case -2146233512
		 return "CORPROF_E_JITMAPS_NOT_ENABLED "
		case -2146233511
		 return "CORPROF_E_INPROC_ALREADY_BEGUN "
		case -2146233510
		 return "CORPROF_E_INPROC_NOT_AVAILABLE "
		case -2146233509
		 return "CORPROF_E_NOT_YET_AVAILABLE "
		case -2146233508
		 return "CORPROF_E_TYPE_IS_PARAMETERIZED "
		case -2146233507
		 return "CORPROF_E_FUNCTION_IS_PARAMETERIZED "
		case -2146233344
		 return "SECURITY_E_XML_TO_ASN_ENCODING "
		case -2146233343
		 return "SECURITY_E_INCOMPATIBLE_SHARE "
		case -2146233342
		 return "SECURITY_E_UNVERIFIABLE "
		case -2146233341
		 return "SECURITY_E_INCOMPATIBLE_EVIDENCE "
		case -2146230273
		 return "CLDB_E_INTERNALERROR "
		case -2146233328
		 return "CORSEC_E_DECODE_SET "
		case -2146233327
		 return "CORSEC_E_ENCODE_SET "
		case -2146233326
		 return "CORSEC_E_UNSUPPORTED_FORMAT "
		case -2146233325
		 return "SN_CRYPTOAPI_CALL_FAILED "
		case -2146233325
		 return "CORSEC_E_CRYPTOAPI_CALL_FAILED "
		case -2146233324
		 return "SN_NO_SUITABLE_CSP "
		case -2146233324
		 return "CORSEC_E_NO_SUITABLE_CSP "
		case -2146233323
		 return "CORSEC_E_INVALID_ATTR "
		case -2146233322
		 return "CORSEC_E_POLICY_EXCEPTION "
		case -2146233321
		 return "CORSEC_E_MIN_GRANT_FAIL "
		case -2146233320
		 return "CORSEC_E_NO_EXEC_PERM "
		case -2146233319
		 return "CORSEC_E_XMLSYNTAX "
		case -2146233318
		 return "CORSEC_E_INVALID_STRONGNAME "
		case -2146233317
		 return "CORSEC_E_MISSING_STRONGNAME "
		case -2146233316
		 return "CORSEC_E_CONTAINER_NOT_FOUND "
		case -2146233315
		 return "CORSEC_E_INVALID_IMAGE_FORMAT "
		case -2146233314
		 return "CORSEC_E_INVALID_PUBLICKEY "
		case -2146233312
		 return "CORSEC_E_SIGNATURE_MISMATCH "
		case -2146233296
		 return "CORSEC_E_CRYPTO "
		case -2146233295
		 return "CORSEC_E_CRYPTO_UNEX_OPER "
		case -2146233286
		 return "CORSECATTR_E_BAD_ATTRIBUTE "
		case -2146233285
		 return "CORSECATTR_E_MISSING_CONSTRUCTOR "
		case -2146233284
		 return "CORSECATTR_E_FAILED_TO_CREATE_PERM "
		case -2146233283
		 return "CORSECATTR_E_BAD_ACTION_ASM "
		case -2146233282
		 return "CORSECATTR_E_BAD_ACTION_OTHER "
		case -2146233281
		 return "CORSECATTR_E_BAD_PARENT "
		case -2146233280
		 return "CORSECATTR_E_TRUNCATED "
		case -2146233279
		 return "CORSECATTR_E_BAD_VERSION "
		case -2146233278
		 return "CORSECATTR_E_BAD_ACTION "
		case -2146233277
		 return "CORSECATTR_E_NO_SELF_REF "
		case -2146233276
		 return "CORSECATTR_E_BAD_NONCAS "
		case -2146233275
		 return "CORSECATTR_E_ASSEMBLY_LOAD_FAILED "
		case -2146233274
		 return "CORSECATTR_E_ASSEMBLY_LOAD_FAILED_EX "
		case -2146233273
		 return "CORSECATTR_E_TYPE_LOAD_FAILED "
		case -2146233272
		 return "CORSECATTR_E_TYPE_LOAD_FAILED_EX "
		case -2146233271
		 return "CORSECATTR_E_ABSTRACT "
		case -2146233270
		 return "CORSECATTR_E_UNSUPPORTED_TYPE "
		case -2146233269
		 return "CORSECATTR_E_UNSUPPORTED_ENUM_TYPE "
		case -2146233268
		 return "CORSECATTR_E_NO_FIELD "
		case -2146233267
		 return "CORSECATTR_E_NO_PROPERTY "
		case -2146233266
		 return "CORSECATTR_E_EXCEPTION "
		case -2146233265
		 return "CORSECATTR_E_EXCEPTION_HR "
		case -2146233264
		 return "ISS_E_ISOSTORE "
		case -2146233248
		 return "ISS_E_OPEN_STORE_FILE "
		case -2146233247
		 return "ISS_E_OPEN_FILE_MAPPING "
		case -2146233246
		 return "ISS_E_MAP_VIEW_OF_FILE "
		case -2146233245
		 return "ISS_E_GET_FILE_SIZE "
		case -2146233244
		 return "ISS_E_CREATE_MUTEX "
		case -2146233243
		 return "ISS_E_LOCK_FAILED "
		case -2146233242
		 return "ISS_E_FILE_WRITE "
		case -2146233241
		 return "ISS_E_SET_FILE_POINTER "
		case -2146233240
		 return "ISS_E_CREATE_DIR "
		case -2146233239
		 return "ISS_E_STORE_NOT_OPEN "
		case -2146233216
		 return "ISS_E_CORRUPTED_STORE_FILE "
		case -2146233215
		 return "ISS_E_STORE_VERSION "
		case -2146233214
		 return "ISS_E_FILE_NOT_MAPPED "
		case -2146233213
		 return "ISS_E_BLOCK_SIZE_TOO_SMALL "
		case -2146233212
		 return "ISS_E_ALLOC_TOO_LARGE "
		case -2146233211
		 return "ISS_E_USAGE_WILL_EXCEED_QUOTA "
		case -2146233210
		 return "ISS_E_TABLE_ROW_NOT_FOUND "
		case -2146233184
		 return "ISS_E_DEPRECATE "
		case -2146233183
		 return "ISS_E_CALLER "
		case -2146233182
		 return "ISS_E_PATH_LENGTH "
		case -2146233181
		 return "ISS_E_MACHINE "
		case -2146233180
		 return "ISS_E_MACHINE_DACL "
		case -2146233264
		 return "ISS_E_ISOSTORE_START "
		case -2146233089
		 return "ISS_E_ISOSTORE_END "
		case -2146232832
		 return "COR_E_APPLICATION "
		case -2147024809
		 return "COR_E_ARGUMENT "
		case -2146233086
		 return "COR_E_ARGUMENTOUTOFRANGE "
		case -2147024362
		 return "COR_E_ARITHMETIC "
		case -2146233085
		 return "COR_E_ARRAYTYPEMISMATCH "
		case -2146233084
		 return "COR_E_CONTEXTMARSHAL "
		case -2146233083
		 return "COR_E_TIMEOUT "
		case -2146232969
		 return "COR_E_KEYNOTFOUND "
		case -2146233024
		 return "COR_E_DEVICESNOTSUPPORTED "
		case -2147352558
		 return "COR_E_DIVIDEBYZERO "
		case -2146233088
		 return "COR_E_EXCEPTION "
		case -2146233082
		 return "COR_E_EXECUTIONENGINE "
		case -2146233081
		 return "COR_E_FIELDACCESS "
		case -2146233033
		 return "COR_E_FORMAT "
		case -2147024885
		 return "COR_E_BADIMAGEFORMAT "
		case -2146234344
		 return "COR_E_ASSEMBLYEXPECTED "
		case -2146234349
		 return "COR_E_TYPEUNLOADED "
		case -2146233080
		 return "COR_E_INDEXOUTOFRANGE "
		case -2147467262
		 return "COR_E_INVALIDCAST "
		case -2146233079
		 return "COR_E_INVALIDOPERATION "
		case -2146233030
		 return "COR_E_INVALIDPROGRAM "
		case -2146233062
		 return "COR_E_MEMBERACCESS "
		case -2146233072
		 return "COR_E_METHODACCESS "
		case -2146233071
		 return "COR_E_MISSINGFIELD "
		case -2146233038
		 return "COR_E_MISSINGMANIFESTRESOURCE "
		case -2146233070
		 return "COR_E_MISSINGMEMBER "
		case -2146233069
		 return "COR_E_MISSINGMETHOD "
		case -2146233034
		 return "COR_E_MISSINGSATELLITEASSEMBLY "
		case -2146233068
		 return "COR_E_MULTICASTNOTSUPPORTED "
		case -2146233048
		 return "COR_E_NOTFINITENUMBER "
		case -2146233047
		 return "COR_E_DUPLICATEWAITOBJECT "
		case -2146233031
		 return "COR_E_PLATFORMNOTSUPPORTED "
		case -2146233067
		 return "COR_E_NOTSUPPORTED "
		case -2147467261
		 return "COR_E_NULLREFERENCE "
		case -2147024882
		 return "COR_E_OUTOFMEMORY "
		case -2146233066
		 return "COR_E_OVERFLOW "
		case -2146233065
		 return "COR_E_RANK "
		case -2146233077
		 return "COR_E_REMOTING "
		case -2146233074
		 return "COR_E_SERVER "
		case -2146233073
		 return "COR_E_SERVICEDCOMPONENT "
		case -2146233078
		 return "COR_E_SECURITY "
		case -2146233076
		 return "COR_E_SERIALIZATION "
		case -2147023895
		 return "COR_E_STACKOVERFLOW "
		case -2146233064
		 return "COR_E_SYNCHRONIZATIONLOCK "
		case -2146233087
		 return "COR_E_SYSTEM "
		case -2146233040
		 return "COR_E_THREADABORTED "
		case -2146233029
		 return "COR_E_OPERATIONCANCELED "
		case -2146233028
		 return "COR_E_NOTCANCELABLE "
		case -2146233063
		 return "COR_E_THREADINTERRUPTED "
		case -2146233056
		 return "COR_E_THREADSTATE "
		case -2146233055
		 return "COR_E_THREADSTOP "
		case -2146233036
		 return "COR_E_TYPEINITIALIZATION "
		case -2146233054
		 return "COR_E_TYPELOAD "
		case -2146233053
		 return "COR_E_ENTRYPOINTNOTFOUND "
		case -2146233052
		 return "COR_E_DLLNOTFOUND "
		case -2147024891
		 return "COR_E_UNAUTHORIZEDACCESS "
		case -2146233075
		 return "COR_E_VERIFICATION "
		case -2146233049
		 return "COR_E_INVALIDCOMOBJECT "
		case -2146233046
		 return "COR_E_COMOBJECTINUSE "
		case -2146233045
		 return "COR_E_SEMAPHOREFULL "
		case -2146233044
		 return "COR_E_WAITHANDLECANNOTBEOPENED "
		case -2146233043
		 return "COR_E_ABANDONEDMUTEX "
		case -2146233035
		 return "COR_E_MARSHALDIRECTIVE "
		case -2146233039
		 return "COR_E_INVALIDOLEVARIANTTYPE "
		case -2146233037
		 return "COR_E_SAFEARRAYTYPEMISMATCH "
		case -2146233032
		 return "COR_E_SAFEARRAYRANKMISMATCH "
		case -2146233023
		 return "COR_E_DATAMISALIGNED "
		case -2147352562
		 return "COR_E_TARGETPARAMCOUNT "
		case -2147475171
		 return "COR_E_AMBIGUOUSMATCH "
		case -2146232831
		 return "COR_E_INVALIDFILTERCRITERIA "
		case -2146232830
		 return "COR_E_REFLECTIONTYPELOAD "
		case -2146232829
		 return "COR_E_TARGET "
		case -2146232828
		 return "COR_E_TARGETINVOCATION "
		case -2146232827
		 return "COR_E_CUSTOMATTRIBUTEFORMAT "
		case -2146232799
		 return "COR_E_FILELOAD "
		case -2146232800
		 return "COR_E_IO "
		case -2146232798
		 return "COR_E_OBJECTDISPOSED "
		case -2146232768
		 return "COR_E_HOSTPROTECTION "
		case -2146232797
		 return "COR_E_FAILFAST "
		case -2146232576
		 return "CLR_E_SHIM_RUNTIMELOAD "
		case -2146232575
		 return "CLR_E_SHIM_RUNTIMEEXPORT "
		case -2146232574
		 return "CLR_E_SHIM_INSTALLROOT "
		case -2146232573
		 return "CLR_E_SHIM_INSTALLCOMP "
		case -2146232319
		 return "VER_E_HRESULT "
		case -2146232318
		 return "VER_E_OFFSET "
		case -2146232317
		 return "VER_E_OPCODE "
		case -2146232316
		 return "VER_E_OPERAND "
		case -2146232315
		 return "VER_E_TOKEN "
		case -2146232314
		 return "VER_E_EXCEPT "
		case -2146232313
		 return "VER_E_STACK_SLOT "
		case -2146232312
		 return "VER_E_LOC "
		case -2146232311
		 return "VER_E_ARG "
		case -2146232310
		 return "VER_E_FOUND "
		case -2146232309
		 return "VER_E_EXPECTED "
		case -2146232308
		 return "VER_E_LOC_BYNAME "
		case -2146232304
		 return "VER_E_UNKNOWN_OPCODE "
		case -2146232303
		 return "VER_E_SIG_CALLCONV "
		case -2146232302
		 return "VER_E_SIG_ELEMTYPE "
		case -2146232300
		 return "VER_E_RET_SIG "
		case -2146232299
		 return "VER_E_FIELD_SIG "
		case -2146232296
		 return "VER_E_INTERNAL "
		case -2146232295
		 return "VER_E_STACK_TOO_LARGE "
		case -2146232294
		 return "VER_E_ARRAY_NAME_LONG "
		case -2146232288
		 return "VER_E_FALLTHRU "
		case -2146232287
		 return "VER_E_TRY_GTEQ_END "
		case -2146232286
		 return "VER_E_TRYEND_GT_CS "
		case -2146232285
		 return "VER_E_HND_GTEQ_END "
		case -2146232284
		 return "VER_E_HNDEND_GT_CS "
		case -2146232283
		 return "VER_E_FLT_GTEQ_CS "
		case -2146232282
		 return "VER_E_TRY_START "
		case -2146232281
		 return "VER_E_HND_START "
		case -2146232280
		 return "VER_E_FLT_START "
		case -2146232279
		 return "VER_E_TRY_OVERLAP "
		case -2146232278
		 return "VER_E_TRY_EQ_HND_FIL "
		case -2146232277
		 return "VER_E_TRY_SHARE_FIN_FAL "
		case -2146232276
		 return "VER_E_HND_OVERLAP "
		case -2146232275
		 return "VER_E_HND_EQ "
		case -2146232274
		 return "VER_E_FIL_OVERLAP "
		case -2146232273
		 return "VER_E_FIL_EQ "
		case -2146232272
		 return "VER_E_FIL_CONT_TRY "
		case -2146232271
		 return "VER_E_FIL_CONT_HND "
		case -2146232270
		 return "VER_E_FIL_CONT_FIL "
		case -2146232269
		 return "VER_E_FIL_GTEQ_CS "
		case -2146232268
		 return "VER_E_FIL_START "
		case -2146232267
		 return "VER_E_FALLTHRU_EXCEP "
		case -2146232266
		 return "VER_E_FALLTHRU_INTO_HND "
		case -2146232265
		 return "VER_E_FALLTHRU_INTO_FIL "
		case -2146232264
		 return "VER_E_LEAVE "
		case -2146232263
		 return "VER_E_RETHROW "
		case -2146232262
		 return "VER_E_ENDFINALLY "
		case -2146232261
		 return "VER_E_ENDFILTER "
		case -2146232260
		 return "VER_E_ENDFILTER_MISSING "
		case -2146232259
		 return "VER_E_BR_INTO_TRY "
		case -2146232258
		 return "VER_E_BR_INTO_HND "
		case -2146232257
		 return "VER_E_BR_INTO_FIL "
		case -2146232256
		 return "VER_E_BR_OUTOF_TRY "
		case -2146232255
		 return "VER_E_BR_OUTOF_HND "
		case -2146232254
		 return "VER_E_BR_OUTOF_FIL "
		case -2146232253
		 return "VER_E_BR_OUTOF_FIN "
		case -2146232252
		 return "VER_E_RET_FROM_TRY "
		case -2146232251
		 return "VER_E_RET_FROM_HND "
		case -2146232250
		 return "VER_E_RET_FROM_FIL "
		case -2146232249
		 return "VER_E_BAD_JMP_TARGET "
		case -2146232248
		 return "VER_E_PATH_LOC "
		case -2146232247
		 return "VER_E_PATH_THIS "
		case -2146232246
		 return "VER_E_PATH_STACK "
		case -2146232245
		 return "VER_E_PATH_STACK_DEPTH "
		case -2146232244
		 return "VER_E_THIS "
		case -2146232243
		 return "VER_E_THIS_UNINIT_EXCEP "
		case -2146232242
		 return "VER_E_THIS_UNINIT_STORE "
		case -2146232241
		 return "VER_E_THIS_UNINIT_RET "
		case -2146232240
		 return "VER_E_THIS_UNINIT_V_RET "
		case -2146232239
		 return "VER_E_THIS_UNINIT_BR "
		case -2146232238
		 return "VER_E_LDFTN_CTOR "
		case -2146232237
		 return "VER_E_STACK_NOT_EQ "
		case -2146232236
		 return "VER_E_STACK_UNEXPECTED "
		case -2146232235
		 return "VER_E_STACK_EXCEPTION "
		case -2146232234
		 return "VER_E_STACK_OVERFLOW "
		case -2146232233
		 return "VER_E_STACK_UNDERFLOW "
		case -2146232232
		 return "VER_E_STACK_EMPTY "
		case -2146232231
		 return "VER_E_STACK_UNINIT "
		case -2146232230
		 return "VER_E_STACK_I_I4_I8 "
		case -2146232229
		 return "VER_E_STACK_R_R4_R8 "
		case -2146232228
		 return "VER_E_STACK_NO_R_I8 "
		case -2146232227
		 return "VER_E_STACK_NUMERIC "
		case -2146232226
		 return "VER_E_STACK_OBJREF "
		case -2146232225
		 return "VER_E_STACK_P_OBJREF "
		case -2146232224
		 return "VER_E_STACK_BYREF "
		case -2146232223
		 return "VER_E_STACK_METHOD "
		case -2146232222
		 return "VER_E_STACK_ARRAY_SD "
		case -2146232221
		 return "VER_E_STACK_VALCLASS "
		case -2146232220
		 return "VER_E_STACK_P_VALCLASS "
		case -2146232219
		 return "VER_E_STACK_NO_VALCLASS "
		case -2146232218
		 return "VER_E_LOC_DEAD "
		case -2146232217
		 return "VER_E_LOC_NUM "
		case -2146232216
		 return "VER_E_ARG_NUM "
		case -2146232215
		 return "VER_E_TOKEN_RESOLVE "
		case -2146232214
		 return "VER_E_TOKEN_TYPE "
		case -2146232213
		 return "VER_E_TOKEN_TYPE_MEMBER "
		case -2146232212
		 return "VER_E_TOKEN_TYPE_FIELD "
		case -2146232211
		 return "VER_E_TOKEN_TYPE_SIG "
		case -2146232210
		 return "VER_E_UNVERIFIABLE "
		case -2146232209
		 return "VER_E_LDSTR_OPERAND "
		case -2146232208
		 return "VER_E_RET_PTR_TO_STACK "
		case -2146232207
		 return "VER_E_RET_VOID "
		case -2146232206
		 return "VER_E_RET_MISSING "
		case -2146232205
		 return "VER_E_RET_EMPTY "
		case -2146232204
		 return "VER_E_RET_UNINIT "
		case -2146232203
		 return "VER_E_ARRAY_ACCESS "
		case -2146232202
		 return "VER_E_ARRAY_V_STORE "
		case -2146232201
		 return "VER_E_ARRAY_SD "
		case -2146232200
		 return "VER_E_ARRAY_SD_PTR "
		case -2146232199
		 return "VER_E_ARRAY_FIELD "
		case -2146232198
		 return "VER_E_ARGLIST "
		case -2146232197
		 return "VER_E_VALCLASS "
		case -2146232196
		 return "VER_E_METHOD_ACCESS "
		case -2146232195
		 return "VER_E_FIELD_ACCESS "
		case -2146232194
		 return "VER_E_DEAD "
		case -2146232193
		 return "VER_E_FIELD_STATIC "
		case -2146232192
		 return "VER_E_FIELD_NO_STATIC "
		case -2146232191
		 return "VER_E_ADDR "
		case -2146232190
		 return "VER_E_ADDR_BYREF "
		case -2146232189
		 return "VER_E_ADDR_LITERAL "
		case -2146232188
		 return "VER_E_INITONLY "
		case -2146232187
		 return "VER_E_THROW "
		case -2146232186
		 return "VER_E_CALLVIRT_VALCLASS "
		case -2146232185
		 return "VER_E_CALL_SIG "
		case -2146232184
		 return "VER_E_CALL_STATIC "
		case -2146232183
		 return "VER_E_CTOR "
		case -2146232182
		 return "VER_E_CTOR_VIRT "
		case -2146232181
		 return "VER_E_CTOR_OR_SUPER "
		case -2146232180
		 return "VER_E_CTOR_MUL_INIT "
		case -2146232179
		 return "VER_E_SIG "
		case -2146232178
		 return "VER_E_SIG_ARRAY "
		case -2146232177
		 return "VER_E_SIG_ARRAY_PTR "
		case -2146232176
		 return "VER_E_SIG_ARRAY_BYREF "
		case -2146232175
		 return "VER_E_SIG_ELEM_PTR "
		case -2146232174
		 return "VER_E_SIG_VARARG "
		case -2146232173
		 return "VER_E_SIG_VOID "
		case -2146232172
		 return "VER_E_SIG_BYREF_BYREF "
		case -2146232170
		 return "VER_E_CODE_SIZE_ZERO "
		case -2146232169
		 return "VER_E_BAD_VARARG "
		case -2146232168
		 return "VER_E_TAIL_CALL "
		case -2146232167
		 return "VER_E_TAIL_BYREF "
		case -2146232166
		 return "VER_E_TAIL_RET "
		case -2146232165
		 return "VER_E_TAIL_RET_VOID "
		case -2146232164
		 return "VER_E_TAIL_RET_TYPE "
		case -2146232163
		 return "VER_E_TAIL_STACK_EMPTY "
		case -2146232162
		 return "VER_E_METHOD_END "
		case -2146232161
		 return "VER_E_BAD_BRANCH "
		case -2146232160
		 return "VER_E_FIN_OVERLAP "
		case -2146232159
		 return "VER_E_LEXICAL_NESTING "
		case -2146232158
		 return "VER_E_VOLATILE "
		case -2146232157
		 return "VER_E_UNALIGNED "
		case -2146232156
		 return "VER_E_INNERMOST_FIRST "
		case -2146232155
		 return "VER_E_CALLI_VIRTUAL "
		case -2146232154
		 return "VER_E_CALL_ABSTRACT "
		case -2146232153
		 return "VER_E_STACK_UNEXP_ARRAY "
		case -2146232152
		 return "VER_E_NOT_IN_GC_HEAP "
		case -2146232151
		 return "VER_E_TRY_N_EMPTY_STACK "
		case -2146232150
		 return "VER_E_DLGT_CTOR "
		case -2146232149
		 return "VER_E_DLGT_BB "
		case -2146232148
		 return "VER_E_DLGT_PATTERN "
		case -2146232147
		 return "VER_E_DLGT_LDFTN "
		case -2146232146
		 return "VER_E_FTN_ABSTRACT "
		case -2146232145
		 return "VER_E_SIG_C_VC "
		case -2146232144
		 return "VER_E_SIG_VC_C "
		case -2146232143
		 return "VER_E_BOX_PTR_TO_STACK "
		case -2146232142
		 return "VER_E_SIG_BYREF_TB_AH "
		case -2146232141
		 return "VER_E_SIG_ARRAY_TB_AH "
		case -2146232140
		 return "VER_E_ENDFILTER_STACK "
		case -2146232139
		 return "VER_E_DLGT_SIG_I "
		case -2146232138
		 return "VER_E_DLGT_SIG_O "
		case -2146232137
		 return "VER_E_RA_PTR_TO_STACK "
		case -2146232136
		 return "VER_E_CATCH_VALUE_TYPE "
		case -2146232135
		 return "VER_E_CATCH_BYREF "
		case -2146232134
		 return "VER_E_FIL_PRECEED_HND "
		case -2146232133
		 return "VER_E_LDVIRTFTN_STATIC "
		case -2146232132
		 return "VER_E_CALLVIRT_STATIC "
		case -2146232131
		 return "VER_E_INITLOCALS "
		case -2146232130
		 return "VER_E_BR_TO_EXCEPTION "
		case -2146232129
		 return "VER_E_CALL_CTOR "
		case -2146232128
		 return "VER_E_VALCLASS_OBJREF_VAR "
		case -2146232127
		 return "VER_E_STACK_P_VALCLASS_OBJREF_VAR "
		case -2146232126
		 return "VER_E_SIG_VAR_PARAM "
		case -2146232125
		 return "VER_E_SIG_MVAR_PARAM "
		case -2146232124
		 return "VER_E_SIG_VAR_ARG "
		case -2146232123
		 return "VER_E_SIG_MVAR_ARG "
		case -2146232122
		 return "VER_E_SIG_GENERICINST "
		case -2146232121
		 return "VER_E_SIG_METHOD_INST "
		case -2146232120
		 return "VER_E_SIG_METHOD_PARENT_INST "
		case -2146232119
		 return "VER_E_SIG_FIELD_PARENT_INST "
		case -2146232118
		 return "VER_E_CALLCONV_NOT_GENERICINST "
		case -2146232117
		 return "VER_E_TOKEN_BAD_METHOD_SPEC "
		case -2146232116
		 return "VER_E_BAD_READONLY_PREFIX "
		case -2146232115
		 return "VER_E_BAD_CONSTRAINED_PREFIX "
		case -2146232114
		 return "VER_E_CIRCULAR_VAR_CONSTRAINTS "
		case -2146232113
		 return "VER_E_CIRCULAR_MVAR_CONSTRAINTS "
		case -2146232112
		 return "VER_E_UNSATISFIED_METHOD_INST "
		case -2146232111
		 return "VER_E_UNSATISFIED_METHOD_PARENT_INST "
		case -2146232110
		 return "VER_E_UNSATISFIED_FIELD_PARENT_INST "
		case -2146232109
		 return "VER_E_UNSATISFIED_BOX_OPERAND "
		case -2146232108
		 return "VER_E_CONSTRAINED_CALL_WITH_NON_BYREF_THIS "
		case -2146232107
		 return "VER_E_CONSTRAINED_OF_NON_VARIABLE_TYPE "
		case -2146232106
		 return "VER_E_READONLY_UNEXPECTED_CALLEE "
		case -2146232105
		 return "VER_E_READONLY_ILLEGAL_WRITE "
		case -2146232104
		 return "VER_E_READONLY_IN_MKREFANY "
		case -2146232103
		 return "VER_E_UNALIGNED_ALIGNMENT "
		case -2146232102
		 return "VER_E_TAILCALL_INSIDE_EH "
		case -2146232101
		 return "VER_E_BACKWARD_BRANCH "
		case -2146232100
		 return "VER_E_CALL_TO_VTYPE_BASE "
		case -2146232099
		 return "VER_E_NEWOBJ_OF_ABSTRACT_CLASS "
		case -2146232096
		 return "VER_E_FIELD_OVERLAP "
		case -2146232080
		 return "VER_E_BAD_PE "
		case -2146232079
		 return "VER_E_BAD_MD "
		case -2146232078
		 return "VER_E_BAD_APPDOMAIN "
		case -2146232077
		 return "VER_E_TYPELOAD "
		case -2146232076
		 return "VER_E_PE_LOAD "
		case -2146232075
		 return "VER_E_WRITE_RVA_STATIC "
		case -2146231296
		 return "CORDBG_E_THREAD_NOT_SCHEDULED "
		case -2146231295
		 return "CORDBG_E_HANDLE_HAS_BEEN_DISPOSED "
		case -2146231294
		 return "CORDBG_E_NONINTERCEPTABLE_EXCEPTION "
		case -2146231293
		 return "CORDBG_E_CANT_UNWIND_ABOVE_CALLBACK "
		case -2146231292
		 return "CORDBG_E_INTERCEPT_FRAME_ALREADY_SET "
		case -2146231291
		 return "CORDBG_E_NO_NATIVE_PATCH_AT_ADDR "
		case -2146231290
		 return "CORDBG_E_MUST_BE_INTEROP_DEBUGGING "
		case -2146231289
		 return "CORDBG_E_NATIVE_PATCH_ALREADY_AT_ADDR "
		case -2146231288
		 return "CORDBG_E_TIMEOUT "
		case -2146231287
		 return "CORDBG_E_CANT_CALL_ON_THIS_THREAD "
		case -2146231286
		 return "CORDBG_E_ENC_INFOLESS_METHOD "
		case -2146231285
		 return "CORDBG_E_ENC_NESTED_HANLDERS "
		case -2146231284
		 return "CORDBG_E_ENC_IN_FUNCLET "
		case -2146231283
		 return "CORDBG_E_ENC_LOCALLOC "
		case -2146231282
		 return "CORDBG_E_ENC_EDIT_NOT_SUPPORTED "
		case -2146231281
		 return "CORDBG_E_FEABORT_DELAYED_UNTIL_THREAD_RESUMED "
		case -2146231280
		 return "CORDBG_E_NOTREADY "
		case -2146231279
		 return "CORDBG_E_CANNOT_RESOLVE_ASSEMBLY "
		case -2146231278
		 return "CORDBG_E_MUST_BE_IN_LOAD_MODULE "
		case -2146231277
		 return "CORDBG_E_CANNOT_BE_ON_ATTACH "
		case 1252371
		 return "CORDBG_S_NOT_ALL_BITS_SET "
		case -2146231276
		 return "CORDBG_E_NGEN_NOT_SUPPORTED "
		case -2146231275
		 return "CORDBG_E_ILLEGAL_SHUTDOWN_ORDER "
		case -2146231274
		 return "CORDBG_E_CANNOT_DEBUG_FIBER_PROCESS "
		case -2146231273
		 return "CORDBG_E_MUST_BE_IN_CREATE_PROCESS "
		case -2146231272
		 return "CORDBG_E_DETACH_FAILED_OUTSTANDING_EVALS "
		case -2146231271
		 return "CORDBG_E_DETACH_FAILED_OUTSTANDING_STEPPERS "
		case -2146231264
		 return "CORDBG_E_CANT_INTEROP_STEP_OUT "
		case -2146231263
		 return "CORDBG_E_DETACH_FAILED_OUTSTANDING_BREAKPOINTS "
		case -2146231262
		 return "CORDBG_E_ILLEGAL_IN_STACK_OVERFLOW "
		case -2146231261
		 return "CORDBG_E_ILLEGAL_AT_GC_UNSAFE_POINT "
		case -2146231260
		 return "CORDBG_E_ILLEGAL_IN_PROLOG "
		case -2146231259
		 return "CORDBG_E_ILLEGAL_IN_NATIVE_CODE "
		case -2146231258
		 return "CORDBG_E_ILLEGAL_IN_OPTIMIZED_CODE "
		case -2146231040
		 return "PEFMT_E_NO_CONTENTS "
		case -2146231039
		 return "PEFMT_E_NO_NTHEADERS "
		case -2146231038
		 return "PEFMT_E_64BIT "
		case -2146231037
		 return "PEFMT_E_NO_CORHEADER "
		case -2146231036
		 return "PEFMT_E_NOT_ILONLY "
		case -2146231035
		 return "PEFMT_E_IMPORT_DLLS "
		case -2146231034
		 return "PEFMT_E_EXE_NOENTRYPOINT "
		case -2146231033
		 return "PEFMT_E_BASE_RELOCS "
		case -2146231032
		 return "PEFMT_E_ENTRYPOINT "
		case -2146231031
		 return "PEFMT_E_ZERO_SIZEOFCODE "
		case -2146231030
		 return "PEFMT_E_BAD_CORHEADER "		
		Case Else
		 Return Str(hr)
	End Select
End Function
'************************************************************************************
'Event sink common procedure & constants
'************************************************************************************
TYPE Events_IDispatchVtbl
   QueryInterface AS DWORD     ' Returns pointers to supported interfaces
   AddRef AS DWORD             ' Increments reference count
   Release AS DWORD            ' Decrements reference count
   GetTypeInfoCount AS DWORD   ' Retrieves the number of type descriptions
   GetTypeInfo AS DWORD        ' Retrieves a description of object's programmable interface
   GetIDsOfNames AS DWORD      ' Maps name of method or property to DispId
   Invoke AS DWORD             ' Calls one of the object's methods, or gets/sets one of its properties
   pVtblAddr AS DWORD          ' Address of the virtual table
   cRef AS DWORD               ' Reference counter
   pthis AS DWORD              ' IUnknown or IDispatch of the control that fires the events
END Type

' ****************************************************************************************
' UI4 AddRef()
' Increments the reference counter.
' ****************************************************************************************
FUNCTION Events_AddRef (BYVAL pCookie AS Events_IDispatchVtbl PTR) AS DWORD
   pCookie->cRef+=1
   FUNCTION = pCookie->cRef
END FUNCTION

' ****************************************************************************************
' UI4 Release()
' Releases our class if there is only a reference to him and decrements the reference counter.
' ****************************************************************************************
FUNCTION Events_Release (BYVAL pCookie AS Events_IDispatchVtbl PTR) AS DWORD
	Dim pVtblAddr AS DWORD
	
	If pCookie->cRef = 1 THEN
		pVtblAddr = pCookie->pVtblAddr
		IF HeapFree(GetProcessHeap(), 0, BYVAL pVtblAddr) THEN
			FUNCTION = 0
			EXIT Function
		ELSE
			FUNCTION = pCookie->cRef
			EXIT FUNCTION
		END IF
	End IF
	pCookie->cRef-=1
	Function = pCookie->cRef
End FUNCTION

' ****************************************************************************************
' HRESULT GetTypeInfoCount([out] *UINT pctinfo)
' ****************************************************************************************
FUNCTION Events_GetTypeInfoCount (BYVAL pCookie AS Events_IDispatchVtbl PTR, BYREF pctInfo AS DWORD) AS LONG
   pctInfo = 0
   Function = S_OK
END FUNCTION

' ****************************************************************************************
' HRESULT GetTypeInfo([in] UINT itinfo, [in] UI4 lcid, [out] **VOID pptinfo)
' ****************************************************************************************
FUNCTION Events_GetTypeInfo (BYVAL pCookie AS Events_IDispatchVtbl PTR, _
   BYVAL itinfo AS DWORD, BYVAL lcid AS DWORD, BYREF pptinfo AS DWORD) AS LONG
   FUNCTION = E_NOTIMPL
END Function

' ****************************************************************************************
' HRESULT GetTypeInfo([in] UINT itinfo, [in] UI4 lcid, [out] **VOID pptinfo)
' ****************************************************************************************
FUNCTION Events_TypeInfo (BYVAL pCookie AS Events_IDispatchVtbl PTR, _
   BYVAL itinfo AS DWORD, BYVAL lcid AS DWORD, BYREF pptinfo AS DWORD) AS LONG
   FUNCTION = E_NOTIMPL
END FUNCTION

' ****************************************************************************************
' HRESULT GetIDsOfNames([in] *GUID riid, [in] **I1 rgszNames, [in] UINT cNames, [in] UI4 lcid, [out] *I4 rgdispid)
' ****************************************************************************************
Function Events_GetIDsOfNames ( BYVAL pCookie AS Events_IDispatchVtbl PTR, _
   BYREF riid as IID, BYVAL rgszNames AS DWORD, BYVAL cNames AS DWORD, BYVAL lcid AS DWORD, BYREF rgdispid AS LONG) AS LONG
   FUNCTION = E_NOTIMPL
End Function

' ****************************************************************************************
' Builds the IDispatch Virtual Table
' ****************************************************************************************
Function Events_BuildVtbl (BYVAL pthis AS DWORD,byval qryptr As dword, ByVal invptr As dword) AS DWORD
   DIM pVtbl AS Events_IDispatchVtbl PTR
   DIM pUnk AS Events_IDispatchVtbl PTR

   pVtbl = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, SIZEOF(*pVtbl))
   IF pVtbl = 0 THEN EXIT FUNCTION
   pVtbl->QueryInterface   = QryPtr
   pVtbl->AddRef           = ProcPTR(Events_AddRef)
   pVtbl->Release          = ProcPTR(Events_Release)
   pVtbl->GetTypeInfoCount = ProcPTR(Events_GetTypeInfoCount)
   pVtbl->GetTypeInfo      = ProcPTR(Events_GetTypeInfo)
   pVtbl->GetIDsOfNames    = ProcPTR(Events_GetIDsOfNames)
   pVtbl->Invoke           = InvPtr
   pVtbl->pVtblAddr        = pVtbl
   pVtbl->pthis            = pthis
   pUnk = VARPTR(pVtbl->pVtblAddr)
   FUNCTION = pUnk
End Function