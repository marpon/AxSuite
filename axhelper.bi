#Include once"win/ocidl.bi"

'covert string to bstr
'please follow with sysfreestring(bstr) after use to avoid memory leak
Function StringToBSTR(cnv_string As String) As BSTR
    Dim sb As BSTR
    Dim As Integer n
    n = (MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, cnv_string, -1, NULL, 0))-1
    sb=SysAllocStringLen(sb,n)
    MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, cnv_string, -1, sb, n)
    Return sb
End Function

'return numeric value from variant
Function Variantv(ByRef v As variant)As Double
	Dim vvar As variant
	VariantChangeTypeEx(@vvar,@v,NULL,VARIANT_NOVALUEPROP,VT_R8)
	Return vvar.dblval
End Function

'return string value from variant
Function Variants(ByRef v As variant)As String
	Dim vvar As variant
	VariantChangeTypeEx(@vvar,@v,NULL,VARIANT_NOVALUEPROP,VT_BSTR)
	Function=*Cast(wstring Ptr,vvar.bstrval)	
End Function

'return bstr value from variant
Function Variantb(ByRef v As variant)As bstr
	Dim vvar As variant
	VariantChangeTypeEx(@vvar,@v,NULL,VARIANT_NOVALUEPROP,VT_BSTR)
	Function=vvar.bstrval	
End Function

'assign variant with value
Declare sub vlet OverLoad(v As variant,x As variant)
Declare sub vlet OverLoad(v As variant,x As string)
Declare sub vlet OverLoad(v As variant,x As Longint)
Declare sub vlet OverLoad(v As variant,x As Ulongint)
Declare sub vlet OverLoad(v As variant,x As Double)

sub vlet(v As variant,x As variant)
	v.vt=vt_variant:v.pvarval=@x
End sub

sub vlet(v As variant,x As string)
	v.vt=vt_bstr:v.bstrval=stringtobstr(x)
End sub

sub vlet(v As variant,x As LongInt)
	v.vt=vt_i8:v.llval=x
End Sub

sub vlet(v As variant,x As ULongInt)
	v.vt=vt_ui8:v.ullval=x
End Sub

sub vlet(v As variant,x As single)
	v.vt=vt_r4:v.fltval=x
End Sub

Sub vlet(v As variant,x As double)
	v.vt=vt_r8:v.dblval=x
End Sub

'assign value to variant ptr
Declare function vptr OverLoad(x As variant)As Variant Ptr
Declare Function vptr OverLoad(x As string)As Variant Ptr
Declare Function vptr OverLoad(x As Longint)As Variant Ptr
Declare Function vptr OverLoad(x As Ulongint)As Variant Ptr
Declare Function vptr OverLoad(x As Double)As Variant Ptr

Function vptr(x As variant)As Variant Ptr
	static v As variant
	v.vt=vt_variant:v.pvarval=@x
	Return @v
End Function

Function vptr(x As string)As Variant Ptr
	Static v As variant
	v.vt=vt_bstr:v.bstrval=stringtobstr(x)
	Return @v
End Function

Function vptr(x As LongInt)As Variant Ptr
	Static v As variant
	v.vt=vt_i8:v.llval=x
	Return @v
End Function

Function vptr(x As ULongInt)As Variant Ptr
	Static v As variant
	v.vt=vt_ui8:v.ullval=x
	Return @v
End Function

Function vptr(x As double)As Variant Ptr
	Static v As variant
	v.vt=vt_r8:v.dblval=x
	Return @v
End Function

'assign variant as pointer to a variable
Declare sub vplet OverLoad(byref v As variant,Byref x As variant)
Declare sub vplet OverLoad(byref v As variant,Byref x As bstr)
Declare sub vplet OverLoad(ByRef v As variant,Byref x As byte)
Declare sub vplet OverLoad(ByRef v As variant,Byref x As ubyte)
Declare sub vplet OverLoad(ByRef v As variant,Byref x As short)
Declare sub vplet OverLoad(ByRef v As variant,Byref x As ushort)
Declare sub vplet OverLoad(ByRef v As variant,Byref x As Integer)
Declare sub vplet OverLoad(ByRef v As variant,Byref x As Uinteger)
'Declare sub vplet OverLoad(ByRef v As variant,Byref x As Longint)
'Declare sub vplet OverLoad(ByRef v As variant,Byref x As Ulongint)
Declare sub vplet OverLoad(ByRef v As variant,Byref x As single)
Declare sub vplet OverLoad(ByRef v As variant,Byref x As Double)

sub vplet(ByRef v As variant,Byref x As variant)
	v.vt=vt_byref Or vt_variant
	v.pvarval=@x
End sub

sub vplet(ByRef v As variant,Byref x As bstr)
	v.vt=vt_byref Or vt_bstr:v.pbstrval=@x
End sub

sub vplet(ByRef v As variant,ByRef x As Byte)
	v.vt=vt_byref Or vt_i1:v.pbval=@x
End Sub

sub vplet(ByRef v As variant,ByRef x As UByte)
	v.vt=vt_byref Or vt_ui1:v.pbval=@x
End Sub

sub vplet(ByRef v As variant,ByRef x As Short)
	v.vt=vt_byref Or vt_i2:v.pival=@x
End Sub

sub vplet(ByRef v As variant,ByRef x As UShort)
	v.vt=vt_byref Or vt_ui2:v.puival=@x
End Sub

sub vplet(ByRef v As variant,Byref x As Integer)
	v.vt=vt_byref Or vt_i4:v.plval=@x
End Sub

sub vplet(ByRef v As variant,Byref x As UInteger)
	v.vt=vt_byref Or vt_ui4:v.pulval=@x
End Sub
/'
sub vplet(ByRef v As variant,Byref x As LongInt)
	v.vt=vt_byref Or vt_i8:v.pllval=@x
End Sub

sub vplet(ByRef v As variant,Byref x As ULongInt)
	v.vt=vt_byref Or vt_ui8:v.pullval=@x
End Sub
'/
sub vplet(ByRef v As variant,Byref x As single)
	v.vt=vt_byref Or vt_r4:v.pfltval=@x
End Sub

Sub vplet(ByRef v As variant,Byref x As double)
	v.vt=vt_byref Or vt_r8:v.pdblval=@x
End Sub
