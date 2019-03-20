Dim Shared putEvtList As Sub(s As String)
Dim Shared clrEvtList As Sub()
Dim Shared EnumvTable As Function(hTree As Long,hParent() As long)As ZString Ptr
Dim Shared EnumInvoke As Function(hTree As Long,hParent() As long)As ZString Ptr
Dim Shared EnumEvent As Function(hTree As Long,hParent() As long)As ZString Ptr
Dim Shared EnumEnum As Function(hTree As Long,hParent() As long)As ZString Ptr
Dim Shared EnumRecord As Function(hTree As Long,hParent() As long)As ZString Ptr
Dim Shared EnumUnion As Function(hTree As Long,hParent() As long)As ZString Ptr
Dim Shared EnumModule As Function(hTree As Long,hParent() As long)As ZString Ptr
Dim Shared EnumAlias As Function(hTree As Long,hParent() As long)As ZString Ptr
Dim Shared EnumCoClass As Function(hTree As Long,hParent() As long)As ZString Ptr