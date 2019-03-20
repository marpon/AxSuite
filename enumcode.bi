Sub mkConstant
	Dim b As zString Ptr
	b=EnumEnum(htree,hparent())	
	sreport =*b:sysfreestring(b)
	b=EnumUnion(htree,hparent())
	sreport &=*b:sysfreestring(b)
	b=EnumAlias(htree,hparent())
	sreport &=*b:sysfreestring(b)
	b=EnumRecord(htree,hparent())
	sreport &=*b:sysfreestring(b)
	showinfo
End Sub

Sub mkModule
	Dim b As zString Ptr
	b=EnumModule(htree,hparent())	
	sreport =*b:sysfreestring(b)
	showinfo	
End sub

Sub mkvtable
	Dim b As zString Ptr
	b=EnumCoClass(htree,hparent())
	sreport =*b:sysfreestring(b)
	b=EnumvTable(htree,hparent())
	sreport &=*b:sysfreestring(b)
	showinfo	
End Sub

Sub mkInvoke
	Dim b As zString Ptr
	b=EnumCoClass(htree,hparent())
	sreport =*b:sysfreestring(b)
	b=EnumInvoke(htree,hparent())
	sreport &=*b:sysfreestring(b)
	showinfo	
End Sub

Sub mkEvent
	Dim b As zString Ptr
	b=EnumEvent(htree,hparent())
	sreport =*b:sysfreestring(b)
	showinfo	
End Sub