'******************************************
'remove match string from source string
'******************************************
Function str_remove(ByRef source As String,byref match As String)As String
	dim as long s=1,lmt
	Dim As String txt
	txt=source	
	lmt=Len(match)
	Do
		s=InStr(s,txt,match)
		If s Then txt=Left(txt,s-1)+Right(txt,Len(txt)-lmt-(s-1))
	Loop While s
	Function=txt
end function

'***********************************************************************
'remove any character in Match string from stxt string
'***********************************************************************
Function Str_removeany(ByRef source As String,ByRef match As String)As String
	dim as long s=1
	Dim As String txt
	txt=source
	Do
		s=Instr(s,txt,Any match)
		If s Then txt=Left(txt,s-1)+Right(txt,Len(txt)-1-(s-1))
	Loop While s
	Function=txt
end Function

'***************************************************************
'count number of parse string, separated by delimiter of source
'***************************************************************
Function str_numparse(ByRef source as string,ByRef delimiter as string)as long
	Dim As Long s=1,c,l
	l=Len(delimiter)
	Do
		s=instr(s,source,delimiter)
		If s then
			c+=1
			s+=l
		end if 
	Loop While s
	Function=c+1
end function
                    
'************************************************************
'parse source string, indexed by delimiter string at idx
'************************************************************                    
Function str_parse(ByRef source As String,Byref delimiter As String,ByVal idx As Long)As String
	Dim As Long s=1,c,l
	l=Len(delimiter)
	Do
		If c=idx-1 then Return mid(source,s,instr(s,source,delimiter)-s)
		s=instr(s,source,delimiter)
		If s then
			c+=1
			s+=l
		end if 
	Loop While s
End Function

'**********************************************************
'find index of match string as parse source by delimiter
'**********************************************************
Function str_parsepos(ByRef match As String,byref source As String,ByRef delimiter As String)As Integer
	Dim As Integer i
	For i=1 To str_numparse(source,delimiter)
		If str_parse(source,delimiter,i)=match Then Return i
	Next
	Function=0
End Function

'************************************************************************
'count number of parse string, separated by any character in delimiter string
'************************************************************************
Function str_numparseany(ByRef source As String,Byref delimiter As String)As Long
	Dim As Long s=1,c
	Do
		s=Instr(s,source,Any delimiter)
		If s Then
			c+=1
			s+=1
		End If 
	Loop While s
	Function=c+1
End Function

'********************************************************************
'parse source string, separated by any character in delimiter string
'********************************************************************                    
Function str_parseany(ByRef source As String,Byref delimiter As String,ByVal idx As Long)As String
	Dim As Long s=1,c
	Do
		If c=idx-1 Then Return Mid(source,s,Instr(s,source,Any delimiter)-s)
		s=Instr(s,source,Any delimiter)
		If s Then
			c+=1
			s+=1
		End If 
	Loop While s
End Function

'**********************************************************
'find index of match string as parse source by delimiter
'**********************************************************
Function str_parseanypos(ByRef match As String,byref source As String,ByRef delimiter As String)As Integer
	Dim As Integer i
	For i=1 To str_numparseany(source,delimiter)
		If str_parseany(source,delimiter,i)=match Then Return i
	Next
	Function=0
End Function

'*****************************************************
'count occurance of delimiter string in source string
'*****************************************************
function str_tally(ByRef source as string,byref delimiter as string)as long
	Dim As Long s=1,c,l=Len(delimiter)
	Do
		s=instr(s,source,delimiter)
		If s then
			c+=1
			s+=l
		end if 
	Loop While s
	Function=c
End function
                                                                                         
'**************************************************************
'count occurance of any character in del string in cmd string
'**************************************************************
function str_tallyany(ByRef source as string,byref delimiter as string)as long
	Dim As Long s=1,c
	Do
		s=InStr(s,source,any delimiter)
		If s then
			c+=1
			s+=1
		end if 
	Loop While s
	Function=c
End function
                                                               
'***********************************************************
'replace first occurance of match string with replace string in source string
'***********************************************************
Function str_replace(ByRef match As String,Byref sreplace As String,Byref source As String)As String
	Dim As Long s=1,c=len(sreplace),l=len(match)
	Dim As String text
	text=source
	Do
		s=InStr(s,text,match)
		If s then
			text=Mid$(text,1,s-1)+sreplace+Mid$(text,s+l,-1)
			s+=c
		end if 
	Loop While s
	Function=text
End Function
                                                 
'****************************************************************
'replace first occurance of any character in del string with r string in cmd string
'*****************************************************************                                                 
Function Str_replaceany(ByRef match As String,Byref sreplace As String,Byref source As String)As String
	Dim As Long s=1,c=len(sreplace)
	Dim As String text
	text=source
	Do 
		s=InStr(s,text,Any match)
		If s then
			text=Mid$(text,1,s-1)+sreplace+Mid$(text,s+1,-1)
			s+=c
		end if 
	Loop While s
	Function=text
End Function

'****************************************************************
'replace all occurance of any character in del string with r string in cmd string
'*****************************************************************                                                 
Function str_ReplaceAll(ByRef match As String,ByRef sreplace As String,ByRef source As String)As String
	Dim text As String
	text=source
	While InStr(text,match)
		text=str_replace(match,sreplace,text)	
	Wend
	Function=text
End Function

'****************************************************************
'replace all occurance of any character in del string with r string in cmd string
'*****************************************************************                                                 
Function str_ReplaceAnyAll(ByRef match As String,ByRef sreplace As String,ByRef source As String)As String
	Dim text As String
	text=source
	While InStr(text,Any match)
		text=str_replaceany(match,sreplace,text)	
	Wend
	Function=text
End Function
