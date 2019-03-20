#Include Once "windows.bi"

	declare Function Filesize(str1 as string)as integer
	declare Function MakeFile(lRet2 AS HGLOBAL, NewFile As String, dret1 as long) As integer

#ifndef __incany_bi__
	#define __incany_bi__

	#macro MemFileAny(_nfile , strFile , _outfile)
		dim is1##_nfile      as long = Filesize(strFile)
		dim lpFile##_nfile   As any ptr

		asm
			.section .text
			jmp .end_##lpFile##_nfile
			.section .data
			.align 16
			.start_##lpFile##_nfile:
			.incbin ##strFile
			.section .text
			.align 16
			.end_##lpFile##_nfile:
			lea eax, .start_##lpFile##_nfile
			mov dword ptr [lpFile##_nfile], eax
		end asm

		MakeFile(lpFile##_nfile, _outfile, is1##_nfile)
	#endmacro

#endif ' __incany_bi__


Function Filesize(str1 as string)as integer
	Dim f As short
	f = FreeFile
	if Open (str1 For Binary Access Read As  #f)=0 then
		return (LOF(f))
		Close #f
	else
		return 0
	end if
END FUNCTION

Function MakeFile(lRet2 AS HGLOBAL, NewFile As String, dret1 as long) As integer
   dim                   AS ubyte ptr dRet2
   Dim lFile             AS short
   dRet2 = LockResource(lRet2)

	print len(dRet2)
   lFile = FreeFile                              'find next open file number
   Open NewFile For Binary As #lFile             'create the new file
		Put #lFile,, dRet2[0], dRet1                  'write the buffer to the file
   Close #lFile                                  'close the file
   Function = 1                                  'worked so return a 1
End Function

'dim str1 as string="babygrid5.exe"
'dim str2 as string="DEF1-Editor.exe"
'MemFileAny(1,"babygrid1.exe",str1)
'MemFileAny(2,"DEF-Editor.exe",str2)