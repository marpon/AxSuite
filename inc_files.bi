#IFNDEF INCFILE_PLUS
   #DEFINE INCFILE_PLUS

   Function MakeFileImport_(TabU1 as ubyte ptr, IncFil_len As Uinteger, NewFile As String) As integer
      Dim lFile AS short
      lFile = FreeFile                           'find next open file number
      Open NewFile For Binary As #lFile          'create the new file
      Put #lFile,, TabU1[0], IncFil_len          'write the buffer to the file
      Close #lFile                               'close the file
      Function = 1                               'worked so return a 1
   End Function

	function StringImport_(TabU1 as ubyte ptr, IncFil_len As Uinteger, Newvar As String)as string
      dim x as integer
		for x=0 to IncFil_len
			Newvar[x]=TabU1[x]
      NEXT
		function = Newvar
   End function

	Function mem_ClipboardSetText(TabU1 as ubyte ptr, IncFil_len As Uinteger) As Long
		dim pstr as zstring ptr
		pstr=TabU1
		Var hGlobalClip = GlobalAlloc(GMEM_MOVEABLE Or GMEM_SHARE, IncFil_len + 1)
		If OpenClipboard(0) Then
			EmptyClipboard()
			Var lpMem = GlobalLock(hGlobalClip)
			If lpMem Then
				CopyMemory(lpMem, pstr, IncFil_len)
				GlobalUnlock(lpMem)
				Function = cast(long, SetClipboardData(CF_TEXT, hGlobalClip))
			End If
			CloseClipboard()
		else
			Function = 0
		End If
	End Function

   #MACRO IncFilePlus(label, file, sectionName, attr)
      #if __FUNCTION__ <> "__FB_MAINPROC__"
         #IfnDef Var__##label##__Import__
            'Dim label As Const UByte Ptr = Any
            Dim label As UByte Ptr = Any
            Dim label##_len As UInteger = Any
         #endif
         #if __FB_DEBUG__
            asm jmp .LT_END_OF_FILE_##label##_DEBUG_JMP
         #else
            ' Switch to/Create the specified section
            #if attr = ""
               asm .section sectionName
            #else
               asm .section sectionName , attr
            #endif
         #endif

         ' Assign a label to the beginning of the file
         asm .LT_START_OF_FILE_##label#:
         asm __##label##__start = .
         ' Include the file
         asm .incbin ##file
         ' Mark the end of the the file
         asm __##label##__len = . - __##label##__start
         asm .LT_END_OF_FILE_##label:
         ' Pad it with a NULL Integer (harmless, yet useful for text files)
         asm .LONG 0
         #if __FB_DEBUG__
            asm .LT_END_OF_FILE_##label##_DEBUG_JMP:
         #else
            ' Switch back to the .text (code) section
            asm .section .text
            asm .balign 16
         #endif
         asm .LT_SKIP_FILE_##label:
         asm mov dword ptr [label ] , offset .LT_START_OF_FILE_##label#
         asm mov dword ptr [label##_len ] , offset __##label##__len

      #else
         #IfnDef Var__##label##__Import__
            Extern "c"
               Extern label As UByte Ptr
               Extern label##_len As UInteger
            End Extern
         #endif
         #if __FB_DEBUG__
            asm jmp .LT_END_OF_FILE_##label##_DEBUG_JMP
         #else
            ' Switch to/create the specified section
            #if attr = ""
               asm .section sectionName
            #else
               asm .section sectionName , attr
            #endif
         #endif
         ' Assign a label to the beginning of the file
         asm .LT_START_OF_FILE_##label#:
         asm __##label##__start = .
         ' Include the file
         asm .incbin ##file
         ' Mark the end of the the file
         asm __##label##__len = . - __##label##__start
         asm .LT_END_OF_FILE_##label:
         ' Pad it with a NULL Integer (harmless, yet useful for text files)
         asm .LONG 0
         asm label:
         asm .int .LT_START_OF_FILE_##label#
         asm label##_len:
         asm .int __##label##__len
         #if __FB_DEBUG__
            asm .LT_END_OF_FILE_##label##_DEBUG_JMP:
         #else
            ' Switch back to the .text (code) section
            asm .section .text
            asm .balign 16
         #endif
         asm .LT_SKIP_FILE_##label:
      #endif

   #endmacro

   #macro IncFileMem(label, file)
      IncFilePlus(label, file,.data, "")         'Use the .data (storage) section (SHARED)
      'IncFilePlus(label, file, firo, "") ' or what you want
   #endmacro

   #macro ImportFile(label, file, NewFile)
      #define Var__##label##__Import__
      #if __FUNCTION__ <> "__FB_MAINPROC__"
         'Dim label As Const Ubyte Ptr = Any
         Dim label As Ubyte Ptr = Any
         Dim label##_len As Uinteger = Any
      #Else
         Extern "c"
            Extern label As Ubyte Ptr
            Extern label##_len As Uinteger
         End Extern
      #endif
      IncFilePlus(label, file,.data, "")         'Use the .data (storage) section (SHARED)
      MakeFileImport_(label, label##_len, NewFile)
   #endmacro

#endif



'function innit1()as any Ptr
' IncFileMem(EXE1, "babygrid1.exe")
' 'print EXE1
' return cast(any ptr ,EXE1)
'END function


IncFileMem(inc1,"ATL.DLL")			'MakeFile(lpFile_inc1, "new path", is1_inc1)
IncFileMem(inc2,"ATL71.DLL")			'MakeFile(lpFile_inc2, "new path", is1_inc2)
'IncFileMem(inc3,"Ax_Disp_Lib.bi")		'MakeFile(lpFile_inc3, "new path", is1_inc3)
'IncFileMem(inc4,"LibAx_Disp.a")		'MakeFile(lpFile_inc4, "new path", is1_inc4)
IncFileMem(inc5,"atlCtl.dll")
IncFileMem(inc6,"atl71Ctl.dll")

'IncFileMem(inc7,"Ax_Lite_Lib.bi")		'MakeFile(lpFile_inc3, "new path", is1_inc3)
'IncFileMem(inc8,"LibAx_Lite.a")		'Mak
IncFileMem(inc9,"Ax_Lite.bi")		'Mak

'IncFileMem(inc10,"libatl.dll.a")		'Mak
'IncFileMem(inc11,"libatl71.dll.a")		'Mak

'MakeFileImport_(inc1, inc1_len, NewFile)
'MakeFile(lpFile_i1, "new path", is1_i1)  'for file 1


Function RegisterServer(hWnd As hwnd, filePath As String, Mode As integer=1)as integer
  Dim hLib As HMODULE
  Dim ProcAd As any ptr
  Dim sMod As String
  Dim res As integer

  if mode=0 THEN
	  sMod = "DllUnregisterServer"
  else
	  sMod = "DllRegisterServer"
  END IF

  hLib = LoadLibrary(filePath)
  ProcAd = GetProcAddress(hLib, sMod)

  res = CallWindowProc(ProcAd, hwnd, 0, 0, 0)

  FreeLibrary hLib
  Function = cast(integer,res)

End Function
