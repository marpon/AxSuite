/'
_BEGIN_RC_  'début de bloc Rc   placer un '  devant _BEGIN_RC_     , pour ne pas prendre en compte

#define Icon1 500
#define IDR_MENU1 10000
#define tlb_1 10001
#define tlb_2 10002
#define tlb_3 10003

IDR_MENU1 MENU
BEGIN
  POPUP "TypeLib"
  BEGIN
    MENUITEM "Enum Reged Lib",tlb_1
    MENUITEM "Open Lib File",tlb_2
    MENUITEM "Exit",tlb_3
  END
END
Icon1 ICON DISCARDABLE "tlb.ico"
_END_RC_  'fin de bloc Rc     placer un '  devant _END_RC_  , pour ne pas prendre en compte
'/
'fin de bloc Rc
#define Icon1 500
#define IDR_MENU1 10000
#define tlb_1 10001
#define tlb_2 10002
#define tlb_3 10003

#Include Once "windows.bi"
Dim Shared AS Hwnd but1
Dim Shared AS Hwnd but2
Dim Shared AS Hwnd edit1
Dim Shared AS Hwnd hwnd
Dim Shared AS Hwnd Label1


Declare Function WinMain(ByVal hInstance As HINSTANCE)As Integer

sub tlb_11()
	MessageBox(NULL, "Tlb1!", "Menu", MB_OK Or MB_ICONERROR)
end sub

sub tlb_22()
	MessageBox(NULL, "Tlb1!", "Menu", MB_OK Or MB_ICONERROR)
end sub

sub tlb_33()
	MessageBox(NULL, "Tlb1!", "Menu", MB_OK Or MB_ICONERROR)
end sub

End WinMain(GetModuleHandle(NULL))

Function WndProc(ByVal hWnd As HWND, _
         ByVal message As UINT, _
         ByVal wParam As WPARAM, _
         ByVal lParam As LPARAM) As LRESULT

   Function = 0
   dim buff                 AS string *255
	dim hDC as HDC

   Select Case(message)

      Case WM_COMMAND

         If HiWord(WPARAM) = BN_CLICKED Then
            If LoWord(WPARAM) = 992 Then
               MessageBox(NULL, "Print Button1 Pushed", "buff", MB_OK Or MB_ICONINFORMATION)
            End If
            If LoWord(WPARAM) = 994 Then

               MessageBox(NULL, "Print Button2 Pushed", "buff", MB_OK Or MB_ICONERROR)
            End If
         END IF


      Case WM_CLOSE
         MessageBox(NULL, "Fermeture du programme", "cloture", MB_OK Or MB_ICONINFORMATION)
         DestroyWindow(hwnd)
      Case WM_DESTROY
         MessageBox(NULL, "Fermeture effective", "Fin", MB_OK Or MB_ICONINFORMATION)
         PostQuitMessage 0

   End Select

   Function = DefWindowProc(hWnd, message, wParam, lParam)
End Function


Function WinMain(ByVal hInstance As HINSTANCE) As Integer


   dim as string WINDOW_CLASS = "WindowClass"
   Dim wMsg                 AS MSG
   Dim wcls                 AS WndClass
   DIM szAppName            AS string

   'dim AS string Icon2 = "E:/Autoit_Portable/App/Aut2Exe/Icons/epub.ico"

   Function = 0

	With wcls
		.style = CS_HREDRAW Or CS_VREDRAW Or CS_DBLCLKS or CS_GLOBALCLASS
		.lpfnWndProc = Cast(WNDPROC, @WndProc)
		.hInstance = hInstance
		.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(Icon1))
		.hCursor = LoadCursor(hInstance, IDC_ARROW)
		.hbrBackground = Cast(HBRUSH, COLOR_BTNFACE + 1)
		.lpszMenuName = Null
		.lpszClassName = strptr(WINDOW_CLASS)
	End With


	If RegisterClass(@wcls) = FALSE Then
		MessageBox(NULL, "RegisterClass('WindowClass') a echoué!", "Erreur!", MB_OK Or MB_ICONERROR)
		Exit FUNCTION
	End If



	hwnd = CreateWindowEx(WS_EX_CLIENTEDGE, WINDOW_CLASS, "win", WS_OVERLAPPED or WS_SYSMENU or WS_MINIMIZEBOX, 350, 250, 400, 300, NULL, NULL, NULL, NULL)


	but1 = CreateWindowEx(NULL, "Button", "button1", WS_VISIBLE Or WS_CHILD or BS_DEFPUSHBUTTON , 50, 180, 120, 35, hwnd, Cast(HMENU, 992), NULL, NULL)
	but2 = CreateWindowEx(NULL, "Button", "button1", WS_VISIBLE Or WS_CHILD, 220, 180, 120, 35, hwnd, Cast(HMENU, 994), NULL, NULL)
	edit1 = CreateWindowEx(NULL, "Edit", "edit", WS_VISIBLE Or WS_CHILD or WS_BORDER or SS_SUNKEN or SS_CENTER, 50, 140, 290, 24, hwnd, Cast(HMENU, 996), NULL, NULL)
	Label1 = CreateWindowEx(NULL, "Static", "info" , WS_VISIBLE Or WS_CHILD or SS_CENTER or SS_SUNKEN, 50, 40, 290, 80, hwnd, Cast(HMENU, 998), NULL, NULL)


	ShowWindow(hwnd, 1)

	Dim AS MSG uMsg

	While GetMessage(@uMsg, NULL, NULL, NULL) <> FALSE
		TranslateMessage(@uMsg)
		DispatchMessage(@uMsg)
	Wend
	UnregisterClass strptr(WINDOW_CLASS), hInstance

	FUNCTION = uMsg.wParam

End Function





