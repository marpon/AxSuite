
Dim Shared As Byte On_Off(9), ORDER(4)
Dim Shared As Integer Screen_Width = 320, Screen_height = 240
Dim As Integer i, cnt
screenres Screen_width, Screen_height, 32

'
'
Sub Draw_Boxes()
Dim As Uinteger C_bck, C_fore, x, y
Dim As Integer i,j,  Size
c_bck = rgb(255,255,255)
c_fore = rgb(255,0,0)
Size = 8
Line (0,0)-(screen_width-1,Screen_Height-1), C_Bck, BF
x = 100
y = 100
For i = 1 To 3
    For j = 1 To 3
        If ON_OFF ((i-1)*3+j) Then Line (x + (j-1) * size *1.5, y + (i-1)*Size * 1.5) - Step (size, size), C_Fore, BF
    Next j
Next i
End Sub
'
'
Sub ALL_ON
For i As Integer = 1 To 9
    ON_Off(i) = 1
Next i
End Sub


WindowTitle "Waiting.... hit ESC to quit"
ORDER(1) = 1
ORDER(2) = 3
ORDER(3) = 9
ORDER(4) = 7

ALL_ON
CNT = 1
Do
    Select Case ORDER(CNT)
    Case 1,7
            ON_OFF(ORDER(CNT)) = 0
            ON_OFF(ORDER(CNT)+1) = 0
            If ORDER(CNT) = 1 Then 
                ON_OFF(ORDER(CNT)+3) = 0
            Else
                ON_OFF(ORDER(CNT)-3) = 0
            End If
    Case 3,9
            ON_OFF(ORDER(CNT)) = 0
            ON_OFF(ORDER(CNT)-1) = 0
            If ORDER(CNT) = 3 Then
                ON_OFF(ORDER(CNT)+3) = 0
            Else
                ON_OFF(ORDER(CNT)-3) = 0
            End If
    End Select
    Draw_Boxes
    CNT += 1
    If CNT > 4 Then CNT = 1
    Sleep 80
    ALL_ON      'reset
Loop Until multikey(1)
 