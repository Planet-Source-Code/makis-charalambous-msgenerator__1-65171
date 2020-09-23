Attribute VB_Name = "VBINPUT"
Sub IntKP(ThisControl As Control, Length%, KeyAscii As Integer)
    
    If Len(ThisControl.Text) >= Length% Then
        KeyAscii = 0
        Beep
        Exit Sub
    End If
    
    If Len(ThisControl.Text) = Length% Then
      If KeyAscii <> 8 Then
        KeyAscii = 0
        Beep
      End If
    Else
      c$ = Chr$(KeyAscii)
      Select Case c$
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", Chr$(8)
        Case "-"
          If ThisControl.SelStart <> 0 Or InStr(ThisControl.Text, "-") <> 0 Then
            KeyAscii = 0
            Beep
          End If
        Case Else
          KeyAscii = 0
          Beep
      End Select
    End If
End Sub

'Point2
'  Sub PointKP(ThisControl, Length%, DecimalNum,KeyAscii)
'  Function PointLF(ThisControl, Min$, Max$)
'  Accepts number with DecimalNum digits to the right of the decimal point.
'  Maximum length = Length%
'  Minimum value = Min$
'  Maximum value = Max$
'  Example 123.49
'
Sub PointKP(ThisControl As Control, Length%, DecimalNum%, KeyAscii As Integer)
    
    If KeyAscii = 13 Then Exit Sub ' Allow The ENTER key
    If KeyAscii = 8 Then Exit Sub ' Allow The BACKspace key
    
    If Len(ThisControl.Text) >= Length% Then
        KeyAscii = 0
        Beep
        Exit Sub
    End If
    
    If Len(ThisControl.Text) <> Length% Then
      StringLength% = Len(ThisControl.Text)
      DecimalPosition% = InStr(ThisControl.Text, ".")
      If StringLength% - DecimalPosition% = DecimalNum% And DecimalPosition% <> 0 Then
        If ThisControl.SelStart < DecimalPosition% Then
          Select Case KeyAscii
            Case 48 To 57
            Case 45 ' "-"
              If ThisControl.SelStart <> 0 Or InStr(ThisControl.Text, "-") <> 0 Then
                KeyAscii = 0
                Beep
              End If
            Case Else
              KeyAscii = 0
              Beep
          End Select
        ElseIf KeyAscii <> 8 Then
          KeyAscii = 0
          Beep
        End If
      Else
        Select Case KeyAscii
          Case 48 To 57
          Case 46 ' "."
            If InStr(ThisControl.Text, ".") <> 0 Then
              KeyAscii = 0
              Beep
            End If
          Case 45
            If ThisControl.SelStart <> 0 Or InStr(ThisControl.Text, "-") <> 0 Then
              KeyAscii = 0
              Beep
            End If
          Case Else
            KeyAscii = 0
            Beep
        End Select
      End If
    End If
End Sub

