VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm038 
   Caption         =   "Frasortering"
   ClientHeight    =   7850
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11328
   OleObjectBlob   =   "frm038.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm038"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub ComboBox2_Change()
    Call TextBox1_Change
End Sub


Private Sub ComboBox4_Change()
    Call TextBox2_Change
End Sub



Public Sub OKButton_Click()
           
    ' Validering for forkert anvendelse af f�r/efter
    
    If ComboBox2.Value = "efter" And ComboBox4.Value = "f�r" Then
        dFunc.msgError = "Forkert anvendelse af f�r/efter"
        SFunc.ShowFunc ("frmMsg")
        GoTo ending
    End If
    
    ' Validering for numeriske v�rdier
    
    Dim cControl As Control
        
    For Each cControl In Me.Controls
        
        control_type = UCase(Left(cControl.Name, 4))
            
        If control_type = "TEXT" Then
           If cControl.Text = "" Then
              cControl.SetFocus
              dFunc.msgError = "Felt skal udfyldes med tal."
              SFunc.ShowFunc ("frmMsg")
              GoTo ending
           End If
           If cControl.Text <> "" Then
              If IsNumeric(cControl.Text) = False Then
                 cControl.SetFocus
                 dFunc.msgError = "Felt skal udfyldes med tal."
                 SFunc.ShowFunc ("frmMsg")
                 GoTo ending
              End If
           End If
        End If
        
    Next cControl
    
    ' Validering for 'efter'
    
    If ComboBox2.Value = "efter" Then
        If Int(TextBox1.Value) > Int(TextBox2.Value) Then
            dFunc.msgError = "V�rdien i 'Fra' skal v�re mindre end v�rdien i 'Til'."
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    Dim antal As Integer
    
    Dim x1 As Variant
    Dim x2 As Variant
    
    ' Reset values
    
    Call Insert_to_sheet("Regler", "J22:O22", "")
    
    'Relationen mellem "stiftelsesdato" og "periode slut"
    
    x1 = TextBox1.Value
    x2 = TextBox2.Value
    
     ' 'F�r' fra foranstilles med minus
    If ComboBox2.Value = "f�r" Then
        x1 = "-" + x1
    End If
    
    ' 'F�r' fra foranstilles med minus
    If ComboBox4.Value = "f�r" Then
        x2 = "-" + x2
    End If
    
    ' Validering for 'f�r'
    
    If ComboBox2.Value = "f�r" Then
        If Int(x1) > Int(x2) Then
            dFunc.msgError = "V�rdien i 'Fra' skal v�re mindre end v�rdien i 'Til'."
            SFunc.ShowFunc ("frmMsg")
            'MsgBox ("V�rdien i 'Fra' skal v�re mindre end v�rdien i 'Til'.")
            GoTo ending
        End If
    End If
    
    ' Validering af 'Stiftelsesdato' kan ligge fra 10 dage f�r til 1081 dage efter 'Periode slut'.
    
    If ComboBox2.Value = "f�r" Then
        If (Int(TextBox1.Value) > 10) Then
            dFunc.msgError = "'Stiftelsesdato' kan minimalt ligge 10 dage f�r 'Periode slut'."
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    If ComboBox2.Value = "f�r" And ComboBox4.Value = "efter" Then
        If (Int(TextBox2.Value) > 1081) Then
            dFunc.msgError = "'Stiftelsesdato' kan maksimalt ligge 1081 dage efter 'Periode slut'."
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    If ComboBox2.Value = "efter" And ComboBox4.Value = "efter" Then
        If (Int(TextBox2.Value) - Int(TextBox1.Value) > 1081) Then
            dFunc.msgError = "'Stiftelsesdato' kan maksimalt ligge 1081 dage efter 'Periode slut'. "
            SFunc.ShowFunc ("frmMsg")
            GoTo ending
        End If
    End If
    
    ' Inds�t v�rdier i  regel
    Call Insert_to_sheet("Regler", "J22:J22", x1)
    Call Insert_to_sheet("Regler", "M22:M22", x2)
    
    ' Aktiver regler
    Call Insert_to_sheet("Regler", "G22:G22", "JA")
    
    ' Skriv svar ned i 'SpmSvar'
    
    ' Relationen mellem "stiftelsesdato" og "periode slut"
    a = "Stiftelsesdato"
    b = "Periode slut"
    VisuTitle = a & " i forhold til " & b
    Worksheets("SpmSvar").Range("C64:C64").Value = VisuTitle
    Worksheets("SpmSvar").Range("D64:D64").Value = TextBox1.Value
    Worksheets("SpmSvar").Range("E64:E64").Value = "dage"
    Worksheets("SpmSvar").Range("F64:F64").Value = ComboBox2.Value
    Worksheets("SpmSvar").Range("G64:G64").Value = TextBox2.Value
    Worksheets("SpmSvar").Range("H64:H64").Value = "dage"
    Worksheets("SpmSvar").Range("I64:I64").Value = ComboBox4.Value
    
    Me.Hide
    SFunc.ShowFunc ("frm021")
       
ending:
End Sub

Private Sub TextBox1_Change()
    
    Count = TextBox1.Value
    DMY = "dage"
    FE = ComboBox2.Value
    
    If Not IsNumeric(Count) Then
        Exit Sub
    End If
    
    x = Count
        
    Worksheets("SpmSvar").Range("L2") = x
    
    Worksheets("SpmSvar").Range("K2") = Count & " " & DMY & " " & FE
    
    If FE = "efter" Then
    
        Worksheets("SpmSvar").Range("L2") = x
        
    ElseIf FE = "f�r" Then
    
        Worksheets("SpmSvar").Range("L2") = -x
        
    End If
    
    Call DrawChart
    
    
End Sub

Private Sub TextBox2_Change()
    
    Count = TextBox2.Value
    DMY = "dage"
    FE = ComboBox4.Value
    
    If Not IsNumeric(Count) Then
        Exit Sub
    End If
    
    x = Count
    
    Worksheets("SpmSvar").Range("L4") = x
    
    Worksheets("SpmSvar").Range("K4") = Count & " " & DMY & " " & FE
    
    If FE = "efter" Then
    
        Worksheets("SpmSvar").Range("L4") = x
        
    ElseIf FE = "f�r" Then
    
        Worksheets("SpmSvar").Range("L4") = -x
        
    End If
    
    Call DrawChart

End Sub


Private Sub DrawChart()

    Dim Fname As String

    Call SaveChart
    Fname = ThisWorkbook.Path & "\temp1.gif"
    Me.Image2.Picture = LoadPicture(Fname)
    Call DeleteFile
    
End Sub

Private Sub SaveChart()

    Dim MyChart As Chart
    Dim Fname As String

    Set MyChart = Sheets("SpmSvar").ChartObjects(1).Chart
    Fname = ThisWorkbook.Path & "\temp1.gif"
    MyChart.Export Filename:=Fname, FilterName:="GIF"
    
End Sub

Sub DeleteFile()

    Dim Fname As String
    On Error Resume Next
    Fname = ThisWorkbook.Path & "\temp1.gif"
    Kill Fname
    On Error GoTo 0
    
End Sub



Public Sub Tilbage_Click()
    Me.Hide
    SFunc.ShowFunc ("frm036")
End Sub

Private Sub UserForm_Initialize()

    Image1.PictureSizeMode = fmPictureSizeModeStretch
    ' Activate sheet
    Worksheets("SpmSvar").Activate
    ActiveWindow.Zoom = 80
    Worksheets("SpmSvar").Range("I1").Select

    a = "Stiftelsesdato"
    b = "Periode slut"
    VisuTitle = a & " i forhold til " & b
    
    Worksheets("SpmSvar").Range("K3") = b
    Worksheets("SpmSvar").Range("J1") = VisuTitle
    Worksheets("SpmSvar").Range("K2") = "20 dage efter"
    Worksheets("SpmSvar").Range("L2") = 20
    Worksheets("SpmSvar").Range("K4") = "30 dage efter"
    Worksheets("SpmSvar").Range("L4") = 30
    
    Call DrawChart
    
    With ComboBox2
        .AddItem "f�r"
        .AddItem "efter"
    End With
    
    With ComboBox4
        .AddItem "f�r"
        .AddItem "efter"
    End With

    
    ' Indl�s tidligere svar fra 'SpmSvar'

    ' Relationen mellem forfaldsdato" og "sidste rettidige betalingsdato"
    TextBox1.Value = Worksheets("SpmSvar").Range("D64:D64").Value
    If Not IsEmpty(Worksheets("SpmSvar").Range("F64:F64").Value) Then ComboBox2.Value = Worksheets("SpmSvar").Range("F64:F64").Value
    TextBox2.Value = Worksheets("SpmSvar").Range("G64:G64").Value
    If Not IsEmpty(Worksheets("SpmSvar").Range("I64:I64").Value) Then ComboBox4.Value = Worksheets("SpmSvar").Range("I64:I64").Value

End Sub

