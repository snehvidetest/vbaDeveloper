VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm032 
   Caption         =   "UserForm1"
   ClientHeight    =   9675
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   13368
   OleObjectBlob   =   "frm032.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm032"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Public Sub CheckBox1_Click()
If CheckBox1.Value = True Then
    TextBox1.Value = ""
    TextBox1.Enabled = False
ElseIf CheckBox1.Value = False Then
    TextBox1.Enabled = True
End If

End Sub

Public Sub CheckBox2_Click()
If CheckBox2.Value = True Then
    TextBox2.Value = ""
    TextBox2.Enabled = False
    CheckBox3.Enabled = False
ElseIf CheckBox2.Value = False Then
    TextBox2.Enabled = True
    CheckBox3.Enabled = True
End If
End Sub


Public Sub CheckBox3_Click()
If CheckBox3.Value = True Then
    CheckBox2.Enabled = False
    TextBox2.Value = "1095"
    TextBox2.Enabled = False
ElseIf CheckBox3.Value = False Then
    CheckBox2.Enabled = True
    TextBox2.Value = ""
    TextBox2.Enabled = True
End If

End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal x As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label3_Click()

End Sub

Public Sub OKButton_Click()
If (CheckBox1.Value = True Or CheckBox2.Value = True) And frm007.OptionButton3.Value = True Then
    dFunc.msgError = "Sp�rgeskemaet kan ikke anvendes p� baggrund af indtastede oplysninger"
    SFunc.ShowFunc ("frmMsg")
    Me.Hide
    SFunc.ShowFunc ("frm002")
    GoTo ending
End If

' Validering af ingen optionbuttons valgt
If OptionButton1.Value = False And OptionButton2.Value = False Then
dFunc.msgError = "V�lg venligst �n af svar mulighederne for at g� videre."
    SFunc.ShowFunc ("frmMsg")
    GoTo ending
End If

' Validering - Negative v�rdier
If TextBox1.Value < 0 Then
    dFunc.msgError = "Der kan ikke indtastes negative v�rdier i antal dage"
    SFunc.ShowFunc ("frmMsg")
    GoTo ending
End If

' Validering - gyldig tal v�rdi i antal dage tekstfelter
If IsNumeric(TextBox1.Value) = False And CheckBox1.Value = False Then
    dFunc.msgError = "Inds�t en gyldig v�rdi i antal dage"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Inds�t en gyldig v�rdi i antal dage")
    GoTo ending
End If

If IsNumeric(TextBox2.Value) = False And CheckBox2.Value = False Then
    dFunc.msgError = "Inds�t en gyldig v�rdi i antal dage"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Inds�t en gyldig v�rdi i antal dage")
    GoTo ending
End If

' Validering - Antal dage tekstboks felter skal udfyldes
If IsEmpty(TextBox1.Value) = True And CheckBox1.Value = False Then
    dFunc.msgError = "Inds�t en v�rdi i antal dage"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Inds�t en v�rdi i antal dage")
    GoTo ending
End If

If IsEmpty(TextBox2.Value) = True And CheckBox2.Value = False Then
    dFunc.msgError = "Inds�t en v�rdi i antal dage"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("Inds�t en v�rdi i antal dage")
    GoTo ending
End If


' Validering p� mindst �n af OptionButtons skal v�re udfyldt
If OptionButton1.Value = False And OptionButton2.Value = False Then
    dFunc.msgError = "V�lg venligst hvor begyndelsestidspunktet for for�ldelsesfristen skal beregnes"
    SFunc.ShowFunc ("frmMsg")
    'MsgBox ("V�lg venligst hvor begyndelsestidspunktet for for�ldelsesfristen skal beregnes")
    GoTo ending
End If

' Advarsels popup hvis "Ved ikke" er valgt
'If CheckBox1.Value = True Then
'    dFunc.msgError = "RIM kan ikke beregne et tidligst muligt for�ldelsestidspunkt for populationen."
'    SFunc.ShowFunc ("frmMsg")
'
'    ' Gruppe 1 deaktiveres
'    Worksheets("Gruppering").Range("C2:C2").Value = "NEJ"
'
'    ' Populations arket �ndres
'    Worksheets("Population").Range("B17:B17").Value = "NEJ"
'
'    'MsgBox ("RIM kan ikke beregne et tidligst muligt for�ldelsestidspunkt for populationen.")
'    GoTo ending
'End If

' Antal dage skrives i Varighed_X for SRB datoen
If OptionButton1.Value = True And (CheckBox1.Value = False And CheckBox2.Value = False) Then
    Worksheets("Regler").Range("J64:J64").Value = CInt(TextBox2.Value) - CInt(TextBox1.Value)
    Worksheets("Regler").Range("J65:J65").Value = CInt(TextBox2.Value) - CInt(TextBox1.Value)
    Worksheets("Regler").Range("J66:J66").Value = CInt(TextBox2.Value) - CInt(TextBox1.Value)
    Worksheets("Regler").Range("J67:J67").Value = CInt(TextBox2.Value) - CInt(TextBox1.Value)
    Worksheets("Regler").Range("J72:J72").Value = CInt(TextBox2.Value) - CInt(TextBox1.Value)
ElseIf OptionButton2.Value = True And (CheckBox1.Value = False And CheckBox2.Value = False) Then
    Worksheets("Regler").Range("J64:J64").Value = CInt(TextBox1.Value) + CInt(TextBox2.Value)
    Worksheets("Regler").Range("J65:J65").Value = CInt(TextBox1.Value) + CInt(TextBox2.Value)
    Worksheets("Regler").Range("J66:J66").Value = CInt(TextBox1.Value) + CInt(TextBox2.Value)
    Worksheets("Regler").Range("J67:J67").Value = CInt(TextBox1.Value) + CInt(TextBox2.Value)
    Worksheets("Regler").Range("J72:J72").Value = CInt(TextBox1.Value) + CInt(TextBox2.Value)
End If

If CheckBox1.Value = False Or CheckBox2.Value = False Then

' Populations arket �ndres
Worksheets("Population").Range("B17:B17").Value = "NEJ"

' Reglerne aktiveres
Worksheets("Regler").Range("G64:G64").Value = "NEJ"
Worksheets("Regler").Range("G65:G65").Value = "NEJ"
Worksheets("Regler").Range("G66:G66").Value = "NEJ"
Worksheets("Regler").Range("G67:G67").Value = "NEJ"
Worksheets("Regler").Range("G72:G72").Value = "NEJ"

Else

' Populations arket �ndres
Worksheets("Population").Range("B17:B17").Value = "JA"

' Reglerne deaktiveres
Worksheets("Regler").Range("G64:G64").Value = "JA"
Worksheets("Regler").Range("G65:G65").Value = "JA"
Worksheets("Regler").Range("G66:G66").Value = "JA"
Worksheets("Regler").Range("G67:G67").Value = "JA"
Worksheets("Regler").Range("G72:G72").Value = "JA"
End If

' Grupper aktiveres
If OptionButton1.Value = True Or OptionButton2.Value = True Then
Worksheets("Gruppering").Range("C2:C2").Value = "JA"
End If

' Indsaet vaerdier i SpmSvar

Worksheets("SpmSvar").Range("C92:C92").Value = Label3.Caption
Worksheets("SpmSvar").Range("C94:C94").Value = Label4.Caption


If OptionButton1.Value = True Then
    Worksheets("SpmSvar").Range("C93:C93").Value = Label7.Caption
Else
    Worksheets("SpmSvar").Range("C93:C93").Value = Label1.Caption
End If

If OptionButton1.Value = True Then
    Worksheets("SpmSvar").Range("D92:D92").Value = "F�r det valgte stamdatafelt"
ElseIf OptionButton2.Value = True Then
    Worksheets("SpmSvar").Range("D92:D92").Value = "Samme dag eller senere end det valgte stamdatafelt"
End If

' Antal dage skrives ned i arket

If TextBox1.Value <> "" Then
    Worksheets("SpmSvar").Range("D93:D93").Value = CInt(TextBox1.Value)
End If
    
If TextBox2.Value <> "" Then
    Worksheets("SpmSvar").Range("D94:D94").Value = CInt(TextBox2.Value)
End If

' "Ved ikke" skrives ned i arket
If CheckBox1.Value = True Then
    Worksheets("SpmSvar").Range("D93:D93").Value = "Ved ikke"
End If

If CheckBox2.Value = True Then
    Worksheets("SpmSvar").Range("D94:D94").Value = "Ved ikke"
End If
    


If CheckBox1.Value = True Or CheckBox2.Value = True Then
    dFunc.msgError = "RIM kan ikke beregne et tidligst muligt for�ldelsestidspunkt den del af populationen, hvor der ikke er indsendt FOKO. Den f�lgende konfiguration ang�r derfor kun fordringer, hvor der er indsendt FOKO"
    SFunc.ShowFunc ("frmMsg")
End If

' Tjek om gruppe 1 skal deaktiveres
If frm014.Forfaldsdato.Value = True And (frm028.CheckBox1.Value = True Or frm028.CheckBox2.Value = True) Then
    Worksheets("Gruppering").Range("C2:C2").Value = "NEJ"
End If
    
If frm014.SRB.Value = True And (frm032.CheckBox1.Value = True Or frm032.CheckBox2.Value = True) Then
    Worksheets("Gruppering").Range("C2:C2").Value = "NEJ"
End If
    
If frm014.Stiftelsesdato.Value = True And (frm029.CheckBox1.Value = True Or frm029.CheckBox2.Value = True) Then
    Worksheets("Gruppering").Range("C2:C2").Value = "NEJ"
End If
    
If frm014.PeriodeStartdato.Value = True And (frm030.CheckBox1.Value = True Or frm030.CheckBox2.Value = True) Then
    Worksheets("Gruppering").Range("C2:C2").Value = "NEJ"
End If
    
If frm014.PeriodeSlutdato.Value = True And (frm031.CheckBox1.Value = True Or frm031.CheckBox2.Value = True) Then
    Worksheets("Gruppering").Range("C2:C2").Value = "NEJ"
End If

If frm014.Stiftelsesdato.Value = True Then
    Me.Hide
    SFunc.ShowFunc ("frm029")
    'frm029.Show
ElseIf frm014.PeriodeStartdato.Value = True Then
    Me.Hide
    SFunc.ShowFunc ("frm030")
    'frm030.Show
ElseIf frm014.PeriodeSlutdato.Value = True Then
    Me.Hide
    SFunc.ShowFunc ("frm031")
    'frm031.Show
Else
    Me.Hide
    SFunc.ShowFunc ("frm039")
    'frm039.Show
End If
ending:

End Sub

Private Sub OptionButton1_Click()
If OptionButton1.Value = True Then
    Label1.Visible = True
    Label2.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = False
    Label7.Visible = False
    CheckBox1.Visible = True
    CheckBox2.Visible = True
    CheckBox3.Visible = True
    TextBox1.Visible = True
    TextBox2.Visible = True
    CheckBox1.Enabled = True
    CheckBox2.Enabled = True
    TextBox1.Enabled = True
    TextBox2.Enabled = True
End If
End Sub

Private Sub OptionButton2_Click()
If OptionButton2.Value = True Then
    Label1.Visible = False
    Label2.Visible = False
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label7.Visible = True
    CheckBox1.Visible = True
    CheckBox2.Visible = True
    TextBox1.Visible = True
    TextBox2.Visible = True
    CheckBox1.Enabled = True
    CheckBox2.Enabled = True
    TextBox1.Enabled = True
    TextBox2.Enabled = True
End If
End Sub


Public Sub Tilbage_Click()
Me.Hide
SFunc.ShowFunc ("frm014")
'frm014.Show
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

Image1.PictureSizeMode = fmPictureSizeModeStretch

If OptionButton1.Value = False And OptionButton2.Value = False Then
    Label1.Visible = False
    Label2.Visible = False
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label7.Visible = False
    CheckBox1.Visible = False
    CheckBox2.Visible = False
    CheckBox3.Visible = False
    TextBox1.Visible = False
    TextBox2.Visible = False
Else
    Label1.Visible = True
    Label2.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label7.Visible = True
    CheckBox1.Visible = True
    CheckBox2.Visible = True
    CheckBox3.Visible = True
    TextBox1.Visible = True
    TextBox2.Visible = True
End If

    If Worksheets("SpmSvar").Range("D92:D92").Value = "Samme dag eller senere end det valgte stamdatafelt" Then
        OptionButton2.Value = True
    ElseIf Worksheets("SpmSvar").Range("D92:D92").Value = "F�r det valgte stamdatafelt" Then
        OptionButton1.Value = True
    End If
    
If Worksheets("SpmSvar").Range("D93:D93").Value = "Ved ikke" Then
    CheckBox1.Value = True
Else
    TextBox1.Value = Worksheets("SpmSvar").Range("D93:D93").Value
End If



If Worksheets("SpmSvar").Range("D94:D94").Value = "Ved ikke" Then
    CheckBox2.Value = True
Else
    TextBox2.Value = Worksheets("SpmSvar").Range("D94:D94").Value
End If
    
    Label12.Font.Size = 15
End Sub
