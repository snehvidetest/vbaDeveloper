VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm002 
   Caption         =   "Populationsafgr�nsning"
   ClientHeight    =   7560
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11292
   OleObjectBlob   =   "frm002.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cboFordringstype_Exit(ByVal Cancel As MSForms.ReturnBoolean)
   
If Not IsError(Application.Match(cboFordringstype, Worksheets("FID_FTYPE_Data").Range("C2:C7561"), 0)) Then
    ' Fordringstype findes
    tt = Application.Match(cboFordringstype, Worksheets("FID_FTYPE_Data").Range("C2:C7561"), 0) + 1
    ws = Worksheets("FID_FTYPE_Data").Range("D" & tt).Value
    lblFtypeTxt.Caption = ws
 
Else
    lblFtypeTxt.Caption = ""
End If

End Sub

Public Sub OKButton_Click()
'validering




If IsNumeric(txtFordringsId) Then
       'OK
       GoTo videre
Else
       dFunc.msgError = "FordringshaverID er forkert udfyldt"
       SFunc.ShowFunc ("frmMsg")
       'frmMsg.Show
       
       'MsgBox "FordringsId er forkert udfyldt"
       txtFordringsId.SetFocus
       GoTo ending
End If
 
videre:

If Len(txtFordringsId) > 4 Or Len(txtFordringsId) < 4 Then
        dFunc.msgError = "FordringshaverID er forkert udfyldt"
        SFunc.ShowFunc ("frmMsg")
        'frmMsg.Show
        'MsgBox "FordringsId er forkert udfyldt"
        txtFordringsId.SetFocus
        GoTo ending
End If

If Len(cboFordringstype) > 7 Or Len(cboFordringstype) < 7 Then
        dFunc.msgError = "FordringsType er forkert udfyldt"
        SFunc.ShowFunc ("frmMsg")
        'frmMsg.Show
        'MsgBox "FordringsType er forkert udfyldt"
        cboFordringstype.SetFocus
        GoTo ending
End If

' Validering af FordringshaverIf findes
 w_txtFordringsId = txtFordringsId
    
If Not IsError(Application.Match(Int(w_txtFordringsId), Worksheets("FID_TXT").Range("A2:A994"), 0)) Then
        ' FordringsId findes
Else
        lblFhaverTxt.Caption = ""
        dFunc.msgError = "FordringshaverID findes ikke"
        SFunc.ShowFunc ("frmMsg")
        'frmMsg.Show
        txtFordringsId.SetFocus
        GoTo ending
End If

' Validering af Fordringstype findes
If Not IsError(Application.Match(cboFordringstype, Worksheets("FID_FTYPE_Data").Range("E2:E474"), 0)) Then
    ' Fordringstype findes
Else
    lblFtypeTxt.Caption = ""
    dFunc.msgError = "Fordringstype findes ikke"
    SFunc.ShowFunc ("frmMsg")
    'frmMsg.Show
    cboFordringstype.SetFocus
    GoTo ending
End If

' Validering af FordringshaverID og Fordringstype kombination er valid
Dim searchFID_FTYPE As String

searchFID_FTYPE = txtFordringsId + cboFordringstype

If Not IsError(Application.Match(searchFID_FTYPE, Worksheets("FID_FTYPE_Data").Range("A2:A7561"), 0)) Then
    ' Kombination af FordringshaverID og Fordringstype findes
    tt = Application.Match(searchFID_FTYPE, Worksheets("FID_FTYPE_Data").Range("A2:A7561"), 0) + 1
    ws = Worksheets("FID_FTYPE_Data").Range("D" & tt).Value
    lblFtypeTxt.Caption = ws
Else
    dFunc.msgError = "Kombination af FordringshaverID og Fordringstype findes ikke."
    SFunc.ShowFunc ("frmMsg")
    'frmMsg.Show
    GoTo ending
End If

If txtModtStart.Value = "" Then
    dFunc.msgError = "Startdatoen for perioden skal udfyldes."
    SFunc.ShowFunc ("frmMsg")
    'frmMsg.Show

    'MsgBox ("Startdatoen for perioden skal udfyldes.")
    GoTo ending
End If

If txtModtStart.Value <> "" Then
    If FormatCheck(txtModtStart, "date") = False Then
        txtModtStart.SetFocus
        GoTo ending
    End If
End If

If txtModtSlut.Value <> "" Then
    If FormatCheck(txtModtSlut, "date") = False Then
        txtModtSlut.SetFocus
        GoTo ending
    End If
End If

If CDate(txtModtStart.Value) < CDate("1 - 9 - 2013") Then
    SFunc.ShowFunc ("frm043")
    'frm043.Show
'    MsgBox ("Fordringshaver har indtastet et begyndelsestidspunkt for modtagelsesperioden, der ligger f�r den 1. september 2013. Som udgangspunkt vil vi ikke konfigurere fordringer, der er modtaget f�r den 1. september 2013, da der er risiko for, at fordringer modtaget f�r den 1. september 2013 har mistet data i forbindelse med konverteringen til EFI/DMI. S�fremt der i populationsafgr�nsningen v�lges en modtagelsesperiode med start f�r den 1. september 2013, skal det afd�kkes, om konverteringen af den afgr�nsede population har medf�rt �ndringer i fordringernes data.")
    GoTo ending
End If

If txtModtSlut.Value <> "" Then
    If CDate(txtModtSlut.Value) < CDate("1 - 9 - 2013") Then
        dFunc.msgError = "Perioden kan ikke ligge f�r den 1. September 2013."
        SFunc.ShowFunc ("frmMsg")
        'frmMsg.Show
        'MsgBox ("Perioden kan ikke ligge f�r den 1. September 2013.")
        txtModtSlut.Value = ""
        GoTo ending
    End If
End If

If txtModtSlut.Value <> "" Then
    If CDate(txtModtStart.Value) > CDate(txtModtSlut.Value) Then
        dFunc.msgError = "Slutperioden kan ikke ligge f�r startperioden."
        SFunc.ShowFunc ("frmMsg")
        'frmMsg.Show
        'MsgBox ("Slutperioden kan ikke ligge f�r startperioden")
        GoTo ending
    End If
End If

If forkertData.Value = False And korrektData.Value = False Then
    dFunc.msgError = "Udfyld venligst sp�rgsm�l omkring fordringshavers registreringspraksis for at fors�tte."
    SFunc.ShowFunc ("frmMsg")
    'frmMsg.Show
    'MsgBox ("Udfyld venligst sp�rgsm�l omkring fordringshavers registreringspraksis for at fors�tte")
    GoTo ending
End If
'If Len(txtModtStart) <> 10 Or Len(txtModtSlut) <> 10 Then
'    MsgBox "Modtagelsesperiode er forkert udfyldt"
'    txtModtStart.SetFocus
'    GoTo ending
'End If


'alt er ok - klar til at opdatere konfiguration

'gem data
Worksheets("SpmSvar").Range("C2:C2").Value = Controls("Label1").Caption
Worksheets("SpmSvar").Range("C3:C3").Value = Controls("Label2").Caption
Worksheets("SpmSvar").Range("C4:C4").Value = Controls("spg3").Caption
Worksheets("SpmSvar").Range("C5:C5").Value = Controls("Label5").Caption

Worksheets("SpmSvar").Range("D2:D2").Value = txtFordringsId.Value
Worksheets("SpmSvar").Range("D3:D3").Value = cboFordringstype.Value
Worksheets("SpmSvar").Range("D4:D4").Value = txtModtStart.Value
Worksheets("SpmSvar").Range("E4:E4").Value = txtModtSlut.Value

If forkertData.Value = True Then
    Worksheets("SpmSvar").Range("D5:D5").Value = "Ja"
ElseIf korrektData.Value = True Then
    Worksheets("SpmSvar").Range("D5:D5").Value = "Nej"
End If

Worksheets("Population").Range("B2:B2").Value = txtFordringsId.Value
Worksheets("Population").Range("B3:B3").Value = cboFordringstype.Value
Worksheets("Population").Range("B4:B4").Value = txtModtStart.Value
Worksheets("Population").Range("B5:B5").Value = txtModtSlut.Value

Me.Hide

' Worksheets("Konfiguration").Activate
' Activate sheet

If forkertData = False Then
    SFunc.ShowFunc ("frm003")
    'frm003.Show
Else
    SFunc.ShowFunc ("frm005")
    'frm005.Show
End If

ending:
End Sub

Public Sub Tilbage_Click()

Me.Hide

SFunc.ShowFunc ("frm001")
'frm001.Show

End Sub

Public Sub txtFordringsId_Exit(ByVal Cancel As MSForms.ReturnBoolean)

If Len(txtFordringsId) = 1 Then
    txtFordringsId.Value = "000" + txtFordringsId.Value
ElseIf Len(txtFordringsId) = 2 Then
    txtFordringsId.Value = "00" + txtFordringsId.Value
ElseIf Len(txtFordringsId) = 3 Then
    txtFordringsId.Value = "0" + txtFordringsId.Value
End If

If IsNumeric(txtFordringsId) Then
     
    w_txtFordringsId = txtFordringsId
    
    If Not IsError(Application.Match(Int(w_txtFordringsId), Worksheets("FID_TXT").Range("A2:A994"), 0)) Then
        ' Fordringstype findes
        tt = Application.Match(Int(w_txtFordringsId), Worksheets("FID_TXT").Range("A2:A994"), 0) + 1
        ws = Worksheets("FID_TXT").Range("B" & tt).Value
        lblFhaverTxt.Caption = ws
    Else
        lblFhaverTxt.Caption = ""
    End If

End If

End Sub

Private Sub txtModtStart_Change()

End Sub

Public Sub UserForm_Initialize()
'Fill JA/NEJ ComboBox
'Controls("Label1").Caption = Worksheets("SpmSvar").Range("B2:B2").Value
Image1.PictureSizeMode = fmPictureSizeModeStretch
' Reset all values ?
Worksheets("Population").Range("B2:B18").Value = ""


With cboFordringstype
.AddItem "AFGD�KN"
.AddItem "AFTILL�"
.AddItem "B�ADMIN"
.AddItem "B�ASKAT"
.AddItem "B�BVMEA"
.AddItem "B�B�DIC"
.AddItem "B�B�KON"
.AddItem "B�DAGB�"
.AddItem "B�DAGFP"
.AddItem "B�DAGSK"
.AddItem "B�FOLKE"
.AddItem "B�GR�NL"
.AddItem "B�NORDI"
.AddItem "B�OVUDS"
.AddItem "B�P�LAG"
.AddItem "B�P�LVD"
.AddItem "B�SKATK"
.AddItem "B�SKATT"
.AddItem "B�TVPLL"
.AddItem "CFCIVIL"
.AddItem "CFOMKCK"
.AddItem "DAEKSIS"
.AddItem "DAFRIAF"
.AddItem "DAGAVEA"
.AddItem "DAKAPIT"
.AddItem "DAKONTR"
.AddItem "DAOPH�V"
.AddItem "DAPAFGI"
.AddItem "DAPAFGR"
.AddItem "DAPROCE"
.AddItem "DAREALR"
.AddItem "DAREGMK"
.AddItem "DAR�VEJ"
.AddItem "DATIMK�"
.AddItem "DBFOLKF"
.AddItem "DFADMIN"
.AddItem "DFAFFIS"
.AddItem "DFAFF�D"
.AddItem "DFAKTIE"
.AddItem "DFAMBDR"
.AddItem "DFBERED"
.AddItem "DFB�RMV"
.AddItem "DFDEGSO"
.AddItem "DFDGMMA"
.AddItem "DFDOBBE"
.AddItem "DFDOMMA"
.AddItem "DFEFBNI"
.AddItem "DFEFDAG"
.AddItem "DFEFMER"
.AddItem "DFERSBI"
.AddItem "DFERSTA"
.AddItem "DFFAKTV"
.AddItem "DFFLEKS"
.AddItem "DFFMBFY"
.AddItem "DFFMTDI"
.AddItem "DFFMUDL"
.AddItem "DFFMUPE"
.AddItem "DFFORBR"
.AddItem "DFFOUDL"
.AddItem "DFFUBDP"
.AddItem "DFFUIHS"
.AddItem "DFGAFAL"
.AddItem "DFGAFLL"
.AddItem "DFGAFPL"
.AddItem "DFGL�DK"
.AddItem "DFGTAOM"
.AddItem "DFGTAOU"
.AddItem "DFG�LKO"
.AddItem "DFG�LST"
.AddItem "DFHUSLE"
.AddItem "DFINBEU"
.AddItem "DFINBIN"
.AddItem "DFINDFI"
.AddItem "DFKFEUB"
.AddItem "DFKFUBI"
.AddItem "DFKOMPG"
.AddItem "DFKONFI"
.AddItem "DFKVSUG"
.AddItem "DFKVSUK"
.AddItem "DFKVSUR"
.AddItem "DFKVSUU"
.AddItem "DFMDEGK"
.AddItem "DFMDEGR"
.AddItem "DFMDEGS"
.AddItem "DFMIBUB"
.AddItem "DFNOBIA"
.AddItem "DFOECDF"
.AddItem "DFOMRIS"
.AddItem "DFOPHOL"
.AddItem "DFPASK�"
.AddItem "DFPFFER"
.AddItem "DFPFMLG"
.AddItem "DFPFOHK"
.AddItem "DFREAHU"
.AddItem "DFREPAT"
.AddItem "DFSOCSI"
.AddItem "DFSPILD"
.AddItem "DFTILL�"
.AddItem "DFTILSB"
.AddItem "DFUDAFG"
.AddItem "DFURETF"
.AddItem "DFVANDL"
.AddItem "DGASAMB"
.AddItem "DGBLYAK"
.AddItem "DGCIVFO"
.AddItem "DGEKSPE"
.AddItem "DGGEGSP"
.AddItem "DGGPD�K"
.AddItem "DGLEVNE"
.AddItem "DGMILJ�"
.AddItem "DGVURDE"
.AddItem "DOADVOM"
.AddItem "DOSAGSO"
.AddItem "DSAKTIE"
.AddItem "DSAKTTK"
.AddItem "DSD�DSB"
.AddItem "DSEJEAV"
.AddItem "DSEJEV�"
.AddItem "DSEUREN"
.AddItem "DSKURSG"
.AddItem "DSMIDPE"
.AddItem "DSPENBS"
.AddItem "DSPENSA"
.AddItem "DSPENSK"
.AddItem "DSREJDV"
.AddItem "DSS�RPE"
.AddItem "DSUREOS"
.AddItem "DAAEOGS"
.AddItem "DAAFSON"
.AddItem "DAAKTIE"
.AddItem "DAARVEA"
.AddItem "FADPE�S"
.AddItem "FAMIARB"
.AddItem "FAUBMBE"
.AddItem "FAUBMDP"
.AddItem "FAUBMEA"
.AddItem "FAUBMEF"
.AddItem "FAUBMET"
.AddItem "FAUBMFD"
.AddItem "FAUBMFY"
.AddItem "FAUBMIA"
.AddItem "FAUBMIJ"
.AddItem "FAUBMIK"
.AddItem "FAUBMIV"
.AddItem "FAUBMKU"
.AddItem "FAUBMOA"
.AddItem "FAUBMOB"
.AddItem "FAUBMOU"
.AddItem "FAUBMOV"
.AddItem "FAUBMPY"
.AddItem "FAUBMSA"
.AddItem "FAUBMUA"
.AddItem "FAUBMUG"
.AddItem "FFBOLOP"
.AddItem "FFBYGGB"
.AddItem "FFBYGTL"
.AddItem "FFDIGEL"
.AddItem "FFEJDG�"
.AddItem "FFEJDSK"
.AddItem "FFFKVUD"
.AddItem "FFHEGNS"
.AddItem "FFJORDF"
.AddItem "FFLANDI"
.AddItem "FFPUMPE"
.AddItem "FFRENHV"
.AddItem "FFRENOH"
.AddItem "FFRENOV"
.AddItem "FFROTTE"
.AddItem "FFSKORS"
.AddItem "FFTINGL"
.AddItem "FFVANDF"
.AddItem "FFVANDL"
.AddItem "FFVEJBE"
.AddItem "FFVEJVL"
.AddItem "FOANKRF"
.AddItem "FOBORAH"
.AddItem "FOBORAV"
.AddItem "FOB�DFO"
.AddItem "FOEGVFT"
.AddItem "FOEGVVE"
.AddItem "FOERSFO"
.AddItem "FOKONFO"
.AddItem "FOOMKFO"
.AddItem "FOROTFO"
.AddItem "FOROT�R"
.AddItem "FOSAGFO"
.AddItem "FOSKOBR"
.AddItem "FOSPATV"
.AddItem "FOSTATA"
.AddItem "FOVAFTV"
.AddItem "FRG�TKO"
.AddItem "FUDIVER"
.AddItem "FUL�NIN"
.AddItem "FUMODRE"
.AddItem "FAAOUYD"
.AddItem "GEAFGOP"
.AddItem "GEBETAL"
.AddItem "GEGEBEJ"
.AddItem "GEINDDR"
.AddItem "GEINDSL"
.AddItem "GELOENI"
.AddItem "GELOENS"
.AddItem "GEOPKRS"
.AddItem "GEOPKR�"
.AddItem "GEOPREB"
.AddItem "GEOPRET"
.AddItem "GERYBET"
.AddItem "GERYKKE"
.AddItem "GETILSI"
.AddItem "KFBEBOE"
.AddItem "KFBEREG"
.AddItem "KFBILL�"
.AddItem "KFBOIST"
.AddItem "KFBOLIN"
.AddItem "KFBOLIR"
.AddItem "KFBOLRE"
.AddItem "KFBOLYD"
.AddItem "KFBOL�G"
.AddItem "KFBOL�N"
.AddItem "KFBOSIK"
.AddItem "KFBOSIR"
.AddItem "KFB�FIR"
.AddItem "KFB�FYD"
.AddItem "KFB�DPL"
.AddItem "KFDAGIN"
.AddItem "KFDAGIR"
.AddItem "KFFLYIR"
.AddItem "KFFLYTT"
.AddItem "KFFMTIR"
.AddItem "KFFMTPB"
.AddItem "KFFMUBT"
.AddItem "KFFMUIR"
.AddItem "KFFORIR"
.AddItem "KFFORUS"
.AddItem "KFG�TIR"
.AddItem "KFG�TKM"
.AddItem "KFG�TKO"
.AddItem "KFG�TKU"
.AddItem "KFHJEHJ"
.AddItem "KFKAUTI"
.AddItem "KFKILAS"
.AddItem "KFKIRKE"
.AddItem "KFKOMTA"
.AddItem "KFLOASO"
.AddItem "KFMUSIK"
.AddItem "KFPLEJE"
.AddItem "KFRENOV"
.AddItem "KFSERVL"
.AddItem "KFSKOLF"
.AddItem "KFSKROT"
.AddItem "KFSPILL"
.AddItem "KFTILIN"
.AddItem "KFUDDL�"
.AddItem "KFUDLEJ"
.AddItem "KFUREBI"
.AddItem "KFURSDP"
.AddItem "KFVAFBL"
.AddItem "KTBRUTT"
.AddItem "KTKTDOM"
.AddItem "KTNETTO"
.AddItem "LIERHVE"
.AddItem "LIMEDIE"
.AddItem "MOCDSKI"
.AddItem "MOCRMFO"
.AddItem "MOEJEBR"
.AddItem "MOGEALM"
.AddItem "MOGRPTL"
.AddItem "MOGR�EJ"
.AddItem "MOGR�UD"
.AddItem "MOHISPL"
.AddItem "MOM�NUM"
.AddItem "MOM�TOS"
.AddItem "MONORGR"
.AddItem "MONULKN"
.AddItem "MOOVF�P"
.AddItem "MOPRMRK"
.AddItem "MOPRREG"
.AddItem "MOREGMK"
.AddItem "MOSKGTP"
.AddItem "MOTRENU"
.AddItem "MOTSFRE"
.AddItem "MOUDLAF"
.AddItem "MOVEJBE"
.AddItem "MOV�GTA"
.AddItem "MO�NPLR"
.AddItem "OMINKAS"
.AddItem "OMKFORT"
.AddItem "PABKMDL"
.AddItem "PABOGAV"
.AddItem "PAB�RBA"
.AddItem "PACFCHF"
.AddItem "PACHOKO"
.AddItem "PACIGAP"
.AddItem "PACIGAR"
.AddItem "PACIGPP"
.AddItem "PACOAEL"
.AddItem "PACOAFG"
.AddItem "PACOANA"
.AddItem "PACOAOG"
.AddItem "PACOAVA"
.AddItem "PACOBEN"
.AddItem "PAEMBBP"
.AddItem "PAEMBVA"
.AddItem "PAENEOG"
.AddItem "PAENSER"
.AddItem "PAFEDTA"
.AddItem "PAFORBB"
.AddItem "PAGEBTP"
.AddItem "PAGL�DL"
.AddItem "PAGVSSP"
.AddItem "PAKAFFE"
.AddItem "PAKASIN"
.AddItem "PAKLOPM"
.AddItem "PAKONIS"
.AddItem "PAKULDI"
.AddItem "PAKULVA"
.AddItem "PAKV�LO"
.AddItem "PAKV�LS"
.AddItem "PALEDNV"
.AddItem "PALOTGV"
.AddItem "PALUDOM"
.AddItem "PALYSTF"
.AddItem "PAMETHA"
.AddItem "PAMIBUB"
.AddItem "PAMINFO"
.AddItem "PAMINVA"
.AddItem "PANICAB"
.AddItem "PAPUNKT"
.AddItem "PAPVCFO"
.AddItem "PAPVCFT"
.AddItem "PAREFSF"
.AddItem "PAR�STF"
.AddItem "PASKAFO"
.AddItem "PASPALO"
.AddItem "PASPFOR"
.AddItem "PASPGOP"
.AddItem "PASPIRI"
.AddItem "PASPKAS"
.AddItem "PASPKLA"
.AddItem "PASPLDV"
.AddItem "PASPLOT"
.AddItem "PASPPUL"
.AddItem "PASPTOT"
.AddItem "PASPUIN"
.AddItem "PASPVDL"
.AddItem "PASPV�D"
.AddItem "PASTEMP"
.AddItem "PASTENK"
.AddItem "PASVOVL"
.AddItem "PATINGL"
.AddItem "PATIPNI"
.AddItem "PATOBAK"
.AddItem "PAVANDV"
.AddItem "PAVINFR"
.AddItem "POANDEP"
.AddItem "POB�DFI"
.AddItem "POB�DGR"
.AddItem "POB�DIS"
.AddItem "POB�DNO"
.AddItem "POB�DNP"
.AddItem "POB�DPO"
.AddItem "POB�DSV"
.AddItem "POERSAN"
.AddItem "POERSPO"
.AddItem "POERSRP"
.AddItem "POKONPO"
.AddItem "POMILIT"
.AddItem "POOFBID"
.AddItem "POSAGPO"
.AddItem "POSKAB�"
.AddItem "POSKAPO"
.AddItem "POTOLAF"
.AddItem "POTVAPO"
.AddItem "POVANDY"
.AddItem "PSACARB"
.AddItem "PSARBGE"
.AddItem "PSARBMB"
.AddItem "PSARBRE"
.AddItem "PSBSKAT"
.AddItem "PSBSKRE"
.AddItem "PSFOREA"
.AddItem "PSRESRE"
.AddItem "PSRESTS"
.AddItem "PSS�RLI"
.AddItem "PSS�MAN"
.AddItem "PSTILL�"
.AddItem "PAAAF�L"
.AddItem "PAAFFLD"
.AddItem "PAAFVOP"
.AddItem "PAALKSV"
.AddItem "PAANSMK"
.AddItem "PAANTIB"
.AddItem "RECIVFO"
.AddItem "REENSKA"
.AddItem "REFORTR"
.AddItem "REINDDR"
.AddItem "RELICEN"
.AddItem "REL�BEJ"
.AddItem "REOPKIS"
.AddItem "REOPKI�"
.AddItem "REOPKRS"
.AddItem "REOPKR�"
.AddItem "REOPKTS"
.AddItem "REOPKT�"
.AddItem "REOPKVA"
.AddItem "REOPREB"
.AddItem "RERETEJ"
.AddItem "RETSAFG"
.AddItem "RFG�TRM"
.AddItem "RFG�TRU"
.AddItem "SFFORAG"
.AddItem "SGDETSG"
.AddItem "SGDETSS"
.AddItem "SGINFRS"
.AddItem "SGMFMAF"
.AddItem "SGMFMBO"
.AddItem "SGMFMIK"
.AddItem "SGMFMUD"
.AddItem "SGMFM�K"
.AddItem "SGMFMAA"
.AddItem "SGMISEU"
.AddItem "SGMISFM"
.AddItem "SGMISLA"
.AddItem "SGMISLC"
.AddItem "SGMISLD"
.AddItem "SGMISLE"
.AddItem "SGMISLF"
.AddItem "SGMISLG"
.AddItem "SGMISLI"
.AddItem "SGMISLJ"
.AddItem "SGMISLK"
.AddItem "SGMISLL"
.AddItem "SGOPSVU"
.AddItem "SGVOKSU"
.AddItem "TUTRANX"
.AddItem "TUUDL�G"
.AddItem "UBUDEUN"
.AddItem "UHEFORS"
.AddItem "UHFORSK"
.AddItem "UHKOMIN"
.AddItem "UHKONVE"
.AddItem "UHTILL�"
.AddItem "UH�GTEF"
.AddItem "underho"
.AddItem "VAMOMSE"
.AddItem "VARENFE"
.AddItem "VATOLDE"
.AddItem "VSARBMB"
.AddItem "VSARGBD"
.AddItem "VSASKAT"
.AddItem "VSDIVEN"
.AddItem "VSDIVMO"
.AddItem "VSFONAC"
.AddItem "VSFONSK"
.AddItem "VSFOUDM"
.AddItem "VSIMPOR"
.AddItem "VSKULAC"
.AddItem "VSL�AFG"
.AddItem "VSMOMSE"
.AddItem "VSRENFE"
.AddItem "VSRENTE"
.AddItem "VSROYAL"
.AddItem "VSSELAC"
.AddItem "VSSELSK"
.AddItem "VSSELSL"
.AddItem "VSSKAKR"
.AddItem "VSSKULB"
.AddItem "VSTOLDE"
.AddItem "VSUDBYT"
.AddItem "VSUDMSK"
.AddItem "VAARBMB"
.AddItem "VAARGBD"
.AddItem "VAASKAT"
End With


' Indl�s tidligere svar 1t4
txtFordringsId.Value = Worksheets("SpmSvar").Range("D2:D2").Value
cboFordringstype.Value = Worksheets("SpmSvar").Range("D3:D3").Value

If txtModtStart.Value = "" Then
    txtModtStart.Value = "01-09-2013"
Else
    txtModtStart.Value = CStr(Worksheets("SpmSvar").Range("D4:D4").Value)
End If

txtModtSlut.Value = CStr(Worksheets("SpmSvar").Range("E4:E4").Value)

If Worksheets("SpmSvar").Range("D5:D5").Value = "" Then
    forkertData.Value = False
    korrektData.Value = False
ElseIf Worksheets("SpmSvar").Range("D5:D5").Value = "Ja" Then
    forkertData.Value = True
ElseIf Worksheets("SpmSvar").Range("D5:D5").Value = "Nej" Then
    korrektData.Value = True
End If

End Sub
