Attribute VB_Name = "dFunc"
Public msgError As String
Public msgYesNo As String
Public msgYesNoTxt As String
Function FOKO_Retracer()
    ' FOKO s�ttes op som i Retracer:
    Call Func.Insert_to_sheet("Regler", "J43:J47", "")
    Call Func.Insert_to_sheet("Regler", "G43:G47", "JA")
    
    ' RIM
    Call Func.Insert_to_sheet("Population", "B16:B16", "NEJ")
    'Call Func.Insert_to_sheet("Population", "B17:B17", "NEJ") ' Sp�rg Patrick

    'frm014.Show

End Function
