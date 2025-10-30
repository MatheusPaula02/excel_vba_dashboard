VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Registro_de_Cantoneiras 
   Caption         =   "Registro de Cantoneiras"
   ClientHeight    =   5295
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5625
   OleObjectBlob   =   "Registro_de_Cantoneiras.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Registro_de_Cantoneiras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Registrar_Click()

Sheet1.unprotect Password:="1234"

Linha = Range("B2").End(xlDown).Row + 1

If Motorista = "" Or IsNull(Motorista) Then
MsgBox "Campo de preenchimento obrigat�rio", vbCritical, "Aten��o"
Motorista.SetFocus
Motorista.BackColor = 10092543
Else

If Data = "" Or IsNull(Data) Then
MsgBox "Campo de preenchimento obrigat�rio", vbCritical, "Aten��o"
Data.SetFocus
Data.BackColor = 10092543
Else

If Quantidade = "" Or IsNull(Quantidade) Then
MsgBox "Campo de preenchimento obrigat�rio", vbCritical, "Aten��o"
Quantidade.SetFocus
Quantidade.BackColor = 10092543
Else

If Transportadora = "" Or IsNull(Transportadora) Then
MsgBox "Campo de preenchimento obrigat�rio", vbCritical, "Aten��o"
Transportadora.SetFocus
Transportadora.BackColor = 10092543
Else

If Placa_Cavalo = "" Or IsNull(Placa_Cavalo) Then
MsgBox "Campo de preenchimento obrigat�rio", vbCritical, "Aten��o"
Placa_Cavalo.SetFocus
Placa_Cavalo.BackColor = 10092543
Else

If Placa_Carreta = "" Or IsNull(Placa_Carreta) Then
MsgBox "Campo de preenchimento obrigat�rio", vbCritical, "Aten��o"
Placa_Carreta.SetFocus
Placa_Carreta.BackColor = 10092543
Else

If Conferente = "" Or IsNull(Conferente) Then
MsgBox "Campo de preenchimento obrigat�rio", vbCritical, "Aten��o"
Conferente.SetFocus
Conferente.BackColor = 10092543
Else

If M�s = "" Or IsNull(M�s) Then
MsgBox "Campo de preenchimento obrigat�rio", vbCritical, "Aten��o"
M�s.SetFocus
M�s.BackColor = 10092543
Else

If Ano = "" Or IsNull(Ano) Then
MsgBox "Campo de preenchimento obrigat�rio", vbCritical, "Aten��o"
Ano.SetFocus
Ano.BackColor = 10092543
Else


Sheet1.Cells(Linha, 2).Value = Data.Value
Sheet1.Cells(Linha, 6).Value = Quantidade.Value
Sheet1.Cells(Linha, 7).Value = Transportadora.Value
Sheet1.Cells(Linha, 8).Value = Placa_Cavalo.Value
Sheet1.Cells(Linha, 9).Value = Placa_Carreta.Value
Sheet1.Cells(Linha, 10).Value = Motorista.Value
Sheet1.Cells(Linha, 11).Value = Conferente.Value
Sheet1.Cells(Linha, 3).Value = M�s.Value
Sheet1.Cells(Linha, 4).Value = Ano.Value
If Cobrar.Value = True Then
Sheet1.Cells(Linha, 5) = "Cobrar"
Else
Sheet1.Cells(Linha, 5) = "Reposi��o"
End If
If Devolu��o.Value = True Then
Sheet1.Cells(Linha, 5) = "Devolu��o"
End If

Unload Registro_de_Cantoneiras

Sheet1.Protect Password:="1234", AllowFiltering:=True, DrawingObjects:=False

MsgBox "Registro feito e salvo com sucesso", vbInformation, "Registro de Cantoneira"

ActiveWorkbook.Save

ActiveWorkbook.RefreshAll

MsgBox "DashBoard Atualizada", vbInformation, "Registro de Cantoneira"
End If
End If
End If
End If
End If
End If
End If
End If
End If

End Sub


Private Sub UserForm_Initialize()

Transportadora.RowSource = "Sheet2!C2:C30"
Conferente.RowSource = "Sheet2!G2:G12"
M�s.RowSource = "Sheet2!I2:I13"

End Sub
