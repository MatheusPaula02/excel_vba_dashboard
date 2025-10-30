VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

Sheet1.Select

Sheet1.unprotect Password:="1234"
Sheet1.Protect Password:="1234", AllowFiltering:=True, DrawingObjects:=False

MsgBox "Por favor, certifique-se de inserir as informações de maneira correta na tabela de Registro de Cantoneira, para que não ocorra erros futuros.", vbInformation, "Registro Cantoneira"

MsgBox "Fazer o bloqueio das células utilizando a mesma senha após ter sido feitas as alterações.", vbInformation, "Registro Cantoneira"

MsgBox "A senha para desbloquer as células para fazer as alterações é: 1234.", vbInformation, "Registro Cantoneira"

Registro_de_Cantoneiras.Show

End Sub
