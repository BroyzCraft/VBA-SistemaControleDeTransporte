VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formLogin 
   Caption         =   "LOGIN"
   ClientHeight    =   8445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6765
   OleObjectBlob   =   "formLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub box_login_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

'Altera a IMG
reset:
img.Picture = LoadPicture()
usuario = formLogin.box_login.Value
On Error Resume Next
caminhoImg = Sheets("login").Range("B:B").Find(usuario).Offset(0, 3)
caminhoArquivo = Environ("USERPROFILE") & "\Desktop\Sistema de Controle de Transporte\img\" & caminhoImg
img.Picture = LoadPicture(caminhoArquivo)

End Sub

Private Sub box_senha_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

'Valida Senha
Dim usuario, senha As String
usuario = formLogin.box_login.Value
senha = formLogin.box_senha.Value

ActiveWorkbook.Unprotect (123)
Sheets("login").Visible = True

e = True
On Error Resume Next
senhaCerta = Worksheets("login").Range("B:B").Find(usuario).Offset(0, 1)

If senha & "" <> senhaCerta & "" Then
    
    MsgBox ("Usuário ou Senha incorretos, Tente novamente!")
    formLogin.box_senha.Value = ""
    formLogin.box_login.Value = ""
    
Else:
    
    formLogin.Hide
    formMenu.Show
    Call scripts.log(usuario, "Logado com sucesso")
    
End If

Sheets("login").Visible = False
ActiveWorkbook.Protect (123)

End Sub

Private Sub bt_cancelar_Click()

'Application.Quit
'DEBUG
formLogin.Hide
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
 If CloseMode = vbFormControlMenu Then
  '  MsgBox "Preencha os dados e clique em OK", vbCritical, "AVISO"
     Cancel = True
End If
    
End Sub
