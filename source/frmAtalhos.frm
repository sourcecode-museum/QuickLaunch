VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmAtalhos 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5010
   Icon            =   "frmAtalhos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.ListView lstApp 
      Height          =   2400
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   4233
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "imgApp"
      SmallIcons      =   "imgApp"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin ComctlLib.ImageList imgApp 
      Left            =   75
      Top             =   2475
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmAtalhos.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAtalhos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long

Private Type PictDesc
  cbSizeofStruct As Long
  picType As Long
  hImage As Long
  xExt As Long
  yExt As Long
End Type
Private Type Guid
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0 To 7) As Byte
End Type

Private Sub Form_Load()
    Dim sPathINI As String
    
    Dim iFor        As Integer
    Dim nAppQtd     As Integer
    Dim sAppDesc    As String
    Dim sAppPath    As String
    
    Dim Icone       As StdPicture
    
    sPathINI = App.Path & "\Config.ini"
    nAppQtd = LerINI("CONFIG", "APP_QTD", sPathINI, 0)
    
    For iFor = 1 To nAppQtd
        sAppDesc = LerINI("CONFIG", "APP_DESC" & iFor, sPathINI)
        sAppPath = LerINI("CONFIG", "APP_PATH" & iFor, sPathINI)
        
        'Verificando existencia do arquivo
        If Dir(sAppPath) <> "" Then
            Set Icone = New StdPicture
            Set Icone = ObterIconeApp(sAppPath)
            
            'Salvando o icone temporario do Aplicativo
            Call SavePicture(Icone, App.Path & "\temp~.ico")

            'Inserindo o icone na lista
            imgApp.ListImages.Add , , LoadPicture(App.Path & "\temp~.ico")

            'Deletanto o icone temporario
            Kill App.Path & "\temp~.ico"

            
            'Incluindo atalho na lista
            lstApp.ListItems.Add(, , sAppDesc, imgApp.ListImages.Count, imgApp.ListImages.Count).Tag = sAppPath
        End If
    Next
    
    lstApp.ListItems.Add(, , "Sair", 1, 1).Tag = "Sair"
    
    Set Icone = Nothing
End Sub

Function ObterIconeApp(ByVal pPathApp As String) As Picture
On Error GoTo TrataErro
    Dim NewPic As Picture
    Dim PicConv As PictDesc
    Dim IGuid As Guid
    Dim hIcon As Long

    hIcon = ExtractIcon(0, pPathApp, 0)
    
    PicConv.cbSizeofStruct = Len(PicConv)
    PicConv.picType = vbPicTypeIcon
    PicConv.hImage = hIcon
    
    IGuid.Data1 = &H20400
    IGuid.Data4(0) = &HC0
    IGuid.Data4(7) = &H46

    With IGuid
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With

    DoEvents
    
    Call OleCreatePictureIndirect(PicConv, IGuid, False, NewPic)
    
    DoEvents
    Set ObterIconeApp = NewPic
    
    Call DestroyIcon(hIcon)
    
    Exit Function
TrataErro:
    MsgBox "Não foi possível ler os modulos para apresentação.", vbCritical, "Modulos"
End Function

Private Sub Form_Resize()
    lstApp.Move Me.Left + 150, Me.Top + 150, Me.Width - 300, Me.Height - 300
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAtalhos = Nothing
    End
End Sub

Private Sub lstApp_ItemClick(ByVal Item As ComctlLib.ListItem)
    If UCase(lstApp.Tag) <> "" Then
        If LCase(Item.Tag) = "sair" Then
            Unload Me
        ElseIf Not Dir(Item.Tag, vbArchive) = "" Then
            Call Shell(Item.Tag, vbNormalFocus)
        End If
        lstApp.Tag = ""
    End If
End Sub

Private Sub lstApp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        lstApp.Tag = "OK"
        lstApp_ItemClick lstApp.SelectedItem
    End If
    
    lstApp.Tag = ""
End Sub

Private Sub lstApp_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbLeftButton Then
    lstApp.Tag = "OK"
  End If
End Sub
