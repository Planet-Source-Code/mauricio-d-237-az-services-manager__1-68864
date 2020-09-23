VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   4200
      Width           =   6975
      Begin VB.CommandButton Command5 
         Caption         =   "Instalar"
         Height          =   375
         Left            =   5280
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Remover"
         Height          =   375
         Left            =   4080
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Pausar"
         Height          =   375
         Left            =   2880
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Detener"
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Iniciar"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3240
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin MSComctlLib.ListView lv 
         Height          =   3015
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nombre"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre para mostrar"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Enumerar"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6735
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Command1_Click()
Dim S As New UpcServiciosyWea

On Error GoTo Mierda
S.ServiceName = lv.ListItems(lv.SelectedItem.Index).Text

 S.Start_Service
 Sleep (2000)
 Command6.Value = True
 
Mierda:
    
End Sub

Private Sub Command2_Click()
Dim S As New UpcServiciosyWea
On Error GoTo Mierda
S.ServiceName = lv.ListItems(lv.SelectedItem.Index).Text
S.Stop_Service
Sleep (2000)

Command6.Value = True

Mierda:

End Sub

Private Sub Command3_Click()
Dim S As New UpcServiciosyWea
On Error GoTo Mierda
S.ServiceName = lv.ListItems(lv.SelectedItem.Index).Text
If Command3.Caption = "Pausar" Then
    S.Pause_Service
    Command3.Caption = "Continuar"
Else
    S.Resume_Service
    Command3.Caption = "Pausar"
End If
Sleep (2000)
Command6.Value = True
Mierda:
End Sub

Private Sub Command4_Click()
Dim S As New UpcServiciosyWea
On Error GoTo Mierda
    S.ServiceName = lv.ListItems(lv.SelectedItem.Index).Text
    S.RemoveService
    Sleep (2000)
    Command6.Value = True
    
Mierda:
End Sub

Private Sub Command5_Click()
Dim S As New UpcServiciosyWea
Dim Sname As String
Dim SdName As String
Dim Spath As String
On Error GoTo Mierda
    
    Sname = InputBox("Ingrese el nombre del servicio", "Nombre del servicio")
    SdName = InputBox("Ingrese el nombre con el cual se mostrará el servicio", "Nombre para mostrar")
    Spath = InputBox("Ingrese la ruta en donde se encuentra el servicio", "Ruta del servicio")
    S.ServiceName = Sname
    S.InstallService Spath, SdName, "", e_ServiceType_Automatic
    Command6.Value = True
    


Mierda:
End Sub

Private Sub Command6_Click()
Dim S As New UpcServiciosyWea
Dim Cs As Collection
Dim E As ListItem
Dim i As Long
lv.ListItems.Clear
With S
    Set Cs = .ListServices
    For i = 1 To Cs.Count Step 3
        Set E = lv.ListItems.Add(, , CStr(Cs(i)), 1, 1) 'CStr(Cs(i)), , , 0)
        E.SubItems(1) = Cs(i + 1)
        E.SubItems(2) = Cs(i + 2)
        DoEvents
    Next i
    
End With
End Sub
