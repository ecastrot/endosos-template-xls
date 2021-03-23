VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHome 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Endosos - Alejandro Giraldo Arango"
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   13905
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   6360
      ScaleHeight     =   1095
      ScaleWidth      =   7335
      TabIndex        =   7
      Top             =   120
      Width           =   7335
      Begin VB.Image iProcess2 
         Height          =   675
         Left            =   4080
         Picture         =   "frmHome.frx":0000
         Top             =   240
         Visible         =   0   'False
         Width           =   3030
      End
      Begin VB.Image iConfig2 
         Height          =   675
         Left            =   240
         Picture         =   "frmHome.frx":6B22
         Top             =   240
         Visible         =   0   'False
         Width           =   3510
      End
      Begin VB.Image iProcess 
         Height          =   675
         Left            =   4080
         Picture         =   "frmHome.frx":E724
         Top             =   240
         Width           =   3015
      End
      Begin VB.Image iConfig 
         Height          =   660
         Left            =   240
         Picture         =   "frmHome.frx":15192
         Top             =   240
         Width           =   3510
      End
   End
   Begin VB.ListBox listTemplates 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7155
      Left            =   360
      TabIndex        =   6
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Frame frmProcess 
      BackColor       =   &H00FF8080&
      Caption         =   "Procesar plantilla"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   3120
      TabIndex        =   18
      Top             =   1560
      Width           =   10575
   End
   Begin VB.Frame frmConfig 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Gestión de plantilla"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   3120
      TabIndex        =   9
      Top             =   1560
      Visible         =   0   'False
      Width           =   10575
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Gestión de columnas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   720
         TabIndex        =   12
         Top             =   960
         Width           =   9135
         Begin VB.TextBox tColResult 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   3720
            MaxLength       =   1
            TabIndex        =   3
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox tColSufix 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   7200
            MaxLength       =   30
            TabIndex        =   5
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox tColPrefix 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   5280
            MaxLength       =   30
            TabIndex        =   4
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox tColSource 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2280
            MaxLength       =   1
            TabIndex        =   2
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox tColName 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   360
            MaxLength       =   15
            TabIndex        =   1
            Top             =   840
            Width           =   1695
         End
         Begin VB.Image iAddCol 
            Height          =   420
            Left            =   3240
            Picture         =   "frmHome.frx":1CAD4
            Top             =   1440
            Width           =   2550
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Col. Destino"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   17
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sufijo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7200
            TabIndex        =   16
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Prefijo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5280
            TabIndex        =   15
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Col. Origen"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   14
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nombre Columna"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   13
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.TextBox tNameTemplate 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2640
         TabIndex        =   0
         Top             =   480
         Width           =   2295
      End
      Begin MSComctlLib.ListView listCols 
         Height          =   3450
         Left            =   720
         TabIndex        =   10
         Top             =   3120
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   6085
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nombre Columna"
            Object.Width           =   3263
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Col. Origen"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Col. Destino"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Prefijo"
            Object.Width           =   4322
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Sufijo"
            Object.Width           =   3881
         EndProperty
      End
      Begin VB.Image iReset 
         Height          =   525
         Left            =   6960
         Picture         =   "frmHome.frx":20316
         Top             =   6720
         Width           =   2550
      End
      Begin VB.Image iErase 
         Height          =   480
         Left            =   4080
         Picture         =   "frmHome.frx":24958
         Top             =   6720
         Width           =   2700
      End
      Begin VB.Image iSave 
         Height          =   480
         Left            =   1200
         Picture         =   "frmHome.frx":28D1A
         Top             =   6720
         Width           =   2700
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nombre Plantilla"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Label label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Plantillas existentes"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   120
      X2              =   13485
      Y1              =   1305
      Y2              =   1305
   End
   Begin VB.Image Image1 
      Height          =   1185
      Left            =   105
      Picture         =   "frmHome.frx":2D0DC
      Top             =   60
      Width           =   4410
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public templateSel As String

Private Sub Command1_Click()
'Call ModFile.savePropertyFile(ModIni.fileConfigPath, ModIni.KEY_TEMPLATES, newTemplates)
Call ModFile.removePropertyFile(ModIni.fileConfigPath, "miClave", "any")

'miClave=miValor

End Sub

Private Sub Form_Load()
For i = 0 To UBound(ModIni.templates)
    Me.listTemplates.AddItem ModIni.templates(i)
Next
'Me.listCols.ListItems.Remove 1
End Sub

Private Sub iAddCol_Click()
If (Me.tColName.Text = "") Then
    MsgBox "Debes ingresar un nombre para la columna", vbCritical, "Endosos Alejandro Giraldo"
    Me.tColName.SetFocus
    Exit Sub
End If

If (Me.tColSource.Text = "") Then
    MsgBox "Debes ingresar la columna de origen", vbCritical, "Endosos Alejandro Giraldo"
    Me.tColSource.SetFocus
    Exit Sub
End If

If Not (isValidCol(Me.tColSource.Text)) Then
    MsgBox "Debe ingresar una columna válida entre A-Z", vbCritical, "Endosos Alejandro Giraldo"
    Me.tColSource.Text = ""
    Me.tColSource.SetFocus
    Exit Sub
End If

If (Me.tColResult.Text = "") Then
    MsgBox "Debes ingresar la columna de destino", vbCritical, "Endosos Alejandro Giraldo"
    Me.tColResult.SetFocus
    Exit Sub
End If

If Not (isValidCol(Me.tColResult.Text)) Then
    MsgBox "Debe ingresar una columna válida entre A-Z", vbCritical, "Endosos Alejandro Giraldo"
    Me.tColResult.Text = ""
    Me.tColResult.SetFocus
    Exit Sub
End If

Dim colName As String
colName = UCase(Me.tColName.Text)

If (existColName(colName)) Then
    MsgBox "Ya existe otra columna con este mismo nombre", vbCritical, "Endosos Alejandro Giraldo"
    Exit Sub
End If

Set li = Me.listCols.ListItems.Add(, , colName)
li.SubItems(1) = Me.tColSource
li.SubItems(2) = Me.tColResult
li.SubItems(3) = Me.tColPrefix
li.SubItems(4) = Me.tColSufix

Call cleanInfoCol
End Sub

Private Sub iConfig_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.iConfig.Visible = False
Me.iConfig2.Visible = True
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Image5_Click()

End Sub

Private Sub iConfig2_Click()
frmConfig.Visible = True
frmProcess.Visible = False

Call iReset_Click
End Sub

Private Sub iErase_Click()
If (Me.listTemplates.ListIndex = -1) Then
    MsgBox "Debes seleccionar la plantilla que quieres borrar", vbCritical, "Endosos Alejandro Giraldo"
    Exit Sub
End If

If (MsgBox("¿Está seguro de borrar la plantilla " & Me.templateSel & "?", vbQuestion + vbYesNo, "Endosos Alejandro Giraldo") = vbYes) Then

    Dim cols As String
    Dim colArray() As String
    cols = ModFile.readPropertyFile(ModIni.fileConfigPath, Me.templateSel & ".cols", "")
    If (cols <> "") Then
        colArray() = Split(cols, ",")
        For i = 0 To UBound(colArray)
            Call ModFile.removePropertyFile(ModIni.fileConfigPath, Me.templateSel & "." & colArray(i))
        Next
    End If
    Call ModFile.removePropertyFile(ModIni.fileConfigPath, Me.templateSel & ".cols")
    ModIni.removeTemplate (Me.listTemplates.ListIndex)
    Me.listTemplates.RemoveItem (Me.listTemplates.ListIndex)
    Call iReset_Click
End If
End Sub

Private Sub iProcess_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.iProcess.Visible = False
Me.iProcess2.Visible = True
End Sub

Private Sub iProcess2_Click()
frmConfig.Visible = False
frmProcess.Visible = True

Call iReset_Click
End Sub

Private Sub iReset_Click()
Me.templateSel = ""
Me.tNameTemplate = ""
Call cleanInfoCol
Me.listTemplates.ListIndex = -1
Me.listCols.ListItems.Clear
Me.tNameTemplate.Enabled = True
If Me.frmConfig.Visible = True Then
    Me.tNameTemplate.SetFocus
End If
End Sub

Private Function cleanInfoCol()
Me.tColName = ""
Me.tColSource = ""
Me.tColResult = ""
Me.tColPrefix = ""
Me.tColSufix = ""
End Function

Private Sub iSave_Click()
If (Me.tNameTemplate.Text = "") Then
    MsgBox "Debes ingresar un nombre para la plantilla", vbCritical
    Exit Sub
End If

Dim templateName As String
templateName = UCase(Me.tNameTemplate.Text)

If (templateSel = "" And existTemplate(templateName)) Then
    MsgBox "Ya existe otra plantilla con este mismo nombre", vbCritical, "Endosos Alejandro Giraldo"
    Exit Sub
End If

If (Me.listCols.ListItems.Count <= 0) Then
    MsgBox "No existen columnas asociadas a la plantilla", vbCritical, "Endosos Alejandro Giraldo"
    Exit Sub
End If

If (Me.templateSel <> "") Then
    'Erase info's col
    Dim colsErase As String
    Dim colEraseArray() As String
    colsErase = ModFile.readPropertyFile(ModIni.fileConfigPath, Me.templateSel & ".cols", "")
    If (colsErase <> "") Then
        colEraseArray() = Split(colsErase, ",")
        For i = 0 To UBound(colEraseArray)
            Call ModFile.removePropertyFile(ModIni.fileConfigPath, Me.templateSel & "." & colEraseArray(i))
        Next
    End If
    Call ModFile.removePropertyFile(ModIni.fileConfigPath, Me.templateSel & ".cols")
Else
    Me.templateSel = Me.tNameTemplate.Text
    Me.listTemplates.AddItem Me.templateSel
    Call ModIni.addTemplate(Me.templateSel)
End If
    
Dim cols As String
cols = ""
Dim j As Integer
For j = 1 To Me.listCols.ListItems.Count
    cols = cols & Me.listCols.ListItems(j).Text & ","
    ModIni.addCol Me.templateSel, Me.listCols.ListItems(j).Text, Me.listCols.ListItems(j).SubItems(1), Me.listCols.ListItems(j).SubItems(2), Me.listCols.ListItems(j).SubItems(3), Me.listCols.ListItems(j).SubItems(4)
Next
cols = Left$(cols, Len(cols) - 1)
ModIni.addCols Me.templateSel, cols

MsgBox "Se guardó correctamente la información de la plantilla", vbInformation, "Endosos Alejandro Giraldo"
Call iReset_Click
End Sub

Private Function existTemplate(templateToCreate As String) As Boolean
For i = 0 To UBound(ModIni.templates)
    If (ModIni.templates(i) = templateToCreate) Then
        existTemplate = True
        Exit Function
    End If
Next
existTemplate = False
End Function

Private Function existColName(colName As String) As Boolean
Dim noItems As Integer
noItems = Me.listCols.ListItems.Count
If (noItems <= 0) Then
    existColName = False
    Exit Function
End If
        
Dim i As Integer
For i = 1 To noItems
    If (Me.listCols.ListItems(i).Text = colName) Then
        existColName = True
        Exit Function
    End If
Next
existColName = False
End Function


Private Function isValidCol(col As String) As Boolean
Dim i As Integer
For i = 0 To UBound(ModIni.COLS_ALLOWED)
    If (col = ModIni.COLS_ALLOWED(i)) Then
        isValidCol = True
        Exit Function
    End If
Next
isValidCol = False
End Function

Private Sub listCols_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 46 And Not Me.listCols.SelectedItem Is Nothing) Then
    If (MsgBox("¿Está seguro de borrar la columna " & Me.listCols.SelectedItem.Text & "?", vbQuestion + vbYesNo, "Endosos Alejandro Giraldo") = vbYes) Then
        Me.listCols.ListItems.Remove Me.listCols.SelectedItem.Index
    End If
End If
End Sub

Private Sub listTemplates_DblClick()
Me.templateSel = ModIni.templates(Me.listTemplates.ListIndex)
Me.tNameTemplate = Me.templateSel

If (Me.frmProcess.Visible = True) Then
    Exit Sub
End If

Me.listCols.ListItems.Clear

Dim cols As String
Dim colArray() As String
cols = ModFile.readPropertyFile(ModIni.fileConfigPath, Me.templateSel & ".cols", "")
If (cols <> "") Then
    colArray() = Split(cols, ",")
    For i = 0 To UBound(colArray)
        Dim colInfo As String
        Dim colInfoArray() As String
        colInfo = ModFile.readPropertyFile(ModIni.fileConfigPath, Me.templateSel & "." & colArray(i), "")
        If (colInfo <> "") Then
            colInfoArray() = Split(colInfo, ",")
            Set li = Me.listCols.ListItems.Add(, , colArray(i))
            li.SubItems(1) = colInfoArray(0)
            li.SubItems(2) = colInfoArray(1)
            li.SubItems(3) = colInfoArray(2)
            li.SubItems(4) = colInfoArray(3)
        End If
    Next
End If
Me.tNameTemplate.Enabled = False
End Sub

Private Sub picOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.iConfig2.Visible = True Then Me.iConfig2.Visible = False
If Me.iProcess2.Visible = True Then Me.iProcess2.Visible = False
Me.iConfig.Visible = True
Me.iProcess.Visible = True
End Sub

Private Sub tColName_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub tColPrefix_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub tColResult_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub tColSource_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub tColSufix_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub tNameTemplate_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
