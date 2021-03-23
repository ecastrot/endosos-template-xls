VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmStart 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Endosos - Alejandro Giraldo Arango"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   8850
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1245
      Left            =   7920
      TabIndex        =   9
      Top             =   2790
      Width           =   645
   End
   Begin VB.PictureBox picProcessing 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4470
      Left            =   1455
      Picture         =   "frnStart.frx":0000
      ScaleHeight     =   4470
      ScaleWidth      =   5895
      TabIndex        =   8
      Top             =   7000
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.TextBox tRow 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   6570
      MaxLength       =   3
      TabIndex        =   7
      Top             =   5610
      Width           =   615
   End
   Begin MSComDlg.CommonDialog excelDialog 
      Left            =   120
      Top             =   2550
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Seleccione el archivo de novedades a procesar"
      Filter          =   "Archivos Excel (xlsx)|*.xlsx|Archivos Excel (xls)|*.xls"
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   5715
      Picture         =   "frnStart.frx":55DDA
      Top             =   4845
      Width           =   795
   End
   Begin VB.Image iProcess 
      Height          =   705
      Left            =   2640
      Picture         =   "frnStart.frx":56DBC
      Top             =   6360
      Width           =   3285
   End
   Begin VB.Image Image4 
      Height          =   525
      Left            =   6480
      Picture         =   "frnStart.frx":5E72A
      Top             =   5520
      Width           =   825
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fila a partir de la cual se pondran los resultados"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   5610
      Width           =   4815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Información del archivo de resultado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   4200
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Información de la plantilla base"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   2640
      Width           =   4335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00719800&
      BorderWidth     =   3
      X1              =   1200
      X2              =   7560
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label lExcelResult 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cargar archivo de resultado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   1590
      TabIndex        =   3
      Top             =   4830
      Width           =   4935
   End
   Begin VB.Label lUploadResult 
      BackColor       =   &H000000FF&
      Height          =   600
      Left            =   7320
      TabIndex        =   2
      Top             =   4560
      Width           =   660
   End
   Begin VB.Label lExcelBase 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cargar archivo a migrar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   1590
      TabIndex        =   1
      Top             =   3390
      Width           =   4935
   End
   Begin VB.Image Image3 
      Height          =   660
      Left            =   1440
      Picture         =   "frnStart.frx":5FE64
      Top             =   3240
      Width           =   5865
   End
   Begin VB.Label lUploadBase 
      BackColor       =   &H000000FF&
      Height          =   600
      Left            =   7080
      TabIndex        =   0
      Top             =   3360
      Width           =   660
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   120
      Picture         =   "frnStart.frx":6C8C6
      Top             =   120
      Width           =   8520
   End
   Begin VB.Image Image2 
      Height          =   660
      Left            =   1440
      Picture         =   "frnStart.frx":AC370
      Top             =   4680
      Width           =   5865
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public excelPathBase As String
Public excelPathResult As String

'Excel base
Public excelAppBase As Excel.Application
Public workbookBase As Excel.Workbook
Public sheetBase As Excel.Worksheet

'Excel result
Public excelAppResult As Excel.Application
Public workbookResult As Excel.Workbook
Public sheetResult As Excel.Worksheet

Private Sub Command1_Click()
Form1.Show
End Sub

Private Sub iProcess_Click()
On Error GoTo closeResources

If (Me.excelPathBase = "") Then
    MsgBox "Debe seleccionar el archivo de excel desde el cual se leeran los datos", vbCritical
    Exit Sub
End If

If (Me.excelPathResult = "") Then
    MsgBox "Debe seleccionar el archivo de excel en el cual se escribirán los datos", vbCritical
    Exit Sub
End If

Call showProcessing
Call loadExcelBase
Call loadExcelResult

Dim hasMoreRows As Boolean
Dim row As Integer
Dim rowsProcessed As Integer

hasMoreRows = True
row = 2
rowsProcessed = 0

'Configuracion temporal de columnas de bas
Dim colResName, colBasId, colBasEmail, colBasPhone, colBasApto, colBasFinancial, colBasFinalcialNit, colBasValue As Integer
colBasName = 2
colBasId = 3
colBasEmail = 4
colBasPhone = 5
colBasApto = 6
colBasFinancial = 7
colBasFinalcialNit = 8
colBasValue = 9

'Configuracion temporal de columnas de resultado
Dim colName, colId, colEmail, colPhone, colApto, colFinancial, colFinancialOther, colFinalcialNit, colValue As Integer
Dim colCity, colLocalidad, colAddressType, colClorCr, colDir1, colDir2, colDir3 As Integer
colName = 15
colId = 16
'colEmail = 15 'No existe en la plantilla destino
'colPhone = 15 'No existe en la plantilla destino
colApto = 13
colFinancial = 1 'Resolver AV VILLAS/BANCO BBVA/ otro
colFinancialOther = 2
colFinalcialNit = 3
colValue = 20
colCity = 4
colLocalidad = 5
colAddressType = 6
colClorCr = 7
colDir1 = 8
colDir2 = 10
colDir3 = 12

Dim rowStartRead As Integer
rowStartRead = Val(Me.tRow)

While hasMoreRows
    Value = Me.sheetBase.Cells(row, 1)
    If (Value = "") Then
        hasMoreRows = False
    Else
        valueFinancial = UCase(Me.sheetBase.Cells(row, colBasFinancial))
        If (valueFinancial = "AV VILLAS" Or valueFinancial = "BANCO BBVA") Then
            valueFinancial = valueFinancial
        Else
            valueFinancial = "Otro"
        End If
        
        Me.sheetResult.Cells(rowStartRead, colName) = Me.sheetBase.Cells(row, colBasName)
        Me.sheetResult.Cells(rowStartRead, colId) = Me.sheetBase.Cells(row, colBasId)
        Me.sheetResult.Cells(rowStartRead, colApto) = Me.sheetBase.Cells(row, colBasApto)
        Me.sheetResult.Cells(rowStartRead, colFinancial) = valueFinancial
        If (valueFinancial = "Otro") Then
            Me.sheetResult.Cells(rowStartRead, colFinancialOther) = UCase(Me.sheetBase.Cells(row, colBasFinancial))
        End If
        Me.sheetResult.Cells(rowStartRead, colFinalcialNit) = Me.sheetBase.Cells(row, colBasFinalcialNit)
        Me.sheetResult.Cells(rowStartRead, colValue) = Me.sheetBase.Cells(row, colBasValue)
        Me.sheetResult.Cells(rowStartRead, colCity) = "MEDELLIN"
'        Me.sheetResult.Cells(rowStartRead, colFecha) = Me.sheetBase.Cells(row, colName)
'        Me.sheetResult.Cells(rowStartRead, colFecha) = Me.sheetBase.Cells(row, colName)
'        Me.sheetResult.Cells(rowStartRead, colFecha) = Me.sheetBase.Cells(row, colName)
'        Me.sheetResult.Cells(rowStartRead, colFecha) = Me.sheetBase.Cells(row, colName)
    
'CDate(sheet.Cells(row, ModConfig.COL_DATE)) & " " & CDate(sheet.Cells(row, ModConfig.COL_HOUR_INI))
        rowsProcessed = rowsProcessed + 1
'        Call writeResults(row)
    End If
    row = row + 1
    rowStartRead = rowStartRead + 1
Wend

Me.workbookBase.Close SaveChanges:=False

Me.workbookResult.Save
Me.workbookResult.Close SaveChanges:=False

Call showProcessing

MsgBox "Finalizó con éxito el procesamiento de las novedades." & vbNewLine & vbNewLine & _
    "Se procesaron " & rowsProcessed & " regisros de novedades.", vbInformation
    
closeResources:
Call closeResourcesBase
Call closeResourcesResult
End Sub

Private Sub writeResults(row)
Dim totalToReport As Integer
If ("HORA EXTRA" = Sheet.Cells(row, ModConfig.COL_TYPE_ROW)) Then
    Sheet.Cells(row, ModConfig.COL_HEDO) = IIf(hedo > 0, Round(hedo / 60, 2), "")
    Sheet.Cells(row, ModConfig.COL_HENO) = IIf(heno > 0, Round(heno / 60, 2), "")
    Sheet.Cells(row, ModConfig.COL_HEDF) = IIf(hedf > 0, Round(hedf / 60, 2), "")
    Sheet.Cells(row, ModConfig.COL_HENF) = IIf(henf > 0, Round(henf / 60, 2), "")
    totalToReport = hedo + heno + hedf + henf
Else
    Sheet.Cells(row, ModConfig.COL_RN) = IIf(heno > 0, Round(heno / 60, 2), "")
    Sheet.Cells(row, ModConfig.COL_RF) = IIf(hedf > 0, Round(hedf / 60, 2), "")
    Sheet.Cells(row, ModConfig.COL_RNF) = IIf(henf > 0, Round(henf / 60, 2), "")
    totalToReport = heno + hedf + henf
End If

If (totMins <> totalToReport) Then
    Call markCheckWithError(row, ModConfig.COL_TYPE_ROW, "El total de horas reportadas no puedieron ser clasificadas en su totalidad según su tipo")
End If

Private Function closeResourcesBase()
Set workbookBase = Nothing
If Not excelAppBase Is Nothing Then
    excelAppBase.Quit
    Set excelAppBase = Nothing
End If
End Function

Private Function closeResourcesResult()
Set Me.workbookResult = Nothing
If Not Me.excelAppResult Is Nothing Then
    Me.excelAppResult.Quit
    Set Me.excelAppResult = Nothing
End If
End Function

Private Sub showProcessing()
If (Me.picProcessing.Visible = False) Then
    Me.picProcessing.Visible = True
    Me.picProcessing.Top = 2565
Else
    Me.picProcessing.Visible = False
End If
End Sub

Private Sub lUploadBase_Click()
excelDialog.ShowOpen
If excelDialog.FileName <> "" Then
    Me.excelPathBase = excelDialog.FileName
    Me.lExcelBase = excelDialog.FileTitle
Else
    Me.excelPathBase = ""
    Me.lExcelBase = "Cargar archivo a migrar"
End If
End Sub

Private Sub lUploadResult_Click()
excelDialog.ShowOpen
If excelDialog.FileName <> "" Then
    Me.excelPathResult = excelDialog.FileName
    Me.lExcelResult = excelDialog.FileTitle
Else
    Me.excelPathResult = ""
    Me.lExcelResult = "Cargar archivo de resultado"
End If
End Sub

Private Function loadExcelBase()
Set excelAppBase = New Excel.Application
Set workbookBase = Me.excelAppBase.Workbooks.Open(FileName:=Me.excelPathBase)
Set sheetBase = Me.workbookBase.Sheets(1)
End Function

Private Function loadExcelResult()
Set excelAppResult = New Excel.Application
Set workbookResult = Me.excelAppResult.Workbooks.Open(FileName:=Me.excelPathResult)
Set sheetResult = Me.workbookResult.Sheets(1)
End Function

