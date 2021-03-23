Attribute VB_Name = "ModIni"
Option Explicit

'Almacena la ruta del archivo de configuraciones
Public fileConfigPath As String

'File Keys
Public Const KEY_TEMPLATES As String = "TT"

Public COLS_ALLOWED() As String

'Vars
Public templates() As String
Public noTemplates As Integer

'Public ROW_START_READ As Integer
'Public COL_TYPE_ROW As Integer
'Public COL_DATE As Integer
'Public COL_HOUR_INI As Integer
'Public COL_HOUR_END As Integer
'Public COL_TOT As Integer
'Public COL_HEDO As Integer
'Public COL_HENO As Integer
'Public COL_HEDF As Integer
'Public COL_HENF As Integer
'Public COL_RN As Integer
'Public COL_RNF As Integer
'Public COL_RF As Integer
'
'Public HOUR_START_D As String
'Public HOUR_END_D As String

Sub main()
COLS_ALLOWED = Split("A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z", ",")

fileConfigPath = App.Path & "\config.ini"

Call loadTemplates
'calendarPath = App.Path & "\config.ini"
'currentYear = Year(Now)
'Call reloadHolidays
'Call loadExcelConfig
'
'HOUR_START_D = ModIni.readPropertyFile(calendarPath, ModIni.K_HOUR_START_D, "06:00:00")
'HOUR_END_D = ModIni.readPropertyFile(calendarPath, ModIni.K_HOUR_END_D, "21:00:00")

Call ModFile.savePropertyFile(fileConfigPath, "miClave", "miValor")


frmHome.Show
End Sub

'Public Sub reloadHolidays()
'ReDim holidays(0)
'Call ModCalendar.loadHolidays(currentYear - 1)
'Call ModCalendar.loadHolidays(currentYear)
'Call ModCalendar.loadHolidays(currentYear + 1)
'End Sub
'
'Public Function loadExcelConfig()
'ROW_START_READ = ModIni.readPropertyFile(calendarPath, ModIni.K_ROW_START_READ, 9)
'COL_TYPE_ROW = ModIni.readPropertyFile(calendarPath, ModIni.K_COL_TYPE_ROW, 5)
'COL_DATE = ModIni.readPropertyFile(calendarPath, ModIni.K_COL_DATE, 6)
'COL_TOT = ModIni.readPropertyFile(calendarPath, ModIni.K_COL_TOT, 9)
'COL_HOUR_INI = ModIni.readPropertyFile(calendarPath, ModIni.K_COL_HOUR_INI, 7)
'COL_HOUR_END = ModIni.readPropertyFile(calendarPath, ModIni.K_COL_HOUR_END, 8)
'COL_HEDO = ModIni.readPropertyFile(calendarPath, ModIni.K_COL_HEDO, 10)
'COL_HENO = ModIni.readPropertyFile(calendarPath, ModIni.K_COL_HENO, 11)
'COL_HEDF = ModIni.readPropertyFile(calendarPath, ModIni.K_COL_HEDF, 12)
'COL_HENF = ModIni.readPropertyFile(calendarPath, ModIni.K_COL_HENF, 13)
'COL_RN = ModIni.readPropertyFile(calendarPath, ModIni.K_COL_RN, 14)
'COL_RNF = ModIni.readPropertyFile(calendarPath, ModIni.K_COL_RNF, 15)
'COL_RF = ModIni.readPropertyFile(calendarPath, ModIni.K_COL_RF, 16)
'End Function

Public Function loadTemplates()
'ReDim Preserve holidays(0)
Dim templatesLine As String
Dim i As Integer

'Dim holidaysByMonth() As String
'Dim countHolidays As Integer
'Dim holidayDate As Date

templatesLine = ModFile.readPropertyFile(fileConfigPath, ModIni.KEY_TEMPLATES, "")
If (templatesLine = "") Then
    ModIni.noTemplates = 0
Else
    ModIni.templates() = Split(templatesLine, ",")
    ModIni.noTemplates = UBound(templates) + 1
End If







'countHolidays = getInitialCount()
'
'For m = 1 To 12
'    holidaysLine = ModIni.readPropertyFile(fileConfigPath, ModConfig.calendarPath, "")
'    holidaysByMonth = Split(holidaysLine, "|")
'    For d = 0 To UBound(holidaysByMonth)
'        If (isValidDay(holidaysByMonth(d))) Then
'            holidayDate = DateSerial(year, m, holidaysByMonth(d))
'            If Not (existHoliday(holidayDate)) Then
'                ReDim Preserve holidays(countHolidays)
'                holidays(countHolidays) = holidayDate
'                countHolidays = countHolidays + 1
'            End If
'        End If
'    Next
'Next
End Function


Public Function addTemplate(nameTemplate As String)
ModIni.noTemplates = ModIni.noTemplates + 1
ReDim Preserve ModIni.templates(ModIni.noTemplates - 1)
ModIni.templates(ModIni.noTemplates - 1) = nameTemplate
Call saveCurrentTemplates
End Function

Public Function removeTemplate(indexTemplate As Integer)
Call ModIni.removeTemplateFromArray(indexTemplate)
ModIni.noTemplates = ModIni.noTemplates - 1
Call saveCurrentTemplates
End Function

Public Function saveCurrentTemplates()
Dim newTemplates As String
newTemplates = Join(ModIni.templates, ",")
'newTemplates = Left$(newTemplates, Len(newTemplates) - 1)
Call ModFile.savePropertyFile(ModIni.fileConfigPath, ModIni.KEY_TEMPLATES, newTemplates)
End Function

Public Function removeTemplateFromArray(indexTemplate As Integer)
Dim i As Integer
For i = indexTemplate To ModIni.noTemplates - 2
ModIni.templates(i) = ModIni.templates(i + 1)
Next
ReDim Preserve ModIni.templates(ModIni.noTemplates - 2)
End Function

Public Function addCol(nameTemplate As String, colName As String, colSource As String, colResult As String, colPrefix As String, colSufix As String)
Dim infoKey As String
Dim infoVal As String
infoKey = nameTemplate & "." & colName
infoVal = colSource & "," & colResult & "," & colPrefix & "," & colSufix
Call ModFile.savePropertyFile(ModIni.fileConfigPath, infoKey, infoVal)
End Function

Public Function addCols(nameTemplate As String, colsName As String)
Dim infoKey As String
Dim infoVal As String
infoKey = nameTemplate & "." & "cols"
Call ModFile.savePropertyFile(ModIni.fileConfigPath, infoKey, colsName)
End Function



