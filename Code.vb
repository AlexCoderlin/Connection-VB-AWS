'--------------------
'   SOURCES
'--------------------

'# https://www.javierrguez.com/ejecutar-consultas-mysql-desde-una-macro-excel-con-odbc/
'# https://www.exceleinfo.com/ejecutar-consulta-sql-desde-excel/
'# https://support.microsoft.com/es-es/help/278973/excelado-demonstrates-how-to-use-ado-to-read-and-write-data-in-excel-w?fbclid=IwAR3cIIIMa_cuRctcjCo1w4fjHcHtlfsPWDnB98p5ab1Y-rFuZvSjlPMygeo
'# https://youtu.be/9AhaTELLzC0  (19.Curso Postgres - Ejecutar consulta SQL desde Excel mediante macro VBA) - sugerido desde FB
'# https://youtu.be/dEGGqa_BdFM (Conectar con SQL Server desde Excel | VBA Excel 2013 #50)
'# https://youtu.be/IOPeQZEqb3U (tutorial vba: Ms Excel y mySQL realizando consultas)
'# https://youtu.be/keahyLxtOfQ (Excel Avanzado 2010 Bases de Datos 1) - Curso Pildoras informaticas


Nota: Con JDBC es posible conectar los datos a DataStudio Google

'--------------------
'   FIRST CODE
'--------------------


ThisWorkbook.Connections("ConnectionName").Refresh
Debug.Print ThisWorkbook.Connections("ConnectionName").ODBCConnection.refreshing 'Prints True
While ThisWorkbook.Connections("ConnectionName").ODBCConnection.refreshing
DoEvents
Wend
Debug.Print "updated"

'El codigo anterior es solucionado con la siguiente linea:
Application.CalculateUntilAsyncQueriesDone


                 Con OLE DB          Con ODBC
+---------------+  +---------------+  +---------------+
|   Programa    |  |   Programa    |  |   Programa    |
+---------------+  |      |        |  |      |        |
|      ADO      |  |     ADO       |  |     ADO       |
+---------------+  |      |        |  |      |        |
|    OLE DB     |  |    OLEDB      |  |    OLEDB (OLE DB especial para
|      +--------+  |      |        |  |      |    comunicación con cualquier ODBC)
|      |  ODBC  |  |      |        |  |     ODBC      |
+------+--------+  |      |        |  |      |        |
| Base de datos |  | Base de datos |  | Base de datos |
+---------------+  +---------------+  +---------------+

'---------------
'Antonio Candelario San Juan - Mcros Excel (VbA)
'---------------
Private Const DRIVER_MYSQL As String = "DRIVER={MySQL ODBC 8.0 ANSI Driver};"
Private Const OPCIONS_MYSQL As String = "SERVER=localhost;PORT=3306;DATABASE=" & "movedb;USER=root;PASSWORD=tupassword;"
Public connMySql As ADODB.Connection
Public rsMySql As New ADODB.Recordset
Sub ConexionMysql()
Dim strDB, strSQL As String
Dim strTabla As String
Dim lngCampos As Long
Dim i As Long
Dim bBien As Boolean
On Error GoTo ControlError
bBien = True
Set connMySql = New ADODB.Connection
connMySql.Open DRIVER_MYSQL & OPCIONS_MYSQL
strTabla = "archivo"
strSQL = "SELECT * FROM " & strTabla & " "
rsMySql.Open strSQL, connMySql
Sheets("ARCHIVO").Cells(2, 1).CopyFromRecordset rsMySql
lngCampos = rsMySql.Fields.Count
For i = 0 To lngCampos - 1
Sheets("ARCHIVO").Cells(1, i + 1).Value = rsMySql.Fields(i).Name
Next
rsMySql.Close: Set rsMySql = Nothing
connMySql.Close: Set connMySql = Nothing
Salir:
On Error Resume Next
If Not bBien Then
MsgBox "NO SE HA PODIDO ACTUALIZAR LA BASE DE DATOS, INTÉNTALO MÁS TARDE."
End If
rsMySql.Close: Set rsMySql = Nothing
connMySql.Close: Set connMySql = Nothing
Exit Sub
ControlError:
bBien = False
Resume Salir
End Sub