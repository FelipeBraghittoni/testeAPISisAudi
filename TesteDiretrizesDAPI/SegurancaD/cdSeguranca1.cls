'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'  Persistable = 1  'Persistable
'  DataBindingBehavior = 0  'vbNone
'  DataSourceBehavior  = 0  'vbNone
'  MTSTransactionMode  = 0  'NotAnMTSObject
'END
'Attribute VB_Name = "cdSeguranca1"
'Attribute VB_GlobalNameSpace = True
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = True
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Attribute VB_Ext_KEY = "Member0" ,"usuario"

Option Explicit On
Imports ADODB.CommandTypeEnum
Imports ADODB.CursorTypeEnum
Imports ADODB.LockTypeEnum
Imports ADODB.ExecuteOptionEnum
Imports ADODB.ObjectStateEnum
Imports Microsoft.VisualBasic.FileIO


<ComClass(cdSeguranca1.ClassId, cdSeguranca1.InterfaceId, cdSeguranca1.EventsId)>
Public Class cdSeguranca1

    'Public Db As ADODB.Connection
    'Public tDirTpDiretriz As ADODB.Recordset

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537859001"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d329001"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c9001"
#End Region

    '#########################################################
    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub
    '###########################

    'Variáveis para abrir as tabelas relacionadas às classes:
    Private Db As ADODB.Connection

    'Para buscar o nome do computador:
    Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Integer) As Integer


    Public Function sGetComputerName() As String
        '* **********************************************************
        '* * Busca o nome do computador onde está rodando o sistema *
        '* **********************************************************

        Dim sBuffer As String
        Dim lBufSize As Integer
        Dim lStatus As Integer

        lBufSize = 255
        sBuffer = New String(" "c, lBufSize)       'String$(lBufSize, " ")
        lStatus = GetComputerName(sBuffer, 12)
        sGetComputerName = ""
        If lStatus <> 0 Then
            sGetComputerName = Left$(sBuffer, InStr(sBuffer, Chr(0)) - 1)
        Else
            MsgBox("O sistema retornou o erro " & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function executaSQL(vSeleciona As String) As Integer
        '* Executa um comando SQL, como DELETE ou UPDATE:

        Dim strConnect As String

        On Error GoTo EexecutaSQL

        'Cria uma instância de ADODB.Connection:
        Db = New ADODB.Connection
        Db.CommandTimeout = 300
        '* ********************************************
        '* Cria uma instância de ADODB.Connection:    *
        '* ********************************************
        strConnect = LeDADOSsys(1)
        Db.Open(strConnect)

        Db.Execute(vSeleciona, , adExecuteNoRecords)

        Db.Close()
        executaSQL = 0

EexecutaSQL:
        If Err.Number <> 0 Then
            executaSQL = Err.Number
            MsgBox(Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function executaSQLAudiHoras(vSeleciona As String, vSelec2 As String) As Integer
        '* Executa um comando SQL, como DELETE ou UPDATE no banco de Dados AudiHoras:

        Dim strConnect As String

        On Error GoTo EexecutaSQLAudiHoras

        'Cria uma instância de ADODB.Connection:
        Db = New ADODB.Connection
        Db.CommandTimeout = 300
        '* ********************************************
        '* Cria uma instância de ADODB.Connection:    *
        '* ********************************************
        strConnect = vSelec2
        Db.Open(strConnect)
        Db.Execute(vSeleciona, , adExecuteNoRecords)

        Db.Close()
        executaSQLAudiHoras = 0

EexecutaSQLAudiHoras:
        If Err.Number <> 0 Then
            executaSQLAudiHoras = Err.Number
            MsgBox(Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function FileExists(Path As String) As Integer
        '* **************************************************
        '* * Verifica se um arquivo existe tentando abri-lo *
        '* **************************************************

        If FileSystem.FileExists(Path) Then
            FileExists = 1
        Else
            FileExists = 0
        End If

        'Dim x As Integer

        'x = FreeFile
        'On Local Error Resume Next
        'Open Path For Input As x
        'If Err.Number = 0 Then
        '    FileExists = 1
        'Else
        '    FileExists = 0
        'End If
        'Err.Number = 0
        'Close #x
    End Function

    '    Public Function CompactaDB() As Integer
    '        'Compacta o banco de dados

    '        'Define as variáveis:
    '        Dim RC As Integer
    '        Dim je As JRO.JetEngine

    '        On Local Error GoTo ECompactaDB

    '        'Define uma instância para JRO:
    '        je = New JRO.JetEngine

    '        'Verifica se o arquivo temporário já existe:
    '        If Dir(".\NewAudiTeste.MDB") <> "" Then Kill ".\NewAudiTeste.MDB"

    '    'Compacta o banco de dados:
    '        je.CompactDatabase "Data Source=.\AudiTeste.MDB;Jet OLEDB:Database Password=etpospfc;",
    '        "Data Source=" & ".\NewAudiTeste.MDB;Jet OLEDB:Database Password=etpospfc;"

    '    'Elimina o DB original:
    '        Kill ".\AudiTeste.MDB"

    '    'Renomeia o novo DB com o nome original:
    '        Name ".\NewAudiTeste.MDB" As ".\AudiTeste.MDB"

    '    CompactaDB = 0

    'ECompactaDB:
    '        If Err.Number <> 0 Then
    '            MsgBox "Classe cdSeguranca - CompactaDB" & Chr$(13) & Chr$(10) & Err.Number & " - " & Err.Description
    '        CompactaDB = Err.Number
    '        End If

    '    End Function

    Public Function GravaDADOSsys(Path As String) As Integer
        '* ******************************************
        '* * Grava o acesso ao DB em .\DADOSDB.SYS  *
        '* ******************************************

        Dim SegUsuario As SegurancaD.segUsuario
        Dim FNum As Integer
        Dim vCurDir As String = ""
        Dim vLenPath As Integer
        Dim vPath As String

        On Error GoTo EGravaDADOSsys
        GravaDADOSsys = 0
        SegUsuario = New SegurancaD.segUsuario

        '* ****************************************
        '* Criptografa os dados a serem gravados: *
        '* ****************************************
        vLenPath = Len(Path)
        vPath = vLenPath.ToString("000")
        vPath = vPath & SegUsuario.cripto(Mid(Path, 1, vLenPath), 5)
        vCurDir = ""
        'Verifica se DADOS.SYS existe. Se sim, elimina-o:
        FNum = 0
        If FileExists("\\Vboxsvr\vm\Sistema AudiTeste SQLExpress\DADOSDB.SYS") = 1 Then
            Kill("\\Vboxsvr\vm\Sistema AudiTeste SQLExpress\DADOSDB.SYS")
            vCurDir = "\\Vboxsvr\vm\Sistema AudiTeste SQLExpress"
            FNum = 1
        Else
            If FileExists("C:\Auditeste\DADOSDB.SYS") = 1 Then
                Kill("C:\Auditeste\DADOSDB.SYS")
                vCurDir = "C:\Auditeste"
                FNum = 2
            Else
                If FileExists("C:\Arquivos de Programas\Auditeste\DADOSDB.SYS") = 1 Then
                    Kill("C:\Arquivos de Programas\Auditeste\DADOSDB.SYS")
                    vCurDir = "C:\Arquivos de Programas\Auditeste"
                    FNum = 3
                Else
                    If FileExists("C:\Program Files\Auditeste\DADOSDB.SYS") = 1 Then
                        Kill("C:\Program Files\Auditeste\DADOSDB.SYS")
                        vCurDir = "C:\Program Files\Auditeste"
                        FNum = 4
                    Else
                        If FileExists("C:\Windows\DADOSDB.SYS") = 1 Then
                            Kill("C:\Windows\DADOSDB.SYS")
                            vCurDir = "C:\Windows"
                            FNum = 5
                        End If
                    End If
                End If
            End If
        End If
        If FNum = 0 Then
            If FileExists("\\Vboxsvr\vm\Sistema AudiTeste SQLExpress\AuditesteSQLe.exe") = 1 Then
                vCurDir = "\\Vboxsvr\vm\Sistema AudiTeste SQLExpress"
                FNum = 1
            Else
                If FileExists("C:\Auditeste\AuditesteSQLe.exe") = 1 Then
                    vCurDir = "C:\Auditeste"
                    FNum = 2
                Else
                    If FileExists("C:\Arquivos de Programas\Auditeste\AuditesteSQLe.exe") = 1 Then
                        vCurDir = "C:\Arquivos de Programas\Auditeste"
                        FNum = 3
                    Else
                        If FileExists("C:\Program Files\Auditeste\AuditesteSQLe.exe") = 1 Then
                            vCurDir = "C:\Program Files\Auditeste"
                            FNum = 4
                        Else
                            vCurDir = "C:\Windows"
                            FNum = 5
                        End If
                    End If
                End If
            End If
        End If
        '* ***********************************************************
        '* ***********************************************************
        '* Se for compilar para instalação na VM, tirar o comentário *
        '* Caso contrário, deixar como comentário antes de compilar: *
        '* ***********************************************************
        '* Início
        '* ***********************************************************
        'vCurDir = System.AppDomain.CurrentDomain.BaseDirectory() 'application.startuppath 'App.Path
        'FNum = FileExists(vCurDir & "\DADOSDB.SYS")
        'If FNum <> 0 Then   'Arquivo existe
        '    Kill(vCurDir & "\DADOSDB.SYS")
        'End If
        '* ***********************************************************
        '* Término
        '* ***********************************************************
        'Grava um arquivo com os dados do DB a ser aberto:
        Dim vPathOut() As Byte = System.Text.Encoding.Default.GetBytes(vPath)
        FileSystem.WriteAllBytes(vCurDir & "\DADOSDB.SYS", vPathOut, False)
        'FNum = FreeFile
        'Open vCurDir & "\DADOSDB.SYS" For Binary Access Write As #FNum     'Grava o arquivo
        'Put #FNum, 1, vPath
        'Close #FNum

EGravaDADOSsys:
        If Err.Number <> 0 Then
            MsgBox("Classe cdSeguranca - GravaDADOSsys" & vbCrLf & Err.Number & " - " & Err.Description)
            GravaDADOSsys = Err.Number
        End If
    End Function

    Public Function LeDADOSsys(tipo As Integer) As String
        'Le dados para identificação do DB a ser aberto:

        'Define variáveis:
        Dim SegUsuario As SegurancaD.segUsuario
        Dim FNum As Integer
        Dim vCurDir As String = ""
        Dim vLenPath As Integer
        Dim vPath As String

        LeDADOSsys = ""
        SegUsuario = New SegurancaD.segUsuario
        'Verifica se DADOS.SYS existe:
        If FileExists("\\Vboxsvr\vm\Sistema AudiTeste SQLExpress\DADOSDB.SYS") = 1 Then
            vCurDir = "\\Vboxsvr\vm\Sistema AudiTeste SQLExpress"
            FNum = 1
        Else
            If FileExists("C:\Auditeste\DADOSDB.SYS") = 1 Then
                vCurDir = "C:\Auditeste"
                FNum = 2
            Else
                If FileExists("C:\Arquivos de Programas\Auditeste\DADOSDB.SYS") = 1 Then
                    vCurDir = "C:\Arquivos de Programas\Auditeste"
                    FNum = 3
                Else
                    If FileExists("C:\Program Files\Auditeste\DADOSDB.SYS") = 1 Then
                        vCurDir = "C:\Program Files\Auditeste"
                        FNum = 4
                    Else
                        If FileExists("C:\Windows\DADOSDB.SYS") = 1 Then
                            vCurDir = "C:\Windows"
                            FNum = 5
                        End If
                    End If
                End If
            End If
        End If
        '* ***********************************************************
        '* ***********************************************************
        '* Se for compilar para instalação na VM, tirar o comentário *
        '* Caso contrário, deixar como comentário antes de compilar: *
        '* ***********************************************************
        '* Início
        '* ***********************************************************
        'vCurDir = System.AppDomain.CurrentDomain.BaseDirectory()
        'FNum = FileExists(vCurDir & "\DADOSDB.SYS")
        '* ***********************************************************
        '* Término
        '* ***********************************************************
        If FNum <> 0 Then   'Arquivo existe
            'Prepara variável para receber os dados lidos:
            LeDADOSsys = New String(" "c, 320) 'String$(320, " ")
            'Abre e lê o arquivo:
            Dim BytesLeDADOSsys() = FileSystem.ReadAllBytes(vCurDir & "\DADOSDB.SYS")
            LeDADOSsys = System.Text.Encoding.Default.GetString(BytesLeDADOSsys)
            'FNum = FreeFile
            'Open vCurDir & "\DADOSDB.SYS" For Binary Access Read As #FNum      'Le DADOSDB.SYS
            'Get #FNum, 1, LeDADOSsys
            'Close #FNum
            '* ********************************
            '* Descriptografa os dados lidos: *
            '* ********************************
            vLenPath = Val(Mid(LeDADOSsys, 1, 3))
            vPath = ""
            vPath = vPath & SegUsuario.DeCripto(Mid(LeDADOSsys, 4, vLenPath), 5)
            'Analisa qual Banco de Dados será lido:
            If InStr(1, vPath, Chr(10), vbBinaryCompare) > 0 Then
                'Banco de dados do Sistema Auditeste:
                If tipo = 1 Then
                    LeDADOSsys = Mid(vPath, 1, InStr(1, vPath, Chr(10), vbBinaryCompare) - 1)
                Else        'Banco de dados do AudiHoras
                    LeDADOSsys = Mid(vPath, InStr(1, vPath, Chr(10), vbBinaryCompare) + 1)
                End If
            Else
                If tipo = 1 Then
                    MsgBox("Erro no acesso ao Banco de Dados do Sistema Auditeste" & Chr(10) & "Entrar em Contato com o Administrador.")
                Else
                    MsgBox("Erro no acesso ao Banco de Dados do AudiHoras" & Chr(10) & "Entrar em Contato com o Administrador.")
                End If
            End If
        Else
            LeDADOSsys = ""
        End If

    End Function

    'Public Sub CreateAccessDB(DBName As String)
    '    Dim adoCat As ADOX.Catalog
    '    Dim Filenm As String, strConn As String

    '    '~~> MDB to be created. Change this to relevant path and filename
    '    Filenm = DBName

    '    '~~> Creating an instance of the ADOX-object.
    '    adoCat = New ADOX.Catalog

    '    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" &
    '        "Data Source = " & Filenm & ";" &
    '        "Jet OLEDB:Database Password = etpospfc;"

    '    '~~> Deleting the existing copy of the database, if any.
    '    On Local Error Resume Next
    '    Kill Filenm
    'On Local Error GoTo 0

    '    '~~> Create the MDB.
    '    adoCat.Create(strConn)

    'End Sub

    Public Function CriatSQLtotAccess(tabela As String) As Boolean
        '**************************************
        '* Exporta uma tabela SQL para Access *
        '* ************************************

        Dim sqlDB As New ADODB.Connection    'SQL
        Dim jetDB As New ADODB.Connection    'Access
        Dim rs As New ADODB.Recordset        'SQL
        Dim rt As New ADODB.Recordset        'Access
        ''''ReDim Tables$(0)

        Dim I As Integer
        Dim SQL As String
        Dim mesa As String
        Dim tbl As String
        Dim txt As String
        Dim ds As String
        Dim conName As String
        Dim f As Integer
        ''''ReDim SQLstr(0) As String
        Dim DBFileSpec As String
        Dim SQLFileSpec As String
        'Dim vCurDir As String
        'Dim FNum As Integer

        On Error GoTo CriaHell
        CriatSQLtotAccess = False

        ' *** Open SQL Server ***
        SQLFileSpec = LeDADOSsys(1)
        sqlDB.Open(SQLFileSpec$)

        ' *** Grab all SQL tables ***
        'SQL$ = "SELECT * FROM INFORMATION_SCHEMA.Tables ORDER by Table_Type, Table_Name"
        'rs.Open SQL$, sqlDB, adOpenForwardOnly, adLockReadOnly
        'While Not rs.EOF
        ' *** Save table name ***
        'ReDim Preserve Tables$(UBound(Tables$) + 1)
        'Tables$(UBound(Tables$)) = rs!Table_Name
        'rs.MoveNext
        'Wend
        'rs.Close

        ' *** Connect to Jet database ***
        jetDB = New ADODB.Connection
        DBFileSpec$ = ".\AudiTeste.mdb"
        jetDB.Open("Provider=Microsoft.Jet.OLEDB.4.0;" &
        "Data Source = " & DBFileSpec$ & ";" &
        "Jet OLEDB:Database Password = etpospfc;")

        '******************************************
        ' *** Cria e carrega a tabela no Access ***
        '******************************************
        mesa$ = tabela
        ' *** Elimina a Tabela Access ***
        SQL$ = "DROP TABLE [" & mesa$ & "] "
        On Error Resume Next
        jetDB.Execute(SQL$, , adExecuteNoRecords)
        On Error GoTo CriaHell
        ' *** Get table info ***
        SQL$ = "SELECT '" & mesa$ & "' FROM INFORMATION_SCHEMA.Tables"
        rs.Open(SQL$, sqlDB, adOpenForwardOnly, adLockReadOnly)
        If Not rs.EOF Then
            ' *** Get Column info ***
            SQL$ = "SELECT Column_Name, Data_Type, Character_Maximum_Length, Is_Nullable "
            SQL$ = SQL$ & "FROM INFORMATION_SCHEMA.Columns WHERE "
            SQL$ = SQL$ & "Table_Name = '" & mesa$ & "'"
            rt.Open(SQL$, sqlDB, adOpenForwardOnly, adLockReadOnly)
            tbl$ = "("
            txt$ = ""
            While Not rt.EOF
                Select Case rt.Fields("Data_Type").Value
                    Case "bit"
                        ' *** Yes/No ***
                        txt$ = "YESNO"
                    Case "tinyint"
                        ' *** Byte ***
                        txt$ = "BYTE"
                    Case "smallint"
                        ' *** Integer ***
                        txt$ = "INTEGER"
                    Case "int"
                        ' *** Long Integer ***
                        txt$ = "LONG"
                    Case "smallmoney", "money"
                        ' *** Currency ***
                        txt$ = "CURRENCY"
                    Case "real"
                        ' *** Single ***
                        txt$ = "SINGLE"
                    Case "float"
                        ' *** Double ***
                        txt$ = "DOUBLE"
                    Case "decimal"
                        ' *** Decimal ***
                        txt$ = "Decimal"
                    Case "datetime", "smalldatetime", "timestamp"
                        ' *** Date/Time ***
                        txt$ = "DATETIME"
                    Case 9
                        ' *** Binary ***
                        txt$ = "BINARY"
                    Case "ntext"
                        ' *** Text ***
                        If rt.Fields("Character_Maximum_length").Value <= 255 Then
                            txt$ = "STRING"
                        Else
                            txt$ = "Memo"
                        End If
                    Case "nvarchar", "varchar", "nchar", "char"
                        ' *** Text ***
                        ds$ = CStr(rt.Fields("Character_Maximum_length").Value)
                        txt$ = "NVARCHAR(" & ds$ & ")"
                    Case "image"
                        ' *** OLE Object ***
                        txt$ = "OLEObject"
                    Case "memo"
                        ' *** Memo or Hyperlink ***
                        txt$ = "Text"
                End Select
                If rt.Fields("Is_Nullable").Value = "NO" Then
                    txt$ = txt$ & " NOT NULL"
                End If
                tbl$ = tbl$ & "[" & rt.Fields("Column_Name").Value & "]" & " " & txt$ & ","
                rt.MoveNext()
            End While
            Mid$(tbl$, tbl$.Length, 1) = ")"
            rt.Close()
            rs.Close()
            ' *** Create the damn thing already! ***
            SQL$ = "CREATE TABLE [" & mesa$ & "] " & tbl$
            jetDB.Execute(SQL$, , adExecuteNoRecords)
            ' *** Find all Primary Keys ***
            SQL$ = "SELECT s.name AS TABLE_SCHEMA, t.name AS TABLE_NAME, k.name AS CONSTRAINT_NAME, k.type_desc AS CONSTRAINT_TYPE, "
            SQL$ = SQL$ & "c.name AS COLUMN_NAME, ic.key_ordinal AS ORDINAL_POSITION "
            SQL$ = SQL$ & "FROM sys.key_constraints AS k "
            SQL$ = SQL$ & "JOIN sys.tables AS t ON t.object_id = k.parent_object_id "
            SQL$ = SQL$ & "JOIN sys.schemas AS s ON s.schema_id = t.schema_id "
            SQL$ = SQL$ & "JOIN sys.index_columns AS ic ON ic.object_id = t.object_id AND ic.index_id = k.unique_index_id "
            SQL$ = SQL$ & "JOIN sys.columns AS c ON c.object_id = t.object_id AND c.column_id = ic.column_id "
            SQL$ = SQL$ & "ORDER BY TABLE_SCHEMA, TABLE_NAME, CONSTRAINT_TYPE, CONSTRAINT_NAME, ORDINAL_POSITION"
            rt.Open(SQL$, sqlDB, adOpenForwardOnly, adLockReadOnly)
            ' *** Set the table's Primary Keys ***
            txt$ = ""
            conName$ = ""
            While Not rt.EOF
                If rt.Fields("Table_Name").Value = mesa$ Then
                    If rt.Fields("Constraint_Type").Value = "PRIMARY_KEY_CONSTRAINT" Then
                        conName$ = rt.Fields("Constraint_Name").Value
                        txt$ = txt$ & "[" & rt.Fields("Column_Name").Value & "],"
                    End If
                End If
                rt.MoveNext()
            End While
            rt.Close()

            If Len(txt$) Then
                Mid$(txt$, Len(txt$), 1) = ")"
                SQL$ = "ALTER TABLE [" & mesa$ & "] ADD CONSTRAINT [" & conName$ & "] PRIMARY KEY (" & txt$
                jetDB.Execute(SQL$, , adExecuteNoRecords)
            End If
        End If
        On Error Resume Next
        rs.Close()
        rt.Close()

        sqlDB.Close()
        jetDB.Close()

        rs = Nothing
        rt = Nothing
        sqlDB = Nothing
        jetDB = Nothing

        CriatSQLtotAccess = True

        Exit Function

CriaHell:
        'MsgBox Err & " - " & Error()
        If Err.Number = 3704 Then
            Resume Next
        End If
        If rs.State = 1 Then
            rs.Close()
            rs = Nothing
        End If
        If rt.State = 1 Then
            rt.Close()
            rt = Nothing
        End If
        If sqlDB.State <> adStateClosed Then
            sqlDB.Close()
            sqlDB = Nothing
        End If
        If jetDB.State <> adStateClosed Then
            jetDB.Close()
            jetDB = Nothing
        End If
        I% = MsgBox("Internal error", vbCritical, "SQLtoAccess" & vbCrLf & vbCrLf & Err.Description)
        'OBJETO.Cursor = Cursors.Default 'Screen.MousePointer = 0 'normal

    End Function

    Public Function CarregatSQLtotAccess(tabela As String, stringSQL As String) As Integer
        '**************************************
        '* Exporta uma tabela SQL para Access *
        '* ************************************

        Dim FuncoesGerais As SegurancaD.FuncoesGerais
        Dim sqlDB As New ADODB.Connection    'SQL
        Dim jetDB As New ADODB.Connection    'Access
        Dim rs As New ADODB.Recordset        'SQL
        Dim rt As New ADODB.Recordset        'Access
        ''''ReDim Tables$(0)

        Dim I As Integer
        Dim SQL As String
        Dim mesa As String
        Dim tbl As String
        Dim txt As String
        Dim ds As String
        Dim conName As String
        Dim f As Integer
        ''''ReDim SQLstr(0) As String
        Dim DBFileSpec As String
        Dim SQLFileSpec As String

        On Error GoTo CarregaHell
        CarregatSQLtotAccess = 1
        FuncoesGerais = New SegurancaD.FuncoesGerais

        mesa$ = tabela
        ' *** Open SQL Server ***
        SQLFileSpec = LeDADOSsys(1)
        sqlDB.Open(SQLFileSpec$)
        ' *** Connect to Jet database ***
        jetDB = New ADODB.Connection
        DBFileSpec$ = ".\AudiTeste.mdb"
        jetDB.Open("Provider=Microsoft.Jet.OLEDB.4.0;" &
        "Data Source = " & DBFileSpec$ & ";" &
        "Jet OLEDB:Database Password = etpospfc;")
        '*********************
        ' *** Fill 'er up! ***
        '*********************
        If stringSQL = "" Or stringSQL = " " Then
            SQL$ = "SELECT * FROM [" & mesa$ & "]"
        Else
            SQL$ = stringSQL
        End If
        rs.Open(SQL$, sqlDB, adOpenForwardOnly, adLockReadOnly)
        rt = New ADODB.Recordset
        rt.Open(mesa$, jetDB, adOpenDynamic, adLockOptimistic, adCmdTableDirect)
        While Not rs.EOF
            'Carrega a tabela:
            rt.AddNew()

            For f% = 0 To rs.Fields.Count - 1
                If rs.Fields(f%).ActualSize > 0 Then
                    If rs.Fields(f%).Type <> 128 Then
                        rt.Fields(f%).Value = rs.Fields(f%).Value 'rt(f%)
                    Else
                        rt.Fields(f%).Value = FuncoesGerais.dataServ()
                    End If
                Else
                    If rs.Fields(f%).Type <> 205 Then
                        rt.Fields(f%).Value = " "
                    Else
                        rt.Fields(f%).Value = DBNull.Value
                    End If
                End If
            Next f%
            rt.Update()
            rs.MoveNext()
        End While
        On Error Resume Next
        rs.Close()
        rt.Close()

        sqlDB.Close()
        jetDB.Close()

        rs = Nothing
        rt = Nothing
        sqlDB = Nothing
        jetDB = Nothing

        CarregatSQLtotAccess = 0

        Exit Function

CarregaHell:
        MsgBox(Err.Number & " - " & Err.Description)
        If Err.Number = 3704 Then
            Resume Next
        End If
        If rs.State = 1 Then
            rs.Close()
            rs = Nothing
        End If
        If rt.State = 1 Then
            rt.Close()
            rt = Nothing
        End If
        If sqlDB.State <> adStateClosed Then
            sqlDB.Close()
            sqlDB = Nothing
        End If
        If jetDB.State <> adStateClosed Then
            jetDB.Close()
            jetDB = Nothing
        End If
        I% = MsgBox("Internal error", vbCritical, "SQLtoAccess" & vbCrLf & vbCrLf & Err.Description)
        ''''Screen.MousePointer = 0 'normal

    End Function

    Public Function executaAccessSQL(vSeleciona As String) As Integer
        '* Executa um comando SQL, como DELETE ou UPDATE:

        On Error GoTo EexecutaAccessSQL
        'Cria uma instância de ADODB.Connection:
        Db = New ADODB.Connection
        'Abre o DB:
        Db.Open("Provider=Microsoft.Jet.OLEDB.4.0;" &
        "Data Source = .\AudiTeste.MDB;" &
        "Jet OLEDB:Database Password = etpospfc;")

        Db.Execute(vSeleciona, , adExecuteNoRecords)
        Db.Close()
        executaAccessSQL = 0

EexecutaAccessSQL:
        If Err.Number <> 0 Then executaAccessSQL = Err.Number

    End Function

    Public Function compactaSQLexpress(vSelec As String) As Integer

        Dim sqlDB As New ADODB.Connection    'SQL
        Dim rs As New ADODB.Recordset        'SQL
        Dim SQLFileSpec As String
        Dim vSQL As String
        Dim vSelec1 As String
        Dim nomeDB As String
        Dim tamanhoInic As Double
        Dim maxTamInic As Double
        Dim tamanhoFinal As Double
        Dim maxTamFinal As Double
        Dim I As Integer
        Dim sBuffer As String

        On Error GoTo eCompactaSQLe
        compactaSQLexpress = 0

        sBuffer = LeDADOSsys(1)
        sBuffer = Mid(sBuffer, 31, 14)
        ' *** Open SQL Server ***
        SQLFileSpec = LeDADOSsys(1)
        sqlDB.Open(SQLFileSpec$)

        '* ****************************************************
        '* Guarda dados iniciais do tamanho do Banco de Dados *
        '* ****************************************************
        vSelec1 = "SELECT DB_NAME() as NomeBD, name as NomeArquivo, size/128 as TamanhoMB, "
        vSelec1 = vSelec1 & "max_size/128 as MaxTamanhoMB FROM sys.database_files"
        rs.Open(vSelec1, sqlDB, adOpenForwardOnly, adLockReadOnly)
        I = 0
        While Not rs.EOF
            If I = 0 Then
                nomeDB = rs.Fields("NomeBD").Value
                tamanhoInic = Math.Round(rs.Fields("tamanhoMB").Value, 2)
            Else
                If I = 1 Then
                    maxTamInic = Math.Round(rs.Fields("MaxTamanhoMB").Value, 2)
                End If
            End If
            rs.MoveNext()
            I = I + 1
        End While
        rs.Close()

        ' *** Compacta SQL - Banco de dados = vSelec ***
        vSQL$ = "DBCC SHRINKDATABASE(" & vSelec & ", 10)"
        sqlDB.Execute(vSQL$)

        '* **************************************************
        '* Guarda dados finais do tamanho do Banco de Dados *
        '* **************************************************
        vSelec1 = "SELECT DB_NAME() as NomeBD, name as NomeArquivo, size/128 as TamanhoMB, "
        vSelec1 = vSelec1 & "max_size/128 as MaxTamanhoMB FROM sys.database_files"
        rs.Open(vSelec1, sqlDB, adOpenForwardOnly, adLockReadOnly)
        I = 0
        While Not rs.EOF
            If I = 0 Then
                nomeDB = rs.Fields("NomeBD").Value
                tamanhoFinal = Math.Round(rs.Fields("tamanhoMB").Value, 2)
            Else
                If I = 1 Then
                    maxTamFinal = Math.Round(rs.Fields("MaxTamanhoMB").Value, 2)
                End If
            End If
            rs.MoveNext()
            I = I + 1
        End While
        rs.Close()

        '* ****************************************************
        '* Exibe o resultado da compactação do Banco de Dados *
        '* ****************************************************
        If sBuffer = "187.35.155.198" Then
            vSelec1 = "Compactação do DB " & vSelec & " no Servidor efetuada com sucesso."
        Else
            vSelec1 = "Compactação do DB " & vSelec & " Local efetuada com sucesso."
        End If
        vSelec1 = vSelec1 & Chr(10)
        vSelec1 = vSelec1 & Chr(10)
        vSelec1 = vSelec1 & "Tamanho inicial (MB): " & tamanhoInic
        vSelec1 = vSelec1 & Chr(10)
        vSelec1 = vSelec1 & Chr(10)
        vSelec1 = vSelec1 & "Tamanho Final (MB) : " & tamanhoFinal
        MsgBox(vSelec1)

eCompactaSQLe:
        If Err.Number <> 0 Then
            compactaSQLexpress = Err.Number
            MsgBox(Err.Number & " - " & Err.Description)
        End If

    End Function
End Class

