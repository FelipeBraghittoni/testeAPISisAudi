'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'  Persistable = 1  'Persistable
'  DataBindingBehavior = 0  'vbNone
'  DataSourceBehavior  = 0  'vbNone
'  MTSTransactionMode  = 0  'NotAnMTSObject
'END
'Attribute VB_Name = "segLog"
'Attribute VB_GlobalNameSpace = True
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = True
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit On
Imports ADODB.CommandTypeEnum
Imports ADODB.CursorTypeEnum
Imports ADODB.LockTypeEnum
Imports ADODB.ExecuteOptionEnum
Imports ADODB.ObjectStateEnum
Imports ADODB
Imports Microsoft.VisualBasic.FileIO
<ComClass(segLog.ClassId, segLog.InterfaceId, segLog.EventsId)>
Public Class segLog

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537859006"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d329006"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c9006"
#End Region

    'local variable(s) to hold property value(s)
    'Private mvardataLog As Date 'local copy
    'Private mvarhoraLog As Date 'local copy
    'Private mvarusuarioLog As String 'local copy
    'Private mvarmicroLog As String 'local copy
    'Private mvardescricaoLog As String 'local copy

    'Variáveis para abrir as tabelas relacionadas às classes:
    Public Db As ADODB.Connection
    Public tsegLog As ADODB.Recordset
    'Variáveis para abrir as tabelas relacionadas às classes (criaDbExpt):
    Public DbExpt As ADODB.Connection
    Public tExpt As ADODB.Recordset
    Public cat As ADOX.Catalog
    Public tbl As ADOX.Table

    Private mvardescricaoLog As String
    Public Property descricaoLog() As String
        Get
            Return mvardescricaoLog
        End Get
        Set(ByVal value As String)
            mvardescricaoLog = value
        End Set
    End Property

    Private mvarmicroLog As String
    Public Property microLog() As String
        Get
            Return mvarmicroLog
        End Get
        Set(ByVal value As String)
            mvarmicroLog = value
        End Set
    End Property

    Private mvarusuarioLog As String
    Public Property usuarioLog() As String
        Get
            Return mvarusuarioLog
        End Get
        Set(ByVal value As String)
            mvarusuarioLog = value
        End Set
    End Property

    Private mvarhoraLog As Date
    Public Property horaLog() As Date
        Get
            Return mvarhoraLog
        End Get
        Set(ByVal value As Date)
            mvarhoraLog = value
        End Set
    End Property

    Private mvardataLog As Date
    Public Property dataLog() As Date
        Get
            Return mvardataLog
        End Get
        Set(ByVal value As Date)
            mvardataLog = value
        End Set
    End Property

    '#########################################################
    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub
    '###########################

    Public Function dbConecta(abreDB As Integer, tipo As Integer, Optional vSelec As String = "") As Integer
        '* ****************************************
        '* * abreDB = se abre ou não o DB:        *
        '* * 0 = abre o DB                        *
        '* * 1 = não abre o DB                    *
        '* *                                      *
        '* * tipo = forma de abrir as tabelas:    *
        '* * 0 = OpenTable                        *
        '* * 1 = OpenDynaset                      *
        '* ****************************************

        Dim strConnect As String, strSQL As String

        On Error GoTo EdbConecta
        Dim vSelect As String = ""

        dbConecta = 0   'ReturnCode se não houver nenhum problema

        'Abre o DB:
        If abreDB = 0 Then
            'Cria uma instância de ADODB.Connection:
            Db = New ADODB.Connection
            Dim cdSeguranca1 As SegurancaD.cdSeguranca1
            cdSeguranca1 = New SegurancaD.cdSeguranca1
            strConnect = cdSeguranca1.LeDADOSsys(1)
            Db.Open(strConnect)
        End If

        'Cria uma instância de ADODB.Recordset
        tsegLog = New ADODB.Recordset
        'Abre a tabela:
        If tipo = 0 Then    'Aberto como OpenTable - ChavePrimaria
            'Dim vSelect As String
            tsegLog.Open("Log", Db, adOpenDynamic, adLockOptimistic, adCmdTableDirect)
        Else
            If vSelec = "" Then
                dbConecta = 1014
            Else
                tsegLog.Open(vSelec, Db, adOpenDynamic)
            End If
        End If

EdbConecta:
        If Err.Number Then
            If Err.Number = 3705 Then
                tsegLog.Close()
                tsegLog.Open(vSelect, Db, adOpenDynamic)
                Resume Next
            Else
                dbConecta = Err.Number
                If Err.Number = 3024 Then dbConecta = 1046
                MsgBox("Classe segLog - dbConecta" & vbCrLf & Err.Number & " - " & Err.Description)
            End If
        End If

    End Function

    Public Function leSeq(vPrimVez As Integer) As Integer
        '* ****************************************
        '* * Lê sequencialmente a tabela          *
        '* *                                      *
        '* * vPrimVez - Se é a primeira leitura   *
        '* * da tabela                            *
        '* ****************************************

        On Error GoTo EleSeq

        leSeq = 0   'ReturnCode se não houver nenhum problema
        'Se não chegou no final do arquivo:
        If Not tsegLog.EOF Then
            If vPrimVez = 0 Then    'Não é a primeira vez
                tsegLog.MoveNext()    'lê 1 linha
            End If
        End If
        'Se não chegou no final do arquivo, carrega propriedades:
        If Not tsegLog.EOF Then
            dataLog = tsegLog.Fields("dataLog").Value
            horaLog = tsegLog.Fields("horaLog").Value
            usuarioLog = RTrim(tsegLog.Fields("usuarioLog").Value)
            microLog = RTrim(tsegLog.Fields("microLog").Value)
            If tsegLog.Fields("descricaoLog").Value <> "" Then
                descricaoLog = tsegLog.Fields("descricaoLog").Value
            Else
                descricaoLog = " "
            End If
        Else
            leSeq = 1016
        End If

EleSeq:
        If Err.Number Then
            leSeq = Err.Number
            MsgBox("Classe segLog - leSeq" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function CriaDBExpt(nome As String) As Integer

        Dim Fld As Fields
        Dim Sel As String
        Dim RC As Integer
        Dim I As Integer
        Dim Msg As Integer
        Dim vAux As String
        Dim cdSeguranca1 As SegurancaD.cdSeguranca1

        'Cria uma instância da classe:
        cdSeguranca1 = New SegurancaD.cdSeguranca1

        On Error GoTo ECriaDBExpt

        Msg = 0     'Define a ação a ser tomada
        RC = cdSeguranca1.FileExists(nome)
        If RC <> 0 Then
            'Abre o DB:
            'Cria uma instância de ADODB.Connection:
            DbExpt = New ADODB.Connection
            DbExpt.Open("Provider=Microsoft.Jet.OLEDB.4.0;" &
            "Data Source = " & nome & ";")
            'Abre as características do BD nome:
            tExpt = DbExpt.OpenSchema(SchemaEnum.adSchemaTables)
            '* ***********************************************
            '* * Verifica se nome possui tabela de Log       *
            '* * Se sim, pergunta se:                        *
            '* * - Coloca os dados no final da tabela        *
            '* * - Limpa a tabela antes de carregar os dados *
            '* ***********************************************
            Do Until tExpt.EOF      'Lista todas as Tabelas do DB
                If tExpt.Fields("TABLE_TYPE").Value = "TABLE" Then
                    'É um DB de Log:
                    If tExpt.Fields("TABLE_NAME").Value = "Log" Then
                        vAux = "Deseja preservar os dados existentes?" & vbCrLf & vbCrLf
                        vAux = vAux & "Tecle SIM para copiar no final da tabela." & vbCrLf
                        vAux = vAux & "Tecle NÃO para limpar os dados antes de exportá-los." & vbCrLf
                        vAux = vAux & "Tecle CANCELAR para interromper o processamento."
                        Msg = MsgBox(vAux, vbYesNoCancel, "Tabela de LOG já existe")
                        Exit Do
                    End If
                End If
                tExpt.MoveNext()
            Loop
        Else        'DB Não Existe
            'Cria o DB Access:
            cat = New ADOX.Catalog
            cat.Create("Provider = Microsoft.Jet.OLEDB.4.0;" &
            "Data Source = " & nome & ";")
            'Cria uma instância de ADODB.Connection:
            DbExpt = New ADODB.Connection
            DbExpt.Open("Provider=Microsoft.Jet.OLEDB.4.0;" &
            "Data Source = " & nome & ";")
        End If
        'Define a ação a ser tomada:
        CriaDBExpt = Msg
        Select Case Msg
            Case 0      'O DB não existe
                'Cria a Tabela Log:
                Sel = "CREATE TABLE Log"
                Sel = Sel & " (dataLog DATETIME"
                Sel = Sel & ", horaLog DATETIME"
                Sel = Sel & ", usuarioLog nvarchar(25)"
                Sel = Sel & ", microLog nvarchar(25)"
                Sel = Sel & ", descricaoLog nvarchar(100))"
                DbExpt.Execute(Sel, , adExecuteNoRecords)
            Case 7      'Os dados da tabela serão eliminados antes de receber os exportados
                Sel = "DELETE * FROM Log"
                DbExpt.Execute(Sel, , adExecuteNoRecords)
        End Select

ECriaDBExpt:
        If Err.Number Then
            MsgBox(Err.Number & " - " & Err.Description)
            CriaDBExpt = Err.Number
        End If

    End Function

    Public Function CriaArqExpt(nome As String) As Integer

        Dim RC As Integer
        Dim I As Integer
        Dim Msg As Integer
        Dim vAux As String
        Dim cdSeguranca1 As SegurancaD.cdSeguranca1

        'Cria uma instância da classe:
        cdSeguranca1 = New SegurancaD.cdSeguranca1

        On Error GoTo ECriaArqExpt

        Msg = 0     'Define a ação a ser tomada
        RC = cdSeguranca1.FileExists(nome)
        If RC <> 0 Then
            'O Arquivo existe:
            vAux = "Deseja preservar os dados existentes?" & vbCrLf & vbCrLf
            vAux = vAux & "Tecle SIM para copiar no final da tabela." & vbCrLf
            vAux = vAux & "Tecle NÃO para limpar os dados antes de exportá-los." & vbCrLf
            vAux = vAux & "Tecle CANCELAR para interromper o processamento."
            Msg = MsgBox(vAux, vbYesNoCancel, "Tabela de LOG já existe")
        Else        'Arquivo Não Existe
            'Cria o Arquivo:
            FileSystem.WriteAllText(nome, "", False)
            'I = FreeFile()                    'Get free file number.
            'Open nome For Output As #I      'Open the file.
            'Close #I
        End If
        'Define a ação a ser tomada:
        CriaArqExpt = Msg
        'Os dados da tabela serão eliminados antes de receber os exportados:
        If Msg = 7 Then
            'Apaga o arquivo:
            Kill(nome)
            'Recria o Arquivo:
            FileSystem.WriteAllText(nome, "", False)
            'I = FreeFile()                    'Get free file number.
            'Open nome For Output As #I      'Open the file.
            'Close #I
        End If

ECriaArqExpt:
        If Err.Number Then
            If Err.Number = 55 Then
                Resume Next
            Else
                CriaArqExpt = Err.Number
            End If
        End If

    End Function

    Public Function inclui(pUsuario As String, vNomeComputador As String, vPalavra As String) As Integer
        'Inclui registro na tabela LOG do sistema

        Dim FuncoesGerais As SegurancaD.FuncoesGerais

        FuncoesGerais = New SegurancaD.FuncoesGerais
        On Error GoTo Einclui
        tsegLog.AddNew()
        tsegLog.Fields("dataLog").Value = FormatDateTime(FuncoesGerais.dataServ(), DateFormat.ShortDate)
        tsegLog.Fields("horaLog").Value = FormatDateTime(FuncoesGerais.dataServ(), DateFormat.ShortTime)
        tsegLog.Fields("usuarioLog").Value = Mid(Trim(pUsuario), 1, 25)
        tsegLog.Fields("microLog").Value = Mid(Trim(vNomeComputador), 1, 20)
        tsegLog.Fields("descricaoLog").Value = Mid(vPalavra, 1, 100)
        tsegLog.Update()
        inclui = 0

Einclui:
        If Err.Number Then
            inclui = Err.Number
            MsgBox("Classe segLog - inclui" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function

    Public Function elimina(vdataLog As Date, vhoraLog As Date, vusuarioLog As String, vmicroLog As String, vDescric As String) As Integer
        '* ****************************************
        '* * Elimina um registro da tabela        *
        '* * Log                                  *
        '* ****************************************

        Dim vSelec As String

        On Error GoTo Eelimina
        elimina = 0     'Se não houver erro
        vSelec = "DELETE FROM Log WHERE "
        vSelec = vSelec & " Convert(datetime, dataLog, 112) = '" & String.Format(vdataLog, "yyyymmdd") & "' "
        'vSelec = vSelec & " AND Convert(datetime, horaLog, 108) = '" & vhoraLog & "' "
        vSelec = vSelec & " AND usuarioLog = '" & vusuarioLog & "' "
        vSelec = vSelec & " AND microLog = '" & vmicroLog & "'"
        vSelec = vSelec & " AND descricaoLog = '" & vDescric & "'"
        Db.Execute(vSelec, , adExecuteNoRecords)

Eelimina:
        If Err.Number Then
            elimina = Err.Number
            MsgBox("Classe ColTpDemissao - elimina" & vbCrLf & Err.Number & " - " & Err.Description)
        End If

    End Function
End Class
