'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'  Persistable = 0  'NotPersistable
'  DataBindingBehavior = 0  'vbNone
'  DataSourceBehavior  = 0  'vbNone
'  MTSTransactionMode  = 0  'NotAnMTSObject
'END
'Attribute VB_Name = "cdSeguranca"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Description = "Provê conexão ao DB Seguranca, e acesso a suas tabelas"
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit On
Imports ADODB.CommandTypeEnum
Imports ADODB.CursorTypeEnum
Imports ADODB.LockTypeEnum
Imports ADODB.ExecuteOptionEnum
Imports ADODB.ObjectStateEnum

<ComClass(cdSeguranca.ClassId, cdSeguranca.InterfaceId, cdSeguranca.EventsId)>
Public Class cdSeguranca

    'Public Db As ADODB.Connection
    'Public tDirTpDiretriz As ADODB.Recordset

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "664501e9-7c28-42a3-989b-8f5537859002"
    Public Const InterfaceId As String = "6d50d7cc-1a12-42ed-b489-1c819d329002"
    Public Const EventsId As String = "2b89fcb2-ff08-4645-899c-5412e42c9002"
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
    Private tUsuario As ADODB.Recordset
    Private tServicos As ADODB.Recordset
    Private tMensagens As ADODB.Recordset
    Private tLog As ADODB.Recordset

    Public Function conectaDB() As Integer

        '* ****************************************
        '* * abreDB = se abre ou não o DB:        *
        '* * 0 = abre o DB                        *
        '* * 1 = não abre o DB                    *
        '* *                                      *
        '* * tipo = forma de abrir as tabelas:    *
        '* * 0 = OpenTable                        *
        '* * 1 = OpenDynaset                      *
        '* ****************************************
        Dim strConnect As String
        Dim vTab As Integer
        Dim cdSeguranca1 As SegurancaD.cdSeguranca1

        On Error GoTo EconectaDB

        conectaDB = 0   'ReturnCode se não houver nenhum problema
        'Cria uma instância de ADODB.Connection:
        Db = New ADODB.Connection
        cdSeguranca1 = New SegurancaD.cdSeguranca1
        'Abre o DB:
        strConnect = cdSeguranca1.LeDADOSsys(1)
        Db.Open(strConnect)

        'Cria as instâncias de ADODB.Recordset
        'tUsuario = New ADODB.Recordset
        tServicos = New ADODB.Recordset
        tMensagens = New ADODB.Recordset
        tLog = New ADODB.Recordset
        'Abre as tabelas:
        vTab = 1
        tUsuario.Open("Usuario", Db, adOpenDynamic, adLockOptimistic, adCmdTableDirect)
        vTab = 2
        tServicos.Open("Servicos", Db, adOpenDynamic, adLockOptimistic, adCmdTableDirect)
        vTab = 3
        tMensagens.Open("Mensagens", Db, adOpenDynamic, adLockOptimistic, adCmdTableDirect)
        vTab = 4
        tLog.Open("Log", Db, adOpenDynamic, adLockOptimistic, adCmdTableDirect)
EconectaDB:
        If Err.Number <> 0 Then
            If Err.Number = 3705 Then
                Select Case vTab
                    Case 1
                        tUsuario.Close()
                        tUsuario.Open("Usuario", Db, adOpenDynamic, adLockOptimistic, adCmdTableDirect)
                    Case 2
                        tServicos.Close()
                        tServicos.Open("Servicos", Db, adOpenDynamic, adLockOptimistic, adCmdTableDirect)
                    Case 3
                        tMensagens.Close()
                        tMensagens.Open("Mensagens", Db, adOpenDynamic, adLockOptimistic, adCmdTableDirect)
                    Case 4
                        tLog.Close()
                        tLog.Open("Log", Db, adOpenDynamic, adLockOptimistic, adCmdTableDirect)
                End Select
                Resume Next
            Else
                conectaDB = Err.Number
                If Err.Number = 3024 Then conectaDB = 1046
                MsgBox("Classe cdSeguranca - conectaDB" & vbCrLf & Err.Number & " - " & Err.Number)
            End If
        End If

    End Function
End Class
