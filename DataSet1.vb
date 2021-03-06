﻿'------------------------------------------------------------------------------
' <autogenerated>
'     This code was generated by a tool.
'     Runtime Version: 1.1.4322.2032
'
'     Changes to this file may cause incorrect behavior and will be lost if 
'     the code is regenerated.
' </autogenerated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On

Imports System
Imports System.Data
Imports System.Runtime.Serialization
Imports System.Xml


<Serializable(),  _
 System.ComponentModel.DesignerCategoryAttribute("code"),  _
 System.Diagnostics.DebuggerStepThrough(),  _
 System.ComponentModel.ToolboxItem(true)>  _
Public Class DataSet1
    Inherits DataSet
    
    Private tablefempUsers As fempUsersDataTable
    
    Public Sub New()
        MyBase.New
        Me.InitClass
        Dim schemaChangedHandler As System.ComponentModel.CollectionChangeEventHandler = AddressOf Me.SchemaChanged
        AddHandler Me.Tables.CollectionChanged, schemaChangedHandler
        AddHandler Me.Relations.CollectionChanged, schemaChangedHandler
    End Sub
    
    Protected Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)
        MyBase.New
        Dim strSchema As String = CType(info.GetValue("XmlSchema", GetType(System.String)),String)
        If (Not (strSchema) Is Nothing) Then
            Dim ds As DataSet = New DataSet
            ds.ReadXmlSchema(New XmlTextReader(New System.IO.StringReader(strSchema)))
            If (Not (ds.Tables("fempUsers")) Is Nothing) Then
                Me.Tables.Add(New fempUsersDataTable(ds.Tables("fempUsers")))
            End If
            Me.DataSetName = ds.DataSetName
            Me.Prefix = ds.Prefix
            Me.Namespace = ds.Namespace
            Me.Locale = ds.Locale
            Me.CaseSensitive = ds.CaseSensitive
            Me.EnforceConstraints = ds.EnforceConstraints
            Me.Merge(ds, false, System.Data.MissingSchemaAction.Add)
            Me.InitVars
        Else
            Me.InitClass
        End If
        Me.GetSerializationData(info, context)
        Dim schemaChangedHandler As System.ComponentModel.CollectionChangeEventHandler = AddressOf Me.SchemaChanged
        AddHandler Me.Tables.CollectionChanged, schemaChangedHandler
        AddHandler Me.Relations.CollectionChanged, schemaChangedHandler
    End Sub
    
    <System.ComponentModel.Browsable(false),  _
     System.ComponentModel.DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Content)>  _
    Public ReadOnly Property fempUsers As fempUsersDataTable
        Get
            Return Me.tablefempUsers
        End Get
    End Property
    
    Public Overrides Function Clone() As DataSet
        Dim cln As DataSet1 = CType(MyBase.Clone,DataSet1)
        cln.InitVars
        Return cln
    End Function
    
    Protected Overrides Function ShouldSerializeTables() As Boolean
        Return false
    End Function
    
    Protected Overrides Function ShouldSerializeRelations() As Boolean
        Return false
    End Function
    
    Protected Overrides Sub ReadXmlSerializable(ByVal reader As XmlReader)
        Me.Reset
        Dim ds As DataSet = New DataSet
        ds.ReadXml(reader)
        If (Not (ds.Tables("fempUsers")) Is Nothing) Then
            Me.Tables.Add(New fempUsersDataTable(ds.Tables("fempUsers")))
        End If
        Me.DataSetName = ds.DataSetName
        Me.Prefix = ds.Prefix
        Me.Namespace = ds.Namespace
        Me.Locale = ds.Locale
        Me.CaseSensitive = ds.CaseSensitive
        Me.EnforceConstraints = ds.EnforceConstraints
        Me.Merge(ds, false, System.Data.MissingSchemaAction.Add)
        Me.InitVars
    End Sub
    
    Protected Overrides Function GetSchemaSerializable() As System.Xml.Schema.XmlSchema
        Dim stream As System.IO.MemoryStream = New System.IO.MemoryStream
        Me.WriteXmlSchema(New XmlTextWriter(stream, Nothing))
        stream.Position = 0
        Return System.Xml.Schema.XmlSchema.Read(New XmlTextReader(stream), Nothing)
    End Function
    
    Friend Sub InitVars()
        Me.tablefempUsers = CType(Me.Tables("fempUsers"),fempUsersDataTable)
        If (Not (Me.tablefempUsers) Is Nothing) Then
            Me.tablefempUsers.InitVars
        End If
    End Sub
    
    Private Sub InitClass()
        Me.DataSetName = "DataSet1"
        Me.Prefix = ""
        Me.Namespace = "http://www.tempuri.org/DataSet1.xsd"
        Me.Locale = New System.Globalization.CultureInfo("en-US")
        Me.CaseSensitive = false
        Me.EnforceConstraints = true
        Me.tablefempUsers = New fempUsersDataTable
        Me.Tables.Add(Me.tablefempUsers)
    End Sub
    
    Private Function ShouldSerializefempUsers() As Boolean
        Return false
    End Function
    
    Private Sub SchemaChanged(ByVal sender As Object, ByVal e As System.ComponentModel.CollectionChangeEventArgs)
        If (e.Action = System.ComponentModel.CollectionChangeAction.Remove) Then
            Me.InitVars
        End If
    End Sub
    
    Public Delegate Sub fempUsersRowChangeEventHandler(ByVal sender As Object, ByVal e As fempUsersRowChangeEvent)
    
    <System.Diagnostics.DebuggerStepThrough()>  _
    Public Class fempUsersDataTable
        Inherits DataTable
        Implements System.Collections.IEnumerable
        
        Private columnUSER_KEY As DataColumn

        Private columnUSER_CODE As DataColumn

        Private columnName As DataColumn

        Private columnHANFORD_ID As DataColumn

        Private columnEMAIL_ADDRESS As DataColumn

        Friend Sub New()
            MyBase.New("fempUsers")
            Me.InitClass()
        End Sub

        Friend Sub New(ByVal table As DataTable)
            MyBase.New(table.TableName)
            If (table.CaseSensitive <> table.DataSet.CaseSensitive) Then
                Me.CaseSensitive = table.CaseSensitive
            End If
            If (table.Locale.ToString <> table.DataSet.Locale.ToString) Then
                Me.Locale = table.Locale
            End If
            If (table.Namespace <> table.DataSet.Namespace) Then
                Me.Namespace = table.Namespace
            End If
            Me.Prefix = table.Prefix
            Me.MinimumCapacity = table.MinimumCapacity
            Me.DisplayExpression = table.DisplayExpression
        End Sub

        <System.ComponentModel.Browsable(False)> _
        Public ReadOnly Property Count() As Integer
            Get
                Return Me.Rows.Count
            End Get
        End Property

        Friend ReadOnly Property USER_KEYColumn() As DataColumn
            Get
                Return Me.columnUSER_KEY
            End Get
        End Property

        Friend ReadOnly Property USER_CODEColumn() As DataColumn
            Get
                Return Me.columnUSER_CODE
            End Get
        End Property

        Friend ReadOnly Property NameColumn() As DataColumn
            Get
                Return Me.columnName
            End Get
        End Property

        Friend ReadOnly Property HANFORD_IDColumn() As DataColumn
            Get
                Return Me.columnHANFORD_ID
            End Get
        End Property

        Friend ReadOnly Property EMAIL_ADDRESSColumn() As DataColumn
            Get
                Return Me.columnEMAIL_ADDRESS
            End Get
        End Property

        Default Public ReadOnly Property Item(ByVal index As Integer) As fempUsersRow
            Get
                Return CType(Me.Rows(index), fempUsersRow)
            End Get
        End Property

        Public Event fempUsersRowChanged As fempUsersRowChangeEventHandler

        Public Event fempUsersRowChanging As fempUsersRowChangeEventHandler

        Public Event fempUsersRowDeleted As fempUsersRowChangeEventHandler

        Public Event fempUsersRowDeleting As fempUsersRowChangeEventHandler

        Public Overloads Sub AddfempUsersRow(ByVal row As fempUsersRow)
            Me.Rows.Add(row)
        End Sub

        Public Overloads Function AddfempUsersRow(ByVal USER_CODE As Integer, ByVal Name As String, ByVal HANFORD_ID As String, ByVal EMAIL_ADDRESS As String) As fempUsersRow
            Dim rowfempUsersRow As fempUsersRow = CType(Me.NewRow, fempUsersRow)
            rowfempUsersRow.ItemArray = New Object() {Nothing, USER_CODE, Name, HANFORD_ID, EMAIL_ADDRESS}
            Me.Rows.Add(rowfempUsersRow)
            Return rowfempUsersRow
        End Function

        Public Function FindByHANFORD_IDEMAIL_ADDRESS(ByVal HANFORD_ID As String, ByVal EMAIL_ADDRESS As String) As fempUsersRow
            Return CType(Me.Rows.Find(New Object() {HANFORD_ID, EMAIL_ADDRESS}), fempUsersRow)
        End Function

        Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
            Return Me.Rows.GetEnumerator
        End Function

        Public Overrides Function Clone() As DataTable
            Dim cln As fempUsersDataTable = CType(MyBase.Clone, fempUsersDataTable)
            cln.InitVars()
            Return cln
        End Function

        Protected Overrides Function CreateInstance() As DataTable
            Return New fempUsersDataTable
        End Function

        Friend Sub InitVars()
            Me.columnUSER_KEY = Me.Columns("USER_KEY")
            Me.columnUSER_CODE = Me.Columns("USER_CODE")
            Me.columnName = Me.Columns("Name")
            Me.columnHANFORD_ID = Me.Columns("HANFORD_ID")
            Me.columnEMAIL_ADDRESS = Me.Columns("EMAIL_ADDRESS")
        End Sub

        Private Sub InitClass()
            Me.columnUSER_KEY = New DataColumn("USER_KEY", GetType(System.Int32), Nothing, System.Data.MappingType.Element)
            Me.Columns.Add(Me.columnUSER_KEY)
            Me.columnUSER_CODE = New DataColumn("USER_CODE", GetType(System.Int32), Nothing, System.Data.MappingType.Element)
            Me.Columns.Add(Me.columnUSER_CODE)
            Me.columnName = New DataColumn("Name", GetType(System.String), Nothing, System.Data.MappingType.Element)
            Me.Columns.Add(Me.columnName)
            Me.columnHANFORD_ID = New DataColumn("HANFORD_ID", GetType(System.String), Nothing, System.Data.MappingType.Element)
            Me.Columns.Add(Me.columnHANFORD_ID)
            Me.columnEMAIL_ADDRESS = New DataColumn("EMAIL_ADDRESS", GetType(System.String), Nothing, System.Data.MappingType.Element)
            Me.Columns.Add(Me.columnEMAIL_ADDRESS)
            Me.Constraints.Add(New UniqueConstraint("Constraint1", New DataColumn() {Me.columnHANFORD_ID, Me.columnEMAIL_ADDRESS}, True))
            Me.columnUSER_KEY.AutoIncrement = True
            Me.columnUSER_KEY.AllowDBNull = False
            Me.columnUSER_KEY.ReadOnly = True
            Me.columnUSER_CODE.AllowDBNull = False
            Me.columnName.ReadOnly = True
            Me.columnHANFORD_ID.AllowDBNull = False
            Me.columnEMAIL_ADDRESS.AllowDBNull = False
        End Sub

        Public Function NewfempUsersRow() As fempUsersRow
            Return CType(Me.NewRow, fempUsersRow)
        End Function

        Protected Overrides Function NewRowFromBuilder(ByVal builder As DataRowBuilder) As DataRow
            Return New fempUsersRow(builder)
        End Function

        Protected Overrides Function GetRowType() As System.Type
            Return GetType(fempUsersRow)
        End Function

        Protected Overrides Sub OnRowChanged(ByVal e As DataRowChangeEventArgs)
            MyBase.OnRowChanged(e)
            If (Not (Me.fempUsersRowChangedEvent) Is Nothing) Then
                RaiseEvent fempUsersRowChanged(Me, New fempUsersRowChangeEvent(CType(e.Row, fempUsersRow), e.Action))
            End If
        End Sub

        Protected Overrides Sub OnRowChanging(ByVal e As DataRowChangeEventArgs)
            MyBase.OnRowChanging(e)
            If (Not (Me.fempUsersRowChangingEvent) Is Nothing) Then
                RaiseEvent fempUsersRowChanging(Me, New fempUsersRowChangeEvent(CType(e.Row, fempUsersRow), e.Action))
            End If
        End Sub

        Protected Overrides Sub OnRowDeleted(ByVal e As DataRowChangeEventArgs)
            MyBase.OnRowDeleted(e)
            If (Not (Me.fempUsersRowDeletedEvent) Is Nothing) Then
                RaiseEvent fempUsersRowDeleted(Me, New fempUsersRowChangeEvent(CType(e.Row, fempUsersRow), e.Action))
            End If
        End Sub

        Protected Overrides Sub OnRowDeleting(ByVal e As DataRowChangeEventArgs)
            MyBase.OnRowDeleting(e)
            If (Not (Me.fempUsersRowDeletingEvent) Is Nothing) Then
                RaiseEvent fempUsersRowDeleting(Me, New fempUsersRowChangeEvent(CType(e.Row, fempUsersRow), e.Action))
            End If
        End Sub

        Public Sub RemovefempUsersRow(ByVal row As fempUsersRow)
            Me.Rows.Remove(row)
        End Sub
    End Class

    <System.Diagnostics.DebuggerStepThrough()> _
    Public Class fempUsersRow
        Inherits DataRow

        Private tablefempUsers As fempUsersDataTable

        Friend Sub New(ByVal rb As DataRowBuilder)
            MyBase.New(rb)
            Me.tablefempUsers = CType(Me.Table, fempUsersDataTable)
        End Sub

        Public Property USER_KEY() As Integer
            Get
                Return CType(Me(Me.tablefempUsers.USER_KEYColumn), Integer)
            End Get
            Set(ByVal value As Integer)
                Me(Me.tablefempUsers.USER_KEYColumn) = value
            End Set
        End Property

        Public Property USER_CODE() As Integer
            Get
                Return CType(Me(Me.tablefempUsers.USER_CODEColumn), Integer)
            End Get
            Set(ByVal value As Integer)
                Me(Me.tablefempUsers.USER_CODEColumn) = value
            End Set
        End Property

        Public Property Name() As String
            Get
                Try
                    Return CType(Me(Me.tablefempUsers.NameColumn), String)
                Catch e As InvalidCastException
                    Throw New StrongTypingException("Cannot get value because it is DBNull.", e)
                End Try
            End Get
            Set(ByVal value As String)
                Me(Me.tablefempUsers.NameColumn) = value
            End Set
        End Property

        Public Property HANFORD_ID() As String
            Get
                Return CType(Me(Me.tablefempUsers.HANFORD_IDColumn), String)
            End Get
            Set(ByVal value As String)
                Me(Me.tablefempUsers.HANFORD_IDColumn) = value
            End Set
        End Property

        Public Property EMAIL_ADDRESS() As String
            Get
                Return CType(Me(Me.tablefempUsers.EMAIL_ADDRESSColumn), String)
            End Get
            Set(ByVal value As String)
                Me(Me.tablefempUsers.EMAIL_ADDRESSColumn) = value
            End Set
        End Property

        Public Function IsNameNull() As Boolean
            Return Me.IsNull(Me.tablefempUsers.NameColumn)
        End Function

        Public Sub SetNameNull()
            Me(Me.tablefempUsers.NameColumn) = System.Convert.DBNull
        End Sub
    End Class
    
    <System.Diagnostics.DebuggerStepThrough()>  _
    Public Class fempUsersRowChangeEvent
        Inherits EventArgs
        
        Private eventRow As fempUsersRow
        
        Private eventAction As DataRowAction
        
        Public Sub New(ByVal row As fempUsersRow, ByVal action As DataRowAction)
            MyBase.New
            Me.eventRow = row
            Me.eventAction = action
        End Sub
        
        Public ReadOnly Property Row As fempUsersRow
            Get
                Return Me.eventRow
            End Get
        End Property
        
        Public ReadOnly Property Action As DataRowAction
            Get
                Return Me.eventAction
            End Get
        End Property
    End Class
End Class
