Imports System.IO
Public Class Picture
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Picture))
        Me.sqlImage = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand
        Me.dsCommon = New System.Data.DataSet
        Me.commonConn = New System.Data.SqlClient.SqlConnection
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'sqlImage
        '
        Me.sqlImage.SelectCommand = Me.SqlSelectCommand1
        Me.sqlImage.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_QUESTION_LIST_ITEM", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("LIST_ITEM_IMAGE", "LIST_ITEM_IMAGE"), New System.Data.Common.DataColumnMapping("LIST_ITEM_IMAGE_TYPE", "LIST_ITEM_IMAGE_TYPE")})})
        '
        'SqlSelectCommand1
        '
        Me.SqlSelectCommand1.CommandText = resources.GetString("SqlSelectCommand1.CommandText")
        Me.SqlSelectCommand1.Connection = Me.commonConn
        '
        'dsCommon
        '
        Me.dsCommon.DataSetName = "NewDataSet"
        Me.dsCommon.Locale = New System.Globalization.CultureInfo("en-US")
        '
        'commonConn
        '
        Me.commonConn.FireInfoMessageEventOnUserErrors = False
        '
        'commonConn
        '
        Me.commonConn.FireInfoMessageEventOnUserErrors = False
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object
    Protected WithEvents sqlImage As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents dsCommon As System.Data.DataSet
    Protected WithEvents SqlSelectCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents commonConn As System.Data.SqlClient.SqlConnection
    Protected myUtility As New Utility
    Protected requestItems As Array

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

#Region "Basic"
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load
        'generate session variables for the incoming query string
        Try
            'clear session
            myUtility.clearSession()
            Response.Expires = 0
            Response.Buffer = True

            Dim requestString As String = Page.Request.Url.Query()
            requestString = requestString.Substring(1)
            requestItems = requestString.Split("&")
            Dim str As String
            Dim Template As String
            Dim Question As String
            Dim ListItem As String
            Dim temparr As Array
            temparr = CType(requestItems(0), String).Split("=")
            Template = CType(temparr(1), String)
            temparr = CType(requestItems(1), String).Split("=")
            Question = CType(temparr(1), String)
            temparr = CType(requestItems(2), String).Split("=")
            ListItem = CType(temparr(1), String)

            Trace.Warn(Template & " " & Question & " " & ListItem)
            'get the particular data item image and binary write it to the window

            'Set the connection based on the machine
            myUtility.getConnection(Me.commonConn)

            Me.SqlSelectCommand1.CommandText &= " WHERE FL.TEMPLATE_KEY=" & Template & " AND FL.QUESTION_ID=" & _
                Question & " AND FL.LIST_ITEM_ID=" & ListItem
            Me.SqlSelectCommand1.CommandText = Me.SqlSelectCommand1.CommandText.Replace("VARCONN.", myUtility.getExtension())
            Me.SqlSelectCommand1.Connection = Me.commonConn
            Me.sqlImage.Fill(Me.dsCommon, "Image")
            Trace.Warn(Me.dsCommon.Tables("Image").Rows.Count)

            'display the image in the window
            Response.ContentType = Me.dsCommon.Tables("Image").Rows(0).Item("LIST_ITEM_IMAGE_TYPE")
            Response.BinaryWrite(Me.dsCommon.Tables("Image").Rows(0).Item("LIST_ITEM_IMAGE"))
        Catch ex As Exception
            Session("Alert") = "There is no picture to display. Attach a picture and try again."
            'Response.Write("<script type='text/javascript'>window.opener.location.href = window.opener.location.href;</script>")
            Response.Write("<script type='text/javascript'>window.close()</script>")
        End Try
    End Sub
#End Region
End Class
