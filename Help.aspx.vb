Imports System.Collections.Specialized
Imports System.xml
Imports System.Security.Principal
Imports System.IO
Public Class Help
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Image1 As System.Web.UI.WebControls.Image
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents Td1 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected myUtility As New Utility
    Protected WithEvents ContentSection As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected WithEvents JSAction As String

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

#Region "Basic"
    'loads the page
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load
        Trace.Warn("Loading Page")
        InitializeComponent()
        Session("JSAction") = ""

        'catch incoming alert messages
        alert(Session("Alert"))

        'clear session
        myUtility.clearSession()

        If (Session.IsNewSession) Then
            Session("Alert") = "Your session has timed out for security purposes.  Please log in again."
        End If

        'get anything on the request line
        myUtility.getRequest(Page.Request.Url.Query())

        'Put user code to initialize the page here
        If (Session("isAuthenticated") <> "True") Then
            Session("CurrentPage") = "./Login.aspx"
            Session("pageSel") = "Login"
            Response.Redirect(Session("CurrentPage"))
        Else
            Session("CurrentPage") = "./Help.aspx"
            Session("pageSel") = "Help"
        End If

        'if the user's password has expired send them to the password reset page
        If (Session("isExpired") = "True") Then
            Session("CurrentPage") = "./Password.aspx"
            Session("pageSel") = "Password"
            Response.Redirect(Session("CurrentPage"))
        End If

        'get the from page
        If Not (IsPostBack) Then
            Session("FromPage") = getFromPage()

            'write content based on the XMLFile URL parameter
            writeContent()
        End If

        Session("JSAction") = JSAction
    End Sub

    'handles error messages
    Public Sub alert(ByVal message As String)
        If (message <> "") Then
            JSAction = "<script type='text/javascript'>window.alert('" & message & "');</script>"
            Session("Alert") = ""
        End If
    End Sub

    'gets the referring or from page and returns it
    Public Function getFromPage() As String
        Dim coll As NameValueCollection
        ' Load ServerVariable collection into NameValueCollection object.
        coll = Request.ServerVariables
        ' Get referring page.
        Return (coll.Item("HTTP_REFERER"))
    End Function
#End Region

#Region "XML Content"
    'wites the page content from the XML file specified on the URL
    Public Sub writeContent()
        Dim reader As XmlReader
        Try
            'get and store the XML file needed
            If (Session("XMLFile") Is Nothing) Then
                reader = New XmlTextReader(MapPath("./FAQ/Introduction.xml"))
            Else
                reader = New XmlTextReader(MapPath(CType(Session("XMLFile"), String)))
            End If

            'process the contents of the XML file into the content labels
            Me.contentSection.Text = ""
            While (reader.Read())
                If (reader.NodeType = XmlNodeType.Element) Then
                    If(reader.LocalName = "Text") Then
                        Me.contentSection.Text &= reader.ReadString
                    End If
                End If
            End While
            reader.Close()
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            reader.Close()
        End Try
    End Sub
#End Region
End Class
