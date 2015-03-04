<%@ Page Trace="false" Language="vb" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" Codebehind="Help.aspx.vb" Inherits="SurveyAdmin.Help" smartNavigation="False" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder" Runat="Server">
<%=Session("JSAction")%>
<h1>User Help</h1>
<p></p>
<asp:label id="ContentSection" Runat="server"></asp:label>
<p></p>
</asp:Content>
