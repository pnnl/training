Imports System.Collections.Specialized
Public Class Groups
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.SqlGroups = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlDeleteCommand1 = New System.Data.SqlClient.SqlCommand
        Me.commonConn = New System.Data.SqlClient.SqlConnection
        Me.SqlInsertCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlUpdateCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlUsers = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlDeleteCommand2 = New System.Data.SqlClient.SqlCommand
        Me.SqlInsertCommand2 = New System.Data.SqlClient.SqlCommand
        Me.SqlSelectCommand2 = New System.Data.SqlClient.SqlCommand
        Me.SqlUpdateCommand2 = New System.Data.SqlClient.SqlCommand
        Me.dsCommon = New System.Data.DataSet
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'SqlGroups
        '
        Me.SqlGroups.DeleteCommand = Me.SqlDeleteCommand1
        Me.SqlGroups.InsertCommand = Me.SqlInsertCommand1
        Me.SqlGroups.SelectCommand = Me.SqlSelectCommand1
        Me.SqlGroups.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "fempGroups", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("seqSurvey", "seqSurvey"), New System.Data.Common.DataColumnMapping("intGroupID", "intGroupID"), New System.Data.Common.DataColumnMapping("strGroupName", "strGroupName"), New System.Data.Common.DataColumnMapping("strUserList", "strUserList")})})
        Me.SqlGroups.UpdateCommand = Me.SqlUpdateCommand1
        '
        'SqlDeleteCommand1
        '
        Me.SqlDeleteCommand1.CommandText = "DELETE FROM VARCONN.fempGroups WHERE (intGroupID = @Original_intGroupID) AND (seqSurv" & _
        "ey = @Original_seqSurvey) AND (strGroupName = @Original_strGroupName) AND (strUs" & _
        "erList = @Original_strUserList)"
        Me.SqlDeleteCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_intGroupID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "intGroupID", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_seqSurvey", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "seqSurvey", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_strGroupName", System.Data.SqlDbType.VarChar, 300, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "strGroupName", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_strUserList", System.Data.SqlDbType.VarChar, 8000, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "strUserList", System.Data.DataRowVersion.Original, Nothing))
        '
        'SqlInsertCommand1
        '
        Me.SqlInsertCommand1.CommandText = "INSERT INTO VARCONN.fempGroups(seqSurvey, intGroupID, strGroupName, strUserList) VALU" & _
        "ES (@seqSurvey, @intGroupID, @strGroupName, @strUserList); SELECT seqSurvey, int" & _
        "GroupID, strGroupName, strUserList FROM VARCONN.fempGroups WHERE (intGroupID = @intG" & _
        "roupID) AND (seqSurvey = @seqSurvey)"
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@seqSurvey", System.Data.SqlDbType.Int, 4, "seqSurvey"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@intGroupID", System.Data.SqlDbType.Int, 4, "intGroupID"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@strGroupName", System.Data.SqlDbType.VarChar, 300, "strGroupName"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@strUserList", System.Data.SqlDbType.VarChar, 8000, "strUserList"))
        '
        'SqlSelectCommand1
        '
        Me.SqlSelectCommand1.CommandText = "SELECT seqSurvey, intGroupID, strGroupName, strUserList FROM VARCONN.fempGroups"
        Me.SqlSelectCommand1.Connection = Me.commonConn
        '
        'SqlUpdateCommand1
        '
        Me.SqlUpdateCommand1.CommandText = "UPDATE VARCONN.fempGroups SET seqSurvey = @seqSurvey, intGroupID = @intGroupID, strGr" & _
        "oupName = @strGroupName, strUserList = @strUserList WHERE (intGroupID = @Origina" & _
        "l_intGroupID) AND (seqSurvey = @Original_seqSurvey) AND (strGroupName = @Origina" & _
        "l_strGroupName) AND (strUserList = @Original_strUserList); SELECT seqSurvey, int" & _
        "GroupID, strGroupName, strUserList FROM VARCONN.fempGroups WHERE (intGroupID = @intG" & _
        "roupID) AND (seqSurvey = @seqSurvey)"
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@seqSurvey", System.Data.SqlDbType.Int, 4, "seqSurvey"))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@intGroupID", System.Data.SqlDbType.Int, 4, "intGroupID"))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@strGroupName", System.Data.SqlDbType.VarChar, 300, "strGroupName"))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@strUserList", System.Data.SqlDbType.VarChar, 8000, "strUserList"))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_intGroupID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "intGroupID", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_seqSurvey", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "seqSurvey", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_strGroupName", System.Data.SqlDbType.VarChar, 300, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "strGroupName", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_strUserList", System.Data.SqlDbType.VarChar, 8000, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "strUserList", System.Data.DataRowVersion.Original, Nothing))
        '
        'SqlUsers
        '
        Me.SqlUsers.DeleteCommand = Me.SqlDeleteCommand2
        Me.SqlUsers.InsertCommand = Me.SqlInsertCommand2
        Me.SqlUsers.SelectCommand = Me.SqlSelectCommand2
        Me.SqlUsers.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "fempUsers", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("seqUserID", "seqUserID"), New System.Data.Common.DataColumnMapping("intUserType", "intUserType"), New System.Data.Common.DataColumnMapping("strLastName", "strLastName"), New System.Data.Common.DataColumnMapping("strFirstName", "strFirstName"), New System.Data.Common.DataColumnMapping("strMiddleName", "strMiddleName"), New System.Data.Common.DataColumnMapping("strNameSuffix", "strNameSuffix"), New System.Data.Common.DataColumnMapping("strAuthenticator", "strAuthenticator"), New System.Data.Common.DataColumnMapping("strWorkPhone", "strWorkPhone"), New System.Data.Common.DataColumnMapping("strHID", "strHID"), New System.Data.Common.DataColumnMapping("strEmail", "strEmail"), New System.Data.Common.DataColumnMapping("datDateCreated", "datDateCreated"), New System.Data.Common.DataColumnMapping("datDateUsed", "datDateUsed"), New System.Data.Common.DataColumnMapping("intCreatorID", "intCreatorID"), New System.Data.Common.DataColumnMapping("blnActive", "blnActive"), New System.Data.Common.DataColumnMapping("blnTraining", "blnTraining"), New System.Data.Common.DataColumnMapping("strCompany", "strCompany"), New System.Data.Common.DataColumnMapping("blnNoEmail", "blnNoEmail")})})
        Me.SqlUsers.UpdateCommand = Me.SqlUpdateCommand2
        '
        'SqlDeleteCommand2
        '
        Me.SqlDeleteCommand2.CommandText = "DELETE FROM VARCONN.fempUsers WHERE (strEmail = @Original_strEmail) AND (strHID = @Or" & _
        "iginal_strHID) AND (blnActive = @Original_blnActive) AND (blnNoEmail = @Original" & _
        "_blnNoEmail) AND (blnTraining = @Original_blnTraining) AND (datDateCreated = @Or" & _
        "iginal_datDateCreated) AND (datDateUsed = @Original_datDateUsed) AND (intCreator" & _
        "ID = @Original_intCreatorID) AND (intUserType = @Original_intUserType) AND (seqU" & _
        "serID = @Original_seqUserID) AND (strAuthenticator = @Original_strAuthenticator)" & _
        " AND (strCompany = @Original_strCompany) AND (strFirstName = @Original_strFirstN" & _
        "ame) AND (strLastName = @Original_strLastName) AND (strMiddleName = @Original_st" & _
        "rMiddleName) AND (strNameSuffix = @Original_strNameSuffix) AND (strWorkPhone = @" & _
        "Original_strWorkPhone)"
        Me.SqlDeleteCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_strEmail", System.Data.SqlDbType.VarChar, 128, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "strEmail", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_strHID", System.Data.SqlDbType.VarChar, 12, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "strHID", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_blnActive", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "blnActive", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_blnNoEmail", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "blnNoEmail", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_blnTraining", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "blnTraining", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_datDateCreated", System.Data.SqlDbType.DateTime, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "datDateCreated", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_datDateUsed", System.Data.SqlDbType.DateTime, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "datDateUsed", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_intCreatorID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "intCreatorID", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_intUserType", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "intUserType", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_seqUserID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "seqUserID", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_strAuthenticator", System.Data.SqlDbType.VarChar, 64, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "strAuthenticator", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_strCompany", System.Data.SqlDbType.VarChar, 64, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "strCompany", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_strFirstName", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "strFirstName", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_strLastName", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "strLastName", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_strMiddleName", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "strMiddleName", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_strNameSuffix", System.Data.SqlDbType.VarChar, 5, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "strNameSuffix", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlDeleteCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_strWorkPhone", System.Data.SqlDbType.VarChar, 32, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "strWorkPhone", System.Data.DataRowVersion.Original, Nothing))
        '
        'SqlInsertCommand2
        '
        Me.SqlInsertCommand2.CommandText = "INSERT INTO VARCONN.fempUsers(intUserType, strLastName, strFirstName, strMiddleName, " & _
        "strNameSuffix, strAuthenticator, strWorkPhone, strHID, strEmail, datDateCreated," & _
        " datDateUsed, intCreatorID, blnActive, blnTraining, strCompany, blnNoEmail) VALU" & _
        "ES (@intUserType, @strLastName, @strFirstName, @strMiddleName, @strNameSuffix, @" & _
        "strAuthenticator, @strWorkPhone, @strHID, @strEmail, @datDateCreated, @datDateUs" & _
        "ed, @intCreatorID, @blnActive, @blnTraining, @strCompany, @blnNoEmail); SELECT s" & _
        "eqUserID, intUserType, strLastName, strFirstName, strMiddleName, strNameSuffix, " & _
        "strAuthenticator, strWorkPhone, strHID, strEmail, datDateCreated, datDateUsed, i" & _
        "ntCreatorID, blnActive, blnTraining, strCompany, blnNoEmail FROM VARCONN.fempUsers W" & _
        "HERE (strEmail = @strEmail) AND (strHID = @strHID)"
        Me.SqlInsertCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@intUserType", System.Data.SqlDbType.Int, 4, "intUserType"))
        Me.SqlInsertCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@strLastName", System.Data.SqlDbType.VarChar, 50, "strLastName"))
        Me.SqlInsertCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@strFirstName", System.Data.SqlDbType.VarChar, 50, "strFirstName"))
        Me.SqlInsertCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@strMiddleName", System.Data.SqlDbType.VarChar, 50, "strMiddleName"))
        Me.SqlInsertCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@strNameSuffix", System.Data.SqlDbType.VarChar, 5, "strNameSuffix"))
        Me.SqlInsertCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@strAuthenticator", System.Data.SqlDbType.VarChar, 64, "strAuthenticator"))
        Me.SqlInsertCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@strWorkPhone", System.Data.SqlDbType.VarChar, 32, "strWorkPhone"))
        Me.SqlInsertCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@strHID", System.Data.SqlDbType.VarChar, 12, "strHID"))
        Me.SqlInsertCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@strEmail", System.Data.SqlDbType.VarChar, 128, "strEmail"))
        Me.SqlInsertCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@datDateCreated", System.Data.SqlDbType.DateTime, 8, "datDateCreated"))
        Me.SqlInsertCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@datDateUsed", System.Data.SqlDbType.DateTime, 8, "datDateUsed"))
        Me.SqlInsertCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@intCreatorID", System.Data.SqlDbType.Int, 4, "intCreatorID"))
        Me.SqlInsertCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@blnActive", System.Data.SqlDbType.Bit, 1, "blnActive"))
        Me.SqlInsertCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@blnTraining", System.Data.SqlDbType.Bit, 1, "blnTraining"))
        Me.SqlInsertCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@strCompany", System.Data.SqlDbType.VarChar, 64, "strCompany"))
        Me.SqlInsertCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@blnNoEmail", System.Data.SqlDbType.Bit, 1, "blnNoEmail"))
        '
        'SqlSelectCommand2
        '
        Me.SqlSelectCommand2.CommandText = "SELECT seqUserID, intUserType, strLastName, strFirstName, strMiddleName, strNameS" & _
        "uffix, strAuthenticator, strWorkPhone, strHID, strEmail, datDateCreated, datDate" & _
        "Used, intCreatorID, blnActive, blnTraining, strCompany, blnNoEmail FROM VARCONN.femp" & _
        "Users"
        '
        'SqlUpdateCommand2
        '
        Me.SqlUpdateCommand2.CommandText = "UPDATE VARCONN.fempUsers SET intUserType = @intUserType, strLastName = @strLastName, " & _
        "strFirstName = @strFirstName, strMiddleName = @strMiddleName, strNameSuffix = @s" & _
        "trNameSuffix, strAuthenticator = @strAuthenticator, strWorkPhone = @strWorkPhone" & _
        ", strHID = @strHID, strEmail = @strEmail, datDateCreated = @datDateCreated, datD" & _
        "ateUsed = @datDateUsed, intCreatorID = @intCreatorID, blnActive = @blnActive, bl" & _
        "nTraining = @blnTraining, strCompany = @strCompany, blnNoEmail = @blnNoEmail WHE" & _
        "RE (strEmail = @Original_strEmail) AND (strHID = @Original_strHID) AND (blnActiv" & _
        "e = @Original_blnActive) AND (blnNoEmail = @Original_blnNoEmail) AND (blnTrainin" & _
        "g = @Original_blnTraining) AND (datDateCreated = @Original_datDateCreated) AND (" & _
        "datDateUsed = @Original_datDateUsed) AND (intCreatorID = @Original_intCreatorID)" & _
        " AND (intUserType = @Original_intUserType) AND (strAuthenticator = @Original_str" & _
        "Authenticator) AND (strCompany = @Original_strCompany) AND (strFirstName = @Orig" & _
        "inal_strFirstName) AND (strLastName = @Original_strLastName) AND (strMiddleName " & _
        "= @Original_strMiddleName) AND (strNameSuffix = @Original_strNameSuffix) AND (st" & _
        "rWorkPhone = @Original_strWorkPhone); SELECT seqUserID, intUserType, strLastName" & _
        ", strFirstName, strMiddleName, strNameSuffix, strAuthenticator, strWorkPhone, st" & _
        "rHID, strEmail, datDateCreated, datDateUsed, intCreatorID, blnActive, blnTrainin" & _
        "g, strCompany, blnNoEmail FROM VARCONN.fempUsers WHERE (strEmail = @strEmail) AND (s" & _
        "trHID = @strHID)"
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@intUserType", System.Data.SqlDbType.Int, 4, "intUserType"))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@strLastName", System.Data.SqlDbType.VarChar, 50, "strLastName"))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@strFirstName", System.Data.SqlDbType.VarChar, 50, "strFirstName"))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@strMiddleName", System.Data.SqlDbType.VarChar, 50, "strMiddleName"))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@strNameSuffix", System.Data.SqlDbType.VarChar, 5, "strNameSuffix"))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@strAuthenticator", System.Data.SqlDbType.VarChar, 64, "strAuthenticator"))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@strWorkPhone", System.Data.SqlDbType.VarChar, 32, "strWorkPhone"))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@strHID", System.Data.SqlDbType.VarChar, 12, "strHID"))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@strEmail", System.Data.SqlDbType.VarChar, 128, "strEmail"))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@datDateCreated", System.Data.SqlDbType.DateTime, 8, "datDateCreated"))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@datDateUsed", System.Data.SqlDbType.DateTime, 8, "datDateUsed"))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@intCreatorID", System.Data.SqlDbType.Int, 4, "intCreatorID"))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@blnActive", System.Data.SqlDbType.Bit, 1, "blnActive"))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@blnTraining", System.Data.SqlDbType.Bit, 1, "blnTraining"))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@strCompany", System.Data.SqlDbType.VarChar, 64, "strCompany"))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@blnNoEmail", System.Data.SqlDbType.Bit, 1, "blnNoEmail"))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_strEmail", System.Data.SqlDbType.VarChar, 128, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "strEmail", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_strHID", System.Data.SqlDbType.VarChar, 12, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "strHID", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_blnActive", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "blnActive", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_blnNoEmail", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "blnNoEmail", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_blnTraining", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "blnTraining", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_datDateCreated", System.Data.SqlDbType.DateTime, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "datDateCreated", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_datDateUsed", System.Data.SqlDbType.DateTime, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "datDateUsed", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_intCreatorID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "intCreatorID", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_intUserType", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "intUserType", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_strAuthenticator", System.Data.SqlDbType.VarChar, 64, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "strAuthenticator", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_strCompany", System.Data.SqlDbType.VarChar, 64, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "strCompany", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_strFirstName", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "strFirstName", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_strLastName", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "strLastName", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_strMiddleName", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "strMiddleName", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_strNameSuffix", System.Data.SqlDbType.VarChar, 5, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "strNameSuffix", System.Data.DataRowVersion.Original, Nothing))
        Me.SqlUpdateCommand2.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Original_strWorkPhone", System.Data.SqlDbType.VarChar, 32, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "strWorkPhone", System.Data.DataRowVersion.Original, Nothing))
        '
        'dsCommon
        '
        Me.dsCommon.DataSetName = "NewDataSet"
        Me.dsCommon.Locale = New System.Globalization.CultureInfo("en-US")
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Protected WithEvents imgTitle As System.Web.UI.WebControls.Image
    Protected WithEvents lblTemplate As System.Web.UI.WebControls.Label
    Protected WithEvents lblQuestion As System.Web.UI.WebControls.Label
    Protected WithEvents ddlPageOptions As System.Web.UI.WebControls.DropDownList
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected WithEvents hlpTemplatePreview As System.Web.UI.WebControls.Image
    Protected WithEvents btnTemplatePreview As System.Web.UI.WebControls.Button
    Protected WithEvents hlpGroupName As System.Web.UI.WebControls.Image
    Protected WithEvents lblGroupName As System.Web.UI.WebControls.Label
    Protected WithEvents hlpGroupList As System.Web.UI.WebControls.Image
    Protected WithEvents lblGroupList As System.Web.UI.WebControls.Label
    Protected WithEvents txtGroupList As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtGroupName As System.Web.UI.WebControls.TextBox
    Protected WithEvents hlpLevel0 As System.Web.UI.WebControls.Image
    Protected WithEvents lblLevel0None As System.Web.UI.WebControls.Label
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents Image1 As System.Web.UI.WebControls.Image
    Protected WithEvents btnStart As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnPrev As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnReload As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnNext As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnLast As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnNew As System.Web.UI.WebControls.ImageButton
    Protected WithEvents dgGroup As System.Web.UI.WebControls.DataGrid
    Protected WithEvents dgUsers As System.Web.UI.WebControls.DataGrid
    Protected WithEvents SqlSelectCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlInsertCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlUpdateCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlDeleteCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlGroups As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents SqlUsers As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents SqlSelectCommand2 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlInsertCommand2 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlUpdateCommand2 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlDeleteCommand2 As System.Data.SqlClient.SqlCommand
    Protected WithEvents dsCommon As System.Data.DataSet
    Protected navButtons As Collection
    Protected requestItems As Array
    Protected WithEvents lblNavBar As System.Web.UI.WebControls.Label
    Protected myNavBar As Navigation = New Navigation
    Protected myUtility As Utility = New Utility
    Protected mailList As String
    Protected WithEvents imgA1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgB1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgC1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgD1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgE1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgF1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgG1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgH1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgI1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgJ1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgK1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgL1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgM1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgN1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgO1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgP1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgQ1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgR1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgS1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgT1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgU1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgV1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgW1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgX1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgY1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgZ1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgA2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgB2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgC2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgD2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgE2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgF2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgG2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgH2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgI2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgJ2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgK2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgL2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgM2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgN2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgO2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgP2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgQ2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgR2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgS2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgT2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgU2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgV2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgW2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgX2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgY2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgZ2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents lblInGroup As System.Web.UI.WebControls.Label
    Protected WithEvents btnInsert1 As System.Web.UI.WebControls.Button
    Protected WithEvents btnUpdate1 As System.Web.UI.WebControls.Button
    Protected WithEvents btnDelete1 As System.Web.UI.WebControls.Button
    Protected WithEvents btnReset1 As System.Web.UI.WebControls.Button
    Protected sqlCommonAdapter As New System.Data.SqlClient.SqlDataAdapter
    Protected sqlCommonAction As New System.Data.SqlClient.SqlCommand
    Protected WithEvents btnUp As System.Web.UI.WebControls.Button
    Protected WithEvents ddlPageOptions2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents btnInsert2 As System.Web.UI.WebControls.Button
    Protected WithEvents btnUpdate2 As System.Web.UI.WebControls.Button
    Protected WithEvents btnDelete2 As System.Web.UI.WebControls.Button
    Protected WithEvents btnReset2 As System.Web.UI.WebControls.Button
    Protected WithEvents commonConn As System.Data.SqlClient.SqlConnection
    Protected pageOptions As Integer
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
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Trace.Warn("Loading Page")
        Session("JSAction") = ""

        'catch incoming alert messages
        alert(Session("Alert"))

        'clear session
        myUtility.clearSession()

        If (Session.IsNewSession) Then
            Session("Alert") = "Your session has timed out for security purposes.  Please log in again."
        End If

        'get the from page
        If Not (IsPostBack) Then
            Session("FromPage") = getFromPage()
        End If

        'set up a production connection to check on user ownership
        Me.commonConn = myUtility.getConnection(Me.commonConn)

        'determine survey ownership if not defined
        If (Session("isSurveyOwner") Is Nothing) Then
            sqlCommonAction.Connection = Me.commonConn
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            myUtility.determineUserOwnership(sqlCommonAction, sqlCommonAdapter, False)
        End If

        'Put user code to initialize the page here
        If (Session("isAuthenticated") <> "True") Then
            Session("CurrentPage") = "./Login.aspx"
            Session("pageSel") = "Login"
            redirect(Session("CurrentPage"))
        ElseIf (myUtility.checkUserType(2, CType(Session("isSurveyOwner"), Boolean), CType(Session("isSurveyDelegate"), Boolean), True)) Then
            redirect(Session("CurrentPage"))
        Else
            Session("CurrentPage") = "./Survey.aspx"
            Session("pageSel") = "Template"
        End If

        'if the user's password has expired send them to the password reset page
        If (Session("isExpired") = "True") Then
            Session("CurrentPage") = "./Password.aspx"
            Session("pageSel") = "Password"
            redirect(Session("CurrentPage"))
        End If

        'generate collection of navigational buttons
        myUtility.makeNavCollection(Me.navButtons, Me.btnStart, Me.btnPrev, Me.btnReload, Me.btnNext, Me.btnLast, Me.btnNew)

        'Check for user selection from the comments list if we're sure it's not a simple post-back
        If Not (IsPostBack) Then
            'get anything we can off the request line if we're not coming back from processing an e-mail address
            If (Session("ProcessEmail") <> "True") Then
                myUtility.getRequest(Page.Request.Url.Query())
            End If

            If (Session("Alpha") Is Nothing Or Session("Alpha") = "") Then
                Session("Alpha") = "A"
            End If

            'load the page data
            If Not (loadData()) Then
                alert("Failed to load the group information for this survey.")
            Else
                'to determine what of the nav bar buttons should be available
                myNavBar.wnb_MoveTo(Me.navButtons, Session("GroupsMax"), Session("GroupsRow"), Switch)
            End If

            'get the template name - mr sneaky user has access and used a link
            If (Session("TemplateName") Is Nothing Or Session("TemplateName") = "") Then
                If Not (setTemplateName()) Then
                    alert("Unable to locate template name.")
                End If
            End If

            'get the survey name - mr sneaky user has access and used a link
            If (Session("SurveyName") Is Nothing Or Session("SurveyName") = "") Then
                If Not (setSurveyName()) Then
                    alert("Unable to locate survey name.")
                End If
            End If

            'Populate the controls on the page
            setControls()

            'set the page options
            myUtility.optionPreSet(Session("intGroupID"), Session("GroupsMax"), Me.pageOptions)

            Session("PageOptions") = Me.pageOptions

            Me.btnInsert1.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
            Me.btnInsert2.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
            Me.btnUpdate1.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
            Me.btnUpdate2.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
            Me.btnDelete1.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
            Me.btnDelete2.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
            Me.btnReset1.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
            Me.btnReset2.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")

            'process an incoming user input request from the person input page
            If (Session("ProcessEmail") = "True") Then
                Me.txtGroupList.Text = Session("PersonInputEmailReturn")
                Dim failure As Boolean = False
                If Not (doInsertUpdate("Update")) Then
                    failure = True
                End If
                Session("ProcessEmail") = "False"

                Dim blnContinue As Boolean = True
                'determine if we need to reload the data
                If Not (loadData(True)) Then
                    blnContinue = False
                    alert("Failed to load the groups for the current survey.")
                End If

                If (blnContinue) Then
                    'to determine what of the nav bar buttons should be available
                    myNavBar.wnb_MoveTo(Me.navButtons, Session("GroupsMax"), Session("GroupsRow"), Switch)

                    'reset the page controls
                    If Not (failure) Then
                        setControls()
                    End If
                    Session("TransExists") = False

                    'set the page options
                    myUtility.optionPreSet(Session("intGroupID"), Session("GroupsMax"), Me.pageOptions)

                    Session("PageOptions") = Me.pageOptions
                End If
            End If
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

    'Set location and re-direct
    Private Sub redirect(Optional ByVal currentPage As String = "")
        If (currentPage = "") Then
            Session("CurrentPage") = Session("FromPage")
            Response.Redirect(Session("CurrentPage"))
        Else
            Session("CurrentPage") = currentPage
            Response.Redirect(Session("CurrentPage"))
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

    'set the template name
    Private Function setTemplateName() As Boolean
        sqlCommonAdapter.SelectCommand = sqlCommonAction
        sqlCommonAction.Connection = Me.commonConn

        Try
            sqlCommonAction.CommandText = "Select strTemplateName from " & myUtility.getExtension() & "fempTemplates where " & _
            "seqTemplate = " & Session("seqTemplate")
            sqlCommonAction.Connection = Me.commonConn
            sqlCommonAdapter.Fill(Me.dsCommon, "TemplateName")
            Session("TemplateName") = Me.dsCommon.Tables("TemplateName").Rows(0).Item("strTemplateName")
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'set the template name
    Private Function setSurveyName() As Boolean
        sqlCommonAdapter.SelectCommand = sqlCommonAction
        sqlCommonAction.Connection = Me.commonConn

        Try
            sqlCommonAction.CommandText = "Select strComments from " & myUtility.getExtension() & "fempSurvey where " & _
            "seqSurvey = " & Session("seqSurvey")
            sqlCommonAction.Connection = Me.commonConn
            sqlCommonAdapter.Fill(Me.dsCommon, "SurveyName")
            Session("SurveyName") = Me.dsCommon.Tables("SurveyName").Rows(0).Item("strComments")
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function
#End Region

#Region "Data Load"
    'Loads data into the form
    Private Function loadData(Optional ByVal override As Boolean = False) As Boolean
        Trace.Warn("Loading Data")
        Try
            'Fill the Groups List
            If (loadGroups()) Then
                'Fill the Group Users List
                If (loadGroupUsers()) Then
                    'Fill the Users List
                    If (loadGroupInfo()) Then
                        Return True
                    Else
                        Return False
                    End If
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    'loads the groups
    Private Function loadGroups() As Boolean
        Trace.Warn("Loading Groups")
        Try
            'try to erase anything if it exists in the table so we don't end up with duplicate data
            If (Me.dsCommon.Tables.Contains("Groups")) Then
                Me.dsCommon.Tables("Groups").Rows.Clear()
            End If

            'select only the groups for the current survey
            Me.SqlSelectCommand1.CommandText = "SELECT seqSurvey, intGroupID, strGroupName, strUserList FROM " & _
            myUtility.getProdExtension() & "fempGroups Where seqSurvey = " & Session("seqSurvey")
            Me.SqlSelectCommand1.Connection = Me.commonConn
            Me.SqlGroups.Fill(Me.dsCommon, "Groups")
            Session("GroupsMax") = Me.dsCommon.Tables("Groups").Rows.Count()

            'determine if there is any data and if the group id has been set
            If (Session("GroupsMax") = 0) Then
                Session("intGroupID") = -1
            ElseIf (Session("intGroupID") Is Nothing) Then
                Session("intGroupID") = 1
            End If

            'make sure the group row never exceeds the maximum groups
            If (Session("GroupsRow") > Session("GroupsMax")) Then
                Session("GroupsRow") = Session("GroupsMax") - 1
            End If

            'if the group id indicates a new record then move the group row to the maximum
            'else set it to just below the group id's number
            If (Session("intGroupID") = -1) Then
                Session("GroupsRow") = Session("GroupsMax")
            Else
                'Session("GroupsRow") = findRow("Groups", Session("intGroupID"))
                Session("GroupsRow") = Session("intGroupID") - 1
            End If

            'set the group name
            If (Session("intGroupID") <> -1) Then
                Session("GroupName") = Me.dsCommon.Tables("Groups").Rows(Session("GroupsRow")).Item("strGroupName")
            End If

            'set the current mailing list from the current group if there is one selected
            If (Session("intGroupID") <> -1) Then
                Me.mailList = Me.dsCommon.Tables("Groups").Rows(Session("GroupsRow")).Item("strUserList")
            Else
                Me.mailList = ""
            End If

            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'loads the group users
    Private Function loadGroupUsers() As Boolean
        Trace.Warn("Loading Group Users")
        Try
            'try to erase anything if it exists in the table so we don't end up with duplicate data
            If (Me.dsCommon.Tables.Contains("GroupUsers")) Then
                Me.dsCommon.Tables("GroupUsers").Rows.Clear()
            End If

            If (Session("intGroupID") > 0) Then
                'setup the ad-hoc sql tools
                sqlCommonAction.Connection = Me.commonConn
                sqlCommonAdapter.SelectCommand = sqlCommonAction

                'get the user information necessary for the group user listing
                If (Me.dsCommon.Tables("Groups").Rows(Session("intGroupID") - 1).Item("strUserList") <> "") Then
                    sqlCommonAction.CommandText = "SELECT seqUserID, intUserType, strLastName + ', ' + strFirstName + ' ' + st" & _
                    "rMiddleName + ' ' + strNameSuffix AS Name, strHID, strEmail FROM " & myUtility.getProdExtension() & "fempUsers WHERE seqUserID in (" & _
                    Me.dsCommon.Tables("Groups").Rows(Session("intGroupID") - 1).Item("strUserList") & ") ORDER BY Name"

                    sqlCommonAdapter.Fill(Me.dsCommon, "GroupUsers")
                    Me.dgGroup.DataSource = Me.dsCommon.Tables("GroupUsers")
                    Me.dgGroup.DataBind()
                    myUtility.checkBoxes(Me.dgGroup)
                    If (Me.dsCommon.Tables("GroupUsers").Rows.Count() = 0) Then
                        Session("GroupUserListRows") = "None"
                    Else
                        Session("GroupUserListRows") = "Some"
                    End If
                Else
                    Session("GroupUserListRows") = "None"
                End If
            Else
                'setup the ad-hoc sql tools
                sqlCommonAction.Connection = Me.commonConn
                sqlCommonAdapter.SelectCommand = sqlCommonAction

                'get the user information necessary for the group user listing
                sqlCommonAction.CommandText = "SELECT seqUserID, intUserType, strLastName + ', ' + strFirstName + ' ' + st" & _
                "rMiddleName + ' ' + strNameSuffix AS Name, strHID, strEmail FROM " & myUtility.getProdExtension() & "fempUsers WHERE seqUserID in ('-1') " & _
                " ORDER BY Name"
                sqlCommonAdapter.Fill(Me.dsCommon, "GroupUsers")
                Me.dgGroup.DataSource = Me.dsCommon.Tables("GroupUsers")
                Me.dgGroup.DataBind()
                If (Me.dsCommon.Tables("GroupUsers").Rows.Count() = 0) Then
                    Session("GroupUserListRows") = "None"
                Else
                    Session("GroupUserListRows") = "Some"
                End If
            End If
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Session("GroupUserListRows") = "None"
            Return False
        End Try
    End Function

    'loads the users
    Private Function loadGroupInfo() As Boolean
        Trace.Warn("Loading Group Information")
        Try
            'try to erase anything if it exists in the table so we don't end up with duplicate data
            If (Me.dsCommon.Tables.Contains("Users")) Then
                Me.dsCommon.Tables("Users").Rows.Clear()
            End If

            'get the current e-mail string
            If (Session("intGroupID") > 0) Then
                If (Me.dsCommon.Tables("Groups").Rows(Session("intGroupID") - 1).Item("strUserList") <> "") Then
                    Me.SqlSelectCommand2.CommandText = "SELECT seqUserID, intUserType, strLastName + ', ' + strFirstName + ' ' + st" & _
                    "rMiddleName + ' ' + strNameSuffix AS Name, strHID, strEmail FROM " & myUtility.getProdExtension() & "fempUsers WHERE seqUserID not in (" & _
                    Me.dsCommon.Tables("Groups").Rows(Session("intGroupID") - 1).Item("strUserList") & ") and upper(strLastName) like '" & _
                    Session("Alpha") & "%' and strEmail <> '' ORDER BY Name"
                Else
                    Me.SqlSelectCommand2.CommandText = "SELECT seqUserID, intUserType, strLastName + ', ' + strFirstName + ' ' + st" & _
                    "rMiddleName + ' ' + strNameSuffix AS Name, strHID, strEmail FROM " & myUtility.getProdExtension() & "fempUsers WHERE upper(strLastName) like '" & _
                    Session("Alpha") & "%' and strEmail <> '' ORDER BY Name"
                End If
            Else
                Me.SqlSelectCommand2.CommandText = "SELECT seqUserID, intUserType, strLastName + ', ' + strFirstName + ' ' + st" & _
                "rMiddleName + ' ' + strNameSuffix AS Name, strHID, strEmail FROM " & myUtility.getProdExtension() & "fempUsers WHERE upper(strLastName) like '" & _
                Session("Alpha") & "%' and strEmail <> '' ORDER BY Name"
            End If

            Me.SqlSelectCommand2.Connection = Me.commonConn
            Me.SqlUsers.Fill(Me.dsCommon, "Users")
            Me.dgUsers.DataSource = Me.dsCommon.Tables("Users")
            Me.dgUsers.DataBind()
            If (Me.dsCommon.Tables("Users").Rows.Count() = 0) Then
                Session("GeneralUserListRows") = "None"
            Else
                Session("GeneralUserListRows") = "Some"
            End If
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Session("GeneralUserListRows") = "None"
            Return False
        End Try
    End Function
#End Region

#Region "Control Settings"
    'populates the controls with the selected group information
    Private Sub setControls()
        Trace.Warn("Setting Controls")
        'set the basic page controls based on whether or not we are working with a new group
        If (Session("intGroupID") > 0) Then
            Try
                Me.txtGroupName.Text = Me.dsCommon.Tables("Groups").Rows(Session("GroupsRow")).Item("strGroupName")
                Me.txtGroupList.Text = ""
                Me.lblNavBar.Text = "Group " & Session("GroupsRow") + 1 & " of " & Session("GroupsMax")
            Catch ex As Exception
                Trace.Warn(ex.ToString)
            End Try
        Else
            Me.txtGroupName.Text = ""
            Me.txtGroupList.Text = ""
            Me.lblNavBar.Text = "Group " & Session("GroupsRow") + 1 & " of " & Session("GroupsMax") + 1
        End If

        'disable inputs on production.
        If (Session("Machine") = "Production") Then
            Me.txtGroupName.Enabled = False
            Me.txtGroupList.Enabled = False
        Else
            Me.txtGroupList.Enabled = True
            Me.txtGroupName.Enabled = True
        End If

        'set the page options
        myUtility.optionPreSet(Session("intGroupID"), Session("GroupsMax"), Me.pageOptions)

        Session("PageOptions") = Me.pageOptions
    End Sub
#End Region

#Region "Page Switch"
    'Takes the user back to the template they were working on
    Private Sub btnUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUp.Click
        Session("JSAction") = ""
        Session("CurrentPage") = "./Survey.aspx"
        Session("pageSel") = "Login"
        Response.Redirect(Session("CurrentPage"))
    End Sub

    'brings up a pop-up window with the template preview in it
    Private Sub btnTemplatePreview_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTemplatePreview.Click
        Session("JSAction") = "<script>newWin = window.open('./Preview.aspx', 'PreviewWindow', 'width=600,height=400,scrollbars=yes,toolbar=no,titlebar=no,status=no,resizable=yes,menubar=no');newWin.focus();</script>"
    End Sub
#End Region

#Region "Find"
    'finds the row of the survey we're looking for
    Private Function findRow(ByVal strTable As String, ByVal userID As Integer) As Integer
        Dim intRow As Integer = 0
        Dim dr As DataRow
        For Each dr In Me.dsCommon.Tables(strTable).Rows()
            If (dr.Item("intGroupID") = userID) Then
                Return intRow
            Else
                intRow += 1
            End If
        Next
        Return -1
    End Function
#End Region

#Region "Grid"
    'Weigh station for processing the checking and unchecking of checkboxes in the grids 
    Public Sub commandBtnClick(ByVal sender As Object, ByVal e As CommandEventArgs)
        Trace.Warn("Header Button Click")
        Select Case e.CommandName
            Case "SelectGroup"
                myUtility.checkBoxes(Me.dgGroup)
            Case "SelectUsers"
                myUtility.checkBoxes(Me.dgUsers)
        End Select
    End Sub

    'resets the user grid to the value passed in from the button
    Public Sub commandGridUserList(ByVal sender As Object, ByVal e As CommandEventArgs)
        Trace.Warn("Grid Alphabet Button Click")
        Session("Alpha") = e.CommandName

        Dim failure As Boolean = False

        If (Me.txtGroupName.Text <> "") Then
            If (Session("intGroupID") <> -1) Then
                failure = doInsertUpdate("Update")
            Else
                failure = doInsertUpdate("Insert")
            End If

            'load the page data
            If Not (loadData(True)) Then
                alert("Failed to load the group information for this survey.")
            End If
        Else
            'load the page data
            If Not (loadData()) Then
                alert("Failed to load the group information for this survey.")
            End If
        End If

        'reset the page controls
        setControls()
        'set the page options
        myUtility.optionPreSet(Session("intGroupID"), Session("GroupsMax"), Me.pageOptions)

        Session("PageOptions") = Me.pageOptions
    End Sub
#End Region

#Region "Nav Bar Events"
    'Moves the user to the first record
    Private Sub wnbGroups_MoveToStart(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnStart.Click
        Session("JSAction") = ""

        If Not (loadGroups()) Then
            alert("Failed to load the groups for this survey.")
        Else
            Try
                Session("GroupsRow") = 0
                Session("intGroupID") = Me.dsCommon.Tables("Groups").Rows(Session("GroupsRow")).Item("intGroupID")
                Session("GroupName") = Me.dsCommon.Tables("Groups").Rows(Session("GroupsRow")).Item("strGroupName")
                loadGroupUsers()
                loadGroupInfo()

                'set the current mailing list from the current group if there is one selected
                If (Session("intGroupID") <> -1) Then
                    Me.mailList = Me.dsCommon.Tables("Groups").Rows(Session("GroupsRow")).Item("strUserList")
                Else
                    Me.mailList = ""
                End If

                myNavBar.wnb_MoveTo(Me.navButtons, Session("GroupsMax"), Session("GroupsRow"), Switch)
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("intGroupID"), Session("GroupsMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the first group.")
            End Try
        End If
    End Sub

    'Moves the user to the previous record
    Private Sub wnbGroups_MovePrev(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnPrev.Click
        Session("JSAction") = ""

        If Not (loadGroups()) Then
            alert("Failed to load the groups for this survey.")
        Else
            Try
                Session("GroupsRow") -= 1
                Session("intGroupID") = Me.dsCommon.Tables("Groups").Rows(Session("GroupsRow")).Item("intGroupID")
                Session("GroupName") = Me.dsCommon.Tables("Groups").Rows(Session("GroupsRow")).Item("strGroupName")
                loadGroupUsers()
                loadGroupInfo()

                'set the current mailing list from the current group if there is one selected
                If (Session("intGroupID") <> -1) Then
                    Me.mailList = Me.dsCommon.Tables("Groups").Rows(Session("GroupsRow")).Item("strUserList")
                Else
                    Me.mailList = ""
                End If

                myNavBar.wnb_MoveTo(Me.navButtons, Session("GroupsMax"), Session("GroupsRow"), Switch)
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("intGroupID"), Session("GroupsMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the previous group.")
            End Try
        End If
    End Sub

    'reloads the page and the data from the database while attempting to locate the user's location
    Private Sub wnbGroups_MoveReload(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnReload.Click
        Session("JSAction") = ""

        If Not (loadData()) Then
            alert("Failed to load the groups for this survey.")
        Else
            Try
                myNavBar.wnb_MoveTo(Me.navButtons, Session("GroupsMax"), Session("GroupsRow"), Switch)
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("intGroupID"), Session("GroupsMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error resetting the current group.")
            End Try
        End If
    End Sub

    'Moves the user to the next record
    Private Sub wnbGroups_MoveNext(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnNext.Click
        Session("JSAction") = ""

        If Not (loadGroups()) Then
            alert("Failed to load the groups for this survey.")
        Else
            Try
                Session("GroupsRow") += 1
                Session("intGroupID") = Me.dsCommon.Tables("Groups").Rows(Session("GroupsRow")).Item("intGroupID")
                Session("GroupName") = Me.dsCommon.Tables("Groups").Rows(Session("GroupsRow")).Item("strGroupName")
                loadGroupUsers()
                loadGroupInfo()

                'set the current mailing list from the current group if there is one selected
                If (Session("intGroupID") <> -1) Then
                    Me.mailList = Me.dsCommon.Tables("Groups").Rows(Session("GroupsRow")).Item("strUserList")
                Else
                    Me.mailList = ""
                End If

                myNavBar.wnb_MoveTo(Me.navButtons, Session("GroupsMax"), Session("GroupsRow"), Switch)
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("intGroupID"), Session("GroupsMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the next group.")
            End Try
        End If
    End Sub

    'Moves the user to the last record
    Private Sub wnbGroups_MoveToEnd(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnLast.Click
        Session("JSAction") = ""

        If Not (loadGroups()) Then
            alert("Failed to load the groups for this survey.")
        Else
            Try
                Session("GroupsRow") = Session("GroupsMax") - 1
                Session("intGroupID") = Me.dsCommon.Tables("Groups").Rows(Session("GroupsRow")).Item("intGroupID")
                Session("GroupName") = Me.dsCommon.Tables("Groups").Rows(Session("GroupsRow")).Item("strGroupName")
                loadGroupUsers()
                loadGroupInfo()

                'set the current mailing list from the current group if there is one selected
                If (Session("intGroupID") <> -1) Then
                    Me.mailList = Me.dsCommon.Tables("Groups").Rows(Session("GroupsRow")).Item("strUserList")
                Else
                    Me.mailList = ""
                End If

                myNavBar.wnb_MoveTo(Me.navButtons, Session("GroupsMax"), Session("GroupsRow"), Switch)
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("intGroupID"), Session("GroupsMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the last group.")
            End Try
        End If
    End Sub

    'Inserts a record into the table
    Private Sub wnbQuestions_InsertRow(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnNew.Click
        Session("JSAction") = ""

        If Not (loadGroups()) Then
            alert("Failed to load the groups for this survey.")
        Else
            Try
                Session("intGroupID") = -1
                Session("GroupsRow") = Session("GroupsMax")
                Me.dsCommon.Tables("Groups").Rows.Add(Me.dsCommon.Tables("Groups").NewRow())
                Me.dsCommon.Tables("Groups").Rows(Session("GroupsMax")).Item("seqSurvey") = Session("seqSurvey")
                Me.dsCommon.Tables("Groups").Rows(Session("GroupsMax")).Item("intGroupID") = Session("GroupsMax")
                Me.dsCommon.Tables("Groups").Rows(Session("GroupsMax")).Item("strGroupName") = ""
                Me.dsCommon.Tables("Groups").Rows(Session("GroupsMax")).Item("strUserList") = ""
                loadGroupUsers()
                loadGroupInfo()

                'set the current mailing list from the current group if there is one selected
                If (Session("intGroupID") <> -1) Then
                    Me.mailList = Me.dsCommon.Tables("Groups").Rows(Session("GroupsRow")).Item("strUserList")
                Else
                    Me.mailList = ""
                End If

                myNavBar.wnb_MoveTo(Me.navButtons, Session("GroupsMax"), Session("GroupsRow"))
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("intGroupID"), Session("GroupsMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error creating new group.")
            End Try
        End If
    End Sub
#End Region

#Region "Page Action"
    'Reacts to the user performing some kind of selection and submission
    Private Sub btnInsert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsert1.Click, btnInsert2.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'refresh the Groups table
        If Not (loadGroups()) Then
            blnContinue = False
            alert("Failed to load the groups for the current survey.  Insert aborted.")
        End If

        'check for possible malicious code entries
        If Not (myUtility.goodInput(Me.txtGroupList, True) And myUtility.goodInput(Me.txtGroupName, True)) Then
            blnContinue = False
            alert("Possible malicious code entry(s).  Insert aborted.")
        End If

        If (blnContinue) Then
            'do the insert
            failure = doInsertUpdate("Insert")

            'reload the data
            If Not (loadData(True)) Then
                blnContinue = False
                alert("Failed to load the groups for the current survey.")
            End If

            If (blnContinue) Then
                'to determine what of the nav bar buttons should be available
                myNavBar.wnb_MoveTo(Me.navButtons, Session("GroupsMax"), Session("GroupsRow"), Switch)

                'reset the page controls
                If Not (failure) Then
                    setControls()
                End If
                Session("TransExists") = False
                'set the page options
                myUtility.optionPreSet(Session("intGroupID"), Session("GroupsMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection and submission
    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate1.Click, btnUpdate2.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'refresh the Groups table
        If Not (loadGroups()) Then
            blnContinue = False
            alert("Failed to load the groups for the current survey.  Update aborted.")
        End If

        'check for possible malicious code entries
        If Not (myUtility.goodInput(Me.txtGroupList, True) And myUtility.goodInput(Me.txtGroupName, True)) Then
            blnContinue = False
            alert("Possible malicious code entry(s).  Update aborted.")
        End If

        If (blnContinue) Then
            'do the update
            failure = doInsertUpdate("Update")

            'reload the data
            If Not (loadData(True)) Then
                blnContinue = False
                alert("Failed to load the groups for the current survey.")
            End If

            If (blnContinue) Then
                'to determine what of the nav bar buttons should be available
                myNavBar.wnb_MoveTo(Me.navButtons, Session("GroupsMax"), Session("GroupsRow"), Switch)

                'reset the page controls
                If Not (failure) Then
                    setControls()
                End If
                Session("TransExists") = False
                'set the page options
                myUtility.optionPreSet(Session("intGroupID"), Session("GroupsMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection and submission
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete1.Click, btnDelete2.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'refresh the Groups table
        If Not (loadGroups()) Then
            blnContinue = False
            alert("Failed to load the groups for the current survey.  Delete aborted.")
        End If

        'check for possible malicious code entries
        If Not (myUtility.goodInput(Me.txtGroupList, True) And myUtility.goodInput(Me.txtGroupName, True)) Then
            blnContinue = False
            alert("Possible malicious code entry(s).  Delete aborted.")
        End If

        If (blnContinue) Then
            'do the delete
            failure = doDelete()

            'reload the data
            If Not (loadData(True)) Then
                blnContinue = False
                alert("Failed to load the groups for the current survey.")
            End If

            If (blnContinue) Then
                'to determine what of the nav bar buttons should be available
                myNavBar.wnb_MoveTo(Me.navButtons, Session("GroupsMax"), Session("GroupsRow"), Switch)

                'reset the page controls
                If Not (failure) Then
                    setControls()
                End If
                Session("TransExists") = False
                'set the page options
                myUtility.optionPreSet(Session("intGroupID"), Session("GroupsMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection and submission
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset1.Click, btnReset2.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'check for possible malicious code entries
        If Not (myUtility.goodInput(Me.txtGroupList, True) And myUtility.goodInput(Me.txtGroupName, True)) Then
            blnContinue = False
            alert("Possible malicious code entry(s).  Reset aborted.")
        End If

        If (blnContinue) Then
            'do the reset
            failure = doReset()

            If (blnContinue) Then
                'to determine what of the nav bar buttons should be available
                myNavBar.wnb_MoveTo(Me.navButtons, Session("GroupsMax"), Session("GroupsRow"), Switch)

                'reset the page controls
                If Not (failure) Then
                    setControls()
                End If
                Session("TransExists") = False
                'set the page options
                myUtility.optionPreSet(Session("intGroupID"), Session("GroupsMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub
#End Region

#Region "Insert/Update Code"
    'drives the commit and roll-back operations of the insert
    Private Function doInsertUpdate(ByVal strAction As String) As Boolean
        Try
            'set up the transaction
            If (myUtility.setupTransaction(sqlCommonAction, Me.commonConn)) Then
                sqlCommonAdapter.SelectCommand = sqlCommonAction

                Dim failure As Boolean = False
                Dim oldMailList As String = Me.mailList
                Me.mailList = ""
                'Check for actual information to insert
                If (Me.txtGroupName.Text <> "") Then
                    'process the text box for HID's, Emails, and Payrolls
                    If (processManualEntry(sqlCommonAction, sqlCommonAdapter)) Then
                        'process the current members of the group listing for deletions
                        If (processGroupList(sqlCommonAction, sqlCommonAdapter)) Then
                            'process the user listing for selections
                            If (processUserList(sqlCommonAction, sqlCommonAdapter)) Then
                                If (commitGroup(sqlCommonAction, sqlCommonAdapter, strAction)) Then
                                    If (commitPermissions(sqlCommonAction, sqlCommonAdapter, oldMailList)) Then
                                        sqlCommonAction.Transaction.Commit()
                                        'load the page data
                                        If Not (loadData()) Then
                                            alert("Failed to load the group information for this survey.")
                                        Else
                                            'to determine what of the nav bar buttons should be available
                                            myNavBar.wnb_MoveTo(Me.navButtons, Session("GroupsMax"), Session("GroupsRow"), Switch)
                                        End If

                                        'Populate the controls on the page
                                        setControls()

                                        'set the page options
                                        myUtility.optionPreSet(Session("intGroupID"), Session("GroupsMax"), Me.pageOptions)

                                        Session("PageOptions") = Me.pageOptions
                                    Else
                                        sqlCommonAction.Transaction.Rollback()
                                        mailList = oldMailList
                                        alert("Failed to process the user list.")
                                        Return False
                                    End If
                                Else
                                    sqlCommonAction.Transaction.Rollback()
                                    mailList = oldMailList
                                    alert("Failed to process the user list.")
                                    Return False
                                End If
                            Else
                                sqlCommonAction.Transaction.Rollback()
                                mailList = oldMailList
                                alert("Failed to process the user list.")
                                Return False
                            End If
                        Else
                            sqlCommonAction.Transaction.Rollback()
                            mailList = oldMailList
                            alert("Failed to process the group list.")
                            Return False
                        End If
                    Else
                        sqlCommonAction.Transaction.Rollback()
                        mailList = oldMailList
                        alert("Failed to process your manual entry.  Check the format of your entry and try again.  Remember to separate your entries with commas and/or by placing them on their own row.")
                        Return False
                    End If
                Else
                    sqlCommonAction.Transaction.Rollback()
                    mailList = oldMailList
                    alert("You must provide a name for this group at a minimum.")
                    Return False
                End If
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Group " & strAction & " failure.")
                Return True
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Group " & strAction & " failure.")
            Return True
        End Try
    End Function

    'attempts to process the manual entry starting the e-mail string build
    Private Function processManualEntry(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            'create an array of users while trimming the white space
            Dim requestItems As Array
            Dim index As Integer = 0
            If (Me.txtGroupList.Text.Length > 0) Then
                Dim strLine As String = ""
                'replace white space text
                Me.txtGroupList.Text = Me.txtGroupList.Text.Replace(vbCrLf, ",")
                requestItems = Me.txtGroupList.Text.Split(",")
                While index < requestItems.Length
                    requestItems(index) = Trim(CType(requestItems(index), String))
                    index += 1
                End While
            End If

            'determine if we actually got anything.  If not then skip this function
            If (requestItems Is Nothing) Then
                Return True
            End If

            'determine if the person exists as a current user.
            index = 0
            Me.txtGroupList.Text = ""
            While index < requestItems.Length
                If (CType(requestItems(index), String) <> "") Then
                    If Not (getHanfordPeopleInfo(CType(requestItems(index), String), sqlCommonAction, sqlCommonAdapter)) Then
                        If (Session("Alert") <> "") Then
                            Session("Alert") &= ", " & CType(requestItems(index), String)
                        Else
                            Session("Alert") = "Could not find the following people: " & CType(requestItems(index), String)
                        End If
                    End If
                    If (Me.dsCommon.Tables.Contains("ExistingUsers")) Then
                        Me.dsCommon.Tables("ExistingUsers").Rows.Clear()
                    End If
                    If (Me.dsCommon.Tables.Contains("PersonData")) Then
                        Me.dsCommon.Tables("PersonData").Rows.Clear()
                    End If
                    If (Me.dsCommon.Tables.Contains("UserIDFetch")) Then
                        Me.dsCommon.Tables("UserIDFetch").Rows.Clear()
                    End If
                End If
                index += 1
            End While
            If (Session("Alert") <> "") Then
                Session("Alert") &= "\nPlease fill in the pop-up windows for the e-mail addresses not found."
                alert(Session("Alert"))
                Session("Alert") = ""
            End If
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'gets the hanford people information for the user we're looking for, returns true if it finds the person(s)
    Private Function getHanfordPeopleInfo(ByVal LookUpString As String, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            'initialize the adhoc data collection items
            If ((LookUpString.ToUpper.StartsWith("H") And IsNumeric(LookUpString.Substring(1))) Or (LookUpString.Length = 7 And IsNumeric(LookUpString))) Then
                'Process the apparent HID
                Return (processHID(LookUpString, sqlCommonAction, sqlCommonAdapter))
            ElseIf ((LookUpString.ToUpper.StartsWith("D") And LookUpString.Length = 6) Or LookUpString.Length = 5) Then
                If (IsNumeric(LookUpString.Substring(LookUpString.Length - 3))) Then
                    'Process the apparent payroll
                    Return (processPayroll(LookUpString, sqlCommonAction, sqlCommonAdapter))
                Else
                    'Process the apparent last name
                    Return (processEmail(LookUpString, sqlCommonAction, sqlCommonAdapter))
                End If
            Else
                'Process the apparent email
                Return (processEmail(LookUpString, sqlCommonAction, sqlCommonAdapter))
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'Process the HID
    Private Function processHID(ByVal LookUpString As String, ByVal sqlCommonAction As SqlClient.SqlCommand, ByVal sqlCommonAdapter As SqlClient.SqlDataAdapter) As Boolean
        Try
            'strip the H from the front end if it exists
            If (LookUpString.ToUpper.StartsWith("H")) Then
                LookUpString = LookUpString.Substring(1)
            End If

            Dim alreadyExists As Boolean = False
            'first check the current list of users and determine if that user already exists before polling the opwhse
            sqlCommonAction.CommandText = "select fu.strHID as hanford_id, fu.seqUserID, hp.hanf_pay_no, fu.strFirstName as first_name, " & _
            "fu.strLastName as last_name, fu.strMiddleName as middle_initial, fu.strNameSuffix as name_suffix, " & _
            "fu.strWorkPhone as phone_no, fu.strEmail as internet_email_address, fu.strCompany as employed_by_hcid, " & _
            "fu.intUserType, fu.blnTraining, fu.blnActive, hp.active_sw from " & myUtility.getProdExtension() & "fempUsers fu, " & _
            "opwhse.dbo.vw_pub_hanford_people hp where fu.strHID = hp.hanford_id " & _
            "and fu.strHID = '" & LookUpString & "'"
            sqlCommonAdapter.Fill(Me.dsCommon, "ExistingUsers")

            'if we found something then they already exist
            If (Me.dsCommon.Tables("ExistingUsers").Rows.Count >= 1) Then
                Session("PersonTable") = Me.dsCommon.Tables("ExistingUsers")
                alreadyExists = True
            End If

            If (alreadyExists) Then
                'if we found something then add their user id to the list of users.
                addToLists(Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("seqUserID"))
                Return True
            Else
                'set up and process the sql statement
                sqlCommonAction.CommandText = "select isnull(max(hanford_id), '') as hanford_id, isnull(max(hanf_pay_no), '') as " & _
                "hanf_pay_no, max(first_name) as first_name, max(last_name) as last_name, isnull(max(middle_initial), '') as " & _
                "middle_initial, isnull(max(name_prefix), '') as name_prefix, isnull(max(name_suffix), '') as name_suffix, " & _
                "isnull(max(phone_no), '') as phone_no, isnull(max(internet_email_address), '') as internet_email_address, " & _
                "isnull(max(employed_by_hcid), '') as employed_by_hcid, isNull(max(active_sw), 'N') as active_sw from  (" & _
                "select hanford_id, hanf_pay_no, first_name, last_name, middle_initial, NULL as name_prefix, name_suffix, phone_no, " & _
                "internet_email_address, employed_by_hcid, active_sw from opwhse.dbo.vw_pub_hanford_people union select hanford_id, " & _
                "pay_no as hanf_pay_no, first_name, last_name, middle_name as middle_initial, NULL as name_prefix, name_suffix, " & _
                "work_phone as phone_no, internet_address as internet_email_address, employed_by_hcid, NULL as active_sw " & _
                "from opwhse.dbo.vw_pub_pnnl_associate union select hanford_id, pnnl_pay_no as hanf_pay_no, first_name, last_name, " & _
                "middle_initial, name_prefix, name_suffix, contact_work_phone as phone_no, internet_email_address, company as " & _
                "employed_by_hcid, NULL as active_sw from opwhse.dbo.vw_pub_bmi_employee) a where " & _
                "hanford_id = '" & LookUpString & "'" & _
                " group by hanford_id, last_name, first_name, middle_initial, internet_email_address"
                sqlCommonAdapter.Fill(Me.dsCommon, "PersonData")

                'if we found something return true otherwise return false
                If (Me.dsCommon.Tables("PersonData").Rows.Count >= 1) Then
                    Session("PersonTable") = Me.dsCommon.Tables("PersonData")
                    'create the new level 0 user
                    If (createUser(sqlCommonAction, sqlCommonAdapter)) Then
                        addToLists(Session("mailListID"))
                        Return True
                    Else
                        Return False
                    End If
                Else
                    Session("PersonTable") = Nothing
                    Return False
                End If
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'Process the Payroll
    Private Function processPayroll(ByVal LookUpString As String, ByVal sqlCommonAction As SqlClient.SqlCommand, ByVal sqlCommonAdapter As SqlClient.SqlDataAdapter) As Boolean
        Try
            'strip the D from the front end if it exists
            If (LookUpString.ToUpper.StartsWith("D")) Then
                LookUpString = LookUpString.Substring(1)
            End If

            Dim alreadyExists As Boolean = False
            'first check the current list of users and determine if that user already exists before polling the opwhse
            sqlCommonAction.CommandText = "select fu.strHID as hanford_id, fu.seqUserID, hp.hanf_pay_no, fu.strFirstName as first_name, " & _
            "fu.strLastName as last_name, fu.strMiddleName as middle_initial, fu.strNameSuffix as name_suffix, " & _
            "fu.strWorkPhone as phone_no, fu.strEmail as internet_email_address, fu.strCompany as employed_by_hcid, " & _
            "fu.intUserType, fu.blnTraining, fu.blnActive, hp.active_sw from " & myUtility.getProdExtension() & "fempUsers fu, " & _
            "opwhse.dbo.vw_pub_hanford_people hp where fu.strHID = hp.hanford_id " & _
            "and hp.hanf_pay_no = '" & LookUpString & "'"
            sqlCommonAdapter.Fill(Me.dsCommon, "ExistingUsers")

            'if we found something return true otherwise return false
            If (Me.dsCommon.Tables("ExistingUsers").Rows.Count >= 1) Then
                Session("PersonTable") = Me.dsCommon.Tables("ExistingUsers")
                alreadyExists = True
            End If

            If (alreadyExists) Then
                'if we found something then add their user id to the list of users.
                addToLists(Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("seqUserID"))
                Return True
            Else
                'set up and process the sql statement
                sqlCommonAction.CommandText = "select isnull(max(hanford_id), '') as hanford_id, isnull(max(hanf_pay_no), '') as " & _
                "hanf_pay_no, max(first_name) as first_name, max(last_name) as last_name, isnull(max(middle_initial), '') as " & _
                "middle_initial, isnull(max(name_prefix), '') as name_prefix, isnull(max(name_suffix), '') as name_suffix, " & _
                "isnull(max(phone_no), '') as phone_no, isnull(max(internet_email_address), '') as internet_email_address, " & _
                "isnull(max(employed_by_hcid), '') as employed_by_hcid, isNull(max(active_sw), 'N') as active_sw from  (" & _
                "select hanford_id, hanf_pay_no, first_name, last_name, middle_initial, NULL as name_prefix, name_suffix, phone_no, " & _
                "internet_email_address, employed_by_hcid, active_sw from opwhse.dbo.vw_pub_hanford_people union select hanford_id, " & _
                "pay_no as hanf_pay_no, first_name, last_name, middle_name as middle_initial, NULL as name_prefix, name_suffix, " & _
                "work_phone as phone_no, internet_address as internet_email_address, employed_by_hcid, NULL as active_sw " & _
                "from opwhse.dbo.vw_pub_pnnl_associate union select hanford_id, pnnl_pay_no as hanf_pay_no, first_name, last_name, " & _
                "middle_initial, name_prefix, name_suffix, contact_work_phone as phone_no, internet_email_address, company as " & _
                "employed_by_hcid, NULL as active_sw from opwhse.dbo.vw_pub_bmi_employee) a where " & _
                "hanf_pay_no = '" & LookUpString & "'" & _
                " group by hanford_id, last_name, first_name, middle_initial, internet_email_address"
                sqlCommonAdapter.Fill(Me.dsCommon, "PersonData")

                'if we found something return true otherwise return false
                If (Me.dsCommon.Tables("PersonData").Rows.Count >= 1) Then
                    Session("PersonTable") = Me.dsCommon.Tables("PersonData")
                    'create the new level 0 user
                    If (createUser(sqlCommonAction, sqlCommonAdapter)) Then
                        addToLists(Session("mailListID"))
                        Return True
                    Else
                        Return False
                    End If
                Else
                    Session("PersonTable") = Nothing
                    Return False
                End If
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'Process the e-mail address
    Private Function processEmail(ByVal LookUpString As String, ByVal sqlCommonAction As SqlClient.SqlCommand, ByVal sqlCommonAdapter As SqlClient.SqlDataAdapter) As Boolean
        Try
            'destroy any data that may exist before using the table
            If (Me.dsCommon.Tables.Contains("ExisitingUsers")) Then
                Me.dsCommon.Tables("ExistingUsers").Rows.Clear()
            End If
            If (Me.dsCommon.Tables.Contains("PersonData")) Then
                Me.dsCommon.Tables("PersonData").Rows.Clear()
            End If

            'determine if the e-mail address looks valid
            If Not (isValidEmail(LookUpString)) Then
                Return False
            End If

            Dim alreadyExists As Boolean = False
            'first check the current list of users and determine if that user already exists before polling the opwhse
            sqlCommonAction.CommandText = "select fu.strHID as hanford_id, fu.seqUserID, fu.strFirstName as first_name, " & _
            "fu.strLastName as last_name, fu.strMiddleName as middle_initial, fu.strNameSuffix as name_suffix, " & _
            "fu.strWorkPhone as phone_no, fu.strEmail as internet_email_address, fu.strCompany as employed_by_hcid, " & _
            "fu.intUserType, fu.blnTraining, fu.blnActive from " & myUtility.getProdExtension() & "fempUsers fu " & _
            " where fu.strEmail = '" & LookUpString & "'"
            sqlCommonAdapter.Fill(Me.dsCommon, "ExistingUsers")

            'if we found something then they already exist
            If (Me.dsCommon.Tables("ExistingUsers").Rows.Count >= 1) Then
                Session("PersonTable") = Me.dsCommon.Tables("ExistingUsers")
                alreadyExists = True
            End If

            If (alreadyExists) Then
                'if we found something then add their user id to the list of users.
                addToLists(Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("seqUserID"))
                Return True
            Else
                'set up and process the sql statement
                sqlCommonAction.CommandText = "select isnull(max(hanford_id), '') as hanford_id, isnull(max(hanf_pay_no), '') as " & _
                "hanf_pay_no, max(first_name) as first_name, max(last_name) as last_name, isnull(max(middle_initial), '') as " & _
                "middle_initial, isnull(max(name_prefix), '') as name_prefix, isnull(max(name_suffix), '') as name_suffix, " & _
                "isnull(max(phone_no), '') as phone_no, isnull(max(internet_email_address), '') as internet_email_address, " & _
                "isnull(max(employed_by_hcid), '') as employed_by_hcid, isNull(max(active_sw), 'N') as active_sw, " & _
                "max(last_name) + ', ' + max(first_name) + ' ' + isNull(max(middle_initial), '') + ' - ' + " & _
                "isNull(max(internet_email_address), '') as NameEmail from  (" & _
                "select hanford_id, hanf_pay_no, first_name, last_name, middle_initial, NULL as name_prefix, name_suffix, phone_no, " & _
                "internet_email_address, employed_by_hcid, active_sw from opwhse.dbo.vw_pub_hanford_people union select hanford_id, " & _
                "pay_no as hanf_pay_no, first_name, last_name, middle_name as middle_initial, NULL as name_prefix, name_suffix, " & _
                "work_phone as phone_no, internet_address as internet_email_address, employed_by_hcid, NULL as active_sw " & _
                "from opwhse.dbo.vw_pub_pnnl_associate union select hanford_id, pnnl_pay_no as hanf_pay_no, first_name, last_name, " & _
                "middle_initial, name_prefix, name_suffix, contact_work_phone as phone_no, internet_email_address, company as " & _
                "employed_by_hcid, NULL as active_sw from opwhse.dbo.vw_pub_bmi_employee) a where " & _
                "internet_email_address = '" & LookUpString & "'" & _
                " group by hanford_id, last_name, first_name, middle_initial, internet_email_address " & _
                "order by last_name, first_name"
                sqlCommonAdapter.Fill(Me.dsCommon, "PersonData")

                'if we found something return true otherwise ask the user to provide a name for the user's e-mail
                If (Me.dsCommon.Tables("PersonData").Rows.Count >= 1) Then
                    Session("PersonTable") = Me.dsCommon.Tables("PersonData")
                    'create the new level 0 user
                    If (createUser(sqlCommonAction, sqlCommonAdapter)) Then
                        addToLists(Session("mailListID"))
                        Return True
                    Else
                        Return False
                    End If
                ElseIf (Session("ProcessEmail") = "True") Then
                    Session("PersonTable") = Nothing
                    If (createUser(sqlCommonAction, sqlCommonAdapter)) Then
                        addToLists(Session("mailListID"))
                        Return True
                    Else
                        Return False
                    End If
                Else
                    Session("PersonTable") = Nothing
                    Session("JSAction") &= "<script>newWin = window.open('./PersonInput.aspx?PersonInputEmail=" & LookUpString & "', '_blank', 'width=600,height=400,scrollbars=yes,toolbar=no,titlebar=no,status=no,resizable=yes,menubar=no');newWin.focus();</script>"
                    Return False
                End If
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'checks to be sure the e-mail address looks valid
    Private Function isValidEmail(ByVal strEmail As String) As Boolean
        Try
            'make sure that the last period is 3 away from the end
            If ((strEmail.Length - strEmail.LastIndexOf(".")) = 4) Then
                'make sure that the distance between the @ and the last period is greater than 2
                If ((strEmail.LastIndexOf(".") - strEmail.LastIndexOf("@")) > 2) Then
                    'make sure there is only one @ in the string
                    Dim ch As Char
                    Dim atCounter As Integer = 0
                    For Each ch In strEmail
                        If (ch = "@") Then
                            atCounter += 1
                        End If
                    Next
                    If (atCounter = 1) Then
                        'make sure there is only one period after the @ in the string
                        Dim strTemp As String = strEmail.Substring(strEmail.LastIndexOf("@") + 1)
                        Dim periodCounter As Integer = 0
                        For Each ch In strTemp
                            If (ch = ".") Then
                                periodCounter += 1
                            End If
                        Next
                        If (periodCounter = 1) Then
                            Return True
                        Else
                            Return False
                        End If
                    Else
                        Return False
                    End If
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'appends a list of user id's and email addresses
    Private Sub addToLists(ByVal userID As String)
        If (mailList.Length = 0) Then
            mailList = userID
        Else
            mailList &= ", " & userID
        End If
    End Sub

    'attempts to insert a user, returns false if it cannot
    Private Function createUser(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, Optional ByVal strEmail As String = "") As Boolean
        Try
            sqlCommonAction.Parameters.Clear()
            If Not (Session("PersonTable") Is Nothing) Then
                'parameterized the text input to allow for things like quotes to get through
                Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@strEmail", System.Data.SqlDbType.VarChar, 128, "strEmail")
                param0.Value = Me.dsCommon.Tables("PersonData").Rows(0).Item("internet_email_address")
                sqlCommonAction.Parameters.Add(param0)
                Dim param1 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@strFirstName", System.Data.SqlDbType.VarChar, 50, "strFirstName")
                param1.Value = Me.dsCommon.Tables("PersonData").Rows(0).Item("first_name")
                sqlCommonAction.Parameters.Add(param1)
                Dim param2 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@strLastName", System.Data.SqlDbType.VarChar, 50, "strLastName")
                param2.Value = Me.dsCommon.Tables("PersonData").Rows(0).Item("last_name")
                sqlCommonAction.Parameters.Add(param2)
                Dim param3 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@strMiddleName", System.Data.SqlDbType.VarChar, 50, "strMiddleName")
                param3.Value = Me.dsCommon.Tables("PersonData").Rows(0).Item("middle_initial")
                sqlCommonAction.Parameters.Add(param3)
                Dim param4 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@strNameSuffix", System.Data.SqlDbType.VarChar, 50, "strNameSuffix")
                param4.Value = Me.dsCommon.Tables("PersonData").Rows(0).Item("name_suffix")
                sqlCommonAction.Parameters.Add(param4)
                Dim param5 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@strWorkPhone", System.Data.SqlDbType.VarChar, 32, "strWorkPhone")
                param5.Value = Me.dsCommon.Tables("PersonData").Rows(0).Item("phone_no")
                sqlCommonAction.Parameters.Add(param5)
                Dim param6 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@strCompany", System.Data.SqlDbType.VarChar, 32, "strCompany")
                param6.Value = Me.dsCommon.Tables("PersonData").Rows(0).Item("employed_by_hcid")
                sqlCommonAction.Parameters.Add(param6)
                Dim param7 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@strHID", System.Data.SqlDbType.VarChar, 12, "strHID")

                'get the user's hid from the stored table
                Dim dt As DataTable = CType(Session("PersonTable"), DataTable)
                param7.Value = dt.Rows(0).Item("hanford_id")
                sqlCommonAction.Parameters.Add(param7)

                'generate a password for this user
                Dim param8 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@strAuthenticator", System.Data.SqlDbType.VarChar, 64, "strAuthenticator")
                param8.Value = myUtility.hashPassword("De7en!stration", param0.value)
                sqlCommonAction.Parameters.Add(param8)

                'insert the new user
                sqlCommonAction.CommandText = "Insert INTO " & myUtility.getProdExtension() & "fempUsers (strHID, strEmail, intUserType, strLastName, strFirstName, " & _
                "strMiddleName, strNameSuffix, strAuthenticator, strWorkPhone, datDateCreated, datDateUsed, intCreatorID, blnActive, " & _
                "blnTraining, strCompany, blnNoEmail) VALUES (@strHID, @strEmail, 0, @strLastName," & _
                "@strFirstName, @strMiddleName, @strNameSuffix, @strAuthenticator, @strWorkPhone,'" & Now & "','1/1/1970'," & _
                Session("UserID") & "," & _
                IIf(Me.dsCommon.Tables("PersonData").Rows(0).Item("active_sw") = "T", 1, 0) & ", 0" & _
                ", @strCompany, 0)"
                sqlCommonAction.ExecuteNonQuery()

                'get the user's sequence number
                sqlCommonAction.CommandText = "Select * from " & myUtility.getProdExtension() & "fempUsers where strMiddleName = @strMiddleName and strLastName = " & _
                "@strLastName and strFirstName = @strFirstName and strNameSuffix = @strNameSuffix"
                sqlCommonAdapter.Fill(Me.dsCommon, "UserIDFetch")
                Session("mailListID") = Me.dsCommon.Tables("UserIDFetch").Rows(0).Item("seqUserID")
                Return True
            Else
                'parameterized the text input to allow for things like quotes to get through
                Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@strEmail", System.Data.SqlDbType.VarChar, 128, "strEmail")
                param0.Value = Session("PersonInputEmailReturn")
                sqlCommonAction.Parameters.Add(param0)
                Dim param1 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@strFirstName", System.Data.SqlDbType.VarChar, 50, "strFirstName")
                param1.Value = Session("PersonInputFirstName")
                sqlCommonAction.Parameters.Add(param1)
                Dim param2 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@strLastName", System.Data.SqlDbType.VarChar, 50, "strLastName")
                param2.Value = Session("PersonInputLastName")
                sqlCommonAction.Parameters.Add(param2)
                Dim param3 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@strMiddleName", System.Data.SqlDbType.VarChar, 50, "strMiddleName")
                param3.Value = Session("PersonInputMiddleInitial")
                sqlCommonAction.Parameters.Add(param3)

                'insert the new user
                sqlCommonAction.CommandText = "Insert INTO " & myUtility.getProdExtension() & "fempUsers (strHID, strEmail, intUserType, strLastName, strFirstName, " & _
                "strMiddleName, strNameSuffix, strAuthenticator, strWorkPhone, datDateCreated, datDateUsed, intCreatorID, blnActive, " & _
                "blnTraining, strCompany, blnNoEmail) VALUES ('', @strEmail, 0, " & _
                "@strLastName, @strFirstName, @strMiddleName, '', '" & _
                myUtility.hashPassword("De7en!stration", param0.Value) & _
                "', '','" & Now & "','1/1/1970'," & Session("UserID") & "," & _
                " 0, 0, '', 0)"
                sqlCommonAction.ExecuteNonQuery()

                'get the user's sequence number
                sqlCommonAction.CommandText = "Select * from " & myUtility.getProdExtension() & "fempUsers where strEmail = @strEmail"
                sqlCommonAdapter.Fill(Me.dsCommon, "UserIDFetch")
                Session("mailListID") = Me.dsCommon.Tables("UserIDFetch").Rows(0).Item("seqUserID")
                Return True
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            sqlCommonAction.Parameters.Clear()
            Return False
        End Try
    End Function

    'attempts to insert permissions for the creator into the permissions table, returns false if it cannot
    Private Function processGroupList(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            Trace.Warn("Checking Users")
            Dim Index As Integer = 0
            While (Index < Me.dgGroup.Items.Count())
                If (CType(Me.dgGroup.Items(Index).Controls(0).Controls(1), CheckBox).Checked = True) Then
                    addToLists(CType(Me.dgGroup.Items(Index).Controls(1), TableCell).Text)
                End If
                Index += 1
            End While
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'attempts to insert permissions for the creator into the permissions table, returns false if it cannot
    Private Function processUserList(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            Trace.Warn("Checking Users")
            Dim Index As Integer = 0
            While (Index < Me.dgUsers.Items.Count())
                If (CType(Me.dgUsers.Items(Index).Controls(0).Controls(1), CheckBox).Checked = True) Then
                    addToLists(CType(Me.dgUsers.Items(Index).Controls(1), TableCell).Text)
                End If
                Index += 1
            End While
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'commits the group data generated to the groups table
    Private Function commitGroup(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal strAction As String) As Boolean
        Try
            'set up the group name parameter
            sqlCommonAction.Parameters.Clear()
            Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@strGroupName", System.Data.SqlDbType.VarChar, 300, "strGroupName")
            param0.Value = Me.txtGroupName.Text
            sqlCommonAction.Parameters.Add(param0)

            'do the insert or update
            If (strAction = "Insert") Then
                If (Session("intGroupID") = -1) Then
                    sqlCommonAction.CommandText = "Insert INTO " & myUtility.getProdExtension() & "fempGroups (seqSurvey, intGroupID, " & _
                    "strGroupName, strUserList) VALUES (" & Session("seqSurvey") & ", " & _
                    Session("GroupsMax") + 1 & ", @strGroupName, '" & Me.mailList & "')"
                    sqlCommonAction.ExecuteNonQuery()
                    Session("intGroupID") = Session("GroupsMax") + 1
                    Return True
                Else
                    'update the group id's
                    Dim lowIndex As Integer = Session("GroupsRow")
                    Dim highIndex As Integer = Me.dsCommon.Tables("Groups").Rows.Count() - 1

                    'update the group id's of the groups
                    While (highIndex >= lowIndex)
                        sqlCommonAction.CommandText = "Update " & myUtility.getProdExtension() & "fempGroups SET intGroupID = (" & _
                        "intGroupID+1) WHERE seqSurvey = " & Session("seqSurvey") & _
                        " and intGroupID = " & _
                        Me.dsCommon.Tables("Groups").Rows(highIndex).Item("intGroupID")
                        sqlCommonAction.ExecuteNonQuery()
                        highIndex -= 1
                    End While

                    'attempt to add the new record to the database
                    sqlCommonAction.CommandText = "Insert INTO " & myUtility.getProdExtension() & "fempGroups (seqSurvey, intGroupID, " & _
                    "strGroupName, strUserList) VALUES (" & Session("seqSurvey") & ", " & _
                    Session("GroupsRow") + 1 & ", @strGroupName, '" & Me.mailList & "')"
                    sqlCommonAction.ExecuteNonQuery()
                    Session("intGroupID") = Session("GroupsRow") + 1
                    Return True
                End If
            Else
                'just update the current record
                sqlCommonAction.CommandText = "Update " & myUtility.getProdExtension() & "fempGroups SET strGroupName = @strGroupName, " & _
                "strUserList = '" & Me.mailList & "' where seqSurvey = " & _
                Session("seqSurvey") & " and intGroupID = " & Session("intGroupID")
                sqlCommonAction.ExecuteNonQuery()
                Return True
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'commits the permissions by getting rid of the old permissions and adding the new permissions
    Private Function commitPermissions(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal strOldList As String) As Boolean
        'variables
        Dim oldList As Array
        Dim newList As Array
        Dim masterList As Array
        Dim index As Integer = 0
        Dim blnProcessOldList As Boolean = True
        Dim blnProcessNewList As Boolean = True
        Dim blnFailure As Boolean = False

        Try
            'create an array of users while trimming the white space
            If Not (strOldList Is Nothing) Then
                If (strOldList.Length > 0) Then
                    Dim strLine As String = ""
                    'replace white space text
                    oldList = strOldList.Split(",")
                    While index < oldList.Length
                        oldList(index) = Trim(CType(oldList(index), String))
                        index += 1
                    End While
                End If

                'determine if we actually got anything.  If not then skip this function
                If (oldList Is Nothing) Then
                    blnProcessOldList = False
                End If
            Else
                blnProcessOldList = False
            End If

            'create an array of users while trimming the white space
            index = 0
            If (Me.mailList.Length > 0) Then
                Dim strLine As String = ""
                'replace white space text
                newList = Me.mailList.Split(",")
                While index < newList.Length
                    newList(index) = Trim(CType(newList(index), String))
                    index += 1
                End While
            End If

            'determine if we actually got anything.  If not then skip this function
            If (newList Is Nothing And oldList Is Nothing) Then
                Return True
            ElseIf (newList Is Nothing) Then
                blnProcessNewList = False
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try

        'get the other group listings and combine them into one master list for the survey
        sqlCommonAction.CommandText = "Select strUserList from " & myUtility.getProdExtension() & "fempGroups where seqSurvey = " & Session("seqSurvey") & _
        "and intGroupID <> " & Session("GroupsRow") + 1
        sqlCommonAdapter.SelectCommand = sqlCommonAction
        sqlCommonAdapter.Fill(Me.dsCommon, "SurveyGroups")
        Dim dr As DataRow
        Dim strTemp As String = ""
        index = 0
        For Each dr In Me.dsCommon.Tables("SurveyGroups").Rows
            If (index = 0) Then
                strTemp &= dr.Item("strUserList")
            Else
                If (strTemp <> "" And dr.Item("strUserList") <> "") Then
                    strTemp &= "," & dr.Item("strUserList")
                Else
                    strTemp &= dr.Item("strUserList")
                End If
            End If
            index += 1
        Next
        'replace white space text
        index = 0
        masterList = strTemp.Split(",")
        While index < masterList.Length
            masterList(index) = Trim(CType(masterList(index), String))
            index += 1
        End While

        'update the permissions for the old list
        If (blnProcessOldList) Then
            index = 0
            blnFailure = False
            While (index < oldList.Length)
                If Not (masterList.IndexOf(masterList, oldList(index)) >= 0) Then
                    Try
                        'attempt to insert a permission record with no permissions if that fails do an update and remove the user's user permission
                        sqlCommonAction.CommandText = "Insert INTO " & myUtility.getProdExtension() & "fempPermission (seqUserID, intCreatorID, seqTemplate, seqSurvey, datCreatedDate, " & _
                        "OWNER_IND, blnDelegate, blnUser) VALUES (" & oldList(index) & ", " & Session("UserID") & ", -1, " & _
                        Session("seqSurvey") & ", '" & Now & "', 0, 0, 0)"
                        sqlCommonAction.ExecuteNonQuery()
                        blnFailure = False
                    Catch ex As Exception
                        Trace.Warn(ex.ToString)
                        blnFailure = True
                    End Try

                    If (blnFailure) Then
                        Try
                            'attempt to update the exisiting record by removing the user permission
                            sqlCommonAction.CommandText = "Update " & myUtility.getProdExtension() & "fempPermission set blnUser = 0 where " & _
                            "seqSurvey = " & Session("seqSurvey") & " and seqUserID = " & oldList(index)
                            sqlCommonAction.ExecuteNonQuery()
                        Catch ex As Exception
                            Trace.Warn(ex.ToString)
                            Return False
                        End Try
                    End If

                    blnFailure = False
                End If
                index += 1
            End While
        End If

        'update the permissions for the new list
        If (blnProcessNewList) Then
            index = 0
            blnFailure = False
            While (index < newList.Length)
                If Not (masterList.IndexOf(masterList, newList(index)) >= 0) Then
                    Try
                        'attempt to insert a permission record with no permissions if that fails do an update and remove the user's user permission
                        sqlCommonAction.CommandText = "Insert INTO " & myUtility.getProdExtension() & "fempPermission (seqUserID, intCreatorID, seqTemplate, seqSurvey, datCreatedDate, " & _
                        "OWNER_IND, blnDelegate, blnUser) VALUES (" & newList(index) & ", " & Session("UserID") & ", -1, " & _
                        Session("seqSurvey") & ", '" & Now & "', 0, 0, 1)"
                        sqlCommonAction.ExecuteNonQuery()
                        blnFailure = False
                    Catch ex As Exception
                        Trace.Warn(ex.ToString)
                        blnFailure = True
                    End Try

                    If (blnFailure) Then
                        Try
                            'attempt to update the exisiting record by removing the user permission
                            sqlCommonAction.CommandText = "Update " & myUtility.getProdExtension() & "fempPermission set blnUser = 1 where " & _
                            "seqSurvey = " & Session("seqSurvey") & " and seqUserID = " & newList(index)
                            sqlCommonAction.ExecuteNonQuery()
                        Catch ex As Exception
                            Trace.Warn(ex.ToString)
                            Return False
                        End Try
                    End If

                    blnFailure = False
                End If
                index += 1
            End While
        End If
        Return True
    End Function
#End Region

#Region "Delete Code"
    'deletes the current Group from the fempGroups table
    Private Function doDelete() As Boolean
        Try
            'start transaction
            If (myUtility.setupTransaction(sqlCommonAction, Me.commonConn)) Then
                Dim failure As Boolean = False

                'do the deletion
                If (deletePermissions(sqlCommonAction, sqlCommonAdapter, Me.mailList)) Then
                    If (deleteGroup(sqlCommonAction, sqlCommonAdapter)) Then
                        sqlCommonAction.Transaction.Commit()
                        'determine if we need to take a step back 
                        If (Session("GroupsMax") = 1) Then
                            Session("GroupsRow") = 0
                            Session("intGroupID") = -1
                            Session("GroupsMax") = 0
                            Me.dsCommon.Tables("Groups").Rows.Add(Me.dsCommon.Tables("Groups").NewRow())
                            Me.dsCommon.Tables("Groups").Rows(Session("GroupsMax")).Item("seqSurvey") = Session("seqSurvey")
                            Me.dsCommon.Tables("Groups").Rows(Session("GroupsMax")).Item("intGroupID") = -1
                            Me.dsCommon.Tables("Groups").Rows(Session("GroupsMax")).Item("strGroupName") = ""
                            Me.dsCommon.Tables("Groups").Rows(Session("GroupsMax")).Item("strUserList") = ""
                        ElseIf (Session("intGroupID") = Session("GroupsMax") And Session("GroupsMax") - Session("GroupsRow") = 1) Then
                            Session("GroupsMax") -= 1
                            Session("GroupsRow") -= 1
                            Session("intGroupID") -= 1
                        Else
                            Session("GroupsMax") -= 1
                        End If
                    Else
                        sqlCommonAction.Transaction.Rollback()
                        alert("Error deleting group from this survey.")
                        failure = True
                    End If
                    Return failure
                Else
                    If (Session("transExists") = True) Then
                        sqlCommonAction.Transaction.Rollback()
                    End If
                    alert("Error deleting group from this survey.")
                    Return True
                End If
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Error deleting group from this survey.")
                Return True
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Error deleting group from this survey.")
            Return True
        End Try
    End Function

    'attempts to delete a group fempGroups table.  
    Private Function deleteGroup(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            'attempt to delete the record from the database
            sqlCommonAction.CommandText = "Delete from " & myUtility.getProdExtension() & "fempGroups where seqSurvey = " & Session("seqSurvey") & _
            " and intGroupID = " & Session("intGroupID")
            sqlCommonAction.ExecuteNonQuery()

            'update the group id's
            Dim lowIndex As Integer = Session("GroupsRow")
            Dim highIndex As Integer = Me.dsCommon.Tables("Groups").Rows.Count()

            'update the group id's of the groups for the current survey
            While (lowIndex < highIndex)
                sqlCommonAction.CommandText = "Update " & myUtility.getProdExtension() & "fempGroups SET intGroupID = (intGroupID-1) " & _
                "WHERE seqSurvey = " & Session("seqSurvey") & " and intGroupID = " & _
                Me.dsCommon.Tables("Groups").Rows(lowIndex).Item("intGroupID")
                sqlCommonAction.ExecuteNonQuery()
                lowIndex += 1
            End While
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'deletes the permissions by setting the user permissions to 0
    Private Function deletePermissions(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal strList As String) As Boolean
        'variables
        Dim List As Array
        Dim masterList As Array
        Dim index As Integer = 0
        Dim blnProcessList As Boolean = True
        Dim blnFailure As Boolean = False

        Try
            'create an array of users while trimming the white space
            If Not (strList Is Nothing) Then
                If (strList.Length > 0) Then
                    Dim strLine As String = ""
                    'replace white space text
                    List = strList.Split(",")
                    While index < List.Length
                        List(index) = Trim(CType(List(index), String))
                        index += 1
                    End While
                End If

                'determine if we actually got anything.  If not then skip this function
                If (List Is Nothing) Then
                    blnProcessList = False
                End If
            Else
                blnProcessList = False
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try

        'get the other group listings and combine them into one master list for the survey
        sqlCommonAction.CommandText = "Select strUserList from " & myUtility.getProdExtension() & "fempGroups where seqSurvey = " & Session("seqSurvey") & _
        "and intGroupID <> " & Session("GroupsRow") + 1
        sqlCommonAdapter.SelectCommand = sqlCommonAction
        sqlCommonAdapter.Fill(Me.dsCommon, "SurveyGroups")
        Dim dr As DataRow
        Dim strTemp As String = ""
        index = 0
        For Each dr In Me.dsCommon.Tables("SurveyGroups").Rows
            If (index = 0) Then
                strTemp &= dr.Item("strUserList")
            Else
                If (strTemp <> "" And dr.Item("strUserList") <> "") Then
                    strTemp &= "," & dr.Item("strUserList")
                Else
                    strTemp &= dr.Item("strUserList")
                End If
            End If
            index += 1
        Next
        'replace white space text
        index = 0
        masterList = strTemp.Split(",")
        While index < masterList.Length
            masterList(index) = Trim(CType(masterList(index), String))
            index += 1
        End While

        'update the permissions for the list
        If (blnProcessList) Then
            index = 0
            While (index < List.Length)
                If Not (masterList.IndexOf(masterList, List(index)) >= 0) Then
                    Try
                        'attempt to update the exisiting record by removing the user permission
                        sqlCommonAction.CommandText = "Update " & myUtility.getProdExtension() & "fempPermission set blnUser = 0 where " & _
                        "seqSurvey = " & Session("seqSurvey") & " and seqUserID = " & List(index)
                        sqlCommonAction.ExecuteNonQuery()
                    Catch ex As Exception
                        Trace.Warn(ex.ToString)
                        Return False
                    End Try
                End If
                index += 1
            End While
        End If
        Return True
    End Function
#End Region

#Region "Reset Code"
    'resets the page, will remove any new item being worked on at the end of the list
    Private Function doReset() As Boolean
        If Not (loadData()) Then
            alert("Failed to load the groups for the current survey.")
        Else
            Try
                If (Session("intGroupID") = -1) Then
                    Me.dsCommon.Tables("Groups").Rows.Add(Me.dsCommon.Tables("Groups").NewRow())
                    Me.dsCommon.Tables("Groups").Rows(Session("GroupsMax")).Item("seqSurvey") = Session("seqSurvey")
                    Me.dsCommon.Tables("Groups").Rows(Session("GroupsMax")).Item("intGroupID") = -1
                    Me.dsCommon.Tables("Groups").Rows(Session("GroupsMax")).Item("strGroupName") = ""
                    Me.dsCommon.Tables("Groups").Rows(Session("GroupsMax")).Item("strUserList") = ""
                End If
                Return False
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error resetting the current group.")
                Return True
            End Try
        End If
    End Function
#End Region
End Class
