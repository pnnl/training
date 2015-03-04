Public Class Navigation
    'Sets what is available during a reload of the page from one of the listed links
    Public Sub wnb_MoveTo(ByRef NavBar As Collection, ByVal MaxRows As Integer, ByVal CurrentRow As Integer, Optional ByVal isSwitch As Boolean = False, Optional ByVal isProduction As Boolean = False)
        If (CurrentRow = 0 And MaxRows = 0) Then
            CType(NavBar.Item(1), ImageButton).Enabled = False
            CType(NavBar.Item(1), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/start_disabled.gif"
            CType(NavBar.Item(2), ImageButton).Enabled = False
            CType(NavBar.Item(2), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/prev_disabled.gif"
            CType(NavBar.Item(4), ImageButton).Enabled = False
            CType(NavBar.Item(4), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/next_disabled.gif"
            CType(NavBar.Item(5), ImageButton).Enabled = False
            CType(NavBar.Item(5), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/end_disabled.gif"
            CType(NavBar.Item(6), ImageButton).Enabled = False
            CType(NavBar.Item(6), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/add_disabled.gif"
        ElseIf (CurrentRow = 0 And MaxRows = 1) Then
            CType(NavBar.Item(1), ImageButton).Enabled = False
            CType(NavBar.Item(1), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/start_disabled.gif"
            CType(NavBar.Item(2), ImageButton).Enabled = False
            CType(NavBar.Item(2), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/prev_disabled.gif"
            CType(NavBar.Item(4), ImageButton).Enabled = False
            CType(NavBar.Item(4), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/next_disabled.gif"
            CType(NavBar.Item(5), ImageButton).Enabled = False
            CType(NavBar.Item(5), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/end_disabled.gif"
            If (isProduction) Then
                CType(NavBar.Item(6), ImageButton).Visible = False
                CType(NavBar.Item(6), ImageButton).Enabled = False
                CType(NavBar.Item(6), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/add_disabled.gif"
            Else
                CType(NavBar.Item(6), ImageButton).Visible = True
                CType(NavBar.Item(6), ImageButton).Enabled = True
                CType(NavBar.Item(6), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/add_up.gif"
            End If
        ElseIf (CurrentRow = 0 And MaxRows > 1) Then
            CType(NavBar.Item(1), ImageButton).Enabled = False
            CType(NavBar.Item(1), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/start_disabled.gif"
            CType(NavBar.Item(2), ImageButton).Enabled = False
            CType(NavBar.Item(2), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/prev_disabled.gif"
            CType(NavBar.Item(4), ImageButton).Enabled = True
            CType(NavBar.Item(4), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/next_up.gif"
            CType(NavBar.Item(5), ImageButton).Enabled = True
            CType(NavBar.Item(5), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/end_up.gif"
            If (isProduction) Then
                CType(NavBar.Item(6), ImageButton).Visible = False
                CType(NavBar.Item(6), ImageButton).Enabled = False
                CType(NavBar.Item(6), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/add_disabled.gif"
            Else
                CType(NavBar.Item(6), ImageButton).Visible = True
                CType(NavBar.Item(6), ImageButton).Enabled = True
                CType(NavBar.Item(6), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/add_up.gif"
            End If
        ElseIf (CurrentRow = MaxRows - 1) Then
            CType(NavBar.Item(1), ImageButton).Enabled = True
            CType(NavBar.Item(1), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/start_up.gif"
            CType(NavBar.Item(2), ImageButton).Enabled = True
            CType(NavBar.Item(2), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/prev_up.gif"
            CType(NavBar.Item(4), ImageButton).Enabled = False
            CType(NavBar.Item(4), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/next_disabled.gif"
            CType(NavBar.Item(5), ImageButton).Enabled = False
            CType(NavBar.Item(5), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/end_disabled.gif"
            If (isProduction) Then
                CType(NavBar.Item(6), ImageButton).Visible = False
                CType(NavBar.Item(6), ImageButton).Enabled = False
                CType(NavBar.Item(6), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/add_disabled.gif"
            Else
                CType(NavBar.Item(6), ImageButton).Visible = True
                CType(NavBar.Item(6), ImageButton).Enabled = True
                CType(NavBar.Item(6), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/add_up.gif"
            End If
        ElseIf (CurrentRow = MaxRows) Then
            CType(NavBar.Item(1), ImageButton).Enabled = True
            CType(NavBar.Item(1), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/start_up.gif"
            CType(NavBar.Item(2), ImageButton).Enabled = True
            CType(NavBar.Item(2), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/prev_up.gif"
            CType(NavBar.Item(4), ImageButton).Enabled = False
            CType(NavBar.Item(4), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/next_disabled.gif"
            CType(NavBar.Item(5), ImageButton).Enabled = False
            CType(NavBar.Item(5), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/end_disabled.gif"
            If (isProduction) Then
                CType(NavBar.Item(6), ImageButton).Visible = False
            Else
                CType(NavBar.Item(6), ImageButton).Visible = True
            End If
            CType(NavBar.Item(6), ImageButton).Enabled = False
            CType(NavBar.Item(6), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/add_disabled.gif"
        Else
            CType(NavBar.Item(1), ImageButton).Enabled = True
            CType(NavBar.Item(1), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/start_up.gif"
            CType(NavBar.Item(2), ImageButton).Enabled = True
            CType(NavBar.Item(2), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/prev_up.gif"
            CType(NavBar.Item(4), ImageButton).Enabled = True
            CType(NavBar.Item(4), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/next_up.gif"
            CType(NavBar.Item(5), ImageButton).Enabled = True
            CType(NavBar.Item(5), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/end_up.gif"
            If (isProduction) Then
                CType(NavBar.Item(6), ImageButton).Visible = False
                CType(NavBar.Item(6), ImageButton).Enabled = False
                CType(NavBar.Item(6), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/add_disabled.gif"
            Else
                CType(NavBar.Item(6), ImageButton).Visible = True
                CType(NavBar.Item(6), ImageButton).Enabled = True
                CType(NavBar.Item(6), ImageButton).ImageUrl = "./MyInfragisticsImages/XpBlue/add_up.gif"
            End If
        End If
    End Sub
End Class
