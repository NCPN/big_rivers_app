Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    OrderByOn = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4560
    DatasheetFontHeight =9
    ItemSuffix =4
    Left =9960
    Top =2085
    Right =14610
    Bottom =10680
    DatasheetGridlinesColor =12632256
    OrderBy ="[frm_User_Roles].[User_role], User_name"
    RecSrcDt = Begin
        0xc71a6a0bef61e340
    End
    RecordSource ="SELECT tsys_User_Roles.* FROM tsys_User_Roles ORDER BY tsys_User_Roles.User_name"
        "; "
    Caption =" Set User Access Roles"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =360
            BackColor =12433075
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =60
                    Top =60
                    Width =2310
                    Height =240
                    Name ="labUser_name"
                    Caption ="User login"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =2430
                    Top =60
                    Width =1770
                    Height =240
                    Name ="labUser_role"
                    Caption ="Application role"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            Height =375
            BackColor =12574431
            Name ="Detail"
            Begin
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =2310
                    Height =255
                    ColumnWidth =2310
                    Name ="txtUser_name"
                    ControlSource ="User_name"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =6192
                    Left =2430
                    Top =60
                    Width =2070
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="cmbUser_role"
                    ControlSource ="User_role"
                    RowSourceType ="Value List"
                    RowSource ="\"read only\";\"view data only\";\"data entry\";\"enter/edit data for the curren"
                        "t season\";\"power user\";\"set user roles, edit lookup values, edit certified d"
                        "ata\";\"admin\";\"create new application releases, update app contact info\""
                    ColumnWidths ="1152;5040"
                    DefaultValue ="\"data entry\""

                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' FORM NAME:    frm_User_Roles
' Description:  Standard form for setting user access privileges
' Data source:  In-line SQL statement based on tsys_User_Roles
' Data access:  add, edit, delete
' Pages:        none
' Functions:    none
' References:   fxnSwitchboardIsOpen
' Source/date:  John R. Boetsch, September 2008
' Revisions:    JRB, 11/5/2009 - Updated to provide power user with edit and delete capability
'               BLC, 6/12/2014 - Revised to use TempVars.Item("UserAccessLevel") vs. cAppMode
'                                for readability & accessibility w/o referencing subform control
' =================================

Dim varOpenArgs As Variant

' ---------------------------------
' SUB:     Form_Open
' Description: Opens sub form & sets controls based on UserAccessLevel
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Created John R. Boetsch, September 2008
'               Adapted 06/12/2014 Bonnie Campbell, June 2014
'               Revised 06/12/2014 blc; Last modified 06/12/2014 blc.
' Revisions:    Bonnie Campbell, June 12, 2014 - Revised to use TempVars.Item("UserAccessLevel") vs. cAppMode
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    varOpenArgs = Me.OpenArgs
    ' Control who can make edits or deletions
    If fxnSwitchboardIsOpen Then
        Select Case TempVars.item("UserAccessLevel")
            Case "admin"
                Me.AllowEdits = True
                Me.AllowDeletions = True
            Case "power user"
                Me.AllowEdits = True
                Me.AllowDeletions = True
            Case Else
                MsgBox "Access denied", , "Cannot open the form ..."
                DoCmd.CancelEvent
        End Select
    Else
        MsgBox "The main database switchboard must be" & vbCrLf & _
            "open for this form to function properly.", , "Cannot open the form ..."
        DoCmd.CancelEvent
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub
