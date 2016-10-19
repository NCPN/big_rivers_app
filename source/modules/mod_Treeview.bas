Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Treeview
' Level:        Framework module
' Version:      1.00
' Description:  treeview functions & procedures
'
' Source/date:  Bonnie Campbell, 8/30/2016
' Revisions:    BLC, 8/30/2016 - 1.00 - initial version
' =================================

' ---------------------------------
'  Properties
' ---------------------------------

' ---------------------------------
'  Subroutines & Functions
' ---------------------------------

' ---------------------------------
' SUB:          LoadTree
' Description:  treeview loading actions
' Assumptions:
'               All static, immovable nodes have 1-letter keys:
'                   R-reference     V-overview      F-feature
'                   T-transect      O-other         A-animal
'                   P-plant         C-cultural      D-disturbance
'                   W-field work    S-scenic        W-weather
'                   O-other         U-unclassified
'
'   s_photo_data -> complete photo data w/ appropriate form data supplied & submitted
'   s_tsys_temp_photo_data -> incomplete, but imported photo files
'
' Parameters:   frm - treeview control's parent form (form)
'               tvw - treeview control to load (treeview)
'               template - query template to load from (string)
'               params - array of parameters to limit data from datasource (variant)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, July 10, 2015 - for NCPN tools
' Revisions:
'   BLC - 7/10/2015 - initial version
'   BLC - 8/31/2016 - load from query or table
' ---------------------------------
Public Sub LoadTree(frm As Form, tvw As Treeview, Template As String, Params As Variant)
On Error GoTo Err_Handler
    
    'exit w/o values
    If Not IsArray(Params) Then GoTo Exit_Handler
    
    'variables
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim strPhotoPath As String

    'default
    strPhotoPath = ""

    'retrieve data
    Set db = CurrentDb
    
    With db
        Set qdf = .QueryDefs("usys_temp_qdf")
        
        With qdf
        
            'check if record exists in site
            .sql = GetTemplate(Template)
            
            '-------------------
            ' set SQL parameters --> .Parameters("") = params()
            '-------------------
            
            '-------------------------------------------------------------------------
            ' NOTE:
            '   param(0) --> reserved for record action RefTable (ReferenceType)
            '   last param(x) --> used as record ID for updates
            '-------------------------------------------------------------------------
            Select Case Template
            
        '-----------------------
        '  SELECT
        '-----------------------
                Case "s_photo_data"
                
                    'use PHOTO_PATH vs. rs!PhotoPath --> always NULL for this query
                    strPhotoPath = PHOTO_PATH
                    
                    '-- required parameters --
'                    .Parameters("PhotoDate") = params(1)
'                    .Parameters("PhotoType") = params(2)
'                    .Parameters("PhotographerID") = params(3)
'                    .Parameters("FileName") = params(4)
'                    .Parameters("NCPNImageID") = params(5)
'                    .Parameters("DirectionFacing") = params(6)
'                    .Parameters("PhotogLocation") = params(7)
'                    .Parameters("IsCloseup") = params(8)
'                    .Parameters("IsInActive") = params(9)
'                    .Parameters("IsSkipped") = params(10)
'                    .Parameters("IsReplacement") = params(11)
'                    .Parameters("LastPhotoUpdate") = params(12)
'
'                    .Parameters("CreateDate") = Now()
'                    .Parameters("CreatedByID") = TempVars("ContactID")
'                    .Parameters("LastModified") = Now()
'                    .Parameters("LastModifiedByID") = TempVars("ContactID")
                
                Case "s_usys_temp_photo_data"
                    
                    'use rs!PhotoPath
                    
                    '-- required parameters --
'                    .Parameters("ptype") = params(1)
                
            End Select
            
            'populate rs
            Set rs = .OpenRecordset(dbOpenDynaset) 'dbOpenDynamic fails w/ Error #3001 Invalid argument
            
            If Not (rs.BOF And rs.EOF) Then
                        
                'determine # records
                rs.MoveLast
                rs.MoveFirst
                
                'iterate
                If rs.RecordCount > 0 Then
                    
                    'variables
                    Dim oTree As MSComctlLib.Treeview
                    Dim strKey As String, strText As String, strDisplayName As String
                    Dim strPhotoType As String
                    Dim strDuplicates As String
                    Dim nodeSelected As node
                    Dim nodeParent As node
                    Dim nodeX As node
                    Dim nodeNew As node
                    
                    'default
                    strPhotoType = "U"
            
                '---------------
                ' load tree
                '---------------
                                        
                    'Create a reference to the TreeView control
                    Set oTree = tvw
                
                    Do While Not rs.EOF
                
                        strPhotoType = rs!PhotoType
                        
                        'select the photo type, immovable node
                        oTree.Nodes(strPhotoType).Selected = True
                                               
                        'selected node = immovable --> highlight selected item only
                        Set oTree.SelectedItem = oTree.DropHighlight
                    
                        'select the appropriate immovable node
                        'oTree.Object.Nodes(strPhotoType).Selected = True
                        oTree.Nodes(strPhotoType).Selected = True
                    
                        'Reference the selected node as the one being added to.
                        Set nodeSelected = oTree.SelectedItem
                
                        If ImmovableNode(nodeSelected) Then
                
                            'add children here
                                        
                            'Relative, Relationship, Key, Text
                            'Unique Key --> absolute path to the file
                            'Displayed Text --> file name w/ extension
                            strKey = IIf(Len(strPhotoPath) = 0, rs!PhotoPath & "\" & rs!PhotoFilename, _
                                        strPhotoPath & rs!PhotoFilename & ".jpg")
                            
                            strDisplayName = Replace(rs!PhotoFilename, ".jpg", "")
                                        
                            'Save key & text to use when node re-added
                            'strKey = nodeSelected.key
                            'strText = nodeSelected.Text
                            
                            'check for duplicate keys
                            If Not IsDuplicateKey(strKey, oTree) Then
                                    
                                'check to see if node was static parent or child -> add only to parents
                                If Len(oTree.SelectedItem.key) > 2 Then
                                    strPhotoType = oTree.SelectedItem.Parent.key
                                    Set nodeParent = oTree.SelectedItem.Parent
                                Else
                                    strPhotoType = oTree.SelectedItem.key
                                    Set nodeParent = oTree.SelectedItem
                                End If
                                    
                                'add node & tag
                                Set nodeX = oTree.Nodes.Add(nodeParent, tvwChild, strKey, strDisplayName)
                                nodeX.Tag = "M|C|" & strKey & "|" & strDisplayName & "|" & strPhotoType 'oTree.SelectedItem.key 'strDisplayName
                                
                                'adjust node font weight/color for incomplete data
                                If Template = "s_usys_temp_photo_data" Then
                                    nodeX.ForeColor = lngRed
                                    nodeX.Bold = True
                                End If
                                
                                'select the relocated node
                                'oTree.SelectedItem = nodX
                                
                                'get the parent key to identify the form to view
                                'MsgBox nodX.Parent, vbInformation
                                TempVars("PhotoType") = strPhotoType
                                frm.lblPhotoTypeValue.Caption = nodeX.Parent
                                
                                'set full photo path
                                TempVars("FullPhotoPath") = ParseString(nodeX.Tag, 2)
                            
                            Else
                                'prepare duplicate message
                                strDuplicates = IIf(Len(strDuplicates) > 0, strDuplicates & ",", vbCrLf & "Skipped duplicates:  ") & strDisplayName
                            
                            End If
                            
                
                        End If
                
                       rs.MoveNext
                       
                    Loop
                
                
                '----------
                
        
'        '-------------------------------
'        '  Node Added to Empty Space
'        '-------------------------------
'        ' update the db table & make it a root node
'        If oTree.DropHighlight Is Nothing Then
'
'            'Save key & text to use when node re-added
'            strKey = nodeSelected.key
'            strText = nodeSelected.Text
'
'             'selected node = immovable --> highlight selected item only
'             If ImmovableNode(nodeSelected) Then Set oTree.SelectedItem = oTree.DropHighlight
'
'
'        '-------------------------------
'        '  Node Added to Another Node
'        '-------------------------------
'        Else
'
'            'get new parent node info
''            If CountInString(PHOTO_TYPES_MAIN, nodeDragged.key) + CountInString(PHOTO_TYPES_OTHER, nodeDragged.key) = 0 Then
''             If Not ImmovableNode(nodeDragged) Then
'
'            'Save key & text to use when node re-added
'            strKey = nodeSelected.key
'            strText = nodeSelected.Text
'
'            'if the selected node is immovable, set the new parent
'            If ImmovableNode(nodeSelected) Then Set oTree.SelectedItem = oTree.DropHighlight
'
'            If Not ImmovableNode(nodeSelected) Then
'                'Delete the current node for the photo
'                oTree.Nodes.Remove nodeSelected.index
'
'                'Add to new location
'                Set nodeNew = oTree.Nodes.Add(oTree.DropHighlight, tvwChild, strKey, strText)
'                nodeNew.Tag = "M|C|" & strKey & "|" & strText & "|" & oTree.DropHighlight.key
'
'                'update photo type
'                TempVars("PhotoType") = oTree.DropHighlight.key
'                frm.Controls("lblPhotoTypeValue").Caption = nodeNew.Parent
'
'                'highlight the new node
'                oTree.SelectedItem = nodeNew
'                oTree.DropHighlight = nodeNew
'            End If
'
'        End If
'    End If

                
                
                
                '--------------
                
                
                End If
            
            End If
            
            'cleanup
            .Close
        
        End With

    End With
                
Exit_Handler:
    'cleanup
    Set qdf = Nothing
    Set db = Nothing
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - LoadTree[mod_Treeview])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          AddChildren
' Description:  add children to treeview node
' Assumptions:  -
' Parameters:   tvw - treeview control
'               nodeParent - parent node for children (node)
'               aryKids - comma separated list of children (string)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, July 10, 2015 - for NCPN tools
' Revisions:
'   BLC - 7/10/2015 - initial version
'   BLC - 6/15/2016 - adapted for big rivers app
' ---------------------------------
Public Sub AddChildren(tvw As Treeview, nodeParent As node, aryKids As String)
'Private Sub AddChildren(tvw As TreeView, nodeParent As String, aryChildren As String)

On Error GoTo Err_Handler
    
    Dim nodeX As node
    Dim aryChildren() As String
    Dim child As Variant
    
    'set the array
    aryChildren = Split(aryKids, ",")
    
    For Each child In aryChildren
        'Set nodeX = tvw.Nodes.Add(nodeParent, tvwChild, , CStr(child))
        Set nodeX = tvw.Nodes.Add(nodeParent, tvwChild, , CStr(child))
        'recursively add children
        AddChildren tvw, nodeX, CStr(child) 'was tvw, nodeX, child
    Next
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AddChildren[Tree form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          FindSpecificNode
' Description:  find a node based on it's
' Assumptions:  -
' Parameters:   tvw - treeview control (treeview object)
'               strFind - item to find (string)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
'  das974, May 6, 2008
'  http://www.vbforums.com/showthread.php?509289-2008-Select-Treeview-node-by-name-or-key
' Adapted:      Bonnie Campbell, July 29, 2015 - for NCPN tools
' Revisions:
'   BLC - 7/29/2015 - initial version
' ---------------------------------
Public Sub FindSpecificNode(ByVal tvw As MSComctlLib.Treeview, strFind As String)
'Private Sub FindSpecificNode(ByVal tvw As MSComctlLib.Treeview, strFind As String)
'    Dim i As Integer
'    'Dim nodes As TreeNode, node As TreeNode
'    Dim nodes As Variant
'    Dim node As mscomctllib.node
'
'    Set nodes = tvw.nodes.Find(strFind, True) '"<selected word>",True)
'
'    'iterate through nodes
'    For Each node In nodes
'        tvw.Focus
'        tvw.SelectedNode = node
'    Next
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (" & Err.Number & " - FindSpecificNode[Tree form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     IsDuplicateKey
' Description:  determine if a node already exists for a given key
' Assumptions:  -
' Parameters:   strKey - key to check (string)
'               tvw - treeview control (treeview object)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, July 29, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/29/2015 - initial version
' ---------------------------------
Public Function IsDuplicateKey(strKey As String, tvw As MSComctlLib.Treeview) As Boolean
On Error GoTo Err_Handler

    Dim tvwNode As node
    Dim item As Variant
    Dim blnIsDupe As Boolean
    
    blnIsDupe = False
    
'    For Each tvwNode In tvwTree.Object.Nodes 'Me.TreeView.nodes
    For Each tvwNode In tvw.Nodes 'Me.TreeView.nodes
    
        If tvwNode.key = strKey Then
           blnIsDupe = True
           Exit For
        End If
    
    Next
    
Exit_Handler:
    IsDuplicateKey = blnIsDupe
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (" & Err.Number & " - IsDuplicateKey[Tree form])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     ImmovableNode
' Description:  indicate if a node can or cannot be moved
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, July 27, 2015 - for NCPN tools
' Revisions:
'   BLC - 7/27/2015 - initial version
'   BLC - 9/1/2016  - added unclassified photo type
' ---------------------------------
Public Function ImmovableNode(node As node) As Boolean

On Error GoTo Err_Handler
    
        'default
        ImmovableNode = False
        
        If CountInString(PHOTO_TYPES_MAIN, node) + CountInString(PHOTO_TYPES_OTHER, node) _
           + CountInString("Unclassified", node) > 0 Then
'            Debug.Print node
            ImmovableNode = True
        End If

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ImmovableNode[Tree form])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          tvwNodeSelect
' Description:  set view as if nodes are selected
' Assumptions:  -
' Parameters:   Node being selected (node object)
' Returns:      -
' Throws:       none
' References:
'  asp1n, March 6, 2004
'  http://www.xtremevbtalk.com/showthread.php?t=133762
' Source/date:  Bonnie Campbell, July 10, 2015 - for NCPN tools
' Adapted:
' Revisions:
'   BLC - 7/27/2015 - initial version
' ---------------------------------
Private Sub tvwNodeSelect(Optional node As node, Optional blnNodeSelected As Boolean)
On Error GoTo Err_Handler
    Dim i As Long
    Dim SelectedNode As node
    Dim colTreeNodes As Collection
    
    If blnNodeSelected Then
        If node.BackColor = vbHighlight Then
            If colTreeNodes.Count > 1 Then
                node.BackColor = vbWindowBackground
                node.ForeColor = vbWindowText
                node.Selected = False
                colTreeNodes.Remove node.key
            End If
            Exit Sub
        End If
    Else
        For i = 0 To colTreeNodes.Count - 1
            Set SelectedNode = colTreeNodes.item(i) 'colTreeNodes.Remove(, 0)
            SelectedNode.BackColor = vbWindowBackground
            SelectedNode.ForeColor = vbWindowText
            colTreeNodes.Remove i
        Next i
    End If
    
    If Not node Is Nothing Then
        node.BackColor = vbHighlight
        node.ForeColor = vbHighlightText
        colTreeNodes.Add node, node.key
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tvwNodeSelect[Tree form])"
    End Select
    Resume Exit_Handler
End Sub