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
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, July 10, 2015 - for NCPN tools
' Revisions:
'   BLC - 7/10/2015 - initial version
' ---------------------------------
Private Sub LoadTree()

On Error GoTo Err_Handler

'LoadTree tvwTree '(tvw)
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - LoadTree[Tree form])"
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
Private Sub AddChildren(tvw As Treeview, nodeParent As node, aryKids As String)
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
Private Sub FindSpecificNode(ByVal tvw As MSComctlLib.Treeview, strFind As String)
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