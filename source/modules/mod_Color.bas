Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Color
' Level:        Framework module
' Version:      1.05
' Description:  color functions & procedures
'
' Source/date:  Bonnie Campbell, 2/9/2015
' Revisions:    BLC, 2/9/2015 - 1.00 - initial version
'               BLC, 5/1/2015 - 1.01 - integrated into Invasives Reporting tool
'               BLC, 11/10/2015 - 1.02 - added additional colors
'               BLC, 5/27/2016 - 1.03 - added additional colors
'               BLC, 6/4/2016  - 1.04 - added HTMLconvert()
'               BLC, 6/24/2016 - 1.05 - replaced Exit_Function > Exit_Handler
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------
' http://cloford.com/resources/colours/500col.htm
' vbGrayText            &H80000011  Grayed (disabled) text
' vbInactiveTitleBar    &H80000003  Color of the title bar for the inactive window
' Andy Pope, March 7, 2003
' http://www.ozgrid.com/forum/showthread.php?t=49072
' Microsoft
' https://msdn.microsoft.com/en-us/library/office/aa195896%28v=office.11%29.aspx
' http://rainbowprod.com/english/powerbuilder/colors.html
' long value = (65536*Blue) + (256*Green) + (Red)
' Anonymous, March 9, 1999
' http://www.vbcode.com/asp/showsn.asp?theID=311
' Convert RGB to LONG:      LONG = B * 65536 + G * 256 + R
' ---------------------------------
Public Const lngGray As Long = 8224125      '?RGB(125, 125, 125)
Public Const lngLtGray As Long = 13882323   '?RGB(211, 211, 211)
Public Const lngGray50 As Long = 8355711    '?RGB(127,127,127) Text 1, Lighter 50% #7F7F7F Gray50
Public Const lngLime As Long = 6750105      '?RGB(153, 255, 102) #99FF66
Public Const lngBlue As Long = 16711680     '?RGB(0, 0, 255) #0000FF
Public Const lngLtOrange As Long = 52479    '?RGB(255,204,0) #FFCC00
Public Const lngLtLime As Long = 6750156    '?RGB(204,255,102) #CCFF66
Public Const lngDkLime As Long = 52377      '?RGB(153,204,0) #99CC00
Public Const lngBrtLime As Long = 3407769   '?RGB(153,255,51) #99FF33
Public Const lngLtGreen As Long = 52224     '?RGB(0,204,0) #00CC00
Public Const lngDkGray As Long = 2375487    '?RGB(63,63,63) #3F3F3F
Public Const lngYelLime As Long = 9699294   '?RGB(222,255,147) #DEFF93
Public Const lngWhite As Long = 16777215    '?RGB(255,255,255) #FFFFFF
Public Const lngBlack As Long = 0           '?RGB(0,0,0) #000000
Public Const lngYellow As Long = 65535      '?RGB(255,255,0) #FFFF00
Public Const lngLtYellow As Long = 14745599 '?RGB(255,255,224) #FFFFE0
Public Const lngBrown As Long = 13107       '?RGB(51,51,0) #333300
Public Const lngPutty As Long = 3355443     '?RGB(51,51,51) #333333
Public Const lngPurple As Long = 9974127    '?RGB(111,49,152) #6F3198
Public Const lngLtBlue As Long = 16777164   '?RGB(204,255,255) #CCFFFF
Public Const lngRed As Long = 255           '?RGB(255,0,0) #FF0000
Public Const lngGreen As Long = 65280       '?RGB(0,255,0) #00FF00
Public Const lngDkGreen As Long = 690698    '?RGB(10,138,10) #0A8A0A
Public Const lngRobinEgg As Long = 16772541 '?RGB(189,237,255) #BDEDFF robin's egg blue
Public Const lngLtCyan As Long = 16777184   '?RGB(224,255,255) #E0FFFF
Public Const lngGrnApple As Long = 1557580  '?RGB(76,196,23) #4CC417
Public Const lngCoral As Long = 527564      '?RGB(255,127,80) #FF7F50
Public Const lngYelGrn As Long = 1560658    '?RGB(82,208,23) #52D017
Public Const lngPink As Long = 10582263     '?RGB(247,120,161) #F778A1 carnation red
Public Const lngOceanBlue As Long = 15492395 '?RGB(43, 101, 236) #2B65EC ocean blue
Public Const lngSedona As Long = 26316      '?RGB(204,102,0) #CC6600
Public Const lngCocoa As Long = 5334161     '?RGB(145, 100, 81) #916451 cocoa brown NPS arrowhead bgd
Public Const lngDkBlueGrn As Long = 4538399 '?RGB(31, 64, 69) #1f4045 dark blue-green NPS arrowhead trees & buffalo outline
Public Const lngCream As Long = 11262179    '?RGB(227, 216, 171) #e3d8ab cream NPS arrowhead mtn & lake
Public Const lngNPSBrown As Long = 2634567  '?RGB(71, 51, 40) #473328 NPS signs
Public Const lngVanilla As Long = 11265523  '?RGB(243, 229, 171) #F3E5AB
Public Const lngGold As Long = 121087       '?RGB(255, 216, 1) #FFD801
Public Const lngMimosa As Long = 16743326   '?RGB(158, 123, 255) #937BFF purple mimosa
Public Const lngMimosaComp As Long = 8060892 '?RGB(220, 255, 122) #DCFF7A
Public Const lngSageGreen As Long = 7965572 '?RGB(132, 139, 121) #848B79
Public Const lngLtSalmon As Long = 10998527 '?RGB(255, 210, 167) #FFD2A7
Public Const lngSalmon As Long = 7051001    '?RGB(249, 150, 107) #F9966B
Public Const lngLtSienna As Long = 3497896  '?RGB(168, 95, 53) #A85F35
Public Const lngSpringGreen As Long = 32768 '?RGB(0,128,0) #008000
Public Const lngTan As Long = 9221330       '?RGB(210,180,140) #D2B48C tan
Public Const lngBurlywood As Long = 8894686 '?RGB(222,184,135) burlywood
Public Const lngLtRose As Long = 11845354   '?RGB(234,190,180) #EABEB4
Public Const lngRose As Long = 11843306     '?RGB(234,182,180) #EAB6B4

' ---------------------------------
'  Subroutines & Functions
' ---------------------------------

' ---------------------------------
' SUB:          ConvertLongToRGB
' Description:  Convert long color value to RGB component values
' Assumptions:  User will call specific color values via dict("R"), dict("G"), dict("B") as needed
' Parameters:   lngValue - long color value
' Returns:      RGB - as dicitionary object (R, G, B components)
' Throws:       none
' References:   none
' Requires:     Microsoft Scriping Runtime, scrrun.dll reference for dictionary object
' Source/date:
' Chirag, March 12, 2008
' http://www.pcreview.co.uk/threads/convert-long-integer-color-value-to-rgb.3447564/
' JTolle, August 21, 2009
' http://stackoverflow.com/questions/1309689/hash-table-associative-array-in-vba
' Adapted:      Bonnie Campbell, March 4, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/4/2015  - initial version
'   BLC - 11/10/2015 - added reference requirements
' ---------------------------------
Public Function ConvertLongToRGB(ByVal lngValue As Long) As Dictionary
On Error GoTo Err_Handler
    Dim dRGB As Dictionary
    Set dRGB = New Dictionary
       
    dRGB("R") = lngValue Mod 256
    dRGB("G") = Int(lngValue / 256) Mod 256
    dRGB("B") = Int(lngValue / 256 / 256) Mod 256
    
    Set ConvertLongToRGB = dRGB
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ConvertLongToRGB[mod_Color])"
    End Select
    Resume Exit_Handler
End Function

' =================================
' FUNCTION:     HTMLConvert
' Description:  converts HTML string value for color to RGB which can be used for control colors
' Parameters:   strHTML - HTML color (make sure you include # otherwise the color won't match)
' Returns:      HTML color as long
' Throws:       none
' References:   none
' Source/date:  Adapted from http://www.access-programmers.co.uk/forums/showthread.php?t=193353
'               by Steve R., 5/21/2010.
'               Created 05/12/2014 blc; Last modified 05/12/2014 blc.
' Revisions:    BLC, 5/12/2014 - initial version
'               BLC, 4/30/2015 - moved from mod_Common_UI to mod_UI
'               BLC, 5/17/2015 - moved from mod_UI to mod_Color & added error handling
'               BLC, 5/18/2015 - renamed, removed fxn prefix
' =================================
Public Function HTMLConvert(strHTML As String) As Long
On Error GoTo Err_Handler
    
    Rem converts a HTML color code number such as #D8B190 to an RGB value.
    HTMLConvert = RGB(CInt("&H" & Mid(strHTML, 2, 2)), CInt("&H" & Mid(strHTML, 4, 2)), CInt("&H" & Mid(strHTML, 6, 2)))

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - HTMLConvert[mod_Color])"
    End Select
    Resume Exit_Handler
End Function