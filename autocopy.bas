Attribute VB_Name = "Module1"
Sub OpenLinkAndCopyAddress()
Attribute OpenLinkAndCopyAddress.VB_ProcData.VB_Invoke_Func = "d\n14"
    Dim currentRow As Long
    Dim linkURL As String
    Dim copyText As String
    Dim clipBoardObj As Object
    
    Set clipBoardObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
    currentRow = ActiveCell.Row
    
    'Column mapping matched your table
    Const LINK_COL = 6       'Column F: AMZ link
    Const NAME_COL = 9       'Column I: recipient name
    Const PHONE_COL = 10     'Column J: phone number
    Const ADDR1_COL = 11     'Column K: ship address1
    Const ADDR2_COL = 12     'Column L: ship address2
    Const CITY_COL = 15      'Column O: ship city
    Const STATE_COL = 16     'Column P: ship state
    Const ZIP_COL = 17       'Column Q: ship postal code
    
    linkURL = Cells(currentRow, LINK_COL).Value
    
    'Combine address data with tab separator
    copyText = Cells(currentRow, NAME_COL).Value & vbTab & _
               Cells(currentRow, PHONE_COL).Value & vbTab & _
               Cells(currentRow, ADDR1_COL).Value & vbTab & _
               Cells(currentRow, ADDR2_COL).Value & vbTab & _
               Cells(currentRow, CITY_COL).Value & vbTab & _
               Cells(currentRow, STATE_COL).Value & vbTab & _
               Cells(currentRow, ZIP_COL).Value
    
    If Trim(linkURL) = "" Then
        MsgBox "No product link on selected row!", vbExclamation
        Exit Sub
    End If
    
    'Copy address to clipboard silently
    clipBoardObj.SetText copyText
    clipBoardObj.PutInClipboard
    
    'Open Amazon link directly without any prompt
    ThisWorkbook.FollowHyperlink linkURL
End Sub

