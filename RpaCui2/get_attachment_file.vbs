Dim olapp: Set olApp = CreateObject("Outlook.Application")

Const MAPI = "MAPI"
Dim src: src = WScript.Arguments(0)
Dim dst: dst = WScript.Arguments(1)
Dim bak: bak = WScript.Arguments(2)
 
Dim nmspace: Set nmspace = olapp.GetNamespace(MAPI)
Dim inbox:   Set inbox = nmspace.GetDefaultFolder(6)
Dim src_fld: Set src_fld = inbox.Folders.Item(src)
Dim bak_fld: Set bak_fld = inbox.Folders.Item(bak)
 
Dim i
Dim atc_fil
Dim cnt_f: cnt_f = 0





src_fld.Items.Sort ["ÆüÉÕ"], True

For Each itm In src_fld.Items
    If itm.Attachments.Count > 0 Then
        For i = 1 To itm.Attachments.Count
            atc_fil = dst + "\" + itm.Attachments.Item(i) 
            itm.Attachments.Item(i).SaveAsFile atc_fil
        Next
        cnt_f = cnt_f + 1
    End If
Next





Dim flg: flg = True
Dim cnt: cnt = 0
Do
    If src_fld.Items.Count = 0 Then
        Exit Do
    End If
    For Each itm In src_fld.Items
        itm.Move bak_fld
    Next
Loop Until flg = False



If cnt_f = 0 Then
    WScript.Quit(1)
ElseIf cnt_f > 1 Then
    WScript.Quit(2)
Else
    '---
End If
