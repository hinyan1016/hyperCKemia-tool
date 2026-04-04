' PPTX → PDF 変換スクリプト
' 使い方: cscript //nologo convert_to_pdf.vbs "C:\full\path\to\slides.pptx"
' 出力: 同じフォルダに slides.pdf を生成

Option Explicit

Dim objPPT, objPres, strInput, strOutput, objFSO

Set objFSO = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count < 1 Then
    WScript.Echo "使い方: cscript //nologo convert_to_pdf.vbs <pptxファイルのフルパス>"
    WScript.Quit 1
End If

strInput = WScript.Arguments(0)

' フルパスに変換
If Not objFSO.FileExists(strInput) Then
    strInput = objFSO.GetAbsolutePathName(strInput)
End If

If Not objFSO.FileExists(strInput) Then
    WScript.Echo "エラー: ファイルが見つかりません - " & strInput
    WScript.Quit 1
End If

' 出力パス (.pptx → .pdf)
strOutput = objFSO.GetParentFolderName(strInput) & "\" & _
            objFSO.GetBaseName(strInput) & ".pdf"

WScript.Echo "変換中: " & strInput

Set objPPT = CreateObject("PowerPoint.Application")
objPPT.Visible = True

Set objPres = objPPT.Presentations.Open(strInput, True, False, False)

' ppSaveAsPDF = 32
objPres.SaveAs strOutput, 32

objPres.Close
objPPT.Quit

Set objPres = Nothing
Set objPPT = Nothing

WScript.Echo "完了: " & strOutput
