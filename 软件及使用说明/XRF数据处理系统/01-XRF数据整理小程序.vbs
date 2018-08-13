Set oExcel = CreateObject("Excel.Application")
oExcel.Visible = True

Set ws=WScript.CreateObject("wscript.shell")
w=ws.CurrentDirectory
Set fso = CreateObject("Scripting.FileSystemObject")
Set oFolder = fso.GetFolder(w)     
Set oFiles = oFolder.Files 
For Each file In oFiles
    sExt = fso.GetExtensionName(file)    
    sExt = LCase(sExt)
    If sExt = "txt" Then
        name = left(file.name, InstrRev(file.name,".")-1)
        Set file=fso.opentextfile(name+".txt")
        s=file.readall
        file.close
        For i = 0 To 14
            s=replace(s,"  "," ")   '去除所有多余的空格
        Next
        s=replace(s,"Sample Identification","Sample_Identification")
        s=replace(s,"Date / Time","Date Time AMPM")
        s=replace(s,"Result Type","Result_Type")
        s=replace(s," (","(")
        's=replace(s,"1/1","1\1")
        s=replace(s,vbCrLf + " S", vbCrLf + ",,,,,S")   '如果不这样处理，这一行的标题将不能和下面的内容对齐
        s=replace(s,vbCrLf + " Time", vbCrLf + ",,,,,Time")
        s = replace(s, " Total", "Total")
        s = replace(s, " Report", "Report")
        Set re = New RegExp
        re.pattern = "\d+/.+\r\n.+\r\n\r\n.+\r\n.+\r\n\r\n.+\r\n"  '匹配上字符串的前几行
        re.Global = True
        Set colMatches = re.execute(s)
        s = re.replace(s,"")    '将前几行替换为空
        s=replace(s," ",",")    '此时已经删除掉前几行，可以放心地替换所有空格为逗号了，因为这一步替换会破坏前几行的格式
        'For Each match In colMatches
            ' msgbox match
        '    s = match + s   '将前几行再加入到字符串中
        'Next
        Set file=fso.createtextfile(name + ".csv")
        file.write s
        file.close    
    End If
Next
'MsgBox "单击确定后将开始生成xls文件，此过程耗时稍长，请耐心等候"
Set oWorkbooks = oExcel.Workbooks.Open(w+"\" + "XRFTransform.xlsm")
oExcel.Run "transform"
oWorkbooks.Close
'oExcel.Quit
Set oWorkbooks= Nothing
Set oExcel= Nothing

For Each file In oFiles
    sExt = fso.GetExtensionName(file)    
    sExt = LCase(sExt)
    If sExt = "csv" Then
        file.attributes=0
        file.delete
    end If
next
'MsgBox "格式转换完毕，请查看新生成的csv与xls文件"