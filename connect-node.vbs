'''On   Error   Resume   Next
wscript.echo  vbcrlf + vbcrlf + vbcrlf
wscript.echo "**************************************************************************************"
wscript.echo "本工具旨在免费提供给大家更多更好的Chia优质更新节点，好的节点离不开大家的提供。 请多多支持。"
wscript.echo "更多优质节点请提交给 https://github.com/chiaprofessor/chianode "
wscript.echo "程序运行时切勿关闭此窗口，每一个小时同步一次节点，保证客户端24小时不掉线。"
wscript.echo "**************************************************************************************"
wscript.echo  vbcrlf + vbcrlf + vbcrlf

''' 1. 获取chia目录官方程序路径

chia_exe = ""    ''' 默认自动查找Chia文件，也可以根据自己电脑自定义，比如:C:\Users\Talent\AppData\Local\chia-blockchain\app-1.1.5\resources\app.asar.unpacked\daemon\chia.exe

Set myws = CreateObject("WScript.Shell")
Set fso=CreateObject("Scripting.FileSystemObject") 

If chia_exe = ""  Then
	chia_exe = GET_CHIA_REG()
	If fso.fileExists(chia_exe) Then
		wscript.echo "已通过注册表找到Chia客户端：" + vbcrlf + chia_exe
	Else
		chia_exe = GET_CHIA_APPDATA()
			If fso.fileExists(chia_exe) Then
				wscript.echo "已通过用户目录遍历方式找到Chia客户端：" + vbcrlf + chia_exe
			Else
				wscript.echo "未找到Chia客户端，请手动修改chia_exe参数进行配置。"
			End If
	End If
	
	
Else
	wscript.echo  "已使用自定义Chia路径: " + vbcrlf + chia_exe

End If
wscript.echo  vbcrlf + vbcrlf

''' 2. 获取当前的活跃节点列表

Dim Nodes
wscript.echo vbcrlf + "开始寻找活跃更新节点..." 
wscript.echo  vbcrlf + vbcrlf + vbcrlf
Nodes = NodeList("http://158.247.225.94/node/")
'''wscript.echo Nodes
Nodearr=split(Nodes,vbCrLf)
Nodenum = Ubound(Nodearr)
wscript.echo vbcrlf + "已找到 " +cstr(Nodenum)+" 个高速节点."
wscript.echo  vbcrlf + vbcrlf + vbcrlf




''' 3. 每一个小时循环连接节点

do

wscript.echo vbcrlf + "开始连接高速节点..." +vbcrlf
	For i=0 to Nodenum-1
		wscript.echo "开始连接第 " + cstr(cint(i+1)) +" 个节点 :  " +vbcrlf
		Set nodestart = CreateObject("WScript.Shell")
		chia_dir=fso.GetParentFolderName(chia_exe)
		nodestart.CurrentDirectory = chia_dir
		Set nodeExec = nodestart.Exec("%COMSPEC% /C """ + chia_exe + """ show -a " + Nodearr(i))
		Do While Not nodeExec.StdOut.AtEndOfStream
			strText = nodeExec.StdOut.ReadAll()
		Loop

	Next
	wscript.echo vbcrlf +  vbcrlf + vbcrlf
	wscript.echo "更新完毕，勿关闭窗口 。 下个小时自动循环更新。"
wscript.sleep 3600000
loop








Function GET_CHIA_REG()
	chiaInstallLocation = CreateObject("Wscript.Shell").RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\chia-blockchain\InstallLocation")
	chiaversion = CreateObject("Wscript.Shell").RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\chia-blockchain\DisplayVersion")
	chia_exe = chiaInstallLocation + "\app-"+chiaversion+"\resources\app.asar.unpacked\daemon\chia.exe"
	GET_CHIA_REG = chia_exe
ENd Function


Function GET_CHIA_APPDATA()
	user_appdata = myws.ExpandEnvironmentStrings("%LOCALAPPDATA%")
	chia_dir = user_appdata + "\chia-blockchain\"
	
	Set fs = fso.GetFolder(chia_dir) 
	Set df = fs.SubFolders
	
	For Each d In df

		chia_exe = d + "\resources\app.asar.unpacked\daemon\chia.exe"
		If fso.fileExists(chia_exe) Then         
			GET_CHIA_APPDATA = chia_exe  
			Exit Function			
		End If
	Next
ENd Function





Function NodeList(url)
    Dim xmlHttp
    Set xmlHttp = CreateObject("Microsoft.XMLHTTP")
    xmlHttp.open "get", url, False
    xmlHttp.send
    Do
    Loop Until xmlHttp.readyState = 4
    NodeList = xmlHttp.responseText
    Set xmlHttp = Nothing
End Function



