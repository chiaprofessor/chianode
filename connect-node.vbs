On   Error   Resume   Next
wscript.echo  vbcrlf + vbcrlf + vbcrlf
wscript.echo "**************************************************************************************"
wscript.echo "������ּ������ṩ����Ҹ�����õ�Chia���ʸ��½ڵ㣬�õĽڵ��벻����ҵ��ṩ�� ����֧�֡�"
wscript.echo "�������ʽڵ����ύ�� https://github.com/chiaprofessor/chianode "
wscript.echo "��������ʱ����رմ˴��ڣ�ÿһ��Сʱͬ��һ�νڵ㣬��֤�ͻ���24Сʱ�����ߡ�"
wscript.echo "**************************************************************************************"
wscript.echo  vbcrlf + vbcrlf + vbcrlf

''' 1. ��ȡchiaĿ¼�ٷ�����·��

chia_exe = ""    ''' Ĭ���Զ�����Chia�ļ���Ҳ���Ը����Լ������Զ��壬����:C:\Users\Talent\AppData\Local\chia-blockchain\app-1.1.5\resources\app.asar.unpacked\daemon\chia.exe

Set myws = CreateObject("WScript.Shell")
Set fso=CreateObject("Scripting.FileSystemObject") 

If chia_exe = ""  Then
	chia_exe = GET_CHIA_REG()
	If fso.fileExists(chia_exe) Then
		wscript.echo "��ͨ��ע����ҵ�Chia�ͻ��ˣ�" + vbcrlf + chia_exe
	Else
		chia_exe = GET_CHIA_APPDATA()
			If fso.fileExists(chia_exe) Then
				wscript.echo "��ͨ���û�Ŀ¼������ʽ�ҵ�Chia�ͻ��ˣ�" + vbcrlf + chia_exe
			Else
				wscript.echo "δ�ҵ�Chia�ͻ��ˣ����ֶ��޸�chia_exe�����������á�"
			End If
	End If
	
	
Else
	wscript.echo  "��ʹ���Զ���Chia·��: " + vbcrlf + chia_exe

End If
wscript.echo  vbcrlf + vbcrlf

''' 2. ��ȡ��ǰ�Ļ�Ծ�ڵ��б�

Dim Nodes
wscript.echo vbcrlf + "��ʼѰ�һ�Ծ���½ڵ�..." 
wscript.echo  vbcrlf + vbcrlf + vbcrlf
Nodes = NodeList("http://158.247.225.94/node/")
'''wscript.echo Nodes
Nodearr=split(Nodes,vbCrLf)
Nodenum = Ubound(Nodearr)
wscript.echo vbcrlf + "���ҵ� " +cstr(Nodenum)+" �����ٽڵ�."
wscript.echo  vbcrlf + vbcrlf + vbcrlf




''' 3. ÿһ��Сʱѭ�����ӽڵ�

do

wscript.echo vbcrlf + "��ʼ���Ӹ��ٽڵ�..." +vbcrlf
	For i=0 to Nodenum-1
		wscript.echo "��ʼ���ӵ� " + cstr(cint(i+1)) +" ���ڵ� :  " +vbcrlf
		Set nodestart = CreateObject("WScript.Shell")
		chia_dir=fso.GetParentFolderName(chia_exe)
		nodestart.CurrentDirectory = chia_dir
		Set nodeExec = nodestart.Exec("%COMSPEC% /C """ + chia_exe + """ show -a " + Nodearr(i))
		Do While Not nodeExec.StdOut.AtEndOfStream
			strText = nodeExec.StdOut.ReadAll()
		Loop

	Next
	wscript.echo vbcrlf +  vbcrlf + vbcrlf
	wscript.echo "������ϣ���رմ��� �� �¸�Сʱ�Զ�ѭ�����¡�"
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



