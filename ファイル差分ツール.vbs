Option Explicit
'On Error Resume Next
Dim w,ws,os,fs,objPB
Dim proname,useProgress
Dim folder1,folder2,strF1path,strF2path,outputf,opf
Dim subName(),subPath(),subSize()
Dim Flengh,StartNow,StartTime
Dim minName,minPath,minSize,minKPath,subKPath
Dim diffNName,diffNPath,diffNSize,dn
Dim diffPName,diffPPath,diffPSize
Dim diffSName,diffSPath,diffSSize
Dim minDPath,subDPath,dp,tdp
Dim diffDPath,diffDP,LenDPath,ldp
Dim LenName,ln,LenPath,lp,LenSize,ls
Dim doPath,doSize
Dim NegWordFo,NegWordFi
Dim excludeFo,excludeFi
Dim excCouFo,excCouFi,hidCou
Dim HiddenFo,HiddenFi
Dim HiddenFos()
Dim CopyFo,CopyFi
Dim strFolder,strFile
Dim copyCouFo,copyCouFi,skipCou
Dim popAns,ubErFr
Dim negFlag
Dim ansSize,proWait
Dim i

' #####################################
'				�ύX�\�ӏ�
'// ��O����
' ���O������(�t�H���_)[","��؂�]
NegWordFo = ""
' ���O������(�t�@�C��)[","��؂�]
NegWordFi = ""

' �B���t�H���_�����O����{�ꍇ�ɂ�萳��ɓ��삵�Ȃ��\��������܂�}
HiddenFo = False
' �B���t�@�C�������O����
HiddenFi = False

'/* �t�H���_�R�s�[
'* 0:�R�s�[���������Ȃ�	{����}
'* 1:�t�H���_1�ɃR�s�[
'* 2:�t�H���_2�ɃR�s�[
'* 3:�����ɃR�s�[
'*/
CopyFo = 0
'/* �t�@�C���R�s�[
'* 0:�R�s�[���������Ȃ�	{����}
'* 1:�t�H���_1�ɃR�s�[
'* 2:�t�H���_2�ɃR�s�[
'* 3:�����ɃR�s�[
'*/
CopyFi = 0


'//���쏈��
' �i�s�x�\��
useProgress = True

' �t�H���_�p�X����
doPath = True
' �t�@�C���T�C�Y����
doSize = True

' #####################################

Set w=WScript
Set ws=w.CreateObject("WScript.Shell")
Set os=w.CreateObject("Shell.Application")
Set fs=createObject("Scripting.FileSystemObject")

StartNow = Now
StartTime = Timer
ansSize = True
copyCouFo = 0
copyCouFi = 0
skipCou = 0

if(not fs.FileExists("FProgressBar.vbs")) then
	useProgress = False
end if
if (useProgress) then _
	Execute fs.OpenTextFile("FProgressBar.vbs", 1).ReadAll()

proname = Left(w.ScriptName,Len(w.ScriptName)-4)
outputf = fs.getParentFolderName(WScript.ScriptFullName) & "\outputs.log"

Set folder1 = os.BrowseForFolder(0, "�t�H���_1��I�����ĉ�����", &H10, 17)
if (folder1 is nothing) then
	ws.Popup "�t�H���_���I������Ă��܂���I",60,proname,16
	w.Quit
end if
strF1path = folder1.Items.Item.Path
Set folder2 = os.BrowseForFolder(0, "�t�H���_2��I�����ĉ�����", &H10, 17)
if (folder2 is nothing) then
	ws.Popup "�t�H���_���I������Ă��܂���I",60,proname,16
	w.Quit
end if
strF2path = folder2.Items.Item.Path

if (strF1path = strF2path) then
	ws.Popup "����̃t�H���_��I�����Ă��܂��I",60,proname,16
	w.Quit
end if

if (HiddenFo) then
	popAns = ws.Popup("�B���t�H���_���O�@�\��ON�ɐݒ肳��Ă��܂�"&vbCrLf& _
		"���̋@�\�͏ꍇ�ɂ�萳��ɓ��삵�Ȃ��\��������܂���"&vbCrLf& _
		"���s���܂����H",60,proname,48+1)
	if (popAns = 2) then
		w.Quit
	end if
end if
if (CopyFo Or CopyFi) then
	popAns = ws.Popup("�t�H���_(�t�@�C��)�R�s�[�@�\��ON�ɐݒ肳��Ă��܂�"&vbCrLf& _
		"���̋@�\�͏ꍇ�ɂ�萳��ɓ��삵�Ȃ��\��������܂���"&vbCrLf& _
		"���s���܂����H",60,proname,48+1)
	if (popAns = 2) then
		w.Quit
	end if
end if

'//��O�����ݒ�
excludeFo = Split(NegWordFo,",")
excludeFi = Split(NegWordFi,",")
excCouFo = UBound(excludeFo)
excCouFi = UBound(excludeFi)

ReDim HiddenFos(0)
HiddenFos(0) = "#"


'//�h���C�u���ڎw��Ή�
if (Right(strF1path,1) = "\") then
	strF1path = Left(strF1path,Len(strF1path)-1)
end if
if (Right(strF2path,1) = "\") then
	strF2path = Left(strF2path,Len(strF2path)-1)
end if



'//�������烁�C��

if (useProgress) then
	Set objPB = new ProgressBar
	objPB.SetTitle "�t�@�C���ǂݍ��ݒ�..."
	objPB.SetProgress 0
end if

ubErFr = True
call FindFolder(fs.GetFolder(strF1path))
minName = subName
minPath = subPath
minSize = subSize
Erase subName
Erase subPath
Erase subSize


if (ubErFr) then
	call push(subName,"")
	call push(subPath,"")
	call push(subSize,0)
end if
minKPath = minPath
Flengh = Len(strF1path)+2
dn = UBound(minKPath)
tdp = ""
proWait = 0
ubErFr = True
for i=0 to dn
	dp = Mid(minKPath(i),Flengh)
	minKPath(i) = dp &"\"& minName(i)
	if(doPath) then
		if (not tdp = dp and not dp = "") then
			if (useProgress) then
				if (proWait > 10) then
					objPB.SetProgress i/dn/4+0.25
					proWait=0
				end if
				proWait = proWait + 1
			end if
			call push(minDPath,dp)
			ubErFr = False
		end if
		tdp = dp
	end if
next
if (ubErFr) then
	call push(minDPath,"")
end if

if (useProgress) then _
	objPB.SetProgress 0.5

ubErFr = True
call FindFolder(fs.GetFolder(strF2path))

if (ubErFr) then
	call push(subName,"")
	call push(subPath,"")
	call push(subSize,0)
end if
subKPath = subPath
Flengh = Len(strF2path)+2
dn = UBound(subKPath)
tdp = ""
ubErFr = True
for i=0 to dn
	dp = Mid(subKPath(i),Flengh)
	subKPath(i) = dp &"\"& subName(i)
	if(doPath) then
		if (not tdp = dp and not dp = "") then
			if (useProgress) then
				if (proWait > 10) then
					objPB.SetProgress i/dn/4+0.75
					proWait = 0
				end if
				proWait = proWait + 1
			end if
			call push(subDPath,dp)
			ubErFr = False
		end if
		tdp = dp
	end if
next
if (ubErFr) then
	call push(subDPath,"")
end if

if(doPath) then _
	call array_diff_path(minDPath,subDPath)

'//��O�t�H���_�ݒ�
hidCou = UBound(HiddenFos)

call array_diff_name(minKPath,subKPath,minSize,subSize)


'//�����X�^�[�g
Set opf = fs.OpenTextFile(outputf,2,True)
opf.WriteLine StartNow&" �����X�^�[�g"&vbCrLf
opf.WriteLine "�t�H���_1�p�X : "&strF1path
opf.WriteLine "�t�H���_2�p�X : "&strF2path &vbCrLf
if(doPath) then
	opf.WriteLine "�t�H���_1�T�u�t�H���_���v��:"&(UBound(minDPath)+1)
	opf.WriteLine "�t�H���_2�T�u�t�H���_���v��:"&(UBound(subDPath)+1)
	opf.WriteLine "�t�H���_���v��:"&(UBound(minDPath)+UBound(subDPath)+2)&vbCrLf
	opf.WriteLine "�f�B���N�g���\�������\�����[�h"&vbCrLf
	if (IsArray(diffDPath)<>False) then
		dn = UBound(diffDPath)
		opf.WriteLine "������:"&(dn+1)&vbCrLf
		if (useProgress) then _
			objPB.SetTitle "�t�H���_�\�������������ݒ�..."
		for i = 0 to dn
			if (useProgress) then
				if (proWait > 50) then
					objPB.SetProgress i/dn
					proWait = 0
				end if
				proWait = proWait + 1
			end if
			opf.WriteLine diffDPath(i)&Space(LenDPath-CountLen(diffDPath(i))+1)&"| "& _
				diffDP(i)
		next
	else
		opf.WriteLine "�S�Ĉ�v���Ă��܂��I"
	end if

	opf.WriteLine vbCrLf&String(LenName+LenSize+LenPath+13,"#")
	opf.WriteLine String(LenName+LenSize+LenPath+13,"#")&vbCrLf
end if

opf.WriteLine "�t�H���_1�t�@�C�����v��:"&(UBound(minName)+1)
opf.WriteLine "�t�H���_2�t�@�C�����v��:"&(UBound(subName)+1)
opf.WriteLine "�S�t�@�C�����v��:"&(UBound(minName)+UBound(subName)+2)&vbCrLf
opf.WriteLine "�t�@�C�����O�����\�����[�h"&vbCrLf
if (IsArray(diffNName)<>False) then
	dn = UBound(diffNName)
	if (dn+1 = UBound(minName)+UBound(subName)+2) then _
		ansSize = False
	opf.WriteLine "������:"&(dn+1)&vbCrLf
	if (useProgress) then _
		objPB.SetTitle "�t�@�C�����O�����������ݒ�..."
	for i = 0 to dn
		if (useProgress) then
				if (proWait > 50) then
					objPB.SetProgress i/dn
					proWait = 0
				end if
				proWait = proWait + 1
			end if
		opf.WriteLine diffNName(i)&Space(LenName-CountLen(diffNName(i))+1)&"| "& _
			diffNSize(i)&Space(LenSize-CountLen(diffNSize(i))+1)&"bytes "&vbTab&"| "& _
			diffNPath(i)
	next
else
	opf.WriteLine "�S�Ĉ�v���Ă��܂��I"
end if

if (doSize) then
	opf.WriteLine vbCrLf&String(LenName+LenSize+LenPath+13,"#")
	opf.WriteLine String(LenName+LenSize+LenPath+13,"#")
	opf.WriteLine vbCrLf&"�t�@�C���T�C�Y�����\�����[�h"&vbCrLf

	if (IsArray(diffSName)<>False) then
		dn = UBound(diffSName)
		opf.WriteLine "������:"&(dn+1)&vbCrLf
		if (useProgress) then _
			objPB.SetTitle "�t�@�C���T�C�Y�����������ݒ�..."
		for i = 0 to dn
			if (useProgress) then
				if (proWait > 50) then
					objPB.SetProgress i/dn
					proWait = 0
				end if
				proWait = proWait + 1
			end if
			opf.WriteLine diffSName(i)&Space(LenName-CountLen(diffSName(i))+1)&"| "& _
				diffSSize(i)&Space(LenSize-CountLen(diffSSize(i))+1)&"bytes "&vbTab&"| "& _
				diffSPath(i)
		next
	else
		if (ansSize) then
			opf.WriteLine "�S�Ĉ�v���Ă��܂��I"
		else
			opf.WriteLine "��r����t�@�C�������݂��܂���I"
		end if
	end if
end if

opf.WriteLine vbCrLf&Now&" �����͐���ɏI�����܂����I"
opf.WriteLine "�������� : "&(Timer-StartTime)&"s"
opf.Close
Set opf = Nothing

Set objPB = Nothing
ws.Popup "�I�����܂����I",60,proname,0

if (copyCouFo > 0) then
	if (copyCouFi > 0) then
		ws.Popup "�t�H���_��"&copyCouFo&"��"&vbCrLf& _
			"�t�@�C����"&copyCouFi&"��"&vbCrLf& _
			"�R�s�[���܂���",60,proname,64
	else
		ws.Popup "�t�H���_��"&copyCouFo&"��"&vbCrLf& _
			"�R�s�[���܂���",60,proname,64
	end if
else
	if (copyCouFi > 0) then
		ws.Popup "�t�@�C����"&copyCouFi&"��"&vbCrLf& _
			"�R�s�[���܂���",60,proname,64
	else
		if (CopyFo Or CopyFi) then
			ws.Popup "�t�H���_(�t�@�C��)�̃R�s�[�͎��s����܂���ł���",60,proname,64
		end if
	end if
end if
if (skipCou > 0) then
	ws.Popup skipCou&"���̃R�s�[���X�L�b�v����܂���",60,proname,48
end if

Erase subName
Erase subPath
Erase subSize

'//�I���n�_


Sub array_diff_path(ByVal mAryP1,ByVal mAryP2)
	Dim i,j
	Dim mi,mj
	mi = UBound(mAryP1)
	mj = UBound(mAryP2)

	ReDim mTmp(mi)

	if (useProgress) then _
		objPB.SetTitle "�t�H���_�\���v�Z��..."
	for i=0 to mi
		if (useProgress) then
			if (proWait > 30) then
				objPB.SetProgress i/mi
				proWait = 0
			end if
			proWait = proWait + 1
		end if
		for j=0 to mj
			if (StrComp(mAryP1(i),mAryP2(j),0)=0) then
				mAryP1(i)="@"
				mAryP2(j)="@"
				exit for
			end if
		next
	next

	ldp=0
	LenDPath=0

	if (useProgress) then _
		objPB.SetTitle "�t�H���_�\�������\���v�Z(1)��..."
	for i=0 to mi
		if (not mAryP1(i)="@") then
			if (useProgress) then
				if (proWait > 30) then
					objPB.SetProgress i/mi
					proWait = 0
				end if
				proWait = proWait + 1
			end if
			negFlag = True
			for j=0 to excCouFo
				if (InStr(minDPath(i),excludeFo(j))>0) then
					negFlag = False
					exit for
				end if
			next
			if (negFlag And HiddenFo) then
				if((fs.GetFolder(strF1path&"\"&minDPath(i)).Attributes And 2) <> 0) then
					negFlag = False
					call push(HiddenFos,minDPath(i))
				end if
			end if
			if (negFlag) then
				call push(diffDPath,minDPath(i))
				ldp = CountLen(minDPath(i))
				if (LenDPath<ldp) then
					LenDPath = ldp
				end if
				call push(diffDP,strF1path&"\"&minDPath(i))
				if (CopyFo = 2 Or CopyFo = 3) then
					strFolder = strF2path&"\"&minDPath(i)
					call FoFiCreate(strFolder)
				end if
			end if
		else
			if (HiddenFo) then
				if((fs.GetFolder(strF1path&"\"&minDPath(i)).Attributes And 2) <> 0) then
					call push(HiddenFos,minDPath(i))
				end if
			end if
		end if
	next

	if (useProgress) then _
		objPB.SetTitle "�t�H���_�\�������\���v�Z(2)��..."
	for i=0 to mj
		if (not mAryP2(i)="@") then
			if (useProgress) then
				if (proWait > 30) then
					objPB.SetProgress i/mj
					proWait = 0
				end if
				proWait = proWait + 1
			end if
			negFlag = True
			for j=0 to excCouFo
				if (InStr(subDPath(i),excludeFo(j))>0) then
					negFlag = False
					exit for
				end if
			next
			if (negFlag And HiddenFo) then
				if((fs.GetFolder(strF2path&"\"&subDPath(i)).Attributes And 2) <> 0) then
					negFlag = False
					call push(HiddenFos,subDPath(i))
				end if
			end if
			if (negFlag) then
				call push(diffDPath,subDPath(i))
				ldp = CountLen(subDPath(i))
				if (LenDPath<ldp) then
					LenDPath = ldp
				end if
				call push(diffDP,strF2path&"\"&subDPath(i))
				if (CopyFo = 1 Or CopyFo = 3) then
					strFolder = strF2path&"\"&minDPath(i)
					call FoFiCreate(strFolder)
				end if
			end if
		end if
	next

End Sub

Sub array_diff_name(ByVal mAryN1,ByVal mAryN2,ByVal mAryS1,ByVal mAryS2)
	Dim i,j
	Dim mi,mj
	Dim arsi
	Dim mTmp(),mt
	Dim cmTmp(),csTmp()
	mi = UBound(mAryN1)
	mj = UBound(mAryN2)

	ReDim mTmp(mi)
	ReDim cmTmp(mi)
	ReDim csTmp(mj)

	if (useProgress) then _
		objPB.SetTitle "�t�@�C�������v�Z��..."

	for i=0 to mj
		csTmp(i) = mAryN2(i)
	next
	if (doSize) then
		for i=0 to mi
			arsi = True
			cmTmp(i) = mAryN1(i)
			if (useProgress) then
				if (proWait > 30) then
					objPB.SetProgress i/mi
					proWait = 0
				end if
				proWait = proWait + 1
			end if
			for j=0 to mj
				if (StrComp(mAryN1(i),mAryN2(j),0)=0) then
					mAryN1(i)="@"
					mAryN2(j)="@"
					if(StrComp(mAryS1(i),mAryS2(j),0)=0) then
						mAryS1(i)="@"
					else
						mTmp(i) = j
					end if
					arsi = False
					exit for
				end if
			next
			if(arsi) then
				mAryS1(i)="@"
			end if
		next
	else
		for i=0 to mi
			cmTmp(i) = mAryN1(i)
			if (useProgress) then
				if (proWait > 30) then
					objPB.SetProgress i/mi
					proWait = 0
				end if
				proWait = proWait + 1
			end if
			for j=0 to mj
				if (StrComp(mAryN1(i),mAryN2(j),0)=0) then
					mAryN1(i)="@"
					mAryN2(j)="@"
					exit for
				end if
			next
		next
	end if

	ln=0
	ls=0
	LenName=0
	LenSize=0

	if (useProgress) then _
		objPB.SetTitle "�t�@�C�����O�����\���v�Z(1)��..."
	for i=0 to mi
		if (not mAryN1(i)="@") then
			if (useProgress) then
				if (proWait > 30) then
					objPB.SetProgress i/mi
					proWait = 0
				end if
				proWait = proWait + 1
			end if
			negFlag = True
			for j=0 to excCouFo
				if (InStr(minPath(i),excludeFo(j))>0) then
					negFlag = False
					exit for
				end if
			next
			if (negFlag) then
				for j=0 to excCouFi
					if (InStr(minName(i),excludeFi(j))>0) then
						negFlag = False
						exit for
					end if
				next
			end if
			if (negFlag And HiddenFi) then
				if((fs.GetFile(minPath(i) & "\" & minName(i)).Attributes And 2) <> 0) then
					negFlag = False
				end if
			end if
			if (negFlag And HiddenFo) then
				for j=1 to hidCou
					if (InStr(minPath(i),HiddenFos(j))>0) then
						negFlag = False
						exit for
					end if
				next
			end if
			if (negFlag) then
				call push(diffNName,minName(i))
				ln = CountLen(minName(i))
				if (LenName<ln) then
					LenName = ln
				end if
				call push(diffNPath,minPath(i))
				lp = CountLen(minSize(i))
				if (LenPath<lp) then
					LenPath = lp
				end if
				call push(diffNSize,minSize(i))
				ls = CountLen(minSize(i))
				if (LenSize<ls) then
					LenSize = ls
				end if
				if (CopyFi = 2 Or CopyFi = 3) then
					strFile = strF2path&"\"&cmTmp(i)
					if (fs.FolderExists(Left(strFile,InstrRev(strFile,"\")-1))) then
						if (not fs.FileExists(strFile)) then
							call fs.CopyFile(strF1path&"\"&cmTmp(i), strFile)
							copyCouFi = copyCouFi + 1
						end if
						if (Err.Number <> 0) then
							skipCou = skipCou + 1
							copyCouFi = copyCouFi - 1
							Err.Clear
						end if
					else
						skipCou = skipCou + 1
					end if
				end if
			end if
		end if
	next

	if (useProgress) then _
		objPB.SetTitle "�t�@�C�����O�����\���v�Z(2)��..."
	for i=0 to mj
		if (not mAryN2(i)="@") then
			if (useProgress) then
				if (proWait > 30) then
					objPB.SetProgress i/mj
					proWait = 0
				end if
				proWait = proWait + 1
			end if
			negFlag = True
			for j=0 to excCouFo
				if (InStr(subPath(i),excludeFo(j))>0) then
					negFlag = False
					exit for
				end if
			next
			if (negFlag) then
				for j=0 to excCouFi
					if (InStr(subName(i),excludeFi(j))>0) then
						negFlag = False
						exit for
					end if
				next
			end if
			if (negFlag And HiddenFi) then
				if((fs.GetFile(subPath(i) & "\" & subName(i)).Attributes And 2) <> 0) then
					negFlag = False
				end if
			end if
			if (negFlag And HiddenFo) then
				for j=1 to hidCou
					if (InStr(subPath(i),HiddenFos(j))>0) then
						negFlag = False
						exit for
					end if
				next
			end if
			if (negFlag) then
				call push(diffNName,subName(i))
				ln = CountLen(subName(i))
				if (LenName<ln) then
					LenName = ln
				end if
				call push(diffNPath,subPath(i))
				lp = CountLen(subPath(i))
				if (LenPath<lp) then
					LenPath = lp
				end if
				call push(diffNSize,subSize(i))
				ls = CountLen(subSize(i))
				if (LenSize<ls) then
					LenSize = ls
				end if
				if (CopyFi = 1 Or CopyFi = 3) then
					strFile = strF1path&"\"&csTmp(i)
					if (fs.FolderExists(Left(strFile,InstrRev(strFile,"\")-1))) then
						if (not fs.FileExists(strFile)) then
							call fs.CopyFile(strF2path&"\"&csTmp(i), strFile)
							copyCouFi = copyCouFi + 1
						end if
						if (Err.Number <> 0) then
							skipCou = skipCou + 1
							copyCouFi = copyCouFi - 1
							Err.Clear
						end if
					else
						skipCou = skipCou + 1
					end if
				end if
			end if
		end if
	next

	if (doSize) then
		if (useProgress) then _
			objPB.SetTitle "�t�@�C���T�C�Y�����\���v�Z��..."
		for i=0 to mi
			if (not mAryS1(i)="@") then
				if (useProgress) then
					if (proWait > 30) then
						objPB.SetProgress i/mi
						proWait = 0
					end if
					proWait = proWait + 1
				end if

				negFlag = True
				for j=0 to excCouFo
					if (InStr(minPath(i),excludeFo(j))>0) then
						negFlag = False
						exit for
					end if
				next
				if (negFlag) then
					for j=0 to excCouFi
						if (InStr(minName(i),excludeFi(j))>0) then
							negFlag = False
							exit for
						end if
					next
				end if
				if (negFlag And HiddenFi) then
					if((fs.GetFile(minPath(i) & "\" & minName(i)).Attributes And 2) <> 0) then
						negFlag = False
					end if
				end if
				if (negFlag And HiddenFi) then
					if((fs.GetFile(subPath(i) & "\" & subName(i)).Attributes And 2) <> 0) then
						negFlag = False
					end if
				end if
				if (negFlag And HiddenFo) then
					for j=1 to hidCou
						if (InStr(minPath(i),HiddenFos(j))>0) then
							negFlag = False
							exit for
						end if
					next
				end if
				if (negFlag) then
					mt = mTmp(i)
					call push(diffSName,minName(i))
					ln = CountLen(minName(i))
					if (LenName<ln) then
						LenName = ln
					end if
					call push(diffSName,subName(mt))
					ln = CountLen(subName(mt))
					if (LenName<ln) then
						LenName = ln
					end if

					call push(diffSPath,minPath(i))
					lp = CountLen(minPath(i))
					if (LenPath<lp) then
						LenPath = lp
					end if
					call push(diffSPath,subPath(mt))
					lp = CountLen(subPath(mt))
					if (LenPath<lp) then
						LenPath = lp
					end if

					call push(diffSSize,minSize(i))
					ls = CountLen(minSize(i))
					if (LenSize<ls) then
						LenSize = ls
					end if
					call push(diffSSize,subSize(mt))
					ls = CountLen(subSize(mt))
					if (LenSize<ls) then
						LenSize = ls
					end if
				end if
			end if
		next
	end if
End Sub

Sub FoFiCreate(ByVal Paths)
	if(fs.FolderExists(Left(Paths,InstrRev(Paths,"\")-1))) then
		if (not fs.FolderExists(Paths)) then
			fs.CreateFolder(Paths)
			copyCouFo = copyCouFo + 1
		end if
		if (Err.Number <> 0) then
			skipCou = skipCou + 1
			copyCouFo = copyCouFo - 1
			Err.Clear
		end if
	else
		call FoFiCreate(Left(Paths,InstrRev(Paths,"\")-1))
		if(fs.FolderExists(Left(Paths,InstrRev(Paths,"\")-1))) then
			if (not fs.FolderExists(Paths)) then
				fs.CreateFolder(Paths)
				copyCouFo = copyCouFo + 1
			end if
			if (Err.Number <> 0) then
				skipCou = skipCou + 1
				copyCouFo = copyCouFo - 1
				Err.Clear
			end if
		end if
	end if
End Sub

Sub FindFolder(ByVal objMainFolder)
	Dim objSubFolder,objFile

	'// �t�H���_������΍ċA
	for each objSubFolder in objMainFolder.SubFolders
		FindFolder objSubFolder
	next

	'// �t�H���_�̒��̃t�@�C������\��
	for each objFile in objMainFolder.files
		call push(subName,objFile.Name)
		call push(subPath,objFile.ParentFolder)
		call push(subSize,objFile.Size)
		ubErFr = False
	next
End Sub

Sub push(ByRef arr,ByVal elm)
	Dim i,tmp : i = 0
	if IsArray(arr) then
		for each tmp in arr
			i = 1
			exit for
		next
		if i=1 then
			ReDim Preserve arr(UBound(arr)+1)
		else
			ReDim arr(0)
		end if
	else
		arr = Array(0)
	end if
	if IsObject(elm) then
		Set arr(UBound(arr)) = elm
	else
		arr(UBound(arr)) = elm
	end If
End Sub

Function CountLen(ByVal data)
	Dim i,c,counter
	counter = 0
	for i = 1 to Len(data)
		c = asc(mid(data, i, 1))
		if c >= &H00 and c <= &H7E then
			counter = counter + 1
		else
			counter = counter + 2
		end if
	next
	CountLen = counter
End Function