<!-- #include file="easp.asp" -->
<head>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<style>
.divscroll{ height:30%;margin-left:10px;overflow-y:scroll; overflow-x:scroll;} 
</style>
</head>
<%
'''��÷�ҳ��Ϣ
'page =  easp.r("page",1)
'easp.wn "page= " & page
'''ȫ�ֱ���
Dim dbfile: dbfile= "a2000.mdb"


'''��ñ����Ϣ
Dim action
action =  easp.R("action",0)

If Not easp.isN(action) then

	Select Case action
		Case "search":
		Dim search
		search = easp.CheckForm(Easp.R("search",0),"",1,"�������ݲ���Ϊ��!")
		
		Case "add":
		Dim uname, ucode
		uname =  easp.CheckForm(Easp.R("name",0),"",1,"��������Ϊ��!:������Ч���û���!")
		ucode = easp.CheckForm(Easp.R("code",0),"",1,"���/���벻��Ϊ��!")

		Case "del":
		Dim uid
		uid =  easp.CheckForm(Easp.R("delid",1),"number",1,"ɾ����Ÿ�ʽ����ȷ!")
	End select
End if
%>
<%
Dim pagename
pagename = "����ͨѶ¼"
easp.wn "<title>" & pagename & "</title>"
easp.wn "<a href='dia.asp'>" & pagename & "</a>"

Set Db = new EasyAsp.db
Db.dbconn = Db.openconn(1,dbfile,"")
Db.Debug = True

''''����

'''���������ģ������
easp.wn(">>>���������ģ������<<< ")
easp.wn "<form method=""post"" action=""dia.asp?action=search"">"
easp.w "����������������ȫ���򲿷ֵ���:<input type=""text"" name=""search"">"
easp.wn "<input type=""submit"" value=""����"">"
easp.wn "</form>"

If action = "search" then
	Dim rs
	Set rs = db.GetRecordDetail("cs","[_name] like '%" & search & "%'")

	easp.wn("�����б�")
	while Not (rs.eof or rs.bof)
		Easp.wn rs("_name") & "'s  code:" & rs("code")
		rs.movenext()
	Wend
	easp.wn " ������: " & (rs.recordcount) & " �� ��¼."
	easp.c(rs)
	easp.wn "========================"
End If

''''��ʾ�б�
easp.wn "---------------------"
easp.wn "����޸ĵ���ȫ�б�"
Dim pagelist
'Dim rs
Set rs = db.GetRecordBySQL("Select * from cs")

pagelist= "<div class='divscroll'><br><table><tr><td>�� ��</td><td>�� ��</td><td>�� ��</td></tr>"

while Not (rs.eof or rs.bof)

	pagelist = pagelist & "<tr><td>" & rs("_name") & "</td><td>code:" & rs("code") & "</td><td>" & "<a href=?action=del&delid="& rs(0) &">ɾ��</a></td></tr>"
	rs.movenext()
Wend
easp.wn(rs.recordcount) & " �� ��¼."
easp.c(rs)

easp.wn pagelist & "</table></div>" '''���html
''''���Ӽ�¼
If action = "add" Then
dim NewID
NewID = Db.Autoid("cs")

	db.ar "cs",array("_name:" & uname,"code:" & ucode,"sortId:" & NewID)
	easp.rr "dia.asp"
End If

easp.wn "----------------------"
easp.wn "������ϵ��:"
easp.wn "<form method=""post"" action=""dia.asp?action=add"">"
easp.wn "����:<input type=""text"" name=""name"">"
easp.wn "����:<input type=""text"" name=""code"">"
easp.wn "<input type=""submit"">"
easp.wn "</form>"

''''ɾ����¼
If action = "del" then
	
	easp.wn "ɾ����� #" & uid & " ����."
	db.DeleteRecord "cs"," id=" & uid
	easp.rr "dia.asp"
'db.DeleteRecord "cs",Array("_name:me")
'db.DeleteRecord "cs"," [_name] like '%me%' "
End If

''''�޸ļ�¼
'db.UpdateRecord "cs",Array("_name:sme"),Array("code:a12345")


''''��ҳ
easp.wn ""
easp.wn "��ҳ����"
'��ͬһ��ҳ���ϸ��ݼ�¼�����ɶ����ҳ������ͬʱ������¼��Ҳ���з�ҳ(Ƕ�׷�ҳ):   

   
Dim rsSon   
'�����ҳ��ÿҳ5����¼����ҳ��ʽΪarray����   
'''������������ Array ����1->��:����, 2->����, 3-> ��������
'''4-> ����

	'
	'db.GPR("array:page:5",
    Set rsSon = db.GPR("array:5",Array("cs","","sortid asc","sortid"))  

	'�ȶ���һ����ҳ������ʽ,��Ҫ������ grp ֮��   
	db.SetPager "default", "{first}{prev}{liststart}{list}{listend}{next}{last},��{jump}ҳ", Array("listlong:9","listsidelong:3","first:��ҳ","prev:��ҳ","next:��ҳ","last:ĩҳ")  
	pagerParent = db.GetPager("default") '���ɸ����ҳ������������Ԥ��   	
	
	If Not rsSon.eof Then '��������м�¼   
	'''asp 3.0 ��ִ���֮һ ��ȡ����˳����������
	son0 = rsSon(0)
	Dim rcount : rcount = 0
	
	'easp.wn "rsSon ��¼����: " & rsSon.recordcount
	do While Not rsSon.eof 'ѭ�����౾ҳ��¼   
						'��ʾ�ļ�¼��  rsSon( n ) ����
        Easp.WC "<div>" & rsSon(0) & "|"  & rsSon(1) & "|" & rsSon(2) & "</div>" '�����������   
        rsSon.movenext()  
		rcount = rcount + 1
		If rcount>= 5 Then Exit do
      loop 
      '���������ҳ������ͨ��ajax��ʽ(js���� ajaxGo(����ID,ҳ��) )��ȡ��ҳ���� 

	'''��ʾ����ҳ
	easp.w("<a href=dia.asp>����ҳ(search ���)</a>" & pagerParent)

	db.pagesize=5
	'Easp.W db.Pager("<div>��{liststart}{list}{listend}ҳ,��{jump}ҳ</div>", Array("link:?page=*"))  

    End If 
    Easp.C(rsSon)  


Easp.c(Db)
%>