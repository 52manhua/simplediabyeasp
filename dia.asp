<!-- #include file="easp.asp" -->
<head>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<style>
.divscroll{ height:30%;margin-left:10px;overflow-y:scroll; overflow-x:scroll;} 
</style>
</head>
<%
'''获得翻页信息
'page =  easp.r("page",1)
'easp.wn "page= " & page
'''全局变量
Dim dbfile: dbfile= "a2000.mdb"


'''获得表格信息
Dim action
action =  easp.R("action",0)

If Not easp.isN(action) then

	Select Case action
		Case "search":
		Dim search
		search = easp.CheckForm(Easp.R("search",0),"",1,"搜索内容不能为空!")
		
		Case "add":
		Dim uname, ucode
		uname =  easp.CheckForm(Easp.R("name",0),"",1,"姓名不能为空!:不是有效的用户名!")
		ucode = easp.CheckForm(Easp.R("code",0),"",1,"编号/代码不能为空!")

		Case "del":
		Dim uid
		uid =  easp.CheckForm(Easp.R("delid",1),"number",1,"删除编号格式不正确!")
	End select
End if
%>
<%
Dim pagename
pagename = "简易通讯录"
easp.wn "<title>" & pagename & "</title>"
easp.wn "<a href='dia.asp'>" & pagename & "</a>"

Set Db = new EasyAsp.db
Db.dbconn = Db.openconn(1,dbfile,"")
Db.Debug = True

''''搜索

'''针对姓名的模糊搜索
easp.wn(">>>针对姓名的模糊搜索<<< ")
easp.wn "<form method=""post"" action=""dia.asp?action=search"">"
easp.w "搜索姓名所包含的全部或部分单词:<input type=""text"" name=""search"">"
easp.wn "<input type=""submit"" value=""搜索"">"
easp.wn "</form>"

If action = "search" then
	Dim rs
	Set rs = db.GetRecordDetail("cs","[_name] like '%" & search & "%'")

	easp.wn("搜索列表")
	while Not (rs.eof or rs.bof)
		Easp.wn rs("_name") & "'s  code:" & rs("code")
		rs.movenext()
	Wend
	easp.wn " 搜索到: " & (rs.recordcount) & " 条 记录."
	easp.c(rs)
	easp.wn "========================"
End If

''''显示列表
easp.wn "---------------------"
easp.wn "兼具修改的完全列表"
Dim pagelist
'Dim rs
Set rs = db.GetRecordBySQL("Select * from cs")

pagelist= "<div class='divscroll'><br><table><tr><td>姓 名</td><td>号 码</td><td>操 作</td></tr>"

while Not (rs.eof or rs.bof)

	pagelist = pagelist & "<tr><td>" & rs("_name") & "</td><td>code:" & rs("code") & "</td><td>" & "<a href=?action=del&delid="& rs(0) &">删除</a></td></tr>"
	rs.movenext()
Wend
easp.wn(rs.recordcount) & " 条 记录."
easp.c(rs)

easp.wn pagelist & "</table></div>" '''输出html
''''增加记录
If action = "add" Then
dim NewID
NewID = Db.Autoid("cs")

	db.ar "cs",array("_name:" & uname,"code:" & ucode,"sortId:" & NewID)
	easp.rr "dia.asp"
End If

easp.wn "----------------------"
easp.wn "增加联系人:"
easp.wn "<form method=""post"" action=""dia.asp?action=add"">"
easp.wn "姓名:<input type=""text"" name=""name"">"
easp.wn "号码:<input type=""text"" name=""code"">"
easp.wn "<input type=""submit"">"
easp.wn "</form>"

''''删除记录
If action = "del" then
	
	easp.wn "删除编号 #" & uid & " 数据."
	db.DeleteRecord "cs"," id=" & uid
	easp.rr "dia.asp"
'db.DeleteRecord "cs",Array("_name:me")
'db.DeleteRecord "cs"," [_name] like '%me%' "
End If

''''修改记录
'db.UpdateRecord "cs",Array("_name:sme"),Array("code:a12345")


''''分页
easp.wn ""
easp.wn "分页测试"
'在同一个页面上根据记录集生成多个分页导航，同时对主记录集也进行分页(嵌套分页):   

   
Dim rsSon   
'父类分页，每页5条记录，分页方式为array数组   
'''后面的数组参数 Array 参数1->表:数量, 2->条件, 3-> 排序依据
'''4-> 主键

	'
	'db.GPR("array:page:5",
    Set rsSon = db.GPR("array:5",Array("cs","","sortid asc","sortid"))  

	'先定义一个分页导航样式,需要紧跟在 grp 之后   
	db.SetPager "default", "{first}{prev}{liststart}{list}{listend}{next}{last},到{jump}页", Array("listlong:9","listsidelong:3","first:首页","prev:上页","next:下页","last:末页")  
	pagerParent = db.GetPager("default") '生成父类分页导航并将数据预存   	
	
	If Not rsSon.eof Then '如果子类有记录   
	'''asp 3.0 奇怪错误之一 读取数据顺序会引起错误
	son0 = rsSon(0)
	Dim rcount : rcount = 0
	
	'easp.wn "rsSon 记录数量: " & rsSon.recordcount
	do While Not rsSon.eof '循环子类本页记录   
						'显示的记录类  rsSon( n ) 决定
        Easp.WC "<div>" & rsSon(0) & "|"  & rsSon(1) & "|" & rsSon(2) & "</div>" '输出子类内容   
        rsSon.movenext()  
		rcount = rcount + 1
		If rcount>= 5 Then Exit do
      loop 
      '生成子类分页导航，通过ajax方式(js函数 ajaxGo(子类ID,页码) )获取分页数据 

	'''显示导航页
	easp.w("<a href=dia.asp>回首页(search 清空)</a>" & pagerParent)

	db.pagesize=5
	'Easp.W db.Pager("<div>第{liststart}{list}{listend}页,到{jump}页</div>", Array("link:?page=*"))  

    End If 
    Easp.C(rsSon)  


Easp.c(Db)
%>