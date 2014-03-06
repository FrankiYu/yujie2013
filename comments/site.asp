<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%> 

<% 
db="site.mdb"
dim conn,connstr,db,pSize
pSize=10'每页的留言数
action=Request.Form("action")

'获取导航条显示的内容
if action="getDH" then
	'获取总留言数
	dim totalLY,todayLy,totalPage,str,today_hf
	set conn=Server.CreateObject("ADODB.CONNECTION")	
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;data Source="&Server.MapPath(db)
	conn.open connstr
	set rs=conn.execute("select count(id) from site" )
	totalLY=rs(0)
	'获取今日留言数
	nian=cstr(year(now()))
	if len(month(now()))=1 then
	yue="0"+cstr(month(now()))
	else
	yue=cstr(month(now()))
	end if
	if len(day(now()))=1 then
	ri="0"+cstr(day(now()))
	else
	ri=cstr(day(now()))
	end if
	sj=nian+"-"+yue+"-"+ri
	set rs_today=conn.execute("select count(id) from site where sj='"&sj&"'" )
	todayLy=rs_today(0)
	'获取今日回复
	set rs_today_hf=conn.execute("select count(id) from hf where sj='"&sj&"'" )
	today_hf=rs_today_hf(0)
	'计算总页数
	totalPage=totalLY\pSize
	if totalLY Mod pSize<>0 then 
		totalPage=totalPage+1
	end if
	str=cstr(totalLY)+"鯑"+cstr(todayLy)+"鯑"+cstr(totalPage)+"鯑"+cstr(pSize)+"鯑"+cstr(today_hf)
	

	
	rs.close 
	conn.close 
	set rs=nothing 
	set conn=nothing 
	Response.Charset="UTF-8" 
	Response.Write(str)
'获取指定页码留言数据
elseif action="get" then
	dim currentPage
	currentPage=Cint(Request.Form("page"))
	set conn=Server.CreateObject("ADODB.CONNECTION")	
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;data Source="&Server.MapPath(db)
	conn.open connstr
	set rs=server.CreateObject("adodb.recordset") 
	if currentPage=1 then
		sqlstr="select top "&pSize&" * from site order by px desc,id desc" 
	else
		sqlstr="select top "&pSize&" * from site where id not in(select top "&(currentPage-1)*pSize&"  id  from site order by px desc,id desc) order by px desc,id desc" 
	end if
	rs.open sqlstr,conn,1,3 
	rs.movefirst
	while not rs.eof   
		str2=""
		str=str+rs("nc")+"囃"+rs("ly")+"囃"+rs("sj")+"囃"+rs("isGF")+"囃"+cstr(rs("id"))+"囃"+cstr(rs("px"))+"艸"
		set rs_hf=server.CreateObject("adodb.recordset") 
		sqlstr_hf="select * from hf where toID="&rs("id")&" order by id desc"
		rs_hf.open sqlstr_hf,conn,1,3
		if rs_hf.bof then
		   str=str+"0"+"彟"
		  else		  
			   rs_hf.movefirst
		
			   while not rs_hf.eof 
				  str2=str2+rs_hf("nc")+"鮬"+rs_hf("ly")+"鮬"+rs_hf("sj")+"鮬"+rs_hf("isGF")+"鮬"+cstr(rs_hf("id"))+"贇"
			   rs_hf.movenext  
			   wend
			   str=str+str2+"彟"
		
		end if
		rs_hf.close
		set rs_hf=nothing 
	
	rs.movenext  
	wend
	
	rs.close 
	conn.close 
	set rs=nothing 
	set conn=nothing 
	Response.Charset="UTF-8"
	Response.Write(str)
'检验登陆状态

elseif action="check" then
	if session("uid")=null or session("uid")="" then
		Response.Write("0")
	else
		Response.Write(session("uid"))
    end if
'新增一条留言
elseif action="add" then
	dim nc,nr,sj,isGF,yzm
	nc=trim(Request.Form("nc"))
	nr=trim(Request.Form("nr"))
	yzm=trim(Request.Form("yzm"))
	'判断验证码
	if trim(session("validateCode")) <> yzm then
        response.write("-2")
        response.end
    end if
	
	
	isGF="0"
	if  session("uid") <> ""  Then
	
		isGF="1"
		nc="官方留言"
	
	end if
	nian=cstr(year(now()))
	if len(month(now()))=1 then
	yue="0"+cstr(month(now()))
	else
	yue=cstr(month(now()))
	end if
	if len(day(now()))=1 then
	ri="0"+cstr(day(now()))
	else
	ri=cstr(day(now()))
	end if
	sj=nian+"-"+yue+"-"+ri
	
	on Error Resume  Next
		set conn=Server.CreateObject("ADODB.CONNECTION")	
		connstr="Provider=Microsoft.Jet.OLEDB.4.0;data Source="&Server.MapPath(db)
		conn.open connstr
		set rs=server.CreateObject("adodb.recordset") 
		sqlstr="select * from site" 
		rs.open sqlstr,conn,1,3 
		rs.addnew 
		rs("nc")=nc
		rs("ly")=nr	
		rs("sj")=sj
		rs("isGF")=isGF
		rs("px")=0
		rs.update 
		
		rs.close 
		conn.close 
		set rs=nothing 
		set conn=nothing 
		Response.write("1")
	if Err Then
		Err.clear
		Response.write("0")
		response.end()
	end if
	
'新增一条回复
elseif action="addHF" then
	dim nc_hf,nr_hf,sj_hf,isGF_hf,toID
	nc_hf=Request.Form("nc_hf")
	nr_hf=Request.Form("nr_hf")
	toID=cInt(Request.Form("toID"))
	isGF_hf="0"
	if  session("uid") <> ""  Then
	
		isGF_hf="1"
		nc_hf="官方回复"
	
	end if
	nian=cstr(year(now()))
	if len(month(now()))=1 then
	yue="0"+cstr(month(now()))
	else
	yue=cstr(month(now()))
	end if
	if len(day(now()))=1 then
	ri="0"+cstr(day(now()))
	else
	ri=cstr(day(now()))
	end if
	sj_hf=nian+"-"+yue+"-"+ri
	
	on Error Resume  Next
		set conn=Server.CreateObject("ADODB.CONNECTION")	
		connstr="Provider=Microsoft.Jet.OLEDB.4.0;data Source="&Server.MapPath(db)
		conn.open connstr
		set rs=server.CreateObject("adodb.recordset") 
		sqlstr="select * from hf" 
		rs.open sqlstr,conn,1,3 
		rs.addnew 
		rs("nc")=nc_hf
		rs("ly")=nr_hf	
		rs("sj")=sj_hf
		rs("isGF")=isGF_hf
		rs("toID")=toID
		rs.update 
		
		rs.close 
		conn.close 
		set rs=nothing 
		set conn=nothing 
		Response.write("1")
	if Err Then
		Err.clear
		Response.write("0")
		response.end()
	end if
'登陆
elseif action="login" then
	dim uid,pwd,num
	uid=Request.Form("uid")
	pwd=Request.Form("pwd")
	
	set conn=Server.CreateObject("ADODB.CONNECTION")	
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;data Source="&Server.MapPath(db)
	conn.open connstr
	set rs=server.CreateObject("adodb.recordset") 
	sqlstr="select count(*) from admin where uid='"&uid&"' and pwd='"&pwd&"'"
	rs.open sqlstr,conn,1,3 	
	rs.movefirst
	while not rs.eof 
	num=rs(0)
	rs.movenext  
	wend
	if num>0 then		
		Session("uid")=uid
		response.Write("1")
	end if
	
	
	rs.close 
	conn.close 
	set rs=nothing 
	set conn=nothing 	

'注销
elseif action="logout" then 
	Session.Abandon()
'置顶
elseif action="zd" then
     dim zid,px,old
	 zid=cint(Request.Form("id"))
	 old=cint(Request.Form("old"))
	 if session("uid") <> "" Then
			set conn=Server.CreateObject("ADODB.CONNECTION")	
			connstr="Provider=Microsoft.Jet.OLEDB.4.0;data Source="&Server.MapPath(db)
			conn.open connstr
			
		   '置顶
		    if old=0 then
				set rs1=conn.execute("select top 1 px from site order by px desc" )
	             px=rs1(0)
			     px=px+1
			     set rs=conn.execute("update site set px="&px&" where id="&zid)
			else
				set rs=conn.execute("update site set px=0 where id="&zid)
			end if
			'取消
			
			Response.Write("1")
			rs.close 
			rs1.close
			conn.close 
			set rs=nothing 
			set rs1=nothing 
			set conn=nothing 
			end if
'删除一条留言
elseif action="del" then
	dim id
	id=cint(Request.Form("id"))
		
	 	if session("uid") <> "" Then
			set conn=Server.CreateObject("ADODB.CONNECTION")	
			connstr="Provider=Microsoft.Jet.OLEDB.4.0;data Source="&Server.MapPath(db)
			conn.open connstr
			set rs=server.CreateObject("adodb.recordset") 
			sqlstr="select * from site"
			rs.open sqlstr,conn,1,3 	
			
			while not rs.eof 
			if rs("id")=id then 
			rs.delete 
			rs.update		 
			else 
			rs.movenext 
			end if 
			wend 
			
			Response.Write("1")
			rs.close 
			conn.close 
			set rs=nothing 
			set conn=nothing 
			end if
'删除一条回复
elseif action="delhf" then
	dim hf_id
	 hf_id=cint(Request.Form("id"))
		
	 	if session("uid") <> "" Then
			set conn=Server.CreateObject("ADODB.CONNECTION")	
			connstr="Provider=Microsoft.Jet.OLEDB.4.0;data Source="&Server.MapPath(db)
			conn.open connstr
			set rs=server.CreateObject("adodb.recordset") 
			
				sqlstr="select * from hf"
				rs.open sqlstr,conn,1,3 	
				
				while not rs.eof 
				if rs("id")=hf_id then 
				rs.delete 
				rs.update		 
				else 
				rs.movenext 
				end if 
				wend 
				
				Response.Write("1")
				rs.close 
				conn.close 
				set rs=nothing 
				set conn=nothing 
				
			
			
	end if
'修改密码
elseif action="changePwd" then
	dim oldPwd,newPwd
	oldPwd=Request.Form("oldPwd")
	newPwd=Request.Form("newPwd")
	
	set conn=Server.CreateObject("ADODB.CONNECTION")	
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;data Source="&Server.MapPath(db)
	conn.open connstr
	set rs=server.CreateObject("adodb.recordset") 
	sqlstr="select count(*) from admin where  pwd='"&oldPwd&"'"
	rs.open sqlstr,conn,1,3 	
	rs.movefirst
	while not rs.eof 
	num=rs(0)
	rs.movenext  
	wend
	if num>0 then		
		Session("uid")=uid
		sql="update admin set pwd='"&newPwd&"'"
		on error resume next
		conn.Execute sql
		if error<>0 then
			response.Write(Err.description)
		else
			response.Write("密码修改成功")
		end if
	else
		response.Write("旧密码输入错误，修改密码失败")
	end if
		rs.close 
		conn.close 
		set rs=nothing 
		set conn=nothing 
	
end if



'获取总页数

%>