<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Conn_User.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/admin_code_article.asp"-->
<%
if CheckUserLogined()=False then
	response.Redirect "User_Login.asp"
end if
dim ClassID,SpecialID
dim SkinID,LayoutID,SkinCount,LayoutCount,ClassMaster,BrowsePurview,AddPurview
ClassID=session("ClassID")
SpecialID=session("SpecialID")
if ClassID="" then
	ClassID=0
else
	ClassID=Clng(ClassID)
end if
if SpecialID="" then
	SpecialID=0
else
	SpecialID=Clng(SpecialID)
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<title>发表文章</title>
<link rel="stylesheet" type="text/css" href="Admin_style.css">
<script language = "JavaScript">
function AddItem(strFileName){
  document.myform.IncludePic.checked=true;
  document.myform.DefaultPicUrl.value=strFileName;
  document.myform.DefaultPicList.options[document.myform.DefaultPicList.length]=new Option(strFileName,strFileName);
  document.myform.DefaultPicList.selectedIndex+=1;
  if(document.myform.UploadFiles.value==''){
	document.myform.UploadFiles.value=strFileName;
  }
  else{
    document.myform.UploadFiles.value=document.myform.UploadFiles.value+"|"+strFileName;
  }
}
function CheckForm()
{
  if (editor.EditMode.checked==true)
	  document.myform.Content.value=editor.HtmlEdit.document.body.innerText;
  else
	  document.myform.Content.value=editor.HtmlEdit.document.body.innerHTML; 

  if (document.myform.ClassID.value=="")
  {
    alert("文章所属栏目不能指定为含有子栏目的栏目！");
	document.myform.ClassID.focus();
	return false;
  }
  if (document.myform.ClassID.value=="0")
  {
    alert("文章所属栏目不能指定为外部栏目！");
	document.myform.ClassID.focus();
	return false;
  }
  if (document.myform.ClassID.value=="-1")
  {
    alert("你没有在此栏目发表文章的权限，请选择其他栏目！");
	document.myform.ClassID.focus();
	return false;
  }

  if (document.myform.Title.value=="")
  {
    alert("文章标题不能为空！");
	document.myform.Title.focus();
	return false;
  }
  //选定课程
  if (document.myform.SpecialID.value==0)
  {
    alert("请指定文章所属课程！");
	document.myform.SpecialID.focus();
	return false;
  }
  /*结束选课程*/
  if (document.myform.Key.value=="")
  {
    alert("关键字不能为空！");
	document.myform.Key.focus();
	return false;
  }
  if (document.myform.Content.value=="")
  {
    alert("文章内容不能为空！");
	editor.HtmlEdit.focus();
	return false;
  }
  if (document.myform.Content.value.length>65536)
  {
    alert("文章内容太长，超出了ACCESS数据库的限制（64K）！建议将文章分成几部分录入。");
	editor.HtmlEdit.focus();
	return false;
  }
  return true;  
}
</script>
</head>
<body leftmargin="5" topmargin="10">
<form method="POST" name="myform" onSubmit="return CheckForm();" action="User_ArticleSave.asp" target="_self">
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="border">
    <tr>
      <td height="22" align="center" class="title"><b>作 者 投 稿 中 心</b></td>
    </tr>
    <tr align="center">
      <td height="266" class="tdbg"><table width="100%" border="0" cellpadding="2" cellspacing="0">
          <tr class="tdbg">
            <td width="102" height="25" align="right"><strong>文章栏目：</strong></td>
            <td width="647"><select name='ClassID'>
                <%call Admin_ShowClass_Option(4,ClassID)%>
              </select>
              <font color="#0000FF">请不要发表在带“*”号的类别</font> </td>
          </tr>
          <tr class="tdbg">
            <td width="102" align="right"><strong>所属课程：</strong></td>
            <td colspan="2"><% call Admin_ShowSpecial_Option(2,SpecialID) %>
              <!-- 用户添加文章的函数,要校难用户权限-->
              <font color="#FF0000">*</font> </td>
          </tr>
          <tr class="tdbg">
            <td width="102" align="right"><strong>文章标题：</strong></td>
            <td colspan="2"><input name="Title" type="text" id="Title" value="" size="50" maxlength="255">
            </td>
          </tr>
          <tr class="tdbg">
            <td width="102" height="20" align="right"><strong>关 键 字：</strong></td>
            <td colspan="2"><input name="Key" type="text"
           id="Key" value="<%=session("Key")%>" size="50" maxlength="255">
              <font color="#0000FF">输入作者或文章内容关键字</font> </td>
          </tr>
          <tr class="tdbg">
            <td width="102" align="right"><strong>任课教师：</strong></td>
            <td colspan="2"><%
call User_ArticleTeacherList()
%><input name="AuthorName" type="hidden"
           id="AuthorName" value="<%=Trim(Request.Cookies("asp163")("UserName"))%>"c>
            </td>
          </tr>
          <tr class="tdbg">
            <td width="102" height="25" align="right"><strong>分页方式：</strong></td>
            <td colspan="2"><select name="PaginationType" id="PaginationType">
                <option value="0" <%if session("PaginationType")=0 then response.write " selected"%>>不分页</option>
                <option value="1" <%if session("PaginationType")=1 then response.write " selected"%>>自动分页</option>
                <option value="2" <%if session("PaginationType")=2 then response.write " selected"%>>手动分页</option>
              </select>
              <font color="#0000FF">手动分页须自己添加分页处，标记符为“</font><font color="#FF0000">[NextPage]</font><font color="#0000FF">”，注意大小写</font></td>
          </tr>
          <tr>
            <td width="102" align="right"><strong>包含图片：</strong></td>
            <td colspan="4"><input name="IncludePic" type="checkbox" id="IncludePic" value="yes">
              是<font color="#0000FF">（如果选中的话会在标题前面显示[图文]）</font></td>
          </tr>
          <tr>
            <td width="102" align="right"><strong>首页图片：</strong></td>
            <td colspan="4"><input name="DefaultPicUrl" type="text" id="DefaultPicUrl" size="56" maxlength="200">
              用在首页的图片文章处显示 <br>
              直接从上传图片中选择：
              <select name="DefaultPicList" id="DefaultPicList" onChange="DefaultPicUrl.value=this.value;">
                <option selected>不指定首页图片</option>
              </select>
              <input name="UploadFiles" type="hidden" id="UploadFiles">
            </td>
          </tr>
        </table>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="border">
          <tr class="tdbg">
            <td width="100%" height="22" align="center" valign="middle">在下面的方框中添加文章内容：（<font color="#FF0000">若你不熟悉以下编辑功能，请勿滥用，直接添加文章即可</font>）</td>
          </tr>
          <tr>
            <td><textarea name="Content" style="display:none"></textarea>
              <iframe ID="editor" src="editor.asp?UserType=User" frameborder=1 scrolling=no width="600" height="405"></iframe></td>
          </tr>
        </table></td>
    </tr>
  </table>
  <div align="center">
    <%dim trs
	  set trs=conn.execute("select SkinID from Skin where IsDefault=True")
	  %>
    <input name="SkinID" type="hidden" id="SkinID" value="?業摤敬?瀼?瑳???????猯牴湯??????????瀼愠楬湧∽敬瑦?昼湯?潣潬??????洦摩潤?????????????????戼?????????洦摩潤??????桓晩?湅整?是湯?戼?????????昼湯?潣潬??????洦摩潤????????湅整?是湯????摴￣??????摴挠汯灳湡∽??整瑸牡慥渠浡??湯整瑮?瑳汹?搢獩汰祡渺湯≥?琯硥慴敲?????????晩慲敭??攢楤潴?猠捲∽摥瑩牯愮灳唿敳呲灹?獕牥?牦浡扥牯敤??捳潲汬湩?潮眠摩桴∽??栠楥桧?????晩慲敭￣???????摴￣?????琯??????琼?汣獡?琢扤≧￣??????琼?楷瑤??∶愠楬湧∽楲桧??瑳潲杮????????猯牴湯??摴￣??????摴挠汯灳湡∽??敳敬瑣渠浡?倢条湩瑡潩呮灹≥椠?倢条湩瑡潩呮灹≥￣????????灯楴湯瘠污敵∽??椥?敳獳潩?倢条湩瑡潩呮灹≥?‰桴湥爠獥潰獮?牷瑩??敳敬瑣摥┢?????灯楴湯￣????????灯楴湯瘠污敵∽??椥?敳獳潩?倢条湩瑡潩呮灹≥??桴湥爠獥潰獮?牷瑩??敳敬瑣摥┢??????灯楴湯￣????????灯楴湯瘠污敵∽??椥?敳獳潩?倢条湩瑡潩呮灹≥?′桴湥爠獥潰獮?牷瑩??敳敬瑣摥┢??????灯楴湯￣???????猯汥捥??扮灳?扮灳?扮灳?扮灳?瑳潲杮?潦瑮挠汯牯∽〣??????是湯??瑳潲杮?潦瑮挠汯牯∽〣?????????????是湯?昼湯?潣潬??????乛硥側条嵥?潦瑮?潦瑮挠汯牯∽〣???????????是湯??摴￣?????琯??????琼?汣獡?琢扤≧￣??????琼?污杩?爢杩瑨?渦獢??摴￣??????摴挠汯灳湡∽????????????????????????瑳潲杮￣???????椼灮瑵渠浡??硡桃牡敐偲条≥琠灹?琢硥?椠??硡桃牡敐偲条≥瘠污敵∽???猠穩???慭汸湥瑧???????????瑳潲杮?琯???????牴￣?????牴￣??????琼?楷瑤??∶愠楬湧∽楲桧??瑳潲杮??????猯牴湯??摴￣??????摴挠汯灳湡∽∴?湩異?慮敭∽湉汣摵健捩?祴数∽档捥扫硯?摩∽湉汣摵健捩?慶畬?礢獥?????????昼湯?潣潬???????ǹ????????????巄??潦瑮?琯???????牴￣?????牴￣??????琼?楷瑤??∶愠楬湧∽楲桧??瑳潲杮??????猯牴湯??摴￣??????摴挠汯灳湡∽∴?湩異?慮敭∽敄慦汵側捩牕?琠灹?琢硥?椠??晥畡瑬楐啣汲?楳敺∽??慭汸湥瑧?㈢???????????????????????牢￣?????????????????????????猼汥捥?慮敭∽敄慦汵側捩楌瑳?摩∽敄慦汵側捩楌瑳?湯桃湡敧∽敄慦汵側捩牕?慶畬?桴獩瘮污敵?￣????????灯楴湯猠汥捥整?????????灯楴湯￣???????猯汥捥??湩異?慮敭∽灕潬摡楆敬?琠灹?栢摩敤≮椠?唢汰慯?汩獥????????琯???????牴￣????琯扡敬￣???琯????牴￣?琯扡敬￣?楤?污杩?挢湥整?￣??瀼￣???┼楤?牴??敳?牴?潣湮攮數畣整∨敳敬瑣匠楫??牦浯匠楫?桷牥?獉敄慦汵?牔敵??┠￣???湩異?慮敭∽歓湩??祴数∽楨摤湥?摩∽歓湩??慶畬???牴???????┼?猠瑥琠獲挽湯?硥捥瑵?猢汥捥?慌潹瑵?映潲?慌潹瑵眠敨敲??晥畡瑬吽畲?湡?慌潹瑵祔数???┠￣???湩異?慮敭∽慌潹瑵??祴数∽楨摤湥?摩∽慌潹瑵??慶畬???牴???????椼灮瑵渠浡??瑣潩≮琠灹?栢摩敤≮椠??瑣潩≮瘠污敵∽摁≤￣???湩異?慮敭∽摁≤琠灹?猢扵業??摩∽摁≤瘠污敵∽???∠漠?楬正∽潤畣敭瑮洮晹牯?捡楴湯?獕牥?瑲捩敬慓敶愮灳?潤畣敭瑮洮晹牯?慴杲瑥?獟汥??￣???扮灳※???椼灮瑵?慮敭∽牐癥敩?琠灹?猢扵業??摩∽牐癥敩?瘠污敵∽???∠漠?楬正∽潤畣敭瑮洮晹牯?捡楴湯?摁業彮牁楴汣健敲楶睥愮灳?潤畣敭瑮洮晹牯?慴杲瑥?扟慬歮????????楤??潦浲?戯摯??瑨汭愾畬?????<%=trs(0)%>">
    <%
	  set trs=conn.execute("select LayoutID from Layout where IsDefault=True and LayoutType=3")
	  %>
    <input name="LayoutID" type="hidden" id="LayoutID" value="<%=trs(0)%>">
    <input name="Action" type="hidden" id="Action" value="Add">
    <input name="Add" type="submit"  id="Add" value=" 添 加 " onClick="document.myform.action='User_ArticleSave.asp';document.myform.target='_self';">
    &nbsp;
    <input
  name="Preview" type="submit"  id="Preview" value=" 预 览 " onClick="document.myform.action='Admin_ArticlePreview.asp';document.myform.target='_blank';">
  </div>
</form>
</body>
</html>
