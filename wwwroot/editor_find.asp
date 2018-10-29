<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>查找</title>

<SCRIPT LANGUAGE="JScript.Encode">
var notfirstclick=0;
function gofind(){
	var dirx=-64000;
	if(window.form1.direction[1].checked==true){dirx=64000;}
	if(notfirstclick==1){
		var xz=window.dialogArguments;
		var yz=xz.HtmlEdit.document.selection.createRange();
		if(dirx==64000){yz.collapse(false);}else{yz.collapse(true);}
		yz.select();
	}
	var opsx=0;
	if(window.form1.matchcase.checked==true){opsx=opsx+4;}
	if(window.form1.wholeword.checked==true){opsx=opsx+2;}
	var x=window.dialogArguments;
	var y=x.HtmlEdit.document.selection.createRange();
	var TextValue=window.form1.findtext.value;
	if (TextValue=="") {
       alert("请输入查找内容");
       return false;
	}
	var z=y.findText(TextValue,dirx,opsx);
	if(z==false){alert("查找完毕");}else{y.select();}
	if(notfirstclick==0){notfirstclick=1;}
}

function initmoondowner()
{
  window.form1.findtext.focus();
}

function clearme()
{
	var x=window.dialogArguments;
	x.find_status=0;
}

function bye()
{
	var x=window.dialogArguments;
	x.find_status=0;
	window.close();
}
</SCRIPT>

<link rel="stylesheet" href="editor_dialog.css" type="text/css">
</head>
<body  onLoad="initmoondowner();" onunload="clearme();" onkeydown="if (event.keyCode == 13) return false;" bgcolor=menu>
<form method="POST" name="form1">
  <table border="0" cellpadding="3" cellspacing="0" width="390" align="center">
    <tr>
      <td colspan="2" width="296" valign="middle">查找内容: <input type="text" name="findtext" size="27"></td>
      <td width="85"><input type="button" value="查找下一个" name="btn_find" onClick="gofind();" style="width:80px;"></td>
    </tr>
    <tr>
      <td width="166" height="13"> 
        <input type="checkbox" name="wholeword" value="ON">完全匹配</td>
      <td rowspan="2" width="122">
        <fieldset style="padding-bottom:5px"><legend>查找方向:</legend><br>
              <input type="radio" value="64000" checked name="direction"> 向上<input type="radio" name="direction" value="-64000" checked> 向下
			  </fieldset>
      </td>
      <td width="85" height="13"> 
        <input type="button" value="取消" name="btn_cancel" onClick="bye();" style="width:80px;"></td>
    </tr>
    <tr>
      <td width="166"><input type="checkbox" name="matchcase" value="ON">匹配任何一个 </td>
      <td width="85"></td>
    </tr>
  </table>
</form>
</body>

</html>