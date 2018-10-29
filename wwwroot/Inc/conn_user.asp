<%
dim UserTableType,Conn_User,db_bbs
UserTableType = "MyPower"					  ' "Dvbbs6.0" --- 整合动网论坛6.0
											  ' "Dvbbs6.1" --- 整合动网论坛6.1	
											  ' "MyPower"  --- 不整合论坛
db_bbs="database/user.mdb"      '数据库文件的位置
Set Conn_User = Server.CreateObject("ADODB.Connection")
Conn_User.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db_bbs)

sub CloseConn_User()
	Conn_User.close
	set Conn_User=nothing
end sub

'MY动力与动网论坛共用的用户数据表
Const db_User_Table="[User]"

'MY动力与动网论坛共用的用户字段名
Const db_User_ID="UserID"						'用户ID
Const db_User_Name="UserName"					'用户名
Const db_User_Sex="Sex"							'性别

'人文学院两课教改网站新增的字段名

Const db_User_TrueName="TrueName"				'真实姓名

Const db_User_StudentNumber="StudentNumber"		'学号

Const db_User_StudentClass="StudentClass"						'班级
Const db_User_College="College"					'学院
Const db_User_ArticleCommentScore="sum(ArticleComment.Score)"

'以上是人文学院两课教改网站新增的字段名
Const db_User_Email="UserEmail"					'Email地址
Const db_User_Homepage="homepage"				'主页
Const db_User_QQ="Oicq"							'QQ
Const db_User_Icq="icq"							'Icq
Const db_User_Msn="msn"							'Msn
Const db_User_Password="UserPassword"			'密码
Const db_User_Question="Quesion"				'忘记密码的提示问题
Const db_User_Answer="Answer"					'问题答案
Const db_User_Sign="sign"						'签名
Const db_User_Face="face"						'头像
Const db_User_FaceWidth="width"					'头像宽度
Const db_User_FaceHeight="height"				'头像高度
Const db_User_RegDate="addDate"					'注册日期
Const db_User_LoginTimes="logins"				'登录次数
Const db_User_LastLoginTime="lastlogin"			'最后登录时间
Const db_User_LastLoginIP="UserLastIP"			'最后登录IP
Const db_User_UserClass="userclass"				'论坛用户等级（登录时用到）


'MY动力使用的用户字段名
Const db_User_LockUser="lockuser"				'是否锁定用户
Const db_User_ArticleCount="ArticleCount"		'发表文章数
Const db_User_ArticleChecked="ArticleChecked"	'已审核文章数
Const db_User_UserLevel="UserLevel"				'用户等级（权限）
Const db_User_UserPoint="UserPoint"				'用户点数
Const db_User_ChargeType="ChargeType"			'计费方式
Const db_User_BeginDate="BeginDate"				'开始日期
Const db_User_Valid_Num="Valid_Num"				'有效期数值
Const db_User_Valid_Unit="Valid_Unit"			'有效期单位


'动网论坛使用的用户字段名
Const db_User_BbsType="bbstype"
Const db_User_Article="Article"
Const db_User_UserGroup="UserGroup"
Const db_User_UserWealth="userWealth"
Const db_User_UserEP="userEP"
Const db_User_UserCP="userCP"
Const db_User_Title="title"
Const db_User_Showre="showre"
Const db_User_Reann="reann"
Const db_User_UserCookies="usercookies"
Const db_User_Birthday="birthday"
Const db_User_UserPhoto="UserPhoto"
Const db_User_UserPower="UserPower"
Const db_User_UserDel="UserDel"
Const db_User_UserIsBest="UserIsBest"
Const db_User_UserInfo="UserInfo"
Const db_User_UserSetting="UserSetting"
Const db_User_UserGroupID="UserGroupID"
Const db_User_TitlePic="TitlePic"
%>

