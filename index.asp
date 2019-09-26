<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>短信发送测试</title>
<script language="javascript">
  function errorCheck() { 
  return(true);
}

 </script>
</head>
<body>
<form  name="sendsms" action="./smsworks.asp" method="POST" >
<table border="2">
   <tr>
      <td>
         <p align="center">短信发送测试</p>
      </td>
   </tr>
	
   <tr>
      <td>	
	<p>用户名:<input type="text" name="username"></p>
     </td>
   </tr>
   <tr>
      <td> 
	  <p>密&nbsp;&nbsp;码:<input type="password" name="pwd"></p>
      </td>
    </tr>
    <tr>
	<td>
	   <p>手&nbsp;&nbsp;机:<input type="text" name="mobiles"></p>
	</td>
    </tr>
    <tr>
	<td>
	   <p>内&nbsp;&nbsp;容:<textarea name="contents" cols="50" rows="5"></textarea></p>
	</td>
     </tr>
     <tr>
	<td>
            <input type="submit" value="发送"  >
            <input type="submit" value="清空" >
	</td>
        
     </tr>
</table>
</form>
</body>
</html>
