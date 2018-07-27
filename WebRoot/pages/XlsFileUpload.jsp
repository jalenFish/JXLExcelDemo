<%@ page contentType="text/html; charset=GBK" language="java"%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
 <meta http-equiv="Content-Type" content="text/html; charset=gbk" />
 <title>数据文件上传</title>
 <script>
     var xmlHttp;
     function uploadxls(contextPath){
         //判断是否选择了文件并且后缀是xls
         if(check()){
             //检验文件是否存在
             xlsfileupload(contextPath)
         }
     }
     //获取上传文件 文件名
     function getFileName(){
         //获取文件名
         var upfile = document.getElementById("uploadxlsfile").value;
         var allfilename = upfile.split("\\");
         var filename = allfilename[allfilename.length-1];
         return filename;
     }
     //
     function xlsfileupload(contextPath){

         document.uploadxlsform.action = contextPath+'/LoadReport';
         //document.uploadxlsform.target="_self";
         document.uploadxlsform.submit();

         alert("通讯录更新完成！");
         window.close();

     }
     function submitAfter(req,res){
         var result = res.getAttr("result");
         if(null!=result&&""!=result){
             alert(result);
         }
     }
     //是否选择了文件并且后缀是xls文件判断
     function check(){
         var upfile = document.getElementById("uploadxlsfile").value;
         if(upfile=="" || upfile==null){
             return false;
         }
         if(upfile.endWith("xls") ){
             return true;
         }else{
             alert("请上传xls文件！");
             return false;
         }
         return true;
     }
     String.prototype.endWith=function(oString){
         var   reg=new RegExp(oString+"$");
         return   reg.test(this);
     }
 </script>
 <style>
  .mainContent {
   width: 100%;
   margin: auto;
   text-align: center;
  }
  .fileTable,form{
   width: 50%;
   margin: auto;
   text-align: center;
  }
 </style>
</head>
<body>
<p>
<div class="mainContent" align="center">
 <div class="logo">
  <img src="/pages/images/logo.jpg" width="200" height="200">
  <img src="/pages/images/name.jpg" width="200" height="200">

 </div>
 <div class="fileTable">
  <form name="uploadxlsform" id="uploadxlsform" enctype="multipart/form-data" action="/ProduceExcel" method="post" >
   <table border="0" cellpadding="0" cellspacing="0" width="100%" id="table1" class="table_style">
    <tr class="table_th">
     <td style="text-align:left">
      &nbsp;&nbsp;请选择上传数据xls文件：
     </td>
    </tr>
    <tr>
     <td style="height:60px;">
      <input type="file" name="uploadxlsfile" id="uploadxlsfile" placeholder="请选择数据文件" size="120" />
     </td>
    </tr>
    <%--<tr class="table_th">
     <td style="text-align:left">
      &nbsp;&nbsp;请选择上传模板xls文件：
     </td>
    </tr>
    <tr>
     <td style="height:60px;">
      <input type="file" name="mubanfile" id="mubanfile" placeholder="请选择数据文件" size="120" />
     </td>
    </tr>--%>
   </table>
   <table>
    <tr>
     <td>
      <input type="submit" value="提交" onclick="return check(this.form)" />
      <%-- <input type="button" value="上 传" name="B1" class="button_common_skyBlue" onclick="uploadxls('<%=request.getContextPath()%>');" />
      &nbsp;&nbsp; --%>
      <input type="reset"  value="重 置" name="B2" class="button_common_skyBlue" />
     </td>
    </tr>
   </table>
  </form>
 </div>

</div>
</body>
</html>