<%@ page contentType="text/html; charset=GBK" language="java"%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
 <meta http-equiv="Content-Type" content="text/html; charset=gbk" />
 <title>�����ļ��ϴ�</title>
 <script>
     var xmlHttp;
     function uploadxls(contextPath){
         //�ж��Ƿ�ѡ�����ļ����Һ�׺��xls
         if(check()){
             //�����ļ��Ƿ����
             xlsfileupload(contextPath)
         }
     }
     //��ȡ�ϴ��ļ� �ļ���
     function getFileName(){
         //��ȡ�ļ���
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

         alert("ͨѶ¼������ɣ�");
         window.close();

     }
     function submitAfter(req,res){
         var result = res.getAttr("result");
         if(null!=result&&""!=result){
             alert(result);
         }
     }
     //�Ƿ�ѡ�����ļ����Һ�׺��xls�ļ��ж�
     function check(){
         var upfile = document.getElementById("uploadxlsfile").value;
         if(upfile=="" || upfile==null){
             return false;
         }
         if(upfile.endWith("xls") ){
             return true;
         }else{
             alert("���ϴ�xls�ļ���");
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
      &nbsp;&nbsp;��ѡ���ϴ�����xls�ļ���
     </td>
    </tr>
    <tr>
     <td style="height:60px;">
      <input type="file" name="uploadxlsfile" id="uploadxlsfile" placeholder="��ѡ�������ļ�" size="120" />
     </td>
    </tr>
    <%--<tr class="table_th">
     <td style="text-align:left">
      &nbsp;&nbsp;��ѡ���ϴ�ģ��xls�ļ���
     </td>
    </tr>
    <tr>
     <td style="height:60px;">
      <input type="file" name="mubanfile" id="mubanfile" placeholder="��ѡ�������ļ�" size="120" />
     </td>
    </tr>--%>
   </table>
   <table>
    <tr>
     <td>
      <input type="submit" value="�ύ" onclick="return check(this.form)" />
      <%-- <input type="button" value="�� ��" name="B1" class="button_common_skyBlue" onclick="uploadxls('<%=request.getContextPath()%>');" />
      &nbsp;&nbsp; --%>
      <input type="reset"  value="�� ��" name="B2" class="button_common_skyBlue" />
     </td>
    </tr>
   </table>
  </form>
 </div>

</div>
</body>
</html>