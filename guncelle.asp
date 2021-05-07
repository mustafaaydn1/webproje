<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/kayitbaglanti.asp" -->
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%
If (CStr(Request("MM_update")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_kayitbaglanti_STRING
    MM_editCmd.CommandText = "UPDATE [kayit ol] SET Kimlik = ?, adi = ?, soyadi = ?, dogtrh = ?, adres = ?, tel = ? WHERE Kimlik = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("Kimlik"), Request.Form("Kimlik"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 201, 1, -1, Request.Form("adi")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 201, 1, -1, Request.Form("soyadi")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 135, 1, -1, MM_IIF(Request.Form("dogtrh"), Request.Form("dogtrh"), null)) ' adDBTimeStamp
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 201, 1, -1, Request.Form("adres")) ' adLongVarChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 5, 1, -1, MM_IIF(Request.Form("tel"), Request.Form("tel"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", -1, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' N/A
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "uyeol.asp"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If
End If
%>
<%
Dim uyeveri__MMColParam
uyeveri__MMColParam = "1"
If (Request.QueryString("Kimlik") <> "") Then 
  uyeveri__MMColParam = Request.QueryString("Kimlik")
End If
%>
<%
Dim uyeveri
Dim uyeveri_cmd
Dim uyeveri_numRows

Set uyeveri_cmd = Server.CreateObject ("ADODB.Command")
uyeveri_cmd.ActiveConnection = MM_kayitbaglanti_STRING
uyeveri_cmd.CommandText = "SELECT * FROM [kayit ol] WHERE Kimlik = ?" 
uyeveri_cmd.Prepared = true
uyeveri_cmd.Parameters.Append uyeveri_cmd.CreateParameter("param1", 5, 1, -1, uyeveri__MMColParam) ' adDouble

Set uyeveri = uyeveri_cmd.Execute
uyeveri_numRows = 0
%>
<form method="post" action="<%=MM_editAction%>" name="form1">
  <table align="center">
    <tr valign="baseline">
      <td nowrap align="right">Kimlik:</td>
      <td><input type="text" name="Kimlik" value="<%=(uyeveri.Fields.Item("Kimlik").Value)%>" size="32"></td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Adi:</td>
      <td><input type="text" name="adi" value="<%=(uyeveri.Fields.Item("adi").Value)%>" size="32"></td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Soyadi:</td>
      <td><input type="text" name="soyadi" value="<%=(uyeveri.Fields.Item("soyadi").Value)%>" size="32"></td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Dogtrh:</td>
      <td><input type="text" name="dogtrh" value="<%=(uyeveri.Fields.Item("dogtrh").Value)%>" size="32"></td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Adres:</td>
      <td><input type="text" name="adres" value="<%=(uyeveri.Fields.Item("adres").Value)%>" size="32"></td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">Tel:</td>
      <td><input type="text" name="tel" value="<%=(uyeveri.Fields.Item("tel").Value)%>" size="32"></td>
    </tr>
    <tr valign="baseline">
      <td nowrap align="right">&nbsp;</td>
      <td><input type="submit" value="Kayd&#305; G&uuml;ncelle&#351;tir"></td>
    </tr>
  </table>
  <input type="hidden" name="MM_update" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= uyeveri.Fields.Item("Kimlik").Value %>">
</form>
<p>&nbsp;</p>
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title></title>
</head>

<body>
</body>
</html>
<%
uyeveri.Close()
Set uyeveri = Nothing
%>
