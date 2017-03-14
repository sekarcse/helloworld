<!--
  ##############################################################################
  # Copyright Secure Data Solutions.
  ##############################################################################
  # MODULE  NAME:     main.asp
  #
  # MODULE  PURPOSE:  This module presents the search for Customer screen.
  #
  #
  ##############################################################################

  ##############################################################################
  #
  # SOURCE O/S: Windows 2000 Server - IIS
  # TARGET O/S: Windows98/Windows2000/WindowsXP
  #
  # SOURCE LANGUAGE: ASP
  # TARGET LANGUAGE: ASP/Javascript/HTML
  #
  # SOURCE HARDWARE: Intel Server
  # TARGET HARDWARE: PC (Desktop)
  # TARGET SOFTWARE: IE Explorer v6
  #
  ##############################################################################

  ##############################################################################
  # MODIFICATION LOG
  #
  # CODE
  # CHANGE ID  MODULE MER NAME     DATE        CHANGE DESCRIPTION
  # =========  =================== ==========  =================================
  # 000000     Jagdeep Duhra       2005-01-03  ORIGINAL VERSION
  #
  ##############################################################################
  -->
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Buffer=TRUE%><!-- #BeginLibraryItem "/library/user_check.lbi" --><%
'If the session variable is False or does not exsist then redirect the user to the unauthorised user page
If Session("usrAuth") = False or IsNull(Session("usrAuth")) = True then
	
	'Redirect to unauthorised to login page
	Response.Redirect"/default.asp?msg=USRNL"
End If
%><!-- #EndLibraryItem --><!--
  ##############################################################################
  #
  #  INCLUDE FILE SECTION
  #
  ##############################################################################
  -->
<!--#include file='globals.asp'-->
<!--#include file='dbconnection.inc'-->
<!--#include file='sql/def_Messages.inc'-->
<!--#include file='sql/def_Customers.inc'-->
<!--#include file='sql/def_Refcodes.inc'-->
<!--#include file='sql/join_CusElecdet.inc'-->

<!--
  ##############################################################################
  #
  #  VARIABLE DECLARATION SECTION
  #
  ##############################################################################
  -->
<%
Dim strQSMsgCode
Dim strQSMulti
Dim strMsgDescr
Dim boolMoreCustomers
Dim boolMoreRefcode
Dim strRefcode1

Dim arrRefcode1
Dim intRefcode1Idx
Dim arrCustomer1
Dim intCustomer1Idx
Dim arrMessage1
Dim intMessage1Idx
%>


<!--
  ##############################################################################
  #
  #  SUBROUTINE DECLARATION SECTION
  #
  ##############################################################################
  -->
<%
Sub redirectToPage(strPage)
	'Reset server objects
	Call disconnectDB
	
	'Redirect to the requested page
	Response.Redirect strPage
End Sub
%>
<!--
  ##############################################################################
  #
  #  MAIN SECTION
  #
  ##############################################################################
  -->
<%
boolDisplayMenuItems = false
strQSMsgCode         = Request.querystring("msg")
strQSMulti           = Request.querystring("multi")
strMsgDescr          = ""
boolMoreCustomers    = false
strRefcode1			 = "CUSFS"

strFMCustomerName    = request.Form("f_searchWord")
strFMMpan		     = request.Form("f_searchWord")
strFMSupplierAccNo   = request.Form("f_searchWord")
strFMSearchType		 = request.Form("f_searchType")

'Remove session variables containing account being worked on
Session.Contents.Remove("customer_id")
Session.Contents.Remove("company_name")
Session.Contents.Remove("location_id")

' Connect to database
Call connectDB

selectRefcodeByCode strRefcode1, adoCon, arrRefcode1, intRefcode1Idx

if request.Form("f_searchType") <> "" Then

	Select Case strFMSearchType

		Case "CustName"
			If Session("usrPriv") = "1" Then
				selectCustomerLikeName strFMCustomerName, adoCon, arrCustomer1, intCustomer1Idx
			else
				selectCustomerLikeNameOthers strFMCustomerName, adoCon, arrCustomer1, intCustomer1Idx
			End if

		Case "MeterRef"
			If Session("usrPriv") = "1" Then
				seljoinCustomerElectricdetailByMpan2 strFMMpan, adoCon, arrCustomer1, intCustomer1Idx
			else
				seljoinCustomerElectricdetailByMpan2Others strFMMpan, adoCon, arrCustomer1, intCustomer1Idx
			End if

		Case "SuppAccNo"
			If Session("usrPriv") = "1" Then
				seljoinCustomerElectricdetailBySuppAccNo2 strFMSupplierAccNo, adoCon, arrCustomer1, intCustomer1Idx
			else
				seljoinCustomerElectricdetailBySuppAccNo2Others strFMSupplierAccNo, adoCon, arrCustomer1, intCustomer1Idx
			End if
		
	End Select
	
	If dbCusRowCount = 1 Then
		Session("customer_id") = dbCusCustomerId
		redirectToPage("customers/customer.asp")
	End if

	If dbjCusUdtRowCount = 1 Then
		Session("customer_id") = dbjCusUdtCustomerId
		redirectToPage("customers/customer.asp")
	End if

End if

If strQSMsgCode <> "" Then
	selectMessageByCode strQSMsgCode, 1, adoCon, arrMessage1, intMessage1Idx

	If dbMsgRowCount = 1 Then
		strMsgDescr = dbMsgMsgDescr
	End If
End If
%>
<!--
  ##############################################################################
  #
  #  HTML SECTION
  #
  ##############################################################################
  -->
<html>
<head>
<title>Monarch Partnership</title>
<meta http-equiv="Content-Type" content="text/html;">
<script language="JavaScript">
<!--
function MM_displayStatusMsg(msgStr)  { //v3.0
	status=msgStr; document.MM_returnValue = true;
}

function MM_findObj(n, d) { //v3.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document); return x;
}
function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
 var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
   var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
   if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}

function setSelectedOption() {
  var selectedIdx=document.custSelect.f_searchMultiRes.selectedIndex;
  document.custSearch.f_companyName.value = document.custSelect.f_searchMultiRes.options[selectedIdx].text;
  check(document.custSearch,document.custSearch.elements.length)
}
//-->
</script>
<script language="javascript" src="validator.js"></script>

<link href="styles.css" rel="stylesheet" type="text/css">
<meta name="keywords" content="Monarch Partnership">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="expires" content="0">
</head>
<body bgcolor="#ffffff" onLoad="document.custSearch.f_searchWord.focus()" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"><!-- #BeginLibraryItem "/Library/header.lbi" --><!--
  ##############################################################################
  # Copyright Secure Data Solutions.
  ##############################################################################
  # MODULE  NAME:     header.lbi
  #
  # MODULE  PURPOSE:  This module defines the header of the screen, displaying
  #					  the menu tabs depending on whether they are required or
  #					  not.
  #
  #
  ##############################################################################

  ##############################################################################
  #
  # SOURCE O/S: Windows 2000 Server - IIS
  # TARGET O/S: Windows98/Windows2000/WindowsXP
  #
  # SOURCE LANGUAGE: ASP
  # TARGET LANGUAGE: ASP/Javascript/HTML
  #
  # SOURCE HARDWARE: Intel Server
  # TARGET HARDWARE: PC (Desktop)
  # TARGET SOFTWARE: IE Explorer v6
  #
  ##############################################################################

  ##############################################################################
  # MODIFICATION LOG
  #
  # CODE
  # CHANGE ID  MODULE MER NAME     DATE        CHANGE DESCRIPTION
  # =========  =================== ==========  =================================
  # 000000     Jagdeep Duhra       2005-01-03  ORIGINAL VERSION
  #
  ##############################################################################
  -->
<!--
  ##############################################################################
  #
  #  INCLUDE FILE SECTION
  #
  ##############################################################################
  -->
<!--#include virtual='/sql/def_Pagelinks.inc'-->
<!--
  ##############################################################################
  #
  #  VARIABLE DECLARATION SECTION
  #
  ##############################################################################
  -->
<%
Dim strPagelinkType
Dim boolMoreLinks
Dim strBgColor
Dim strMenuText

Dim arrPagelink1
Dim intPagelink1Idx
%>
<!--
  ##############################################################################
  #
  #  SUBROUTINE DECLARATION SECTION
  #
  ##############################################################################
  -->
<!--
  ##############################################################################
  #
  #  MAIN SECTION
  #
  ##############################################################################
  -->
<%
strPagelinkType = "HEADER"
selectPagelinkByType strPagelinkType, adoCon, arrPagelink1, intPagelink1Idx

If dbPglRowCount>0 Then
	boolMoreLinks = true
Else
	boolMoreLinks = false
End if
%>
<!--
  ##############################################################################
  #
  #  HTML SECTION
  #
  ##############################################################################
  -->
<style>
    .contact a
    {
        margin-left:8px;
        color: #fff;
        vertical-align: middle;
        font-size: 14px;
        font-family: "Lato-Bold";
        padding-right: 10px;
        text-decoration: none;
    }
    .contact a:hover
    {
        color: #2c3e50;
        transition: 0.3s ease-in;
        text-decoration: none;
        cursor:pointer;
    }
    
    .contact a:hover {
        -webkit-transition: 0.3s ease-in;
    }
    
    .footer {
        background: #1a1c27;
        padding: 34px 0px;
        color: #fff;
        text-align:center;
    }
    .footer a,
    .footer a:hover,
    .footer a:visited{
        color: #fff;
        text-decoration:none;
    }
    
    .login p {
        font-size: 24px;
        padding-bottom: 10px;
        margin: 0 0 10px;
    }

    .login p, .login label {
        font-family: "MyriadPro-Regular";
        font-weight: bold;
        color: #2c3e50;
    }
    
    .form-group {
    margin-bottom: 15px;
    }

    * {
        box-sizing: border-box;
    }
    
     .login label {
    font-size: 14px;
}

.login label {
    display: inline-block;
    max-width: 100%;
    margin-bottom: 5px;
    }
     .login .form-group input {
    margin-bottom: 30px;
}

.form-control {
    display: block;
    width: 100%;
    height: 34px;
    padding: 6px 12px;
    font-size: 14px;
    line-height: 1.42857;
    color: #555;
    background-color: #fff;
    background-image: none;
    border: 1px solid #ccc;
    border-radius: 4px;
    box-shadow: inset 0 1px 1px rgba(0,0,0,0.075);
    -webkit-transition: border-color ease-in-out 0.15s,box-shadow ease-in-out 0.15s;
    transition: border-color ease-in-out 0.15s,box-shadow ease-in-out 0.15s;
}

 .login .form-group button:hover {
    background-color: #fff;
    color: #ea4b4b;
    -webkit-transition: 0.3s ease-in;
    transition: 0.3s ease-in;
}
 .login .form-group button {
    margin-top: 42px;
    background: #ea4b4b;
    color: #fff;
    margin-right:10px;
    border: 1px solid #ea4b4b;
    padding: 13px 0px;
    font-size: 14px;
    font-family: "MyriadPro-Regular";
    font-weight: bold;
    text-transform: uppercase;
    -webkit-transition: 0.3s ease-out;
    transition: 0.3s ease-out;
}
    
    
</style>
<div style="background-color: #ea4b4b;padding:16px 0;">
        <div class="contact">
            &nbsp;<a href="tel: 020 8835 3535"><i class="fa fa-phone" aria-hidden="true"></i>020 8835
                3535</a>   <a href="mailto:info@monarchpartnership.co.uk"><i class="fa fa-envelope"
                    aria-hidden="true"></i>info@monarchpartnership.co.uk</a>
        </div>
</div>
<table width="750" height="70" border="0" cellspacing="0" cellpadding="0" bgcolor="<%=menuBgColor%>">
  <tr>
    <td align="right" width="10" valign="top" rowspan="2">&nbsp;</td>
    <td rowspan="2"><img src="images/moto_logo_003a.jpg" width="320" height="60"></td>
    <td align="right">
<%
if Session("customer_id") <> "" Then
%>	
	<table class="textheader">
        <tr>
          <td>Customer Id: </td>
          <td><%=Session("customer_id")%> </td>
        </tr>
        <tr>
          <td>Company Name: </td>
          <td><%=Session("company_name")%> </td>
        </tr>
        <tr>
          <td>Location Id: </td>
          <td><%=Session("location_id")%> </td>
        </tr>
        <tr>
          <td colspan="2">&nbsp;</td>
        </tr>
      </table>
<%
Else
%>
	&nbsp;
<%
End if
%>	

	</td>
  </tr>
  <tr>
    <td align="right" width="610" valign="bottom"><%
IF boolDisplayMenuItems Then
%>
      <table cellpadding="0" cellspacing="0">
        <tr>
          <%
DO WHILE ((dbPglRowCount>0) and (boolMoreLinks))
	IF dbPglLabel=strPageId Then
		strBgColor	  =dbPglColour
		strPageColour = dbPglColour
		strMenuText="<font class=""TopMenuLinkSelected"">" & dbPglLabel & "<font>"
	Else
		strBgColor=menuTabColorOff	
		strMenuText="<a href=""" & dbPglLink & """ class=""TopMenuLink"">" & dbPglLabel & "</a>"
	End if
%>
          <td height="25" width="8" bgcolor="<%=strBgColor%>"><img src="images/menu/l_menuTab.gif" height="25"></td>
          <td height="25" width="60" bgcolor="<%=strBgColor%>" align="center"><%=strMenuText%></td>
          <td height="25" width="8" bgcolor="<%=strBgColor%>"><img src="images/menu/r_menuTab.gif" height="25"></td>
          <%
	boolMoreLinks=getNextPagelink(arrPagelink1, intPagelink1Idx)
Loop
%>
          <td height="25"><img src="Library/images/pixel-transparent.gif" width="10" height="1"></td>
        </tr>
        <tr>
          <%
boolMoreLinks=getFirstPagelink(arrPagelink1, intPagelink1Idx)
DO WHILE (dbPglRowCount>0) and (boolMoreLinks)
	IF dbPglLabel=strPageId Then
		strBgColor=dbPglColour
	Else
		strBgColor=menuBgColor	
	End if
%>
          <td colspan="3" bgcolor="<%=strBgColor%>"><img src="Library/images/pixel-transparent.gif" width="1" height="1"></td>
          <%
	boolMoreLinks=getNextPagelink(arrPagelink1, intPagelink1Idx)
Loop
%>
          <td><img src="Library/images/pixel-transparent.gif" width="1" height="1"></td>
        </tr>
      </table>
      <%
Else
%>
      <img src="Library/images/pixel-transparent.gif" width="1" height="1">
      <%
End if
%>
    </td>
  </tr>
</table>
<!--
  ##############################################################################
  #
  #  CLOSE SECTION
  #
  ##############################################################################
  -->
<%
'Call closePagelink
%>
<!-- #EndLibraryItem -->
    <table width="750" border="0" cellspacing="0" cellpadding="0">
        <tr>
            <td>
                <img src="images/pixel-transparent.gif" width="1" height="10">
            </td>
        </tr>
        <tr>
            <td align="center">
                <form name="custSearch" method="post" action="main.asp">
                <div class="login" style="background:#f5f7f8;padding:30px;width:430px;margin:0 auto;">
                    <p>
                        Search for Customer</p>
        
                    <div class="form-group">
                        <label>
                            Search:</label>
                            <input name="f_searchWord" tabindex="1" size="30" maxlength="30" placeholder="Search" class="username form-control">
                            <input name="r_f_searchWord" type="hidden" value="Invalid Search criteria">
                    </div>
                    <div class="form-group">
                        <label>
                            In:</label>
                           <select name="f_searchType" class="texttiny" size="1" tabindex="2">
                                            <%
boolMoreRefcode = getFirstRefcode(arrRefcode1, intRefcode1Idx)
DO WHILE boolMoreRefcode
	If dbRfcValue = "CustName" Then
                                            %>
                                            <option value="<%=dbRfcValue%>" selected="selected">
                                                <%=dbRfcDescr%></option>
                                            <%
	Else
                                            %>
                                            <option value="<%=dbRfcValue%>">
                                                <%=dbRfcDescr%></option>
                                            <%
	End If

	boolMoreRefcode = getNextRefcode(arrRefcode1, intRefcode1Idx)

Loop
                                            %>
                                        </select>
                                        <input name="r_f_locType" type="hidden" value="Invalid Type">
                                        <input name="r_f_fieldType" type="hidden" value="Field type">
                    </div>
                    <div class="form-group">
                        <button class="submit" tabindex="3" value="New customer" type="button" onclick="javascript: window.location='customers/new_customer.asp'">New customer</button>
                        <button class="submit" tabindex="2" value="Search" type="button" onclick="javascript: check(custSearch,custSearch.elements.length)">Search</button>
                        
                    </div>
                </div>
                </form>
            </td>
        </tr>
        <tr>
            <td width="750" valign="top" height="282">
                <table width="700" cellpadding="0" cellspacing="0">
                    <tr>
                        <td height="20">
                            <img src="../images/pixel-transparent.gif" width="1" height="20">
                        </td>
                    </tr>
                    <%
If strFMSearchType = "CustName" Then
                    %>
                    <tr>
                        <td>
                            <table width="700" cellpadding="0" cellspacing="0">
                                <tr class="texttabletitle">
                                    <td width="700">
                                        Company Name
                                    </td>
                                </tr>
                                <%
	boolMoreCustomers = getFirstCustomer(arrCustomer1, intCustomer1Idx)
	DO WHILE boolMoreCustomers
                                %>
                                <tr class="texttablenormal">
                                    <td width="700" align="left">
                                        <a href="customers\customer.asp?customer=<%=Server.URLEncode(dbCusCompanyName)%>"
                                            class="SearchLink">
                                            <%=dbCusCompanyName%></a>
                                    </td>
                                </tr>
                                <%
		boolMoreCustomers = getNextCustomer(arrCustomer1, intCustomer1Idx)
	Loop
ElseIf strFMSearchType = "MeterRef" Then
                                %>
                                <tr>
                                    <td>
                                        <table width="700" cellpadding="0" cellspacing="0">
                                            <tr class="texttabletitle">
                                                <td width="220">
                                                    Company
                                                </td>
                                                <td width="160">
                                                    Mpan 1
                                                </td>
                                                <td width="160">
                                                    Mpan 2
                                                </td>
                                                <td width="160">
                                                    Mpan 3
                                                </td>
                                            </tr>
                                            <%
	boolMoreCustomers = getFirstJoinCusUdt(arrCustomer1, intCustomer1Idx)
	DO WHILE boolMoreCustomers
                                            %>
                                            <tr class="texttablenormal">
                                                <td width="220" align="left">
                                                    <a href="customers\customer.asp?customer=<%=Server.URLEncode(dbjCusUdtCompanyName)%>"
                                                        class="SearchLink">
                                                        <%=dbjCusUdtCompanyName%></a>
                                                </td>
                                                <td width="160" align="left">
                                                    <a href="customers\customer.asp?customer=<%=Server.URLEncode(dbjCusUdtCompanyName)%>"
                                                        class="SearchLink">
                                                        <%=dbjCusUdtMpan1%></a>
                                                </td>
                                                <td width="160" align="left">
                                                    <a href="customers\customer.asp?customer=<%=Server.URLEncode(dbjCusUdtCompanyName)%>"
                                                        class="SearchLink">
                                                        <%=dbjCusUdtMpan2%></a>
                                                </td>
                                                <td width="160" align="left">
                                                    <a href="customers\customer.asp?customer=<%=Server.URLEncode(dbjCusUdtCompanyName)%>"
                                                        class="SearchLink">
                                                        <%=dbjCusUdtMpan3%></a>
                                                </td>
                                                </a>
                                            </tr>
                                            <%
		boolMoreCustomers = getNextJoinCusUdt(arrCustomer1, intCustomer1Idx)
	Loop
Elseif strFMSearchType = "SuppAccNo" Then
                                            %>
                                            <tr>
                                                <td>
                                                    <table width="700" cellpadding="0" cellspacing="0">
                                                        <tr class="texttabletitle">
                                                            <td width="460">
                                                                Company
                                                            </td>
                                                            <td width="240">
                                                                Supplier Account No.
                                                            </td>
                                                        </tr>
                                                        <%
	boolMoreCustomers = getFirstJoinCusUdt(arrCustomer1, intCustomer1Idx)
	DO WHILE boolMoreCustomers
                                                        %>
                                                        <tr class="texttablenormal">
                                                            <td width="460" align="left">
                                                                <a href="customers\customer.asp?customer=<%=Server.URLEncode(dbjCusUdtCompanyName)%>"
                                                                    class="SearchLink">
                                                                    <%=dbjCusUdtCompanyName%></a>
                                                            </td>
                                                            <td width="240" align="left">
                                                                <a href="customers\customer.asp?customer=<%=Server.URLEncode(dbjCusUdtCompanyName)%>"
                                                                    class="SearchLink">
                                                                    <%=dbjCusUdtSupplierAccNo%></a>
                                                            </td>
                                                            </a>
                                                        </tr>
                                                        <%
		boolMoreCustomers = getNextJoinCusUdt(arrCustomer1, intCustomer1Idx)
	Loop
End if
                                                        %>
                                                        <tr>
                                                            <td colspan="5" align="right">
                                                                <input name="f_action" type="hidden" value="insert">
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                        </table>
<!-- #BeginLibraryItem "/Library/footer.lbi" -->
<div class="footer">
    © 2017 All Rights Reserved. <a href="http://www.monarchpartnership.co.uk/">Monarch Partnership
        Ltd</a>.
</div>
<!-- #EndLibraryItem --><!--
  ##############################################################################
  #
  #  CLOSE SECTION
  #
  ##############################################################################
  -->
<%
'Call closeMessage
'Call closeCustomer
Call disconnectDB
%>
</body>
</html>
