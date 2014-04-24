<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

  <xsl:template match="Launch">
    <xsl:variable name="date">
      <xsl:value-of select ="substring(@datetime, 1, 10)"/>
    </xsl:variable>
    <xsl:variable name="time">
      <xsl:value-of select ="substring(@datetime, 10)"/>
    </xsl:variable>
    <xsl:variable name="total_tests">
      <xsl:value-of select ="count(Testcase)"/>
    </xsl:variable>
    <xsl:variable name="passed_tests">
      <xsl:value-of select ="count(Testcase[Result='Pass'])"/>
    </xsl:variable>
    <xsl:variable name="failed_tests">
      <xsl:value-of select ="count(Testcase[Result='Fail'])"/>
    </xsl:variable>
	<xsl:variable name="error_tests">
      <xsl:value-of select ="count(Testcase[Result='Error'])"/>
    </xsl:variable>
	<xsl:variable name="abort_tests">
      <xsl:value-of select ="count(Testcase[Result='Aborted'])"/>
    </xsl:variable>
	<xsl:variable name="notrunning_tests">
      <xsl:value-of select ="count(Testcase[Result='Not Running'])"/>
    </xsl:variable>
    <xsl:variable name="percentage">
      <xsl:value-of select="round(($passed_tests div $total_tests) * 100)"/>%
    </xsl:variable>
	<xsl:variable name="elapsed_sec">
      <xsl:value-of select ="floor(sum(Testcase/ElapsedTime)) mod 60"/>
    </xsl:variable>
	<xsl:variable name="elapsed_min">
      <xsl:value-of select ="floor(sum(Testcase/ElapsedTime) div 60) mod 60"/>:
    </xsl:variable>	
	<xsl:variable name="elapsed_hour">
      <xsl:value-of select ="floor(sum(Testcase/ElapsedTime) div 3600) mod 24"/>:
    </xsl:variable>
	<!-- xsl:variable name="elapsed_day" -->
      <!-- xsl:value-of select ="floor(sum(Testcase/ElapsedTime) div 86400)"/ -->
    <!-- /xsl:variable -->
	<xsl:variable name="elapsed_time">
      <xsl:value-of select ="concat($elapsed_hour,$elapsed_min,$elapsed_sec)"/>
    </xsl:variable>
    <html>
      <head>
        <style type="text/css">
          body {font-family:"Calibri", Arial, sans-serif;}
          h1 {}
          h2 {}
          h3 {}

          table {
          border-width: 2px;
          border-spacing: 2px;
          border-style: solid;
          border-color: black;
          border-collapse: collapse;
          background-color: white;
          }
          th {
          border-width: 2px;
          padding: 1px;
          text-align:left;
          border-style: inset;
          border-color: black;
          background-color:rgb(0,51,153);
          color:white;
          }
          td {
          border-width: 2px;
          padding: 1px;
          text-align:left;
          border-style: inset;
          border-color: black;
          background-color:white;
          }
          td.center {
          text-align:center;
          }
          td.large {
          font-size:300%;
          }
          td.pass {
          text-align:center;
          color:black;
          background-color: green;
          }
          td.fail {
          text-align: center;
          color:black;
          background-color: red;
          }
          td.notrunning {
          text-align: center;
          color:black;
          background-color: gray;
          }
		  td.error {
          text-align: center;
          color:black;
          background-color: yellow;
          }
		  td.aborted {
          text-align: center;
          color:black;
          background-color: orange;
          }
        </style>
      </head>
      <body>
      <h1>
        <xsl:value-of select ="@name"/>
      </h1>

      <h2>Summary</h2>
        <table border="1">
          <tr>
            <th>Start Time</th>
            <td><xsl:value-of select ="$date"/>, <xsl:value-of select="$time"/></td>
            <td class="large" rowspan="9">
              <xsl:value-of select="$percentage"/>
            </td>
          </tr>
          <tr>
            <th>Number of Tests</th>
            <td class="center"><xsl:value-of select ="$total_tests"/></td>
          </tr>
          <tr>
            <th>Tests Passed</th>
            <td class="center">
              <xsl:value-of select ="$passed_tests"/>
            </td>
          </tr>
          <tr>
            <th>Tests Failed</th>
            <td class="center">
              <xsl:value-of select ="$failed_tests"/>
            </td>
          </tr>
		  <tr>
            <th>Tests Error</th>
            <td class="center">
              <xsl:value-of select ="$error_tests"/>
            </td>
          </tr>
		  <tr>
            <th>Tests Aborted</th>
			<td class="center">
              <xsl:value-of select ="$abort_tests"/>
            </td>
          </tr>
		  <tr>
            <th>Tests Not Running</th>
            <td class="center">
              <xsl:value-of select ="$notrunning_tests"/>
            </td>
          </tr>
		  <tr>
            <th>Elapsed Time</th>
            <td class="center">
              <xsl:value-of select ="$elapsed_time"/>
            </td>
          </tr>
		  <tr>
            <th>Browser</th>
            <td class="center">
              <xsl:value-of select ="@Browser_Type"/>
            </td>
          </tr>
        </table>
        </body>
      </html>

	<h2>Card Information</h2>
	<table border="1">
		<tr>
		  <th>Agent Date Time</th>
		  <td><xsl:value-of select ="@Agent_Date_Time"/></td>
		</tr>
		<tr>
		  <th>Agent Model Unity</th>
		  <td><xsl:value-of select ="@Agent_Model_Unity"/></td>
		</tr>
		<tr>
		  <th>Agent App Firmware Version</th>
		  <td><xsl:value-of select ="@Agent_App_Firmware_Version"/></td>
		</tr>
		<tr>
		  <th>Agent App Firmware Label</th>
		  <td><xsl:value-of select ="@Agent_App_Firmware_Label"/></td>
		</tr>
		<tr>
		  <th>Agent Boot Firmware Version</th>
		  <td><xsl:value-of select ="@Agent_Boot_Firmware_Version"/></td>
		</tr>
		<tr>
		  <th>Agent Boot Firmware Label</th>
		  <td><xsl:value-of select ="@Agent_Boot_Firmware_Label"/></td>
		</tr>
		<tr>
		  <th>Agent Serial Number</th>
		  <td><xsl:value-of select ="@Agent_Serial_Number"/></td>
		</tr>
		<tr>
		  <th>Agent Manufacture Date</th>
		  <td><xsl:value-of select ="@Agent_Manufacture_Date"/></td>
		</tr>
		<tr>
		  <th>Agent Hardware Version</th>
		  <td><xsl:value-of select ="@Agent_Hardware_Version"/></td>
		</tr>
		<tr>
		  <th>GDD Version</th>
		  <td><xsl:value-of select ="@GDD_Version"/></td>
		</tr>
		<tr>
		  <th>FDM Version</th>
		  <td><xsl:value-of select ="@FDM_Version"/></td>
		</tr>
	</table>
	  
    <h2>Overview</h2>
	<!-- Write the overview table -->
    <table border="1">
      <!-- First write the table header -->
      <tr>
        <th> Index </th>
		<th align="left"> Test Case </th>
        <th align="left"> Elapsed Time </th>
        <th align="left"> Spreadsheet Result Link </th>
		<th align="left"> Testcase Log Link </th>
        <th align="center">Result</th>
      </tr>
      <!-- Fill out each row of the table via the overview template -->
      <xsl:apply-templates select="Testcase" mode="overview"/>
    </table>
  </xsl:template>


  <xsl:template match="Testcase" mode="overview">
    <!-- Determine the result / pass css style name for this test result -->
    <xsl:variable name ="result_style">
      <!-- If the test passed, the style is 'pass' otherwise it's 'fail' -->
      <xsl:choose>
        <xsl:when test="Result='Pass'">pass</xsl:when>
		<xsl:when test="Result='Fail'">fail</xsl:when>
		<xsl:when test="Result='Error'">error</xsl:when>
		<xsl:when test="Result='Aborted'">aborted</xsl:when>
        <xsl:otherwise>notrunning</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
	<xsl:variable name="sec">
      <xsl:value-of select ="floor(ElapsedTime mod 60)"/>
    </xsl:variable>
	<xsl:variable name="min">
      <xsl:value-of select ="floor(ElapsedTime div 60) mod 60"/>:
    </xsl:variable>	
	<xsl:variable name="hour">
      <xsl:value-of select ="floor(ElapsedTime div 3600) mod 24"/>:
    </xsl:variable>
	<!-- xsl:variable name="day" -->
      <!-- xsl:value-of select ="floor(ElapsedTime div 86400)"/ --> 
    <!-- /xsl:variable -->
	<xsl:variable name="time">
      <xsl:value-of select ="concat($hour,$min,$sec)"/>
    </xsl:variable>
	<xsl:variable name="Spreadsheet">
      <xsl:value-of select ="SpreadsheetLink"/>
    </xsl:variable>

	<xsl:variable name="report_date">
      <xsl:value-of select ="substring-after(substring-before($Spreadsheet,'.xls'),'-')"/>
    </xsl:variable>
	
	<xsl:variable name="Spreadsheetpath">
      ../../<xsl:value-of select ="SpreadsheetLink"/>
    </xsl:variable>
	<xsl:variable name="SpreadsheetLink">
      <xsl:value-of select ="substring-before($Spreadsheetpath,$Spreadsheet)"/>Testcase%20Result/<xsl:value-of select ="$report_date"/>/<xsl:value-of select ="$Spreadsheet"/>
    </xsl:variable>
	
	
	<xsl:variable name="TestcaseLog">
      <xsl:value-of select ="TestcaseLogLink"/>
    </xsl:variable>
	<xsl:variable name="TestcaseLogpath">
      ../../<xsl:value-of select ="TestcaseLogLink"/>
    </xsl:variable>
	<xsl:variable name="TestcaseLogLink">
      <xsl:value-of select ="substring-before($TestcaseLogpath,$TestcaseLog)"/>Log%20Recording/<xsl:value-of select ="$report_date"/>/<xsl:value-of select ="$TestcaseLog"/>
    </xsl:variable>
    <tr>
      <td>
        <xsl:value-of select ="position()"/>
      </td>
	  
	  <td>
        <xsl:element name="a">
          <xsl:value-of select ="@name"/>
        </xsl:element>
      </td>

      <td class="center">
        <xsl:value-of select="$time"/>
      </td>
	  
	  <td>
        <xsl:element name="a">
          <xsl:attribute name="href"><xsl:value-of select ="$SpreadsheetLink"/></xsl:attribute>
          <xsl:value-of select ="$Spreadsheet"/>
        </xsl:element>
      </td>
	  <td>
        <xsl:element name="a">
          <xsl:attribute name="href"><xsl:value-of select ="$TestcaseLogLink"/></xsl:attribute>
          <xsl:value-of select ="$TestcaseLog"/>
        </xsl:element>
      </td>
      <!-- td class='pass|fail'>Result</td> -->
      <xsl:element name="td">
        <xsl:attribute name="class"><xsl:value-of select="$result_style"/></xsl:attribute>
        <xsl:value-of select ="Result"/>
      </xsl:element>
    </tr>
  </xsl:template>
 
</xsl:stylesheet>
