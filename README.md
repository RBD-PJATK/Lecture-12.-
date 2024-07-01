# Lecture 12. ADO. Database API

<h3>Abstract</h3>

<p>Lecture 12 presents a method to access databases in programs written in 3GL
(= <i>3rd Generation Language</i>), i.e. in imperative programming languages like Visual Basic and
C++.  This method uses the programmer's interface called ADO (= <i>ActiveX Data Objects</i>). 
It is a language independent API based on predefined objects <i>Connection</i>, <i> Command</i> and <i>Recordset</i>.
MS Access implements ADO.</p>

<p>In the second part of the lecture we mention ASP (= <i>Active Server Pages</i>) and applications of ADO in ASP.
ASP is a script language that extends HTML. We present the new library developed by Microsoft called ADO.NET.
It is the version of ADO tailored for web applications on .NET platform.</p>

<hr><h3><a name="api">Call Level Interface</a></h3> 

<p>A database application often performs a set of operations on the database
without any communication with the user. It can be implemented with calls to
the <code>RunSQL</code> method of the object <code>DoCmd</code>. However, it has
some drawbacks.  First, we cannot process data as it is read row by row.
Second, the object <code>DoCmd</code> is not available in other development
environments.  There is another method called <i>CLI</i> (=
<i>Call Level Interface</i>). The rules of CLI are part of the SQL
standard.  This lecture presents one of CLI interfaces called <i>ADO</i> (=
<i>ActiveX Data Objects</i>).  ADO has been developed by Microsoft.  The
lecture describes also another interface of this family.  It is the database
interface for Java called <i>JDBC</i> (= <i>Java DataBase Connectivity</i>). 
JDBC has been developed by Sun.</p>

<hr><h3><a name="Model">Programming model of ADO</a></h3>

<p><i>The programming model</i> defines the collection of objects and their
methods which facilitate accessing and updating various data sources and
databases among them. Here are the fundamental rules.</p>

<ol>
<li>Connection to a data source is accomplished by means of an object of the <code>Connection</code>
	class.
<li>An SQL statement to be executed is presented to the data source.
<li>The data source executes this statement.  If this is a SELECT statement, the resulting rows will be stored in
	the returned object	
	<code>Recordset</code>.  The application can browse through them.
<li>If necessary, the application can update these rows by means 
	of methods of the <code>Recordset</code> object.
<li>The errors that have occurred during the connection to the database and the execution of the statement may be detected. 
</ol>


<hr><h3><a name="Podstawowe obiekty">Fundamental objects of ADO</a></h3>

<dl>
<dt><code>Connection</code>
<dd>Its is the root of the hierarchy of classes of ADO. We use it when we connect to a data source.
<dt><code>Recordset</code>
<dd>It represents the set of records returned by a data source. We use it to process these records. By means of
a <code>Recordset</code> you can browse through them, modify and delete them and add new records. 
At one time the <code>Recordset</code> object exhibits only one record called <i>the current record</i>.
<dt><code>Command</code>
<dd>It represents an SQL statement.
<dt><code>Error</code>
<dd>It represents an error encountered by ADO.
</dl>

<h4><a name="Uzycie">Connection</a></h4>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad12/images/11_1.png"></p>

<p>In order to connect to a database, you have to create an object of the <code>Connection</code>
class.</p>

<pre>Dim cnCurrent As ADODB.Connection
Set cnCurrent = CurrentProject.Connection</pre>

<p>Object <i>cnCurrent</i> allows accessing all data of the current MS Access database by means of SQL statements.</p>

<p>We will use sample database schema with tables <i>Customers</i> (with columns <i>Cust_id</i> and <i>Name</i>)
and <i>Products</i> (with column <i>Prod_id</i>).</p>

<h5><a name="remote-con">Remote database connection through ODBC</a></h5>

<p>In the example below we connect to an Oracle database. We log on as
user <i>scott</i> with password <i>tiger</i>.  We assume that we have
defined ODBC data source name <code>DSN=scott</code>.  Of course the connecting
command does not reveal that this is an Oracle database, because ODBC hides the details
of the data source. Therefore databases of other vendors are referenced the same way.</p>

<pre>Dim cnCurrent As ADODB.Connection
Set cnCurrent = New ADODB.Connection
cnCurrent.ConnectionString = "DSN=scott;UID=scott;PWD=tiger;"
cnCurrent.Open</pre>

<p>In the examples we will use a sample table <i>Emp</i> installed in <i>scott</i>'s schema.</p>

<p>If <code>ConnectionString</code> contains the <code>Provider</code>
parameter, we can set
additional parameters specific for the indicated data provider.</p>

<ul>
<li><code>Provider = "SQLOLEDB"</code> (Microsoft OLE DB Provider for SQL Server)
<li><code>Provider = "MSDAORA"</code> (Microsoft OLE DB Provider for Oracle)
<li><code>Provider = "MSDASQL"</code> (Microsoft OLE DB Provider for ODBC) is the default setting.
</ul>

<p>For example:</p>

<pre>Dim cnn As ADODB.Connection
Set cnn = New ADODB.Connection
cnn.Provider = "SQLOLEDB"
cnn.Open "Data Source=srv;Initial Catalog=pubs;", "scott", "tiger"</pre>

<h4>Recordset</h4>

<p>In order to declare and create a <code>Recordset</code>, use:</p>

<pre>Dim rsCustomers As ADODB.Recordset
Set rsCustomers = New ADODB.Recordset</pre>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad12/images/11_2.png"></p>

<p>To open this <code>Recordset</code>, use its
<code>Open</code> method.</p>

<pre>rsCustomers.Open "Customers", cnCurrent</pre>

<p>To close the <code>Recordset</code> and the <code>Connection</code> and to release
	their memory, write. 

<pre>rsCustomers.Close
cnCurrent.Close
Set rsCustomers = Nothing
Set cnCurrent = Nothing</pre>

<p>You can reference fields of a <code>Recordset</code> the same way as fields of
a form. The exclamation mark is the reference to an item of the
<code>Fields</code> collection, e.g.</p>

<pre>MsgBox rsCustomers!Name</pre>

<p>The following methods are used to navigate to another record of the <code>Recordset</code>.</p>

<ul>
<li><code>MoveFirst</code>
<li><code>MoveLast</code>
<li><code>MoveNext</code>
<li><code>MovePrevious</code>
</ul>

<p>The following properties indicate the position inside the <code>Recordset</code>.</p>

<ul>
<li><code>BOF</code> is true, if the current record is positioned before the first one.
<li><code>EOF</code> is true, if the current record is positioned after the last one.
</ul>

<p>You can scan all records of the <code>Recordset</code>.</p>

<pre>Do Until rsCustomers.EOF
  MsgBox rsCustomers!Name
  rsCustomers.MoveNext
Loop</pre>

<p><a name="select-exam"/>The <code>Recordset</code> can be based on an SQL query
(we assume that <code>txtName</code> is a field of the 
<code>customers</code> form), e.g.
</p>

<pre>rsCustomers.Open "SELECT * FROM Customers", cnCurrent</pre>

<pre>strSQL = "SELECT * FROM Customers" &amp; _
         "WHERE Name = '" &amp; Forms!Customers!txtName.Value &amp; "'"
rsCustomers.Open strSQL, cnCurrent</pre>


<hr><h3><a name="Wykonywanie">Using SQL statements</a></h3>

<p>There are several ways to execute an SQL statement in code written in VBA.
You can use one of the following methods.</p>

<dl>
<dt><code>RunSQL</code> of object <code>RunSQL</code>
<dd>We have described it during previous lectures. This method can be used only 
	inside MS Access.
<dt><code>Execute</code> of object <code>Connection</code>
<dd>If the executed statement is SELECT, the returned <code>Recordset</code>
	will be read-only.
<dt><code>Execute</code> of object <code>Command</code>
<dd>If the executed statement is SELECT, the returned <code>Recordset</code>
	will be read-only.
<dt><code>Open</code> of object <code>Recordset</code>
<dd>It can also execute INSERT, UPDATE and DELETE.
</dl>

<p>Now we present several examples of data manipulation.
Each operation is illustrated twice. The first example uses 
the <code>Execute</code> method of the <code>Connection</code> object,
while the second example uses methods of the <code>Recordset</code>
object.</p>

<h4><a name="insert-exam">Examples of INSERT</a></h4>

<pre>Dim cnCurrent As ADODB.Connection
Set cnCurrent = CurrentProject.Connection
strSQL = "INSERT INTO Customers(Cust_id, Name) VALUES ('" _
             &amp; Me!txtCust_id.Value &amp; "','" _
             &amp; Me!txtName.Value &amp; "')"
cnCurrent.Execute strSQL</pre>

<pre>rsCustomers.AddNew
rsCustomers!Cust_id = Me!txtCust_id.Value
rsCustomers!Name = InputBox("Enter Name:")
rsCustomers.Update</pre>

<h4>Examples of DELETE</h4>

<pre>Dim cnCurrent As ADODB.Connection
Set cnCurrent = CurrentProject.Connection
strSQL = "DELETE * FROM Customers " &amp; _
         "WHERE Cust_id = " &amp; Me!txtCust_id.Value
cnCurrent.Execute strSQL</pre>

<p>If the current record of <code>Recordset</code> <code>rsCustomers</code> is the one to be updated,
we can do it in the following way:</p>

<pre>rsCustomers.Delete
rsCustomers.MoveNext
If rsCustomers.EOF Then
  rsCustomers.MoveLast
End If</pre>

<p>Note that we have to move the pointer of the current record after
the deleted record (<code>rsCustomers.MoveNext</code>).  Furthermore, if
this operation moves beyond the last record, we have to set the pointer 
to the last record (<code>rsCustomers.MoveLast</code>).</p>


<h4>Examples of UPDATE</h4>

<p>We assume that <code>txtCust_id</code> is a field of the
<code>customers</code> form.</p>

<pre>Dim cnCurrent As ADODB.Connection
Set cnCurrent = CurrentProject.Connection
strSQL = "UPDATE Customers SET Name = '" &amp; txtName.Value &amp; "'" &amp; _
         "WHERE Cust_id = '" &amp; Me!txtCust_id.Value &amp; "'"
cnCurrent.Execute strSQL</pre>

<p>If the current record of <code>Recordset</code> <code>rsCustomers</code> is the one to be updated,
we can do it in the following way:</p>

<pre>rsCustomers!Name = InputBox("Enter Name:")
rsCustomers.Update</pre>

<p>Let us assume that we want to update the <i>Emp</i> table so that the <i>Job</i> of every
salesman is set to accountant. The simplest way to it is to run the following SQL statement:</p>

<pre>UPDATE Emp
SET Job = "Accountant"
WHERE Job = "Salesman"</pre>

<p>If we want to use the programming language, we can write a program that
retrieves subsequent records from the <i>Emp</i> table. For every record, the
program checks whether the current record describes a salesman.  If it does,
the program will modify this record by changing the job to accountant.  
Access to records of the table is performed by the
<code>Recordset</code> object.  At the beginning we create it and supply the
specification of a record source, then we use methods <code>MoveFirst</code>
and <code>MoveNext</code> to visit all returned records.</p>

<pre>Dim cnCurrent As ADODB.Connection
Set cnCurrent = CurrentProject.Connection

Dim rsEmp As ADODB.Recordset
Set rsEmp = New ADODB.Recordset

rsEmp.Open "Emp", cnCurrent
rsEmp.MoveFirst

Do Until rsEmp.EOF
  If rsEmp!Job = "Salesman" Then 
    rsEmp!Job = "Accountant"
    rsEmp.Update
  End If

  rsEmp.MoveNext
Loop

rsEmp.Close
Set rsEmp = Nothing</pre>

<p><b>Warning:</b> There are data sources that do not allow modifying the database this way.</p>

<p>The order of visited records may be determined by an index that has been
created for the table. If there is an index on the <i>Ename</i> column of the
<i>Emp</i> table, then we can browse the records of <i>Emp</i> in the sequence
defined by this index. The index for a Recordset must be fixed before the
call to method <code>MoveFirst</code>.<p>

<pre>rsEmp.Index = "Ename"</pre>

<hr><h3><a name="remore-ops">Remote operations</a></h3>

<p>The most important advantage of ADO is the possibility to access
heterogeneous data sources in a uniform way.  A programmer just sets
the <code>ConnectionString</code> property of the 
<code>ConnectionString</code> object and then accesses the data from the data
source by universal classes and methods of ADO (see <a
href="#remote-con">example</a>). ADO can be used in all programs written in
VBA and in scriptlets inside ASP (= Active Server Pages).</p>

<p>We show several procedure that connect to the database through ODBC and
perform some operations on its data. For the time being we do not care about errors.  We will
consider them <a href="#Uzycie1">soon</a>.

<h4>Raise salaries of all employees who earn less than 2000</h4>

<pre>Public Sub SalRise()
  Dim cnn As ADODB.Connection
  Set cnn = New ADODB.Connection
  cnn.ConnectionString = "DSN=scott;UID=scott;PWD=tiger;"
  cnn.Open

  Dim strSQL As String
  strSQL = "UPDATE Emp SET Sal=Sal*1.1 WHERE Sal &lt; 2000"

  cnn.Execute strSQL

  cnn.Close
  Set cnn = Nothing

End Sub</pre>

<h4>Display names of all employees</h4>

<pre>Public Sub Show_Emps()
  Dim cnn As ADODB.Connection
  Set cnn = New ADODB.Connection
  cnn.ConnectionString = "DSN=scott;UID=scott;PWD=tiger;"
  cnn.Open

  Dim rsEmps As ADODB.Recordset
  Set rsEmps = New ADODB.Recordset
  rsEmps.Open "Emp", cnn
  rsEmps.MoveFirst
  Do Until rsEmps.EOF
      MsgBox rsEmps!Ename
      rsEmps.MoveNext
  Loop

  rsEmps.Close
  cnn.Close
  Set rsEmps = Nothing
  Set cnn = Nothing

End Sub </pre>

<h4>Display the employees with the highest salary</h4>

<pre>Public Sub EmpHighSal()
  Dim cnn As ADODB.Connection
  Set cnn = New ADODB.Connection
  cnn.ConnectionString = "DSN=scott;UID=scott;PWD=tiger;"
  cnn.Open

  Dim rsEmps As ADODB.Recordset
  Set rsEmps = New ADODB.Recordset
  Dim strSQL As String
  strSQL = "SELECT Ename, Sal FROM Emp WHERE Sal = (SELECT Max(Sal) FROM Emp)"

  rsEmps.Open strSQL, cnn

  rsEmps.MoveFirst
  Do Until rsEmps.EOF
     MsgBox rsEmps!Ename & " Sal = " & rsEmps!Sal
     rsEmps.MoveNext
  Loop

  rsEmps.Close
  cnn.Close
  Set rsEmps = Nothing
  Set cnn = Nothing

End Sub</pre>

<hr><h3><a name="Transak">Transactions</a></h3>

<p><i>A database transaction</i> is a sequence of INSERT, UPDATE and
DELETE statements that is performed as a whole.  Either all or none of these operations
are executed.  If you are using ADO, you have to start the transaction
explicitly, because
by default all SQL statements are single item transactions with <i>auto-commit</i>.
ADO transactions may be nested.</p>

<p>The following methods of object <code>Connection</code> manipulate transactions.</p>

<dl>
<dt><code>BeginTrans</code>
<dd>It starts a new transaction and returns the nesting level of this transaction.
<dt><code>CommitTrans</code>
<dd>It commits the changes made by the current transaction and closes
	this transaction.
<dt><code>RollbackTrans</code> 
<dd>It cancels the changes made by the current transaction and closes this transaction.
</dl>

<p>The same effect can be achieved by the following SQL statements executed
by the <code>Execute</code> method.</p>

<ul>
<li><code>BEGIN TRANSACTION</code>
<li><code>COMMIT</code>
<li><code>ROLLBACK</code>
</ul>

<p><b>Warning:</b> There are data providers that do not implement transactions.</p>

<hr><h3><i>Command</i> <a name="Command">Object</a></h3>

<p>The <code>Command</code> object represents an SQL statement to be executed
by the data source.</p>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad12/images/11_3.png"></p>

<p>The <code>Command</code> object is useful when you need to run the same statement more
than once or when you need a parameterized statement (we will not consider parameters here).
We have already shown how to execute SQL statements by means of
<a href="#insert-exam"><code>Connection</code></a>
and <a href="#select-exam"><code>Recordset</code></a>.</p>

<pre>Dim strCnn As String
strCnn = "DSN=scott;UID=scott;PWD=tiger;"

Dim cnn As ADODB.Connection
Set cnn = New ADODB.Connection
cnn.Open strCnn

Dim cmdChange As ADODB.Command
Set cmdChange = New ADODB.Command
Set cmdChange.ActiveConnection = cnn

Dim strSQL As String
strSQL = "UPDATE Emp SET Job = 'Accountant' " &amp; _
         "WHERE Job = 'Salesman'"
cmdChange.CommandText = strSQL

cmdChange.Execute</pre>

<hr><h3><a name="Uzycie1">Using collection <i>Errors</i> and object <i>Error</i></a></h3>

<p>Every call to a method of an ADO object may cause one or more errors
which
are reported by the data source. Each error is represented as a separate
object of the <code>Error</code> class.  The <code>Errors</code> collection contains
all errors reported by the most recent command. When a new command generates
errors, this collection is cleared and filled up with new errors.</p>

<p>Apart from ADO errors there are also VBA errors that can occur during the
execution of the code. These errors are stored in the <code>Err</code>
object described by the previous lectures on VBA.</p>

<p>An object of the <code>Error</code> type has the following properties.</p>

<dl>
<dt><code>Description</code>
<dd>The text that describes the error.
<dt><code>Number</code>
<dd>The number of the error.
<dt><code>Source</code>
<dd>The object that reported the error.
<dt><code>SQLState</code> and <code>NativeError</code>
<d>The information from the data source.
</dl>

<p>As an example, we will code the error handling into the previous example.
We will catch errors that arise during the execution of
<font color="red"><code>cmdChange.Execute</code></font>.</p>

<pre>Dim strCnn As String
strCnn = "DSN=scott;UID=scott;PWD=tiger;"

Dim cnn As ADODB.Connection
Set cnn = New ADODB.Connection
cnn.Open strCnn

Dim cmdChange As ADODB.Command
Set cmdChange = New ADODB.Command
Set cmdChange.ActiveConnection = cnn

Dim strSQL As String
strSQL = "UPDATE Emp SET Job = 'Accountant' " &amp; _
         "WHERE Job = 'Salesman'"
cmdChange.CommandText = strSQL

<font color="blue"><code>On Error GoTo Err_Execute</code></font>
<font color="red"><code>cmdChange.Execute</code></font>

<font color="blue"><code>Err_Execute:</code></font>
  Dim errLoop As ADODB.Error
  If cnn.Errors.Count > 0 Then
    For Each errLoop In cnn.Errors
      MsgBox "Errno: " & errLoop.Number & vbCr & errLoop.Description
    Next errLoop
  End If
  <font color="blue"><code>Resume Next</code></font></pre>

<hr><h3><a name="Skrypty">Server-side script language ASP</a></h3>

<p>A database application on the Internet contains code with commands which
generate HTML documents and retrieve data from the database. The most
popular tool to create such applications are so called <i>server pages</i>,
i.e. scripts written in languages like PHP, ASP and JSP. We will shortly
sketch the structure of server pages developed in the ASP language (= <i>Active
Server Pages</i>) invented at Microsoft. We start from some basic
information on HTML.</p>

<hr><h3><a name="HTML">HTML</a></h3>

<p>HTML is used to build text documents enriched by markup.  The markup has
the form of tags which determine the structure of a document from the point
of view of a web browser.  The tags indicated the beginning and the end of the
title, the header, the body etc.</p>

<pre>&lt;HTML&gt;
   &lt;HEAD&gt;
     &lt;TITLE&gt;The title of the document&lt;/TITLE&gt;
   &lt;/HEAD&gt; 
   &lt;BODY&gt;
     The content of the page.
   &lt;/BODY&gt;
&lt;/HTML&gt;</pre>


<h4>Paragraph</h4>


<pre>&lt;P&gt;The text of the paragraph.&lt;/P&gt;</pre>

<h4>Unordered list</h4>

<table border="1"  align="center">
<tr>	<th>HTML
	<th>The view in a browser
<tr><td>
<pre>&lt;UL&gt;
  &lt;LI&gt;Fruits&lt;/LI&gt;
  &lt;LI&gt;Vegetables&lt;/LI&gt;   
  &lt;LI&gt;Fish&lt;/LI&gt;
  &lt;LI&gt;Meat&lt;/LI&gt;
  &lt;LI&gt;Chicken&lt;/LI&gt;
&lt;/UL&gt;</pre>
<td>
<UL>
 <LI>Fruits</LI>
 <LI>Vegetables</LI>
 <LI>Fish</LI>
 <LI>Meat</LI>
 <LI>Chicken</LI>
</UL></table>

<h4>Ordered list</h4>

<table border="1"  align="center">
<tr>	<th>HTML
	<th>The view in a browser
<tr><td>
<pre>&lt;OL&gt;
  &lt;LI&gt;Fruits&lt;/LI&gt;
  &lt;LI&gt;Vegetables&lt;/LI&gt;   
  &lt;LI&gt;Fish&lt;/LI&gt;
  &lt;LI&gt;Meat&lt;/LI&gt;
  &lt;LI&gt;Chicken&lt;/LI&gt;
&lt;/OL&gt;</pre>
<td>
<OL>
 <LI>Fruits</LI>
 <LI>Vegetables</LI>
 <LI>Fish</LI>
 <LI>Meat</LI>
 <LI>Chicken</LI>
</OL></table>


<h4>Table</h4>

<p>A table is defined row by row. Each row is surrounded by tags
<code>&lt;TR&gt;...&lt;/TR&gt;</code> while each cell is delimited by
tags <code>&lt;TD>...&lt;/TD></code>.</p>

<table border="1"  align="center">
<tr>	<th>HTML
	<th>The view in a browser
<tr><td>
<pre>
<code>&lt;TABLE border="1"&gt;
  &lt;TR&gt;&lt;TD&gt;One&lt;/TD&gt;&lt;TD&gt;Two&lt;/TD&gt;&lt;/TR&gt;
  &lt;TR&gt;&lt;TD&gt;Three&lt;/TD&gt;&lt;TD&gt;Four&lt;/TD&gt;&lt;/TR&gt;
&lt;/TABLE&gt;</pre>
<td>
<TABLE border="1" align="center">
<TR><TD>One</TD><TD>Two</TD></TR>
<TR><TD>Three</TD><TD>Four</TD></TR>
</TABLE>
</table>

<p>The <code>border="1"</code> attribute says that the table has the visible border
(the default value of this attribute is 0, i.e. no frame).</p>


<h4>Text formatting</h4>

<table border="1"  align="center">
<tr>	<th>HTML
	<th>The view in a browser
<tr><td><pre>&lt;B&gt;Do not forget!&lt;/B&gt;</pre>
<td><B>Do not forget!</B>
<tr><td><pre>&lt;I&gt;Do not walk!&lt;/I&gt;</pre>
<td><I>Do not walk!</I>
<tr><td><pre>&lt;U&gt;Never give in!&lt;/U&gt;</pre>
<td><U>Never give in!</U>
</table>

<h4>Hyperlink</h4>

<p>The hyperlink <a href="http://www.pjwstk.edu.pl/">Visit our school</a> can be defined as:

<pre>&lt;A href="http://www.pjwstk.edu.pl/"&gt;Visit our school&lt;/A&gt;</pre>

<p>Tags <code>&lt;A&gt;</code>  and <code>&lt;/A&gt;</code>  define a hyperlink 
to another document.  Therefore they are metadata.</p>

<h4>Image</h4>

<p>In order to place an image onto the document, use the following single tag.</p>

<pre>&lt;IMG SRC="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad12/images/11_1.png"/&gt;</pre>

<P align="center"><IMG SRC="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad12/images/11_1.png"/></p>

<h4>Form</h4>

<p>A form is a complex structure built up from text, text fields and buttons.</p>


<table border="1" align="center">
<tr>	<th>HTML
	<th>The view in a browser
<tr><td>
<pre>
&lt;FORM METHOD=&quot;POST&quot;
      ACTION=&quot;http://xeon/display&quot;&gt;
  Enter table name:&lt;br&gt;
    &lt;INPUT TYPE=&quot;text&quot;
           NAME=&quot;name&quot;&gt;&lt;br&gt;
  Press button:&lt;br&gt;
    &lt;INPUT TYPE=&quot;submit&quot;
           VALUE=&quot;Display&quot;&gt;
&lt;/FORM&gt;</pre>
<td>
<FORM METHOD="POST"
      ACTION="http://xeon/display">
  Enter table name:<br>
    <INPUT TYPE="text"
           NAME="name"><br>
  Press button:<br>
    <INPUT TYPE="submit"
           VALUE="Display">
</FORM>
</table>

<p>The <code>ACTION</code> attribute of <code>FORM</code> indicates that the
program running in the Internet will process data entered by a
user into this form.</p>

<hr><h3><a name="ASP">ASP</a></h3>

<p>ASP is a direct extension of HTML with scriptlets (fragments of code) written
in Visual Basic. The scriptlets are surrounded by tags 
<code>&lt;%</code> and <code>%&gt;</code>.</p>

<p>ASP can connect to a database by means of ADO, retrieve some data from this database
and place it in the response document. If there are data items sent from a form or 
enclosed in the URL, ASP can use them for example in SQL statements in the script.</p>

<p>The interpreter of ASP creates the following objects.</p>

<dl>
<dt><code>Server</code>
<dd>The handle to the server. This is the factory object for
	objects <code>Connection</code> and <code>Recordset</code>.
<dt><code>Application</code>
<dd>The application as a whole.
<dt><code>Session</code>
<dd>The current session of the user.  It can store certain objects, e.g.
	<code>Session("RecSet")</code> is the reference to a record set created previously.
<dt><code>Request</code>
<dd>The data sent by the user in the request, e.g. <code>Request.Form("Name")</code> 
	represents the value entered into field <i>Name</i> of the form.
<dt><code>Response</code>
<dd>The document HTML to be sent to the user.  For example, 
	<code>Response.Write("Enter data:")</code> adds this text to the resulting
	HTML document.
</dl>


<h4>The form</h4>

<p>This is the HTML document that displays the form which will be used to enter data
of a customer.  This is bare HTML and therefore it is also an ASP that contains no scriptlets.</p>

<pre>&lt;HTML&gt;

&lt;HEAD&gt;
  &lt;TITLE&gt; Customer data &lt;/TITLE&gt;
&lt;/HEAD&gt;

&lt;BODY&gt; 

Please fill in the following form:

&lt;FORM METHOD=POST ACTION=&quot;r_new.asp&quot;&gt;

&lt;TABLE&gt;
&lt;TR&gt;&lt;TD&gt;First Name&lt;/TD&gt;
    &lt;TD&gt;&lt;INPUT TYPE=TEXT NAME="FirstName" SIZE=20&gt;&lt;/TD&gt;&lt;/TR&gt;
&lt;TR&gt;&lt;TD&gt;Last Name&lt;/TD&gt;
    &lt;TD&gt;&lt;INPUT TYPE=TEXT NAME="LastName"  SIZE=20&gt;&lt;/TD&gt;&lt;/TR&gt;
&lt;TR&gt;&lt;TD&gt;Phone&lt;/TD&gt;
    &lt;TD&gt;&lt;INPUT TYPE=TEXT NAME="Phone"     SIZE=20&gt;&lt;/TD&gt;&lt;/TR&gt;
&lt;TR&gt;&lt;TD&gt;Address&lt;/TD&gt;
    &lt;TD&gt;&lt;INPUT TYPE=TEXT NAME="Address"   SIZE=40&gt;&lt;/TD&gt;&lt;/TR&gt;
&lt;/TABLE&gt;

&lt;INPUT TYPE=SUBMIT VALUE="Send"&gt;
&lt;INPUT TYPE=RESET VALUE="Clear"&gt;

&lt;/FORM&gt;

&lt;/BODY&gt;
&lt;/HTML&gt;</pre>

<p>Here is script <code>r_new.asp</code> that inserts the customer data sent
by the user by means of a form.</p>

<pre>&lt;HTML&gt;
&lt;HEAD&gt;&lt;TITLE&gt;Processed customer data&lt;/TITLE&gt;&lt;/HEAD&gt; 
&lt;BODY&gt; 

&lt;% 
  SET cnn = <font color="red"><code>Server.CreateObject("ADODB.Connection")</code></font> 
  <font color="red"><code>cnn.Open</code></font>("DSN=LocalServer;database=Shop;UID=Shop;PWD=xyx#123") 
    
  IF (Request.Form("FirstName") = "") OR (Request.Form("LastName") = "")  THEN 
    Response.Write("&lt;BR&gt;One of the fields is empty.") 
  ELSE 
    'Check if the data on this customer is already stored
    SET rst = <font color="#FF0000"&gt;Server.CreateObject(&quot;ADODB.Recordset&quot;</font&gt;) 
    SQL = "SELECT * FROM Customers " 
    SQL = SQL + " WHERE FirstName = " &amp; "'" &amp; Request.Form("FirstName") &amp; "'" 
    SQL = SQL + " AND LastName = " &amp; "'" &amp; Request.Form("LastName") &amp; "'" 
    <font color="red"><code>rst.Open</code></font> SQL, cnn, 3, 3 
   
    IF rst.RecordCount &gt; 0 THEN 
      Response.Write(" &lt;BR&gt;We already have data on this customer.") 
    ELSE 
      Response.Write(" &lt;BR&gt;We do not have data on this customer.") 

      SET rstNew = Server.CreateObject("ADODB.Recordset") 
      SQLNew = "INSERT INTO Customers (FirstName, LastName, Address, Phone) " 
      SQLNew = SQLNew + " VALUES ('" &amp; Request.Form("FirstName") 
      SQLNew = SQLNew + "', '" &amp; Request.Form("LastName") 
      SQLNew = SQLNew + "', '" &amp; Request.Form("Address") 
      SQLNew = SQLNew + "', '" &amp; Request.Form("Phone") 
      SQLNew = SQLNew + "')" 
      <font color="red"><code>rstNew.Open</code></font> SQLNew, cnn, 3, 3 
   
      Response.Write(" &lt;BR&gt;Data successfully inserted.") 
    END IF 
  END IF
%&gt; 

&lt;/BODY&gt; 
&lt;/HTML&gt;</pre>

<hr><h3><a name="Adonet">ADO.NET</a></h3>

<p>In ADO.NET we use the <code>DataSet</code> object to process the result of an
SQL statement.  Its structure is more complex than the corresponding 
the <code>Recordset</code> object of ADO.  In particular a <code>DataSet</code> can
contain a number of tables connected with relationships based pair foreign
key-primary key. SQL statements are represented by objects of class
<code>SqlDataAdapter</code>.</p>

<p>Here is an example usage of <code>DataSet</code>.</p>

<pre>'Establish connection to the database
Dim conn as New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" &amp; _
			        "Data Source=F:\ASPNET\data\library.mdb")
'Create the statement
Dim da As New SqlDataAdapter("select * from Customers", conn)

'Execute the query and fill the data set
 
Dim ds As New DataSet()
da.Fill(ds)</pre>


<h4>Example retrieval</h4>

<p>Data from the database will be displayed on a web page (ASP.NET). The example
is taken from the book by Chris Payne, <i>Sams Teach Yourself ASP.NET in 21 Days</i>.
The access to data from the table is facilitated by objects of class
<code>DataTable</code> (that represents tables) and
<code>DataRow</code> (that represents rows).</p>
 
<pre>&lt;%@ Page Language="VB" %&gt;
&lt;%@ Import Namespace="System.Data" %&gt;
&lt;%@ Import Namespace="System.Data.OleDb" %&gt;

&lt;script runat="server"&gt;

Sub Page_Load(obj as object, e as eventargs)
  Dim objConn as new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" &amp; _
                                     "Data Source=F:\ASPNET\data\banking.mdb")
  Dim objCmd as new OleDbDataAdapter("select * from Customers", objConn)
  Dim ds as DataSet = new DataSet()
  objCmd.Fill(ds, "Customers")

  Dim <font color="red"><code>dTable</code></font> as <font color="red"><code>DataTable</code></font> = ds.Tables("Customers")
  Dim <font color="red"><code>dRows()</code></font> as <font color="red"><code>DataRow</code></font> = dTable.Select(Nothing, Nothing, _
			                       DataViewRowState.CurrentRows)
  Dim I, J as integer
  Dim strOutput as string
  For I = 0 to dRows.Length-1
    For J = 0 to dTable.Columns.Count-1
      strOutput = strOutput &amp; dTable.Columns(J).ColumnName _
                  &amp; " = " &amp; dRows(I)(J).ToString &amp; "&lt;br&gt;"
    Next
  Next

  Response.Write(strOutput)
End Sub

&lt;/script&gt;

&lt;html&gt;
&lt;body&gt;
  &lt;!-- The page created in object <i>Response</i>!--&gt;
&lt;/body&gt;
&lt;/html&gt;</pre>

<h4>Example of update</h4>

<p>In order to execute INSERT, UPDATE and DELETE statements, we use
objects of the <code>OleDbCommand</code> class and their
<code>ExecuteNonQuery</code> method.  Here is the procedure that executes updates.
The <code>strSQL</code> variable  stores the text of the SQL statement, while
<code>Conn</code> is an object of the <code>OleDbConnection</code> class).

<pre>Sub ExecuteStatement(strSQL As String)
  Dim objCmd as new OleDbCommand(strSQL, Conn)

  Try
    objCmd.Connection.Open()
    objCmd.ExecuteNonQuery()
  Catch ex As Exception
    errorText = "An error occurred during update."
  End Try

  objCmd.Connection.Close()
End Sub</pre>

<p>Note the new method of exception handling.  We used clauses
<code>Try ... Catch ... End Try</code> which is similar to Java.</p>

<hr><h3><a name="Podsumowanie">Summary</a></h3>

<p>Inside a database application a programming language is used to:</p>

<ol>
<li>process data that require iteration and choice;
<li>check the correctness of data and remove errors;
<li>response appropriately to the reported errors;
<li>co-operate with other applications and remote databases;
<li>reuse the same code.
</ol>

<p>Lecture 12 presented the method to access databases in programs written 
in imperative programming languages.  This method is based on library ADO and its
objects <i>Connection</i>, <i> Command</i> and <i>Recordset</i> and their properties
and methods.</p>

<p>We mentioned applications of ADO in ASP and ADO.NET.
ASP is a script language that extends HTML. ADO.NET is 
the new library developed by Microsoft and used to access databases.</p>

<hr><h3><a name="Slownik">Dictionary</a></h3>

<dl>

<dt><a href="#Command">Command</a>
<dd>The object that represents the SQL statement to be executed by the data source. 

<dt><a href="#Podstawowe obiekty">Connection</a>
<dd>It is the root of the hierarchy of classes of ADO. 
	The application uses it to connect to a data source. 

<dt><a href="#Uzycie1">Error</a>
<dd>The object that represents errors reported by a data source during 
	the execution of an SQL statement.

<dt><a href="#Podstawowe obiekty">Recordset</a>
<dd>The object that represents the set of all records of a table
or the set of records returned by a query. 

<dt><a href="#Skrypty">server script</a>
<dd>A script run by the Web server. It produces an HTML document that is sent to the user.
	Scripts contain commands that generate tags of HTML and connect to databases.
	The most popular server script languages are  PHP, ASP, ASP.NET and JSP.
<dt><a href="#Transak">transaction</a>
<dd>A sequence of INSERT, UPDATE and DELETE statements that is performed
as a whole. Either all or none of these operations are executed.
</dl>

<hr><h3><a name="Zadania">Exercise</a></h3>

<ol>
<li>Create the <i>Persons</i> table with following columns:
	<ul>
	<li><i>Person_id</i> (Autonumber),
	<li><i>First Name</i> (Text) and 
	<li><i>Last Name</i> (Text).
	</ul>
<li>Create a form that is not based on any table.
<li>Add unbounded fields <i>Person_id</i>, <i>First Name</i> and <i>Last Name</i> to this
	form.
<li>Add three command button to this form:
	<ul>
	<li><i>Display first person</i> (by <i>Person_id</i>).
	<li><i>Display next person</i> (by <i>Person_id</i>).
	<li><i>Delete person</i>.
	</ul>
<li>Write event procedures for these buttons. Use ADO in these procedures.
<li>Add code that handles errors that may be reported during
	database operations.
</ol>
