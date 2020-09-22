<div align="center">

## Dynamically Add WebBrowser Control at runtime without a Reference \(V\. 2\)

<img src="PIC200541321873471.gif">
</div>

### Description

Allows VB applications to determine at run-time if Internet Explorer (4.0 or later) is installed, and if so, creates a WebBrowser. If not, a trappable error allows program to continue. This is an enhanced version of a previous article, showing how to capture the events of the created object.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dave Slinn](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dave-slinn.md)
**Level**          |Beginner
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dave-slinn-dynamically-add-webbrowser-control-at-runtime-without-a-reference-v-2__1-59979/archive/master.zip)





### Source Code

<font face="Tahoma" size="2"><p>On a VB6 Form, add a Button control in the upper
left corner and a Listbox control underneath the button and down the left hand
side of the form.&nbsp; Events raised by the WebBrowser will be displayed in the
listbox.</p>
<p>Place the following code into a standard VB 6.0 form.</p></font>
<hr>
<p><font face="Courier New" size="2"><font color="#0000FF">Private WithEvents
</font>m_WebControl <font color="#0000FF">As</font> VBControlExtender</font></p>
<p><font face="Courier New" size="2"><font color="#0000FF">Private Sub</font> Form_Resize()<br>
<font color="#0000FF">On Error Resume Next</font><br>
&nbsp;&nbsp; Me.List1.Height = Me.ScaleHeight - Me.List1.Top<br>
<br>
&nbsp;&nbsp; <font color="#008000">' resize webbrowser to fill form next to listbox</font><br>
&nbsp;&nbsp; <font color="#0000FF">If Not </font>m_WebControl
<font color="#0000FF">Is Nothing Then</font><br>
&nbsp;&nbsp;&nbsp;&nbsp; m_WebControl.Move Me.List1.Left + Me.List1.Width + 30, 0, ScaleWidth - (Me.List1.Left + Me.List1.Width + 30), ScaleHeight<br>
&nbsp;&nbsp; <font color="#0000FF">End If</font><br>
<font color="#0000FF">End Sub</font><br>
<br>
<font color="#0000FF">Private Sub </font>Command1_Click()<br>
<font color="#0000FF">On Error GoTo </font>ErrHandler<br>
<br>
&nbsp;&nbsp; ' attempting to add WebBrowser here ('Shell.Explorer.2' is registered<br>
&nbsp;&nbsp; ' with Windows if a recent (&gt;= 4.0) version of Internet Explorer is installed<br>
&nbsp;&nbsp; <font color="#0000FF">Set </font>m_WebControl = Controls.Add(&quot;Shell.Explorer.2&quot;, &quot;webctl&quot;, Me)<br>
<br>
&nbsp;&nbsp; <font color="#008000">' if we got to here, there was no problem creating the WebBrowser</font><br>
&nbsp;&nbsp; <font color="#008000">' so we should size it properly and ensure it's visible</font><br>
&nbsp;&nbsp; m_WebControl.Move Me.List1.Left + Me.List1.Width + 30, 0, ScaleWidth - (Me.List1.Left + Me.List1.Width + 30), ScaleHeight<br>
&nbsp;&nbsp; m_WebControl.Visible = <font color="#0000FF">True</font><br>
<br>
&nbsp;&nbsp; <font color="#008000">' use the Navigate method of the WebBrowser control to open a</font><br>
&nbsp;&nbsp; <font color="#008000">' web page</font><br>
&nbsp;&nbsp; m_WebControl.object.navigate &quot;http://www.planet-source-code.com&quot;<br>
<br>
<font color="#0000FF">Exit Sub</font><br>
ErrHandler:<br>
&nbsp;&nbsp; MsgBox &quot;Could not create WebBrowser control&quot;, vbInformation<br>
<font color="#0000FF">End Sub</font><br>
<br>
<font color="#0000FF">Private Sub </font>m_WebControl_ObjectEvent(Info
<font color="#0000FF">As</font> EventInfo)<br>
<font color="#0000FF">On Error GoTo</font> ErrHandler<br>
<br>
&nbsp;&nbsp; <font color="#0000FF">Dim</font> i <font color="#0000FF">As Integer</font><br>
&nbsp;&nbsp; <font color="#0000FF">Dim</font> evp <font color="#0000FF">As</font> EventParameter<br>
<br>
&nbsp;&nbsp; <font color="#008000">' display the event that was raised in the listbox</font><br>
&nbsp;&nbsp; Me.List1.AddItem &quot;Event Raised: &quot; &amp; Info.Name<br>
&nbsp;&nbsp; <font color="#0000FF">For Each </font>evp <font color="#0000FF">In</font> Info.EventParameters<br>
&nbsp;&nbsp; &nbsp;&nbsp; Me.List1.AddItem &quot; &quot; &amp; evp.Name &amp; &quot; (&quot; &amp; evp.Value &amp; &quot;)&quot;<br>
&nbsp;&nbsp; <font color="#0000FF">Next</font> evp<br>
<br>
&nbsp;&nbsp; Me.List1.ListIndex = Me.List1.NewIndex<br>
<font color="#0000FF">Exit Sub</font><br>
ErrHandler:<br>
&nbsp;&nbsp; <font color="#0000FF">If </font>Err.Number = -2147024809
<font color="#0000FF">Then</font><br>
&nbsp;&nbsp; &nbsp;&nbsp; Me.List1.AddItem &quot; &quot; &amp; evp.Name &amp; &quot; (#ERROR)&quot;<br>
&nbsp;&nbsp; &nbsp;&nbsp; <font color="#0000FF">Resume Next</font><br>
&nbsp;&nbsp; <font color="#0000FF">End If</font><br>
<font color="#0000FF">End Sub</font></p>

