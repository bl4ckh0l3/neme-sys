<%
sub generateTree(path, elementsMap)
	set fso = CreateObject("Scripting.FileSystemObject")
	vDir = path
	root = Server.MapPath(vDir) & "\"
	set fold = fso.getFolder(root)
	
	if fold.subfolders.count > 0 then
		for each f in fold.subfolders
			fpath = vDir & f.name & "/"			
			call generateTree(fpath, elementsMap)
		next
	end if
	if fold.files.count > 0 then
		for each fil in fold.files	
			elementsMap.add fil.name, vDir
		next
	end if	
	set fso = nothing
end sub


function getfoldlink(d, c, f, p)
	if d <> "" then
	
		' needs to be clickable
		getfoldlink = "<a href='#' style='cursor:hand' " & _
			"onclick='flip(""" & d & """);" & _
			"this.blur();return false;'>" & _
			"<img id='i" & d & "' class=" & c & _
			" src=img/plus.gif vspace=0 hspace=2 border=0>" & _
			"<img src=img/folder.gif hspace=2 border=0></a>&nbsp;" & _
			"<a target=_blank href=" & p & getsftitle(f) & _
			">" & f.name & "</a></div><div id='" & d & "'" & _
			" display=none style='display:none'>"
	else
	
		' can't be clickable
		getfoldlink = "<div><img id='i" & d & "' " & _
			"class=" & c & " src=img/plus.gif vspace=0 " & _
			"hspace=2 visibility=hidden style='visibility:hidden'><img" & _
			" src=img/folder.gif hspace=2>&nbsp;<a " & _
			"target=_blank href=" & p & getsftitle(f) & _
			">" & f.name & "</a></div>"
	end if
end function

function getfoldlinkname(d, c, f, p)
	if d <> "" then
	
		' needs to be clickable
		getfoldlinkname = "<a href='#' style='cursor:hand' " & _
			"onclick='flip(""" & d & """);" & _
			"this.blur();return false;'>" & _
			"<img id='i" & d & "' class=" & c & _
			" src=img/plus.gif vspace=0 hspace=2 border=0>" & _
			"<img src=img/folder.gif hspace=2 border=0></a>&nbsp;" & f.name & _
			"</div><div id='" & d & "'" & _
			" display=none style='display:none'>"
	else
	
		' can't be clickable
		getfoldlinkname = "<div><img id='i" & d & "' " & _
			"class=" & c & " src=img/plus.gif vspace=0 " & _
			"hspace=2 visibility=hidden style='visibility:hidden'><img" & _
			" src=img/folder.gif hspace=2>&nbsp;" & f.name & "</div>"
	end if
end function

function getfilelink(c, fold, file)
	getfilelink = "<div><img class=" & c & " src=img/file.gif" & _
		" hspace=2>&nbsp;<a href=" & fold & file.name & _
		getfiletitle(file) & ">" & file.name & "</a></div>"
end function

function getfilename(c, fold, file)
	getfilename = "<div><img class=" & c & " src=img/file.gif" & _
		" hspace=2>&nbsp;" & file.name & "</div>"
end function

function getfiletitle(file)
	getfiletitle = " title='Size: " & _
		formatnumber(file.size/1024, 2, -1, 0, -1) & _
		" kb" & vbCrLf & getDL(file) & "'"
end function

function getsftitle(fold)
	getsftitle = " title='" & getsfc(fold) & _
	vbCrLf & getfc(fold) & _
	vbCrLf & getfs(fold) & _
	vbCrLf & getDL(fold) & "'"
end function

function getDL(o)
	d = o.dateLastModified
	getDL = "Last mod: " & formatdatetime(d, 2) & _
		" " & formatdatetime(d, 3)
end function

function getfc(fold)
	getfc = fCount(fold.files.count)
end function

function getsfc(fold)
	getsfc = sfCount(fold.subfolders.count)
end function

function getfs(fold)
	getfs = "Size: " & bToMB(fold.size)
end function 

function bToMB(b)
	bToMB = formatnumber(b/1024/1024, 2, -1, 0, -1) & " MB"
end function

function fCount(c)
	fCount = formatnumber(c, 0, -1, 0, -1) & " file" & _
		suffix(c)
end function
	
function sfCount(c)
	sfCount = formatnumber(c, 0, -1, 0, -1) & _
		" subfolder" & suffix(c)
end function

function suffix(c)
	if c <> 1 then suffix = "s"
end function
%>
