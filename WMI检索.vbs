set ie = createobject("internetexplorer.application")
ie.toolbar = 0:ie.navigate "about:blank"
url = inputbox("Select * from ","请输入查询的目标","Win32_BIOS")
if url = "" then wscript.quit
ie.visible = 1
ie.document.write "<title>" & url & "</title>"
for each i in getobject("winmgmts:\\.\root\cimv2").execquery("select * from " & url)
  with ie.document
    .write "<hr>"
    .write "<table><tr valign=top><td colspan=2>"
    .write "<font color=red><b>" & i.path_.path & "</b></font>"
    .write "</td></tr><tr valign=top><td>"
    .write "<table border=1>"
    .write "<tr><td colspan=2><font color=blue><b>属性</b></font></td></tr>"
    for each j in i.properties_
      .write "<tr><td>"
      .write j.name
      .write "</td>"
      .write "<td>"
      if typename(j.value) = "Variant()" then
        for each k in j.value
          .write k & "<br>"
        next
      else
        .write j.value
      end if
      .write "</td></tr>"
    next
    .write "</table>"
    .write "</td><td>"
    .write "<table border=1>"
    .write "<tr><td><font color=blue><b>方法</b></font></td></tr>"
    for each j in i.methods_
      .write "<tr><td>"
      .write j.name
      .write "</td></tr>"
    next
    .write "</table>"
    .write "</td></tr></table>"
  end with
next