<%
Set xml = Server.CreateObject("MSXML2.ServerXMLHTTP")
url = "http://localhost/trydropdown/retrivedata.asp" & filename
xml.Open "GET", url, False
xml.Send
redim arr(0)
Set NodeList = xml.responseXml.selectNodes("//Customers/Customer")

Response.write("SET: "&NodeList.length&"<br>")
response.write(" <table border=1 >")
    response.write("<tr>")
    response.write("<th>Customer ID</th>")
    response.write("<th>Name</th>")
    response.write("<th>Country</th>")
dim scan:scan=0
            dim baseNode
            for scan=0 to NodeList.length-1
                set baseNode =      NodeList(scan)
                if not (baseNode Is Nothing) then
                    response.write("<tr>")
                    response.Write("<td>" &  baseNode.selectSingleNode("customerid").text &" </td>")
                    response.Write("<td> " &  baseNode.selectSingleNode("name").text &" </td>")
                    response.Write("<td>  " &  baseNode.selectSingleNode("country").text &" </td>")
                    response.write("</tr>")
                 
 
                else
                    response.Write(" basenode missing<br>")
                end if
            next

response.write("</table>")

%>