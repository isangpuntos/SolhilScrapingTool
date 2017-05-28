Imports System.IO
Imports System.Net
Imports System.Runtime.InteropServices
Imports System.Text
Imports HtmlAgilityPack
Imports iTextSharp.text.pdf
Imports iTextSharp.text.pdf.parser
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form1

    Dim applicationList As List(Of String) = New List(Of String)
    Dim summaryHeaders As New Dictionary(Of String, String) From {{"Reference", "A"}, {"Application Validated", "B"}, {"Address", "J"}, {"Proposal", "F"}, {"Status", "G"}}
    Dim detailsHeaders As New Dictionary(Of String, String) From {{"Application Type", "E"}, {"Ward", "H"}, {"Parish", "I"}, {"Applicant Name", "K"}, {"Agent Name", "W"}, {"Agent Company Name", "X"}, {"Agent Address", "Y"}}
    Dim contactsHeaders As New Dictionary(Of String, String) From {{"Work Phone Number", "AA"}, {"Email Address", "Z"}}
    Dim xlApp As Excel.Application
    Dim xlWorkBook As Excel.Workbook
    Dim xlWorkSheet As Excel.Worksheet
    Dim misValue = System.Reflection.Missing.Value

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        WebBrowser1.Navigate(TextBox1.Text)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim LoadedUrl As String = ""
        Try
            LoadedUrl = WebBrowser1.Url.ToString()
        Catch ex As Exception

        End Try
        If LoadedUrl.ToLower.Contains("searchresult") And Not WebBrowser1.IsBusy Then
            xlApp = New Microsoft.Office.Interop.Excel.Application()

            If xlApp Is Nothing Then
                MessageBox.Show("Excel is not properly installed!!")
                Return
            End If
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            '								Target Determination Date	Page URL Link 	Page URL
            xlWorkSheet = DirectCast(xlWorkBook.Worksheets.Item(1), Excel.Worksheet)
            xlWorkSheet.Cells(1, 1) = "Planning Ref"
            xlWorkSheet.Cells(1, 2) = "Date Validated"
            xlWorkSheet.Cells(1, 3) = "Category"
            xlWorkSheet.Cells(1, 4) = "Brief Description"
            xlWorkSheet.Cells(1, 5) = "Application_Type"
            xlWorkSheet.Cells(1, 6) = "Proposal"
            xlWorkSheet.Cells(1, 7) = "Status"
            xlWorkSheet.Cells(1, 8) = "Ward"
            xlWorkSheet.Cells(1, 9) = "Parish"
            xlWorkSheet.Cells(1, 10) = "Site address"
            xlWorkSheet.Cells(1, 11) = "Applicant Name"
            xlWorkSheet.Cells(1, 12) = "Applicant Address"
            xlWorkSheet.Cells(1, 13) = "Applicant Title"
            xlWorkSheet.Cells(1, 14) = "Applicant First Name"
            xlWorkSheet.Cells(1, 15) = "Applicant Last Name"
            xlWorkSheet.Cells(1, 16) = "Applicant Address 1"
            xlWorkSheet.Cells(1, 17) = "Applicant Address 2"
            xlWorkSheet.Cells(1, 18) = "Applicant Address 3"
            xlWorkSheet.Cells(1, 19) = "Applicant Town"
            xlWorkSheet.Cells(1, 20) = "Applicant County"
            xlWorkSheet.Cells(1, 21) = "Applicant Post Code"
            xlWorkSheet.Cells(1, 22) = "Applicant Telephone"
            xlWorkSheet.Cells(1, 23) = "Agent Name"
            xlWorkSheet.Cells(1, 24) = "Agent Company Name"
            xlWorkSheet.Cells(1, 25) = "Agent Address"
            xlWorkSheet.Cells(1, 26) = "Agent email"
            xlWorkSheet.Cells(1, 27) = "Agent Telephone"
            xlWorkSheet.Cells(1, 28) = "Council"
            xlWorkSheet.Cells(1, 29) = "Region"
            xlWorkSheet.Cells(1, 30) = "Target Determination Date"
            xlWorkSheet.Cells(1, 31) = "Page URL link"
            xlWorkSheet.Cells(1, 32) = "Page URL"
            ParseListSearchResults()
            ParseTable()
            xlWorkBook.SaveAs("Trial.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
                              Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            xlWorkBook.Close(True, misValue, misValue)
            xlApp.Quit()
            Marshal.ReleaseComObject(xlWorkSheet)
            Marshal.ReleaseComObject(xlWorkBook)
            Marshal.ReleaseComObject(xlApp)

        End If

    End Sub

    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        WebBrowser1 = TryCast(sender, WebBrowser)
        TextBox1.Text = WebBrowser1.Url.ToString()
        ProgressBar1.MarqueeAnimationSpeed = 0
        Debug.WriteLine(WebBrowser1.Document.Cookie)
    End Sub

    Delegate Sub SetHtmlVoidDelegate([text] As String)

    Public Sub ParseListSearchResults()
        Try
            Dim doc As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlDocument()
            doc.LoadHtml(WebBrowser1.DocumentText)
            Dim totalSearched As String = doc.DocumentNode.SelectSingleNode("//div[@id='searchResultsContainer']/p[@class='pager top']/span[@class='showing']").InnerText
            Dim ctr As Int32 = 1
            Dim totSearch As Int32 = Convert.ToInt32(totalSearched.Substring(totalSearched.IndexOf("of") + 3))

            While applicationList.Count() < totSearch
                If ctr <> 1 Then
                    WebBrowser1.Navigate("https://publicaccess.solihull.gov.uk/online-applications/pagedSearchResults.do?action=page&searchCriteria.page=" + ctr.ToString())
                    While (WebBrowser1.ReadyState <> WebBrowserReadyState.Complete)
                        System.Windows.Forms.Application.DoEvents()
                    End While
                    doc.LoadHtml(WebBrowser1.DocumentText)
                    TextBox1.Text = WebBrowser1.Url.ToString()
                End If
                For Each li As HtmlNode In doc.DocumentNode.SelectNodes("//div[@class='col-a']/ul[@id='searchresults']/li[@class='searchresult']/a")
                    applicationList.Add(li.GetAttributeValue("href", ""))
                Next
                ctr = ctr + 1
            End While
        Catch ex As Exception
            Return
        End Try
    End Sub

    Private Sub ParseTable()
        Dim doc As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlDocument()
        Dim TableIds As String() = {"simpleDetailsTable", "applicationDetails", "agents"}
        Dim TabNames As String() = {"summary", "details", "contacts"}
        Dim divForAgent As String = ""
        Dim attribForAgent As String = "id"
        Dim rowCell As Integer = 2
        For Each app As String In applicationList
            For itr As Integer = 0 To TableIds.Count() - 1
                Try
                    Dim tableUrl = "https://publicaccess.solihull.gov.uk/online-applications/applicationDetails.do?activeTab=" + TabNames(itr) + "&keyVal=" + app.Substring(app.IndexOf("keyVal=") + 7, 13)
                    ServicePointManager.Expect100Continue = True
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
                    ServicePointManager.ServerCertificateValidationCallback = Function() True
                    If TabNames(itr) = "contacts" Then
                        divForAgent = "/div[@class='agents']"
                        attribForAgent = "class"
                    Else
                        divForAgent = ""
                        attribForAgent = "id"
                    End If
                    Using client As New WebClient
                        Dim htmlCode As String = client.DownloadString(tableUrl)
                        doc.LoadHtml(htmlCode)
                    End Using
                    For Each table As HtmlNode In doc.DocumentNode.SelectNodes("//div[@class='tabcontainer']" + divForAgent + "/table[@" + attribForAgent + "='" + TableIds(itr) + "']")
                        For Each row As HtmlNode In table.SelectNodes("tr")
                            Dim tableData As String = row.SelectSingleNode("th").InnerText.Trim()
                            If summaryHeaders.ContainsKey(tableData) And TabNames(itr) = "summary" Then
                                Dim textValue As String = row.SelectSingleNode("td").InnerText.Trim
                                xlWorkSheet.Cells(rowCell, summaryHeaders.Item(tableData)) = textValue
                                Debug.WriteLine(textValue)
                            ElseIf detailsHeaders.ContainsKey(tableData) And TabNames(itr) = "details" Then
                                Dim textValue As String = row.SelectSingleNode("td").InnerText.Trim
                                xlWorkSheet.Cells(rowCell, detailsHeaders.Item(tableData)) = textValue
                                Debug.WriteLine(textValue)
                            ElseIf contactsHeaders.ContainsKey(tableData) And TabNames(itr) = "contacts" Then
                                Dim textValue As String = row.SelectSingleNode("td").InnerText.Trim
                                xlWorkSheet.Cells(rowCell, contactsHeaders.Item(tableData)) = textValue
                                Debug.WriteLine(textValue)
                            End If
                        Next
                    Next
                Catch ex As Exception

                End Try
            Next
            ParsePdf("https://publicaccess.solihull.gov.uk/online-applications/applicationDetails.do?activeTab=documents&keyVal=" + app.Substring(app.IndexOf("keyVal=") + 7, 13), rowCell)
            rowCell += 1
            Debug.WriteLine("")
        Next
    End Sub

    Private Sub ParsePdf(docUrl As String, rowCell As Integer)
        Try
            Dim doc As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlDocument()
            Dim pdfLink As String = ""
            Dim checkBoxValue As String = ""
            Dim text As StringBuilder = New StringBuilder()
            ServicePointManager.Expect100Continue = True
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            ServicePointManager.ServerCertificateValidationCallback = Function() True

            Dim client As New WebClientEx(New CookieContainer())
            Dim htmlCode As String = client.DownloadString(docUrl)
            doc.LoadHtml(htmlCode)
            htmlCode = doc.DocumentNode.SelectNodes("//div[@id='idox']/div[@id='pa']/div[@class='container']/div[@class='content']/div[@class='tabcontainer toplevel']")(0).InnerHtml
            htmlCode = htmlCode.Substring(htmlCode.IndexOf("<table "), htmlCode.IndexOf("</table>") + 8 - htmlCode.IndexOf("<table "))
            doc.LoadHtml(htmlCode)
            For Each table As HtmlNode In doc.DocumentNode.SelectNodes("table")
                For Each row As HtmlNode In table.SelectNodes("tr")
                    If row.InnerHtml.Trim().StartsWith("<td>") Then
                        Dim x As String = row.SelectNodes("td")(2).InnerText
                        If row.SelectNodes("td")(2).InnerText = "Application Form" Then
                            pdfLink = "https://publicaccess.solihull.gov.uk" + row.SelectNodes("td")(5).SelectSingleNode("a").Attributes("href").Value
                            checkBoxValue = row.SelectNodes("td")(0).SelectSingleNode("input").Attributes("value").Value
                        End If
                    End If
                Next
            Next

            If pdfLink IsNot "" Then
                client.DownloadFile(New Uri(pdfLink), "appForm.pdf")
                Dim pdfReader As PdfReader = New PdfReader("appForm.pdf")

                For page As Integer = 1 To pdfReader.NumberOfPages
                    Dim strategy As ITextExtractionStrategy = New SimpleTextExtractionStrategy()
                    Dim currentText As String = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy)
                    currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)))
                    text.Append(currentText)
                Next
                pdfReader.Close()

                Dim fullText As String = text.ToString()
                Dim title As String = (fullText.Substring(fullText.IndexOf("Title: ") + 7, fullText.IndexOf("First Name:") - (fullText.IndexOf("Title: ") + 7)))
                Dim fName As String = (fullText.Substring(fullText.IndexOf("First Name: ") + 12, fullText.IndexOf("Surname:") - (fullText.IndexOf("First Name: ") + 12)))
                Dim sName As String = (fullText.Substring(fullText.IndexOf("Surname: ") + 9, fullText.IndexOf("Company name:") - (fullText.IndexOf("Surname: ") + 9)))
                Dim address As String = (fullText.Substring(fullText.IndexOf("Street address: ") + 16, fullText.IndexOf("Telephone number:") - (fullText.IndexOf("Street address: ") + 16)))
                Dim townCity As String = (fullText.Substring(fullText.IndexOf("Town/City: ") + 11, fullText.IndexOf("Fax number:") - (fullText.IndexOf("Town/City: ") + 11)))
                Dim Country As String = (fullText.Substring(fullText.IndexOf("Country: ") + 9, fullText.IndexOf("Email address:") - (fullText.IndexOf("Country: ") + 9)))
                Dim postCode As String = (fullText.Substring(fullText.IndexOf("Postcode: ") + 10, fullText.IndexOf("Are you an agent acting on behalf of the applicant?") - (fullText.IndexOf("Postcode: ") + 10)))

                xlWorkSheet.Cells(rowCell, "M") = title.Replace(vbCr, "").Replace(vbLf, "")
                xlWorkSheet.Cells(rowCell, "N") = fName.Replace(vbCr, "").Replace(vbLf, "")
                xlWorkSheet.Cells(rowCell, "O") = sName.Replace(vbCr, "").Replace(vbLf, "")
                xlWorkSheet.Cells(rowCell, "P") = address.Split(vbLf)(0)
                xlWorkSheet.Cells(rowCell, "Q") = If(address.Split(vbLf).Count() > 1, address.Split(vbLf)(1), "")
                xlWorkSheet.Cells(rowCell, "R") = If(address.Split(vbLf).Count() > 2, address.Split(vbLf)(2), "")
                xlWorkSheet.Cells(rowCell, "S") = townCity.Replace(vbCr, "").Replace(vbLf, "")
                xlWorkSheet.Cells(rowCell, "T") = Country.Replace(vbCr, "").Replace(vbLf, "")
                xlWorkSheet.Cells(rowCell, "U") = postCode.Replace(vbCr, "").Replace(vbLf, "")

            End If
        Catch ex As Exception

        End Try
        System.IO.File.Delete("appForm.pdf")
    End Sub

End Class
