Sub scrape_quotes_dual()
    Dim ie As InternetExplorer
    Dim doc As HTMLDocument
    Dim classElements As Object
    Dim paragraphElement As Object
    Dim strongElements As Object
    Dim strongElement As Object
    Dim urlRange As Range
    Dim urlCell As Range
    
    ' Define the range of cells containing the URLs
    Set urlRange = Range("A1:A3") ' Update the range as per your requirement
    
    Set ie = New InternetExplorer
    ie.Visible = True
    
    ' Loop through each URL in the range
    For Each urlCell In urlRange
        ie.navigate urlCell.value
        
        ' Wait for the webpage to finish loading
        Do While ie.Busy Or ie.readyState <> 4: DoEvents: Loop
        
        Set doc = ie.document
        
        ' Pull elements from a specific class
        Set classElements = doc.getElementsByClassName("column-half")
        
        ' Extract values from the first and second <strong> tags within the first <p> tag of the first class element
        Dim extractedData1 As String
        Dim extractedData2 As String
        extractedData1 = ""
        extractedData2 = ""
        
        If Not classElements Is Nothing Then
            Dim classElement As Object
            Set classElement = classElements.Item(0) ' Retrieve the first element from the collection
            
            ' Retrieve the first <p> tag within the class element
            Set paragraphElement = classElement.getElementsByTagName("p").Item(0)
            
            If Not paragraphElement Is Nothing Then
                ' Retrieve all <strong> tags within the paragraph element
                Set strongElements = paragraphElement.getElementsByTagName("strong")
                
                If Not strongElements Is Nothing Then
                    ' Retrieve the first <strong> tag
                    Set strongElement = strongElements.Item(0)
                    If Not strongElement Is Nothing Then
                        extractedData1 = strongElement.innerText
                    End If
                    
                    ' Retrieve the second <strong> tag
                    Set strongElement = strongElements.Item(1)
                    If Not strongElement Is Nothing Then
                        extractedData2 = strongElement.innerText
                    End If
                End If
            End If
        End If
        
        ' Output the extracted data to adjacent columns
        Cells(urlCell.Row, 3).value = extractedData1
        Cells(urlCell.Row, 4).value = extractedData2
    Next urlCell
    
    ' Clean up
    ie.Quit
    Set ie = Nothing
End Sub
