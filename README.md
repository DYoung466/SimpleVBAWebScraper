# VBA WebScraper

This is an example of a webscraper, which can iterate through and pull specific values from a webpage and place them into an Excel document. It is currently set to pull account date values from a companies Gov.uk site. The code will iterate through a cell range of company urls, and retrieve the end date of the account as well as its due date. 

## How to run
### Required Excel references:

* Microsft Office HTML Object Library
* Microsft Internet controls

[Setting up Visual Basic modules and references in Microsoft Excel](https://oxylabs.io/blog/web-scraping-excel-vba)

The main.vba file contains the visual basic script. This should be placed within a new module in a Visual Basic project in Excel. 
</br>`Set urlRange = Range("A1:A3")` - This is the parameter range of the urls. </br>A company url currently is defined as: </br>`https://find-and-update.company-information.service.gov.uk/company/{company number}`. 
</br>To change the output range, simply amend these parameters at the end of the script: `Cells(urlCell.Row, 3).value = extractedData1`.

