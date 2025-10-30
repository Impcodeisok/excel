# Various examples of Excel code

## Missing Zip Codes

User has adress data missing zipcode information and needed to validate the data and generate a correct Zip code.

Below is the sample data I used to test the method.

![alt text](https://github.com/Impcodeisok/excel/blob/main/FZFinisheddata.jpg "Source Example")

Since there is not a 1 to 1 relationship between zipcode and town/city a simple xlookup wouldn't suffice.

If we manually enter the partial address into google maps it returns the zip code, but the user has 4000 adresses to do so we need to automate this process.

Using power query I created a custom function to lookup the address from the data we did have and then return the full adress data that google was responding with.

``` Mcode
let check = (chk) =>
    let
    Source = Web.BrowserContents("https://www.google.com/maps/place/"&chk),
    #"Extracted Table From Html" = Html.Table(Source, {{"Maps says", ".JpCtJf:nth-last-child(3)"}})
    in
    #"Extracted Table From Html"
in check
```

Then we apply this function to our source data as a new column and we get the following.
![alt text](https://github.com/Impcodeisok/excel/blob/main/FZFinisheddata.jpg "Example output")

We see it returns nothing for the last value, which was bad data.  We're unlikely to recieve a false positive result with the method given because we've got the majority of the address.  One probable issue is the above method probably will not scale to the full 4000 address list without adding a wait state. See https://medium.com/@AndreAlessi/building-delays-into-power-bi-api-queries-function-invokeafter-and-google-maps-api-68b475c73a2c for what that would involve.

[Example File](https://github.com/Impcodeisok/excel/blob/main/Find%20zip.xlsx "Find Zip example")


## Reformatting Data example:

Data from https://foresightbi.com.ng/microsoft-power-bi/dirty-data-samples-to-practice-on/ Example 4.
Data needed to be reformatted from the existing messy format to a more compact and functional form as seen in the Image below.

![alt text](https://github.com/Impcodeisok/excel/blob/main/goal.jpg "Data to reformat"
)
Although Power Query is the better option for fun I tried resolving our format with formulas first.  In a small dataset like this it can sometimes be the right answer if you need something quick as a one off.

First, I created a copy of the data on a new sheet named “Dirty copy” and modified it manually to show a ship mode value for every relevant column to make life simpler.

From here I copy the order and date columns into a new sheet called “Formula method”

Then I created a column for Ship mode, Segment and Sales mirroring the desired format.
For finding the Ship mode for each order ID we combine xlookup and filter.

The xlookup portion of the formula “XLOOKUP($C2&$D2,'Dirty copy'!$A:$A&'Dirty copy'!$B:$B,'Dirty copy'!$C:$N)>0)” returns a true and false array for each column in the original format, giving us a way to then filter for the correct Ship mode and Segment, respectively from rows 2/3 of our source data.

The full formula is “=FILTER('Dirty copy'!$C$2:$N$2,XLOOKUP($C2&$D2,'Dirty copy'!$A:$A&'Dirty copy'!$B:$B,'Dirty copy'!$C:$N)>0)” for Ship mode
And
“=FILTER('Dirty copy'!$C$3:$N$3,XLOOKUP($C2&$D2,'Dirty copy'!$A:$A&'Dirty copy'!$B:$B,'Dirty copy'!$C:$N)>0)” for Segment.

The final step is return the value of the given order, for which we use 
“=FILTER(XLOOKUP($C2&$D2,'Dirty copy'!$A:$A&'Dirty copy'!$B:$B,'Dirty copy'!$C:$N),XLOOKUP($C2&$D2,'Dirty copy'!$A:$A&'Dirty copy'!$B:$B,'Dirty copy'!$C:$N)<>0)”

Once we’ve gotten our results I copy the raw data and paste as value into the sheet “Formula method final output” and format it as a table, apply sorting and would typically then copy the worksheet to a new workbook to forward to the stakeholder who needed it.

But the more correct method here is probably using Power Query to reformat the data.  It handles larger data sets more efficiently and gives you a repeatable process you can easily automate if a task is common enough to warrant it.

For that method we select our data on “Dirty copy” and select data/data from table range.

We remove the two columns at the top and promote our next row to headers.

Next I change data type for “Order Date” to date type and remove errors to purge the grand total row at the end of our data.

I replaced null with 0 in column 3-14 so when I unpivot the columns the step will succeed since it rejects nulls.
Before that I rename column 3-14 to reflect the Ship mode and Segment type as a Dash delimited value.
Then we select order ID and Order Data and unpivot other columns.
This reshapes our data as below

![alt text](https://github.com/Impcodeisok/excel/blob/main/pivot.jpg "Unpivot")

Then we split the Ship mode and Segment value into two separate columns and filter sales to remove all values of 0
Lastly, we apply the desired sorting and convert Sales to a decimal value.
We then close and load and the results are output on the “Power query Method” worksheet.
Full PQ code as follows:

let
    Source = Excel.CurrentWorkbook(){[Name="Table13"]}[Content],
    #"Removed Top Rows" = Table.Skip(Source,2),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Order Date", type date}}),
    #"Removed Errors" = Table.RemoveRowsWithErrors(#"Changed Type", {"Order Date"}),
    #"Replaced Value" = Table.ReplaceValue(#"Removed Errors",null,0,Replacer.ReplaceValue,{"Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14"}),
    #"Renamed Columns" = Table.RenameColumns(#"Replaced Value",{{"Column3", "First Class - Consumer"}, {"Column4", "First Class - Corporate"}, {"Column5", "First Class - Home Office"}, {"Column6", "Same Day - Consumer"}, {"Column7", "Same Day - Corporate"}, {"Column8", "Same Day - Home Office"}, {"Column9", "Second Class - Consumer"}, {"Column12", "Standard Class - Consumer"}, {"Column10", "Second Class - Corporate"}, {"Column11", "Second Class - Home Office"}, {"Column13", "Standard Class - Corporate"}, {"Column14", "Standard Class - Home Office"}}),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Renamed Columns", {"Order ID", "Order Date"}, "Attribute", "Value"),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Unpivoted Other Columns", "Attribute", Splitter.SplitTextByDelimiter(" - ", QuoteStyle.Csv), {"Attribute.1", "Attribute.2"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Split Column by Delimiter",{{"Attribute.1", "Shipe Mode"}, {"Attribute.2", "Segment"}, {"Value", "Sales"}}),
    #"Filtered Rows" = Table.SelectRows(#"Renamed Columns1", each ([Sales] <> 0)),
    #"Reordered Columns" = Table.ReorderColumns(#"Filtered Rows",{"Shipe Mode", "Segment", "Order ID", "Order Date", "Sales"}),
    #"Sorted Rows" = Table.Sort(#"Reordered Columns",{{"Shipe Mode", Order.Ascending}, {"Segment", Order.Ascending}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Sorted Rows",{{"Sales", type number}})
in
    #"Changed Type1"

File with all data and solutions:

[Example File](https://github.com/Impcodeisok/excel/blob/main/4.-Badly-Structured-Sales-Data-4.xlsx "Badly structured sales data")



