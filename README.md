### Various examples of Excel code

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

We see it returns nothing for the last value, which was bad data.  We're unlikely to recieve a false positive result with the method given because we've got the majority of the address.  One probable issue is the above method probably will not scale to the full 4000 address list without adding a wait state.
