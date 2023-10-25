'Open application after closing in the end
Systemutil.CloseProcessByName "FlightsGUI.exe"
Systemutil.Run "C:\Program Files (x86)\Micro Focus\UFT One\samples\Flights Application\FlightsGUI.exe"

'screenshot before logging in
WpfWindow("Micro Focus MyFlight Sample").CaptureBitmap "C:\Users\kuchi\Desktop\Desktop\SQC\Screenshots\beforelogin" & DataTable("iteration", dtGlobalSheet) & ".png", True

'login screen
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set DataTable("username", dtGlobalSheet) @@ hightlight id_;_-232978912_;_script infofile_;_ZIP::ssf110.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("password").SetSecure DataTable("password", dtGlobalSheet) @@ hightlight id_;_-232971808_;_script infofile_;_ZIP::ssf111.xml_;_

'standard checkpoint for the "OK" button
WpfWindow("Micro Focus MyFlight Sample").WpfButton("OK").Check CheckPoint("OK_3") @@ hightlight id_;_1972983912_;_script infofile_;_ZIP::ssf112.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("OK").Click @@ hightlight id_;_2079076408_;_script infofile_;_ZIP::ssf113.xml_;_

'screenshot after logging in
WpfWindow("Micro Focus MyFlight Sample").CaptureBitmap "C:\Users\kuchi\Desktop\Desktop\SQC\Screenshots\afterlogin" & DataTable("iteration", dtGlobalSheet) & ".png", True

'Bitmap checkpoint to track change of image
WpfWindow("Micro Focus MyFlight Sample").WpfObject("Seattle to San Francisco,").Check CheckPoint("Seattle to San Francisco,  all inclusive_3") @@ hightlight id_;_1970545072_;_script infofile_;_ZIP::ssf114.xml_;_

'Text checkpoint to track "John Smith
WpfWindow("Micro Focus MyFlight Sample").WpfObject("John Smith").Check CheckPoint("John Smith_3") @@ hightlight id_;_1970534568_;_script infofile_;_ZIP::ssf115.xml_;_

'Standard checkpoint for "Find Flights" button
WpfWindow("Micro Focus MyFlight Sample").WpfButton("FIND FLIGHTS").Check CheckPoint("FIND FLIGHTS_3") @@ hightlight id_;_2111916056_;_script infofile_;_ZIP::ssf116.xml_;_

'screenshot before finding flights 
WpfWindow("Micro Focus MyFlight Sample").CaptureBitmap "C:\Users\kuchi\Desktop\Desktop\SQC\Screenshots\beforeFindingFlights" & DataTable("iteration", dtGlobalSheet) & ".png", True

'Entering the flight details: Departure location, Destination, Tickets, Date and Class
WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("fromCity").Select DataTable("fly_from", dtGlobalSheet) @@ hightlight id_;_2092006416_;_script infofile_;_ZIP::ssf120.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("toCity").Select DataTable("fly_to", dtGlobalSheet) @@ hightlight id_;_2092007520_;_script infofile_;_ZIP::ssf124.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfImage("WpfImage_4").Click 135,98 @@ hightlight id_;_1974419096_;_script infofile_;_ZIP::ssf125.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("Class").Select DataTable("class", dtGlobalSheet) @@ hightlight id_;_1974421880_;_script infofile_;_ZIP::ssf130.xml_;_
' Get the value from the DataTable
numOfTickets = DataTable("Tickets", dtGlobalSheet)

' Check if the number of tickets is greater than 5
If numOfTickets > 5 Then
    ' If greater than 5, set it to 1 in the DataTable
    DataTable("Tickets", dtGlobalSheet) = 1

    ' Fail the test step
    Reporter.ReportEvent micFail, "Number of Tickets Exceeded", "Number of tickets exceeded the limit and has been set to 1 in the DataTable."
End If
WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("numOfTickets").Select DataTable("Tickets", dtGlobalSheet) @@ hightlight id_;_1972987512_;_script infofile_;_ZIP::ssf132.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("FIND FLIGHTS").Click

'screenshot after finding flights 
WpfWindow("Micro Focus MyFlight Sample").CaptureBitmap "C:\Users\kuchi\Desktop\Desktop\SQC\Screenshots\afterFindingFlights" & DataTable("iteration", dtGlobalSheet) & ".png", True

'To search and select a flight after entering flight parameters
WpfWindow("Micro Focus MyFlight Sample").WpfTable("flightsDataGrid").SelectCell 0,1 @@ hightlight id_;_1974422456_;_script infofile_;_ZIP::ssf134.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("SELECT FLIGHT").Click @@ hightlight id_;_1974422168_;_script infofile_;_ZIP::ssf135.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("passengerName").Set DataTable("Passenger", dtGlobalSheet) @@ hightlight id_;_1974423656_;_script infofile_;_ZIP::ssf136.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("ORDER").Click @@ hightlight id_;_2111917832_;_script infofile_;_ZIP::ssf137.xml_;_

'Screenshot after ordering the flight
WpfWindow("Micro Focus MyFlight Sample").CaptureBitmap "C:\Users\kuchi\Desktop\Desktop\SQC\Screenshots\afterOrder" & DataTable("iteration", dtGlobalSheet) & ".png", True

'To start a new search
WpfWindow("Micro Focus MyFlight Sample").WpfButton("NEW SEARCH").Click @@ hightlight id_;_2111917928_;_script infofile_;_ZIP::ssf138.xml_;_

'To close application
WpfWindow("Micro Focus MyFlight Sample").Close @@ hightlight id_;_3936182_;_script infofile_;_ZIP::ssf139.xml_;_

'All test cases have been designed to pass to avoid termination of the iterations.
