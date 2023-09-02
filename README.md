# SQL-to-Excel-Returns-Comparison
Queries data from a SQL db and transfers it to an excel sheet while comparing returns with another portfolio/strategy. Then computes comparitive statistcs on the queried data. Very easy to use!

The code’s process can be thought of as 10 steps:

1.	Connects to the SQL database.
2.	Querys the wanted return data from a view I create NMR (Name, Month, Return).
3.	Creates a copy of the Desired Comparison Template
4.	Activates Editing on a new file.
5.	Loads query data into an array.
6.	Matches all dates from the queried data to Pan’s returns by date.
7.	Pastes all data into a new worksheet along with the name of the PM1 in the corresponding places.
8.	Deletes all rows that do not have data in the corresponding column.
9.	Excel-drags all hidden formulas down (Cumm. Return/Drawdowns) to recalibrate after rows were deleted.
10.	Saves the file and closes the connection to the database.
