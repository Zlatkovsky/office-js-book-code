Excel.run(function(context) {

        // Create proxy objects to represent the "real" workbook objects that we'll be working with.  More information on proxy objects  will be presented in the very next section of this chapter.

        var table = context.workbook.tables.getItem("PopulationTable");

        var nameColumn = table.columns.getItem("City");
        var latestPopulationColumn = table.columns.getItem(
            "7/1/2014 population estimate");
        var earliestCensusColumn = table.columns.getItem(
            "4/1/1990 census population");

        // Now, load the values for each of the three columns that we want to read from.  Note that, to support batching operations together (again, you'll see more in the upcoming sections of this chapter), the load doesn't *actually* happen until we do a "context.sync()", as below.

        nameColumn.load("values");
        latestPopulationColumn.load("values");
        earliestCensusColumn.load("values");

        return context.sync()
            .then(function() {
                // Create an in-memory representation of the data, using an array that will contain JSON objects representing each city
                var cityData = [];

                // Start at i = 1 (that is, 2nd row of the table -- remember the 0-indexing) in order to skip the header.
                for (var i = 1; i < nameColumn.values.length; i++) {
                    // A couple of the cities don't have data for  1990, so skip over those.

                    // Note that because the "values" is a 2D array (even though, in this particular case, it's just a single column), need to extract out the 0th element of each row.
                    var population1990 = earliestCensusColumn.values[i][0];

                    if (typeof population1990 !== "number") {
                        // Skip this iteration of the loop, and move to the next one.
                        continue;
                    }

                    // Otherwise, push the data into the in-memory store
                    cityData.push({
                        name: nameColumn.values[i][0],
                        growth: latestPopulationColumn.values[i][0] -
                            earliestCensusColumn.values[i][0]
                    });
                }

                var sorted = cityData.sort(function(city1, city2) {
                    return city2.growth - city1.growth;
                    // Note the opposite order from the usual  "first minus second" -- because want to sort in descending order rather than ascending.
                });
                var top10 = sorted.slice(0, 10);

                // The data is now all ready to be written to the output worksheet, which we'll call "Top 10 Growing Cities". However, if an existing worksheet (presumably from some previous run of this function) already exists, we do not want to error out. So, begin by checking whether the other worksheet exists, and, if so, delete it.

                var existingWorksheetIfAny = context.workbook.worksheets
                    .getItemOrNull("Top 10 Growing Cities");

                // Because everything on the JavaScript layer is a proxy object, there is no way to know if a worksheet exists or not without doing a "sync". So, do that before proceeding.

                return context.sync().then(function() {
                    // existingWorksheetIfAny is a Worksheet proxy object, so loading can't set the object itself to null. But it *can* set the isNull property:

                    if (!existingWorksheetIfAny.isNull) {
                        existingWorksheetIfAny.delete();
                    }

                    // Now that we've computed the data, create a new worksheet for the output.
                    var outputSheet = context.workbook.worksheets.add(
                        "Top 10 Growing Cities");

                    var sheetHeader = outputSheet.getRange("B2:D2");
                    sheetHeader.values = [
                        ["Top 10 Growing Cities", "", ""]
                    ];
                    sheetHeader.merge();
                    sheetHeader.format.font.bold = true;
                    sheetHeader.format.font.size = 14;

                    var tableHeader = outputSheet.getRange("B4:D4");
                    tableHeader.values = [
                        ["Rank", "City", "Population Growth"]
                    ];
                    var table = outputSheet.tables.add(
                        "B4:D4", true /*hasHeaders*/ );

                    // Could use a "for i = 0; i < array.length; i++" but using an often-more-convenient  ".forEach" approach
                    top10.forEach(function(item, index) {
                        table.rows.add(
                            null /* null means "add to end" */ , [
                                [index + 1, item.name, item.growth]
                            ]);
                        // Note: even though adding just a single row, the API still expects a 2D array for  consistency and interoperability with  Range.values.
                    });

                    // Auto-fit the column widths, and set uniform thousands-separator number formatting on the "Population" column of the table.
                    table.getRange().getEntireColumn().format
                        .autofitColumns();
                    table.getDataBodyRange().getLastColumn()
                        .numberFormat = "#,##";


                    // Finally, with the table in place, add a chart:

                    var fullTableRange = table.getRange();

                    // For the chart, no need to show the "Rank", so only use the city's name and population delta
                    var dataRangeForChart =
                        fullTableRange.getColumn(1).getBoundingRect(
                            fullTableRange.getLastColumn());

                    // A note on the function call above: Range.getBoundingRect can be thought of like a "get range between" function, creating a new range spanning between this object (in our case, the column at index 1, which is the "City" column --  remember that everything in Office.js is  zero-indexed!), and the last column of the table  ("Population Growth").

                    var chart = outputSheet.charts.add(
                        Excel.ChartType.columnClustered,
                        dataRangeForChart,
                        Excel.ChartSeriesBy.columns);

                    chart.title.text =
                        "Population Growth between 1990 and 2014";

                    var tableEndRow =
                        3 /* row #4 -- remember that we're 0-indexed */ +
                        1 /* the table header */ +
                        top10.length /* presumably 10 */ ;

                    var chartPositionStart = outputSheet.getRange("F2");
                    chart.setPosition(
                        chartPositionStart,
                        chartPositionStart.getOffsetRange(
                            19 /* 19 rows down, i.e., 20 rows in total */ ,
                            9 /* 9 columns to the right, so 10 in total */
                        )
                    );

                    outputSheet.activate();
                });
            });
    }
).catch(function(error) {
    app.showNotification("Error", error);
    // Log additional information to the console, if applicable:
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
