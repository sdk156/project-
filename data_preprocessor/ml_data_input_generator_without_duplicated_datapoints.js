// this program takes the csv data files of collected data and converts to format as specified by vignesh
//column_1 to column_6 : sensor data, column 7 : (VERY_FRESH,EARLY_SPOILED,HALF_SPOILED,FULL_SPOILED)
//for algorithm's friendlyness, each datapoint (1 datapoint = S1,S2,S3,S4,S5,S6 sensor data) of all trials are considered as one datapoint and fed into the algorithm
//as per sarathy anna's sensor reading methodology, each trial is done in the following way.
// from_seconds    |    to_seconds    |    description
//      0          |       500        |    air resistance or baseline votage
//      501        |      1000        |    shrimp ambience( shrimp is kept inside the sensor kit box for analysis)
//      1001       |       1500       |    shrimp removed from kit box. i.e sensor recovery characteristics
//      1501       |       2000       |    shrimp ambience( shrimp is kept inside the sensor kit box for analysis)
//      2001       |       2500       |    shrimp removed from kit box. i.e sensor recovery characteristics
//      2501       |       3000       |    shrimp ambience( shrimp is kept inside the sensor kit box for analysis)
//      3001       |       3500       |    shrimp removed from kit box. i.e sensor recovery characteristics
//
//for training the model, only the shrimp ambience part of data is required
//hence this program is used to extract only the shrimp ambience part of data and places it in a single csv file with all trials combined that is, each trial is also considered
// as a separate datapoint of 6 sensors
// outline of final csv file columns is as follows...
// Sensor_1 | Sensor_2 | Sensor_3 | Sensor_4 | Sensor_5 | Sensor_6 | Classification | Day 
// Classification tells one of the four decisions of shrimp quality VERY_FRESH,EARLY_SPOILED,HALF_SPOILED,FULL_SPOILED
// Day column is added in the final outline becoz, the model is expected to do future predictions of spoilage duration also.
// in this program duplicate datapoints are checked among four decisions. that is, whether same value of all six sensor value exsists among four decisions or not

const Exceljs = require('exceljs'); //import module
// as per the  module, workbook is a file and worksheet is a sheet inside a workbook file
//hence create an workbook instance, then create a worksheet inside that workbook and start processing it
//  below variables should be defined by user to secify the csv files from which data is to be extracted and place those files in the same folder as this js file
const filename = [
        "fresh 1.csv", "fresh 2.csv", "fresh 3.csv",
        "Day 2 trail 1 .csv", "Day 2 trail 2.csv", "Day 2 trail 3.csv", "Day 4 trail 1.csv", "Day 4 trail 2.csv",
        "Half spoiled 1.csv", "Half spoiled 2.csv", "Half-spoiled 3.csv", "HS 1_trial 1.csv", "HS 2_trail 2.csv", "HS_trial 3.csv",
        "full spoil trail 1.csv", "full spoil trail 2.csv", "fully spoil trail 3.csv"
    ] //insetr the filenames from which the sensor data has to be read 
const shrimpQualityClassificationInFileName = [ ///data classification based on the fileName[]
    'VERY_FRESH', 'VERY_FRESH', 'VERY_FRESH',
    'EARLY_SPOILED', 'EARLY_SPOILED', 'EARLY_SPOILED', 'EARLY_SPOILED', 'EARLY_SPOILED',
    'HALF_SPOILED', 'HALF_SPOILED', 'HALF_SPOILED', 'HALF_SPOILED', 'HALF_SPOILED', 'HALF_SPOILED',
    'FULL_SPOILED', 'FULL_SPOILED', 'FULL_SPOILED'
]; //four classifications of shrimp quality
const shrimpQuality = "FULL_SPOILED"; //one of the four classifications of data for naming purposes of csv file
const dayCountFileName = [
        0, 0, 0,
        2, 2, 2, 4, 4, 4,
        6, 6, 6, 8, 8, 8,
        10, 10, 10
    ] // nth day data recording based on the filename[]
const last_copy_row = 3501; // since first row is a header row, total 
const desired_row_count = 500 // shrimp ambeince datapoints row count in each one set of data
const finalworkbook_name = "finaldata_unique_datapoints.csv"; // filename of final csv file with all trials of one particular day


//initialization of workbook instance
const finalWorkbook = new Exceljs.Workbook();
//adding a worksheet in the workbook
var finalWorksheet = finalWorkbook.addWorksheet("sheet1");
var nth_set_rows_final_sheet = 0; //updated inside the loop to count the number of sets of shrimp ambience data is written into finalsheet for data insertion purposes


// Promise.resolve(() => {
//from experiments done with this exceljs library, it is seen that, when saving the csv object to a file by using csv.writefile function, each time when saving the file,
// the previous file was overwritten. that is actually good. but the overwriting is happened with in the inner 'for' loop to ensure file is written after getting 
//necessary columns from naother csv file. even this is correct procedure to be followed, while writing the file, previous call to write file conflicts with current call
//and data is stored incorrectly in the file. but when saving a single column in the file using writeFile function only once, data is correctly stored. hence, if all the file
// columns are saved to object and writefile function is executed only once, no file writing conflicts will occur. to implement such mechanism in node (node is async by nature)
// I have used promise  to check all code to all data writing tasks to worksheet object have been completed and if all code done, wring the object to file
// this is done by counting number of times the data is written into object. if the count equals 18 that is 6 sensor values with 3 trials for each sensor, resolve a promise which 
// then writes the finalworksheet object to a csv file
// above is tried but the fact is that it is not a good implementtion for handling async events. because, we have to hardcode the number 18 somewhere and is not a scallable .
//as a good practice, i used async/await keywords to have a sequential approach as coded below.
// refer asyc/await  https://javascript.info/async-await
var finalSheetAllRows = new Array(); /// for storing the finalsheet rows before storing it into the final worksheet all rows at once,
var duplicateRowFounderSet = new Set(); // it is a set datatype because, the input data to the ML mode should not contain any duplicate data
var fromRow_finalSheet = 2; // this variable calculates the starting row of next rows set of finalsheet, in which data has to be inputted
finalWorksheet.addRow(['S1', 'S2', 'S3', 'S4', 'S5', 'S6', 'Decision', 'Age'])
let asyncFunction = async() => {
    console.log("Total csv data file count: ", filename.length);
    console.log("max populated row would be: ", 3 * desired_row_count * filename.length + 1); //+1 is because, first row is a header
    var set_iterator = duplicateRowFounderSet.values();
    var duplicateRowCount = 0; //to indicate the number of duplicate rows in the given files
    for (let i = 0; i < filename.length; i++) {
        let fromRow_sourceSheet = 2;
        // read from a file
        let workbook = new Exceljs.Workbook(); // for storing the workbook objects of csv files defined in filename[]
        let worksheet = await workbook.csv.readFile(filename[i]); // for storing the worksheet objects of csv files defined in filename[]

        for (let j = 1; j <= 3; j++) { //each csv file has three times the shrimp ambience hence  1<=j<=3// refer the comments written at the beginning
            fromRow_sourceSheet = fromRow_sourceSheet + desired_row_count; //shrimp ambience data rows
            console.log(filename[i] + " First Row :" + fromRow_sourceSheet + "  " + finalworkbook_name + " First Row : " + fromRow_finalSheet);

            // number 2 in the above formula is due to the fact that the first row is header row. and the rows of worksheet object is a 1-indexed array not 0-indexed
            let finalsheetRows_temp = worksheet.getRows(fromRow_sourceSheet, desired_row_count); //returns an iterable object, hence forEach is to be done and each row is parsed using .values property as shown below
            // console.log("Extracted rows count", finalsheetRows_temp.length);

            finalsheetRows_temp.forEach(function(value, rowIndex, rowArray) { // (current array element, index in finalsheetRows_temp, finalsheetRows_temp)
                // rowArray and finalsheetRows_temp referes to the same object as per docs of forEach() function of Array


                rowArray[rowIndex] = value.values.slice(2, 8); //slice method extracts the array from 2 to one element before index 8
                //we use slice mwthod to remove the s.no field. so slice method should be 1,7 but, exceljs library aslyaws returns the row with first position (index 0) empty
                //                   so, that an 1 indexed array can be constructed, hence to eliminate s.no field, we have to start extracting the array from index 2.

                //inthe below code, before adding the row into the finalsheetAllRows array, duplicate checking of six sensor values is done and then addded into the it

                if (!(duplicateRowFounderSet.has(rowArray[rowIndex].toString()))) { //it is converted into string for comparison purposes, if not converted to string, even if both rows are same, they cant comapre it and take it as two different tobjects
                    // hence converted to string and then checking is done

                    duplicateRowFounderSet.add(rowArray[rowIndex].toString()); // for comparing purposes, it is converted into string and then stored into the array
                    // console.log("[ " + set_iterator.next().value.toString() + " ]", duplicateRowFounderSet.size);
                    rowArray[rowIndex].push(shrimpQualityClassificationInFileName[i]); // pushing the shrimp quality into the sliced Array
                    rowArray[rowIndex].push(dayCountFileName[i]); //pushing the shrimp age into the sliced array
                    finalSheetAllRows.push(rowArray[rowIndex]); //insert the modified row into the rows array for adding it ito the workseet in bulk

                    fromRow_finalSheet++;
                } else {
                    // console.log("dupilcate found at: " + filename[i] + " row: " + (fromRow_sourceSheet + 1) + " [ " + rowArray[rowIndex].toString() + " ] duplicate rows count: " + duplicateRowCount++);
                }
                fromRow_sourceSheet++;


            });
            // above line of code extracts a single datapoint of six columns from source csv and appends the row with shrimp quality and age of shrimp in days. 
            // slice method (removes first element aof array and returns it) is used to extract only the sensor data as a single row of source csv has s.no as first element of each row.hence,from index 1 to index 7(excluded) 
            //then, push method is used to push the corressponding shrimp quality and age of shrimp in days as defined in filename 


            console.log(filename[i] + " Last Row :" + (fromRow_sourceSheet - 1) + "  " + finalworkbook_name + " Last Row : " + (fromRow_finalSheet - 1));
            //(fromRow_sourceSheet-1) and (fromRow_finalSheet-1) is used instead of fromRow_sourceSheet and fromRow_finalSheet only is because, starting both variables are incremented at tthe end of its iteration inside forEach
            //hence the incremented value of at the end of each iteration is used for the next iteration starting, hence the incre,ented value-1 is showed in console , which is ccorrect
            //but in the formula above in row extraction, this is not used, because, starting number row and total rows from that number is specified.
            //if end row number had asked, the end number of row would have been 501 if starting number row is 2

        }
        console.log();

    }
    finalWorksheet.addRows(finalSheetAllRows);
    await finalWorkbook.csv.writeFile(finalworkbook_name).then(() => {
        console.log(finalworkbook_name + ' saved successfully');
        console.log("Set() length", duplicateRowFounderSet.size)
    }).catch((err) => {
        console.log(finalworkbook_name + 'file save failed');
    });
}


asyncFunction();