const Exceljs = require('exceljs'); //import module
// as per the  module, workbook is a file and worksheet is a sheet inside a workbook file
//hence create an workbook instance, then create a worksheet inside that workbook and start processing it
//  below variables should be defined by user to secify the csv files from which data is to be extracted and place those files in the same folder as this js file
const filename = ["full spoil trail 1.csv", "full spoil trail 2.csv", "fully spoil trail 3.csv"] //insetr the filenames from which the sensor data has to be read 
const shrimpQuality = "FULL_SPOILED"; //one of the four classifications of data for naming purposes of csv file
const dayCount = 10; // nth of the data recording

const fianlworkbook_name = "day_" + dayCount + "_" + shrimpQuality + "" + ".csv"; // filename of final csv file with all trials of one particular day
// const finalworksheet = "sheet1"
var workbook = new Array();
var worksheet = new Array();
// write to a file
const finalWorkbook = new Exceljs.Workbook();
var finalWorksheet = finalWorkbook.addWorksheet("sheet1");

var columnInsertCompletePromise = new Promise((resolve, reject) => {

});

var finalWorksheetColumn = 2;
const maxColumns = 18;
const lastCopyColumn = 7;
var total_columns = 1;
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
let asyncFunction = async() => {
        for (let i = 0; i < filename.length; i++) {
            // read from a file
            workbook[i] = new Exceljs.Workbook(); //create an workbook object to read a file
            worksheet[i] = await workbook[i].csv.readFile(filename[i]); //.then((readSheets) => {
            var finalWorksheetColumn = 2 + i; //insert columns of all sensors of one trial at a time by leaving spaces for trials of same sensor of other trials
            for (let j = 2; j <= lastCopyColumn; j++) {
                finalWorksheet.getColumn(finalWorksheetColumn).values = await worksheet[i].getColumn(j).values;
                console.log("Trial data CSV file: " + (i + 1) + " Column: " + j + "--->" + fianlworkbook_name + " Column : " + finalWorksheetColumn);
                finalWorksheetColumn += 3;
                // if (total_columns == maxColumns) {
                //     finalWorksheet.csv.writeFile(fianlworkbook_name).then(() => {
                //         console.log('file save successfull');

                //     }).catch((err) => {
                //         console.log('file save unsuccessfull', err);
                //     });
                // }
            }
            console.log();
            // });
            // .then(() => {
            //     finalWorksheet.csv.writeFile(fianlworkbook_name).then(() => {
            //         console.log('file save successfull');

            //     }).catch((err) => {
            //         console.log('file save unsuccessfull', err);
            //     }); //reading a csv file
            // });
            // return Promise.resolve(finalWorkbook);
        }
        const row_count = finalWorksheet.getColumn(2).values.length; //counting how many rows , to fill the first column with s.no.. from 2nd column, columns are already populated...hence we acan use any column number for row count.. he re 2 is used
        console.log("row count: ", row_count);
        const column1_value = new Array(row_count); // place holder variable of column1. this array is used to store the serial numbers and then directly assigned using the the above method already used
        column1_value[0] = "S.no"; //header of first column
        for (let i = 1; i < row_count - 1; i++) {
            column1_value[i] = i;
        }
        finalWorksheet.getColumn(1).values = column1_value; //filling first column
        await finalWorkbook.csv.writeFile(fianlworkbook_name).then(() => {
            console.log(fianlworkbook_name + ' saved successfully');
        }).catch((err) => {
            console.log(fianlworkbook_name + 'file save failed');
        });
    }
    // asyncFunction().then((finalbook) => {
    //     finalbook.csv.writeFile(fianlworkbook_name).then(() => {
    //         console.log('file save successfull');

//     }).catch((err) => {
//         console.log('file save unsuccessfull', err);
//     });
// });
// columnInsertCompletePromise.then((finalbook) => {
//     finalbook.csv.writeFile(fianlworkbook_name).then(() => {
//         console.log('file save successfull');

//     }).catch((err) => {
//         console.log('file save unsuccessfull', err);
//     });
// });


asyncFunction();