

// Import the file system module
const fs = require('fs');
const path = require('path');
// excell paxkage
const ExcelJS = require('exceljs');


// Get the current date and time
const now = new Date();
const year = now.getFullYear();
const month = String(now.getMonth() + 1).padStart(2, '0'); // Ensure 2 digits for month
const day = String(now.getDate()).padStart(2, '0');
const hours = String(now.getHours()).padStart(2, '0');
const minutes = String(now.getMinutes()).padStart(2, '0');
const seconds = String(now.getSeconds()).padStart(2, '0');

// Create a file-safe timestamp (removing colons and adding dashes)
const currentTime = `${year}-${month}-${day}_${hours}-${minutes}-${seconds}`;











class jsonGen{



        ensureDirectoryExistence(filePath) {

            if (!filePath) {
                throw new Error('File path is undefined');
            }

            const dir = path.dirname(filePath);
            if (!fs.existsSync(dir)) {
            fs.mkdirSync(dir, { recursive: true });
            }
        }
         //  Array to json 
        ArraytoJson(data,filename,foldername){
            
            
        
            // Convert array to JSON string

                const jsonArray = JSON.stringify(data, null, 2);
                const outputFolder = path.join(__dirname, foldername?foldername:"outPut");
            
                const file = filename;  // Concatenate filename and timestamp

                // Ensure the output folder exists
                if (!fs.existsSync(outputFolder)) {
                    fs.mkdirSync(outputFolder, { recursive: true });  // Create folder if it doesn't exist
                }

                // Define the output path with the file-safe name
                const outputPath = path.join(outputFolder, `${file}.json`);

                // Write the JSON string to a file
                fs.writeFile(outputPath, jsonArray, (err) => {
                    if (err) {
                        console.error('Error writing file:', err);
                        return;
                    }
                    console.log('JSON file has been saved.');
                });
                        


        }
        //    read excell data
        readExcelFile(filepath,sheetnumber) {
            const workbook = new ExcelJS.Workbook();

            return workbook.xlsx.readFile(filepath)
            .then(() => {
                console.log('Available Sheets:', workbook.worksheets.map(sheet => sheet.name));

                // Get the specified worksheet
                const worksheet = workbook.getWorksheet(sheetnumber?sheetnumber:1);

                if (!worksheet) {
                throw new Error('Worksheet not found');
                }

                // Initialize an array to hold the final data
                const data = [];
                const headers = []; // To store the header keys

                worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
                // Remove the first empty item from each row
                const rowData = row.values.slice(1);

                // On the first row, store headers
                if (rowNumber === 1) {
                    headers.push(...rowData); // Spread operator to add all headers
                } else {
                    // Create an object for subsequent rows using the headers as keys
                    const rowObject = {};
                    rowData.forEach((value, index) => {
                    rowObject[headers[index]] = value; // Map headers to their respective values
                    });
                    data.push(rowObject); // Push the object to data array
                }
                });

                return data; // Return the data at the end of the promise chain
            })
            .catch(error => {
                console.error('Error reading Excel file:', error.message);
                throw error; // Re-throw the error for further handling if necessary
            });
        }
        // excell to json
        ExceltoJson(excelpath,sheetnumber){
          let excelData = this.readExcelFile(excelpath,sheetnumber?sheetnumber:1)
           excelData.then(res=>{
            console.log(res);
            this.ArraytoJson(res,"ExcellTojson"+currentTime,"ExcellToJson")
           }).catch(err=>{
                 console.log("error",err);
            
           })
            
        }
        // json to excell
        JsontoExcell(data,outputFilename){

            if (!outputFilename) {
                outputFilename = `JsontoExcell/JsontoExcell-${currentTime}.xlsx`;
             }else{ 
                outputFilename="JsontoExcell/" + outputFilename + currentTime + ".xlsx";
             }

            // ./ is must before path
            const json=require(data)
            console.log(json);
            
            // Create a new workbook
            const workbook = new ExcelJS.Workbook();
            // Add a new worksheet
            const worksheet = workbook.addWorksheet('Sheet1');
            // Add column headers based on the keys in the first object of the JSON data
            const columns = Object.keys(json[0]).map(key => ({ header: key, key }));
            worksheet.columns = columns;
            // Add rows based on the JSON data
            json.forEach(item => {
                worksheet.addRow(item);
            });


            this.ensureDirectoryExistence(outputFilename)

            workbook.xlsx.writeFile(outputFilename);
            console.log(`Excel file has been created at ${outputFilename}`);


        }


}


module.exports = jsonGen;