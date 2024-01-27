const XLSX = require('xlsx');

// Function to analyze the Excel file and print results
function analyzeFile(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);

    // Helper function to convert time in HH:MM format to minutes
    function convertTimeToMinutes(time) {
        const [hours, minutes] = time.split(':').map(Number);
        return hours * 60 + minutes;
    }

    // Helper function to get a unique date string in a consistent format
    function getDateString(date) {
        return new Date(date).toLocaleDateString();
    }

    // Helper function to check if there are consecutive days considering repeated dates
    function hasConsecutiveDays(entries, consecutiveDays) {
        let currentDateCount = 1;

        for (let i = 1; i < entries.length; i++) {
            const currentDate = new Date(entries[i]['Time']).toLocaleDateString();
            const previousDate = new Date(entries[i - 1]['Time']).toLocaleDateString();

            if (currentDate !== previousDate) {
                currentDateCount = 1;
            } else {
                currentDateCount++;
            }

            if (currentDateCount >= consecutiveDays) {
                return true;
            }
        }

        return false;
    }

    const consecutiveDays = 7; // Define consecutiveDays here

    const uniqueEmployeeNames = [...new Set(data.map(e => e['Employee Name']))];

    uniqueEmployeeNames.forEach(employeeName => {
        const entriesForEmployee = data.filter(e => e['Employee Name'] === employeeName && getDateString(e['Time']) !== 'Invalid Date');

        if (entriesForEmployee.length >= consecutiveDays && hasConsecutiveDays(entriesForEmployee, consecutiveDays)) {
            console.log(`Employee ${employeeName} has worked for ${consecutiveDays} consecutive days.`);
        }

        const shiftHours = convertTimeToMinutes(entriesForEmployee['Timecard Hours (as Time)']);
        if (shiftHours > 840) {
            console.log(`Employee ${entry['Employee Name']} has worked for more than 14 hours in a single shift.`);
        }
    });
}
        
// Provide the path to your Excel file
const filePath = './file.xlsx';

// Call the function to analyze the file
analyzeFile(filePath);