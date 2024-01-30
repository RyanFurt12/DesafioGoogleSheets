const { google } = require('googleapis');

const auth = new google.auth.GoogleAuth({
    keyFile: 'credenciais.json',
    scopes: 'https://www.googleapis.com/auth/spreadsheets',
});

const SPREADSHEET_ID = '1H46PE3AoifpFUNom2HeM7n5cXVyOOX8ES300AeBbv8s';
const RANGE = 'engenharia_de_software!A2:H';

const calculateSituation = (average, totalFails, totalClasses) => 
{
    if (totalFails > 0.25 * totalClasses) {
        return 'Reprovado por Falta';
    } else if (average < 50) {
        return 'Reprovado por Nota';
    } else if (50 <= average && average < 70) {
        return 'Exame Final';
    } else {
        return 'Aprovado';
    }
};

const updateSheet = async () => {
    const client = await auth.getClient();
    const sheetsInstance = google.sheets({ version: 'v4', auth: client });

    try {
        const response = await sheetsInstance.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: RANGE,
        });

        const values = response.data.values;

        const { requests } = require('./ConditionalFormatRule.js');

        const request = {
            spreadsheetId: SPREADSHEET_ID,
            resource: {
              requests: requests,
            },
            auth: client,
        };
        
        try {
            const response = await sheetsInstance.spreadsheets.batchUpdate(request);
            console.log('Style update successful\n');
        } catch (error) {
            console.error('Error in batch update:', error.message);
        }
        
        let totalClasses = parseInt(values[0][0].match(/\d+/))
        console.log(`--Software Engineering--\n Total Classes: ${totalClasses}\n`);

        for (const row of values) {
            let [studentId, studentName, totalFails, p1, p2, p3, situation, finalGrade] = row
                    .map((value, index) => (index === 1 || index === 6) ? value : Number(value));

            if (isNaN(studentId) || studentId === "") continue;
            
            const average = (p1 + p2 + p3) / 3;
            situation = calculateSituation(average, totalFails, totalClasses);

            let naf = 0;
            finalGrade = average.toFixed(2);

            if (situation === 'Exame Final') {
                naf = Math.ceil((70 - average) * 2 + average);
                // finalGrade = Math.round((average + naf) / 2);
            }

            const updateData = [[situation, naf]];
            const updateRange = `engenharia_de_software!G${studentId + 3}:H${studentId + 3}`;

            await sheetsInstance.spreadsheets.values.update({
                spreadsheetId: SPREADSHEET_ID,
                range: updateRange,
                valueInputOption: 'RAW',
                resource: { values: updateData },
            });

            console.log(`Name: ${studentName}, Situation: ${situation}, NAF: ${naf.toFixed(2)}, Final Grade: ${finalGrade}`);
        }
    } catch (error) {
        console.error('Error reading Google Sheet:', error.message);
    }
};

const main = async () => {
  await updateSheet();
};

console.clear();
main();
