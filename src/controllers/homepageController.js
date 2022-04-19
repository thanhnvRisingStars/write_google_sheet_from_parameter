import moment from "moment";
const { GoogleSpreadsheet } = require('google-spreadsheet');

const PRIVATE_KEY = 'YOUR_PRIVATE_KEY'
const CLIENT_EMAIL = 'CLIENT_MAIL'
const SHEET_ID = 'GOOGLE_SPREADSHEET_ID';

let getHomepage = async (req, res) => {
    return res.render("homepage.ejs");
};

let getGoogleSheet = async (req, res) => {
    try {
        let currentDate = new Date();

        const format = "HH:mm DD/MM/YYYY"

        let formatedDate = moment(currentDate).format(format);

        // Initialize the sheet - doc ID is the long id in the sheets URL
        const doc = new GoogleSpreadsheet(SHEET_ID);

        // Initialize Auth - see more available options at https://theoephraim.github.io/node-google-spreadsheet/#/getting-started/authentication

        await doc.useServiceAccountAuth({
            client_email: CLIENT_EMAIL,
            private_key: PRIVATE_KEY,
        });

        await doc.loadInfo(); // loads document properties and worksheets

        const sheet = doc.sheetsByTitle[req.query.sheetname]; // or use doc.sheetsById[id] or doc.sheetsByTitle[title]
        console.log('-----------');
        await sheet.loadCells({ // GridRange object
            startRowIndex: 0, endRowIndex: 100, startColumnIndex:0, endColumnIndex: 200
          });
        const cell = sheet.getCell(req.query.row-1,req.query.col-1);
        cell.value = req.query.data;
        await sheet.saveUpdatedCells(); // saves both cells in one API call

        return res.send('Writing data to Google Sheet succeeds!')
    }
    catch (e) {
        console.log(e);
        return res.send('Oops! Something wrongs, check logs console for detail ... ')
    }
}

module.exports = {
    getHomepage: getHomepage,
    getGoogleSheet: getGoogleSheet
};
