const {google} = require("googleapis")
const keys = require("./key.json")
const Excel = require('exceljs');
const {Translate} = require('@google-cloud/translate').v2;
const key = "AIzaSyDQ41_OF_nYymEUj331n0wyFUQUUwGcR5A"
const client = new google.auth.JWT(
    keys.client_email,
    null,
    keys.private_key,
    ["https://www.googleapis.com/auth/spreadsheets"]
)

client.authorize((err,token)=>{
    if(err){
        console.log(err)
    } else if (token) {
       gsrun(client)
    }
})



async function gsrun(cl) {
    const projId= "excel-into-sheet"
    const wb =  new Excel.Workbook()
    const translate = new Translate({key:key});
    const gsapi = google.sheets({version:"v4", auth: cl})

    let excelFile = await wb.xlsx.readFile("center.xlsx");
    let ws = excelFile.getWorksheet("عمان");
    let data =  ws.getSheetValues()
    data = data.map(function(r){
        return [r[6]]
    })

    data.length = data.length - 38;
    //console.log(data)

    let [translations] = await translate.translate(data, "en");
    translations = Array.isArray(translations) ? translations : [translations];
    console.log('Translations:');
    let newArr = []
    translations.forEach((translation, i) => {
      console.log(`${data[i]} => (en) ${translation}`);
      
      newArr.push([translation])
    });
    console.log(newArr)

    //console.log("after trans",data)


    const updateData = {
        spreadsheetId: '16O4d2Mgh5CVidKJM7jvsOUUr9gDkkwMymmdTq8qNWZo',
        range: "test!A1",
        valueInputOption: "USER_ENTERED",
        resource: {values:newArr}
    }
    let resault = await gsapi.spreadsheets.values.update(updateData);
}