const express = require("express"); 
const {auth} = require("google-auth-library");

const run = require('./helper/google-sheet');
const bodyParser = require('body-parser');
const ejs = require('ejs');
const fs = require('fs');
const { google } = require('googleapis');

const { GoogleSpreadsheet } = require("google-spreadsheet");


const credentials = JSON.parse(fs.readFileSync('google-client-secret.json', 'utf-8')); 
const { client_secret: clientSecret, client_id: clientId, redirect_uris: redirectUris, } = credentials.installed; 
const oAuth2Client = new google.auth.OAuth2( clientId, clientSecret, redirectUris[0], ); 
const SCOPES = [ 'https://www.googleapis.com/auth/spreadsheets', ];
const url = oAuth2Client.generateAuthUrl({ access_type: 'offline', scope: SCOPES, });
console.info(`authUrl: ${url}`);


const app = express(); 

var range2 = "";
var range = "";
var spreadsheetId = "";
var sheetID = "";


app.set("view engine", "ejs"); 
app.use(express.urlencoded({extended: true}));
app.use(express.static("public"));

app.get("/",(req, res) => { 
    res.render("info");
});

app.post("/selected", async(req, res) =>{
    range2 = req.body.selected;
    const rangeValue = range2.split(',');

    range = rangeValue[0];
    sheetID = rangeValue[1];

    console.log('request value : ' + range);

    res.render("insert");
})

app.post("/check", async (req, res) => {
    //const spreadsheetId = await "15ci4qvWgF1SE_9EmHRNLrFKdZC06xZRzvo6xNi2T2lA";
 
    spreadsheetId =  req.body.selected;  
    console.log('type of : '+ typeof spreadsheetId)
    console.log('spreadsheetId :' + spreadsheetId);
    const auth = new google.auth.GoogleAuth({
      keyFile: "credentials.json",
      scopes: "https://www.googleapis.com/auth/drive",
    });
   
    const getTitle = async (auth, spreadsheetId) => {
      const client = await auth.getClient();
      const sheets = await google.sheets({ version: "v4", auth: client });
      const result = (await sheets.spreadsheets.get({spreadsheetId})).data.sheets.map((sheets)=>{
        return sheets.properties.title;
      })
      return result;
    };

  const getSheetID = async (auth, spreadsheetId) => {
      const client = await auth.getClient();
      const sheets = await google.sheets({ version: "v4", auth: client });
      const sheetID = (await sheets.spreadsheets.get({spreadsheetId})).data.sheets.map((sheets)=>{
        return sheets.properties.sheetId
      })
      return sheetID;
    };
    
    const result = await getTitle(auth, spreadsheetId); 
    sheetID = await getSheetID(auth,spreadsheetId);
    let result_entries = Object.values(result);
    console.log('dataType : ' + typeof result);
    console.log('result : ' + result_entries);


    result.push(spreadsheetId);
    res.render("index", {result,sheetID});
  });
   


async function getAuthSheet(){
  const auth = new google.auth.GoogleAuth({
    "keyFile" : "credentials.json", //서비스 계정키
    "scopes" : "https://www.googleapis.com/auth/spreadsheets", //읽고 쓰기 권한 허용
  });

  const client = await auth.getClient();

  const googleSheets = google.sheets({version: "v4", auth: client});

    const spreadsheetId = "15ci4qvWgF1SE_9EmHRNLrFKdZC06xZRzvo6xNi2T2lA"; //스프레드 시트 ID

  return{
    auth,
    client,
    googleSheets,
    spreadsheetId
  }
}

app.get('/sheetList', async(req, res)=>{
  
  const {googleSheets, auth, spreadsheetId} = await getAuthSheet();

  const sheetList = await googleSheets.spreadsheets.get({
    auth,
    spreadsheetId
  })

  res.send(sheetList.data);
})

app.post('/delete', async(req, res) =>{

   const name = req.body.name;
   console.log("삭제되는 이름:" + name);
   const doc = new GoogleSpreadsheet(spreadsheetId);

   const deletRow = async (req, res) =>{
   const credentials = JSON.parse(fs.readFileSync('credentials.json'))
    
    await doc.useServiceAccountAuth({
      client_email : credentials.client_email,
      private_key : credentials.private_key
    })

    await doc.loadInfo();
 
   let sheet = doc.sheetsByTitle[range];
 
   let rows = await sheet.getRows();
   console.log('sheet : ' + sheet);
   console.log(rows);
 
   for(let index = 0; index < rows.length; index++){
     const row = rows[index];
     if(row[req] == res){
       await rows[index].delete();
       break;
     }
   }
 }

 deletRow('Name', name);
})

app.post('/update', async(req, res) =>{
  const doc = new GoogleSpreadsheet(spreadsheetId);

  const update = async(keyvalue, oldValue, newValue) =>{
     const credentials = JSON.parse(fs.readFileSync('credentials.json'))
    
    await doc.useServiceAccountAuth({
      client_email : credentials.client_email,
      private_key : credentials.private_key
    })

    await doc.loadInfo();
 
   let sheet = doc.sheetsByTitle[range];
 
   let rows = await sheet.getRows();
   console.log('sheet : ' + sheet);
   console.log(rows);

   for(let index = 0; index <rows.length; index++){
      const row = rows[index];
      if(row[keyvalue] == oldValue){
        rows[index][keyvalue] = newValue;
        await rows[index].save();
        break;
      }
   }
  }
})

app.post("/append", async (req, res) => {
    const {Name, Age, Gendar, Address, Memo} = req.body; // 요청되는 body안의 값 객체로 전달
    console.log({Name, Age, Gendar, Address, Memo});
    //허용 범위 설정
    
    const auth = new google.auth.GoogleAuth({
        "keyFile" : "credentials.json", //서비스 계정키
        "scopes" : "https://www.googleapis.com/auth/spreadsheets", //읽고 쓰기 권한 허용
    });

    const client = await auth.getClient();

    const googleSheets = google.sheets({version: "v4", auth: client});
    var spId = spreadsheetId;
    //const spreadsheetId = spreadsheetId; //스프레드 시트 ID
    console.log('spreadsheetId : '  + spreadsheetId);
    console.log('range : ' + range);

    await googleSheets.spreadsheets.values.append({
        auth,
        spreadsheetId,
        range : range+"!A2:F", // 행열 범위
        valueInputOption : "USER_ENTERED", //사용자
        resource : {    
           values : [
                [Name, Age, Gendar, Address, Memo]
            ],
        }
    });

    //res.send(getRows.data);
});

// 내부 시트 생성
app.post("/addSheet", async (req, res) => {

   // const spreadsheetId = "15ci4qvWgF1SE_9EmHRNLrFKdZC06xZRzvo6xNi2T2lA"; //스프레드 시트 ID

    const {title} = req.body;

    const auth = new google.auth.GoogleAuth({
        "keyFile" : "credentials.json", //서비스 계정키
        "scopes" : "https://www.googleapis.com/auth/drive", //모든 한 부여
    });

  const client = await auth.getClient();
  const sheets = google.sheets({version: "v4", auth: client}); 

   sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    "resource" : {
        "requests": [{
              "addSheet": {
                "properties": {
                  "title": title,
              }
            }
          }]
        }
      });
    });

app.get("/deleteSheet", async (req,res) =>{
  
  const auth = new google.auth.GoogleAuth({
      "keyFile" : "credentials.json", //서비스 계정키
      "scopes" : "https://www.googleapis.com/auth/drive", //모든 한 부여
  });
console.log('삭제할 시트 : ' + sheetID);
const client = await auth.getClient();
const sheets = google.sheets({version: "v4", auth: client}); 

  sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    "resource" :{
      "requests": [{
        "deleteSheet" :{
            "sheetId" : sheetID,
        }
      }]
    }
  })
})
  //새로운 시트 생성
  app.post("/new", async (req, res) =>{
    const spreadsheetId = "15ci4qvWgF1SE_9EmHRNLrFKdZC06xZRzvo6xNi2T2lA"; //스프레드 시트 ID
    const range = "sheet1!A:F5";
    const title = req.body.title;

    console.log(title);

   await run.create(spreadsheetId,range, title);
  })

app.listen(1337, (req, res) => console.log("running 1337")); // 서버 포트 설정
