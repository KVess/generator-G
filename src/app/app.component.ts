import { Component } from '@angular/core';
import * as xlsx from 'xlsx';
import * as moment from 'moment';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  title = 'generator-g';
  script: string[] = [];
  afficher: string = ""
  excelFile: any;
  private arrayBuffer: any;
  private worksheet: any;

  constructor(){
    for (let index = 0; index < 10; index++) {
      this.script.push('Je suis l\'homme le plus heureux et j\'aide mes proches');  
    }

    console.log('Insertion dans une chaine de charactère');
    var test = moment(this.ExcelDateToJSDate(43105)).format('MM-DD-YYYY');
    console.log('New Date', test.toString()); 
    console.log('ParseInt', parseInt('NC')); 

    this.afficher = ""
    this.script.forEach(elmt => {
      this.afficher += elmt + '\n';
    });
  }

  loadFile(event: any) {
    console.log("Chargement effectué");
    this.excelFile = event.target.files[0];
    const classe = {
      libelle: null,
      ecole: null,
      date_creation: null,
      nbre_etudiant: null
    };
    // this.fileReader(this.excelFile, classe);
    this.fileReader(this.excelFile);
  }

  // private fileReader(file: any, line: any) {
  private fileReader(file: any) {
    console.log("File reader exécuté"); 
    const fileReader = new FileReader();

    fileReader.onload = (e) => {
      this.arrayBuffer = fileReader.result;
      const data = new Uint8Array(this.arrayBuffer);
      const arr = [];

      for (let i = 0; i !== data.length; i++) {
        arr[i] = String.fromCharCode(data[i]);
      }
      const bstr = arr.join('');

      var workbook = xlsx.read(bstr, {type:"binary"});
      var first_sheet_name = workbook.SheetNames[2];
      var worksheet = workbook.Sheets[first_sheet_name];
      console.log(xlsx.utils.sheet_to_json(worksheet,{raw:true}));
      this.makeQuerry(xlsx.utils.sheet_to_json(worksheet,{raw:true}));

      // const bstr = arr.join('');
      // // const workbook = xlsx.read(bstr, { type: 'binary', cellDates: true });
      // const workbook = xlsx.read(bstr, { type: 'binary'});
      // const first_sheet_name = workbook.SheetNames[0];

      // const worksheet = workbook.Sheets[first_sheet_name];
      // this.worksheet = xlsx.utils.sheet_to_json(worksheet, {raw:true});
      // console.log(worksheet);
      /**
       * Call matching function
       */
      // this.matchingCell(this.worksheet, line);
    };
    fileReader.readAsArrayBuffer(file);
  }

  private makeQuerry(jsonCollection: any){
    this.script = [];
    // if(jsonCollection.length > 100){
    if(false){
      (jsonCollection as []).forEach(tab => {
        (tab as []).forEach(jObject => {
          this.script.push(this.jsonToQuerry(jObject));
        });
      });
    } else {
      (jsonCollection as []).forEach(jObject => {
        if (jObject[' Price '] === 'NC') {
          console.log('vérification', parseInt(jObject[' Price ']));
        }

        // if (parseInt(jObject[' Price ']) !== NaN) {
        if (!isNaN(jObject[' Price '])) {
          this.script.push(this.jsonToQuerry(jObject));
        }
      });
    }

    // console.log('tableau de script', this.script);

    this.afficher = ""
    this.script.forEach(elmt => {
      this.afficher += elmt + '\n';
    });
  }

  private jsonToQuerry(jsonObject: any) {
    var stockID = '31';
    var querry: string = 'INSERT INTO tblStockDetails(FKStockID,TradingDate,IsUnlisted,Price,Volume,StockValue,ResOff,ResDem,AdjustedPrice,ResDemPrice,ResOffPrice,IsActive,IsDelete) VALUES (' + stockID + ', ';

    querry += '\'' + moment(this.ExcelDateToJSDate(+jsonObject['Date'])).format('MM-DD-YYYY') + '\', 0, ' + jsonObject[' Price '] + ', ' + jsonObject['Volume'] + ', ' + jsonObject['Value'] + ', ' + jsonObject['ResOff'] + ', ' + jsonObject['ResDem'] + ', 0, 0, 0, 1, 0)';
    
    return querry;

  }

  private ExcelDateToJSDate(serial: number) {
    var utc_days  = Math.floor(serial - 25569);
    var utc_value = utc_days * 86400;                                        
    var date_info = new Date(utc_value * 1000);
 
    var fractional_day = serial - Math.floor(serial) + 0.0000001;
 
    var total_seconds = Math.floor(86400 * fractional_day);
 
    var seconds = total_seconds % 60;
 
    total_seconds -= seconds;
 
    var hours = Math.floor(total_seconds / (60 * 60));
    var minutes = Math.floor(total_seconds / 60) % 60;
 
    return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
 }
}
