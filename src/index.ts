import { parse } from 'ts-command-line-args';
import {ethers, Wallet} from 'ethers';
import os from 'os'
import Excel from 'exceljs'
import * as fs from 'fs';
import * as path from 'path';
import {format} from 'date-fns';

interface GenKeyArguments {
    filepath?: string,
    filename?: string,
    num?: number,
}

interface KeyInfo {
    phrase: string,
    address: string,
    privateKey: string,
    publicKey: string,
}

//parse command line argruments
const args = parse<GenKeyArguments>({
    filepath: {type: String, alias: 'p', description: 'save file path', optional: true},
    filename: {type: String, alias: 'f', description: 'save file name', optional: true},
    num:  {type: Number, alias: 'n', description: 'number of generate key', optional: true},
},
// {
//     helpArg: 'help',
//     headerContentSections: [{ header: 'My Example Config', content: 'Thanks for using Our Awesome Library' }],
//     footerContentSections: [{ header: 'Footer', content: `Copyright: Big Faceless Corp. inc.` }],
// }
);

let { filepath, filename, num } =  {
    filepath: args.filepath || os.homedir(),
    filename: args.filename || "ethereum_key.xlsx",
    num: args.num || 100
}

//gen key
let keyArr: KeyInfo[] = [];
console.log('%s %s %s %s', 'address', 'phrase', 'privateKey', 'publicKey');
for(let idx = 0; idx < num; idx++) {
    let wallet = ethers.HDNodeWallet.createRandom(undefined, undefined, undefined);
    console.log('%s "%s" %s %s', wallet.address, wallet.mnemonic?.phrase, wallet.privateKey, wallet.publicKey)
    keyArr.push({
        phrase : wallet.mnemonic?.phrase || '',
        address: wallet.address,
        privateKey : wallet.privateKey,
        publicKey : wallet.publicKey,
    });
}

let workbook = new Excel.Workbook();
let worksheet = workbook.addWorksheet();
worksheet.columns = [
                    {header: "PHRASE", key: 'phrase', width:90},
                    {header: "ADDRESS", key: 'address', width: 45},
                    {header: "PRIVATE KEY", key: 'privateKey', width: 70},
                    {header: "PUBLIC KEY", key: 'publicKey', width: 70},
                ];
// header row set bold and center
let headerRow = worksheet.getRow(1);
headerRow.eachCell(cell => {
    cell.style = {
        font: {
            bold: true,
        },
        alignment:{
            horizontal: 'center',
        },
    };
});

//add data
worksheet.addRows(keyArr);

//.xlsx
if(!filename.endsWith('.xlsx')) {
    filename += '.xlsx';
}

//write file
let writeFilePath = path.join(filepath, filename);
//exist file save new other file
if(fs.existsSync(writeFilePath)) {
    const filenameOption = path.parse(filename);    
    filename = filenameOption.name + "_" + format(new Date(), 'yyyyMMddHHmmss') +filenameOption.ext;
    writeFilePath = path.join(filepath, filename);
}

//write file
console.log("key file save path: ", writeFilePath);
writeKeyFile(workbook, writeFilePath);
async function writeKeyFile(workbook: Excel.Workbook, writeFilePath: string){
    await workbook.xlsx.writeFile(writeFilePath);
}
console.log("save key file sucessfully!");

// node .\src\index.js --filepath=C:\Users\cxw\Desktop\ --filename=ethereum-key1111.xlsx --num=20
// node .\src\index.js --p=C:\Users\cxw\Desktop\ --f=ethereum-key --n=20
