import e from 'express';
import express from 'express';

import { downloadImages } from '../commons/downloadImages';

const xl = require('excel4node');

const fs = require('fs');
const path = require('path');
const request = require('request');

const router = express.Router();
import { logger } from '../commons/log';



var download = function(uri, uploadFolder, callback){
  request.head(uri, function(err, res, body){
    request(uri).pipe(fs.createWriteStream(uploadFolder)).on('close', callback);
  });
};

var setConfigWorkSheet = function(ws) {
  //style
  ws.column(1).setWidth(10);
  ws.column(2).setWidth(15);
  ws.column(3).setWidth(50);
  ws.column(4).setWidth(20);

  ws.row(1).setHeight(25);
  ws.row(1).freeze();

  const headStyle = {
    alignment:{
      vertical:'center'
    },
    fill:{
      type:'pattern',
      patternType:'solid',
      fgColor:'#d3d3d3'
    },
    font:{
      bold:true
    }
  };

  // header
  ws.cell(1,1).string('ID').style(headStyle);
  ws.cell(1,2).string('사진').style(headStyle);
  ws.cell(1,3).string('상품명').style(headStyle);
  ws.cell(1,4).string('S/N CODE').style(headStyle);
};


var setDataInWorkSheet = function(ws, item, index) {
  ws.row(index+2).setHeight(120);
  ws.cell(index+2,1).string(item.itemId.toString());

  const lastIdx = item.imgLinkTh.lastIndexOf('/');
  const uploadFolder = path.resolve('./downloadImages'+item.imgLinkTh.substring(lastIdx));

  let pic = ws.addImage({
    path: uploadFolder,
    type: 'picture',
    position: {
      type: 'twoCellAnchor',
      from: {
        row: index+2,
        colOff: "1mm",
        col: 2,
        rowOff: "1mm"
      },
      to: {
        row: index+2,
        colOff: "25mm",
        col: 2,
        rowOff: "45mm",
      },
    },
  });
  pic.editAs = "twoCell";


  ws.cell(index+2,3).string(item.orgName);
  ws.cell(index+2,4).string(item.snCode);
}


const makeWorkBook = async (rows, response) => {

  var wb = new xl.Workbook();
  var ws = wb.addWorksheet('세탁');

  setConfigWorkSheet(ws);


  const dataWritePromise = rows.map((item, index) => {
    setDataInWorkSheet(ws, item, index);
  });


  const promiseResult = await Promise.all(dataWritePromise);
  console.log('PromiseResult : ', promiseResult);

  if (promiseResult) {
    wb.write('/ExcelFile.xlsx',response);
    return 200;
  }
  else return 404;


};

var cleanDirectory =  function() {
  logger.info('FINISH');
  logger.info('---------------');

  const directory2 = 'downloadImages';

  fs.readdir(directory2, (err, files) => {
    if (err) throw err;

    for (const file of files) {
      fs.unlink(path.join(directory2, file), err => {
        if (err) throw err;
      });
    }
  });
}

router.post('/', async (req, res, next) => {

    res.setHeader('Access-Control-Allow-Origin', '*');
  
    logger.info('request comming.'+JSON.parse(req.body.rows).length);
  
    try {
      const downres = await downloadImages(JSON.parse(req.body.rows));
      if(downres == 200) {

        setTimeout(async () => {
          let excelRes = await makeWorkBook(JSON.parse(req.body.rows), res);
    
          console.log('excel res : ', excelRes)
    
          if(excelRes == 200) {
            res.status(excelRes);
            logger.warn('SEND XLSX SUCCESS');
          } else {
            res.status(excelRes).json(null);
            logger.warn('Fail');
          }
          cleanDirectory();
        },1000);
      } else {
        res.status(excelRes).json(null);
        logger.warn('Fail');
      }
  
    } catch (err) {
      logger.error('API ERROR');
      next(err);
    }
  });

export default router;