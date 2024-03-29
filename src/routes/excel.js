import e from 'express';
import express from 'express';

import { downloadImages } from '../commons/downloadImages';
import { makeWorkBook } from '../commons/worksheet';

const fs = require('fs');
const path = require('path');

const router = express.Router();
import { logger } from '../commons/log';



var cleanDirectory =  function() {
  logger.info('FINISH');
  logger.info('---------------');

  const directory2 = 'images';

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
    // res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    // res.setHeader("Content-Disposition", "attachment; filename=" + req.body.fileName);
  
    console.log(req.body.type)
    console.log(req.body.rows.length)
    console.log(req.body.filename)
    logger.info('request comming type : '+ req.body.type);
    logger.info('request comming data : '+ req.body.rows.length);
  
    try {
      const downres = await downloadImages(req.body.rows);
      logger.info('downres : good')
      if(downres == 200) {

        setTimeout(async () => {
          let excelRes = await makeWorkBook(req.body.rows, req.body.type , res, req.body.filename);
    
          logger.info('excel res : good')
    
          if(excelRes == 200) {
            res.status(excelRes);
            logger.info('SEND XLSX SUCCESS');
          } else {
            res.status(excelRes).json(null);
            logger.warn('SEND XLSX FAIL');
          }
          cleanDirectory();
        },10000);
      } else {
        res.status(excelRes).json(null);
        logger.warn('DOWNLOAD IMAGE FAIL');
      }
  
    } catch (err) {
      logger.error('API ERROR');
      next(err);
    }
  });

export default router;