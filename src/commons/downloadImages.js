const fs = require('fs');
const path = require('path');
const request = require('request');

var download = function(uri, uploadFolder, callback){
    request.head(uri, function(err, res, body){
      request(uri).pipe(fs.createWriteStream(uploadFolder)).on('close', callback);
    });
};
  
  
var downloadImages = async function (fileList) {
    const imgCropPromise = fileList.map(async (item) => {

        if(item.srcType == 'o') {
            const lastIdx = item.imgLinkTh.lastIndexOf('/');
            const uploadFolder = path.resolve('./images'+item.imgLinkTh.substring(lastIdx));

            download(item.imgLinkTh, uploadFolder, function(){});
        }
    });

    const promiseResult = await Promise.all(imgCropPromise);

    if (promiseResult) return 200;

    return 404;
}

export { downloadImages }