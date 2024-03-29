
const xl = require('excel4node');
const path = require('path');

var setConfigWorkSheet = function(ws, type) {

    //style
    ws.column(1).setWidth(10);
    ws.column(2).setWidth(10);
    ws.column(3).setWidth(15);
    ws.column(4).setWidth(20);
    ws.column(5).setWidth(20);
    ws.column(6).setWidth(20);
    ws.column(7).setWidth(20);
    ws.column(8).setWidth(15);
    ws.column(9).setWidth(15);

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
            bold:true,
            size:14
        }
    };


    // header
    ws.cell(1,1).string('상품 ID').style(headStyle);
    ws.cell(1,2).string('구분').style(headStyle);
    ws.cell(1,3).string('브랜드').style(headStyle);
    ws.cell(1,4).string('이미지').style(headStyle);
    ws.cell(1,5).string('상품명').style(headStyle);
    ws.cell(1,6).string('S/N 코드').style(headStyle);
    ws.cell(1,7).string('세탁코드').style(headStyle);
    ws.cell(1,8).string('세탁중 시작일').style(headStyle);

    if(type == 'calc') {
        ws.cell(1,9).string('세탁 입고일').style(headStyle);
        ws.cell(1,10).string('일반세탁').style(headStyle);
        ws.cell(1,11).string('특수세탁').style(headStyle);
        ws.cell(1,12).string('오점제거').style(headStyle);
        ws.cell(1,13).string('경수선').style(headStyle);
        ws.cell(1,14).string('브랜드수선').style(headStyle);
        ws.cell(1,15).string('총 세탁 요금').style(headStyle);
    }

};

  
var setDataInWorkSheet = function(ws,type , item, index) {


    let boldStyle = {
        alignment:{
            vertical: 'top'
        },
        font:{
            bold:true,
            size:14
        }
    };

    let wrapStyle = {
        alignment:{
            wrapText: true,
            vertical: 'top'
        },
        font:{
            size:12
        }
    };

    let topAlignmentStyle = {
        alignment:{
            vertical: 'top'
        },
        font:{
            size:12
        }
    }

    ws.row(index+2).setHeight(120);
    ws.cell(index+2,1).string(item.srcId.toString()).style(topAlignmentStyle);
    ws.cell(index+2,2).string(item.washType).style(topAlignmentStyle);
    ws.cell(index+2,3).string(item.brand).style(topAlignmentStyle);

    if(item.imgLinkTh != null) {
        const lastIdx = item.imgLinkTh.lastIndexOf('/');
        const uploadFolder = path.resolve('./images'+item.imgLinkTh.substring(lastIdx));
    
        let pic = ws.addImage({
            path: uploadFolder,
            type: 'picture',
            position: {
                type: 'twoCellAnchor',
                from: {
                    row: index+2,
                    colOff: "1mm",
                    col: 4,
                    rowOff: "1mm"
                },
                to: {
                    row: index+2,
                    colOff: "25mm",
                    col: 4,
                    rowOff: "45mm",
                },
            },
        });
        pic.editAs = "twoCell";

        ws.cell(index+2,5).string(item.name).style(wrapStyle);
    } else {
        ws.cell(index+2,4).string('');
        ws.cell(index+2,5).string('');
    }

    ws.cell(index+2,6).string(item.snCode).style(boldStyle);
    ws.cell(index+2,7).string(item.washCode).style(boldStyle);
    ws.cell(index+2,8).string(item.outDate).style(wrapStyle);

    if(type == 'calc') {
        ws.cell(index+2,9).string(item.inDate).style(wrapStyle);
        ws.cell(index+2,10).string(item.defaultPrice).style(topAlignmentStyle);
        ws.cell(index+2,11).string(item.specialPrice).style(topAlignmentStyle);
        ws.cell(index+2,12).string(item.pollutionPrice).style(topAlignmentStyle);
        ws.cell(index+2,13).string(item.lightPrice).style(topAlignmentStyle);
        ws.cell(index+2,14).string(item.brandPrice).style(topAlignmentStyle);
        ws.cell(index+2,15).string(item.totalPrice).style(boldStyle);
    }
}
  
  
var makeWorkBook = async (rows, type , response, fileName) => {

    var wb = new xl.Workbook();
    var ws = wb.addWorksheet('세탁');

    setConfigWorkSheet(ws, type);
    const dataWritePromise = rows.map((item, index) => {
        setDataInWorkSheet(ws, type , item, index);
    });
    const promiseResult = await Promise.all(dataWritePromise);

    if (promiseResult) {
        wb.write('/'+fileName,response);
        return 200;
    }
    else return 404;

};


export { makeWorkBook }