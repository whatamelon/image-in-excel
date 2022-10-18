
const xl = require('excel4node');
const path = require('path');

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
    const uploadFolder = path.resolve('./images'+item.imgLinkTh.substring(lastIdx));

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
  
  
var makeWorkBook = async (rows, response) => {

    var wb = new xl.Workbook();
    var ws = wb.addWorksheet('세탁');

    setConfigWorkSheet(ws);


    const dataWritePromise = rows.map((item, index) => {
        setDataInWorkSheet(ws, item, index);
    });


    const promiseResult = await Promise.all(dataWritePromise);

    if (promiseResult) {
        wb.write('/ExcelFile.xlsx',response);
        return 200;
    }
    else return 404;


};


  export { makeWorkBook }