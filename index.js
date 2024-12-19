const objectInfo = require('./objectInfo.json'); // Relative path to the file
const profileInfo = require('./profileInfo.json');
const data = require('./data.json'); // Relative path to the file
var xl = require('excel4node');

var permission_Map = prepareMap();

function setUpInitialHeader(ws,wb)
{
    var ObjectHeaderstyle = wb.createStyle({
        font: {
          color: 'black',
          size: 12,
          bold: true
        },
        alignment: {
          horizontal: 'center',  // Horizontal alignment (center)
          vertical: 'center'    // Vertical alignment (center)
        },
        fill: {
            type: 'pattern',      // Type of fill (pattern)
            patternType: 'solid', // Fill pattern (solid)
            bgColor: '#b8d5f1',   // Background color (use if you want background in patterns)
            fgColor: '#b8d5f1'    // Fill color (actual cell background)
        },
        border: {
            left: {
                style: 'thin',      // Style of the left border
                color: '#000000'    // Color of the left border
            },
            right: {
                style: 'thin',      // Style of the right border
                color: '#000000'    // Color of the right border
            },
            top: {
                style: 'thin',      // Style of the top border
                color: '#000000'    // Color of the top border
                },
            bottom: {
                style: 'thin',      // Style of the bottom border
                color: '#000000'    // Color of the bottom border
            }
        }
      });

    ws.cell(1, 1, 3, 1, true).string('Object Names').style(ObjectHeaderstyle);
    ws.cell(1, 2, 3, 2, true).string('Object API Names').style(ObjectHeaderstyle);

    object_current_row_index = 4;
    for(let i = 0;i<objectInfo.length;i++)
    {
        ws.cell(object_current_row_index, 1).string(objectInfo[i].Label);
        ws.cell(object_current_row_index, 2).string(objectInfo[i].QualifiedApiName);

        object_current_row_index++;
    }
}


function setUpProfilePermission(ws,wb)
{

    let ExColIndex = 3;
    let currentRow = 0;
    let lastColmIndex = 0;
    let key;
    for(let i = 0;i<objectInfo.length;i++)
    {
      ExColIndex = 3;
        for(let j = 0;j<profileInfo.length;j++)
        {
            if(i == 0)
            {
                let profileHeaderstyle = wb.createStyle({
                    font: {
                      color: 'black',
                      size: 12,
                      bold: true
                    },
                    alignment: {
                        vertical: 'center'    // Vertical alignment (center)
                    },
                    fill: {
                        type: 'pattern',      // Type of fill (pattern)
                        patternType: 'solid', // Fill pattern (solid)
                        bgColor: '#3e9cd9',   // Background color (use if you want background in patterns)
                        fgColor: '#3e9cd9'    // Fill color (actual cell background)
                    },
                    border: {
                        left: {
                            style: 'thin',      // Style of the left border
                            color: '#000000'    // Color of the left border
                        },
                        right: {
                            style: 'thin',      // Style of the right border
                            color: '#000000'    // Color of the right border
                        },
                        top: {
                            style: 'thin',      // Style of the top border
                            color: '#000000'    // Color of the top border
                            },
                        bottom: {
                            style: 'thin',      // Style of the bottom border
                            color: '#000000'    // Color of the bottom border
                        }
                    }
                  });
                ws.cell(2, ExColIndex, 2, ExColIndex + 5, true).string(profileInfo[j].Name).style(profileHeaderstyle);
                writethirdRow(ws,ExColIndex,wb);
            }

            currentRow = i + 4;
            key = objectInfo[i].QualifiedApiName+profileInfo[j].Name;
            if(permission_Map.has(key))
            {
                writePermission(currentRow,ExColIndex,permission_Map.get(key),ws);
            }
            else
            {
                let AllFalsePermissionObj = {
                    "PermissionsRead": false,
                    "PermissionsCreate": false,
                    "PermissionsEdit": false,
                    "PermissionsDelete": false,
                    "PermissionsViewAllRecords": false,
                    "PermissionsModifyAllRecords": false,
                }

                writePermission(currentRow,ExColIndex,AllFalsePermissionObj,ws);
            }

            ExColIndex += 6;
        }

        lastColmIndex = ExColIndex - 1;
    }


    let profileFirstHeader = wb.createStyle({
        font: {
          color: 'white',
          size: 12,
          bold: true
        },
        fill: {
            type: 'pattern',      // Type of fill (pattern)
            patternType: 'solid', // Fill pattern (solid)
            bgColor: '#0a598c',   // Background color (use if you want background in patterns)
            fgColor: '#0a598c'    // Fill color (actual cell background)
        }
      });

    ws.cell(1, 3, 1,lastColmIndex, true).string('Profile Names').style(profileFirstHeader);
   
}


function main()
{
    var wb = new xl.Workbook();
    var ws = wb.addWorksheet('Sheet 1');

    setUpInitialHeader(ws,wb);

    ws.column(1).setWidth(30);
    ws.column(2).setWidth(30); 

    setUpProfilePermission(ws,wb);

    wb.write('Excel.xlsx');

    console.log('Done..');
}

function prepareMap()
{
    const permissionMap = new Map();
    for(let i = 0;i<data.length;i++)
    {
        let key = data[i].SobjectType+data[i].Parent.Profile.Name;
        permissionMap.set(key,data[i]);
    }

    return permissionMap;
}


function writethirdRow(ws,Colindex,wb)
{
    let PermissionHeaderstyle = wb.createStyle({
        font: {
          color: 'black',
          size: 12,
          bold: true
        },
        fill: {
            type: 'pattern',      // Type of fill (pattern)
            patternType: 'solid', // Fill pattern (solid)
            bgColor: '#ca5a22',   // Background color (use if you want background in patterns)
            fgColor: '#ca5a22'    // Fill color (actual cell background)
        }
      });
    let array_row_title = ['R','C','E','D','View All','Modify All'];

    for(let i = 0;i<array_row_title.length;i++)
    {
        ws.cell(3, Colindex).string(array_row_title[i]).style(PermissionHeaderstyle);
        Colindex++;
    }
}


function writePermission(row,col,permission,ws)
{
     ws.cell(row,col).string(permission.PermissionsRead ? 'Y' : 'N');
     ws.cell(row,col+1).string(permission.PermissionsCreate ? 'Y' : 'N');
     ws.cell(row,col+2).string(permission.PermissionsEdit ? 'Y' : 'N');
     ws.cell(row,col+3).string(permission.PermissionsDelete ? 'Y' : 'N');
     ws.cell(row,col+4).string(permission.PermissionsViewAllRecords ? 'Y' : 'N');
     ws.cell(row,col+5).string(permission.PermissionsModifyAllRecords ? 'Y' : 'N');
}

main();