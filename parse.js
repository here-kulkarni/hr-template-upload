var JSZip = require('jszip');
var Docxtemplater = require('docxtemplater');
var fs = require('fs');
var path = require('path');

function createOffer(data, docFile) {
    const bufferObject = docFile.buffer
    var arrayBuffer = new ArrayBuffer(bufferObject.length);
    var typedArray = new Uint8Array(arrayBuffer);
    for (var i = 0; i < bufferObject.length; ++i) {
        typedArray[i] = bufferObject[i];
    }
     // jszip 2.6 only support to Docxtemplater
    var zip = new JSZip(typedArray);
    var doc = new Docxtemplater();
    doc.loadZip(zip);
    doc.setData(
        myVariables(data)
    );
    try {
        doc.render()
        
    }
    catch (error) {
        var e = {
            message: error.message,
            name: error.name,
            stack: error.stack,
            properties: error.properties,
        }
        console.log(JSON.stringify({ error: e }));
        throw error;
    }

    var buf = doc.getZip()
        .generate({ type: 'nodebuffer' });

    fs.writeFileSync(path.resolve(__dirname + '/downloads', `${data["full_name"] || data["Full Name"]}.docx`), buf);
   
}


function pad(a, position, extra) {
    if (a?.length <= position) {
        return a;
    }
    var b = "\t\t\t\t\t\t";
    if (extra) b = b + "\t";
    return a?.substr(0, position) + b + a.substr(position).trim();
}

function myVariables(data) {
    data.full_name_32 = pad(data.full_name, 32);
    data.address_1_32 = pad(data.address_1, 32, true);
    data.address_2_32 = pad(data.address_2, 32, true);
    return data;
}


function parseSheet(data, docFile) {
   
    let head = data[0];

    let vars = [];
    for (var k in head) {
        if (head[k] == "S/o") {
            vars[k] = "so";
        }else
        if (head[k] == "Full Name") {
            vars[k] = "full_name";
        }else

        if (head[k] == "Address 1") {
            vars[k] = "address_1";
        }else

        if (head[k] == "Address 2") {
            vars[k] = "address_2";
        }else if (head[k] == "IPC Location") {
            vars[k] = "iPC_Location";
        }else

        if (head[k] == "City") {
            vars[k] = "city";
        }else
        if (head[k] == "Pin") {
            vars[k] = "pin";
        }
        if (head[k] == "Title") {
            vars[k] = "title";
        }else

        if (head[k] == "Date of Offer") {
            vars[k] = "date_of_offer";
        }else
        if (head[k] == "Date of Joining") {
            vars[k] = "date_of_joining";
        }
        // else {
        //     if(head[k] !== undefined){
        //         const lowerCaseValue= head[k].toLowerCase()?.replace(' ','_')
        //         console.log(lowerCaseValue,'lowerCaseValue')
        //         vars[k] = lowerCaseValue
        //     }
           
        // }
    }
    for (var k = 1; k < data.length; k++) {
        let raw = data[k];
        let trData = {};
        for (var a in raw) {
            if (vars[a]) {
                trData[vars[a]] = raw[a];
              
            }
        }
     
        createOffer(trData, docFile)

    }
}

module.exports = {
    parseSheet
}