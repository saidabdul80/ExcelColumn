var results = []
window.wdata = "";
window.result1 = [];
window.result2 = [];

var ExcelToJSON = function() {

    this.parseExcel = function (file, name, col, callback, num) {
      var reader = new FileReader();

      reader.onload = function(e) {
        var data = e.target.result;
        var workbook = XLSX.read(data, {
          type: 'binary'
        });
        var sheetsx = []
        workbook.SheetNames.forEach(function(sheetName) {
            sheetsx.push(sheetName)
        })         
        if(!sheetsx.includes(name)){
            alert("file name not exist on "+ num+ ", You can try interchange sheetname")
        }        
        var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[name]);
        callback(JSON.stringify(XL_row_object));

        //var json_object = JSON.stringify(XL_row_object.push({col:col}));        
        //console.log(XL_row_object)
        //return json_object;
       // console.log(XL_row_object)    
      };

      reader.onerror = function(ex) {
        console.log(ex);
      };

      reader.readAsBinaryString(file);
    };
    
};

function handleFileSelect(evt) {
    var sheet1name = document.getElementById("sheet1name").value;
    var sheet2name = document.getElementById("sheet2name").value;
    var column1 = document.getElementById("column1").value;
    var column2 = document.getElementById("column2").value;
    var elt = document.querySelector("#fileToUpload");
    if(sheet1name == "" || sheet2name =="" || column1 == "" || column2 == ""){        
        alert("all field is required")
        return 0;
    }
    if(elt.files.length < 2 || elt.files.length > 2){
        alert("please upload only two excel files")
        return 0;
    }
    var file1 = elt.files[0]; // FileList object
    var file2 = elt.files[1]; // FileList object
    //file number is in other they are in destination;
    document.getElementById("filenamex").innerHTML =file1.name+", "+ file2.name;
    var xl2json = new ExcelToJSON();  

    xl2json.parseExcel(file2, sheet2name, column2, function(res){
        alert("ready to generate")
        result2 = res
    }, "file 2");
    xl2json.parseExcel(file1, sheet1name, column1, function(res){        
        alert("ready to generate");
        result1 = res;
        document.querySelector("#generateK").removeAttribute('disabled');
    }, "file 1");
  //  console.log(result1, result2);
    
    
    /* 
    xl2json.parseExcel(files2); */
}


function generate(){
    //console.log(results);
    var column1 = document.getElementById("column1").value;
    var column2 = document.getElementById("column2").value;
    var rate = document.getElementById("rate").value;    
    var wb = XLSX.utils.book_new();
    wb.Props = {
            Title: "Similar data Result by column "+ column1 +" & "+ column2,
            Subject: "Test",
            Author: "Abdullahi Kawu",
            CreatedDate: new Date()
    };
    var find = false;
    var row1 = "";
    var row2 = "", item, split1, split2, wsimilar="";
    if(result1.length > result2.length){
        //results =  JSON.parse(result1).filter((item)=>{        
            item = JSON.parse(result1);
            item2 = JSON.parse(result2);
            for(var k in item){
                find = false;
                row1 = k;
                for(var l in item2){
                    /* if(item[k][column1] === 'undefined' || typeof item2[l][column2] === 'undefined' ){
                        alert("cell of a column name is empty; check row "+" "+k+"in file 1 & row "+ l +"in file 2")
                        location.reload();
                    } */

                    try{         
                        if(!item[k].hasOwnProperty(column1)){
                            break
                        }else{
                            //console.log(stringSimilarity.compareTwoStrings(item[k][column1].toLowerCase(), item2[l][column2].toLowerCase()))
                            if(stringSimilarity.compareTwoStrings(item[k][column1].toLowerCase(), item2[l][column2].toLowerCase()) >rate){
                                find = true;     
                                row2 = l;
                                break;           
                            }/* else if(item[k][column1].toLowerCase().includes(item2[l][column2].toLowerCase()) || item2[l][column2].toLowerCase().includes(item[k][column1].toLowerCase())){
                                find = true;     
                                row2 = l;
                                break;               
                            }else{
                                split1 = item[k][column1].replace("[,;.]"," ").toLowerCase().split(" ");
                                split2 = item2[l][column2].replace("[,;.]"," ").toLowerCase().split(" ");
                                split1 = split1.filter(i=>{return i != "";})
                                split2 = split2.filter(i=>{return i != "";})
                                //console.log(split1,split2);
                                //wdata += " ; "+ split1.join(", ")+" ## "+ split2.join(", ");
                                for (let index = 0; index < split1.length; index++) {
                                    if(split1[index].length > 3){
                                        find = split2.includes(split1[index]);
                                        wsimilar = split1[index]
                                        if(find){                                                                              
                                            index =  split1.length;
                                        } 
                                    }
                                }
                                /* find = split1.some(x=>{split2.includes(x)});
                                if(split1.includes("plastiras")){
                                    console.log(true)
                                } 
                                if(find){                                    
                                    row2 = l;
                                    break;
                                }
                            }   */               
                        }       
                    }catch(err){
                     console.log(err)   
                     /* alert("cell of a column name is empty; check row "+" "+k+"in file 1 & row "+ l +"in file 2")
                        location.reload(); */
                    }
                }

                if(find){
                    row1 =Number(row1) + 1;
                    row2 =Number(row2) + 1;
                    item[k].row ="row "+ row1+" & "+row2+ " similar with: "+ wsimilar;
                    results.push(item[k]);
                }
            }            
    }else{
        item = JSON.parse(result2);
        item2 = JSON.parse(result1);
        for(var k in item){
            find = false;
            row1 = k;
            for(var l in item2){
                if(item[k][column1].includes(item2[l][column2]) || item2[l][column2].includes(item[k][column1])){
                    find = true;     
                    row2 = l;
                    break;               
                    }                 
            }
            if(find){
                results.push(item[k].push({row:k +" & "+ l}));
            }
        }  
    }
    console.log(results);
    var ws_data = []
    for(let x in results){
        ws_data.push(Object.values(results[x]))
    }
    console.log(ws_data);
    var ws;
    wb.SheetNames.push('result');
    /* 
    for (let index = 0; index < results.length; index++) {
        ws_data = results[index];
    } */
    ws = XLSX.utils.aoa_to_sheet(ws_data);
    wb.Sheets['result'] = ws;            
    var wbout = XLSX.write(wb, {bookType:'xlsx',  type: 'binary'});
    function s2ab(s) {
            var buf = new ArrayBuffer(s.length);
            var view = new Uint8Array(buf);
            for (var i=0; i<s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
            return buf;
            
    }
    var dx = new Date();
    saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), 'result_'+dx+'_.xlsx');
}

(function(){
    $("#generateL").show();
    document.getElementById('generateL').addEventListener('click', handleFileSelect, false);
    document.getElementById("generateK").addEventListener("click", generate,false);
})()