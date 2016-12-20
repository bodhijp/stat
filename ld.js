google.charts.load('current', {'packages':['table']});
google.charts.setOnLoadCallback(drawCn);

function drawTable(docurl , sort ) {
	
	var load = function(col){
		
		var qstr ='https://docs.google.com/spreadsheets/d/'+ docurl + '/gviz/tq?tq='				  
		  + encodeURIComponent('select B,J where not (' + col +' matches "")');
		  //order by B desc

		var tblDiv = document.getElementById('tbl_'+col);
		tblDiv.innerHTML="";
		var csvn = "报名表_"+ $(".am-active")[0].innerText.trim() +"--" + tblDiv.parentElement.innerText.trim();
	  
		var query = new google.visualization.Query(qstr); 

		query.send(function(response){
			var data = response.getDataTable();

            if ( sort === 'cn'){
                data.addColumn('string', '缩写');
                var rowcnt = data.getNumberOfRows();
                for (var i = 0; i < rowcnt; i++) {
                    var v = data.getValue(i, 0);
                    var py = pinyinUtil.getFirstLetter(v, false);
                    data.setValue(i,2,py);
                }
               data.sort(2);
            }
                            
			var table = new google.visualization.Table(tblDiv);
			table.draw(data, {showRowNumber: true,alternatingRowStyle:false, width: '100%', height: '100%'});	  

			// This must be a hyperlink
			$("#dl_"+col).on('click', function (event) {

                $(this).attr({ 'download': csvn + '.xls', 'href': getXls($("#tbl_" + col + " table")[0]), 'target': '_blank' }); 

				//exportTableToCSV.apply(this, [ $("#tbl_" + col + " table") , csvn+'.csv']);
				//window.location.href = getXls($("#tbl_" + col + " table")[0]);
				// IF CSV, don't do event.preventDefault() or return false
				// We actually need this to be a typical hyperlink
			});

		});	  
			
	};
	
	load('F');
	load('G');
	load('H');
	load('I'); 
    //console.log(pinyinUtil.getFirstLetter('好', false));
}

function drawCn(){
	drawTable('1oAurf_DFFJh30t1vSjlQOu8ij3c4xRTTDRtq0qdNGOM' , 'cn');
}

function drawJp(){
	drawTable('1r-f1LEApo-4frBcI5m1hESZPRJTFoD6o6QplCCvtx-o');
}

$("#chg_cn").on('click', function (event) {
	drawCn();
	$("#chg_jp").removeClass("am-active");
	$("#chg_cn").addClass("am-active");
});

$("#chg_jp").on('click', function (event) {
	drawJp();	
	$("#chg_cn").removeClass("am-active");
	$("#chg_jp").addClass("am-active");
});

function printTable(divid){
	var thediv = document.getElementById(divid);
	if(thediv==null)return;
	var tbl  = thediv.querySelector("table");
	if(tbl==null)return;


	var frm = document.createElement('iframe');
	frm.src="javascript:false";

	var where = document.getElementsByTagName('script')[0];
	where.parentNode.insertBefore(frm, where);

	var doc = frm.contentWindow.document;
	doc.write(tbl.outerHTML);

	frm.contentWindow.print();

}


function getXls(table){  
        var uri = 'data:application/vnd.ms-excel;base64,';
        var template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" '+
        'xmlns:x="urn:schemas-microsoft-com:office:excel" '+
        'xmlns="http://www.w3.org/TR/REC-html40">'+
        '<head>'+
            '<!--[if gte mso 9]>'+
            '<xml>'+
                '<x:ExcelWorkbook>'+
                    '<x:ExcelWorksheets>'+
                        '<x:ExcelWorksheet>'+
                            '<x:WorksheetOptions>'+
                                '<x:DisplayGridlines/>'+
                            '</x:WorksheetOptions>'+
                        '</x:ExcelWorksheet>'+
                    '</x:ExcelWorksheets>'+
                '</x:ExcelWorkbook>'+
            '</xml>'+
            '<![endif]-->'+
        '</head>'+
        '<body>'+
            '<table>{table}</table>'+
        '</body>'+
        '</html>';
    
        /**
         * 转码 base 64
         * window.btoa能从ascii/二进制流中创建一个base64编码的字符串
         * escape编码  unescape 解码字符串
         * encodeURIComponent编码  DecodeURIComponent 解码字符串
         **/
        var base64 = function(s) {
            return window.btoa(unescape(encodeURIComponent(s)));
        };

        //返回替换完具体数据的xls模板
        var getXlsXml = function(template,data) {
            return template.replace(/{(\w+)}/g,data);
        };

        //返回资源链接
        return uri+base64(getXlsXml(template, table.innerHTML));
    };



function exportTableToCSV($table, filename) {

    var $rows = $table.find('tr:has(td),tr:has(th)'),

        // Temporary delimiter characters unlikely to be typed by keyboard
        // This is to avoid accidentally splitting the actual contents
        tmpColDelim = String.fromCharCode(11), // vertical tab character
        tmpRowDelim = String.fromCharCode(0), // null character

        // actual delimiter characters for CSV format
        colDelim = '","',
        rowDelim = '"\r\n"',

        // Grab text from table into CSV formatted string
        csv = '"' + $rows.map(function (i, row) {
            var $row = $(row), $cols = $row.find('td,th');

            return $cols.map(function (j, col) {
                var $col = $(col), text = $col.text();

                return text.replace(/"/g, '""'); // escape double quotes

            }).get().join(tmpColDelim);

        }).get().join(tmpRowDelim)
            .split(tmpRowDelim).join(rowDelim)
            .split(tmpColDelim).join(colDelim) + '"',



        // Data URI
        csvData = 'data:application/csv;charset=utf-8,' + encodeURIComponent(csv);

        //console.log(csv);

        if (window.navigator.msSaveBlob) { // IE 10+
            //alert('IE' + csv);
            window.navigator.msSaveOrOpenBlob(new Blob([csv], {type: "text/plain;charset=utf-8;"}), "csvname.csv")
        } 
        else {
            $(this).attr({ 'download': filename, 'href': csvData, 'target': '_blank' }); 
        }
}

function print2pdf(){
	var doc = new jsPDF();

	// We'll make our own renderer to skip this editor
	var specialElementHandlers = {
		'#editor': function(element, renderer){
			return true;
		}
	};

	// All units are in the set measurement for the document
	// This can be changed to "pt" (points), "mm" (Default), "cm", "in"
	doc.fromHTML($('#tbl_F').get(0), 15, 15, {
		'width': 170, 
		'elementHandlers': specialElementHandlers
	});
	
	doc.save('Test.pdf');
}