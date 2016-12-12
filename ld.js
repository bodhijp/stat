google.charts.load('current', {'packages':['table']});
google.charts.setOnLoadCallback(drawTable);




function drawTable() {
	
	var load = function(col){
		
		var qstr =  'https://docs.google.com/spreadsheets/d/1oAurf_DFFJh30t1vSjlQOu8ij3c4xRTTDRtq0qdNGOM/gviz/tq?tq='				  
		  + encodeURIComponent('select B where not (' + col +' matches "")');
		  //order by B desc
	  
		var query = new google.visualization.Query(qstr); 

		query.send(function(response){
			var data = response.getDataTable();
			var table = new google.visualization.Table(document.getElementById('tbl_'+col));
			table.draw(data, {showRowNumber: true,alternatingRowStyle:false, width: '100%', height: '100%'});	  
		});	  
			
	};
	
	load('F');
	load('G');
	load('H');
	load('I');
  

}