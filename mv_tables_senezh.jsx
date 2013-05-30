var MY_TS = "����� ������� 1";				//����� �������
var MY_MC = "�� �������";					//����� �������� �����
var MY_HC = "��������";						//����� �������� �����
var MY_MT = "�� ����� �������";				//����� ������ ��������� ������
var MY_HT = "�� ������� ��������";			//����� ������ ������ �������� �����

var MY_Q = "12.4 ��";	//����� ������


with (app) {
	try {
		var doc = activeDocument;
	} catch (error) {
		alert('No documents are open!');
		exit();
	}
	
  var tableErrs = 0;
  for (var i = 0; i < selection.length; i++){
	
	try {
		var text = selection[i].parentStory;
	} catch (error) {
//		if
		alert("������ ������ ���� � ������, �� �� � �������" +  " -> " +  selection[i].parent.constructor.name + " -> " +  selection[i].parent.parent.constructor.name);
		exit();
	}
    alert("Story has " + text.tables.length + " tables");
	for (var i = 0; i < text.tables.length; i++){
    mvTable(text.tables[i]);
    fixTableWidth(text.tables[i], getColumnWidth(text));
	}
  }

}

function mvTable(table) {
//  alert("table style: " + table.appliedTableStyle.name);
  table.clearTableStyleOverrides();
  table.appliedTableStyle = doc.tableStyles.itemByName(MY_TS);

// ���������� ������ ���� ������� �������
    with(table.cells.everyItem()){
	  minimumHeight = 0;
	  autoGrow = true;
      appliedCellStyle = doc.cellStyles.itemByName(MY_MC);
      clearCellStyleOverrides();
	  
	  with(paragraphs.everyItem()){
		applyParagraphStyle(doc.paragraphStyles.itemByName(MY_MT), true);
	  }
	}

	
// ���������� ������ �������� ������� �������	
  countHeadRows = function(table){	//������� �������� ������
/*	var hR = 0;
	var header = table.rows[hR].cells;
	var hh
	for(var i = 0; i < header.length; i++){
	  hh = header[i].columnSpan;
	  alert("Cell " + header[i].name + "span from" + hh + "columns");
	  if (header[i].columnSpan > 1){
		alert("����� �������")
	  } else {
		alert("Cell " + header[i].name + "span from" + header[i].rowSpan + "rows");
	  }
	} //*/  return(2);
  }
	try {
	  var hR = countHeadRows(table);
	  var hC = table.rows[hR-1].cells.count();
	  var selection = 0;
	  for (var i = 0; i < hR; i++){
	    alert("���: " + table.rows[0].cells[0].name + ", " + table.rows[i].cells[table.rows[i].cells.count()-1].name);
	    selection =  table.rows[i].cells.everyItem();
		alert(selection.constructor.name);
 //*/
	  with(selection){
//	    alert("cell style: " + appliedCellStyle.name);
/*	    appliedCellStyle = doc.cellStyles.itemByName(MY_HC);
 	    with(paragraphs.everyItem()){
		  applyParagraphStyle(doc.paragraphStyles.itemByName(MY_HT), true);
	    }//*/
	    rowType = RowTypes.headerRow;
	  }
	  };
	} catch (error) {
	  if (table.headerRowCount == 1){
	    alert("o-ops! " + error);
	    return;
	  } else {
	    alert(error + " " + "in" + ++tableErrs + " tables"); 	// ����� ����� ����������� �������, ����� ����� ������� �� ������������ ����� 
											// ���� ������ ��������� ��������, � ������� ��������� ��������
	  }
	}
}

function fixTableWidth(table, myWidth) {
  alert("Column width: " +  myWidth/table.columnCount);
/*  with(table.columns.everyItem()){
    width = myWidth/table.columnCount;
  } //*/
  var m = [];  // ���� � ��������
  var w = [];  // ������ �������
  var c = 0;   // ���� �����
  var o = [];  // ���� �� ������������ (������������ ������)
  var f = 0;   // ����� ������ ������������ �������
  for(var i = 0; i < table.columnCount; i++){
    var cn = table.columns[i];
	m[i] = 0;
	w[i] = 0;
	o[i] = 0;
	for(var j = 0; j < cn.cells.length; j++){
	  var cnw = cn.cells[j];
	  m[i] += parseInt(cnw.words.count());
//	  alert("���� � ������� " + i + ": " + cnw.words.count() + "��� " + m[i]);
	} 
//  	alert("���� � ������� " + i + ": " + m[i]);	
	c += m[i];
  }
  alert("���� � ������� " + i + ": " + c);
  for(var i = 0; i < table.columnCount; i++){
    var cn = table.columns[i];
    w[i] = myWidth*m[i]/c;
	cn.width = tryNewWidth(cn, w[i], myWidth);
  }
}

// ��������������� �������
function tryNewWidth(column, width, max){
	column.width = width;
    for(var j = 0; j < column.cells.length; j++){
//	  alert("������������: " + column.cells[j].overflows);
	  while (column.cells[j].overflows) {
//	  	alert (column.width);
	    column.width += 1;
		if (column.width > max) return(max);
	  }
	}
	return (column.width);  
}
/*	  for(var k = 0; k < cnw.words.length; k++ ){
	    var d = cnw.words[k].characters.length;
		w[i] = (w[i] < d) ? d : w[i];
		alert("���. ����� ����� � ������� " + i + ": " + w[i] + ", ����� �������� �����" + d);
	  } //*/


// ���������� ������ ������ ������� ������. ������� ����, � ���� ���� ����� �������������.
function getColumnWidth(text) {
  var tf = text.textColumns[0].parentTextFrames[0];
  try{
	var tfWidth = tf.geometricBounds[3]-tf.geometricBounds[1];
	var pref = tf.textFramePreferences;
	if (pref.useFixedColumnWidth) {
		return pref.textColumnFixedWidth;
	} else {
		return (tfWidth - pref.textColumnGutter*(pref.textColumnCount-1))/pref.textColumnCount;
	}
  } catch (error) {
    alert(error);
  }	
}