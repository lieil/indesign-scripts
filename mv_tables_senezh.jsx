//v.1.2

var MY_TS = "����� ������� 1";				//����� �������
var MY_MC = "�� �������";					//����� �������� �����
var MY_HC = "��������";						//����� �������� �����
var MY_MT = "�� ����� �������";				//����� ������ ��������� ������
var MY_HT = "�� ������� ��������";			//����� ������ ������ �������� �����


with (app) {
	try {
		var doc = activeDocument;
	} catch (error) {
		alert('No documents are open!');
		exit();
	}
	
	var styleFlag = tryStyles(MY_TS, MY_MC, MY_HC, MY_MT, MY_HT) || "";
	if (styleFlag != ""){
		alert("������������ ��������� �����: \"" + styleFlag +"\"");
		exit();
	}
	
	var OldX_UNITS = doc.viewPreferences.horizontalMeasurementUnits;
	doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.POINTS;
	
	for (var i = 0; i < selection.length; i++){
		var text = getSelectionStory(selection[i]);
		if (text == undefined || text == "app") {
			corrExit("������� ��������� ����!");
		}
	
		alert("Story has " + text.tables.length + " tables");
		for (var j = 0; j < text.tables.length; j++){
			var over = text.overflows;
//			alert (over);
			mvTable(text.tables[j]);
			fixTableWidth(text.tables[j], getColumnWidth(text), over);
			}
	}
	corrExit();
};

function mvTable(table) {
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
	var hR = countHeadRows(table);
	  
	for (var i = 0; i < hR; i++){
	    with(table.rows[i].cells.everyItem()){
			appliedCellStyle = doc.cellStyles.itemByName(MY_HC);
			with(paragraphs.everyItem()){
				applyParagraphStyle(doc.paragraphStyles.itemByName(MY_HT), true);
			}	
			try { //��� ������ ����� ������ ������ ������ ������ � ������, � ��������� - ������.
				rowType = RowTypes.headerRow;
			} catch (error){};
	    }
	}
}

// ���������� ������ �������� � ������� (table) ��� �������� ������ ������� (myWidth)
function fixTableWidth(table, myWidth, over) {
	var hR = findLongestRow(table);
	var mrcs = table.rows[hR].cells;
	var maxRow = mrcs.count();
	var m = [];  // ���� � ��������
	var c = 0;   // ���� �����
	var fix = [];  // ������������
	var d = 0; // �������� ������������
	for(var i = 0; i < maxRow; i++){
		m[i] = 1;
		fix[i] = 0;
		for(var j = 0, cn = mrcs[i].parentColumn.cells; j < cn.length; j++){
			m[i] += parseInt(cn[j].words.count());
		} // ���� �� ��� ������, ��� ����� � ����� ����������� �� ������ �������, � ���� ����������.
		c += m[i];
	}
	if ( over){ //���� ���� ����������
		for(var i = 0, colWidth = 0; i < maxRow; i++){
			colWidth = myWidth*m[i]/c;
			mrcs[i].width = colWidth;
		}
	} else {
		for(var i = 0, colWidth = 0; i < maxRow; i++){
			colWidth = myWidth*m[i]/c;
			mrcs[i].width = colWidth;
			d += tryNewWidth(mrcs[i], myWidth);			
			if (colWidth < mrcs[i].width) fix[i] = mrcs[i].width - colWidth;
		}
		if (d > 1){
			for(var i = 0, colWidth = 0; i < maxRow; i++){
				if(fix[i] <= 0){
					colWidth = mrcs[i].width;
					mrcs[i].width = ((colWidth - d) > 0) ? (colWidth - d) : 1;
					tryNewWidth(mrcs[i], myWidth);
					d += mrcs[i].width - colWidth;
				}
			}
			if (d > 1) 	alert("��������, � �� ���� ��� ������ ����� ��� �������! �� ������ " + parseInt(d) + "pt");
		}
	}
	with(table.cells.everyItem()){
		minimumHeight = 0;
		autoGrow = true;
	}
}

// "��������" ������ �������� � ������������� ��� ������������
function tryNewWidth(cell, max){
	var delta = 0;
    for(var j = 0, cn = cell.parentColumn.cells; j < cn.length; j++){
		cn[j].contents; //���� �� ���������� - ������-�� �� ��������� overflow.
		if (cn[j].overflows) {
			cn[j].width += 1;
			delta++;
			j--;
			if (cell.width >= max) {
				cell.width = max;
				alert("������� �� ���������� � �������!");
				return max;
				}
		}
	}
	return delta;
}

//������� �������� ������, ������ �� �������������, ��� ������ ������ ���������� ������������ ���������� �����	  
function countHeadRows(table){
	return table.rows[0].cells[0].rowSpan;
  }

//������� ������ ������ ����� �����, � ������� ��� ������������ �����  
function findLongestRow(table){
	var hR = countHeadRows(table);
	if  (hR >= table.rows.count() - 1) return table.rows.count() - 1;
	var r = table.rows[hR].cells;
	for(var i = 0; i < r.length; i++){
		if(r[i].columnSpan > 1){
			if  (hR >= table.rows.count() - 1) return table.rows.count() - 1;
			hR++;
			i = 0;
			r = table.rows[hR].cells;
		}
	}
	return hR;
}
  
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

// ��������� ������� ������������ ������
function tryStyles(ts, mc, hc, mt, ht){
	var flag = [false,false,false,false,false];
	for (var i = 0; i < doc.tableStyles.length; i++){
		flag[0] = (ts == doc.tableStyles[i].name) || flag[0];
	}
	for (var i = 0; i < doc.cellStyles.length; i++){
		flag[1] = (mc == doc.cellStyles[i].name) || flag[1];
		flag[2] = (hc == doc.cellStyles[i].name) || flag[2];
	}
	for (var i = 0; i < doc.paragraphStyles.length; i++){
		flag[3] = (mt == doc.paragraphStyles[i].name) || flag[3];
		flag[4] = (ht == doc.paragraphStyles[i].name) || flag[4];
	}		
	for (var i = 0, str = ""; i < 5; i++){
		if (!flag[i]) str += arguments[i] + ", "; 
	}
	str = str.slice(0, -2);
	return str;
}

// ������� �������� ��������� Story
function getSelectionStory(sel){
	if (sel.constructor.name == "Application") return "app";
	if (sel.constructor.name != "TextFrame" && sel.parent.constructor.name == "Page"){
		return "app";
	} else {
		try {
			return sel.parentStory;
		} catch(error) {
			return getSelectionStory(sel.parent);
		}	
	}
}

function corrExit(mess){
	doc.viewPreferences.horizontalMeasurementUnits = OldX_UNITS;
	if ((mess != "") && (mess != undefined) ) alert(mess);
	exit();
	}
