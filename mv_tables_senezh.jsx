var MY_TS = "Стиль таблицы 1";				//стиль таблицы
var MY_MC = "МВ Таблица";					//стиль основных ячеек
var MY_HC = "Головные";						//стиль головных ячеек
var MY_MT = "МВ Шрифт таблицы";				//стиль абзаца основного текста
var MY_HT = "МВ таблица головные";			//стиль абзаца текста головных ячеек

var MY_Q = "12.4 пт";	//квант ширины


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
		alert("Курсор должен быть в тексте, но не в таблице" +  " -> " +  selection[i].parent.constructor.name + " -> " +  selection[i].parent.parent.constructor.name);
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

// присвоение стилей всем ячейкам таблицы
    with(table.cells.everyItem()){
	  minimumHeight = 0;
	  autoGrow = true;
      appliedCellStyle = doc.cellStyles.itemByName(MY_MC);
      clearCellStyleOverrides();
	  
	  with(paragraphs.everyItem()){
		applyParagraphStyle(doc.paragraphStyles.itemByName(MY_MT), true);
	  }
	}

	
// присвоение стилей головным ячейкам таблицы	

	try {
	  var hR = countHeadRows(table);
	  var hC = table.rows[hR-1].cells.count();
	  var selection = 0;
	  for (var i = 0; i < hR; i++){
//	    table.rows[i].rowType = RowTypes.headerRow;
		alert("Строку " + i + " добавили в шапку");
 //*/
	    with(table.rows[i].cells.everyItem()){

	      appliedCellStyle = doc.cellStyles.itemByName(MY_HC);
		    alert("cell style: " + appliedCellStyle.name); alert("cell style: " + appliedCellStyle.name);
 	      with(paragraphs.everyItem()){
	  	  applyParagraphStyle(doc.paragraphStyles.itemByName(MY_HT), true);
	      }//*/
	    
	    }
	  }
//	  table.headerRowCount = hR;
	} catch (error) {
	  if (table.headerRowCount == 1){
	    alert("o-ops! " + error);
	    return;
	  } else {
	    alert(error + " " + "in " + ++tableErrs + " tables"); 	// нужно будет проработать вариант, когда шапка состоит из объединенных ячеек 
											// пока просто считаются таблички, в которых возникает проблема
	  }
	}
}

function fixTableWidth(table, myWidth) {
  var hR = countHeadRows(table);
  maxRow = table.rows[hR].cells.count();
  alert("Column width: " +  myWidth/maxRow);
/*  with(table.columns.everyItem()){
    width = myWidth/table.columnCount;
  } //*/
  var m = [];  // слов в колонках
  var w = [];  // ширина колонки
  var c = 0;   // слов всего
  var o = [];  // есть ли переполнение (фиксирование ширины)
  var f = 0;   // общая ширина фиксированых колонок
  for(var i = 0; i < maxRow; i++){
    var cn = table.rows[hR].columns[i];
	m[i] = 0;
	w[i] = 0;
	o[i] = 0;
	for(var j = 0; j < cn.cells.length; j++){
	  var cnw = cn.cells[j];
	  m[i] += parseInt(cnw.words.count());
//	  alert("Слов в ячейках " + i + ": " + cnw.words.count() + "или " + m[i]);
	} 
//  	alert("Слов в колонке " + i + ": " + m[i]);	
	c += m[i];
  }
  alert("Слов в таблице " + i + ": " + c);
  for(var i = 0; i < maxRow; i++){
    var cn = table.rows[hR].columns[i];
    w[i] = myWidth*m[i]/c;
	cn.width = tryNewWidth(cn, w[i], myWidth);
  }
}

// вспомогательные функции
function tryNewWidth(column, width, max){
	column.width = width;
    for(var j = 0; j < column.cells.length; j++){
//	  alert("переполнение: " + column.cells[j].overflows);
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
		alert("мах. длина слова в колонке " + i + ": " + w[i] + ", длина текущего слова" + d);
	  } //*/
	  
function countHeadRows(table){	//считаем головные строки
/*	var hR = 0;
	var header = table.rows[hR].cells;
	var hh
	for(var i = 0; i < header.length; i++){
	  hh = header[i].columnSpan;
	  alert("Cell " + header[i].name + "span from" + hh + "columns");
	  if (header[i].columnSpan > 1){
		alert("много колонок")
	  } else {
		alert("Cell " + header[i].name + "span from" + header[i].rowSpan + "rows");
	  }
	} //*/  return(2);
  }	  


// возвращает ширину первой колонки текста. Спасибо тому, у кого этот кусок позаимстовала.
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