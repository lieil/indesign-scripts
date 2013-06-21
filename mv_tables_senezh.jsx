//v.0.2

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
	
  for (var i = 0; i < selection.length; i++){
	
	try {
		var text = selection[i].parentStory;
	} catch (error) {
//		if	здесь должна быть учтена ситуация, когда выделена часть текста внутри таблицы. Пока просто выдает ошибку.
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
	var hR = countHeadRows(table);
	  
	for (var i = 0; i < hR; i++){
	    with(table.rows[i].cells.everyItem()){
	      appliedCellStyle = doc.cellStyles.itemByName(MY_HC);
 	      with(paragraphs.everyItem()){
			applyParagraphStyle(doc.paragraphStyles.itemByName(MY_HT), true);
	      }	
		  try { //тип строки можно задать только первой ячейке в строке, в остальных - ошибка.
			  rowType = RowTypes.headerRow;
		  } catch (error){}
	    }
	}
}

// вычисление ширины столбцов в таблице (table) под заданную ширину столбца (myWidth)
function fixTableWidth(table, myWidth) {
	var hR = findLongestRow(table);
	var mrcs = table.rows[hR].cells
	var maxRow = mrcs.count();
	var k = 0;
	var m = [];  // слов в колонках
	var c = 0;   // слов всего
	var fix = [];  // есть ли переполнение (фиксирование ширины)
	var delta = 0;   // величина переполнения
	for(var i = 0; i < maxRow; i++){
		m[i] = 1;
		fix[i] = false;
		for(var j = 0, cn = mrcs[i].parentColumn.cells; j < cn.length; j++){
			m[i] += parseInt(cn[j].words.count());
		}
		c += m[i];
	}
	for(var i = 0; i < maxRow; i++){
//		alert("Столбец" + i + ": "+ fix[i]);
		if (!fix[i]){
			w = ((myWidth*m[i]/c - delta - 4) < 0) ? myWidth*m[i]/c : (myWidth*m[i]/c - delta);
			alert("Столбец" + i + ": "+ parseInt(w) + "пт, ширина" + parseInt(myWidth*m[i]/c) + ", поправка" + delta);
			delta = tryNewWidth(mrcs[i], w, myWidth) + (myWidth*m[i]/c - w);
			if (delta > 0) fix[i] = true;
		}
		if ((i >= maxRow-1) && (delta > 0)){
			alert("Переполнение таблицы, возвращаемся");
						i = 0;
			k++;
			if (k > 3) break;
		}
	}
}

// "Примерка" ширины столбцов и корректировка при переполнении
function tryNewWidth(cell, width, max){
	var d = 0;
	cell.width = width;
//*
    for(var j = 0, cn = cell.parentColumn.cells; j < cn.length; j++){
		cn[j].contents; //если не обратиться - почему-то не считается overflow.
		if (cn[j].overflows) {
//			alert("переполнение: " + cn[j].name + ", " + cn[j].contents);
//			cn[j].contents;
			cn[j].width += 1;
			d++;
			j--;
			if (cell.width >= max) {
				cell.width = max;
				alert("Таблица не помещается в колонку!")
				break;
				}
		}
	}
//*/
	return d;
}
/*
	  for(var k = 0; k < cnw.words.length; k++ ){
	    var d = cnw.words[k].characters.length;
		w[i] = (w[i] < d) ? d : w[i];
		alert("мах. длина слова в колонке " + i + ": " + w[i] + ", длина текущего слова" + d);
	  } //*/

//считаем головные строки, исходя из предположения, что первая ячейка объединяет максимальное количество строк	  
function countHeadRows(table){
	return table.rows[0].cells[0].rowSpan;
  }

//находит первую строку после шапки, в которой нет объединенных строк  
function findLongestRow(table){
	var hR = countHeadRows(table), r = table.rows[hR].cells;
	for(var i = 0; i < r.length; i++){
		if(r[i].columnSpan > 1){
			hR++;
			i = 0;
		}
	}
	return hR;
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