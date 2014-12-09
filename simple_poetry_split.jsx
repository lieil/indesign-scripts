//v.0

/*
/* Исходные данные: поток графоманских стихов из ворда
/* заголовки - КАПСЛОКОМ 
/* перед заголовком может быть посвящение в одну строку 
/* в конце стихотворения - дата формата "\d{0,2}\s*?(январ|феврал|март|апрел|ма|июн|июл|август|сентябр|октябр|ноябр|декабр)[яаь]\s+?\d{2,4}"
/* строки рифмуются через одну, последние буквы совпадают или дают пары е-и
/* строфы отбиты абзацем
/*
/* что происходит:
/* после даты и перед посвящением разрывает страницу и переносит на следующую 
/* присваивает заголовку стиль "заголовок"
/* присваивает дате стиль "дата"
/* присваивает посвящению стиль "посвящение"
/* каждой первой строке строфы назначает стиль "поэзия_первый"
/* 
/* после этого можно запускать sZam
/*/

//стилевые константы
S_HEADER = "Header_stih"; //заголовок стиха
S_DATA = "date"; // дата создания стиха
S_INSCRIBE = "inscribe";  // стиль строки с посвящением
S_POETRY1 = "poetry-first";  //  первая строка строфы
S_POETRY = "poetry";  // основной стиль стихов
S_CHORUS = "chorus";  // выделение слова "припев"

with (app) {
	try {
		var doc = activeDocument;
	} catch (error) {
		alert('No documents are open!');
		exit();
	}
	
	var styleFlag = tryStyles(S_HEADER, S_DATA, S_INSCRIBE, S_POETRY1, S_POETRY, S_CHORUS) || "";
	if (styleFlag != ""){
		alert("Неопределены следующие стили: \"" + styleFlag +"\"");
		exit();
	}
	
	try {
		
	} catch (error) {
		alert('error');
		exit();
	}

}	

// проверяет наличие используемых стилей
function tryStyles(styles){
	alert(arguments.length);
	var flag = [];
	for(j = 0; j < 6; j++){
		flag[j] = false;
		for (var i = 0; i < doc.paragraphStyles.length; i++){
			flag[j] = (arguments[j] == doc.paragraphStyles[i].name) || flag[j];
		}		
	}
	for (var i = 0, str = ""; i < 5; i++){
		if (!flag[i]) str += arguments[i] + ", "; 
	}
	str = str.slice(0, -2);
	return str;
}
