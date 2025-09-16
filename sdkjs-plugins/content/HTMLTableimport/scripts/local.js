/*
 (c) VNexsus 2023

 Licensed under the Apache License, Version 2.0 (the "License");
 you may not use this file except in compliance with the License.
 You may obtain a copy of the License at

     http://www.apache.org/licenses/LICENSE-2.0

 Unless required by applicable law or agreed to in writing, software
 distributed under the License is distributed on an "AS IS" BASIS,
 WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 See the License for the specific language governing permissions and
 limitations under the License.
 */
 
window.translate = function(term){
	return window.localization[parent.Common.Locale.getCurrentLanguage()][term] ? window.localization[parent.Common.Locale.getCurrentLanguage()][term] : window.localization["en"][term] || "";
}

window.localization = {
	"en":{
		
		"sourcePaneTitle": "Import data from file",
		"buttonSelect": "Select",
		"buttonClear": "Clear",
		
		"destinationPaneTitle": "Import data to",
		"destinationPaneNewWorksheet": "new worksheet",
		"destinationPaneExistingWorksheet": "existing worksheet",
		
		"warningTitle": "Warning",
		"warningFileNotSelected": "Select file for import",
		"warningFileEmpty": "Selected file is empty"
		
	},
	"ru":{

		"sourcePaneTitle": "Загрузить таблицу из файла",
		"buttonSelect": "Выбрать",
		"buttonClear": "Очистить",

		"destinationPaneTitle": "Импортировать данные в",
		"destinationPaneNewWorksheet": "новый лист",
		"destinationPaneExistingWorksheet": "существующий лист",
		
		"warningTitle": "Предупреждение",
		"warningFileNotSelected": "Выберите файл для импорта",
		"warningFileEmpty": "Выбранный файл пуст"
	}
}
