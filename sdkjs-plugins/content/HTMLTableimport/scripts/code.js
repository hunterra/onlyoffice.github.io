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

(function(window, undefined){

	window.model = null;

    window.Asc.plugin.init = function()
    {
		
		window.model = new importModel();
		ko.applyBindings(window.model);		
		
	};

	window.Asc.plugin.selectFile = function() {
		if(document.getElementById("selectFile").value == translate('buttonClear')){
			window.model.data(undefined);
			document.getElementById("selectFile").value = translate('buttonSelect')
			return;
		}
		if(window["AscDesktopEditor"])
			window["AscDesktopEditor"]["OpenFilenameDialog"]("Файлы c HTML таблицами (*.html *.xls)", false, function(_file){
					var file = _file;
					if (Array.isArray(file)){
						window.model.filepath(file[0]);
						file = file[0];
						window["AscDesktopEditor"]["loadLocalFile"](file,window.Asc.plugin.logoSelected)
					}
			});		
		else
			document.getElementById("uploadLogo").click();
	}
	
	window.Asc.plugin.logoSelected = function(data) {
		if(data.length > 0){
			window.model.data(new TextDecoder().decode(data));
			document.getElementById("selectFile").value = translate('buttonClear')
		}
	}
	
	window.Asc.plugin.logoUploaded = function(e) {
		var arrFiles = e.target.files;
		if(arrFiles.length && arrFiles.length > 0 && arrFiles[0].type && arrFiles[0].type.indexOf('text')!=-1) {
			var oFileReader = new FileReader();
			oFileReader.onloadend = function() {
				window.data.logo = oFileReader.result;
				document.getElementById("selectFile").value = translate('buttonClear')
			}
			oFileReader.readAsDataURL(arrFiles[0]);
		}
	}
	
	window.Asc.plugin.getSheets = function(){
		var sheets = [];
		parent.Asc.editor.Sheets.forEach(function(sheet, index){
			sheets.push(new Sheet({name: sheet.Name, index: index}))
		});
		return sheets;
	}
	
	window.Asc.plugin.getActiveSheetIndex = function(){
		return parent.Asc.editor.GetActiveSheet().Index;
	}

	window.Asc.plugin.setSheet = function(sheet){
		Asc.scope.sheet = sheet;
		window.Asc.plugin.callCommand(function() {
			var oSheets = Api.GetSheets();
			oSheets[Asc.scope.sheet.index].SetActive();
		}, false, true, function(){
			var oSheet = parent.Asc.editor.GetActiveSheet();
			var oActiveCell = oSheet.GetActiveCell();
			var sAddress = oActiveCell.GetAddress(false, false, "xlA1", false);
			sheet.celladdress(sAddress);
		});
	}

	window.Asc.plugin.getCellAddress = function(){
		var mask = $(window.parent.document).find(".modals-mask");
		if(mask.length != 0)
			$(mask[0]).css("display", "none");
		
		var plugin_window = $(parent.document.getElementById("iframe_" + Asc.plugin.guid)).closest(".asc-window");

		var win = new parent.SSE.Views.CellRangeDialog({
			handler: function(dlg, result){
					if (result == 'ok') {
						var txt = dlg.getSettings();
						var cellrng = txt.split(":");
						window.model.worksheet().celladdress(cellrng[0].replace(/\$/g,""));
					}
				}
			}).on('close', function() {
				plugin_window.css("display", "block");
				$(mask[0]).css("display", "block");
			});
		plugin_window.css("display", "none");
		var xy = {left: 0, top: 0};
		win.show(xy.left + 160, xy.top + 125);
		win.setSettings({
			api     : parent.g_asc_plugins.api, //me.api,
			range	: window.model.worksheet().celladdress(),
			type    : parent.Asc.c_oAscSelectionDialogType.FormatTable//PivotTableData
		});
	}

    window.Asc.plugin.button = function(id){
		if(id == 0){
			if(!window.model.filepath()){
				window.parent.Common.UI.warning({title: translate("warningTitle"), msg: translate("warningFileNotSelected")});
				return;
			}
			else if(!window.model.data()){
				window.parent.Common.UI.warning({title: translate("warningTitle"), msg: translate("warningFileEmpty")});
				return;
			}
			window.Asc.plugin.import();
			this.executeCommand("close", "");
		}
		else
			this.executeCommand("close", "");
	};

	window.Asc.plugin.import = function(){
		var parser = new DOMParser();
		var doc = parser.parseFromString(window.model.data(),"text/html");
		// TODO: handle errors
		
		var oWorksheet;
		if(!window.model.newworksheet())
			oWorksheet = parent.Asc.editor.Sheets[window.model.worksheet().index];
		else{
			parent.Asc.editor.AddSheet(window.model.filename().substring(0,(window.model.filename().lastIndexOf(".") != -1 ? window.model.filename().lastIndexOf(".") : 30)));
			oWorksheet = parent.Asc.editor.GetActiveSheet();
		}
		var oStartCell = oWorksheet.GetRange((!window.model.newworksheet() ? window.model.worksheet().celladdress() : "A1"));
		var startrow = oStartCell.GetRow();
		var startcol = oStartCell.GetCol();
		var currentrow = startrow;
		var currentcol = startcol;

		var border = false;
		var nodes = doc.getElementsByTagName("*");
		for(i = 0; i < nodes.length; i++){
			switch(nodes[i].nodeName){
				case "P":
				case "SPAN":
				case "DIV":
					var oCell = oWorksheet.GetRangeByNumber(currentrow, currentcol);
					oCell.SetValue(nodes[i].innerText);
					border = false;
					currentrow++;
					break;
				case "TABLE":
					if(nodes[i].border && nodes[i].border != "")
						border = true;
					tablestart = true;
					break;
				case "TR":
					currentcol = startcol;
					if(!tablestart) 
						currentrow++;
					else
						tablestart = false;
					break;
				case "TD":
					var oCell = oWorksheet.GetRangeByNumber(currentrow, currentcol);
					oCell.SetValue(nodes[i].innerText);
					if(nodes[i].style.backgroundColor && nodes[i].style.backgroundColor != ""){
						var re = /rgb\((\d+),\s*(\d+),\s*(\d+)\)/i;
						var color = nodes[i].style.backgroundColor.match(re);
						if(color.length == 4){
							oCell.SetFillColor(parent.g_asc_plugins.api.CreateColorFromRGB(parseInt(color[1]),parseInt(color[2]),parseInt(color[3])));
						}
					}
					if(border){
						oCell.SetBorders("Left", parent.g_asc_plugins.api.CreateColorFromRGB(0,0,0));
						oCell.SetBorders("Bottom", "Thin", parent.g_asc_plugins.api.CreateColorFromRGB(0,0,0));
						oCell.SetBorders("Right", "Thin", parent.g_asc_plugins.api.CreateColorFromRGB(0,0,0));
						oCell.SetBorders("Top", "Thin", parent.g_asc_plugins.api.CreateColorFromRGB(0,0,0));
					}
					currentcol++;
					break;
			}
		}
		parent.Asc.editor.controller.view.resize();
	}

})(window, undefined);
