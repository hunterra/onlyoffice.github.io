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

var Sheet = function(data){
	var self = this; 
	self.name = data ? data.name : "";
	self.index = data ? data.index : undefined;
	self.celladdress = ko.observable(data ? (data.celladdress ? data.celladdress : "A1") : "A1");

	self.getcelladdress = function(){
		window.Asc.plugin.getCellAddress();
	}
}


var importModel = function(){
	var self = this;
	
	self.worksheets = ko.observableArray(window.Asc.plugin.getSheets());
	self.newworksheet = ko.observable(true);
	self.worksheet = ko.observable(self.worksheets()[window.Asc.plugin.getActiveSheetIndex()]);
	self.filename = ko.observable();
	self.filepath = ko.observable();
	self.filepath.subscribe(function(path){
		if(self.filepath() && self.filepath().lastIndexOf("\\") != -1)
			self.filename(self.filepath().substring(self.filepath().lastIndexOf("\\")+1))
		else if(self.filepath() && self.filepath().lastIndexOf("/") != -1)
			self.filename(self.filepath().substring(self.filepath().lastIndexOf("/")+1))
		else
			self.filename(undefined);
	});
	self.data = ko.observable();
	
	self.worksheet.subscribe(function(sheet){
		window.Asc.plugin.setSheet(sheet);
	});
	self.data.subscribe(function(data){
		if(!data)
			self.filepath(undefined)
	});
}
