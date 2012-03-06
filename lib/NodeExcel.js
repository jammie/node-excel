var
	worksheet = require('./Worksheet'),
	documentProperties = require('./DocumentProperties'),
	documentSecurity = require('./DocumentSecurity'),
	excelStyle = require('./Style');

module.exports.create = function () {
	var
		self = {},
		properties = documentProperties.create(),
		security = documentSecurity.create(),
		worksheetCollection = [],
		activeSheetIndex = 0,
		namedRanges = [],
		cellXfSupervisor = excelStyle.create(),
		cellXfCollection,
		cellStyleXfCollection;

	function disconnectWorksheets() {
		worksheetCollection.forEach(function(worksheet) {
			worksheet.disconnectCells();
		});
		worksheetCollection = [];
	}

	function getProperties() {
		return properties;
	}

	function setProperties(newProperties) {
		properties = newProperties;
	}

	function getSecurity() {
		return security;
	}

	function setSecurity(newSecurity) {
		security = newSecurity;
	}

	function getAllSheets() {
		return worksheetCollection;
	}

	function getSheetCount() {
		return worksheetCollection.length;
	}

	function getSheetByName(name) {
		if (name === undefined) {
			name = '';
		}

		var sheetCount = getSheetCount();
		for (var i = 0; i < sheetCount; i++) {
			if (worksheetCollection[i].getTitle() === name) {
				return worksheetCollection[i];
			}
		}

		return null;
	}

	self.disconnectWorksheets = disconnectWorksheets;
	self.getProperties = getProperties;
	self.setProperties = setProperties;
	self.getSecurity = getSecurity;
	self.setSecurity = setSecurity;
	self.getAllSheets = getAllSheets;
	self.getSheetByName = getSheetByName;
	self.getSheetCount = getSheetCount;

	worksheetCollection.push(worksheet.create(self));

	return self;
}