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

	function getActiveSheet() {
		return worksheetCollection[activeSheetIndex];
	}

	function setActiveSheetIndex(index) {
		if (index > worksheetCollection.length - 1) {
			throw new Error('Sheet index is out of bounds');
		} else {
			activeSheetIndex = index;
		}
		return getActiveSheet();
	}

	function getSheet(index) {
		if (index > worksheetCollection.length - 1) {
			throw new Error('Sheet index is out of bounds');
		} else {
			return worksheetCollection[index];
		}
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

	function setActiveSheetIndexByName(name) {
		var worksheet = getSheetByName(name);
		if (worksheet) {
			setActiveSheetIndex(worksheet.getParent().getIndex(worksheet));
			return worksheet;
		}

		throw new Error('Workbook does not contain sheet:' + name);
	}

	function addSheet(sheet, sheetIndex) {
		if (sheetIndex) {
			worksheetCollection.splice(sheetIndex, 0, sheet);
			if (activeSheetIndex >= sheetIndex) {
				activeSheetIndex++;
			}
		} else {
			worksheetCollection.push(sheet);
		}
		return sheet;
	}

	function createSheet(sheetIndex) {
		var newSheet = worksheet.create(self);
		addSheet(newSheet, sheetIndex);
		return newSheet;
	}

	function removeSheetByIndex(index) {
		if (index > worksheetCollection.length - 1) {
			throw new Error('Sheet index is out of bounds');
		} else {
			worksheetCollection.splice(index, 1);
		}
	}

	function getIndex(sheet) {
		for (var i in worksheetCollection) {
			if (worksheetCollection[i].getHashCode() === sheet.getHashCode()) {
				return i;
			}
		}
	}

	function setIndexByName(sheetName, newIndex) {
		var 
			oldIndex = getIndex(getSheetByName(sheetName)),
			sheet = worksheetCollection.splice(oldIndex, 1);

		worksheetCollection.splice(newIndex, 0, sheet);
		return newIndex;
	}

	function getSheetNames() {
		var 
			sheetNames = [],
			worksheetCount = getSheetCount();
			
		for (var i = 0; i < worksheetCount; i++) {
			sheetNames.push(getSheet(i).getTitle());
		}
		return sheetNames;
	}

	self.disconnectWorksheets = disconnectWorksheets;
	self.getProperties = getProperties;
	self.setProperties = setProperties;
	self.getSecurity = getSecurity;
	self.setSecurity = setSecurity;
	self.getAllSheets = getAllSheets;
	self.getSheetByName = getSheetByName;
	self.getSheetCount = getSheetCount;
	self.addSheet = addSheet;
	self.getSheet = getSheet;
	self.createSheet = createSheet;
	self.getActiveSheet = getActiveSheet;
	self.getIndex = getIndex;
	self.setIndexByName = setIndexByName;
	self.setActiveSheetIndex = setActiveSheetIndex;
	self.setActiveSheetIndexByName = setActiveSheetIndexByName;
	self.getSheetNames = getSheetNames;

	worksheetCollection.push(worksheet.create(self));

	return self;
}