module.exports.create = function (worksheet) {
	var
		self = {},

		parent = worksheet,
		currentObject = null,
		currentObjectId = null,
		currentCellIsDirty = true,
		cellCache = {};

	function isDataSet(coord) {
		if (coord === currentObjectId) {
			return true;
		}

		return (cellCache[coord] !== undefined);
	}

	function addCacheData(coord, cell) {
		cellCache[coord] = cell;
		return cell;
	}

	function updateCacheData(cell) {
		return addCacheData(cell.getCoordinate(), cell);
	}

	function deleteCacheData(coord) {
		if (coord === currentObjectId) {
			currentObject.detach();
			currentObjectId = null;
			currentObject = null;
		}

		if (typeof(cellCache[coord]) === 'object') {
			cellCache[coord].detach();

			//Not sure if this is the correct way to unset the object property.
			cellCache[coord] = null;
		}

		currentCellIsDirty = false;
	}

	function getCellList() {
		var cellAddresses = [];
		for (var property in cellCache) {
			cellAddresses.push(property);
		}
		return cellAddresses;
	}

	function unsetWorksheetCells() {

	}

	self.unsetWorksheetCells = unsetWorksheetCells;
	self.isDataSet = isDataSet;
	self.updateCacheData = updateCacheData;
	self.deleteCacheData = deleteCacheData;
	self.getCellList = getCellList;

	return self;
}