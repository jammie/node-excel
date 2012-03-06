var
	cachedObjectStorageFactory = require('./CachedObjectStorageFactory').create(),
	PageSetupModule = require('./Worksheet/PageSetup'),
	PageMarginsModule = require('./Worksheet/PageMargins'),
	HeaderFooterModule = require('./Worksheet/HeaderFooter'),
	SheetViewModule = require('./Worksheet/SheetView'),
	ProtectionModule = require('./Worksheet/Protection'),
	RowDimensionModule = require('./Worksheet/RowDimension'),
	ColumnDimensionModule = require('./Worksheet/ColumnDimension');

module.exports.create = function (excel, title) {

	var constants = {
		break: {
			none: 0,
			row: 1,
			column: 2
		},
		sheetState: {
			visible: 'visible',
			hidden: 'hidden',
			veryHidden: 'veryHidden'
		},
		invalidCharactersForTitle: {
			list: ['*', ':', '/', '\\', '?', '[', ']'],
			regExp: /[\*\:\/\\\?\[\]]/i
		}
	};

	var
		self = {},

		parent,
		cellCollection = null,
		rowDimensions = [],
		defaultRowDimension = RowDimensionModule.create(null),
		columnDimensions = [],
		defaultColumnDimension = ColumnDimensionModule.create(null),
		drawingCollection = null,
		chartCollection = null,
		workSheetTitle,
		sheetState,
		pageSetup = PageSetupModule.create(),
		pageMargins = PageMarginsModule.create(),
		headerFooter = HeaderFooterModule.create(),
		sheetView = SheetViewModule.create(),
		protection = ProtectionModule.create(),
		styles = [],
		conditionalStylesCollection = [],
		cellCollectionIsSorted = false,
		breaks = [],
		mergeCells = [],
		protectedCells = [],
		autoFilter = '',
		freezePane = '',
		showGridLines = true,
		printGridLines = true,
		showRowColHeaders = true,
		showSummaryBelow = true,
		showSummaryRight = true,
		comments = [],
		activeCell = 'A1',
		selectedCell = 'A1',
		cachedHighestColumn = 'A',
		cachedHighestRow = 1,
		rightToLeft = false,
		hyperLinkCollection = [],
		dataValidationCollection = [],
		tabColor,
		dirty = true,
		hash = null;

	function getParent() {
		return parent;
	}

	function getTitle() {
		return workSheetTitle;
	}

	function getInvalidTitleCharacters() {
		return constants.invalidCharactersForTitle.list;
	}

	function checkSheetTitle(title) {
		if (title.search(constants.invalidCharactersForTitle.regExp) !== -1) {
			throw new Error('Invalid character found in sheet title, must not include: ' + getInvalidTitleCharacters());
		}

		if (title.length > 31) {
			throw new Error('Maximum 31 characters allowed in sheet title.');
		}

		return title;
	}

	function setTitle(title) {
		if (getTitle() === title) {
			return self;
		}

		checkSheetTitle(title);

		var oldTitle = getTitle();

		if (getParent().getSheetByName(title)) {

			if (title.length > 29) {
				title = title.substring(0, 29);
			}

			var i = 1;
			while (getParent().getSheetByName(title + ' ' + i)) {
				i++;
				if ((i === 10) && (title.length > 28)) {
					title = title.substring(0, 28);
				} else if ((i === 100) && (title.length > 27)) {
					title = title.substring(0, 27);
				}
			}

			var altTitle = title + ' ' + i;
			return setTitle(altTitle);
		}

		workSheetTitle = title;
		return self;
	}

	function disconnectCells() {

	}

	function initialise() {
		if (excel === undefined) {
			parent = null;
		} else {
			parent = excel;
		}

		if (title === undefined) {
			setTitle('sheet');
		} else {
			setTitle(title);
		}

		drawingCollection = {};
		chartCollection = {};

		cellCollection = cachedObjectStorageFactory.getInstance(self);
	}

	self.disconnectCells = disconnectCells;
	self.getParent = getParent;
	self.setTitle = setTitle;
	self.getTitle = getTitle;
	self.getInvalidTitleCharacters = getInvalidTitleCharacters;

	initialise();

	return self;
}