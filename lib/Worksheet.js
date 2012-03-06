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
		defaultRowDimension = null,
		columnDimensions = [],
		defaultColumnDimension = null,
		drawingCollection = null,
		chartCollection = null,
		workSheetTitle,
		sheetState,
		pageSetup,
		pageMargins,
		headerFooter,
		sheetView,
		protection,
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

		if (getParent().getSheetByName(title)) {
			var me = true;
			//TODO copy sheet naming logic to add a number to title.
		}

		workSheetTitle = title;
		return self;
	}

	function disconnectCells() {

	}

	self.disconnectCells = disconnectCells;
	self.getParent = getParent;
	self.setTitle = setTitle;
	self.getTitle = getTitle;
	self.getInvalidTitleCharacters = getInvalidTitleCharacters;

	return self;
}