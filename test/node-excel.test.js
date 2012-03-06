var nodeExcel = require('../lib/NodeExcel');

describe('Node Excel', function() {

	describe('#create()', function() {
		it('creates a new NodeExcel object', function() {
			var excel = nodeExcel.create();
			excel.should.be.a('object');
		});
	});

	describe('#disconnectWorksheets()', function() {
		it('should disconnect all worksheets and set worksheets to a blank array', function() {
			var excel = nodeExcel.create();
			// Need to add a sheet before we can successfully test disconnection
			excel.disconnectWorksheets();
			excel.getAllSheets().length.should.equal(0);
		});
	});

	describe('#getAllSheets()', function() {
		it('returns all worksheets as an array');
	});

	describe('#getSheetCount()', function() {
		it('return the number of worksheets in the current workbook');
	});	

	describe('#getSheet()', function() {
		it('return the sheet at the specified index');
		it('throw an error if specified index does not exist');
	});

	describe('#getActiveSheet()', function() {
		it('return the currently active worksheet');
	});

	describe('#setActiveSheetIndex()', function() {
		it('sets the currently active sheet to the given index and returns new active sheet');
	});

	describe('#setActiveSheetIndexByName()', function() {
		it('sets the currently active sheet by sheet name and returns new active sheet');
	});

	describe('#getSheetByName()', function() {
		it('return a worksheet object with a valid name');
		it('return null for a non existent name');
	});

	describe('#createSheet()', function() {
		it('returns a new sheet object');
	});

	describe('#addSheet()', function() {
		it('adds a new sheet object at the specified index');
	});

	describe('#removeSheetByIndex()', function() {
		it('removes the sheet at the given index');
		it('throw an error if specified index does not exist');
	});

	describe('#getIndex()', function() {
		it('returns the index for a given sheet');
	});

	describe('#setIndexByName()', function() {
		it('set index for sheet by sheet name and return newIndex');
	});

	describe('#getSheetNames()', function() {
		it('return an array of sheet names');
	});


});