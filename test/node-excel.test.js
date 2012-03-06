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

	describe('#getSheetCount()', function() {
		it('return the number of worksheets in the current workbook');
	});

	describe('#getSheetByName()', function() {
		it('return a worksheet object with a valid name');
		it('return null for a non existent name');
	});

});