var
	nodeExcel = require('../lib/NodeExcel'),
	worksheet = require('../lib/Worksheet');

describe('Worksheet', function() {

	describe('#create()', function() {

		it('returns an empty worksheet', function() {
			var excel = nodeExcel.create();
			var newWorkSheet = worksheet.create(excel);
			newWorkSheet.should.be.a('object');
		});
	});

	describe('#setTitle()', function() {
		var excel;
		var newWorkSheet;

		beforeEach(function() {
			excel = nodeExcel.create();
			newWorkSheet = worksheet.create(excel);
		});

		it('returns worksheet object when same name passed in', function() {
			newWorkSheet = newWorkSheet.setTitle('sheet');
			newWorkSheet.should.be.a('object');
		});

		it('returns worksheet object with new title when a new title is set', function() {
			newWorkSheet = newWorkSheet.setTitle('my sheet');
			var title = newWorkSheet.getTitle();
			title.should.be.equal('my sheet');
		});

		it('returns a worksheet with a valid name of a sheet that already exists with a number added.');
		it('returns a worksheet with a valid name(31 characters) of a sheet that already exists truncated');

		it('throws an error when a invalid character is used for the title', function() {
			try {
				newWorkSheet = newWorkSheet.setTitle('my sheet *');
				false.should.be.ok;
			} catch (error) {
				true.should.be.ok;
			}
		});

		it('throws an error when a title of more than 31 characters is used', function() {
			try {
				newWorkSheet = newWorkSheet.setTitle('my sheet title is longer than 31 characters so will throw an error');
				false.should.be.ok;
			} catch (error) {
				true.should.be.ok;
			}
		});
	});

	describe('#getParent()', function() {
		it('return a nodeExcel object');
	});

	describe('#getInvalidTitleCharacters()', function() {
		it('todo');
	});

	describe('#getTitle()', function() {
		it('todo');
	});

	describe('#disconnectCells()', function() {
		it('todo');
	});

});