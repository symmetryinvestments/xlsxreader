// TODO: add const access to getter functions, such as `getRow`
module xlsxreader;

import std.algorithm.iteration : filter, map, joiner;
import std.algorithm.mutation : reverse;
import std.algorithm.searching : all, canFind, startsWith;
import std.algorithm.sorting : sort;
import std.array : array, empty, front, popFront;
import std.ascii : isDigit;
import std.conv : to;
import std.datetime : DateTime, Date, TimeOfDay;
import std.exception : enforce;
import std.file : read, exists, readText;
import std.format : format;
import std.range : tee;
import std.regex;
import std.stdio;
import std.traits : isIntegral, isFloatingPoint, isSomeString;
import std.typecons : tuple, Nullable, nullable;
import std.utf : byChar;
import std.variant : Algebraic, visit;
import std.zip;

import dxml.dom;

@safe:

///
struct Pos {
	// zero based
	size_t row;
	// zero based
	size_t col;
}

///
alias Data = Algebraic!(bool,long,double,string,DateTime,Date,TimeOfDay);

///
struct Cell {
	string loc;
	size_t row; // row[r]
	string t; // s or n, s for pointer, n for value, stored in v
	string r; // c[r]
	string v; // c.v the value or ptr
	string f; // c.f the formula
	Data value;
	Pos position;

	bool canConvertTo(CellType ct) const @trusted {
		auto b = (bool l) {
			switch(ct) {
				case CellType.datetime: return false;
				case CellType.timeofday: return false;
				case CellType.date: return false;
				default: return true;
			}
		};

		auto dt = (DateTime l) {
			switch(ct) {
				case CellType.datetime: return true;
				case CellType.timeofday: return true;
				case CellType.date: return true;
				case CellType.double_: return true;
				default: return false;
			}
		};

		auto de = (Date l) {
			switch(ct) {
				case CellType.datetime: return false;
				case CellType.timeofday: return false;
				case CellType.date: return true;
				case CellType.double_: return true;
				case CellType.long_: return true;
				default: return false;
			}
		};

		auto tod = (TimeOfDay l) {
			switch(ct) {
				case CellType.datetime: return false;
				case CellType.timeofday: return true;
				case CellType.double_: return true;
				default: return false;
			}
		};

		auto l = (long l) {
			switch(ct) {
				case CellType.date: return !tryConvertTo!Date(l);
				case CellType.long_: return true;
				case CellType.string_: return true;
				case CellType.double_: return true;
				case CellType.bool_: return l == 0 || l == 1;
				default: return false;
			}
		};

		auto d = (double l) {
			switch(ct) {
				case CellType.string_: return true;
				case CellType.double_: return true;
				case CellType.datetime: return !tryConvertTo!DateTime(l);
				case CellType.date: return !tryConvertTo!Date(l);
				case CellType.timeofday: return !tryConvertTo!TimeOfDay(l);
				default: return false;
			}
		};

		auto s = (string l) {
			switch(ct) {
				case CellType.string_: return true;
				case CellType.bool_: return !tryConvertTo!bool(l);
				case CellType.long_: return !tryConvertTo!long(l);
				case CellType.double_: return !tryConvertTo!double(l);
				case CellType.datetime: return !tryConvertTo!DateTime(l);
				case CellType.date: return !tryConvertTo!Date(l);
				case CellType.timeofday: return !tryConvertTo!TimeOfDay(l);
				default: return false;
			}
		};

		return this.value.visit!(
				(bool l) => b(l),
				(long lo) => l(lo),
				(double l) => d(l),
				(string l) => s(l),
				(DateTime l) => dt(l),
				(Date l) => de(l),
				(TimeOfDay l) => tod(l),
				() => false)
			();
	}

	bool convertToBool() @trusted const {
		return convertTo!bool(this.value);
	}

	long convertToLong() @trusted const {
		return convertTo!long(this.value);
	}

	double convertToDouble() @trusted const {
		return convertTo!double(this.value);
	}

	string convertToString() @trusted const {
		return convertTo!string(this.value);
	}

	Date convertToDate() @trusted const {
		return convertTo!Date(this.value);
	}

	TimeOfDay convertToTimeOfDay() @trusted const {
		return convertTo!TimeOfDay(this.value);
	}

	DateTime convertToDateTime() @trusted const {
		return convertTo!DateTime(this.value);
	}
}

//
enum CellType {
	datetime,
	timeofday,
	date,
	bool_,
	double_,
	long_,
	string_
}

import std.ascii : toUpper;
///
struct Sheet {
	Cell[] cells;
	Cell[][] table;
	Pos maxPos;

	string toString() const @trusted {
		import std.format : formattedWrite;
		import std.array : appender;
		long[] maxCol = new long[](maxPos.col + 1);
		foreach(row; this.table) {
			foreach(idx, Cell col; row) {
				string s = col.value.visit!(
						(bool l) => to!string(l),
						(long l) => to!string(l),
						(double l) => format("%.4f", l),
						(string l) => l,
						(DateTime l) => l.toISOExtString(),
						(Date l) => l.toISOExtString(),
						(TimeOfDay l) => l.toISOExtString(),
						() => "")
					();

				maxCol[idx] = maxCol[idx] < s.length ? s.length : maxCol[idx];
			}
		}
		maxCol[] += 1;

		auto app = appender!string();
		foreach(row; this.table) {
			foreach(idx, Cell col; row) {
				string s = col.value.visit!(
						(bool l) => to!string(l),
						(long l) => to!string(l),
						(double l) => format("%.4f", l),
						(string l) => l,
						(DateTime l) => l.toISOExtString(),
						(Date l) => l.toISOExtString(),
						(TimeOfDay l) => l.toISOExtString(),
						() => "")
					();
				formattedWrite(app, "%*s, ", maxCol[idx], s);
			}
			formattedWrite(app, "\n");
		}
		return app.data;
	}

	void printTable() const {
		writeln(this.toString());
	}

	// Column

	Iterator!T getColumn(T)(size_t col, size_t startColumn = 0, size_t endColumn = size_t.max) {
		auto c = this.iterateColumn!T(col, startColumn, endColumn);
		return Iterator!T(c.array);
	}

	private enum t = q{
	Iterator!(%1$s) getColumn%2$s(size_t col, size_t startColumn = 0, size_t endColumn = size_t.max) {
		return getColumn!(%1$s)(col, startColumn, endColumn);
	}
	};
	static foreach(T; ["long", "double", "string", "Date", "TimeOfDay",
			"DateTime"])
	{
		mixin(format(t, T, T[0].toUpper ~ T[1 .. $]));
	}

	ColumnUntyped iterateColumnUntyped(size_t col, size_t startColumn = 0, size_t endColumn = size_t.max) {
		return ColumnUntyped(&this, col, startColumn, endColumn);
	}

	Column!(T) iterateColumn(T)(size_t col, size_t startColumn = 0, size_t endColumn = size_t.max) {
		return Column!(T)(&this, col, startColumn, endColumn);
	}

	Column!(long) iterateColumnLong(size_t col, size_t startColumn = 0, size_t endColumn = size_t.max) {
		return Column!(long)(&this, col, startColumn, endColumn);
	}

	Column!(double) iterateColumnDouble(size_t col, size_t startColumn = 0, size_t endColumn = size_t.max) {
		return Column!(double)(&this, col, startColumn, endColumn);
	}

	Column!(string) iterateColumnString(size_t col, size_t startColumn = 0, size_t endColumn = size_t.max) {
		return Column!(string)(&this, col, startColumn, endColumn);
	}

	Column!(DateTime) iterateColumnDateTime(size_t col, size_t startColumn = 0, size_t endColumn = size_t.max)
	{
		return Column!(DateTime)(&this, col, startColumn, endColumn);
	}

	Column!(Date) iterateColumnDate(size_t col, size_t startColumn = 0, size_t endColumn = size_t.max) {
		return Column!(Date)(&this, col, startColumn, endColumn);
	}

	Column!(TimeOfDay) iterateColumnTimeOfDay(size_t col, size_t startColumn = 0, size_t endColumn = size_t.max)
	{
		return Column!(TimeOfDay)(&this, col, startColumn, endColumn);
	}

	// Row

	Iterator!T getRow(T)(size_t row, size_t startColumn = 0, size_t endColumn = size_t.max) {
		auto c = this.iterateRow!T(row, startColumn, endColumn);
		return Iterator!T(c.array); // TODO: why .array?
	}

	private enum t2 = q{
	Iterator!(%1$s) getRow%2$s(size_t row, size_t startColumn = 0, size_t endColumn = size_t.max) {
		return getRow!(%1$s)(row, startColumn, endColumn);
	}
	};
	static foreach(T; ["long", "double", "string", "Date", "TimeOfDay",
			"DateTime"])
	{
		mixin(format(t2, T, T[0].toUpper ~ T[1 .. $]));
	}

	RowUntyped iterateRowUntyped(size_t row, size_t startColumn = 0, size_t endColumn = size_t.max) {
		return RowUntyped(&this, row, startColumn, endColumn);
	}

	Row!(T) iterateRow(T)(size_t row, size_t startColumn = 0, size_t endColumn = size_t.max) {
		return Row!(T)(&this, row, startColumn, endColumn);
	}

	Row!(long) iterateRowLong(size_t row, size_t startColumn = 0, size_t endColumn = size_t.max) {
		return Row!(long)(&this, row, startColumn, endColumn);
	}

	Row!(double) iterateRowDouble(size_t row, size_t startColumn = 0, size_t endColumn = size_t.max) {
		return Row!(double)(&this, row, startColumn, endColumn);
	}

	Row!(string) iterateRowString(size_t row, size_t startColumn = 0, size_t endColumn = size_t.max) {
		return Row!(string)(&this, row, startColumn, endColumn);
	}

	Row!(DateTime) iterateRowDateTime(size_t row, size_t startColumn = 0, size_t endColumn = size_t.max)
	{
		return Row!(DateTime)(&this, row, startColumn, endColumn);
	}

	Row!(Date) iterateRowDate(size_t row, size_t startColumn = 0, size_t endColumn = size_t.max) {
		return Row!(Date)(&this, row, startColumn, endColumn);
	}

	Row!(TimeOfDay) iterateRowTimeOfDay(size_t row, size_t startColumn = 0, size_t endColumn = size_t.max)
	{
		return Row!(TimeOfDay)(&this, row, startColumn, endColumn);
	}
}

struct Iterator(T) {
	T[] data;

	this(T[] data) {
		this.data = data;
	}

	@property bool empty() const pure nothrow @nogc {
		return this.data.empty;
	}

	void popFront() {
		this.data.popFront();
	}

	@property T front() {
		return this.data.front;
	}

	inout(typeof(this)) save() inout pure nothrow @nogc {
		return this;
	}

	// Request random access.
	inout(T)[] array() inout @safe pure nothrow @nogc {
		return data;
	}
}

///
struct RowUntyped {
	Sheet* sheet;
	const size_t row;
	const size_t startColumn;
	const size_t endColumn;
	size_t cur;

	this(Sheet* sheet, in size_t row, in size_t startColumn = 0, in size_t endColumn = size_t.max) pure nothrow @nogc {
		assert(sheet.table.length == sheet.maxPos.row + 1);
		this.sheet = sheet;
		this.row = row;
		this.startColumn = startColumn;
		this.endColumn = endColumn != size_t.max ? endColumn : sheet.maxPos.col + 1;
		this.cur = this.startColumn;
	}

	@property bool empty() const pure nothrow @nogc {
		return this.cur >= this.endColumn;
	}

	void popFront() pure nothrow @nogc {
		++this.cur;
	}

	inout(typeof(this)) save() inout pure nothrow @nogc {
		return this;
	}

	@property inout(Cell) front() inout pure nothrow @nogc {
		return this.sheet.table[this.row][this.cur];
	}
}

///
struct Row(T) {
	RowUntyped ru;
	T front;

	this(Sheet* sheet, size_t row, size_t startColumn = 0, size_t endColumn = size_t.max) {
		ru = RowUntyped(sheet, row, startColumn, endColumn);
		read();
	}

	@property bool empty() const pure nothrow @nogc {
		return this.ru.empty;
	}

	void popFront() {
		this.ru.popFront();
		if(!this.empty) {
			this.read();
		}
	}

	inout(typeof(this)) save() inout pure nothrow @nogc {
		return this;
	}

	private void read() {
		this.front = convertTo!T(this.ru.front.value);
	}

	bool canConvertTo(CellType ct) const {
		for(size_t it = this.ru.startColumn; it < this.ru.endColumn; ++it) {
			if(!this.ru.sheet.table[this.ru.row][it].canConvertTo(ct)) {
				return false;
			}
		}
		return true;
	}
}

///
struct ColumnUntyped {
	Sheet* sheet;
	const size_t col;
	const size_t startRow;
	const size_t endRow;
	size_t cur;

	this(Sheet* sheet, size_t col, size_t startRow = 0, size_t endRow = size_t.max) {
		assert(sheet.table.length == sheet.maxPos.row + 1);
		this.sheet = sheet;
		this.col = col;
		this.startRow = startRow;
		this.endRow = endRow != size_t.max ? endRow : sheet.maxPos.row + 1;
		this.cur = this.startRow;
	}

	@property bool empty() const pure nothrow @nogc {
		return this.cur >= this.endRow;
	}

	void popFront() {
		++this.cur;
	}

	inout(typeof(this)) save() inout pure nothrow @nogc {
		return this;
	}

	@property Cell front() {
		return this.sheet.table[this.cur][this.col];
	}
}

///
struct Column(T) {
	ColumnUntyped cu;

	T front;

	this(Sheet* sheet, size_t col, size_t startRow = 0, size_t endRow = size_t.max) {
		this.cu = ColumnUntyped(sheet, col, startRow, endRow);
		this.read();
	}

	@property bool empty() const pure nothrow @nogc {
		return this.cu.empty;
	}

	void popFront() {
		this.cu.popFront();
		if(!this.empty) {
			this.read();
		}
	}

	inout(typeof(this)) save() inout pure nothrow @nogc {
		return this;
	}

	private void read() {
		this.front = convertTo!T(this.cu.front.value);
	}

	bool canConvertTo(CellType ct) const {
		for(size_t it = this.cu.startRow; it < this.cu.endRow; ++it) {
			if(!this.cu.sheet.table[it][this.cu.col].canConvertTo(ct)) {
				return false;
			}
		}
		return true;
	}
}

unittest {
	import std.range : isForwardRange;
	import std.meta : AliasSeq;
	static foreach(T; AliasSeq!(long,double,DateTime,TimeOfDay,Date,string)) {{
		alias C = Column!T;
		alias R = Row!T;
		alias I = Iterator!T;
		static assert(isForwardRange!C, C.stringof);
		static assert(isForwardRange!R, R.stringof);
		static assert(isForwardRange!I, I.stringof);
	}}
}

Date longToDate(long d) {
	// modifed from https://www.codeproject.com/Articles/2750/
	// Excel-Serial-Date-to-Day-Month-Year-and-Vice-Versa

	// Excel/Lotus 123 have a bug with 29-02-1900. 1900 is not a
	// leap year, but Excel/Lotus 123 think it is...
	if(d == 60) {
		return Date(1900, 2,  29);
	} else if(d < 60) {
		// Because of the 29-02-1900 bug, any serial date
		// under 60 is one off... Compensate.
		++d;
	}

	// Modified Julian to DMY calculation with an addition of 2415019
	int l = cast(int)d + 68569 + 2415019;
	int n = int(( 4 * l ) / 146097);
	l = l - int(( 146097 * n + 3 ) / 4);
	int i = int(( 4000 * ( l + 1 ) ) / 1461001);
	l = l - int(( 1461 * i ) / 4) + 31;
	int j = int(( 80 * l ) / 2447);
	int nDay = l - int(( 2447 * j ) / 80);
	l = int(j / 11);
	int nMonth = j + 2 - ( 12 * l );
	int nYear = 100 * ( n - 49 ) + i + l;
	return Date(nYear, nMonth, nDay);
}

long dateToLong(Date d) {
	// modifed from https://www.codeproject.com/Articles/2750/
	// Excel-Serial-Date-to-Day-Month-Year-and-Vice-Versa

	// Excel/Lotus 123 have a bug with 29-02-1900. 1900 is not a
	// leap year, but Excel/Lotus 123 think it is...
	if(d.day == 29 && d.month == 2 && d.year == 1900) {
		return 60;
	}

	// DMY to Modified Julian calculated with an extra subtraction of 2415019.
	long nSerialDate =
			int(( 1461 * ( d.year + 4800 + int(( d.month - 14 ) / 12) ) ) / 4) +
			int(( 367 * ( d.month - 2 - 12 *
				( ( d.month - 14 ) / 12 ) ) ) / 12) -
				int(( 3 * ( int(( d.year + 4900
				+ int(( d.month - 14 ) / 12) ) / 100) ) ) / 4) +
				d.day - 2415019 - 32075;

	if(nSerialDate < 60) {
		// Because of the 29-02-1900 bug, any serial date
		// under 60 is one off... Compensate.
		nSerialDate--;
	}

	return nSerialDate;
}

unittest {
	auto ds = [ Date(1900,2,1), Date(1901, 2, 28), Date(2019, 06, 05) ];
	foreach(d; ds) {
		long l = dateToLong(d);
		Date r = longToDate(l);
		assert(r == d, format("%s %s", r, d));
	}
}

TimeOfDay doubleToTimeOfDay(double s) {
	import core.stdc.math : lround;
	double secs = (24.0 * 60.0 * 60.0) * s;

	// TODO not one-hundred my lround is needed
	int secI = to!int(lround(secs));

	return TimeOfDay(secI / 3600, (secI / 60) % 60, secI % 60);
}

double timeOfDayToDouble(TimeOfDay tod) {
	long h = tod.hour * 60 * 60;
	long m = tod.minute * 60;
	long s = tod.second;
	return (h + m + s) / (24.0 * 60.0 * 60.0);
}

unittest {
	auto tods = [ TimeOfDay(23, 12, 11), TimeOfDay(11, 0, 11),
		 TimeOfDay(0, 0, 0), TimeOfDay(0, 1, 0),
		 TimeOfDay(23, 59, 59), TimeOfDay(0, 0, 0)];
	foreach(tod; tods) {
		double d = timeOfDayToDouble(tod);
		assert(d <= 1.0, format("%s", d));
		TimeOfDay r = doubleToTimeOfDay(d);
		assert(r == tod, format("%s %s", r, tod));
	}
}

double datetimeToDouble(DateTime dt) {
	double d = dateToLong(dt.date);
	double t = timeOfDayToDouble(dt.timeOfDay);
	return d + t;
}

DateTime doubleToDateTime(double d) {
	long l = cast(long)d;
	Date dt = longToDate(l);
	TimeOfDay t = doubleToTimeOfDay(d - l);
	return DateTime(dt, t);
}

unittest {
	auto ds = [ Date(1900,2,1), Date(1901, 2, 28), Date(2019, 06, 05) ];
	auto tods = [ TimeOfDay(23, 12, 11), TimeOfDay(11, 0, 11),
		 TimeOfDay(0, 0, 0), TimeOfDay(0, 1, 0),
		 TimeOfDay(23, 59, 59), TimeOfDay(0, 0, 0)];
	foreach(d; ds) {
		foreach(tod; tods) {
			DateTime dt = DateTime(d, tod);
			double dou = datetimeToDouble(dt);

			Date rd = longToDate(cast(long)dou);
			assert(rd == d, format("%s %s", rd, d));

			double rest = dou - cast(long)dou;
			TimeOfDay rt = doubleToTimeOfDay(dou - cast(long)dou);
			assert(rt == tod, format("%s %s", rt, tod));

			DateTime r = doubleToDateTime(dou);
			assert(r == dt, format("%s %s", r, dt));
		}
	}
}

Date stringToDate(string s) {
	import std.array : split;
	import std.string : indexOf;

	if(s.indexOf('/') != -1) {
		auto sp = s.split('/');
		enforce(sp.length == 3, format("[%s]", sp));
		return Date(to!int(sp[2]), to!int(sp[1]), to!int(sp[0]));
	} else {
		return longToDate(to!long(s));
	}
}

bool tryConvertTo(T,S)(S var) @trusted {
	return !(tryConvertToImpl!T(Data(var)).isNull());
}

Nullable!(T) tryConvertToImpl(T)(Data var) {
	try {
		return nullable(convertTo!T(var));
	} catch(Exception e) {
		return Nullable!T();
	}
}

T convertTo(T)(in Data var) @trusted {
	static if(is(T == Data)) {
		return var;
	} else static if(isSomeString!T) {
		return var.visit!(
			(bool l) => to!string(l),
			(long l) => to!string(l),
			(double l) => format("%f", l),
			(string l) => l,
			(DateTime l) => l.toISOExtString(),
			(Date l) => l.toISOExtString(),
			(TimeOfDay l) => l.toISOExtString(),
			() => "")
		();
	} else static if(is(T == bool)) {
		if(var.type != typeid(bool) && var.type != typeid(long)
		   && var.type == typeid(string))
		{
			throw new Exception("Can not convert " ~ var.type.toString() ~
								" to bool");
		}
		return var.visit!(
			(bool l) => l,
			(long l) => l == 0 ? false : true,
			(string l) => to!bool(l),
			(double l) => false,
			(DateTime l) => false,
			(Date l) => false,
			(TimeOfDay l) => false,
			() => false)
		();
	} else static if(isIntegral!T) {
		return var.visit!(
			(bool l) => to!T(l),
			(long l) => to!T(l),
			(double l) => to!T(l),
			(string l) => to!T(l),
			(DateTime l) => to!T(dateToLong(l.date)),
			(Date l) => to!T(dateToLong(l)),
			(TimeOfDay l) => to!T(0),
			() => to!T(0))
		();
	} else static if(isFloatingPoint!T) {
		return var.visit!(
			(bool l) => to!T(l),
			(long l) => to!T(l),
			(double l) => to!T(l),
			(string l) => to!T(l),
			(DateTime l) => to!T(dateToLong(l.date)),
			(Date l) => to!T(dateToLong(l)),
			(TimeOfDay l) => to!T(0),
			() => T.init)
		();
	} else static if(is(T == DateTime)) {
		return var.visit!(
			(bool l) => DateTime.init,
			(long l) => doubleToDateTime(to!long(l)),
			(double l) => doubleToDateTime(l),
			(string l) => doubleToDateTime(to!double(l)),
			(DateTime l) => l,
			(Date l) => DateTime(l, TimeOfDay.init),
			(TimeOfDay l) => DateTime(Date.init, l),
			() => DateTime.init)
		();
	} else static if(is(T == Date)) {
		import std.math : lround;

		return var.visit!(
			(bool l) => Date.init,
			(long l) => longToDate(l),
			(double l) => longToDate(lround(l)),
			(string l) => stringToDate(l),
			(DateTime l) => l.date,
			(Date l) => l,
			(TimeOfDay l) => Date.init,
			() => Date.init)
		();
	} else static if(is(T == TimeOfDay)) {
		import std.math : lround;

		return var.visit!(
			(bool l) => TimeOfDay.init,
			(long l) => TimeOfDay.init,
			(double l) => doubleToTimeOfDay(l - cast(long)l),
			(string l) => doubleToTimeOfDay(
				to!double(l) - cast(long)to!double(l)
			),
			(DateTime l) => l.timeOfDay,
			(Date l) => TimeOfDay.init,
			(TimeOfDay l) => l,
			() => TimeOfDay.init)
		();
	}
	assert(false, T.stringof);
}

private ZipArchive readFile(in string filename) @trusted {
	enforce(exists(filename), "File with name " ~ filename ~ " does not exist");
	return new typeof(return)(read(filename));
}

struct SheetNameId {
	string name;
	int id;
	string rid;
}

string convertToString(in ubyte[] d) @trusted {
	import std.encoding;
	auto b = getBOM(d);
	switch(b.schema) {
		case BOM.none:
			return cast(string)d;
		case BOM.utf8:
			return cast(string)(d[3 .. $]);
		case BOM.utf16be: goto default;
		case BOM.utf16le: goto default;
		case BOM.utf32be: goto default;
		case BOM.utf32le: goto default;
		default:
			string ret;
			transcode(d, ret);
			return ret;
	}
}

SheetNameId[] sheetNames(in string filename) @trusted {
	auto file = readFile(filename);
	auto ams = file.directory;
	immutable wbStr = "xl/workbook.xml";
	if(wbStr !in ams) {
		return SheetNameId[].init;
	}
	ubyte[] wb = file.expand(ams[wbStr]);
	string wbData = convertToString(wb);

	auto dom = parseDOM(wbData);
	assert(dom.children.length == 1);
	auto workbook = dom.children[0];
	string sheetName = workbook.name == "workbook"
		? "sheets" : "s:sheets";
	assert(workbook.name == "workbook" || workbook.name == "s:workbook");
	auto sheetsRng = workbook.children.filter!(c => c.name == sheetName);
	assert(!sheetsRng.empty);
	return sheetsRng.front.children
		.map!(s => SheetNameId(
					s.attributes.filter!(a => a.name == "name").front.value,
					s.attributes.filter!(a => a.name == "sheetId").front
						.value.to!int(),
					s.attributes.filter!(a => a.name == "r:id").front.value,
				)
		)
		.array
		.sort!((a, b) => a.id < b.id)
		.release;
}

unittest {
	auto r = sheetNames("multitable.xlsx");
	assert(r[0].name == "wb1");
	assert(r[0].id == 1);
}

struct Relationships {
	string id;
	string file;
}

Relationships[string] parseRelationships(ZipArchive za, ArchiveMember am) @trusted {
	ubyte[] d = za.expand(am);
	string relData = convertToString(d);
	auto dom = parseDOM(relData);
	assert(dom.children.length == 1);
	auto rel = dom.children[0];
	assert(rel.name == "Relationships");
	auto relRng = rel.children.filter!(c => c.name == "Relationship");
	assert(!relRng.empty);

	Relationships[string] ret;
	foreach(r; relRng) {
		Relationships tmp;
		tmp.id = r.attributes.filter!(a => a.name == "Id")
			.front.value;
		tmp.file = r.attributes.filter!(a => a.name == "Target")
			.front.value;
		ret[tmp.id] = tmp;
	}
	return ret;
}

Sheet readSheet(in string filename, in string sheetName) {
	SheetNameId[] sheets = sheetNames(filename);
	auto sRng = sheets.filter!(s => s.name == sheetName);
	enforce(!sRng.empty, "No sheet with name " ~ sheetName
			~ " found in file " ~ filename);
	return readSheetImpl(filename, sRng.front.rid);
}

string eatXlPrefix(string fn) {
	foreach(p; ["xl//", "/xl/"]) {
		if(fn.startsWith(p)) {
			return fn[p.length .. $];
		}
	}
	return fn;
}

Sheet readSheetImpl(in string filename, in string rid) @trusted {
	scope(failure) {
		writefln("Failed at file '%s' and sheet '%s'", filename, rid);
	}
	auto file = readFile(filename);
	auto ams = file.directory;
	immutable ss = "xl/sharedStrings.xml";
	Data[] sharedStrings = (ss in ams)
		? readSharedEntries(file, ams[ss])
		: [];
	//logf("%s", sharedStrings);

	Relationships[string] rels = parseRelationships(file,
			ams["xl/_rels/workbook.xml.rels"]);

	Relationships* sheetRel = rid in rels;
	enforce(sheetRel !is null, format("Could not find '%s' in '%s'", rid,
				filename));
	string shrFn = eatXlPrefix(sheetRel.file);
	string fn = "xl/" ~ shrFn;
	ArchiveMember* sheet = fn in ams;
	enforce(sheet !is null, format("sheetRel.file orig '%s', fn %s not in [%s]",
				sheetRel.file, fn, ams.keys()));

	Sheet ret;
	ret.cells = insertValueIntoCell(readCells(file, *sheet), sharedStrings);
	Pos maxPos;
	foreach(ref c; ret.cells) {
		c.position = toPos(c.r);
		maxPos = elementMax(maxPos, c.position);
	}
	ret.maxPos = maxPos;
	ret.table = new Cell[][](ret.maxPos.row + 1, ret.maxPos.col + 1);
	foreach(c; ret.cells) {
		ret.table[c.position.row][c.position.col] = c;
	}
	// debug writeln("xlsxreader: sheet members: cells.length:", ret.cells.length, " table.length:", ret.table.length, " table[0].length:", ret.table[0].length, " ret.maxPos:", ret.maxPos);
	return ret;
}

Data[] readSharedEntries(ZipArchive za, ArchiveMember am) @trusted {
	ubyte[] ss = za.expand(am);
	string ssData = convertToString(ss);
	auto dom = parseDOM(ssData);
	Data[] ret;
	if(dom.type != EntityType.elementStart) {
		return ret;
	}
	assert(dom.children.length == 1);
	auto sst = dom.children[0];
	assert(sst.name == "sst");
	if(sst.type != EntityType.elementStart || sst.children.empty) {
		return ret;
	}
	auto siRng = sst.children.filter!(c => c.name == "si");
	foreach(si; siRng) {
		if(si.type != EntityType.elementStart) {
			continue;
		}
		//ret ~= extractData(si);
		string tmp;
		foreach(tORr; si.children) {
			if(tORr.name == "t" && tORr.type == EntityType.elementStart
					&& !tORr.children.empty)
			{
				ret ~= Data(convert(tORr.children[0].text));
			} else if(tORr.name == "r") {
				foreach(r; tORr.children.filter!(r => r.name == "t")) {
					if(r.type == EntityType.elementStart && !r.children.empty) {
						tmp ~= r.children[0].text;
					}
				}
			} else {
				ret ~= Data.init;
			}
		}
		if(!tmp.empty) {
			ret ~= Data(convert(tmp));
		}
	}
	return ret;
}

string extractData(DOMEntity!string si) {
	string tmp;
	foreach(tORr; si.children) {
		if(tORr.name == "t") {
			if(!tORr.attributes.filter!(a => a.name == "xml:space").empty) {
				return "";
			} else if(tORr.type == EntityType.elementStart
					&& !tORr.children.empty)
			{
				return tORr.children[0].text;
			} else {
				return "";
			}
		} else if(tORr.name == "r") {
			foreach(r; tORr.children.filter!(r => r.name == "t")) {
				tmp ~= r.children[0].text;
			}
		}
	}
	if(!tmp.empty) {
		return tmp;
	}
	assert(false);
}

private bool canConvertToLong(in string s) {
	if(s.empty) {
		return false;
	}
	return s.byChar.all!isDigit();
}

private immutable rs = r"[\+-]{0,1}[0-9][0-9]*\.[0-9]*";
private auto rgx = ctRegex!rs;

private bool canConvertToDouble(in string s) {
	auto cap = matchAll(s, rgx);
	return cap.empty || cap.front.hit != s ? false : true;
}

Data convert(in string s) @trusted {
	struct ToRe {
		string from;
		string to;
	}

	immutable ToRe[] toRe = [
		ToRe( "&amp;", "&"),
		ToRe( "&gt;", "<"),
		ToRe( "&lt;", ">"),
		ToRe( "&quot;", "\""),
		ToRe( "&apos;", "'")
	];

	string replaceStrings(string s) {
		import std.algorithm.searching : canFind;
		import std.array : replace;
		foreach(tr; toRe) {
			while(canFind(s, tr.from)) {
				s = s.replace(tr.from, tr.to);
			}
		}
		return s;
	}

	if(canConvertToLong(s)) {
		return Data(to!long(s));
	} else if(canConvertToDouble(s)) {
		return Data(to!double(s));
	} else {
		return Data(replaceStrings(s));
	}
}

Cell[] readCells(ZipArchive za, ArchiveMember am) @trusted {
	Cell[] ret;
	ubyte[] ss = za.expand(am);
	string ssData = convertToString(ss);
	auto dom = parseDOM(ssData);
	assert(dom.children.length == 1);
	auto ws = dom.children[0];
	if(ws.name != "worksheet") {
		return ret;
	}
	auto sdRng = ws.children.filter!(c => c.name == "sheetData");
	assert(!sdRng.empty);
	if(sdRng.front.type != EntityType.elementStart) {
		return ret;
	}
	auto rows = sdRng.front.children
		.filter!(r => r.name == "row");

	foreach(row; rows) {
		if(row.type != EntityType.elementStart || row.children.empty) {
			continue;
		}
		foreach(c; row.children.filter!(r => r.name == "c")) {
			Cell tmp;
			tmp.row = row.attributes.filter!(a => a.name == "r")
				.front.value.to!size_t();
			tmp.r = c.attributes.filter!(a => a.name == "r")
				.front.value;
			auto t = c.attributes.filter!(a => a.name == "t");
			if(t.empty) {
				// we assume that no t attribute means direct number
				//writefln("Found a strange empty cell \n%s", c);
			} else {
				tmp.t = t.front.value;
			}
			if(tmp.t == "s" || tmp.t == "n") {
				if(c.type == EntityType.elementStart) {
					auto v = c.children.filter!(c => c.name == "v");
					//enforce(!v.empty, format("r %s", tmp.row));
					if(!v.empty && v.front.type == EntityType.elementStart
							&& !v.front.children.empty)
					{
						tmp.v = v.front.children[0].text;
					} else {
						tmp.v = "";
					}
				}
			} else if(tmp.t == "inlineStr") {
				auto is_ = c.children.filter!(c => c.name == "is");
				tmp.v = extractData(is_.front);
			} else if(c.type == EntityType.elementStart) {
				auto v = c.children.filter!(c => c.name == "v");
				if(!v.empty && v.front.type == EntityType.elementStart
						&& !v.front.children.empty)
				{
					tmp.v = v.front.children[0].text;
				}
			}
			if(c.type == EntityType.elementStart) {
				auto f = c.children.filter!(c => c.name == "f");
				if(!f.empty && f.front.type == EntityType.elementStart) {
					tmp.f = f.front.children[0].text;
				}
			}
			ret ~= tmp;
		}
	}
	return ret;
}

Cell[] insertValueIntoCell(Cell[] cells, Data[] ss) @trusted {
	immutable excepted = ["n", "s", "b", "e", "str", "inlineStr"];
	immutable same = ["n", "e", "str", "inlineStr"];
	foreach(ref Cell c; cells) {
		assert(canFind(excepted, c.t) || c.t.empty,
				format("'%s' not in [%s]", c.t, excepted));
		if(c.t.empty) {
			c.value = convert(c.v);
		} else if(canFind(same, c.t)) {
			c.value = convert(c.v);
		} else if(c.t == "b") {
			//logf("'%s' %s", c.v, c);
			c.value = c.v == "1";
		} else {
			if(!c.v.empty) {
				size_t idx = to!size_t(c.v);
				//logf("'%s' %s", c.v, idx);
				c.value = ss[idx];
			}
		}
	}
	return cells;
}

Pos toPos(in string s) {
	import std.string : indexOfAny;
	import std.math : pow;
	ptrdiff_t fn = s.indexOfAny("0123456789");
	enforce(fn != -1, s);
	size_t row = to!size_t(to!long(s[fn .. $]) - 1);
	size_t col = 0;
	string colS = s[0 .. fn];
	foreach(idx, char c; colS) {
		col = col * 26 + (c - 'A' + 1);
	}
	return Pos(row, col - 1);
}

unittest {
	assert(toPos("A1").col == 0);
	assert(toPos("Z1").col == 25);
	assert(toPos("AA1").col == 26);
}

Pos elementMax(Pos a, Pos b) {
	return Pos(a.row < b.row ? b.row : a.row,
			a.col < b.col ? b.col : a.col);
}

@trusted unittest {
	import std.math : approxEqual;
	auto r = readSheet("multitable.xlsx", "wb1");
	assert(approxEqual(r.table[12][5].value.get!double(), 26.74),
			format("%s", r.table[12][5])
		);

	assert(approxEqual(r.table[13][5].value.get!double(), -26.74),
			format("%s", r.table[13][5])
		);
}

@trusted unittest {
	import std.algorithm.comparison : equal;
	auto s = readSheet("multitable.xlsx", "wb1");
	auto r = s.iterateRow!long(15, 1, 6);

	assert(r.canConvertTo(CellType.long_));
	assert(r.canConvertTo(CellType.double_));
	assert(r.canConvertTo(CellType.string_));

	auto expected = [1, 2, 3, 4, 5];
	assert(equal(r, expected), format("%s", r));

	auto r2 = s.getRow!long(15, 1, 6);
	assert(equal(r, expected));

	auto it = s.iterateRowLong(15, 1, 6);
	assert(equal(r2, it));

	auto it2 = s.iterateRowUntyped(15, 1, 6)
		.map!(it => format("%s", it))
		.array;
}

@trusted unittest {
	import std.algorithm.comparison : equal;
	auto s = readSheet("multitable.xlsx", "wb2");
	//writefln("%s\n%(%s\n%)", s.maxPos, s.cells);
	auto rslt = s.iterateColumn!Date(1, 1, 6);
	auto rsltUt = s.iterateColumnUntyped(1, 1, 6)
		.map!(it => format("%s", it))
		.array;
	assert(!rsltUt.empty);

	auto target = [Date(2019,5,01), Date(2016,12,27), Date(1976,7,23),
		 Date(1986,7,2), Date(2038,1,19)
	];
	assert(equal(rslt, target), format("\n%s\n%s", rslt, target));

	auto it = s.getColumn!Date(1, 1, 6);
	assert(equal(rslt, it));

	auto it2 = s.getColumnDate(1, 1, 6);
	assert(equal(rslt, it2));
}

@trusted unittest {
	import std.algorithm.comparison : equal;
	auto s = readSheet("multitable.xlsx", "Sheet3");
	writeln(s.table[0][0].value.type());
	assert(s.table[0][0].value.peek!long(),
			format("%s", s.table[0][0].value));
	assert(s.table[0][0].canConvertTo(CellType.bool_));
}

@trusted unittest {
	import std.file : dirEntries, SpanMode;
	import std.traits : EnumMembers;
	foreach(de; dirEntries("xlsx_files/", "*.xlsx", SpanMode.depth)
			.filter!(a => a.name != "xlsx_files/data03.xlsx"))
	{
		//writeln(de.name);
		auto sn = sheetNames(de.name);
		foreach(s; sn) {
			auto sheet = readSheet(de.name, s.name);
			foreach(cell; sheet.cells) {
				foreach(T; [EnumMembers!CellType]) {
					auto cc = cell.canConvertTo(T);
				}
			}
		}
	}
}

unittest {
	import std.algorithm.comparison : equal;
	auto sheet = readSheet("testworkbook.xlsx", "ws1");
	//writefln("%(%s\n%)", sheet.cells);
	//writeln(sheet.toString());
	//assert(sheet.table[2][3].value.get!long() == 1337);

	auto c = sheet.getColumnLong(3, 2, 5);
	auto r = [1337, 2, 3];
	assert(equal(c, r), format("%s", c));

	auto c2 = sheet.getColumnString(4, 2, 5);
	string f2 = sheet.table[2][4].convertToString();
	assert(f2 == "hello", f2);
	f2 = sheet.table[3][4].convertToString();
	assert(f2 == "sil", f2);
	f2 = sheet.table[4][4].convertToString();
	assert(f2 == "foo", f2);
	auto r2 = ["hello", "sil", "foo"];
	assert(equal(c2, r2), format("%s", c2));
}

@trusted unittest {
	import std.math : approxEqual;
	auto sheet = readSheet("toto.xlsx", "Trades");
	writefln("%(%s\n%)", sheet.cells);

	auto r = sheet.getRowString(1, 0, 2).array;

	double d = to!double(r[1]);
	assert(approxEqual(d, 38204642.510000));
}
