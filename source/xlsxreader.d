/* TODO: Qualify `std.zip.ZipArchive` members as, at least, `pure` and then the
   functions in there that use them.
*/
module xlsxreader;

import core.time : Duration;
import std.algorithm.iteration : filter, map;
import std.algorithm.searching : all, canFind, startsWith;
import std.algorithm.sorting : sort;
import std.array : array, empty, front, popFront;
import std.conv : to;
import std.datetime : DateTime, Date, TimeOfDay;
import std.datetime.stopwatch : StopWatch, AutoStart;
import std.exception : enforce;
import std.file : read, exists, readText;
import std.format : format;
import std.stdio;
import std.traits : isIntegral, isFloatingPoint, isSomeString;
import std.typecons : Nullable, nullable;
import std.zip;
import mir.algebraic : Algebraic;
import dxml.dom : DOMEntity, EntityType, parseDOM;

// disabled for now for faster builds
// version = ctRegex_test;

version(xlsxreader_profileGC)
	enum runCount = 1;
else version(xlsxreader_benchmark)
    enum runCount = 10;

/** Cell position row 0-based offset. */
alias RowOffset = uint;

/** Cell position column 0-based offset. */
alias ColOffset = uint;

/** Cell Position.

	Excel has a limit of 1,048,576 rows and 16,384 columns per sheet so use
	32-bit precision now.

	See: https://support.socrata.com/hc/en-us/articles/115005306167-Limitations-of-Excel-and-CSV-Downloads
 */
struct Pos {
	RowOffset row;
	ColOffset col;
}

/// Cell Data.
alias Data = Algebraic!(bool, long, double, string, DateTime, Date, TimeOfDay);

/// Sheet Cell.
struct Cell {
	string loc; ///< Location.
	RowOffset row; ///< Row. row[r]
	string t; // s or n, s for pointer, n for value, stored in v
	string r; // c[r]
	string v; // c.v the value or ptr
	string f; // c.f the formula
	string xmlValue; ///< Value stored in cell.
	Pos position; ///< Position of cell.
}

/// Cell Type.
enum CellType {
	datetime,
	timeofday,
	date,
	bool_,
	double_,
	long_,
	string_
}

/// Sheet.
struct Sheet {
	import std.ascii : toUpper;

	@property const(Cell)[] cells() const return scope @safe pure nothrow @nogc {
		return _cells;
	}

	const string name;			///< Name of sheet.
	private Cell[] _cells;		///< Cells of sheet.
	Cell[][] table;				// TODO: make this read-only and lazily constructed behind a property
	const Pos maxPos;

	@property string toString() const scope @safe pure {
		import std.array : appender;
		auto result = appender!(typeof(return));
		toString(result);
		return result.data[];
	}

	void toString(Sink)(ref scope Sink sink) const scope {
		import std.format : formattedWrite;
		long[] maxCol = new long[](maxPos.col + 1);
		foreach (const ref row; this.table) {
			foreach (const idx, const ref Cell col; row) {
				const string s = col.xmlValue;

				maxCol[idx] = maxCol[idx] < s.length ? s.length : maxCol[idx];
			}
		}
		maxCol[] += 1;

		foreach (const ref row; this.table) {
			foreach (const idx, const ref Cell col; row)
				sink.formattedWrite("%*s, ", maxCol[idx], col.xmlValue);
			sink.formattedWrite("\n");
		}
	}

	void printTable() const scope @trusted {
		writeln(this);			// uses toString(Sink)
	}

	// Column

	Iterator!T getColumn(T)(size_t col, size_t startRow, size_t endRow) return scope @trusted {
		auto c = this.iterateColumn!T(col, startRow, endRow);
		return typeof(return)(c.array);
	}

	private enum t = q{
	Iterator!(%1$s) getColumn%2$s(size_t col, size_t startRow, size_t endRow) return scope @trusted {
		return getColumn!(%1$s)(col, startRow, endRow);
	}
	};
	static foreach (T; ["long", "double", "string", "Date", "TimeOfDay",
			"DateTime"])
	{
		mixin(format(t, T, T[0].toUpper ~ T[1 .. $]));
	}

	ColumnUntyped iterateColumnUntyped(size_t col, size_t startRow, size_t endRow) return scope @trusted {
		return typeof(return)(&this, col, startRow, endRow);
	}

	Column!(T) iterateColumn(T)(size_t col, size_t startRow, size_t endRow) return scope @trusted {
		return typeof(return)(&this, col, startRow, endRow);
	}

	Column!(long) iterateColumnLong(size_t col, size_t startRow, size_t endRow) return scope @trusted {
		return typeof(return)(&this, col, startRow, endRow);
	}

	Column!(double) iterateColumnDouble(size_t col, size_t startRow, size_t endRow) return scope @trusted {
		return typeof(return)(&this, col, startRow, endRow);
	}

	Column!(string) iterateColumnString(size_t col, size_t startRow, size_t endRow) return scope @trusted {
		return typeof(return)(&this, col, startRow, endRow);
	}

	Column!(DateTime) iterateColumnDateTime(size_t col, size_t startRow, size_t endRow) return scope @trusted {
		return typeof(return)(&this, col, startRow, endRow);
	}

	Column!(Date) iterateColumnDate(size_t col, size_t startRow, size_t endRow) return scope @trusted {
		return typeof(return)(&this, col, startRow, endRow);
	}

	Column!(TimeOfDay) iterateColumnTimeOfDay(size_t col, size_t startRow, size_t endRow) return scope @trusted {
		return typeof(return)(&this, col, startRow, endRow);
	}

	// Row

	Iterator!T getRow(T)(size_t row, size_t startColumn, size_t endColumn) return scope @trusted {
		return typeof(return)(this.iterateRow!T(row, startColumn, endColumn).array); // TODO: why .array?
	}

	private enum t2 = q{
	Iterator!(%1$s) getRow%2$s(size_t row, size_t startColumn, size_t endColumn) return scope @trusted {
		return getRow!(%1$s)(row, startColumn, endColumn);
	}
	};
	static foreach (T; ["long", "double", "string", "Date", "TimeOfDay", "DateTime"])
	{
		mixin(format(t2, T, T[0].toUpper ~ T[1 .. $]));
	}

	RowUntyped iterateRowUntyped(size_t row, size_t startColumn, size_t endColumn) return scope @trusted {
		return typeof(return)(&this, row, startColumn, endColumn);
	}

	Row!(T) iterateRow(T)(size_t row, size_t startColumn, size_t endColumn) return scope @trusted {
		return typeof(return)(&this, row, startColumn, endColumn);
	}

	Row!(long) iterateRowLong(size_t row, size_t startColumn, size_t endColumn) return scope @trusted {
		return typeof(return)(&this, row, startColumn, endColumn);
	}

	Row!(double) iterateRowDouble(size_t row, size_t startColumn, size_t endColumn) return scope @trusted {
		return typeof(return)(&this, row, startColumn, endColumn);
	}

	Row!(string) iterateRowString(size_t row, size_t startColumn, size_t endColumn) return scope @trusted {
		return typeof(return)(&this, row, startColumn, endColumn);
	}

	Row!(DateTime) iterateRowDateTime(size_t row, size_t startColumn, size_t endColumn) return scope @trusted {
		return typeof(return)(&this, row, startColumn, endColumn);
	}

	Row!(Date) iterateRowDate(size_t row, size_t startColumn, size_t endColumn) return scope @trusted {
		return typeof(return)(&this, row, startColumn, endColumn);
	}

	Row!(TimeOfDay) iterateRowTimeOfDay(size_t row, size_t startColumn, size_t endColumn) return scope @trusted {
		return typeof(return)(&this, row, startColumn, endColumn);
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

	this(Sheet* sheet, size_t row, size_t startColumn, size_t endColumn) pure nothrow @nogc @safe {
		assert(sheet.table.length == sheet.maxPos.row + 1);
		this.sheet = sheet;
		this.row = row;
		this.startColumn = startColumn;
		this.endColumn = endColumn;
		this.cur = this.startColumn;
	}

	@property bool empty() const pure nothrow @nogc @safe {
		return this.cur >= this.endColumn;
	}

	void popFront() pure nothrow @nogc @safe {
		++this.cur;
	}

	inout(typeof(this)) save() inout pure nothrow @nogc @safe {
		return this;
	}

	@property inout(Cell) front() inout pure nothrow @nogc @safe {
		return this.sheet.table[this.row][this.cur];
	}
}

/// Sheet Row.
struct Row(T) {
	RowUntyped ru;
	T front;

	this(Sheet* sheet, size_t row, size_t startColumn, size_t endColumn) {
		this.ru = RowUntyped(sheet, row, startColumn, endColumn);
		this.read();
	}

	@property bool empty() const pure nothrow @nogc {
		return this.ru.empty;
	}

	void popFront() {
		this.ru.popFront();
		if (!this.empty) {
			this.read();
		}
	}

	inout(typeof(this)) save() inout pure nothrow @nogc {
		return this;
	}

	private void read() {
		this.front = convertTo!T(this.ru.front.xmlValue);
	}
}

///
struct ColumnUntyped {
	Sheet* sheet;
	const size_t col;
	const size_t startRow;
	const size_t endRow;
	size_t cur;

	this(Sheet* sheet, size_t col, size_t startRow, size_t endRow) @safe {
		assert(sheet.table.length == sheet.maxPos.row + 1);
		this.sheet = sheet;
		this.col = col;
		this.startRow = startRow;
		this.endRow = endRow;
		this.cur = this.startRow;
	}

	@property bool empty() const pure nothrow @nogc @safe {
		return this.cur >= this.endRow;
	}

	void popFront() @safe {
		++this.cur;
	}

	inout(typeof(this)) save() inout pure nothrow @nogc @safe {
		return this;
	}

	@property Cell front() @safe {
		return this.sheet.table[this.cur][this.col];
	}
}

/// Sheet Column.
struct Column(T) {
	ColumnUntyped cu;

	T front;

	this(Sheet* sheet, size_t col, size_t startRow, size_t endRow) {
		this.cu = ColumnUntyped(sheet, col, startRow, endRow);
		this.read();
	}

	@property bool empty() const pure nothrow @nogc {
		return this.cu.empty;
	}

	void popFront() {
		this.cu.popFront();
		if (!this.empty) {
			this.read();
		}
	}

	inout(typeof(this)) save() inout pure nothrow @nogc {
		return this;
	}

	private void read() {
		this.front = convertTo!T(this.cu.front.xmlValue);
	}
}

version(xlsxreader_test) @safe unittest {
	import std.range : isForwardRange;
	import std.meta : AliasSeq;
	static foreach (T; AliasSeq!(long,double,DateTime,TimeOfDay,Date,string)) {{
		alias C = Column!T;
		alias R = Row!T;
		alias I = Iterator!T;
		static assert(isForwardRange!C, C.stringof);
		static assert(isForwardRange!R, R.stringof);
		static assert(isForwardRange!I, I.stringof);
	}}
}

Date longToDate(long d) @safe {
	// modifed from https://www.codeproject.com/Articles/2750/
	// Excel-Serial-Date-to-Day-Month-Year-and-Vice-Versa

	// Excel/Lotus 123 have a bug with 29-02-1900. 1900 is not a
	// leap year, but Excel/Lotus 123 think it is...
	if (d == 60) {
		return Date(1900, 2,  29);
	} else if (d < 60) {
		// Because of the 29-02-1900 bug, any serial date
		// under 60 is one off... Compensate.
		++d;
	}

	// Modified Julian to DMY calculation with an addition of 2415019
	int l = cast(int)d + 68569 + 2415019;
	const int n = int(( 4 * l ) / 146097);
	l = l - int(( 146097 * n + 3 ) / 4);
	const int i = int(( 4000 * ( l + 1 ) ) / 1461001);
	l = l - int(( 1461 * i ) / 4) + 31;
	const int j = int(( 80 * l ) / 2447);
	const nDay = l - int(( 2447 * j ) / 80);
	l = int(j / 11);
	const int nMonth = j + 2 - ( 12 * l );
	const int nYear = 100 * ( n - 49 ) + i + l;
	return Date(nYear, nMonth, nDay);
}

long dateToLong(Date d) @safe {
	// modifed from https://www.codeproject.com/Articles/2750/
	// Excel-Serial-Date-to-Day-Month-Year-and-Vice-Versa

	// Excel/Lotus 123 have a bug with 29-02-1900. 1900 is not a
	// leap year, but Excel/Lotus 123 think it is...
	if (d.day == 29 && d.month == 2 && d.year == 1900) {
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

	if (nSerialDate < 60) {
		// Because of the 29-02-1900 bug, any serial date
		// under 60 is one off... Compensate.
		nSerialDate--;
	}

	return nSerialDate;
}

version(xlsxreader_test) @safe unittest {
	auto ds = [ Date(1900,2,1), Date(1901, 2, 28), Date(2019, 06, 05) ];
	foreach (const d; ds) {
		const long l = dateToLong(d);
		const Date r = longToDate(l);
		assert(r == d, format("%s %s", r, d));
	}
}

TimeOfDay doubleToTimeOfDay(double s) @safe {
	import core.stdc.math : lround;
	const double secs = (24.0 * 60.0 * 60.0) * s;

	// TODO not one-hundred my lround is needed
	const int secI = to!int(lround(secs));

	return TimeOfDay(secI / 3600, (secI / 60) % 60, secI % 60);
}

double timeOfDayToDouble(TimeOfDay tod) @safe {
	const long h = tod.hour * 60 * 60;
	const long m = tod.minute * 60;
	const long s = tod.second;
	return (h + m + s) / (24.0 * 60.0 * 60.0);
}

version(xlsxreader_test) @safe unittest {
	auto tods = [ TimeOfDay(23, 12, 11), TimeOfDay(11, 0, 11),
		 TimeOfDay(0, 0, 0), TimeOfDay(0, 1, 0),
		 TimeOfDay(23, 59, 59), TimeOfDay(0, 0, 0)];
	foreach (const tod; tods) {
		const double d = timeOfDayToDouble(tod);
		assert(d <= 1.0, format("%s", d));
		TimeOfDay r = doubleToTimeOfDay(d);
		assert(r == tod, format("%s %s", r, tod));
	}
}

double datetimeToDouble(DateTime dt) @safe {
	const double d = dateToLong(dt.date);
	const double t = timeOfDayToDouble(dt.timeOfDay);
	return d + t;
}

DateTime doubleToDateTime(double d) @safe {
	long l = cast(long)d;
	Date dt = longToDate(l);
	TimeOfDay t = doubleToTimeOfDay(d - l);
	return DateTime(dt, t);
}

version(xlsxreader_test) @safe unittest {
	auto ds = [ Date(1900,2,1), Date(1901, 2, 28), Date(2019, 06, 05) ];
	auto tods = [ TimeOfDay(23, 12, 11), TimeOfDay(11, 0, 11),
		 TimeOfDay(0, 0, 0), TimeOfDay(0, 1, 0),
		 TimeOfDay(23, 59, 59), TimeOfDay(0, 0, 0)];
	foreach (const d; ds) {
		foreach (const tod; tods) {
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

Date stringToDate(string s) @safe {
	import std.array : split;
	import std.string : indexOf;

	if (s.indexOf('/') != -1) {
		auto sp = s.split('/');
		enforce(sp.length == 3, format("[%s]", sp));
		return Date(to!int(sp[2]), to!int(sp[1]), to!int(sp[0]));
	} else {
		return longToDate(to!long(s));
	}
}

bool tryConvertTo(T,S)(S var) {
	return !(tryConvertToImpl!T(Data(var)).isNull());
}

Nullable!(T) tryConvertToImpl(T)(Data var) {
	try {
		return nullable(convertTo!T(var));
	} catch(Exception e) {
		return Nullable!T();
	}
}

T convertTo(T)(string var) @safe {
	import std.math : lround;
	static if (isSomeString!T) {
		return to!T(var);
	} else static if (is(T == bool)) {
		return var == "1";
	} else static if (isIntegral!T) {
		return to!T(var);
	} else static if (isFloatingPoint!T) {
		return to!T(var);
	} else static if (is(T == DateTime)) {
		if (var.canConvertToLong()) {
			return doubleToDateTime(to!long(var));
		} else if (var.canConvertToDouble()) {
			return doubleToDateTime(to!double(var));
		}
		enforce(false, "Can not convert '" ~ var ~ "' to a DateTime");
		assert(false, "Unreachable");
	} else static if (is(T == Date)) {
		if (var.canConvertToLong()) {
			return longToDate(to!long(var));
		} else if (var.canConvertToDouble()) {
			return longToDate(lround(to!double(var)));
		}
		return stringToDate(var);
	} else static if (is(T == TimeOfDay)) {
		const double l = to!double(var);
		return doubleToTimeOfDay(l - cast(long)l);
	} else {
		static assert(false, T.stringof ~ " not supported");
	}
}

private ZipArchive readFile(in string filename) @trusted {
	enforce(exists(filename), "File with name " ~ filename ~ " does not exist");
	return new typeof(return)(read(filename));
}

private static immutable workbookXMLPath = "xl/workbook.xml";
private static immutable sharedStringXMLPath = "xl/sharedStrings.xml";
private static immutable relsXMLPath = "xl/_rels/workbook.xml.rels";

private static expandTrusted(ZipArchive za, ArchiveMember de) @trusted /* TODO: pure */ {
    enum tme = true;            // time me
    static if (tme) auto sw = StopWatch(AutoStart.yes);
	auto ret = za.expand(de);
    static if (tme) writeln("expand length:", ret.length, " took: ", sw.peek());
    return ret;
}

struct Relationships {
	string id;
	string file;
}

alias RelationshipsById = Relationships[string];

/// File.
struct File {
	// TODO: convert to constructor this(in string filename)?
	static typeof(this) fromPath(in string filename) @trusted {
        return typeof(return)(filename, new ZipArchive(read(filename)));
	}

    private SheetNameId[] sheetNameIds() @safe /* TODO: pure */ {
		auto dom = workbookDOM();
        if (dom.children.length != 1)
            return [];

		auto workbook = dom.children[0];
        if (workbook.name != "workbook" &&
            workbook.name != "s:workbook")
            return [];

		const sheetName = workbook.name == "workbook" ? "sheets" : "s:sheets";
		auto sheetsRng = workbook.children.filter!(c => c.name == sheetName);
        if (sheetsRng.empty)
            return [];

		return sheetsRng.front.children.map!(
            // TODO: optimize by using indexOf or find
			s => SheetNameId(s.attributes
							  .filter!(a => a.name == "name")
							  .front
							  .value
							  .specialCharacterReplacementReverseLazy.to!string(),
							 s.attributes
							  .filter!(a => a.name == "sheetId")
							  .front
							  .value
							  .to!int(),
							 s.attributes
							  .filter!(a => a.name == "r:id")
							  .front.value)).array;
    }

	/// Returns: lazy range over sheets.
    auto bySheet() @safe {
        return sheetNameIds.map!((const scope SheetNameId sheetNameId) {
                return extractSheet(_za, relationships, filename, sheetNameId.rid, sheetNameId.name);
			});
    }

	/// Get (and cache) DOM.
	private DOMEntity!(string) workbookDOM() @safe /* TODO: pure */ {
		import dxml.parser : Config, SkipComments, SkipPI, SplitEmpty, ThrowOnEntityRef;
		auto ent = workbookXMLPath in _za.directory;
		// TODO: use enforce(ent ! is null); instead?
		if (ent is null)
			return typeof(return).init;
		if (_wbDOM == _wbDOM.init) {
            auto est = _za.expandTrusted(*ent).convertToString();
            enum tme = true;           // time me
            static if (tme) auto sw = StopWatch(AutoStart.yes);
            _wbDOM = est.parseDOM!(Config(SkipComments.no, // TODO: change to SkipComments.yes and validate
                                        SkipPI.no, // TODO: change to SkipPI.yes and validate
                                        SplitEmpty.no, // default is ok
                                        ThrowOnEntityRef.yes))(); // default is ok
			enforce(_wbDOM.children.length == 1,
					"Expected a single DOM child but got " ~ _wbDOM.children.length.to!string);
            static if (tme) writeln("parseDOM length:", est.length, " took: ", sw.peek());
		}
		return _wbDOM;
	}

	private RelationshipsById relationships() @safe /* TODO: pure */ {
		if (_rels is null)
			_rels = parseRelationships(_za.directory[relsXMLPath]);
		return _rels;
	}

	private RelationshipsById parseRelationships(ArchiveMember am) @safe {
		auto est = _za.expandTrusted(am).convertToString();
        // import std.digest : digest;
        // import std.digest.md : MD5;
        // writeln("am.name:", am.name, " md5:", est.digest!MD5);
		auto dom = est.parseDOM();
		enforce(dom.children.length == 1,
				"Expected a single DOM child but got " ~ dom.children.length.to!string);

		auto rel = dom.children[0];
		enforce(rel.name == "Relationships",
				"Expected rel.name to be \"Relationships\" but was " ~ rel.name);

		typeof(return) ret;
		static if (is(typeof(ret.reserve(size_t.init)) == void)) {
			/* Use reserve() when AA gets it or `RelationshipsById` is a custom hash
			 * map. */
			ret.reserve(rel.children.length);
		}
		foreach (ref r; rel.children.filter!(c => c.name == "Relationship")) {
			Relationships tmp;
			tmp.id = r.attributes.filter!(a => a.name == "Id").front.value;
			tmp.file = r.attributes.filter!(a => a.name == "Target").front.value;
			ret[tmp.id] = tmp;
		}
		enforce(!ret.empty);
		return ret;
	}

	const string filename;
	private ZipArchive _za;
	DOMEntity!string _wbDOM;	///< Workbook.
	private RelationshipsById _rels;
}

/// benchmark reading of "multitable.xlsx"
version(xlsxreader_benchmark) @safe unittest {
	import std.meta : AliasSeq;
    static void use_bySheet_multitable() @trusted {
        File file = File.fromPath("multitable.xlsx");
        foreach (ref sheet; file.bySheet) {
        }
    }
	alias funs = AliasSeq!(use_bySheet_multitable);
	auto results = benchmarkMin!(funs)(runCount);
	foreach (const i, fun; funs) {
		writeln(fun.stringof, " took ", results[i]);
	}
}

/// benchmark reading of "50xP_sheet1.xlsx"
version(xlsxreader_benchmark) @safe unittest {
	enum path = "50xP_sheet1.xlsx";
	import std.meta : AliasSeq;
	static void use_sheetNamesAndreadSheet() @trusted {
		foreach (const i, const ref s; sheetNames(path)) {
			auto sheet = readSheet(path, s.name);
		}
	}
    static void use_bySheet() @trusted {
        File file = File.fromPath(path);
		size_t i = 0;
        foreach (ref sheet; file.bySheet) {
			i++;
        }
    }
	alias funs = AliasSeq!(use_sheetNamesAndreadSheet,
						   use_bySheet);
	auto results = benchmarkMin!(funs)(runCount);
	foreach (const i, fun; funs) {
		writeln(fun.stringof[0 .. $-2], "(\"", path, "\") took ", results[i]);
	}
}

/// Sheet name, id and rid.
struct SheetNameId {
	string name;
	int id;
	string rid;
}

/// Strip BOM and convert ubyte[] to a string.
string convertToString(inout(ubyte)[] d) @trusted {
	import std.encoding : getBOM, BOM, transcode;
	const b = getBOM(d);
	switch (b.schema) {
	case BOM.none:
		return cast(string)d;
	case BOM.utf8:
		return cast(string)(d[3 .. $]);
	case BOM.utf16be:
	case BOM.utf16le:
	case BOM.utf32be:
	case BOM.utf32le:
		goto default;
	default:
		string ret;
		transcode(d, ret);
		return ret;
	}
}

/// Read sheet names stored in `filename`.
SheetNameId[] sheetNames(in string filename) @trusted {
	return File.fromPath(filename).sheetNameIds();
}

version(xlsxreader_test) @safe unittest {
	auto r = sheetNames("multitable.xlsx");
	assert(r[0].name == "wb1");
	assert(r[0].id == 1);
}

version(xlsxreader_test) @safe unittest {
	auto r = sheetNames("sheetnames.xlsx");
	assert(r[0].name == "A & B ;", r[0].name);
	assert(r[0].id == 1);
}

deprecated("use File.relationships(ArchiveMember am) instead")
RelationshipsById parseRelationships(ZipArchive za, ArchiveMember am) @safe {
	auto dom = za.expandTrusted(am)
				 .convertToString()
				 .parseDOM();
	enforce(dom.children.length == 1,
			"Expected a single DOM child but got " ~ dom.children.length.to!string);

	auto rel = dom.children[0];
	enforce(rel.name == "Relationships",
			"Expected rel.name to be \"Relationships\" but was " ~ rel.name);

	typeof(return) ret;
	static if (is(typeof(ret.reserve(size_t.init)) == void)) {
		/* Use reserve() when AA gets it or `RelationshipsById` is a custom hash
		 * map. */
		ret.reserve(rel.children.length);
	}
	foreach (ref r; rel.children.filter!(c => c.name == "Relationship")) {
		Relationships tmp;
		tmp.id = r.attributes.filter!(a => a.name == "Id").front.value;
		tmp.file = r.attributes.filter!(a => a.name == "Target").front.value;
		ret[tmp.id] = tmp;
	}
	enforce(!ret.empty);
	return ret;
}

/// Read sheet named `sheetName` from `filename`.
Sheet readSheet(in string filename, in string sheetName) @safe {
	const SheetNameId[] sheets = sheetNames(filename);
	auto sRng = sheets.filter!(s => s.name == sheetName);
	enforce(!sRng.empty, "No sheet with name " ~ sheetName
			~ " found in file " ~ filename);
	return readSheetImpl(filename, sRng.front.rid, sheetName);
}

string eatXlPrefix(scope return string fn) @safe pure nothrow @nogc {
    static immutable xlPrefixes = ["xl//", "/xl/"];
	foreach (const p; xlPrefixes) {
		// TODO: use fn.skipOver("xl//", "/xl/") when it’s nothrow @nogc
		if (fn.startsWith(p)) {
			return fn[p.length .. $];
		}
	}
	return fn;
}

Sheet readSheetImpl(in string filename,
                    in string rid, in string sheetName) @safe {
	scope(failure)
		writefln("Failed at file '%s' and sheet '%s'", filename, rid);
	auto za = readFile(filename);
	auto rels = parseRelationships(za, za.directory[relsXMLPath]);
    return extractSheet(za, rels, filename, rid, sheetName);
}

private Sheet extractSheet(ZipArchive za,
						   in RelationshipsById rels,
						   in string filename,
                           in string rid, in string sheetName) @trusted {
	string[] ss;
	if (ArchiveMember* amPtr = sharedStringXMLPath in za.directory)
		ss = readSharedEntries(za, *amPtr);
	//logf("%s", ss);

	const Relationships* sheetRel = rid in rels;
	enforce(sheetRel !is null, format("Could not find '%s' in '%s'", rid,
                                      filename));
	const fn = "xl/" ~ eatXlPrefix(sheetRel.file);
	ArchiveMember* sheet = fn in za.directory;
	enforce(sheet !is null, format("sheetRel.file orig '%s', fn %s not in [%s]",
				sheetRel.file, fn, za.directory.keys()));

	auto cells = insertValueIntoCell(readCells(za, *sheet), ss);

	Pos maxPos;
	foreach (ref c; cells) {
		c.position = toPos(c.r);
		maxPos = elementMax(maxPos, c.position);
	}

	auto table = new Cell[][](maxPos.row + 1, maxPos.col + 1);
	foreach (const ref c; cells) {
		table[c.position.row][c.position.col] = c;
	}

	return typeof(return)(sheetName, cells, table, maxPos);
}

string[] readSharedEntries(ZipArchive za, ArchiveMember am) @trusted {
	ubyte[] ss = za.expand(am);
	string ssData = convertToString(ss);
	auto dom = parseDOM(ssData);
	string[] ret;
	if (dom.type != EntityType.elementStart) {
		return ret;
	}
	assert(dom.children.length == 1);
	auto sst = dom.children[0];
	assert(sst.name == "sst");
	if (sst.type != EntityType.elementStart || sst.children.empty) {
		return ret;
	}
	auto siRng = sst.children.filter!(c => c.name == "si");
	foreach (ref si; siRng) {
		if (si.type != EntityType.elementStart) {
			continue;
		}
		//ret ~= extractData(si);
		string tmp;
		foreach (ref tORr; si.children) {
			if (tORr.name == "t" && tORr.type == EntityType.elementStart
					&& !tORr.children.empty)
			{
				//ret ~= Data(convert(tORr.children[0].text));
				ret ~= tORr.children[0].text.removeSpecialCharacter();
			} else if (tORr.name == "r") {
				foreach (ref r; tORr.children.filter!(r => r.name == "t")) {
					if (r.type == EntityType.elementStart && !r.children.empty) {
						tmp ~= r.children[0].text.removeSpecialCharacter();
					}
				}
			} else {
				//ret ~= Data.init;
				ret ~= "";
			}
		}
		if (!tmp.empty) {
			//ret ~= Data(convert(tmp));
			ret ~= tmp.removeSpecialCharacter();
		}
	}
	return ret;
}

string extractData(DOMEntity!string si) {
	string tmp;
	foreach (ref tORr; si.children) {
		if (tORr.name == "t") {
			if (!tORr.attributes.filter!(a => a.name == "xml:space").empty) {
				return "";
			} else if (tORr.type == EntityType.elementStart
					&& !tORr.children.empty)
			{
				return tORr.children[0].text;
			} else {
				return "";
			}
		} else if (tORr.name == "r") {
			foreach (ref r; tORr.children.filter!(r => r.name == "t")) {
				tmp ~= r.children[0].text;
			}
		}
	}
	if (!tmp.empty) {
		return tmp;
	}
	assert(false);
}

private bool canConvertToLong(in string s) @safe pure nothrow @nogc {
	import std.utf : byChar;
	import std.ascii : isDigit;
	if (s.empty)
		return false;
	return s.byChar.all!isDigit();
}

version(ctRegex_test)
version(unittest)
{
	import std.regex : ctRegex, matchAll;
	private static immutable rs = r"[\+-]{0,1}[0-9][0-9]*\.[0-9]*";
	private static immutable rgx = ctRegex!rs;
	private bool canConvertToDoubleOld(in string s) @safe {
		auto cap = matchAll(s, rgx);
		return cap.empty || cap.front.hit != s ? false : true;
	}
}

private bool canConvertToDouble(string s) pure @safe nothrow @nogc {
	if (s.startsWith('+', '-')) {
		s = s[1 .. $];
	}
	if (s.empty) {
		return false;
	}

	if (s[0] < '0' || s[0] > '9') { // at least one in [0-9]
		return false;
	}

	s = s[1 .. $];

	if (s.empty) {
		return true;
	}

	while (!s.empty && s[0] >= '0' && s[0] <= '9') {
		s = s[1 .. $];
	}
	if (s.empty) {
		return true;
	}
	if (s[0] != '.') {
		return false;
	}
	s = s[1 .. $];
	if (s.empty) {
		return true;
	}

	while (!s.empty && s[0] >= '0' && s[0] <= '9') {
		s = s[1 .. $];
	}

	return s.empty;
}

version(xlsxreader_test) @safe unittest {
	static struct Test {
		string tt;
		bool rslt;
	}
	auto tests =
		[ Test("-", false)
		, Test("0.0", true)
		, Test("-0.", true)
		, Test("-0.0", true)
		, Test("-0.a", false)
		, Test("-0.0", true)
		, Test("-1100.0", true)
		];
	foreach (const t; tests) {
        version(ctRegex_test)
            assert(canConvertToDouble(t.tt) == canConvertToDoubleOld(t.tt) ,
                   format("%s %s %s %s", t.tt,
                          canConvertToDouble(t.tt),
                          canConvertToDoubleOld(t.tt),
                          t.rslt));
		assert(canConvertToDouble(t.tt) == t.rslt,
               format("%s %s %s", t.tt,
                      canConvertToDouble(t.tt),
                      t.rslt));
	}
}

string removeSpecialCharacter(string s) {
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
		// TODO: use substitute.array
		import std.array : replace;
		foreach (const ref tr; toRe) {
			while (canFind(s, tr.from)) {
				s = s.replace(tr.from, tr.to);
			}
		}
		return s;
	}

	return replaceStrings(s);
}

Cell[] readCells(ZipArchive za, ArchiveMember am) @trusted {
	Cell[] ret;
	ubyte[] ss = za.expand(am);
	string ssData = convertToString(ss);
	auto dom = parseDOM(ssData);
	assert(dom.children.length == 1);
	auto ws = dom.children[0];
	if (ws.name != "worksheet") {
		return ret;
	}
	auto sdRng = ws.children.filter!(c => c.name == "sheetData");
	assert(!sdRng.empty);
	if (sdRng.front.type != EntityType.elementStart) {
		return ret;
	}
	auto rows = sdRng.front.children
		.filter!(r => r.name == "row");

	foreach (ref row; rows) {
		if (row.type != EntityType.elementStart || row.children.empty) {
			continue;
		}
		foreach (ref c; row.children.filter!(r => r.name == "c")) {
			Cell tmp;
			tmp.row = row.attributes.filter!(a => a.name == "r")
				.front.value.to!RowOffset();
			tmp.r = c.attributes.filter!(a => a.name == "r")
				.front.value;
			auto t = c.attributes.filter!(a => a.name == "t");
			if (t.empty) {
				// we assume that no t attribute means direct number
				//writefln("Found a strange empty cell \n%s", c);
			} else {
				tmp.t = t.front.value;
			}
			if (tmp.t == "s" || tmp.t == "n") {
				if (c.type == EntityType.elementStart) {
					auto v = c.children.filter!(c => c.name == "v");
					//enforce(!v.empty, format("r %s", tmp.row));
					if (!v.empty && v.front.type == EntityType.elementStart
							&& !v.front.children.empty)
					{
						tmp.v = v.front.children[0].text;
					} else {
						tmp.v = "";
					}
				}
			} else if (tmp.t == "inlineStr") {
				auto is_ = c.children.filter!(c => c.name == "is");
				tmp.v = extractData(is_.front);
			} else if (c.type == EntityType.elementStart) {
				auto v = c.children.filter!(c => c.name == "v");
				if (!v.empty && v.front.type == EntityType.elementStart
						&& !v.front.children.empty)
				{
					tmp.v = v.front.children[0].text;
				}
			}
			if (c.type == EntityType.elementStart) {
				auto f = c.children.filter!(c => c.name == "f");
				if (!f.empty && f.front.type == EntityType.elementStart) {
					tmp.f = f.front.children[0].text;
				}
			}
			ret ~= tmp;
		}
	}
	return ret;
}

Cell[] insertValueIntoCell(Cell[] cells, string[] ss) @trusted {
	immutable excepted = ["n", "s", "b", "e", "str", "inlineStr"];
	immutable same = ["n", "e", "str", "inlineStr"];
	foreach (ref Cell c; cells) {
		assert(canFind(excepted, c.t) || c.t.empty,
				format("'%s' not in [%s]", c.t, excepted));
		if (c.t.empty) {
			//c.xmlValue = convert(c.v);
			c.xmlValue = c.v.removeSpecialCharacter();
		} else if (canFind(same, c.t)) {
			//c.xmlValue = convert(c.v);
			c.xmlValue = c.v.removeSpecialCharacter();
		} else if (c.t == "b") {
			//logf("'%s' %s", c.v, c);
			//c.xmlValue = c.v == "1";
			c.xmlValue = c.v.removeSpecialCharacter();
		} else {
			if (!c.v.empty) {
				size_t idx = to!size_t(c.v);
				//logf("'%s' %s", c.v, idx);
				//c.xmlValue = ss[idx];
				c.xmlValue = ss[idx];
			}
		}
	}
	return cells;
}

Pos toPos(in string s) @safe pure {
	import std.string : indexOfAny;
	import std.math : pow;
	ptrdiff_t fn = s.indexOfAny("0123456789");
	enforce(fn != -1, s);
	RowOffset row = to!RowOffset(to!int(s[fn .. $]) - 1);
	ColOffset col = 0;
	string colS = s[0 .. fn];
	foreach (const idx, char c; colS) {
		col = col * 26 + (c - 'A' + 1);
	}
	return Pos(row, col - 1);
}

version(xlsxreader_test) @safe pure unittest {
	assert(toPos("A1").col == 0);
	assert(toPos("Z1").col == 25);
	assert(toPos("AA1").col == 26);
}

Pos elementMax(Pos a, Pos b) @safe pure nothrow @nogc {
    import std.algorithm.comparison : max;
	return Pos(max(a.row, b.row),
               max(a.col, b.col));
}

deprecated("this function is unused and will therefore be removed")
string specialCharacterReplacement(string s) @safe pure nothrow {
	import std.algorithm.iteration : substitute;
    import std.array : replace;
	// TODO: use substitute.array
	return s.replace("\"", "&quot;")
			.replace("'", "&apos;")
			.replace("<", "&lt;")
			.replace(">", "&gt;")
			.replace("&", "&amp;");
}

version(xlsxreader_test) @safe pure nothrow unittest {
	assert("&".specialCharacterReplacement == "&amp;");
}

deprecated("use specialCharacterReplacementReverseLazy instead")
string specialCharacterReplacementReverse(string s) @safe pure nothrow {
    import std.array : replace;
	return s.replace("&quot;", "\"")
			.replace("&apos;", "'")
			.replace("&lt;", "<")
			.replace("&gt;", ">")
			.replace("&amp;", "&");
}

version(none) // disabled because specialCharacterReplacementReverse is deprecated
version(xlsxreader_test) @safe pure nothrow unittest {
	assert("&quot;".specialCharacterReplacementReverse == "\"");
	assert("&apos;".specialCharacterReplacementReverse == "'");
	assert("&lt;".specialCharacterReplacementReverse == "<");
	assert("&gt;".specialCharacterReplacementReverse == ">");
	assert("&amp;".specialCharacterReplacementReverse == "&");
}

auto specialCharacterReplacementReverseLazy(string s) @safe pure /* TODO: nothrow @nogc */ {
    import std.algorithm.iteration : substitute;
	return s.substitute!("&quot;", "\"",
						 "&apos;", "'",
						 "&lt;", "<",
						 "&gt;", ">",
						 "&amp;", "&");
}

version(xlsxreader_test) @safe pure unittest {
	assert("&quot;".specialCharacterReplacementReverseLazy.equal("\""));
	assert("&apos;".specialCharacterReplacementReverseLazy.equal("'"));
	assert("&lt;".specialCharacterReplacementReverseLazy.equal("<"));
	assert("&gt;".specialCharacterReplacementReverseLazy.equal(">"));
	assert("&amp;".specialCharacterReplacementReverseLazy.equal("&"));
}

version(xlsxreader_test) @safe unittest {
	import std.math : isClose;
	auto r = readSheet("multitable.xlsx", "wb1");
	assert(isClose(r.table[12][5].xmlValue.to!double(), 26.74),
			format("%s", r.table[12][5])
		);

	assert(isClose(r.table[13][5].xmlValue.to!double(), -26.74),
			format("%s", r.table[13][5])
		);
}

version(xlsxreader_test) @safe unittest {
	auto s = readSheet("multitable.xlsx", "wb1");
	auto r = s.iterateRow!long(15, 1, 6);

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

version(xlsxreader_test) @safe unittest {
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

version(xlsxreader_test) @safe unittest {
	auto s = readSheet("multitable.xlsx", "Sheet3");
	// writeln(s.table[0][0].xmlValue);
	assert(s.table[0][0].xmlValue.to!long(),
			format("%s", s.table[0][0].xmlValue));
	//assert(s.table[0][0].canConvertTo(CellType.bool_));
}

version(xlsxreader_test) unittest {
	import std.file : dirEntries, SpanMode;
	import std.traits : EnumMembers;
	foreach (const de; dirEntries("xlsx_files/", "*.xlsx", SpanMode.depth)
			.filter!(a => a.name != "xlsx_files/data03.xlsx"))
	{
		//writeln(de.name);
		auto sn = sheetNames(de.name);
		foreach (const ref s; sn) {
			auto sheet = readSheet(de.name, s.name);
			foreach (const ref cell; sheet.cells) {
			}
		}
	}
}

version(xlsxreader_test) @safe unittest {
	auto sheet = readSheet("testworkbook.xlsx", "ws1");
	//writefln("%(%s\n%)", sheet.cells);
	//writeln(sheet.toString());
	assert(sheet.table[2][3].xmlValue.to!long() == 1337);

	auto c = sheet.getColumnLong(3, 2, 5);
	auto r = [1337, 2, 3];
	assert(equal(c, r), format("%s %s", c, sheet.toString()));

	auto c2 = sheet.getColumnString(4, 2, 5);
	string f2 = sheet.table[2][4].xmlValue;
	assert(f2 == "hello", f2);
	f2 = sheet.table[3][4].xmlValue;
	assert(f2 == "sil", f2);
	f2 = sheet.table[4][4].xmlValue;
	assert(f2 == "foo", f2);
	auto r2 = ["hello", "sil", "foo"];
	assert(equal(c2, r2), format("%s", c2));
}

version(xlsxreader_test) @safe unittest {
	import std.math : isClose;
	auto sheet = readSheet("toto.xlsx", "Trades");
	// writefln("%(%s\n%)", sheet.cells);

	auto r = sheet.getRowString(1, 0, 2).array;

	const double d = to!double(r[1]);
	assert(isClose(d, 38204642.510000));
}

version(xlsxreader_test) @safe unittest {
	const sheet = readSheet("leading_zeros.xlsx", "Sheet1");
	auto a2 = sheet.cells.filter!(c => c.r == "A2");
	assert(!a2.empty);
	assert(a2.front.xmlValue == "0012", format("%s", a2.front));
}

version(xlsxreader_test) @safe unittest {
	auto s = readSheet("datetimes.xlsx", "Sheet1");
	//writefln("%s\n%(%s\n%)", s.maxPos, s.cells);
	auto rslt = s.iterateColumn!DateTime(0, 0, 2);
	assert(!rslt.empty);

	auto target =
		[ DateTime(Date(1986,1,11), TimeOfDay.init)
		, DateTime(Date(1986,7,2), TimeOfDay.init)
		];
	assert(equal(rslt, target), format("\ngot: %s\nexp: %s\ntable %s", rslt
				, target, s.toString()));
}

version(xlsxreader_test) {
	import std.algorithm.comparison : equal;
}

/** Variant of Phobos `benchmark` that, instead of sum all run times, returns
	minimum of all run times as that is a more stable metric.
 */
private Duration[funs.length] benchmarkMin(funs...)(uint n) if (funs.length >= 1) {
    import std.algorithm.comparison : min;
    Duration[funs.length] result;
    auto sw = StopWatch(AutoStart.yes);
    foreach (const i, fun; funs) {
        result[i] = Duration.max;
		foreach (const j; 0 .. n) {
            sw.reset();
			sw.start();
            fun();
            sw.stop();
            result[i] = min(result[i], sw.peek());
        }
    }
    return result;
}
