module xslxreader;

import std.algorithm : filter, map, sort, all, joiner, each;
import std.datetime : DateTime, Date, TimeOfDay;
import std.array : array;
import std.ascii : isDigit;
import std.conv : to;
import std.exception : enforce;
import std.file : read, exists, readText;
import std.format : format;
import std.range : tee;
import std.regex;
import std.stdio;
import std.traits : isIntegral, isFloatingPoint, isSomeString;
import std.typecons : tuple;
import std.utf : byChar;
import std.variant;
import std.zip;

import dxml.dom;

struct Pos {
	// zero based
	size_t row;
	// zero based
	size_t col;
}

alias Data = Algebraic!(long,double,string,DateTime,Date,TimeOfDay);

struct Cell {
	string loc;
	size_t row; // row[r]
	string t; // s or n, s for pointer, n for value, stored in v
	string r; // c[r]
	string v; // c.v the value or ptr
	string f; // c.f the formula
	Data value;
	Pos position;
}

struct Sheet {
	Cell[] cells;
	Data[][] table;
	Pos maxPos;

	void printTable() {
		long[] maxCol = new long[](maxPos.col + 1);
		foreach(row; this.table) {
			foreach(idx, Data col; row) {
				string s = col.visit!(
						(long l) => to!string(l),
						(double l) => to!string(l),
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

		foreach(row; this.table) {
			foreach(idx, Data col; row) {
				string s = col.visit!(
						(long l) => to!string(l),
						(double l) => to!string(l),
						(string l) => l,
						(DateTime l) => l.toISOExtString(),
						(Date l) => l.toISOExtString(),
						(TimeOfDay l) => l.toISOExtString(),
						() => "")
					();
				//writef("%*s, ", maxCol[idx], s);
			}
			//writeln();
		}
	}

	// Column

	Column!(T) iterateColumn(T)(size_t col, size_t start, size_t end) {
		return Column!(T)(&this, col, start, end);
	}

	Column!(long) iterateColumnLong(size_t col, size_t start, size_t end) {
		return Column!(long)(&this, col, start, end);
	}

	Column!(double) iterateColumnDouble(size_t col, size_t start, size_t end) {
		return Column!(double)(&this, col, start, end);
	}

	Column!(string) iterateColumnString(size_t col, size_t start, size_t end) {
		return Column!(string)(&this, col, start, end);
	}

	Column!(DateTime) iterateColumnDateTime(size_t col, size_t start,
			size_t end)
	{
		return Column!(DateTime)(&this, col, start, end);
	}

	Column!(Date) iterateColumnDate(size_t col, size_t start, size_t end) {
		return Column!(Date)(&this, col, start, end);
	}

	Column!(TimeOfDay) iterateColumnTimeOfDay(size_t col, size_t start,
			size_t end)
	{
		return Column!(TimeOfDay)(&this, col, start, end);
	}

	// Row

	Row!(T) iterateRow(T)(size_t row, size_t start, size_t end) {
		return Row!(T)(&this, row, start, end);
	}

	Row!(long) iterateRowLong(size_t row, size_t start, size_t end) {
		return Row!(long)(&this, row, start, end);
	}

	Row!(double) iterateRowDouble(size_t row, size_t start, size_t end) {
		return Row!(double)(&this, row, start, end);
	}

	Row!(string) iterateRowString(size_t row, size_t start, size_t end) {
		return Row!(string)(&this, row, start, end);
	}

	Row!(DateTime) iterateRowDateTime(size_t row, size_t start,
			size_t end)
	{
		return Row!(DateTime)(&this, row, start, end);
	}

	Row!(Date) iterateRowDate(size_t row, size_t start, size_t end) {
		return Row!(Date)(&this, row, start, end);
	}

	Row!(TimeOfDay) iterateRowTimeOfDay(size_t row, size_t start,
			size_t end)
	{
		return Row!(TimeOfDay)(&this, row, start, end);
	}
}

struct Row(T) {
	Sheet* sheet;
	const size_t row;
	size_t start;
	size_t end;
	size_t cur;

	T front;

	this(Sheet* sheet, size_t row, size_t start, size_t end) {
		this.sheet = sheet;
		this.row = row;
		this.start = start;
		this.end = end;
		this.cur = this.start;
		this.read();
	}

	@property bool empty() const {
		return this.cur >= this.end;
	}

	void popFront() {
		++this.cur;
		if(!this.empty) {
			this.read();
		}
	}

	private void read() {
		this.front = convertTo!T(this.sheet.table[this.row][this.cur]);
	}
}

struct Column(T) {
	Sheet* sheet;
	const size_t col;
	size_t start;
	size_t end;
	size_t cur;

	T front;

	this(Sheet* sheet, size_t col, size_t start, size_t end) {
		this.sheet = sheet;
		this.col = col;
		this.start = start;
		this.end = end;
		this.cur = this.start;
		this.read();
	}

	@property bool empty() const {
		return this.cur >= this.end;
	}

	void popFront() {
		++this.cur;
		if(!this.empty) {
			this.read();
		}
	}

	private void read() {
		this.front = convertTo!T(this.sheet.table[this.cur][this.col]);
	}
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
    if(d.day == 29 && d.month == 02 && d.year ==1900) {
        return 60;
	}

    // DMY to Modified Julian calculated with an extra subtraction of 2415019.
    long nSerialDate =
            int(( 1461 * ( d.year + 4800 + int(( d.month - 14 ) / 12) ) ) / 4) +
            int(( 367 * ( d.month - 2 - 12 * ( ( d.month - 14 ) / 12 ) ) ) / 12) -
            int(( 3 * ( int(( d.year + 4900 + int(( d.month - 14 ) / 12) ) / 100) ) ) / 4) +
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
	import std.math : lround;
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

T convertTo(T)(Data var) {
	static if(isSomeString!T) {
		return var.visit!(
				(long l) => to!string(l),
				(double l) => to!string(l),
				(string l) => l,
				(DateTime l) => l.toISOExtString(),
				(Date l) => l.toISOExtString(),
				(TimeOfDay l) => l.toISOExtString(),
				() => "")
			();
	} else static if(isIntegral!T) {
		return var.visit!(
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
	assert(false);
}


private ZipArchive readFile(string filename) {
	enforce(exists(filename), "File with name " ~ filename ~ " does not
			exist");

	auto file = new ZipArchive(read(filename));
	return file;
}

struct SheetNameId {
	string name;
	int id;
}

SheetNameId[] sheetNames(string filename) {
	auto file = readFile(filename);
	auto ams = file.directory;
	immutable wbStr = "xl/workbook.xml";
	enforce(wbStr in ams, "No workbook found");
	ubyte[] wb = file.expand(ams[wbStr]);
	string wbData = cast(string)wb;

	auto dom = parseDOM(wbData);
	assert(dom.children.length == 1);
	auto workbook = dom.children[0];
	assert(workbook.name == "workbook");
	auto sheetsRng = workbook.children.filter!(c => c.name == "sheets");
	assert(!sheetsRng.empty);
	return sheetsRng.front.children
		.map!(s => SheetNameId(
					s.attributes.filter!(a => a.name == "name").front.value,
					s.attributes.filter!(a => a.name == "sheetId").front
						.value.to!int()
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

Sheet readSheet(string filename, string sheetName) {
	SheetNameId[] sheets = sheetNames(filename);
	auto sRng = sheets.filter!(s => s.name == sheetName);
	enforce(!sRng.empty, "No sheet with name " ~ sheetName
			~ " found in file " ~ filename);
	return readSheet(filename, sRng.front.id);
}

Sheet readSheet(string filename, int sheetId) {
	auto file = readFile(filename);
	auto ams = file.directory;
	immutable ss = "xl/sharedStrings.xml";
	enforce(ss in ams, "No sharedStrings found");
	Data[] sharedStrings = readSharedEntries(file, ams[ss]);
	writeln(sharedStrings);

	immutable wsStr = "xl/worksheets/sheet" ~ to!string(sheetId) ~ ".xml";
	enforce(wsStr in ams, wsStr ~ " Not found in "
			~ ams.keys().joiner(" ").to!string()
		);
	Sheet ret;
	ret.cells = insertValueIntoCell(readCells(file, ams[wsStr]), sharedStrings);
	Pos maxPos;
	foreach(ref c; ret.cells) {
		c.position = toPos(c.r);
		maxPos = elementMax(maxPos, c.position);
	}
	ret.maxPos = maxPos;
	ret.table = new Data[][](ret.maxPos.row + 1, ret.maxPos.col + 1);
	foreach(c; ret.cells) {
		ret.table[c.position.row][c.position.col] = c.value;
	}
	return ret;
}

Data[] readSharedEntries(ZipArchive za, ArchiveMember am) {
	ubyte[] ss = za.expand(am);
	string ssData = cast(string)ss;
	auto dom = parseDOM(ssData);
	assert(dom.children.length == 1);
	auto sst = dom.children[0];
	assert(sst.name == "sst");
	auto siRng = sst.children.filter!(c => c.name == "si");
	assert(!siRng.empty);
	return siRng
		.map!(si => si.children[0])
		.tee!(t => assert(t.name == "t"))
		//.tee!(t => writeln(t))
		.map!(t => Data(convert(t.children[0].text)))
		.array;
}

private bool canConvertToLong(string s) {
	return s.byChar.all!isDigit();
}

private immutable rs = r"[0-9][0-9]*\.[0-9]*";
private auto rgx = ctRegex!rs;

private bool canConvertToDouble(string s) {
	auto cap = matchAll(s, rgx);
	return cap.empty || cap.front.hit != s ? false : true;
}

Data convert(string s) {
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
		import std.algorithm : canFind;
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

Cell[] readCells(ZipArchive za, ArchiveMember am) {
	ubyte[] ss = za.expand(am);
	string ssData = cast(string)ss;
	auto dom = parseDOM(ssData);
	assert(dom.children.length == 1);
	auto ws = dom.children[0];
	assert(ws.name == "worksheet");
	auto sdRng = ws.children.filter!(c => c.name == "sheetData");
	assert(!sdRng.empty);
	auto rows = sdRng.front.children
		.filter!(r => r.name == "row");

	Cell[] ret;
	foreach(row; rows) {
		foreach(c; row.children.filter!(r => r.name == "c")) {
			Cell tmp;
			tmp.row = row.attributes.filter!(a => a.name == "r")
				.front.value.to!size_t();
			tmp.r = c.attributes.filter!(a => a.name == "r")
				.front.value;
			auto t = c.attributes.filter!(a => a.name == "t");
			if(t.empty) {
				writeln("Found a strange empty cell");
				continue;
			}
			tmp.t = t.front.value;
			auto v = c.children.filter!(c => c.name == "v");
			enforce(!v.empty);
			tmp.v = v.front.children[0].text;
			auto f = c.children.filter!(c => c.name == "f");
			if(!f.empty) {
				tmp.f = f.front.children[0].text;
			}
			ret ~= tmp;
		}
	}
	return ret;
}

Cell[] insertValueIntoCell(Cell[] cells, Data[] ss) {
	foreach(ref Cell c; cells) {
		assert(c.t == "n" || c.t == "s", format("%s", c));
		if(c.t == "n") {
			c.value = convert(c.v);
		} else {
			size_t idx = to!size_t(c.v);
			c.value = ss[idx];
		}
	}
	return cells;
}

Pos toPos(string s) {
	import std.algorithm : reverse;
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

unittest {
	import std.math : approxEqual;
	auto r = readSheet("multitable.xlsx", "wb1");
	assert(approxEqual(r.table[12][5].get!double(), 26.74),
			format("%s", r.table[12][5])
		);
}

unittest {
	import std.algorithm.comparison : equal;
	auto s = readSheet("multitable.xlsx", "wb1");
	auto r = s.iterateRow!long(15, 1, 6);
	assert(equal(r, [1, 2, 3, 4, 5]), format("%s", r));
}

unittest {
	import std.algorithm.comparison : equal;
	auto s = readSheet("multitable.xlsx", "wb2");
	writefln("%s\n%(%s\n%)", s.maxPos, s.cells);
	auto rslt = s.iterateColumn!Date(1, 1, 6);
	auto target = [Date(2019,5,01), Date(2016,12,27), Date(1976,7,23),
		 Date(1986,7,2), Date(2038,1,19)
	];
	assert(equal(rslt, target), format("\n%s\n%s", rslt, target));
}
