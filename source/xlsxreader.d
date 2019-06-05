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
				writef("%*s, ", maxCol[idx], s);
			}
			writeln();
		}
	}

	Row!(T) iterateRow(T)(size_t row, size_t start, size_t end) {
		return Row!(T)(row, start, end, start);
	}
}

T convertTo(T)(Variant var) {
	if(is(T == Variant)) {
		return var;
	} else static if(isSomeString!T) {
		return var.coerce!T();
	} else static if(is(T == DateTime)) {
	} else static if(is(T == Date)) {
	} else static if(is(T == TimeOfDay)) {
	}
	assert(false);
}

struct Row(T) {
	size_t row;
	size_t start;
	size_t end;
	size_t cur;
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

bool canConvertToLong(string s) {
	return s.byChar.all!isDigit();
}

immutable rs = r"[0-9][0-9]*\.[0-9]*";
auto rgx = ctRegex!rs;

bool canConvertToDouble(string s) {
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
	writefln("%s\n%(%s\n%)", r.maxPos, r.cells);
	assert(approxEqual(r.table[12][5].get!double(), 26.74),
			format("%s", r.table[12][5])
		);
}
