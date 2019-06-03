module xslxreader;

import std.algorithm : filter, map, sort, all, joiner;
import std.typecons : tuple;
import std.ascii : isDigit;
import std.array : array;
import std.regex;
import std.conv : to;
import std.utf : byChar;
import std.exception : enforce;
import std.file : read, exists, readText;
import std.range : tee;
import std.stdio;
import std.variant;
import std.zip;

import dxml.dom;

alias Sheet = Variant[][];

struct SheetNameId {
	string name;
	int id;
}

private ZipArchive readFile(string filename) {
	enforce(exists(filename), "File with name " ~ filename ~ " does not
			exist");

	auto file = new ZipArchive(read(filename));
	return file;
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

Variant[] readSharedEntries(ZipArchive za, ArchiveMember am) {
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
		.map!(t => Variant(convert(t.children[0].text)))
		.array;
}

bool canConvertToLong(string s) {
	return s.byChar.all!isDigit();
}

immutable rs = r"[0-9][0-9]*\.[0-9]*";
auto rgx = regex(rs);

bool canConvertToDouble(string s) {
	auto cap = matchAll(s, rgx);
	return cap.empty || cap.front.hit != s ? false : true;
}

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

Variant convert(string s) {
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
		return Variant(to!long(s));
	} else if(canConvertToDouble(s)) {
		return Variant(to!double(s));
	} else {
		return Variant(replaceStrings(s));
	}
}

struct Cell {
	string loc;
	size_t row;
	string t;
	string v;
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
	auto tmp = sdRng.front.children
		.filter!(r => r.name == "row")
		.map!(row => row.children
			.map!(rc => tuple(
				row.attributes.filter!(rowA => rowA.name == "r").front.value
					.to!long(),
				row.children
			))
		)
		.joiner
		.map!(c => c[1].map!(v => tuple(
					c[0],
					v.attributes.filter!(cr => cr.name == "r").front.value,
					v.attributes.filter!(cr => cr.name == "t").front.value,
					v.children.filter!(i => i.name == "v")
				)
			)
		)
		.joiner
		.map!(v => v[3].map!(vi => tuple(
					v[0],
					v[1],
					v[2],
					vi.children[0].text
				)
			)
		)
		.joiner
		//.map!(t1 => tuple(
		//		t1[0],
		//		t1[1].attributes.filter!(ra => ra.name == "r").front.value,
		//		t1[1]
		//	)
		//k)
		//.map!(r => tuple(
		//	r.attributes.filter!(ra => ra.name == "r").front.value.to!size_t(),
		//	r)
		//)
		//.map!(rc => tuple(rc[0], rc[1].children
		//			.filter!(c => c.name == "c").array)
		//)
		.array;
	writefln("%(%s\n%)", tmp);
	writeln(typeof(tmp).stringof);
	return Cell[].init;
}

Sheet readSheet(string filename, int sheetId) {
	auto file = readFile(filename);
	auto ams = file.directory;
	immutable ss = "xl/sharedStrings.xml";
	enforce(ss in ams, "No sharedStrings found");
	Variant[] sharedStrings = readSharedEntries(file, ams[ss]);
	writeln(sharedStrings);

	immutable wsStr = "xl/worksheets/sheet" ~ to!string(sheetId) ~ ".xml";
	enforce(wsStr in ams, wsStr ~ " Not found in "
			~ ams.keys().joiner(" ").to!string()
		);
	Cell[] cells = readCells(file, ams[wsStr]);
	return Sheet.init;
}

unittest {
	auto r = readSheet("multitable.xlsx", "wb1");
	writeln(r);
}
