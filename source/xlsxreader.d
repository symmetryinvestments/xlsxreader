module xslxreader;

import std.algorithm : filter, map;
import std.array : array;
import std.conv : to;
import std.exception : enforce;
import std.file : read, exists, readText;
import std.stdio;
import std.variant;
import std.zip;

import dxml.dom;

alias Sheet = Variant[][];

struct SheetNameId {
	string name;
	int id;
}

SheetNameId[] sheetNames(string filename) {
	enforce(exists(filename), "File with name " ~ filename ~ " does not
			exist");	

	auto file = new ZipArchive(read(filename));
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
		.array;
}

unittest {
	auto r = sheetNames("multitable.xlsx");
	assert(r[0].name == "wb1");
	assert(r[0].id == 1);
}
