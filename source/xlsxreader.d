module xslxreader;

import std.zip;
import std.file : exists, readText;
import std.exception : enforce;
import std.variant;

alias Sheet = Variant[][];

string[] sheetNames(string filename) {
	enforce(exists(filename), "File with name " ~ filename ~ " does not
			exist");	
}
