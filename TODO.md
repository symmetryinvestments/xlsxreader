- Clean up `readSharedEntries`
- Replace row.attributes.filter!(a => a.name == "r").front with row.attributeNamed("r")
- Make `readCells` safe
- Cache parts or whole calculation of `dom` in `readCells`
- In `readCells()`, can we reserve `ret` by looking up dimensions
  somewhere in DOM. Print string passed to DOM and look into it.
- In `readSharedEntries()`, can we reserve `ret` by looking up dimensions
  somewhere in DOM. Print string passed to DOM and look into it.
- Replace calls to array append `~=` with `Appender.put()` taking a range if possible
- Clean up `readCells`
- Clean up `insertValueIntoCell`
- Avoid cast to `immutable` in `convertToString` and return `inout(ubyte)` instead
- Call `assumeUnique` at the end of `readCells` if `Cell.members` are
  `immutable`.
- TODO: 1. contruct this lazily in Sheet upon use and cache
- TODO: 2. deprecate it
- const Relationships* sheetRel = rid in rels; // TODO: move this calculation to caller and pass Relationships as rels
- Replace ret ~= tORr.children[0].text.specialCharacterReplacementReverseLazy.to!string; with
  ret.put(tORr.children[0].text.specialCharacterReplacementReverseLazy)
- Make sure all the calls `expandTrusted`, `convertToString`, `parseDOM` are
  only called once (and cached) for every ArchiveMember by moving them into
  caching members of `File`.
- Move body of `parseRelationships` into `File.parseRelationships` and make the
  module-scope public `parseRelationships` a thin wrapper on top of it. Cache
  the calculation of `parseRelationships` in an AA mapping from `ArchiveMember`
  to `RelationshipsById` if needed and store that AA in private `File` member.
- Clean up `parseRelationships`
