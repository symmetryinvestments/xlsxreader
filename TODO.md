- Make `readCells` safe
- Cache parts or whole calculation of `dom` in `readCells`
- Replace calls to array append `~=` with `Appender.put()` taking a range if possible
- Clean up `parseRelationships` and `readSharedEntries`
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
