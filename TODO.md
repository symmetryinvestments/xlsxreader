- Make sure all the calls `expandTrusted`, `convertToString`, `parseDOM` are
  only called once (and cached) for every ArchiveMember by moving them into
  caching members of `File`.
- Move body of `parseRelationships` into `File.parseRelationships` and make the
  module-scope public `parseRelationships` a thin wrapper on top of it. Cache
  the calculation of `parseRelationships` in an AA mapping from `ArchiveMember`
  to `RelationshipsById` if needed and store that AA in private `File` member.
- Use a single call to `substitute().array()` in `specialCharacterReplacementReverse`
- Use a single call to `substitute().array()` in `specialCharacterReplacement`
